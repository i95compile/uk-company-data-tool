"""Enrich the lead lists in output/<AREA>/ with publicly listed business contact details.

Usage:
  python enrich.py --limit 25 --shuffle          test run on 25 random companies
  python enrich.py --areas M,B                   whole areas
  python enrich.py                               everything (resumable; re-run to continue)
Options: --search  also try the DuckDuckGo HTML page as a fallback (opt-in, slow, stops
                   for good at the first CAPTCHA/anomaly page)
         --no-progress  plain logs
Reads   output/<AREA>/leads*.xlsx (never modified)
Writes  enriched/<AREA>/leads*_enriched.xlsx, enriched/enrichment_report.txt,
        enriched/progress.jsonl (resume state), enriched/enrich.log
Needs   pip install -r requirements.txt

How a website is accepted (never guessed):
  High    the site shows the CompanyNumber (UK companies must display it on their website)
  Medium  exact company name in the page title/footer/h1 AND the registered postcode on the site
  Uncertain  name matches but nothing else; URL recorded, no contacts extracted
Only business channels are kept: role emails on the company's own domain (info@, sales@...),
free-mail addresses named after the company, phone numbers listed on the site.
"""
import argparse
import json
import logging
import os
import random
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path
from urllib.parse import parse_qs, quote_plus, unquote, urljoin, urlsplit
from urllib.robotparser import RobotFileParser

import dns.exception
import dns.resolver
import pandas as pd
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm

import main as leads   # reuse the lead-list Excel writer and column layout

# ---- CONFIG ----
CONTACT_EMAIL = os.environ.get("ENRICH_CONTACT_EMAIL", "").strip()   # sent in the User-Agent so site owners can reach you
WORKERS = 5
GLOBAL_GAP, DOMAIN_GAP, SEARCH_GAP = 1.0, 3.0, 5.0   # seconds between requests
HTTP_TIMEOUT = 10
MAX_CANDIDATES = 15                   # guessed domains per company
MAX_EXTRA_PAGES = 4                   # pages fetched per site besides the homepage
# ----------------

HERE = Path(__file__).parent
SRC_DIR, OUT_DIR = HERE / "output", HERE / "enriched"
PROGRESS = OUT_DIR / "progress.jsonl"
USER_AGENT = f"LeadEnrichBot/1.0 (business contact research; contact: {CONTACT_EMAIL})"
ROBOTS_TOKEN = "LeadEnrichBot"
HEADERS = {"User-Agent": USER_AGENT, "Accept": "text/html,application/xhtml+xml",
           "Accept-Language": "en-GB,en;q=0.8"}
LEGAL = {"ltd", "limited", "llp", "plc"}
FILLER = {"uk", "services", "service", "construction", "constructions", "building", "builders",
          "contractors", "contracting", "group", "solutions", "and", "the", "co", "company"}
TLDS = [".co.uk", ".com", ".uk"]
ROLE_LOCALS = {"info", "hello", "contact", "enquiries", "enquiry", "office", "admin", "sales",
               "accounts", "bookings", "quotes", "quote", "estimating", "support", "mail",
               "team", "general", "reception", "jobs", "work", "projects", "business"}
FREEMAIL = {"gmail.com", "googlemail.com", "hotmail.com", "hotmail.co.uk", "outlook.com",
            "live.com", "live.co.uk", "yahoo.com", "yahoo.co.uk", "icloud.com", "aol.com",
            "btinternet.com", "sky.com", "msn.com"}
# Registers/directories show company numbers too, but are not the company's own site.
AGGREGATORS = ("gov.uk", "company-information.service.gov.uk", "endole.co.uk",
               "opencorporates.com", "companieshouse.id", "find-and-update.company-information.service.gov.uk",
               "linkedin.com", "facebook.com", "instagram.com", "twitter.com", "x.com", "tiktok.com",
               "youtube.com", "checkcompany.co.uk", "companycheck.co.uk", "duedil.com",
               "bizdb.co.uk", "yell.com", "companiesintheuk.co.uk", "ukcompanyinfo.co.uk",
               "company-information.org.uk", "find-company.co.uk", "thegazette.co.uk",
               "google.com", "bing.com", "duckduckgo.com", "trustpilot.com", "checkatrade.com",
               "mybuilder.com", "ratedpeople.com", "trustatrader.com", "bark.com", "yelp.co.uk",
               "192.com", "cylex-uk.co.uk", "freeindex.co.uk", "companyhub.nz", "companieslist.co.uk")
PARKED = ("domain is for sale", "buy this domain", "this domain may be for sale",
          "domain parking", "is parked free", "parked domain", "hugedomains", "sedo.com",
          "dan.com", "afternic")
PAGE_KEYS = [("contact", 0), ("about", 1), ("privacy", 2), ("terms", 3), ("legal", 3),
             ("cookie", 4)]
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:\+44\s?\(0\)\s?|\+44\s?|(?<!\d)0)(?:\d[\s-]?){9,10}(?!\d)")
NEW_COLS = [("Enrichment Status", 20), ("Enrichment Detail", 40), ("Website", 34),
            ("Website Confidence", 12), ("Verification Evidence", 40),
            ("Verification Source URL", 40), ("Contact Page URL", 40), ("Business Email", 30),
            ("Email Source URL", 40), ("Business Phone", 18), ("Phone Source URL", 40),
            ("Candidate Website (unverified)", 34), ("Found Via", 14), ("Checked On", 12)]

log = logging.getLogger("enrich")


# ---------------- polite HTTP ----------------
class Blocked(Exception):
    """robots.txt disallows the URL."""


class SearchStopped(Exception):
    """The search page answered with a CAPTCHA/anomaly page; search is off for this run."""


class Throttle:
    """At most one request per GLOBAL_GAP overall and one per `gap` seconds per host."""
    def __init__(self):
        self.lock = threading.Lock()
        self.next_any, self.next_host = 0.0, {}

    def wait(self, host, gap=DOMAIN_GAP):
        with self.lock:
            t = max(time.monotonic(), self.next_any, self.next_host.get(host, 0.0))
            self.next_any, self.next_host[host] = t + GLOBAL_GAP, t + gap
        time.sleep(max(0.0, t - time.monotonic()))


THROTTLE = Throttle()
RESOLVER = dns.resolver.Resolver()
RESOLVER.lifetime = RESOLVER.timeout = 4
_robots, _robots_lock = {}, threading.Lock()
SEARCH_OK = threading.Event()
SEARCH_STOP = []          # reason the search fallback was switched off, if it was


def site_key(host):
    host = host.lower().split(":")[0]
    return host[4:] if host.startswith("www.") else host


def is_aggregator(host):
    h = site_key(host)
    return any(h == a or h.endswith("." + a) for a in AGGREGATORS)


def robots_allows(url):
    p = urlsplit(url)
    origin = f"{p.scheme}://{p.netloc.lower()}"
    with _robots_lock:
        rp = _robots.get(origin)
    if rp is None:
        rp = RobotFileParser()
        try:
            status, _, _, body = http_get(origin + "/robots.txt", check_robots=False)
            if status in (401, 403):
                rp.disallow_all = True
            elif status == 200:
                rp.parse(body.decode("utf-8", "replace").splitlines())
            else:
                rp.allow_all = True
        except requests.RequestException:
            rp.allow_all = True
        with _robots_lock:
            _robots[origin] = rp
    return rp.can_fetch(ROBOTS_TOKEN, url)


def http_get(url, check_robots=True, gap=DOMAIN_GAP):
    """Throttled GET; redirects followed by hand so each hop is throttled and robots-checked.
    Returns (status, final_url, content_type, body_bytes)."""
    for _ in range(6):
        if check_robots and not robots_allows(url):
            raise Blocked(url)
        THROTTLE.wait(site_key(urlsplit(url).netloc), gap)
        r = requests.get(url, headers=HEADERS, timeout=HTTP_TIMEOUT, allow_redirects=False,
                         stream=True)
        try:
            if r.status_code in (301, 302, 303, 307, 308) and r.headers.get("Location"):
                url = urljoin(url, r.headers["Location"])
                if urlsplit(url).scheme not in ("http", "https"):
                    raise requests.RequestException("redirect to non-http scheme")
                continue
            body = b""
            for chunk in r.iter_content(65536):
                body += chunk
                if len(body) > 1_500_000:
                    break
            return r.status_code, url, r.headers.get("Content-Type", ""), body
        finally:
            r.close()
    raise requests.TooManyRedirects(url)


def fetch_html(url):
    """(body, final_url) for a 200 HTML page, else (None, final_url or None). Blocked propagates."""
    try:
        status, final, ctype, body = http_get(url)
    except Blocked:
        raise
    except requests.RequestException as e:
        log.debug("fetch failed %s: %r", url, e)
        return None, None
    ok = status == 200 and ("html" in ctype.lower() or not ctype)
    return (body, final) if ok else (None, final)


# ---------------- matching helpers ----------------
def norm(s):
    s = re.sub(r"['’`]", "", (s or "").lower()).replace("&", " and ")
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def name_tokens(name):
    toks = norm(name).split()
    while toks and toks[-1] in LEGAL:
        toks.pop()
    return toks


def candidate_domains(name):
    toks = name_tokens(name)
    core = [t for t in toks if t not in FILLER] or toks
    variants = (toks, core, [t for t in toks if t != "and"], toks + ["uk"], core + ["uk"])
    bases = []
    for sep in ("", "-"):                           # joined forms first, hyphenated after
        for b in variants:
            v = sep.join(b)
            if len(v.replace("-", "")) >= 4 and v not in bases:
                bases.append(v)
    return [b + t for b in bases for t in TLDS][:MAX_CANDIDATES]


def resolve_host(domain):
    for host in (domain, "www." + domain):
        for rtype in ("A", "AAAA"):
            try:
                RESOLVER.resolve(host, rtype)
                return host
            except dns.exception.DNSException:
                continue
    return None


def phone_norm(raw):
    d = re.sub(r"\D", "", raw)
    if d.startswith("440"):
        d = "0" + d[3:]
    elif d.startswith("44"):
        d = "0" + d[2:]
    return d if len(d) in (10, 11) and d[0] == "0" and d[1] in "12378" else None


def analyse(body, url, co):
    """Facts from one page: verification hits, contact candidates, links to follow."""
    soup = BeautifulSoup(body, "html.parser")
    mailtos = [unquote(a["href"][7:].split("?")[0]).strip()
               for a in soup.find_all("a", href=re.compile(r"^mailto:", re.I))]
    tels = [a["href"][4:] for a in soup.find_all("a", href=re.compile(r"^tel:", re.I))]
    links = []
    for a in soup.find_all("a", href=True):
        txt = (a.get_text(" ") + " " + a["href"]).lower()
        for key, prio in PAGE_KEYS:
            if key in txt:
                links.append((prio, urljoin(url, a["href"].split("#")[0])))
                break
    title = norm(soup.title.get_text()) if soup.title else ""
    for t in soup(["script", "style", "noscript", "template"]):
        t.decompose()
    blocks = (soup.find_all(["footer", "h1"]) + soup.find_all(id=re.compile("footer", re.I))
              + soup.find_all(class_=re.compile("footer", re.I)))
    head = norm(" ".join(b.get_text(" ") for b in blocks))
    text = re.sub(r"\s+", " ", soup.get_text(" "))
    num = co["number"]
    nums = [num] + ([num.lstrip("0")] if num[0] == "0" and len(num.lstrip("0")) >= 7 else [])
    pc = re.sub(r"\s+", "", co["postcode"].upper())
    name = " ".join(co["tokens"])
    return {
        "num": any(re.search(rf"(?<![A-Za-z0-9]){re.escape(n)}(?![A-Za-z0-9])", text) for n in nums),
        "name": len(name) >= 4 and any(f" {name} " in f" {s} " for s in (title, head)),
        "pc": len(pc) >= 5 and bool(re.search(
            rf"\b{re.escape(pc[:-3])}\s?{re.escape(pc[-3:])}\b", text.upper())),
        "parked": any(p in text.lower() for p in PARKED),
        "emails": mailtos + [m.rstrip(".") for m in EMAIL_RE.findall(text)],
        "phones": [(t, True) for t in tels] + [(m, False) for m in PHONE_RE.findall(text)],
        "links": links,
    }


def pick_email(emails, site, tokens):
    """First business email: role address on the site's domain, or free-mail named after the firm."""
    for e in emails:
        local, _, dom = e.lower().partition("@")
        if dom == site or dom.endswith("." + site):
            if local in ROLE_LOCALS:
                return e.lower()
        elif dom in FREEMAIL and any(len(t) >= 4 and t in local for t in tokens):
            return e.lower()
    return None


def fmt_phone(d):
    """UK display format: 020 7946 0018, 0161 496 0018, 0306 999 0018, 01632 960018, 07700 900018."""
    if d[:2] == "02":
        return f"{d[:3]} {d[3:7]} {d[7:]}"
    if d[:2] in ("03", "08") or d[:3] == "011" or (d[:2] == "01" and d[3] == "1"):
        return f"{d[:4]} {d[4:7]} {d[7:]}"
    return f"{d[:5]} {d[5:]}"


def pick_phone(phones):
    for raw, via_tel in sorted(phones, key=lambda p: not p[1]):     # tel: links first
        if d := phone_norm(unquote(raw)):
            return fmt_phone(d)
    return None


# ---------------- per-company pipeline ----------------
def check_site(host, co):
    """Fetch homepage + up to MAX_EXTRA_PAGES linked pages and judge the match.
    Returns (state, info) with state in verified / uncertain / failed / parked / blocked."""
    body = final = None
    for scheme in ("https", "http"):
        try:
            body, final = fetch_html(f"{scheme}://{host}/")
        except Blocked:
            return "blocked", {"detail": f"robots.txt disallows {host}"}
        if body:
            break
    if not body:
        return "failed", {"detail": f"{host}: no HTML homepage"}
    p = urlsplit(final)
    base, site = f"{p.scheme}://{p.netloc}", site_key(p.netloc)
    if is_aggregator(site):
        return "failed", {"detail": f"{host} redirects to directory {site}"}
    pages = [(final, analyse(body, final, co))]
    if pages[0][1]["parked"]:
        return "parked", {"detail": f"{site} looks parked/for sale"}

    seen, queue = {final.rstrip("/")}, []
    for prio, u in sorted(pages[0][1]["links"]):
        if site_key(urlsplit(u).netloc) == site and u.rstrip("/") not in seen:
            seen.add(u.rstrip("/"))
            queue.append(u)
    if not any("contact" in u.lower() for u in queue):
        queue.insert(0, base + "/contact")
    contact_url = None
    for u in queue[:MAX_EXTRA_PAGES]:
        if any(f["num"] for _, f in pages) and contact_url:
            break                                   # verified and contact page seen
        try:
            b, fu = fetch_html(u)
        except Blocked:
            continue
        if b:
            pages.append((fu, analyse(b, fu, co)))
            if "contact" in u.lower() and not contact_url:
                contact_url = fu

    hit = lambda k: next((u for u, f in pages if f[k]), None)
    num_url, name_url, pc_url = hit("num"), hit("name"), hit("pc")
    if num_url:
        conf, ev, src = "High", f"Company number {co['number']} shown on site", num_url
    elif name_url and pc_url:
        conf, ev, src = "Medium", f"Company name in title/heading/footer and postcode {co['postcode']} on site", pc_url
    elif name_url:
        return "uncertain", {"candidate": base, "detail": f"Name matches on {name_url}; "
                             "no company number or postcode found"}
    else:
        return "failed", {"detail": f"{site}: no company number/name+postcode"}

    # contacts: contact page first, then the rest
    ordered = sorted(pages, key=lambda pf: pf[0] != contact_url)
    email = phone = email_src = phone_src = None
    for u, f in ordered:
        if not email and (e := pick_email(f["emails"], site, co["tokens"])):
            email, email_src = e, u
        if not phone and (ph := pick_phone(f["phones"])):
            phone, phone_src = ph, u
    return "verified", {"website": base, "confidence": conf, "evidence": ev, "source": src,
                        "contact_url": contact_url, "email": email, "email_src": email_src,
                        "phone": phone, "phone_src": phone_src}


def search_hosts(co):
    """Opt-in DuckDuckGo HTML fallback. Returns candidate hosts (directories filtered out)."""
    if not SEARCH_OK.is_set():
        return []
    q = f'"{co["name"]}" {co["postcode"]}'
    url = "https://html.duckduckgo.com/html/?kl=uk-en&q=" + quote_plus(q)
    try:
        status, _, _, body = http_get(url, gap=SEARCH_GAP)
    except Blocked:
        SEARCH_OK.clear()
        raise SearchStopped("robots.txt disallows the search page")
    except requests.RequestException:
        return []
    text = body.decode("utf-8", "replace")
    if status != 200 or "anomaly" in text.lower() or "captcha" in text.lower():
        SEARCH_OK.clear()
        raise SearchStopped(f"search answered HTTP {status} / anomaly page")
    hosts = []
    for a in BeautifulSoup(text, "html.parser").select("a.result__a"):
        href = a.get("href", "")
        if "uddg=" in href:
            href = parse_qs(urlsplit(href).query).get("uddg", [""])[0]
        h = urlsplit(href).netloc.lower()
        if h and not is_aggregator(h) and h not in hosts:
            hosts.append(h)
    return hosts[:4]


def enrich_company(co, prior, use_search):
    """Full pipeline for one company -> result dict (stored in progress.jsonl)."""
    res = dict(prior or {})
    if not res.get("guessed"):
        resolved, blocked = 0, False
        for domain in candidate_domains(co["name"]):
            host = resolve_host(domain)
            if not host:
                continue
            resolved += 1
            state, info = check_site(host, co)
            if state == "verified":
                return {**info, "status": f"Verified - {info['confidence']}", "via": "domain guess",
                        "detail": f"{resolved} guessed domain(s) live", "guessed": True}
            if state == "uncertain" and not res.get("candidate"):
                res.update(candidate=info["candidate"], uncertain_detail=info["detail"])
            blocked |= state == "blocked"
        res.update(guessed=True, resolved=resolved, blocked=blocked)
    if use_search and not res.get("searched") and SEARCH_OK.is_set():
        try:
            hosts = search_hosts(co)
        except SearchStopped as e:
            SEARCH_STOP.append(str(e))
            hosts = None                            # not marked as searched: retried next run
        for host in hosts or []:
            state, info = check_site(host, co)
            if state == "verified":
                return {**info, "status": f"Verified - {info['confidence']}", "via": "web search",
                        "detail": "found via DuckDuckGo, verified on site", "guessed": True,
                        "searched": True}
            if state == "uncertain" and not res.get("candidate"):
                res.update(candidate=info["candidate"], uncertain_detail=info["detail"])
        if hosts is not None:
            res["searched"] = True
    if res.get("candidate"):
        res.update(status="Uncertain match", detail=res["uncertain_detail"])
    elif res.get("blocked"):
        res.update(status="No website found", detail="a candidate site was blocked by robots.txt")
    elif res.get("resolved"):
        res.update(status="No website found",
                   detail=f"{res['resolved']} guessed domain(s) live, none verified")
    else:
        res.update(status="No website found", detail="no guessed domain resolved")
    if not use_search:
        res["detail"] += " (web search not used)"
    return res


# ---------------- I/O ----------------
def load_progress():
    done = {}
    if PROGRESS.exists():
        for line in PROGRESS.read_text(encoding="utf-8").splitlines():
            try:
                rec = json.loads(line)
                done[rec["number"]] = rec
            except (ValueError, KeyError):
                continue                            # half-written last line after a crash
    return done


def load_leads(areas):
    frames = []
    for f in sorted(SRC_DIR.glob("*/leads*.xlsx")):
        area = f.parent.name
        if areas and area not in areas:
            continue
        d = pd.read_excel(f, dtype={"Company Number": str, "Postcode": str})
        d["_area"], d["_file"] = area, f.name
        frames.append(d)
    assert frames, f"No lead files found in {SRC_DIR}"
    return pd.concat(frames, ignore_index=True)


def result_columns(rec):
    r = (rec or {}).get("result") or {}
    if not r:
        return {"Enrichment Status": "Not processed"}
    return {"Enrichment Status": r.get("status"), "Enrichment Detail": r.get("detail"),
            "Website": r.get("website"), "Website Confidence": r.get("confidence"),
            "Verification Evidence": r.get("evidence"), "Verification Source URL": r.get("source"),
            "Contact Page URL": r.get("contact_url"), "Business Email": r.get("email"),
            "Email Source URL": r.get("email_src"), "Business Phone": r.get("phone"),
            "Phone Source URL": r.get("phone_src"),
            "Candidate Website (unverified)": r.get("candidate"),
            "Found Via": r.get("via"), "Checked On": rec.get("checked")}


def write_outputs(df, done):
    cols = [c for c in leads.COLUMNS] + NEW_COLS
    extra = pd.DataFrame([result_columns(done.get(n)) for n in df["Company Number"]],
                         index=df.index)
    full = pd.concat([df, extra.reindex(columns=[c for c, _ in NEW_COLS])], axis=1)
    full["Companies House Link"] = leads.COMPANY_URL + full["Company Number"]
    for (area, fname), g in full.groupby(["_area", "_file"], sort=True):
        folder = OUT_DIR / area
        folder.mkdir(parents=True, exist_ok=True)
        leads.write_one(folder / fname.replace(".xlsx", "_enriched.xlsx"), g, cols)
    return full


def write_report(full, sel_numbers, run_secs, newly, use_search, search_note):
    s = full[full["Company Number"].isin(sel_numbers)]
    n = len(s)
    st = s["Enrichment Status"].fillna("")
    ver = s[st.str.startswith("Verified")]
    pct = lambda k, d: f"{k} ({100 * k / d:.1f}%)" if d else f"{k}"
    remaining = (full["Enrichment Status"] == "Not processed").sum()
    per = run_secs / newly if newly else None
    lines = [
        "ENRICHMENT REPORT", f"Date: {date.today():%d/%m/%Y}",
        f"Companies in selected lead files: {len(full)}",
        f"Companies in this selection (processed now or earlier): {n}",
        f"Websites verified: {pct(len(ver), n)}",
        f"  High (company number on site): {(st == 'Verified - High').sum()}",
        f"  Medium (name + postcode on site): {(st == 'Verified - Medium').sum()}",
        f"Uncertain matches (candidate URL only, no contacts taken): {pct((st == 'Uncertain match').sum(), n)}",
        f"No website found: {pct((st == 'No website found').sum(), n)}",
        f"  no guessed domain resolved: {s['Enrichment Detail'].fillna('').str.startswith('no guessed').sum()}",
        f"  domain(s) live but not verified: {s['Enrichment Detail'].fillna('').str.contains('none verified').sum()}",
        f"  blocked by robots.txt: {s['Enrichment Detail'].fillna('').str.contains('robots').sum()}",
        f"Errors (retried next run): {(st == 'Error').sum()}",
        f"Business email found: {pct(ver['Business Email'].notna().sum(), n)} "
        f"(of verified: {pct(ver['Business Email'].notna().sum(), len(ver))})",
        f"Business phone found: {pct(ver['Business Phone'].notna().sum(), n)} "
        f"(of verified: {pct(ver['Business Phone'].notna().sum(), len(ver))})",
        f"Contact page found: {pct(ver['Contact Page URL'].notna().sum(), n)}",
        f"Verified but missing email: {ver['Business Email'].isna().sum()}; "
        f"missing phone: {ver['Business Phone'].isna().sum()}",
        f"Web search fallback: {'on' if use_search else 'off'}{search_note}",
        f"Processing time this run: {run_secs:.0f} s for {newly} companies"
        + (f" ({per:.1f} s/company)" if per else ""),
    ]
    if per and remaining:
        lines.append(f"Not yet processed: {remaining}; at this pace ~{remaining * per / 3600:.1f} h")
    (OUT_DIR / "enrichment_report.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return lines


def main():
    ap = argparse.ArgumentParser(description="Add public business contact details to lead lists")
    ap.add_argument("--areas", help="comma-separated postcode areas (default: all)")
    ap.add_argument("--limit", type=int, default=0, help="only the first N companies of the selection")
    ap.add_argument("--shuffle", action="store_true",
                    help="random but repeatable order (seed 1), so --limit gives a fair sample")
    ap.add_argument("--search", action="store_true", help="opt-in DuckDuckGo HTML fallback")
    ap.add_argument("--no-progress", action="store_true")
    args = ap.parse_args()
    areas = {a.strip().upper() for a in args.areas.split(",")} if args.areas else None
    if not CONTACT_EMAIL:
        sys.exit("Set ENRICH_CONTACT_EMAIL to an address site owners can reach you at "
                 "(it is sent in the User-Agent).")

    OUT_DIR.mkdir(exist_ok=True)
    logging.basicConfig(filename=OUT_DIR / "enrich.log", level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s")
    df = load_leads(areas)
    done = load_progress()
    order = list(range(len(df)))
    if args.shuffle:
        random.Random(1).shuffle(order)
    sel = df.iloc[order[:args.limit] if args.limit else order]

    def needs_work(num):
        r = (done.get(num) or {}).get("result")
        if not r or r.get("status") == "Error":
            return True
        return args.search and not r.get("searched") and not r["status"].startswith("Verified")

    todo = [r for r in sel.itertuples(index=False) if needs_work(r[1])]
    print(f"{len(sel)} companies selected, {len(sel) - len(todo)} already done, {len(todo)} to process")
    if args.search:
        SEARCH_OK.set()
    lock, t0, search_note = threading.Lock(), time.time(), ""

    def work(row):
        co = {"name": row[0], "number": row[1],
              "postcode": row[3] if isinstance(row[3], str) else "",
              "tokens": name_tokens(row[0])}
        prior = (done.get(co["number"]) or {}).get("result")
        try:
            res = enrich_company(co, prior if prior and prior.get("status") != "Error" else None,
                                 args.search)
        except Exception as e:                      # one bad site must not stop the run
            log.exception("company %s failed", co["number"])
            res = {"status": "Error", "detail": repr(e)[:200]}
        rec = {"number": co["number"], "checked": f"{date.today():%d/%m/%Y}", "result": res}
        with lock:
            done[co["number"]] = rec
            if res["status"] != "Error":
                with open(PROGRESS, "a", encoding="utf-8") as f:
                    f.write(json.dumps(rec) + "\n")
        log.info("%s %s: %s | %s", co["number"], co["name"], res["status"], res.get("detail"))
        return co, res

    show = not args.no_progress and sys.stdout.isatty()
    with ThreadPoolExecutor(WORKERS) as ex, \
            tqdm(total=len(todo), desc="Enriching", unit="co", disable=not show) as bar:
        futs = [ex.submit(work, r) for r in todo]
        for fut in as_completed(futs):
            co, res = fut.result()
            if res["status"].startswith("Verified") or not show:
                tqdm.write(f"  {co['name']}: {res['status']} {res.get('website') or ''}")
            bar.update(1)
    if SEARCH_STOP:
        search_note = f" - stopped early: {SEARCH_STOP[0]}"

    full = write_outputs(df, done)
    lines = write_report(full, set(sel["Company Number"]), time.time() - t0, len(todo),
                         args.search, search_note)
    print("\n".join(lines))


if __name__ == "__main__":
    main()

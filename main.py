"""Turn a Companies House Free Company Data Product CSV into formatted Excel lead lists,
one folder per postcode area.

Usage:  python main.py [input.csv] [--areas M,B,LS] [--no-progress]   (default: all areas)
Output: output/<AREA>/{leads.xlsx,report.txt}, output/summary.xlsx, output/summary_report.txt
Needs:  pip install -r requirements.txt
Edit the CONFIG block below to change SIC / mode.
Company-level fields only; no officer or personal data is read or written.

Big files: the CSV is streamed in chunks, twice. Pass 1 counts addresses (as 64-bit hashes);
pass 2 applies the cheap filters per chunk so only surviving rows stay in memory.
"""
import argparse
import ctypes
import re
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter
from tqdm import tqdm

# ---- CONFIG ----
SIC_CODES = ["41", "43"]           # leading digits of SICCode.SicText_1..4
MODE = "new"                       # "new" or "accounts_due"
DAYS = 90                          # window size in days
SHARED_ADDRESS_THRESHOLD = 10
VIRTUAL_KEYWORDS = ["suite", "unit", "floor", "office", "box"]
CHUNKSIZE = 200_000
MAX_ROWS_PER_FILE = 1_000_000      # Excel's hard limit is 1,048,576 rows per sheet

# ----------------

# ================= LEADS PIPELINE =================
OUT_DIR = Path(__file__).parent / "output"
UNKNOWN = "_unknown"
COMPANY_URL = "https://find-and-update.company-information.service.gov.uk/company/"
SIC_COLS = [f"SICCode.SicText_{i}" for i in range(1, 5)]
SIC_PREFIXES = tuple(SIC_CODES)
ADDR_COLS = ["RegAddress.CareOf", "RegAddress.POBox", "RegAddress.AddressLine1",
             "RegAddress.AddressLine2", "RegAddress.PostTown", "RegAddress.County",
             "RegAddress.Country", "RegAddress.PostCode"]
REQUIRED = ["CompanyName", "CompanyNumber", "CompanyCategory", "CompanyStatus",
            "IncorporationDate", "Accounts.NextDueDate", *ADDR_COLS, *SIC_COLS]
STRIP_FIRST = ["CompanyStatus", "CompanyCategory", "RegAddress.PostCode", "IncorporationDate",
               "Accounts.NextDueDate", *SIC_COLS]       # needed clean before filtering
KEEP = ["CompanyName", "CompanyNumber", *ADDR_COLS, "_sic", "_ind",
        "IncorporationDate_d", "Accounts.NextDueDate_d", "area"]
COLUMNS = [("Company Name", 36), ("Company Number", 15), ("Address", 60), ("Postcode", 11),
           ("Industry", 40), ("SIC Code", 9), ("Incorporation Date", 18),
           ("Accounts Due Date", 18), ("Possible Virtual Office", 22),
           ("Companies House Link", 24)]
DATE_COLS = {"Incorporation Date", "Accounts Due Date"}
STEPS = ["not_active", "not_plc", "no_sic", "date", "shared", "dupe"]


def mem_mb(peak=False):
    """Process memory in MB (Windows working set via ctypes; elsewhere peak RSS)."""
    if sys.platform == "win32":
        class PMC(ctypes.Structure):
            _fields_ = [("cb", ctypes.c_ulong), ("PageFaultCount", ctypes.c_ulong)] + [
                (n, ctypes.c_size_t) for n in (
                    "PeakWorkingSetSize", "WorkingSetSize", "QuotaPeakPagedPoolUsage",
                    "QuotaPagedPoolUsage", "QuotaPeakNonPagedPoolUsage",
                    "QuotaNonPagedPoolUsage", "PagefileUsage", "PeakPagefileUsage")]
        k32, psapi = ctypes.windll.kernel32, ctypes.windll.psapi
        k32.GetCurrentProcess.restype = ctypes.c_void_p
        psapi.GetProcessMemoryInfo.argtypes = [ctypes.c_void_p, ctypes.POINTER(PMC), ctypes.c_ulong]
        pmc = PMC(); pmc.cb = ctypes.sizeof(PMC)
        psapi.GetProcessMemoryInfo(k32.GetCurrentProcess(), ctypes.byref(pmc), pmc.cb)
        return (pmc.PeakWorkingSetSize if peak else pmc.WorkingSetSize) / 2**20
    import resource
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024


def read_chunks(path, cols):
    """Stream the CSV in chunks, loading only `cols` (header whitespace stripped)."""
    want = set(cols)
    reader = pd.read_csv(path, usecols=lambda c: c.strip() in want, dtype=str,
                         chunksize=CHUNKSIZE, keep_default_na=False,
                         encoding_errors="replace", low_memory=False)
    for n, chunk in enumerate(reader, 1):
        chunk.columns = chunk.columns.str.strip()
        yield n, chunk


def addr_hash(chunk):
    """64-bit hash of each row's normalised address (lowercase, punctuation/space collapsed)."""
    s = chunk[ADDR_COLS[0]].str.cat([chunk[c] for c in ADDR_COLS[1:]], sep=" ")
    s = s.str.lower().str.replace(r"[^a-z0-9]+", " ", regex=True).str.strip()
    return pd.util.hash_pandas_object(s, index=False).to_numpy()


def postcode_area(pc):
    """M1 1AA -> M, SW1A 1AA -> SW, ME4 4AA -> ME; missing/unparseable -> _unknown."""
    pc = pc.str.upper().str.replace(r"\s+", "", regex=True)
    return pc.str.extract(r"^([A-Z]{1,2})(?=\d)")[0].fillna(UNKNOWN)


PROGRESS = True     # set in main(): off with --no-progress or when stdout is not a terminal


def bar(desc, total, unit="rows"):
    return tqdm(total=total, desc=desc, unit=unit, unit_scale=(unit == "rows"),
                mininterval=0.5, disable=not PROGRESS or total == 0)


def advance(pb, k):
    """Update a row bar; if the estimate was too low, raise the total so it ends at 100%."""
    if pb.total and pb.n + k > pb.total:
        pb.total = pb.n + k
    pb.update(k)


def estimate_rows(path):
    """Exact row count for small files; otherwise file size / average bytes per row of the
    first 10,000 rows. Returns None if it can't be estimated (bar then has no total)."""
    try:
        size = path.stat().st_size
        with open(path, "rb") as f:
            header = f.readline()
            head = [f.readline() for _ in range(10_000)]
        if not head[-1]:        # file ended inside the sample: count exactly
            return sum(len(c) for c in pd.read_csv(path, usecols=[0], dtype=str,
                                                   chunksize=CHUNKSIZE, encoding_errors="replace"))
        return int((size - len(header)) / (sum(map(len, head)) / len(head)))
    except (OSError, ValueError):
        return None


def shared_address_hashes(path, total=None):
    """Pass 1: hashes of addresses used by more than SHARED_ADDRESS_THRESHOLD companies."""
    parts = []
    with bar("Pass 1/2: counting addresses", total) as pb:
        for _, chunk in read_chunks(path, ADDR_COLS):
            parts.append(addr_hash(chunk))
            advance(pb, len(chunk))
    if not parts:
        return np.array([], dtype=np.uint64)
    hashes, counts = np.unique(np.concatenate(parts), return_counts=True)
    return hashes[counts > SHARED_ADDRESS_THRESHOLD]


def count(stats, step, area):
    for a, n in area.value_counts().items():
        if n:
            stats[a][step] += int(n)


def process_chunk(chunk, shared, lo, hi, date_col, areas, stats, outside, bad_dates):
    """Pass 2 on one chunk: assign area, apply every filter, return the surviving rows."""
    for c in STRIP_FIRST:
        chunk[c] = chunk[c].str.strip()
    area = postcode_area(chunk["RegAddress.PostCode"])
    if areas is not None:
        m = area.isin(areas)
        outside[0] += int((~m).sum())
        chunk, area = chunk[m], area[m]
    count(stats, "read", area)

    def keep(mask, step, chunk, area):
        count(stats, step, area[~mask])
        return chunk[mask], area[mask]

    chunk, area = keep(chunk["CompanyStatus"].str.lower() == "active", "not_active", chunk, area)
    chunk, area = keep(chunk["CompanyCategory"].str.contains("Private Limited Company",
                       case=False, regex=False), "not_plc", chunk, area)
    # first SIC field starting with a wanted prefix -> (code, description)
    code = pd.Series("", index=chunk.index)
    ind = code.copy()
    for col in SIC_COLS:
        m = chunk[col].str.extract(r"^\s*(\d+)\s*-?\s*(.*)$")
        hit = (code == "") & m[0].str.startswith(SIC_PREFIXES).fillna(False).astype(bool)
        code, ind = code.mask(hit, m[0]), ind.mask(hit, m[1].str.strip())
    chunk, area = keep((code != "").to_numpy(), "no_sic", chunk, area)
    chunk = chunk.assign(_sic=code[chunk.index], _ind=ind[chunk.index])
    for f in ("IncorporationDate", "Accounts.NextDueDate"):
        chunk[f + "_d"] = pd.to_datetime(chunk[f], format="%d/%m/%Y", errors="coerce")
        bad = chunk[chunk[f + "_d"].isna()]
        for num, raw, a in zip(bad["CompanyNumber"], bad[f], area[bad.index]):
            stats[a]["bad_dates"] += 1
            if len(bad_dates[a]) < 50:
                bad_dates[a].append((num, f, raw))
    chunk, area = keep(chunk[date_col + "_d"].between(lo, hi).to_numpy(), "date", chunk, area)
    chunk, area = keep(~np.isin(addr_hash(chunk), shared), "shared", chunk, area)
    chunk = chunk.assign(area=area)
    for c in ("CompanyName", "CompanyNumber", *ADDR_COLS[:-1]):
        chunk[c] = chunk[c].str.strip()
    return chunk[KEEP]


def build_output(df):
    """Survivors -> the final column layout (plus a helper _area column)."""
    pat = re.compile(r"\b(" + "|".join(map(re.escape, VIRTUAL_KEYWORDS)) + r")\b", re.I)
    rows = list(zip(*(df[c] for c in ADDR_COLS)))
    return pd.DataFrame({
        "Company Name": df["CompanyName"],
        "Company Number": df["CompanyNumber"],
        "Address": [", ".join(p for p in r if p) for r in rows],
        "Postcode": df["RegAddress.PostCode"],
        "Industry": df["_ind"].astype("category"),
        "SIC Code": df["_sic"].astype("category"),
        "Incorporation Date": df["IncorporationDate_d"],
        "Accounts Due Date": df["Accounts.NextDueDate_d"],
        "Possible Virtual Office": ["Yes" if pat.search(" ".join(p for p in r[:-1] if p))
                                    else "No" for r in rows],
        "Companies House Link": COMPANY_URL + df["CompanyNumber"],
        "_area": df["area"],
    })


def write_one(path, frame, cols):
    """One xlsx file in constant_memory mode: format/width/filter/freeze first, then rows in order."""
    names = [n for n, _ in cols]
    wb = xlsxwriter.Workbook(str(path), {"constant_memory": True, "strings_to_urls": False})
    ws = wb.add_worksheet("Leads")
    head = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78",
                          "valign": "vcenter"})
    date = wb.add_format({"num_format": "dd/mm/yyyy"})
    link = wb.add_format({"font_color": "#0563C1", "underline": 1})
    for i, (_, w) in enumerate(cols):
        ws.set_column(i, i, w)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(frame), len(cols) - 1)
    for i, n in enumerate(names):
        ws.write_string(0, i, n, head)
    for r, row in enumerate(frame[names].itertuples(index=False), 1):
        for c, v in enumerate(row):
            if pd.isna(v):
                continue
            n = names[c]
            if n == "Companies House Link":     # HYPERLINK() formula: no 65,530-link sheet cap
                ws.write_formula(r, c, f'=HYPERLINK("{v}","View on Companies House")', link,
                                 "View on Companies House")
            elif n in DATE_COLS:
                ws.write_datetime(r, c, v.to_pydatetime(), date)
            else:
                ws.write_string(r, c, str(v))
    wb.close()


def write_leads(folder, frame):
    """Write leads.xlsx, or leads_part1.xlsx, leads_part2.xlsx... above MAX_ROWS_PER_FILE rows."""
    folder.mkdir(parents=True, exist_ok=True)
    for old in folder.glob("leads*.xlsx"):
        old.unlink()
    n_parts = -(-len(frame) // MAX_ROWS_PER_FILE)
    names = []
    for k in range(n_parts):
        name = "leads.xlsx" if n_parts == 1 else f"leads_part{k + 1}.xlsx"
        write_one(folder / name, frame.iloc[k * MAX_ROWS_PER_FILE:(k + 1) * MAX_ROWS_PER_FILE],
                  [c for c in COLUMNS if c[0] in frame.columns])
        names.append(name)
    return names


def area_report(area, st, bad, ctx, files):
    f = ctx["date_col"]
    lines = [f"DATA QUALITY REPORT - postcode area {area}",
             f"Input file: {ctx['input']}",
             f"Mode: {MODE} ({DAYS} days: {ctx['lo']:%d/%m/%Y} to {ctx['hi']:%d/%m/%Y}); SIC {SIC_CODES}",
             f"Rows read (this area): {st['read']}",
             f"Dropped - status not Active: {st['not_active']}",
             f"Dropped - category not Private Limited Company: {st['not_plc']}",
             f"Dropped - no SIC code starting with {SIC_CODES}: {st['no_sic']}",
             f"Dropped - {f} outside window (or unparseable): {st['date']}",
             f"Dropped - address shared by more than {SHARED_ADDRESS_THRESHOLD} companies: {st['shared']}",
             f"Dropped - duplicate CompanyNumber: {st['dupe']}",
             f"Flagged possible virtual office: {st['virtual']}",
             f"Final row count: {st['final']}",
             f"Output: {', '.join(files)}",
             f"Missing/unparseable dates among rows reaching the date step: {st['bad_dates']}"]
    if area == UNKNOWN:
        lines.insert(4, "(Rows with a missing or unparseable postcode.)")
    lines += [f"  {n}: {fld} = {raw!r}" for n, fld, raw in bad[:50]]
    if st["bad_dates"] > 50:
        lines.append(f"  ... and {st['bad_dates'] - 50} more")
    return lines


def run_pipeline(args):
    t0 = time.time()
    assert MODE in ("new", "accounts_due"), "MODE must be 'new' or 'accounts_due'"
    here = Path(__file__).parent
    path = Path(args.csv) if args.csv else sorted(here.glob("*.csv"))[0]
    header = pd.read_csv(path, nrows=0, dtype=str, encoding_errors="replace").columns.str.strip()
    missing = [c for c in REQUIRED if c not in header]
    assert not missing, f"Missing required columns: {missing}"

    today = pd.Timestamp.today().normalize()
    if MODE == "new":
        date_col, lo, hi = "IncorporationDate", today - pd.Timedelta(days=DAYS), today
    else:
        date_col, lo, hi = "Accounts.NextDueDate", today, today + pd.Timedelta(days=DAYS)

    total = estimate_rows(path)
    shared = shared_address_hashes(path, total)

    stats = defaultdict(Counter)          # area -> step -> count
    bad_dates = defaultdict(list)         # area -> [(CompanyNumber, field, raw)]
    outside = [0]                         # rows skipped because of --areas
    kept, rows_read, kept_rows = [], 0, 0
    with bar("Pass 2/2: filtering", total) as pb:
        for _, chunk in read_chunks(path, REQUIRED):
            rows_read += len(chunk)
            res = process_chunk(chunk, shared, lo, hi, date_col, args.areas, stats, outside,
                                bad_dates)
            kept.append(res)
            kept_rows += len(res)
            pb.set_postfix(kept=f"{kept_rows:,}", refresh=False)
            advance(pb, len(chunk))

    df = pd.concat(kept, ignore_index=True) if kept else pd.DataFrame(columns=KEEP)
    del kept
    dupes = df.duplicated("CompanyNumber")
    count(stats, "dupe", df.loc[dupes, "area"])
    df = df[~dupes].sort_values(date_col + "_d", ascending=(MODE == "accounts_due"),
                                kind="stable")
    out = build_output(df)
    del df

    files = {}
    wb = bar("Writing Excel files", out["_area"].nunique(), unit="area")
    for area, g in out.groupby("_area", observed=True, sort=True):
        wb.set_postfix_str(str(area), refresh=False)
        st = stats[area]
        st["final"] = len(g)
        st["virtual"] = int((g["Possible Virtual Office"] == "Yes").sum())
        files[area] = write_leads(OUT_DIR / area, g)
        ctx = dict(input=path.name, date_col=date_col, lo=lo, hi=hi)
        (OUT_DIR / area / "report.txt").write_text(
            "\n".join(area_report(area, st, bad_dates[area], ctx, files[area])) + "\n",
            encoding="utf-8")
        wb.update(1)
    wb.close()

    areas = sorted(set(a for a in stats if stats[a]["read"]) | (args.areas or set()))
    write_summary(areas, stats, files)
    totals = Counter()
    for a in areas:
        totals.update({k: v for k, v in stats[a].items()})
    runtime = time.time() - t0
    lines = ["SUMMARY REPORT", f"Input file: {path.name}",
             f"Mode: {MODE} ({DAYS} days: {lo:%d/%m/%Y} to {hi:%d/%m/%Y}); SIC {SIC_CODES}; "
             f"areas: {', '.join(sorted(args.areas)) if args.areas else 'all'}",
             f"Rows read (whole file): {rows_read}"]
    if args.areas:
        lines.append(f"Skipped - area not in --areas: {outside[0]}")
    lines += [f"Areas seen: {len(areas)}; with results: {len(files)}",
              f"Rows with missing/unparseable postcode ({UNKNOWN}): {stats[UNKNOWN]['read']}",
              f"Rows read (selected areas): {totals['read']}",
              f"Dropped - status not Active: {totals['not_active']}",
              f"Dropped - category not Private Limited Company: {totals['not_plc']}",
              f"Dropped - no SIC code starting with {SIC_CODES}: {totals['no_sic']}",
              f"Dropped - {date_col} outside window (or unparseable): {totals['date']}",
              f"Dropped - address shared by more than {SHARED_ADDRESS_THRESHOLD} companies: {totals['shared']}",
              f"Dropped - duplicate CompanyNumber: {totals['dupe']}",
              f"Flagged possible virtual office: {totals['virtual']}",
              f"Final row count: {totals['final']}",
              f"Missing/unparseable dates among rows reaching the date step: {totals['bad_dates']}",
              f"Runtime: {runtime:.1f} s", f"Peak memory: {mem_mb(peak=True):.0f} MB"]
    (OUT_DIR / "summary_report.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    tqdm.write(f"Done: {rows_read:,} rows read, {totals['final']:,} kept, {len(files)} areas "
               f"written, {runtime:.1f} s. Reports: {OUT_DIR / 'summary_report.txt'}")


def write_summary(areas, stats, files):
    cols = {"Area": areas}
    for key, title in [("read", "Rows Read"), ("not_active", "Dropped - Not Active"),
                       ("not_plc", "Dropped - Not Private Ltd"), ("no_sic", "Dropped - No SIC Match"),
                       ("date", "Dropped - Outside Date Window"), ("shared", "Dropped - Shared Address"),
                       ("dupe", "Dropped - Duplicate"), ("final", "Rows After Filters"),
                       ("virtual", "Possible Virtual Office")]:
        cols[title] = [stats[a][key] for a in areas]
    cols["Output Path"] = [", ".join(f"{a}/{f}" for f in files[a]) if a in files else "" for a in areas]
    df = pd.DataFrame(cols)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT_DIR / "summary.xlsx", engine="xlsxwriter") as xw:
        df.to_excel(xw, index=False, sheet_name="Summary")
        ws = xw.sheets["Summary"]
        ws.set_column(0, len(df.columns) - 2, 20)
        ws.set_column(len(df.columns) - 1, len(df.columns) - 1, 40)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("csv", nargs="?", help="input CSV (default: first .csv in this folder)")
    ap.add_argument("--areas", help="comma-separated postcode areas, e.g. M,B,LS (default: all)")
    ap.add_argument("--no-progress", action="store_true", help="disable progress bars")
    args = ap.parse_args()
    global PROGRESS
    PROGRESS = not args.no_progress and sys.stdout.isatty()
    args.areas = {a.strip().upper().replace("_UNKNOWN", UNKNOWN)
                  for a in args.areas.split(",") if a.strip()} if args.areas else None
    run_pipeline(args)


if __name__ == "__main__":
    main()

# UK Company Data Tool

Turns the Companies House **Free Company Data Product** (a ~2.7 GB CSV of every live UK company) into filtered, formatted Excel lead lists, one folder per postcode area. An optional second step looks for each company's own website and records publicly listed business contact details, but only when the site can be verified as belonging to that company.

The default configuration finds **active private limited companies in construction (SIC 41 / 43) incorporated in the last 90 days**.

## How it works

| Script | Input | Output |
|---|---|---|
| `main.py` | Companies House CSV | `output/<AREA>/leads.xlsx`, per-area `report.txt`, `output/summary.xlsx`, `output/summary_report.txt` |
| `enrich.py` | `output/<AREA>/leads*.xlsx` (read-only) | `enriched/<AREA>/leads*_enriched.xlsx`, `enriched/enrichment_report.txt`, `enriched/progress.jsonl`, `enriched/enrich.log` |

### `main.py`: lead lists

The CSV is never loaded whole. It is streamed with `pandas.read_csv(chunksize=200_000)`, reading only the columns the pipeline needs, in two passes:

1. **Pass 1** hashes each row's normalised registered address to a 64-bit integer. Addresses shared by more than `SHARED_ADDRESS_THRESHOLD` companies (formation agents, accountants, virtual offices) are marked.
2. **Pass 2** filters each chunk and keeps only the rows that survive:
   - status is Active
   - category is Private Limited Company
   - at least one SIC code starts with a configured prefix
   - incorporation date (or accounts due date) falls in the window
   - the address is not a shared address

   Duplicate company numbers are then removed.

Survivors are written with `xlsxwriter` in `constant_memory` mode. Each sheet has a header style, date formats, filters, frozen panes and a Companies House link per row. Areas with more than 1,000,000 rows are split into `leads_part1.xlsx`, `leads_part2.xlsx`, and so on, to stay under Excel's sheet limit. Every area gets a data-quality report showing how many rows each filter dropped.

Only company-level fields are read. No officer or personal data is used.

### `enrich.py`: website and contact enrichment (optional)

For each company, `enrich.py` guesses likely domains from the company name, checks which ones resolve in DNS, and fetches the homepage and a few linked pages (contact, about, privacy, and so on). A website is accepted only when:

- **High confidence:** the site shows the company's registration number, which UK companies must display on their website.
- **Medium confidence:** the exact company name appears in the title, heading or footer, **and** the registered postcode appears on the site.
- **Uncertain:** the name matches but nothing else does. The URL is recorded, but no contacts are taken from it.

Only business channels are kept: role addresses (`info@`, `sales@`, and so on) on the company's own domain, free-mail addresses named after the company, and phone numbers listed on the site. Each value is recorded with its source URL.

The crawler is deliberately polite. It obeys `robots.txt`, makes at most 1 request per second overall and 1 every 3 seconds per host, follows redirects manually so each hop is checked, caps page size, and identifies itself with a contact address in its User-Agent. Progress is saved after every company, so an interrupted run resumes where it stopped.

## Prerequisites

- Python 3.11 or newer (developed on Python 3.13 with pandas 3.0)
- About 3 GB of free disk space for the input CSV
- Internet access, for `enrich.py` only

## Installation

```bash
git clone https://github.com/<your-username>/uk-company-data-tool.git
cd uk-company-data-tool
python -m venv .venv
# Windows: .venv\Scripts\activate    macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

### Get the input data

1. Download **BasicCompanyDataAsOneFile-YYYY-MM-DD.zip** from the Companies House [Free Company Data Product](https://download.companieshouse.gov.uk/en_output.html) page.
2. Unzip it into the project folder.

The data is published under the Open Government Licence. The CSV is excluded from Git by `.gitignore`.

## Usage

```bash
# All postcode areas, using the first .csv in the project folder
python main.py

# Explicit input file, selected postcode areas only
python main.py BasicCompanyDataAsOneFile-2026-10-01.csv --areas M,B,LS

# Plain log output (no progress bars)
python main.py --no-progress
```

```bash
# enrich.py needs an address that site owners can contact you at
export ENRICH_CONTACT_EMAIL=you@example.com        # PowerShell: $env:ENRICH_CONTACT_EMAIL="you@example.com"

python enrich.py --limit 25 --shuffle   # trial run on 25 random companies
python enrich.py --areas M,B            # whole areas
python enrich.py                        # everything; re-run to resume
python enrich.py --search               # also try DuckDuckGo as a fallback (slow, opt-in)
```

## Configuration

Edit the `CONFIG` block at the top of each script.

**`main.py`**

| Setting | Default | Meaning |
|---|---|---|
| `SIC_CODES` | `["41", "43"]` | Leading digits of the SIC codes to keep |
| `MODE` | `"new"` | `"new"`: incorporated in the last `DAYS` days. `"accounts_due"`: accounts due in the next `DAYS` days |
| `DAYS` | `90` | Window size, relative to today's date |
| `SHARED_ADDRESS_THRESHOLD` | `10` | Drop addresses used by more than this many companies |
| `VIRTUAL_KEYWORDS` | `suite, unit, floor, office, box` | Words that set the "Possible Virtual Office" flag |
| `CHUNKSIZE` | `200_000` | Rows per chunk. Lower it to reduce memory use |
| `MAX_ROWS_PER_FILE` | `1_000_000` | Split threshold for Excel files |

**`enrich.py`**

| Setting | Default | Meaning |
|---|---|---|
| `ENRICH_CONTACT_EMAIL` (environment variable) | required | Sent in the User-Agent |
| `WORKERS` | `5` | Parallel worker threads |
| `GLOBAL_GAP`, `DOMAIN_GAP`, `SEARCH_GAP` | `1.0`, `3.0`, `5.0` s | Minimum delay between requests, overall, per host and for search |
| `MAX_CANDIDATES` | `15` | Guessed domains checked per company |
| `MAX_EXTRA_PAGES` | `4` | Pages fetched per site besides the homepage |

## Performance

These figures come from a full run of `main.py` on a 2.7 GB Free Company Data Product CSV (5,704,711 rows) on a Windows 11 desktop:

| Metric | Result |
|---|---|
| Runtime (both passes plus writing 122 area workbooks) | 150 s |
| Peak memory (Windows peak working set) | 504 MB |
| Rows kept (SIC 41/43, incorporated in the last 90 days) | 10,218 |

Memory depends mainly on two things: the address-hash array from pass 1 (8 bytes per row, about 45 MB for 5.7 M rows) and the number of rows that survive the filters. Wide filters, such as all SIC codes or a long `DAYS` window, keep more rows in memory before writing.

`enrich.py` is network-bound by design. A 2,000-company trial averaged about **7.4 s per company** and verified websites for **3.5%** of companies.

## Limitations

- **Date window is relative to today.** Re-running on a later day gives a different window. Use a fresh Companies House snapshot to get recent incorporations.
- **Address matching is exact after normalisation** (lowercase, punctuation removed). Variants of the same address are not merged, and 64-bit hash collisions are theoretically possible, though unlikely.
- **"Possible Virtual Office" is a keyword heuristic.** Many genuine trading addresses contain words like "Unit" or "Suite".
- **Registered address is not trading address.** Many small companies register at an accountant's or formation agent's address.
- **Enrichment coverage is low.** Most new companies have no website yet. Results marked *Uncertain* must not be treated as verified contacts.
- **The DuckDuckGo fallback** depends on an HTML page that can change or rate-limit. It switches itself off at the first CAPTCHA.
- **No automated test suite yet.**
- **Peak memory reporting** uses the Windows API on Windows and `resource` elsewhere. Elsewhere it reports peak RSS, so it is not directly comparable.

## Data protection

The input is public data, but the enriched output contains business contact details. If you use them for marketing, you are responsible for complying with UK GDPR and PECR. For example, unsolicited marketing email to sole traders and partnerships needs consent, and every message needs an opt-out.

## Licence

No licence has been chosen yet. All rights reserved by the author until one is added.

# Vietnamese Bank Financial Report Downloader

Downloads annual audited financial reports (BCTC) for 26 Vietnamese listed banks from [cafef.vn](https://cafef.vn), covering 2010–2025.

## Banks covered

ABB, ACB, BAB, BID, BVB, CTG, EIB, HDB, KLB, LPB, MBB, MSB, NAB, NVB, OCB, PGB, SCB, SGB, SHB, SSB, STB, TCB, TPB, VCB, VIB, VPB

## Report types downloaded (per bank per year)

| Tag | Report name |
|---|---|
| `Consolidated` | Báo cáo tài chính hợp nhất năm YYYY (đã kiểm toán) |
| `Annual` | Báo cáo tài chính năm YYYY (đã kiểm toán) |

> Some banks use one naming convention, others use the other. The script tries both and downloads whichever exists.

## Setup

Requires Python 3.13.

```bash
python3.13 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## Usage

```bash
# Download everything (all 26 banks, 2010–2025)
.venv/bin/python download_bctc.py

# Dry run — see what would be downloaded without fetching files
.venv/bin/python download_bctc.py --dry-run

# Specific banks only
.venv/bin/python download_bctc.py --banks vcb tcb acb

# Specific years only
.venv/bin/python download_bctc.py --years 2022 2023 2024

# Retry broken links (run after a normal download pass)
.venv/bin/python download_bctc.py --retry-broken
```

## Output

```
financial_reports/
├── ABB/
│   ├── ABB_BCTC_Consolidated_2010.pdf
│   ├── ABB_BCTC_Consolidated_2011.pdf
│   └── ...
├── ACB/
│   └── ...
└── download_log.json
```

Each file is named `{BANK}_BCTC_{Type}_{YEAR}.pdf`.

`download_log.json` records every bank/year/type with one of these statuses:

| Status | Meaning |
|---|---|
| `ok` | Downloaded successfully |
| `exists` | Already on disk, skipped |
| `not_found` | cafef has no record for this bank/year/type |
| `broken_link` | cafef has a record but the CDN file is gone |
| `api_error` | API request failed |

## Broken link recovery

When a URL returns 404, the script automatically tries alternative CDN hosts (`cafef1.mediacdn.vn`, `cafef.mediacdn.vn`, `static.cafef.vn`). If all fail, run `--retry-broken` which re-queries the API across all exchange types (HOSE/HNX/UPCOM) to find a working URL.

## Known gaps

21 bank-years have no report available on cafef at all (both naming variants absent):

- **ABB** 2012 · **BAB** 2010–2011 · **BVB** 2010 · **KLB** 2010
- **LPB** 2011 · **NVB** 2012 · **OCB** 2011 · **SCB** 2011, 2020–2024
- **SGB** 2010, 2017 · **TPB** 2011

> SCB 2020–2024: SCB was placed under special State Bank control in 2022 and stopped publishing reports on cafef.

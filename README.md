# Vietnamese Bank Financial Report Downloader & Parser

Downloads annual audited financial reports (BCTC) for Vietnamese banks from [cafef.vn](https://cafef.vn) (with [vietstock.vn](https://finance.vietstock.vn) as fallback), and extracts key financial metrics into Excel using a vision LLM.

## Banks covered

ABB, ACB, AGR, BAB, BID, BVB, CTG, EIB, HDB, KLB, LPB, MBB, MSB, NAB, NVB, OCB, PCB, PGB, SCB, SGB, SHB, SSB, STB, TCB, TPB, VAB, VBB, VCB, VIB, VPB

*(29 banks, 2010–2024)*

## Scripts

| Script | Purpose |
|---|---|
| `download_bctc.py` | Download PDFs from cafef.vn (vietstock fallback) |
| `parse_bctc.py` | Extract metrics from PDFs → Excel (full run) |
| `fill_missing.py` | Fill empty cells in an existing Excel from PDFs |

---

## Setup

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example .env   # then fill in your keys
```

> `pdf2image` requires `poppler`: `brew install poppler`

---

## 1. Download PDFs — `download_bctc.py`

```bash
# Download everything (all banks, 2010–2024)
python download_bctc.py

# Dry run — see what would be downloaded
python download_bctc.py --dry-run

# Specific banks / years
python download_bctc.py --banks vcb tcb acb --years 2022 2023 2024

# Retry broken links after a normal pass
python download_bctc.py --retry-broken
```

### Report types

| Tag | Report name |
|---|---|
| `Consolidated` | Báo cáo tài chính hợp nhất năm YYYY (đã kiểm toán) |
| `Annual` | Báo cáo tài chính năm YYYY (đã kiểm toán) |

When cafef has no record, the script falls back to **vietstock.vn** (`POST /data/getdocument`) automatically.

### Output

```
financial_reports/
├── ABB/
│   ├── ABB_BCTC_Consolidated_2018.pdf
│   ├── ABB_BCTC_Annual_2010.pdf
│   └── ...
├── ACB/
│   └── ...
└── download_log.json
```

`download_log.json` statuses:

| Status | Meaning |
|---|---|
| `ok` | Downloaded successfully |
| `exists` | Already on disk, skipped |
| `not_found` | No record on cafef or vietstock |
| `broken_link` | Record found but CDN file is gone |
| `api_error` | API request failed after retries |

---

## 2. Parse PDFs → Excel — `parse_bctc.py`

Uses a vision LLM to read scanned PDF pages and extract financial metrics.

```bash
# Full run (all banks, all years)
python parse_bctc.py --model claude-sonnet-4-5-20250929

# Specific banks / years
python parse_bctc.py --banks abb acb --years 2015 2016

# Use OpenAI-compatible proxy (default)
python parse_bctc.py --provider openai --model gpt-5.4-mini

# Use Anthropic API directly
ANTHROPIC_API_KEY=sk-ant-... python parse_bctc.py --provider anthropic --model claude-3-5-sonnet-20241022
```

### Extracted fields

| Column | Description |
|---|---|
| `TỔNG CỘNG TÀI SẢN` | Total assets |
| `XIII. Lợi nhuận sau thuế (XI-XII)` | Net profit after tax |
| `III. Tiền gửi của khách hàng` | Customer deposits |
| `1. Cho vay và cho thuê tài chính khách hàng` | Customer loans & financial leases |
| `X. Chi phí dự phòng rủi ro tín dụng` | Credit loss provision expense |
| `Chi phí hoạt động` | Operating expenses |
| `Tổng thu nhập hoạt động` | Total operating income |
| `TỔNG NỢ PHẢI TRẢ` | Total liabilities |
| `TỔNG NỢ PHẢI TRẢ VÀ VỐN CHỦ SỞ HỮU` | Total liabilities & equity |
| `Total Equity` | Shareholders' equity |

Output: `financial_reports/bctc_data.xlsx`

---

## 3. Fill missing cells — `fill_missing.py`

Targeted re-extraction for rows that are still empty in the Excel.

```bash
# Fill all missing rows
python fill_missing.py --model claude-sonnet-4-5-20250929

# Fill specific MCPs only
python fill_missing.py --mcp VBB AGR --model claude-sonnet-4-5-20250929

# Higher concurrency for speed (may hit rate limits)
python fill_missing.py --concurrency 4 --model claude-sonnet-4-5-20250929
```

> Zeros are treated as empty and will be overwritten with extracted values.

### MCP name mapping

Some banks use different codes in the Excel vs. their folder name:

| Excel MCP | Folder |
|---|---|
| `AGRB` | `AGR/` |
| `PVCOMBANK` | `PCB/` |

---

## LLM proxy

All vision extraction routes through `https://llm.ngtlam.com/v1` (OpenAI-compatible). Available models include `claude-sonnet-4-5-20250929`, `gpt-5.4-mini`, `gpt-5.5`.

---

## Known gaps

- **BAB** 2010–2011 · **OCB** 2011 · **SGB** 2010 · **TPB** 2011 — no PDFs on cafef or vietstock
- **VBB** 2010–2015 — bank not yet listed / no public reports for those years
- **SCB** 2020–2024 — placed under State Bank special control in 2022, stopped publishing
- **PVCOMBANK / PCB** — use `--banks pcb` to download; appears as `PVCOMBANK` in the Excel

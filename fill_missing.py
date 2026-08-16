#!/usr/bin/env python3
"""
fill_missing.py — Find rows with NaN in bctc_data.xlsx, extract from PDFs,
and write the values back.

Usage:
    python fill_missing.py [--xlsx financial_reports/bctc_data.xlsx]
                           [--model claude-sonnet-4-5-20250929]
                           [--dpi 150] [--concurrency 2]
"""

import argparse
import asyncio
import base64
import io
import json
import os
import sys
from pathlib import Path

import httpx
import pandas as pd
from dotenv import load_dotenv
from pdf2image import convert_from_path

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────

API_BASE = os.getenv("LLM_API_BASE", "")
API_KEY  = os.getenv("LLM_API_KEY", "")
MODEL    = os.getenv("LLM_MODEL", "gpt-5.4-mini")

REPORTS_DIR  = Path("financial_reports")
DEFAULT_XLSX = REPORTS_DIR / "bctc_data.xlsx"

DPI         = 150
MAX_PAGES   = 14
CONCURRENCY = 2
RETRY_DELAY = 3.0
MAX_RETRIES = 3

# Map MCP prefix → actual folder name (when they differ)
FOLDER_MAP: dict[str, str] = {
    "AGRB": "AGR",
    "PVCOMBANK": "PCB",   # cafef symbol is pcb
}
# MCP prefixes with no PDFs available — skip silently
NO_PDF: set[str] = set()

FIELDS = [
    "TỔNG CỘNG TÀI SẢN",
    "XIII. Lợi nhuận sau thuế (XI-XII)",
    "III. Tiền gửi của khách hàng",
    "1. Cho vay và cho thuê tài chính khách hàng",
    "X. Chi phí dự phòng rủi ro tín dụng",
    "Chi phí hoạt động",
    "Tổng thu nhập hoạt động",
    "TỔNG NỢ PHẢI TRẢ",
    "TỔNG NỢ PHẢI TRẢ VÀ VỐN CHỦ SỞ HỮU",
    "Total Equity",
]

SYSTEM_PROMPT = """\
You are a financial data extraction assistant for Vietnamese bank annual reports.
Given an image of a page from a Vietnamese bank financial statement (BCTC), extract
the exact numeric values for the requested line items.

Rules:
- Return ONLY a JSON object with the field names as keys and integer values (no commas,
  no units). Use null if a field is not present on this page.
- Values are typically in millions of VND (triệu đồng) — extract the number as printed.
- For "TỔNG NỢ PHẢI TRẢ" return only the subtotal line, NOT "TỔNG NỢ PHẢI TRẢ VÀ VỐN CHỦ SỞ HỮU".
- If a value appears with parentheses e.g. (496,149) it is negative — return as negative int.
- Use the current-year column (leftmost value column), not the prior-year comparison column.
- "Total Equity" = Vốn chủ sở hữu (total equity / shareholders equity subtotal).
- "Tổng thu nhập hoạt động" may also appear as "Thu nhập hoạt động thuần" or
  "Tổng thu nhập hoạt động thuần" — treat them as equivalent.
- Do not include any explanation, only the JSON object.\
"""

USER_PROMPT = """\
Extract these line items from the financial statement page (null for any not found on this page):
{fields}

Return JSON only.\
"""


# ── Helpers ───────────────────────────────────────────────────────────────────

def color(t: str, c: str) -> str:
    return f"\033[{c}m{t}\033[0m" if sys.stdout.isatty() else t

ok_  = lambda t: color(t, "32")
err_ = lambda t: color(t, "31")
dim_ = lambda t: color(t, "2")
yel_ = lambda t: color(t, "33")


def find_pdf(symbol: str, year: int) -> Path | None:
    """Return best matching PDF for (symbol, year), or None."""
    folder = REPORTS_DIR / FOLDER_MAP.get(symbol.upper(), symbol.upper())
    if not folder.exists():
        return None
    # Prefer Consolidated > Annual
    for tag in ("Consolidated", "Annual"):
        pdf = folder / f"{folder.name}_BCTC_{tag}_{year}.pdf"
        if pdf.exists():
            return pdf
    # Fallback: any pdf containing the year
    candidates = sorted(folder.glob(f"*_{year}.pdf"))
    return candidates[0] if candidates else None


def page_to_b64(page) -> str:
    buf = io.BytesIO()
    page.save(buf, format="JPEG", quality=82)
    return base64.b64encode(buf.getvalue()).decode()


def parse_json(raw: str) -> dict:
    text = raw.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text.strip())


# ── Multimodal probe ──────────────────────────────────────────────────────────

# 1×1 white JPEG in base64 — minimal valid image for a vision probe
_PROBE_IMG = (
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8U"
    "HRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgN"
    "DRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIy"
    "MjIyMjL/wAARCAAEAAQDASIAAhEBAxEB/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAA"
    "AAAAAAAAAAAAAAAAAAAA/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/EABQRAQAAAAAAAAAAAAAAA"
    "AAAAP/aAAwDAQACEQMRAD8AJQAA/9k="
)


async def probe_vision(client: httpx.AsyncClient) -> None:
    """
    Send a minimal image to the configured model.
    Raises SystemExit if the model doesn't support vision input.
    """
    payload = {
        "model": MODEL,
        "max_tokens": 5,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "text", "text": "Reply OK"},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{_PROBE_IMG}"}},
            ],
        }],
    }
    try:
        r = await client.post(f"{API_BASE}/chat/completions", json=payload, timeout=30)
        data = r.json()
        # A 4xx with a message about images/vision means no multimodal support
        err_msg = (data.get("error") or {}).get("message", "")
        if r.status_code >= 400 and any(
            kw in err_msg.lower()
            for kw in ["does not support", "vision not", "no vision", "multimodal not", "not support image"]
        ):
            print(err_(f"✗ Model '{MODEL}' does not support vision input: {err_msg}"))
            raise SystemExit(1)
    except httpx.HTTPError as exc:
        print(err_(f"✗ Vision probe failed (network): {exc}"))
        raise SystemExit(1)


# ── LLM call ─────────────────────────────────────────────────────────────────

async def extract_page(
    client: httpx.AsyncClient,
    b64: str,
    fields: list[str],
) -> dict:
    payload = {
        "model": MODEL,
        "max_tokens": 512,
        "temperature": 0,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": [
                {"type": "text",
                 "text": USER_PROMPT.format(fields="\n".join(f"- {f}" for f in fields))},
                {"type": "image_url",
                 "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
            ]},
        ],
    }
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = await client.post(f"{API_BASE}/chat/completions", json=payload, timeout=180)
            r.raise_for_status()
            data = r.json()
            content = data.get("choices", [{}])[0].get("message", {}).get("content", "")
            if not content or not content.strip():
                # Empty response — model refused or returned nothing; skip this page
                return {}
            return parse_json(content)
        except (httpx.HTTPError, json.JSONDecodeError, KeyError) as exc:
            if attempt < MAX_RETRIES:
                await asyncio.sleep(RETRY_DELAY * attempt)
            else:
                print(err_(f"    ✗ LLM error: {exc}"))
    return {}


async def extract_missing_from_pdf(
    client: httpx.AsyncClient,
    pdf: Path,
    missing_fields: list[str],
    sem: asyncio.Semaphore,
) -> dict:
    async with sem:
        try:
            pages = convert_from_path(pdf, dpi=DPI, last_page=MAX_PAGES)
        except Exception as exc:
            print(err_(f"    ✗ PDF render: {exc}"))
            return {}

        remaining = list(missing_fields)
        accumulated: dict = {}

        for i, page in enumerate(pages):
            if not remaining:
                break
            b64 = page_to_b64(page)
            result = await extract_page(client, b64, remaining)
            for f, v in result.items():
                if v is not None and f not in accumulated:
                    accumulated[f] = v
            remaining = [f for f in remaining if accumulated.get(f) is None]
            found_now = [f for f in missing_fields if result.get(f) is not None]
            if found_now:
                print(dim_(f"    p{i+1}: +{len(found_now)} field(s)"))

        return accumulated


# ── Main ─────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--xlsx",        default=str(DEFAULT_XLSX))
    p.add_argument("--model",       default=MODEL)
    p.add_argument("--dpi",         type=int, default=DPI)
    p.add_argument("--concurrency", type=int, default=CONCURRENCY)
    p.add_argument("--mcp",         nargs="+", metavar="MCP",
                   help="Only fill these MCP codes e.g. --mcp VBB ACB")
    return p.parse_args()


async def run(args):
    global MODEL, DPI
    MODEL = args.model
    DPI   = args.dpi

    xlsx = Path(args.xlsx)
    df = pd.read_excel(xlsx, header=0, index_col=0)
    df.index.name = "row_key"

    # Normalise column names (strip ASCII + non-breaking spaces)
    df.columns = [str(c).replace('\xa0', ' ').strip() for c in df.columns]

    # Only process rows that have at least one NaN in extractable fields
    extractable = [f for f in FIELDS if f in df.columns]

    # Treat 0 as empty
    df[extractable] = df[extractable].replace(0, pd.NA)

    missing_mask = df[extractable].isnull().any(axis=1)
    target_rows = df[missing_mask].copy()

    # Filter to specific MCPs if requested
    if args.mcp:
        wanted = [m.upper() for m in args.mcp]
        target_rows = target_rows[target_rows["MCP"].str.upper().isin(wanted)]

    print(f"Rows with missing data : {len(target_rows)}")
    print(f"Model                  : {MODEL}\n")

    sem = asyncio.Semaphore(args.concurrency)
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}

    async with httpx.AsyncClient(headers=headers, follow_redirects=True) as probe_client:
        print(f"Probing vision support for model '{MODEL}' … ", end="", flush=True)
        await probe_vision(probe_client)
        print(ok_("ok"))

    async def process_row(client, idx, row):
        symbol = str(row.get("MCP", "")).strip().upper()
        year   = int(row.get("Năm", 0))

        if not symbol or not year:
            return idx, {}

        if symbol in NO_PDF:
            print(dim_(f"  - {symbol} {year}: no PDFs available, skipping"))
            return idx, {}

        missing_fields = [f for f in extractable if pd.isna(row.get(f))]
        if not missing_fields:
            return idx, {}

        pdf = find_pdf(symbol, year)
        if pdf is None:
            print(yel_(f"  ! {symbol} {year}: no PDF found"))
            return idx, {}

        print(f"  ↓ {symbol} {year}  ({len(missing_fields)} missing)  [{pdf.name}]")
        extracted = await extract_missing_from_pdf(client, pdf, missing_fields, sem)
        return idx, extracted

    async with httpx.AsyncClient(headers=headers, follow_redirects=True) as client:
        tasks = [
            process_row(client, idx, row)
            for idx, row in target_rows.iterrows()
        ]
        results = await asyncio.gather(*tasks)

    # Write extracted values back into df and save
    filled = 0
    for idx, extracted in results:
        for field, value in extracted.items():
            if field in df.columns:
                current = df.at[idx, field]
                if pd.isna(current) or current == 0:
                    df.at[idx, field] = value
                    filled += 1

    df.to_excel(xlsx, index=True)
    print(f"\n{ok_(f'Filled {filled} cells')} → {xlsx}")


if __name__ == "__main__":
    args = parse_args()
    asyncio.run(run(args))

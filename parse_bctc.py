#!/usr/bin/env python3
"""
parse_bctc.py — Extract key financial metrics from Vietnamese bank BCTC PDFs
using a vision LLM and write to Excel.

Providers:
  --provider openai     OpenAI-compatible endpoint (default: llm.ngtlam.com)
  --provider anthropic  Anthropic API directly (claude-3-5-sonnet-20241022)

Usage:
    python parse_bctc.py [--dir financial_reports] [--out bctc_data.xlsx]
                         [--banks abb acb ...] [--years 2010 2011 ...]
                         [--provider anthropic] [--model claude-3-5-sonnet-20241022]
                         [--dpi 150] [--concurrency 3]
"""

import argparse
import asyncio
import base64
import io
import json
import os
import sys
from pathlib import Path

import anthropic
import httpx
import pandas as pd
from dotenv import load_dotenv
from pdf2image import convert_from_path

load_dotenv()

# ── Configuration ──────────────────────────────────────────────────────────────

# OpenAI-compatible proxy
OPENAI_API_BASE = os.getenv("LLM_API_BASE", "")
OPENAI_API_KEY  = os.getenv("LLM_API_KEY", "")
OPENAI_MODEL    = os.getenv("LLM_MODEL", "gpt-5.4-mini")

# Anthropic direct
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL   = "claude-3-5-sonnet-20241022"

# Active provider/model — overridden by CLI args
PROVIDER = "openai"
API_BASE = OPENAI_API_BASE
API_KEY  = OPENAI_API_KEY
MODEL    = OPENAI_MODEL
REPORTS_DIR = Path("financial_reports")
OUTPUT_FILE = Path("bctc_data.xlsx")

DPI          = 150    # render resolution; 150 is fast, raise to 200 for blurry scans
MAX_PAGES    = 12     # look at at most this many pages per PDF (saves tokens)
CONCURRENCY  = 3      # parallel PDFs
RETRY_DELAY  = 3.0
MAX_RETRIES  = 3

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
]

SYSTEM_PROMPT = """\
You are a financial data extraction assistant for Vietnamese bank annual reports.
Given an image of a page from a Vietnamese bank financial statement (BCTC), extract
the exact numeric values for the requested line items.

Rules:
- Return ONLY a JSON object with the field names as keys and integer values (no commas,
  no units). Use null if a field is not present on this page.
- Values are typically in millions of VND (triệu đồng) — extract the number as printed.
- For "TỔNG NỢ PHẢI TRẢ" return only the subtotal (not "TỔNG NỢ PHẢI TRẢ VÀ VỐN CHỦ SỞ HỮU").
- If a value appears with parentheses (e.g. (496,149)) it is negative — return as negative int.
- Use the current-year column (leftmost value column), not the prior-year comparison column.
- Do not include any explanation, only the JSON object.\
"""

USER_PROMPT = """\
Extract these line items from the financial statement page (return null for any not found):
{fields}

Return JSON only.\
"""


# ── Helpers ────────────────────────────────────────────────────────────────────

def color(text: str, code: str) -> str:
    return f"\033[{code}m{text}\033[0m" if sys.stdout.isatty() else text

def ok_(t: str) -> str:  return color(t, "32")
def err_(t: str) -> str: return color(t, "31")
def dim_(t: str) -> str: return color(t, "2")


def page_to_b64(page) -> str:
    """Convert a PIL image (PDF page) to base64 JPEG."""
    buf = io.BytesIO()
    page.save(buf, format="JPEG", quality=82)
    return base64.b64encode(buf.getvalue()).decode()


def merge_results(pages_data: list[dict]) -> dict:
    """
    Merge per-page extraction results: take the first non-null value seen for
    each field across all pages.
    """
    merged: dict = {f: None for f in FIELDS}
    for page_result in pages_data:
        for field in FIELDS:
            if merged[field] is None and page_result.get(field) is not None:
                merged[field] = page_result[field]
        if all(v is not None for v in merged.values()):
            break   # all fields found, no need to look further
    return merged


# ── LLM extraction ─────────────────────────────────────────────────────────────

def _parse_llm_json(raw: str) -> dict:
    """Strip markdown fences and parse JSON from LLM response."""
    text = raw.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text.strip())


# ── Multimodal probe ──────────────────────────────────────────────────────────

_PROBE_IMG = (
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8U"
    "HRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgN"
    "DRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIy"
    "MjIyMjL/wAARCAAEAAQDASIAAhEBAxEB/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAA"
    "AAAAAAAAAAAAAAAAAAAA/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/EABQRAQAAAAAAAAAAAAAAA"
    "AAAAP/aAAwDAQACEQMRAD8AJQAA/9k="
)


async def probe_vision_openai(client: httpx.AsyncClient) -> None:
    """Fail fast if the active OpenAI-compatible model doesn't support vision."""
    payload = {
        "model": MODEL,
        "max_tokens": 5,
        "messages": [{"role": "user", "content": [
            {"type": "text", "text": "Reply OK"},
            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{_PROBE_IMG}"}},
        ]}],
    }
    try:
        r = await client.post(f"{API_BASE}/chat/completions", json=payload, timeout=30)
        data = r.json()
        err_msg = (data.get("error") or {}).get("message", "")
        if r.status_code >= 400 and any(
            kw in err_msg.lower()
            for kw in ["does not support", "vision not", "no vision", "multimodal not", "not support image"]
        ):
            print(err_(f"✗ Model '{MODEL}' does not support vision: {err_msg}"))
            raise SystemExit(1)
    except httpx.HTTPError as exc:
        print(err_(f"✗ Vision probe failed: {exc}"))
        raise SystemExit(1)


async def _extract_page_openai(
    client: httpx.AsyncClient,
    b64_image: str,
    fields: list[str],
) -> dict:
    user_content = [
        {"type": "text", "text": USER_PROMPT.format(fields="\n".join(f"- {f}" for f in fields))},
        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_image}"}},
    ]
    payload = {
        "model": MODEL,
        "max_tokens": 512,
        "temperature": 0,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": user_content},
        ],
    }
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = await client.post(f"{API_BASE}/chat/completions", json=payload, timeout=180)
            resp.raise_for_status()
            content = resp.json().get("choices", [{}])[0].get("message", {}).get("content", "")
            if not content or not content.strip():
                return {}
            return _parse_llm_json(content)
        except (httpx.HTTPError, json.JSONDecodeError, KeyError) as exc:
            if attempt < MAX_RETRIES:
                await asyncio.sleep(RETRY_DELAY * attempt)
            else:
                print(err_(f"    ✗ OpenAI error: {exc}"))
    return {}


async def _extract_page_anthropic(
    aclient: anthropic.AsyncAnthropic,
    b64_image: str,
    fields: list[str],
) -> dict:
    user_content = [
        {
            "type": "image",
            "source": {"type": "base64", "media_type": "image/jpeg", "data": b64_image},
        },
        {"type": "text", "text": USER_PROMPT.format(fields="\n".join(f"- {f}" for f in fields))},
    ]
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            msg = await aclient.messages.create(
                model=MODEL,
                max_tokens=512,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_content}],
            )
            return _parse_llm_json(msg.content[0].text)
        except (anthropic.APIError, json.JSONDecodeError, IndexError) as exc:
            if attempt < MAX_RETRIES:
                await asyncio.sleep(RETRY_DELAY * attempt)
            else:
                print(err_(f"    ✗ Anthropic error: {exc}"))
    return {}


async def extract_page(
    client,          # httpx.AsyncClient (openai) or anthropic.AsyncAnthropic
    b64_image: str,
    fields: list[str],
) -> dict:
    if PROVIDER == "anthropic":
        return await _extract_page_anthropic(client, b64_image, fields)
    return await _extract_page_openai(client, b64_image, fields)


async def extract_pdf(
    client: httpx.AsyncClient,
    pdf_path: Path,
    sem: asyncio.Semaphore,
) -> dict:
    """Render pages, send to LLM one at a time, merge results."""
    async with sem:
        print(f"  ↓ {pdf_path.name}")
        try:
            pages = convert_from_path(pdf_path, dpi=DPI, last_page=MAX_PAGES)
        except Exception as exc:
            print(err_(f"    ✗ PDF render failed: {exc}"))
            return {f: None for f in FIELDS}

        # Identify which pages are likely financial-statement pages (skip cover pages)
        # by checking if any keyword appears — we do a quick text probe via the model
        # on the first few pages and merge results.
        remaining_fields = list(FIELDS)
        pages_data: list[dict] = []

        for i, page in enumerate(pages):
            if not remaining_fields:
                break
            b64 = page_to_b64(page)
            result = await extract_page(client, b64, remaining_fields)
            pages_data.append(result)
            # Remove fields we've already found
            remaining_fields = [f for f in remaining_fields if not result.get(f)]
            found = [f for f in FIELDS if result.get(f) is not None]
            if found:
                print(dim_(f"    p{i+1}: found {len(found)} field(s)"))

        return merge_results(pages_data)


# ── Main ───────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(description="Parse BCTC PDFs → Excel using vision LLM")
    p.add_argument("--dir", default=str(REPORTS_DIR))
    p.add_argument("--out", default=str(OUTPUT_FILE))
    p.add_argument("--banks",    nargs="+", metavar="SYMBOL")
    p.add_argument("--years",    nargs="+", type=int, metavar="YEAR")
    p.add_argument("--provider", default="openai", choices=["openai", "anthropic"],
                   help="openai = CLIProxyAPI (default), anthropic = Anthropic API directly")
    p.add_argument("--model",    default=None,
                   help="Override model (default: gpt-5.4-mini for openai, claude-3-5-sonnet-20241022 for anthropic)")
    p.add_argument("--dpi",      type=int, default=DPI)
    p.add_argument("--concurrency", type=int, default=CONCURRENCY)
    return p.parse_args()


async def run(args) -> None:
    global MODEL, DPI, PROVIDER, API_BASE, API_KEY
    PROVIDER = args.provider
    DPI      = args.dpi

    if PROVIDER == "anthropic":
        MODEL = args.model or ANTHROPIC_MODEL
        api_key = ANTHROPIC_API_KEY or os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            print(err_("Set ANTHROPIC_API_KEY env var to use --provider anthropic"))
            return
    else:
        MODEL    = args.model or OPENAI_MODEL
        API_BASE = OPENAI_API_BASE
        API_KEY  = OPENAI_API_KEY
        # Enforce multimodal support before processing any PDFs
        oai_headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
        async with httpx.AsyncClient(headers=oai_headers, follow_redirects=True) as probe:
            print(f"Probing vision support for model '{MODEL}' … ", end="", flush=True)
            await probe_vision_openai(probe)
            print(ok_("ok"))

    reports_dir = Path(args.dir)
    pdfs = sorted(reports_dir.glob("*/*_Consolidated_*.pdf"))

    # Also include Annual type
    pdfs += sorted(reports_dir.glob("*/*_Annual_*.pdf"))
    pdfs = sorted(set(pdfs))

    # Filter
    filtered = []
    for pdf in pdfs:
        parts = pdf.stem.split("_")
        symbol = parts[0].lower()
        try:
            year = int(parts[-1])
        except ValueError:
            continue
        if args.banks and symbol not in [b.lower() for b in args.banks]:
            continue
        if args.years and year not in args.years:
            continue
        filtered.append((pdf, symbol, year))

    if not filtered:
        print("No PDFs matched.")
        return

    print(f"Extracting {len(filtered)} PDF(s) with provider={PROVIDER}, model={MODEL}, dpi={DPI}, concurrency={args.concurrency}\n")

    sem = asyncio.Semaphore(args.concurrency)
    rows = []

    async def extract_with_meta(client, pdf: Path, symbol: str, year: int) -> tuple:
        result = await extract_pdf(client, pdf, sem)
        return symbol, year, result

    if PROVIDER == "anthropic":
        llm_client = anthropic.AsyncAnthropic(api_key=ANTHROPIC_API_KEY or os.environ["ANTHROPIC_API_KEY"])
        tasks = [extract_with_meta(llm_client, pdf, symbol, year) for pdf, symbol, year in filtered]
        results = await asyncio.gather(*tasks)
        await llm_client.close()
    else:
        headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
        async with httpx.AsyncClient(headers=headers, follow_redirects=True) as client:
            tasks = [extract_with_meta(client, pdf, symbol, year) for pdf, symbol, year in filtered]
            results = await asyncio.gather(*tasks)

    for symbol, year, result in results:
        row = {"MCP": f"{symbol.upper()}{year}", "Năm": year, **result}
        rows.append(row)
        print(ok_(f"  ✓ {symbol.upper()} {year}"))

    if not rows:
        print("No results.")
        return

    cols = ["MCP", "Năm"] + FIELDS
    df = pd.DataFrame(rows, columns=cols).sort_values(["MCP"]).set_index("MCP")

    out = Path(args.out)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="BCTC")

    print(f"\nWrote {len(rows)} rows → {out}")


if __name__ == "__main__":
    args = parse_args()
    asyncio.run(run(args))

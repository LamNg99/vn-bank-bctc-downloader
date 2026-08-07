#!/usr/bin/env python3
"""
Download annual audited financial reports for Vietnamese banks from cafef.vn.

Three report types are fetched per bank/year (whichever exist):
  1. "Báo cáo tài chính hợp nhất năm {year} (đã kiểm toán)"   → *_Consolidated_*
  2. "Báo cáo tài chính công ty mẹ năm {year} (đã kiểm toán)" → *_CongTyMe_*
  3. "Báo cáo tài chính năm {year} (đã kiểm toán)"             → *_Annual_*
     (used by some banks — e.g. LPB, OCB, PGB, TPB — that don't split reports)

Usage:
    python download_bctc.py [--dry-run] [--banks abb acb ...] [--years 2020 2021 ...]
"""

import asyncio
import argparse
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import httpx

# ── Configuration ──────────────────────────────────────────────────────────────

BANKS = [
    "abb", "acb", "bab", "bid", "bvb", "ctg", "eib", "hdb", "klb", "lpb",
    "mbb", "msb", "nab", "nvb", "ocb", "pgb", "scb", "sgb", "shb", "ssb",
    "stb", "tcb", "tpb", "vcb", "vib", "vpb",
]

YEARS = list(range(2010, 2025))   # 2010 … 2024 inclusive

API_URL  = "https://cafef.vn/du-lieu/Ajax/PageNew/FileBCTC.ashx"
OUT_DIR  = Path("financial_reports")
LOG_FILE = OUT_DIR / "download_log.json"

# Polite concurrency — don't hammer the server
MAX_CONCURRENT = 4
RETRY_DELAY    = 2.0   # seconds between retries
MAX_RETRIES    = 3

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    ),
    "Referer": "https://cafef.vn/",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8",
}


# ── Report target definitions ──────────────────────────────────────────────────

@dataclass
class ReportTarget:
    """Describes one type of annual audited report to look for."""
    label: str           # human description for logs
    file_tag: str        # inserted into output filename, e.g. "HopNhat"
    exact_name: str      # filled in at runtime with .format(year=year)
    # Keywords that must ALL appear in Name for the loose fallback match:
    loose_must: list[str] = field(default_factory=list)
    # Keywords that must NOT appear in Name during the loose match:
    loose_must_not: list[str] = field(default_factory=list)


REPORT_TARGETS: list[ReportTarget] = [
    ReportTarget(
        label="Hợp nhất (consolidated)",
        file_tag="Consolidated",
        exact_name="Báo cáo tài chính hợp nhất năm {year} (đã kiểm toán)",
        loose_must=["hợp nhất", "kiểm toán"],
        loose_must_not=["công ty mẹ"],
    ),
    # ReportTarget(
    #     label="Công ty mẹ (parent company)",
    #     file_tag="CongTyMe",
    #     exact_name="Báo cáo tài chính công ty mẹ năm {year} (đã kiểm toán)",
    #     loose_must=["công ty mẹ", "kiểm toán"],
    #     loose_must_not=["hợp nhất", "quý"],
    # ),
    ReportTarget(
        label="Báo cáo tài chính năm (plain annual — LPB, OCB, PGB, TPB etc.)",
        file_tag="Annual",
        exact_name="Báo cáo tài chính năm {year} (đã kiểm toán)",
        # No "hợp nhất", no "công ty mẹ" — plain un-split annual report
        loose_must=["kiểm toán"],
        loose_must_not=["hợp nhất", "công ty mẹ", "quý"],
    ),
]


# ── Helpers ────────────────────────────────────────────────────────────────────

def find_report(
    data: list[dict], year: int, target: ReportTarget
) -> Optional[dict]:
    """
    Try exact name match first; fall back to keyword match on annual (Quarter==5)
    entries that satisfy loose_must / loose_must_not.
    """
    exact = target.exact_name.format(year=year)
    for item in data:
        if item.get("Name", "").strip() == exact:
            return item

    # Loose fallback — must be annual (Quarter == 5) and contain the year
    for item in data:
        name = item.get("Name", "")
        if item.get("Quarter") != 5:
            continue
        if str(year) not in name:
            continue
        if not all(kw in name for kw in target.loose_must):
            continue
        if any(kw in name for kw in target.loose_must_not):
            continue
        return item

    return None


def file_extension(url: str) -> str:
    lower = url.lower().split("?")[0]
    for ext in (".pdf", ".xlsx", ".xls", ".zip", ".rar"):
        if lower.endswith(ext):
            return ext
    return ".pdf"   # safest default for BCTC


def color(text: str, code: str) -> str:
    return f"\033[{code}m{text}\033[0m" if sys.stdout.isatty() else text

ok_  = lambda t: color(t, "32")   # green
err_ = lambda t: color(t, "31")   # red
dim_ = lambda t: color(t, "2")    # dim


# ── CDN fallback hosts (tried in order when the primary URL 404s) ──────────────

CDN_FALLBACKS: list[tuple[str, str]] = [
    # Try every known CDN host variant when the primary URL 404s
    ("cafefnew.mediacdn.vn", "cafef1.mediacdn.vn"),   # most common fix
    ("cafefnew.mediacdn.vn", "cafef.mediacdn.vn"),
    ("cafef.mediacdn.vn",    "cafef1.mediacdn.vn"),
    ("cafef.mediacdn.vn",    "cafefnew.mediacdn.vn"),
    ("cafef1.mediacdn.vn",   "cafefnew.mediacdn.vn"),
    ("cafef1.mediacdn.vn",   "cafef.mediacdn.vn"),
    ("cafefnew.mediacdn.vn", "static.cafef.vn"),
    ("cafef.mediacdn.vn",    "static.cafef.vn"),
    ("cafef1.mediacdn.vn",   "static.cafef.vn"),
]


# ── Core async functions ───────────────────────────────────────────────────────

async def fetch_report_list(
    client: httpx.AsyncClient, symbol: str, year: int
) -> Optional[list[dict]]:
    params = {"Symbol": symbol.upper(), "Type": "1", "Year": str(year)}
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = await client.get(API_URL, params=params, timeout=30)
            resp.raise_for_status()
            result = resp.json()
            if result.get("Success"):
                return result.get("Data") or []
            return []
        except (httpx.HTTPError, json.JSONDecodeError) as exc:
            if attempt < MAX_RETRIES:
                await asyncio.sleep(RETRY_DELAY * attempt)
            else:
                print(err_(f"  ✗ API {symbol.upper()} {year}: {exc}"))
    return None


async def resolve_url(client: httpx.AsyncClient, url: str) -> Optional[str]:
    """
    HEAD-check the URL; if it returns 404 try known CDN host alternatives.
    Returns the first working URL, or None if all candidates fail.
    """
    candidates: list[str] = [url]
    for old_host, new_host in CDN_FALLBACKS:
        if old_host in url:
            candidates.append(url.replace(old_host, new_host, 1))

    # deduplicate while preserving order
    seen: set[str] = set()
    unique: list[str] = []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            unique.append(c)

    for candidate in unique:
        try:
            r = await client.head(candidate, timeout=15)
            if r.status_code < 400:
                if candidate != url:
                    print(color(f"    ↪ resolved via alt CDN: {candidate}", "33"))
                return candidate
            if r.status_code == 404 and candidate == url:
                continue   # try fallbacks
        except Exception:
            continue

    return None   # all candidates returned 4xx/5xx or errored


async def download_file(
    client: httpx.AsyncClient, url: str, dest: Path
) -> tuple[bool, str]:
    """
    Download url → dest.
    Returns (success, resolved_url_or_error_message).
    """
    resolved = await resolve_url(client, url)
    if resolved is None:
        msg = f"404 on all CDN candidates for {url}"
        print(err_(f"  ✗ Broken link: {url}"))
        return False, msg

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            async with client.stream("GET", resolved, timeout=120) as resp:
                resp.raise_for_status()
                with open(dest, "wb") as fh:
                    async for chunk in resp.aiter_bytes(8192):
                        fh.write(chunk)
            return True, resolved
        except Exception as exc:
            if attempt < MAX_RETRIES:
                await asyncio.sleep(RETRY_DELAY * attempt)
            else:
                print(err_(f"  ✗ Download failed ({resolved}): {exc}"))
                return False, str(exc)
    return False, "max retries exceeded"


async def process(
    client: httpx.AsyncClient,
    symbol: str,
    year: int,
    sem: asyncio.Semaphore,
    log: list[dict],
    dry_run: bool,
) -> None:
    """Fetch the report list once, then try every target type."""
    async with sem:
        data = await fetch_report_list(client, symbol, year)

        if data is None:
            log.append({"symbol": symbol, "year": year, "type": "all", "status": "api_error"})
            return

        any_found = False

        for target in REPORT_TARGETS:
            report = find_report(data, year, target)
            entry: dict = {
                "symbol": symbol,
                "year": year,
                "type": target.file_tag,
            }

            if not report:
                entry["status"] = "not_found"
                log.append(entry)
                continue

            any_found = True
            link: str = report.get("Link", "").strip()
            if not link:
                entry["status"] = "no_link"
                log.append(entry)
                continue

            ext  = file_extension(link)
            name = f"{symbol.upper()}_BCTC_{target.file_tag}_{year}{ext}"
            dest = OUT_DIR / symbol.upper() / name
            entry["file"] = str(dest)
            entry["url"]  = link
            entry["name"] = report.get("Name", "")

            if dest.exists():
                print(dim_(f"  ✓ {symbol.upper()} {year} [{target.file_tag}]: already exists, skipping"))
                entry["status"] = "exists"
                log.append(entry)
                continue

            if dry_run:
                print(f"  [dry-run] {symbol.upper()} {year} [{target.file_tag}]: → {dest}")
                entry["status"] = "dry_run"
                log.append(entry)
                continue

            print(f"  ↓ {symbol.upper()} {year} [{target.file_tag}]: {report['Name']}")
            success, detail = await download_file(client, link, dest)
            if success:
                entry["status"] = "ok"
                entry["resolved_url"] = detail   # may differ from original link
            else:
                entry["status"] = "broken_link" if "404" in detail else "download_error"
                entry["error"] = detail
            log.append(entry)

            await asyncio.sleep(0.3)   # polite gap between downloads

        if not any_found:
            print(dim_(f"  – {symbol.upper()} {year}: no matching reports found"))


# ── Broken-link recovery ───────────────────────────────────────────────────────

# Exchange type codes tried when recovering broken links
EXCHANGE_TYPES = [1, 2, 3]   # 1=HOSE, 2=HNX, 3=UPCOM


async def recover_broken(
    client: httpx.AsyncClient,
    entry: dict,
    sem: asyncio.Semaphore,
) -> dict:
    """
    For one broken-link log entry, try fetching the report list with every
    exchange Type and look for the same report Name.  Returns an updated entry.
    """
    symbol    = entry["symbol"]
    year      = entry["year"]
    file_tag  = entry["type"]
    want_name = entry.get("name", "")
    dest      = Path(entry["file"])

    # Find the matching ReportTarget so we can call find_report
    target = next((t for t in REPORT_TARGETS if t.file_tag == file_tag), None)

    async with sem:
        for type_code in EXCHANGE_TYPES:
            params = {"Symbol": symbol.upper(), "Type": str(type_code), "Year": str(year)}
            try:
                resp = await client.get(API_URL, params=params, timeout=30)
                resp.raise_for_status()
                result = resp.json()
                if not result.get("Success"):
                    continue
                data: list[dict] = result.get("Data") or []
            except Exception:
                continue

            # Try exact name match first, then target-based search
            candidates = [d for d in data if d.get("Name", "").strip() == want_name]
            if not candidates and target:
                found = find_report(data, year, target)
                if found:
                    candidates = [found]

            for candidate in candidates:
                link = candidate.get("Link", "").strip()
                if not link:
                    continue
                resolved = await resolve_url(client, link)
                if resolved is None:
                    continue

                # Found a working URL — download it
                print(f"  ↺ {symbol.upper()} {year} [{file_tag}] recovered via Type={type_code}")
                success, detail = await download_file(client, link, dest)
                if success:
                    new_entry = dict(entry)
                    new_entry["status"]       = "ok"
                    new_entry["url"]          = link
                    new_entry["resolved_url"] = detail
                    new_entry["recovered_via_type"] = type_code
                    return new_entry

        # All exchange types exhausted — still broken
        print(err_(f"  ✗ {symbol.upper()} {year} [{file_tag}]: permanently broken"))
        return entry   # unchanged


async def run_recovery(log_path: Path) -> None:
    """Read the log, retry every broken_link entry, write updated log."""
    if not log_path.exists():
        print(err_(f"Log not found: {log_path}"))
        return

    with open(log_path, encoding="utf-8") as fh:
        log: list[dict] = json.load(fh)

    broken = [e for e in log if e.get("status") == "broken_link"]
    if not broken:
        print("No broken links in log — nothing to recover.")
        return

    print(f"Recovering {len(broken)} broken links …\n")
    sem = asyncio.Semaphore(MAX_CONCURRENT)

    async with httpx.AsyncClient(headers=HEADERS, follow_redirects=True) as client:
        tasks = [recover_broken(client, e, sem) for e in broken]
        results = await asyncio.gather(*tasks)

    # Merge recovered entries back into the log (match by symbol+year+type)
    index: dict[tuple, int] = {
        (e["symbol"], e["year"], e["type"]): i
        for i, e in enumerate(log)
        if e.get("status") == "broken_link"
    }
    for updated in results:
        key = (updated["symbol"], updated["year"], updated["type"])
        if key in index:
            log[index[key]] = updated

    with open(log_path, "w", encoding="utf-8") as fh:
        json.dump(log, fh, ensure_ascii=False, indent=2)

    recovered  = sum(1 for r in results if r.get("status") == "ok")
    still_bad  = len(results) - recovered
    print()
    print("─" * 55)
    print(ok_(f"  Recovered   : {recovered}"))
    if still_bad:
        print(err_(f"  Still broken: {still_bad}  (permanently unavailable on cafef CDN)"))
    print(f"  Log updated : {log_path}")
    print("─" * 55)

    perm_broken = [r for r in results if r.get("status") == "broken_link"]
    if perm_broken:
        print()
        print(color("  Permanently broken (source file missing from cafef):", "31"))
        for r in perm_broken:
            print(color(f"    {r['symbol'].upper()} {r['year']} [{r['type']}]  {r.get('url', '')}", "31"))


# ── Main ───────────────────────────────────────────────────────────────────────

async def run(banks: list[str], years: list[int], dry_run: bool) -> None:
    OUT_DIR.mkdir(exist_ok=True)
    sem  = asyncio.Semaphore(MAX_CONCURRENT)
    log: list[dict] = []

    target_labels = " | ".join(f"[{t.file_tag}] {t.exact_name.format(year='YYYY')}" for t in REPORT_TARGETS)
    total = len(banks) * len(years)
    print(f"Banks  : {', '.join(b.upper() for b in banks)}")
    print(f"Years  : {years[0]}–{years[-1]}")
    print(f"Tasks  : {total} bank-year pairs  (concurrency={MAX_CONCURRENT})")
    print(f"Types  : {target_labels}")
    print(f"Output : {OUT_DIR.resolve()}")
    if dry_run:
        print(color("  *** DRY RUN — no files will be downloaded ***", "33"))
    print()

    async with httpx.AsyncClient(headers=HEADERS, follow_redirects=True) as client:
        tasks = [
            process(client, symbol, year, sem, log, dry_run)
            for symbol in banks
            for year in years
        ]
        await asyncio.gather(*tasks)

    # Persist full log
    with open(LOG_FILE, "w", encoding="utf-8") as fh:
        json.dump(log, fh, ensure_ascii=False, indent=2)

    # Summary
    counts: dict[str, int] = {}
    for r in log:
        counts[r["status"]] = counts.get(r["status"], 0) + 1

    print()
    print("─" * 55)
    print(ok_(f"  Downloaded  : {counts.get('ok', 0)}"))
    print(dim_(f"  Existed     : {counts.get('exists', 0)}"))
    print(dim_(f"  Not found   : {counts.get('not_found', 0)}"))
    if counts.get("broken_link"):
        print(err_(f"  Broken links: {counts['broken_link']}  (see log for URLs)"))
    if counts.get("api_error") or counts.get("download_error"):
        print(err_(f"  Other errors: {counts.get('api_error', 0) + counts.get('download_error', 0)}"))
    print(f"  Log         : {LOG_FILE}")
    print("─" * 55)

    # Print broken links to stdout for easy copy-paste / investigation
    broken = [r for r in log if r.get("status") == "broken_link"]
    if broken:
        print()
        print(color("  Broken links (all CDN alternatives failed):", "33"))
        for r in broken:
            print(color(f"    {r['symbol'].upper()} {r['year']} [{r['type']}]  {r.get('url', '')}", "33"))


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Download Vietnamese bank BCTC from cafef.vn")
    p.add_argument(
        "--banks", nargs="+", metavar="SYMBOL",
        default=BANKS,
        help="Bank symbols to download (default: all 26)",
    )
    p.add_argument(
        "--years", nargs="+", type=int, metavar="YEAR",
        default=YEARS,
        help="Years to download (default: 2010–2025)",
    )
    p.add_argument(
        "--dry-run", action="store_true",
        help="Print what would be downloaded without actually downloading",
    )
    p.add_argument(
        "--retry-broken", action="store_true",
        help=(
            "Re-attempt all broken_link entries in the log by trying every "
            "exchange Type (1/2/3). Run after a normal download pass."
        ),
    )
    return p.parse_args()


if __name__ == "__main__":
    args = parse_args()
    if args.retry_broken:
        asyncio.run(run_recovery(LOG_FILE))
    else:
        banks = [b.lower() for b in args.banks]
        asyncio.run(run(banks, sorted(args.years), args.dry_run))

from __future__ import annotations

import asyncio
import re
from urllib.parse import quote

import httpx
from playwright.async_api import async_playwright

from drive2_vin_audit import THREAD, MAX_PAGE, VIN_RE, FRAME_RE, clean_code, decode_ravenol, save_excel

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.7",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}


def html_to_text(html: str) -> str:
    html = re.sub(r"(?is)<script.*?</script>|<style.*?</style>", " ", html)
    text = re.sub(r"(?s)<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", text)


async def fetch_one(client: httpx.AsyncClient, sem: asyncio.Semaphore, n: int):
    url = THREAD if n == 1 else f"{THREAD}?page={n}#comments"
    async with sem:
        try:
            r = await client.get(url, timeout=15, follow_redirects=True)
            text = html_to_text(r.text)
            status = str(r.status_code)
        except Exception as exc:
            return n, "", f"error: {type(exc).__name__}"
    return n, text, status


async def collect_codes_fast():
    sem = asyncio.Semaphore(16)
    async with httpx.AsyncClient(headers=HEADERS, http2=True) as client:
        results = await asyncio.gather(*(fetch_one(client, sem, n) for n in range(1, MAX_PAGE + 1)))
    rows = []
    seen_mentions = set()
    for n, text, status in sorted(results):
        found = 0
        for regex, kind in ((VIN_RE, "VIN"), (FRAME_RE, "Frame")):
            for match in regex.finditer(text):
                code = clean_code(match.group(0))
                if code and (n, code) not in seen_mentions:
                    seen_mentions.add((n, code))
                    snippet = text[max(0, match.start()-120):match.end()+120]
                    rows.append({"page": n, "code": code, "kind": kind, "snippet": snippet[:500], "drive2_status": status})
                    found += 1
        print(f"Drive2 page {n}: {found} codes, status={status}", flush=True)
    return rows


async def main():
    mentions = await collect_codes_fast()
    unique_codes = sorted({r["code"] for r in mentions})
    decoded = {}
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(locale="ru-RU", user_agent=HEADERS["User-Agent"])
        pages = [await context.new_page() for _ in range(4)]
        sem = asyncio.Semaphore(4)

        async def decode_one(idx: int, code: str):
            async with sem:
                page = pages[idx % len(pages)]
                result = await decode_ravenol(page, code)
                print(f"Ravenol {idx+1}/{len(unique_codes)} {code}: {result['manufacturer']} ({result['ravenol_status']})", flush=True)
                return code, result

        pairs = await asyncio.gather(*(decode_one(i, code) for i, code in enumerate(unique_codes)))
        decoded = dict(pairs)
        await browser.close()
    save_excel(mentions, decoded)


if __name__ == "__main__":
    asyncio.run(main())

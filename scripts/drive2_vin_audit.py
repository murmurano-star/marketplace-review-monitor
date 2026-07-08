from __future__ import annotations

import asyncio
import re
from collections import Counter
from pathlib import Path
from urllib.parse import quote

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

THREAD = "https://www.drive2.ru/o/b/517173162262135362/"
MAX_PAGE = 144
OUT = Path("reports/drive2_vin_by_manufacturer.xlsx")
VIN_RE = re.compile(r"(?<![A-Z0-9])[A-HJ-NPR-Z0-9]{17}(?![A-Z0-9])", re.I)
FRAME_RE = re.compile(r"(?<![A-Z0-9])(?:[A-Z]{1,5}[A-Z0-9]{0,6}-[A-Z0-9]{5,12})(?![A-Z0-9])", re.I)


def clean_code(value: str) -> str:
    return re.sub(r"[^A-Z0-9-]", "", value.upper())


def guess_brand_from_vin(vin: str) -> str | None:
    prefixes = {
        "WVW": "Volkswagen", "WVG": "Volkswagen", "XW8": "Volkswagen",
        "WAU": "Audi", "WBA": "BMW", "WBS": "BMW", "WDD": "Mercedes-Benz",
        "WDB": "Mercedes-Benz", "ZFA": "Fiat", "ZAR": "Alfa Romeo",
        "VF1": "Renault", "VF3": "Peugeot", "VF7": "Citroën",
        "TMB": "Škoda", "TMA": "Hyundai", "KMH": "Hyundai", "Z94": "Hyundai",
        "KNA": "Kia", "KNE": "Kia", "XWE": "Kia", "KNM": "Renault Samsung",
        "JTD": "Toyota", "JT1": "Toyota", "JT2": "Toyota", "JTE": "Toyota",
        "JTM": "Toyota", "JTN": "Toyota", "JTK": "Toyota", "JMB": "Mitsubishi",
        "JMZ": "Mazda", "JN1": "Nissan", "JN8": "Nissan", "JNK": "Infiniti",
        "JHM": "Honda", "JHL": "Honda", "JF1": "Subaru", "JF2": "Subaru",
        "JS2": "Suzuki", "JS3": "Suzuki", "KL1": "Chevrolet", "XUF": "Chevrolet",
        "1G1": "Chevrolet", "1GC": "Chevrolet", "1FA": "Ford", "1FM": "Ford",
        "WF0": "Ford", "X9F": "Ford", "SAL": "Land Rover", "SAJ": "Jaguar",
        "YV1": "Volvo", "YS3": "Saab", "LVS": "Ford (China)", "LSG": "SAIC/GM",
        "LHG": "Honda (China)", "LVV": "Chery", "LGW": "Great Wall",
        "LDC": "Dongfeng/Citroën", "LGB": "Nissan (China)", "XTA": "LADA",
        "XTT": "УАЗ", "XTC": "КАМАЗ", "XTH": "ГАЗ", "X7L": "Renault Russia",
    }
    return prefixes.get(vin[:3])


async def get_text(page, url: str, wait_ms: int = 500) -> tuple[str, str]:
    try:
        response = await page.goto(url, wait_until="domcontentloaded", timeout=45000)
        await page.wait_for_timeout(wait_ms)
        text = await page.locator("body").inner_text(timeout=10000)
        return text, str(response.status if response else "")
    except Exception as exc:
        return "", f"error: {type(exc).__name__}"


async def collect_codes(page):
    rows = []
    seen_mentions = set()
    empty_streak = 0
    for n in range(1, MAX_PAGE + 1):
        url = THREAD if n == 1 else f"{THREAD}?page={n}#comments"
        text, status = await get_text(page, url, 700)
        found = []
        for regex, kind in ((VIN_RE, "VIN"), (FRAME_RE, "Frame")):
            for match in regex.finditer(text):
                code = clean_code(match.group(0))
                if code and (n, code) not in seen_mentions:
                    seen_mentions.add((n, code))
                    snippet = text[max(0, match.start()-120):match.end()+120].replace("\n", " ")
                    found.append((code, kind, snippet[:500]))
        if not found:
            empty_streak += 1
        else:
            empty_streak = 0
        for code, kind, snippet in found:
            rows.append({"page": n, "code": code, "kind": kind, "snippet": snippet, "drive2_status": status})
        print(f"Drive2 page {n}: {len(found)} codes, status={status}", flush=True)
    return rows


async def decode_ravenol(page, code: str):
    url = f"https://podbor.ravenol.ru/search/?q={quote(code)}"
    text, status = await get_text(page, url, 350)
    normalized = " ".join(text.split())
    brand = None
    model = None
    engine = None
    result_status = "not recognized"

    patterns = [
        r"Марка\s*[:—-]?\s*([A-Za-zА-Яа-яЁё0-9 ._-]{2,40})",
        r"Производитель\s*[:—-]?\s*([A-Za-zА-Яа-яЁё0-9 ._-]{2,40})",
    ]
    for pat in patterns:
        m = re.search(pat, normalized, re.I)
        if m:
            brand = re.split(r"\s{2,}|Модель|Код модели|Кузов|Двигатель", m.group(1))[0].strip(" :-")
            break

    # Breadcrumb / page-result heuristics.
    known = ["Mercedes-Benz", "Volkswagen", "Hyundai", "Toyota", "Nissan", "Mitsubishi", "Mazda", "Honda", "Kia", "BMW", "Audi", "Škoda", "Skoda", "Renault", "Peugeot", "Citroën", "Citroen", "Ford", "Chevrolet", "Opel", "Volvo", "Subaru", "Suzuki", "Lexus", "Infiniti", "Land Rover", "Jaguar", "Porsche", "LADA", "УАЗ", "ГАЗ", "Chery", "Geely", "Haval", "Great Wall", "FAW", "JAC", "BYD", "Dongfeng", "Fiat"]
    if not brand:
        for name in known:
            if re.search(rf"\b{re.escape(name)}\b", normalized, re.I):
                brand = name
                break
    mm = re.search(r"Модель\s*[:—-]?\s*([^|]{2,70}?)(?:Код модели|Кузов|Двигатель|Топливо|$)", normalized, re.I)
    if mm:
        model = mm.group(1).strip(" :-")
    me = re.search(r"Двигатель\s*[:—-]?\s*([A-Z0-9._-]{2,25})", normalized, re.I)
    if me:
        engine = me.group(1)

    if brand:
        result_status = "recognized by Ravenol"
    else:
        brand = guess_brand_from_vin(code) if len(code) == 17 and "-" not in code else None
        if brand:
            result_status = "Ravenol unclear; WMI fallback"
        elif "автомобиль не найден" in normalized.lower() or "ничего не найдено" in normalized.lower():
            result_status = "not found by Ravenol"
        elif not text:
            result_status = f"request failed ({status})"
        else:
            result_status = "Ravenol response ambiguous"
    return {"manufacturer": brand or "Не определён", "model": model or "", "engine": engine or "", "ravenol_status": result_status, "ravenol_url": url, "ravenol_http": status}


def save_excel(mentions, decoded):
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "VIN по производителям"
    headers = ["VIN / Frame-код", "Автопроизводитель", "Модель", "Двигатель", "Тип кода", "Статус проверки", "Количество упоминаний", "Страницы Drive2", "Ссылка Ravenol"]
    ws.append(headers)
    grouped = {}
    for row in mentions:
        grouped.setdefault(row["code"], []).append(row)
    for code in sorted(grouped, key=lambda c: (decoded[c]["manufacturer"], c)):
        group = grouped[code]
        d = decoded[code]
        ws.append([code, d["manufacturer"], d["model"], d["engine"], group[0]["kind"], d["ravenol_status"], len(group), ", ".join(map(str, sorted({r['page'] for r in group}))), d["ravenol_url"]])

    raw = wb.create_sheet("Исходные упоминания")
    raw.append(["Страница", "VIN / Frame-код", "Тип", "Фрагмент комментария", "HTTP Drive2"])
    for r in mentions:
        raw.append([r["page"], r["code"], r["kind"], r["snippet"], r["drive2_status"]])

    summary = wb.create_sheet("Сводка")
    summary.append(["Автопроизводитель", "Уникальных VIN / Frame"])
    counts = Counter(decoded[c]["manufacturer"] for c in grouped)
    for maker, count in counts.most_common():
        summary.append([maker, count])
    summary.append([])
    summary.append(["Показатель", "Значение"])
    summary.append(["Страниц Drive2 проверено", MAX_PAGE])
    summary.append(["Всего упоминаний кодов", len(mentions)])
    summary.append(["Уникальных VIN / Frame", len(grouped)])
    summary.append(["Распознано Ravenol", sum(1 for d in decoded.values() if d["ravenol_status"] == "recognized by Ravenol")])
    summary.append(["Не определено", sum(1 for d in decoded.values() if d["manufacturer"] == "Не определён")])

    header_fill = PatternFill("solid", fgColor="F58220")
    for sheet in wb.worksheets:
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        for col in range(1, sheet.max_column + 1):
            max_len = max((len(str(sheet.cell(r, col).value or "")) for r in range(1, min(sheet.max_row, 250) + 1)), default=8)
            sheet.column_dimensions[get_column_letter(col)].width = min(max(max_len + 2, 12), 45)
    raw.column_dimensions["D"].width = 70
    ws.column_dimensions["I"].width = 55
    wb.save(OUT)
    print(f"Saved {OUT}", flush=True)


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(locale="ru-RU", user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36")
        drive_page = await context.new_page()
        mentions = await collect_codes(drive_page)
        unique_codes = sorted({r["code"] for r in mentions})
        decoded = {}
        ravenol_page = await context.new_page()
        for idx, code in enumerate(unique_codes, 1):
            decoded[code] = await decode_ravenol(ravenol_page, code)
            print(f"Ravenol {idx}/{len(unique_codes)} {code}: {decoded[code]['manufacturer']} ({decoded[code]['ravenol_status']})", flush=True)
            await asyncio.sleep(0.2)
        await browser.close()
    save_excel(mentions, decoded)


if __name__ == "__main__":
    asyncio.run(main())

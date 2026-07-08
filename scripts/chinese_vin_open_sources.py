from __future__ import annotations

import asyncio
import html
import json
import re
from collections import defaultdict
from pathlib import Path
from urllib.parse import quote, unquote, urlparse, parse_qs

import httpx
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUT_XLSX = Path('reports/chinese_vin_open_sources.xlsx')
OUT_JSON = Path('reports/chinese_vin_open_sources.json')
VIN_RE = re.compile(r'(?<![A-Z0-9])[A-HJ-NPR-Z0-9]{17}(?![A-Z0-9])', re.I)

BRANDS = {
    'Chery / EXEED / OMODA / JAECOO / Jetour': ['LVV', 'LVM'],
    'Geely': ['L6T', 'LB3'],
    'GWM / Haval / Tank / Ora / Wey': ['LGW'],
    'Changan': ['LS5'],
    'BYD': ['LC0', 'LPE'],
    'GAC / Trumpchi': ['LMG'],
    'SAIC MG / Roewe': ['LSJ'],
    'SAIC Maxus': ['LSF', 'LSH', 'LSK'],
    'Wuling / Baojun': ['LZW'],
    'DFSK / Dongfeng Sokon': ['LVZ'],
    'BAIC': ['LNB'],
    'Foton': ['LVA', 'LVB', 'LVC'],
    'JAC': ['LJ1'],
    'FAW / Bestune / Hongqi': ['LFB', 'LFP'],
    'Lifan': ['LLV'],
}
PREFIX_TO_BRAND = {p: b for b, ps in BRANDS.items() for p in ps}
TERMS = {
    'Chery': ['LVV','LVM'], 'EXEED': ['LVV'], 'OMODA': ['LVV'],
    'Haval': ['LGW'], 'Tank': ['LGW'], 'Geely': ['L6T','LB3'],
    'Changan': ['LS5'], 'BYD': ['LC0','LPE'], 'GAC': ['LMG'],
    'MG Roewe': ['LSJ'], 'Maxus': ['LSF','LSH','LSK'], 'Wuling': ['LZW'],
    'DFSK': ['LVZ'], 'BAIC': ['LNB'], 'Foton': ['LVA','LVB','LVC'],
    'JAC': ['LJ1'], 'Bestune Hongqi': ['LFB','LFP'], 'Lifan': ['LLV'],
}

TRANSLIT = {**{str(i): i for i in range(10)}, **dict(zip('ABCDEFGHJKLMNPRSTUVWXYZ', [1,2,3,4,5,6,7,8,1,2,3,4,5,7,9,2,3,4,5,6,7,8,9]))}
WEIGHTS = [8,7,6,5,4,3,2,10,0,9,8,7,6,5,4,3,2]

def check_digit(vin: str) -> bool:
    try: total = sum(TRANSLIT[c] * w for c, w in zip(vin, WEIGHTS))
    except KeyError: return False
    return vin[8] == ('X' if total % 11 == 10 else str(total % 11))

def brand_for(vin: str): return PREFIX_TO_BRAND.get(vin[:3])

def context(text: str, vin: str):
    i=text.find(vin)
    return re.sub(r'\s+',' ',text[max(0,i-160):i+len(vin)+220]).strip() if i>=0 else ''

def clean_url(u: str):
    u=html.unescape(u.strip())
    if 'uddg=' in u:
        try: u=unquote(parse_qs(urlparse(u).query)['uddg'][0])
        except Exception: pass
    return u.rstrip(').,]')

def urls_from(text: str):
    vals=re.findall(r'https?://[^\s<>"\']+',text)+re.findall(r'\]\((https?://[^)]+)\)',text)
    out=[]; seen=set()
    for u in vals:
        u=clean_url(u)
        if any(x in u for x in ['google.com/search','bing.com/search','duckduckgo.com/html']): continue
        if u not in seen: seen.add(u); out.append(u)
    return out

async def fetch(client,url,timeout=14):
    try:
        r=await client.get(url,timeout=timeout,follow_redirects=True)
        return r.status_code,r.text
    except Exception as e: return 0,f'ERROR {type(e).__name__}'

async def search(client,q,sem):
    sources=[
        f'https://r.jina.ai/http://www.bing.com/search?q={quote(q)}',
        f'https://html.duckduckgo.com/html/?q={quote(q)}',
    ]
    parts=[]; urls=[]
    async with sem:
        for u in sources:
            s,t=await fetch(client,u)
            parts.append(f'\n{s} {u}\n{t[:150000]}'); urls += urls_from(t)
    return q,urls[:20],'\n'.join(parts)

async def read(client,u,sem):
    async with sem:
        for x in ('https://r.jina.ai/http://'+u.split('://',1)[-1],u):
            s,t=await fetch(client,x,16)
            if s==200 and len(t)>100: return u,s,t
        return u,s,t

async def main():
    queries=[]
    for term,prefixes in TERMS.items():
        queries += [f'"{term}" VIN',f'"{term}" "VIN-код"',f'site:drive2.ru "{term}" VIN',f'site:sudact.ru "{term}" VIN']
        for p in prefixes: queries += [f'"{p}" "{term}"',f'"{p}" VIN']
    queries += ['Росстандарт отзыв Chery VIN приложение','Росстандарт отзыв Haval VIN приложение','Росстандарт отзыв Geely VIN приложение','auction Chery full VIN','auction Haval full VIN','auction Geely full VIN']
    queries=list(dict.fromkeys(queries))
    headers={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36','Accept-Language':'ru-RU,ru;q=0.9,en;q=0.7'}
    sem=asyncio.Semaphore(12)
    records=[]; candidate=[]
    async with httpx.AsyncClient(headers=headers,http2=True,limits=httpx.Limits(max_connections=20)) as client:
        results=await asyncio.gather(*(search(client,q,sem) for q in queries))
        for q,urls,blob in results:
            candidate += urls
            up=blob.upper()
            for m in VIN_RE.finditer(up):
                vin=m.group(); b=brand_for(vin)
                if b: records.append({'vin':vin,'brand':b,'source':'search snippet','url':'','query':q,'context':context(up,vin),'check_digit_valid':check_digit(vin)})
        priority=['rst.gov.ru','sudact.ru','drive2.ru','auto.ru','drom.ru','auction','bid','vin','forum','gov.ru','pdf']
        unique=sorted(set(candidate),key=lambda u:sum(3 for x in priority if x in u.lower()),reverse=True)[:120]
        pages=await asyncio.gather(*(read(client,u,sem) for u in unique))
        for u,s,t in pages:
            up=t.upper()
            for m in VIN_RE.finditer(up):
                vin=m.group(); b=brand_for(vin)
                if b: records.append({'vin':vin,'brand':b,'source':'web page','url':u,'query':'','context':context(up,vin),'check_digit_valid':check_digit(vin),'http_status':s})
    dedup={(r['vin'],r.get('url',''),r.get('query','')):r for r in records}
    records=sorted(dedup.values(),key=lambda r:(r['brand'],r['vin'],r.get('url','')))
    OUT_XLSX.parent.mkdir(parents=True,exist_ok=True)
    OUT_JSON.write_text(json.dumps(records,ensure_ascii=False,indent=2),encoding='utf-8')
    wb=Workbook(); ws=wb.active; ws.title='Упоминания VIN'
    ws.append(['VIN','Бренд / производитель','Контрольная цифра','Источник','URL','Поисковый запрос','Контекст цитирования'])
    for r in records: ws.append([r['vin'],r['brand'],'корректна' if r['check_digit_valid'] else 'не прошла проверку',r['source'],r.get('url',''),r.get('query',''),r.get('context','')])
    sm=wb.create_sheet('Сводка'); sm.append(['Бренд / производитель','Уникальных VIN','Упоминаний'])
    groups=defaultdict(list)
    for r in records: groups[r['brand']].append(r)
    for b,rs in sorted(groups.items()): sm.append([b,len({r['vin'] for r in rs}),len(rs)])
    sm.append([]); sm.append(['Всего уникальных VIN',len({r['vin'] for r in records})]); sm.append(['Проверенных URL',len(unique)])
    fill=PatternFill('solid',fgColor='F58220')
    for sh in wb.worksheets:
        for c in sh[1]: c.fill=fill; c.font=Font(bold=True,color='FFFFFF'); c.alignment=Alignment(wrap_text=True,vertical='top')
        sh.freeze_panes='A2'; sh.auto_filter.ref=sh.dimensions
        for row in sh.iter_rows():
            for c in row: c.alignment=Alignment(wrap_text=True,vertical='top')
        for col in range(1,sh.max_column+1):
            ml=max((len(str(sh.cell(r,col).value or '')) for r in range(1,min(sh.max_row,200)+1)),default=10)
            sh.column_dimensions[get_column_letter(col)].width=min(max(ml+2,12),70)
    ws.column_dimensions['E'].width=55; ws.column_dimensions['G'].width=90
    wb.save(OUT_XLSX)
    print(json.dumps({'records':len(records),'unique_vins':len({r['vin'] for r in records}),'urls':len(unique)},ensure_ascii=False))

if __name__=='__main__': asyncio.run(main())

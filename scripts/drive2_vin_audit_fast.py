from __future__ import annotations

import asyncio
from pathlib import Path

import scripts.drive2_vin_audit as audit


async def decode_all_fast(context, codes: list[str], debug_dir: Path):
    results = {}
    semaphore = asyncio.Semaphore(5)
    lock = asyncio.Lock()
    total = len(codes)
    completed = 0

    async def worker(code: str):
        nonlocal completed
        async with semaphore:
            result = await audit.decode_with_ravenol(context, code, debug_dir)
            await asyncio.sleep(0.15)
        async with lock:
            results[code] = result
            completed += 1
            print(f"Ravenol completed {completed}/{total}: {code} — {result.status}", flush=True)

    await asyncio.gather(*(worker(code) for code in codes))
    return results


audit.decode_all = decode_all_fast

if __name__ == "__main__":
    raise SystemExit(audit.main())

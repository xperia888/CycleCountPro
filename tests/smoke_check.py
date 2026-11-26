"""Smoke test: open index.html and monitor.html in headless Chrome at a few screen sizes.

Fails if a page throws an uncaught JavaScript error or shows nothing.
Screenshots go to tests/out/. Usage: python3 tests/smoke_check.py
"""
import asyncio, functools, http.server, sys, threading
from pathlib import Path
from playwright.async_api import async_playwright

REPO = Path(__file__).resolve().parents[1]
OUT = Path(__file__).resolve().parent / 'out'
OUT.mkdir(exist_ok=True)
PORT = 8780
PAGES = ['index.html', 'monitor.html']
SIZES = [(1920, 1080), (1366, 768), (390, 844)]  # TV / laptop / phone

failed = []


def check(name, ok, detail=''):
    print(f"{'PASS' if ok else 'FAIL'}  {name}{'  — ' + detail if detail and not ok else ''}", flush=True)
    if not ok:
        failed.append(name)


async def main():
    class Quiet(http.server.SimpleHTTPRequestHandler):
        def log_message(self, *a):
            pass
    handler = functools.partial(Quiet, directory=str(REPO))
    site = http.server.ThreadingHTTPServer(('127.0.0.1', PORT), handler)
    threading.Thread(target=site.serve_forever, daemon=True).start()
    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(executable_path='/usr/bin/google-chrome', args=['--no-sandbox'])
            for page_name in PAGES:
                for w, h in SIZES:
                    ctx = await browser.new_context(viewport={'width': w, 'height': h})
                    pg = await ctx.new_page()
                    errors = []
                    pg.on('pageerror', lambda e: errors.append(str(e)))
                    label = f'{page_name} {w}x{h}'
                    try:
                        await pg.goto(f'http://127.0.0.1:{PORT}/{page_name}', wait_until='load', timeout=30000)
                        await pg.wait_for_timeout(2000)
                        text = (await pg.inner_text('body')).strip()
                        check(f'{label}: page shows content', len(text) > 0)
                        check(f'{label}: no JavaScript errors', not errors, '; '.join(errors[:3]))
                    except Exception as e:
                        check(f'{label}: page opens', False, repr(e))
                    await pg.screenshot(path=str(OUT / f"{page_name.replace('.html', '')}_{w}x{h}.png"), full_page=True)
                    await ctx.close()
            await browser.close()
    finally:
        site.shutdown()
    print(f'\n{"All checks passed" if not failed else f"{len(failed)} check(s) failed"}')
    sys.exit(1 if failed else 0)


asyncio.run(main())

import asyncio, sys
from playwright.async_api import async_playwright
import openpyxl
async def run(sample, outname):
    txt=open(sample,encoding='utf-8').read()
    async with async_playwright() as p:
        b = await p.chromium.launch(executable_path='/opt/pw-browsers/chromium-1194/chrome-linux/chrome', args=['--no-sandbox'])
        ctx = await b.new_context(accept_downloads=True)
        pg = await ctx.new_page()
        await pg.goto('file:///home/claude/work/index.html')
        await pg.click('#tabBtnPaste')
        await pg.fill('#pasteNameInput', outname)
        await pg.fill('#pasteArea', txt)
        await pg.click('#pasteAddBtn'); await pg.click('#parseBtn')
        await pg.wait_for_selector('#issueSummary', state='visible')
        print((await pg.inner_text('#logArea')).split('\n')[2])
        print((await pg.inner_text('#issueSummary')).replace('\n',' | ')[:300])
        async with pg.expect_download() as dl: await pg.click('#exportBtn')
        d=await dl.value; path=f'/home/claude/work/tests/{outname}.xlsx'; await d.save_as(path)
        await b.close(); return path
for sample,name in [('wuxi_sample.txt','wuxi'),('xiongqiang_sample.txt','xiongqiang')]:
    try: path=asyncio.run(run(sample,name))
    except Exception as e: print(name,'ERR',e); continue
    import subprocess; print(subprocess.run(['python3','-I','xl_cat.py',path],capture_output=True,text=True).stdout)

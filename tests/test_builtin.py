import asyncio, json, sys
sys.path.insert(0,'/home/claude/work/tests')
from playwright.async_api import async_playwright
import openpyxl
from test_ai import mk

INIT = """
window.__calls = {sample:0, saved:null, prompt:null};
window.claude = { use: async (name) => {
  if(name==='sample'){
    const f = async (input, opts) => {
      window.__calls.sample++; window.__calls.prompt = input;
      const text = '```json\\n' + window.__MOCK + '\\n```';
      opts && opts.onText && opts.onText({text, delta:text});
      return { text, truncated:false, modelTierApplied:'default' };
    };
    f.json = async()=>{ throw {code:'capability_removed'} };
    return f;
  }
  if(name==='downloads'){
    return { save: async ({filename,data}) => { window.__calls.saved = {filename, size: data.byteLength}; window.__dl = data; return {status:'saved'}; } };
  }
  return null;
} };
"""
async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch(executable_path='/opt/pw-browsers/chromium-1194/chrome-linux/chrome', args=['--no-sandbox'])
        ctx = await b.new_context(viewport={'width':420,'height':900})
        pg = await ctx.new_page()
        errs=[]; pg.on('pageerror', lambda e: errs.append(str(e)))
        reqs=[]; pg.on('request', lambda r: reqs.append(r.url) if r.url.startswith('http') else None)
        await pg.add_init_script("window.__MOCK = " + json.dumps(json.dumps(mk(), ensure_ascii=False)) + ";" + INIT)
        await pg.goto('file:///home/claude/work/index.html')
        await pg.wait_for_timeout(600)
        print('key toggle visible:', await pg.is_visible('#aiCfgToggle'), '| builtin note visible:', await pg.is_visible('#aiBuiltinNote'), '| inapp banner:', await pg.is_visible('#inappBanner'))
        await pg.fill('#aiTopic','胸腔闭式引流护理团体标准基础知识')
        await pg.click('#aiGenBtn')
        await pg.wait_for_function("document.getElementById('aiStatus').textContent.includes('已生成')", timeout=15000)
        print((await pg.inner_text('#aiStatus')).replace('\n',' | '))
        print('sample calls', await pg.evaluate('window.__calls.sample'), '| prompt mentions topic:', '胸腔闭式引流' in (await pg.evaluate('window.__calls.prompt')))
        await pg.click('#exportBtn')
        await pg.wait_for_function("window.__calls.saved !== null", timeout=15000)
        print('saved via downloads capability:', await pg.evaluate('window.__calls.saved'))
        data = await pg.evaluate("Array.from(new Uint8Array(window.__dl))")
        open('/home/claude/work/tests/out_builtin.xlsx','wb').write(bytes(data))
        wb=openpyxl.load_workbook('/home/claude/work/tests/out_builtin.xlsx')
        print([ (n, wb[n].max_row-1) for n in ['单选题','多选题','是非题']], [c.value for c in wb['试卷信息'][2]][:6])
        print('external requests:', [u for u in reqs if 'localhost' not in u])
        print('ERRORS', errs)
        await pg.screenshot(path='/home/claude/work/tests/m_builtin.png')
        await b.close()
asyncio.run(main())

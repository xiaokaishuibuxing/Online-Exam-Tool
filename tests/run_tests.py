#!/usr/bin/env python3
"""回归测试：每次改动后、上传前先跑一遍。

用法（在仓库根目录）：
    python3 tests/run_tests.py            # 先自动构建，再跑全部测试
    python3 tests/run_tests.py --no-build # 不重新构建，直接测现有 index.html

依赖：pip install playwright python-docx openpyxl（并已安装 Chromium）
全部通过时退出码为 0，有任何一项失败为 1。
"""
import asyncio, json, os, pathlib, re, subprocess, sys, tempfile

import openpyxl

ROOT = pathlib.Path(__file__).resolve().parent.parent
SAMPLES = ROOT / 'tests' / 'samples'
EXPECTED = json.loads((ROOT / 'tests' / 'expected.json').read_text(encoding='utf-8'))
VERSION = json.loads((ROOT / 'src' / 'version.json').read_text(encoding='utf-8'))
TMP = pathlib.Path(tempfile.mkdtemp(prefix='examtool_'))

results = []


def check(name, ok, detail=''):
    results.append((name, bool(ok), detail))
    print(('  ✓ ' if ok else '  ✗ ') + name + ('' if ok or not detail else f'\n      → {detail}'))


def read_xlsx(path):
    wb = openpyxl.load_workbook(path)
    single = [r[7].value for r in wb['单选题'].iter_rows(min_row=2)]
    multi = [r[9].value for r in wb['多选题'].iter_rows(min_row=2)]
    tf = [r[1].value for r in wb['是非题'].iter_rows(min_row=2)]
    opts_ok = all(all(r[i].value for i in (1, 2, 3, 4)) for r in wb['单选题'].iter_rows(min_row=2))
    info = [c.value for c in wb['试卷信息'][2]][:6]
    return {'single': ''.join(x or '?' for x in single), 'multi': multi, 'tf': tf, 'opts_ok': opts_ok, 'info': info,
            'sheets': wb.sheetnames, 'stems': [r[0].value for r in wb['单选题'].iter_rows(min_row=2)]}


def make_docx(text, path):
    import docx
    d = docx.Document()
    for line in text.split('\n'):
        if line.strip():
            d.add_paragraph(line)
    d.save(str(path))


def launch(p):
    exe = os.environ.get('CHROMIUM_PATH')
    if not exe:
        for cand in sorted(pathlib.Path('/opt/pw-browsers').glob('chromium-*/chrome-linux/chrome')):
            exe = str(cand)
    kw = {'args': ['--no-sandbox']}
    if exe:
        kw['executable_path'] = exe
    return p.chromium.launch(**kw)


async def new_page(browser, html, **ctx_kw):
    ctx = await browser.new_context(accept_downloads=True, **ctx_kw)
    pg = await ctx.new_page()
    pg.errors = []
    pg.requests = []
    pg.on('pageerror', lambda e: pg.errors.append(str(e)))
    pg.on('request', lambda r: pg.requests.append(r.url))
    await pg.goto(html.as_uri())
    return pg


async def add_paste(pg, name, text):
    await pg.click('#tabBtnPaste')
    await pg.fill('#pasteNameInput', name)
    await pg.fill('#pasteArea', text)
    await pg.click('#pasteAddBtn')


async def parse(pg):
    await pg.click('#parseBtn')
    await pg.wait_for_selector('#exportPanel', state='visible')


async def export_to(pg, out):
    """点导出；若有问题需要二次确认（按钮变成"再点一次"），就再点一次。"""
    async with pg.expect_download(timeout=15000) as dl:
        await pg.click('#exportBtn')
        await pg.wait_for_timeout(300)
        if '再点一次' in await pg.inner_text('#exportBtn'):
            await pg.click('#exportBtn')
    d = await dl.value
    await d.save_as(str(out))
    return d.suggested_filename


async def summary(pg):
    cls = await pg.get_attribute('#issueSummary', 'class')
    txt = await pg.inner_text('#issueSummary')
    return cls, txt


# ---------------------------------------------------------------------------
async def t_samples(browser, html):
    print('\n【1】固定测试题：粘贴文字 + 上传 .docx，答案、选项、总分全部核对')
    for name, exp in EXPECTED.items():
        text = (SAMPLES / f'{name}.txt').read_text(encoding='utf-8')
        for mode in ('paste', 'docx'):
            pg = await new_page(browser, html)
            if mode == 'paste':
                await add_paste(pg, name, text)
            else:
                dp = TMP / f'{name}.docx'
                make_docx(text, dp)
                await pg.set_input_files('#fileInput', str(dp))
            await parse(pg)
            cls, txt = await summary(pg)
            out = TMP / f'{name}_{mode}.xlsx'
            await export_to(pg, out)
            r = read_xlsx(out)
            tag = f'{name}/{mode}'
            check(f'{tag}：单选答案', r['single'] == exp['single'], f"{r['single']} != {exp['single']}")
            check(f'{tag}：多选答案', r['multi'] == exp['multi'], f"{r['multi']} != {exp['multi']}")
            check(f'{tag}：判断答案', r['tf'] == exp['tf'], f"{r['tf']} != {exp['tf']}")
            check(f'{tag}：选项已分开（A–D 均有内容）', r['opts_ok'])
            if exp.get('total'):
                check(f'{tag}：总分 {exp["total"]}，每题 5 分', r['info'][2] == str(exp['total']) and r['info'][3:6] == ['5', '5', '5'], str(r['info']))
            if not exp.get('no_standard_check'):
                check(f'{tag}：导出前检查通过', 'good' in cls, txt[:200])
            check(f'{tag}：页面无脚本报错', not pg.errors, str(pg.errors))
            await pg.context.close()


async def t_inline_stem(browser, html):
    print('\n【2】题干内嵌答案：括号里的答案要从题干中清掉')
    pg = await new_page(browser, html)
    await add_paste(pg, 'inline_bracket', (SAMPLES / 'inline_bracket.txt').read_text(encoding='utf-8'))
    await parse(pg)
    out = TMP / 'inline_bracket_stems.xlsx'
    await export_to(pg, out)
    stems = read_xlsx(out)['stems']
    check('括号答案已从题干清除', all(not re.search(r'[（(]\s*[A-H]\s*[）)]', s or '') for s in stems), str(stems))
    pg2 = await new_page(browser, html)
    await add_paste(pg2, 'inline_line', (SAMPLES / 'inline_line.txt').read_text(encoding='utf-8'))
    await parse(pg2)
    out2 = TMP / 'inline_line_stems.xlsx'
    await export_to(pg2, out2)
    wb = openpyxl.load_workbook(out2)
    tf_stems = [r[0].value for r in wb['是非题'].iter_rows(min_row=2)]
    sg_stems = [r[0].value for r in wb['单选题'].iter_rows(min_row=2)]
    check('「解析」内容没有混进题干', all('解析' not in (s or '') and '这是判断' not in (s or '') for s in tf_stems + sg_stems), str(tf_stems + sg_stems))
    check('「答案：」行没有被当成答案区起点（题目没丢）', len(sg_stems) == 2 and len(tf_stems) == 2, f'{len(sg_stems)}/{len(tf_stems)}')
    await pg.context.close(); await pg2.context.close()


async def t_missing_and_confirm(browser, html):
    print('\n【3】答案缺失：标成未识别；有问题时导出需二次确认')
    text = (SAMPLES / 'wuxi.txt').read_text(encoding='utf-8')
    text = re.sub(r'多选\n\n(?:\d\.[A-D]+\n)+', '多选\n\n', text)   # 删掉多选答案
    pg = await new_page(browser, html)
    await add_paste(pg, 'missing', text)
    await parse(pg)
    cls, txt = await summary(pg)
    flagged = await pg.locator('.qrow.flag').count()
    check('5 道多选题标为未识别（橙色）', flagged == 5, str(flagged))
    check('问题列表出现', 'bad' in cls and '未选择答案' in txt, txt[:200])
    btn0 = await pg.inner_text('#exportBtn')
    got = None
    try:
        async with pg.expect_download(timeout=1500) as dl:
            await pg.click('#exportBtn')
        got = await dl.value
    except Exception:
        got = None
    btn1 = await pg.inner_text('#exportBtn')
    check('第一次点击不导出，只提示', got is None and '再点一次' in btn1, btn1)
    out = TMP / 'missing.xlsx'
    name = await export_to(pg, out)
    check('第二次点击才导出', name.endswith('.xlsx'))
    # 修改任意内容会取消"确认"状态
    await pg.wait_for_function("!document.getElementById('exportBtn').disabled")
    await pg.click('#exportBtn')
    armed = '再点一次' in await pg.inner_text('#exportBtn')
    await pg.locator('.qstem-edit').first.fill('改一下题干')
    btn2 = await pg.inner_text('#exportBtn')
    check('修改后取消二次确认状态', armed and '再点一次' not in btn2 and '生成' in btn2, f'{armed} / {btn2}')
    await pg.context.close()


async def t_edit(browser, html):
    print('\n【4】页面上改题：改题干/选项/答案、删题、加题；导出内容跟着改')
    text = (SAMPLES / 'insulin.txt').read_text(encoding='utf-8')
    pg = await new_page(browser, html)
    await add_paste(pg, 'edit', text)
    await parse(pg)
    blocks = pg.locator('.qtype-block')
    first = blocks.nth(0).locator('.qrow.qedit').first
    await first.locator('.qstem-edit').fill('改过的第一题（  ）')
    await first.locator('.opt-text').nth(0).fill('改过的选项A')
    await first.locator('.opt-row input:not(.opt-text)').nth(3).check()          # 单选答案改成 D
    await blocks.nth(1).locator('.qrow.qedit').first.locator('.opt-row input:not(.opt-text)').nth(0).uncheck()   # 多选 ABCD -> BCD
    await blocks.nth(2).locator('.qrow.qedit').first.locator('.tf-select').select_option('正确')   # 判断 错误 -> 正确
    out = TMP / 'edit1.xlsx'
    await export_to(pg, out)
    r = read_xlsx(out)
    wb = openpyxl.load_workbook(out)
    check('改答案：单选第 1 题 → D', r['single'][0] == 'D', r['single'])
    check('改答案：多选第 1 题 → BCD', r['multi'][0] == 'BCD', str(r['multi']))
    check('改答案：判断第 1 题 → 正确', r['tf'][0] == '正确', str(r['tf']))
    check('改题干与选项写进了导出文件', r['stems'][0] == '改过的第一题（  ）' and wb['单选题']['B2'].value == '改过的选项A', f"{r['stems'][0]} / {wb['单选题']['B2'].value}")

    # 删一道单选题：题数 9、总分 95 → 提示
    await blocks.nth(0).locator('.qdel').first.click()
    cls, txt = await summary(pg)
    check('删题后提示"标准是 10 题，现在是 9 题"', '标准是 10 题，现在是 9 题' in txt, txt[:300])
    check('删题后提示总分 95', '总分是 95 分，不是 100 分' in txt, txt[:300])
    # 加一道空白题：提示未选择答案 / 题干为空
    await blocks.nth(0).locator('text=添加一题').click()
    cls, txt = await summary(pg)
    check('新增空白题后提示题干为空、有效选项少于 2 个', '题干为空' in txt and '有效选项少于 2 个' in txt, txt[:400])
    new_row = blocks.nth(0).locator('.qrow.qedit').last
    await new_row.locator('.qstem-edit').fill('新增的题（  ）')
    for i, t in enumerate(['甲', '乙', '丙', '丁']):
        await new_row.locator('.opt-text').nth(i).fill(t)
    await new_row.locator('.opt-row input:not(.opt-text)').nth(1).check()
    cls, txt = await summary(pg)
    check('补全新增题后检查恢复通过', 'good' in cls, txt[:300])
    # 单选勾两个不可能（radio）；多选只勾 1 个 → 提示
    m_row = blocks.nth(1).locator('.qrow.qedit').first
    for i in (1, 2, 3):
        await m_row.locator('.opt-row input:not(.opt-text)').nth(i).uncheck()
    await m_row.locator('.opt-row input:not(.opt-text)').nth(1).check()
    cls, txt = await summary(pg)
    check('多选只选 1 个答案时提示', '多选题至少要有 2 个答案' in txt, txt[:300])
    await pg.context.close()


async def t_counts(browser, html):
    print('\n【5】题数检查：标准题数 + 试卷自己写的题数')
    text = (SAMPLES / 'insulin.txt').read_text(encoding='utf-8')
    # 删掉单选第 10 题（含选项）和它的答案 → 试卷写"共10题"，实际 9 题
    text = re.sub(r'10\. [^\n]*\n(?:A\.[^\n]*\n?)+\n?', '', text, count=1)
    pg = await new_page(browser, html)
    await add_paste(pg, 'counts', text)
    await parse(pg)
    cls, txt = await summary(pg)
    check('标准题数：单选标准是 10 题，现在是 9 题', '标准是 10 题，现在是 9 题' in txt, txt[:400])
    check('试卷自己写的题数：写的是 10 题，识别到 9 题', '试卷里写的单选题是 10 题' in txt, txt[:400])
    check('总题数：试卷写 20 题，识别到 19 题', '试卷里写的总题数是 20 题' in txt, txt[:500])
    # 标准题数填 0 = 不检查
    await pg.click('#cfgToggle')
    await pg.fill('#cfg_expect_single', '0')
    await pg.dispatch_event('#cfg_expect_single', 'input')
    cls, txt = await summary(pg)
    check('标准题数填 0 后不再检查标准', '标准是' not in txt, txt[:300])
    await pg.context.close()


async def t_unsupported(browser, html):
    print('\n【6】选到 .doc / .wps / .pdf：给出明确提示')
    pg = await new_page(browser, html)
    files = {}
    for fn in ('旧试卷.doc', '试卷.wps', '试卷.pdf', '试卷.txt', '图片.png'):
        p = TMP / fn
        p.write_bytes(b'x')
        files[fn] = str(p)
    await pg.set_input_files('#fileInput', list(files.values()))
    shown = await pg.is_visible('#fileNotice')
    txt = await pg.inner_text('#fileNotice')
    check('出现提示框', shown)
    check('.doc 提示另存为 .docx', '旧试卷.doc' in txt and '另存为' in txt)
    check('.wps 提示 WPS 格式', '试卷.wps' in txt and 'WPS' in txt)
    check('.pdf 提示', '试卷.pdf' in txt and 'PDF' in txt)
    check('.txt 提示用粘贴文字', '试卷.txt' in txt and '粘贴文字' in txt)
    check('其他格式也有提示', '图片.png' in txt)
    check('这些文件没有被加入列表', await pg.locator('.file-row').count() == 0)
    dp = TMP / 'ok.docx'
    make_docx((SAMPLES / 'inline_bracket.txt').read_text(encoding='utf-8'), dp)
    await pg.set_input_files('#fileInput', [str(dp), files['旧试卷.doc']])
    check('混选时：.docx 正常加入，.doc 仍有提示', await pg.locator('.file-row').count() == 1 and '旧试卷.doc' in await pg.inner_text('#fileNotice'))
    await pg.set_input_files('#fileInput', [str(dp)])
    check('全部是 .docx 时提示消失', not await pg.is_visible('#fileNotice'))
    acc = await pg.get_attribute('#fileInput', 'accept')
    check('文件选择框不限制扩展名（手机上才能选到并看到提示）', not acc, str(acc))
    await pg.context.close()


async def t_version(browser, html):
    print('\n【7】版本号')
    pg = await new_page(browser, html)
    await pg.clock.set_fixed_time('2026-11-10T09:00:00') if hasattr(pg, 'clock') else None
    await pg.reload()
    txt = await pg.inner_text('#verBadge')
    check(f'角落显示 v{VERSION["version"]} 与更新日期', f'v{VERSION["version"]} · {VERSION["date"]}' in txt, txt)
    check('30 天内显示 NEW 标签', await pg.is_visible('#verNew'))
    await pg.click('#verBadge')
    body = await pg.inner_text('#verPanel')
    check('点击展开本次更新内容', VERSION['changelog'][0]['items'][0] in body)
    await pg.click('#verClose')
    check('可关闭', not await pg.is_visible('#verPanel'))
    foot = await pg.inner_text('footer')
    check('页脚版本号与角标一致', f'v{VERSION["version"]} · {VERSION["date"]}' in foot, foot)
    await pg.context.close()
    # 超过 30 天不再显示 NEW
    pg = await new_page(browser, html)
    if hasattr(pg, 'clock'):
        await pg.clock.set_fixed_time('2027-01-20T09:00:00')
        await pg.reload()
        check('超过 30 天后 NEW 标签自动消失', not await pg.is_visible('#verNew'))
    await pg.context.close()


async def t_offline(browser, html):
    print('\n【8】离线版：没有 AI 出题页，没有联网代码，运行时没有任何网络请求')
    pg = await new_page(browser, html)
    check('没有「AI 出题」标签和面板', await pg.locator('#tabBtnAi, #paneAi, #aiGenBtn, #aiKey').count() == 0)
    check('默认打开「上传 .docx」页', await pg.evaluate("document.getElementById('paneFile').classList.contains('active')"))
    text = (SAMPLES / 'wuxi.txt').read_text(encoding='utf-8')
    await add_paste(pg, 'offline', text)
    await parse(pg)
    await export_to(pg, TMP / 'offline.xlsx')
    ext = [u for u in pg.requests if not u.startswith(('file:', 'data:', 'blob:'))]
    check('整个流程没有任何网络请求', not ext, str(ext))
    src = html.read_text(encoding='utf-8')
    app = re.search(r'<script id="app-logic">(.*?)</script>', src, re.S).group(1)
    bad = [k for k in ('fetch(', 'XMLHttpRequest', 'sendBeacon', 'WebSocket', 'deepseek', 'DeepSeek', 'aiGenBtn') if k in app]
    check('业务代码里没有 fetch / XHR / WebSocket / DeepSeek', not bad, str(bad))
    await pg.context.close()


def t_build():
    print('\n【9】构建')
    out = TMP / 'rebuilt.html'
    subprocess.run([sys.executable, str(ROOT / 'build.py'), '--out', str(out)], check=True, capture_output=True)
    same = out.read_text(encoding='utf-8') == (ROOT / 'index.html').read_text(encoding='utf-8')
    check('index.html 与 src/ 源码一致（没有忘记重新构建）', same)
    on = TMP / 'online.html'
    r = subprocess.run([sys.executable, str(ROOT / 'build.py'), '--edition', 'online', '--out', str(on)], capture_output=True, text=True)
    check('在线版也能从同一份源码构建', r.returncode == 0 and 'aiGenBtn' in on.read_text(encoding='utf-8'), r.stderr[:200])


async def main():
    from playwright.async_api import async_playwright
    if '--no-build' not in sys.argv:
        subprocess.run([sys.executable, str(ROOT / 'build.py')], check=True)
    html = ROOT / 'index.html'
    t_build()
    async with async_playwright() as p:
        browser = await launch(p)
        for fn in (t_samples, t_inline_stem, t_missing_and_confirm, t_edit, t_counts, t_unsupported, t_version, t_offline):
            try:
                await fn(browser, html)
            except Exception as e:  # noqa
                check(f'{fn.__name__} 运行出错', False, repr(e)[:300])
        await browser.close()
    failed = [r for r in results if not r[1]]
    print(f'\n==== 共 {len(results)} 项，通过 {len(results) - len(failed)} 项，失败 {len(failed)} 项 ====')
    for n, _, d in failed:
        print('  ✗', n, ('→ ' + d) if d else '')
    sys.exit(1 if failed else 0)


if __name__ == '__main__':
    asyncio.run(main())

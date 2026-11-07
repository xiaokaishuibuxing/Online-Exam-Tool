#!/usr/bin/env python3
"""把 src/ 拼成单文件 HTML。

用法：
  python3 build.py                 # 离线版 → index.html（默认，部署用这个）
  python3 build.py --edition online  # 在线版（带 AI 出题）→ dist/online.html

源码里用标记区分两个版本：
  <!--@AI:BEGIN--> … <!--@AI:END-->   只在线版保留（HTML）
  /*@AI:BEGIN*/ … /*@AI:END*/         只在线版保留（JS / CSS）
  <!--@OFF:BEGIN--> … <!--@OFF:END--> 只离线版保留
构建离线版时，@AI 区块整段删除，所以离线版里没有 AI 出题页，也没有任何联网代码。
"""
import argparse, json, pathlib, re, sys

root = pathlib.Path(__file__).parent
src = root / 'src'

ap = argparse.ArgumentParser()
ap.add_argument('--edition', choices=['offline', 'online'], default='offline')
ap.add_argument('--out', default=None)
args = ap.parse_args()

ver = json.loads((src / 'version.json').read_text(encoding='utf-8'))
offline = args.edition == 'offline'

MARK = re.compile(r'(?:<!--|/\*)@(AI|OFF):BEGIN(?:-->|\*/)(.*?)(?:<!--|/\*)@\1:END(?:-->|\*/)', re.S)

def apply_edition(text):
    keep = 'OFF' if offline else 'AI'
    return MARK.sub(lambda m: m.group(2) if m.group(1) == keep else '', text)

html = (src / 'template.html').read_text(encoding='utf-8')
app = (src / 'app.js').read_text(encoding='utf-8')
html = apply_edition(html)
app = apply_edition(app)

label = '%s · v%s' % ('离线版' if offline else '在线版', ver['version'])
html = (html.replace('__EDITION_LABEL__', 'v%s · %s' % (ver['version'], 'Offline Edition' if offline else 'Online Edition'))
            .replace('__VERSION_LABEL__', 'v%s · %s' % (ver['version'], ver['date']))
            .replace('__FILE_ACTIVE__', ' active' if offline else ''))
app = app.replace('__VERSION_JSON__', json.dumps(ver, ensure_ascii=False))

out = (html.replace('__JSZIP__', (src / 'jszip.min.js').read_text(encoding='utf-8'))
           .replace('__TEMPLATE_DATA__', (src / 'template_data.js').read_text(encoding='utf-8'))
           .replace('__APP_JS__', app))

assert '__' not in re.sub(r'[A-Za-z0-9_]__[A-Za-z0-9_]', '', ''.join(re.findall(r'__[A-Z_]+__', out))) or True
left = [x for x in re.findall(r'__[A-Z_]{4,}__', out) if x not in ('__proto__',)]
if left:
    sys.exit('未替换的占位符：%s' % sorted(set(left)))

if args.out:
    dst = pathlib.Path(args.out)
else:
    dst = root / ('index.html' if offline else 'dist/online.html')
dst.parent.mkdir(parents=True, exist_ok=True)
dst.write_text(out, encoding='utf-8')
print('built %s edition -> %s (%d bytes, v%s)' % (args.edition, dst, len(out.encode('utf-8')), ver['version']))

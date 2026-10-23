import sys, pathlib
d = pathlib.Path(__file__).parent / 'src'
t = (d/'template.html').read_text(encoding='utf-8')
out = (t.replace('__JSZIP__', (d/'jszip.min.js').read_text(encoding='utf-8'))
        .replace('__TEMPLATE_DATA__', (d/'template_data.js').read_text(encoding='utf-8'))
        .replace('__APP_JS__', (d/'app.js').read_text(encoding='utf-8')))
dst = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else pathlib.Path(__file__).parent / 'index.html')
dst.write_text(out, encoding='utf-8')
print('built', dst, len(out))

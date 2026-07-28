import glob, re

archivos = glob.glob('TBD_VE_*.html')

def clave(a):
    m = re.search(r'TBD_VE_(.+)\.html', a)
    v = m.group(1)
    return (0, int(v)) if v.isdigit() else (1, v)

archivos = sorted(archivos, key=clave)

botones = ''.join(
    f'<a class="btn" href="{a}">{a.replace("TBD_VE_", "").replace(".html", "")}</a>'
    for a in archivos
)

html = f'''<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Reportes TBD</title>
<style>
  *{{box-sizing:border-box}}
  body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
       padding:18px;max-width:480px;margin:auto;background:#fff;color:#202124}}
  h1{{font-size:17px;margin:0 0 2px}}
  .sub{{font-size:12px;color:#5f6368;margin-bottom:16px}}
  .grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
  .btn{{display:flex;align-items:center;justify-content:center;padding:16px 4px;
       background:#f1f3f4;color:#1a56db;border:1px solid #dadce0;border-radius:12px;
       text-decoration:none;font-weight:700;font-size:15px;text-align:center}}
  .btn:active{{background:#1a56db;color:#fff}}
</style></head>
<body>
<h1>Reportes TBD</h1>
<div class="sub">Elegi tu VE</div>
<div class="grid">{botones}</div>
</body></html>'''

open('index.html', 'w').write(html)
print('index.html regenerado con', len(archivos), 'botones')

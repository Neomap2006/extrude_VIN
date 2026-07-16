import sqlite3
import os
import glob

files = sorted(glob.glob(r'Scripts\backup_velasto_*.sqlite'), reverse=True)
if not files:
    print('Tidak ada file backup!')
    exit()

db_file = files[0]
conn = sqlite3.connect(db_file)
cur = conn.cursor()

cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
tables = [r[0] for r in cur.fetchall() if not r[0].startswith('_')]

data = {}
for t in tables:
    cur.execute(f'SELECT * FROM "{t}" LIMIT 50')
    cols = [d[0] for d in cur.description]
    rows = [list(row) for row in cur.fetchall()]
    cur.execute(f'SELECT COUNT(*) FROM "{t}"')
    total = cur.fetchone()[0]
    data[t] = {'cols': cols, 'rows': rows, 'total': total}

conn.close()

def make_table_html(t):
    info = data[t]
    header = ''.join(f'<th>{c}</th>' for c in info['cols'])
    body = ''
    for row in info['rows']:
        cells = ''.join(f'<td title="{str(v) if v else ""}">{str(v) if v else ""}</td>' for v in row)
        body += f'<tr>{cells}</tr>'
    return f'''
    <div class="panel" id="p_{t}">
      <div class="info">Tabel: <b>{t}</b> &mdash; Total: <b>{info["total"]:,} rows</b> (50 pertama)</div>
      <div class="wrap"><table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>
    </div>'''

tabs_html = ''
for t in tables:
    tabs_html += f'<div class="tab" onclick="show(\'{t}\')">{t}<br><span style="color:#475569;font-size:0.7rem">{data[t]["total"]:,} rows</span></div>'

panels_html = ''.join(make_table_html(t) for t in tables)

html = f"""<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<title>Velasto DB Backup Viewer</title>
<style>
body{{font-family:Consolas,monospace;background:#0f172a;color:#e2e8f0;margin:0;padding:20px}}
h1{{color:#38bdf8;font-size:1.4rem;margin-bottom:20px}}
.tabs{{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:20px}}
.tab{{padding:6px 14px;background:#1e293b;border:1px solid #334155;border-radius:6px;cursor:pointer;font-size:0.8rem;color:#94a3b8;text-align:center}}
.tab:hover,.tab.active{{background:#0284c7;color:#fff;border-color:#0284c7}}
.panel{{display:none}}.panel.active{{display:block}}
.info{{color:#64748b;font-size:0.8rem;margin-bottom:10px}}
table{{border-collapse:collapse;width:100%;font-size:0.75rem}}
th{{background:#1e293b;padding:8px 10px;text-align:left;border:1px solid #334155;color:#38bdf8;position:sticky;top:0}}
td{{padding:6px 10px;border:1px solid #1e293b;color:#cbd5e1;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
tr:nth-child(even){{background:#0f172a}}tr:hover{{background:#1e3a5f}}
.wrap{{overflow-x:auto;max-height:500px;overflow-y:auto}}
</style>
</head>
<body>
<h1>Velasto DB Backup Viewer</h1>
<div class="tabs" id="tabs">{tabs_html}</div>
{panels_html}
<script>
function show(t){{
  document.querySelectorAll('.panel').forEach(function(p){{p.classList.remove('active')}});
  document.querySelectorAll('.tab').forEach(function(b){{b.classList.remove('active')}});
  document.getElementById('p_'+t).classList.add('active');
  event.currentTarget.classList.add('active');
}}
document.querySelector('.tab').click();
</script>
</body></html>"""

out = os.path.join('Scripts', 'backup_viewer.html')
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'OK: {out}')

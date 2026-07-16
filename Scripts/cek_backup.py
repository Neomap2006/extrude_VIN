import sqlite3
import os
import glob

# Cari file backup terbaru
files = sorted(glob.glob(r'Scripts\backup_velasto_*.sqlite'), reverse=True)
if not files:
    print('Tidak ada file backup ditemukan!')
    exit()

db_file = files[0]
size_mb = os.path.getsize(db_file) / (1024 * 1024)
print(f'File  : {db_file}')
print(f'Size  : {size_mb:.2f} MB')
print()

conn = sqlite3.connect(db_file)
cur = conn.cursor()

cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
tables = [r[0] for r in cur.fetchall()]

print(f"{'TABEL':<35} {'ROWS':>10}")
print('-' * 48)
total = 0
for t in tables:
    if t.startswith('_'):
        continue
    cur.execute(f'SELECT COUNT(*) FROM "{t}"')
    count = cur.fetchone()[0]
    total += count
    print(f'{t:<35} {count:>10,}')

print('-' * 48)
print(f"{'TOTAL':<35} {total:>10,}")

print()
print('=== INFO BACKUP ===')
cur.execute('SELECT key, value FROM _backup_info')
for k, v in cur.fetchall():
    print(f'  {k}: {v}')

conn.close()

# -*- coding: utf-8 -*-
import sys
import psycopg

sys.stdout.reconfigure(encoding='utf-8')

conn = psycopg.connect('host=localhost port=5432 user=postgres password=postgres dbname=taichung_house')
cur = conn.cursor()
cur.execute("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public';")
tables = [r[0] for r in cur.fetchall()]
print('Tables in DB:', tables)

for t in tables:
    cur.execute(f"SELECT count(*) FROM {t};")
    cnt = cur.fetchone()[0]
    print(f"  {t}: {cnt} rows")

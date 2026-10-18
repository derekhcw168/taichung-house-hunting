# -*- coding: utf-8 -*-
import sys
import psycopg

sys.stdout.reconfigure(encoding='utf-8')

conn = psycopg.connect('host=localhost port=5432 user=postgres password=postgres dbname=taichung_house')
cur = conn.cursor()
cur.execute("SELECT column_name, data_type FROM information_schema.columns WHERE table_name = 'communities';")
print('Columns in communities:')
for c in cur.fetchall():
    print(f"  {c[0]} ({c[1]})")

cur.execute("SELECT * FROM communities LIMIT 2;")
print('Sample rows:')
for r in cur.fetchall():
    print(r)

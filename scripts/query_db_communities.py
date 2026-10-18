# -*- coding: utf-8 -*-
import sys
import psycopg

sys.stdout.reconfigure(encoding='utf-8')

conn = psycopg.connect('host=localhost port=5432 user=postgres password=postgres dbname=taichung_house')
cur = conn.cursor()

names = [
    '長億城', '工學天下', '龍邦綠園道', '文心高第', '大東家華園',
    '假日半島', '文心大地', '文心及第', '大安黃金城', '京城特區',
    '長億南園', '大墩人文'
]

print("=== CHECKING COMMUNITIES IN DATABASE ===")
for name in names:
    cur.execute("SELECT id, name, district, address, total_units, completion_year, builder FROM communities WHERE name LIKE %s;", (f"%{name}%",))
    rows = cur.fetchall()
    if rows:
        for r in rows:
            print(f"ID: {r[0]:<3} | 名稱: {r[1]:<16} | 區域: {r[2]} | 地址: {r[3]} | 戶數: {r[4]} | 完工年: {r[5]} | 建商: {r[6]}")
    else:
        print(f"❌ 未在資料庫找到: {name}")

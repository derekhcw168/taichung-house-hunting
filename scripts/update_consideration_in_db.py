import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import psycopg

conn = psycopg.connect('host=localhost port=5432 user=postgres password=postgres dbname=taichung_house')
cur = conn.cursor()

# 1. Update 大安紫金城 (id=3) -> consider
cur.execute("UPDATE properties SET decision_status = 'consider', updated_at = NOW() WHERE id = 3")
print("Updated 大安紫金城 status to consider")

# 2. Update 比佛利山莊13樓 (id=79) -> consider
cur.execute("UPDATE properties SET decision_status = 'consider', updated_at = NOW() WHERE id = 79")
print("Updated 比佛利山莊13樓 status to consider")

# 3. Insert or check 長億南園8樓
cur.execute("SELECT id FROM properties WHERE community_name = '長億南園8樓'")
row = cur.fetchone()
if not row:
    cur.execute('''
        INSERT INTO properties (
            code, community_name, city, district, address, title, price_total, unit_price,
            rooms, living_rooms, bathrooms, layout_raw, floor_info, current_floor, total_floors,
            age, total_area, indoor_area, attached_area, indoor_total, public_area, land_area,
            public_ratio, public_ratio_has_parking, public_ratio_desc, parking_type, parking_desc,
            orientation, management_fee, decision_status, original_source_platform, original_url,
            showing_agent, created_at, updated_at
        ) VALUES (
            'PROP_106', '長億南園8樓', '台中市', '南區', '台中市南區德富路390號8樓之1',
            '長億南園8樓優質大空間｜室內32坪大三房附坡道機械車位', 998, 23.63,
            3, 2, 2, '3房2廳2衛', '8F/14F', 8, 14,
            34, 42.23, 28.60, 3.45, 32.05, 10.18, 5.05,
            19.2, true, '含機械車位約2.55坪，純公設約7.63坪', '坡道機械式', 'B1-127坡道機械車位',
            '坐北朝南', '1800元/月', 'consider', '永義房屋', 'https://buy.yungyi-house.com.tw/house/YC1038914',
            '永義房屋周小姐', NOW(), NOW()
        ) RETURNING id
    ''')
    new_id = cur.fetchone()[0]
    print(f"Inserted 長億南園8樓 with ID {new_id}")
else:
    cur.execute("UPDATE properties SET decision_status = 'consider', updated_at = NOW() WHERE id = %s", (row[0],))
    print(f"長億南園8樓 already exists (id={row[0]}), updated status to consider")

# 4. Insert or check 歌德堡8樓
cur.execute("SELECT id FROM properties WHERE community_name = '歌德堡8樓'")
row = cur.fetchone()
if not row:
    cur.execute('''
        INSERT INTO properties (
            code, community_name, city, district, address, title, price_total, unit_price,
            rooms, living_rooms, bathrooms, layout_raw, floor_info, current_floor, total_floors,
            age, total_area, indoor_area, attached_area, indoor_total, public_area, land_area,
            public_ratio, public_ratio_has_parking, public_ratio_desc, parking_type, parking_desc,
            orientation, management_fee, decision_status, original_source_platform, original_url,
            showing_agent, created_at, updated_at
        ) VALUES (
            'PROP_107', '歌德堡8樓', '台中市', '南區', '台中市南區德富路60號8樓之1',
            '歌德堡次頂樓大三房｜中山醫大慶雙鐵商圈裝潢好免整理', 988, 32.96,
            3, 2, 2, '3房2廳2衛', '8F/9F', 8, 9,
            28, 29.98, 20.03, 2.04, 22.07, 7.91, 6.08,
            26.38, false, '超低公設比僅26.4%，室內得房率高達73.6%', '無車位', '專用配屬機車位B1-59',
            '大門朝東北客廳朝西北', '未載明', 'consider', '永義房屋', 'https://buy.yungyi-house.com.tw',
            '永義房屋周小姐', NOW(), NOW()
        ) RETURNING id
    ''')
    new_id = cur.fetchone()[0]
    print(f"Inserted 歌德堡8樓 with ID {new_id}")
else:
    cur.execute("UPDATE properties SET decision_status = 'consider', updated_at = NOW() WHERE id = %s", (row[0],))
    print(f"歌德堡8樓 already exists (id={row[0]}), updated status to consider")

# 5. Insert or check 大慶國宅10樓
cur.execute("SELECT id FROM properties WHERE community_name = '大慶國宅10樓'")
row = cur.fetchone()
if not row:
    cur.execute('''
        INSERT INTO properties (
            code, community_name, city, district, address, title, price_total, unit_price,
            rooms, living_rooms, bathrooms, layout_raw, floor_info, current_floor, total_floors,
            age, total_area, indoor_area, attached_area, indoor_total, public_area, land_area,
            public_ratio, public_ratio_has_parking, public_ratio_desc, parking_type, parking_desc,
            orientation, management_fee, decision_status, original_source_platform, original_url,
            showing_agent, created_at, updated_at
        ) VALUES (
            'PROP_108', '大慶國宅10樓', '台中市', '南區', '台中市南區工學二街16號10樓',
            '大慶國宅次高樓景觀大三房｜921後新法規耐震建築室內近28坪', 998, 22.43,
            3, 2, 2, '3房2廳2衛', '10F/12F', 10, 12,
            24, 44.50, 25.37, 2.40, 27.77, 16.73, 8.50,
            37.59, false, '大社區千坪中庭花園，公設比約37.6%', '無車位(現承租平車)', '現況承租管委會地下平面車位1500元/月',
            '客廳朝東大門朝南', '1500元/月', 'consider', '永義房屋', 'https://buy.yungyi-house.com.tw',
            '永義房屋周小姐', NOW(), NOW()
        ) RETURNING id
    ''')
    new_id = cur.fetchone()[0]
    print(f"Inserted 大慶國宅10樓 with ID {new_id}")
else:
    cur.execute("UPDATE properties SET decision_status = 'consider', updated_at = NOW() WHERE id = %s", (row[0],))
    print(f"大慶國宅10樓 already exists (id={row[0]}), updated status to consider")

conn.commit()
conn.close()
print("All PostgreSQL updates successfully committed!")

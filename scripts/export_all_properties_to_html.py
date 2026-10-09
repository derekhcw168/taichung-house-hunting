import psycopg
import json
import re
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML_FILE = os.path.join(BASE_DIR, "591看屋物件地圖與比較分析.html")

def export_db_to_html():
    conn = psycopg.connect('host=localhost port=5432 user=postgres password=postgres dbname=taichung_house')
    cur = conn.cursor()
    cur.execute("""
        SELECT 
            p.id, p.code, p.community_name, p.title, p.price_total, p.unit_price,
            p.rooms, p.living_rooms, p.bathrooms, p.layout_raw, p.floor_info,
            p.current_floor, p.total_floors, p.age, p.total_area, p.indoor_area,
            p.attached_area, p.indoor_total, p.public_area, p.public_ratio,
            p.public_ratio_has_parking, p.public_ratio_desc, p.parking_type,
            p.parking_desc, p.orientation, p.management_fee, p.decision_status,
            p.original_source_platform, p.original_url, p.district,
            COUNT(pi.id) as img_count,
            MIN(pi.file_relpath) as sample_img
        FROM properties p
        LEFT JOIN property_images pi ON p.id = pi.property_id
        GROUP BY p.id
        ORDER BY 
            CASE p.decision_status
                WHEN 'consider' THEN 1
                WHEN 'pending' THEN 2
                WHEN 'rejected' THEN 3
                WHEN 'deleted' THEN 4
                ELSE 5
            END,
            p.id ASC;
    """)
    rows = cur.fetchall()
    cols = [d[0] for d in cur.description]
    
    properties = []
    coords = {
        '南屯區': (24.138, 120.645),
        '南區': (24.118, 120.662),
        '西區': (24.145, 120.660),
        '西屯區': (24.165, 120.645),
        '北區': (24.160, 120.680),
        '北屯區': (24.175, 120.690),
        '東區': (24.135, 120.695),
        '中區': (24.140, 120.680)
    }

    for r in rows:
        item = dict(zip(cols, r))
        for k, v in item.items():
            if hasattr(v, '__float__'):
                item[k] = float(v)

        dist = item.get('district') or '南屯區'
        base_lat, base_lng = coords.get(dist, (24.138, 120.650))
        jitter_lat = ((item['id'] * 17) % 50 - 25) * 0.0006
        jitter_lng = ((item['id'] * 23) % 50 - 25) * 0.0006
        lat = round(base_lat + jitter_lat, 5)
        lng = round(base_lng + jitter_lng, 5)

        mgmt_fee_str = item.get('management_fee') or '未載明'
        mgmt_num = int(re.sub(r'[^0-9]', '', mgmt_fee_str)) if re.sub(r'[^0-9]', '', mgmt_fee_str) else 0

        floor_val = item.get('floor_info')
        if not floor_val or 'None' in str(floor_val):
            if item.get('current_floor') and item.get('total_floors'):
                floor_val = f"{item['current_floor']}/{item['total_floors']}F"
            else:
                floor_val = "未載明"

        layout_val = item.get('layout_raw')
        if not layout_val or 'None' in str(layout_val):
            layout_val = "未載明"

        p_obj = {
            "id": item['id'],
            "code": item.get('code') or f"PROP_{item['id']:03d}",
            "name": item.get('community_name') or '未具名',
            "title": item.get('title') or '',
            "price": item.get('price_total') or 0,
            "unitPrice": item.get('unit_price') or 0,
            "layout": layout_val,
            "rooms": item.get('rooms') or 3,
            "baths": item.get('bathrooms') or 2,
            "floor": floor_val,
            "floorNum": item.get('current_floor') or 5,
            "totalFloors": item.get('total_floors') or 10,
            "age": f"{item['age']}年" if item.get('age') else '未載明',
            "ageNum": item.get('age') or 30,
            "totalArea": item.get('total_area') or 0,
            "mainArea": item.get('indoor_area') or 0,
            "subArea": item.get('attached_area') or 0,
            "indoor": item.get('indoor_total') or 25,
            "parkingType": item.get('parking_type') or '無車位',
            "parkingNote": item.get('parking_desc') or item.get('parking_type') or '無車位',
            "mgmtFee": mgmt_fee_str,
            "mgmtFeeNum": mgmt_num,
            "publicRatio": f"{item['public_ratio']}%" if item.get('public_ratio') else '未載明',
            "publicRatioNum": item.get('public_ratio') or 25,
            "publicRatioDetail": item.get('public_ratio_desc') or '',
            "district": item.get('district') or '南屯區',
            "area": item.get('district') or '台中市',
            "orientation": item.get('orientation') or '未載明',
            "note": item.get('title') or '',
            "lat": lat,
            "lng": lng,
            "url": item.get('original_url') or '',
            "category": item.get('decision_status') or 'pending',
            "themeCategory": 'plane' if (item.get('parking_type') and '平面' in item.get('parking_type')) else 'view',
            "isNew": False,
            "source": item.get('original_source_platform') or '591',
            "imgCount": item.get('img_count') or 0,
            "sampleImg": item.get('sample_img') or ''
        }
        properties.append(p_obj)

    conn.close()

    print(f"Generated {len(properties)} properties from PostgreSQL.")

    js_array = "let allProperties = " + json.dumps(properties, ensure_ascii=False, indent=2) + ";\n"

    with open(HTML_FILE, 'r', encoding='utf-8') as f:
        content = f.read()

    # Regex replace the existing allProperties declaration
    # Starts with (const|let) allProperties = [ ... ];
    pattern = r'(const|let)\s+allProperties\s*=\s*\[[\s\S]*?\];'
    if not re.search(pattern, content):
        raise ValueError("Could not find allProperties array in HTML file!")

    new_content = re.sub(pattern, js_array.strip(), content, count=1)

    with open(HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(new_content)

    print(f"Successfully updated 591看屋物件地圖與比較分析.html with all {len(properties)} properties!")

if __name__ == '__main__':
    export_db_to_html()

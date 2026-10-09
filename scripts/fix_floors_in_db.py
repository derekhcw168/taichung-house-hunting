import psycopg
import urllib.request
import ssl
import json
import re
import sys

if sys.stdout:
    sys.stdout.reconfigure(encoding='utf-8')

ctx = ssl._create_unverified_context()
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Device': 'pc',
    'deviceid': 'xyz123'
}

def get_591_details(house_id):
    url = f'https://bff.591.com.tw/v1/house/sale/detail?id={house_id}'
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, context=ctx, timeout=8) as resp:
        data = json.loads(resp.read().decode('utf-8'))
        base_info = data.get('data', {}).get('baseInfo', {})
        info_list = base_info.get('info', [])
        floor_str = None
        age_str = None
        mgmt_str = None
        for item in info_list:
            if item.get('name') == '樓層':
                floor_str = item.get('value')
            elif item.get('name') == '屋齡':
                age_str = item.get('value')
            elif item.get('name') == '管理費':
                mgmt_str = item.get('value')
        
        # Parse floor_str (e.g., "9F/22F" or "9/22F" or "9F")
        cur_floor = None
        tot_floors = None
        if floor_str:
            m = re.search(r'(\d+)\s*F?\s*[/／]\s*(\d+)\s*F?', floor_str)
            if m:
                cur_floor = int(m.group(1))
                tot_floors = int(m.group(2))
                floor_str = f"{cur_floor}/{tot_floors}F"
            else:
                m_single = re.search(r'(\d+)\s*F?', floor_str)
                if m_single:
                    cur_floor = int(m_single.group(1))

        age_val = None
        if age_str:
            m_age = re.search(r'([0-9.]+)', age_str)
            if m_age:
                age_val = float(m_age.group(1))

        return {
            'floor': floor_str,
            'current_floor': cur_floor,
            'total_floors': tot_floors,
            'age': age_val,
            'mgmt_fee': mgmt_str
        }

def main():
    conn = psycopg.connect('host=localhost port=5432 user=postgres password=postgres dbname=taichung_house')
    cur = conn.cursor()

    cur.execute("SELECT id, community_name, floor_info, current_floor, total_floors, original_url FROM properties ORDER BY id;")
    rows = cur.fetchall()

    print(f"Total properties in DB: {len(rows)}")

    # Known manual/specific fixes for Yungyi/Sinyi if needed
    special_fixes = {
        74: {'floor': '4/13F', 'current_floor': 4, 'total_floors': 13}, # 永義 7475810: 4F/共13F
        78: {'floor': '13/14F', 'current_floor': 13, 'total_floors': 14}, # 永義 7300671: 13F/共14F
    }

    updated_count = 0

    for r in rows:
        pid, name, f_info, c_floor, t_floor, url = r
        
        # Check special fixes first
        if pid in special_fixes:
            fix = special_fixes[pid]
            cur.execute("""
                UPDATE properties 
                SET floor_info = %s, current_floor = %s, total_floors = %s
                WHERE id = %s
            """, (fix['floor'], fix['current_floor'], fix['total_floors'], pid))
            print(f"✅ ID {pid} [{name}]: 更新樓層為 {fix['floor']} ({fix['current_floor']}/{fix['total_floors']})")
            updated_count += 1
            continue

        # If 591 URL
        if url and '591' in url:
            m_id = re.search(r'(?:detail/2/|sale/|/sale/|house/)(\d+)', url)
            if m_id:
                hid = m_id.group(1)
                try:
                    res = get_591_details(hid)
                    new_floor = res['floor']
                    new_cur = res['current_floor']
                    new_tot = res['total_floors']
                    
                    if new_floor and (f_info in ['高樓層', '中樓層', '低樓層', None, ''] or not c_floor or new_floor != f_info):
                        sql_parts = ["floor_info = %s"]
                        sql_vals = [new_floor]
                        if new_cur is not None:
                            sql_parts.append("current_floor = %s")
                            sql_vals.append(new_cur)
                        if new_tot is not None:
                            sql_parts.append("total_floors = %s")
                            sql_vals.append(new_tot)
                        if res['age'] is not None:
                            sql_parts.append("age = %s")
                            sql_vals.append(res['age'])
                        if res['mgmt_fee']:
                            sql_parts.append("management_fee = %s")
                            sql_vals.append(res['mgmt_fee'])
                        
                        sql_vals.append(pid)
                        cur.execute(f"UPDATE properties SET {', '.join(sql_parts)} WHERE id = %s", sql_vals)
                        print(f"✅ ID {pid} [{name}]: 舊樓層 [{f_info}] -> 新精確樓層 [{new_floor}] (現居 {new_cur}F / 總高 {new_tot}F)")
                        updated_count += 1
                except Exception as e:
                    print(f"⚠️ ID {pid} [{name}] 591 解析失敗: {e}")

    conn.commit()
    conn.close()
    print(f"\n🎉 樓層補正完成！共更新了 {updated_count} 筆物件。")

if __name__ == '__main__':
    main()

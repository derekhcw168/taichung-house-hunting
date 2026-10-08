import psycopg
import re, os, sys

sys.stdout.reconfigure(encoding='utf-8')

def import_921():
    print("Reading 921 disaster markdown file...")
    md_path = '台中921受影響社區大樓完整查核名冊與避雷對照表.md'
    if not os.path.exists(md_path):
        print(f"File not found: {md_path}")
        return

    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.splitlines()
    records = []
    current_category = ""
    
    # Simple markdown table parser
    in_table = False
    headers = []
    
    for line in lines:
        line_s = line.strip()
        if line_s.startswith('### 第一類：'):
            current_category = "第一類：全倒／倒塌拆除原地或易地重建"
        elif line_s.startswith('### 第二類：'):
            current_category = "第二類：全倒拆除但未原地重建（改為公園日照或懸宕）"
        elif line_s.startswith('### 第三類：'):
            current_category = "第三類：半倒／結構受創補強列管（未拆除修繕解管）"
        elif line_s.startswith('### 第四類：'):
            current_category = "第四類：網路討論提及但非倒塌拆除（舊耐震法規老屋）"
        elif line_s.startswith('|') and ('---' not in line_s):
            cols = [c.strip() for c in line_s.split('|')[1:-1]]
            if not cols:
                continue
            if '行政區' in cols[0] or '社區名稱' in cols[0]:
                headers = cols
                in_table = True
                continue
            
            # Data row
            if len(cols) >= 3 and current_category:
                # Clean bold markdown markers
                c_clean = [re.sub(r'[*`]', '', c) for c in cols]
                records.append({
                    'category': current_category,
                    'raw_cols': c_clean
                })
        elif not line_s.startswith('|'):
            in_table = False

    print(f"Extracted {len(records)} raw 921 records.")

    conn = psycopg.connect("host=localhost port=5432 user=postgres password=postgres dbname=taichung_house", autocommit=True)
    with conn.cursor() as cur:
        # Clear existing records
        cur.execute("TRUNCATE TABLE community_earthquake_records RESTART IDENTITY CASCADE;")
        
        inserted = 0
        for r in records:
            cat = r['category']
            cols = r['raw_cols']
            district = ""
            comm_name = ""
            damage = cat
            legal = ""
            repair = ""
            rebuilt = ""
            notes = ""

            if "第一類" in cat:
                # | 行政區（原區） | 社區大樓原名 | 災前樓層／戶數 | 921 災損狀況與司法判定 | 重建途徑與重建後新案名 | 現況與購屋備註 |
                district = cols[0] if len(cols) > 0 else ""
                comm_name = cols[1] if len(cols) > 1 else ""
                damage = cols[3] if len(cols) > 3 else "全倒拆除"
                rebuilt = cols[4] if len(cols) > 4 else ""
                notes = f"災前: {cols[2] if len(cols)>2 else ''} | 處置: {cols[4] if len(cols)>4 else ''} | 備註: {cols[5] if len(cols)>5 else ''}"
            elif "第二類" in cat:
                # | 行政區（原區） | 社區大樓原名 | 災前狀況 | 921 災損狀況與司法判定 | 最終處置與現狀 | 買房避坑備註 |
                district = cols[0] if len(cols) > 0 else ""
                comm_name = cols[1] if len(cols) > 1 else ""
                damage = cols[3] if len(cols) > 3 else "全倒拆除"
                repair = cols[4] if len(cols) > 4 else ""
                notes = f"最終處置: {cols[4] if len(cols)>4 else ''} | 避坑: {cols[5] if len(cols)>5 else ''}"
            elif "第三類" in cat:
                # | 行政區 | 社區大樓名稱 | 座落地點 | 921 歷史受損與列管紀錄 | 修繕補強歷程與現況 | 買房評估與安全性警示 |
                district = cols[0] if len(cols) > 0 else ""
                comm_name = cols[1] if len(cols) > 1 else ""
                legal = cols[3] if len(cols) > 3 else ""
                repair = cols[4] if len(cols) > 4 else ""
                damage = "半倒/結構受損列管修復"
                notes = f"座落: {cols[2] if len(cols)>2 else ''} | 警示: {cols[5] if len(cols)>5 else ''}"
            elif "第四類" in cat:
                # | 社區名稱 | 實際狀況與真相釐清 | 買房安全與現況判讀 |
                comm_name = cols[0] if len(cols) > 0 else ""
                damage = "未倒塌/非全半倒"
                notes = f"真相釐清: {cols[1] if len(cols)>1 else ''} | 現況判讀: {cols[2] if len(cols)>2 else ''}"

            if not comm_name:
                continue

            # First ensure community exists in communities table if not already
            cur.execute("""
                INSERT INTO communities (name, district)
                VALUES (%s, %s)
                ON CONFLICT (name) DO UPDATE SET district = COALESCE(communities.district, EXCLUDED.district)
                RETURNING id;
            """, (comm_name, district))
            cid = cur.fetchone()[0]

            cur.execute("""
                INSERT INTO community_earthquake_records 
                (community_id, original_community_name, district, damage_level, legal_status, repair_status, rebuilt_name, notes)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """, (cid, comm_name, district, damage, legal, repair, rebuilt, notes))
            inserted += 1

        print(f"Successfully inserted {inserted} 921 earthquake disaster records into community_earthquake_records!")
    conn.close()

if __name__ == '__main__':
    import_921()

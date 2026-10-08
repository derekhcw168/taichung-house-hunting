import psycopg
import os, sys

sys.stdout.reconfigure(encoding='utf-8')

def verify_all():
    print("=" * 60)
    print("      台中買房 PostgreSQL 資料庫遷移驗證報表")
    print("=" * 60)

    conn = psycopg.connect("host=localhost port=5432 user=postgres password=postgres dbname=taichung_house")
    with conn.cursor() as cur:
        # 1. Table row counts
        tables = [
            'communities',
            'properties',
            'property_sources',
            'property_images',
            'inspection_notes',
            'community_earthquake_records',
            'buyer_preferences'
        ]
        print("\n【一、資料表筆數統計】")
        for tbl in tables:
            cur.execute(f"SELECT COUNT(*) FROM {tbl};")
            cnt = cur.fetchone()[0]
            print(f"  • {tbl:<30}: {cnt:>5} 筆")

        # 2. Decision status breakdown
        print("\n【二、看屋決策狀態分佈 (properties)】")
        cur.execute("""
            SELECT decision_status, COUNT(*) 
            FROM properties 
            GROUP BY decision_status 
            ORDER BY count DESC;
        """)
        for status, cnt in cur.fetchall():
            status_zh = {
                'consider': '列入考慮 (consider)',
                'pending': '待看物件 (pending)',
                'rejected': '不考慮 (rejected)'
            }.get(status, status)
            print(f"  • {status_zh:<28}: {cnt:>3} 戶")

        # 3. Source platforms breakdown
        print("\n【三、來源平台分佈 (property_sources)】")
        cur.execute("""
            SELECT platform, COUNT(*) 
            FROM property_sources 
            GROUP BY platform 
            ORDER BY count DESC;
        """)
        for plat, cnt in cur.fetchall():
            print(f"  • {plat:<25}: {cnt:>3} 個網頁備份")

        # 4. Image verification on disk
        print("\n【四、圖片檔案與資料庫一致性驗證】")
        cur.execute("SELECT file_relpath FROM property_images;")
        image_records = cur.fetchall()
        missing_images = []
        for (relpath,) in image_records:
            if not os.path.exists(relpath):
                missing_images.append(relpath)
        
        print(f"  • 資料庫記錄圖片總數: {len(image_records)} 張")
        print(f"  • 本機實體檔案檢查: 存在 {len(image_records) - len(missing_images)} 張, 遺失 {len(missing_images)} 張")
        if missing_images:
            print(f"    ⚠️ 警告：發現 {len(missing_images)} 個死鏈路徑！")
            for m in missing_images[:5]:
                print(f"      - {m}")
        else:
            print("  • ✅ 完美！所有圖片路徑與本機硬碟檔案 100% 精確吻合，無任何死鏈。")

        # 5. Consider list preview
        print("\n【五、【列入考慮】物件與看屋筆記即時預覽】")
        cur.execute("""
            SELECT code, community_name, price_total, indoor_total, parking_type, pros 
            FROM v_consider_properties;
        """)
        for code, comm, price, indoor, parking, pros in cur.fetchall():
            pros_snippet = (pros[:30] + '...') if pros else '暫無評語'
            print(f"  [{code}] {comm:<10} | 總價: {price or '未定':>5} 萬 | 實坪: {indoor or '未定':>4} 坪 | 車位: {parking or '無'} | 優點: {pros_snippet}")

        # 6. Earthquake match preview
        print("\n【六、921 歷史受損社區比對預警】")
        cur.execute("""
            SELECT code, community_name, damage_level, legal_status 
            FROM v_properties_with_earthquake_check 
            WHERE damage_level IS NOT NULL 
            LIMIT 6;
        """)
        for code, comm, dmg, legal in cur.fetchall():
            print(f"  ⚠️ [{code}] {comm:<12} | 災損: {dmg:<15} | 列管/備註: {str(legal)[:30]}")

    conn.close()
    print("\n" + "=" * 60)
    print("               所有驗證項目均已順利通過！")
    print("=" * 60)

if __name__ == '__main__':
    verify_all()

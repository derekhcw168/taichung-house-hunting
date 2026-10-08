import psycopg
import sys

sys.stdout.reconfigure(encoding='utf-8')

def create_views():
    conn = psycopg.connect("host=localhost port=5432 user=postgres password=postgres dbname=taichung_house", autocommit=True)
    with conn.cursor() as cur:
        # 1. Link inspection_notes to properties
        print("Linking inspection_notes to properties...")
        cur.execute("""
            UPDATE inspection_notes inote
            SET property_id = p.id
            FROM properties p
            WHERE p.community_name ILIKE '%' || inote.community_name || '%'
               OR inote.community_name ILIKE '%' || p.community_name || '%';
        """)

        # 2. View: v_active_properties (決策狀態統計與重要指標)
        cur.execute("""
            CREATE OR REPLACE VIEW v_active_properties AS
            SELECT 
                p.id,
                p.code,
                p.decision_status,
                p.community_name,
                p.district,
                p.title,
                p.price_total,
                p.unit_price,
                p.rooms,
                p.living_rooms,
                p.bathrooms,
                p.layout_raw,
                p.floor_info,
                p.age,
                p.indoor_total,
                p.total_area,
                p.parking_type,
                p.parking_desc,
                p.management_fee,
                p.original_source_platform,
                p.original_url,
                COUNT(pi.id) AS image_count
            FROM properties p
            LEFT JOIN property_images pi ON p.id = pi.property_id
            GROUP BY p.id
            ORDER BY 
                CASE p.decision_status
                    WHEN 'consider' THEN 1
                    WHEN 'pending' THEN 2
                    WHEN 'rejected' THEN 3
                    ELSE 4
                END,
                p.price_total ASC;
        """)

        # 3. View: v_consider_properties (所有列入考慮之物件與深度筆記)
        cur.execute("""
            CREATE OR REPLACE VIEW v_consider_properties AS
            SELECT 
                p.code,
                p.community_name,
                p.district,
                p.price_total,
                p.indoor_total,
                p.parking_type,
                p.floor_info,
                p.age,
                inote.status_conclusion,
                inote.pros,
                inote.cons,
                inote.suggested_offer_low,
                inote.suggested_offer_high,
                inote.renovation_estimate,
                p.original_url
            FROM properties p
            LEFT JOIN inspection_notes inote ON p.id = inote.property_id
            WHERE p.decision_status = 'consider'
            ORDER BY p.price_total ASC;
        """)

        # 4. View: v_properties_with_earthquake_check (物件與921名冊避雷對照表)
        cur.execute("""
            CREATE OR REPLACE VIEW v_properties_with_earthquake_check AS
            SELECT 
                p.code,
                p.community_name,
                p.district,
                p.decision_status,
                p.price_total,
                p.age,
                eq.damage_level,
                eq.legal_status,
                eq.repair_status,
                eq.notes AS earthquake_notes
            FROM properties p
            LEFT JOIN community_earthquake_records eq 
                ON p.community_name ILIKE '%' || eq.original_community_name || '%'
                OR eq.original_community_name ILIKE '%' || p.community_name || '%'
            ORDER BY 
                CASE WHEN eq.damage_level IS NOT NULL THEN 1 ELSE 2 END,
                p.community_name;
        """)

        # 5. View: v_property_image_gallery (圖片畫廊檢視表)
        cur.execute("""
            CREATE OR REPLACE VIEW v_property_image_gallery AS
            SELECT 
                p.code,
                p.community_name,
                p.decision_status,
                pi.image_category,
                pi.file_name,
                pi.file_relpath,
                ROUND(pi.file_size / 1024.0, 1) AS size_kb,
                pi.original_url
            FROM property_images pi
            JOIN properties p ON pi.property_id = p.id
            ORDER BY p.code, pi.sort_order;
        """)

        print("All SQL Views successfully created!")
    conn.close()

if __name__ == '__main__':
    create_views()

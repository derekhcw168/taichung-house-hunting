import psycopg
import re, os, sys

sys.stdout.reconfigure(encoding='utf-8')

def import_preferences_and_notes():
    conn = psycopg.connect("host=localhost port=5432 user=postgres password=postgres dbname=taichung_house", autocommit=True)
    with conn.cursor() as cur:
        # 1. buyer_preferences
        print("Importing buyer preferences...")
        cur.execute("TRUNCATE TABLE buyer_preferences RESTART IDENTITY CASCADE;")
        preferences = [
            ("必要條件", "開價1000萬之內", True, 1, "總價 <= 1000 萬元"),
            ("必要條件", "電梯大樓，不要公寓", True, 2, "建築型態為電梯大樓"),
            ("必要條件", "兩房以上", True, 3, "格局房數 >= 2 房"),
            ("必要條件", "兩衛", True, 4, "衛浴數量 >= 2 衛"),
            ("必要條件", "有工作陽台", True, 5, "需具備工作/後陽台"),
            ("必要條件", "不要二樓和頂樓", True, 6, "所在樓層 != 2 且 所在樓層 != 總樓層"),
            ("加分條件", "一個機車停車格", False, 1, "社區有分配或好租機車位"),
            ("加分條件", "三房", False, 2, "格局房數 >= 3 房"),
            ("加分條件", "汽車停車位", False, 3, "附坡道平面或機械車位 (平車最優)"),
            ("加分條件", "中高樓層", False, 4, "所在樓層在總樓層之一半以上"),
            ("加分條件", "無障礙空間", False, 5, "社區進出平順有無障礙坡道與電梯直達"),
            ("加分條件", "主+陽坪數 > 20坪", False, 6, "主建物 + 附屬建物室內空間 > 20 坪"),
            ("區域偏好", "第1優先：南屯區 (74快速道路以內)", False, 1, "行政區為南屯區且在74內"),
            ("區域偏好", "第2優先：南區", False, 2, "行政區為南區"),
            ("區域偏好", "第3優先：西區", False, 3, "行政區為西區")
        ]
        for cat, name, mand, prio, cond in preferences:
            cur.execute("""
                INSERT INTO buyer_preferences (category, item_name, is_mandatory, priority, condition_detail)
                VALUES (%s, %s, %s, %s, %s);
            """, (cat, name, mand, prio, cond))
        print(f"Inserted {len(preferences)} buyer preferences.")

        # 2. inspection_notes
        print("Importing inspection notes...")
        cur.execute("TRUNCATE TABLE inspection_notes RESTART IDENTITY CASCADE;")
        
        # 信義房屋帶看物件
        sinyi_notes = [
            ("皇后大道", "信義房屋", "2026-10-04", "10:40", "不列入考慮", "", "黎明路生活圈，已看過淘汰", None, None, None, "https://maps.app.goo.gl/qXaLyUB8rdzujHWE9", "7F / 31.72坪，總價998萬"),
            ("尚品大昌街華廈", "信義房屋", "2026-10-04", "11:15", "不列入考慮", "", "公益商圈，室內偏小淘汰", None, None, None, "https://maps.app.goo.gl/TZJK62goor4f64wD7", "5F / 27.11坪，總價958萬"),
            ("大墩人文", "信義房屋", "2026-10-04", None, "列入考慮", "緊鄰南屯路與黎明路商圈、大墩明星學區、近捷運南屯站。開價降至950萬，符合歷史高點成交區間。921曾列黃單但已修繕合格解除列管銷案。", "33年國宅老屋，需整修水電浴室廚具", 880.0, 915.0, "基本預算50.9萬；含兩間浴室重做約72.9萬；含氣密窗約82.4萬", "https://maps.app.goo.gl/ZHpsvCGCH7XoPaGy9", "11F / 31.54坪，信義案號1975QD"),
            ("比佛利山莊", "信義房屋", "2026-10-04", None, "不列入考慮", "", "忠明南路近中山醫，已看過淘汰", None, None, None, "https://maps.app.goo.gl/SGKuWZKwP8384boQ9", "7F / 38.02坪，總價898萬"),
            ("美和街56號華廈", "信義房屋", "2026-10-04", None, "不列入考慮", "", "低總價明亮三房車位，已看過淘汰", None, None, None, "https://maps.app.goo.gl/E2GSgV3npqt9z4bx8", "6F / 30.49坪，總價860萬"),
            ("工學天下", "信義房屋", "2026-10-04", None, "列入考慮", "15樓次頂樓三房，高樓採光視野極佳，總登記坪數40.10坪，總價928萬單價約23.14萬親民", "永義帶看其他低樓層戶已淘汰，此特定15樓戶列入考慮", 850.0, 900.0, None, "https://maps.app.goo.gl/J1ZFGVCt6d3wqcCZ6", "15F / 40.10坪，信義案號4262WL"),
            ("佳麗堡大廈", "信義房屋", "2026-09-28", "14:30", "不列入考慮", "", "陝西東五街立人國中旁，已看過淘汰", None, None, None, "https://maps.app.goo.gl/yRCntfHP54mV8Y53A?g_st=ic", "3F / 31.76坪，總價798萬"),
            ("名流園邸", "信義房屋", "2026-09-28", None, "不列入考慮", "", "青島西街近捷運，已看過淘汰", None, None, None, "https://maps.app.goo.gl/rrX9Z7EFsRjzAfYdA?g_st=ic", "6F / 32.63坪，總價893萬"),
            ("金馬雙星上品座", "信義房屋", "2026-09-28", None, "不列入考慮", "", "五權南路國圖館生活圈，已看過淘汰", None, None, None, "https://maps.app.goo.gl/23Y1cPgxLNoHmN3cA?g_st=ic", "7F / 34.70坪，總價798萬"),
        ]

        # 永義房屋帶看物件
        yungyi_notes = [
            ("台中珍愛", "永義房屋", "2026-09-20", "14:40", "不列入考慮", "", "已看過，不符合需求淘汰", None, None, None, "https://maps.app.goo.gl/7bAejMHjZrsTw2hs8?g_st=ic", ""),
            ("蘇活大街", "永義房屋", "2026-09-20", "15:15", "不列入考慮", "", "已看過，不符合需求淘汰", None, None, None, "https://maps.app.goo.gl/KbCo6YSQBADXE9K37?g_st=ic", ""),
            ("中友生活家", "永義房屋", "2026-09-20", "15:55", "不列入考慮", "", "已看過，不符合需求淘汰", None, None, None, "https://maps.app.goo.gl/eLupL5KZ6V3hjzrW9?g_st=ic", ""),
            ("向上園邸", "永義房屋", "2026-09-20", "16:30", "不列入考慮", "", "已看過，不符合需求淘汰", None, None, None, "https://maps.app.goo.gl/hdiaxZW5JP2RsucB8?g_st=ic", ""),
            ("佛羅里達社區", "永義房屋", "2026-09-20", "17:15", "列入考慮", "廣三佛羅里達，忠明南路園道第一排，高樓層採光空間佳，近健康公園與商圈", "需向仲介調閱不動產說明書確認漏水及公設維護狀況", None, None, None, "https://maps.app.goo.gl/uVYJkEEghSB7kkCD9?g_st=ic", "廣三佛羅里達，案號YE0157095"),
            ("劃江山", "永義房屋", "2026-10-03", "14:00", "不列入考慮", "", "已看過，不符合需求淘汰", None, None, None, "https://maps.app.goo.gl/JgurdKbLXUKdQRcV6?g_st=ic", ""),
            ("工學天下", "永義房屋", "2026-10-03", "14:40", "不列入考慮", "", "永義帶看其他戶別不符需求淘汰 (注意信義帶看之15樓戶列入考慮)", None, None, None, "https://maps.app.goo.gl/ZT9teVqixxjccFG38?g_st=ic", ""),
            ("長億城陽光區", "永義房屋", "2026-10-03", "15:15", "列入考慮（注意921）", "工學商圈生活機能極強，大造鎮格局好", "曾名列921半倒受災名冊經碳纖維鋼板補強解管，轉手有心理抗性", None, None, None, "https://maps.app.goo.gl/n7LBmjanKizkiQKv5?g_st=ic", "帶看兩戶，案號YE0019696"),
            ("大東家南園", "永義房屋", "2026-10-03", "16:30", "列入考慮", "興大商圈學府路文教生活機能強，大東家建設品質穩定，未列入921重大震損", "需確認水電管線翻修情況與機械車位進出動線", None, None, None, "https://maps.app.goo.gl/F1ULWW6cjdVgw1vm9?g_st=ic", "高樓大三房車位"),
        ]

        all_notes = sinyi_notes + yungyi_notes
        for comm, realtor, vdate, vtime, status, pros, cons, low, high, reno, gmap, notes in all_notes:
            cur.execute("""
                INSERT INTO inspection_notes (
                    community_name, realtor_company, visit_date, visit_time, status_conclusion,
                    pros, cons, suggested_offer_low, suggested_offer_high, renovation_estimate,
                    google_maps_url, notes
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
            """, (comm, realtor, vdate, vtime, status, pros, cons, low, high, reno, gmap, notes))
            
            # Also ensure community exists in communities table
            cur.execute("""
                INSERT INTO communities (name)
                VALUES (%s)
                ON CONFLICT (name) DO NOTHING;
            """, (comm,))

        print(f"Inserted {len(all_notes)} field inspection notes!")
    conn.close()

if __name__ == '__main__':
    import_preferences_and_notes()

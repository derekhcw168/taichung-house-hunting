# -*- coding: utf-8 -*-
"""
scripts/generate_dadun_renwen_negotiation_presentation.py
-------------------------------------------------------
生成《大墩人文11樓_屋主見面議價簡報.pptx》
高品質 16:9 寬螢幕商業簡報（13.333 x 7.5 英吋）

核心目標：
專為今晚與大墩人文 11 樓（南屯路二段 207 之 6 號 11 樓之 2）屋主見面議價設計。
完整呈現：
1. 封面與執行摘要：自住家庭誠意前來、開誠佈公、數據支撐、理性協商
2. 大墩人文物件現況剖析：基本規格、優點肯定（起家厝善緣傳承、實坪扎實、通風格局）、缺點與物理折舊硬傷
3. 信義不動產說明書深度查核揭露：雙衛外牆滲水壁癌、窗框漏水、34年水電未換、公設未獨立登記、陽台既存違建、物業協調中
4. 賣方底牌與獲利解碼：6人公同共有繼承渴望變現、零房貸無抵押純拿現金、土增稅僅3,293元微乎其微、空屋每月乾耗成本
5. 老屋安全基礎翻新預算扣抵：盛祥統包工程 8 大工項純基礎工程報價 120 萬逐項解剖，非惡意殺價，而是安全剛性開銷
6. 台中市南屯區 30 年以上老牌大型社區實價登錄歷史大盤走勢：75棟社區、6,473筆交易，大盤量縮 79.7%，價格反轉確立
7. 南屯 10 大指標老牌名宅實登全線回檔鐵證：嶺東率先破 2 字頭，五期/黎明指標案齊跌 20%~38%
8. 大墩人文歷史成交價量深度剖析：長年均價 19~20 萬，近期 2 樓成交掉至 25.5 萬，開價 30 萬踩在歷史天花板
9. 買方強力替代方案（BATNA）震撼施壓：南區「靜園」美裝潢水電全換免花 120 萬可直入且附車位，大墩不談買方立即轉移
10. 全套買方議價與降價說帖策略：50條全攻略精華、金融管制與老屋限貸年限縮短月付沉重
11. 客觀行情與修繕成本精算推導：850萬總預算 - 120萬必要修繕 = 購買極限 730 萬（710 萬起手，保留 20 萬撮合至 730 萬）
12. 簽約特約底線與善意成全：共有人出席授權保證、10萬真金白銀斡旋支票促成簽字、承諾用心善待起家厝
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# 色彩配置
NAVY = RGBColor(15, 32, 67)          # #0F2043 封面主底色
PRIMARY_BLUE = RGBColor(27, 54, 93)  # #1B365D 標題深藍
SLATE = RGBColor(44, 94, 138)        # #2C5E8A 次藍
LIGHT_BG = RGBColor(248, 250, 252)   # #F8FAFC 投影片背景
WHITE = RGBColor(255, 255, 255)
CARD_BG = RGBColor(255, 255, 255)
BORDER_GRAY = RGBColor(226, 232, 240)
ROW_ALT = RGBColor(248, 250, 252)
GOLD = RGBColor(217, 140, 20)        # #D98C14 強調金
RED = RGBColor(192, 57, 43)          # #C0392B 警示紅
GREEN = RGBColor(39, 174, 96)        # #27AE60 支撐綠 / 成功綠
TEXT_DARK = RGBColor(30, 41, 59)     # #1E293B
TEXT_MUTED = RGBColor(100, 116, 139) # #64748B

base_dir = r"d:\GoogleDrive\01_個人_家庭與生活\02_家庭_安莉的資料\小寶的買房子資料"
target_dir = os.path.join(base_dir, "大墩人文")
nantun_chart_dir = os.path.join(base_dir, "南屯區30年以上大型社區走勢圖")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

def set_shape_flat(shape, fill_color, border_color=None, border_width=1):
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(border_width)
    else:
        shape.line.fill.background()

def set_cell_bg(cell, rgb_color):
    cell.fill.solid()
    cell.fill.fore_color.rgb = rgb_color

def add_header(slide, title_text, category="大墩人文 11 樓・屋主見面議價專用溝通簡報"):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, LIGHT_BG)

    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    set_shape_flat(top_bar, PRIMARY_BLUE)

    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.32), Inches(6.5), Inches(0.32))
    set_shape_flat(pill, RGBColor(234, 242, 250), SLATE, 1)
    tf_p = pill.text_frame
    tf_p.word_wrap = True
    tf_p.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_p = tf_p.paragraphs[0]
    p_p.text = category
    p_p.font.name = "微軟正黑體"
    p_p.font.size = Pt(10)
    p_p.font.bold = True
    p_p.font.color.rgb = SLATE

    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(11.733), Inches(0.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "微軟正黑體"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = PRIMARY_BLUE

def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=BORDER_GRAY):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    set_shape_flat(card, bg_color, border_color, 1)
    return card

# ==========================================
# SLIDE 1: 封面
# ==========================================
s1 = prs.slides.add_slide(blank_layout)
bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
set_shape_flat(bg1, NAVY)

dec = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.8), Inches(0.15), Inches(3.2))
set_shape_flat(dec, GOLD)

tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(1.6), Inches(11.3), Inches(3.6))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1 = tf1.paragraphs[0]
p1.text = "大墩人文 11 樓之 2（專任捷運大三房）"
p1.font.name = "微軟正黑體"
p1.font.size = Pt(22)
p1.font.bold = True
p1.font.color.rgb = GOLD

p2 = tf1.add_paragraph()
p2.text = "屋主見面議價與善意溝通簡報"
p2.font.name = "微軟正黑體"
p2.font.size = Pt(36)
p2.font.bold = True
p2.font.color.rgb = WHITE
p2.space_before = Pt(8)

p3 = tf1.add_paragraph()
p3.text = "【物件現況優缺評估 ｜ 產調瑕疵與翻修扣抵 ｜ 南屯75大社區實登走勢 ｜ 買方客觀預算與雙贏讓利提案】"
p3.font.name = "微軟正黑體"
p3.font.size = Pt(13)
p3.font.color.rgb = RGBColor(190, 210, 240)
p3.space_before = Pt(12)

card_foot = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.6), Inches(11.733), Inches(1.15))
set_shape_flat(card_foot, RGBColor(22, 45, 85), RGBColor(40, 75, 130), 1)
tf_foot = card_foot.text_frame
tf_foot.vertical_anchor = MSO_ANCHOR.MIDDLE
p_foot = tf_foot.paragraphs[0]
p_foot.text = (
    "報告對象：大墩人文全體共有人長輩 ｜ 報告人：誠意自住家庭買方 ｜ 標的：台中市南屯區南屯路二段207之6號11樓之2\n"
    "核心訴求：客觀數據攤平討論、全套老屋安全翻修扣除、珍惜善待起家厝善緣、促成實質變現落袋為安"
)
p_foot.font.name = "微軟正黑體"
p_foot.font.size = Pt(10.5)
p_foot.font.color.rgb = WHITE
p_foot.alignment = PP_ALIGN.CENTER

# ==========================================
# SLIDE 2: 執行摘要：買方核心心態與 4 大協商維度
# ==========================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "執行摘要：買方不卑不亢、理性出價的 4 大協商維度")

exec_cards = [
    ("❤️ 肯定起家厝與真誠自住", "我們曾在此生活圈租屋，對大墩人文社區有深刻感情。我們是紮實的純自住家庭，非投資炒作客，承諾購入後投入上百萬細心翻修水電與廚衛，好好延續長輩起家厝的溫暖生命力！", GREEN, Inches(0.8)),
    ("🛠️ 120 萬老屋翻新剛性支出", "34 年老屋水電管線暗管未換、電箱僅 70A、雙衛浴外牆漏水壁癌（信義產調載明）。盛祥統包報價 8 大純基礎安全工項高達 120 萬，這筆保命支出必須在購屋總價如實扣除。", RED, Inches(3.8)),
    ("📉 南屯實登大盤量縮 79.7%", "南屯 75 棟老牌社區年成交量由 694 件雪崩至 141 件！嶺東率先破 2 字頭，五期指標案接連回檔。大墩人文長年行情僅 19~20 萬，近期 2 樓跌至 25.5 萬，950 萬開價已在歷史頂峰。", GOLD, Inches(6.8)),
    ("🎯 850 萬家庭預算上限封頂", "自備款與家庭資金上限 850 萬，扣除 120 萬基礎工程款，購屋極限天花板鎖定 730 萬。隨身備妥 10 萬真金白銀斡旋支票，價格合意當場簽約，協助長輩 6 人全額純現金落袋為安！", PRIMARY_BLUE, Inches(9.8))
]

for title, desc, col, left in exec_cards:
    add_card(s2, left, Inches(1.5), Inches(2.75), Inches(5.3))
    tb = s2.shapes.add_textbox(left + Inches(0.2), Inches(1.7), Inches(2.35), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(13.5)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 3: 物件現況全面解剖（優點肯定 vs 物理硬傷）
# ==========================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "大墩人文 11 樓現況剖析：優點充分肯定，老舊瑕疵客觀面對")

s3_boxes = [
    ("💎 買方高度認同之物件優點（誠意購買基石）", [
        ("室內專有實坪扎實", "登記 31.54 坪全部為主建物(30.07坪)加陽台(1.47坪)，室內大三房開闊好住，無現代新大樓35%灌水虛坪。"),
        ("南屯成熟生活圈地段", "近萬和宮文化圈、永春國小、南屯捷運站，生活採買、就學便利，是買方非常喜歡且具感情的生活微環境。"),
        ("11 樓高樓層採光通風", "次頂樓高視野，邊間雙面採光，白日自然光充足，空氣對流佳，住起來舒服自在。"),
        ("自住起家厝人文底蘊", "長輩惜售傳承，室內格局方正、格局無穿堂等重大風水瑕疵，是自住家庭安居樂業的理想骨架。")
    ], GREEN, Inches(0.8)),
    ("⚠️ 不可忽視之屋況與公設抗性（議價折讓基礎）", [
        ("34 年高屋齡物理老化", "民國 81 年 11 月完工，屋齡已近 34 年，室內水電管線、泥作防水皆達物理壽命終點，必須大幅翻新。"),
        ("無獨立產權車位硬傷", "地下1層使照登記為全棟防空避難兼停車，每年需碰運氣抽籤；南屯三房無車位轉手抗性極大，市價折減150~200萬。"),
        ("高樓層迎風與水壓考量", "11 樓老舊鋁窗漏風漏雨，且高樓層管線水壓偏小，需另外加裝加壓馬達與全面更換氣密隔音窗。"),
        ("物業協調中與基金未明", "信義產調載明物業公司協調中、住戶規約未附、公共基金未明，未來34年老電梯更換每戶面臨集資攤派風險。")
    ], RED, Inches(6.8))
]

for title, items, col, left in s3_boxes:
    add_card(s3, left, Inches(1.5), Inches(5.7), Inches(5.3))
    tb = s3.shapes.add_textbox(left + Inches(0.25), Inches(1.7), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(14)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    for sub, desc in items:
        p_sub = tf.add_paragraph()
        p_sub.text = f"• {sub}："
        p_sub.font.name = "微軟正黑體"
        p_sub.font.size = Pt(10.5)
        p_sub.font.bold = True
        p_sub.font.color.rgb = col
        p_sub.space_before = Pt(8)
        
        p_desc = tf.add_paragraph()
        p_desc.text = desc
        p_desc.font.name = "微軟正黑體"
        p_desc.font.size = Pt(9.5)
        p_desc.font.color.rgb = TEXT_DARK
        p_desc.line_spacing = 1.2

# ==========================================
# SLIDE 4: 信義不動產說明書深度查核揭露（白紙黑字鐵證）
# ==========================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "信義產調報告深度查核：不動產說明書白紙黑字登載的 5 大瑕疵鐵證")

table_s4 = s4.shapes.add_table(6, 4, Inches(0.8), Inches(1.5), Inches(11.733), Inches(4.5))
tbl4 = table_s4.table
tbl4.columns[0].width = Inches(2.2)
tbl4.columns[1].width = Inches(3.0)
tbl4.columns[2].width = Inches(3.2)
tbl4.columns[3].width = Inches(3.333)

headers_s4 = ["說明書查核項目", "產調頁碼與官方登載內容", "法律與現況解讀", "對屋主議價之實質訴求"]
for idx, h in enumerate(headers_s4):
    cell = tbl4.cell(0, idx)
    set_cell_bg(cell, PRIMARY_BLUE)
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.name = "微軟正黑體"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

rows_s4 = [
    ("滲漏水與壁癌瑕疵", "說明書第4頁第一類、第17頁第29項：白紙黑字載明「廁所外牆、窗框有滲漏水、壁癌」，近半年未修。", "雙衛浴防水層已老化破裂失效，窗框填縫硬化，非表面刷漆可解決，需打除見底重做。", "兩間衛浴泥作防水22萬＋氣密窗13萬，共35萬剛性支出必須全額自房價折抵。"),
    ("34年水電管線未換", "說明書第16頁第19項：明確勾選「水管、電管線於產權持有期間未曾更新」，配電箱僅70A。", "老舊電線氧化走火風險高，生鏽鐵管隨時爆管，完全無法負擔現代多台冷氣與廚房高功率家電。", "全室太平洋電線抽換、配電箱加大至100A與不銹鋼雙壓接水管需21.9萬，安全必做款。"),
    ("陽台加窗屬既存違建", "說明書第4頁第三類、第17頁第33項：明載「陽台加窗屬未經合法程序之違章建築」。", "若遭檢舉或市府公安查報面臨即報即拆且無補償，信義合約要求買方自行承擔被拆除風險。", "買方承擔違法被拆風險與拆除還原費用，開價應給予合理產權瑕疵折讓。"),
    ("物業協調中／規約未附", "說明書第5頁第4~5項：管委會物業公司「在協調中」，住戶規約「目前尚未檢附」。", "大樓物業管理處於協調動盪期，公共門禁安防有真空隱憂，公設權利義務一團黑盒子。", "管理不穩定社區中古行情本應折減5%~10%，買方承擔未知管理風險。"),
    ("地下室共有無獨立車位", "說明書第27~28頁使照存根：地下1層登記為防空避難兼停車，無獨立車位產權。", "每年需碰運氣抽籤，南屯路大墩路車位極難租，在外租車位20年額外支出70~80萬。", "無產權車位在南屯三房市場行情折減150~200萬，不可硬套附車位之單價。")
]

for r_idx, r_data in enumerate(rows_s4, start=1):
    for c_idx, val in enumerate(r_data):
        cell = tbl4.cell(r_idx, c_idx)
        set_cell_bg(cell, ROW_ALT if r_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.name = "微軟正黑體"
        p.font.size = Pt(9.5)
        p.font.color.rgb = TEXT_DARK
        if c_idx in [0, 3]:
            p.font.bold = True
        if c_idx == 3:
            p.font.color.rgb = RED

card_s4_tip = add_card(s4, Inches(0.8), Inches(6.1), Inches(11.733), Inches(1.0), bg_color=RGBColor(254, 242, 242), border_color=RED)
tb_s4_tip = s4.shapes.add_textbox(Inches(1.0), Inches(6.15), Inches(11.333), Inches(0.9))
tf_s4_tip = tb_s4_tip.text_frame
tf_s4_tip.word_wrap = True
p_tip1 = tf_s4_tip.paragraphs[0]
p_tip1.text = "🎯 買方核心陳述：這份信義產調報告是經紀人員工親自勘驗簽認的法定文件，所有瑕疵白紙黑字清清楚楚。"
p_tip1.font.name = "微軟正黑體"
p_tip1.font.size = Pt(11)
p_tip1.font.bold = True
p_tip1.font.color.rgb = RED

p_tip2 = tf_s4_tip.add_paragraph()
p_tip2.text = "買方不是看屋後空口嫌東嫌西，而是完全根據說明書揭露的客觀事實，計算交屋後必須承擔的真實工程與法律負擔，請長輩務必在總價上實質折抵。"
p_tip2.font.name = "微軟正黑體"
p_tip2.font.size = Pt(10)
p_tip2.font.color.rgb = TEXT_DARK
p_tip2.space_before = Pt(2)

# ==========================================
# SLIDE 5: 賣方底牌與財務現況分析（落袋為安、無稅務負擔）
# ==========================================
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "賣方底牌與利益解碼：零房貸純拿現金、低土增稅，變現落袋為安最明智")

s5_cards = [
    ("1. 6人公同共有繼承，大家最希望變現分產", "• 本戶於 113 年 7 月發生繼承，登記為 6 人公同共有。\n• 產調揭露其中 5 位長輩甚至連權狀費用都未領，產權手續繁複。\n• 房子空置未住，6 位共有人大家最期盼的就是順利賣掉、公平變現分配遺產現金，落袋為安最圓滿！", Inches(0.8), PRIMARY_BLUE),
    ("2. 謄本完全無抵押權，賣多少全額純拿現金！", "• 地政謄本他項權利部白紙黑字「無他項權利登記」，賣方完全零房貸壓力！\n• 成交價無論 710 萬還是 730 萬，全額都是實質純現金入袋！\n• 屋主沒有任何高額貸款本利和或賠售套牢問題，爽快讓利成就一段買賣善緣！", Inches(4.8), GREEN),
    ("3. 土地增值稅僅 3,293 元，微乎其微零負擔！", "• 信義說明書第 9 頁土地增值稅概算表：前次移轉為113年7月，一般稅率合計僅需繳納 3,293 元（自用更僅 1,646 元）！\n• 繼承取得完全沒有龐大土增稅成本，賣方切勿拿稅金當阻擋議價的藉口！", Inches(8.8), GOLD)
]

for title, desc, left, col in s5_cards:
    add_card(s5, left, Inches(1.5), Inches(3.7), Inches(3.3), border_color=col)
    tb = s5.shapes.add_textbox(left + Inches(0.2), Inches(1.65), Inches(3.3), Inches(3.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(8)

card_s5_bot = add_card(s5, Inches(0.8), Inches(5.0), Inches(11.733), Inches(2.1), bg_color=RGBColor(240, 249, 255), border_color=SLATE)
tb_s5_bot = s5.shapes.add_textbox(Inches(1.0), Inches(5.1), Inches(11.333), Inches(1.9))
tf_s5_bot = tb_s5_bot.text_frame
tf_s5_bot.word_wrap = True

p_bot_t = tf_s5_bot.paragraphs[0]
p_bot_t.text = "💡 屋主繼續空置持有的「隱形失血與跌價風險」試算："
p_bot_t.font.name = "微軟正黑體"
p_bot_t.font.size = Pt(13)
p_bot_t.font.bold = True
p_bot_t.font.color.rgb = PRIMARY_BLUE

bot_points = [
    ("固定支出乾耗", "空屋每月管理費 1,615 元＋公設電費水費約 2,000 元，多拖半年就平白燒掉 1.5~2 萬元現金，每年還要繳納房屋稅與地價稅。"),
    ("大環境量縮修正", "央行第七波信用管制下，南屯老屋成交量暴跌 8 成，價格緩跌確立。房子繼續掛售只會變滯銷物件，半年後買方出價恐跌破 700 萬。"),
    ("家庭變現落袋為安", "今天只要點頭成交，6 位共有人每人可立即分得約 120 萬現金，把虛擬的不動產轉化為實體資產，既圓滿了長輩遺願，也免去日後管理煩惱！")
]

for sub, desc in bot_points:
    p_bp = tf_s5_bot.add_paragraph()
    p_bp.text = f"• {sub}：{desc}"
    p_bp.font.name = "微軟正黑體"
    p_bp.font.size = Pt(10)
    p_bp.font.color.rgb = TEXT_DARK
    p_bp.space_before = Pt(4)

# ==========================================
# SLIDE 6: 盛祥空間工程 8 大純基礎修繕預算 120 萬逐項解剖
# ==========================================
s6 = prs.slides.add_slide(blank_layout)
add_header(s6, "老屋安全必做工程：盛祥統包工程 120 萬報價單逐項解剖（非奢華裝潢）")

table_s6 = s6.shapes.add_table(10, 5, Inches(0.8), Inches(1.4), Inches(11.733), Inches(4.5))
tbl6 = table_s6.table
tbl6.columns[0].width = Inches(1.0)
tbl6.columns[1].width = Inches(2.2)
tbl6.columns[2].width = Inches(4.733)
tbl6.columns[3].width = Inches(1.8)
tbl6.columns[4].width = Inches(2.0)

headers_s6 = ["項次", "工程類別", "主要施作內容與規格摘要", "預估金額 (NTD)", "佔比與剛性說明"]
for idx, h in enumerate(headers_s6):
    cell = tbl6.cell(0, idx)
    set_cell_bg(cell, PRIMARY_BLUE)
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.name = "微軟正黑體"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

rows_s6 = [
    ("一", "保護與拆除清運", "公共梯廳防撞、拆除舊發霉木作櫃體天花板、雙衛廚具打到底、4車合法清運", "$162,000", "13.4% (前屋主裝潢負資產)"),
    ("二", "水電管線與電力系統", "全室抽換太平洋一級銅線、總電箱加大匯流排(75~100A)、SUS304不銹鋼雙壓接水管", "$219,000", "18.2% (居家用電用水保命款)"),
    ("三", "泥作修補、壁癌防水", "剔除粉化壁癌抗鹼打底、雙衛浴彈性水泥防水三道(樑下試水48hr)、全室地壁磚", "$220,000", "18.3% (信義產調瑕疵根治)"),
    ("四", "鋁門窗與氣密隔音", "11樓高樓客廳大落地氣密窗、三房5+5膠合強化隔音窗、三合一通風門、窗框塞水路", "$130,000", "10.8% (高樓防颱抗漏風漏水)"),
    ("五", "雙衛浴設備與五金", "TOTO/Derek省水馬桶×2、防潮發泡板浴櫃面盆×2、主衛乾濕分離拉門、多功能暖風機", "$106,000", "8.8% (基本衛生與長輩防滑)"),
    ("六", "廚房系統廚具三機", "240cm一字型防潮廚具、耐磨石英石檯面、櫻花/林內防乾燒瓦斯爐與強力油煙機", "$137,000", "11.4% (34年老瓦斯管防洩漏)"),
    ("七", "全室油漆粉刷工程", "全室裂縫AB膠抗裂填縫、全批土兩道研磨細平、得利/虹牌健康無毒低甲醛乳膠漆", "$93,000", "7.7% (921細紋整平粉刷)"),
    ("八", "門扇地坪局部包管", "臥房實木門×3、衛浴防潮門×2、管線走頂局部平釘包管、全室直鋪SPC石塑地板", "$138,000", "11.5% (老地磚膨拱整平隔音)"),
    ("合計", "全室純基礎工程總額", "8 大工種老屋保命純基礎翻新總工程款（盛祥工程實打實估價單）", "$1,205,000", "100.0% (約 120 萬元整)")
]

for r_idx, r_data in enumerate(rows_s6, start=1):
    for c_idx, val in enumerate(r_data):
        cell = tbl6.cell(r_idx, c_idx)
        is_total = (r_idx == 9)
        set_cell_bg(cell, RGBColor(254, 242, 242) if is_total else (ROW_ALT if r_idx % 2 == 1 else WHITE))
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.name = "微軟正黑體"
        p.font.size = Pt(9.5) if not is_total else Pt(10.5)
        p.font.color.rgb = RED if is_total else TEXT_DARK
        if c_idx in [0, 1, 3] or is_total:
            p.font.bold = True
        if c_idx in [0, 3]:
            p.alignment = PP_ALIGN.CENTER

card_s6_foot = add_card(s6, Inches(0.8), Inches(6.05), Inches(11.733), Inches(1.1), bg_color=RGBColor(240, 249, 255), border_color=PRIMARY_BLUE)
tb_s6_f = s6.shapes.add_textbox(Inches(1.0), Inches(6.1), Inches(11.333), Inches(1.0))
tf_s6_f = tb_s6_f.text_frame
tf_s6_f.word_wrap = True

p_f1 = tf_s6_f.paragraphs[0]
p_f1.text = "🎯 買方出價公式精算：家庭總預算 850 萬 － 老屋純基礎工程 120 萬 ＝ 購屋總價極限天花板 730 萬"
p_f1.font.name = "微軟正黑體"
p_f1.font.size = Pt(11.5)
p_f1.font.bold = True
p_f1.font.color.rgb = PRIMARY_BLUE

p_f2 = tf_s6_f.add_paragraph()
p_f2.text = "這份估價單完全沒有水晶吊燈、豪奢木作壁板或頂級家電，全是為了能安心住進去的水電防水安全工程。每一分錢都省不掉，買方必須拿現金或修繕貸款支付，因此必須從屋主售價中完全折減！"
p_f2.font.name = "微軟正黑體"
p_f2.font.size = Pt(9.8)
p_f2.font.color.rgb = TEXT_DARK
p_f2.space_before = Pt(3)

# ==========================================
# SLIDE 7: 南屯區 30 年以上老牌大社區實登大盤走勢（圖表分析）
# ==========================================
s7 = prs.slides.add_slide(blank_layout)
add_header(s7, "南屯大盤實價登錄走勢：75 棟社區量縮 79.7%，價格全面反轉回修")

macro_img = os.path.join(nantun_chart_dir, "00_台中市南屯區30年以上大型社區_大盤匯總走勢圖.png")
if os.path.exists(macro_img):
    s7.shapes.add_picture(macro_img, Inches(0.8), Inches(1.5), Inches(6.8), Inches(5.3))

add_card(s7, Inches(7.8), Inches(1.5), Inches(4.733), Inches(5.3))
tb_s7 = s7.shapes.add_textbox(Inches(8.0), Inches(1.65), Inches(4.333), Inches(4.9))
tf_s7 = tb_s7.text_frame
tf_s7.word_wrap = True

p_s7_t = tf_s7.paragraphs[0]
p_s7_t.text = "📊 南屯老牌社區大盤 4 大權威結論"
p_s7_t.font.name = "微軟正黑體"
p_s7_t.font.size = Pt(14)
p_s7_t.font.bold = True
p_s7_t.font.color.rgb = PRIMARY_BLUE

s7_points = [
    ("樣本涵蓋全區最完整", "普查南屯全區 75 棟 30 年以上大型社區（14,671 戶、平均屋齡 32.4 年），純住宅實登有效交易 5,437 筆，經 591/樂居/CTOPMap 交叉驗證。"),
    ("成交量雪崩 79.7%", "年成交量自高峰 694 件一路崩跌至 115 年僅 141 件（量縮近 8 成）！央行第七波信用管制後，買氣瞬間急凍，『量先價行』確立價格走跌。"),
    ("歷史頂峰見頂下修", "113 年大盤中位數衝上 29.1 萬/坪歷史極值，114~115 年已反轉回落至 27.8 萬/坪，嶺東精科生活圈更是率先摜破 2 字頭！"),
    ("計量模型三年預估", "計量經濟六模組預估：大盤 116 年探至 24.4 萬 ➜ 117 年 21.3 萬 ➜ 118 年 20.6 萬（累計修正 -25.9%），逐步回歸 106~109 年合理價值帶(15.4~19.8萬)！")
]

for title, desc in s7_points:
    p_p = tf_s7.add_paragraph()
    p_p.text = f"• {title}："
    p_p.font.name = "微軟正黑體"
    p_p.font.size = Pt(11)
    p_p.font.bold = True
    p_p.font.color.rgb = RED if "雪崩" in title or "頂峰" in title else SLATE
    p_p.space_before = Pt(8)
    
    p_pd = tf_s7.add_paragraph()
    p_pd.text = desc
    p_pd.font.name = "微軟正黑體"
    p_pd.font.size = Pt(9.5)
    p_pd.font.color.rgb = TEXT_DARK
    p_pd.line_spacing = 1.25

# ==========================================
# SLIDE 8: 南屯 10 大類似老牌指標名宅實登全數反轉向下（鐵證總表）
# ==========================================
s8 = prs.slides.add_slide(blank_layout)
add_header(s8, "南屯指標社區實證：周邊 10 大老牌大案實登全數見頂回跌（鐵證總表）")

table_s8 = s8.shapes.add_table(11, 7, Inches(0.8), Inches(1.5), Inches(11.733), Inches(4.2))
tbl8 = table_s8.table
tbl8.columns[0].width = Inches(2.0)
tbl8.columns[1].width = Inches(1.2)
tbl8.columns[2].width = Inches(1.8)
tbl8.columns[3].width = Inches(1.8)
tbl8.columns[4].width = Inches(1.7)
tbl8.columns[5].width = Inches(1.2)
tbl8.columns[6].width = Inches(2.033)

headers_s8 = ["指標社區建案", "總戶數", "歷史最高單價", "最新成交單價", "116~118預估", "三年修正", "實登走勢特徵與意義"]
for idx, h in enumerate(headers_s8):
    cell = tbl8.cell(0, idx)
    set_cell_bg(cell, PRIMARY_BLUE)
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.name = "微軟正黑體"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

bench_s8 = [
    ("理想貴族 (大墩二街)", "732 戶", "35.5 萬 (113/06)", "27.5 萬 (115/09)", "21.9 ➜ 18.7 萬", "-24.6%", "五期700戶最大案，高檔大修近8萬"),
    ("宏全世界廣場 (文心路)", "436 戶", "37.6 萬 (113/07)", "28.4 萬 (115/07)", "24.1 ➜ 21.1 萬", "-21.6%", "七期南側捷運名廈，高點急凍跌近10萬"),
    ("中友綠園邸 (黎明路)", "431 戶", "39.8 萬 (113/06)", "27.7 萬 (115/04)", "25.3 ➜ 22.2 萬", "-21.0%", "黎明商圈大案，最新成交自近40萬重挫"),
    ("凱撒風景園邸 (忠勇路)", "423 戶", "29.8 萬 (112/07)", "19.6 萬 (115/07)", "21.2 ➜ 18.5 萬", "-21.3%", "🔥 嶺東400戶大案最新摜破 2 字頭！"),
    ("龍觀天下 (大墩十一街)", "420 戶", "32.0 萬 (113/12)", "23.4 萬 (115/07)", "24.4 ➜ 21.3 萬", "-21.4%", "五期知名名宅由32萬重修至23.4萬"),
    ("皇后大道 (黎明路)", "388 戶", "37.6 萬 (114/06)", "27.2 萬 (115/08)", "23.3 ➜ 20.3 萬", "-22.2%", "黎明新村商圈指標案高檔回跌逾10萬"),
    ("黎明大地 (嶺東路)", "376 戶", "32.0 萬 (112/10)", "18.7 萬 (115/07)", "19.0 ➜ 16.5 萬", "-22.2%", "🔥 最新成交跌至18.7萬，大跌逾4成"),
    ("陽光綠意参 (文心路)", "366 戶", "35.8 萬 (113/08)", "24.5 萬 (115/05)", "24.5 ➜ 21.3 萬", "-22.3%", "文心森林公園生活圈大樓跌破25萬"),
    ("黎明春秋 (向上路)", "361 戶", "39.2 萬 (113/08)", "30.8 萬 (115/02)", "27.7 ➜ 24.2 萬", "-21.4%", "向上商圈名宅高點見頂量縮橫盤"),
    ("現代經典 (南屯路二段)", "222 戶", "33.0 萬 (112/11)", "30.2 萬 (115/04)", "27.4 ➜ 24.0 萬", "-20.5%", "同在南屯路生活圈，價格亦緩步向下")
]

for r_idx, r_data in enumerate(bench_s8, start=1):
    for c_idx, val in enumerate(r_data):
        cell = tbl8.cell(r_idx, c_idx)
        set_cell_bg(cell, ROW_ALT if r_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.name = "微軟正黑體"
        p.font.size = Pt(9.5)
        p.font.color.rgb = TEXT_DARK
        if c_idx in [0, 3, 5]:
            p.font.bold = True
        if c_idx == 5:
            p.font.color.rgb = RED

card_s8_foot = add_card(s8, Inches(0.8), Inches(5.85), Inches(11.733), Inches(1.25), bg_color=RGBColor(254, 242, 242), border_color=RED)
tb_s8_f = s8.shapes.add_textbox(Inches(1.0), Inches(5.95), Inches(11.333), Inches(1.05))
tf_s8_f = tb_s8_f.text_frame
tf_s8_f.word_wrap = True

p_s8_1 = tf_s8_f.paragraphs[0]
p_s8_1.text = "🎯 房仲與買方一致客觀研判：全南屯老牌大社區無一倖免！高點已過，量縮回檔是整個區域的大勢所趨。"
p_s8_1.font.name = "微軟正黑體"
p_s8_1.font.size = Pt(11)
p_s8_1.font.bold = True
p_s8_1.font.color.rgb = RED

p_s8_2 = tf_s8_f.add_paragraph()
p_s8_2.text = "南屯區這些戶數更多、甚至附車位的老牌名宅，成交價紛紛自 35~39 萬高點重修至 2 字頭，外圍嶺東更已摜破 20 萬。大墩人文若堅持開價 30 萬以上，完全脫離當前市場大盤走勢，只會陷入乏人問津的滯銷僵局！"
p_s8_2.font.name = "微軟正黑體"
p_s8_2.font.size = Pt(9.8)
p_s8_2.font.color.rgb = TEXT_DARK
p_s8_2.space_before = Pt(3)

# ==========================================
# SLIDE 9: 大墩人文歷史成交價量深度剖析（走勢圖與歷史錨定）
# ==========================================
s9 = prs.slides.add_slide(blank_layout)
add_header(s9, "大墩人文歷史實登全解讀：長年均價20萬，近期2樓成交跌至25.5萬")

dadun_img = os.path.join(nantun_chart_dir, "22_大墩人文_成交行情走勢圖.png")
if os.path.exists(dadun_img):
    s9.shapes.add_picture(dadun_img, Inches(0.8), Inches(1.5), Inches(6.8), Inches(5.3))

add_card(s9, Inches(7.8), Inches(1.5), Inches(4.733), Inches(5.3))
tb_s9 = s9.shapes.add_textbox(Inches(8.0), Inches(1.65), Inches(4.333), Inches(4.9))
tf_s9 = tb_s9.text_frame
tf_s9.word_wrap = True

p_s9_t = tf_s9.paragraphs[0]
p_s9_t.text = "🏢 大墩人文 39 筆實登關鍵事實"
p_s9_t.font.name = "微軟正黑體"
p_s9_t.font.size = Pt(14)
p_s9_t.font.bold = True
p_s9_t.font.color.rgb = PRIMARY_BLUE

dadun_points = [
    ("長年合理行情僅 19~20 萬/坪", "樂居與內政部實登歷史顯示：2015 至 2022 年長達 7 年間，大墩人文常態成交總價多在 600~680 萬，折合單價長期穩定在 19~20 萬/坪。"),
    ("114年10月2樓成交跌至 25.5 萬", "最新實登顯示 207 號之 2 的 2 樓戶，成交總價僅 800 萬、單價 25.5 萬/坪！社區價格回調訊號明確，市場已拒絕買單 30 萬以上高價。"),
    ("開價 950 萬踩在歷史最頂峰", "全社區 34 年歷史最高單價為 30.3 萬/坪。屋主目前開價 950 萬（30.12 萬/坪），依然死守在歷史天花板，在量縮大盤下完全缺乏競爭力。"),
    ("計量模型預估回修至 22.8 萬", "南屯計量模型動態推算：大墩人文三年預期價格將收斂至 116 年 26.0 萬 ➜ 117 年 23.6 萬 ➜ 118 年 22.8 萬（累計修正 -20.6%），向長期支撐帶(17.8~19.2萬)靠攏！")
]

for title, desc in dadun_points:
    p_p = tf_s9.add_paragraph()
    p_p.text = f"• {title}："
    p_p.font.name = "微軟正黑體"
    p_p.font.size = Pt(11)
    p_p.font.bold = True
    p_p.font.color.rgb = RED if "2樓" in title or "天花板" in title else SLATE
    p_p.space_before = Pt(8)
    
    p_pd = tf_s9.add_paragraph()
    p_pd.text = desc
    p_pd.font.name = "微軟正黑體"
    p_pd.font.size = Pt(9.5)
    p_pd.font.color.rgb = TEXT_DARK
    p_pd.line_spacing = 1.25

# ==========================================
# SLIDE 10: 買方強力替代方案（BATNA）：南區「靜園」美裝潢附平車施壓
# ==========================================
s10 = prs.slides.add_slide(blank_layout)
add_header(s10, "買方客觀替代選擇：永義房屋帶看「靜園」美裝潢直入＋附車位強力比價")

table_s10 = s10.shapes.add_table(5, 3, Inches(0.8), Inches(1.5), Inches(11.733), Inches(3.8))
tbl10 = table_s10.table
tbl10.columns[0].width = Inches(2.5)
tbl10.columns[1].width = Inches(4.6)
tbl10.columns[2].width = Inches(4.633)

headers_s10 = ["評估指標", "南區【靜園社區 11 樓】(永義帶看中・替代方案)", "南屯【大墩人文 11 樓】(本案溝通標的)"]
for idx, h in enumerate(headers_s10):
    cell = tbl10.cell(0, idx)
    set_cell_bg(cell, PRIMARY_BLUE)
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.name = "微軟正黑體"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

rows_s10 = [
    ("基礎工程與室內裝修", "原屋主已全室重拉水電、泥作衛浴全新翻修，配備高質感美裝潢，買方完全免掏 120 萬翻修款，拎包即可入住！", "34年未翻新之原始屋況，水電暗管、配電箱、雙衛浴漏水壁癌全需打掉重練，買方須另掏 120 萬現金翻修。"),
    ("停車位條件與產權", "擁有專屬獨立固定平面汽車車位，上下班停車無憂，親友探訪便利。", "無獨立產權車位，每年需碰運氣抽籤；周邊車位一位難求，外租每月需 3,000 元。"),
    ("工作通勤與地緣便利", "買方先生在霧峰上班，南區動線緊鄰台 74 線與中投公路，通勤時間節省 15~20 分鐘。", "買方曾在本生活圈租屋，生活習慣熟悉、有地緣感情（唯一最大情感優勢）。"),
    ("買方當前談判決策", "若大墩人文價格卡死在 800 多萬以上，買方今天見面後將立刻轉向永義房屋洽談「靜園」！", "若屋主長輩願意成全讓利，在扣除 120 萬翻修款後以 710~730 萬成交，買方當場優先簽約！")
]

for r_idx, r_data in enumerate(rows_s10, start=1):
    for c_idx, val in enumerate(r_data):
        cell = tbl10.cell(r_idx, c_idx)
        set_cell_bg(cell, ROW_ALT if r_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.name = "微軟正黑體"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_DARK
        if c_idx == 0:
            p.font.bold = True
        if r_idx == 4 and c_idx == 2:
            p.font.bold = True
            p.font.color.rgb = GREEN

card_s10_warn = add_card(s10, Inches(0.8), Inches(5.6), Inches(11.733), Inches(1.3), bg_color=RGBColor(254, 242, 242), border_color=RED)
tb_s10_w = s10.shapes.add_textbox(Inches(1.0), Inches(5.7), Inches(11.333), Inches(1.1))
tf_s10_w = tb_s10_w.text_frame
tf_s10_w.word_wrap = True

p_w1 = tf_s10_w.paragraphs[0]
p_w1.text = "⚠️ 買方坦率告知：我們對大墩人文有深厚感情，但家庭財務紀律必須理智！"
p_w1.font.name = "微軟正黑體"
p_w1.font.size = Pt(12)
p_w1.font.bold = True
p_w1.font.color.rgb = RED

p_w2 = tf_s10_w.add_paragraph()
p_w2.text = "「靜園」不用花 120 萬翻修、還附平面車位、上班更近。大墩人文如果屋主不願意在總價扣除 120 萬修繕費，買方理智一回神就會直接轉去買「靜園」了。請長輩把握今天當面洽談的良機，不要讓最有誠意的買方流失！"
p_w2.font.name = "微軟正黑體"
p_w2.font.size = Pt(10.5)
p_w2.font.color.rgb = TEXT_DARK
p_w2.space_before = Pt(3)

# ==========================================
# SLIDE 11: 金融風暴與老屋房貸授信緊縮（買方無法追高的客觀障礙）
# ==========================================
s11 = prs.slides.add_slide(blank_layout)
add_header(s11, "金融環境劇變：央行信用管制與銀行老屋限貸，買方絕無追高空間")

s11_cards = [
    ("1. 央行第七波信用管制，市場買氣急凍", "• 央行史上最嚴厲打炒房措施，房市游資全面退潮。\n• 買方態度轉趨極度觀望保守，目前市場上僅剩財務健全的極少數剛性首購族。\n• 失去投資炒作買盤拉抬，開價過高的老屋只會淪為乏人問津的滯銷孤兒！", Inches(0.8), RED),
    ("2. 銀行法 72-2 滿水位，鑑價極端保守", "• 各大公私立銀行不動產放款水位接近警戒線，排隊撥款拉長至 2~3 個月。\n• 銀行鑑價防線僅約 750~800 萬，高出鑑價的落差銀行不借，必須全部拿現金補齊。\n• 買方是一般受薪家庭，根本拿不出額外百萬現金去填補估價落差！", Inches(4.8), GOLD),
    ("3. 34 年老屋年限縮短，月付本利暴增！", "• 銀行遵循「屋齡＋貸款年限 ≤ 50」風控鐵律，34年老屋無法承做30/40年期房貸。\n• 貸款年限被硬生生壓縮至 15~20 年，每月還款本息暴增一萬多元，掐死受薪家庭收支比上限！\n• 總價唯有鎖死在 730 萬以內，買方每月現金流才能安全過關！", Inches(8.8), PRIMARY_BLUE)
]

for title, desc, left, col in s11_cards:
    add_card(s11, left, Inches(1.5), Inches(3.7), Inches(5.3), border_color=col)
    tb = s11.shapes.add_textbox(left + Inches(0.2), Inches(1.7), Inches(3.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 12: 買方 50 條議價全攻略核心精粹（談判說法矩陣）
# ==========================================
s12 = prs.slides.add_slide(blank_layout)
add_header(s12, "買方議價理由 50 條攻略精華：五大核心維度說法矩陣")

table_s12 = s12.shapes.add_table(6, 4, Inches(0.8), Inches(1.5), Inches(11.733), Inches(4.5))
tbl12 = table_s12.table
tbl12.columns[0].width = Inches(2.2)
tbl12.columns[1].width = Inches(2.8)
tbl12.columns[2].width = Inches(3.2)
tbl12.columns[3].width = Inches(3.533)

headers_s12 = ["維度分類", "涵蓋核心條號", "客觀事實與數據支撐", "現場溝通說法與台詞要旨"]
for idx, h in enumerate(headers_s12):
    cell = tbl12.cell(0, idx)
    set_cell_bg(cell, PRIMARY_BLUE)
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.name = "微軟正黑體"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

rows_s12 = [
    ("一、屋況瑕疵與基礎翻修", "BP-01 ~ BP-10 (共10條)", "盛祥工程報價 1,205,000 元、信義產調雙衛漏水壁癌、窗框滲水。", "「水電管線與雙衛防水攸關居家生命安全，120萬純基礎工程一毛錢都省不掉，這筆剛性支出必須全額自開價扣抵。」"),
    ("二、社區公設與管理提撥", "BP-11 ~ BP-18 (共8條)", "34年電梯機組老化、物業管理協調中、無獨立公設登記、無產權車位。", "「34年電梯隨時面臨上百萬更換集資，無獨立車位每年抽籤且南屯轉手抗性大，老社區管理不透明風險應折減房價。」"),
    ("三、實登行情與估價落差", "BP-19 ~ BP-27 (共9條)", "歷史常態 19~20 萬、2樓最新成交 25.5 萬、銀行鑑價上限 750~800 萬。", "「實登最新2樓已跌回25.5萬，開價950萬脫離現況；銀行鑑價落差銀行不借，買方出價緊扣客觀鑑價，絕不盲目追高。」"),
    ("四、競品替代與地段抗性", "BP-28 ~ BP-35 (共8條)", "永義帶看南區「靜園」美裝潢免修繕、捷運步行實測近10分、嫌惡設施。", "「南區靜園水電裝潢全翻新且附車位，大墩人文若不實質讓利，買方今天立刻轉向簽約；買方承擔老屋未來轉手風險。」"),
    ("五、信用管制與心理攻防", "BP-36 ~ BP-52 (共17條)", "6人繼承純拿現金、零房貸無壓力、土增稅僅3,293元、備妥10萬支票。", "「長輩繼承取得零房貸，賣掉全額純拿現金落袋為安，土增稅微乎其微；我們備妥10萬支票隨時履約，請長輩圓滿成全。」")
]

for r_idx, r_data in enumerate(rows_s12, start=1):
    for c_idx, val in enumerate(r_data):
        cell = tbl12.cell(r_idx, c_idx)
        set_cell_bg(cell, ROW_ALT if r_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.name = "微軟正黑體"
        p.font.size = Pt(9.5)
        p.font.color.rgb = TEXT_DARK
        if c_idx in [0, 1]:
            p.font.bold = True
        if c_idx == 3:
            p.font.bold = True
            p.font.color.rgb = PRIMARY_BLUE

card_s12_f = add_card(s12, Inches(0.8), Inches(6.1), Inches(11.733), Inches(0.95), bg_color=RGBColor(240, 249, 255), border_color=PRIMARY_BLUE)
tb_s12_f = s12.shapes.add_textbox(Inches(1.0), Inches(6.15), Inches(11.333), Inches(0.85))
tf_s12_f = tb_s12_f.text_frame
tf_s12_f.word_wrap = True
p_f_t = tf_s12_f.paragraphs[0]
p_f_t.text = "🎯 現場談判準則：不疾不徐、有條不紊，以數據與說明書為盾，以真誠善待起家厝為矛，促成屋主共鳴讓步。"
p_f_t.font.name = "微軟正黑體"
p_f_t.font.size = Pt(11)
p_f_t.font.bold = True
p_f_t.font.color.rgb = PRIMARY_BLUE

# ==========================================
# SLIDE 13: 買方具體出價推導與降價訴求提案（710~730 萬）
# ==========================================
s13 = prs.slides.add_slide(blank_layout)
add_header(s13, "降價訴求推導：基於客觀行情與修繕成本的理性出價方案")

s13_boxes = [
    ("📊 買方財務極限與出價推導數學模型", [
        ("家庭全部購屋總資金極限", "新台幣 850 萬元（含自有自備款與銀行房貸上限，無多餘資金）"),
        ("老屋必做純基礎翻新支出", "扣除新台幣 120 萬元（盛祥統包報價單：水電、泥作、氣密窗、廚衛）"),
        ("大墩人文購買總價極限天花板", "新台幣 730 萬元（折合單價約 23.14 萬/坪，已充分高於歷史長期行情）"),
        ("今晚斡旋起手誠意出價", "新台幣 710 萬元（折合單價約 22.51 萬/坪，保留 20 萬磋商撮合彈性）")
    ], PRIMARY_BLUE, Inches(0.8)),
    ("🤝 屋主降價讓利能獲得的 4 大實質好處", [
        ("純現金落袋為安，零違約風險", "買方財務審核穩健、自備款充裕，保證順利撥貸履約，絕非投機客亂出價。"),
        ("6 人共有人每人現拿 120 萬", "順利處分空置遺產，立即解除 6 人多年分產懸念，長輩人人拿百萬現金。"),
        ("立刻止血空屋折舊與每月水電費", "立即免除每月管理費與水電乾耗，更避開未來南屯老屋續跌 15%~25% 風險。"),
        ("善緣傳承，起家厝交給珍惜它的家庭", "買方承諾投入重金修復水電泥作、長久安居，延續長輩起家厝的生命力！")
    ], GREEN, Inches(6.8))
]

for title, items, col, left in s13_boxes:
    add_card(s13, left, Inches(1.5), Inches(5.7), Inches(5.3))
    tb = s13.shapes.add_textbox(left + Inches(0.25), Inches(1.7), Inches(5.2), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(13.5)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    for sub, desc in items:
        p_sub = tf.add_paragraph()
        p_sub.text = f"• {sub}："
        p_sub.font.name = "微軟正黑體"
        p_sub.font.size = Pt(10.5)
        p_sub.font.bold = True
        p_sub.font.color.rgb = col
        p_sub.space_before = Pt(8)
        
        p_desc = tf.add_paragraph()
        p_desc.text = desc
        p_desc.font.name = "微軟正黑體"
        p_desc.font.size = Pt(9.5)
        p_desc.font.color.rgb = TEXT_DARK
        p_desc.line_spacing = 1.25

# ==========================================
# SLIDE 14: 簽約特約底線與產權防線（保障雙方履約安全）
# ==========================================
s14 = prs.slides.add_slide(blank_layout)
add_header(s14, "合約安全防線：保障雙方交易安全之簽約特約底線")

s14_cards = [
    ("1. 全體公同共有人出席或合法授權", "• 說明書記載 5 人未會同申請、欠繳書狀費。\n• 簽約時堅持全體共有人親自到場；若委任需備妥戶政印鑑證明正本與合法公證授權書，杜絕日後無權代理爭端！", Inches(0.8), RED),
    ("2. 賣方自行負擔並清繳書狀規費", "• 契約明訂因未領權狀所產生之書狀規費、補發規費及欠繳稅費，一律由賣方自行負擔完竣，不得延誤登記過戶法定時程。", Inches(4.8), GOLD),
    ("3. 成數不足解約防護特約條款", "• 34 年老屋審核極嚴，若銀行核貸成數不足 8 成，非因買方個人信用瑕疵者，雙方得無條件解約退款，確保資金不被套牢。", Inches(8.8), PRIMARY_BLUE)
]

for title, desc, left, col in s14_cards:
    add_card(s14, left, Inches(1.5), Inches(3.7), Inches(5.3), border_color=col)
    tb = s14.shapes.add_textbox(left + Inches(0.2), Inches(1.7), Inches(3.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 15: 收尾逼單：10萬真金白銀斡旋支票促成簽約雙贏
# ==========================================
s15 = prs.slides.add_slide(blank_layout)
add_header(s15, "收尾促成簽約：隨身備妥 10 萬真金白銀斡旋支票，成全雙贏善緣！")

card_s15 = add_card(s15, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.3), bg_color=RGBColor(234, 250, 241), border_color=GREEN)
tb_s15 = s15.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(11.0), Inches(4.7))
tf_s15 = tb_s15.text_frame
tf_s15.word_wrap = True

p_s15_1 = tf_s15.paragraphs[0]
p_s15_1.text = "🏆 買方向全體長輩屋主的最真誠表白（現場一槌定音金句）："
p_s15_1.font.name = "微軟正黑體"
p_s15_1.font.size = Pt(16)
p_s15_1.font.bold = True
p_s15_1.font.color.rgb = GREEN

p_s15_2 = tf_s15.add_paragraph()
p_s15_2.text = (
    "「各位長輩好，我們今天帶著最大的誠意坐在這裡，不是為了挑毛病砍價，而是把所有市場數據與老屋現況攤開跟長輩理性商量。\n\n"
    "我們桌上這張 10 萬元銀行斡旋支票就是真金白銀！只要長輩願意體恤我們自住年輕家庭，理解 34 年老屋需要投入 120 萬翻修水電與衛浴防水的現實，願意在 710 萬至 730 萬元成全讓利，這張支票當場轉為定金，我們立刻動筆簽約！\n\n"
    "各位長輩是繼承取得，成本極低、土增稅才三千多塊，今天順利成交就是全額幾百萬純現金落袋為安，6 位長輩大家都能圓滿分產，更不用再每個月白白繳納管理費水電費。\n\n"
    "我們是誠意自住，買下來會好好重新整頓、用心善待這個起家厝，延續長輩的溫暖福氣。請長輩今晚就成全我們，大家圓圓滿滿結下一段美好善緣！」"
)
p_s15_2.font.name = "微軟正黑體"
p_s15_2.font.size = Pt(11.5)
p_s15_2.font.color.rgb = TEXT_DARK
p_s15_2.line_spacing = 1.35
p_s15_2.space_before = Pt(12)

# 儲存簡報
output_pptx = os.path.join(target_dir, "大墩人文11樓_屋主見面議價簡報.pptx")
prs.save(output_pptx)
print(f"Presentation saved successfully to: {output_pptx}")

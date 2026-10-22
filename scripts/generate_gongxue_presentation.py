# -*- coding: utf-8 -*-
"""
生成《工學天下15樓_房仲向屋主議價溝通簡報.pptx》
高品質 16:9 商業簡報，具備圖表、數據卡片、結構對比與精美配色
借鑑大墩人文與長億南園的專業簡報方法：
1. 檔名：工學天下15樓_房仲向屋主議價溝通簡報.pptx (及相容 .ppt)
2. 完整納入南區大盤走勢（53 棟大型社區成交量暴跌 76.4%）
3. 完整納入南區 10 大類似老牌指標社區實登折線走勢圖與深度剖析專章（4 頁鐵證大篇幅）
4. 完整納入工學天下實價登錄全紀錄（129筆）價量走勢圖與深度解析（最新摜破 2 字頭至 19.58 萬/坪）
5. 完整呈現 12 大說服理由（信義產調 *7 梯間虛坪縮水3坪、13 曾為須注意建物/921黃單、31.7年老屋水電未換雙衛浴無窗、完全無車位、屋主85年持有30年大賺落袋、同棟141號4F附平車20.8萬、正樓下13F無車位610萬/15.6萬、正樓上16F無車位530萬/13.2萬、同社區23間待售多殺多、雙重競品南屯大墩人文與南區長億南園施壓、出價精算 628~650 萬）
6. 只要屋主願意談，買方願意立即下斡旋金促成簽約
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

if sys.stdout:
    sys.stdout.reconfigure(encoding='utf-8')

# 色彩配置
NAVY = RGBColor(27, 54, 93)       # #1B365D 主深藍
DARK_BLUE = RGBColor(15, 32, 67)  # #0F2043 封面底色
SLATE = RGBColor(44, 94, 138)     # #2C5E8A 次藍
LIGHT_BG = RGBColor(248, 250, 252)# #F8FAFC 淺灰背景
WHITE = RGBColor(255, 255, 255)
CARD_BG = RGBColor(255, 255, 255)
BORDER_GRAY = RGBColor(226, 232, 240)
ROW_ALT = RGBColor(248, 250, 252)
GOLD = RGBColor(217, 140, 20)     # #D98C14 重點金
RED = RGBColor(192, 57, 43)       # #C0392B 警示紅
GREEN = RGBColor(39, 174, 96)     # #27AE60 成功綠
TEXT_DARK = RGBColor(30, 41, 59)  # #1E293B
TEXT_MUTED = RGBColor(100, 116, 139) # #64748B

base_dir = r"d:\GoogleDrive\01_個人_家庭與生活\02_家庭_安莉的資料\小寶的買房子資料"
target_dir = os.path.join(base_dir, "工學天下15樓")
trend_dir = os.path.join(base_dir, "南區25年以上大型社區走勢圖")

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

def add_header(slide, title_text, category="工學天下 15 樓之 1・房仲向屋主議價溝通簡報"):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, LIGHT_BG)
    
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p_cat = tf.paragraphs[0]
    p_cat.text = category.upper()
    p_cat.font.name = "微軟正黑體"
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = GOLD
    
    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.name = "微軟正黑體"
    p_title.font.size = Pt(20)
    p_title.font.bold = True
    p_title.font.color.rgb = NAVY
    p_title.space_before = Pt(2)

def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=BORDER_GRAY):
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    set_shape_flat(card, bg_color, border_color, 1)
    return card

# ==========================================
# SLIDE 1: 封面
# ==========================================
s1 = prs.slides.add_slide(blank_layout)
bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
set_shape_flat(bg1, DARK_BLUE)

tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.0))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1_sub1 = tf1.paragraphs[0]
p1_sub1.text = "工學天下 15 樓之 1（次頂樓溫馨三房・銷售編號 4262WL）"
p1_sub1.font.name = "微軟正黑體"
p1_sub1.font.size = Pt(22)
p1_sub1.font.bold = True
p1_sub1.font.color.rgb = GOLD

p1_main = tf1.add_paragraph()
p1_main.text = "房仲向屋主議價溝通簡報"
p1_main.font.name = "微軟正黑體"
p1_main.font.size = Pt(36)
p1_main.font.bold = True
p1_main.font.color.rgb = WHITE
p1_main.space_before = Pt(8)

p1_desc = tf1.add_paragraph()
p1_desc.text = "【市場大盤反轉急凍・南區10大指標社區全數走跌・工學天下全紀錄價量剖析・同社區23間競售賣壓・南屯大墩人文與南區長億南園競品威脅・掌握稀缺成熟自住買盤】"
p1_desc.font.name = "微軟正黑體"
p1_desc.font.size = Pt(14)
p1_desc.font.color.rgb = RGBColor(200, 215, 235)
p1_desc.space_before = Pt(14)

p1_foot = tf1.add_paragraph()
p1_foot.text = "報告單位：信義房屋專案團隊  |  標的：工學北路 141 號 15 樓之 1（40.10坪無車位）  |  買方狀況：成熟中年自住家庭・做足功課・屋主願談可立即下斡旋金"
p1_foot.font.name = "微軟正黑體"
p1_foot.font.size = Pt(11)
p1_foot.font.color.rgb = RGBColor(150, 175, 205)
p1_foot.space_before = Pt(36)

# ==========================================
# SLIDE 2: 執行摘要（4大關鍵局勢）
# ==========================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "執行摘要：屋主當下面臨的 4 大關鍵決策局勢")

exec_cards = [
    ("📉 大盤確認反轉下修", "第七波信用管制後，南區 53 棟老牌社區成交量暴跌 76.4%！中位數自 26.9 萬見頂回落，價格正式進入『緩跌』格局，近兩年買盤極端稀缺，越晚賣越難賣。", RED, Inches(0.8)),
    ("🏢 同社區 23 戶競售賣壓", "同社區 591 即有高達 23 戶在售，三房佔 22 戶！鄰近老牌社區更有數十戶在賣，委售大塞車，削價競爭劇烈，誰先讓利誰先拿現金！", GOLD, Inches(3.8)),
    ("⚡ 兩大強烈競品威脅", "買方正同時洽談南屯『大墩人文 11 樓』（室內 31.5 坪無灌水梯間）與南區『長億南園 8 樓』（室內 32 坪且水電已翻新），本案屋主若無意願協調，買方隨時轉移簽約！", SLATE, Inches(6.8)),
    ("🎯 把握稀缺成熟自住買方", "年輕人只買新小宅，買 31.7 年老屋的只剩中年家庭。買方純自住，只要屋主願意談，可立即下斡旋金，是目前市場極罕見的實質好買盤！", GREEN, Inches(9.8))
]

for title, desc, col, left in exec_cards:
    add_card(s2, left, Inches(1.5), Inches(2.75), Inches(5.2))
    tb = s2.shapes.add_textbox(left + Inches(0.2), Inches(1.7), Inches(2.35), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(14)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(11)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 3: 宏觀鐵證 1【南區大盤暴跌 76.4%】
# ==========================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "宏觀鐵證 1：南區大盤暴跌 76.4%，價格正式向下反轉")

mkt_img_path = os.path.join(trend_dir, "00_台中市南區25年以上大型社區_大盤匯總走勢圖.png")
if os.path.exists(mkt_img_path):
    s3.shapes.add_picture(mkt_img_path, Inches(0.8), Inches(1.5), Inches(6.8), Inches(5.2))

add_card(s3, Inches(7.8), Inches(1.5), Inches(4.733), Inches(5.2))
tb_s3 = s3.shapes.add_textbox(Inches(8.1), Inches(1.7), Inches(4.133), Inches(4.7))
tf_s3 = tb_s3.text_frame
tf_s3.word_wrap = True

p_s3_t = tf_s3.paragraphs[0]
p_s3_t.text = "📊 全南區大盤關鍵數據"
p_s3_t.font.name = "微軟正黑體"
p_s3_t.font.size = Pt(15)
p_s3_t.font.bold = True
p_s3_t.font.color.rgb = NAVY

s3_points = [
    ("大盤樣本完整", "全南區 53 棟 25 年以上老牌大社區、涵蓋 21,208 戶、累計 7,022 筆成交實價登錄。"),
    ("歷史見頂反轉", "113 年大盤中位數衝上 26.9 萬/坪歷史極值，第七波信用管制後急速反轉下修至 24.6 萬/坪。"),
    ("成交量雪崩 76.4%", "年成交量自高峰 738 件崩縮至 115 年僅 174 件，市場動能徹底急凍！"),
    ("『量先價行』鐵律", "沒有成交量支撐，老屋價格只有緩跌修正一途，不可再心存追高幻想！")
]

for title, desc in s3_points:
    p_p = tf_s3.add_paragraph()
    p_p.text = f"• {title}"
    p_p.font.name = "微軟正黑體"
    p_p.font.size = Pt(12)
    p_p.font.bold = True
    p_p.font.color.rgb = RED if "雪崩" in title or "鐵律" in title else SLATE
    p_p.space_before = Pt(8)
    
    p_pd = tf_s3.add_paragraph()
    p_pd.text = desc
    p_pd.font.name = "微軟正黑體"
    p_pd.font.size = Pt(10.5)
    p_pd.font.color.rgb = TEXT_DARK
    p_pd.space_before = Pt(2)

# ==========================================
# SLIDE 4: 大盤未來走勢預估
# ==========================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "大盤走勢預估：房市趨冷只會緩跌，房屋年齡越老越便宜")

s4_cards = [
    ("1. 信用管制掐死市場流動性", "• 央行第七波信用管制全面收緊房貸成數。\n• 銀行鑑價極度保守，放款滿水位持續延燒。\n• 實質房貸利率站穩 2.5%～3.0%，買方負擔能力大幅縮水，只會持續觀望砍價。", Inches(0.8), RED),
    ("2. 30 年老屋金融大關縮限", "• 工學天下已 31.7 年，跨越 30 年金融門檻。\n• 銀行內規『屋齡＋貸款年限』嚴格受限，無法貸滿 30/40 年長天期。\n• 買方每月還款本金加重，老屋接手門檻越來越高，未來只剩全現金買方。", Inches(4.8), GOLD),
    ("3. 未來三年價格逐年回測", "• 根據南區大盤模型推估：\n   - 116 年下探：21.1 萬/坪\n   - 117 年下探：18.0 萬/坪\n   - 118 年回測歷史支撐：15.8 萬/坪\n• 趁目前 20 幾萬高檔反轉初期賣出，是資產最佳安全落袋窗口！", Inches(8.8), NAVY)
]

for title, desc, left, col in s4_cards:
    add_card(s4, left, Inches(1.5), Inches(3.7), Inches(5.2))
    tb = s4.shapes.add_textbox(left + Inches(0.2), Inches(1.7), Inches(3.3), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(14)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(11)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 5: 鐵證總表【南區 10 大指標老牌社區實登全數反轉向下】
# ==========================================
s5_bench = prs.slides.add_slide(blank_layout)
add_header(s5_bench, "客觀實證：全南區 10 大類似指標老牌社區實價登錄全數反轉向下（鐵證總表）")

table_shape = s5_bench.shapes.add_table(11, 7, Inches(0.8), Inches(1.5), Inches(11.733), Inches(4.2))
tbl_b = table_shape.table
tbl_b.columns[0].width = Inches(2.1)
tbl_b.columns[1].width = Inches(1.2)
tbl_b.columns[2].width = Inches(1.8)
tbl_b.columns[3].width = Inches(1.8)
tbl_b.columns[4].width = Inches(1.7)
tbl_b.columns[5].width = Inches(1.2)
tbl_b.columns[6].width = Inches(1.933)

headers_bench = ["指標社區建案", "總戶數", "歷史最高單價", "115年最新成交", "113➜115均價", "跌幅", "實登走勢反轉特徵"]
for col_idx, h_text in enumerate(headers_bench):
    cell = tbl_b.cell(0, col_idx)
    set_cell_bg(cell, NAVY)
    p = cell.text_frame.paragraphs[0]
    p.text = h_text
    p.font.name = "微軟正黑體"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

bench_rows = [
    ("長億城攬翠區", "669 戶", "29.5 萬 (113/06)", "26.9 萬 (115/06)", "26.3 ➜ 23.2 萬", "-11.9%", "工學大案高點已過，量能大縮85%"),
    ("工學天下 (本社區)", "452 戶", "29.0 萬 (113/07)", "19.6 萬 (115/06)", "24.7 ➜ 20.7 萬", "-16.1%", "最新單價暴跌，正式摜破 2 字頭！"),
    ("長億南園", "171 戶", "27.2 萬 (113/05)", "21.4 萬 (115/09)", "25.1 ➜ 21.8 萬", "-13.1%", "德富路指標回檔，9樓僅782萬"),
    ("龍邦綠園道", "322 戶", "28.5 萬 (114/02)", "23.6 萬 (115/08)", "23.6 ➜ 21.3 萬", "-9.5%", "綠園道知名名宅高檔回落、買盤縮手"),
    ("文心高第", "266 戶", "34.5 萬 (113/05)", "30.4 萬 (115/07)", "31.3 ➜ 28.5 萬", "-8.9%", "東興路高樓名廈見頂反轉修正"),
    ("大東家華園", "227 戶", "27.7 萬 (113/03)", "20.3 萬 (115/07)", "24.5 ➜ 23.1 萬", "-5.9%", "學府商圈最新成交逼近 1 字頭！"),
    ("假日半島", "219 戶", "33.2 萬 (113/10)", "24.8 萬 (115/07)", "29.4 ➜ 25.3 萬", "-14.0%", "復興路指標大樓下修逾 14%"),
    ("文心大地", "207 戶", "28.4 萬 (113/07)", "24.1 萬 (115/05)", "26.4 ➜ 22.5 萬", "-15.1%", "文心南路聚落量能雪崩、價格下挫"),
    ("文心及第", "259 戶", "35.7 萬 (113/06)", "18.9 萬 (115/03)", "30.6 ➜ 18.9 萬", "-38.2%", "歷史極值暴跌，最新成交跌破 2 字頭！"),
    ("京城特區", "388 戶", "25.9 萬 (113/03)", "18.6 萬 (115/05)", "22.4 ➜ 20.9 萬", "-6.9%", "學府路大案連跌，最新跌至 18.6 萬！")
]

for row_idx, r_data in enumerate(bench_rows, start=1):
    for col_idx, text in enumerate(r_data):
        cell = tbl_b.cell(row_idx, col_idx)
        set_cell_bg(cell, ROW_ALT if row_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = text
        p.font.name = "微軟正黑體"
        p.font.size = Pt(9.5)
        p.font.color.rgb = TEXT_DARK
        if col_idx in [0, 3, 5]:
            p.font.bold = True
        if col_idx == 5:
            p.font.color.rgb = RED

card_b_foot = add_card(s5_bench, Inches(0.8), Inches(5.85), Inches(11.733), Inches(1.25), bg_color=RGBColor(254, 242, 242), border_color=RED)
tb_bf = s5_bench.shapes.add_textbox(Inches(1.0), Inches(5.95), Inches(11.333), Inches(1.05))
tf_bf = tb_bf.text_frame
tf_bf.word_wrap = True
p_bf1 = tf_bf.paragraphs[0]
p_bf1.text = "🎯 房仲議價關鍵說服點：全南區老牌大社區無一倖免！高點已過，量縮價跌是整個大盤的系統性下修"
p_bf1.font.name = "微軟正黑體"
p_bf1.font.size = Pt(11.5)
p_bf1.font.bold = True
p_bf1.font.color.rgb = RED

p_bf2 = tf_bf.add_paragraph()
p_bf2.text = "這 10 個指標老牌社區全部在 113 年見頂後反轉向下！其中工學天下、京城特區、文心及第最新成交更已直接跌破 2 字頭！工學天下絕非個案，此時若不趁高檔反轉初期讓利成交，後續跌幅只會更深！"
p_bf2.font.name = "微軟正黑體"
p_bf2.font.size = Pt(10)
p_bf2.font.color.rgb = TEXT_DARK
p_bf2.space_before = Pt(3)

# ==========================================
# SLIDE 6: 指標實證 1【工學商圈大案全面回檔】長億城・工學天下・京城特區
# ==========================================
s6_g1 = prs.slides.add_slide(blank_layout)
add_header(s6_g1, "指標實證 1【工學商圈大案全面回檔】長億城・工學天下・京城特區")

g1_data = [
    ("長億城攬翠區（669戶・屋齡30.9年）", "03_長億城攬翠區陽光區_成交行情走勢圖.png",
     "• 113年高點：29.5 萬/坪 (8樓)\n• 115年最新成交：26.9 萬/坪\n• 均價自 26.3 萬跌至 23.2 萬 (-11.9%)\n• 年成交量自高峰 63 件雪崩至 9 件！\n• 669 戶大社區委售競爭激烈，價格持續下探。"),
    ("工學天下（452戶・屋齡31.7年）", "07_工學天下_成交行情走勢圖.png",
     "• 113年高點：29.0 萬/坪 (9樓)\n• 115年最新成交：19.6 萬/坪 (7樓)\n• 🔥 最新成交正式跌破 2 字頭！\n• 均價自 24.7 萬跌至 20.7 萬 (-16.1%)\n• 買盤對 30 年以上無車位老屋縮手，殺價猛烈。"),
    ("京城特區（388戶・屋齡29.1年）", "11_京城特區_成交行情走勢圖.png",
     "• 113年高點：25.9 萬/坪 (18樓)\n• 115年最新成交：18.6 萬/坪 (3樓)\n• 🔥 跌破 2 字頭，單價見 18.6 萬低點！\n• 均價自 22.4 萬下修至 20.9 萬 (-6.9%)\n• 學府路核心生活圈老屋買氣急凍、去化極慢。")
]

for idx, (title, img_name, desc) in enumerate(g1_data):
    left = Inches(0.8 + idx * 4.0)
    img_p = os.path.join(trend_dir, img_name)
    if os.path.exists(img_p):
        s6_g1.shapes.add_picture(img_p, left, Inches(1.5), width=Inches(3.7))
    
    card = add_card(s6_g1, left, Inches(3.95), Inches(3.7), Inches(3.1))
    tb = s6_g1.shapes.add_textbox(left + Inches(0.15), Inches(4.05), Inches(3.4), Inches(2.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(11.5)
    p_t.font.bold = True
    p_t.font.color.rgb = NAVY
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.2
    p_d.space_before = Pt(4)

# ==========================================
# SLIDE 7: 指標實證 2【復興與文心核心聚落反轉】文心及第・假日半島・文心高第
# ==========================================
s7_g2 = prs.slides.add_slide(blank_layout)
add_header(s7_g2, "指標實證 2【復興與文心核心聚落反轉】文心及第・假日半島・文心高第")

g2_data = [
    ("文心及第（259戶・屋齡27.3年）", "27_文心及第文心南路_成交行情走勢圖.png",
     "• 113年高點：35.7 萬/坪 (10樓)\n• 115年最新成交：18.9 萬/坪 (14樓)\n• 🔥 高點暴跌 38.2%，高樓層亦破 2 字頭！\n• 均價自 30.6 萬腰斬至 18.9 萬\n• 文心南路捷運生活圈老屋面臨買盤嚴苛砍價。"),
    ("假日半島（219戶・屋齡32.4年）", "38_假日半島_成交行情走勢圖.png",
     "• 113年高點：33.2 萬/坪 (9樓)\n• 115年最新成交：24.8 萬/坪 (5樓)\n• 均價自 29.4 萬跌至 25.3 萬 (-14.0%)\n• 復興路二段名宅高檔震盪下修\n• 交易量能萎縮逾八成，買方出價保守謹慎。"),
    ("文心高第（266戶・屋齡24.9年）", "26_文心高第_成交行情走勢圖.png",
     "• 113年高點：34.5 萬/坪 (12樓)\n• 115年最新成交：30.4 萬/坪 (2樓)\n• 均價自 31.3 萬下修至 28.5 萬 (-8.9%)\n• 東興路指標建案高檔反轉\n• 買盤難耐高總價，委售時間由 2 個月拉長至半年以上。")
]

for idx, (title, img_name, desc) in enumerate(g2_data):
    left = Inches(0.8 + idx * 4.0)
    img_p = os.path.join(trend_dir, img_name)
    if os.path.exists(img_p):
        s7_g2.shapes.add_picture(img_p, left, Inches(1.5), width=Inches(3.7))
    
    card = add_card(s7_g2, left, Inches(3.95), Inches(3.7), Inches(3.1))
    tb = s7_g2.shapes.add_textbox(left + Inches(0.15), Inches(4.05), Inches(3.4), Inches(2.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(11.5)
    p_t.font.bold = True
    p_t.font.color.rgb = NAVY
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.2
    p_d.space_before = Pt(4)

# ==========================================
# SLIDE 8: 指標實證 3【德富、園道與學府聚落破底】長億南園・龍邦綠園道・大東家華園・大安黃金城
# ==========================================
s8_g3 = prs.slides.add_slide(blank_layout)
add_header(s8_g3, "指標實證 3【德富、園道與學府聚落破底】長億南園・龍邦綠園道・大東家華園・大安黃金城")

g3_data = [
    ("長億南園（171戶）", "28_長億南園_成交行情走勢圖.png",
     "• 高點：27.2 萬 (113年)\n• 最新：21.4 萬 (115/09)\n• 均價：25.1 ➜ 21.8 萬 (-13.1%)\n• 9 樓總價僅 782 萬，永義積極洽談本組買方！"),
    ("龍邦綠園道（322戶）", "15_龍邦綠園道_成交行情走勢圖.png",
     "• 高點：28.5 萬 (114年)\n• 最新：23.6 萬 (115/08)\n• 均價：23.6 ➜ 21.3 萬 (-9.5%)\n• 園道首排支撐鬆動，價格高檔回落。"),
    ("大東家華園（227戶）", "35_大東家華園_成交行情走勢圖.png",
     "• 高點：27.7 萬 (113年)\n• 最新：20.3 萬 (115/07)\n• 🔥 最新成交逼近 1 字頭！\n• 均價：24.5 ➜ 23.1 萬 (-5.9%)。"),
    ("大安黃金城（307戶）", "18_大安黃金城大廈_成交行情走勢圖.png",
     "• 高點：34.3 萬 (114年)\n• 最新：25.2 萬 (115/06)\n• 均價：26.6 ➜ 23.7 萬 (-11.1%)\n• 中興大學生活圈買盤全面縮手。")
]

for idx, (title, img_name, desc) in enumerate(g3_data):
    left = Inches(0.8 + idx * 2.98)
    img_p = os.path.join(trend_dir, img_name)
    if os.path.exists(img_p):
        s8_g3.shapes.add_picture(img_p, left, Inches(1.5), width=Inches(2.78))
    
    card = add_card(s8_g3, left, Inches(3.45), Inches(2.78), Inches(3.6))
    tb = s8_g3.shapes.add_textbox(left + Inches(0.1), Inches(3.55), Inches(2.58), Inches(3.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(10.5)
    p_t.font.bold = True
    p_t.font.color.rgb = NAVY
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(9)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.15
    p_d.space_before = Pt(3)

# ==========================================
# SLIDE 9: 工學天下實價登錄全紀錄（129筆）價量走勢全解析
# ==========================================
s9 = prs.slides.add_slide(blank_layout)
add_header(s9, "工學天下全紀錄價量剖析：115年最新成交摜破20萬至19.58萬，成交量大縮84%")

gx_chart_path = os.path.join(target_dir, "工學天下15樓_成交單價價量走勢圖.png")
if os.path.exists(gx_chart_path):
    s9.shapes.add_picture(gx_chart_path, Inches(0.8), Inches(1.5), Inches(6.8), Inches(5.2))

add_card(s9, Inches(7.8), Inches(1.5), Inches(4.733), Inches(5.2))
tb_s9 = s9.shapes.add_textbox(Inches(8.1), Inches(1.7), Inches(4.133), Inches(4.7))
tf_s9 = tb_s9.text_frame
tf_s9.word_wrap = True

p_s9_t = tf_s9.paragraphs[0]
p_s9_t.text = "🏢 工學天下實登關鍵鐵證"
p_s9_t.font.name = "微軟正黑體"
p_s9_t.font.size = Pt(15)
p_s9_t.font.bold = True
p_s9_t.font.color.rgb = NAVY

s9_points = [
    ("115/06 最新無車位跌破 20 萬", "189號7F之3 無車位成交 19.58 萬/坪（折合本戶 40.1 坪僅值 785 萬）！"),
    ("115/01 最新無車位僅 660 萬", "189號3F之2 無車位成交總價更僅 660 萬（20.77 萬/坪），總價定錨極低！"),
    ("年度成交均價大跌 16.1%", "年度均價自 113 年 24.68 萬大跌至 115 年 20.70 萬，年成交量自 25 件急凍至 4 件！"),
    ("長期合理歷史均價僅 17 萬", "101~109 年合理區間為 12~16 萬，樂居全期均價 17.02 萬。本案開 23.14 萬完全偏離行情！")
]

for title, desc in s9_points:
    p_p = tf_s9.add_paragraph()
    p_p.text = f"• {title}"
    p_p.font.name = "微軟正黑體"
    p_p.font.size = Pt(12)
    p_p.font.bold = True
    p_p.font.color.rgb = RED if "跌破" in title or "大跌" in title else SLATE
    p_p.space_before = Pt(8)
    
    p_pd = tf_s9.add_paragraph()
    p_pd.text = desc
    p_pd.font.name = "微軟正黑體"
    p_pd.font.size = Pt(10.5)
    p_pd.font.color.rgb = TEXT_DARK
    p_pd.space_before = Pt(2)

# ==========================================
# SLIDE 10: 理由 1【說明書 *7 條：主建物灌水電樓梯間，實際縮水 3 坪】
# ==========================================
s10 = prs.slides.add_slide(blank_layout)
add_header(s10, "說服理由 1：產調說明書 *7 條鐵證揭露，主建物含電樓梯間室內縮水 3 坪！")

s10_cards = [
    ("信義房屋產調說明書白紙黑字鐵證", "• 注意事項第 *7 條明確記載：\n  『本物件經信義房屋確認主建物登記面積（27.1坪）含該層之（電）樓梯間，因此可能造成室內實際面積小於主建物登記之面積』。\n• 早期建商將當層 3 戶共用的 2 部電梯廳與公共樓梯間塞入主建物登記！\n• 扣除分攤梯間公設（約 2.5～3.5 坪），室內實質可用面積僅約 23.5～24 坪，實質公設比超過 30%！", Inches(0.8), RED),
    ("買方理性爭取折讓 60 萬虛坪價差", "• 買方複看時拿捲尺一對即發現室內實坪嚴重縮水。\n• 買方表示：『我花錢買 27.1 坪主建物，結果裡面有 3 坪是門外的梯間公設！這 3 坪虛坪算下來就要 60 萬元！』\n• 屋主本來就應該將梯間公設虛坪價差實質折讓給買方，這是白紙黑字事實，談判桌上無法反駁！", Inches(6.8), GOLD)
]

for title, desc, left, col in s10_cards:
    add_card(s10, left, Inches(1.5), Inches(5.7), Inches(5.2))
    tb = s10.shapes.add_textbox(left + Inches(0.3), Inches(1.8), Inches(5.1), Inches(4.6))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(15)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(12)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 11: 理由 2【說明書 13 條：曾列「須注意之建築物（921黃單）」，貸款審核極嚴】
# ==========================================
s11 = prs.slides.add_slide(blank_layout)
add_header(s11, "說服理由 2：產調 13 條註記曾為須注意建築物（921黃單），銀行貸款保守")

s11_cards = [
    ("信義說明書第 13 條警示條款", "• 說明書白紙黑字記載：\n  『本棟建物為須注意之建築物，業已修復完成並取得建築師專業鑑定報告，建議承買人簽約前確認金融機構貸款條件』。\n• 依法必須揭露，看屋買方十組有八組看到『須注意之建築物』即嚇退！\n• 銀行內部亦有歷史列管紀錄，鑑價與成數審核極其保守！", Inches(0.8), RED),
    ("31.7 年老屋＋無車位＋黃單三重限制", "• 屋齡已達 31.7 年，加上無固定車位，銀行貸款年限與成數遭到大幅壓縮。\n• 買方必須自行準備充裕現金彌補自備款落差。\n• 這組成熟中年買方願包容歷史瑕疵、理性承接，是千載難逢的誠意買盤，屋主價格務必實質讓步！", Inches(6.8), SLATE)
]

for title, desc, left, col in s11_cards:
    add_card(s11, left, Inches(1.5), Inches(5.7), Inches(5.2))
    tb = s11.shapes.add_textbox(left + Inches(0.3), Inches(1.8), Inches(5.1), Inches(4.6))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(15)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(12)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 12: 理由 3【115年最新無車位成交已跌破 20 萬至 19.58 萬/坪】
# ==========================================
s12 = prs.slides.add_slide(blank_layout)
add_header(s12, "說服理由 3：115 年最新無車位成交摜破 20 萬，最低總價僅 660 萬！")

table_shape_s12 = s12.shapes.add_table(5, 4, Inches(0.8), Inches(1.6), Inches(11.733), Inches(4.0))
tbl_s12 = table_shape_s12.table
tbl_s12.columns[0].width = Inches(2.2)
tbl_s12.columns[1].width = Inches(2.2)
tbl_s12.columns[2].width = Inches(2.4)
tbl_s12.columns[3].width = Inches(4.933)

headers_s12 = ["成交時間與戶別", "車位狀況 / 坪數", "成交總價 / 單價", "對本案（141號15F之1）的定錨意義"]
for col_idx, h in enumerate(headers_s12):
    cell = tbl_s12.cell(0, col_idx)
    set_cell_bg(cell, NAVY)
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.name = "微軟正黑體"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

s12_rows = [
    ("115/06 189號7樓之3", "無車位 (49.54坪)", "970 萬\n(19.58 萬/坪)", "最新一筆無車位實登，直接跌破 20 萬大關！折合本戶僅值 785 萬。"),
    ("115/01 189號3樓之2", "無車位 (31.78坪)", "660 萬\n(20.77 萬/坪)", "115 年最新無車位 3 房最低成交總價定錨在 660 萬！"),
    ("114/01 189號5樓之3", "機械車位 (54.10坪)", "1,050 萬\n(19.41 萬/坪)", "含車位都才 19.41 萬/坪！以此單價乘本戶 40.1 坪僅值 778 萬。"),
    ("【本案開價】141號15F", "無車位 (40.10坪)", "928 萬\n(23.14 萬/坪)", "嚴重脫節！比最新無車位成交價（19.58萬）貴了近兩成，無人追價！")
]

for row_idx, r_data in enumerate(s12_rows, start=1):
    for col_idx, t in enumerate(r_data):
        cell = tbl_s12.cell(row_idx, col_idx)
        set_cell_bg(cell, ROW_ALT if row_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = t
        p.font.name = "微軟正黑體"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_DARK
        if col_idx in [0, 2]:
            p.font.bold = True
        if row_idx == 4 and col_idx == 2:
            p.font.color.rgb = RED

tb_s12_bot = s12.shapes.add_textbox(Inches(0.8), Inches(5.8), Inches(11.733), Inches(1.1))
tf_b12 = tb_s12_bot.text_frame
p_b12 = tf_b12.paragraphs[0]
p_b12.text = "💡 房仲核心建言：最新無車位成交已實質跌回 19.58 萬/坪，總價最低僅 660 萬。照最新實登計算，本戶天花板頂多 780 萬，完全不可能有買方去追 23 萬以上開價！"
p_b12.font.name = "微軟正黑體"
p_b12.font.size = Pt(12)
p_b12.font.bold = True
p_b12.font.color.rgb = RED

# ==========================================
# SLIDE 13: 理由 4【完全無車位致命硬傷 vs 同棟 141 號 4 樓附平車才賣 20.8 萬/坪】
# ==========================================
s13 = prs.slides.add_slide(blank_layout)
add_header(s13, "說服理由 4：完全無車位致命硬傷 vs 同棟 141 號附平車才賣 20.8 萬/坪！")

s13_cards = [
    ("工學商圈停車極難，無車位是最大抗性", "• 樓下為工學市場與商圈，周邊臨停極度困難。\n• 三房家庭大多有開車通勤需求，無固定車位是中古轉手最大抗性。\n• 買方若在外租車位，一個月 3,000 元，住 20 年就要多付 72 萬租金！這筆折價必然由房價扣除！", Inches(0.8), RED),
    ("同棟 141 號 4 樓附平面車位僅賣 20.86 萬/坪", "• 買方調出同一棟（141 號 4 樓之 2）113/03 實登紀錄：\n  附一個坡道平面大車位（47.95 坪），成交單價僅 20.86 萬/坪（總價 1,000 萬）！\n• 買方質疑：『人家同棟有平面大車位在最熱時才賣 20.8 萬，你們這間完全沒車位憑什麼開 23.14 萬？』", Inches(6.8), GOLD)
]

for title, desc, left, col in s13_cards:
    add_card(s13, left, Inches(1.5), Inches(5.7), Inches(5.2))
    tb = s13.shapes.add_textbox(left + Inches(0.3), Inches(1.8), Inches(5.1), Inches(4.6))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(15)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(12)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 14: 理由 5【同棟正樓下 13 樓 610 萬與正樓上 16 樓 530 萬實登定錨！】
# ==========================================
s14 = prs.slides.add_slide(blank_layout)
add_header(s14, "說服理由 5：同棟正樓下 13 樓 610 萬與正樓上 16 樓 530 萬實登定錨！")

s14_cards = [
    ("同棟正樓下 141 號 13 樓之 1 (無車位)", "• 坪數：39.19 坪 (與本案格局同座向)\n• 成交時間：民國 109 年 4 月\n• 成交總價：610 萬 (單價 15.57 萬/坪)\n• 正樓下無車位正常市場價即為 610 萬！", Inches(0.8), SLATE),
    ("同棟正樓上 141 號 16 樓之 1 (無車位)", "• 坪數：40.30 坪 (頂樓戶無車位)\n• 成交時間：民國 108 年 1 月\n• 成交總價：530 萬 (單價 13.15 萬/坪)\n• 正樓上無車位正常市場價僅 530 萬！", Inches(4.8), NAVY),
    ("信義行情表區段最低單價", "• 門牌：155 號 3 樓之 1\n• 成交時間：114 年 9 月\n• 成交總價：630 萬 (單價 12.54 萬/坪)\n• 買方已被 610~660 萬定錨，出價 600 多萬合情合理！", Inches(8.8), GOLD)
]

for title, desc, left, col in s14_cards:
    add_card(s14, left, Inches(1.5), Inches(3.7), Inches(5.2))
    tb = s14.shapes.add_textbox(left + Inches(0.2), Inches(1.7), Inches(3.3), Inches(4.7))
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
    p_d.font.size = Pt(11)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 15: 理由 6【31.7年老屋水電未換＋雙衛浴無窗，120萬修繕預算扣除】
# ==========================================
s15 = prs.slides.add_slide(blank_layout)
add_header(s15, "說服理由 6：31.7 年老屋管線未曾更換＋雙衛浴無窗，120 萬翻修剛性支出")

s15_cards = [
    ("屋況客觀限制與必要工程清單", "• 說明書勾選『水電管線未曾更換』，32 年老舊電線與冷熱水管必須全數重拉！\n• 雙衛浴皆無對外窗（暗衛），必須打除見底重做防水層與加裝暖風排氣系統。\n• 前陽台加窗屬既存違建需重整，更換隔音防水氣密窗。\n• 廚房老舊流理台櫥櫃換新、全室油漆泥作，統包基礎工程報價達 110～120 萬！", Inches(0.8), RED),
    ("買方是純自住，預算嚴格扣除工程款", "• 買方不是投資客隨便粉刷轉租，而是家庭長期自住，居家安全不可妥協。\n• 120 萬翻修工程款是實打實的現金支出，必須從購屋總預算中倒扣。\n• 買方出價是經過精密財務精算後的真實極限，絕非惡意殺價！", Inches(6.8), SLATE)
]

for title, desc, left, col in s15_cards:
    add_card(s15, left, Inches(1.5), Inches(5.7), Inches(5.2))
    tb = s15.shapes.add_textbox(left + Inches(0.3), Inches(1.8), Inches(5.1), Inches(4.6))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(15)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(12)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 16: 理由 7【同社區高達 23 間在售多殺多危機】
# ==========================================
s16 = prs.slides.add_slide(blank_layout)
add_header(s16, "說服理由 7：同社區高達 23 間在售多殺多危機，早一步成交勝過被鄰居踐踏")

s16_cards = [
    ("同社區 23 戶在售大塞車", "• 目前 591 網站上『工學天下』同社區即有高達 23 戶同時銷售（三房佔 22 戶）！\n• 開價區間自 788 萬至 1,398 萬，社區待售供給量極度龐大。\n• 買方進來看工學天下，一次可以挑選比較多間，選擇極多！", Inches(0.8), GOLD),
    ("誰先讓利，誰就能先把數百萬現金抱回家", "• 現在市場買氣急凍，只要社區有急售鄰居撐不住率先降價成交在 17~18 萬，後面所有買方出價會再殺兩成！\n• 趁本組買方最喜歡 15 樓視野、尚未被別戶吸走前成交，才不會陷入社區多殺多泥淖！", Inches(6.8), RED)
]

for title, desc, left, col in s16_cards:
    add_card(s16, left, Inches(1.5), Inches(5.7), Inches(5.2))
    tb = s16.shapes.add_textbox(left + Inches(0.3), Inches(1.8), Inches(5.1), Inches(4.6))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(15)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(12)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 17: 理由 8【屋主 85 年買入持有 30 年成本僅 300 多萬，大賺出場落袋為安】
# ==========================================
s17 = prs.slides.add_slide(blank_layout)
add_header(s17, "說服理由 8：屋主 85 年買入持有整整 30 年成本僅 300 多萬，大賺出場落袋為安")

table_shape_s17 = s17.shapes.add_table(5, 3, Inches(0.8), Inches(1.6), Inches(11.733), Inches(3.8))
tbl_s17 = table_shape_s17.table
tbl_s17.columns[0].width = Inches(3.2)
tbl_s17.columns[1].width = Inches(4.266)
tbl_s17.columns[2].width = Inches(4.266)

headers_s17 = ["財務 / 稅務項目", "歷史數據與條件", "本次成交預估結果"]
for col_idx, h in enumerate(headers_s17):
    cell = tbl_s17.cell(0, col_idx)
    set_cell_bg(cell, NAVY)
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.name = "微軟正黑體"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

s17_rows = [
    ("屋主原始買入成本", "民國 85 年 5 月登記：約 320～400 萬 (8~10 萬/坪)", "本金早已 100% 回收，無任何房貸壓力"),
    ("產權持有時間", "民國 85 年至今：已整整滿 30 年！", "第一手屋主，資產累積期長"),
    ("成交實質獲利評估", "在 600 多萬價位順利成交", "淨賺 250 萬～350 萬純現金落袋！"),
    ("房屋現況持有成本", "目前空屋，每月白繳 1,943 元管理費", "及時變現，省去每月管理費與老屋跌價折舊風險！")
]

for row_idx, r_data in enumerate(s17_rows, start=1):
    for col_idx, t in enumerate(r_data):
        cell = tbl_s17.cell(row_idx, col_idx)
        set_cell_bg(cell, ROW_ALT if row_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = t
        p.font.name = "微軟正黑體"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_DARK
        if col_idx == 0:
            p.font.bold = True

tb_s17_p = s17.shapes.add_textbox(Inches(0.8), Inches(5.6), Inches(11.733), Inches(1.3))
tf_p17 = tb_s17_p.text_frame
p_p17 = tf_p17.paragraphs[0]
p_p17.text = "🏆 獲利了結論點：屋主持有滿 30 年，無論成交 630 萬還是 680 萬，都是倒賺 250~350 萬現金落袋！\n在房市反轉向下之際，及時把數百萬現金抱在身上，比繼續空置老屋承擔跌價風險聰明太多！"
p_p17.font.name = "微軟正黑體"
p_p17.font.size = Pt(12)
p_p17.font.bold = True
p_p17.font.color.rgb = GREEN

# ==========================================
# SLIDE 18: 理由 9【強烈競品雙重施壓：南屯『大墩人文』與南區『長億南園』】
# ==========================================
s18 = prs.slides.add_slide(blank_layout)
add_header(s18, "說服理由 9：強烈競品雙重施壓，南屯大墩人文與南區長億南園隨時簽約！")

table_shape_s18 = s18.shapes.add_table(5, 4, Inches(0.8), Inches(1.6), Inches(11.733), Inches(3.8))
tbl_s18 = table_shape_s18.table
tbl_s18.columns[0].width = Inches(2.2)
tbl_s18.columns[1].width = Inches(3.1)
tbl_s18.columns[2].width = Inches(3.1)
tbl_s18.columns[3].width = Inches(3.333)

headers_s18 = ["評估指標", "南屯【大墩人文 11 樓】", "南區【長億南園 8 樓】", "工學天下 15 樓之 1 (本案)"]
for col_idx, h in enumerate(headers_s18):
    cell = tbl_s18.cell(0, col_idx)
    set_cell_bg(cell, NAVY)
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.name = "微軟正黑體"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

s18_rows = [
    ("地段生活圈", "南屯捷運站、大墩森林公園富人商圈", "南區德富路商圈、近忠明南路", "南區工學市場商圈、高工路生活圈"),
    ("室內實質坪數", "室內 31.5 坪 (純室內，無灌水梯間！)", "室內 32 坪超大空間 (附機上車位)", "權狀 40.1 坪 (主建物灌梯間，實質 23.5 坪)"),
    ("屋況整修現況", "需基礎翻修，但戶數 188 戶單純質感佳", "6 年前整理過水電衛浴，永義積極撮合", "31.7 年老屋管線未曾更換，需全室翻修"),
    ("買方當下態度", "買方曾租住有感情，信義房屋積極洽談中！", "永義房屋積極催促下斡旋！", "若本戶價格踩太硬，買方立刻轉身簽別間！")
]

for row_idx, r_data in enumerate(s18_rows, start=1):
    for col_idx, t in enumerate(r_data):
        cell = tbl_s18.cell(row_idx, col_idx)
        set_cell_bg(cell, ROW_ALT if row_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = t
        p.font.name = "微軟正黑體"
        p.font.size = Pt(9.5)
        p.font.color.rgb = TEXT_DARK
        if col_idx == 0:
            p.font.bold = True

tb_s18_w = s18.shapes.add_textbox(Inches(0.8), Inches(5.6), Inches(11.733), Inches(1.3))
tf_w18 = tb_s18_w.text_frame
p_w18 = tf_w18.paragraphs[0]
p_w18.text = "⚠️ 經紀人急迫預警：買方手上有大墩人文與長億南園兩大強烈競品，哪一間先讓利談進預算就簽哪一間！若我們堅持高價不談，買方轉頭簽約，本案將徹底失去難得的自住好買方！"
p_w18.font.name = "微軟正黑體"
p_w18.font.size = Pt(12)
p_w18.font.bold = True
p_w18.font.color.rgb = RED

# ==========================================
# SLIDE 19: 理由 10 & 11【買方「實登最低價再低兩成」出價精算與斡旋方案】
# ==========================================
s19 = prs.slides.add_slide(blank_layout)
add_header(s19, "說服理由 10 & 11：買方『比實登最低價再低兩成』科學精算與出價方案")

s19_cards = [
    ("方案 A：實登最低價打 8 折（★最推薦首選）", "• 基準：115/06 最新無車位實登 19.58 萬/坪 (折合本戶 785 萬)。\n• 精算：785 萬 × 0.8 (打 8 折) ＝ 628 萬 (15.66 萬/坪)！\n• 扣除 120 萬翻修＋35 萬梯間虛坪，出價有憑有據！", Inches(0.8), GOLD),
    ("方案 B：信義房屋 7 折收斡方案", "• 基準：目前開價 928 萬 × 0.7 ＝ 650 萬 (16.21 萬/坪)！\n• 符合信義店頭收斡慣例，總價剛好站上 650 萬整數關卡。\n• 能迅速促成信義店頭向屋主收斡見面談！", Inches(4.8), NAVY),
    ("方案 C：同條件 15 樓歷史成交打 8 折封頂", "• 基準：112/08 149 號 15 樓無車位成交 840 萬 × 0.8 ＝ 672 萬 (16.76 萬/坪)！\n• 買方家庭談判之封頂絕對天花板，絕不再加價！", Inches(8.8), GREEN)
]

for title, desc, left, col in s19_cards:
    add_card(s19, left, Inches(1.5), Inches(3.7), Inches(5.2))
    tb = s19.shapes.add_textbox(left + Inches(0.2), Inches(1.7), Inches(3.3), Inches(4.7))
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
    p_d.font.size = Pt(11)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 20: 臨門一腳促成簽約
# ==========================================
s20 = prs.slides.add_slide(blank_layout)
add_header(s20, "臨門一腳促成簽約：『只要屋主願意談，買方願意立即下斡旋金！』")

add_card(s20, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.2), RGBColor(234, 250, 241), GREEN)
tb_s20 = s20.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(11.0), Inches(4.6))
tf_s20 = tb_s20.text_frame
tf_s20.word_wrap = True

p_s20_1 = tf_s20.paragraphs[0]
p_s20_1.text = "🎯 房仲逼單最關鍵的一句話（激發屋主落袋為安與急迫感）："
p_s20_1.font.name = "微軟正黑體"
p_s20_1.font.size = Pt(16)
p_s20_1.font.bold = True
p_s20_1.font.color.rgb = GREEN

p_s20_2 = tf_s20.add_paragraph()
p_s20_2.text = (
    "「大哥／大姐，現在市場上光看不買的假客人很多，但真正有誠意的自住買方少之又少。\n\n"
    "這組自住買方非常有誠意：只要大哥／大姐願意談，買方願意立即下斡旋金，隨時配合我們正式簽約履約！\n\n"
    "您民國 85 年買 300 多萬，現在成交已經讓您穩穩淨賺兩三百萬現金出場，還能完美避開後面南區老屋量縮價跌與同社區 23 戶削價競爭的風險。\n\n"
    "而且買方甚至承諾會花上百萬把全室水電與防水重新翻新、好好善待這個起家厝。信義房屋那邊今天還在積極催買方去簽南屯的大墩人文，請大哥／大姐今天就做決定，給我們授權簽字，我們立刻幫您把數百萬現金穩穩落袋！」"
)
p_s20_2.font.name = "微軟正黑體"
p_s20_2.font.size = Pt(13)
p_s20_2.font.color.rgb = TEXT_DARK
p_s20_2.line_spacing = 1.3
p_s20_2.space_before = Pt(14)

# ==========================================
# SLIDE 21: 封底
# ==========================================
s21 = prs.slides.add_slide(blank_layout)
bg21 = s21.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
set_shape_flat(bg21, DARK_BLUE)

tb21 = s21.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.333), Inches(3.5))
tf21 = tb21.text_frame
tf21.word_wrap = True

p21_1 = tf21.paragraphs[0]
p21_1.text = "抓緊大盤反轉前夕  •  把握唯一誠意自住買盤"
p21_1.font.name = "微軟正黑體"
p21_1.font.size = Pt(28)
p21_1.font.bold = True
p21_1.font.color.rgb = GOLD
p21_1.alignment = PP_ALIGN.CENTER

p21_2 = tf21.add_paragraph()
p21_2.text = "恭祝 屋主資產高點變現・圓滿落袋為安大贏家！"
p21_2.font.name = "微軟正黑體"
p21_2.font.size = Pt(22)
p21_2.font.bold = True
p21_2.font.color.rgb = WHITE
p21_2.alignment = PP_ALIGN.CENTER
p21_2.space_before = Pt(16)

p21_3 = tf21.add_paragraph()
p21_3.text = "信義房屋專案團隊 敬製  |  工學天下 15 樓之 1"
p21_3.font.name = "微軟正黑體"
p21_3.font.size = Pt(13)
p21_3.font.color.rgb = RGBColor(190, 210, 240)
p21_3.alignment = PP_ALIGN.CENTER
p21_3.space_before = Pt(24)

# 儲存
pptx_path = os.path.join(target_dir, "工學天下15樓_房仲向屋主議價溝通簡報.pptx")
prs.save(pptx_path)
print(f"PowerPoint 簡報已生成: {pptx_path}")

# 相容 .ppt
ppt_compat = os.path.join(target_dir, "工學天下15樓_房仲向屋主議價溝通簡報.ppt")
prs.save(ppt_compat)
print(f"PowerPoint 相容檔案已生成: {ppt_compat}")

print("簡報生成全部完成！")

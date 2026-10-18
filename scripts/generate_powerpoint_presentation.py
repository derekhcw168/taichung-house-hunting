# -*- coding: utf-8 -*-
"""
生成升級版《長億南園8樓_房仲向屋主議價說服簡報.pptx》
納入全南區 10 大指標老牌社區實價登錄全面反轉向下圖表與數據
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
target_dir = os.path.join(base_dir, "長億南園8樓")
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

def add_header(slide, title_text, category="長億南園 8 樓・房仲議價說服專案報告"):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, LIGHT_BG)
    
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    set_shape_flat(top_bar, NAVY)

    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.35), Inches(4.5), Inches(0.32))
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

    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.68), Inches(11.7), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "微軟正黑體"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = NAVY

def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=BORDER_GRAY):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    set_shape_flat(card, bg_color, border_color, 1)
    return card

# ==========================================
# SLIDE 1: 封面
# ==========================================
s1 = prs.slides.add_slide(blank_layout)
bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
set_shape_flat(bg1, DARK_BLUE)

dec = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.8), Inches(0.15), Inches(3.2))
set_shape_flat(dec, GOLD)

tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(1.6), Inches(11), Inches(3.5))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1 = tf1.paragraphs[0]
p1.text = "長億南園 8 樓"
p1.font.name = "微軟正黑體"
p1.font.size = Pt(24)
p1.font.bold = True
p1.font.color.rgb = GOLD

p2 = tf1.add_paragraph()
p2.text = "房仲向屋主議價溝通說帖與大盤反轉專案簡報"
p2.font.name = "微軟正黑體"
p2.font.size = Pt(32)
p2.font.bold = True
p2.font.color.rgb = WHITE
p2.space_before = Pt(8)

p3 = tf1.add_paragraph()
p3.text = "【南區 10 大指標老牌社區實登全線反轉・量縮76%下行週期・實質4戶競售剖析・把握唯一自住買盤】"
p3.font.name = "微軟正黑體"
p3.font.size = Pt(14)
p3.font.color.rgb = RGBColor(190, 210, 240)
p3.space_before = Pt(12)

card_foot = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.6), Inches(11.733), Inches(1.1))
set_shape_flat(card_foot, RGBColor(22, 45, 85), RGBColor(40, 75, 130), 1)
tf_foot = card_foot.text_frame
tf_foot.vertical_anchor = MSO_ANCHOR.MIDDLE
p_foot = tf_foot.paragraphs[0]
p_foot.text = "報告單位：永義房屋專案團隊  |  標的：長億南園 8 樓（德富路 390 號）  |  買方狀況：成熟中年家庭・純自住・已付 5 萬元斡旋現金"
p_foot.font.name = "微軟正黑體"
p_foot.font.size = Pt(12)
p_foot.font.color.rgb = WHITE
p_foot.alignment = PP_ALIGN.CENTER

# ==========================================
# SLIDE 2: 執行摘要 (Executive Summary)
# ==========================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "執行摘要：屋主當下面臨的 4 大關鍵決策局勢")

cards_data_s2 = [
    ("📉 大盤與南區 10 大社區全反轉", "第七波管制後南區老屋成交量暴跌 76.4%！長億城、工學天下、文心及第等 10 大社區全面見頂反轉，部分單價已暴跌破 2 字頭，越晚賣越難賣。", RED, Inches(0.8), Inches(1.5)),
    ("🏢 社區實質 4 戶競售", "591 上 23 筆刊登為房仲重複洗版，全社區實質僅 4 戶在賣！4 樓開 898 萬搶低總、3 樓巨露台平車大降 70 萬，同社區削價競爭激烈。", GOLD, Inches(6.8), Inches(1.5)),
    ("⚡ 南屯大墩人文強烈威脅", "買方同時積極考慮信義房屋推薦之南屯『大墩人文』。南屯地段精華保值、近捷運與好市多，若本戶價格談不攏，買方將立即轉向大墩人文！", SLATE, Inches(0.8), Inches(4.3)),
    ("🎯 把握稀缺中年自住買方", "年輕人只買新小宅，願意買 34 年老屋的僅剩中年家庭。買方一家要自住，已備妥『5 萬元斡旋現金』，出價衡量自身財務極限，錯過難再遇到！", GREEN, Inches(6.8), Inches(4.3))
]

for title, desc, color, left, top in cards_data_s2:
    card = add_card(s2, left, top, Inches(5.7), Inches(2.6))
    
    tb_t = s2.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), Inches(5.3), Inches(0.5))
    tf_t = tb_t.text_frame
    p_t = tf_t.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(16)
    p_t.font.bold = True
    p_t.font.color.rgb = color
    
    tb_d = s2.shapes.add_textbox(left + Inches(0.2), top + Inches(0.7), Inches(5.3), Inches(1.7))
    tf_d = tb_d.text_frame
    tf_d.word_wrap = True
    p_d = tf_d.paragraphs[0]
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(12)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3

# ==========================================
# SLIDE 3: 宏觀大盤實證（南區 53 棟老牌社區走勢圖）
# ==========================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "客觀實證 1：台中南區老牌大盤暴跌 76.4%，價格正式向下反轉")

macro_img = os.path.join(trend_dir, "00_台中市南區25年以上大型社區_大盤匯總走勢圖.png")
if os.path.exists(macro_img):
    s3.shapes.add_picture(macro_img, Inches(0.8), Inches(1.5), width=Inches(7.2))

right_card = add_card(s3, Inches(8.2), Inches(1.5), Inches(4.3), Inches(5.5))
tb_r3 = s3.shapes.add_textbox(Inches(8.4), Inches(1.7), Inches(3.9), Inches(5.1))
tf_r3 = tb_r3.text_frame
tf_r3.word_wrap = True

p_r3_1 = tf_r3.paragraphs[0]
p_r3_1.text = "📊 全南區大盤關鍵數據"
p_r3_1.font.name = "微軟正黑體"
p_r3_1.font.size = Pt(15)
p_r3_1.font.bold = True
p_r3_1.font.color.rgb = NAVY

points_s3 = [
    ("樣本覆蓋全面", "全南區 53 棟 25 年以上大型社區、涵蓋 21,208 戶、累計 7,022 筆實登。"),
    ("歷史見頂反轉", "113 年大盤中位數衝上 26.9 萬/坪歷史極值，第七波信用管制後急速反轉下修至 24.6 萬/坪。"),
    ("成交量雪崩 76.4%", "年成交量自高峰 738 件崩縮至 115 年僅 174 件，市場動能徹底急凍！"),
    ("『量先價行』鐵律", "沒有成交量支撐，老屋價格只有緩跌修正一途，不可再心存追高幻想！")
]

for title, desc in points_s3:
    p_t = tf_r3.add_paragraph()
    p_t.text = f"• {title}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = RED
    p_t.space_before = Pt(8)
    
    p_d = tf_r3.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = TEXT_DARK

# ==========================================
# SLIDE 4: 未來三年走勢預估（緩跌＋老屋貶值）
# ==========================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "客觀實證 2：房市趨冷只會緩跌，房屋年齡越老越便宜")

s4_cards = [
    ("1. 房貸信用管制掐死買氣", "• 央行第七波信用管制全面收緊房貸成數。\n• 銀行鑑價極度保守，滿水位狀況持續延燒。\n• 實質房貸利率站穩 2.5%～3.0%，買方負擔能力大幅縮水，只會持續觀望砍價。", Inches(0.8)),
    ("2. 35 年老屋金融大關縮限", "• 長億南園已 34.7 年，跨越 35 年金融門檻。\n• 銀行內規『屋齡＋貸款年限』嚴格受限，無法貸滿 30/40 年長天期。\n• 買方每月還款本金加重，老屋接手門檻越來越高，未來只剩全現金買方。", Inches(4.8)),
    ("3. 未來三年價格逐年回測", "• 根據南區大盤模型推估：\n  - 116 年下探：21.1 萬/坪\n  - 117 年下探：18.0 萬/坪\n  - 118 年回測歷史支撐：15.8 萬/坪\n• 趁目前 20 幾萬高檔反轉初期賣出，是資產最佳安全落袋窗口！", Inches(8.8))
]

for title, desc, left in s4_cards:
    add_card(s4, left, Inches(1.5), Inches(3.7), Inches(5.3))
    tb = s4.shapes.add_textbox(left + Inches(0.2), Inches(1.7), Inches(3.3), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(14)
    p_t.font.bold = True
    p_t.font.color.rgb = NAVY
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(11)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(10)

# ==========================================
# SLIDE 5 (NEW): 南區 10 大指標老牌社區實登行情全數見頂反轉（對照總表）
# ==========================================
s5_bench = prs.slides.add_slide(blank_layout)
add_header(s5_bench, "客觀實證 3：全南區 10 大指標老牌社區實價登錄全數反轉向下（鐵證總表）")

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
    ("工學天下", "452 戶", "29.0 萬 (113/07)", "19.6 萬 (115/06)", "24.7 ➜ 20.7 萬", "-16.1%", "最新單價暴跌，正式跌破 2 字頭！"),
    ("龍邦綠園道", "322 戶", "28.5 萬 (114/02)", "23.6 萬 (115/08)", "23.6 ➜ 21.3 萬", "-9.5%", "綠園道知名名宅高檔回落、買盤縮手"),
    ("文心高第", "266 戶", "34.5 萬 (113/05)", "30.4 萬 (115/07)", "31.3 ➜ 28.5 萬", "-8.9%", "東興路高樓名廈見頂反轉修正"),
    ("大東家華園", "227 戶", "27.7 萬 (113/03)", "20.3 萬 (115/07)", "24.5 ➜ 23.1 萬", "-5.9%", "學府商圈最新成交逼近 1 字頭！"),
    ("假日半島", "219 戶", "33.2 萬 (113/10)", "24.8 萬 (115/07)", "29.4 ➜ 25.3 萬", "-14.0%", "復興路指標大樓下修逾 14%"),
    ("文心大地", "207 戶", "28.4 萬 (113/07)", "24.1 萬 (115/05)", "26.4 ➜ 22.5 萬", "-15.1%", "文心南路聚落量能雪崩、價格下挫"),
    ("文心及第", "259 戶", "35.7 萬 (113/06)", "18.9 萬 (115/03)", "30.6 ➜ 18.9 萬", "-38.2%", "歷史極值暴跌，最新成交跌破 2 字頭！"),
    ("大安黃金城大廈", "307 戶", "34.3 萬 (114/04)", "25.2 萬 (115/06)", "26.6 ➜ 23.7 萬", "-11.1%", "興大生活圈大案支撐瓦解、買氣急凍"),
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
p_bf2.text = "這 10 個比長億南園規模更大、知名度更高的老牌社區，全部在 113 年見頂後反轉向下！其中工學天下、京城特區、文心及第最新成交更已直接跌破 2 字頭！長億南園絕非個案，此時若不趁高檔反轉初期讓利成交，後續跌幅只會更深！"
p_bf2.font.name = "微軟正黑體"
p_bf2.font.size = Pt(10)
p_bf2.font.color.rgb = TEXT_DARK
p_bf2.space_before = Pt(3)

# ==========================================
# SLIDE 6 (NEW): 指標實證 1【工學商圈大案全面回檔】長億城・工學天下・京城特區
# ==========================================
s6_g1 = prs.slides.add_slide(blank_layout)
add_header(s6_g1, "指標實證 1【工學商圈大案全面回檔】長億城・工學天下・京城特區")

g1_data = [
    ("長億城攬翠區（669戶・屋齡30.9年）", "03_長億城攬翠區陽光區_成交行情走勢圖.png",
     "• 113年高點：29.5 萬/坪 (8樓)\n• 115年最新成交：26.9 萬/坪\n• 均價自 26.3 萬跌至 23.2 萬 (-11.9%)\n• 年成交量自高峰 63 件雪崩至 9 件！\n• 669 戶大社區委售競爭激烈，價格持續下探。"),
    ("工學天下（452戶・屋齡31.7年）", "07_工學天下_成交行情走勢圖.png",
     "• 113年高點：29.0 萬/坪 (9樓)\n• 115年最新成交：19.6 萬/坪 (7樓)\n• 🔥 最新成交正式跌破 2 字頭！\n• 均價自 24.7 萬跌至 20.7 萬 (-16.1%)\n• 買盤對 30 年以上無障礙老屋縮手，殺價猛烈。"),
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
# SLIDE 7 (NEW): 指標實證 2【復興與文心核心聚落反轉】文心高第・假日半島・文心及第
# ==========================================
s7_g2 = prs.slides.add_slide(blank_layout)
add_header(s7_g2, "指標實證 2【復興與文心核心聚落反轉】文心高第・假日半島・文心及第")

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
# SLIDE 8 (NEW): 指標實證 3【園道與大安聚落破底】龍邦綠園道・大東家華園・文心大地・大安黃金城
# ==========================================
s8_g3 = prs.slides.add_slide(blank_layout)
add_header(s8_g3, "指標實證 3【園道與大安聚落破底】龍邦綠園道・大東家華園・文心大地・大安黃金城")

g3_data = [
    ("龍邦綠園道（322戶）", "15_龍邦綠園道_成交行情走勢圖.png",
     "• 高點：28.5 萬 (114年)\n• 最新：23.6 萬 (115/08)\n• 均價：23.6 ➜ 21.3 萬\n• 園道首排支撐鬆動，價格下修 -9.5%。"),
    ("大東家華園（227戶）", "35_大東家華園_成交行情走勢圖.png",
     "• 高點：27.7 萬 (113年)\n• 最新：20.3 萬 (115/07)\n• 🔥 最新成交逼近 1 字頭！\n• 均價：24.5 ➜ 23.1 萬\n• 復興學府生活圈老屋買盤大退潮。"),
    ("文心大地（207戶）", "41_文心大地_成交行情走勢圖.png",
     "• 高點：28.4 萬 (113年)\n• 最新：24.1 萬 (115/05)\n• 均價：26.4 ➜ 22.5 萬\n• 均價下修 -15.1%，量能急凍、多案滯銷。"),
    ("大安黃金城（307戶）", "18_大安黃金城大廈_成交行情走勢圖.png",
     "• 高點：34.3 萬 (114年)\n• 最新：25.2 萬 (115/06)\n• 均價：26.6 ➜ 23.7 萬\n• 均價下修 -11.1%，中興大學生活圈買盤縮手。")
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
# SLIDE 9: 長億南園最新實價登錄走勢（21.4萬低點定錨）
# ==========================================
s9 = prs.slides.add_slide(blank_layout)
add_header(s9, "社區行情定錨：長億南園最新成交已跌至 21.4 萬/坪")

comm_img = os.path.join(trend_dir, "28_長億南園_成交行情走勢圖.png")
if os.path.exists(comm_img):
    s9.shapes.add_picture(comm_img, Inches(0.8), Inches(1.5), width=Inches(7.2))

add_card(s9, Inches(8.2), Inches(1.5), Inches(4.3), Inches(5.5))
tb_s9 = s9.shapes.add_textbox(Inches(8.4), Inches(1.7), Inches(3.9), Inches(5.1))
tf_s9 = tb_s9.text_frame
tf_s9.word_wrap = True

p_s9_h = tf_s9.paragraphs[0]
p_s9_h.text = "🏢 長億南園最新實登實證"
p_s9_h.font.name = "微軟正黑體"
p_s9_h.font.size = Pt(15)
p_s9_h.font.bold = True
p_s9_h.font.color.rgb = NAVY

s9_points = [
    ("115 年 9 月最新成交定錨", "同社區德富路 4 樓，含車位總價僅 825 萬，拆算單價已跌回 21.4 萬/坪（含車均攤僅 22.1 萬）。"),
    ("115 年 4 月成交低點", "同社區 9 樓總價更僅 782 萬，買方出價信心大增。"),
    ("實價登錄更見 20.6 萬紀錄", "同社區 12 樓曾出現單價 20.6 萬/坪成交紀錄，打破屋主 23 萬以上開價幻想。"),
    ("歷史長期均價 13～17 萬", "105～110 年長達六年合理區間為 13～17 萬/坪。前兩年極端高點已過，買方態度堅定理性，絕不追高！")
]

for title, desc in s9_points:
    p_t = tf_s9.add_paragraph()
    p_t.text = f"• {title}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = RED
    p_t.space_before = Pt(8)
    
    p_d = tf_s9.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = TEXT_DARK

# ==========================================
# SLIDE 10: 社區真實 4 戶競售格局（破解23筆洗版）
# ==========================================
s10 = prs.slides.add_slide(blank_layout)
add_header(s10, "社區競爭真相：591 帳面 23 筆為房仲洗版，實質『4 戶』激烈競售")

table_shape10 = s10.shapes.add_table(5, 5, Inches(0.8), Inches(1.6), Inches(11.733), Inches(4.0))
tbl10 = table_shape10.table

tbl10.columns[0].width = Inches(1.8)
tbl10.columns[1].width = Inches(2.2)
tbl10.columns[2].width = Inches(2.2)
tbl10.columns[3].width = Inches(2.5)
tbl10.columns[4].width = Inches(3.033)

headers_s10 = ["戶別", "坪數與格局", "目前開價 / 單價", "房仲刊登狀況", "市場競爭威脅評估"]
for col_idx, h_text in enumerate(headers_s10):
    cell = tbl10.cell(0, col_idx)
    set_cell_bg(cell, NAVY)
    p = cell.text_frame.paragraphs[0]
    p.text = h_text
    p.font.name = "微軟正黑體"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

rows_s10 = [
    ("【本案】8 樓大三房", "42.23 坪 (室內 32 坪)\n3房2廳2衛+機上車", "998 萬\n(23.6 萬/坪)", "9 家房仲重複聯賣洗版！\n歷經三次降價求售", "同社區主力競品，屋主急售心態外溢，開價已被定錨於 998 萬。"),
    ("10 樓高樓三房", "35.61 坪 (室內 27 坪)\n3房2廳2衛+機上車", "958 萬\n(26.9 萬/坪)", "9 家房仲重複聯賣洗版！\n主打免千萬高樓三房", "總價比 8 樓便宜 40 萬，同樣 9 家房仲全力推銷，強烈瓜分看屋客。"),
    ("4 樓大三房", "37.29 坪 (室內 28 坪)\n3房2廳2衛+休旅車", "898 萬\n(24.1 萬/坪)", "中信房屋承辦刊登\n社區最低總價門檻", "開價直接殺入 800 多萬，買方第一眼首選，吸走大部分首購自住客。"),
    ("3 樓巨露台平車戶", "66.28 坪 (含露台31坪)\n3房2廳2衛+坡道平車", "1,258 萬\n(18.9 萬/坪)", "3 家房仲刊登\n近期主動降價 70 萬", "社區極罕見平面車位＋31坪大露台，單價殺至 18.9 萬，讓利誠意十足。")
]

for row_idx, row_data in enumerate(rows_s10, start=1):
    for col_idx, text in enumerate(row_data):
        cell = tbl10.cell(row_idx, col_idx)
        set_cell_bg(cell, ROW_ALT if row_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = text
        p.font.name = "微軟正黑體"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_DARK
        if col_idx in [0, 2]:
            p.font.bold = True

tb_s10_bot = s10.shapes.add_textbox(Inches(0.8), Inches(5.8), Inches(11.733), Inches(1.1))
tf_bot10 = tb_s10_bot.text_frame
p_b10 = tf_bot10.paragraphs[0]
p_b10.text = "💡 房仲致屋主建言：這 4 戶在同一個社區搶同一組買方。『誰願意先合理讓利，誰就能先把數百萬現金抱回家！』若我們堅持 998 萬不放，只會淪為幫 4 樓與 10 樓免費抬轎！"
p_b10.font.name = "微軟正黑體"
p_b10.font.size = Pt(11)
p_b10.font.bold = True
p_b10.font.color.rgb = RED

# ==========================================
# SLIDE 11: 南屯「大墩人文」競品強烈威脅
# ==========================================
s11 = prs.slides.add_slide(blank_layout)
add_header(s11, "外部競爭威脅：買方同時評估南屯『大墩人文』，隨時轉向！")

add_card(s11, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.5), bg_color=WHITE, border_color=SLATE)
tb_l11 = s11.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_l11 = tb_l11.text_frame
tf_l11.word_wrap = True

p_l11_h = tf_l11.paragraphs[0]
p_l11_h.text = "📍 南屯區「大墩人文」極致優勢"
p_l11_h.font.name = "微軟正黑體"
p_l11_h.font.size = Pt(16)
p_l11_h.font.bold = True
p_l11_h.font.color.rgb = NAVY

dadun_points = [
    ("門牌價值碾壓南區", "大墩路精華生活圈，近好市多、迪卡儂、豐樂雕塑公園，地段抗跌保值性遠高於南區。"),
    ("捷運綠線正核心", "步行 8 分鐘直達捷運南屯站與南屯市場，交通與商圈機能完勝德富路。"),
    ("社區規模精緻單純", "大墩人文僅 75 戶，出入單純管理極佳，公基金充沛，未來轉手性強。"),
    ("信義房屋專案極力促成", "信義房屋業務正在積極拉買方下斡旋大墩人文，買方這週末即將做最後決定！")
]

for title, desc in dadun_points:
    p_t = tf_l11.add_paragraph()
    p_t.text = f"✔ {title}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = GREEN
    p_t.space_before = Pt(8)
    
    p_d = tf_l11.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = TEXT_DARK

add_card(s11, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.5), bg_color=RGBColor(254, 249, 231), border_color=GOLD)
tb_r11 = s11.shapes.add_textbox(Inches(7.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_r11 = tb_r11.text_frame
tf_r11.word_wrap = True

p_r11_h = tf_r11.paragraphs[0]
p_r11_h.text = "⚠️ 長億南園必須面對的現實劣勢"
p_r11_h.font.name = "微軟正黑體"
p_r11_h.font.size = Pt(16)
p_r11_h.font.bold = True
p_r11_h.font.color.rgb = RED

ny_disadv = [
    ("無障礙硬傷", "進電梯前有 4～5 階樓梯，推嬰兒車或長輩輪椅極不便，年輕與長輩買方第一眼扣分。"),
    ("機車位難抽", "社區 250 戶但機車位極少，中籤率低，機車日常停放困難。"),
    ("雙衛浴無對外窗", "兩間浴室皆無窗，潮濕問題需仰賴抽風，居住舒適度受限。"),
    ("921 歷史抗性", "曾被列需注意建物，雖已解除列管，但買方心理仍存顧慮。")
]

for title, desc in ny_disadv:
    p_t = tf_r11.add_paragraph()
    p_t.text = f"✘ {title}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = RED
    p_t.space_before = Pt(8)
    
    p_d = tf_r11.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = TEXT_DARK

# ==========================================
# SLIDE 12: 買方家庭財務精算極限與整修預算扣除
# ==========================================
s12 = prs.slides.add_slide(blank_layout)
add_header(s12, "買方出價邏輯：非盲目砍價，而是精算自備款與入住整修的務實極限")

add_card(s12, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.5))
tb_l12 = s12.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_l12 = tb_l12.text_frame
tf_l12.word_wrap = True

p_l12_h = tf_l12.paragraphs[0]
p_l12_h.text = "🛠️ 買方入住必備之整修預算扣除"
p_l12_h.font.name = "微軟正黑體"
p_l12_h.font.size = Pt(15)
p_l12_h.font.bold = True
p_l12_h.font.color.rgb = NAVY

cost_items = [
    ("全室油漆與細部修補", "全室批土油漆、裂縫修補，維持家庭居住質感", "約 12～15 萬元"),
    ("雙衛浴暖風乾燥工程", "雙衛無窗，必須安裝高階暖風機與排風管線", "約 8～10 萬元"),
    ("老舊氣密窗與窗框防水", "主臥與廚房外推陽台防水補強、窗框更新", "約 15～20 萬元"),
    ("空調與燈具系統更新", "全室冷氣管線重新整理、節能空調與燈光配置", "約 18～25 萬元"),
    ("清潔、搬遷與零星工程", "全室細部清潔、門鎖換新、零星木作修繕", "約 8～10 萬元"),
    ("整修預算總計", "買方自備款必須額外扣除的實體支出", "約 60～80 萬元")
]

for name, desc, cost in cost_items:
    p = tf_l12.add_paragraph()
    p.text = f"• {name}：{cost}"
    p.font.name = "微軟正黑體"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = RED if name == "整修預算總計" else TEXT_DARK
    p.space_before = Pt(6)
    
    p_sub = tf_l12.add_paragraph()
    p_sub.text = f"  {desc}"
    p_sub.font.name = "微軟正黑體"
    p_sub.font.size = Pt(9.5)
    p_sub.font.color.rgb = TEXT_MUTED

add_card(s12, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.5))
tb_r12 = s12.shapes.add_textbox(Inches(7.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_r12 = tb_r12.text_frame
tf_r12.word_wrap = True

p_r12_h = tf_r12.paragraphs[0]
p_r12_h.text = "💰 買方財務自備款與月付極限"
p_r12_h.font.name = "微軟正黑體"
p_r12_h.font.size = Pt(15)
p_r12_h.font.bold = True
p_r12_h.font.color.rgb = NAVY

fin_points = [
    ("貸款成數保守上限 7 成", "35 年老屋銀行鑑價保守，自備款至少需 3 成以上，現金壓力極大。"),
    ("自備款硬性扣除裝修費", "買方手頭現金需同時應付 3 成頭期款與 70 萬裝修款，已無多餘彈性追價。"),
    ("房貸利率 2.6% 月付壓力重", "現行利率走高，若總價拉高至 900 萬，每月本息還款將突破 4.2 萬元，超出家庭負擔上限。"),
    ("出價是誠實的成家極限", "買方出價並非刻意砍屋主，而是把所有現金掏出來的『真誠成家底線』！")
]

for title, desc in fin_points:
    p_t = tf_r12.add_paragraph()
    p_t.text = f"✔ {title}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(11.5)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE
    p_t.space_before = Pt(8)
    
    p_d = tf_r12.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = TEXT_DARK

# ==========================================
# SLIDE 13: 買方條件優勢與 5 萬元現金斡旋誠意
# ==========================================
s13 = prs.slides.add_slide(blank_layout)
add_header(s13, "買方特質：稀缺純自住中年買家，已備妥『5 萬元現金斡旋』")

s13_cards = [
    ("1. 成熟穩健自住家庭", "• 一家四口自住需求，看重室內 32 坪大空間。\n• 珍惜屋主水電衛浴整理基礎，非投機客炒作。\n• 工作穩定、財務清白，成交履約最安全！", Inches(0.8), GREEN),
    ("2. 5 萬元現金斡旋在桌上", "• 買方態度堅決，已於店頭支付 5 萬現金斡旋。\n• 絕非隨便看看的假客人，只要屋主同意隨時簽約！\n• 帶真金白銀進場的買方，全南區屈指可數。", Inches(4.8), GOLD),
    ("3. 錯過不再有下一組", "• 年輕世代只買外圍小新屋，完全抗拒老陳屋。\n• 能承接 34 年老屋的自住客群極度稀缺！\n• 一旦這組買方流失，下次見客恐再等一年！", Inches(8.8), RED)
]

for title, desc, left, color in s13_cards:
    add_card(s13, left, Inches(1.5), Inches(3.7), Inches(5.3))
    tb = s13.shapes.add_textbox(left + Inches(0.2), Inches(1.7), Inches(3.3), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(15)
    p_t.font.bold = True
    p_t.font.color.rgb = color
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(11)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(10)

# ==========================================
# SLIDE 14: 屋主資產獲利了結與自住免稅清單
# ==========================================
s14 = prs.slides.add_slide(blank_layout)
add_header(s14, "屋主利益極大化：107年670萬買入持有滿8年，獲利了結＋享400萬免稅！")

add_card(s14, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.5), bg_color=WHITE, border_color=GREEN)
tb_l14 = s14.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_l14 = tb_l14.text_frame
tf_l14.word_wrap = True

p_l14_h = tf_l14.paragraphs[0]
p_l14_h.text = "📊 屋主財務帳面試算（絕對贏家）"
p_l14_h.font.name = "微軟正黑體"
p_l14_h.font.size = Pt(16)
p_l14_h.font.bold = True
p_l14_h.font.color.rgb = GREEN

calc_rows = [
    ("原始購入時間", "民國 107 年 6 月 10 日（實登 #49 筆可考）"),
    ("原始購入總價", "670 萬元（單價僅 15.9 萬/坪）"),
    ("資產持有期間", "整整持有滿 8 年，充分享受居住與資產增值"),
    ("本次建議成交價", "約 700 多萬元（大幅高於原始購入成本）"),
    ("帳面淨獲利", "全數收回 670 萬成本外，現賺數十萬現金！"),
    ("自住免稅優惠", "持有逾6年符合自住規定，享 400 萬免稅額，超額僅課 10%！")
]

for title, val in calc_rows:
    p = tf_l14.add_paragraph()
    p.text = f"• {title}：{val}"
    p.font.name = "微軟正黑體"
    p.font.size = Pt(11)
    p.font.bold = True if "獲利" in title or "免稅" in title else False
    p.font.color.rgb = RED if "獲利" in title else (GREEN if "免稅" in title else TEXT_DARK)
    p.space_before = Pt(8)

add_card(s14, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.5), bg_color=RGBColor(244, 246, 249), border_color=SLATE)
tb_r14 = s14.shapes.add_textbox(Inches(7.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_r14 = tb_r14.text_frame
tf_r14.word_wrap = True

p_r14_h = tf_r14.paragraphs[0]
p_r14_h.text = "🏆 為什麼現在賣是『真正的大贏家』？"
p_r14_h.font.name = "微軟正黑體"
p_r14_h.font.size = Pt(16)
p_r14_h.font.bold = True
p_r14_h.font.color.rgb = NAVY

why_win = [
    ("避開未來三年跌勢", "大盤即將下修至 15～18 萬支撐區，晚賣一年少賺幾十萬。"),
    ("免繳高額房地合一稅", "享受自住 400 萬免稅額，獲利全部是實質進戶頭的現金。"),
    ("資產重新配置至高利商品", "把資金拿去買高股息 ETF 或定存（年息 4～5%），每年穩賺數十萬利息，不用再煩惱老屋維修管線！"),
    ("鎖定 8 年投資回報", "將 670 萬的陳年老屋換成現金，安心享受退休或家庭生活！")
]

for title, desc in why_win:
    p_t = tf_r14.add_paragraph()
    p_t.text = f"✔ {title}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(11.5)
    p_t.font.bold = True
    p_t.font.color.rgb = NAVY
    p_t.space_before = Pt(8)
    
    p_d = tf_r14.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = TEXT_DARK

# ==========================================
# SLIDE 15: 促成簽字逼單話術與特約保障
# ==========================================
s15 = prs.slides.add_slide(blank_layout)
add_header(s15, "促成決策：把握這組已備妥現金的客人，今天簽字落袋為安！")

card_callout = add_card(s15, Inches(0.8), Inches(1.5), Inches(11.733), Inches(3.0), bg_color=RGBColor(234, 250, 241), border_color=GREEN)
tb_s15_c = s15.shapes.add_textbox(Inches(1.1), Inches(1.7), Inches(11.133), Inches(2.6))
tf_s15_c = tb_s15_c.text_frame
tf_s15_c.word_wrap = True

p_s15_h = tf_s15_c.paragraphs[0]
p_s15_h.text = "🏆 房仲收尾逼單最關鍵的一句話（激發屋主落袋為安與急迫感）"
p_s15_h.font.name = "微軟正黑體"
p_s15_h.font.size = Pt(15)
p_s15_h.font.bold = True
p_s15_h.font.color.rgb = GREEN

p_s15_quote = tf_s15_c.add_paragraph()
p_s15_quote.text = (
    "「大哥／大姐，現在市場上光看不買的假客人很多，但願意掏出真金白銀的人少之又少。\n\n"
    "這組自住買方非常有誠意，已經把『5 萬元的現金斡旋』直接擺在我們店頭桌上，隨時等著只要您點頭就正式簽約履約！\n\n"
    "您 107 年買 670 萬，買方出價已經讓您穩穩賺錢出場，又享有 8 年自住免稅優勢，還能完美避開後面南區老屋量縮價跌的風險。\n\n"
    "信義房屋那邊今天還在打電話催買方去談南屯的大墩人文，請大哥／大姐今天就痛快做決定，給我們授權簽字，我們立刻幫您把七百多萬的現金穩穩落袋！」"
)
p_s15_quote.font.name = "微軟正黑體"
p_s15_quote.font.size = Pt(11)
p_s15_quote.font.color.rgb = TEXT_DARK
p_s15_quote.line_spacing = 1.25
p_s15_quote.space_before = Pt(8)

card_agree = add_card(s15, Inches(0.8), Inches(4.7), Inches(11.733), Inches(2.3), bg_color=RGBColor(253, 237, 236), border_color=RED)
tb_s15_a = s15.shapes.add_textbox(Inches(1.1), Inches(4.85), Inches(11.133), Inches(2.0))
tf_s15_a = tb_s15_a.text_frame
tf_s15_a.word_wrap = True

p_s15_ah = tf_s15_a.paragraphs[0]
p_s15_ah.text = "🛡️ 買方斡旋唯一保留特約條款（純粹財務保命條款，其餘抗坑條款全數剔除）"
p_s15_ah.font.name = "微軟正黑體"
p_s15_ah.font.size = Pt(13)
p_s15_ah.font.bold = True
p_s15_ah.font.color.rgb = RED

p_s15_at = tf_s15_a.add_paragraph()
p_s15_at.text = (
    "「本件買賣若買方於簽約後向金融機構申請本標的物房屋貸款，獲核准之貸款金額未達買賣成交總價百分之七十（70%）者"
    "（非因買方個人重大信用瑕疵所致），買方得主張無條件解除本買賣契約（或斡旋契約）。賣方及仲介方應於買方出具銀行審核結果後三日內，"
    "全額無息返還買方已支付之全部斡旋金及價款，雙方互不負違約賠償責任及仲介服務費。」"
)
p_s15_at.font.name = "微軟正黑體"
p_s15_at.font.size = Pt(10.5)
p_s15_at.font.color.rgb = TEXT_DARK
p_s15_at.line_spacing = 1.2
p_s15_at.space_before = Pt(6)

# ==========================================
# SLIDE 16: 封底
# ==========================================
s16 = prs.slides.add_slide(blank_layout)
bg16 = s16.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
set_shape_flat(bg16, DARK_BLUE)

tb16 = s16.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.333), Inches(3.5))
tf16 = tb16.text_frame
tf16.word_wrap = True

p16_1 = tf16.paragraphs[0]
p16_1.text = "抓緊大盤反轉前夕  •  把握唯一誠意自住買盤"
p16_1.font.name = "微軟正黑體"
p16_1.font.size = Pt(28)
p16_1.font.bold = True
p16_1.font.color.rgb = GOLD
p16_1.alignment = PP_ALIGN.CENTER

p16_2 = tf16.add_paragraph()
p16_2.text = "恭祝 屋主資產安全變現・圓滿落袋為安大贏家！"
p16_2.font.name = "微軟正黑體"
p16_2.font.size = Pt(22)
p16_2.font.bold = True
p16_2.font.color.rgb = WHITE
p16_2.alignment = PP_ALIGN.CENTER
p16_2.space_before = Pt(16)

p16_3 = tf16.add_paragraph()
p16_3.text = "永義房屋專案團隊 敬製  |  長億南園 8 樓"
p16_3.font.name = "微軟正黑體"
p16_3.font.size = Pt(13)
p16_3.font.color.rgb = RGBColor(190, 210, 240)
p16_3.alignment = PP_ALIGN.CENTER
p16_3.space_before = Pt(24)

# 儲存
pptx_path = os.path.join(target_dir, "長億南園8樓_房仲向屋主議價說服簡報.pptx")
prs.save(pptx_path)
print(f"🎉 升級版 PowerPoint 簡報已生成 (共 16 頁): {pptx_path}")

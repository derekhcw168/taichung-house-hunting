# -*- coding: utf-8 -*-
"""
生成《長億南園8樓_房仲向屋主議價溝通簡報.pptx》
高品質 16:9 商業簡報，具備圖表、數據卡片、結構對比與精美配色
嚴格落實使用者最新修改與指示：
1. 檔名改為：長億南園8樓_房仲向屋主議價溝通簡報.pptx
2. "實質4戶競售剖析" -> "同社區4戶競售剖析"
3. "同一社區即有4戶同時銷售，競爭激烈。鄰近社區總計也有數十戶在賣。"
4. 買方曾租住南屯「大墩人文」有感情，正由信義房屋積極洽談中（約188戶，室內達31.5坪）
5. 買方願意立即下斡旋金（若屋主願意談，無具體金額）
6. 震後宅對比加入大慶國宅
7. 說服理由 7 裝修部分新增：原本舊有的裝潢、衣櫃、固定書桌，買方都需要拆除清運，重新施作泥作、粉刷、油漆等工程，因此買方也會有額外的成本需要負擔。
8. 屋主獲利出場說帖採用「在較高價位順利成交」、「賣屋收回 670 萬本金加獲利」、「把現金穩穩落袋」、「今天就做決定」
9. 加強近兩年買方稀缺說明與原因
10. 完整呈現 10 大說服理由
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

def add_header(slide, title_text, category="長億南園 8 樓・房仲向屋主議價溝通簡報"):
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
p1_sub1.text = "長億南園 8 樓"
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
p1_desc.text = "【市場大盤反轉急凍・同社區4戶競售剖析・南屯大墩人文競品威脅・掌握稀缺中年自住買盤】"
p1_desc.font.name = "微軟正黑體"
p1_desc.font.size = Pt(15)
p1_desc.font.color.rgb = RGBColor(200, 215, 235)
p1_desc.space_before = Pt(14)

p1_foot = tf1.add_paragraph()
p1_foot.text = "報告單位：永義房屋專案團隊  |  標的：長億南園 8 樓（德富路 390 號）  |  買方狀況：成熟中年家庭・自住・屋主願談可立即下斡旋金"
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
    ("📉 大盤確認反轉下修", "第七波信用管制後，南區 53 棟老牌社區成交量暴跌 76.4%！單價中位數自 26.9 萬見頂回落，價格正式進入『緩跌』格局，近兩年買盤極端稀缺，越晚賣越難賣。", RED, Inches(0.8)),
    ("🏢 同社區 4 戶競售剖析", "同社區即有 4 戶在售，附近社區類似總價更有數十戶在賣！4 樓開 898 萬搶低總、3 樓巨露台平車大降 70 萬，周邊供給大量湧出，削價競爭劇烈。", GOLD, Inches(3.8)),
    ("⚡ 南屯大墩人文強烈威脅", "買方曾租住南屯『大墩人文』社區有感情，目前正由信義房屋積極洽談中。若本案屋主無意願協調溝通，買方將轉向信義房屋購買大墩人文！", SLATE, Inches(6.8)),
    ("🎯 把握稀缺中年自住買方", "近年年輕人只買新成屋小宅，買 34 年老屋的只剩中年家庭。買方純自住，只要屋主願意談，可立即下斡旋金，是目前市場極罕見的實質好買盤！", GREEN, Inches(9.8))
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
# SLIDE 3: 理由 1【南區大盤暴跌 76.4%】
# ==========================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "說服理由 1：南區大盤暴跌 76.4%，價格正式向下反轉")

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
    ("2. 35 年老屋金融大關縮限", "• 長億南園已 34.7 年，跨越 35 年金融門檻。\n• 銀行內規『屋齡＋貸款年限』嚴格受限，無法貸滿 30/40 年長天期。\n• 買方每月還款本金加重，老屋接手門檻越來越高，未來只剩全現金買方。", Inches(4.8), GOLD),
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
# SLIDE 5: 理由 2【實登最新跌至 21.4 萬/坪，9樓僅782萬】
# ==========================================
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "說服理由 2：實登最新成交跌至 21.4 萬/坪，同社區9樓僅782萬無法追高")

cy_img_path = os.path.join(trend_dir, "28_長億南園_成交行情走勢圖.png")
if os.path.exists(cy_img_path):
    s5.shapes.add_picture(cy_img_path, Inches(0.8), Inches(1.5), Inches(6.8), Inches(5.2))

add_card(s5, Inches(7.8), Inches(1.5), Inches(4.733), Inches(5.2))
tb_s5 = s5.shapes.add_textbox(Inches(8.1), Inches(1.7), Inches(4.133), Inches(4.7))
tf_s5 = tb_s5.text_frame
tf_s5.word_wrap = True

p_s5_t = tf_s5.paragraphs[0]
p_s5_t.text = "🏢 長億南園最新實登實證"
p_s5_t.font.name = "微軟正黑體"
p_s5_t.font.size = Pt(15)
p_s5_t.font.bold = True
p_s5_t.font.color.rgb = NAVY

s5_points = [
    ("115 年 9 月最新成交定錨", "同社區德富路 4 樓，含車位總價僅 825 萬，拆算單價已跌回 21.4 萬/坪（含車均攤僅 22.1 萬）。"),
    ("同社區 9 樓總價僅 782 萬", "115 年 4 月成交同社區 9 樓總價更只有 782 萬，買方無法出高價追價！"),
    ("實價登錄更見 20.6 萬紀錄", "同社區 12 樓曾出現單價 20.6 萬/坪成交紀錄，打破屋主 23 萬以上開價幻想。"),
    ("歷史長期均價 13～17 萬", "105～110 年長達六年合理區間為 13～17 萬/坪。前兩年極端高點已過，買方態度堅定理性，無法追高！")
]

for title, desc in s5_points:
    p_p = tf_s5.add_paragraph()
    p_p.text = f"• {title}"
    p_p.font.name = "微軟正黑體"
    p_p.font.size = Pt(12)
    p_p.font.bold = True
    p_p.font.color.rgb = RED if "782" in title or "21.4" in title else SLATE
    p_p.space_before = Pt(8)
    
    p_pd = tf_s5.add_paragraph()
    p_pd.text = desc
    p_pd.font.name = "微軟正黑體"
    p_pd.font.size = Pt(10.5)
    p_pd.font.color.rgb = TEXT_DARK
    p_pd.space_before = Pt(2)

# ==========================================
# SLIDE 6: 理由 3【同社區即有 4 戶競售】
# ==========================================
s6 = prs.slides.add_slide(blank_layout)
add_header(s6, "說服理由 3：同一社區即有 4 戶同時銷售，周邊更有數十戶在賣")

table_shape = s6.shapes.add_table(5, 5, Inches(0.8), Inches(1.5), Inches(11.733), Inches(4.0))
tbl = table_shape.table

tbl.columns[0].width = Inches(2.3)
tbl.columns[1].width = Inches(2.3)
tbl.columns[2].width = Inches(1.8)
tbl.columns[3].width = Inches(2.3)
tbl.columns[4].width = Inches(3.033)

headers_s6 = ["戶別", "坪數與格局", "目前開價 / 單價", "銷售現況", "市場競爭威脅評估"]
for col_idx, h_text in enumerate(headers_s6):
    cell = tbl.cell(0, col_idx)
    set_cell_bg(cell, NAVY)
    p = cell.text_frame.paragraphs[0]
    p.text = h_text
    p.font.name = "微軟正黑體"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

rows_s6 = [
    ("【本案】8 樓大三房", "42.23 坪 (室內 32 坪)\n3房2廳2衛+機上車", "998 萬\n(23.6 萬/坪)", "主力推廣戶別\n歷經三次降價求售", "同社區主力競品，開價已被市場定錨於 998 萬。"),
    ("10 樓高樓三房", "35.61 坪 (室內 27 坪)\n3房2廳2衛+機上車", "958 萬\n(26.9 萬/坪)", "主打免千萬高樓三房\n積極推銷中", "總價比 8 樓便宜 40 萬，強烈瓜分看屋客。"),
    ("4 樓大三房", "37.29 坪 (室內 28 坪)\n3房2廳2衛+休旅車", "898 萬\n(24.1 萬/坪)", "社區最低總價門檻\n承辦經紀人主推", "開價直接殺入 800 多萬，買方第一眼首選，吸走大部分首購自住客。"),
    ("3 樓巨露台平車戶", "66.28 坪 (含露台31坪)\n3房2廳2衛+坡道平車", "1,258 萬\n(18.9 萬/坪)", "近期主動降價 70 萬\n屋主讓利積極", "社區極罕見平面車位＋31坪大露台，單價殺至 18.9 萬，讓利誠意十足。")
]

for row_idx, row_data in enumerate(rows_s6, start=1):
    for col_idx, text in enumerate(row_data):
        cell = tbl.cell(row_idx, col_idx)
        set_cell_bg(cell, ROW_ALT if row_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = text
        p.font.name = "微軟正黑體"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_DARK
        if col_idx in [0, 2]:
            p.font.bold = True

tb_s6_bot = s6.shapes.add_textbox(Inches(0.8), Inches(5.8), Inches(11.733), Inches(1.1))
tf_bot = tb_s6_bot.text_frame
p_b = tf_bot.paragraphs[0]
p_b.text = "💡 房仲核心建言：同社區即有 4 戶在售，附近社區類似總價物件總計更有數十戶同時銷售！「誰先適度讓利，誰就能先把數百萬現金抱回家」！\n堅持高價只會變成替其他物件做免費抬轎。"
p_b.font.name = "微軟正黑體"
p_b.font.size = Pt(12)
p_b.font.bold = True
p_b.font.color.rgb = RED

# ==========================================
# SLIDE 7: 理由 4【921 震前老屋抗性與產調安全疑慮】
# ==========================================
s7 = prs.slides.add_slide(blank_layout)
add_header(s7, "說服理由 4：相同價位可買『921 震後宅』，震前老屋存在安全疑慮")

s7_cards = [
    ("1. 921 震災曾列『需注意建物』紀錄", "• 同社區產調白紙黑字載明：『本物件曾被台中市政府列為需注意建物，雖目前已無列管，但收定後及簽約時勿承諾貸款成數』。\n• 買方是純自住，對震災歷史紀錄非常敏感，存在天然的耐震安全心理抗性。\n• 銀行內部亦有歷史紀錄，鑑價與成數極度保守！", RED, Inches(0.8)),
    ("2. 相同價位市場可選『921 震後宅』", "• 長億南園民國 80 年興建（屋齡 34.7 年），適用 921 震前舊耐震法規。\n• 買方直言：『在目前 800、900 萬價位，市場上完全可以挑到 921 震後興建、耐震法規提升的中古大樓(大慶國宅)，對地震疑慮完全消除！』，買方對大慶國宅亦展現高度興趣。\n• 震前老屋在現行市場上具有無法抹滅的實質折價抗性！", GOLD, Inches(6.8))
]

for title, desc, col, left in s7_cards:
    add_card(s7, left, Inches(1.5), Inches(5.7), Inches(5.2))
    tb = s7.shapes.add_textbox(left + Inches(0.3), Inches(1.8), Inches(5.1), Inches(4.6))
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
# SLIDE 8: 理由 5【年輕人只買新小宅，買 34 年老屋的買家只剩中年家庭】
# ==========================================
s8 = prs.slides.add_slide(blank_layout)
add_header(s8, "說服理由 5：近年年輕人只買新小宅，買 34 年老屋的買家只剩中年家庭")

s8_boxes = [
    ("現代年輕首購族（25～35歲）", "❌ 絕不買 30 幾年陳屋\n• 偏好重劃區預售屋、新成屋、小宅兩房。\n• 重視公設會館、飯店物業管理、電動車充電樁。\n• 看到 34 年無公設外觀老舊大樓，直接轉頭放棄。", Inches(0.8), RED),
    ("中年成熟家庭買方（40～55歲）", "✅ 老牌大樓唯一真實買盤！\n• 重視室內實質大坪數（喜愛本戶室內 32 坪大空間）。\n• 財務務實健全、自備款充裕、珍惜整理過的水電衛浴。\n• 剛好就是本組買方！近兩年買盤極端稀缺，錯過難再尋！", Inches(6.8), GREEN)
]

for title, desc, left, col in s8_boxes:
    add_card(s8, left, Inches(1.5), Inches(5.7), Inches(5.2))
    tb = s8.shapes.add_textbox(left + Inches(0.3), Inches(1.8), Inches(5.1), Inches(4.6))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(16)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(12)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(14)

# ==========================================
# SLIDE 9: 理由 6【買方曾租住南屯『大墩人文』有感情，信義房屋洽談中】
# ==========================================
s9 = prs.slides.add_slide(blank_layout)
add_header(s9, "說服理由 6：買方曾租住南屯大墩人文有感情，信義房屋正積極洽談中")

table_shape_s9 = s9.shapes.add_table(5, 3, Inches(0.8), Inches(1.6), Inches(11.733), Inches(3.8))
tbl_s9 = table_shape_s9.table
tbl_s9.columns[0].width = Inches(2.5)
tbl_s9.columns[1].width = Inches(4.6)
tbl_s9.columns[2].width = Inches(4.633)

headers_s9 = ["評估指標", "南屯區【大墩人文】(約 188 戶・信義積極洽談中)", "南區【長億南園 8 樓】(本案)"]
for col_idx, h in enumerate(headers_s9):
    cell = tbl_s9.cell(0, col_idx)
    set_cell_bg(cell, NAVY)
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.name = "微軟正黑體"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

s9_rows = [
    ("地段價值與增值性", "南屯成熟富人商圈，地段保值抗跌性全台中名列前茅。", "南區老舊市區，受十三期新成屋比價排擠。"),
    ("交通與生活機能", "緊鄰文心森林公園商圈、捷運綠線、南屯好市多(Costco)。", "南區德富路商圈，生活機能便利但無大型好市多與精華捷運。"),
    ("房屋格局與空間", "正三房格局，全社區約 188 戶規模適中，質感佳，室內達 31.5 坪滿意度高。", "室內 32 坪超大空間（本戶唯一最大勝出亮點！）。"),
    ("買方當下態度", "買方曾租住大墩人文有深厚感情，信義房屋積極洽談中！", "若本戶屋主無意願協調溝通，買家只好轉向買大墩人文！")
]

for row_idx, r_data in enumerate(s9_rows, start=1):
    for col_idx, t in enumerate(r_data):
        cell = tbl_s9.cell(row_idx, col_idx)
        set_cell_bg(cell, ROW_ALT if row_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = t
        p.font.name = "微軟正黑體"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_DARK
        if col_idx == 0:
            p.font.bold = True

tb_s9_warn = s9.shapes.add_textbox(Inches(0.8), Inches(5.6), Inches(11.733), Inches(1.3))
tf_w = tb_s9_warn.text_frame
p_w = tf_w.paragraphs[0]
p_w.text = "⚠️ 經紀人急迫預警：買方曾租住大墩人文有深厚感情，信義房屋正積極帶他們洽談！若大哥／大姐無意願協調溝通，買方就只好轉向向信義房屋購買大墩人文了！"
p_w.font.name = "微軟正黑體"
p_w.font.size = Pt(12)
p_w.font.bold = True
p_w.font.color.rgb = RED

# ==========================================
# SLIDE 10: 理由 7【雙衛浴無窗＋舊裝潢清運泥作工程龐大，買方無法追高】
# ==========================================
s10 = prs.slides.add_slide(blank_layout)
add_header(s10, "說服理由 7：雙衛浴皆無對外窗＋舊裝潢清運泥作工程龐大，買方無法追高")

s10_cards = [
    ("屋況客觀限制與必備整修款", "• 肯定屋主 6 年前水電衛浴翻新用心，但已有 6 年自然折舊。\n• 原本舊有裝潢、衣櫃、固定書桌，買方均需拆除清運，重新施作泥作、粉刷、油漆等工程，額外成本負擔沉重！\n• 雙衛浴皆無對外窗（暗房衛浴），通風排濕不良，須加裝獨立抽風暖風乾燥系統。\n• 扣除拆除清運與泥作整修必備支出，買方購屋實質預算已達上限。", Inches(0.8)),
    ("買方無法追高，絕非惡意殺價", "• 買方是純自住家庭，財務預算客觀嚴格。\n• 銀行因 34.7 年高屋齡與 921 歷史，核貸成數極其保守，買方自備款負擔沈重！\n• 扣除拆除清運、泥作粉刷與水電整修成本後，買方已出到自身務實極限，真的無法出高價追價！", Inches(6.8))
]

for title, desc, left in s10_cards:
    add_card(s10, left, Inches(1.5), Inches(5.7), Inches(5.2))
    tb = s10.shapes.add_textbox(left + Inches(0.3), Inches(1.8), Inches(5.1), Inches(4.6))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(15)
    p_t.font.bold = True
    p_t.font.color.rgb = NAVY
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(11.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 11: 理由 8【屋主 107 年 670 萬買入滿 8 年，大賺出場＋享 400 萬免稅額】
# ==========================================
s11 = prs.slides.add_slide(blank_layout)
add_header(s11, "說服理由 8：107 年買 670 萬持有滿 8 年，大賺出場＋享 400 萬免稅額")

table_shape_s11 = s11.shapes.add_table(5, 3, Inches(0.8), Inches(1.6), Inches(11.733), Inches(3.8))
tbl_s11 = table_shape_s11.table
tbl_s11.columns[0].width = Inches(3.2)
tbl_s11.columns[1].width = Inches(4.266)
tbl_s11.columns[2].width = Inches(4.266)

headers_s11 = ["財務 / 稅務項目", "歷史數據與條件", "本次成交預估結果"]
for col_idx, h in enumerate(headers_s11):
    cell = tbl_s11.cell(0, col_idx)
    set_cell_bg(cell, NAVY)
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.name = "微軟正黑體"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

s11_rows = [
    ("屋主當初買入成本", "民國 107 年 6 月購入：670 萬 (15.9 萬/坪)", "本金 670 萬 100% 全數收回"),
    ("產權持有時間", "107 年 7 月登記至今：已整整滿 8 年！", "已度過房地合一高稅率重課期"),
    ("成交實質獲利評估", "在較高價位順利成交", "完全覆蓋成本與裝修折舊，淨賺數十萬現金！"),
    ("房地合一自住免稅額", "持有滿 8 年，符合自住設籍滿 6 年條件", "享有 400 萬免稅額，獲利免繳稅，全額現金落袋！")
]

for row_idx, r_data in enumerate(s11_rows, start=1):
    for col_idx, t in enumerate(r_data):
        cell = tbl_s11.cell(row_idx, col_idx)
        set_cell_bg(cell, ROW_ALT if row_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = t
        p.font.name = "微軟正黑體"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_DARK
        if col_idx == 0:
            p.font.bold = True

tb_s11_p = s11.shapes.add_textbox(Inches(0.8), Inches(5.6), Inches(11.733), Inches(1.3))
tf_p11 = tb_s11_p.text_frame
p_p11 = tf_p11.paragraphs[0]
p_p11.text = "🏆 獲利了結論點：屋主買 670 萬已享受 8 年居住增值，今天賣出保證倒賺幾十萬現金且不用繳房地合一稅！\n在房市反轉向下之際，及時把現金抱在身上，比繼續套在 35 年老屋承擔跌價風險聰明太多！"
p_p11.font.name = "微軟正黑體"
p_p11.font.size = Pt(12)
p_p11.font.bold = True
p_p11.font.color.rgb = GREEN

# ==========================================
# SLIDE 12: 理由 9 & 10【十三期新案磁吸比價＋高利環境收回現金創造被動收益】
# ==========================================
s12 = prs.slides.add_slide(blank_layout)
add_header(s12, "說服理由 9 & 10：十三期新案夾擊磁吸，現金落袋理財勝過套牢老屋")

s12_cards = [
    ("理由 9：十三期新案交屋潮磁吸比價", "• 南區緊鄰十三期重劃區，周邊大量現代化新成屋陸續完工交屋。\n• 具備高預算的買盤全面被吸走，對比出 34 年老舊大樓在管理、公設與外觀上的巨大折舊落差。\n• 二手老屋市場邊緣化日益嚴重，現在不賣，未來的買家只會越來越少！", Inches(0.8), GOLD),
    ("理由 10：收回數百萬現金創造穩定收益", "• 目前央行利率與金融收益處於歷史高檔。\n• 賣屋收回 670 萬本金加獲利，投入高股息或穩健定存，每年穩領 20～30 萬被動收入！\n• 完全免除老屋折舊、房屋稅、地價稅與管理費，資金效益遠勝套牢在緩跌老屋！", Inches(6.8), GREEN)
]

for title, desc, left, col in s12_cards:
    add_card(s12, left, Inches(1.5), Inches(5.7), Inches(5.2))
    tb = s12.shapes.add_textbox(left + Inches(0.3), Inches(1.8), Inches(5.1), Inches(4.6))
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
    p_d.space_before = Pt(14)

# ==========================================
# SLIDE 13: 收尾逼單促成簽字
# ==========================================
s13 = prs.slides.add_slide(blank_layout)
add_header(s13, "收尾逼單促成：『只要屋主願意談，買方願意立即下斡旋金！』")

add_card(s13, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.2), RGBColor(234, 250, 241), GREEN)
tb_s13 = s13.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(11.0), Inches(4.6))
tf_s13 = tb_s13.text_frame
tf_s13.word_wrap = True

p_s13_1 = tf_s13.paragraphs[0]
p_s13_1.text = "🎯 房仲逼單最關鍵的一句話（激發屋主落袋為安與急迫感）："
p_s13_1.font.name = "微軟正黑體"
p_s13_1.font.size = Pt(16)
p_s13_1.font.bold = True
p_s13_1.font.color.rgb = GREEN

p_s13_2 = tf_s13.add_paragraph()
p_s13_2.text = (
    "「大哥／大姐，現在市場上光看不買的假客人很多，但真正有誠意的買方少之又少。\n\n"
    "這組自住買方非常有誠意：只要大哥／大姐願意談，買方願意立即下斡旋金，隨時配合我們正式簽約履約！\n\n"
    "您 107 年買 670 萬，現在成交已經讓您穩穩賺錢出場，又享有滿 8 年自住 400 萬免稅優勢，還能完美避開後面南區老屋量縮價跌與新案夾擊的風險。\n\n"
    "信義房屋那邊今天還在積極催買方去簽南屯的大墩人文，請大哥／大姐今天就做決定，給我們授權簽字，我們立刻幫您把現金穩穩落袋！」"
)
p_s13_2.font.name = "微軟正黑體"
p_s13_2.font.size = Pt(13)
p_s13_2.font.color.rgb = TEXT_DARK
p_s13_2.line_spacing = 1.3
p_s13_2.space_before = Pt(14)

# ==========================================
# SLIDE 14: 封底
# ==========================================
s14 = prs.slides.add_slide(blank_layout)
bg14 = s14.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
set_shape_flat(bg14, DARK_BLUE)

tb14 = s14.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.333), Inches(3.5))
tf14 = tb14.text_frame
tf14.word_wrap = True

p14_1 = tf14.paragraphs[0]
p14_1.text = "抓緊大盤反轉前夕  •  把握唯一誠意自住買盤"
p14_1.font.name = "微軟正黑體"
p14_1.font.size = Pt(28)
p14_1.font.bold = True
p14_1.font.color.rgb = GOLD
p14_1.alignment = PP_ALIGN.CENTER

p14_2 = tf14.add_paragraph()
p14_2.text = "恭祝 屋主資產安全變現・圓滿落袋為安大贏家！"
p14_2.font.name = "微軟正黑體"
p14_2.font.size = Pt(22)
p14_2.font.bold = True
p14_2.font.color.rgb = WHITE
p14_2.alignment = PP_ALIGN.CENTER
p14_2.space_before = Pt(16)

p14_3 = tf14.add_paragraph()
p14_3.text = "永義房屋專案團隊 敬製  |  長億南園 8 樓"
p14_3.font.name = "微軟正黑體"
p14_3.font.size = Pt(13)
p14_3.font.color.rgb = RGBColor(190, 210, 240)
p14_3.alignment = PP_ALIGN.CENTER
p14_3.space_before = Pt(24)

# 儲存新檔名
new_pptx_path = os.path.join(target_dir, "長億南園8樓_房仲向屋主議價溝通簡報.pptx")
prs.save(new_pptx_path)
print(f"PowerPoint 簡報已生成: {new_pptx_path}")

# 同步存為 .ppt 結尾相容檔名（若使用者直接以 ppt 搜尋）
new_ppt_compat = os.path.join(target_dir, "長億南園8樓_房仲向屋主議價溝通簡報.ppt")
prs.save(new_ppt_compat)

# 刪除舊的「說服簡報」檔名，以保持目錄整潔
old_pptx = os.path.join(target_dir, "長億南園8樓_房仲向屋主議價說服簡報.pptx")
if os.path.exists(old_pptx):
    try:
        os.remove(old_pptx)
        print(f"舊簡報檔已刪除更名: {old_pptx}")
    except Exception as e:
        pass

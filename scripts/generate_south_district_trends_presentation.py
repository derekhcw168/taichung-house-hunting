# -*- coding: utf-8 -*-
"""
生成《台中市南區老牌大型社區實價登錄趨勢與後續三年下跌預估分析.pptx》
納入：
1. 全南區 25 年以上老牌大型社區大盤價量走勢圖與歷史週期趨勢分析
2. 後續三年（116～118年）每年預估下跌 7.5% 階梯式走勢模型
3. 價格走勢下跌的四大維度成因深度剖析（供需面、產業面、市場面、金融面）
4. 全南區 10 大指標老牌大社區的個別價量走勢圖與年跌 7.5% 獨立預估數據卡片
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
NAVY = RGBColor(15, 32, 67)         # #0F2043 封面主底色
PRIMARY_BLUE = RGBColor(27, 54, 93) # #1B365D 標題深藍
SLATE = RGBColor(44, 94, 138)       # #2C5E8A 次藍
ACCENT_CYAN = RGBColor(14, 165, 233)# #0EA5E9 數據藍
LIGHT_BG = RGBColor(248, 250, 252)  # #F8FAFC 投影片背景
WHITE = RGBColor(255, 255, 255)
CARD_BG = RGBColor(255, 255, 255)
BORDER_GRAY = RGBColor(226, 232, 240)
ROW_ALT = RGBColor(248, 250, 252)
GOLD = RGBColor(217, 140, 20)       # #D98C14 強調金
RED = RGBColor(192, 57, 43)         # #C0392B 警示紅
GREEN = RGBColor(39, 174, 96)       # #27AE60 支撐綠
TEXT_DARK = RGBColor(30, 41, 59)    # #1E293B
TEXT_MUTED = RGBColor(100, 116, 139)# #64748B

base_dir = r"d:\GoogleDrive\01_個人_家庭與生活\02_家庭_安莉的資料\小寶的買房子資料"
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

def add_header(slide, title_text, category="台中市南區 25 年以上老牌大社區・實價登錄趨勢與後續三年預估專題報告"):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, LIGHT_BG)
    
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    set_shape_flat(top_bar, PRIMARY_BLUE)

    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.32), Inches(5.8), Inches(0.32))
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

    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(11.7), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "微軟正黑體"
    p.font.size = Pt(19)
    p.font.bold = True
    p.font.color.rgb = PRIMARY_BLUE

def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=BORDER_GRAY):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    set_shape_flat(card, bg_color, border_color, 1)
    return card

# 10 大社區數據（含年跌 -7.5% 獨立試算）
bench_data = [
    {
        'rank': 1,
        'label': '長億城攬翠區',
        'name': '長億城攬翠區陽光區',
        'households': 669,
        'age': 30.9,
        'build_date': '1995-11-30',
        'address': '工學路116號',
        'peak_p': 29.5,
        'peak_date': '113年06月 (8樓)',
        'latest_p': 26.9,
        'latest_date': '115年06月 (10樓)',
        'avg113': 26.3,
        'avg115': 23.2,
        'drop_pct': -11.9,
        'p116': 21.5,
        'p117': 19.9,
        'p118': 18.4,
        'cum_drop': -20.7,
        'supp': '13.8～16.1 萬/坪',
        'img': '03_長億城攬翠區陽光區_成交行情走勢圖.png',
        'desc': '工學商圈龍頭大社區。實登在 113 年衝上 29.5 萬後反轉回檔，115 年均價下修至 23.2 萬。年跌 7.5% 預估下，117 年將跌破 2 字頭至 19.9 萬，118 年探至 18.4 萬。669 戶龐大存量賣壓與二手削價競爭劇烈。'
    },
    {
        'rank': 2,
        'label': '工學天下',
        'name': '工學天下',
        'households': 452,
        'age': 31.7,
        'build_date': '1995-01-23',
        'address': '工學北路135號',
        'peak_p': 29.0,
        'peak_date': '113年07月 (9樓)',
        'latest_p': 19.6,
        'latest_date': '115年06月 (7樓)',
        'avg113': 24.7,
        'avg115': 20.7,
        'drop_pct': -16.1,
        'p116': 19.1,
        'p117': 17.7,
        'p118': 16.4,
        'cum_drop': -20.8,
        'supp': '13.6～16.5 萬/坪',
        'img': '07_工學天下_成交行情走勢圖.png',
        'desc': '屋齡 31.7 年，與長億南園條件極為相近。113 年高點達 29.0 萬，115 年 6 月最新成交直接摜破 2 字頭暴跌至 19.6 萬！年跌 7.5% 預估 116 年下探 19.1 萬，118 年探至 16.4 萬。老屋限貸與缺乏無障礙設施使買盤縮手。'
    },
    {
        'rank': 3,
        'label': '龍邦綠園道',
        'name': '龍邦綠園道',
        'households': 322,
        'age': 33.4,
        'build_date': '1993-05-19',
        'address': '南和路197號',
        'peak_p': 28.5,
        'peak_date': '114年02月 (5樓)',
        'latest_p': 23.6,
        'latest_date': '115年08月 (16樓)',
        'avg113': 23.6,
        'avg115': 21.3,
        'drop_pct': -9.5,
        'p116': 19.7,
        'p117': 18.2,
        'p118': 16.9,
        'cum_drop': -20.7,
        'supp': '13.8～16.8 萬/坪',
        'img': '15_龍邦綠園道_成交行情走勢圖.png',
        'desc': '南和路綠園道首排景觀名宅。114 年 28.5 萬高檔過後買盤急凍，115 年成交已降至 23.6 萬。年跌 7.5% 預估 116 年將跌入 1 字頭（19.7萬），118 年探至 16.9 萬。大盤下行期景觀溢價亦全面收斂。'
    },
    {
        'rank': 4,
        'label': '文心高第',
        'name': '文心高第',
        'households': 266,
        'age': 24.9,
        'build_date': '2001-11-20',
        'address': '東興路一段29號',
        'peak_p': 34.5,
        'peak_date': '113年05月 (12樓)',
        'latest_p': 30.4,
        'latest_date': '115年07月 (2樓)',
        'avg113': 31.3,
        'avg115': 28.5,
        'drop_pct': -8.9,
        'p116': 26.4,
        'p117': 24.4,
        'p118': 22.6,
        'cum_drop': -20.7,
        'supp': '16.5～20.0 萬/坪',
        'img': '26_文心高第_成交行情走勢圖.png',
        'desc': '東興路高樓名廈，屋齡約 24.9 年。113 年突破 34 萬天花板後，115 年均價回跌至 28.5 萬。年跌 7.5% 預估 116 年回修至 26.4 萬，118 年下修至 22.6 萬。買方對 3 字頭老屋抗拒強烈，委售週期大幅拉長。'
    },
    {
        'rank': 5,
        'label': '大東家華園',
        'name': '大東家華園',
        'households': 227,
        'age': 31.2,
        'build_date': '1995-07-28',
        'address': '復興路二段30巷15號',
        'peak_p': 27.7,
        'peak_date': '113年03月 (4樓)',
        'latest_p': 20.3,
        'latest_date': '115年07月 (3樓)',
        'avg113': 24.5,
        'avg115': 23.1,
        'drop_pct': -5.9,
        'p116': 21.4,
        'p117': 19.8,
        'p118': 18.3,
        'cum_drop': -20.8,
        'supp': '13.5～16.2 萬/坪',
        'img': '35_大東家華園_成交行情走勢圖.png',
        'desc': '復興路與學府商圈老牌大樓。實登最新 115 年 7 月成交殺至 20.3 萬/坪逼近 1 字頭！年跌 7.5% 預估 117 年摜破 2 字頭至 19.8 萬，118 年跌至 18.3 萬。傳統市區老屋缺乏新題材，價格加速回測歷史支撐。'
    },
    {
        'rank': 6,
        'label': '假日半島',
        'name': '假日半島',
        'households': 219,
        'age': 32.4,
        'build_date': '1994-05-18',
        'address': '復興路二段71巷21弄31號',
        'peak_p': 33.2,
        'peak_date': '113年10月 (9樓)',
        'latest_p': 24.8,
        'latest_date': '115年07月 (5樓)',
        'avg113': 29.4,
        'avg115': 25.3,
        'drop_pct': -14.0,
        'p116': 23.4,
        'p117': 21.6,
        'p118': 20.0,
        'cum_drop': -20.9,
        'supp': '14.0～17.0 萬/坪',
        'img': '38_假日半島_成交行情走勢圖.png',
        'desc': '復興路二段指名大樓。113 年底曾見 33.2 萬極值，但 115 年均價重挫 14.0% 至 25.3 萬，最新成交 24.8 萬。年跌 7.5% 預估 116 年下修至 23.4 萬，118 年跌至 20.0 萬保衛戰。成交量能萎縮逾八成。'
    },
    {
        'rank': 7,
        'label': '文心大地',
        'name': '文心大地',
        'households': 207,
        'age': 27.4,
        'build_date': '1999-05-18',
        'address': '文心南十路380號',
        'peak_p': 28.4,
        'peak_date': '113年07月 (12樓)',
        'latest_p': 24.1,
        'latest_date': '115年05月 (10樓)',
        'avg113': 26.4,
        'avg115': 22.5,
        'drop_pct': -15.1,
        'p116': 20.8,
        'p117': 19.3,
        'p118': 17.8,
        'cum_drop': -20.9,
        'supp': '13.5～16.5 萬/坪',
        'img': '41_文心大地_成交行情走勢圖.png',
        'desc': '文心南路商圈指標建案。113 年均價 26.4 萬，115 年均價滑落至 22.5 萬，跌幅達 15.1%！年跌 7.5% 預估 117 年將跌破 2 字頭至 19.3 萬，118 年探至 17.8 萬。買方受信用管制壓抑，出價高度謹慎。'
    },
    {
        'rank': 8,
        'label': '文心及第',
        'name': '文心及第(文心南路)',
        'households': 259,
        'age': 27.3,
        'build_date': '1999-06-12',
        'address': '文心南路846號',
        'peak_p': 35.7,
        'peak_date': '113年06月 (10樓)',
        'latest_p': 18.9,
        'latest_date': '115年03月 (14樓)',
        'avg113': 30.6,
        'avg115': 18.9,
        'drop_pct': -38.2,
        'p116': 17.5,
        'p117': 16.2,
        'p118': 15.0,
        'cum_drop': -20.6,
        'supp': '14.0～17.5 萬/坪',
        'img': '27_文心及第文心南路_成交行情走勢圖.png',
        'desc': '全南區最震撼之破底案例！113 年 6 月創 35.7 萬天價，但 115 年 3 月 14 樓頂樓次高戶直接暴跌至 18.9 萬/坪成交，摜破 2 字頭！年跌 7.5% 預估 116 年 17.5 萬，118 年探至 15.0 萬歷史支撐底線。'
    },
    {
        'rank': 9,
        'label': '大安黃金城大廈',
        'name': '大安黃金城大廈',
        'households': 307,
        'age': 32.2,
        'build_date': '1994-07-27',
        'address': '忠孝路46號',
        'peak_p': 34.3,
        'peak_date': '114年04月 (6樓)',
        'latest_p': 25.2,
        'latest_date': '115年06月 (2樓)',
        'avg113': 26.6,
        'avg115': 23.7,
        'drop_pct': -11.1,
        'p116': 21.9,
        'p117': 20.3,
        'p118': 18.8,
        'cum_drop': -20.7,
        'supp': '13.0～15.8 萬/坪',
        'img': '18_大安黃金城大廈_成交行情走勢圖.png',
        'desc': '興大生活圈大案。114 年短暫衝上 34.3 萬後熄火，115 年成交下挫至 25.2 萬，均價跌 11.1%。年跌 7.5% 預估 116 年 21.9 萬，118 年跌破 2 字頭至 18.8 萬。貸款成數受限與維修老化問題持續浮現。'
    },
    {
        'rank': 10,
        'label': '京城特區',
        'name': '京城特區',
        'households': 388,
        'age': 29.1,
        'build_date': '1997-09-02',
        'address': '學府路168巷3號',
        'peak_p': 25.9,
        'peak_date': '113年03月 (18樓)',
        'latest_p': 18.6,
        'latest_date': '115年05月 (3樓)',
        'avg113': 22.4,
        'avg115': 20.9,
        'drop_pct': -6.9,
        'p116': 19.3,
        'p117': 17.9,
        'p118': 16.5,
        'cum_drop': -21.1,
        'supp': '13.2～16.0 萬/坪',
        'img': '11_京城特區_成交行情走勢圖.png',
        'desc': '興大學府路 388 戶大社區。113 年高點 25.9 萬，115 年 5 月最新成交直接殺至 18.6 萬見 1 字頭！年跌 7.5% 預估 116 年續探 19.3 萬，118 年探至 16.5 萬。戶數多且賣盤重疊，買氣急凍加速價格尋底。'
    }
]

# ==========================================
# SLIDE 1: 封面
# ==========================================
s1 = prs.slides.add_slide(blank_layout)
bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
set_shape_flat(bg1, NAVY)

dec = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.8), Inches(0.15), Inches(3.2))
set_shape_flat(dec, GOLD)

tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(1.6), Inches(11.2), Inches(3.5))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1 = tf1.paragraphs[0]
p1.text = "台中市南區 25 年以上老牌大型社區"
p1.font.name = "微軟正黑體"
p1.font.size = Pt(24)
p1.font.bold = True
p1.font.color.rgb = GOLD

p2 = tf1.add_paragraph()
p2.text = "實價登錄趨勢與後續三年下跌預估分析報告"
p2.font.name = "微軟正黑體"
p2.font.size = Pt(32)
p2.font.bold = True
p2.font.color.rgb = WHITE
p2.space_before = Pt(8)

p3 = tf1.add_paragraph()
p3.text = "【大盤量縮 76.4% 轉折確立 ｜ 每年下修 7.5% 趨勢模型 ｜ 10 大指標社區深度剖析 ｜ 供需・產業・市場・金融四大成因全景解析】"
p3.font.name = "微軟正黑體"
p3.font.size = Pt(13)
p3.font.color.rgb = RGBColor(190, 210, 240)
p3.space_before = Pt(12)

card_foot = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.6), Inches(11.733), Inches(1.1))
set_shape_flat(card_foot, RGBColor(22, 45, 85), RGBColor(40, 75, 130), 1)
tf_foot = card_foot.text_frame
tf_foot.vertical_anchor = MSO_ANCHOR.MIDDLE
p_foot = tf_foot.paragraphs[0]
p_foot.text = "數據範疇：全南區 53 棟 25 年以上大型集合住宅社區（總計 21,417 戶、平均屋齡 31.0 年） ｜ 實登總計：10,896 筆 ｜ 分析日期：2026 年 10 月"
p_foot.font.name = "微軟正黑體"
p_foot.font.size = Pt(11)
p_foot.font.color.rgb = WHITE
p_foot.alignment = PP_ALIGN.CENTER

# ==========================================
# SLIDE 2: 執行摘要 (Executive Summary)
# ==========================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "執行摘要：南區老牌大社區房市大盤 4 大核心研判結論")

cards_data_s2 = [
    ("📉 大盤轉折確立・量縮 76.4%", 
     "全南區 53 棟老牌大案於 113 年衝上 26.9 萬/坪天花板後見頂，央行第七波信用管制後，115 年成交量自年均 738 件崩縮至 174 件（暴跌 76.4%），『量先價行』確立市場全面反轉。", 
     RED, Inches(0.8), Inches(1.5)),
    ("📉 未來三年每年下修 7.5%", 
     "建立理性軟著陸預估模型，以 115 年均價為基準按年跌 7.5% 階梯式修正：大盤中位數預估 116 年 22.8 萬 ➜ 117 年 21.1 萬 ➜ 118 年 19.5 萬，三年累計跌幅達 20.6%，回測歷史合理支撐帶。", 
     GOLD, Inches(6.8), Inches(1.5)),
    ("🏢 十大指標建案接連破底", 
     "長億城、工學天下、文心及第、京城特區等南區 10 大標竿社區全數高檔下挫！其中工學天下(19.6萬)、文心及第(18.9萬)、京城特區(18.6萬)最新成交已直接摜破 2 字頭，比價定錨效應確立。", 
     SLATE, Inches(0.8), Inches(4.3)),
    ("🎯 四大成因共振拉動下行", 
     "【供需面】2.1萬戶庫存賣壓與新成屋交屋潮；【產業面】薪資所得停滯與老屋翻修成本暴增；【市場面】追價心理破滅與流動性枯竭；【金融面】央行重拳與 30/35 年老屋貸款年限腰斬。", 
     GREEN, Inches(6.8), Inches(4.3))
]

for title, desc, color, left, top in cards_data_s2:
    card = add_card(s2, left, top, Inches(5.7), Inches(2.6))
    tb_t = s2.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), Inches(5.3), Inches(0.5))
    tf_t = tb_t.text_frame
    p_t = tf_t.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(15.5)
    p_t.font.bold = True
    p_t.font.color.rgb = color
    
    tb_d = s2.shapes.add_textbox(left + Inches(0.2), top + Inches(0.7), Inches(5.3), Inches(1.75))
    tf_d = tb_d.text_frame
    tf_d.word_wrap = True
    p_d = tf_d.paragraphs[0]
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(11)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.35

# ==========================================
# SLIDE 3: 大盤實價登錄全景圖與歷史週期分析
# ==========================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "大盤全景解構：南區 53 棟老牌大社區實登價量全圖與歷史週期剖析")

macro_img = os.path.join(trend_dir, "00_台中市南區25年以上大型社區_大盤匯總走勢圖.png")
if os.path.exists(macro_img):
    s3.shapes.add_picture(macro_img, Inches(0.8), Inches(1.5), width=Inches(7.2))

add_card(s3, Inches(8.2), Inches(1.5), Inches(4.3), Inches(5.5))
tb_s3 = s3.shapes.add_textbox(Inches(8.4), Inches(1.7), Inches(3.9), Inches(5.1))
tf_s3 = tb_s3.text_frame
tf_s3.word_wrap = True

p_s3_h = tf_s3.paragraphs[0]
p_s3_h.text = "📊 南區老牌社區大盤四大週期剖析"
p_s3_h.font.name = "微軟正黑體"
p_s3_h.font.size = Pt(14)
p_s3_h.font.bold = True
p_s3_h.font.color.rgb = PRIMARY_BLUE

cycle_points = [
    ("101～106年【早期溫和盤整期】", "全區均價落在 11～14 萬/坪，成交量平穩扎實，老屋以低總價高性價比為主力客群。", SLATE),
    ("107～112年【低利資金推升主升期】", "受全球寬鬆熱錢與通膨預期帶動，單價突破 20 萬大關，全區年度成交突破 700 件高峰。", PRIMARY_BLUE),
    ("113～115年【見頂回落量縮急凍期】", "113 年見 26.9 萬歷史天花板，央行第七波信用管制後量縮逾 76%，115 年中位數下修至 24.6 萬。", RED),
    ("116～118年【年跌7.5%溫和下行期】", "預估每年按 -7.5% 階梯式修正，118 年下探 19.5 萬，回測歷史長期價值支撐帶（15.2～18.0萬）。", GOLD)
]

for title, desc, col in cycle_points:
    p_t = tf_s3.add_paragraph()
    p_t.text = f"• {title}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(11)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    p_t.space_before = Pt(8)
    
    p_d = tf_s3.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.25

# ==========================================
# SLIDE 4: 造成價格走勢下跌的原因分析（四大維度架構總覽）
# ==========================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "成因全景架構：促使老牌大型社區價格持續下修的四大驅動力")

drivers_s4 = [
    ("🏗️ 整體供需面", "龐大老屋存量與新成交夾殺", 
     "• 53 棟老案庫存逾 2.1 萬戶，轉冷後委售件數暴增\n• 周邊烏日高鐵、13期新成屋交屋潮強烈瓜分客源\n• 年輕首購族排斥老陳屋，自住買盤面臨世代斷層", 
     Inches(0.8), PRIMARY_BLUE),
    ("🏭 產業與經濟面", "實質所得停滯與翻修成本暴增", 
     "• 非高薪科技聚落，受薪族薪資停滯無力追價\n• 室內翻修工料大漲 30~50%，翻新費反噬老屋殘值\n• 成熟商圈發展飽和，缺乏重磅新增建設支撐溢價", 
     Inches(3.8), SLATE),
    ("📉 市場與交易心理", "量先價行與預期追價破滅", 
     "• 成交量暴跌 76.4%，流動性枯竭迫使賣方防線崩潰\n• 買方心理從『恐慌搶買』轉為『理性觀望等更跌』\n• 工學天下等指標社區摜破 2 字頭引發比價連鎖效應", 
     Inches(6.8), GOLD),
    ("🏦 金融與信用管制", "央行重拳與老屋貸款大門關閉", 
     "• 央行第七波管制砍房貸成數，第二戶限貸 5 成無寬限\n• 銀行法 72-2 滿水位警戒，房貸利率實質站上 3%\n• 30/35 年老屋貸款年限被砍至 20/15 年，月付金翻倍", 
     Inches(9.8), RED)
]

for title, sub, desc, left, col in drivers_s4:
    add_card(s4, left, Inches(1.5), Inches(2.78), Inches(5.5), border_color=col)
    
    top_bar_c = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.5), Inches(2.78), Inches(0.1))
    set_shape_flat(top_bar_c, col)
    
    tb = s4.shapes.add_textbox(left + Inches(0.12), Inches(1.7), Inches(2.54), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(15)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_s = tf.add_paragraph()
    p_s.text = sub
    p_s.font.name = "微軟正黑體"
    p_s.font.size = Pt(10)
    p_s.font.color.rgb = TEXT_MUTED
    p_s.space_before = Pt(2)
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(12)

# ==========================================
# SLIDE 5: 深度成因剖析 1 —— 整體供需面
# ==========================================
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "深度成因剖析 1【整體供需面】：龐大老屋存量賣壓、新成屋交屋潮夾殺與買盤斷層")

sd_cards = [
    ("1. 逾 2.1 萬戶老屋龐大存量，社區內部削價競爭",
     "• 供給母體龐大：南區僅這 53 棟 25 年以上老牌大案，總戶數就高達 21,417 戶。像長億城、工學天下等案動輒三四百戶甚至六百戶以上。\n"
     "• 委售重疊競爭：房市降溫後，每棟社區平均在售物件由 2~3 戶激增至 6~12 戶。各房仲多品牌洗版聯賣，造成同社區屋主彼此搶客削價求現。\n"
     "• 去化時間劇增：平均委售去化天數由過去的 45 天大幅拉長至 180 天以上，中古供給嚴重過剩。", Inches(0.8)),
    ("2. 鄰近新興重劃區交屋潮分流，優質自住買盤被吸走",
     "• 新重劃區夾殺：南區緊鄰烏日高鐵特區、13 期重劃區等新興熱區，近期迎來龐大預售完工交屋潮。\n"
     "• 產品代差顯著：新大樓採用最新 921 震後結構耐震法規、配備坡道平面車位、多功能接待會館與無障礙全平動線，大幅降低維護困擾。\n"
     "• 預算排擠效應：總價 1,000～1,400 萬區間，買方寧可選擇外圍新成屋 2 房或小 3 房，直接吸走老陳屋的潛在買客。", Inches(4.8)),
    ("3. 買方世代斷層與少子化，願意接盤中古陳屋者稀缺",
     "• 青年首購偏好新小宅：新一代首購族對 30 年以上老舊管線、斑駁外觀與無車位抗拒感極強，幾乎不考慮中古老陳屋。\n"
     "• 接盤主力嚴重萎縮：目前願意承接 30 年老陳屋的買方，幾乎僅剩具備實際家庭成員居住需求的中年自住換屋客。\n"
     "• 買方結構失衡：買盤基底急遽縮小，形成『賣方急售、買方稀缺』的嚴重供過於求結構性失衡。", Inches(8.8))
]

for title, desc, left in sd_cards:
    add_card(s5, left, Inches(1.5), Inches(3.7), Inches(5.4))
    tb = s5.shapes.add_textbox(left + Inches(0.18), Inches(1.7), Inches(3.34), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = PRIMARY_BLUE
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(10)

# ==========================================
# SLIDE 6: 深度成因剖析 2 —— 產業面與實體經濟
# ==========================================
s6 = prs.slides.add_slide(blank_layout)
add_header(s6, "深度成因剖析 2【產業面與實體經濟】：薪資所得追價乏力與裝修工料暴增反噬殘值")

ind_cards = [
    ("1. 實質薪資成長停滯，房貸負擔能力抵達天花板",
     "• 缺乏高薪科技產業加持：南區非中科或南科科技產業聚落，居民多為傳統受薪階級、公教人員與傳產從業人員。\n"
     "• 可支配所得受通膨侵蝕：在日常物價、水電與生活成本高漲下，薪資實質增幅追不上房價前波漲幅。\n"
     "• 房貸所得比超過警戒：若以南區老屋均價 25 萬計算，3 房總價近千萬，房貸月負擔已超過家庭月收入 45% 警戒線，買盤完全無力追高。", Inches(0.8)),
    ("2. 缺工缺料引發裝修成本暴漲，直接扣抵老屋交易殘值",
     "• 水電泥作大漲 30%～50%：近年室內裝修因缺工工資倍增與建材原料暴漲，翻修 30 年以上老屋動輒需 150～250 萬現金支出。\n"
     "• 現金自備門檻暴增：買方購置老屋後必須立即投入大筆現金翻修管線與衛浴，這筆開銷無法併入房貸融資。\n"
     "• 房價殘值被反向折抵：買方在談判時必然將 150～200 萬的裝潢翻修預算直接從購屋出價中扣抵，老屋本體價值被嚴重反噬！", Inches(4.8)),
    ("3. 成熟商圈缺乏新增大型重大題材，資本向外分流",
     "• 商圈發展成熟飽和：工學商圈、興大商圈與復興商圈機能雖好，但都市更新困難、腹地有限，缺乏大型旗艦投資帶動。\n"
     "• 重磅產業外溢有限：台中近年重大產業發展集中於中科二期、水湳經貿與烏日高鐵娛樂城，南區缺乏新增大型就業題材。\n"
     "• 資本追逐新興軸線：投資置產資金全面流向具備高科技題材的新興軌道軸線，老牌社區淪為純防守型市場。", Inches(8.8))
]

for title, desc, left in ind_cards:
    add_card(s6, left, Inches(1.5), Inches(3.7), Inches(5.4))
    tb = s6.shapes.add_textbox(left + Inches(0.18), Inches(1.7), Inches(3.34), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(10)

# ==========================================
# SLIDE 7: 深度成因剖析 3 —— 市場面與交易心理
# ==========================================
s7 = prs.slides.add_slide(blank_layout)
add_header(s7, "深度成因剖析 3【市場面與交易心理】：『量先價行』確立、追價破滅與破 2 字頭定錨效應")

mkt_cards = [
    ("1. 『量先價行』鐵律顯現，成交量崩跌 76.4% 逼退賣方",
     "• 流動性冰封：南區老牌大社區 115 年成交量暴跌 76.4%，全年僅剩 174 筆成交。歷史經驗顯示，量縮持續 6 個月以上價格必跌！\n"
     "• 零成交心理恐慌：許多社區連續 3~5 個月零成交，委託期限到期未售，使賣方原本堅持高價的心態開始動搖。\n"
     "• 急售屋主率先破防：部分面臨資金調度、換屋或資產重配之屋主，開始帶頭主動大降價，拉開全面降價序幕。", Inches(0.8)),
    ("2. 買方心理逆轉：從『恐慌搶買』轉為『理性觀望等更跌』",
     "• 追價動能徹底熄火：前兩年『現在不買以後更貴』的狂熱心理徹底破滅，市場預期轉向『晚買享折扣』。\n"
     "• 議價空間大幅拉大：買方看屋態度從過去的加價競奪，轉為開口先砍 15%~25%，堅決不追高。\n"
     "• 斡旋成交門檻提高：買方出價衡量自身財務極限，若賣方讓利幅度不足，買方轉頭即走、絕不戀棧。", Inches(4.8)),
    ("3. 指標社區摜破 2 字頭，形成無可逆轉之比價定錨",
     "• 骨牌式破底成交：工學天下最新殺到 19.6 萬、文心及第跌至 18.9 萬、京城特區更見 18.6 萬低點！\n"
     "• 2 字頭心理關卡失守：原本屋主死守的 23~26 萬防線被市場實登狠狠撕裂，市場定錨價格直接被壓制在 19~21 萬。\n"
     "• 鄰近社區比價殺跌：買方直接拿工學天下與文心及第的 1 字頭成交作為殺價範本，高開價社區完全喪失議價主導權。", Inches(8.8))
]

for title, desc, left in mkt_cards:
    add_card(s7, left, Inches(1.5), Inches(3.7), Inches(5.4))
    tb = s7.shapes.add_textbox(left + Inches(0.18), Inches(1.7), Inches(3.34), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = GOLD
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(10)

# ==========================================
# SLIDE 8: 深度成因剖析 4 —— 金融面與信用管制
# ==========================================
s8 = prs.slides.add_slide(blank_layout)
add_header(s8, "深度成因剖析 4【金融面與信用管制】：央行重拳打擊、銀行滿水位與老屋貸款大限")

fin_cards = [
    ("1. 央行第七波信用管制重拳，多屋與換屋資金鏈徹底斷裂",
     "• 嚴打名下有房族群：名下有房者購置第一戶房貸取消寬限期；自然人第二戶限貸成數砍至 5 成且無寬限期。\n"
     "• 換屋買盤動能被掐死：原本想『先買後賣』的換屋族群因自備款成數暴增 3~4 成而被迫縮手觀望。\n"
     "• 投資置產客全面退場：短期置產槓桿資金被徹底清出場外，市場買盤僅剩預算極為有限的純自住剛需。", Inches(0.8)),
    ("2. 銀行法 72-2 滿水位限制，實質房貸利率站穩 3% 高檔",
     "• 行庫排隊審核成常態：各大公民營銀行不動產放款逼近 30% 法定上限，普遍實施『限貸令』與以價制量策略。\n"
     "• 借款利率顯著攀升：首購實質利率自 1.8% 攀升至 2.5%~2.8%，非首購更直接站穩 3.0%~3.5% 之高檔。\n"
     "• 負擔能力大幅縮水：利息支出增加直接吃掉買方的月還款預算，相當於房屋可負擔總價被直接削去 10%~15%！", Inches(4.8)),
    ("3. 30年/35年老屋金融大關縮限：年限腰斬至 20 年或 15 年",
     "• 屋齡加年限嚴格受限：銀行內部審核遵守『屋齡＋貸款年限 ≤ 50 或 55』之內規，30 年以上老屋無法貸滿 30/40 年。\n"
     "• 月供負擔直接翻倍：貸款 700 萬若從 30 年縮短至 20 年，每月還款自 2.6 萬暴增至 3.7 萬元！\n"
     "• 封殺首購接盤意願：年輕自住客每月現金流無法負擔老屋的短年限高月繳，迫使老陳屋只能等待全現金或高自備款買方。", Inches(8.8))
]

for title, desc, left in fin_cards:
    add_card(s8, left, Inches(1.5), Inches(3.7), Inches(5.4))
    tb = s8.shapes.add_textbox(left + Inches(0.18), Inches(1.7), Inches(3.34), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = RED
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.3
    p_d.space_before = Pt(10)

# ==========================================
# SLIDE 9: 後續三年每年下跌 7.5% 模型建構與測算邏輯
# ==========================================
s9 = prs.slides.add_slide(blank_layout)
add_header(s9, "趨勢預估模型：後續三年（116～118年）每年下修 7.5% 測算邏輯與軟著陸演進")

add_card(s9, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.5))
tb_s9_l = s9.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_s9_l = tb_s9_l.text_frame
tf_s9_l.word_wrap = True

p_ml_h = tf_s9_l.paragraphs[0]
p_ml_h.text = "📐 模型建構機制與測算公式"
p_ml_h.font.name = "微軟正黑體"
p_ml_h.font.size = Pt(15)
p_ml_h.font.bold = True
p_ml_h.font.color.rgb = PRIMARY_BLUE

m_desc = [
    ("預估核心假設", "在央行信用管制不鬆綁、實質利率維持高檔與老屋貸款受限下，南區老牌社區價格呈現每年 -7.5% 的溫和階梯式下行回調。"),
    ("數學演進公式", "第 t+1 年預估單價 P(t+1) = P(t) × (1 - 0.075) = P(t) × 0.925。\n三年累計折現率：(0.925)^3 = 0.7916（三年累計跌幅約 -20.6%～-20.8%）。"),
    ("軟著陸而非崩盤", "年跌 7.5% 兼顧市場交易流動性與自住買盤承接力，避免了恐慌性拋售斷頭，屬於理性的資產去槓桿與均值回歸。"),
    ("長期歷史支撐驗證", "全南區 106～109 年長期中位數支撐帶為 15.2～18.0 萬/坪。三年下修 -20.6% 後，大盤價格（19.5萬）剛好平穩靠攏該支撐帶上緣！")
]

for title, desc in m_desc:
    p_t = tf_s9_l.add_paragraph()
    p_t.text = f"• {title}："
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(11)
    p_t.font.bold = True
    p_t.font.color.rgb = SLATE
    p_t.space_before = Pt(8)
    
    p_d = tf_s9_l.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.25

add_card(s9, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.5))
tb_s9_r = s9.shapes.add_textbox(Inches(7.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_s9_r = tb_s9_r.text_frame
tf_s9_r.word_wrap = True

p_mr_h = tf_s9_r.paragraphs[0]
p_mr_h.text = "📊 大盤後續三年每年下跌 7.5% 數據演進表"
p_mr_h.font.name = "微軟正黑體"
p_mr_h.font.size = Pt(15)
p_mr_h.font.bold = True
p_mr_h.font.color.rgb = RED

steps_data = [
    ("【115年 當前現況】", "中位數：24.6 萬/坪  ｜  年度成交量急縮至 174 件", "大盤衝破 26.9 萬天花板後第一階段回修，買氣急凍觀望。", PRIMARY_BLUE),
    ("【116年 預估 (-7.5%)】", "中位數：22.8 萬/坪  ｜  累積跌幅：-7.5%", "賣方防線開始鬆動，同社區多戶委售削價搶跑，單價下破 23 萬。", SLATE),
    ("【117年 預估 (-7.5%)】", "中位數：21.1 萬/坪  ｜  累積跌幅：-14.4%", "2 字頭保衛戰開打，更多指標建案見 1 字頭成交，買方掌握全面定價權。", GOLD),
    ("【118年 預估 (-7.5%)】", "中位數：19.5 萬/坪  ｜  累積跌幅：-20.6%", "大盤中位數正式失守 2 字頭，回測歷史長期支撐區間（15.2～18.0萬）。", RED)
]

for title, num, note, col in steps_data:
    p_t = tf_s9_r.add_paragraph()
    p_t.text = f"{title} {num}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(11)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    p_t.space_before = Pt(8)
    
    p_n = tf_s9_r.add_paragraph()
    p_n.text = note
    p_n.font.name = "微軟正黑體"
    p_n.font.size = Pt(9.5)
    p_n.font.color.rgb = TEXT_DARK
    p_n.line_spacing = 1.2

# ==========================================
# SLIDE 10: 南區 10 大指標老牌大社區實登走勢與後續三年預估總對照表
# ==========================================
s10 = prs.slides.add_slide(blank_layout)
add_header(s10, "指標對照總表：南區 10 大指標老牌社區實登行情與後續三年 (-7.5%/年) 預估總覽")

table_shape = s10.shapes.add_table(11, 9, Inches(0.8), Inches(1.5), Inches(11.733), Inches(4.6))
tbl = table_shape.table
tbl.columns[0].width = Inches(1.8)
tbl.columns[1].width = Inches(1.1)
tbl.columns[2].width = Inches(1.5)
tbl.columns[3].width = Inches(1.4)
tbl.columns[4].width = Inches(1.2)
tbl.columns[5].width = Inches(1.1)
tbl.columns[6].width = Inches(1.1)
tbl.columns[7].width = Inches(1.1)
tbl.columns[8].width = Inches(1.433)

headers_bench = ["指標社區建案", "總戶數/屋齡", "歷史最高單價", "115最新成交", "115年均價", "116預估(-7.5%)", "117預估(-7.5%)", "118預估(-7.5%)", "三年累積跌幅"]
for col_idx, h_text in enumerate(headers_bench):
    cell = tbl.cell(0, col_idx)
    set_cell_bg(cell, PRIMARY_BLUE)
    p = cell.text_frame.paragraphs[0]
    p.text = h_text
    p.font.name = "微軟正黑體"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

for row_idx, b in enumerate(bench_data, start=1):
    row_vals = [
        b['label'],
        f"{b['households']}戶 / {b['age']}年",
        f"{b['peak_p']:.1f}萬 ({b['peak_date'].split(' ')[0]})",
        f"{b['latest_p']:.1f}萬 ({b['latest_date'].split(' ')[0]})",
        f"{b['avg115']:.1f} 萬/坪",
        f"{b['p116']:.1f} 萬/坪",
        f"{b['p117']:.1f} 萬/坪",
        f"{b['p118']:.1f} 萬/坪",
        f"{b['cum_drop']:.1f}%"
    ]
    for col_idx, text in enumerate(row_vals):
        cell = tbl.cell(row_idx, col_idx)
        set_cell_bg(cell, ROW_ALT if row_idx % 2 == 1 else WHITE)
        p = cell.text_frame.paragraphs[0]
        p.text = text
        p.font.name = "微軟正黑體"
        p.font.size = Pt(9.0)
        p.font.color.rgb = TEXT_DARK
        p.alignment = PP_ALIGN.CENTER
        if col_idx in [0, 3, 7, 8]:
            p.font.bold = True
        if col_idx == 8:
            p.font.color.rgb = RED

card_t_foot = add_card(s10, Inches(0.8), Inches(6.25), Inches(11.733), Inches(0.9), bg_color=RGBColor(254, 242, 242), border_color=RED)
tb_tf = s10.shapes.add_textbox(Inches(1.0), Inches(6.3), Inches(11.333), Inches(0.8))
tf_tf = tb_tf.text_frame
tf_tf.word_wrap = True
p_tf = tf_tf.paragraphs[0]
p_tf.text = "💡 數據洞察：十大指標案最新成交已有工學天下(19.6萬)、文心及第(18.9萬)、京城特區(18.6萬)跌破 2 字頭！按每年 7.5% 溫和下修模型，至 118 年除文心高第與假日半島外，其餘社區均價將全面回落至 15～18 萬歷史支撐區間。"
p_tf.font.name = "微軟正黑體"
p_tf.font.size = Pt(10)
p_tf.font.bold = True
p_tf.font.color.rgb = RED

# ==========================================
# SLIDES 11~20: 10 大指標社區獨立個別走勢圖與深度剖析
# ==========================================
for idx, b in enumerate(bench_data, start=1):
    s_comm = prs.slides.add_slide(blank_layout)
    add_header(s_comm, f"社區行情實證 {idx}【{b['label']}】（{b['households']}戶・屋齡{b['age']}年）實登走勢與後續三年預估")
    
    img_path = os.path.join(trend_dir, b['img'])
    if os.path.exists(img_path):
        s_comm.shapes.add_picture(img_path, Inches(0.8), Inches(1.5), width=Inches(7.2))
    
    add_card(s_comm, Inches(8.2), Inches(1.5), Inches(4.3), Inches(5.5))
    tb_c = s_comm.shapes.add_textbox(Inches(8.4), Inches(1.65), Inches(3.9), Inches(5.2))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    
    p_ch = tf_c.paragraphs[0]
    p_ch.text = f"🏢 {b['name']} 實登關鍵指標"
    p_ch.font.name = "微軟正黑體"
    p_ch.font.size = Pt(14)
    p_ch.font.bold = True
    p_ch.font.color.rgb = PRIMARY_BLUE
    
    comm_points = [
        ("歷史最高成交紀錄", f"{b['peak_p']:.1f} 萬/坪（{b['peak_date']}）", RED),
        ("115 年最新實登成交", f"{b['latest_p']:.1f} 萬/坪（{b['latest_date']}）", GOLD),
        ("純住宅年度均價走勢", f"113年 {b['avg113']:.1f}萬 -> 115年 {b['avg115']:.1f}萬 ({b['drop_pct']:.1f}%)", PRIMARY_BLUE),
        ("後續三年年跌 -7.5% 預估", f"116年 {b['p116']:.1f}萬 (-7.5%)\n117年 {b['p117']:.1f}萬 (-7.5%)\n118年 {b['p118']:.1f}萬 (三年累計 {b['cum_drop']:.1f}%)", RED),
        ("長期合理價值支撐帶", f"{b['supp']}", GREEN),
        ("走勢深度分析與成因", f"{b['desc']}", TEXT_DARK)
    ]
    
    for title, val, col in comm_points:
        p_t = tf_c.add_paragraph()
        p_t.text = f"• {title}："
        p_t.font.name = "微軟正黑體"
        p_t.font.size = Pt(10.5)
        p_t.font.bold = True
        p_t.font.color.rgb = col
        p_t.space_before = Pt(5)
        
        p_v = tf_c.add_paragraph()
        p_v.text = val
        p_v.font.name = "微軟正黑體"
        p_v.font.size = Pt(9.5)
        p_v.font.color.rgb = TEXT_DARK
        p_v.line_spacing = 1.2

# ==========================================
# SLIDE 21: 總結建言與買賣雙方決策指引
# ==========================================
s21 = prs.slides.add_slide(blank_layout)
add_header(s21, "戰略決策指引：老牌社區下行週期確立，買賣雙方的理性因應對策")

add_card(s21, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.5), border_color=RED)
tb_s21_l = s21.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_s21_l = tb_s21_l.text_frame
tf_s21_l.word_wrap = True

p_sl_h = tf_s21_l.paragraphs[0]
p_sl_h.text = "🏷️ 給賣方（屋主）的實戰建言：及時讓利、現金落袋"
p_sl_h.font.name = "微軟正黑體"
p_sl_h.font.size = Pt(15)
p_sl_h.font.bold = True
p_sl_h.font.color.rgb = RED

seller_points = [
    ("不要成為替鄰居免費抬轎的犧牲品", "同社區在售戶數普遍在 4～10 戶以上，市場只有極少數中年自住買方願意接手。誰先願意適度讓利，誰就能先把數百萬現金安全抱回家！"),
    ("避免每年 7.5% 跌幅侵蝕多年獲利", "按模型測算，越晚賣價格只會越低，若再拖兩年，以 40 坪物件計算相當於房屋總值憑空蒸發 150～200 萬元以上！"),
    ("屋齡跨過 35 年將徹底喪失金融流動性", "屋齡一年比一年老，銀行 30/40 年貸款大門徹底關閉，未來接盤者貸款受限，流動性枯竭將導致折價幅度更為劇烈。"),
    ("把握目前高檔反轉初期的高利潤窗口", "若多年前以 13～17 萬買入，目前即便在 20～22 萬成交依然大賺數百萬，此時落袋為安才是最大贏家。")
]

for title, desc in seller_points:
    p_t = tf_s21_l.add_paragraph()
    p_t.text = f"• {title}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(11)
    p_t.font.bold = True
    p_t.font.color.rgb = RED
    p_t.space_before = Pt(8)
    
    p_d = tf_s21_l.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.25

add_card(s21, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.5), border_color=GREEN)
tb_s21_r = s21.shapes.add_textbox(Inches(7.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_s21_r = tb_s21_r.text_frame
tf_s21_r.word_wrap = True

p_sr_h = tf_s21_r.paragraphs[0]
p_sr_h.text = "🎯 給買方（買家）的出價指引：理性定錨、嚴守防線"
p_sr_h.font.name = "微軟正黑體"
p_sr_h.font.size = Pt(15)
p_sr_h.font.bold = True
p_sr_h.font.color.rgb = GREEN

buyer_points = [
    ("堅決以『實登最低價再低兩成』起手出價", "實登行情已全線見頂反轉，出價應直接參照最新成交低點打 8 折（約 18～19 萬/坪），絕不替屋主過去兩年的狂熱買單。"),
    ("將 150～200 萬裝修翻新費用預先扣除", "30 年以上老屋水電管線與衛浴必須重新翻新，這筆不可融資的現金支出必須直接在出價總價中全額折減。"),
    ("審慎檢視銀行貸款年限與自備款承受力", "確認銀行是否能核貸 7 成以上以及年限是否受限（20年 vs 30年），斡旋單務必加註『房貸未滿七成無條件解約退斡』保護條款。"),
    ("善用跨社區與跨生活圈競品替代談判", "同時評估南屯（如大墩人文）或新重劃區新小宅，以隨時可轉向的客觀心態談判，牢牢掌握議價主導權！")
]

for title, desc in buyer_points:
    p_t = tf_s21_r.add_paragraph()
    p_t.text = f"• {title}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(11)
    p_t.font.bold = True
    p_t.font.color.rgb = GREEN
    p_t.space_before = Pt(8)
    
    p_d = tf_s21_r.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.25

out_ppt_name = "台中市南區老牌大型社區實價登錄趨勢與後續三年下跌預估分析.pptx"
out_paths = [
    os.path.join(trend_dir, out_ppt_name),
    os.path.join(base_dir, out_ppt_name),
    os.path.join(base_dir, "長億南園8樓", out_ppt_name)
]

for p in out_paths:
    os.makedirs(os.path.dirname(p), exist_ok=True)
    prs.save(p)
    print(f"✅ PowerPoint 簡報已成功生成 ({len(prs.slides)} 頁): {p}")

print("\n🎉 全部 PowerPoint 簡報生成完成！")

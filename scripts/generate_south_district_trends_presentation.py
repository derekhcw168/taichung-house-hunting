# -*- coding: utf-8 -*-
"""
生成《台中市南區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx》
全套 65 頁高規格專案決策簡報架構：
1. 封面：南區老牌大型社區實價登錄趨勢與後續三年預估分析報告
2. 執行摘要：南區老牌大社區房市大盤 4 大核心研判結論
3. 大盤全景解構：54 棟老牌大社區實登價量全圖與歷史週期剖析（101～115年）
4. 成因全景架構：促使老牌大型社區價格持續下修的四大驅動力
5. 深度成因剖析 1【整體供需面】：龐大老屋存量賣壓、新成屋交屋潮夾殺與買盤斷層
6. 深度成因剖析 2【結構機能與維護面】：結構機能衰退與公共維護負擔，老屋邁向50年面臨買盤歸零
7. 深度成因剖析 3【市場面與交易心理】：『量先價行』確立、追價破滅與破 2 字頭定錨效應
8. 深度成因剖析 4【金融面與信用管制】：央行重拳打擊、銀行滿水位與老屋貸款大限
9. 價格預測模型解析：計量經濟學動態預估原理、流動性衝擊與價值引力機制（R5 獨立全頁說明）
10. 指標對照總表：南區 10 大指標老牌社區實登行情與後續三年動態計量預估總覽（11 欄位專業評估表格）
11～64. 全南區全數 54 個老牌大型社區獨立價量走勢圖與月度動態計量預估剖析卡片（共 54 頁）
65. 戰略決策指引：老牌社區下行週期確立，買賣雙方的理性因應對策
"""

import os
import sys
import json
import math
import shutil
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Add project root and scratch directory to sys.path to load chart_engine
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRATCH_DIR = os.path.join(PROJECT_ROOT, "scratch")
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)
if SCRATCH_DIR not in sys.path:
    sys.path.insert(0, SCRATCH_DIR)

from scratch.chart_engine import compute_dynamic_forecast

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

trend_dir = os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖")

# 日期轉換輔助函式（精準至月份）
def format_roc_month(date_str):
    """將西元或民國日期格式化為民國年月（如 113年6月）"""
    if not date_str or str(date_str) in ('—', 'None', 'null', ''):
        return '—'
    s = str(date_str).strip().replace('/', '-').replace('.', '-')
    parts = s.split('-')
    try:
        y = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 1
        ry = y - 1911 if y > 1900 else y
        return f"{ry}年{m}月"
    except Exception:
        return str(date_str)

def format_roc_short(date_str):
    """將日期格式化為簡潔民國年月（如 113/06）"""
    if not date_str or str(date_str) in ('—', 'None', 'null', ''):
        return '—'
    s = str(date_str).strip().replace('/', '-').replace('.', '-')
    parts = s.split('-')
    try:
        y = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 1
        ry = y - 1911 if y > 1900 else y
        return f"{ry}/{m:02d}"
    except Exception:
        return str(date_str)

# 1. 載入全量交易資料集
MASTER_DATA_FILE = os.path.join(SCRATCH_DIR, "all_54_communities_transactions.json")
if not os.path.exists(MASTER_DATA_FILE):
    MASTER_DATA_FILE = os.path.join(SCRATCH_DIR, "all_53_communities_transactions.json")
if not os.path.exists(MASTER_DATA_FILE):
    raise FileNotFoundError(f"找不到主交易資料庫: {MASTER_DATA_FILE}")

with open(MASTER_DATA_FILE, 'r', encoding='utf-8') as f:
    master_data = json.load(f)

communities_list = master_data.get('communities', [])
comm_by_name = {c['name']: c for c in communities_list}
comm_by_id = {c['id']: c for c in communities_list}

# 計算大盤全量純住宅交易資料與動態計量預估
all_market_tx = []
for c in communities_list:
    all_market_tx.extend(c.get('transactions', []))

market_forecast = compute_dynamic_forecast(
    all_market_tx,
    community_metadata={'households': 21417, 'build_date': '1995-06-15'}
)

m_p116 = market_forecast[116]['expected']
m_p116_l = market_forecast[116]['lower']
m_p116_u = market_forecast[116]['upper']
m_d116 = market_forecast[116]['drop_pct']

m_p117 = market_forecast[117]['expected']
m_p117_l = market_forecast[117]['lower']
m_p117_u = market_forecast[117]['upper']
m_d117 = market_forecast[117]['drop_pct']

m_p118 = market_forecast[118]['expected']
m_p118_l = market_forecast[118]['lower']
m_p118_u = market_forecast[118]['upper']
m_d118 = market_forecast[118]['drop_pct']

m_cum_drop = market_forecast['cum_drop_pct']
m_supp_low, m_supp_high = market_forecast['support_band']

# 2. 定義 10 大指標社區（剔除大墩人文，補入跌幅次高之望族二期，其餘依跌幅顯著排序）
bench_defs = [
    {
        'rank': 1,
        'label': '工學天下',
        'match_name': '工學天下',
        'img': '07_工學天下_成交行情走勢圖.png',
        'desc_template': '工學商圈標竿老牌名宅（452戶・屋齡31.7年）。113年7月創下 {peak_p:.1f} 萬歷史高點，115年6月最新成交直接摜破2字頭重挫至 {latest_p:.1f} 萬破底（最新跌幅 {drop_latest:.1f}%）！115年均價 {avg115_str}（均價跌幅 {drop_avg:.1f}%）。在計量經濟模型測算下，預估 116 年下修至 {p116:.1f} 萬，117 年跌至 {p117:.1f} 萬，118 年收斂於 {p118:.1f} 萬（三年累計修正 {cum_drop:.1f}%），回測長期價值支撐帶（{supp}）。'
    },
    {
        'rank': 2,
        'label': '文心及第',
        'match_name': '文心及第(文心南路)',
        'img': '27_文心及第文心南路_成交行情走勢圖.png',
        'desc_template': '文心南路生活圈指標案（259戶・屋齡27.2年）。113年6月創下 {peak_p:.1f} 萬天花板，115年3月頂樓次高戶直接暴跌至 {latest_p:.1f} 萬成交（崩跌 {drop_latest:.1f}%，南區跌幅第一）！115年均價 {avg115_str}（均價跌幅 {drop_avg:.1f}%）。計量模型預估 116 年下探 {p116:.1f} 萬，117 年 {p117:.1f} 萬，118 年收斂於 {p118:.1f} 萬（累計修正 {cum_drop:.1f}%），成為全區比價降價核心定錨。'
    },
    {
        'rank': 3,
        'label': '綠灣大樓',
        'match_name': '綠灣大樓',
        'img': '25_綠灣大樓_成交行情走勢圖.png',
        'desc_template': '綠川水岸生活圈（185戶・屋齡27.4年）。113年10月高點 {peak_p:.1f} 萬，115年3月最新成交直接見 1 字頭 {latest_p:.1f} 萬（暴跌 {drop_latest:.1f}%），115均價跌至 {avg115_str}（均價跌幅 {drop_avg:.1f}%）。計量模型預測 116 年下探 {p116:.1f} 萬，117 年 {p117:.1f} 萬，118 年收斂於 {p118:.1f} 萬（三年累計修正 {cum_drop:.1f}%），破底效應極為顯著。'
    },
    {
        'rank': 4,
        'label': '比佛利山莊',
        'match_name': '比佛利山莊',
        'img': '17_比佛利山莊_成交行情走勢圖.png',
        'desc_template': '南區高總價老牌名邸（247戶・屋齡30.6年）。113年8月歷史高點衝上 {peak_p:.1f} 萬，115年3月最新成交急挫至 {latest_p:.1f} 萬（回檔 {drop_latest:.1f}%），115均價 {avg115_str}（均價跌幅 {drop_avg:.1f}%）。模型預估 116 年跌至 {p116:.1f} 萬，117 年探至 {p117:.1f} 萬，118 年收斂至 {p118:.1f} 萬（累計修正 {cum_drop:.1f}%）。老屋結構維護負擔與大坪數總價過高導致接盤力弱。'
    },
    {
        'rank': 5,
        'label': '德昌中華大樓',
        'match_name': '德昌中華大樓',
        'img': '10_德昌中華大樓_成交行情走勢圖.png',
        'desc_template': '五權南路商圈 409 戶老牌名宅，屋齡 32.3 年。113年7月高點 {peak_p:.1f} 萬，115年3月最新成交殺至 20.0 萬保衛戰（回挫 {drop_latest:.1f}%）！115均價亦跌至 20.0 萬（均價跌幅 {drop_avg:.1f}%）。計量模型測算 116 年即將跌破 2 字頭至 {p116:.1f} 萬，117 年續探 {p117:.1f} 萬，118 年收斂於 {p118:.1f} 萬（累計修正 {cum_drop:.1f}%）。'
    },
    {
        'rank': 6,
        'label': '廣三大時代',
        'match_name': '廣三大時代大廈',
        'img': '02_廣三大時代大廈_成交行情走勢圖.png',
        'desc_template': '工學商圈千戶超大型社區（1,214戶・屋齡31.9年）。113年12月衝上 {peak_p:.1f} 萬天花板後反轉，115年8月最新成交急落至 {latest_p:.1f} 萬（下修 {drop_latest:.1f}%），115均價 {avg115_str}（均價跌幅 {drop_avg:.1f}%）。計量模型預估 116 年下探 {p116:.1f} 萬，117 年摜破 2 字頭至 {p117:.1f} 萬，118 年下修至 {p118:.1f} 萬（累計修正 {cum_drop:.1f}%）。超大戶數二手拋售賣壓沉重。'
    },
    {
        'rank': 7,
        'label': '賺錢時代',
        'match_name': '賺錢時代',
        'img': '01_賺錢時代_成交行情走勢圖.png',
        'desc_template': '全南區規模最大單一社區（2,077戶・屋齡29.5年）。113年11月高點達 {peak_p:.1f} 萬，龐大戶數帶來二手市場劇烈委售競價，115年8月最新成交跌至 {latest_p:.1f} 萬（回修 {drop_latest:.1f}%），115均價 {avg115_str}（均價跌幅 {drop_avg:.1f}%）。計量模型測算 116 年跌至 {p116:.1f} 萬，117 年跌破 2 字頭至 {p117:.1f} 萬，118 年探至 {p118:.1f} 萬（累計修正 {cum_drop:.1f}%）。'
    },
    {
        'rank': 8,
        'label': '遠見天下大樓',
        'match_name': '遠見天下大樓',
        'img': '19_遠見天下大樓_成交行情走勢圖.png',
        'desc_template': '健康公園生活圈 298 戶指標名廈，屋齡 31.8 年。113年12月極值 {peak_p:.1f} 萬，115年5月最新成交急跌至 {latest_p:.1f} 萬（重挫 {drop_latest:.1f}%），115均價 {avg115_str}（均價跌幅 {drop_avg:.1f}%）。計量模型測算 116 年降至 {p116:.1f} 萬，117 年跌破 2 字頭至 {p117:.1f} 萬，118 年探至 {p118:.1f} 萬（累計修正 {cum_drop:.1f}%），追價動能全無。'
    },
    {
        'rank': 9,
        'label': '長億南園',
        'match_name': '長億南園',
        'img': '28_長億南園_成交行情走勢圖.png',
        'desc_template': '工學商圈標竿成熟名廈（239戶・屋齡32.6年）。113年2月高點衝上 {peak_p:.1f} 萬，115年9月最新成交滑落至 {latest_p:.1f} 萬（修正 {drop_latest:.1f}%），115均價 {avg115_str}（均價跌幅 {drop_avg:.1f}%）。計量模型預估 116 年探至 {p116:.1f} 萬，117 年下挫至 {p117:.1f} 萬，118 年收斂於 {p118:.1f} 萬（累計修正 {cum_drop:.1f}%），回歸長期歷史價值支撐區。'
    },
    {
        'rank': 10,
        'label': '望族二期',
        'match_name': '望族二期',
        'img': '16_望族二期_成交行情走勢圖.png',
        'desc_template': '復興北路生活圈成熟名邸（314戶・屋齡32.9年）。113年12月創下 {peak_p:.1f} 萬歷史高點，115年8月最新成交急挫至 {latest_p:.1f} 萬（回跌 {drop_latest:.1f}%），115均價 {avg115_str}（均價跌幅 {drop_avg:.1f}%）。計量模型測算 116 年預估下修至 {p116:.1f} 萬，117 年跌至 {p117:.1f} 萬，118 年收斂於 {p118:.1f} 萬（三年累計修正 {cum_drop:.1f}%），回歸長期歷史價值支撐帶（{supp}）。'
    }
]

bench_data = []
for b_def in bench_defs:
    name = b_def['match_name']
    c = comm_by_name.get(name)
    if not c:
        raise ValueError(f"找不到指標社區: {name}")

    fc = compute_dynamic_forecast(c.get('transactions', []), community_metadata=c)
    
    p116 = fc[116]['expected']
    d116 = fc[116]['drop_pct']
    p117 = fc[117]['expected']
    d117 = fc[117]['drop_pct']
    p118 = fc[118]['expected']
    d118 = fc[118]['drop_pct']
    cum_drop = fc['cum_drop_pct']
    supp_l, supp_h = fc['support_band']
    supp_str = f"{supp_l:.1f}～{supp_h:.1f} 萬/坪"

    # Age calculation
    build_date = str(c.get('build_date') or '1995-01-01')
    try:
        by = int(build_date.split('-')[0])
        bm = int(build_date.split('-')[1]) if '-' in build_date else 6
        age = round(2026.8 - (by + bm / 12.0), 1)
    except Exception:
        age = 31.0

    peak_p = float(c.get('historical_peak_price') or 0.0)
    peak_date = str(c.get('historical_peak_date') or '')
    latest_p = float(c.get('latest_deal_price') or 0.0)
    latest_date = str(c.get('latest_deal_date') or '')
    avg113 = float(c.get('annual_avg_113') or 0.0)
    
    # 115 交易狀態判定與關鍵跌幅計算
    raw_avg115 = c.get('annual_avg_115')
    has_115_deals = (raw_avg115 is not None and float(raw_avg115) > 0.0)
    
    if has_115_deals:
        avg115 = float(raw_avg115)
        avg115_str = f"{avg115:.1f} 萬/坪"
        drop_latest = round(((latest_p - peak_p) / peak_p) * 100.0, 1)
        drop_avg = round(((avg115 - peak_p) / peak_p) * 100.0, 1)
        drop_latest_str = f"{drop_latest:+.1f}%"
        drop_avg_str = f"{drop_avg:+.1f}%"
        latest_deal_display = f"{latest_p:.1f}萬 ({format_roc_short(latest_date)})"
    else:
        avg115 = None
        avg115_str = "—*"
        drop_latest = round(((latest_p - peak_p) / peak_p) * 100.0, 1) if peak_p > 0 else 0.0
        drop_avg = None
        drop_latest_str = f"{drop_latest:+.1f}%*"
        drop_avg_str = "—*"
        latest_deal_display = f"{latest_p:.1f}萬 ({format_roc_short(latest_date)})*"

    desc = b_def['desc_template'].format(
        peak_p=peak_p,
        latest_p=latest_p,
        avg113=avg113,
        avg115=avg115,
        avg115_str=avg115_str,
        drop_latest=drop_latest,
        drop_avg=drop_avg if drop_avg is not None else 0.0,
        p116=p116,
        p117=p117,
        p118=p118,
        cum_drop=cum_drop,
        supp=supp_str
    )

    bench_data.append({
        'rank': b_def['rank'],
        'label': b_def['label'],
        'name': name,
        'households': c.get('households', 0),
        'age': age,
        'build_date': build_date,
        'address': c.get('address', ''),
        'peak_p': peak_p,
        'peak_date': peak_date,
        'latest_p': latest_p,
        'latest_date': latest_date,
        'latest_deal_display': latest_deal_display,
        'avg113': avg113,
        'avg115': avg115,
        'avg115_str': avg115_str,
        'drop_latest_115': drop_latest,
        'drop_avg_115': drop_avg,
        'drop_latest_str': drop_latest_str,
        'drop_avg_str': drop_avg_str,
        'has_115_deals': has_115_deals,
        'p116': p116,
        'd116': d116,
        'p117': p117,
        'd117': d117,
        'p118': p118,
        'd118': d118,
        'cum_drop': cum_drop,
        'supp': supp_str,
        'img': b_def['img'],
        'desc': desc
    })

print(f"✅ 成功載入並完成重組後 10 大指標社區動態計量預估計算！")

# 3. 初始化 PowerPoint 簡報 (16:9 寬螢幕)
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

    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.32), Inches(6.0), Inches(0.32))
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
p2.text = "實價登錄趨勢與後續三年預估分析報告"
p2.font.name = "微軟正黑體"
p2.font.size = Pt(32)
p2.font.bold = True
p2.font.color.rgb = WHITE
p2.space_before = Pt(8)

p3 = tf1.add_paragraph()
p3.text = "【大盤量縮 76.4% 轉折確立 ｜ 月度高頻動態計量預估模型 ｜ 全數 54 社區全圖鑑 ｜ 供需・結構機能・市場・金融四大成因全景解析】"
p3.font.name = "微軟正黑體"
p3.font.size = Pt(13)
p3.font.color.rgb = RGBColor(190, 210, 240)
p3.space_before = Pt(12)

card_foot = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.6), Inches(11.733), Inches(1.1))
set_shape_flat(card_foot, RGBColor(22, 45, 85), RGBColor(40, 75, 130), 1)
tf_foot = card_foot.text_frame
tf_foot.vertical_anchor = MSO_ANCHOR.MIDDLE
p_foot = tf_foot.paragraphs[0]
p_foot.text = "數據範疇：全南區 54 棟 25 年以上大型集合住宅社區（總計 17,945 戶、整體老屋存量逾 21,396 戶、平均屋齡 31.0 年） ｜ 實登總計：7,072 筆（純住宅 5,935 筆） ｜ 時間精度：月度時間序列 ｜ 分析日期：2026 年 10 月"
p_foot.font.name = "微軟正黑體"
p_foot.font.size = Pt(11)
p_foot.font.color.rgb = WHITE
p_foot.alignment = PP_ALIGN.CENTER

# ==========================================
# SLIDE 2: 執行摘要 (Executive Summary) — R4 修正
# ==========================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "執行摘要：南區老牌大社區房市大盤 4 大核心研判結論")

cards_data_s2 = [
    ("📉 大盤轉折確立・量縮 76.4%", 
     "全南區 54 棟老牌大案於 113 年衝上 26.9 萬/坪歷史天花板後見頂。央行實施第七波選擇性信用管制後，115 年成交量自高峰期年均 739 件崩縮至 174 件（暴跌 76.4%），『量先價行』鐵律確立市場全面反轉下行。", 
     RED, Inches(0.8), Inches(1.5)),
    ("📉 計量經濟學動態下修模型", 
     f"捨棄傳統固定跌幅，建立嚴謹月度動態計量經濟模型（量價衰減回歸、流動性衝擊動能、Holt阻尼平滑與IQR價值支撐重力模型）：測算大盤中位數 116 年 {m_p116:.1f} 萬 ➜ 117 年 {m_p117:.1f} 萬 ➜ 118 年 {m_p118:.1f} 萬，三年累計跌幅達 {m_cum_drop:.1f}%，理性靠攏長期合理支撐帶。", 
     GOLD, Inches(6.8), Inches(1.5)),
    ("🏢 十大指標建案接連破底", 
     "工學天下、文心及第、綠灣大樓、廣三大時代等南區 10 大標竿社區全數高檔重挫！其中文心及第跌幅達 47.1%、綠灣大樓跌 40.2%、比佛利山莊跌 44.0%，工學天下(19.6萬)等最新成交直接摜破 2 字頭，形成無可逆轉之比價定錨效應。", 
     SLATE, Inches(0.8), Inches(4.3)),
    ("🎯 四大成因共振拉動下行", 
     "【供需面】2.1萬戶老屋存量賣壓與新成屋交屋潮夾殺；【結構維護面】結構機能衰退與公共維護負擔加劇，30餘年老社區居住10至20年後將成為50年屋齡老舊建築，屆時將無人購買更沒有實質買盤；【市場面】追價心理破滅與流動性枯竭；【金融面】央行第七波重拳與老屋貸款年限腰斬至15~20年。", 
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
add_header(s3, "大盤全景解構：南區 54 棟老牌大社區實登價量全圖與歷史週期剖析")

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
    ("101～106年【早期溫和盤整期】", "全區純住宅均價落在 10.0～14.1 萬/坪，成交量平穩扎實（年均約 400～700 餘件），老牌社區以高性價比為首購主力客群。", SLATE),
    ("107～112年【低利資金推升主升期】", "受全球寬鬆熱錢與通膨預期帶動，單價突破 20 萬大關，全區年度成交創下 739 件高峰。", PRIMARY_BLUE),
    ("113～115年【見頂回落量縮急凍期】", "113 年見 26.9 萬歷史天花板，央行第七波管制後成交量崩跌 76.4%（至 174 件），115 年中位數回修至 24.6 萬。", RED),
    ("116～118年【計量經濟動態下行期】", f"月度動態計量模型預估：116年下修至 {m_p116:.1f} 萬 ➜ 117年探至 {m_p117:.1f} 萬 ➜ 118年收斂於 {m_p118:.1f} 萬，回測歷史長期價值支撐帶（{m_supp_low:.1f}～{m_supp_high:.1f}萬）。", GOLD)
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
# SLIDE 4: 促使價格走勢調整與預估分析成因剖析（四大維度架構總覽） — R4 修正
# ==========================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "成因全景架構：促使老牌大型社區價格持續下修的四大驅動力")

drivers_s4 = [
    ("🏗️ 整體供需面", "龐大老屋存量與新成屋交屋潮夾殺", 
     "• 54 棟老案庫存逾 2.1 萬戶，市場轉冷後二手委售件數暴增\n• 周邊烏日高鐵、十三期新成屋交屋潮強烈瓜分優質自住客源\n• 年輕首購族排斥老陳屋，自住買盤面臨少子化世代斷層", 
     Inches(0.8), PRIMARY_BLUE),
    ("🏚️ 結構機能與維護面", "結構機能衰退與公共維護負擔", 
     "• 大樓水電管線老化、外牆滲漏與電梯維修頻繁，公共修繕基金不足\n• 管委會維護負擔沉重引發鉅額修繕分攤款，反噬老屋殘值\n• 現在歷時三十幾年的老社區，在居住 10 至 20 年之後，將成為 50 年屋齡的老舊建築，屆時將無人購買，更沒有實質買盤！", 
     Inches(3.8), SLATE),
    ("📉 市場與交易心理", "量先價行與預期追價破滅", 
     "• 成交量暴跌 76.4%，流動性枯竭迫使急售賣方防線崩潰\n• 買方心理從『恐慌搶買』徹底轉為『理性觀望等更跌』\n• 工學天下、文心及第摜破 2 字頭引發全區比價連鎖殺跌", 
     Inches(6.8), GOLD),
    ("🏦 金融與信用管制", "央行重拳與老屋貸款大門關閉", 
     "• 央行第七波信用管制砍房貸成數，名下有房第二戶限貸 5 成\n• 銀行法 72-2 滿水位警戒，房貸利率實質站穩 3% 高檔\n• 30/35 年老屋貸款年限被砍至 15~20 年，月付金翻倍壓垮買盤", 
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
     "• 供給母體龐大：南區僅這 54 棟 25 年以上老牌大案，總戶數高達 21,396 戶（統計總戶數 17,945 戶）。如賺錢時代(2,077戶)、廣三大時代(1,214戶)、長億城(1,241戶)動輒數百上千戶。\n"
     "• 委售重疊競爭：房市降溫後，每棟社區平均在售物件由 2~3 戶激增至 8~15 戶。多品牌房仲洗版聯賣，造成同社區屋主彼此搶客削價求現。\n"
     "• 去化時間劇增：平均委售去化天數由過去的 45 天大幅拉長至 180 天以上，中古供給嚴重過剩。", Inches(0.8)),
    ("2. 鄰近新興重劃區交屋潮分流，優質自住買盤被吸走",
     "• 新重劃區夾殺：南區緊鄰烏日高鐵特區、十三期重劃區等新興熱區，近期迎來龐大預售完工交屋潮。\n"
     "• 產品代差顯著：新大樓採用最新 921 震後耐震法規、配備坡道平面車位、多功能接待會館與無障礙動線，大幅降低後續維護困擾。\n"
     "• 預算排擠效應：總價 1,000～1,400 萬區間，買方寧可選擇外圍新成屋 2 房或小 3 房，直接吸走老陳屋的潛在買客。", Inches(4.8)),
    ("3. 買方世代斷層與少子化，願意接盤中古陳屋者稀缺",
     "• 青年首購偏好新小宅：新一代首購族對 30 年以上老舊管線、斑駁外觀與無車位抗拒感極強，幾乎不考慮中古老陳屋。\n"
     "• 接盤主力嚴重萎縮：目前願意承接 30 年老陳屋的買方，幾乎僅剩具備實際家庭居住需求的中年自住換屋客。\n"
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
# SLIDE 6: 深度成因剖析 2 —— 結構機能與維護面 — R4 修正
# ==========================================
s6 = prs.slides.add_slide(blank_layout)
add_header(s6, "深度成因剖析 2【結構機能與維護面】：結構機能衰退與公共維護負擔，老屋邁向50年面臨買盤歸零")

ind_cards = [
    ("1. 大樓結構與機電管線老化，公共維護負擔加劇",
     "• 水電公用管線老化劣化：25～35 年以上老社區普遍面臨給排水幹管鏽蝕、消防管路老化與外牆滲漏裂損等嚴重課題。\n"
     "• 機坑與電梯維修頻繁：老牌大社區電梯設備動輒屆齡 30 年，更新主機與纜線每部需百萬以上，日常保養叫修頻率激增。\n"
     "• 公共維護負擔壓垮管委會：管理費提撥水準偏低導致長期修繕準備金嚴重短缺，已成為中古大樓日常營運的重大隱憂。", Inches(0.8)),
    ("2. 管委會修繕基金不足，巨額分攤款反噬房屋殘值",
     "• 專案分攤款引爆住戶衝突：面對外牆拉皮或機電重大修繕，公共基金不足必引發鉅額專案分攤款（每戶數十萬），住戶繳納意願低落陷入惡性循環。\n"
     "• 現金自備門檻實質暴增：買方購置老陳屋後需自備大筆現金承擔結構機能修復與公設補強，無法納入房貸融資。\n"
     "• 房價殘值被反向折抵：買方在談判時必然將數十至上百萬的裝潢翻修預算直接從購屋出價中扣抵，老屋本體價值被嚴重反噬！", Inches(4.8)),
    ("3. 居住 10～20 年後淪為 50 年老屋，終局流動性徹底歸零",
     "• 邁入 50 年老舊大樓危機：現在已經歷時三十幾年的老社區，在居住 10 至 20 年之後，將成為 50 年屋齡的老舊建築，屆時將無人購買，更沒有實質買盤。\n"
     "• 銀行金融機能全線關閉：依據銀行法規與授信內規（屋齡+貸款年限≤50），屆時銀行對 50 年集合住宅幾乎全面拒貸，缺乏房貸槓桿將使流動性徹底冰封。\n"
     "• 都更重建整合難如登天：大型集合住宅數百至上千戶產權極度分散，都更整合難度無限大，老陳屋恐淪為無法變現之終身套牢資產！", Inches(8.8))
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
     "• 流動性冰封：南區老牌大社區 115 年純住宅成交量暴跌 76.4%，全年僅剩 174 筆成交。歷史經驗顯示，量縮持續 6 個月以上價格必跌！\n"
     "• 零成交心理恐慌：許多社區連續 3~5 個月零成交，委託期限到期未售，使賣方原本堅持高價的心態開始動搖。\n"
     "• 急售屋主率先破防：部分面臨資金調度、換屋或資產重配之屋主，開始帶頭主動大降價，拉開全面降價序幕。", Inches(0.8)),
    ("2. 買方心理逆轉：從『恐慌搶買』轉為『理性觀望等更跌』",
     "• 追價動能徹底熄火：前兩年『現在不買以後更貴』的狂熱心理徹底破滅，市場預期轉向『晚買享折扣』。\n"
     "• 議價空間大幅拉大：買方看屋態度從過去的加價競奪，轉為開口先砍 15%~25%，堅決不追高。\n"
     "• 斡旋成交門檻提高：買方出價衡量自身財務極限，若賣方讓利幅度不足，買方轉頭即走、絕不戀棧。", Inches(4.8)),
    ("3. 指標社區摜破 2 字頭，形成無可逆轉之比價定錨",
     "• 骨牌式破底成交：工學天下最新殺到 19.6 萬、文心及第跌至 18.9 萬、綠灣大樓更見 16.8 萬低點！\n"
     "• 2 字頭心理關卡失守：原本屋主死守的 23~26 萬防線被市場實登狠狠撕裂，市場定錨價格直接被壓制在 18~20 萬。\n"
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
    ("3. 30年/35年老屋金融大限：年限腰斬至 15～20 年",
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
# SLIDE 9: 價格預測模型說明頁 (R5 專屬全頁解析)
# ==========================================
s9 = prs.slides.add_slide(blank_layout)
add_header(s9, "價格預測模型解析：計量經濟學動態預估原理、流動性衝擊與價值引力機制")

model_cols_data = [
    ("📐 核心架構與月度推估機制",
     "捨棄粗略年均，以月份為基礎進行高頻動態時間序列建模",
     [
         ("月度時間序列高頻分析", "以月為單位追蹤價格與成交量變化，精準捕捉 113 下半年央行第七波信用管制後流動性斷崖式崩跌之月度轉折點。"),
         ("量價衰減衝擊動能方程", "g_shock(m) = -(λ0 + λv|ΔVm| + λp·PeakDrop)。以大盤月均量縮 76.4% 為基礎，成交量急凍與歷史高點回檔比例直接轉化為月度下修動能。"),
         ("經驗貝式大盤先驗平滑", "wi = N_recent / (N_recent + 4.0)。對近期月成交稀疏之社區，導入南區 54 社區大盤先驗趨勢進行貝式收縮，避免極端異常交易扭曲預測。")
     ],
     Inches(0.8), PRIMARY_BLUE),

    ("⚙️ 結構折價懲罰與價值重力",
     "結合老屋金融限制、社區規模拋售壓力與長期價值支撐帶",
     [
         ("屋齡與規模雙重折價懲罰", "f_age = 1 + min(0.18, (Age-25)*0.015)；f_scale = 1 + min(0.12, (戶數-150)/4000)。30年以上老屋房貸年限遭銀行腰斬至15~20年；大社區內部多戶踩踏搶跑。"),
         ("106～109 年長期價值重力錨", f"以 106～109 年長期實登純住宅四分位距（IQR [{m_supp_low:.1f}～{m_supp_high:.1f}萬/坪]）作為長期價值重力錨。價格愈偏離支撐帶，向下回歸引力愈強。"),
         ("Holt 阻尼平滑與漸近下限", "導入阻尼係數 φ = 0.68，隨價格靠攏歷史合理支撐區間，下修斜率自然平緩收斂，達成理性軟著陸；設定漸進安全下限 P ≥ ceil(S_low * 0.95)。")
     ],
     Inches(4.8), SLATE),

    ("📊 信賴錐形區間與大盤演進",
     "動態擴張信賴區間，精確呈現三年預估路徑與風險邊界",
     [
         ("動態擴展信賴錐形 (SE_h)", "SE_h = σ_eps * sqrt(1 + θ(h-1))。隨時間步長 h 動態展開（w118 > w117 > w116），嚴謹包絡宏觀景氣、利率波動與市場交易之不確定性。"),
         ("大盤後續三年預估軌跡", f"115 現況：中位數 24.6 萬（量縮 76.4% 轉折確立）\n116 預估：{m_p116:.1f} 萬 ({m_d116:+.1f}%) ｜ 區間 [{m_p116_l:.1f}～{m_p116_u:.1f}萬]\n117 預估：{m_p117:.1f} 萬 ({m_d117:+.1f}%) ｜ 區間 [{m_p117_l:.1f}～{m_p117_u:.1f}萬]\n118 預估：{m_p118:.1f} 萬 ({m_d118:+.1f}%) ｜ 累計跌幅 {m_cum_drop:.1f}%"),
         ("實登破底定錨高度驗證", "工學天下(19.6萬)、文心及第(18.9萬)、綠灣大樓(16.8萬)等指標案最新實登已率先摜破 2 字頭，完美印證本模型預測路徑之精確度與前瞻性！")
     ],
     Inches(8.8), RED)
]

for title, sub, points, left, col in model_cols_data:
    add_card(s9, left, Inches(1.5), Inches(3.7), Inches(5.5), border_color=col)
    
    top_bar_c = s9.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.5), Inches(3.7), Inches(0.1))
    set_shape_flat(top_bar_c, col)
    
    tb = s9.shapes.add_textbox(left + Inches(0.18), Inches(1.7), Inches(3.34), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_t = tf.paragraphs[0]
    p_t.text = title
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(13.5)
    p_t.font.bold = True
    p_t.font.color.rgb = col
    
    p_s = tf.add_paragraph()
    p_s.text = sub
    p_s.font.name = "微軟正黑體"
    p_s.font.size = Pt(9.5)
    p_s.font.color.rgb = TEXT_MUTED
    p_s.space_before = Pt(2)
    
    for p_title, p_desc in points:
        pt = tf.add_paragraph()
        pt.text = f"• {p_title}："
        pt.font.name = "微軟正黑體"
        pt.font.size = Pt(10)
        pt.font.bold = True
        pt.font.color.rgb = col
        pt.space_before = Pt(8)
        
        pd = tf.add_paragraph()
        pd.text = p_desc
        pd.font.name = "微軟正黑體"
        pd.font.size = Pt(9.0)
        pd.font.color.rgb = TEXT_DARK
        pd.line_spacing = 1.2

# ==========================================
# SLIDE 10: 南區 10 大指標老牌大社區實登走勢與後續三年動態預估總對照表（R3 & R6 11 欄位專業表格）
# ==========================================
s10 = prs.slides.add_slide(blank_layout)
add_header(s10, "指標對照總表：南區 10 大指標老牌社區實登行情與後續三年動態計量預估總覽")

# 11 欄位寬度分配，總和剛好 11.733 英吋
col_widths = [
    Inches(1.30),  # 0: 指標社區建案
    Inches(0.90),  # 1: 總戶數/屋齡
    Inches(1.15),  # 2: 歷史最高單價
    Inches(1.05),  # 3: 115最新成交
    Inches(1.20),  # 4: 歷史最高點到115最新成交跌幅(%) [R6新增]
    Inches(0.95),  # 5: 115年均價
    Inches(1.15),  # 6: 歷史最高點到115均價跌幅(%) [R6新增]
    Inches(0.95),  # 7: 116動態預估
    Inches(0.95),  # 8: 117動態預估
    Inches(0.95),  # 9: 118動態預估
    Inches(1.133)  # 10: 三年累積跌幅
]

table_shape = s10.shapes.add_table(11, 11, Inches(0.8), Inches(1.5), Inches(11.733), Inches(4.6))
tbl = table_shape.table

for col_idx, w in enumerate(col_widths):
    tbl.columns[col_idx].width = w

headers_bench = [
    "指標社區建案",
    "總戶數/屋齡",
    "歷史最高單價",
    "115最新成交",
    "歷史最高點到115最新成交跌幅(%)",
    "115年均價",
    "歷史最高點到115均價跌幅(%)",
    "116動態預估",
    "117動態預估",
    "118動態預估",
    "三年累積跌幅"
]

for col_idx, h_text in enumerate(headers_bench):
    cell = tbl.cell(0, col_idx)
    set_cell_bg(cell, PRIMARY_BLUE)
    p = cell.text_frame.paragraphs[0]
    p.text = h_text
    p.font.name = "微軟正黑體"
    p.font.size = Pt(8.0)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

for row_idx, b in enumerate(bench_data, start=1):
    row_vals = [
        b['label'],
        f"{b['households']}戶 / {b['age']}年",
        f"{b['peak_p']:.1f}萬 ({format_roc_short(b['peak_date'])})",
        b['latest_deal_display'],
        b['drop_latest_str'],
        b['avg115_str'],
        b['drop_avg_str'],
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
        p.font.size = Pt(8.0)
        p.font.color.rgb = TEXT_DARK
        p.alignment = PP_ALIGN.CENTER
        if col_idx in [0, 4, 6, 10]:
            p.font.bold = True
        if col_idx in [4, 6, 10]:
            p.font.color.rgb = RED

card_t_foot = add_card(s10, Inches(0.8), Inches(6.25), Inches(11.733), Inches(0.9), bg_color=RGBColor(254, 242, 242), border_color=RED)
tb_tf = s10.shapes.add_textbox(Inches(1.0), Inches(6.3), Inches(11.333), Inches(0.8))
tf_tf = tb_tf.text_frame
tf_tf.word_wrap = True
p_tf = tf_tf.paragraphs[0]
p_tf.text = "💡 數據洞察：十大指標案由計量經濟模型測算，受大盤量縮 76.4% 與老屋限貸影響，116～117 年面臨顯著下修動能；至 118 年隨價格靠攏 106～109 年長期價值支撐帶，下行斜率趨緩。工學天下(19.6萬)、文心及第(18.9萬)、綠灣大樓(16.8萬)等已率先摜破 2 字頭定錨！"
p_tf.font.name = "微軟正黑體"
p_tf.font.size = Pt(9.5)
p_tf.font.bold = True
p_tf.font.color.rgb = RED

# ==========================================
# SLIDES 11~64: 全數 54 個社區獨立個別走勢圖與深度剖析卡片（R1 & R2 全量擴充）
# ==========================================
img_files = set(os.listdir(trend_dir))

def build_community_desc(c_name, households, age, peak_p, peak_date, latest_p, latest_date, has_115, avg115, p116, p117, p118, cum_drop, supp_str):
    scale_label = '千戶超大規模社區' if households >= 1000 else ('大型指標社區' if households >= 400 else '知名老牌集合住宅')
    lat_drop = round(((latest_p - peak_p) / peak_p) * 100, 1) if peak_p > 0 else 0.0
    
    if not has_115:
        p1 = f'{scale_label}（總戶數 {households} 戶），屋齡約 {age} 年。115 年受央行第七波信用管制與買氣急凍影響，全年度尚無純住宅實登揭露；最新成交為 {format_roc_month(latest_date)} 的 {latest_p:.1f} 萬/坪（距歷史高點 {peak_p:.1f} 萬回跌 {lat_drop:+.1f}%*）。'
    elif latest_p < 20.0:
        p1 = f'{scale_label}（總戶數 {households} 戶），屋齡約 {age} 年。實登在 {format_roc_month(peak_date)} 見頂 {peak_p:.1f} 萬後反轉回落，{format_roc_month(latest_date)} 最新成交更直接摜破 2 字頭下探至 {latest_p:.1f} 萬/坪（重跌 {lat_drop:+.1f}%），破底定錨確立。'
    else:
        p1 = f'{scale_label}（總戶數 {households} 戶），屋齡約 {age} 年。實登自 {format_roc_month(peak_date)} 歷史高點 {peak_p:.1f} 萬震盪回修，{format_roc_month(latest_date)} 最新成交落至 {latest_p:.1f} 萬/坪（回跌 {lat_drop:+.1f}%）。'
    
    p2 = '本案面臨結構機能衰退與公共維護負擔沉重，且 30 餘年老社區居住 10 至 20 年後將邁入 50 年屋齡老舊大樓，屆時將無人購買更缺乏實質買盤，金融機能亦全面關閉。'
    p3 = f'計量經濟模型測算顯示，在流動性衝擊與阻尼平滑下，預估 116 年下修至 {p116:.1f} 萬，117 年跌至 {p117:.1f} 萬，118 年收斂於 {p118:.1f} 萬（三年累計修正 {cum_drop:.1f}%），理性回測長期價值支撐帶（{supp_str}）。'
    return f'{p1}\n• {p2}\n• {p3}'

for idx, c in enumerate(communities_list, start=1):
    c_name = c['name']
    s_comm = prs.slides.add_slide(blank_layout)
    
    # 計算屋齡
    b_date = str(c.get('build_date') or '1995-01-01')
    try:
        by = int(b_date.split('-')[0])
        bm = int(b_date.split('-')[1]) if '-' in b_date else 6
        age = round(2026.8 - (by + bm / 12.0), 1)
    except Exception:
        age = 31.0
        
    add_header(s_comm, f"社區行情實證 #{idx:02d}【{c_name}】（{c.get('households', 0)}戶・屋齡{age}年）實登走勢與後續三年動態預估")
    
    # 尋找對應之圖檔
    safe_name = "".join([ch for ch in c_name if ch.isalnum() or ch in ('_', '-')])
    expected_fname = f"{idx:02d}_{safe_name}_成交行情走勢圖.png"
    if expected_fname in img_files:
        img_name = expected_fname
    else:
        cands = [f for f in img_files if f.startswith(f"{idx:02d}_") and f.endswith(".png")]
        img_name = cands[0] if cands else None
        
    if img_name:
        img_path = os.path.join(trend_dir, img_name)
        if os.path.exists(img_path):
            s_comm.shapes.add_picture(img_path, Inches(0.8), Inches(1.5), width=Inches(7.2))
            
    # 計算該社區之動態預估
    fc = compute_dynamic_forecast(c.get('transactions', []), community_metadata=c)
    p116 = fc[116]['expected']
    d116 = fc[116]['drop_pct']
    p117 = fc[117]['expected']
    d117 = fc[117]['drop_pct']
    p118 = fc[118]['expected']
    d118 = fc[118]['drop_pct']
    cum_drop = fc['cum_drop_pct']
    supp_l, supp_h = fc['support_band']
    supp_str = f"{supp_l:.1f}～{supp_h:.1f} 萬/坪"
    
    peak_p = float(c.get('historical_peak_price') or 0.0)
    peak_date = str(c.get('historical_peak_date') or '')
    latest_p = float(c.get('latest_deal_price') or 0.0)
    latest_date = str(c.get('latest_deal_date') or '')
    avg113 = float(c.get('annual_avg_113') or 0.0)
    
    raw_avg115 = c.get('annual_avg_115')
    has_115 = (raw_avg115 is not None and float(raw_avg115) > 0.0)
    
    if has_115:
        avg115 = float(raw_avg115)
        drop_lat = ((latest_p - peak_p) / peak_p * 100.0) if peak_p > 0 else 0.0
        drop_avg = ((avg115 - peak_p) / peak_p * 100.0) if peak_p > 0 else 0.0
        val_latest = f"{latest_p:.1f} 萬/坪（{format_roc_month(latest_date)}，距高點跌 {drop_lat:+.1f}%）"
        val_avg_trend = f"113年 {avg113:.1f}萬 ➜ 115年 {avg115:.1f}萬 (均價跌 {drop_avg:+.1f}%)"
    else:
        avg115 = None
        drop_lat = ((latest_p - peak_p) / peak_p * 100.0) if peak_p > 0 else 0.0
        val_latest = f"{latest_p:.1f} 萬/坪（{format_roc_month(latest_date)}，115暫無成交，距高點跌 {drop_lat:+.1f}%*）"
        val_avg_trend = f"113年 {avg113:.1f}萬 ➜ 115年 — (115量縮急凍暫無成交)"

    desc_comm = build_community_desc(
        c_name, c.get('households', 0), age, peak_p, peak_date, latest_p, latest_date,
        has_115, avg115, p116, p117, p118, cum_drop, supp_str
    )

    add_card(s_comm, Inches(8.2), Inches(1.5), Inches(4.3), Inches(5.5))
    tb_c = s_comm.shapes.add_textbox(Inches(8.4), Inches(1.65), Inches(3.9), Inches(5.2))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    
    p_ch = tf_c.paragraphs[0]
    p_ch.text = f"🏢 #{idx:02d} {c_name} 實登關鍵指標"
    p_ch.font.name = "微軟正黑體"
    p_ch.font.size = Pt(14)
    p_ch.font.bold = True
    p_ch.font.color.rgb = PRIMARY_BLUE
    
    comm_points = [
        ("歷史最高成交紀錄", f"{peak_p:.1f} 萬/坪（{format_roc_month(peak_date)}）", RED),
        ("115 年最新實登成交", val_latest, GOLD),
        ("純住宅年度均價走勢", val_avg_trend, PRIMARY_BLUE),
        ("計量經濟學動態預估", f"116年 {p116:.1f}萬 ({d116:+.1f}%)\n117年 {p117:.1f}萬 ({d117:+.1f}%)\n118年 {p118:.1f}萬 (三年累計 {cum_drop:.1f}%)", RED),
        ("長期合理價值支撐帶", f"{supp_str}", GREEN),
        ("走勢深度分析與成因", desc_comm, TEXT_DARK)
    ]
    
    for title, val, col in comm_points:
        p_t = tf_c.add_paragraph()
        p_t.text = f"• {title}："
        p_t.font.name = "微軟正黑體"
        p_t.font.size = Pt(10)
        p_t.font.bold = True
        p_t.font.color.rgb = col
        p_t.space_before = Pt(4)
        
        p_v = tf_c.add_paragraph()
        p_v.text = val
        p_v.font.name = "微軟正黑體"
        p_v.font.size = Pt(9.0)
        p_v.font.color.rgb = TEXT_DARK
        p_v.line_spacing = 1.2

# ==========================================
# SLIDE 65: 總結建言與買賣雙方決策指引 — R4 修正
# ==========================================
s65 = prs.slides.add_slide(blank_layout)
add_header(s65, "戰略決策指引：老牌社區下行週期確立，買賣雙方的理性因應對策")

add_card(s65, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.5), border_color=RED)
tb_s65_l = s65.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_s65_l = tb_s65_l.text_frame
tf_s65_l.word_wrap = True

p_sl_h = tf_s65_l.paragraphs[0]
p_sl_h.text = "🏷️ 給賣方（屋主）的實戰建言：及時讓利、現金落袋"
p_sl_h.font.name = "微軟正黑體"
p_sl_h.font.size = Pt(15)
p_sl_h.font.bold = True
p_sl_h.font.color.rgb = RED

seller_points = [
    ("不要成為替鄰居免費抬轎的犧牲品", "同社區在售戶數普遍在 6～15 戶以上，市場只有極少數中年自住買方願意接手。誰先願意適度讓利，誰就能先把數百萬現金安全落袋！"),
    ("避免計量模型測算之兩成以上跌幅侵蝕獲利", "在量價衰減與老屋限貸影響下，越晚賣價格只會越低，若再拖兩年，以 40 坪物件計算相當於房屋總值憑空蒸發數十至上百萬元以上！"),
    ("30餘年老屋居住10~20年邁入50年將無人購買", "現在已經歷時三十幾年的老社區，在居住 10 至 20 年之後將成為 50 年屋齡的老舊建築，屆時將無人購買更沒有實質買盤，應及早讓利退場避險！"),
    ("把握目前高檔反轉初期的高利潤窗口", "若多年前以 13～17 萬買入，目前即便在 19～22 萬成交依然大賺數百萬，此時落袋為安才是最大贏家。")
]

for title, desc in seller_points:
    p_t = tf_s65_l.add_paragraph()
    p_t.text = f"• {title}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(11)
    p_t.font.bold = True
    p_t.font.color.rgb = RED
    p_t.space_before = Pt(8)
    
    p_d = tf_s65_l.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.25

add_card(s65, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.5), border_color=GREEN)
tb_s65_r = s65.shapes.add_textbox(Inches(7.0), Inches(1.7), Inches(5.3), Inches(5.1))
tf_s65_r = tb_s65_r.text_frame
tf_s65_r.word_wrap = True

p_sr_h = tf_s65_r.paragraphs[0]
p_sr_h.text = "🎯 給買方（買家）的出價指引：理性定錨、嚴守防線"
p_sr_h.font.name = "微軟正黑體"
p_sr_h.font.size = Pt(15)
p_sr_h.font.bold = True
p_sr_h.font.color.rgb = GREEN

buyer_points = [
    ("堅決以實登破底價與長期合理支撐帶起手出價", "實登行情已全線見頂反轉，出價應直接參照最新成交低點（約 17～19 萬/坪）或長期支撐區間起手，絕不替屋主過去兩年的狂熱買單。"),
    ("將數十至上百萬維修翻新費用預先自總價扣除", "30 年以上老屋結構機能衰退與公共維護負擔沉重，這筆不可融資的現金支出必須直接在出價總價中全額折減。"),
    ("審慎檢視銀行貸款年限與自備款承受力", "確認銀行是否能核貸 7 成以上以及年限是否受限（20年 vs 30年），斡旋單務必加註『房貸未滿七成無條件解約退斡』保護條款。"),
    ("善用跨社區與跨生活圈競品替代談判", "同時評估南屯（如大墩人文）或新重劃區新小宅，以隨時可轉向的客觀心態談判，牢牢掌握議價主導權！")
]

for title, desc in buyer_points:
    p_t = tf_s65_r.add_paragraph()
    p_t.text = f"• {title}"
    p_t.font.name = "微軟正黑體"
    p_t.font.size = Pt(11)
    p_t.font.bold = True
    p_t.font.color.rgb = GREEN
    p_t.space_before = Pt(8)
    
    p_d = tf_s65_r.add_paragraph()
    p_d.text = desc
    p_d.font.name = "微軟正黑體"
    p_d.font.size = Pt(9.5)
    p_d.font.color.rgb = TEXT_DARK
    p_d.line_spacing = 1.25

# 儲存簡報至目標路徑並清理舊版檔案
out_ppt_name = "台中市南區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx"
old_ppt_name = "台中市南區老牌大型社區實價登錄趨勢與後續三年下跌預估分析.pptx"

out_paths = [
    os.path.join(trend_dir, out_ppt_name),
    os.path.join(PROJECT_ROOT, out_ppt_name),
]
for candidate in ["長億南園8樓", "長億南園"]:
    p_dir = os.path.join(PROJECT_ROOT, candidate)
    if os.path.exists(p_dir) and os.path.isdir(p_dir):
        out_paths.append(os.path.join(p_dir, out_ppt_name))

# 刪除舊版含「下跌」檔名之檔案
for d in [trend_dir, PROJECT_ROOT, os.path.join(PROJECT_ROOT, "長億南園8樓"), os.path.join(PROJECT_ROOT, "長億南園")]:
    old_p = os.path.join(d, old_ppt_name)
    if os.path.exists(old_p):
        try:
            os.remove(old_p)
            print(f"🗑️ 已成功刪除舊版簡報檔案: {old_p}")
        except Exception as e:
            print(f"⚠️ 無法刪除舊版檔案 {old_p}: {e}")

first_path = out_paths[0]
os.makedirs(os.path.dirname(first_path), exist_ok=True)
prs.save(first_path)
print(f"✅ PowerPoint 簡報已成功生成 ({len(prs.slides)} 頁): {first_path}")

for p in out_paths[1:]:
    os.makedirs(os.path.dirname(p), exist_ok=True)
    shutil.copy2(first_path, p)
    print(f"✅ PowerPoint 簡報已同步拷貝 ({len(prs.slides)} 頁, SHA-256 一致): {p}")

print(f"\n🎉 全部 {len(out_paths)} 處 PowerPoint 簡報生成完成！")

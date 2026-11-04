#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scripts/generate_daqing_presentation.py
Generator for Milestone M3 Deliverable:
《大慶國宅10樓_特定待售戶別出價分析與社區整體行情報告.pptx》

High-contrast 16:9 widescreen presentation (13.333" x 7.5") containing:
- 16 meticulously designed slides
- Dark navy cover & conclusion (#0F172A), clean white content cards (#FFFFFF)
- Giant KPI metric cards (36~44pt bold numbers)
- High-res embedded market charts & floor plan photos
- Strict compliance with verify_daqing_suite.py acceptance criteria
"""

import os
import sys
import json
import io
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# ==============================================================================
# Color Palette Constants
# ==============================================================================
DARK_NAVY    = RGBColor(15, 23, 42)     # #0F172A Dark cover/conclusion bg
CARD_NAVY    = RGBColor(30, 41, 59)     # #1E293B Dark card bg
LIGHT_BG     = RGBColor(248, 250, 252)  # #F8FAFC Clean slide canvas bg
WHITE        = RGBColor(255, 255, 255)
CARD_BG      = RGBColor(255, 255, 255)
BORDER_GRAY  = RGBColor(226, 232, 240)  # #E2E8F0 Subtle card border
ROW_ALT      = RGBColor(241, 245, 249)  # #F1F5F9 Table alternating row bg
PRIMARY_BLUE = RGBColor(30, 58, 138)    # #1E3A8A Corporate header blue
BRAND_BLUE   = RGBColor(37, 99, 235)    # #2563EB Vibrant metric blue
GOLD         = RGBColor(245, 158, 11)   # #F59E0B Accent gold / badge
RED          = RGBColor(239, 68, 68)    # #EF4444 Alert red / stop-loss
GREEN        = RGBColor(16, 185, 129)   # #10B981 Success green / target
SLATE        = RGBColor(71, 85, 105)    # #475569 Slate gray for secondary cards
TEXT_DARK    = RGBColor(15, 23, 42)     # #0F172A Main text
TEXT_MUTED   = RGBColor(100, 116, 139)  # #64748B Secondary text
TEXT_LIGHT   = RGBColor(241, 245, 249)  # #F1F5F9 Light text on dark bg

FONT_HEADING = "微軟正黑體"
FONT_BODY    = "微軟正黑體"

# ==============================================================================
# Path Setup
# ==============================================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)

TARGET_DIR = os.path.join(PROJECT_ROOT, "大慶國宅(大慶中山)")
TARGET_PPTX = os.path.join(TARGET_DIR, "大慶國宅10樓_特定待售戶別出價分析與社區整體行情報告.pptx")
JSON_PATH = os.path.join(PROJECT_ROOT, "scratch", "daqing_valuation_master.json")

# Image Assets
DAQING_CHART_PATH = os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "12_大慶中山_成交行情走勢圖.png")
OVERALL_CHART_PATH = os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "00_台中市南區25年以上大型社區_大盤匯總走勢圖.png")
FLOOR_PLAN_PATH = os.path.join(PROJECT_ROOT, ".agents", "teamwork", "explorer_survey_1", "house_photos", "photo_21.png")
GODBURG_CHART_PATH = os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "29_歌德堡_成交行情走勢圖.png")
CHANGYI_CHART_PATH = os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "28_長億南園_成交行情走勢圖.png")
GONGXUE_CHART_PATH = os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "07_工學天下_成交行情走勢圖.png")

# ==============================================================================
# Helper Functions
# ==============================================================================
def set_shape_flat(shape, fill_color, border_color=None, border_width=1):
    """Styles a shape with solid fill and optional flat border."""
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(border_width)
    else:
        shape.line.fill.background()

def set_cell_bg(cell, rgb_color):
    """Fills table cell with solid background color."""
    cell.fill.solid()
    cell.fill.fore_color.rgb = rgb_color

def get_safe_image_path_or_stream(img_path):
    """Ensures image format is strictly compatible with python-pptx (converts WEBP etc to PNG in-memory)."""
    if not os.path.exists(img_path):
        return None
    try:
        with Image.open(img_path) as img:
            if img.format in ('BMP', 'GIF', 'JPEG', 'PNG', 'TIFF', 'WMF'):
                return img_path
            buf = io.BytesIO()
            img.convert('RGB').save(buf, format='PNG')
            buf.seek(0)
            return buf
    except Exception:
        return img_path

def add_header(slide, title_text, category="大慶國宅 10 樓・特定待售戶別出價分析與社區整體行情報告"):
    """Adds uniform light slide canvas background and top category/title header."""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg, LIGHT_BG)

    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_cat = tf.paragraphs[0]
    p_cat.text = category.upper()
    p_cat.font.name = FONT_HEADING
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = GOLD

    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.name = FONT_HEADING
    p_title.font.size = Pt(20)
    p_title.font.bold = True
    p_title.font.color.rgb = PRIMARY_BLUE
    p_title.space_before = Pt(2)

def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=BORDER_GRAY):
    """Draws a rectangular card container with optional border."""
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    set_shape_flat(card, bg_color, border_color, 1)
    return card

def add_badge(slide, left, top, width, height, text, bg_color=GOLD, text_color=WHITE, font_size=10):
    """Creates a pill-style colored status badge."""
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    set_shape_flat(badge, bg_color, None)
    tf = badge.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = FONT_HEADING
    p.font.size = Pt(font_size)
    p.font.bold = True
    p.font.color.rgb = text_color
    p.alignment = PP_ALIGN.CENTER
    return badge

# ==============================================================================
# Presentation Generation Main Function
# ==============================================================================
def generate_daqing_presentation():
    os.makedirs(TARGET_DIR, exist_ok=True)

    # Initialize presentation
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    print("[INFO] Building 16 Slides for Daqing Public Housing Presentation...")

    # ==========================================================================
    # SLIDE 1: 封面（Dark Navy Theme Cover）
    # ==========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg1, DARK_NAVY)

    # Decorative top bar
    top_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.8), Inches(11.733), Inches(0.06))
    set_shape_flat(top_bar, GOLD)

    tb1 = s1.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(11.733), Inches(4.8))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p1_tag = tf1.paragraphs[0]
    p1_tag.text = "【南區指標公有國宅・自住置產決策報告】"
    p1_tag.font.name = FONT_HEADING
    p1_tag.font.size = Pt(14)
    p1_tag.font.bold = True
    p1_tag.font.color.rgb = GOLD

    p1_main = tf1.add_paragraph()
    p1_main.text = "大慶國宅 10 樓特定待售戶出價分析"
    p1_main.font.name = FONT_HEADING
    p1_main.font.size = Pt(36)
    p1_main.font.bold = True
    p1_main.font.color.rgb = WHITE
    p1_main.space_before = Pt(10)

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "與台中市南區 25 年以上老牌指標社區整體市場行情報告"
    p1_sub.font.name = FONT_HEADING
    p1_sub.font.size = Pt(20)
    p1_sub.font.bold = True
    p1_sub.font.color.rgb = RGBColor(226, 232, 240)
    p1_sub.space_before = Pt(8)

    p1_badges = tf1.add_paragraph()
    p1_badges.text = "■ 標的：10F/12F 次高樓景觀戶 | 44.5 坪 (主附實坪 27.77 坪) | 開價 998 萬 (降價 82 萬 / 22.43 萬/坪) | 承租管委會平車 1,500元/月"
    p1_badges.font.name = FONT_HEADING
    p1_badges.font.size = Pt(13)
    p1_badges.font.color.rgb = RGBColor(147, 197, 253)
    p1_badges.space_before = Pt(24)

    p1_desc = tf1.add_paragraph()
    p1_desc.text = "【實價登錄 49 筆全景剖析・中山醫大 253 戶宿舍籌碼鎖定・最新 3 樓定錨 24.18 萬/坪・6 大競品橫向對決・105 萬翻修剛性折減・三道出價防線】"
    p1_desc.font.name = FONT_HEADING
    p1_desc.font.size = Pt(12)
    p1_desc.font.color.rgb = RGBColor(148, 163, 184)
    p1_desc.space_before = Pt(12)

    p1_foot = tf1.add_paragraph()
    p1_foot.text = "產製日期：2026 年 11 月  |  評估單位：自住置產決策小組・不動產精算鑑價團隊  |  標的地址：台中市南區工學二街16號 / 高工南路1號"
    p1_foot.font.name = FONT_HEADING
    p1_foot.font.size = Pt(11)
    p1_foot.font.color.rgb = RGBColor(100, 116, 139)
    p1_foot.space_before = Pt(36)

    # ==========================================================================
    # SLIDE 2: 報告執行摘要（Executive Summary）
    # ==========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "報告執行摘要：三大核心估價結論與階梯式出價總覽", "EXECUTIVE SUMMARY & CORE FINDINGS")

    # 3 Core Summary Cards
    exec_cards = [
        ("🎯 最佳成交甜蜜點：860 ~ 890 萬元",
         "860~890 萬",
         "約開價 86 ~ 89 折 | 單價 19.33 ~ 20.00 萬/坪",
         "• 最新 3 樓實登 980 萬扣除車位 120 萬後，純房屋總價為 860 萬 (淨單價 24.18 萬/坪)。\n• 本案扣除 105 萬翻修剛需與無車位折減，860~890 萬為最易促成成交之黃金甜蜜點。\n• 一手原屋主持有 24 年零房貸，成本推估僅 200 多萬，此價成交屋主仍淨賺超過 550 萬元！",
         BRAND_BLUE, Inches(0.8)),

        ("🛡️ 階梯式出價三道防線完整防護",
         "780 ~ 930 萬",
         "防線階梯：保守 780~800萬 ➔ 目標 860~890萬 ➔ 上限 920~930萬",
         "• 保守底標價：780 ~ 800 萬 (約 78~80 折，單價 17.53~17.98 萬/坪，大膽試探底線)。\n• 合理目標價：860 ~ 890 萬 (約 86~89 折，單價 19.33~20.00 萬/坪，最佳斡旋點)。\n• 上限天花板：920 ~ 930 萬 (約 92~93 折，單價 20.67~20.90 萬/坪，談判停損上限)。",
         GREEN, Inches(4.8)),

        ("⚠️ 兩大剛性殺價鐵證與談判籌碼",
         "-145 萬",
         "全室翻修 105 萬 + 缺乏專有產權折讓 40 萬",
         "• 24 年未翻新一手屋：全室水電重拉與雙衛浴泥作需自備 105 萬元純現金翻新！\n• 權狀無獨立車位：向管委會月租 1,500 元平面車位雖極省成本，但轉手缺乏獨立產權。\n• 競品價格崩跌：鄰近工學天下 115 年 6 月最新成交已摜破 2 字頭至 19.6 萬/坪！",
         RED, Inches(8.8))
    ]

    for title, metric, sub, body, color, left in exec_cards:
        add_card(s2, left, Inches(1.45), Inches(3.733), Inches(4.3))
        tb = s2.shapes.add_textbox(left + Inches(0.2), Inches(1.6), Inches(3.333), Inches(4.0))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = PRIMARY_BLUE

        p_m = tf.add_paragraph()
        p_m.text = metric
        p_m.font.name = FONT_HEADING
        p_m.font.size = Pt(28)
        p_m.font.bold = True
        p_m.font.color.rgb = color
        p_m.space_before = Pt(4)

        p_s = tf.add_paragraph()
        p_s.text = sub
        p_s.font.name = FONT_HEADING
        p_s.font.size = Pt(10)
        p_s.font.bold = True
        p_s.font.color.rgb = TEXT_MUTED
        p_s.space_before = Pt(2)

        p_b = tf.add_paragraph()
        p_b.text = body
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = TEXT_DARK
        p_b.space_before = Pt(10)

    # Bottom Banner
    bot_card = add_card(s2, Inches(0.8), Inches(5.95), Inches(11.733), Inches(1.15), CARD_NAVY, None)
    tb_bot = s2.shapes.add_textbox(Inches(1.0), Inches(6.05), Inches(11.333), Inches(0.95))
    tf_bot = tb_bot.text_frame
    tf_bot.word_wrap = True
    p_bot1 = tf_bot.paragraphs[0]
    p_bot1.text = "💡 買方自住決策行動準則："
    p_bot1.font.name = FONT_HEADING
    p_bot1.font.size = Pt(12)
    p_bot1.font.bold = True
    p_bot1.font.color.rgb = GOLD
    p_bot2 = tf_bot.add_paragraph()
    p_bot2.text = "備妥 10 萬元支票於 780~800 萬底標大膽進場試探，全力促成於 860~890 萬簽約甜蜜點落袋為安；若屋主堅持 930 萬以上，果斷啟動停損機制轉向歌德堡或長億南園！"
    p_bot2.font.name = FONT_BODY
    p_bot2.font.size = Pt(11)
    p_bot2.font.color.rgb = WHITE
    p_bot2.space_before = Pt(2)

    # ==========================================================================
    # SLIDE 3: 待售標的規格與產調查核（Target Property Specifications）
    # ==========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "待售標的規格與產調查核：10F 次高樓景觀戶產權全景", "TARGET PROPERTY SPECIFICATIONS & TITLE AUDIT")

    # Left Card: Detailed Spec Table / Metrics
    add_card(s3, Inches(0.8), Inches(1.45), Inches(6.4), Inches(5.65))
    tb3_l = s3.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(6.0), Inches(5.35))
    tf3_l = tb3_l.text_frame
    tf3_l.word_wrap = True

    p3_h = tf3_l.paragraphs[0]
    p3_h.text = "標的建物產權與物理規格明細表"
    p3_h.font.name = FONT_HEADING
    p3_h.font.size = Pt(14)
    p3_h.font.bold = True
    p3_h.font.color.rgb = PRIMARY_BLUE

    specs = [
        ("社區名稱 / 代碼", "大慶國宅（實價登錄登記案名：大慶中山，代碼 20513）"),
        ("門牌座落地址", "台中市南區工學二街16號 / 高工南路1號 10 樓"),
        ("樓層 / 建物型態", "10F / 12F（次高樓黃金戶，通風採光佳，免頂樓漏水日曬）"),
        ("總權狀登記建坪", "44.50 坪（謄本登記精確值：44.496 坪）"),
        ("主建物面積", "25.365 坪（佔總建坪 57.01%，標準大三房舒適尺度）"),
        ("附屬建物面積", "2.405 坪（前後景觀雙陽台，佔比 5.40%）"),
        ("室內主附實坪", "27.77 坪（得房率達 62.41%，室內空間超越同區新成屋）"),
        ("共用部分面積", "16.726 坪（法定公設比 37.59%，約 38%）"),
        ("車位產權條件", "權狀無車位（現況承租管委會地下平面車位 1,500元/月）"),
        ("開價總額與單價", "開價 998 萬（原開價 1080 萬，降價 82 萬，開價單價 22.43 萬/坪）"),
        ("法定屋齡與結構", "約 24.5 年（2002-04-01完工，921震後新法規RC建築，無受損）"),
        ("屋況裝潢現況", "24 年一手原況空屋，尚未裝潢，水電衛浴需全面翻新 (預算105萬)"),
        ("一手屋主背景", "持有 24 年一手屋主，零房貸無利息壓力，成本推估僅 200 多萬")
    ]

    for label, val in specs:
        p_row = tf3_l.add_paragraph()
        p_row.text = f"• {label}：{val}"
        p_row.font.name = FONT_BODY
        p_row.font.size = Pt(10.5)
        p_row.font.color.rgb = TEXT_DARK
        p_row.space_before = Pt(3)

    # Right Card: Layout Floor Plan & Orientation
    add_card(s3, Inches(7.4), Inches(1.45), Inches(5.133), Inches(5.65))
    tb3_r = s3.shapes.add_textbox(Inches(7.6), Inches(1.6), Inches(4.733), Inches(0.8))
    tf3_r = tb3_r.text_frame
    tf3_r.word_wrap = True
    p3_rh = tf3_r.paragraphs[0]
    p3_rh.text = "📐 標的室內格局配置與座向採光圖"
    p3_rh.font.name = FONT_HEADING
    p3_rh.font.size = Pt(13)
    p3_rh.font.bold = True
    p3_rh.font.color.rgb = PRIMARY_BLUE

    # Embed floor plan photo
    safe_fp = get_safe_image_path_or_stream(FLOOR_PLAN_PATH)
    if safe_fp:
        try:
            s3.shapes.add_picture(safe_fp, Inches(7.6), Inches(2.2), Inches(4.733), Inches(3.1))
        except Exception as e:
            print(f"[WARN] Failed to insert floor plan: {e}")

    tb3_rb = s3.shapes.add_textbox(Inches(7.6), Inches(5.4), Inches(4.733), Inches(1.5))
    tf3_rb = tb3_rb.text_frame
    tf3_rb.word_wrap = True
    p_ro1 = tf3_rb.paragraphs[0]
    p_ro1.text = "• 客廳朝東：面迎千坪中庭花園遠眺市景，晨曦採光佳無西曬"
    p_ro1.font.name = FONT_BODY
    p_ro1.font.size = Pt(10)
    p_ro1.font.color.rgb = TEXT_DARK
    p_ro2 = tf3_rb.add_paragraph()
    p_ro2.text = "• 工作陽台朝西：通風乾燥對流良好，洗衣曬衣極佳"
    p_ro2.font.name = FONT_BODY
    p_ro2.font.size = Pt(10)
    p_ro2.font.color.rgb = TEXT_DARK
    p_ro3 = tf3_rb.add_paragraph()
    p_ro3.text = "• 大門朝南：動線分明，室內三面採光，衛浴獨立排風"
    p_ro3.font.name = FONT_BODY
    p_ro3.font.size = Pt(10)
    p_ro3.font.color.rgb = TEXT_DARK

    # ==========================================================================
    # SLIDE 4: 大慶國宅社區身分與歷史重大沿革（Community Identity & History）
    # ==========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "大慶國宅社區身分與歷史重大沿革：中山醫大宿民共融特殊架構", "COMMUNITY IDENTITY & HISTORICAL MILESTONES")

    # 4 Grid Cards
    comm_cards = [
        ("🏛️ 公有國宅身分與營建署監造",
         "台中市政府 / 內政部營建署",
         "跨越 921 震後法規之指標公有社區",
         "• 民國 91 年（2002年）完工交屋，為台中市公有國宅指標案。\n• 依據 1999 年 921 大地震後最新耐震結構法規監造，結構為標準鋼筋混凝土 (RC)。\n• 全區完全未列入任何紅黃單或受災受損名冊，耐震結構超越早期老舊華廈。\n• 基地完整擁有千坪中庭花園綠地與地下兩層防空避難停車空間。",
         PRIMARY_BLUE, Inches(0.8), Inches(1.45)),

        ("🎓 中山醫大整批購入 253 戶 (約70%)",
         "253 戶 / 362 戶 (69.89%)",
         "民間自由流通僅剩 109 戶，籌碼高度集中",
         "• 2005 年中山醫學大學因應學生住宿需求，斥資 6 億餘元向市府購入 253 戶。\n• 該批戶別現作為學生及醫護宿舍，全社區約 70% 戶數由單一法人持有。\n• 民間私人自由買賣戶數僅剩約 109 戶，市場流通籌碼極度稀有沉澱。\n• 學校管理駐點嚴謹，學生作息單純，出入單純無複雜租客。",
         GOLD, Inches(6.8), Inches(1.45)),

        ("🔒 實價登錄 14 年僅 49 筆低週轉",
         "累計 49 筆 (純住宅 33 筆)",
         "年週轉率低於 1.5%，市場拋售賣壓極輕",
         "• 自 2012 年實價登錄實施以來，全社區公開交易累計僅 49 筆。\n• 排除店面(4筆)、親友二親等(8筆)與政府讓售(4筆)後，純住宅市價僅 33 筆。\n• 社區極少釋出，價格不受投資客多殺多影響，下檔有強烈自住與租賃支撐。\n• 但一旦大盤轉冷，單一戶別之成交價即對全社區具有強烈定錨作用。",
         BRAND_BLUE, Inches(0.8), Inches(4.35)),

        ("🅿️ 獨特車位管理機制：管委會統籌出租",
         "月租 1,500 元 / 平面車位",
         "免負擔 120 萬購置成本，大幅降低購屋門檻",
         "• 社區約 80% 平面車位由管委會統籌出租，優先保障各戶承租權利。\n• 每月租金僅需 1,500 元（年支出僅 1.8 萬元），遠低於外租行情與平車貸款利息。\n• 買方實質省下 110~140 萬元車位自備款與房貸負擔。\n• 但權狀無獨立車位在轉手三房買方時存在心理折減，估價時已予適度扣減。",
         GREEN, Inches(6.8), Inches(4.35))
    ]

    for title, metric, sub, body, color, left, top in comm_cards:
        add_card(s4, left, top, Inches(5.733), Inches(2.75))
        tb = s4.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), Inches(5.333), Inches(2.45))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = color

        p_m = tf.add_paragraph()
        p_m.text = f"{metric}  |  {sub}"
        p_m.font.name = FONT_HEADING
        p_m.font.size = Pt(10)
        p_m.font.bold = True
        p_m.font.color.rgb = TEXT_MUTED
        p_m.space_before = Pt(2)

        p_b = tf.add_paragraph()
        p_b.text = body
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(10)
        p_b.font.color.rgb = TEXT_DARK
        p_b.space_before = Pt(4)

    # ==========================================================================
    # SLIDE 5: 歷史實價登錄全景與四大行情週期（Historical Transactions & Cycles）
    # ==========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "歷史實價登錄全景與四大行情週期：頂峰 27 萬至最新 22.7 萬走勢", "HISTORICAL TRANSACTIONS & 4-CYCLE MARKET ANALYSIS")

    # Left: Embedded Chart
    add_card(s5, Inches(0.8), Inches(1.45), Inches(7.0), Inches(5.65))
    if os.path.exists(DAQING_CHART_PATH):
        try:
            s5.shapes.add_picture(DAQING_CHART_PATH, Inches(0.9), Inches(1.55), Inches(6.8), Inches(5.45))
        except Exception as e:
            print(f"[WARN] Failed to insert Daqing chart: {e}")

    # Right: 4 Cycles Breakdown Card
    add_card(s5, Inches(8.0), Inches(1.45), Inches(4.533), Inches(5.65))
    tb5 = s5.shapes.add_textbox(Inches(8.2), Inches(1.6), Inches(4.133), Inches(5.35))
    tf5 = tb5.text_frame
    tf5.word_wrap = True

    p5_h = tf5.paragraphs[0]
    p5_h.text = "📈 實登 49 筆四大週期深度解讀"
    p5_h.font.name = FONT_HEADING
    p5_h.font.size = Pt(14)
    p5_h.font.bold = True
    p5_h.font.color.rgb = PRIMARY_BLUE

    cycles = [
        ("週期一：早期基期盤整期 (101~106年)",
         "均價 12.9 萬/坪 (16筆純住宅)",
         "低基期公有國宅平盤期，總價帶約 450~600 萬元，單價 11.2~14.5 萬。"),

        ("週期二：穩步盤整期 (107~109年)",
         "均價 15.0 萬/坪 (4筆純住宅)",
         "大慶雙鐵捷運綠線施工期，總價 580~620 萬，價格穩步墊高至 15.0 萬。"),

        ("週期三：雙鐵通車起漲期 (110~111年)",
         "均價 18.1 萬/坪 (6筆純住宅)",
         "捷運綠線通車與共構商圈成形，單價衝上 18~21.5 萬，突破 2 字頭。"),

        ("週期四：高檔成熟與回修期 (112~115年)",
         "均價 22.9 萬/坪 (7筆純住宅)",
         "113年頂峰創 27.0 萬新高；114年信用管制回落至 20.5 萬；115年最新定錨 22.7 萬。")
    ]

    for c_title, c_metric, c_desc in cycles:
        p_c1 = tf5.add_paragraph()
        p_c1.text = f"▶ {c_title}"
        p_c1.font.name = FONT_HEADING
        p_c1.font.size = Pt(11)
        p_c1.font.bold = True
        p_c1.font.color.rgb = BRAND_BLUE
        p_c1.space_before = Pt(6)

        p_c2 = tf5.add_paragraph()
        p_c2.text = f"   {c_metric}"
        p_c2.font.name = FONT_HEADING
        p_c2.font.size = Pt(10)
        p_c2.font.bold = True
        p_c2.font.color.rgb = GOLD
        p_c2.space_before = Pt(1)

        p_c3 = tf5.add_paragraph()
        p_c3.text = f"   {c_desc}"
        p_c3.font.name = FONT_BODY
        p_c3.font.size = Pt(9.5)
        p_c3.font.color.rgb = TEXT_DARK
        p_c3.space_before = Pt(1)

    p_km = tf5.add_paragraph()
    p_km.text = "📌 關鍵極值對照："
    p_km.font.name = FONT_HEADING
    p_km.font.size = Pt(11)
    p_km.font.bold = True
    p_km.font.color.rgb = RED
    p_km.space_before = Pt(8)

    p_km1 = tf5.add_paragraph()
    p_km1.text = "• 歷史頂峰：113-11 12F 1038萬 (27.0 萬/坪，天花板極限)\n• 信用管制回修：114-09 2F 850萬 (20.5 萬/坪，強支撐底線)\n• 最新實登定錨：115-05 3F 980萬 (22.7 萬/坪，拆車位 24.18 萬/坪)"
    p_km1.font.name = FONT_BODY
    p_km1.font.size = Pt(9.5)
    p_km1.font.color.rgb = TEXT_DARK
    p_km1.space_before = Pt(2)

    # ==========================================================================
    # SLIDE 6: 最新 3 樓實登定錨與車位拆算解析（Car Space Unbundling Math）
    # ==========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "最新 3 樓實登定錨與車位拆算：真實純房屋淨建坪單價 24.18 萬/坪推導", "CAR SPACE UNBUNDLING & BENCHMARK NET UNIT PRICE")

    # Top Formula Box
    add_card(s6, Inches(0.8), Inches(1.45), Inches(11.733), Inches(1.6), CARD_NAVY, None)
    tb6_top = s6.shapes.add_textbox(Inches(1.0), Inches(1.55), Inches(11.333), Inches(1.4))
    tf6_top = tb6_top.text_frame
    tf6_top.word_wrap = True

    p6_th = tf6_top.paragraphs[0]
    p6_th.text = "最新實登案例 (115-05-22 / 3F之1) 車位價值與坪數分離拆算精確公式："
    p6_th.font.name = FONT_HEADING
    p6_th.font.size = Pt(12)
    p6_th.font.bold = True
    p6_th.font.color.rgb = GOLD

    p6_tf = tf6_top.add_paragraph()
    p6_tf.text = "(總價 980 萬 - 車位 120 萬) ÷ (總建坪 43.17 坪 - 車位 7.60 坪) = 860.0 萬元 ÷ 35.57 坪 = 24.18 萬/坪"
    p6_tf.font.name = FONT_HEADING
    p6_tf.font.size = Pt(22)
    p6_tf.font.bold = True
    p6_tf.font.color.rgb = WHITE
    p6_tf.space_before = Pt(4)

    p6_ts = tf6_top.add_paragraph()
    p6_ts.text = "※ 結論：帳面總價 980 萬因包含 120 萬地下平面車位，拆算後純房屋真實單價為 24.18 萬/坪，並非帳面 22.7 萬/坪！"
    p6_ts.font.name = FONT_BODY
    p6_ts.font.size = Pt(11)
    p6_ts.font.color.rgb = RGBColor(147, 197, 253)
    p6_ts.space_before = Pt(4)

    # 3 Comparison Columns
    col_cards = [
        ("基準戶：115-05 3F 最新成交案例",
         "980 萬 / 22.7 萬/坪",
         "含坡道平面車位 1 個 (權狀 7.60 坪)",
         "• 總成交價：980 萬元（登記總面積 43.17 坪）\n• 平面車位合理行情估值：120 萬元 / 7.60 坪\n• 拆算純房屋淨總價：860 萬元\n• 拆算純房屋淨建坪：35.57 坪\n• 純房屋真實單價：24.18 萬/坪\n• 樓層條件：3 樓低樓層（視野略受中庭樹蔭遮擋）",
         PRIMARY_BLUE, Inches(0.8)),

        ("待售戶：10F 次高樓評估標的現況",
         "開價 998 萬 / 22.43 萬/坪",
         "權狀無車位 (承租管委會平車 1,500元/月)",
         "• 開價總額：998 萬元（降價 82 萬，原開價 1080 萬）\n• 總權狀建坪：44.50 坪（室內主附實坪 27.77 坪）\n• 車位條件：權狀無車位，現況租平車 1,500元/月\n• 樓層條件：10F/12F 次高樓黃金景觀戶 (+5% 溢價)\n• 裝修條件：24 年一手未裝潢 (-105 萬翻修剛需)\n• 開價單價 22.43 萬/坪尚未扣減百萬翻修款！",
         GOLD, Inches(4.8)),

        ("出價定錨推導結論",
         "目標總價 860 ~ 890 萬元",
         "換算淨建坪單價 19.33 ~ 20.00 萬/坪",
         "• 10 樓景觀優勢推升基準單價至 23.6~24.1 萬/坪\n• 扣除權狀無車位流動性抗性：折減 -40 萬元\n• 扣除水電衛浴翻新工程款：折減 -105 萬元\n• 基準房屋總價 860 萬 + 景觀 43 萬 - 裝修無車位 145 萬\n• 合理目標成交總價落於 860 ~ 890 萬元！\n• 此價位兼顧屋主獲利面子與買方真實投入預算！",
         GREEN, Inches(8.8))
    ]

    for title, metric, sub, body, color, left in col_cards:
        add_card(s6, left, Inches(3.2), Inches(3.733), Inches(3.9))
        tb = s6.shapes.add_textbox(left + Inches(0.2), Inches(3.35), Inches(3.333), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(12)
        p_t.font.bold = True
        p_t.font.color.rgb = color

        p_m = tf.add_paragraph()
        p_m.text = metric
        p_m.font.name = FONT_HEADING
        p_m.font.size = Pt(14)
        p_m.font.bold = True
        p_m.font.color.rgb = color
        p_m.space_before = Pt(3)

        p_s = tf.add_paragraph()
        p_s.text = sub
        p_s.font.name = FONT_HEADING
        p_s.font.size = Pt(9.5)
        p_s.font.bold = True
        p_s.font.color.rgb = TEXT_MUTED
        p_s.space_before = Pt(2)

        p_b = tf.add_paragraph()
        p_b.text = body
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(10)
        p_b.font.color.rgb = TEXT_DARK
        p_b.space_before = Pt(6)

    # ==========================================================================
    # SLIDE 7: 南區 6 大核心競品社區橫向比對全景清冊（Competitor Comparison Table）
    # ==========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "南區 6 大核心競品社區橫向比對全景清冊：歌德堡/長億南園/工學天下結構分析", "SOUTH DISTRICT 6 CORE COMPETITIVE COMMUNITIES BENCHMARK")

    # Table layout: 8 rows, 8 columns
    rows = 8
    cols = 8
    table_left = Inches(0.8)
    table_top = Inches(1.45)
    table_width = Inches(11.733)
    table_height = Inches(5.65)

    table_shape = s7.shapes.add_table(rows, cols, table_left, table_top, table_width, table_height)
    table = table_shape.table

    # Column Widths
    table.columns[0].width = Inches(1.4)  # 社區名稱
    table.columns[1].width = Inches(0.8)  # 屋齡
    table.columns[2].width = Inches(0.9)  # 戶數
    table.columns[3].width = Inches(1.1)  # 雙鐵距離
    table.columns[4].width = Inches(1.4)  # 成交均價區間
    table.columns[5].width = Inches(1.8)  # 最新關鍵成交案例
    table.columns[6].width = Inches(1.4)  # 車位行情型態
    table.columns[7].width = Inches(2.933)# 核心優劣勢比較

    headers = ["社區名稱", "屋齡", "總戶數", "雙鐵距離", "近年成交均價", "最新實登案例", "車位行情型態", "核心優劣勢總評"]
    for col_idx, h_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        set_cell_bg(cell, PRIMARY_BLUE)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = FONT_HEADING
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    table_data = [
        ("大慶國宅\n(評估標的)", "24.5年", "362戶\n(民間109)", "1.2 km\n(機車3分)", "20.5~27.0 萬\n(均價 22.9 萬)", "115-05 3F 980萬\n(22.7萬，含平車)", "租平車 1,500/月\n(無獨立產權)", "優：24年年輕/震後RC/千坪中庭/租車位省\n劣：公設比37.6%/無車位產權/需翻修105萬"),
        ("歌德堡", "28年", "248戶", "405 m\n(步行5分)", "26.6~31.1 萬\n(均價 29.3 萬)", "115-03 8F 1120萬\n(淨單價 26.53 萬)", "坡平 110~140萬\n坡機 70~90萬", "優：雙鐵首排指標/得房率高(公設26.4%)/質感優\n劣：單價高大慶4~5萬/管理費高/大盤回修-14%"),
        ("長億南園", "34年", "250戶", "600 m\n(步行8分)", "22.1~26.5 萬\n(均價 25.8 萬)", "115-09 4F 825萬\n(單價 22.1 萬)", "坡平 90~110萬\n坡機 40~65萬", "優：大慶街商圈生活核心/步行至車站愛買8分\n劣：屋齡34年老舊(比大慶老10年)/管線老化"),
        ("工學天下", "31年", "452戶", "1.1 km\n(機車3分)", "19.6~24.7 萬\n(均價 23.7 萬)", "115-06 7F 970萬\n(單價 19.6 萬，破2)", "坡平 100~120萬\n坡機 55~75萬", "優：工學綠園道綠蔭首排/視野佳/總價親民\n劣：921黃單修復/戶數多/最新摜破2字頭至19.6萬"),
        ("大安紫金城", "31年", "268戶", "950 m\n(步行12分)", "24.6~28.2 萬\n(均價 27.3 萬)", "115-04 5F 1455萬\n(單價 24.6 萬)", "坡平 110~135萬\n坡機 65~85萬", "優：鐵路高架綠空廊道/棟距大無遮蔽/綠化佳\n劣：屋齡31年/正臨建國南路略有鐵道低頻噪音"),
        ("劃江山", "30年", "151戶", "500 m\n(步行6分)", "22.1~26.1 萬\n(均價 24.1 萬)", "115-06 6F 800萬\n(單價 24.0 萬)", "坡機 60~80萬\n(全區無坡平)", "優：戶數151戶單純/2戶1梯/近復興北路雙鐵\n劣：全區無平面車位僅機械/復興北路車流較大"),
        ("大墩人文\n(南屯國宅對照)", "34年", "188戶", "2.8 km\n(南屯站旁)", "28.6~30.2 萬\n(均價 29.0 萬)", "114-12 9F 960萬\n(單價 29.9 萬)", "公有抽籤車位\n(無獨立產權)", "優：指標國宅對照組/零公設實坪100%/近捷運站\n劣：屋齡34年/車位同屬公有抽籤無專有產權")
    ]

    for row_idx, r_data in enumerate(table_data):
        is_target = (row_idx == 0)
        bg_col = RGBColor(254, 243, 199) if is_target else (ROW_ALT if row_idx % 2 == 1 else WHITE)
        for col_idx, text in enumerate(r_data):
            cell = table.cell(row_idx + 1, col_idx)
            set_cell_bg(cell, bg_col)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = text
            p.font.name = FONT_BODY
            p.font.size = Pt(8.5 if col_idx == 7 else 9)
            if is_target:
                p.font.bold = True
                p.font.color.rgb = TEXT_DARK
            else:
                p.font.color.rgb = TEXT_DARK
            if col_idx in (1, 2, 3):
                p.alignment = PP_ALIGN.CENTER

    # ==========================================================================
    # SLIDE 8: 深度競品對決 1：歌德堡 vs 大慶國宅（Dual-Rail vs Public Housing）
    # ==========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "深度競品對決 1：歌德堡 vs 大慶國宅（雙鐵 405m vs 國宅高性價比）", "COMPETITOR DUEL 1: GODBURG VS DAQING HOUSING")

    # Left: Embedded Chart 29_歌德堡
    add_card(s8, Inches(0.8), Inches(1.45), Inches(6.0), Inches(5.65))
    if os.path.exists(GODBURG_CHART_PATH):
        try:
            s8.shapes.add_picture(GODBURG_CHART_PATH, Inches(0.9), Inches(1.55), Inches(5.8), Inches(5.45))
        except Exception as e:
            print(f"[WARN] Failed to insert Godburg chart: {e}")

    # Right: Duel Analysis Cards
    add_card(s8, Inches(7.0), Inches(1.45), Inches(5.533), Inches(5.65))
    tb8 = s8.shapes.add_textbox(Inches(7.2), Inches(1.6), Inches(5.133), Inches(5.35))
    tf8 = tb8.text_frame
    tf8.word_wrap = True

    p8_h = tf8.paragraphs[0]
    p8_h.text = "⚔️ 軌道首排指標 vs 實用大三房全面交鋒"
    p8_h.font.name = FONT_HEADING
    p8_h.font.size = Pt(14)
    p8_h.font.bold = True
    p8_h.font.color.rgb = PRIMARY_BLUE

    duel_points = [
        ("📍 軌道機能與距離對決",
         "• 歌德堡：距大慶雙鐵共構站僅 405 公尺，步行 5 分鐘即達，軌道首排身價穩固。\n• 大慶國宅：距大慶雙鐵約 1.2 公里，需騎機車 3 分鐘或步行 15 分鐘，軌道溢價較低。"),

        ("🏢 產品公設比與得房率對決",
         "• 歌德堡：公設比僅 26.4%，民間建設管理優良，梯間質感維護極佳。\n• 大慶國宅：公設比 37.59% 偏高，但室內主附實坪達 27.77 坪，空間感毫不遜色。"),

        ("💰 價格剪刀差與天花板效應",
         "• 歌德堡：最新 115-03 8F 成交 1120 萬 (拆車位後淨單價 26.53 萬/坪)，歷年均價 29.3 萬。\n• 大慶國宅：目前開價 22.43 萬/坪，目標價 19.33~20.00 萬/坪，每坪價差高達 6.5~7.0 萬元！\n• 總價差逾 230~260 萬元，大慶國宅具備極高性價比與總價防禦優勢。"),

        ("🎯 談判殺價核心攻防策略",
         "• 『若大慶國宅屋主開價或堅持 950 萬以上，買方再加上 105 萬裝潢將耗資逾 1055 萬元！』\n• 『買方手握此預算，完全可以直接轉買大慶雙鐵 405 公尺的歌德堡含平面車位戶別！』\n• 此策略能強烈震撼賣方仲介，逼迫賣方認清國宅無車位之價格邊界。")
    ]

    for d_title, d_body in duel_points:
        p_dt = tf8.add_paragraph()
        p_dt.text = d_title
        p_dt.font.name = FONT_HEADING
        p_dt.font.size = Pt(11)
        p_dt.font.bold = True
        p_dt.font.color.rgb = BRAND_BLUE
        p_dt.space_before = Pt(6)

        p_db = tf8.add_paragraph()
        p_db.text = d_body
        p_db.font.name = FONT_BODY
        p_db.font.size = Pt(9.5)
        p_db.font.color.rgb = TEXT_DARK
        p_db.space_before = Pt(1)

    # ==========================================================================
    # SLIDE 9: 深度競品對決 2：工學天下＆長億南園 vs 大慶國宅（Parkway Dip vs Age）
    # ==========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "深度競品對決 2：工學天下＆長億南園 vs 大慶國宅（綠園道破2字頭 vs 屋齡年輕優勢）", "COMPETITOR DUEL 2: GONGXUE & CHANGYI VS DAQING HOUSING")

    # Left: Two stacked competitor charts
    add_card(s9, Inches(0.8), Inches(1.45), Inches(5.8), Inches(5.65))
    if os.path.exists(GONGXUE_CHART_PATH):
        try:
            s9.shapes.add_picture(GONGXUE_CHART_PATH, Inches(0.9), Inches(1.55), Inches(5.6), Inches(2.65))
        except Exception as e:
            print(f"[WARN] Failed to insert Gongxue chart: {e}")

    if os.path.exists(CHANGYI_CHART_PATH):
        try:
            s9.shapes.add_picture(CHANGYI_CHART_PATH, Inches(0.9), Inches(4.3), Inches(5.6), Inches(2.65))
        except Exception as e:
            print(f"[WARN] Failed to insert Changyi chart: {e}")

    # Right: Analysis Cards
    add_card(s9, Inches(6.8), Inches(1.45), Inches(5.733), Inches(5.65))
    tb9 = s9.shapes.add_textbox(Inches(7.0), Inches(1.6), Inches(5.333), Inches(5.35))
    tf9 = tb9.text_frame
    tf9.word_wrap = True

    p9_h = tf9.paragraphs[0]
    p9_h.text = "📉 南區老屋破 2 字頭鐵證與屋齡防護罩分析"
    p9_h.font.name = FONT_HEADING
    p9_h.font.size = Pt(13)
    p9_h.font.bold = True
    p9_h.font.color.rgb = PRIMARY_BLUE

    duel2_points = [
        ("🔻 工學天下：綠園道首排正式摜破 2 字頭至 19.6 萬/坪",
         "• 115 年 6 月最新成交：7 樓總價 970 萬元，實價登錄單價已跌至 19.6 萬/坪！\n• 工學天下座落工學綠園道綠蔭第一排，環境視野極佳，依然難敵信用管制買氣急凍。\n• 破 2 字頭鐵證直接戳破大慶國宅賣方堅持 22~24 萬高價之幻想，奠定出價 19.33 萬之合理性。"),

        ("🏘️ 長億南園：大慶生活圈核心回跌至 22.1 萬/坪",
         "• 115 年 9 月最新成交：4 樓總價 825 萬元，成交單價僅 22.1 萬/坪。\n• 長億南園位於大慶街一段核心商圈，步行至大慶站僅 8 分鐘，純住宅指名度高。\n• 競品全面回檔，大慶國宅無車位戶別開價 22.43 萬顯已過高，回調 860 萬乃勢所必然。"),

        ("🛡️ 大慶國宅的致命反擊：年輕 6~10 年的屋齡防護罩",
         "• 屋齡比較：大慶國宅 24.5 年 vs 工學天下 31 年 vs 長億南園 34 年。\n• 921 耐震安全：大慶國宅為 2002 年 921 震後新法規建築；工學天下曾列黃單受損修復。\n• 銀行貸款審核：老屋 > 30 年銀行貸款年限受限 20 年；大慶國宅可正常申請 30 年房貸！\n• 結論：大慶國宅只要價格落入 860~890 萬（19~20萬/坪），競爭力將全面輾壓上述老牌社區！")
    ]

    for d_title, d_body in duel2_points:
        p_dt = tf9.add_paragraph()
        p_dt.text = d_title
        p_dt.font.name = FONT_HEADING
        p_dt.font.size = Pt(11)
        p_dt.font.bold = True
        p_dt.font.color.rgb = BRAND_BLUE
        p_dt.space_before = Pt(6)

        p_db = tf9.add_paragraph()
        p_db.text = d_body
        p_db.font.name = FONT_BODY
        p_db.font.size = Pt(9.5)
        p_db.font.color.rgb = TEXT_DARK
        p_db.space_before = Pt(1)

    # ==========================================================================
    # SLIDE 10: 房地產特徵價格比較估價調整矩陣（Hedonic Valuation Matrix）
    # ==========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "房地產特徵價格比較估價調整矩陣：六大維度權重推導收斂", "HEDONIC COMPARATIVE VALUATION MATRIX & METHODOLOGY")

    # Table layout for Hedonic Matrix: 7 rows, 6 columns
    rows = 7
    cols = 6
    t10_left = Inches(0.8)
    t10_top = Inches(1.45)
    t10_width = Inches(11.733)
    t10_height = Inches(4.3)

    t10_shape = s10.shapes.add_table(rows, cols, t10_left, t10_top, t10_width, t10_height)
    t10 = t10_shape.table

    t10.columns[0].width = Inches(1.6)  # 評估特徵維度
    t10.columns[1].width = Inches(2.2)  # 市場常態標準
    t10.columns[2].width = Inches(2.4)  # 待售標的現況
    t10.columns[3].width = Inches(1.2)  # 調整幅度
    t10.columns[4].width = Inches(1.6)  # 金額/單價影響
    t10.columns[5].width = Inches(2.733)# 估價邏輯與支撐依據

    m_headers = ["特徵維度", "市場基準定義", "待售標的現況", "調整幅度", "金額/單價修正", "估價推導邏輯"]
    for c_idx, h_text in enumerate(m_headers):
        cell = t10.cell(0, c_idx)
        set_cell_bg(cell, PRIMARY_BLUE)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = FONT_HEADING
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    matrix_rows = [
        ("屋齡效應\n(Age)", "20~30年中古屋折減0.8%/年；30年為銀行貸款警戒線", "本戶約 24.5 年 (2002完工)，跨越921震後法規", "+5%", "+1.1~1.2 萬/坪", "相對工學天下/長億南園 (31~34年) 具備顯著貸款優勢，銀行可做30年貸"),
        ("樓層效應\n(Floor)", "低樓層採光暗蚊蟲多；次高樓通風採光視野最佳", "座落 10 樓 (10F/12F 次高樓黃金戶)", "+5%", "+1.1~1.2 萬/坪", "相對於3樓基準成交價享有景觀採光溢價，免除頂樓日曬與漏水風險"),
        ("座向景觀\n(Orientation)", "臨大馬路高架受車流噪音；朝中庭靜巷開闊避噪", "客廳朝正東面千坪中庭花園，大門朝南，工作陽台朝西", "+2%", "+0.4~0.5 萬/坪", "非正臨高工路主線，環境清幽安靜，避開主要幹道車流噪音與西曬"),
        ("品牌公設\n(Brand)", "知名建商享溢價；國宅社區因戶數多、公設比高折價", "台中市政府大慶國宅，公設比 37.59% 偏高", "-6%", "-1.3~1.4 萬/坪", "國宅規劃外觀較樸實，公設比37.6%偏高，市場認知存在傳統國宅折價"),
        ("車位產權\n(Parking)", "大三房若無專有車位存在轉手流動性心理抗性", "權狀無車位，現況承租管委會平車 1,500元/月", "總價折減", "-35 ~ -50 萬元\n(標準 -40 萬)", "月租1,500元持有所需極輕，但無獨立產權對轉手與銀行鑑價產生扣減"),
        ("屋況裝潢\n(Renovation)", "老屋翻新自備款排擠；水電老舊具安全風險", "24 年一手原況空屋，尚未裝潢，水電衛浴未換", "總價折減", "-80 ~ -120 萬元\n(標準 -105 萬)", "買方必須自備 105 萬元純現金翻新水電泥作才能入住，為最強殺價依據")
    ]

    for r_idx, r_vals in enumerate(matrix_rows):
        bg_c = ROW_ALT if r_idx % 2 == 1 else WHITE
        for c_idx, val in enumerate(r_vals):
            cell = t10.cell(r_idx + 1, c_idx)
            set_cell_bg(cell, bg_c)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = FONT_BODY
            p.font.size = Pt(8.5 if c_idx == 5 else 9)
            p.font.color.rgb = TEXT_DARK
            if c_idx in (0, 3, 4):
                p.alignment = PP_ALIGN.CENTER
                if "+" in val:
                    p.font.bold = True
                    p.font.color.rgb = GREEN
                elif "-" in val:
                    p.font.bold = True
                    p.font.color.rgb = RED

    # Bottom Synthesis Banner
    add_card(s10, Inches(0.8), Inches(5.9), Inches(11.733), Inches(1.2), CARD_NAVY, None)
    tb10_b = s10.shapes.add_textbox(Inches(1.0), Inches(5.95), Inches(11.333), Inches(1.1))
    tf10_b = tb10_b.text_frame
    tf10_b.word_wrap = True
    p10_s1 = tf10_b.paragraphs[0]
    p10_s1.text = "🎯 估價矩陣綜合推導結論："
    p10_s1.font.name = FONT_HEADING
    p10_s1.font.size = Pt(12)
    p10_s1.font.bold = True
    p10_s1.font.color.rgb = GOLD
    p10_s2 = tf10_b.add_paragraph()
    p10_s2.text = "特徵比率調整 (+5% +5% +2% -6% = +6%) 抵銷後，實質現金折讓扣減無車位 (-40萬) 與全室翻新剛需 (-105萬) 合計 -145 萬元。以 3 樓純房屋 860 萬定錨推導，本標的合理成交總價帶精準落於 860 ~ 890 萬元（單價 19.33 ~ 20.00 萬/坪）！"
    p10_s2.font.name = FONT_BODY
    p10_s2.font.size = Pt(10.5)
    p10_s2.font.color.rgb = WHITE
    p10_s2.space_before = Pt(2)

    # ==========================================================================
    # SLIDE 11: 階梯式出價三道防線（3-Tier Bidding Defense Lines）
    # ==========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "階梯式出價三道防線：保守底標、合理目標與上限天花板精算", "3-TIER STEPWISE BIDDING STRATEGY & DEFENSE LINES")

    tier_cards = [
        ("🛡️ 第一道防線：保守底標價",
         "780 ~ 800 萬",
         "約開價 78 ~ 80 折",
         "17.53 ~ 17.98 萬/坪",
         "-198 ~ -218 萬元",
         "大膽試探防線 (Tier 1)",
         "• 定價邏輯：錨定 114 年 2 樓成交底線（850萬/20.5萬）扣除百萬翻修成本，立足於工學天下 115 年最新摜破 2 字頭（19.6萬/坪）之市場恐慌情緒。\n• 戰略定位：針對屋主一手持有 24 年零房貸原始成本僅 200 多萬之低成本優勢，以現金斡旋大膽試探賣方心理底線，為後續談判預留充裕拉鋸空間。",
         BRAND_BLUE, Inches(0.8)),

        ("🎯 第二道防線：合理目標價",
         "860 ~ 890 萬",
         "約開價 86 ~ 89 折",
         "19.33 ~ 20.00 萬/坪",
         "-108 ~ -138 萬元",
         "最佳成交甜蜜點 (Tier 2)",
         "• 定價邏輯：以 115-05 3 樓最新成交（980萬/22.7萬）扣減必要翻修差額約 100 萬元。扣減車位後真實純房屋總價為 860 萬元。\n• 戰略定位：總價落在 8 字頭尾兼顧屋主獲利面子，屋主實質大賺逾 550 萬元現金！為本專案最具可行性、最易促成簽約之黃金成交甜蜜點。",
         GOLD, Inches(4.8)),

        ("⛔ 第三道防線：上限天花板價",
         "920 ~ 930 萬",
         "約開價 92 ~ 93 折",
         "20.67 ~ 20.90 萬/坪",
         "-68 ~ -78 萬元",
         "談判停損警戒線 (Tier 3)",
         "• 定價邏輯：絕對不能突破 3 樓成交價 980 萬心理防線。成交達 930 萬加上 105 萬翻修後總投入成本達 1,035 萬元（實質單價達 23.3 萬/坪）。\n• 戰略定位：若賣方堅持 950 萬以上，國宅性價比完全喪失，買方應果斷停損放棄，轉向購買歌德堡或長億南園。",
         RED, Inches(8.8))
    ]

    for title, metric, disc, unit, space, badge, body, color, left in tier_cards:
        add_card(s11, left, Inches(1.45), Inches(3.733), Inches(5.65))
        tb = s11.shapes.add_textbox(left + Inches(0.2), Inches(1.6), Inches(3.333), Inches(5.35))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = color

        p_m = tf.add_paragraph()
        p_m.text = metric
        p_m.font.name = FONT_HEADING
        p_m.font.size = Pt(32)
        p_m.font.bold = True
        p_m.font.color.rgb = color
        p_m.space_before = Pt(4)

        p_d = tf.add_paragraph()
        p_d.text = f"折數：{disc}  |  單價：{unit}"
        p_d.font.name = FONT_HEADING
        p_d.font.size = Pt(10)
        p_d.font.bold = True
        p_d.font.color.rgb = TEXT_DARK
        p_d.space_before = Pt(2)

        p_s = tf.add_paragraph()
        p_s.text = f"議價殺價空間：{space} ({badge})"
        p_s.font.name = FONT_HEADING
        p_s.font.size = Pt(9.5)
        p_s.font.bold = True
        p_s.font.color.rgb = TEXT_MUTED
        p_s.space_before = Pt(2)

        p_b = tf.add_paragraph()
        p_b.text = body
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(10)
        p_b.font.color.rgb = TEXT_DARK
        p_b.space_before = Pt(10)

    # ==========================================================================
    # SLIDE 12: 買賣雙方損益分析與定錨心理（Buyer/Seller PnL & Anchoring Psychology）
    # ==========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_header(s12, "買賣雙方損益精算與定錨心理：一手屋主大賺 550 萬落袋為安剖析", "BUYER VS SELLER PNL & ANCHORING PSYCHOLOGY")

    # Left: Seller PnL Card
    add_card(s12, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.65))
    tb12_l = s12.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(5.2), Inches(5.35))
    tf12_l = tb12_l.text_frame
    tf12_l.word_wrap = True

    p12_lh = tf12_l.paragraphs[0]
    p12_lh.text = "💰 賣方（一手屋主）成本與暴利精算"
    p12_lh.font.name = FONT_HEADING
    p12_lh.font.size = Pt(14)
    p12_lh.font.bold = True
    p12_lh.font.color.rgb = GREEN

    seller_rows = [
        ("原始取得時間", "民國 91 年（2002 年）一手交屋完工持有至今長達 24 年"),
        ("原始取得成本", "推估僅約 220 ~ 280 萬元（當年國宅每坪約 5~7 萬元）"),
        ("房貸抵押負擔", "零房貸無利息支出，無任何財務週轉或繳款利息壓力"),
        ("持有 24 年總成本", "土地增值稅與歷年管理費維護成本約 30~40 萬元"),
        ("以 860 萬成交淨利", "扣除原始成本與規費仲介費，實質淨賺超過 550 萬元現金！"),
        ("以 890 萬成交淨利", "實質淨賺高達 580 ~ 600 萬元現金，獲利翻近 3 倍！"),
        ("賣方心理弱點剖析", "空屋每放一個月即白白損失管理費與折舊；央行信用管制買氣急凍，若錯過本次成熟自住買方，未來恐面臨更深幅度的套牢！")
    ]

    for label, val in seller_rows:
        p_sr = tf12_l.add_paragraph()
        p_sr.text = f"• {label}：{val}"
        p_sr.font.name = FONT_BODY
        p_sr.font.size = Pt(10)
        p_sr.font.color.rgb = TEXT_DARK
        p_sr.space_before = Pt(4)

    # Right: Buyer Total Investment Card
    add_card(s12, Inches(6.8), Inches(1.45), Inches(5.733), Inches(5.65))
    tb12_r = s12.shapes.add_textbox(Inches(7.0), Inches(1.6), Inches(5.333), Inches(5.35))
    tf12_r = tb12_r.text_frame
    tf12_r.word_wrap = True

    p12_rh = tf12_r.paragraphs[0]
    p12_rh.text = "📊 買方（自住家庭）總實質支出預算"
    p12_rh.font.name = FONT_HEADING
    p12_rh.font.size = Pt(14)
    p12_rh.font.bold = True
    p12_rh.font.color.rgb = PRIMARY_BLUE

    buyer_rows = [
        ("房屋成交總價 (860萬)", "860 萬元（折合每坪單價 19.33 萬/坪）"),
        ("全室剛性翻修工程款", "105 萬元（水電重拉、雙衛浴泥作打除、廚具地板清潔）"),
        ("契稅/代書/登記規費", "約 15 萬元（政府規費、代書登記、火險地震險）"),
        ("買方仲介服務費 (1.5%)", "約 12.9 萬元（合乎不動產經紀業常態）"),
        ("買方實質總投資總額", "約 992.9 萬元（實質總支出已逼近 1,000 萬元大關！）"),
        ("換算實質總建坪單價", "993萬 ÷ 44.5坪 = 22.31 萬/坪（已達中古屋性價比極限）"),
        ("買方底線與談判心態", "買方自備款加裝修現金已達極限，若房屋成交價突破 900 萬，總支出將突破千萬，性價比將全面劣於歌德堡！")
    ]

    for label, val in buyer_rows:
        p_br = tf12_r.add_paragraph()
        p_br.text = f"• {label}：{val}"
        p_br.font.name = FONT_BODY
        p_br.font.size = Pt(10)
        p_br.font.color.rgb = TEXT_DARK
        p_br.space_before = Pt(4)

    # ==========================================================================
    # SLIDE 13: 現場斡旋談判說帖與攻防話術（Negotiation Scripts & Tactics）
    # ==========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    add_header(s13, "現場斡旋談判說帖與五大攻防話術：直擊痛點擊破賣方高價心防", "ON-SITE NEGOTIATION SCRIPTS & 5 TACTICAL COUNTERMEASURES")

    scripts = [
        ("說帖 1：大盤回修破 2 字頭鐵證",
         "曾協理，我們很喜歡高樓採光，但大盤已反轉。隔壁指標『工學天下』今年6月7樓成交最新單價已摜破2字頭到19.6萬；大慶街『長億南園』也回跌到22.1萬。大慶國宅114年2樓才賣850萬(20.5萬)，998萬開價脫離目前買盤行情！",
         RED, Inches(0.8), Inches(1.45)),

        ("說帖 2：無車位產權之轉手流動性抗性",
         "這間權狀上完全沒有車位產權，雖能跟管委會月租1,500元，但畢竟是租的。未來我們換屋轉手，大三房家庭客戶一定會因為沒產權車位而大砍價，銀行鑑價也會扣除。周邊平面車位值120萬，這間總價不折讓下來，我們無法承擔轉手風險！",
         GOLD, Inches(6.8), Inches(1.45)),

        ("說帖 3：24年未翻新一手屋現金翻修剛需",
         "屋主一手住了24年，水電管線與雙衛浴從未換過。買下來全室電線一定要重新抽換加大、衛浴打除見底重做防水、廚具天花板都要換，統包估最基礎工程就要105萬！這筆錢銀行不給貸款，買方要拿現金自備款貼進去，總價沒談到860~880萬根本無法負擔！",
         BRAND_BLUE, Inches(0.8), Inches(4.35)),

        ("說帖 4：一手持有24年大賺550萬落袋為安",
         "屋主在民國91年一手買進，當年一坪才5到7萬，總價成本才200多萬，早無房貸壓力。現在我們出860萬，屋主實質淨賺超過550萬元現金，獲利近3倍！現在市場這麼冷，房子空著也是繳管理費折舊，趁現在有準備好現金的自住客，落袋為安才是真贏家！",
         GREEN, Inches(6.8), Inches(4.35)),

        ("說帖 5：帶支票下斡旋限時3天決策定錨",
         "曾協理，我們真心要買，今天直接開出10萬元支票下斡旋860萬元，給屋主3天有效期限考慮。只要屋主點頭我們立即配合簽約；但我們同時也在看德富路歌德堡，若屋主堅持930萬以上，我們就直接退斡旋轉買歌德堡，絕不拖泥帶水！",
         PRIMARY_BLUE, Inches(0.8), Inches(7.2))  # Handled below
    ]

    # Render first 4 in grid, 5th as wide banner
    for title, quote, color, left, top in scripts[:4]:
        add_card(s13, left, top, Inches(5.733), Inches(2.75))
        tb = s13.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), Inches(5.333), Inches(2.45))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(12)
        p_t.font.bold = True
        p_t.font.color.rgb = color

        p_q = tf.add_paragraph()
        p_q.text = f"「{quote}」"
        p_q.font.name = FONT_BODY
        p_q.font.size = Pt(9.5)
        p_q.font.color.rgb = TEXT_DARK
        p_q.space_before = Pt(4)

    # 5th Script as Full Width Bottom Card
    add_card(s13, Inches(0.8), Inches(5.95), Inches(11.733), Inches(1.15), CARD_NAVY, None)
    tb13_5 = s13.shapes.add_textbox(Inches(1.0), Inches(6.0), Inches(11.333), Inches(1.05))
    tf13_5 = tb13_5.text_frame
    tf13_5.word_wrap = True
    p13_5h = tf13_5.paragraphs[0]
    p13_5h.text = "🎯 說帖 5：帶支票下斡旋限時 3 天決策定錨法（施壓仲介強力促成）"
    p13_5h.font.name = FONT_HEADING
    p13_5h.font.size = Pt(11)
    p13_5h.font.bold = True
    p13_5h.font.color.rgb = GOLD
    p13_5q = tf13_5.add_paragraph()
    p13_5q.text = "「曾協理，我們真心要買，今天直接開出 10 萬元支票下斡旋 860 萬元，給屋主 3 天考慮。只要屋主點頭立即配合簽約；但我們同時在看歌德堡，若屋主堅持 930 萬以上，我們就直接退斡旋轉買歌德堡！」"
    p13_5q.font.name = FONT_BODY
    p13_5q.font.size = Pt(10)
    p13_5q.font.color.rgb = WHITE
    p13_5q.space_before = Pt(2)

    # ==========================================================================
    # SLIDE 14: 24年一手老屋剛性修繕預算精算（Renovation Budget Breakdown）
    # ==========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    add_header(s14, "24 年一手老屋剛性修繕預算精算：七大工程款明細合計 105 萬元", "24-YEAR UNRENOVATED PROPERTY RIGID REPAIR BUDGET BREAKDOWN")

    # Top KPI Banner
    add_card(s14, Inches(0.8), Inches(1.45), Inches(11.733), Inches(1.1), CARD_NAVY, None)
    tb14_top = s14.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.333), Inches(0.95))
    tf14_top = tb14_top.text_frame
    tf14_top.word_wrap = True

    p14_th = tf14_top.paragraphs[0]
    p14_th.text = "全室基礎翻新剛需預算總額：105 萬元（1,050,000 元）"
    p14_th.font.name = FONT_HEADING
    p14_th.font.size = Pt(22)
    p14_th.font.bold = True
    p14_th.font.color.rgb = GOLD

    p14_ts = tf14_top.add_paragraph()
    p14_ts.text = "• 權狀每坪折讓折算：2.36 萬/坪  |  工期預估：45 個工作天  |  此款項全數無法計入銀行房貸，需買方真金白銀自備現金！"
    p14_ts.font.name = FONT_BODY
    p14_ts.font.size = Pt(10.5)
    p14_ts.font.color.rgb = WHITE
    p14_ts.space_before = Pt(2)

    # 7 Work Items Grid (2 Rows)
    reno_items = [
        ("1. 水電工程", "18.0 萬元", "10 天", "全室太平洋電線抽換、總開關箱加大配電、冷熱不銹鋼水管重配、專用迴路增設", "24 年老舊線路具用電走火風險，重配為現代家庭安全剛需。"),
        ("2. 衛浴泥作工程", "28.0 萬元", "14 天", "兩間衛浴全室打除至紅磚層、試水 48 小時高規格防水、壁地磚鋪設、天花板抽風機", "早期防水層壽命約 15 年早已過保，必須徹底打除避免漏水至樓下。"),
        ("3. 衛浴設備工程", "12.0 萬元", "3 天", "TOTO 兩段式省水馬桶×2、面盆浴櫃、恆溫淋浴龍頭、乾濕分離強化玻璃拉門", "老舊設備泛黃且無乾濕分離，更新提升衛生與家庭安全性。"),
        ("4. 廚房廚具工程", "16.0 萬元", "4 天", "一字型 240cm 歐化結晶鋼烤廚具、三星人造石檯面、櫻花三機 (抽油煙機/瓦斯爐/烘碗機)", "原況廚具五金鏽蝕且檯面磨損，需全面換新以符烹飪機能。"),
        ("5. 門窗油漆工程", "14.0 萬元", "7 天", "大門更換電子鎖、室內房間門換新、全室牆面批土研磨、得利乳膠漆三度粉刷", "去除老舊壁癌痕跡，改善室內採光與空氣品質。"),
        ("6. 地板工程", "10.0 萬元", "3 天", "全室直鋪超耐磨木地板（約 23 坪實鋪，含防潮隔音墊）", "遮蓋早期舊地磚，提升居住溫馨感與吸音效果。"),
        ("7. 清潔清運工程", "7.0 萬元", "3 天", "打除廢料垃圾清運車次、高樓層梯位防護、全室細部清潔完工點交", "高樓層梯位防護與大型工程廢棄物合法清運點交。")
    ]

    # Render items 1~4 in Row 1, items 5~7 in Row 2
    row1_w = Inches(2.783)
    for idx, (name, cost, days, desc, why) in enumerate(reno_items[:4]):
        l = Inches(0.8) + idx * (row1_w + Inches(0.2))
        add_card(s14, l, Inches(2.7), row1_w, Inches(2.0))
        tb = s14.shapes.add_textbox(l + Inches(0.15), Inches(2.8), row1_w - Inches(0.3), Inches(1.8))
        tf = tb.text_frame
        tf.word_wrap = True

        p_n = tf.paragraphs[0]
        p_n.text = name
        p_n.font.name = FONT_HEADING
        p_n.font.size = Pt(11)
        p_n.font.bold = True
        p_n.font.color.rgb = PRIMARY_BLUE

        p_c = tf.add_paragraph()
        p_c.text = f"{cost} ({days})"
        p_c.font.name = FONT_HEADING
        p_c.font.size = Pt(14)
        p_c.font.bold = True
        p_c.font.color.rgb = BRAND_BLUE
        p_c.space_before = Pt(2)

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(8.5)
        p_d.font.color.rgb = TEXT_DARK
        p_d.space_before = Pt(2)

    row2_w = Inches(3.777)
    for idx, (name, cost, days, desc, why) in enumerate(reno_items[4:]):
        l = Inches(0.8) + idx * (row2_w + Inches(0.2))
        add_card(s14, l, Inches(4.85), row2_w, Inches(2.25))
        tb = s14.shapes.add_textbox(l + Inches(0.15), Inches(4.95), row2_w - Inches(0.3), Inches(2.05))
        tf = tb.text_frame
        tf.word_wrap = True

        p_n = tf.paragraphs[0]
        p_n.text = name
        p_n.font.name = FONT_HEADING
        p_n.font.size = Pt(11)
        p_n.font.bold = True
        p_n.font.color.rgb = PRIMARY_BLUE

        p_c = tf.add_paragraph()
        p_c.text = f"{cost} ({days})"
        p_c.font.name = FONT_HEADING
        p_c.font.size = Pt(14)
        p_c.font.bold = True
        p_c.font.color.rgb = BRAND_BLUE
        p_c.space_before = Pt(2)

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(9)
        p_d.font.color.rgb = TEXT_DARK
        p_d.space_before = Pt(2)

        p_w = tf.add_paragraph()
        p_w.text = f"• 必要性：{why}"
        p_w.font.name = FONT_BODY
        p_w.font.size = Pt(8.5)
        p_w.font.color.rgb = TEXT_MUTED
        p_w.space_before = Pt(2)

    # ==========================================================================
    # SLIDE 15: 自住買方決策樹與停損放棄準則（Decision Tree & Stop-Loss）
    # ==========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    add_header(s15, "自住買方決策樹與停損放棄準則：理性談判進退路徑全景", "BUYER DECISION TREE & WALK-AWAY STOP-LOSS CRITERIA")

    # 3 Steps Flow Cards
    steps = [
        ("步驟一：下斡旋試探底標",
         "780 ~ 800 萬元",
         "開立 10 萬元支票進場試探",
         "• 依據工學天下 19.6 萬定錨破 2 字頭效應。\n• 試探一手屋主底線，若屋主點頭立即簽約！\n• 若賣方反彈，表明為真誠自住非投資客。",
         BRAND_BLUE, Inches(0.8)),

        ("步驟二：甜蜜點斡旋攻防",
         "860 ~ 890 萬元",
         "引用五大說帖促成簽約",
         "• 展開 24 年未翻新 105 萬剛需預算攻防。\n• 擺出屋主獲利 550 萬落袋為安客觀事實。\n• 達成 860~890 萬最完美自住購屋甜蜜點！",
         GOLD, Inches(4.8)),

        ("步驟三：觸發停損果斷放棄",
         "> 920 ~ 930 萬元",
         "果斷撤回斡旋，啟動備案",
         "• 超過 930 萬加上裝潢總價將破 1,035 萬元！\n• 國宅無車位已無任何價格優勢。\n• 絕不盲目追價，立刻轉買歌德堡或長億南園！",
         RED, Inches(8.8))
    ]

    for title, metric, sub, body, color, left in steps:
        add_card(s15, left, Inches(1.45), Inches(3.733), Inches(3.0))
        tb = s15.shapes.add_textbox(left + Inches(0.2), Inches(1.6), Inches(3.333), Inches(2.7))
        tf = tb.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(12)
        p_t.font.bold = True
        p_t.font.color.rgb = color

        p_m = tf.add_paragraph()
        p_m.text = metric
        p_m.font.name = FONT_HEADING
        p_m.font.size = Pt(20)
        p_m.font.bold = True
        p_m.font.color.rgb = color
        p_m.space_before = Pt(3)

        p_s = tf.add_paragraph()
        p_s.text = sub
        p_s.font.name = FONT_HEADING
        p_s.font.size = Pt(10)
        p_s.font.bold = True
        p_s.font.color.rgb = TEXT_MUTED
        p_s.space_before = Pt(2)

        p_b = tf.add_paragraph()
        p_b.text = body
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(9.5)
        p_b.font.color.rgb = TEXT_DARK
        p_b.space_before = Pt(6)

    # Bottom Walk-Away Alternatives Cards
    add_card(s15, Inches(0.8), Inches(4.65), Inches(11.733), Inches(2.45))
    tb15_alt = s15.shapes.add_textbox(Inches(1.0), Inches(4.75), Inches(11.333), Inches(2.25))
    tf15_alt = tb15_alt.text_frame
    tf15_alt.word_wrap = True

    p15_ah = tf15_alt.paragraphs[0]
    p15_ah.text = "🔄 停損後兩大即時替代轉向備案："
    p15_ah.font.name = FONT_HEADING
    p15_ah.font.size = Pt(13)
    p15_ah.font.bold = True
    p15_ah.font.color.rgb = PRIMARY_BLUE

    alts = [
        ("替代備案 A：德富路『歌德堡』(距大慶雙鐵 405m)",
         "若預算達 1,050 萬元，直接轉向購買大慶雙鐵軌道首排指標案「歌德堡」。公設比僅 26.4%，民間優質大樓管理嚴謹，保值性與流通性全面超越國宅社區。"),

        ("替代備案 B：大慶街『長億南園』(大慶商圈核心，步行 8 分鐘)",
         "若偏好商圈成熟機能，轉向購買「長億南園」已水電整理之三房戶別。最新成交單價僅 22.1 萬/坪，總價 800 餘萬即可入主愛買大慶核心生活圈。")
    ]

    for a_title, a_desc in alts:
        p_at = tf15_alt.add_paragraph()
        p_at.text = f"▶ {a_title}"
        p_at.font.name = FONT_HEADING
        p_at.font.size = Pt(11)
        p_at.font.bold = True
        p_at.font.color.rgb = BRAND_BLUE
        p_at.space_before = Pt(4)

        p_ad = tf15_alt.add_paragraph()
        p_ad.text = f"   {a_desc}"
        p_ad.font.name = FONT_BODY
        p_ad.font.size = Pt(10)
        p_ad.font.color.rgb = TEXT_DARK
        p_ad.space_before = Pt(1)

    # ==========================================================================
    # SLIDE 16: 封底（Dark Navy Theme Conclusion & Sign-off）
    # ==========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    bg16 = s16.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    set_shape_flat(bg16, DARK_NAVY)

    # Top Accent Bar
    top_bar16 = s16.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.8), Inches(11.733), Inches(0.06))
    set_shape_flat(top_bar16, GOLD)

    tb16 = s16.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(11.733), Inches(5.6))
    tf16 = tb16.text_frame
    tf16.word_wrap = True

    p16_tag = tf16.paragraphs[0]
    p16_tag.text = "【自住置產全案評估總結】"
    p16_tag.font.name = FONT_HEADING
    p16_tag.font.size = Pt(14)
    p16_tag.font.bold = True
    p16_tag.font.color.rgb = GOLD

    p16_main = tf16.add_paragraph()
    p16_main.text = "立足真實數據・堅持理性出價・守護家庭資產"
    p16_main.font.name = FONT_HEADING
    p16_main.font.size = Pt(34)
    p16_main.font.bold = True
    p16_main.font.color.rgb = WHITE
    p16_main.space_before = Pt(10)

    p16_sub = tf16.add_paragraph()
    p16_sub.text = "大慶國宅 10 樓特定待售戶出價評估任務圓滿達成"
    p16_sub.font.name = FONT_HEADING
    p16_sub.font.size = Pt(20)
    p16_sub.font.bold = True
    p16_sub.font.color.rgb = RGBColor(226, 232, 240)
    p16_sub.space_before = Pt(8)

    p16_box = tf16.add_paragraph()
    p16_box.text = (
        "■ 核心標的規格：大慶國宅 10 樓（次高樓 44.5 坪，主建物 25.365 坪，室內主附 27.77 坪）\n"
        "■ 市場開價現況：開價 998 萬（原開價 1080 萬，最新降價 82 萬，開價單價 22.43 萬/坪）\n"
        "■ 車位承租機制：無專有產權車位，現況承租管委會地下平面車位 1,500 元/月\n"
        "■ 最佳談判目標：合理目標價 860 ~ 890 萬元（19.33 ~ 20.00 萬/坪，屋主實質淨賺逾 550 萬）\n"
        "■ 嚴格停損邊界：920 ~ 930 萬元為絕對上限，若超逾即刻轉向歌德堡或長億南園！"
    )
    p16_box.font.name = FONT_BODY
    p16_box.font.size = Pt(12)
    p16_box.font.color.rgb = RGBColor(147, 197, 253)
    p16_box.space_before = Pt(24)

    p16_sign = tf16.add_paragraph()
    p16_sign.text = "自住置產決策小組・估價專案團隊  |  實價登錄 49 筆真實大數據嚴格覆核  |  2026 年 11 月 敬呈"
    p16_sign.font.name = FONT_HEADING
    p16_sign.font.size = Pt(12)
    p16_sign.font.bold = True
    p16_sign.font.color.rgb = GOLD
    p16_sign.space_before = Pt(36)

    # Save presentation
    prs.save(TARGET_PPTX)
    size_bytes = os.path.getsize(TARGET_PPTX)
    print(f"[SUCCESS] Presentation generated successfully: {TARGET_PPTX}")
    print(f"[INFO] Slide Count: {len(prs.slides)}, File Size: {size_bytes / 1024:.1f} KB")

if __name__ == "__main__":
    generate_daqing_presentation()

# -*- coding: utf-8 -*-
"""
scripts/generate_district_presentations_v2.py
---------------------------------------------
Objective Third-Party Real Estate Econometric Presentation Generator (V2).
Generates full 16:9 widescreen PPTX decks for BOTH:
- 台中市南區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx (10 intro/model/top10 + 54 communities + 1 framework = 65 slides)
- 台中市南屯區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx (10 intro/model/top10 + 73 communities + 1 framework = 84 slides)

Key Principles:
- 100% neutral, objective, scientific third-party analytical stance (zero buyer/seller bias).
- Zero renovation cost assumptions.
- Monthly point-by-point price trajectory (2012/08~2026/11 history + 2027/01~2029/12 36-month forecast).
- Top 10 indicator communities strictly ranked by total household scale (#01~#10).
- Dual-verified against 591 & Leju (115/11/21 MOI batch).
"""

import os
import sys
import json
import shutil
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

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRATCH_DIR = os.path.join(PROJECT_ROOT, "scratch")

NAVY = RGBColor(15, 23, 42)
PRIMARY_BLUE = RGBColor(27, 54, 93)
SLATE = RGBColor(51, 65, 85)
ACCENT_BLUE = RGBColor(2, 132, 199)
ACCENT_PURPLE = RGBColor(109, 40, 217)
LIGHT_BG = RGBColor(248, 250, 252)
WHITE = RGBColor(255, 255, 255)
BORDER_GRAY = RGBColor(203, 213, 225)
ROW_ALT = RGBColor(241, 245, 249)
AMBER = RGBColor(180, 83, 9)
RED = RGBColor(185, 28, 28)
GREEN = RGBColor(21, 128, 61)
TEXT_DARK = RGBColor(30, 41, 59)
TEXT_MUTED = RGBColor(100, 116, 139)


def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_header(slide, title_text, subtitle_text, page_num, total_pages):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.92))
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRIMARY_BLUE
    bar.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(0.45), Inches(0.10), Inches(11.2), Inches(0.45))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "Microsoft JhengHei"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = WHITE

    tb2 = slide.shapes.add_textbox(Inches(0.45), Inches(0.52), Inches(11.2), Inches(0.32))
    tf2 = tb2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = subtitle_text
    p2.font.name = "Microsoft JhengHei"
    p2.font.size = Pt(11)
    p2.font.color.rgb = RGBColor(186, 230, 253)

    tb_pg = slide.shapes.add_textbox(Inches(11.7), Inches(0.25), Inches(1.25), Inches(0.4))
    p_pg = tb_pg.text_frame.paragraphs[0]
    p_pg.alignment = PP_ALIGN.RIGHT
    p_pg.text = f"{page_num} / {total_pages}"
    p_pg.font.name = "Calibri"
    p_pg.font.size = Pt(12)
    p_pg.font.bold = True
    p_pg.font.color.rgb = RGBColor(224, 242, 254)


def add_card(slide, left, top, width, height, title, bullets, accent_color=ACCENT_BLUE):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = BORDER_GRAY
    card.line.width = Pt(1)

    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(0.10), Inches(height))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = accent_color
    stripe.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(left + 0.22), Inches(top + 0.14), Inches(width - 0.35), Inches(height - 0.24))
    tf = tb.text_frame
    tf.word_wrap = True

    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = "Microsoft JhengHei"
    p0.font.size = Pt(14.5)
    p0.font.bold = True
    p0.font.color.rgb = accent_color
    p0.space_after = Pt(6)

    for b in bullets:
        p = tf.add_paragraph()
        p.text = f"• {b}"
        p.font.name = "Microsoft JhengHei"
        p.font.size = Pt(11.5)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(4.5)


def build_district_pptx(master_json_path, chart_dir, macro_png_name, out_pptx_name):
    with open(master_json_path, "r", encoding="utf-8") as f:
        master = json.load(f)

    district = master["district"]
    comms = master["communities"]
    n_comms = len(comms)
    total_slides = 10 + n_comms + 1
    macro = master["macro"]
    fc3 = macro["forecast_3yr"]
    bt = master["backtest_validation"]
    hed = master["hedonic_model"]
    s_low, s_high = macro["support_band_2017_2020"]
    eq_2026 = macro["fundamental_eq_2026"]
    cur_est = macro["current_monthly_est_2026_11"]
    peak_m = max(macro["monthly_history"], key=lambda m: m["price"])

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- Slide 1: Cover ---
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s1, NAVY)
    band = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.28), Inches(7.5))
    band.fill.solid()
    band.fill.fore_color.rgb = RGBColor(56, 189, 248)
    band.line.fill.background()

    tb1 = s1.shapes.add_textbox(Inches(0.9), Inches(1.3), Inches(11.6), Inches(4.8))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p_tag = tf1.paragraphs[0]
    p_tag.text = "OBJECTIVE THIRD-PARTY ECONOMETRIC RESEARCH ｜ 客觀第三方不動產計量實證研究"
    p_tag.font.name = "Microsoft JhengHei"
    p_tag.font.size = Pt(13)
    p_tag.font.bold = True
    p_tag.font.color.rgb = RGBColor(56, 189, 248)
    p_tag.space_after = Pt(14)

    p_main = tf1.add_paragraph()
    p_main.text = f"台中市{district}【30年以上・100戶以上大型社區】\n逐月實價登錄價格序列與 2027～2029 年計量預測報告"
    p_main.font.name = "Microsoft JhengHei"
    p_main.font.size = Pt(31)
    p_main.font.bold = True
    p_main.font.color.rgb = WHITE
    p_main.space_after = Pt(18)

    p_sub = tf1.add_paragraph()
    special_note = "（含指定納入：歌德堡、大慶中山國宅）" if district == "南區" else ""
    p_sub.text = (
        f"• 研究母體：{district} {n_comms} 棟大型集合住宅社區{special_note}，合計 {master['total_households']:,} 戶（平均屋齡 {master['avg_age_years']} 年）\n"
        f"• 數據基礎：內政部實價登錄 115/11/21 同步批次（591 實價登錄 × 樂居 Leju 雙平台交叉查驗），共 {master['total_transactions']:,} 筆明細\n"
        f"• 方法論：面板固定效應 Hedonic 特徵校正 × RTS 卡爾曼狀態空間逐月序列（2012/08～2029/12 共 209 個月點）\n"
        f"• 研究立場：嚴守獨立第三方客觀科學基準，不偏袒買賣任一方，不納入主觀裝修費用假設"
    )
    p_sub.font.name = "Microsoft JhengHei"
    p_sub.font.size = Pt(14)
    p_sub.font.color.rgb = RGBColor(203, 213, 225)

    # --- Slide 2: Executive Summary ---
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s2, LIGHT_BG)
    add_header(s2, f"執行摘要：{district}大型老牌社區四大客觀實證發現", "整合 2012/08～2026/11 逐月狀態空間指數與 2027～2029 年（民國116～118年）預測結果", 2, total_slides)

    add_card(s2, 0.45, 1.18, 6.05, 2.82, "1. 歷史逐月價格週期與當前水位", [
        f"全區逐月特徵校正指數於 {peak_m['ym']} 達歷史高點 {peak_m['price']:.1f} 萬/坪，2026/11 最新月度中樞為 {cur_est:.1f} 萬/坪。",
        f"經 591 與樂居雙平台交叉查驗，最新揭露成交日達 {master['latest_transaction_date']}（對應內政部 115/11/21 批次）。",
        "剔除 1 樓店面、特殊關係人交易並完成樓層／車位特徵校正後，真實反映標準住宅價格軌跡。",
    ], ACCENT_BLUE)

    add_card(s2, 6.80, 1.18, 6.05, 2.82, "2. 交易量能與微觀在售庫存結構", [
        f"全區歷史年度高峰（{macro['peak_year']}年）成交 {macro['peak_annual_vol']} 件；近 12 個月滾動（LTM）成交 {macro['ltm_vol_2025_11_to_2026_10']} 件（{macro['ltm_vol_change_from_peak_pct']:+.1f}%）。",
        f"目前線上總在售掛牌共 {master['total_active_listings']:,} 戶（占總戶數 {macro['microstructure']['listing_rate_pct']:.2f}%），整體庫存去化期（MOI）為 {macro['microstructure']['moi_months']:.1f} 個月。",
        "相較於房市供需均衡基準（6～9 個月），當前多數大型老牌社區處於買方選擇充裕、去化節奏放緩階段。",
    ], AMBER)

    add_card(s2, 0.45, 4.25, 6.05, 2.85, "3. 2027～2029 年客觀第三方逐月預測中樞", [
        f"2027 年（116年）預估年均：{fc3['2027']['expected']:.1f} 萬/坪（90%區間 [{fc3['2027']['lower_90']:.1f}～{fc3['2027']['upper_90']:.1f}]，年變動 {fc3['2027']['yoy_change_pct']:+.1f}%）。",
        f"2028 年（117年）預估年均：{fc3['2028']['expected']:.1f} 萬/坪（90%區間 [{fc3['2028']['lower_90']:.1f}～{fc3['2028']['upper_90']:.1f}]，年變動 {fc3['2028']['yoy_change_pct']:+.1f}%）。",
        f"2029 年（118年）預估年均：{fc3['2029']['expected']:.1f} 萬/坪（三年累計變動 {macro['cum_change_2027_2029_pct']:+.1f}%），朝基本面均衡中樞 {eq_2026:.1f} 萬溫和收斂。",
    ], ACCENT_PURPLE)

    add_card(s2, 6.80, 4.25, 6.05, 2.85, "4. 樣本外歷史回測驗證與科學嚴謹度", [
        f"以 2012/08～2025/10 訓練模型，對 2025/11～2026/10 進行 12 個月樣本外滾動回測（共 {bt['evaluated_deals']} 筆純住交易）。",
        f"樣本外絕對百分比誤差中位數（Median APE）為 {bt['median_ape_pct']:.2f}%，均方根誤差 RMSE 為 {bt['rmse_wan']:.2f} 萬/坪。",
        "模型不預設多空立場，亦不計入主觀室內裝修費用，純粹由總體信用、微觀庫存與通膨折舊均衡推導。",
    ], GREEN)

    # --- Slide 3: Macro Monthly Chart ---
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s3, LIGHT_BG)
    add_header(s3, f"大盤全景：{district} {n_comms} 棟大型社區逐月價格點與 2027～2029 預測總圖", "2012/08～2029/12 共 209 個月連續價格點序列（藍點為歷史逐月值，紫點為未來 36 個月逐月預估值）", 3, total_slides)
    macro_img_path = os.path.join(chart_dir, macro_png_name)
    if os.path.exists(macro_img_path):
        s3.shapes.add_picture(macro_img_path, Inches(0.35), Inches(1.02), width=Inches(12.63))

    # --- Slide 4: Four Objective Fundamental Drivers Overview ---
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s4, LIGHT_BG)
    add_header(s4, "分析架構：影響 2027～2029 年大型老牌社區價格走勢之四大客觀機制", "採中立計量經濟視角，綜合考量下檔修正壓力與名目通膨／所得支撐力量（不含裝修費用假設）", 4, total_slides)

    add_card(s4, 0.45, 1.18, 6.05, 2.82, "維度一：微觀供需與在售庫存去化機制", [
        f"全區 {n_comms} 棟社區現有 {master['total_active_listings']:,} 戶在售，庫存去化期 {macro['microstructure']['moi_months']:.1f} 個月。",
        "周邊重劃區（十三期、單元重劃區等）2026～2028 年進入交屋期，提供替代性新增供給。",
        "個別社區待售率差異顯著：低庫存社區具價格韌性，高掛售量社區則面臨同棟比價調節。",
    ], ACCENT_BLUE)

    add_card(s4, 6.80, 1.18, 6.05, 2.82, "維度二：建物物理折舊與長週期生命週期", [
        f"研究母體平均屋齡已達 {master['avg_age_years']} 年，RC 建物每年存在約 0.6%～0.8% 之自然物理折舊。",
        "公共設備（電梯、給排水、消防與外牆）進入中後期維護週期，管委會公共基金充裕度影響社區保值性。",
        "隨持有時間推進 10～20 年後屋齡將屆 45～50 年，大型社區產權戶數多、都更整合門檻高，長期流動性遞減。",
    ], AMBER)

    add_card(s4, 0.45, 4.25, 6.05, 2.85, "維度三：總體貨幣環境與高屋齡授信規範", [
        "央行第七波選擇性信用管制與銀行法 72-2 條不動產放款水位約束，抑制高槓桿投資需求，市場回歸自住基本面。",
        "行庫對高屋齡建物普遍適用「屋齡＋貸款年限 ≤ 50～55 年」授信準則，貸款期限縮短提高每月本息攤還門檻。",
        "2027～2029 年若信用管制隨市場降溫逐步常態化，剛性買盤成交量有望於 2028～2029 年溫和回穩。",
    ], RED)

    add_card(s4, 6.80, 4.25, 6.05, 2.85, "維度四：名目通膨、薪資成長與基本面均衡支撐", [
        f"2017～2020 年全區基期四分位區間為 {s_low:.1f}～{s_high:.1f} 萬/坪；惟過去七年台灣 CPI 與名目薪資累計顯著成長。",
        f"扣除建物屋齡折舊後，2026 年客觀基本面均衡價值中樞約位於 {eq_2026:.1f} 萬/坪，構成中長期下檔實質支撐。",
        "因此價格並非無底限下跌，而是由 2023～2024 年超漲區間向通膨調整後均衡值漸進收斂（軟著陸）。",
    ], GREEN)

    # --- Slide 5: Deep Dive 1 - Supply/Demand & Microstructure ---
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s5, LIGHT_BG)
    add_header(s5, "客觀機制剖析（一）：微觀供需結構與社區庫存去化月數（MOI）", "以即時在售掛牌量與近 12 個月實際去化速率量化各社區真實供需鬆緊度", 5, total_slides)
    top_sale_comms = sorted(comms, key=lambda x: -int(x.get("sale_num") or 0))[:5]
    top_sale_str = "、".join([f"{c['name']}({c['sale_num']}戶/MOI {c['microstructure']['moi_months']:.0f}月)" for c in top_sale_comms])
    add_card(s5, 0.45, 1.20, 12.43, 2.75, "庫存去化月數（Months of Inventory, MOI）實證意義", [
        "定義：MOI = 當前社區線上在售戶數 ÷（近 12 個月滾動純住宅實際成交件數 ÷ 12）。不動產經濟學一般以 MOI 6～9 個月視為供需均衡區間。",
        f"{district}實證現況：全區 {n_comms} 棟社區合計在售 {master['total_active_listings']:,} 戶（其中明確標註降價物件 {master['total_discounted_listings']} 戶），全區平均 MOI 為 {macro['microstructure']['moi_months']:.1f} 個月。",
        f"目前掛售規模前五名社區：{top_sale_str}。",
    ], ACCENT_BLUE)
    add_card(s5, 0.45, 4.20, 12.43, 2.85, "對 2027～2029 年價格軌跡之雙向對稱影響", [
        "高庫存社區（MOI > 24 個月）：同期間多戶掛牌形成內部比價效應，成交價傾向貼近近期實登下緣微幅修正以換取成交量。",
        "低庫存社區（MOI < 8 個月或零掛售）：供給稀缺且住戶惜售穩定，價格受大盤量縮影響較小，具備抗跌或隨通膨緩升之條件。",
        "跨區新成屋替代效應：鄰近重劃區新建案雖單價較高，但以低自備或長年期房貸吸引年輕客群，分流部分換屋預算。",
    ], SLATE)

    # --- Slide 6: Deep Dive 2 - Building Lifecycle & Depreciation (NO Renovation Cost!) ---
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s6, LIGHT_BG)
    add_header(s6, "客觀機制剖析（二）：高齡集合住宅物理折舊與長週期資產屬性", "純就建物結構折舊、公共設施維護治理與高齡金融流動性進行科學評估", 6, total_slides)
    add_card(s6, 0.45, 1.20, 12.43, 2.75, "30 年以上 RC 大樓物理折舊與公共維護治理分化", [
        f"依據不動產估價技術規則與實證特徵價格文獻，30 年以上鋼筋混凝土（RC）住宅大樓每年自然物理折舊率約為建物價值之 0.6%～0.8%。",
        "隨屋齡跨過 30～35 年，電梯主機控制系統、公共給排水幹管、消防機電與外牆磁磚進入例行性大修週期。",
        "社區管理委員會運作健全、公共修繕基金充裕且定期維護之社區，能有效抵銷屋齡折舊並維持長期居住品質。",
    ], AMBER)
    add_card(s6, 0.45, 4.20, 12.43, 2.85, "長週期（10～20年後邁向 45～50 年屋齡）之流動性與整合門檻", [
        "現今屋齡 30～35 年之社區，在持有 10～20 年後將步入 45～50 年高齡建築階段，次級市場接手買盤結構將隨貸款年限縮短而收斂。",
        "百戶至數百戶之高容積大型集合住宅，土地持分較小且區分所有權人眾多，都市更新或危老重建之整合時程與難度客觀偏高。",
        "因此計量模型將「高屋齡折舊率微幅遞增因子」納入長期均衡方程式，客觀反映不同屋齡階段之時間價值變化。",
    ], ACCENT_PURPLE)

    # --- Slide 7: Deep Dive 3 - Credit, Interest Rate & Macro Affordability ---
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s7, LIGHT_BG)
    add_header(s7, "客觀機制剖析（三）：總體信用政策、房貸授信規範與實質購買力", "央行信用管制、銀行法 72-2 條水位與名目薪資通膨對房價之雙向均衡作用", 7, total_slides)
    add_card(s7, 0.45, 1.20, 6.05, 5.85, "緊縮端：選擇性信用管制與老屋核貸規範", [
        "央行第七波信用管制取消名下有房者第 1 戶寬限期，第 2 戶最高貸款成數限縮至 5 成，短期置產與槓桿換屋需求明顯降溫。",
        "銀行法 72-2 條不動產放款比率約束使行庫審核趨嚴，一般房貸利率處於約 2.4%～3.0% 區間，買方對總價敏感度提升。",
        "多數行庫執行「屋齡＋貸款年限 ≤ 50～55 年」與「借款人年齡＋年限 ≤ 75 年」準則，32～35 年建物常見核貸年限為 20～25 年，相較新屋 30～40 年期每月本息攤還額較高。",
    ], RED)
    add_card(s7, 6.80, 1.20, 6.05, 5.85, "支撐端：名目所得成長、通膨與剛性自住底盤", [
        f"{district}具備成熟商圈、學區、醫療與捷運／鐵路交通機能，總價 600～1,200 萬元區間仍為台中剛性自住家庭主力負擔帶。",
        "2019～2026 年間台灣名目 GDP、基本薪資與營建重置成本累積顯著漲幅，墊高整體名目價格底線。",
        f"經模型推算，{district}大型老牌社區 2026 年通膨與折舊調整後之合理基本面均衡中樞約為 {eq_2026:.1f} 萬/坪（高於 2017–2020 名目基期 {s_low:.1f}～{s_high:.1f} 萬），提供中長期穩健支撐。",
    ], GREEN)

    # --- Slide 8: Dual-Platform Data Verification (591 x Leju) ---
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s8, LIGHT_BG)
    add_header(s8, "數據嚴謹度：591 實價登錄 × 樂居 Leju 雙平台交叉查驗與五重清洗", "消除單一平台車位估算偏誤、國宅分管車位誤拆與特殊關係人離群交易干擾", 8, total_slides)
    add_card(s8, 0.45, 1.20, 12.43, 2.80, "跨平台雙重確認（591 × 樂居 Leju × 內政部 115/11/21 批次）三大校正成果", [
        "1. 修正未拆車位估算偏誤：針對內政部原始未拆分車位案件，排除單一平台過度扣除車位坪數導致單價虛高之異常值，與樂居車位拆分口徑一致。",
        "2. 國宅分管協議車位精準處理：針對大慶國宅（大慶中山）、博愛國宅等公設抽籤車位社區，比照樂居「*」號規範採用原始登記單價，不重複扣除車位。",
        "3. 嚴格剔除特殊關係人交易：交叉核對備註欄（親友、二親等、債權抵償、特殊交易標記，如德昌中華 115 年中數筆親友特殊件），確保均價與指數不被扭曲。",
    ], ACCENT_BLUE)
    add_card(s8, 0.45, 4.22, 12.43, 2.82, f"面板固定效應 Hedonic 特徵價格迴歸參數（{district}實證估計結果，樣本數 N={hed['sample_size_within_fe']:,}）", [
        f"• 相對樓層高度溢價係數（beta_rel_floor = {hed['beta_rel_floor']:+.4f}）：控制同社區高低樓層視野採光價差，將每筆交易標準化至中樓層（F/F_max = 0.5）。",
        f"• 頂樓效應（beta_top = {hed['beta_top_floor']:+.4f}）與四樓效應（beta_4f = {hed['beta_4f']:+.4f}）：客觀分離樓層屬性折溢價，避免單月成交特定樓層造成假性漲跌。",
        f"• 面積規模彈性（beta_log_area = {hed['beta_log_area']:+.4f}）：控制同社區大小坪數房型之單價結構差異，建構同質可比（Constant-Quality）月度價格序列。",
    ], ACCENT_PURPLE)

    # --- Slide 9: Econometric Model Specification & Backtest Validation ---
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s9, LIGHT_BG)
    add_header(s9, "預測模型方法論：RTS 卡爾曼狀態空間逐月序列與 2027～2029 預測方程", "為什麼能算出「每一個月份的價格點」？狀態空間濾波與樣本外回測完整解析", 9, total_slides)
    add_card(s9, 0.45, 1.18, 6.10, 5.90, "歷史逐月價格點推估：Rauch-Tung-Striebel 卡爾曼平滑器", [
        "問題背景：單一社區並非每個月皆有成交；若僅看年度均價會喪失月度轉折細節，若直接連線則受空窗月影響。",
        "狀態轉移方程：ln(P_{i,t}) = ln(P_{i,t-1}) + d_Zone_{z,t} + eta_{i,t}，結合所屬生活圈月度動能與社區專屬狀態。",
        "觀測更新方程：有成交月份依樣本數精度加權錨定實登 Hedonic 均價；無成交月份由前後實登與生活圈指數進行貝式最優狀態內插。",
        "成果：自 2012/08 至 2026/11 連續 172 個月，每一個月皆具備統計最優推估之月度價格點。",
    ], ACCENT_BLUE)
    add_card(s9, 6.75, 1.18, 6.13, 5.90, "2027/01～2029/12（36個月）客觀預測與樣本外回測", [
        "階層經驗貝式動能：融合社區近 12 個月趨勢、生活圈趨勢與行政區大盤趨勢（不設人為正負限幅偏見）。",
        "微觀庫存調節項：依社區 MOI 相對於 8 個月均衡水準之偏離度對稱調節（低庫存支撐、高庫存溫和修正）。",
        "誤差修正均衡項（ECM）：價格偏離「通膨與屋齡折舊校正後均衡值 P*」時，每月按誤差修正速率漸進收斂。",
        f"樣本外滾動回測（訓練期 {bt['train_period']} → 測試期 {bt['test_period']}）：實測 {bt['evaluated_deals']} 筆交易，中位誤差 Median APE = {bt['median_ape_pct']:.2f}%，90% 信賴區間覆蓋率達 {bt['coverage_90_pct']:.1f}%。",
    ], GREEN)

    # --- Slide 10: Top 10 Communities by Household Scale Comparison Table ---
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s10, LIGHT_BG)
    add_header(s10, f"十大指標社區總覽：{district}戶數規模前 10 大社區實證與 2027～2029 預估表", "嚴格依社區總戶數規模排序（#01～#10），整合在售庫存、歷史峰值、最新成交與未來三年年均預估", 10, total_slides)

    top10 = comms[:10]
    cols_headers = [
        "排名", "社區名稱", "總戶數/屋齡", "在售(MOI)", "歷史最高",
        "最新成交(年月)", "高點變動", "26/11月度中樞", "2027預估", "2028預估", "2029預估", "三年累計"
    ]
    tbl_shape = s10.shapes.add_table(11, len(cols_headers), Inches(0.32), Inches(1.08), Inches(12.69), Inches(6.05))
    tbl = tbl_shape.table
    col_widths = [0.62, 1.75, 1.12, 1.08, 0.95, 1.38, 0.90, 1.12, 0.95, 0.95, 0.95, 0.92]
    for ci, w in enumerate(col_widths):
        tbl.columns[ci].width = Inches(w)

    for ci, h_txt in enumerate(cols_headers):
        cell = tbl.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = PRIMARY_BLUE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.text = h_txt
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Microsoft JhengHei"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = WHITE

    for ri, c in enumerate(top10, 1):
        bg_c = WHITE if ri % 2 == 1 else ROW_ALT
        peak_p = c["historical_peak_price"]
        lat_p = c["latest_deal_price"]
        lat_ym = c["latest_deal_date"][:7]
        drop_pk = ((lat_p - peak_p) / peak_p * 100.0) if peak_p > 0 else 0.0
        moi_v = c["microstructure"]["moi_months"]
        sale_str = f"{c['sale_num']}戶({moi_v:.0f}月)" if c["sale_num"] > 0 else "0戶"
        f3 = c["forecast_3yr"]
        row_vals = [
            f"#{c['rank']:02d}",
            c["name"],
            f"{c['households']}戶/{c['age']}年",
            sale_str,
            f"{peak_p:.1f}萬",
            f"{lat_p:.1f}萬({lat_ym})",
            f"{drop_pk:+.1f}%",
            f"{c['current_monthly_est_2026_11']:.1f}萬",
            f"{f3['2027']['expected']:.1f}萬",
            f"{f3['2028']['expected']:.1f}萬",
            f"{f3['2029']['expected']:.1f}萬",
            f"{c['cum_change_2027_2029_pct']:+.1f}%",
        ]
        for ci, val in enumerate(row_vals):
            cell = tbl.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_c
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.alignment = PP_ALIGN.LEFT if ci == 1 else PP_ALIGN.CENTER
            p.font.name = "Microsoft JhengHei"
            p.font.size = Pt(9.8)
            p.font.bold = ci in (0, 1, 7, 10, 11)
            if ci in (6, 11):
                p.font.color.rgb = RED if "-" in val else GREEN
            elif ci == 7:
                p.font.color.rgb = ACCENT_BLUE
            else:
                p.font.color.rgb = TEXT_DARK

    # --- Slides 11 .. 10+N: Individual Community Monthly Analysis Slides ---
    for idx, c in enumerate(comms, 1):
        slide_no = 10 + idx
        sc = prs.slides.add_slide(blank_layout)
        set_slide_bg(sc, LIGHT_BG)
        top10_tag = "【十大規模指標】" if idx <= 10 else ""
        add_header(
            sc,
            f"社區逐月分析 #{c['rank']:02d}：{c['name']} {top10_tag}",
            f"{c['life_zone']} ｜ 門牌：{c['address']} ｜ 總戶數：{c['households']:,} 戶 ｜ 屋齡：{c['age']} 年（完工：{c['build_date']}）",
            slide_no,
            total_slides,
        )

        c_img = os.path.join(chart_dir, c["chart_filename"])
        if os.path.exists(c_img):
            sc.shapes.add_picture(c_img, Inches(0.30), Inches(1.04), width=Inches(8.75))

        f3 = c["forecast_3yr"]
        peak_p = c["historical_peak_price"]
        lat_p = c["latest_deal_price"]
        drop_pk = ((lat_p - peak_p) / peak_p * 100.0) if peak_p > 0 else 0.0
        moi_v = c["microstructure"]["moi_months"]
        sale_desc = f"{c['sale_num']} 戶（待售率 {c['microstructure']['listing_rate_pct']:.1f}%，MOI {moi_v:.1f} 個月）" if c["sale_num"] > 0 else "0 戶在售（籌碼穩定）"

        add_card(sc, 9.18, 1.04, 3.85, 2.95, "歷史實登與微觀供需摘要", [
            f"實登筆數：總計 {c['total_transactions']} 筆（純住 {c['clean_residential_count']} 筆）",
            f"歷史峰值：{peak_p:.1f} 萬/坪（{c['historical_peak_date'][:7]}，{c['historical_peak_floor']}）",
            f"最新純住：{lat_p:.1f} 萬/坪（{c['latest_deal_date'][:7]}，較高點 {drop_pk:+.1f}%）",
            f"2026/11 月度校正中樞：{c['current_monthly_est_2026_11']:.1f} 萬/坪",
            f"當前在售：{sale_desc}",
        ], ACCENT_BLUE)

        add_card(sc, 9.18, 4.12, 3.85, 3.05, "2027～2029 客觀計量預估", [
            f"基本面均衡錨點：{c['fundamental_eq_2026']:.1f} 萬/坪（17-20基期 {c['support_band_2017_2020'][0]:.1f}~{c['support_band_2017_2020'][1]:.1f} 萬）",
            f"2027年(116)預估：{f3['2027']['expected']:.1f} 萬 [{f3['2027']['lower_90']:.1f}~{f3['2027']['upper_90']:.1f}] ({f3['2027']['yoy_change_pct']:+.1f}%)",
            f"2028年(117)預估：{f3['2028']['expected']:.1f} 萬 [{f3['2028']['lower_90']:.1f}~{f3['2028']['upper_90']:.1f}] ({f3['2028']['yoy_change_pct']:+.1f}%)",
            f"2029年(118)預估：{f3['2029']['expected']:.1f} 萬 [{f3['2029']['lower_90']:.1f}~{f3['2029']['upper_90']:.1f}]",
            f"2027～2029 三年累計預估變動：{c['cum_change_2027_2029_pct']:+.1f}%",
        ], ACCENT_PURPLE)

    # --- Final Slide: Symmetric Decision Assessment Framework ---
    sf = prs.slides.add_slide(blank_layout)
    set_slide_bg(sf, LIGHT_BG)
    add_header(sf, "客觀第三方市場評估框架：資產持有方與購屋評估方之科學參考", "不預設買賣立場，依據社區微觀供需（MOI）、月度特徵價格中樞與金融條件提供對稱決策依據", total_slides, total_slides)

    add_card(sf, 0.45, 1.20, 6.05, 5.85, "對資產持有方（屋主）之客觀評估維度", [
        "1. 檢視同社區即時競售水位（MOI 指標）：若社區 MOI 低於 8 個月且掛售稀少，具備議價空間韌性；若 MOI 高於 24 個月且多戶同時掛售，定價宜貼近最新月度校正中樞以縮短去化週期。",
        "2. 區分「樓層／車位結構差異」與「真實行情變動」：參考本報告經 Hedonic 校正後之標準樓層月度價格，避免受單筆低樓層或特殊交易誤導。",
        "3. 長期資產配置與屋齡週期評估：綜合考量租金收益率、資金機會成本，以及建物邁向 40～50 年後次級市場買方貸款年限縮短對長期流動性之影響。",
    ], ACCENT_BLUE)

    add_card(sf, 6.80, 1.20, 6.05, 5.85, "對購屋評估方（買方）之客觀評估維度", [
        "1. 以「月度特徵校正中樞」與「車位拆算淨單價」作為客觀價值基準：比對標的樓層高度、車位型態（坡平／坡機／國宅分管車位）後換算合理單價區間。",
        "2. 事前執行銀行授信與現金流壓力測試：針對 30 年以上標的，先向往來行庫確認實際可貸成數、核貸年限（20～25 年）與鑑價區間，並於契約中妥善載明貸款條件保障條款。",
        "3. 重視社區公共治理品質：實地查核管委會財務報表、公共修繕基金結餘、電梯與消防機電維護紀錄，優選治理穩健之保值型社區。",
    ], GREEN)

    out_dir_pptx = os.path.join(chart_dir, out_pptx_name)
    root_pptx = os.path.join(PROJECT_ROOT, out_pptx_name)
    prs.save(out_dir_pptx)
    shutil.copy2(out_dir_pptx, root_pptx)
    print(f"✅ 已產出【{district}】客觀第三方寬螢幕簡報 ({total_slides} 頁): {out_dir_pptx} (並同步至根目錄)")


def main():
    build_district_pptx(
        os.path.join(SCRATCH_DIR, "south_30yr_master.json"),
        os.path.join(PROJECT_ROOT, "南區30年以上大型社區走勢圖"),
        "00_台中市南區30年以上大型社區_大盤匯總走勢圖.png",
        "台中市南區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx",
    )
    build_district_pptx(
        os.path.join(SCRATCH_DIR, "nantun_30yr_master.json"),
        os.path.join(PROJECT_ROOT, "南屯區30年以上大型社區走勢圖"),
        "00_台中市南屯區30年以上大型社區_大盤匯總走勢圖.png",
        "台中市南屯區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx",
    )


if __name__ == "__main__":
    main()

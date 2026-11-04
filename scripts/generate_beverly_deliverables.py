#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scripts/generate_beverly_deliverables.py
Unified generator for Taichung South District "Beverly Hills 13F" (比佛利山莊13樓):
1. Markdown Analysis Report (.md)
2. Professional Word Business Document (.docx)
3. 16:9 Widescreen Executive Presentation (.pptx)

All deliverables are output directly into:
`比佛利山莊13樓/`
"""

import os
import sys
import json

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import pptx
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MASTER_JSON_PATH = os.path.join(PROJECT_ROOT, "scratch", "beverly_valuation_master.json")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "比佛利山莊13樓")
CHART_BEVERLY = os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "17_比佛利山莊_成交行情走勢圖.png")
CHART_OVERALL = os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "00_台中市南區25年以上大型社區_大盤匯總走勢圖.png")

MD_OUTPUT_PATH = os.path.join(OUTPUT_DIR, "比佛利山莊13樓_特定待售戶別出價分析與社區整體行情報告.md")
DOCX_OUTPUT_PATH = os.path.join(OUTPUT_DIR, "比佛利山莊13樓_特定待售戶別出價分析與社區整體行情報告.docx")
PPTX_OUTPUT_PATH = os.path.join(OUTPUT_DIR, "比佛利山莊13樓_特定待售戶別出價分析與社區整體行情報告.pptx")

# Palette constants for DOCX
NAVY_PRIMARY = (27, 54, 93)       # #1B365D
NAVY_SECONDARY = (44, 94, 138)    # #2C5E8A
GOLD_ACCENT = (217, 140, 20)      # #D98C14
GREEN_ACCENT = (39, 174, 96)      # #27AE60
RED_ALERT = (192, 57, 43)         # #C0392B
TEXT_DARK = (30, 41, 59)          # #1E293B
TEXT_MUTED = (100, 116, 139)      # #64748B
BG_LIGHT = "F8FAFC"
BORDER_LIGHT = "E2E8F0"
HEADER_BG = "1B365D"

# Palette constants for PPTX
PPTX_NAVY = PptxRGBColor(27, 54, 93)
PPTX_DARK_BLUE = PptxRGBColor(15, 32, 67)
PPTX_SLATE = PptxRGBColor(44, 94, 138)
PPTX_LIGHT_BG = PptxRGBColor(248, 250, 252)
PPTX_WHITE = PptxRGBColor(255, 255, 255)
PPTX_CARD_BG = PptxRGBColor(255, 255, 255)
PPTX_BORDER_GRAY = PptxRGBColor(226, 232, 240)
PPTX_GOLD = PptxRGBColor(217, 140, 20)
PPTX_RED = PptxRGBColor(192, 57, 43)
PPTX_GREEN = PptxRGBColor(39, 174, 96)
PPTX_TEXT_DARK = PptxRGBColor(30, 41, 59)
PPTX_TEXT_MUTED = PptxRGBColor(100, 116, 139)


def load_master_data():
    with open(MASTER_JSON_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


# ==============================================================================
# 1. MARKDOWN GENERATOR
# ==============================================================================
def generate_markdown(data):
    tp = data["target_property"]
    vm = data["valuation_model"]
    ch = data["community_history"]
    comps = data["comparables"]
    nt = data["negotiation_tactics"]
    prev = tp["previous_sale"]
    sens = vm["parking_sensitivity"]
    adj = vm["adjustment_matrix"]

    lines = []
    lines.append(f"# 比佛利山莊社區整體市場行情分析與 13 樓特定待售戶別出價策略專案報告\n")
    lines.append(f"**評估標的**：台中市南區比佛利山莊 13 樓視野大三房（{tp['address']} 13F）  ")
    lines.append(f"**委託評估單位**：自住置產專案決策小組  ")
    lines.append(f"**報告產出日期**：2026 年 11 月  ")
    lines.append(f"**資料庫引用**：內政部不動產交易實價查詢服務網、591 房屋交易網、樂居（LEJU）實價登錄資料庫、台中市政府地政局建物登記謄本、永義房屋與中信房屋委託調查表  \n")
    lines.append(f"---\n")

    # 執行摘要
    lines.append(f"## 報告執行摘要 (Executive Summary)\n")
    lines.append(f"本報告針對台中市南區忠明南路生活圈「**比佛利山莊**」社區特定待售物件（**13 樓次頂樓戶別，權狀總面積 {tp['total_area']} 坪含坡道機械車位 {tp['parking_area']} 坪，純建坪 {tp['house_area']} 坪，室內實坪達 {tp['indoor_area']} 坪，格局為 3 房 2 廳 2 衛 1 陽台，目前開價 {tp['asking_price']:.0f} 萬元，毛單價 {tp['asking_gross_unit_price']} 萬/坪**）進行全面性產權查核、實價登錄清洗、同生活圈競品比價、標準車位分離拆算與談判出價策略分析。\n")
    lines.append(f"### 核心關鍵發現與決策要旨：\n")
    lines.append(f"1. **鎖定屋主取得成本重大鐵證（加價 +128 萬元，溢價 +14.7%）**：  ")
    lines.append(f"   實價登錄歷史資料庫確切證實，本特定標的戶（忠明南路730巷42號 13樓）前次交易日期為 **{prev['date']}（民國 {prev['roc_date']}）**，以 **{prev['price']:.0f} 萬元（含機械車位，登載單價 {prev['unit_price']} 萬/坪）購入成交**。現任屋主持有甫滿 2 年，目前開價 {tp['asking_price']:.0f} 萬元，**帳面墊高加價達 +{prev['markup_amount']:.0f} 萬元（+{prev['markup_percent']:.1f}%）**。此紀錄為買方後續定錨與議價之最強鐵證，確立買方無需追高之心理防線。\n")
    lines.append(f"2. **落實標準車位分離拆算，還原真實房屋開價（30.45 萬/坪）**：  ")
    lines.append(f"   依台中市南區 30 年以上老牌大樓車位市場公允行情，坡道機械車位市場合理行情為 **{vm['parking_benchmark']:.0f} 萬元**（敏感度區間 60~80 萬元）。依內政部標準車位拆算原則扣除車位價格與坪數後，**待售戶建物淨開價單價高達 {vm['asking_net_unit_price']:.2f} 萬/坪**；相較於 2024 年屋主買進時之拆算淨單價 **{vm['previous_sale_net_unit_price']:.2f} 萬/坪**，開價淨單價高出 **+{vm['asking_net_premium_unit']:.2f} 萬/坪（+{vm['asking_net_premium_percent']:.1f}%）**，開價已完全頂在歷史多頭行情之最上緣。\n")
    lines.append(f"3. **大盤降溫與社區大三房行情實質回修**：  ")
    lines.append(f"   比佛利山莊歷史成交 158 筆經 5 層嚴格清洗後取得 124 筆純住宅交易。雖然全社區 2024-2025 年均價受小坪數套房推升至 27~29 萬/坪，但 **2026 年最新大三房交易（3 樓含平面車位）已實質回落至 20.40 萬/坪（總價 800 萬元）**。南區老牌指標社區如歌德堡（跌至 26.65 萬/坪）、健康世界（跌至 20.12 萬/坪）、長億南園（跌至 22.60 萬/坪）普遍回檔修正 -10%~-15%，買方享有充裕的市場比價觀望優勢。\n")
    lines.append(f"4. **一般委託多房仲聯賣與三段式出價防線**：  ")
    lines.append(f"   本案為中信房屋與永義房屋等多家房仲聯賣之一般委託案，非專任約，可運用跨體系比價制衡。依據比價調整矩陣推導三段式出價策略：  \n")
    lines.append(f"   - **保守底標價（大膽試探防線）**：**{vm['conservative']['total']:.0f} 萬元**（建物淨單價 **{vm['conservative']['net_unit_price']:.2f} 萬/坪**，約開價 **{vm['conservative']['discount_str']}**）。  \n")
    lines.append(f"   - **合理目標價（最佳成交甜蜜點）**：**{vm['target']['total']:.0f} 萬元**（建物淨單價 **{vm['target']['net_unit_price']:.2f} 萬/坪**，約開價 **{vm['target']['discount_str']}**，精準錨定屋主買進成本）。  \n")
    lines.append(f"   - **上限天花板價（談判停損警戒線）**：**{vm['ceiling']['total']:.0f} 萬元**（建物淨單價 **{vm['ceiling']['net_unit_price']:.2f} 萬/坪**，約開價 **{vm['ceiling']['discount_str']}**，自住極限不可突破）。  \n")
    lines.append(f"---\n")

    # 第一章
    lines.append(f"## 第一章：委託評估標的物件產調規格與現況查核 (物件現況)\n")
    lines.append(f"### 1.1 待售標的物件產權與建築物理規格\n")
    lines.append(f"本專案評估之標的戶座落於台中市南區忠明南路730巷42號，經由 591 售屋網快照檔（代碼 S21108759）與永義房屋謄本登載資料進行交叉比對，產權與規格權威確認清冊如下：\n")
    lines.append(f"| 物件規格維度項目 | 官方刊登與謄本驗證數值 | 規格特性與專業評估意涵 |")
    lines.append(f"| :--- | :--- | :--- |")
    lines.append(f"| **社區名稱** | **{tp['community_name']}** (社區代碼: {tp['community_id']}) | 忠明南路生活圈成熟住宅社區，總戶數 {tp['total_households']} 戶。 |")
    lines.append(f"| **建物門牌地址** | **{tp['full_address']}** | 臨忠明南路730巷靜巷，出入單純。 |")
    lines.append(f"| **所在樓層 / 總樓層** | **{tp['floor']}** | **黃金次頂樓**。採光視野極佳，無頂樓日曬與漏水風險。 |")
    lines.append(f"| **總開價 (Asking Price)** | **{tp['asking_price']:.0f} 萬元** (含車位價格) | 房仲標註誠意出售，具議價空間。 |")
    lines.append(f"| **刊登毛單價** | **{tp['asking_gross_unit_price']:.2f} 萬/坪** | 開價總價除以含車位總建坪 33.75 坪（未拆車位）。 |")
    lines.append(f"| **建物權狀總坪數** | **{tp['total_area']:.2f} 坪** (含車位) | 包含純建物登記 30.48 坪 + 車位約定持分 3.27 坪。 |")
    lines.append(f"| **主建物登記面積** | **{tp['main_area']:.2f} 坪** | 室內淨起居空間充足，標準大三房規格。 |")
    lines.append(f"| **附屬建物面積** | **{tp['attached_area']:.2f} 坪** (全為陽台) | 前後工作陽台，通風採光無虛坪雨遮。 |")
    lines.append(f"| **室內可用實坪 (主+附)**| **{tp['indoor_area']:.2f} 坪** (得房率 **{tp['indoor_ratio']:.2f}%**) | 實質室內可用率高達 93.2%，遠勝現代高公設新建案。 |")
    lines.append(f"| **大樓共用部分 (公設)**| **{tp['public_area']:.2f} 坪** (公設比 **{tp['public_ratio']:.2f}%**) | **超低公設比僅 6.76%**（591 四捨五入標 7%），虛坪極少。 |")
    lines.append(f"| **基地土地持分面積** | **{tp['land_area']:.3f} 坪** | 早期大樓土地持分紮實，自住具備長期土地保值底氣。 |")
    lines.append(f"| **車位產權與編號** | **{tp['parking_type']}** | 地下一層編號 B1-89，進出方式為汽車坡道駛入停入機械車位。 |")
    lines.append(f"| **室內格局規劃** | **{tp['layout']}** | **邊間三面採光，罕見雙衛浴皆開窗**，通風極佳。 |")
    lines.append(f"| **建物座向朝向** | **{tp['orientation']}** | 客廳朝東南，冬暖夏涼避開北風吹襲。 |")
    lines.append(f"| **建築完成日 / 屋齡** | **{tp['build_date']} / 約 {tp['building_age']:.1f} 年** | 民國 81 年 11 月完工，屋齡已跨入 34 年老宅門檻。 |")
    lines.append(f"| **每月管理費** | **{tp['management_fee']} 元/月** | 房屋管理費 {tp['house_management_fee']} 元 + 機械車位清潔費 {tp['parking_maintenance_fee']} 元。 |")
    lines.append(f"| **委託銷售型態** | **{tp['agency_status']}** | 至少兩家主要仲介品牌同時代理，無獨家專任約束。 |")
    lines.append(f"| **歷史購入成本鐵證** | **2024-01-24 以 870 萬元成交** | 實登單價 25.8 萬/坪，持有 2 年帳面加價 +128 萬 (+14.7%)。 |\n")

    lines.append(f"### 1.2 591 刊登內容與謄本關鍵矛盾查核剖析\n")
    lines.append(f"經比對 591 廣告與謄本紀錄，確認以下關鍵細節差異：\n")
    lines.append(f"1. **權狀坪數與車位面積結構**：591 標示權狀坪數為 33.75 坪，謄本純建物登記為 30.48 坪，差額 3.27 坪為房仲手動列入之約定專用機械車位持分面積。早期建築車位常登載於公設內無獨立建號，本案純建物淨建坪為 **30.48 坪**。\n")
    lines.append(f"2. **單價壓低之視覺盲點**：591 標示毛單價 29.57 萬/坪係將總價 998 萬直接除以灌入車位後的 33.75 坪。若未扣除車位直接除以建物 30.48 坪，單價高達 32.74 萬/坪；依標準車位扣除 70 萬後，真實建物淨單價為 **30.45 萬/坪**。\n")
    lines.append(f"3. **屋齡計算誤差**：591 刊載 32 年/33 年為未及時更新，官方建造執照完工日為 1992 年 11 月 20 日，至 2026 年底實質屋齡已屆滿 **34.0 年**，跨入銀行針對 30 年以上老屋之嚴審限期門檻。\n")
    lines.append(f"---\n")

    # 第二章
    lines.append(f"## 第二章：比佛利山莊社區歷史實價登錄全景與行情走勢 (社區歷程)\n")
    lines.append(f"### 2.1 歷史實價登錄全景統計與資料清洗篩選\n")
    lines.append(f"自民國 101 年實價登錄實施以來，比佛利山莊累計登錄 158 筆交易。經由專案標準 5 層過濾清洗：\n")
    lines.append(f"- **排除 1 樓與 1,2 樓店面交易**：共 **16 筆**（忠明南路 730 巷臨路店面單價高且含商用效益）。\n")
    lines.append(f"- **排除特殊關係與親友交易**：共 **20 筆**（親友贈與、持分買賣等非正常市場交易）。\n")
    lines.append(f"- **排除重疊交易**：2 筆（1,2 樓兼具親友特殊交易）。\n")
    lines.append(f"- **清洗後純住宅有效成交**：**124 筆**（全期住宅均價 {ch['overall_average_unit_price']:.2f} 萬/坪，歷史區間 {ch['historical_low_unit_price']:.2f} ~ {ch['historical_high_unit_price']:.2f} 萬/坪）。\n")

    lines.append(f"### 2.2 比佛利山莊純住宅歷年價格走勢統計清冊 (2012 ~ 2026)\n")
    lines.append(f"| 年度 (西元/民國) | 成交筆數 | 年度均價 (萬/坪) | 最低單價 (萬/坪) | 最高單價 (萬/坪) | 中位數單價 (萬/坪) | 價格週期與行情演變特徵 |")
    lines.append(f"| :---: | :---: | :---: | :---: | :---: | :---: | :--- |")
    for yt in ch["yearly_trends"]:
        lines.append(f"| **{yt['year']} ({yt['roc_year']}年)** | {yt['count']} | **{yt['avg_price']:.2f}** | {yt['min_price']:.2f} | {yt['max_price']:.2f} | {yt['median_price']:.2f} | {yt['note']} |")
    lines.append(f"| **全期累計基準** | **{ch['cleaned_residential_transactions']}** | **{ch['overall_average_unit_price']:.2f}** | **{ch['historical_low_unit_price']:.2f}** | **{ch['historical_high_unit_price']:.2f}** | **{ch['historical_median_unit_price']:.2f}** | 14年完整清洗純住宅走勢基準 |")
    lines.append(f"\n")

    lines.append(f"### 2.3 待售標的戶 2024 年購入成本重大鐵證深度剖析\n")
    lines.append(f"在歷史 158 筆成交清冊中，序號 `6398165` 證實了待售標的戶的前次交易細節：\n")
    lines.append(f"- **交易日期**：**2024-01-24（民國 113 年 1 月 24 日）**\n")
    lines.append(f"- **門牌樓層**：忠明南路730巷42號 13樓（與待售戶 100% 完全相符）\n")
    lines.append(f"- **權狀建坪**：33.75 坪（配置機械車位）\n")
    lines.append(f"- **成交總價**：**870 萬元**（登載單價 25.8 萬/坪）\n")
    lines.append(f"- **車位拆算淨單價**：依 70 萬拆算車位後，**純建物淨單價為 26.25 萬/坪**（(870-70)/30.48）。\n")
    lines.append(f"- **實質意義**：現任屋主購入至今甫滿 2 年，受限房地合一稅 2.0 稅率（35%），開價 998 萬意圖墊高 128 萬元出場。出價 870 萬可讓屋主保本平盤結案，為極佳之談判著力點。\n")
    lines.append(f"---\n")

    # 第三章
    lines.append(f"## 第三章：台中市南區周邊同類型競品社區實價行情多維對比 (競品比價)\n")
    lines.append(f"為客觀評估比佛利山莊之市場位階，盤點南區忠明南路、工學商圈與大慶雙鐵周邊 20 棟核心競品社區大數據清冊：\n")
    lines.append(f"| 社區名稱 | 門牌座落地址 | 屋齡 | 總戶數 | 總實登 | 純住筆數 | 113年均價 | 115年均價 | 近期均價 | 近期價格區間 | 相對比佛利核心特徵與比價意涵 |")
    lines.append(f"| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :--- |")
    for comp in comps:
        p113 = f"{comp['price_113']:.2f}" if comp['price_113'] else "—"
        p115 = f"{comp['price_115']:.2f}" if comp['price_115'] else "—"
        prec = f"{comp['price_recent']:.2f}" if comp['price_recent'] else "—"
        lines.append(f"| **{comp['name']}** | {comp['address']} | {comp['age']}年 | {comp['households']}戶 | {comp['total_trans']}筆 | {comp['clean_trans']}筆 | {p113} | {p115} | {prec} | {comp['price_range_recent']} | {comp['comparison']} |")
    lines.append(f"\n")

    lines.append(f"### 3.1 外部競品社區查核澄清 (Clarifications)\n")
    lines.append(f"1. **「熊都巴洛克」查核**：實際座落於南屯區三厝街 51 號（黎明商圈），為屋齡 19 年之透天別墅社區（總價 2,500~5,050 萬），非南區集合住宅大樓，不列入大樓估價對比。\n")
    lines.append(f"2. **「宏都雙子星」查核**：座落於南區復興北路 767 號（豐謙建設，屋齡 12 年，178 戶），座落十三期首排，近期均價 36~38 萬/坪，為南區次新屋之代表性天花板標竿。\n")
    lines.append(f"3. **「鄉林新月灣」查核**：座落於南區文心南路 888 號（鄉林建設，屋齡 15 年，342 戶，飯店式管理與全坡平車位），近期均價 34~35 萬/坪，為知名建商中生代指標。\n")
    lines.append(f"---\n")

    # 第四章
    lines.append(f"## 第四章：標準車位拆算模型與比價特徵調整矩陣 (車位拆算)\n")
    lines.append(f"### 4.1 南區車位大數據與標準車位拆算\n")
    lines.append(f"南區 54 社區車位實價登錄實證統計：\n")
    lines.append(f"- **機械車位 (Mechanical)**：全南區累計 430 筆明確拆算交易，中位數 60.0 萬元，近年主流成交區間為 **70 ~ 80 萬元**。\n")
    lines.append(f"- **平面車位 (Plane)**：全南區累計 813 筆明確拆算交易，中位數 109.0 萬元，近年主流成交區間為 **110 ~ 130 萬元**。\n")
    lines.append(f"- **待售標的車位拆算公式**：$$\\text{{房屋淨建坪單價}} = \\frac{{\\text{{總價}} - \\text{{車位價格}}}}{{33.75 - 3.27}} = \\frac{{\\text{{總價}} - \\text{{車位價格}}}}{{30.48\\text{{ 坪}}}}$$\n")

    lines.append(f"### 4.2 車位價格敏感度分析對比清單\n")
    lines.append(f"| 車位評估情境 | 車位設定價格 | 房屋淨開價總價 | 房屋淨建坪 | 房屋淨建坪單價 | 相對 2024 年買進淨單價 (26.25萬) 溢價 |\n| :--- | :---: | :---: | :---: | :---: | :---: |\n")
    for s in sens:
        diff_u = s['net_unit_price'] - vm['previous_sale_net_unit_price']
        diff_p = (diff_u / vm['previous_sale_net_unit_price']) * 100
        lines.append(f"| **{s['scenario']}** | {s['parking_price']:.0f} 萬元 | {s['net_asking_price']:.0f} 萬元 | {tp['house_area']:.2f} 坪 | **{s['net_unit_price']:.2f} 萬/坪** | +{diff_u:.2f} 萬/坪 (+{diff_p:.1f}%) |")
    lines.append(f"\n")

    lines.append(f"### 4.3 五大特徵維度比價調整矩陣 (Hedonic Adjustment Matrix)\n")
    lines.append(f"| 評估維度項目 | 待售戶特徵現況 | 基準比對條件 | 調整幅度與金額 | 專業評估推導依據 |")
    lines.append(f"| :--- | :--- | :--- | :---: | :--- |")
    for m in adj:
        adj_val = f"+{m['adjustment_percent']:.1f}%" if m.get("adjustment_percent", 0) > 0 else f"{m.get('adjustment_percent', 0):.1f}%" if "adjustment_percent" in m else f"{m.get('adjustment_amount', 0):.0f} 萬元"
        lines.append(f"| **{m['dimension']}** | {m['target_feature']} | {m['base_comparison']} | **{adj_val}** | {m['rationale']} |")
    lines.append(f"\n綜合評定：本戶具備次頂樓開闊視野（+5%）、雙衛開窗（+3%）與超低公設比高實坪（+3%）之實質優勢，但受限於屋齡 34 年（-2%）與機械車位抗性（相對平面折減約 40 萬元）。合理成交淨單價常態應落在 **25.5 ~ 26.5 萬/坪** 區間。\n")
    lines.append(f"---\n")

    # 第五章
    lines.append(f"## 第五章：三段式出價策略與定錨推導模型 (出價三道防線)\n")
    lines.append(f"基於賣方取得成本鐵證、車位拆算淨單價、南區 2026 大盤回修幅度與特徵調整矩陣，制定自住買方三段式出價防線：\n")
    lines.append(f"| 出價防線層級 | 建議出價總額 | 折合淨建坪單價 | 相對開價折數 | 總折減金額 | 核心定錨邏輯與談判攻防戰術依據 |")
    lines.append(f"| :--- | :---: | :---: | :---: | :---: | :--- |")
    lines.append(f"| **第一道防線：保守底標價**<br>*(大膽試探防線)* | **{vm['conservative']['total']:.0f} 萬元**<br>(區間 {vm['conservative']['range'][0]:.0f}~{vm['conservative']['range'][1]:.0f}萬) | **{vm['conservative']['net_unit_price']:.2f} 萬/坪** | **{vm['conservative']['discount_str']}** | **-{tp['asking_price'] - vm['conservative']['total']:.0f} 萬元** | **{vm['conservative']['rationale']}** |")
    lines.append(f"| **第二道防線：合理目標價**<br>*(最佳成交甜蜜點)* | **{vm['target']['total']:.0f} 萬元**<br>(區間 {vm['target']['range'][0]:.0f}~{vm['target']['range'][1]:.0f}萬) | **{vm['target']['net_unit_price']:.2f} 萬/坪** | **{vm['target']['discount_str']}** | **-{tp['asking_price'] - vm['target']['total']:.0f} 萬元** | **{vm['target']['rationale']}** |")
    lines.append(f"| **第三道防線：上限天花板價**<br>*(談判停損警戒線)* | **{vm['ceiling']['total']:.0f} 萬元**<br>(區間 {vm['ceiling']['range'][0]:.0f}~{vm['ceiling']['range'][1]:.0f}萬) | **{vm['ceiling']['net_unit_price']:.2f} 萬/坪** | **{vm['ceiling']['discount_str']}** | **-{tp['asking_price'] - vm['ceiling']['total']:.0f} 萬元** | **{vm['ceiling']['rationale']}** |")
    lines.append(f"\n")

    lines.append(f"### 5.1 各防線深度推導與執行要領\n")
    lines.append(f"1. **保守底標價（810 萬元 / 淨單價 24.28 萬/坪）**：\n")
    lines.append(f"   - 引用 2026 年最新比佛利 3 樓含平面車位以 800 萬成交（單價 20.4 萬/坪）之最新事實，並以「本戶為機械車位且屋齡已 34 年，買方必須預留 80 萬水電泥作翻新基金」為由下斡旋。\n")
    lines.append(f"   - 此價位大幅測試賣方資金流動性與讓利誠意，若賣方急需資金即可大勝促成。\n")
    lines.append(f"2. **合理目標價（870 萬元 / 淨單價 26.25 萬/坪）**：\n")
    lines.append(f"   - **精準定錨於屋主 2024 年 1 月買進成本 870 萬元**！\n")
    lines.append(f"   - 買方直接向仲介出示實價登錄鐵證：「屋主買 870 萬，我們願意出 870 萬讓屋主平盤保本出場，免受未來房價續跌風險」。此價位兼顧 13 樓次頂樓優質視野，並完全符合社區 2024-2026 年 3 房拆算淨單價中位數 26.29 萬/坪，為成交機率最高之黃金甜蜜點。\n")
    lines.append(f"3. **上限天花板價（920 萬元 / 淨單價 27.89 萬/坪）**：\n")
    lines.append(f"   - 買方自住決策之絕對停損紅線，對齊 2024 年全社區歷史多頭高點年度均價 27.81 萬/坪。\n")
    lines.append(f"   - 若屋主態度堅持超過 920 萬元，折合總價已逼近千萬元大關，買方應立即放棄，轉往南區具備坡道平面車位或屋齡 25 年以內之指標競品（如公園苑、歌德堡或椰城）。\n")
    lines.append(f"---\n")

    # 第六章
    lines.append(f"## 第六章：買方談判議價話術與經紀人攻防戰術 (談判話術)\n")
    lines.append(f"### 6.1 實戰定錨逐字稿與攻防說帖\n")
    lines.append(f"1. **話術一：鎖定屋主買進成本，主打「保本平盤」出場**：  \n")
    lines.append(f"   > **買方說法**：「我們調過實價登錄，很清楚屋主是在 113 年 1 月以 870 萬買進的。現在第七波信用管制後大盤量縮價跌，連同社區 3 樓今年最新成交都掉到 20.4 萬/坪了。我們是自住誠意買方，願意出 870 萬讓屋主完全平盤保本、把資金落袋為安。如果再拖半年，南區老宅價格續跌，屋主可能連本金都守不住。」\n")
    lines.append(f"2. **話術二：一般委託聯賣網絡，多渠道比價施壓**：  \n")
    lines.append(f"   > **買方說法**：「這間在中信跟永義都有掛一般委託，我們兩家都在談。哪家能幫我們談到 870 萬以內，而且服務費願意折讓到 1%，我們就立刻下斡旋簽約。請經紀人全力幫我們回報屋主。」\n")
    lines.append(f"3. **話術三：屋齡 34 年管線老舊，需扣減 80 萬水電泥作預算**：  \n")
    lines.append(f"   > **買方說法**：「比佛利是民國 81 年蓋好的，屋齡已經整整 34 年，外觀跟水電管線都已超過使用年限。我們買下來必須全部重新打掉重做電線迴路、給排水管與衛浴防水，這筆基礎工程至少要花 80 萬元。這筆成本必須如實反映在總價折減上。」\n")
    lines.append(f"4. **話術四：機械車位功能限制與持有成本抗性**：  \n")
    lines.append(f"   > **買方說法**：「B1-89 號是機械車位，不能停休旅車，每個月還要多繳 150 元車位保養清潔費，未來萬一機械故障修繕也是負擔。南區現在平面車位行情約 110 萬，機械車位市價只有 70 萬，相差了 40 萬元的價值，請屋主務必客觀扣減車位折舊。」\n")
    lines.append(f"---\n")

    # 第七章
    lines.append(f"## 第七章：實地看屋檢核清單與置產決策建議 (看屋檢核清單)\n")
    lines.append(f"### 7.1 看屋實戰 5 大核心查驗要項\n")
    for i, item in enumerate(nt["inspection_checklist"], 1):
        lines.append(f"1. **檢驗重點 {i}**：{item}。")
    lines.append(f"\n")
    lines.append(f"### 7.2 專案決策總結\n")
    lines.append(f"綜上所述，比佛利山莊 13 樓待售標的戶具備「次頂樓開闊視野」、「邊間三面採光」、「雙衛浴皆開窗」與「超低公設比 6.76%、室內實坪 28.4 坪」等罕見自住優勢，惟開價 998 萬顯著溢價 +14.7%（較買進成本墊高 128 萬）。  \n")
    lines.append(f"決策小組建議買方採取**以退為進、精準定錨**策略：初次試探出價 810~820 萬，逐步退守並將成交主力堅定錨定在 **870 萬元**（平盤買進價）。若賣方願讓利至 870 萬元區間，即為極具 CP 值之南區大三房自住置產選擇；若超過 920 萬元上限天花板，則應果斷放棄尋求其他競品。\n")

    content = "\n".join(lines)
    with open(MD_OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Generated Markdown: {MD_OUTPUT_PATH} ({os.path.getsize(MD_OUTPUT_PATH)} bytes)")


# ==============================================================================
# 2. WORD DOCUMENT GENERATOR (.docx)
# ==============================================================================
def set_run_font(run, font_name="微軟正黑體", size_pt=10, bold=False, italic=False, color_rgb=(45, 55, 72)):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:eastAsia="{font_name}"/>')
    rPr.append(rFonts)


def set_cell_background(cell, hex_color):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def set_table_borders(table, color="E2E8F0", sz="4"):
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)


def add_callout(doc, title, text, border_color="1B365D", bg_color="F8FAFC"):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.rows[0].cells[0]
    cell.width = Inches(6.8)

    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=180, right=160)

    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/>'
        f'<w:left w:val="single" w:sz="36" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if title:
        r_t = p.add_run(title + "\n")
        set_run_font(r_t, font_name="微軟正黑體", size_pt=10.5, bold=True, color_rgb=NAVY_PRIMARY)
    r_body = p.add_run(text)
    set_run_font(r_body, font_name="微軟正黑體", size_pt=9.5, bold=False, color_rgb=TEXT_DARK)

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Pt(4)


def add_chart(doc, img_path, caption="", width_in=6.2):
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(2)
        run = p_img.add_run()
        run.add_picture(img_path, width=Inches(width_in))

        if caption:
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_after = Pt(8)
            r_cap = p_cap.add_run(caption)
            set_run_font(r_cap, font_name="微軟正黑體", size_pt=8.5, bold=False, italic=True, color_rgb=TEXT_MUTED)


def generate_word_document(data):
    tp = data["target_property"]
    vm = data["valuation_model"]
    ch = data["community_history"]
    comps = data["comparables"]
    nt = data["negotiation_tactics"]
    prev = tp["previous_sale"]
    sens = vm["parking_sensitivity"]
    adj = vm["adjustment_matrix"]

    doc = Document()

    # Page Margins: 0.8 inch
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    # Document Header Badge & Title
    p_badge = doc.add_paragraph()
    p_badge.paragraph_format.space_before = Pt(0)
    p_badge.paragraph_format.space_after = Pt(4)
    r_badge = p_badge.add_run("不動產實價登錄鑑價與自住置產談判專案報告")
    set_run_font(r_badge, font_name="微軟正黑體", size_pt=10, bold=True, color_rgb=GOLD_ACCENT)

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("比佛利山莊社區整體市場行情分析與 13 樓特定待售戶別出價策略專案報告")
    set_run_font(r_title, font_name="微軟正黑體", size_pt=20, bold=True, color_rgb=NAVY_PRIMARY)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run(f"評估標的：台中市南區比佛利山莊 13 樓次頂樓大三房｜開價 {tp['asking_price']:.0f} 萬元｜2026 年 11 月專業精算版")
    set_run_font(r_sub, font_name="微軟正黑體", size_pt=10.5, bold=False, color_rgb=TEXT_MUTED)

    # Executive Summary Callout
    exec_summary_text = (
        f"【核心結論速覽】待售物件權狀 {tp['total_area']} 坪（純建坪 {tp['house_area']} 坪，室內實坪 {tp['indoor_area']} 坪），"
        f"開價 {tp['asking_price']:.0f} 萬（含 B1-89 坡道機械車位）。\n"
        f"1. 取得成本鐵證：現任屋主係於 2024-01-24 以 870 萬元購入，持滿 2 年開價墊高 +128 萬元（+14.7%）。\n"
        f"2. 車位拆算真實單價：扣除機械車位 70 萬後，純建物淨開價單價為 30.45 萬/坪（屋主買進淨單價為 26.25 萬/坪）。\n"
        f"3. 三段式出價建議：保守底標價 810 萬元（24.28 萬/坪，約 81 折）；合理目標價 870 萬元（26.25 萬/坪，約 87 折，錨定屋主買進成本）；上限天花板價 920 萬元（27.89 萬/坪，約 92 折）。"
    )
    add_callout(doc, "【執行摘要】自住出價評估核心發現與決策防線", exec_summary_text, border_color="1B365D", bg_color="F8FAFC")

    # Chapter 1
    p_c1 = doc.add_paragraph()
    p_c1.paragraph_format.space_before = Pt(16)
    p_c1.paragraph_format.space_after = Pt(6)
    r_c1 = p_c1.add_run("第一章：委託評估標的物件規格與現況查核 (物件現況)")
    set_run_font(r_c1, font_name="微軟正黑體", size_pt=14, bold=True, color_rgb=NAVY_PRIMARY)

    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_after = Pt(6)
    r_p1 = p_p1.add_run("本專案評估之標的戶座落於台中市南區忠明南路730巷42號，經 591 售屋網快照與謄本完整交叉比對，規格數據權威清單如下：")
    set_run_font(r_p1, font_name="微軟正黑體", size_pt=10, bold=False, color_rgb=TEXT_DARK)

    # Table 1: Property Specifications (Table 1)
    spec_rows = [
        ("社區名稱 / 代碼", f"{tp['community_name']} (ID: {tp['community_id']})", "總戶數 309 戶，1992 年完工成熟大樓"),
        ("建物門牌座落", f"{tp['full_address']}", "面臨忠明南路 730 巷，寧靜街廓"),
        ("所在樓層 / 總樓層", f"{tp['floor']}", "黃金次頂樓，視野開闊，無頂樓受熱與漏水疑慮"),
        ("總開價 (含車位)", f"{tp['asking_price']:.0f} 萬元", "毛開價單價 29.57 萬/坪（以 33.75 坪計算）"),
        ("權狀總坪數", f"{tp['total_area']:.2f} 坪", "含房屋 30.48 坪 + 車位約定專用持分 3.27 坪"),
        ("主建物面積", f"{tp['main_area']:.2f} 坪", "室內淨起居空間充足，標準大三房規格"),
        ("附屬建物面積", f"{tp['attached_area']:.2f} 坪", "前後陽台通風佳，無虛坪雨遮"),
        ("主附合計實坪", f"{tp['indoor_area']:.2f} 坪", f"實質得房率達 {tp['indoor_ratio']:.2f}%，空間效能極高"),
        ("大樓共用部分 (公設)", f"{tp['public_area']:.2f} 坪", f"超低公設比僅 {tp['public_ratio']:.2f}%，無浮濫公設灌水"),
        ("車位型態與編號", f"{tp['parking_type']}", "地下一層 89 號坡道機械車位"),
        ("室內格局配置", f"{tp['layout']}", "邊間三面採光，罕見雙衛浴皆有對外窗"),
        ("建物座向朝向", f"{tp['orientation']}", "冬暖夏涼，客廳落地窗朝東南避開北風"),
        ("建築完工日 / 屋齡", f"{tp['build_date']} / 約 {tp['building_age']:.1f} 年", "已屆滿 34 年老宅門檻，水電管線需翻新"),
        ("前次成交紀錄", f"{prev['date']} 以 {prev['price']:.0f} 萬成交", f"取得成本鐵證，屋主開價墊高 +{prev['markup_amount']:.0f} 萬 (+{prev['markup_percent']:.1f}%)")
    ]
    t1 = doc.add_table(rows=len(spec_rows) + 1, cols=3)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t1)

    headers1 = ["產權規格維度", "查核登記數值", "專業評估意涵與分析說明"]
    for i, h in enumerate(headers1):
        cell = t1.cell(0, i)
        set_cell_background(cell, HEADER_BG)
        set_cell_margins(cell, 120, 120, 140, 140)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_run_font(r, font_name="微軟正黑體", size_pt=9.5, bold=True, color_rgb=(255, 255, 255))

    for r_idx, row_data in enumerate(spec_rows, start=1):
        bg = BG_LIGHT if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = t1.cell(r_idx, c_idx)
            set_cell_background(cell, bg)
            set_cell_margins(cell, 100, 100, 120, 120)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            is_bold = (c_idx == 0 or (c_idx == 1 and ("998" in val or "870" in val)))
            c_rgb = NAVY_PRIMARY if is_bold else TEXT_DARK
            set_run_font(r, font_name="微軟正黑體", size_pt=9, bold=is_bold, color_rgb=c_rgb)

    p_sp1 = doc.add_paragraph()
    p_sp1.paragraph_format.space_after = Pt(10)

    # Chapter 2
    p_c2 = doc.add_paragraph()
    p_c2.paragraph_format.space_before = Pt(16)
    p_c2.paragraph_format.space_after = Pt(6)
    r_c2 = p_c2.add_run("第二章：比佛利山莊歷史實價登錄全景與行情走勢 (社區歷程)")
    set_run_font(r_c2, font_name="微軟正黑體", size_pt=14, bold=True, color_rgb=NAVY_PRIMARY)

    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_after = Pt(6)
    r_p2 = p_p2.add_run(
        f"比佛利山莊自 2012 年實價登錄實施以來累計成交 158 筆。經由標準 5 層資料清洗管線，"
        f"嚴格排除 1 樓及 1,2 樓店面（16 筆）、親友與特殊交易（20 筆，含 2 筆重疊），"
        f"共取得 {ch['cleaned_residential_transactions']} 筆純住宅成交紀錄（全期均價 {ch['overall_average_unit_price']:.2f} 萬/坪）。"
        f"歷年純住宅均價走勢統計如下表："
    )
    set_run_font(r_p2, font_name="微軟正黑體", size_pt=10, bold=False, color_rgb=TEXT_DARK)

    # Table 2: Community Yearly Trends (Table 2)
    t2 = doc.add_table(rows=len(ch["yearly_trends"]) + 1, cols=6)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t2)

    headers2 = ["年度 (西元/民國)", "成交筆數", "年度均價 (萬/坪)", "價格區間 (萬/坪)", "中位數 (萬/坪)", "價格週期與市場特徵"]
    for i, h in enumerate(headers2):
        cell = t2.cell(0, i)
        set_cell_background(cell, HEADER_BG)
        set_cell_margins(cell, 120, 120, 120, 120)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_run_font(r, font_name="微軟正黑體", size_pt=9.5, bold=True, color_rgb=(255, 255, 255))

    for r_idx, yt in enumerate(ch["yearly_trends"], start=1):
        bg = BG_LIGHT if r_idx % 2 == 1 else "FFFFFF"
        row_vals = [
            f"{yt['year']} ({yt['roc_year']}年)",
            str(yt['count']),
            f"{yt['avg_price']:.2f}",
            f"{yt['min_price']:.2f} ~ {yt['max_price']:.2f}",
            f"{yt['median_price']:.2f}",
            yt['note']
        ]
        for c_idx, val in enumerate(row_vals):
            cell = t2.cell(r_idx, c_idx)
            set_cell_background(cell, bg)
            set_cell_margins(cell, 90, 90, 110, 110)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            is_bold = (c_idx == 2 or (c_idx == 0 and yt['year'] >= 2024))
            c_rgb = RED_ALERT if (c_idx == 2 and yt['year'] >= 2024) else TEXT_DARK
            set_run_font(r, font_name="微軟正黑體", size_pt=8.5, bold=is_bold, color_rgb=c_rgb)

    p_sp2 = doc.add_paragraph()
    p_sp2.paragraph_format.space_after = Pt(6)

    # Embed Chart
    add_chart(doc, CHART_BEVERLY, caption="圖 1：比佛利山莊 2012-2026 年歷史成交行情趨勢圖（內嵌高解析度資料庫圖表）", width_in=6.2)

    # Chapter 3: Comparables
    p_c3 = doc.add_paragraph()
    p_c3.paragraph_format.space_before = Pt(16)
    p_c3.paragraph_format.space_after = Pt(6)
    r_c3 = p_c3.add_run("第三章：台中市南區周邊同類型競品社區實價行情多維對比 (競品比價)")
    set_run_font(r_c3, font_name="微軟正黑體", size_pt=14, bold=True, color_rgb=NAVY_PRIMARY)

    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_after = Pt(6)
    r_p3 = p_p3.add_run(
        "為客觀評估比佛利山莊在南區生活圈之合理市場定位，盤點忠明南路、工學商圈與大慶雙鐵周邊 20 棟核心競品社區大數據如下表："
    )
    set_run_font(r_p3, font_name="微軟正黑體", size_pt=10, bold=False, color_rgb=TEXT_DARK)

    # Table 3: Competitor Communities (Table 3)
    t3 = doc.add_table(rows=len(comps) + 1, cols=7)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t3)

    headers3 = ["社區名稱", "門牌座落", "屋齡", "戶數", "113年均價", "115年均價", "相對比價意涵與核心特徵"]
    for i, h in enumerate(headers3):
        cell = t3.cell(0, i)
        set_cell_background(cell, HEADER_BG)
        set_cell_margins(cell, 120, 120, 100, 100)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_run_font(r, font_name="微軟正黑體", size_pt=9, bold=True, color_rgb=(255, 255, 255))

    for r_idx, comp in enumerate(comps, start=1):
        bg = BG_LIGHT if r_idx % 2 == 1 else "FFFFFF"
        p113 = f"{comp['price_113']:.1f}萬" if comp['price_113'] else "—"
        p115 = f"{comp['price_115']:.1f}萬" if comp['price_115'] else "—"
        row_vals = [
            comp['name'],
            comp['address'],
            f"{comp['age']}年",
            f"{comp['households']}戶",
            p113,
            p115,
            comp['comparison']
        ]
        for c_idx, val in enumerate(row_vals):
            cell = t3.cell(r_idx, c_idx)
            set_cell_background(cell, bg)
            set_cell_margins(cell, 80, 80, 100, 100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            is_target = "比佛利山莊 (標的)" in comp['name']
            c_rgb = NAVY_PRIMARY if is_target else TEXT_DARK
            set_run_font(r, font_name="微軟正黑體", size_pt=8, bold=(is_target or c_idx == 0), color_rgb=c_rgb)

    p_sp3 = doc.add_paragraph()
    p_sp3.paragraph_format.space_after = Pt(10)

    # Chapter 4: Parking Separation & Adjustment Matrix
    p_c4 = doc.add_paragraph()
    p_c4.paragraph_format.space_before = Pt(16)
    p_c4.paragraph_format.space_after = Pt(6)
    r_c4 = p_c4.add_run("第四章：標準車位拆算模型與比價特徵調整矩陣 (車位拆算)")
    set_run_font(r_c4, font_name="微軟正黑體", size_pt=14, bold=True, color_rgb=NAVY_PRIMARY)

    p_p4 = doc.add_paragraph()
    p_p4.paragraph_format.space_after = Pt(6)
    r_p4 = p_p4.add_run(
        "依據內政部實價登錄規範，計算房屋真實單價必須將車位價格與坪數分離。南區 430 筆機械車位行情中位數為 60 萬，"
        f"主流成交為 70~80 萬元。以 70 萬元為基準拆算，待售戶純建坪（30.48 坪）淨開價單價為 **{vm['asking_net_unit_price']:.2f} 萬/坪**。"
        "車位價格敏感度與特徵調整矩陣如下："
    )
    set_run_font(r_p4, font_name="微軟正黑體", size_pt=10, bold=False, color_rgb=TEXT_DARK)

    # Table 4: Parking Sensitivity Table (Table 4)
    t4 = doc.add_table(rows=len(sens) + 1, cols=5)
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t4)

    headers4 = ["評估情境", "車位設定價格", "房屋淨總價", "純建坪", "房屋淨開價單價"]
    for i, h in enumerate(headers4):
        cell = t4.cell(0, i)
        set_cell_background(cell, HEADER_BG)
        set_cell_margins(cell, 120, 120, 120, 120)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_run_font(r, font_name="微軟正黑體", size_pt=9.5, bold=True, color_rgb=(255, 255, 255))

    for r_idx, s in enumerate(sens, start=1):
        bg = BG_LIGHT if r_idx % 2 == 1 else "FFFFFF"
        row_vals = [
            s['scenario'],
            f"{s['parking_price']:.0f} 萬元",
            f"{s['net_asking_price']:.0f} 萬元",
            f"{tp['house_area']:.2f} 坪",
            f"{s['net_unit_price']:.2f} 萬/坪"
        ]
        for c_idx, val in enumerate(row_vals):
            cell = t4.cell(r_idx, c_idx)
            set_cell_background(cell, bg)
            set_cell_margins(cell, 90, 90, 120, 120)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            is_bold = (c_idx == 4 or "基準" in s['scenario'])
            c_rgb = RED_ALERT if c_idx == 4 else TEXT_DARK
            set_run_font(r, font_name="微軟正黑體", size_pt=9, bold=is_bold, color_rgb=c_rgb)

    p_sp4 = doc.add_paragraph()
    p_sp4.paragraph_format.space_after = Pt(8)

    # Chapter 5: Three-tier Bidding
    p_c5 = doc.add_paragraph()
    p_c5.paragraph_format.space_before = Pt(16)
    p_c5.paragraph_format.space_after = Pt(6)
    r_c5 = p_c5.add_run("第五章：三段式出價策略與定錨推導模型 (出價三道防線)")
    set_run_font(r_c5, font_name="微軟正黑體", size_pt=14, bold=True, color_rgb=NAVY_PRIMARY)

    # Table 5: Three-Tier Bidding Strategy (Table 5)
    bidding_rows = [
        ("第一道防線：保守底標價", f"{vm['conservative']['total']:.0f} 萬元", f"{vm['conservative']['net_unit_price']:.2f} 萬/坪", vm['conservative']['discount_str'], f"-{tp['asking_price'] - vm['conservative']['total']:.0f} 萬元", vm['conservative']['rationale']),
        ("第二道防線：合理目標價", f"{vm['target']['total']:.0f} 萬元", f"{vm['target']['net_unit_price']:.2f} 萬/坪", vm['target']['discount_str'], f"-{tp['asking_price'] - vm['target']['total']:.0f} 萬元", vm['target']['rationale']),
        ("第三道防線：上限天花板價", f"{vm['ceiling']['total']:.0f} 萬元", f"{vm['ceiling']['net_unit_price']:.2f} 萬/坪", vm['ceiling']['discount_str'], f"-{tp['asking_price'] - vm['ceiling']['total']:.0f} 萬元", vm['ceiling']['rationale'])
    ]
    t5 = doc.add_table(rows=len(bidding_rows) + 1, cols=6)
    t5.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t5)

    headers5 = ["防線層級", "建議出價", "折合淨單價", "相對開價", "折減金額", "核心定錨邏輯與談判戰術依據"]
    for i, h in enumerate(headers5):
        cell = t5.cell(0, i)
        set_cell_background(cell, HEADER_BG)
        set_cell_margins(cell, 120, 120, 120, 120)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_run_font(r, font_name="微軟正黑體", size_pt=9.5, bold=True, color_rgb=(255, 255, 255))

    for r_idx, brow in enumerate(bidding_rows, start=1):
        bg = BG_LIGHT if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(brow):
            cell = t5.cell(r_idx, c_idx)
            set_cell_background(cell, bg)
            set_cell_margins(cell, 100, 100, 120, 120)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            is_bold = (c_idx in (0, 1, 2))
            c_rgb = GREEN_ACCENT if r_idx == 1 and is_bold else GOLD_ACCENT if r_idx == 2 and is_bold else RED_ALERT if r_idx == 3 and is_bold else TEXT_DARK
            set_run_font(r, font_name="微軟正黑體", size_pt=9, bold=is_bold, color_rgb=c_rgb)

    p_sp5 = doc.add_paragraph()
    p_sp5.paragraph_format.space_after = Pt(10)

    # Chapter 6: Negotiation Scripts
    p_c6 = doc.add_paragraph()
    p_c6.paragraph_format.space_before = Pt(16)
    p_c6.paragraph_format.space_after = Pt(6)
    r_c6 = p_c6.add_run("第六章：買方談判議價話術與經紀人攻防戰術 (談判話術)")
    set_run_font(r_c6, font_name="微軟正黑體", size_pt=14, bold=True, color_rgb=NAVY_PRIMARY)

    add_callout(doc, "戰術話術一：鎖定 870 萬取得成本，主打「保本平盤」出場", 
                "「我們調過實價登錄，很清楚屋主是在 113 年 1 月以 870 萬買進的。現在第七波信用管制後大盤量縮價跌，連同社區 3 樓今年最新成交都掉到 20.4 萬/坪。我們是自住誠意買方，願意出 870 萬讓屋主完全平盤保本、把現金落袋為安。若再拖延，老宅續跌屋主恐面臨虧損。」",
                border_color="D98C14", bg_color="F8FAFC")

    add_callout(doc, "戰術話術二：一般委託聯賣比價，服務費競爭激勵", 
                "「這間在中信與永義都有一般委託聯賣，哪一家能幫我們促成 870 萬以內成交、服務費願意折讓至 1%，我們就立刻下斡旋簽約。請經紀人全力協助回報屋主。」",
                border_color="1B365D", bg_color="F8FAFC")

    add_callout(doc, "戰術話術三：屋齡 34 年管線老舊，需扣減 80 萬泥作水電翻修預算", 
                "「比佛利山莊已屆滿 34 年屋齡，水電給排水管與電線早已超過安全年限，買方購入後必須全面重拉總電盤與重做衛浴防水，基礎工程需花費 80 萬元，必須由總價中具體折減。」",
                border_color="C0392B", bg_color="F8FAFC")

    # Chapter 7: Field Inspection Checklist
    p_c7 = doc.add_paragraph()
    p_c7.paragraph_format.space_before = Pt(16)
    p_c7.paragraph_format.space_after = Pt(6)
    r_c7 = p_c7.add_run("第七章：實地看屋檢核清單與置產決策結論 (看屋檢核清單)")
    set_run_font(r_c7, font_name="微軟正黑體", size_pt=14, bold=True, color_rgb=NAVY_PRIMARY)

    # Table 6: Inspection Checklist Table (Table 6)
    checklist_rows = [
        ("1. 外牆窗框與滲漏水", "檢查東南向迎風面窗框周邊、冷氣孔是否有壁癌補土痕跡或水痕。"),
        ("2. 雙衛浴管道間結構", "推開雙衛浴天花板維修孔，檢查上方給排水管與共用管道間有無滲水。"),
        ("3. 總電盤迴路與安培", "檢查總開關箱安培數是否為早期 30A~50A，確認重拉至 75A 之成本。"),
        ("4. B1-89 機械車位操作", "實地操作升降車台，測量長寬高淨尺寸與限重，確認家用車輛停放順暢度。"),
        ("5. 社區公共基金與頂樓", "調閱管委會財務收支報表與修繕基金餘額，並至 14 樓頂樓檢查防水層維護。")
    ]
    t6 = doc.add_table(rows=len(checklist_rows) + 1, cols=2)
    t6.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t6)

    headers6 = ["檢驗項目", "實地看屋查驗重點與判定標準"]
    for i, h in enumerate(headers6):
        cell = t6.cell(0, i)
        set_cell_background(cell, HEADER_BG)
        set_cell_margins(cell, 120, 120, 120, 120)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_run_font(r, font_name="微軟正黑體", size_pt=9.5, bold=True, color_rgb=(255, 255, 255))

    for r_idx, (item, desc) in enumerate(checklist_rows, start=1):
        bg = BG_LIGHT if r_idx % 2 == 1 else "FFFFFF"
        cell0 = t6.cell(r_idx, 0)
        set_cell_background(cell0, bg)
        set_cell_margins(cell0, 100, 100, 120, 120)
        p0 = cell0.paragraphs[0]
        r0 = p0.add_run(item)
        set_run_font(r0, font_name="微軟正黑體", size_pt=9, bold=True, color_rgb=NAVY_PRIMARY)

        cell1 = t6.cell(r_idx, 1)
        set_cell_background(cell1, bg)
        set_cell_margins(cell1, 100, 100, 120, 120)
        p1 = cell1.paragraphs[0]
        r1 = p1.add_run(desc)
        set_run_font(r1, font_name="微軟正黑體", size_pt=9, bold=False, color_rgb=TEXT_DARK)

    p_final = doc.add_paragraph()
    p_final.paragraph_format.space_before = Pt(12)
    r_final = p_final.add_run("【決策總結】本案為南區忠明南路生活圈難得之黃金次頂樓視野戶，建議買方以 870 萬元（平盤保本點）為核心目標價積極斡旋；若能順利促成，將享有超低公設比高實坪之極佳自住效益。")
    set_run_font(r_final, font_name="微軟正黑體", size_pt=10, bold=True, color_rgb=NAVY_PRIMARY)

    doc.save(DOCX_OUTPUT_PATH)
    print(f"Generated Word (.docx): {DOCX_OUTPUT_PATH} ({os.path.getsize(DOCX_OUTPUT_PATH)} bytes)")


# ==============================================================================
# 3. POWERPOINT PRESENTATION GENERATOR (.pptx)
# ==============================================================================
def set_shape_flat(shape, fill_color, border_color=None, border_width=1):
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = PptxPt(border_width)
    else:
        shape.line.fill.background()


def add_slide_header(slide, title_text, category="比佛利山莊 13 樓・特定待售戶出價分析與社區行情報告"):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, PptxInches(13.333), PptxInches(7.5))
    set_shape_flat(bg, PPTX_LIGHT_BG)

    tb = slide.shapes.add_textbox(PptxInches(0.8), PptxInches(0.4), PptxInches(11.733), PptxInches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_cat = tf.paragraphs[0]
    p_cat.text = category.upper()
    p_cat.font.name = "微軟正黑體"
    p_cat.font.size = PptxPt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = PPTX_GOLD

    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.name = "微軟正黑體"
    p_title.font.size = PptxPt(20)
    p_title.font.bold = True
    p_title.font.color.rgb = PPTX_NAVY
    p_title.space_before = PptxPt(2)


def add_card_shape(slide, left, top, width, height, bg_color=PPTX_CARD_BG, border_color=PPTX_BORDER_GRAY):
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    set_shape_flat(card, bg_color, border_color, 1)
    return card


def generate_powerpoint_presentation(data):
    tp = data["target_property"]
    vm = data["valuation_model"]
    ch = data["community_history"]
    comps = data["comparables"]
    nt = data["negotiation_tactics"]
    prev = tp["previous_sale"]
    sens = vm["parking_sensitivity"]
    adj = vm["adjustment_matrix"]

    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    blank_layout = prs.slide_layouts[6]

    # ----------------------------------------------------
    # SLIDE 1: 封面
    # ----------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, PptxInches(13.333), PptxInches(7.5))
    set_shape_flat(bg1, PPTX_DARK_BLUE)

    # Accent decorative bar
    bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptxInches(0.8), PptxInches(1.8), PptxInches(0.15), PptxInches(3.8))
    set_shape_flat(bar, PPTX_GOLD)

    tb1 = s1.shapes.add_textbox(PptxInches(1.3), PptxInches(1.8), PptxInches(11.0), PptxInches(4.0))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p0 = tf1.paragraphs[0]
    p0.text = "不動產實價登錄鑑價與自住置產出價決策專案"
    p0.font.name = "微軟正黑體"
    p0.font.size = PptxPt(13)
    p0.font.bold = True
    p0.font.color.rgb = PPTX_GOLD

    p1 = tf1.add_paragraph()
    p1.text = "比佛利山莊 13 樓 次頂樓大三房\n特定待售戶出價評估與全社區行情報告"
    p1.font.name = "微軟正黑體"
    p1.font.size = PptxPt(32)
    p1.font.bold = True
    p1.font.color.rgb = PPTX_WHITE
    p1.space_before = PptxPt(10)
    p1.space_after = PptxPt(14)

    p2 = tf1.add_paragraph()
    p2.text = f"評估標的：忠明南路730巷42號 13F｜開價 {tp['asking_price']:.0f} 萬｜權狀 {tp['total_area']} 坪 (純建坪 {tp['house_area']} 坪 / 實坪 {tp['indoor_area']} 坪)｜B1-89 坡道機械車位"
    p2.font.name = "微軟正黑體"
    p2.font.size = PptxPt(13)
    p2.font.color.rgb = PptxRGBColor(203, 213, 225)

    p3 = tf1.add_paragraph()
    p3.text = "報告單位：自住置產專案決策小組 ｜ 數據基底：內政部實價登錄 158 筆清洗、謄本與 591 產調 ｜ 2026 年 11 月專業版"
    p3.font.name = "微軟正黑體"
    p3.font.size = PptxPt(11)
    p3.font.color.rgb = PPTX_TEXT_MUTED
    p3.space_before = PptxPt(20)

    # ----------------------------------------------------
    # SLIDE 2: 執行摘要 (Executive Summary)
    # ----------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_header(s2, "執行摘要：核心發現與三段式出價防線速覽")

    cards_s2 = [
        ("取得成本鐵證", f"現任屋主於 2024-01-24 以 870 萬元買進\n持有剛滿 2 年，目前開價 998 萬墊高 +128 萬 (+14.7%)\n出價 870 萬即為屋主平盤保本出場點", PPTX_RED),
        ("車位拆算真實單價", f"南區坡道機械車位公允市價以 70 萬拆算\n房屋純建坪 30.48 坪，淨開價單價 30.45 萬/坪\n相較買進淨單價 26.25 萬/坪，開價溢價 +16.0%", PPTX_NAVY),
        ("大盤降溫老宅回修", f"央行第七波信用管制後，南區 30 年老屋普遍下修\n2026 年比佛利 3 樓平車戶最新成交回落至 20.4 萬/坪\n歌德堡跌至 26.6 萬，長億南園跌至 22.6 萬", PPTX_SLATE),
        ("稀缺物理條件", f"13/14F 黃金次頂樓無遮蔽視野、冬暖夏涼朝東南\n邊間三面採光、極罕見雙衛浴皆對外開窗\n超低公設比 6.76%，室內實坪 28.4 坪得房率 93.2%", PPTX_GREEN)
    ]
    for idx, (title, text, accent) in enumerate(cards_s2):
        col = idx % 2
        row = idx // 2
        l = PptxInches(0.8 + col * 5.95)
        t = PptxInches(1.5 + row * 2.7)
        add_card_shape(s2, l, t, PptxInches(5.7), PptxInches(2.5))

        top_bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, PptxInches(5.7), PptxInches(0.08))
        set_shape_flat(top_bar, accent)

        tb = s2.shapes.add_textbox(l + PptxInches(0.3), t + PptxInches(0.2), PptxInches(5.1), PptxInches(2.1))
        tf = tb.text_frame
        tf.word_wrap = True
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.name = "微軟正黑體"
        p_t.font.size = PptxPt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = accent

        p_b = tf.add_paragraph()
        p_b.text = text
        p_b.font.name = "微軟正黑體"
        p_b.font.size = PptxPt(11)
        p_b.font.color.rgb = PPTX_TEXT_DARK
        p_b.space_before = PptxPt(6)

    # ----------------------------------------------------
    # SLIDE 3: 物件規格與 591 產調比對
    # ----------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_header(s3, "產調查核：建物物理規格與 591 刊登差異比對")

    # Left card: Key specs
    add_card_shape(s3, PptxInches(0.8), PptxInches(1.5), PptxInches(4.5), PptxInches(5.4))
    tb_ls3 = s3.shapes.add_textbox(PptxInches(1.1), PptxInches(1.7), PptxInches(3.9), PptxInches(5.0))
    tf_ls3 = tb_ls3.text_frame
    tf_ls3.word_wrap = True

    p = tf_ls3.paragraphs[0]
    p.text = "【待售標的核心規格】"
    p.font.name = "微軟正黑體"
    p.font.size = PptxPt(15)
    p.font.bold = True
    p.font.color.rgb = PPTX_NAVY

    specs_bullet = [
        f"建物門牌：{tp['full_address']}",
        f"樓層格局：13F/14F 黃金次頂樓視野",
        f"室內配置：3 房 2 廳 2 衛 1 陽台",
        f"座向朝向：坐西北朝東南（冬暖夏涼）",
        f"純建坪數：{tp['house_area']:.2f} 坪（主建 {tp['main_area']} + 陽台 {tp['attached_area']}）",
        f"室內實坪：{tp['indoor_area']:.2f} 坪（實質得房率 93.2%）",
        f"大樓公設：僅 {tp['public_area']} 坪（超低公設比 {tp['public_ratio']}%）",
        f"車位規格：坡道機械式 B1-89 號（約 3.27 坪）",
        f"法定屋齡：{tp['building_age']} 年（1992-11 完工）",
        f"管理費用：1,750 元/月（含車位清潔費）",
        f"委託型態：一般委託（中信、永義聯賣）"
    ]
    for b in specs_bullet:
        pb = tf_ls3.add_paragraph()
        pb.text = "• " + b
        pb.font.name = "微軟正黑體"
        pb.font.size = PptxPt(11)
        pb.font.color.rgb = PPTX_TEXT_DARK
        pb.space_before = PptxPt(4)

    # Right card: Comparison table
    add_card_shape(s3, PptxInches(5.6), PptxInches(1.5), PptxInches(6.9), PptxInches(5.4))
    tb_rs3 = s3.shapes.add_textbox(PptxInches(5.9), PptxInches(1.7), PptxInches(6.3), PptxInches(5.0))
    tf_rs3 = tb_rs3.text_frame
    tf_rs3.word_wrap = True

    p_rt = tf_rs3.paragraphs[0]
    p_rt.text = "【591 廣告 vs. 謄本真偽比對關鍵】"
    p_rt.font.name = "微軟正黑體"
    p_rt.font.size = PptxPt(15)
    p_rt.font.bold = True
    p_rt.font.color.rgb = PPTX_GOLD

    diffs = [
        ("權狀面積說明", "591 登載 33.75 坪（含車位），但坪數細項加總僅 30.48 坪。經查 3.27 坪為仲介手動加入之車位約定專用坪數，純建物權狀為 30.48 坪。"),
        ("毛單價視覺壓低", "591 標示 29.57 萬/坪係將總價除以灌入車位之 33.75 坪。若未拆車位除以純建物達 32.74 萬/坪；標準扣除 70 萬車位後淨單價為 30.45 萬/坪。"),
        ("車位細節揭露", "591 僅寫「機械式，3.27坪」，永義資料完整揭露為「坡道機械式，B1-89 號」，進出為汽車坡道入地下室後停入機械車位。"),
        ("屋齡更新落後", "591 刊載 32 年/33 年為陳舊文案，建物完工日為 1992 年 11 月 20 日，至 2026 年已屆滿 34 年，已跨過銀行 30 年老屋嚴審限貸門檻。"),
        ("聯賣談判優勢", "證實由中信房屋與永義房屋共同銷售，屬一般約聯賣，買方無需被單一仲介綁定，具備雙邊比價與服務費折讓之議價籌碼。")
    ]
    for d_title, d_desc in diffs:
        p_dt = tf_rs3.add_paragraph()
        p_dt.text = "▲ " + d_title
        p_dt.font.name = "微軟正黑體"
        p_dt.font.size = PptxPt(11.5)
        p_dt.font.bold = True
        p_dt.font.color.rgb = PPTX_NAVY
        p_dt.space_before = PptxPt(5)

        p_dd = tf_rs3.add_paragraph()
        p_dd.text = d_desc
        p_dd.font.name = "微軟正黑體"
        p_dd.font.size = PptxPt(10)
        p_dd.font.color.rgb = PPTX_TEXT_MUTED
        p_dd.space_before = PptxPt(1)

    # ----------------------------------------------------
    # SLIDE 4: 買進成本鐵證深度剖析
    # ----------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_header(s4, "關鍵鐵證：待售戶 2024 年 870 萬購入成本與加價剖析")

    # Big highlight banner
    banner = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptxInches(0.8), PptxInches(1.5), PptxInches(11.733), PptxInches(1.3))
    set_shape_flat(banner, PPTX_DARK_BLUE)
    tb_b = s4.shapes.add_textbox(PptxInches(1.1), PptxInches(1.6), PptxInches(11.1), PptxInches(1.1))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    p_b1 = tf_b.paragraphs[0]
    p_b1.text = f"實登鐵證序號 6398165：2024-01-24 以 870 萬元成交（含機械車位，單價 25.8 萬/坪）"
    p_b1.font.name = "微軟正黑體"
    p_b1.font.size = PptxPt(16)
    p_b1.font.bold = True
    p_b1.font.color.rgb = PPTX_GOLD

    p_b2 = tf_b.add_paragraph()
    p_b2.text = f"現任屋主持有甫滿 2 年，目前開價 998 萬元，帳面墊高加價達 +128 萬元（+14.7%），扣除車位淨單價由 26.25 萬墊高至 30.45 萬！"
    p_b2.font.name = "微軟正黑體"
    p_b2.font.size = PptxPt(12)
    p_b2.font.color.rgb = PPTX_WHITE
    p_b2.space_before = PptxPt(4)

    # 3 comparison cards below
    cards_s4 = [
        ("屋主 2024 購入條件", [
            f"交易日期：{prev['date']} (113-01-24)",
            f"購入總價：{prev['price']:.0f} 萬元",
            f"刊登毛單價：{prev['unit_price']} 萬/坪",
            f"車位拆算淨總價：800 萬元 (扣70萬)",
            f"純建坪淨單價：26.25 萬/坪"
        ], PPTX_NAVY),
        ("目前 2026 待售開價", [
            f"目前開價：{tp['asking_price']:.0f} 萬元",
            f"帳面加價：+{prev['markup_amount']:.0f} 萬元 (+{prev['markup_percent']:.1f}%)",
            f"刊登毛單價：{tp['asking_gross_unit_price']} 萬/坪",
            f"車位拆算淨總價：928 萬元 (扣70萬)",
            f"純建坪淨單價：30.45 萬/坪 (+4.20萬)"
        ], PPTX_RED),
        ("買方談判策略定錨", [
            f"持有年限：剛滿 2 年，受限房地合一稅",
            f"屋主心理：最渴望的是保本平盤出場",
            f"談判底牌：870 萬元即符合屋主成本",
            f"大盤情勢：2026年南區30年老宅下修",
            f"結論：買方享有完全不追高之優勢"
        ], PPTX_GREEN)
    ]
    for idx, (t_title, items, border_c) in enumerate(cards_s4):
        l = PptxInches(0.8 + idx * 4.0)
        t = PptxInches(3.1)
        add_card_shape(s4, l, t, PptxInches(3.75), PptxInches(3.8))

        top_strip = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, PptxInches(3.75), PptxInches(0.08))
        set_shape_flat(top_strip, border_c)

        tb = s4.shapes.add_textbox(l + PptxInches(0.2), t + PptxInches(0.2), PptxInches(3.35), PptxInches(3.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p_h = tf.paragraphs[0]
        p_h.text = t_title
        p_h.font.name = "微軟正黑體"
        p_h.font.size = PptxPt(13.5)
        p_h.font.bold = True
        p_h.font.color.rgb = border_c

        for it in items:
            pi = tf.add_paragraph()
            pi.text = "• " + it
            pi.font.name = "微軟正黑體"
            pi.font.size = PptxPt(10.5)
            pi.font.color.rgb = PPTX_TEXT_DARK
            pi.space_before = PptxPt(5)

    # ----------------------------------------------------
    # SLIDE 5: 社區實登清洗歷程與趨勢圖
    # ----------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_header(s5, "社區走勢：比佛利山莊 124 筆純住宅成交軌跡與大盤回修")

    # Left: Embed Chart
    if os.path.exists(CHART_BEVERLY):
        s5.shapes.add_picture(CHART_BEVERLY, PptxInches(0.8), PptxInches(1.5), PptxInches(6.8), PptxInches(5.4))

    # Right: Summary card
    add_card_shape(s5, PptxInches(7.9), PptxInches(1.5), PptxInches(4.6), PptxInches(5.4))
    tb_rs5 = s5.shapes.add_textbox(PptxInches(8.1), PptxInches(1.7), PptxInches(4.2), PptxInches(5.0))
    tf_rs5 = tb_rs5.text_frame
    tf_rs5.word_wrap = True

    p_rh = tf_rs5.paragraphs[0]
    p_rh.text = "【實登清洗與行情走勢重點】"
    p_rh.font.name = "微軟正黑體"
    p_rh.font.size = PptxPt(14)
    p_rh.font.bold = True
    p_rh.font.color.rgb = PPTX_NAVY

    points_s5 = [
        ("實登總量與清洗", "全社區 158 筆登錄，排除 1 樓店面 16 筆與親友特殊 20 筆，取得 124 筆純住宅真實走勢。"),
        ("三大發展週期", "2012-2019 基期盤整（均價 12~14 萬）；2020-2024 資金推升（衝上 27.8 萬）；2025-2026 大盤回修。"),
        ("大三房行情回落", "2026 年最新 3 樓大坪數戶別（含平面車位）以 800 萬成交，折合單價下探至 20.4 萬/坪。"),
        ("3 房常態中位數", "2024-2026 年社區 3 房格局扣除車位後淨單價中位數為 26.29 萬/坪（均價 26.41 萬/坪）。"),
        ("待售開價頂天", "待售戶淨單價開價 30.45 萬/坪，高出常態中位數達 +15.8%，明顯脫離 2026 年買方市場現況。")
    ]
    for pt_t, pt_d in points_s5:
        ppt = tf_rs5.add_paragraph()
        ppt.text = "▲ " + pt_t
        ppt.font.name = "微軟正黑體"
        ppt.font.size = PptxPt(11)
        ppt.font.bold = True
        ppt.font.color.rgb = PPTX_GOLD
        ppt.space_before = PptxPt(5)

        ppd = tf_rs5.add_paragraph()
        ppd.text = pt_d
        ppd.font.name = "微軟正黑體"
        ppd.font.size = PptxPt(9.5)
        ppd.font.color.rgb = PPTX_TEXT_DARK
        ppd.space_before = PptxPt(1)

    # ----------------------------------------------------
    # SLIDE 6: 車位拆算模型與敏感度分析
    # ----------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_header(s6, "估價模型：南區車位分離拆算與淨單價敏感度分析")

    # Left: Methodology card
    add_card_shape(s6, PptxInches(0.8), PptxInches(1.5), PptxInches(5.6), PptxInches(5.4))
    tb_ls6 = s6.shapes.add_textbox(PptxInches(1.1), PptxInches(1.7), PptxInches(5.0), PptxInches(5.0))
    tf_ls6 = tb_ls6.text_frame
    tf_ls6.word_wrap = True

    p_lm = tf_ls6.paragraphs[0]
    p_lm.text = "【車位價格分離標準公式】"
    p_lm.font.name = "微軟正黑體"
    p_lm.font.size = PptxPt(15)
    p_lm.font.bold = True
    p_lm.font.color.rgb = PPTX_NAVY

    p_f = tf_ls6.add_paragraph()
    p_f.text = "房屋淨建坪單價 = (成交總價 - 車位總價) ÷ (總坪數 - 車位坪數)\n= (總價 - 車位總價) ÷ 30.48 坪"
    p_f.font.name = "微軟正黑體"
    p_f.font.size = PptxPt(11)
    p_f.font.bold = True
    p_f.font.color.rgb = PPTX_GOLD
    p_f.space_before = PptxPt(6)
    p_f.space_after = PptxPt(8)

    car_benchmarks = [
        "南區大數據基準：機械車位 430 筆中位數 60 萬，主流成交 70~80 萬元；平面車位 813 筆中位數 109 萬，主流 110~130 萬。",
        "比佛利本社區實證：2021 年機械車位拆算 60~80 萬，2022 年機械車位拆算 100 萬，2024 年平面車位拆算 100 萬元。",
        "本案基準設定：坡道機械車位以 70 萬元為基準評估（敏感度區間 60~80 萬元）。",
        "還原真實開價：總價 998 萬扣除 70 萬車位後，純建坪淨開價單價為 30.45 萬/坪。",
        "還原買進單價：2024 年 870 萬扣除 70 萬車位後，純建坪買進淨單價為 26.25 萬/坪。"
    ]
    for cb in car_benchmarks:
        pcb = tf_ls6.add_paragraph()
        pcb.text = "• " + cb
        pcb.font.name = "微軟正黑體"
        pcb.font.size = PptxPt(10.5)
        pcb.font.color.rgb = PPTX_TEXT_DARK
        pcb.space_before = PptxPt(5)

    # Right: Sensitivity Cards
    right_cards = [
        ("保守情境（車位 60 萬元）", "房屋淨總價：938 萬元\n純建物淨單價：30.77 萬/坪\n相對 2024 年買進單價溢價：+4.52 萬 (+17.2%)", PPTX_RED),
        ("基準情境（車位 70 萬元）", "房屋淨總價：928 萬元\n純建物淨單價：30.45 萬/坪\n相對 2024 年買進單價溢價：+4.20 萬 (+16.0%)", PPTX_GOLD),
        ("寬鬆情境（車位 80 萬元）", "房屋淨總價：918 萬元\n純建物淨單價：30.12 萬/坪\n相對 2024 年買進單價溢價：+3.87 萬 (+14.7%)", PPTX_NAVY)
    ]
    for idx, (st, sd, sc) in enumerate(right_cards):
        top_y = PptxInches(1.5 + idx * 1.85)
        add_card_shape(s6, PptxInches(6.8), top_y, PptxInches(5.7), PptxInches(1.7))

        strip = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptxInches(6.8), top_y, PptxInches(0.12), PptxInches(1.7))
        set_shape_flat(strip, sc)

        tb_sc = s6.shapes.add_textbox(PptxInches(7.1), top_y + PptxInches(0.15), PptxInches(5.2), PptxInches(1.4))
        tf_sc = tb_sc.text_frame
        tf_sc.word_wrap = True

        p_sct = tf_sc.paragraphs[0]
        p_sct.text = st
        p_sct.font.name = "微軟正黑體"
        p_sct.font.size = PptxPt(13)
        p_sct.font.bold = True
        p_sct.font.color.rgb = sc

        p_scb = tf_sc.add_paragraph()
        p_scb.text = sd
        p_scb.font.name = "微軟正黑體"
        p_scb.font.size = PptxPt(10.5)
        p_scb.font.color.rgb = PPTX_TEXT_DARK
        p_scb.space_before = PptxPt(3)

    # ----------------------------------------------------
    # SLIDE 7: 南區 20+ 棟同型競品社區行情全景
    # ----------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_header(s7, "競品對比：台中市南區 20+ 棟老牌社區與次新標竿全景清單")

    # Table of 8 key peer communities
    table_shape7 = s7.shapes.add_table(9, 6, PptxInches(0.8), PptxInches(1.5), PptxInches(11.733), PptxInches(5.4))
    t7 = table_shape7.table
    t7.columns[0].width = PptxInches(1.8)
    t7.columns[1].width = PptxInches(2.5)
    t7.columns[2].width = PptxInches(0.9)
    t7.columns[3].width = PptxInches(1.2)
    t7.columns[4].width = PptxInches(1.2)
    t7.columns[5].width = PptxInches(4.133)

    h_titles7 = ["社區名稱", "門牌地址", "屋齡", "113年均價", "115年均價", "相對比價意涵與核心特徵"]
    for i, h in enumerate(h_titles7):
        c = t7.cell(0, i)
        c.fill.solid()
        c.fill.fore_color.rgb = PPTX_NAVY
        p = c.text_frame.paragraphs[0]
        p.text = h
        p.font.name = "微軟正黑體"
        p.font.size = PptxPt(10)
        p.font.bold = True
        p.font.color.rgb = PPTX_WHITE

    peers_display = [
        ("比佛利山莊 (標的)", "忠明南路730巷24號", "34年", "27.95萬", "28.25萬", "超低公設比 6.8%，室內 28.4 坪。13F 開價 998 萬（淨單價 30.45 萬）。"),
        ("佛羅里達", "忠明南路763號", "28年", "27.83萬", "25.90萬", "正對面直接鄰居。屋齡新 6 年，近年均價 27.9 萬/坪，115 年回修。"),
        ("家綠寶大廈", "忠明南路780號", "30年", "28.66萬", "26.00萬", "同路段鄰近社區。戶數較少，近年均價 27.45 萬/坪。"),
        ("歌德堡", "德富路96巷2號", "28年", "31.13萬", "26.65萬", "大慶雙鐵 405m。113 年衝上 31.1 萬後，115 年實質回修至 26.65 萬。"),
        ("公園苑", "德富路122號", "26年", "30.66萬", "—", "德富路生活圈指標。屋齡 26 年維護優良，近年均價 29.14 萬/坪。"),
        ("健康世界", "大慶街二段7巷4號", "33年", "26.29萬", "20.12萬", "中山醫商圈 33 年同齡老牌，115 年最新揭露均價下探至 20.12 萬。"),
        ("長億南園", "大慶街一段79號", "34年", "26.47萬", "22.60萬", "同為 34 年屋齡競品，近期均價回落至 22.60 萬/坪。"),
        ("宏都雙子星 (次新)", "復興北路767號", "12年", "38.00萬", "36.50萬", "十三期首排次新屋標竿天花板，均價 36~38 萬/坪。")
    ]
    for r_idx, (cn, ca, cage, cp13, cp15, cdesc) in enumerate(peers_display, start=1):
        row_vals = [cn, ca, cage, cp13, cp15, cdesc]
        is_target = "比佛利山莊" in cn
        for c_idx, val in enumerate(row_vals):
            c = t7.cell(r_idx, c_idx)
            c.fill.solid()
            c.fill.fore_color.rgb = PptxRGBColor(241, 245, 249) if is_target else (PPTX_WHITE if r_idx % 2 == 0 else PPTX_LIGHT_BG)
            p = c.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "微軟正黑體"
            p.font.size = PptxPt(9)
            p.font.bold = (is_target or c_idx == 0)
            p.font.color.rgb = PPTX_RED if is_target and c_idx == 0 else PPTX_TEXT_DARK

    # ----------------------------------------------------
    # SLIDE 8: 五大特徵調整矩陣
    # ----------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_header(s8, "比價調整：五大特徵維度價值修正矩陣 (Hedonic Adjustment)")

    cards_s8 = [
        ("樓層景觀 (+5.0%)", "13F/14F 黃金次頂樓視野", "視野開闊無遮蔽，無頂樓日曬與漏水風險，具備良好採光，相對中低樓層享有 +5% 溢價溢幅。"),
        ("採光格局 (+3.0%)", "邊間三面採光、雙衛皆開窗", "罕見雙衛浴皆有對外窗，通風防潮防霉，邊間通風極佳，為老牌社區極稀缺之優質格局。"),
        ("公設坪效 (+3.0%)", "超低公設比 6.76%、實坪 28.4 坪", "得房率高達 93.2%，實質可用空間達 28.4 坪，無虛坪雨遮灌水，實用坪效享有 +3% 修正。"),
        ("建築屋齡 (-2.0%)", "法定屋齡 34 年 (1992 完工)", "跨過 30 年大關，水電管線已超過耐用年限，銀行貸款審核嚴格，予以 -2% 成本折減。"),
        ("車位型態 (-40 萬元)", "坡道機械車位 (B1-89)", "受限於升降保養成本與休旅車載重尺寸限制，相對平面車位（110 萬）實質折減 40 萬元。")
    ]
    for idx, (atitle, afeat, aexp) in enumerate(cards_s8):
        y = PptxInches(1.5 + idx * 1.1)
        add_card_shape(s8, PptxInches(0.8), y, PptxInches(11.733), PptxInches(0.95))

        badge = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(1.1), y + PptxInches(0.15), PptxInches(2.5), PptxInches(0.65))
        is_pos = "+" in atitle
        set_shape_flat(badge, PPTX_NAVY if is_pos else PPTX_RED)
        tb_bg = badge.text_frame
        p_bg = tb_bg.paragraphs[0]
        p_bg.text = atitle
        p_bg.font.name = "微軟正黑體"
        p_bg.font.size = PptxPt(11)
        p_bg.font.bold = True
        p_bg.font.color.rgb = PPTX_WHITE
        p_bg.alignment = PP_ALIGN.CENTER

        tb_tx = s8.shapes.add_textbox(PptxInches(3.8), y + PptxInches(0.1), PptxInches(8.4), PptxInches(0.75))
        tf_tx = tb_tx.text_frame
        tf_tx.word_wrap = True
        p_th = tf_tx.paragraphs[0]
        p_th.text = afeat
        p_th.font.name = "微軟正黑體"
        p_th.font.size = PptxPt(11)
        p_th.font.bold = True
        p_th.font.color.rgb = PPTX_NAVY

        p_td = tf_tx.add_paragraph()
        p_td.text = aexp
        p_td.font.name = "微軟正黑體"
        p_td.font.size = PptxPt(9.5)
        p_td.font.color.rgb = PPTX_TEXT_DARK
        p_td.space_before = PptxPt(2)

    # ----------------------------------------------------
    # SLIDE 9: 三段式出價策略矩陣
    # ----------------------------------------------------
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_header(s9, "出價防線：三段式出價梯隊與具體定錨區間")

    tiers = [
        ("第一道防線：保守底標價", f"{vm['conservative']['total']:.0f} 萬元", f"淨單價 {vm['conservative']['net_unit_price']:.2f} 萬/坪", vm['conservative']['discount_str'], f"-{tp['asking_price'] - vm['conservative']['total']:.0f} 萬元", [
            "大膽試探防線，建議區間 800 ~ 820 萬",
            "錨定 2026 年比佛利 3 樓平車 800 萬 (20.4萬/坪) 回修事實",
            "提出屋齡 34 年老態需預留 80 萬水電泥作翻修預算",
            "藉此測試屋主資金壓力與讓利誠意"
        ], PPTX_GREEN),
        ("第二道防線：合理目標價", f"{vm['target']['total']:.0f} 萬元", f"淨單價 {vm['target']['net_unit_price']:.2f} 萬/坪", vm['target']['discount_str'], f"-{tp['asking_price'] - vm['target']['total']:.0f} 萬元", [
            "黃金成交甜蜜點，建議區間 860 ~ 880 萬",
            "精準錨定屋主 2024-01-24 買進成本 870 萬元",
            "給予屋主平盤保本出場，免受老屋續跌虧損",
            "完全符合社區 3 房近年拆算淨單價中位數 26.29 萬/坪"
        ], PPTX_GOLD),
        ("第三道防線：上限天花板價", f"{vm['ceiling']['total']:.0f} 萬元", f"淨單價 {vm['ceiling']['net_unit_price']:.2f} 萬/坪", vm['ceiling']['discount_str'], f"-{tp['asking_price'] - vm['ceiling']['total']:.0f} 萬元", [
            "談判停損警戒線，建議區間 910 ~ 920 萬",
            "買方自住極限紅線，對齊 2024 多頭年均價 27.81 萬",
            "若屋主堅持突破 920 萬，強烈建議放棄談判",
            "轉向南區附平面車位或更年輕競品（如公園苑、椰城）"
        ], PPTX_RED)
    ]
    for idx, (tname, ttotal, tunit, tdisc, tcut, tpoints, tcolor) in enumerate(tiers):
        l = PptxInches(0.8 + idx * 4.0)
        t = PptxInches(1.5)
        add_card_shape(s9, l, t, PptxInches(3.75), PptxInches(5.4))

        top_strip = s9.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, PptxInches(3.75), PptxInches(0.12))
        set_shape_flat(top_strip, tcolor)

        tb = s9.shapes.add_textbox(l + PptxInches(0.2), t + PptxInches(0.2), PptxInches(3.35), PptxInches(5.0))
        tf = tb.text_frame
        tf.word_wrap = True

        p_tn = tf.paragraphs[0]
        p_tn.text = tname
        p_tn.font.name = "微軟正黑體"
        p_tn.font.size = PptxPt(13)
        p_tn.font.bold = True
        p_tn.font.color.rgb = tcolor

        p_tot = tf.add_paragraph()
        p_tot.text = ttotal
        p_tot.font.name = "微軟正黑體"
        p_tot.font.size = PptxPt(26)
        p_tot.font.bold = True
        p_tot.font.color.rgb = PPTX_NAVY
        p_tot.space_before = PptxPt(4)

        p_det = tf.add_paragraph()
        p_det.text = f"{tunit} ｜ {tdisc} ｜ 折減 {tcut}"
        p_det.font.name = "微軟正黑體"
        p_det.font.size = PptxPt(10)
        p_det.font.color.rgb = PPTX_TEXT_MUTED
        p_det.space_before = PptxPt(2)
        p_det.space_after = PptxPt(8)

        for tp in tpoints:
            ptp = tf.add_paragraph()
            ptp.text = "• " + tp
            ptp.font.name = "微軟正黑體"
            ptp.font.size = PptxPt(10.5)
            ptp.font.color.rgb = PPTX_TEXT_DARK
            ptp.space_before = PptxPt(4)

    # ----------------------------------------------------
    # SLIDE 10: 談判話術與經紀人攻防戰術
    # ----------------------------------------------------
    s10 = prs.slides.add_slide(blank_layout)
    add_slide_header(s10, "實戰話術：房仲溝通逐字稿與談判攻防戰術")

    tactics = [
        ("話術 1：鎖定 870 萬成本，主打「平盤保本」出場", 
         "「我們很清楚屋主在 113 年 1 月以 870 萬買進。大盤已經反轉，連同社區 3 樓今年都跌到 20.4 萬/坪。我們自住誠意買方出 870 萬讓屋主完全平盤保本出場，把資金安全落袋；再拖延下去恐怕連成本都守不住。」", PPTX_GOLD),
        ("話術 2：一般委託聯賣，發揮雙仲介比價制衡", 
         "「這間在中信跟永義都有一般委託聯賣。哪一家能幫我們促成 870 萬成交、服務費願意折讓至 1%，我們就立刻下斡旋簽約。請經紀人全力協助回報屋主。」", PPTX_NAVY),
        ("話術 3：屋齡 34 年管線折舊，需扣減 80 萬裝修預算", 
         "「社區已屆滿 34 年，水電管線與衛浴早已過使用年限，買方必須重新施作全室配線、給排水與泥作防水，這筆 80 萬基礎工程成本必須具體反映在總價折減上。」", PPTX_SLATE),
        ("話術 4：機械車位抗性，反映 40 萬升降保養折價", 
         "「B1-89 為機械車位，無法停放休旅車且每月需繳清潔費，相較南區平面車位 110 萬行情，機械車位公允市價僅 70 萬，請屋主務必扣除 40 萬價差。」", PPTX_RED)
    ]
    for idx, (ttitle, tquote, tcolor) in enumerate(tactics):
        y = PptxInches(1.5 + idx * 1.35)
        add_card_shape(s10, PptxInches(0.8), y, PptxInches(11.733), PptxInches(1.2))

        bar = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptxInches(0.8), y, PptxInches(0.12), PptxInches(1.2))
        set_shape_flat(bar, tcolor)

        tb = s10.shapes.add_textbox(PptxInches(1.1), y + PptxInches(0.1), PptxInches(11.2), PptxInches(1.0))
        tf = tb.text_frame
        tf.word_wrap = True

        pt = tf.paragraphs[0]
        pt.text = ttitle
        pt.font.name = "微軟正黑體"
        pt.font.size = PptxPt(12.5)
        pt.font.bold = True
        pt.font.color.rgb = tcolor

        pq = tf.add_paragraph()
        pq.text = tquote
        pq.font.name = "微軟正黑體"
        pq.font.size = PptxPt(10.5)
        pq.font.color.rgb = PPTX_TEXT_DARK
        pq.space_before = PptxPt(3)

    # ----------------------------------------------------
    # SLIDE 11: 看屋實戰清單與決策結論
    # ----------------------------------------------------
    s11 = prs.slides.add_slide(blank_layout)
    add_slide_header(s11, "看屋清單與置產決策結論：以退為進、精準定錨")

    # Left: Checklist
    add_card_shape(s11, PptxInches(0.8), PptxInches(1.5), PptxInches(6.0), PptxInches(5.4))
    tb_ls11 = s11.shapes.add_textbox(PptxInches(1.1), PptxInches(1.7), PptxInches(5.4), PptxInches(5.0))
    tf_ls11 = tb_ls11.text_frame
    tf_ls11.word_wrap = True

    p_chk = tf_ls11.paragraphs[0]
    p_chk.text = "【實地看屋 5 大實戰查驗要項】"
    p_chk.font.name = "微軟正黑體"
    p_chk.font.size = PptxPt(14)
    p_chk.font.bold = True
    p_chk.font.color.rgb = PPTX_NAVY

    checks = [
        ("外牆窗框與滲漏水", "檢查東南向迎風面窗框周邊、冷氣孔有無壁癌補土痕跡或水痕。"),
        ("雙衛浴天花板管道間", "推開雙衛天花板維修孔，檢查上方給排水管與排氣管道間有無下層漏水。"),
        ("總開關箱安培數", "檢查總開關安培數是否為早期 30A~50A，確認重新拉線至 75A 之預算。"),
        ("B1-89 機械車位操作", "實地操作升降車台，測量長寬高限值與平穩度，確認家電車輛停放順暢度。"),
        ("社區公共基金現況", "調閱管委會公共修繕基金餘額，並至 14 樓頂樓檢視防水層維護品質。")
    ]
    for ctitle, cdesc in checks:
        pct = tf_ls11.add_paragraph()
        pct.text = "✔ " + ctitle
        pct.font.name = "微軟正黑體"
        pct.font.size = PptxPt(11)
        pct.font.bold = True
        pct.font.color.rgb = PPTX_GOLD
        pct.space_before = PptxPt(5)

        pcd = tf_ls11.add_paragraph()
        pcd.text = cdesc
        pcd.font.name = "微軟正黑體"
        pcd.font.size = PptxPt(9.5)
        pcd.font.color.rgb = PPTX_TEXT_DARK
        pcd.space_before = PptxPt(1)

    # Right: Final Conclusion
    add_card_shape(s11, PptxInches(7.1), PptxInches(1.5), PptxInches(5.4), PptxInches(5.4))
    tb_rs11 = s11.shapes.add_textbox(PptxInches(7.4), PptxInches(1.7), PptxInches(4.8), PptxInches(5.0))
    tf_rs11 = tb_rs11.text_frame
    tf_rs11.word_wrap = True

    p_fc = tf_rs11.paragraphs[0]
    p_fc.text = "【置產決策綜合評定】"
    p_fc.font.name = "微軟正黑體"
    p_fc.font.size = PptxPt(14)
    p_fc.font.bold = True
    p_fc.font.color.rgb = PPTX_NAVY

    conclusions = [
        ("物件整體優勢", "次頂樓開闊視野、邊間三面採光、雙衛皆開窗、超低公設比 6.76%、實坪 28.4 坪，自住空間極度舒適。"),
        ("主要抗性考量", "屋齡已屆滿 34 年老宅門檻，水電需重施作；車位為機械式非平面車位，需扣減 40 萬車位價值。"),
        ("核心出價建議", "強烈建議以 870 萬元（屋主 2024 年取得成本）為核心目標價。此價格既符合社區 3 房行情，又給予屋主平盤保本出場。"),
        ("停損警戒防線", "上限天花板為 920 萬元。若屋主強硬堅持超過 920 萬，建議買方果斷放棄，轉向周邊具平面車位或更年輕之競品。")
    ]
    for ctitle, cdesc in conclusions:
        pct = tf_rs11.add_paragraph()
        pct.text = "★ " + ctitle
        pct.font.name = "微軟正黑體"
        pct.font.size = PptxPt(11)
        pct.font.bold = True
        pct.font.color.rgb = PPTX_GREEN if "出價" in ctitle else PPTX_NAVY
        pct.space_before = PptxPt(5)

        pcd = tf_rs11.add_paragraph()
        pcd.text = cdesc
        pcd.font.name = "微軟正黑體"
        pcd.font.size = PptxPt(10)
        pcd.font.color.rgb = PPTX_TEXT_DARK
        pcd.space_before = PptxPt(1)

    prs.save(PPTX_OUTPUT_PATH)
    print(f"Generated PowerPoint (.pptx): {PPTX_OUTPUT_PATH} ({os.path.getsize(PPTX_OUTPUT_PATH)} bytes)")


# ==============================================================================
# MAIN EXECUTION
# ==============================================================================
def main():
    print("=" * 60)
    print("Generating Beverly Hills 13F Deliverables Suite...")
    print("=" * 60)

    data = load_master_data()
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    generate_markdown(data)
    generate_word_document(data)
    generate_powerpoint_presentation(data)

    print("=" * 60)
    print("All deliverables generated successfully!")
    print("=" * 60)


if __name__ == "__main__":
    main()

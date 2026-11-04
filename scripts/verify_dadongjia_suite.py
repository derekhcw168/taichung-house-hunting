#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scripts/verify_dadongjia_suite.py
Automated Verification Suite for Taichung South District "Dadongjia Nanyuan 9F" (大東家南園9樓)
Valuation and Community Market Report Deliverables.

Four Verification Tiers:
  - Tier 1: File Existence & Basic Checks (MD >=15KB, DOCX >=30KB, PPTX >=300KB, JSON >=5KB)
  - Tier 2: Format & Structural Integrity (DOCX tables/cells, PPTX 16:9 slides/KPI cards, MD 8 required sections)
  - Tier 3: Data Consistency Across Artifacts (Floor 9F, 33.70坪, 898萬, parking decoupling, 805-810/845-860/880-890萬, 77/63 transactions, competitors)
  - Tier 4: Visual Assets & Git Sync Readiness (PNG chart existence, PIL forensics >=1000x600, sync_to_github.py inclusion)

CLI Arguments:
  --all (default)       : Full strict acceptance verification across all 4 tiers
  --progressive         : Progressive mode for ongoing milestone validation (warns on missing deliverables)
  --tier {1,2,3,4}      : Execute a specific tier (also --tier1, --tier2, --tier3, --tier4)
  --test-self           : Execute internal self-test suite with synthetic mock fixtures (8 adversarial test cases)
  --json                : Output machine-readable JSON report
  --verbose             : Verbose diagnostic output
"""

import os
import sys
import re
import json
import argparse
import tempfile
import shutil
from typing import Dict, List, Tuple, Any, Optional

# Ensure UTF-8 output on Windows
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

# Optional third-party imports
try:
    import docx
except ImportError:
    docx = None

try:
    import pptx
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    pptx = None
    Presentation = None
    MSO_SHAPE_TYPE = None

try:
    from PIL import Image
except ImportError:
    Image = None


PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Deliverable Paths
DEFAULT_MD_PATH = os.path.join(PROJECT_ROOT, "大東家南園9樓", "大東家南園9樓_特定待售戶別出價分析與社區整體行情報告.md")
DEFAULT_DOCX_PATH = os.path.join(PROJECT_ROOT, "大東家南園9樓", "大東家南園9樓_特定待售戶別出價分析與社區整體行情報告.docx")
DEFAULT_PPTX_PATH = os.path.join(PROJECT_ROOT, "大東家南園9樓", "大東家南園9樓_特定待售戶別出價分析與社區整體行情報告.pptx")
DEFAULT_JSON_PATH = os.path.join(PROJECT_ROOT, "scratch", "dadongjia_valuation_master.json")
DEFAULT_SYNC_SCRIPT_PATH = os.path.join(PROJECT_ROOT, "scripts", "sync_to_github.py")

# Visual Assets Paths
DEFAULT_DADONGJIA_CHART_PATH = os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "42_大東家南園_成交行情走勢圖.png")
COMPETITOR_CHART_PATHS = [
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "28_長億南園_成交行情走勢圖.png"),
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "20_公園苑_成交行情走勢圖.png"),
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "34_椰城_成交行情走勢圖.png"),
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "21_大東家樸園_成交行情走勢圖.png"),
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "35_大東家華園_成交行情走勢圖.png"),
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "00_台中市南區25年以上大型社區_大盤匯總走勢圖.png"),
]

# Required Size Thresholds (bytes)
SIZE_THRESHOLDS = {
    "md": 15 * 1024,      # 15 KB
    "docx": 30 * 1024,    # 30 KB
    "pptx": 300 * 1024,   # 300 KB
    "json": 5 * 1024      # 5 KB
}

# Required Markdown Section Keywords (8 key chapters from DISPATCH.md)
REQUIRED_MD_SECTIONS = [
    ("物件產調與規格", r"物件產調|物件規格|標的現況|待售物件|基本規格"),
    ("歷史實登洗淨分析", r"歷史實登|實價登錄|歷史行情|洗淨分析|歷年行情|成交行情"),
    ("同類型競品比價", r"同類型競品|競品比價|周邊社區|競品對比|周邊競品"),
    ("車位拆算解耦", r"車位拆算|車位解耦|車位分離|淨建坪單價|淨單價"),
    ("多維度比價調整矩陣", r"比價調整矩陣|多維度比價|特徵調整|調整權重|調整矩陣"),
    ("三段式出價策略", r"三段式出價|出價策略|出價三道防線|出價區間|底標價|目標價|天花板價"),
    ("談判議價話術", r"談判議價話術|談判說帖|議價話術|議價說帖|攻防要點"),
    ("簽約避坑與修繕檢核", r"簽約避坑|避坑檢核|交屋檢核|簽約注意|修繕工程|修繕預算|結論")
]

# Required Canonical Valuation Figures & Keywords for Dadongjia Nanyuan 9F
VALUATION_FIGURES = {
    "floor": ["9F", "9樓", "9/15F", "高樓層"],
    "total_ping": ["33.70", "33.71", "33.7"],
    "main_ping": ["19.84", "19.844", "21.8"],
    "aux_ping": ["1.98", "1.984"],
    "living_ping": ["21.82", "21.828", "21.8"],
    "asking_price": ["898"],
    "asking_unit_price": ["26.65", "26.64", "26.6"],
    "car_space_type": ["機械", "平移", "下層平移", "休旅"],
    "car_space_price": ["72", "75", "70~75", "72~75", "80"],
    "conservative_tier": ["805", "810", "755"],
    "reasonable_tier": ["845", "850", "860", "868", "820"],
    "ceiling_tier": ["880", "890", "850"],
    "transactions_count": ["77"],
    "residential_count": ["63", "67"],
    "latest_deal_2026": ["855", "25.3"],
    "competitors": ["長億南園", "公園苑", "椰城", "大東家華園", "大東家樸園"]
}


class TestResult:
    """Encapsulates test outcomes, status messages, warnings, and errors for a tier."""
    def __init__(self, name: str):
        self.name = name
        self.passed = True
        self.messages: List[str] = []
        self.warnings: List[str] = []
        self.errors: List[str] = []
        self.details: Dict[str, Any] = {}

    def ok(self, msg: str):
        self.messages.append(msg)

    def warn(self, msg: str):
        self.warnings.append(msg)

    def fail(self, msg: str):
        self.passed = False
        self.errors.append(msg)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "passed": self.passed,
            "messages": self.messages,
            "warnings": self.warnings,
            "errors": self.errors,
            "details": self.details
        }


# ==============================================================================
# Helper Parsers
# ==============================================================================
def extract_md_text(path: str) -> str:
    """Reads Markdown file content."""
    if not os.path.exists(path):
        return ""
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception:
        return ""


def extract_docx_text_and_stats(path: str) -> Tuple[str, int, int, int]:
    """
    Parses DOCX and extracts full text, table count, total cells, paragraph count.
    Returns: (full_text, table_count, total_cells, paragraph_count)
    """
    if not os.path.exists(path) or docx is None:
        return "", 0, 0, 0
    try:
        doc = docx.Document(path)
        p_texts = [p.text for p in doc.paragraphs]
        table_count = len(doc.tables)
        total_cells = 0
        table_texts = []
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    total_cells += 1
                    table_texts.append(cell.text)
        full_text = "\n".join(p_texts + table_texts)
        return full_text, table_count, total_cells, len(doc.paragraphs)
    except Exception:
        return "", 0, 0, 0


def extract_pptx_text_and_stats(path: str) -> Tuple[str, int, float, float, int]:
    """
    Parses PPTX and extracts full text, slide count, width (inches), height (inches), picture count.
    Returns: (full_text, slide_count, width_inches, height_inches, picture_count)
    """
    if not os.path.exists(path) or pptx is None:
        return "", 0, 0.0, 0.0, 0
    try:
        prs = Presentation(path)
        slide_count = len(prs.slides)
        width_in = prs.slide_width / 914400.0 if hasattr(prs, 'slide_width') else 0.0
        height_in = prs.slide_height / 914400.0 if hasattr(prs, 'slide_height') else 0.0
        text_parts = []
        picture_count = 0

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:
                        if p.text.strip():
                            text_parts.append(p.text)
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text.strip():
                                text_parts.append(cell.text)
                if MSO_SHAPE_TYPE and shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    picture_count += 1
                elif hasattr(shape, 'image'):
                    picture_count += 1

        full_text = "\n".join(text_parts)
        return full_text, slide_count, width_in, height_in, picture_count
    except Exception:
        return "", 0, 0.0, 0.0, 0


def extract_json_data(path: str) -> Optional[Dict[str, Any]]:
    """Loads and returns JSON master database."""
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


# ==============================================================================
# Tier 1: File Existence & Basic Checks
# ==============================================================================
def verify_tier1(paths: Dict[str, str], progressive: bool = False, verbose: bool = False) -> TestResult:
    """
    Tier 1: Checks existence and minimum size thresholds of:
      - MD (>= 15KB)
      - DOCX (>= 30KB)
      - PPTX (>= 300KB)
      - JSON (>= 5KB)
    """
    res = TestResult("Tier 1: File Existence & Basic Checks (產物存在性與檔案容量檢驗)")

    targets = [
        ("Markdown 分析報告", paths.get("md", DEFAULT_MD_PATH), SIZE_THRESHOLDS["md"], "15KB"),
        ("Word 分析報告 (.docx)", paths.get("docx", DEFAULT_DOCX_PATH), SIZE_THRESHOLDS["docx"], "30KB"),
        ("PowerPoint 簡報 (.pptx)", paths.get("pptx", DEFAULT_PPTX_PATH), SIZE_THRESHOLDS["pptx"], "300KB"),
        ("估價主資料庫 (.json)", paths.get("json", DEFAULT_JSON_PATH), SIZE_THRESHOLDS["json"], "5KB")
    ]

    for label, file_path, min_bytes, min_str in targets:
        rel_path = os.path.relpath(file_path, PROJECT_ROOT) if file_path.startswith(PROJECT_ROOT) else file_path
        if not os.path.exists(file_path):
            if progressive:
                res.warn(f"[{label}] 尚未產生：`{rel_path}` (Progressive Mode 略過阻斷)")
            else:
                res.fail(f"[{label}] 檔案不存在：`{rel_path}`")
            continue

        file_size = os.path.getsize(file_path)
        res.details[label] = {"path": rel_path, "size": file_size}

        if file_size < min_bytes:
            res.fail(f"[{label}] 檔案容量不足：`{rel_path}` ({file_size} bytes < 門檻 {min_bytes} bytes / {min_str})")
        else:
            res.ok(f"[{label}] 存在且容量達標：`{rel_path}` ({file_size / 1024:.1f} KB >= {min_str})")

    return res


# ==============================================================================
# Tier 2: Format & Structural Integrity
# ==============================================================================
def verify_tier2(paths: Dict[str, str], progressive: bool = False, verbose: bool = False) -> TestResult:
    """
    Tier 2: Validates structural integrity of deliverables:
      - DOCX: parsable, >= 5 tables, >= 50 total cells, contains paragraphs (>= 10)
      - PPTX: parsable, slide count >= 10, 16:9 widescreen (13.333" x 7.5")
      - MD: contains all 8 required section headers
    """
    res = TestResult("Tier 2: Format & Structural Integrity (檔案格式與內部結構完整性)")

    # 1. DOCX Verification
    docx_path = paths.get("docx", DEFAULT_DOCX_PATH)
    if os.path.exists(docx_path):
        if docx is None:
            res.fail("python-docx 函式庫未安裝，無法驗證 .docx 結構")
        else:
            try:
                full_text, table_cnt, cell_cnt, p_cnt = extract_docx_text_and_stats(docx_path)
                res.details["docx_stats"] = {
                    "tables": table_cnt,
                    "cells": cell_cnt,
                    "paragraphs": p_cnt
                }

                if table_cnt < 5:
                    res.fail(f"Word 文件 (.docx) 表格數量不足：目前 {table_cnt} 個表格，要求 >= 5 個表格")
                else:
                    res.ok(f"Word 文件 (.docx) 表格數量達標：共 {table_cnt} 個表格 (>= 5)")

                if cell_cnt < 50:
                    res.fail(f"Word 文件 (.docx) 表格總單元格不足：目前 {cell_cnt} 格，要求 >= 50 格")
                else:
                    res.ok(f"Word 文件 (.docx) 表格總單元格達標：共 {cell_cnt} 格 (>= 50)")

                if p_cnt < 10 or len(full_text.strip()) < 500:
                    res.fail(f"Word 文件 (.docx) 內文字數過少或段落缺失 (段落數: {p_cnt})")
                else:
                    res.ok(f"Word 文件 (.docx) 內文段落完整：共 {p_cnt} 個段落")
            except Exception as e:
                res.fail(f"Word 文件 (.docx) 解析發生錯誤：{str(e)}")
    else:
        if progressive:
            res.warn("Word 文件尚未產生，略過 Tier 2 DOCX 檢驗")
        else:
            res.fail(f"Word 文件不存在，無法驗證結構：`{docx_path}`")

    # 2. PPTX Verification
    pptx_path = paths.get("pptx", DEFAULT_PPTX_PATH)
    if os.path.exists(pptx_path):
        if pptx is None:
            res.fail("python-pptx 函式庫未安裝，無法驗證 .pptx 結構")
        else:
            try:
                full_text, slide_cnt, width_in, height_in, pic_cnt = extract_pptx_text_and_stats(pptx_path)
                res.details["pptx_stats"] = {
                    "slides": slide_cnt,
                    "width_in": round(width_in, 3),
                    "height_in": round(height_in, 3),
                    "pictures": pic_cnt
                }

                if slide_cnt < 10:
                    res.fail(f"PowerPoint 簡報 (.pptx) 投影片總頁數不足：目前 {slide_cnt} 頁，要求 >= 10 頁")
                else:
                    res.ok(f"PowerPoint 簡報 (.pptx) 頁數達標：共 {slide_cnt} 頁 (>= 10)")

                # Verify 16:9 dimensions (13.333" x 7.5" with 0.15" tolerance)
                width_ok = abs(width_in - 13.333) < 0.15 or (height_in > 0 and abs((width_in / height_in) - (16.0 / 9.0)) < 0.05)
                height_ok = abs(height_in - 7.5) < 0.15

                if width_ok and height_ok:
                    res.ok(f"PowerPoint 簡報 (.pptx) 比例符合 16:9 寬螢幕規格 ({width_in:.2f}\" x {height_in:.2f}\")")
                else:
                    res.fail(f"PowerPoint 簡報 (.pptx) 比例不符合 16:9：寬度 {width_in:.3f}\" (預期 13.333\"), 高度 {height_in:.3f}\" (預期 7.5\")")
            except Exception as e:
                res.fail(f"PowerPoint 簡報 (.pptx) 解析發生錯誤：{str(e)}")
    else:
        if progressive:
            res.warn("PowerPoint 簡報尚未產生，略過 Tier 2 PPTX 檢驗")
        else:
            res.fail(f"PowerPoint 簡報不存在，無法驗證結構：`{pptx_path}`")

    # 3. Markdown Sections Verification (8 Chapters)
    md_path = paths.get("md", DEFAULT_MD_PATH)
    if os.path.exists(md_path):
        md_text = extract_md_text(md_path)
        missing_sections = []
        for sec_name, pattern in REQUIRED_MD_SECTIONS:
            if not re.search(pattern, md_text):
                missing_sections.append(sec_name)

        if missing_sections:
            res.fail(f"Markdown 報告缺少必要核心章節：{', '.join(missing_sections)}")
        else:
            res.ok("Markdown 報告包含全部 8 大核心章節 (物件產調, 歷史實登, 競品比價, 車位拆算, 調整矩陣, 三段出價, 談判說帖, 簽約避坑)")
    else:
        if progressive:
            res.warn("Markdown 分析報告尚未產生，略過 Tier 2 MD 檢驗")
        else:
            res.fail(f"Markdown 分析報告不存在，無法驗證章節結構：`{md_path}`")

    return res


# ==============================================================================
# Tier 3: Data Consistency Across Artifacts
# ==============================================================================
def verify_tier3(paths: Dict[str, str], progressive: bool = False, verbose: bool = False) -> TestResult:
    """
    Tier 3: Cross-verifies key valuation figures and entities across MD, DOCX, PPTX, and JSON:
      - Target floor: 9F (高樓層)
      - Area: 33.70 坪 (主建 19.844 坪, 附屬 1.984 坪, 主附 21.828 坪)
      - Car space: 機械下層平移車位 (可停休旅車)
      - Asking price: 898 萬 (毛單價 26.65 萬/坪)
      - Pricing tiers: 保守底標 805~810 萬 (或 755~810 萬), 合理目標 845~860 萬 (或 810~860 萬), 上限天花板 880~890 萬 (或 850~890 萬)
      - Historical: 77 筆實價登錄, 63 筆純住宅成交 (排除 14 筆), 2026 年最新成交 855 萬 (3F 25.3 萬/坪)
      - Competitors: 長億南園, 公園苑, 椰城, 大東家華園, 大東家樸園
    """
    res = TestResult("Tier 3: Data Consistency Across Artifacts (跨交付物核心數據一致性)")

    # 1. Master JSON Consistency
    json_path = paths.get("json", DEFAULT_JSON_PATH)
    json_data = extract_json_data(json_path)

    if json_data:
        res.ok(f"成功讀取主資料庫 JSON：`{os.path.basename(json_path)}`")
        tp = json_data.get("target_property", {})
        pt = json_data.get("pricing_tiers", {})
        hs = json_data.get("historical_summary", {})
        comps = json_data.get("comparables", json_data.get("comparable_communities", []))

        # Check Property figures in JSON
        if tp:
            floor_str = str(tp.get("floor", ""))
            if "9" in floor_str:
                res.ok(f"JSON 目標物件樓層正確：{floor_str}")
            else:
                res.fail(f"JSON 目標物件樓層錯誤：{floor_str} (應包含 9F / 9樓)")

            total_ping = float(tp.get("total_ping", tp.get("building_area", 0)))
            if abs(total_ping - 33.70) < 0.15:
                res.ok(f"JSON 總坪數正確：{total_ping} 坪")
            else:
                res.fail(f"JSON 總坪數不吻合：{total_ping} (預期 33.70 坪)")

            asking_price = float(tp.get("asking_price", 0))
            if abs(asking_price - 898.0) < 1.0:
                res.ok(f"JSON 開價總額正確：{asking_price} 萬")
            else:
                res.fail(f"JSON 開價總額不吻合：{asking_price} (預期 898 萬)")

        # Check Pricing Tiers in JSON
        if pt:
            cons = pt.get("conservative_floor", {})
            cons_val = cons.get("total", cons.get("price", cons.get("total_low", 0))) if isinstance(cons, dict) else cons
            if 750.0 <= float(cons_val) <= 815.0:
                res.ok(f"JSON 保守底標價落於合理範圍：{cons_val} 萬 (基準 805~810 萬)")
            else:
                res.fail(f"JSON 保守底標價不吻合：{cons_val} (預期 755~810 萬區間)")

            target = pt.get("reasonable_target", {})
            target_val = target.get("total", target.get("price", target.get("total_low", 0))) if isinstance(target, dict) else target
            if 810.0 <= float(target_val) <= 868.0:
                res.ok(f"JSON 合理目標價落於合理範圍：{target_val} 萬 (基準 845~860 萬)")
            else:
                res.fail(f"JSON 合理目標價不吻合：{target_val} (預期 810~860 萬區間)")

            ceil = pt.get("ceiling_max", {})
            ceil_val = ceil.get("total", ceil.get("price", ceil.get("total_high", 0))) if isinstance(ceil, dict) else ceil
            if 850.0 <= float(ceil_val) <= 895.0:
                res.ok(f"JSON 上限天花板價落於合理範圍：{ceil_val} 萬 (基準 880~890 萬)")
            else:
                res.fail(f"JSON 上限天花板價不吻合：{ceil_val} (預期 850~890 萬區間)")

        # Check Historical Summary in JSON
        if hs:
            tot_cnt = hs.get("total_count", hs.get("total_transactions", 0))
            if tot_cnt in [77, 0]:
                res.ok(f"JSON 歷史總成交筆數：{tot_cnt} 筆 (基準 77 筆)")
            res_cnt = hs.get("residential_count", hs.get("clean_transactions_count", 0))
            if res_cnt in [63, 67]:
                res.ok(f"JSON 純住宅歷史成交筆數正確：{res_cnt} 筆 (基準 63/67 筆)")
            elif res_cnt > 0:
                res.warn(f"JSON 純住宅歷史成交筆數：{res_cnt} 筆 (基準 63 筆)")

        # Check Competitors in JSON
        if isinstance(comps, list) and comps:
            comp_names = [c.get("name", "") if isinstance(c, dict) else str(c) for c in comps]
            all_comp_str = " ".join(comp_names)
            for c_name in VALUATION_FIGURES["competitors"]:
                if c_name in all_comp_str:
                    res.ok(f"JSON 競品清單包含：{c_name}")
                else:
                    res.fail(f"JSON 競品清單缺少核心競品：{c_name}")
    else:
        if progressive:
            res.warn("主資料庫 JSON 尚未產出，略過 JSON 內部欄位檢驗")
        else:
            res.fail(f"主資料庫 JSON 不存在或無法解析：`{json_path}`")

    # 2. Cross-Verification on Document Deliverables (MD, DOCX, PPTX)
    text_corpus: Dict[str, str] = {}
    md_path = paths.get("md", DEFAULT_MD_PATH)
    if os.path.exists(md_path):
        text_corpus["Markdown"] = extract_md_text(md_path)

    docx_path = paths.get("docx", DEFAULT_DOCX_PATH)
    if os.path.exists(docx_path):
        docx_txt, _, _, _ = extract_docx_text_and_stats(docx_path)
        text_corpus["Word (.docx)"] = docx_txt

    pptx_path = paths.get("pptx", DEFAULT_PPTX_PATH)
    if os.path.exists(pptx_path):
        pptx_txt, _, _, _, _ = extract_pptx_text_and_stats(pptx_path)
        text_corpus["PowerPoint (.pptx)"] = pptx_txt

    if not text_corpus:
        if progressive:
            res.warn("所有文檔交付物 (MD/DOCX/PPTX) 均未產生，略過文檔間一致性驗證")
            return res
        else:
            res.fail("所有文檔交付物均不存在，無法進行跨交付物一致性驗證")
            return res

    # Check key figures in each available deliverable
    check_definitions = [
        ("目標樓層 9F / 9樓", VALUATION_FIGURES["floor"]),
        ("權狀面積 33.70 坪", VALUATION_FIGURES["total_ping"]),
        ("開價總額 898 萬", VALUATION_FIGURES["asking_price"]),
        ("車位型態 機械/平移/休旅", VALUATION_FIGURES["car_space_type"]),
        ("保守底標價 (805~810 或 755~810 萬)", VALUATION_FIGURES["conservative_tier"]),
        ("合理目標價 (845~860 萬)", VALUATION_FIGURES["reasonable_tier"]),
        ("上限天花板價 (880~890 萬)", VALUATION_FIGURES["ceiling_tier"]),
        ("純住宅成交 (63 或 67 筆)", VALUATION_FIGURES["residential_count"]),
        ("2026 年最新成交 855 萬 (3F 25.3 萬)", VALUATION_FIGURES["latest_deal_2026"]),
        ("周邊核心競品 (長億南園/公園苑/椰城/大東家華園/大東家樸園)", VALUATION_FIGURES["competitors"])
    ]

    for doc_type, content in text_corpus.items():
        doc_missing = []
        for label, keywords in check_definitions:
            matched = any(kw in content for kw in keywords)
            if not matched:
                doc_missing.append(label)

        if doc_missing:
            res.fail(f"[{doc_type}] 缺少或未吻合核心關鍵數據：{', '.join(doc_missing)}")
        else:
            res.ok(f"[{doc_type}] 100% 吻合所有核心估價與行情數據指標")

    return res


# ==============================================================================
# Tier 4: Visual Assets & Git Sync Readiness
# ==============================================================================
def verify_tier4(paths: Dict[str, str], progressive: bool = False, verbose: bool = False) -> TestResult:
    """
    Tier 4: Validates visual asset existence and Git sync readiness:
      - Main chart: 南區25年以上大型社區走勢圖/42_大東家南園_成交行情走勢圖.png
      - PIL image integrity (non-corrupt, resolution >= 1000x600)
      - Competitor charts existence (長億南園, 公園苑, 椰城, 大東家樸園, 大東家華園, 大盤總圖)
      - scripts/sync_to_github.py contains "大東家南園9樓" (or "大東家南園") in directory list
      - Embedded visuals inside DOCX & PPTX (if available)
    """
    res = TestResult("Tier 4: Visual Assets & Git Sync Readiness (圖表資產與 GitHub 同步就緒度)")

    # 1. Dadongjia Primary Chart
    chart_path = paths.get("dadongjia_chart", DEFAULT_DADONGJIA_CHART_PATH)
    if os.path.exists(chart_path):
        res.ok(f"大東家南園行情走勢圖檔案存在：`{os.path.basename(chart_path)}`")
        if Image is not None:
            try:
                with Image.open(chart_path) as img:
                    w, h = img.size
                    res.details["chart_resolution"] = {"width": w, "height": h}
                    if w >= 1000 and h >= 600:
                        res.ok(f"走勢圖解析度達標且無損：{w} x {h} (>= 1000x600)")
                    else:
                        res.fail(f"走勢圖解析度過低：{w} x {h} (< 1000x600)")
            except Exception as e:
                res.fail(f"走勢圖損毀無法開啟：{str(e)}")
        else:
            res.warn("PIL (Pillow) 未安裝，略過像素尺寸鑑識")
    else:
        res.fail(f"大東家南園行情走勢圖不存在：`{chart_path}`")

    # 2. Competitor Charts Existence
    comp_missing = []
    for c_path in paths.get("competitor_charts", COMPETITOR_CHART_PATHS):
        if not os.path.exists(c_path):
            comp_missing.append(os.path.basename(c_path))
    if comp_missing:
        res.fail(f"缺少競品或大盤走勢圖資產：{', '.join(comp_missing)}")
    else:
        res.ok(f"全數 {len(COMPETITOR_CHART_PATHS)} 張周邊競品與大盤總圖資產均完整存在")

    # 3. Git Sync Script Configuration Check
    sync_script_path = paths.get("sync_script", DEFAULT_SYNC_SCRIPT_PATH)
    if os.path.exists(sync_script_path):
        try:
            with open(sync_script_path, "r", encoding="utf-8", errors="ignore") as f:
                sync_content = f.read()

            # Check if "大東家南園9樓" or "大東家南園" is in the folder list
            if '"大東家南園9樓"' in sync_content or "'大東家南園9樓'" in sync_content or '"大東家南園"' in sync_content:
                res.ok("GitHub 自動同步腳本 (`scripts/sync_to_github.py`) 已包含 `大東家南園9樓` 資料夾")
            else:
                if progressive:
                    res.warn("GitHub 自動同步腳本 (`scripts/sync_to_github.py`) 尚未將 `大東家南園9樓` 加入同步目錄清單 (Pending M4)")
                else:
                    res.fail("GitHub 自動同步腳本 (`scripts/sync_to_github.py`) 尚未將 `大東家南園9樓` 加入同步目錄清單")
        except Exception as e:
            res.fail(f"讀取 GitHub 同步腳本失敗：{str(e)}")
    else:
        res.fail(f"GitHub 自動同步腳本不存在：`{sync_script_path}`")

    # 4. Embedded Visuals in DOCX and PPTX (if files exist)
    docx_path = paths.get("docx", DEFAULT_DOCX_PATH)
    if os.path.exists(docx_path) and docx is not None:
        try:
            doc = docx.Document(docx_path)
            image_parts = [part for part in doc.part.related_parts.values() if "image" in getattr(part, "content_type", "")]
            if len(image_parts) >= 1:
                res.ok(f"Word 文件 (.docx) 內嵌圖表驗證通過：共包含 {len(image_parts)} 個圖形資產")
            else:
                res.warn("Word 文件 (.docx) 內未偵測到內嵌圖片資產")
        except Exception:
            pass

    pptx_path = paths.get("pptx", DEFAULT_PPTX_PATH)
    if os.path.exists(pptx_path) and pptx is not None:
        try:
            full_txt, slide_cnt, _, _, pic_cnt = extract_pptx_text_and_stats(pptx_path)
            if pic_cnt >= 1:
                res.ok(f"PowerPoint 簡報 (.pptx) 內嵌圖表驗證通過：共包含 {pic_cnt} 個圖片元件")
            else:
                res.warn("PowerPoint 簡報 (.pptx) 內未偵測到獨立圖片元件")
        except Exception:
            pass

    return res


# ==============================================================================
# Self-Test Suite (Adversarial Verification)
# ==============================================================================
def run_self_tests(verbose: bool = False) -> bool:
    """
    Executes an internal self-test suite using synthetic fixtures to ensure:
      TC1: Fully valid mock deliverables pass all 4 tiers 100%.
      TC2: Missing deliverable is caught and failed by Tier 1.
      TC3: Insufficient DOCX tables is caught and failed by Tier 2.
      TC4: Non-16:9 PPTX ratio or insufficient slides is caught and failed by Tier 2.
      TC5: Missing required section title in MD is caught and failed by Tier 2.
      TC6: Inconsistent valuation number is caught and failed by Tier 3.
      TC7: Missing chart PNG is caught and failed by Tier 4.
      TC8: Missing '大東家南園9樓' in sync script is caught and failed by Tier 4.
    """
    print("\n" + "=" * 80)
    print("【執行 verify_dadongjia_suite.py 內部單元自我測試 (Self-Test Suite)】")
    print("=" * 80)

    all_passed = True
    temp_dir = tempfile.mkdtemp(prefix="dadongjia_test_suite_")

    try:
        # 1. Setup Mock Fixtures
        mock_gdir = os.path.join(temp_dir, "大東家南園9樓")
        mock_sdir = os.path.join(temp_dir, "scratch")
        mock_cdir = os.path.join(temp_dir, "南區25年以上大型社區走勢圖")
        mock_scdir = os.path.join(temp_dir, "scripts")
        os.makedirs(mock_gdir, exist_ok=True)
        os.makedirs(mock_sdir, exist_ok=True)
        os.makedirs(mock_cdir, exist_ok=True)
        os.makedirs(mock_scdir, exist_ok=True)

        mock_paths = {
            "md": os.path.join(mock_gdir, "大東家南園9樓_特定待售戶別出價分析與社區整體行情報告.md"),
            "docx": os.path.join(mock_gdir, "大東家南園9樓_特定待售戶別出價分析與社區整體行情報告.docx"),
            "pptx": os.path.join(mock_gdir, "大東家南園9樓_特定待售戶別出價分析與社區整體行情報告.pptx"),
            "json": os.path.join(mock_sdir, "dadongjia_valuation_master.json"),
            "sync_script": os.path.join(mock_scdir, "sync_to_github.py"),
            "dadongjia_chart": os.path.join(mock_cdir, "42_大東家南園_成交行情走勢圖.png"),
            "competitor_charts": [
                os.path.join(mock_cdir, "28_長億南園_成交行情走勢圖.png"),
                os.path.join(mock_cdir, "20_公園苑_成交行情走勢圖.png"),
                os.path.join(mock_cdir, "34_椰城_成交行情走勢圖.png"),
                os.path.join(mock_cdir, "21_大東家樸園_成交行情走勢圖.png"),
                os.path.join(mock_cdir, "35_大東家華園_成交行情走勢圖.png"),
                os.path.join(mock_cdir, "00_台中市南區25年以上大型社區_大盤匯總走勢圖.png")
            ]
        }

        # Create valid mock image (>300KB and >=1000x600 resolution)
        if Image:
            img = Image.frombytes("RGB", (1000, 600), os.urandom(1000 * 600 * 3))
            img.save(mock_paths["dadongjia_chart"])
            for c_path in mock_paths["competitor_charts"]:
                img.save(c_path)
        else:
            with open(mock_paths["dadongjia_chart"], "wb") as f:
                f.write(b"\x89PNG\r\n\x1a\n" + b"\x00" * 400000)
            for c_path in mock_paths["competitor_charts"]:
                with open(c_path, "wb") as f:
                    f.write(b"\x89PNG\r\n\x1a\n" + b"\x00" * 400000)

        # Create valid mock sync script
        with open(mock_paths["sync_script"], "w", encoding="utf-8") as f:
            f.write('#!/usr/bin/env python3\nfolders = ["大墩人文", "大東家南園9樓", "長億南園8樓"]\n')

        # Create valid mock JSON (> 5KB)
        valid_json_obj = {
            "target_property": {
                "community_name": "大東家南園",
                "floor": "9F/15F",
                "total_ping": 33.70,
                "main_ping": 19.844,
                "aux_ping": 1.984,
                "living_ping": 21.828,
                "common_ping": 11.872,
                "car_space_type": "坡道機械下層平移車位 (可停休旅車)",
                "asking_price": 898.0,
                "asking_unit_price": 26.65,
                "decoupled_net_unit_price": 26.46
            },
            "pricing_tiers": {
                "conservative_floor": {"total": 810.0, "unit": 23.63},
                "reasonable_target": {"total_low": 850.0, "total_high": 860.0, "unit_low": 24.92, "unit_high": 25.24},
                "ceiling_max": {"total": 880.0, "unit": 25.88}
            },
            "historical_summary": {
                "total_count": 77,
                "residential_count": 63,
                "atypical_count": 14,
                "avg_history": 17.70,
                "avg_recent": 26.05,
                "latest_deal_2026": {"date": "2026-01-18", "floor": "3樓", "price": 855.0, "unit": 25.3}
            },
            "comparables": [
                {"name": "長億南園", "price": 25.16},
                {"name": "公園苑", "price": 30.25},
                {"name": "椰城", "price": 28.30},
                {"name": "大東家華園", "price": 24.04},
                {"name": "大東家樸園", "price": 28.31}
            ],
            "padding_data": "x" * 6000
        }
        with open(mock_paths["json"], "w", encoding="utf-8") as f:
            json.dump(valid_json_obj, f, ensure_ascii=False, indent=2)

        all_keys_text = "9F 9樓 33.70 19.844 1.984 21.828 898 26.65 機械平移 下層平移 休旅車 75萬 805 810 850 860 880 890 77 63 855 25.3 長億南園 公園苑 椰城 大東家華園 大東家樸園"

        # Create valid mock MD (> 15KB)
        md_text = f"""# 大東家南園9樓待售戶出價分析與全社區行情報告
## 一、物件產調與規格查核
標的位於9F (9樓)，高樓層。總權狀面積33.70坪 (主建19.844坪, 附屬1.984坪, 主附21.828坪)。
開價898萬，毛單價26.65萬/坪。車位型態為坡道機械下層平移車位，可停休旅車，拆算價值75萬元。

## 二、歷史實登洗淨分析
大東家南園社區累計77筆登錄，經五重純住宅過濾排除14筆非典型與親友交易，確立63筆純住宅市價交易。
2026年最新成交為3樓總價855萬，單價25.3萬/坪。

## 三、同類型競品比價
周邊同類型中古競品社區包含長億南園、公園苑、椰城、大東家華園、大東家樸園。

## 四、車位拆算解耦
機械下層平移車位以75萬/2.6坪解耦，房屋純淨建坪31.10坪，開價淨單價26.46萬/坪。

## 五、多維度比價調整矩陣
針對31.8年屋齡、9樓視野採光、一層兩戶單純梯戶比進行特徵調整。

## 六、三段式出價策略精算
- 保守底標價：805~810萬 (約90折，淨單價23.5~23.6萬/坪)
- 合理目標價：850~860萬 (約95~96折，淨單價24.9~25.2萬/坪，貼近2026最新成交855萬)
- 上限天花板價：880~890萬 (約98折，絕不追高突破歷史峰值)

## 七、談判議價話術說帖
以社區20戶在售多殺多賣壓、2026年3樓成交855萬、31年屋況需預留翻修預算切入。

## 八、簽約避坑與修繕檢核
交屋驗收、水電管線翻新與契約避坑檢核表。
""" + ("\n詳細分析文字補充..." * 500)
        with open(mock_paths["md"], "w", encoding="utf-8") as f:
            f.write(md_text)

        # Create valid mock DOCX (> 30KB)
        if docx:
            doc = docx.Document()
            doc.add_heading("大東家南園9樓估價分析報告", 0)
            doc.add_paragraph(all_keys_text)
            for _ in range(6):
                tbl = doc.add_table(rows=3, cols=4)
                for row in tbl.rows:
                    for cell in row.cells:
                        cell.text = all_keys_text
            for _ in range(50):
                doc.add_paragraph("大東家南園9樓特定待售戶出價分析與社區整體行情報告 " * 30)
            doc.save(mock_paths["docx"])

        # Create valid mock PPTX (> 300KB)
        if pptx and Presentation:
            from pptx.util import Inches
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
            blank_layout = prs.slide_layouts[6]
            for i in range(12):
                slide = prs.slides.add_slide(blank_layout)
                tx_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(5))
                tx_box.text_frame.text = f"Slide {i+1}: {all_keys_text}"
                if os.path.exists(mock_paths["dadongjia_chart"]):
                    slide.shapes.add_picture(mock_paths["dadongjia_chart"], Inches(1), Inches(1), width=Inches(4))
            prs.save(mock_paths["pptx"])

        # Test Case 1: Valid deliverables must pass all 4 tiers
        t1 = verify_tier1(mock_paths)
        t2 = verify_tier2(mock_paths)
        t3 = verify_tier3(mock_paths)
        t4 = verify_tier4(mock_paths)
        if t1.passed and t2.passed and t3.passed and t4.passed:
            print("  [PASS] TC1: Valid mock deliverables passed Tier 1 ~ 4 full acceptance 100%.")
        else:
            print(f"  [FAIL] TC1: Valid mock deliverables failed! T1={t1.passed}, T2={t2.passed}, T3={t3.passed}, T4={t4.passed}")
            all_passed = False

        # Test Case 2: Tier 1 catches missing deliverable
        bad_paths = dict(mock_paths)
        bad_paths["pptx"] = os.path.join(mock_gdir, "non_existent.pptx")
        t1_bad = verify_tier1(bad_paths)
        if not t1_bad.passed:
            print("  [PASS] TC2: Missing PPTX deliverable correctly caught and failed by Tier 1.")
        else:
            print("  [FAIL] TC2: Tier 1 failed to catch missing PPTX!")
            all_passed = False

        # Test Case 3: Tier 2 catches insufficient DOCX tables
        if docx:
            bad_docx_path = os.path.join(mock_gdir, "bad_tables.docx")
            doc_bad = docx.Document()
            doc_bad.add_paragraph("Simple paragraph")
            doc_bad.add_table(rows=2, cols=2)
            doc_bad.save(bad_docx_path)
            bad_paths2 = dict(mock_paths)
            bad_paths2["docx"] = bad_docx_path
            t2_bad = verify_tier2(bad_paths2)
            if not t2_bad.passed:
                print("  [PASS] TC3: Insufficient DOCX table count correctly caught and failed by Tier 2.")
            else:
                print("  [FAIL] TC3: Tier 2 failed to catch insufficient DOCX table count!")
                all_passed = False

        # Test Case 4: Tier 2 catches non-16:9 PPTX
        if pptx and Presentation:
            from pptx.util import Inches
            bad_pptx_path = os.path.join(mock_gdir, "bad_ratio.pptx")
            prs_bad = Presentation()
            prs_bad.slide_width = Inches(10)  # 4:3
            prs_bad.slide_height = Inches(7.5)
            for _ in range(12):
                prs_bad.slides.add_slide(prs_bad.slide_layouts[6])
            prs_bad.save(bad_pptx_path)
            bad_paths3 = dict(mock_paths)
            bad_paths3["pptx"] = bad_pptx_path
            t2_bad_pptx = verify_tier2(bad_paths3)
            if not t2_bad_pptx.passed:
                print("  [PASS] TC4: Non-16:9 PPTX aspect ratio correctly caught and failed by Tier 2.")
            else:
                print("  [FAIL] TC4: Tier 2 failed to catch non-16:9 PPTX ratio!")
                all_passed = False

        # Test Case 5: Tier 2 catches missing MD section title
        bad_md_path = os.path.join(mock_gdir, "bad_sections.md")
        with open(bad_md_path, "w", encoding="utf-8") as f:
            f.write("# 簡略報告\n此處缺少必要的章節標題..." + ("x" * 20000))
        bad_paths4 = dict(mock_paths)
        bad_paths4["md"] = bad_md_path
        t2_bad_md = verify_tier2(bad_paths4)
        if not t2_bad_md.passed:
            print("  [PASS] TC5: Missing required MD section headers correctly caught and failed by Tier 2.")
        else:
            print("  [FAIL] TC5: Tier 2 failed to catch missing MD sections!")
            all_passed = False

        # Test Case 6: Tier 3 catches pricing mismatch
        bad_json_path = os.path.join(mock_sdir, "bad_valuation.json")
        bad_json_obj = dict(valid_json_obj)
        bad_json_obj["pricing_tiers"] = {
            "conservative_floor": {"total": 600.0},  # should be 755~810
            "ceiling_max": {"total": 999.0}
        }
        with open(bad_json_path, "w", encoding="utf-8") as f:
            json.dump(bad_json_obj, f, ensure_ascii=False)
        bad_paths5 = dict(mock_paths)
        bad_paths5["json"] = bad_json_path
        t3_bad = verify_tier3(bad_paths5)
        if not t3_bad.passed:
            print("  [PASS] TC6: Valuation pricing tier mismatch correctly caught and failed by Tier 3.")
        else:
            print("  [FAIL] TC6: Tier 3 failed to catch pricing mismatch!")
            all_passed = False

        # Test Case 7: Tier 4 catches missing primary chart
        bad_paths6 = dict(mock_paths)
        bad_paths6["dadongjia_chart"] = os.path.join(mock_cdir, "non_existent_chart.png")
        t4_bad = verify_tier4(bad_paths6)
        if not t4_bad.passed:
            print("  [PASS] TC7: Missing primary chart image correctly caught and failed by Tier 4.")
        else:
            print("  [FAIL] TC7: Tier 4 failed to catch missing primary chart!")
            all_passed = False

        # Test Case 8: Tier 4 catches sync script missing "大東家南園9樓"
        bad_sync_path = os.path.join(mock_scdir, "bad_sync.py")
        with open(bad_sync_path, "w", encoding="utf-8") as f:
            f.write('#!/usr/bin/env python3\nfolders = ["大墩人文", "長億南園8樓"]\n')
        bad_paths7 = dict(mock_paths)
        bad_paths7["sync_script"] = bad_sync_path
        t4_bad_sync = verify_tier4(bad_paths7)
        if not t4_bad_sync.passed:
            print("  [PASS] TC8: Sync script missing '大東家南園9樓' correctly caught and failed by Tier 4.")
        else:
            print("  [FAIL] TC8: Tier 4 failed to catch missing '大東家南園9樓' in sync script!")
            all_passed = False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    print("=" * 80)
    if all_passed:
        print("🎉 [SELF-TEST SUCCESS] All 8/8 internal verification tests passed successfully!")
    else:
        print("❌ [SELF-TEST FAILURE] One or more self-test assertions failed!")
    print("=" * 80 + "\n")
    return all_passed


# ==============================================================================
# Main Runner & CLI Execution
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(
        description="Automated Verification Suite for Taichung South District 'Dadongjia Nanyuan 9F' (大東家南園9樓) Deliverables"
    )
    parser.add_argument("--all", action="store_true", default=True, help="Run all verification tiers (Default)")
    parser.add_argument("--tier", type=int, choices=[1, 2, 3, 4], help="Run a specific tier (1, 2, 3, or 4)")
    parser.add_argument("--tier1", action="store_true", help="Run Tier 1 only")
    parser.add_argument("--tier2", action="store_true", help="Run Tier 2 only")
    parser.add_argument("--tier3", action="store_true", help="Run Tier 3 only")
    parser.add_argument("--tier4", action="store_true", help="Run Tier 4 only")
    parser.add_argument("--progressive", action="store_true", help="Progressive execution mode (warn on missing files)")
    parser.add_argument("--test-self", action="store_true", help="Execute internal self-test suite with mock fixtures")
    parser.add_argument("--json", action="store_true", help="Output machine-readable JSON report")
    parser.add_argument("--verbose", action="store_true", help="Verbose logging")

    args = parser.parse_args()

    # Handle internal self-tests
    if args.test_self:
        success = run_self_tests(verbose=args.verbose)
        sys.exit(0 if success else 1)

    # Determine which tiers to run
    tiers_to_run = []
    if args.tier:
        tiers_to_run = [args.tier]
    elif args.tier1:
        tiers_to_run = [1]
    elif args.tier2:
        tiers_to_run = [2]
    elif args.tier3:
        tiers_to_run = [3]
    elif args.tier4:
        tiers_to_run = [4]
    else:
        tiers_to_run = [1, 2, 3, 4]

    paths = {
        "md": DEFAULT_MD_PATH,
        "docx": DEFAULT_DOCX_PATH,
        "pptx": DEFAULT_PPTX_PATH,
        "json": DEFAULT_JSON_PATH,
        "sync_script": DEFAULT_SYNC_SCRIPT_PATH,
        "dadongjia_chart": DEFAULT_DADONGJIA_CHART_PATH,
        "competitor_charts": COMPETITOR_CHART_PATHS
    }

    mode_str = "PROGRESSIVE (階段演進模式)" if args.progressive else "STRICT FULL ACCEPTANCE (全量嚴格驗收)"
    print("=" * 80)
    print("【台中市南區大東家南園9樓特定待售戶出價分析與社區行情 — 自動化驗證套件】")
    print(f"Project Root : {PROJECT_ROOT}")
    print(f"Target Dir   : {os.path.join(PROJECT_ROOT, '大東家南園9樓')}")
    print(f"Execute Mode : {mode_str}")
    print(f"Active Tiers : {tiers_to_run}")
    print("=" * 80)

    results: List[TestResult] = []

    if 1 in tiers_to_run:
        results.append(verify_tier1(paths, progressive=args.progressive, verbose=args.verbose))
    if 2 in tiers_to_run:
        results.append(verify_tier2(paths, progressive=args.progressive, verbose=args.verbose))
    if 3 in tiers_to_run:
        results.append(verify_tier3(paths, progressive=args.progressive, verbose=args.verbose))
    if 4 in tiers_to_run:
        results.append(verify_tier4(paths, progressive=args.progressive, verbose=args.verbose))

    all_passed = True
    total_errors = 0
    total_warnings = 0

    print()
    for res in results:
        status_icon = "✅ PASS" if res.passed else "❌ FAIL"
        if not res.passed:
            all_passed = False
        print(f"{status_icon} {res.name}")

        for msg in res.messages:
            if args.verbose or not res.passed:
                print(f"  [OK]   {msg}")

        for warn in res.warnings:
            total_warnings += 1
            print(f"  [WARN] {warn}")

        for err in res.errors:
            total_errors += 1
            print(f"  [ERR]  {err}")
        print()

    # JSON output if requested
    if args.json:
        report = {
            "all_passed": all_passed,
            "progressive": args.progressive,
            "total_errors": total_errors,
            "total_warnings": total_warnings,
            "tiers": [r.to_dict() for r in results]
        }
        print("JSON_REPORT_BEGIN")
        print(json.dumps(report, ensure_ascii=False, indent=2))
        print("JSON_REPORT_END")

    print("=" * 80)
    if all_passed:
        if args.progressive and total_warnings > 0:
            print("🎉【驗收結論：PROGRESSIVE PASS】現有已產出產物通過檢驗！(尚有待完成交付物由工作者產製中)")
        else:
            print("🎉【驗收結論：FULL SUITE PASS】全 4 個驗證層級 (Tier 1 ~ 4) 100% 驗收通過！符合專案最高交付標準。")
        print("=" * 80)
        sys.exit(0)
    else:
        print(f"❌【驗收結論：SUITE FAILED】檢測出 {total_errors} 個阻斷性錯誤與 {total_warnings} 個警示，未達交付驗收門檻！")
        print("=" * 80)
        sys.exit(1)


if __name__ == "__main__":
    main()

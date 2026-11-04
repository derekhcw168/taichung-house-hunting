#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scripts/verify_beverly_suite.py
Automated 4-Tier Verification Suite for Taichung South District "Beverly Hills 13F" (比佛利山莊13樓)
Valuation and Community Market Report Deliverables.

Four Verification Tiers:
  - Tier 1: File Existence & Basic Checks (MD >=15KB, DOCX >=30KB, PPTX >=200KB, JSON >=5KB)
  - Tier 2: Format & Structural Integrity (DOCX tables/cells/images, PPTX 16:9 slides/images, MD required chapters, JSON schema)
  - Tier 3: 100% Cross-Deliverable Numerical Consistency (998萬, 33.75坪, 30.48坪, 25.43坪, 2.98坪, 28.41坪, 2.06坪, 3.27坪 B1-89, 13F/14F, 2024-01-24 purchased at 870萬, 3-tier bidding strategy)
  - Tier 4: Visual Assets & Git Sync Readiness (PNG chart resolution >= 2000x1500, sync_to_github.py inclusion)

CLI Arguments:
  --all (default)       : Full strict acceptance verification across all 4 tiers
  --progressive         : Progressive mode for ongoing milestone validation (warns on missing deliverables)
  --tier {1,2,3,4}      : Execute a specific tier (also --tier1, --tier2, --tier3, --tier4)
  --test-self           : Execute internal self-test suite with synthetic mock fixtures
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
DEFAULT_MD_PATH = os.path.join(PROJECT_ROOT, "比佛利山莊13樓", "比佛利山莊13樓_特定待售戶別出價分析與社區整體行情報告.md")
DEFAULT_DOCX_PATH = os.path.join(PROJECT_ROOT, "比佛利山莊13樓", "比佛利山莊13樓_特定待售戶別出價分析與社區整體行情報告.docx")
DEFAULT_PPTX_PATH = os.path.join(PROJECT_ROOT, "比佛利山莊13樓", "比佛利山莊13樓_特定待售戶別出價分析與社區整體行情報告.pptx")
DEFAULT_JSON_PATH = os.path.join(PROJECT_ROOT, "scratch", "beverly_valuation_master.json")
DEFAULT_SYNC_SCRIPT_PATH = os.path.join(PROJECT_ROOT, "scripts", "sync_to_github.py")

# Visual Assets Paths
DEFAULT_BEVERLY_CHART_PATH = os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "17_比佛利山莊_成交行情走勢圖.png")
COMPETITOR_CHART_PATHS = [
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "00_台中市南區25年以上大型社區_大盤匯總走勢圖.png"),
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "29_歌德堡_成交行情走勢圖.png"),
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "20_公園苑_成交行情走勢圖.png"),
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "28_長億南園_成交行情走勢圖.png"),
]

# Required Size Thresholds (bytes)
SIZE_THRESHOLDS = {
    "md": 15 * 1024,      # 15 KB
    "docx": 30 * 1024,    # 30 KB
    "pptx": 200 * 1024,   # 200 KB
    "json": 5 * 1024      # 5 KB
}

# Required Markdown Section Keywords
REQUIRED_MD_SECTIONS = [
    ("物件現況", r"物件現況|物件產調|標的現況|待售物件"),
    ("社區歷程", r"社區歷程|歷史行情|實價登錄歷程|成交行情"),
    ("競品比價", r"競品比價|競品對比|同型競品|周邊社區"),
    ("車位拆算", r"車位拆算|車位分離|車位價格拆算|淨建坪單價"),
    ("出價三道防線", r"出價三道防線|三段式出價|出價策略|出價區間"),
    ("談判話術", r"談判話術|談判說帖|議價話術|議價說帖|攻防要點")
]

# Required Canonical Valuation Figures for Beverly Hills 13F
VALUATION_FIGURES = {
    "floor": ["13F", "13樓", "次頂樓", "13/14F", "13 / 14"],
    "asking_price": ["998"],
    "total_ping": ["33.75"],
    "house_ping": ["30.48"],
    "main_ping": ["25.43"],
    "aux_ping": ["2.98"],
    "indoor_ping": ["28.41"],
    "common_ping": ["2.06"],
    "public_ratio": ["6.76", "6.77", "6.8", "6.1"],
    "parking_ping": ["3.27"],
    "parking_type": ["B1-89", "B1 - 89", "機械", "坡道機械"],
    "smoking_gun_date": ["2024-01-24", "2024/01/24", "113年1月24日", "113/01/24", "113年1月", "2024-01", "2024年1月"],
    "smoking_gun_price": ["870"],
    "parking_benchmark": ["70", "60", "80"],
    "asking_net_unit_price": ["30.45"],
    "conservative_tier": ["810", "800", "820", "24.28", "24.3", "24.2"],
    "target_tier": ["870", "860", "880", "26.25", "26.2", "26.3"],
    "ceiling_tier": ["920", "910", "27.89", "27.9", "27.8"],
    "competitors": ["佛羅里達", "家綠寶", "歌德堡", "公園苑", "宏都雙子星", "鄉林新月灣"]
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


def extract_docx_text_and_stats(path: str) -> Tuple[str, int, int, int, int]:
    """
    Parses DOCX and extracts full text, table count, total cells, paragraph count, and image count.
    Returns: (full_text, table_count, total_cells, paragraph_count, image_count)
    """
    if not os.path.exists(path) or docx is None:
        return "", 0, 0, 0, 0
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

        # Count embedded image parts
        image_count = 0
        if hasattr(doc, 'part') and hasattr(doc.part, 'related_parts'):
            for part in doc.part.related_parts.values():
                content_type = getattr(part, 'content_type', '')
                if 'image' in content_type:
                    image_count += 1

        return full_text, table_count, total_cells, len(doc.paragraphs), image_count
    except Exception:
        return "", 0, 0, 0, 0


def extract_pptx_text_and_stats(path: str) -> Tuple[str, int, float, float, int]:
    """
    Parses PPTX and extracts full text, slide count, width (inches), height (inches), and picture count.
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
# Tier 1: File Existence & Non-trivial Size Thresholds
# ==============================================================================
def verify_tier1(paths: Dict[str, str], progressive: bool = False, verbose: bool = False) -> TestResult:
    """
    Tier 1: Checks existence and minimum size thresholds of:
      - MD (>= 15KB)
      - DOCX (>= 30KB)
      - PPTX (>= 200KB)
      - JSON (>= 5KB)
    """
    res = TestResult("Tier 1: File Existence & Non-trivial Size Thresholds (產物存在性與檔案容量檢驗)")

    targets = [
        ("Markdown 分析報告", paths.get("md", DEFAULT_MD_PATH), SIZE_THRESHOLDS["md"], "15KB"),
        ("Word 分析報告 (.docx)", paths.get("docx", DEFAULT_DOCX_PATH), SIZE_THRESHOLDS["docx"], "30KB"),
        ("PowerPoint 簡報 (.pptx)", paths.get("pptx", DEFAULT_PPTX_PATH), SIZE_THRESHOLDS["pptx"], "200KB"),
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
# Tier 2: Internal Structural & Schema Integrity
# ==============================================================================
def verify_tier2(paths: Dict[str, str], progressive: bool = False, verbose: bool = False) -> TestResult:
    """
    Tier 2: Validates structural and schema integrity of deliverables:
      - DOCX: parsable, >= 5 tables, >= 50 total cells, embedded images exist
      - PPTX: parsable, slide count >= 10, 16:9 widescreen (13.333" x 7.5" +/- 0.05")
      - MD: contains 6 core chapters (物件現況, 社區歷程, 競品比價, 車位拆算, 出價三道防線, 談判話術)
      - JSON: schema keys (target_property, valuation_model, community_history, comparables)
    """
    res = TestResult("Tier 2: Internal Structural & Schema Integrity (檔案格式、內部結構與綱要完整性)")

    # 1. DOCX Verification
    docx_path = paths.get("docx", DEFAULT_DOCX_PATH)
    if os.path.exists(docx_path):
        if docx is None:
            res.fail("python-docx 函式庫未安裝，無法驗證 .docx 結構")
        else:
            try:
                full_text, table_cnt, cell_cnt, p_cnt, img_cnt = extract_docx_text_and_stats(docx_path)
                res.details["docx_stats"] = {
                    "tables": table_cnt,
                    "cells": cell_cnt,
                    "paragraphs": p_cnt,
                    "images": img_cnt
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

                if img_cnt < 1:
                    res.fail("Word 文件 (.docx) 未偵測到內嵌圖片資產，要求必須內嵌高清走勢圖")
                else:
                    res.ok(f"Word 文件 (.docx) 內嵌圖表驗證通過：共 {img_cnt} 個圖片資產")
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

                # Verify 16:9 widescreen dimensions (width 13.333" +/- 0.05", height 7.5" +/- 0.05")
                width_ok = abs(width_in - 13.333) <= 0.06 or (height_in > 0 and abs((width_in / height_in) - (16.0 / 9.0)) < 0.05)
                height_ok = abs(height_in - 7.5) <= 0.06

                if width_ok and height_ok:
                    res.ok(f"PowerPoint 簡報 (.pptx) 比例符合 16:9 寬螢幕規格 ({width_in:.3f}\" x {height_in:.3f}\")")
                else:
                    res.fail(f"PowerPoint 簡報 (.pptx) 比例不符合 16:9：寬度 {width_in:.3f}\" (預期 13.333\" +/- 0.05\"), 高度 {height_in:.3f}\" (預期 7.5\" +/- 0.05\")")
            except Exception as e:
                res.fail(f"PowerPoint 簡報 (.pptx) 解析發生錯誤：{str(e)}")
    else:
        if progressive:
            res.warn("PowerPoint 簡報尚未產生，略過 Tier 2 PPTX 檢驗")
        else:
            res.fail(f"PowerPoint 簡報不存在，無法驗證結構：`{pptx_path}`")

    # 3. Markdown Sections Verification
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
            res.ok("Markdown 報告包含全部 6 大核心章節 (物件現況, 社區歷程, 競品比價, 車位拆算, 出價三道防線, 談判話術)")
    else:
        if progressive:
            res.warn("Markdown 分析報告尚未產生，略過 Tier 2 MD 檢驗")
        else:
            res.fail(f"Markdown 分析報告不存在，無法驗證章節結構：`{md_path}`")

    # 4. Master JSON Schema Verification
    json_path = paths.get("json", DEFAULT_JSON_PATH)
    if os.path.exists(json_path):
        json_data = extract_json_data(json_path)
        if json_data is None:
            res.fail(f"估價主資料庫 JSON 解析失敗，非有效 JSON 格式：`{json_path}`")
        else:
            required_keys = ["target_property", "valuation_model", "community_history", "comparables"]
            missing_keys = [k for k in required_keys if k not in json_data]
            if missing_keys:
                res.fail(f"估價主資料庫 JSON 缺少核心綱要鍵值：{', '.join(missing_keys)}")
            else:
                res.ok(f"估價主資料庫 JSON 綱要完整：已驗證核心鍵值 ({', '.join(required_keys)})")
    else:
        if progressive:
            res.warn("估價主資料庫 JSON 尚未產生，略過 Tier 2 JSON 綱要檢驗")
        else:
            res.fail(f"估價主資料庫 JSON 不存在：`{json_path}`")

    return res


# ==============================================================================
# Tier 3: 100% Cross-Deliverable Numerical Consistency
# ==============================================================================
def verify_tier3(paths: Dict[str, str], progressive: bool = False, verbose: bool = False) -> TestResult:
    """
    Tier 3: Cross-verifies exact numerical consistency across MD, DOCX, PPTX, and JSON:
      - Asking price: 998 萬
      - Total area: 33.75 坪
      - Pure house area: 30.48 坪
      - Main indoor area: 25.43 坪
      - Attached area: 2.98 坪
      - Indoor total area: 28.41 坪
      - Common area: 2.06 坪, public ratio ~6.76% (or 6.1% ~ 6.8%)
      - Parking: 3.27 坪 mechanical parking B1-89
      - Floor: 13F / 14F
      - Previous transaction smoking gun: 2024-01-24 purchased at 870 萬元
      - 3-tier bidding strategy (conservative ~810萬/24.28萬, target ~870萬/26.25萬, ceiling ~920萬/27.89萬, asking net ~30.45萬)
    """
    res = TestResult("Tier 3: 100% Cross-Deliverable Numerical Consistency (跨交付物核心數值一致性)")

    # 1. Master JSON Consistency
    json_path = paths.get("json", DEFAULT_JSON_PATH)
    json_data = extract_json_data(json_path)

    if json_data:
        res.ok(f"成功讀取主資料庫 JSON：`{os.path.basename(json_path)}`")
        tp = json_data.get("target_property", {})
        vm = json_data.get("valuation_model", {})
        ch = json_data.get("community_history", {})
        comps = json_data.get("comparables", [])

        # Check Target Property figures in JSON
        if tp:
            floor_str = str(tp.get("floor", ""))
            if "13" in floor_str:
                res.ok(f"JSON 目標物件樓層正確：{floor_str}")
            else:
                res.fail(f"JSON 目標物件樓層錯誤：{floor_str} (應包含 13F / 13樓)")

            # Total ping (33.75)
            tot_p = float(tp.get("total_area", tp.get("total_ping", 0)))
            if abs(tot_p - 33.75) < 0.05:
                res.ok(f"JSON 總建坪正確：{tot_p} 坪")
            else:
                res.fail(f"JSON 總建坪不吻合：{tot_p} (預期 33.75)")

            # Pure house ping (30.48)
            house_p = float(tp.get("house_area", tp.get("net_building_area", 0)))
            if abs(house_p - 30.48) < 0.05:
                res.ok(f"JSON 純建物面積正確：{house_p} 坪")
            else:
                res.fail(f"JSON 純建物面積不吻合：{house_p} (預期 30.48)")

            # Main indoor ping (25.43)
            main_p = float(tp.get("main_area", tp.get("main_indoor_area", 0)))
            if abs(main_p - 25.43) < 0.05:
                res.ok(f"JSON 主建物面積正確：{main_p} 坪")
            else:
                res.fail(f"JSON 主建物面積不吻合：{main_p} (預期 25.43)")

            # Attached ping (2.98)
            att_p = float(tp.get("attached_area", 0))
            if abs(att_p - 2.98) < 0.05:
                res.ok(f"JSON 附屬建物面積正確：{att_p} 坪")
            else:
                res.fail(f"JSON 附屬建物面積不吻合：{att_p} (預期 2.98)")

            # Indoor total (28.41)
            ind_p = float(tp.get("indoor_area", tp.get("indoor_total_area", 0)))
            if abs(ind_p - 28.41) < 0.05:
                res.ok(f"JSON 室內合計面積正確：{ind_p} 坪")
            else:
                res.fail(f"JSON 室內合計面積不吻合：{ind_p} (預期 28.41)")

            # Common ping (2.06)
            com_p = float(tp.get("public_area", tp.get("common_area", 0)))
            if abs(com_p - 2.06) < 0.05:
                res.ok(f"JSON 公設面積正確：{com_p} 坪")
            else:
                res.fail(f"JSON 公設面積不吻合：{com_p} (預期 2.06)")

            # Public ratio (~6.76% or 6.1% ~ 6.8%)
            pub_r = float(tp.get("public_ratio", 0))
            if (6.0 <= pub_r <= 6.9):
                res.ok(f"JSON 公設比正確：{pub_r}% (符合公設比 6.1% ~ 6.8% 範圍)")
            else:
                res.fail(f"JSON 公設比不吻合：{pub_r}% (預期 6.76% 或 6.1% ~ 6.8%)")

            # Parking ping (3.27)
            prk_p = float(tp.get("parking_area", 0))
            if abs(prk_p - 3.27) < 0.05:
                res.ok(f"JSON 車位面積正確：{prk_p} 坪")
            else:
                res.fail(f"JSON 車位面積不吻合：{prk_p} (預期 3.27)")

            # Parking type / number
            prk_type = str(tp.get("parking_type", tp.get("parking_desc", "")))
            if "B1-89" in prk_type or "機械" in prk_type:
                res.ok(f"JSON 車位型態規格正確：{prk_type}")
            else:
                res.fail(f"JSON 車位型態規格不吻合：{prk_type} (應標註 B1-89 機械車位)")

            # Asking price (998)
            ask_p = float(tp.get("asking_price", 0))
            if abs(ask_p - 998.0) < 1.0:
                res.ok(f"JSON 開價總額正確：{ask_p} 萬")
            else:
                res.fail(f"JSON 開價總額不吻合：{ask_p} (預期 998)")

            # Previous purchase smoking gun
            prev_sale = tp.get("previous_sale", {})
            if prev_sale:
                ps_date = str(prev_sale.get("date", ""))
                ps_price = float(prev_sale.get("price", 0))
                if "2024-01-24" in ps_date or "2024/01/24" in ps_date or "113" in ps_date:
                    res.ok(f"JSON 前次交易取得日期鐵證正確：{ps_date}")
                else:
                    res.fail(f"JSON 前次交易取得日期鐵證不吻合：{ps_date} (預期 2024-01-24)")

                if abs(ps_price - 870.0) < 1.0:
                    res.ok(f"JSON 前次交易取得價格鐵證正確：{ps_price} 萬")
                else:
                    res.fail(f"JSON 前次交易取得價格鐵證不吻合：{ps_price} (預期 870 萬)")
            else:
                res.fail("JSON 缺少 previous_sale 前次交易取得成本欄位")

        # Check Valuation Model in JSON
        if vm:
            # Parking benchmark
            pb = float(vm.get("parking_benchmark", 0))
            if 60.0 <= pb <= 80.0:
                res.ok(f"JSON 車位拆算基準價格合理：{pb} 萬 (基準 70 萬)")
            else:
                res.fail(f"JSON 車位拆算基準價格異常：{pb} 萬 (預期 60~80 萬，基準 70 萬)")

            # Asking net unit price (30.45)
            anup = float(vm.get("asking_net_unit_price", 0))
            if abs(anup - 30.45) < 0.15:
                res.ok(f"JSON 開價淨建坪單價正確：{anup} 萬/坪")
            else:
                res.fail(f"JSON 開價淨建坪單價不吻合：{anup} (預期 30.45 萬/坪)")

            # Conservative tier (800~820 萬, ~24.28 萬/坪)
            cons = vm.get("conservative", {})
            cons_tot = float(cons.get("total", 0)) if isinstance(cons, dict) else float(cons)
            if 800.0 <= cons_tot <= 825.0:
                res.ok(f"JSON 保守底標價總額正確：{cons_tot} 萬 (預期 810 萬或 800~820 萬)")
            else:
                res.fail(f"JSON 保守底標價總額不吻合：{cons_tot} (預期 810 萬或 800~820 萬)")

            # Target tier (860~880 萬, ~26.25 萬/坪, 錨定屋主買進價 870 萬)
            targ = vm.get("target", {})
            targ_tot = float(targ.get("total", 0)) if isinstance(targ, dict) else float(targ)
            if 860.0 <= targ_tot <= 880.0:
                res.ok(f"JSON 合理目標價總額正確：{targ_tot} 萬 (預期 870 萬或 860~880 萬)")
            else:
                res.fail(f"JSON 合理目標價總額不吻合：{targ_tot} (預期 870 萬或 860~880 萬)")

            # Ceiling tier (910~925 萬, ~27.89 萬/坪)
            ceil = vm.get("ceiling", {})
            ceil_tot = float(ceil.get("total", 0)) if isinstance(ceil, dict) else float(ceil)
            if 910.0 <= ceil_tot <= 925.0:
                res.ok(f"JSON 上限天花板價總額正確：{ceil_tot} 萬 (預期 920 萬或 910~920 萬)")
            else:
                res.fail(f"JSON 上限天花板價總額不吻合：{ceil_tot} (預期 920 萬或 910~920 萬)")

        # Check Competitors in JSON
        if isinstance(comps, list) and comps:
            comp_names = [c.get("name", "") if isinstance(c, dict) else str(c) for c in comps]
            all_comp_str = " ".join(comp_names)
            found_comps = [c_name for c_name in VALUATION_FIGURES["competitors"] if c_name in all_comp_str]
            if len(found_comps) >= 2:
                res.ok(f"JSON 競品清單包含核心社區：{', '.join(found_comps)}")
            else:
                res.fail(f"JSON 競品清單缺少足夠的核心競品社區 (至少需包含佛羅里達/家綠寶/歌德堡/公園苑等 2 個以上)")
    else:
        if progressive:
            res.warn("主資料庫 JSON 尚未產出，略過 JSON 內部欄位檢驗")
        else:
            res.fail(f"主資料庫 JSON 不存在或無法解析：`{json_path}`")

    # 2. Cross-Verification on Document Deliverables (MD, DOCX, PPTX)
    text_corpus: Dict[str, str] = {}
    md_path = paths.get("md", DEFAULT_MD_PATH)
    if os.path.exists(md_path):
        text_corpus["Markdown 報告"] = extract_md_text(md_path)

    docx_path = paths.get("docx", DEFAULT_DOCX_PATH)
    if os.path.exists(docx_path):
        docx_txt, _, _, _, _ = extract_docx_text_and_stats(docx_path)
        text_corpus["Word 文件 (.docx)"] = docx_txt

    pptx_path = paths.get("pptx", DEFAULT_PPTX_PATH)
    if os.path.exists(pptx_path):
        pptx_txt, _, _, _, _ = extract_pptx_text_and_stats(pptx_path)
        text_corpus["PowerPoint 簡報 (.pptx)"] = pptx_txt

    if not text_corpus:
        if progressive:
            res.warn("所有文檔交付物 (MD/DOCX/PPTX) 均未產生，略過文檔間一致性驗證")
            return res
        else:
            res.fail("所有文檔交付物均不存在，無法進行跨交付物一致性驗證")
            return res

    # Check key figures in each available deliverable
    check_definitions = [
        ("目標樓層 13F / 次頂樓", VALUATION_FIGURES["floor"]),
        ("開價總額 998 萬", VALUATION_FIGURES["asking_price"]),
        ("總權狀坪數 33.75 坪", VALUATION_FIGURES["total_ping"]),
        ("純建物建坪 30.48 坪", VALUATION_FIGURES["house_ping"]),
        ("主建物面積 25.43 坪", VALUATION_FIGURES["main_ping"]),
        ("附屬建物面積 2.98 坪", VALUATION_FIGURES["aux_ping"]),
        ("室內主附合計 28.41 坪", VALUATION_FIGURES["indoor_ping"]),
        ("共用部分公設 2.06 坪", VALUATION_FIGURES["common_ping"]),
        ("公設比約 6.76% (或 6.1% ~ 6.8%)", VALUATION_FIGURES["public_ratio"]),
        ("車位面積 3.27 坪", VALUATION_FIGURES["parking_ping"]),
        ("車位型態規格 (B1-89 機械車位)", VALUATION_FIGURES["parking_type"]),
        ("前次交易日期鐵證 (2024-01-24 / 113年1月)", VALUATION_FIGURES["smoking_gun_date"]),
        ("前次交易取得成本 870 萬", VALUATION_FIGURES["smoking_gun_price"]),
        ("車位拆算開價淨單價 30.45 萬/坪", VALUATION_FIGURES["asking_net_unit_price"]),
        ("保守底標出價 (810 萬 / 800~820 萬 / 24.28 萬)", VALUATION_FIGURES["conservative_tier"]),
        ("合理目標出價 (870 萬 / 860~880 萬 / 26.25 萬)", VALUATION_FIGURES["target_tier"]),
        ("上限天花板出價 (920 萬 / 910~920 萬 / 27.89 萬)", VALUATION_FIGURES["ceiling_tier"]),
        ("南區核心競品社區 (佛羅里達/家綠寶/歌德堡/公園苑等)", VALUATION_FIGURES["competitors"])
    ]

    for doc_type, content in text_corpus.items():
        doc_missing = []
        for label, keywords in check_definitions:
            matched = any(kw in content for kw in keywords)
            if not matched:
                doc_missing.append(label)

        if doc_missing:
            res.fail(f"[{doc_type}] 缺少或未吻合核心關鍵數據指標：{', '.join(doc_missing)}")
        else:
            res.ok(f"[{doc_type}] 100% 吻合全部 18 項核心估價規格與行情數據指標")

    return res


# ==============================================================================
# Tier 4: Visual Assets & Git Sync Readiness
# ==============================================================================
def verify_tier4(paths: Dict[str, str], progressive: bool = False, verbose: bool = False) -> TestResult:
    """
    Tier 4: Validates visual asset existence, PIL forensics, and Git sync readiness:
      - Main chart: 南區25年以上大型社區走勢圖/17_比佛利山莊_成交行情走勢圖.png
      - PIL image integrity (non-corrupt, resolution >= 2000x1500)
      - Competitor / market overview charts existence
      - scripts/sync_to_github.py contains "比佛利山莊13樓" in directory list
      - Embedded visuals inside DOCX & PPTX (if files exist)
    """
    res = TestResult("Tier 4: Visual Assets & Git Sync Readiness (圖表資產與 GitHub 同步就緒度)")

    # 1. Beverly Hills Primary Chart
    chart_path = paths.get("beverly_chart", DEFAULT_BEVERLY_CHART_PATH)
    if os.path.exists(chart_path):
        res.ok(f"比佛利山莊行情走勢圖檔案存在：`{os.path.basename(chart_path)}`")
        if Image is not None:
            try:
                with Image.open(chart_path) as img:
                    w, h = img.size
                    res.details["chart_resolution"] = {"width": w, "height": h}
                    if w >= 2000 and h >= 1500:
                        res.ok(f"走勢圖解析度達標且無損：{w} x {h} (>= 2000x1500 高清門檻)")
                    else:
                        res.fail(f"走勢圖解析度過低：{w} x {h} (< 2000x1500 門檻)")
            except Exception as e:
                res.fail(f"走勢圖損毀無法開啟：{str(e)}")
        else:
            res.warn("PIL (Pillow) 未安裝，略過像素尺寸鑑識")
    else:
        res.fail(f"比佛利山莊行情走勢圖不存在：`{chart_path}`")

    # 2. Market Overview Chart Existence
    comp_missing = []
    for c_path in paths.get("competitor_charts", COMPETITOR_CHART_PATHS):
        if not os.path.exists(c_path):
            comp_missing.append(os.path.basename(c_path))
    if comp_missing:
        res.warn(f"部分參考大盤或競品走勢圖未於預期路徑找到：{', '.join(comp_missing)}")
    else:
        res.ok(f"大盤走勢圖與周邊競品圖表資產均完整存在 ({len(COMPETITOR_CHART_PATHS)} 張)")

    # 3. Git Sync Script Configuration Check
    sync_script_path = paths.get("sync_script", DEFAULT_SYNC_SCRIPT_PATH)
    if os.path.exists(sync_script_path):
        try:
            with open(sync_script_path, "r", encoding="utf-8", errors="ignore") as f:
                sync_content = f.read()

            # Check if "比佛利山莊13樓" is in the folder list
            if '"比佛利山莊13樓"' in sync_content or "'比佛利山莊13樓'" in sync_content:
                res.ok("GitHub 自動同步腳本 (`scripts/sync_to_github.py`) 已包含 `比佛利山莊13樓` 資料夾")
            else:
                if progressive:
                    res.warn("GitHub 自動同步腳本 (`scripts/sync_to_github.py`) 尚未將 `比佛利山莊13樓` 加入同步目錄清單 (Pending M4)")
                else:
                    res.fail("GitHub 自動同步腳本 (`scripts/sync_to_github.py`) 尚未將 `比佛利山莊13樓` 加入同步目錄清單")
        except Exception as e:
            res.fail(f"讀取 GitHub 同步腳本失敗：{str(e)}")
    else:
        res.fail(f"GitHub 自動同步腳本不存在：`{sync_script_path}`")

    # 4. Embedded Visuals in DOCX and PPTX (if files exist)
    docx_path = paths.get("docx", DEFAULT_DOCX_PATH)
    if os.path.exists(docx_path) and docx is not None:
        try:
            _, _, _, _, img_cnt = extract_docx_text_and_stats(docx_path)
            if img_cnt >= 1:
                res.ok(f"Word 文件 (.docx) 內嵌圖表驗證通過：共包含 {img_cnt} 個圖形資產")
            else:
                res.warn("Word 文件 (.docx) 內未偵測到內嵌圖片資產")
        except Exception:
            pass

    pptx_path = paths.get("pptx", DEFAULT_PPTX_PATH)
    if os.path.exists(pptx_path) and pptx is not None:
        try:
            _, _, _, _, pic_cnt = extract_pptx_text_and_stats(pptx_path)
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
      TC2: Missing deliverable (e.g. missing PPTX) is caught and failed by Tier 1.
      TC3: Insufficient file size (e.g. DOCX < 30KB) is caught and failed by Tier 1.
      TC4: Insufficient DOCX tables (< 5) is caught and failed by Tier 2.
      TC5: Non-16:9 PPTX ratio or insufficient slides (< 10) is caught and failed by Tier 2.
      TC6: Missing required section title in MD is caught and failed by Tier 2.
      TC7: Master JSON missing required root schema key is caught and failed by Tier 2.
      TC8: Target property numerical mismatch in JSON is caught and failed by Tier 3.
      TC9: Pricing tier mismatch is caught and failed by Tier 3.
      TC10: Missing primary chart or resolution < 2000x1500 is caught and failed by Tier 4.
      TC11: Missing '比佛利山莊13樓' in sync script is caught and failed by Tier 4.
    """
    print("\n" + "=" * 80)
    print("【執行 verify_beverly_suite.py 內部單元自我測試 (Self-Test Suite)】")
    print("=" * 80)

    all_passed = True
    temp_dir = tempfile.mkdtemp(prefix="beverly_test_suite_")

    try:
        # 1. Setup Mock Fixtures
        mock_bdir = os.path.join(temp_dir, "比佛利山莊13樓")
        mock_sdir = os.path.join(temp_dir, "scratch")
        mock_cdir = os.path.join(temp_dir, "南區25年以上大型社區走勢圖")
        mock_scdir = os.path.join(temp_dir, "scripts")
        os.makedirs(mock_bdir, exist_ok=True)
        os.makedirs(mock_sdir, exist_ok=True)
        os.makedirs(mock_cdir, exist_ok=True)
        os.makedirs(mock_scdir, exist_ok=True)

        mock_paths = {
            "md": os.path.join(mock_bdir, "比佛利山莊13樓_特定待售戶別出價分析與社區整體行情報告.md"),
            "docx": os.path.join(mock_bdir, "比佛利山莊13樓_特定待售戶別出價分析與社區整體行情報告.docx"),
            "pptx": os.path.join(mock_bdir, "比佛利山莊13樓_特定待售戶別出價分析與社區整體行情報告.pptx"),
            "json": os.path.join(mock_sdir, "beverly_valuation_master.json"),
            "sync_script": os.path.join(mock_scdir, "sync_to_github.py"),
            "beverly_chart": os.path.join(mock_cdir, "17_比佛利山莊_成交行情走勢圖.png"),
            "competitor_charts": [
                os.path.join(mock_cdir, "00_台中市南區25年以上大型社區_大盤匯總走勢圖.png"),
                os.path.join(mock_cdir, "29_歌德堡_成交行情走勢圖.png"),
            ]
        }

        # Create valid mock image (>2000x1500 resolution, e.g. 2925x2035, with noise to prevent empty zip compression)
        if Image:
            img = Image.new("RGB", (2925, 2035), color=(27, 54, 93))
            noise = Image.frombytes("RGB", (400, 400), os.urandom(400 * 400 * 3))
            img.paste(noise, (100, 100))
            img.save(mock_paths["beverly_chart"])
            for c_path in mock_paths["competitor_charts"]:
                img.save(c_path)
        else:
            with open(mock_paths["beverly_chart"], "wb") as f:
                f.write(b"\x89PNG\r\n\x1a\n" + b"\x00" * 400000)
            for c_path in mock_paths["competitor_charts"]:
                with open(c_path, "wb") as f:
                    f.write(b"\x89PNG\r\n\x1a\n" + b"\x00" * 400000)

        # Create valid mock sync script containing "比佛利山莊13樓"
        with open(mock_paths["sync_script"], "w", encoding="utf-8") as f:
            f.write('#!/usr/bin/env python3\nfolders = ["大墩人文", "歌德堡", "比佛利山莊13樓", "長億南園8樓"]\n')

        # Create valid mock JSON (> 5KB)
        valid_json_obj = {
            "target_property": {
                "address": "台中市南區忠明南路730巷42號",
                "floor": "13F/14F",
                "asking_price": 998.0,
                "total_area": 33.75,
                "house_area": 30.48,
                "main_area": 25.43,
                "attached_area": 2.98,
                "indoor_area": 28.41,
                "public_area": 2.06,
                "public_ratio": 6.76,
                "parking_area": 3.27,
                "parking_type": "坡道機械式 (B1-89)",
                "layout": "3房2廳2衛1陽台",
                "building_age": 34.0,
                "previous_sale": {
                    "date": "2024-01-24",
                    "price": 870.0,
                    "unit_price": 25.8
                }
            },
            "valuation_model": {
                "parking_benchmark": 70.0,
                "asking_net_unit_price": 30.45,
                "conservative": {"total": 810.0, "net_unit_price": 24.28, "discount": 0.81},
                "target": {"total": 870.0, "net_unit_price": 26.25, "discount": 0.87},
                "ceiling": {"total": 920.0, "net_unit_price": 27.89, "discount": 0.92}
            },
            "community_history": {
                "total_records": 158,
                "clean_residential": 124,
                "storefronts": 16,
                "special_transactions": 20
            },
            "comparables": [
                {"name": "佛羅里達", "price": 26.5},
                {"name": "家綠寶", "price": 25.8},
                {"name": "歌德堡", "price": 26.65},
                {"name": "公園苑", "price": 31.08}
            ],
            "padding_data": "x" * 6000
        }
        with open(mock_paths["json"], "w", encoding="utf-8") as f:
            json.dump(valid_json_obj, f, ensure_ascii=False, indent=2)

        all_keys_text = "13F 13樓 次頂樓 998 33.75 30.48 25.43 2.98 28.41 2.06 6.76 3.27 B1-89 機械 2024-01-24 870 70 30.45 810 24.28 870 26.25 920 27.89 佛羅里達 家綠寶 歌德堡 公園苑"

        # Create valid mock MD (> 15KB)
        md_text = f"""# 比佛利山莊13樓出價分析與社區行情報告
## 一、物件現況與產權調閱
標的位於13F (13樓次頂樓)。總建坪33.75坪，純建物30.48坪 (主建物25.43坪, 附屬建物2.98坪, 室內主附合計28.41坪, 公設2.06坪, 公設比6.76%)。
車位配置：3.27坪，坡道機械車位編號B1-89。開價998萬，換算毛單價29.57萬/坪。
產調重大發現：屋主於2024-01-24以870萬元購入，本次開價998萬意圖短期套利。

## 二、社區歷程與歷史實價登錄行情
比佛利山莊全社區清洗排除非典型交易與店面後，歷史純住宅交易穩定。

## 三、競品比價分析
周邊比價包含佛羅里達、家綠寶、歌德堡、公園苑等南區老牌大型社區。

## 四、車位拆算與淨建坪單價推導
南區老牌社區機械車位公允市價為70萬，拆除車位後淨房屋總價928萬，淨建坪單價為30.45萬/坪。

## 五、出價三道防線精算
- 保守底標價：810萬 (淨單價24.28萬/坪)
- 合理目標價：870萬 (淨單價26.25萬/坪，嚴格錨定屋主買進價870萬)
- 上限天花板價：920萬 (淨單價27.89萬/坪)

## 六、談判話術與議價說帖
以屋主持有2年買進成本870萬為定錨鐵證，並以機械車位抗性切入。
""" + ("\n詳細分析補充文字..." * 500)
        with open(mock_paths["md"], "w", encoding="utf-8") as f:
            f.write(md_text)

        # Create valid mock DOCX (> 30KB)
        if docx:
            doc = docx.Document()
            doc.add_heading("比佛利山莊13樓估價分析報告", 0)
            doc.add_paragraph(all_keys_text)
            for _ in range(6):
                tbl = doc.add_table(rows=3, cols=4)
                for row in tbl.rows:
                    for cell in row.cells:
                        cell.text = all_keys_text
            # Add embedded image
            if os.path.exists(mock_paths["beverly_chart"]):
                doc.add_picture(mock_paths["beverly_chart"])
            # Add padding text to reach size threshold
            for _ in range(50):
                doc.add_paragraph("比佛利山莊13樓特定待售戶出價分析與社區整體行情報告 " * 30)
            doc.save(mock_paths["docx"])

        # Create valid mock PPTX (> 200KB)
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
                # Add picture shape to simulate visual embedding
                if os.path.exists(mock_paths["beverly_chart"]):
                    slide.shapes.add_picture(mock_paths["beverly_chart"], Inches(1), Inches(1), width=Inches(4))
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
            if not t1.passed:
                print(f"    T1 Errors: {t1.errors}")
            if not t2.passed:
                print(f"    T2 Errors: {t2.errors}")
            if not t3.passed:
                print(f"    T3 Errors: {t3.errors}")
            if not t4.passed:
                print(f"    T4 Errors: {t4.errors}")
            all_passed = False

        # Test Case 2: Tier 1 catches missing deliverable
        bad_paths = dict(mock_paths)
        bad_paths["pptx"] = os.path.join(mock_bdir, "non_existent.pptx")
        t1_bad = verify_tier1(bad_paths)
        if not t1_bad.passed:
            print("  [PASS] TC2: Missing PPTX deliverable correctly caught and failed by Tier 1.")
        else:
            print("  [FAIL] TC2: Tier 1 failed to catch missing PPTX!")
            all_passed = False

        # Test Case 3: Tier 1 catches undersized deliverable
        bad_paths_size = dict(mock_paths)
        undersized_docx = os.path.join(mock_bdir, "undersized.docx")
        with open(undersized_docx, "wb") as f:
            f.write(b"PK\x03\x04" + b"\x00" * 1000)
        bad_paths_size["docx"] = undersized_docx
        t1_bad_size = verify_tier1(bad_paths_size)
        if not t1_bad_size.passed:
            print("  [PASS] TC3: Undersized DOCX file correctly caught and failed by Tier 1.")
        else:
            print("  [FAIL] TC3: Tier 1 failed to catch undersized DOCX!")
            all_passed = False

        # Test Case 4: Tier 2 catches insufficient DOCX tables
        if docx:
            bad_docx_path = os.path.join(mock_bdir, "bad_tables.docx")
            doc_bad = docx.Document()
            doc_bad.add_paragraph("Simple paragraph")
            doc_bad.add_table(rows=2, cols=2)  # only 1 table (< 5)
            doc_bad.save(bad_docx_path)
            bad_paths2 = dict(mock_paths)
            bad_paths2["docx"] = bad_docx_path
            t2_bad = verify_tier2(bad_paths2)
            if not t2_bad.passed:
                print("  [PASS] TC4: Insufficient DOCX table count correctly caught and failed by Tier 2.")
            else:
                print("  [FAIL] TC4: Tier 2 failed to catch insufficient DOCX table count!")
                all_passed = False

        # Test Case 5: Tier 2 catches non-16:9 PPTX
        if pptx and Presentation:
            from pptx.util import Inches
            bad_pptx_path = os.path.join(mock_bdir, "bad_ratio.pptx")
            prs_bad = Presentation()
            prs_bad.slide_width = Inches(10)  # 4:3 ratio (10 x 7.5)
            prs_bad.slide_height = Inches(7.5)
            for _ in range(12):
                prs_bad.slides.add_slide(prs_bad.slide_layouts[6])
            prs_bad.save(bad_pptx_path)
            bad_paths3 = dict(mock_paths)
            bad_paths3["pptx"] = bad_pptx_path
            t2_bad_pptx = verify_tier2(bad_paths3)
            if not t2_bad_pptx.passed:
                print("  [PASS] TC5: Non-16:9 PPTX aspect ratio correctly caught and failed by Tier 2.")
            else:
                print("  [FAIL] TC5: Tier 2 failed to catch non-16:9 PPTX ratio!")
                all_passed = False

        # Test Case 6: Tier 2 catches missing MD section title
        bad_md_path = os.path.join(mock_bdir, "bad_sections.md")
        with open(bad_md_path, "w", encoding="utf-8") as f:
            f.write("# 簡略報告\n此處缺少車位拆算與談判話術章節標題..." + ("x" * 20000))
        bad_paths4 = dict(mock_paths)
        bad_paths4["md"] = bad_md_path
        t2_bad_md = verify_tier2(bad_paths4)
        if not t2_bad_md.passed:
            print("  [PASS] TC6: Missing required MD section headers correctly caught and failed by Tier 2.")
        else:
            print("  [FAIL] TC6: Tier 2 failed to catch missing MD sections!")
            all_passed = False

        # Test Case 7: Tier 2 catches missing JSON schema key
        bad_json_path = os.path.join(mock_sdir, "bad_schema.json")
        bad_json_obj = dict(valid_json_obj)
        del bad_json_obj["valuation_model"]
        with open(bad_json_path, "w", encoding="utf-8") as f:
            json.dump(bad_json_obj, f, ensure_ascii=False)
        bad_paths_schema = dict(mock_paths)
        bad_paths_schema["json"] = bad_json_path
        t2_bad_schema = verify_tier2(bad_paths_schema)
        if not t2_bad_schema.passed:
            print("  [PASS] TC7: Master JSON missing schema key 'valuation_model' correctly caught by Tier 2.")
        else:
            print("  [FAIL] TC7: Tier 2 failed to catch missing JSON schema key!")
            all_passed = False

        # Test Case 8: Tier 3 catches asking price mismatch in JSON
        bad_json_price_path = os.path.join(mock_sdir, "bad_price.json")
        bad_json_price_obj = dict(valid_json_obj)
        bad_json_price_obj["target_property"] = dict(valid_json_obj["target_property"])
        bad_json_price_obj["target_property"]["asking_price"] = 1200.0  # mismatch from 998
        with open(bad_json_price_path, "w", encoding="utf-8") as f:
            json.dump(bad_json_price_obj, f, ensure_ascii=False)
        bad_paths_price = dict(mock_paths)
        bad_paths_price["json"] = bad_json_price_path
        t3_bad_price = verify_tier3(bad_paths_price)
        if not t3_bad_price.passed:
            print("  [PASS] TC8: Target property asking price mismatch correctly caught and failed by Tier 3.")
        else:
            print("  [FAIL] TC8: Tier 3 failed to catch asking price mismatch!")
            all_passed = False

        # Test Case 9: Tier 3 catches pricing tier mismatch in JSON
        bad_json_path5 = os.path.join(mock_sdir, "bad_valuation.json")
        bad_json_obj5 = dict(valid_json_obj)
        bad_json_obj5["valuation_model"] = dict(valid_json_obj["valuation_model"])
        bad_json_obj5["valuation_model"]["conservative"] = {"total": 750.0}  # should be 800~825
        with open(bad_json_path5, "w", encoding="utf-8") as f:
            json.dump(bad_json_obj5, f, ensure_ascii=False)
        bad_paths5 = dict(mock_paths)
        bad_paths5["json"] = bad_json_path5
        t3_bad = verify_tier3(bad_paths5)
        if not t3_bad.passed:
            print("  [PASS] TC9: Valuation pricing tier mismatch correctly caught and failed by Tier 3.")
        else:
            print("  [FAIL] TC9: Tier 3 failed to catch pricing mismatch!")
            all_passed = False

        # Test Case 10: Tier 4 catches missing or low-res primary chart
        bad_paths6 = dict(mock_paths)
        bad_paths6["beverly_chart"] = os.path.join(mock_cdir, "non_existent_chart.png")
        t4_bad = verify_tier4(bad_paths6)
        if not t4_bad.passed:
            print("  [PASS] TC10: Missing primary chart image correctly caught and failed by Tier 4.")
        else:
            print("  [FAIL] TC10: Tier 4 failed to catch missing primary chart!")
            all_passed = False

        # Test Case 11: Tier 4 catches sync script missing "比佛利山莊13樓"
        bad_sync_path = os.path.join(mock_scdir, "bad_sync.py")
        with open(bad_sync_path, "w", encoding="utf-8") as f:
            f.write('#!/usr/bin/env python3\nfolders = ["大墩人文", "長億南園8樓"]\n')
        bad_paths7 = dict(mock_paths)
        bad_paths7["sync_script"] = bad_sync_path
        t4_bad_sync = verify_tier4(bad_paths7)
        if not t4_bad_sync.passed:
            print("  [PASS] TC11: Sync script missing '比佛利山莊13樓' correctly caught and failed by Tier 4.")
        else:
            print("  [FAIL] TC11: Tier 4 failed to catch missing '比佛利山莊13樓' in sync script!")
            all_passed = False

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    print("=" * 80)
    if all_passed:
        print("🎉 [SELF-TEST SUCCESS] All 11/11 internal verification tests passed successfully!")
    else:
        print("❌ [SELF-TEST FAILURE] One or more self-test assertions failed!")
    print("=" * 80 + "\n")
    return all_passed


# ==============================================================================
# Main Runner & CLI Execution
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(
        description="Automated 4-Tier Verification Suite for Taichung South District 'Beverly Hills 13F' (比佛利山莊13樓) Deliverables"
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
        "beverly_chart": DEFAULT_BEVERLY_CHART_PATH,
        "competitor_charts": COMPETITOR_CHART_PATHS
    }

    mode_str = "PROGRESSIVE (階段演進模式)" if args.progressive else "STRICT FULL ACCEPTANCE (全量嚴格驗收)"
    print("=" * 80)
    print("【台中市南區比佛利山莊13樓特定待售戶出價分析與社區行情 — 自動化驗證套件】")
    print(f"Project Root : {PROJECT_ROOT}")
    print(f"Target Dir   : {os.path.join(PROJECT_ROOT, '比佛利山莊13樓')}")
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

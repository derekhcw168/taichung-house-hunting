#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scripts/verify_goteborg_suite.py
Automated Verification Suite for Taichung South District "Goteborg" (歌德堡)
Valuation and Community Market Report Deliverables.

Four Verification Tiers:
  - Tier 1: File Existence & Basic Checks (MD >=15KB, DOCX >=30KB, PPTX >=300KB, JSON >=5KB)
  - Tier 2: Format & Structural Integrity (DOCX tables/cells, PPTX 16:9 slides, MD required sections)
  - Tier 3: Data Consistency Across Artifacts (Floor 8F, 29.98坪, 988萬, 810/870-885/930萬, 83 residential, 11 atypical, competitors)
  - Tier 4: Visual Assets & Git Sync Readiness (PNG chart existence, PIL forensics, sync_to_github.py inclusion)

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
DEFAULT_MD_PATH = os.path.join(PROJECT_ROOT, "歌德堡", "歌德堡_特定待售戶別出價分析與社區整體行情報告.md")
DEFAULT_DOCX_PATH = os.path.join(PROJECT_ROOT, "歌德堡", "歌德堡_特定待售戶別出價分析與社區整體行情報告.docx")
DEFAULT_PPTX_PATH = os.path.join(PROJECT_ROOT, "歌德堡", "歌德堡_特定待售戶別出價分析與社區整體行情報告.pptx")
DEFAULT_JSON_PATH = os.path.join(PROJECT_ROOT, "scratch", "goteborg_valuation_master.json")
DEFAULT_SYNC_SCRIPT_PATH = os.path.join(PROJECT_ROOT, "scripts", "sync_to_github.py")

# Visual Assets Paths
DEFAULT_GOTEBORG_CHART_PATH = os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "29_歌德堡_成交行情走勢圖.png")
COMPETITOR_CHART_PATHS = [
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "20_公園苑_成交行情走勢圖.png"),
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "34_椰城_成交行情走勢圖.png"),
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "28_長億南園_成交行情走勢圖.png"),
    os.path.join(PROJECT_ROOT, "南區25年以上大型社區走勢圖", "00_台中市南區25年以上大型社區_大盤匯總走勢圖.png"),
]

# Required Size Thresholds (bytes)
SIZE_THRESHOLDS = {
    "md": 15 * 1024,      # 15 KB
    "docx": 30 * 1024,    # 30 KB
    "pptx": 300 * 1024,   # 300 KB
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

# Required Canonical Valuation Figures
VALUATION_FIGURES = {
    "floor": ["8F", "8樓", "次頂樓", "8/9F"],
    "total_ping": ["29.98"],
    "main_ping": ["20.03"],
    "aux_ping": ["2.04"],
    "common_ping": ["7.91"],
    "public_ratio": ["26.38", "26.4"],
    "car_space": ["無汽車車位", "無車位", "B1-59"],
    "asking_price": ["988"],
    "asking_unit_price": ["32.96"],
    "conservative_tier": ["810", "27.0"],
    "reasonable_tier": ["870", "885", "29.0", "29.5"],
    "ceiling_tier": ["930", "31.0"],
    "transactions_count": ["83"],
    "atypical_count": ["11"],
    "price_2024": ["31.13"],
    "price_2026": ["26.65"],
    "price_drop": ["14.4", "-14.4"],
    "competitors": ["公園苑", "椰城", "長億南園", "廣三大時代"]
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
      - DOCX: parsable, >= 5 tables, >= 50 total cells, contains paragraphs
      - PPTX: parsable, slide count >= 10, 16:9 widescreen (13.333" x 7.5")
      - MD: contains required section headers (物件現況, 社區歷程, 競品比價, 車位拆算, 出價三道防線, 談判話術)
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

                # Verify 16:9 dimensions (13.333" x 7.5" with 0.1" tolerance)
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

    return res


# ==============================================================================
# Tier 3: Data Consistency Across Artifacts
# ==============================================================================
def verify_tier3(paths: Dict[str, str], progressive: bool = False, verbose: bool = False) -> TestResult:
    """
    Tier 3: Cross-verifies key valuation figures and entities across MD, DOCX, PPTX, and JSON:
      - Target floor: 8F (次頂樓)
      - Area: 29.98 坪 (主建 20.03, 附屬 2.04, 公設 7.91, 公設比 26.38%)
      - Car space: 無汽車車位 (機車位 B1-59)
      - Asking price: 988 萬 (32.96 萬/坪)
      - Pricing tiers: 保守 810 萬 (27.0 萬/坪), 合理 870~885 萬 (29.0~29.5 萬/坪), 上限 930 萬 (31.0 萬/坪)
      - Historical: 83 筆純住宅成交 (排除 11 筆非典型), 2024 年均價 31.13 萬, 2026 年最新均價 26.65 萬 (-14.4%)
      - Competitors: 公園苑, 椰城, 長億南園, 廣三大時代
    """
    res = TestResult("Tier 3: Data Consistency Across Artifacts (跨交付物核心數據一致性)")

    # 1. Master JSON Consistency
    json_path = paths.get("json", DEFAULT_JSON_PATH)
    json_data = extract_json_data(json_path)

    if json_data:
        res.ok(f"成功讀取主資料庫 JSON：`{os.path.basename(json_path)}`")
        # Validate JSON keys & fields
        tp = json_data.get("target_property", {})
        pt = json_data.get("pricing_tiers", {})
        hs = json_data.get("historical_summary", {})
        comps = json_data.get("comparables", json_data.get("comparable_communities", []))

        # Check Property figures in JSON
        if tp:
            floor_str = str(tp.get("floor", ""))
            if "8" in floor_str:
                res.ok(f"JSON 目標物件樓層正確：{floor_str}")
            else:
                res.fail(f"JSON 目標物件樓層錯誤：{floor_str} (應包含 8F / 8樓)")

            total_ping = float(tp.get("total_ping", 0))
            if abs(total_ping - 29.98) < 0.05:
                res.ok(f"JSON 總坪數正確：{total_ping} 坪")
            else:
                res.fail(f"JSON 總坪數不吻合：{total_ping} (預期 29.98)")

            asking_price = float(tp.get("asking_price", 0))
            if abs(asking_price - 988.0) < 1.0:
                res.ok(f"JSON 開價總額正確：{asking_price} 萬")
            else:
                res.fail(f"JSON 開價總額不吻合：{asking_price} (預期 988)")

        # Check Pricing Tiers in JSON
        if pt:
            cons = pt.get("conservative_floor", {})
            cons_val = cons.get("total", cons.get("price", 0)) if isinstance(cons, dict) else cons
            if abs(float(cons_val) - 810.0) < 5.0:
                res.ok(f"JSON 保守底標價正確：{cons_val} 萬 (預期 810 萬)")
            else:
                res.fail(f"JSON 保守底標價不吻合：{cons_val} (預期 810 萬)")

            ceil = pt.get("ceiling_max", {})
            ceil_val = ceil.get("total", ceil.get("price", 0)) if isinstance(ceil, dict) else ceil
            if abs(float(ceil_val) - 930.0) < 5.0:
                res.ok(f"JSON 上限天花板價正確：{ceil_val} 萬 (預期 930 萬)")
            else:
                res.fail(f"JSON 上限天花板價不吻合：{ceil_val} (預期 930 萬)")

        # Check Historical Summary in JSON
        if hs:
            res_cnt = hs.get("residential_count", hs.get("clean_transactions_count", 0))
            if res_cnt == 83:
                res.ok("JSON 純住宅歷史成交筆數正確：83 筆")
            elif res_cnt > 0:
                res.warn(f"JSON 純住宅歷史成交筆數：{res_cnt} 筆 (基準 83 筆)")

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
        ("目標樓層 8F / 次頂樓", VALUATION_FIGURES["floor"]),
        ("權狀面積 29.98 坪", VALUATION_FIGURES["total_ping"]),
        ("開價總額 988 萬", VALUATION_FIGURES["asking_price"]),
        ("開價單價 32.96 萬/坪", VALUATION_FIGURES["asking_unit_price"]),
        ("保守底標價 810 萬", ["810"]),
        ("上限天花板價 930 萬", ["930"]),
        ("純住宅成交 83 筆", ["83"]),
        ("排除非典型 11 筆", ["11"]),
        ("2024 年均價 31.13 萬", ["31.13"]),
        ("2026 年最新均價 26.65 萬", ["26.65"]),
        ("周邊核心競品 (公園苑/椰城/長億南園/廣三大時代)", VALUATION_FIGURES["competitors"])
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
      - Main chart: 南區25年以上大型社區走勢圖/29_歌德堡_成交行情走勢圖.png
      - PIL image integrity (non-corrupt, resolution >= 1000x600)
      - Competitor charts existence
      - scripts/sync_to_github.py contains "歌德堡" in directory list
      - Embedded visuals inside DOCX & PPTX (if available)
    """
    res = TestResult("Tier 4: Visual Assets & Git Sync Readiness (圖表資產與 GitHub 同步就緒度)")

    # 1. Goteborg Primary Chart
    chart_path = paths.get("goteborg_chart", DEFAULT_GOTEBORG_CHART_PATH)
    if os.path.exists(chart_path):
        res.ok(f"歌德堡行情走勢圖檔案存在：`{os.path.basename(chart_path)}`")
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
        res.fail(f"歌德堡行情走勢圖不存在：`{chart_path}`")

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

            # Check if "歌德堡" is in the folder list
            if '"歌德堡"' in sync_content or "'歌德堡'" in sync_content:
                res.ok("GitHub 自動同步腳本 (`scripts/sync_to_github.py`) 已包含 `歌德堡` 資料夾")
            else:
                if progressive:
                    res.warn("GitHub 自動同步腳本 (`scripts/sync_to_github.py`) 尚未將 `歌德堡` 加入同步目錄清單 (Pending M4)")
                else:
                    res.fail("GitHub 自動同步腳本 (`scripts/sync_to_github.py`) 尚未將 `歌德堡` 加入同步目錄清單")
        except Exception as e:
            res.fail(f"讀取 GitHub 同步腳本失敗：{str(e)}")
    else:
        res.fail(f"GitHub 自動同步腳本不存在：`{sync_script_path}`")

    # 4. Embedded Visuals in DOCX and PPTX (if files exist)
    docx_path = paths.get("docx", DEFAULT_DOCX_PATH)
    if os.path.exists(docx_path) and docx is not None:
        try:
            doc = docx.Document(docx_path)
            # Check for embedded image parts
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
      TC8: Missing '歌德堡' in sync script is caught and failed by Tier 4.
    """
    print("\n" + "=" * 80)
    print("【執行 verify_goteborg_suite.py 內部單元自我測試 (Self-Test Suite)】")
    print("=" * 80)

    all_passed = True
    temp_dir = tempfile.mkdtemp(prefix="goteborg_test_suite_")

    try:
        # 1. Setup Mock Fixtures
        mock_gdir = os.path.join(temp_dir, "歌德堡")
        mock_sdir = os.path.join(temp_dir, "scratch")
        mock_cdir = os.path.join(temp_dir, "南區25年以上大型社區走勢圖")
        mock_scdir = os.path.join(temp_dir, "scripts")
        os.makedirs(mock_gdir, exist_ok=True)
        os.makedirs(mock_sdir, exist_ok=True)
        os.makedirs(mock_cdir, exist_ok=True)
        os.makedirs(mock_scdir, exist_ok=True)

        mock_paths = {
            "md": os.path.join(mock_gdir, "歌德堡_特定待售戶別出價分析與社區整體行情報告.md"),
            "docx": os.path.join(mock_gdir, "歌德堡_特定待售戶別出價分析與社區整體行情報告.docx"),
            "pptx": os.path.join(mock_gdir, "歌德堡_特定待售戶別出價分析與社區整體行情報告.pptx"),
            "json": os.path.join(mock_sdir, "goteborg_valuation_master.json"),
            "sync_script": os.path.join(mock_scdir, "sync_to_github.py"),
            "goteborg_chart": os.path.join(mock_cdir, "29_歌德堡_成交行情走勢圖.png"),
            "competitor_charts": [
                os.path.join(mock_cdir, "20_公園苑_成交行情走勢圖.png"),
                os.path.join(mock_cdir, "34_椰城_成交行情走勢圖.png"),
                os.path.join(mock_cdir, "28_長億南園_成交行情走勢圖.png"),
                os.path.join(mock_cdir, "00_台中市南區25年以上大型社區_大盤匯總走勢圖.png")
            ]
        }

        # Create valid mock image (>300KB and >=1000x600 resolution)
        if Image:
            img = Image.frombytes("RGB", (1000, 600), os.urandom(1000 * 600 * 3))
            img.save(mock_paths["goteborg_chart"])
            for c_path in mock_paths["competitor_charts"]:
                img.save(c_path)
        else:
            with open(mock_paths["goteborg_chart"], "wb") as f:
                f.write(b"\x89PNG\r\n\x1a\n" + b"\x00" * 400000)
            for c_path in mock_paths["competitor_charts"]:
                with open(c_path, "wb") as f:
                    f.write(b"\x89PNG\r\n\x1a\n" + b"\x00" * 400000)

        # Create valid mock sync script
        with open(mock_paths["sync_script"], "w", encoding="utf-8") as f:
            f.write('#!/usr/bin/env python3\nfolders = ["大墩人文", "歌德堡", "長億南園8樓"]\n')

        # Create valid mock JSON (> 5KB)
        valid_json_obj = {
            "target_property": {
                "community_name": "歌德堡",
                "floor": "8F/9F",
                "total_ping": 29.98,
                "main_ping": 20.03,
                "aux_ping": 2.04,
                "common_ping": 7.91,
                "public_ratio": 26.38,
                "car_space_type": "無汽車車位 (機車位 B1-59)",
                "asking_price": 988.0,
                "asking_unit_price": 32.96
            },
            "pricing_tiers": {
                "conservative_floor": {"total": 810.0, "unit": 27.02},
                "reasonable_target": {"total_low": 870.0, "total_high": 885.0, "unit_low": 29.02, "unit_high": 29.52},
                "ceiling_max": {"total": 930.0, "unit": 31.02}
            },
            "historical_summary": {
                "residential_count": 83,
                "atypical_count": 11,
                "avg_2024": 31.13,
                "avg_2026": 26.65,
                "drop_percent": -14.4
            },
            "comparables": [
                {"name": "公園苑", "price": 31.08},
                {"name": "椰城", "price": 30.08},
                {"name": "長億南園", "price": 24.99},
                {"name": "廣三大時代", "price": 25.89}
            ],
            "padding_data": "x" * 6000
        }
        with open(mock_paths["json"], "w", encoding="utf-8") as f:
            json.dump(valid_json_obj, f, ensure_ascii=False, indent=2)

        all_keys_text = "8F 29.98 20.03 2.04 7.91 26.38 988 32.96 810 870 930 83 11 31.13 26.65 公園苑 椰城 長億南園 廣三大時代 B1-59"

        # Create valid mock MD (> 15KB)
        md_text = f"""# 歌德堡出價分析報告
## 一、物件現況與產權調閱
標的位於8F (8樓)，次頂樓。權狀面積29.98坪 (主建20.03坪, 附屬2.04坪, 公設7.91坪, 公設比26.38%)。
車位配置：無汽車車位，專屬機車位B1-59。開價988萬，單價32.96萬/坪。

## 二、社區歷程與歷史實價登錄行情
全社區排除11筆非典型交易，共83筆純住宅成交。
2024年均價31.13萬，2026年最新成交均價26.65萬，波段回檔-14.4%。

## 三、競品比價分析
周邊比價包含公園苑、椰城、長億南園、廣三大時代等同型老牌社區。

## 四、車位拆算與淨建坪單價推導
本案無汽車車位，開價單價即淨單價32.96萬/坪，相對同樓層實登高估。

## 五、出價三道防線精算
- 保守底標價：810萬 (單價27.0萬/坪)
- 合理目標價：870~885萬 (單價29.0~29.5萬/坪)
- 上限天花板價：930萬 (單價31.0萬/坪)

## 六、談判話術與議價說帖
以無車位抗性與2026年房價回檔14.4%為核心話術切入。
""" + ("\n詳細分析補充文字..." * 500)
        with open(mock_paths["md"], "w", encoding="utf-8") as f:
            f.write(md_text)

        # Create valid mock DOCX (> 30KB)
        if docx:
            doc = docx.Document()
            doc.add_heading("歌德堡估價分析報告", 0)
            doc.add_paragraph(all_keys_text)
            for _ in range(6):
                tbl = doc.add_table(rows=3, cols=4)
                for row in tbl.rows:
                    for cell in row.cells:
                        cell.text = all_keys_text
            # Add padding text to reach size threshold
            for _ in range(50):
                doc.add_paragraph("歌德堡特定待售戶出價分析與社區整體行情報告 " * 30)
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
                # Add picture shape to simulate visual embedding
                if os.path.exists(mock_paths["goteborg_chart"]):
                    slide.shapes.add_picture(mock_paths["goteborg_chart"], Inches(1), Inches(1), width=Inches(4))
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
            doc_bad.add_table(rows=2, cols=2) # only 1 table
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
            prs_bad.slide_width = Inches(10)  # 4:3 ratio (10 x 7.5)
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
            "conservative_floor": {"total": 750.0},  # should be 810
            "ceiling_max": {"total": 930.0}
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
        bad_paths6["goteborg_chart"] = os.path.join(mock_cdir, "non_existent_chart.png")
        t4_bad = verify_tier4(bad_paths6)
        if not t4_bad.passed:
            print("  [PASS] TC7: Missing primary chart image correctly caught and failed by Tier 4.")
        else:
            print("  [FAIL] TC7: Tier 4 failed to catch missing primary chart!")
            all_passed = False

        # Test Case 8: Tier 4 catches sync script missing "歌德堡"
        bad_sync_path = os.path.join(mock_scdir, "bad_sync.py")
        with open(bad_sync_path, "w", encoding="utf-8") as f:
            f.write('#!/usr/bin/env python3\nfolders = ["大墩人文", "長億南園8樓"]\n')
        bad_paths7 = dict(mock_paths)
        bad_paths7["sync_script"] = bad_sync_path
        t4_bad_sync = verify_tier4(bad_paths7)
        if not t4_bad_sync.passed:
            print("  [PASS] TC8: Sync script missing '歌德堡' correctly caught and failed by Tier 4.")
        else:
            print("  [FAIL] TC8: Tier 4 failed to catch missing '歌德堡' in sync script!")
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
        description="Automated Verification Suite for Taichung South District 'Goteborg' (歌德堡) Deliverables"
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
        "goteborg_chart": DEFAULT_GOTEBORG_CHART_PATH,
        "competitor_charts": COMPETITOR_CHART_PATHS
    }

    mode_str = "PROGRESSIVE (階段演進模式)" if args.progressive else "STRICT FULL ACCEPTANCE (全量嚴格驗收)"
    print("=" * 80)
    print("【台中市南區歌德堡特定待售戶出價分析與社區行情 — 自動化驗證套件】")
    print(f"Project Root : {PROJECT_ROOT}")
    print(f"Target Dir   : {os.path.join(PROJECT_ROOT, '歌德堡')}")
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

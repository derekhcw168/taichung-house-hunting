#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scripts/verify_bargaining_points.py
Automated Verification Suite for Dadun Renwen Bargaining Points Guide:
「大墩人文/大墩人文屋主議價與殺價理由50條全攻略.md」

Milestone 1 Test Oracle & Verification Suite
Adheres to Tier 1 ~ Tier 4 Assertions defined in explorer_survey_bargain_3/analysis.md
and PROJECT.md Interface Contracts.

Usage:
    python scripts/verify_bargaining_points.py
    python scripts/verify_bargaining_points.py --target <path_to_markdown>
    python scripts/verify_bargaining_points.py --test-self
    python scripts/verify_bargaining_points.py --help

Exit Codes:
    0 = All assertions passed 100%
    1 = Assertion failure or missing targets
"""

import os
import sys
import re
import argparse
import json
import tempfile
from typing import Dict, List, Set, Tuple, Any, Optional

# Ensure UTF-8 output in Windows environments
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

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_TARGET_PATH = os.path.join(PROJECT_ROOT, "大墩人文", "大墩人文屋主議價與殺價理由50條全攻略.md")

# Six canonical categories specified in PROJECT.md and explorer_survey_bargain_3
VALID_CATEGORIES = {
    "屋況瑕疵與基礎翻修成本折抵",
    "社區公設老化與公共管理提撥",
    "實價登錄行情與銀行估價落差",
    "地段環境與周邊抗性",
    "央行信用管制與房貸政策風險",
    "賣方心理攻防與談判讓步戰術"
}

# Local Dadun Renwen project document keywords
LOCAL_DOC_KEYWORDS = [
    "信義", "不動產說明書", "產調", "說明書",
    "盛祥", "報價單", "統包", "基礎工程",
    "樂居", "實價登錄", "實登",
    "說帖", "可以用來殺價的說帖",
    "修繕", "修繕的項目",
    "公同共有", "大墩人文", "11樓"
]

# Authoritative external property experts, channels and sources (10+ required)
AUTHORITATIVE_EXPERTS = [
    "35線上賞屋",
    "地產秘密客",
    "喬王",
    "房產阿北",
    "國際超級房仲",
    "邱愛莉",
    "愛莉",
    "李奕農",
    "徐佳馨",
    "安娜",
    "好房網",
    "591",
    "理財周刊",
    "PTT",
    "樂居",
    "估價師",
    "銀行鑑價"
]

def clean_text(s: str) -> str:
    """Strip extra spaces and markdown symbols for clean length check."""
    return re.sub(r'[\s`*_#]+', '', s)

def parse_guide(content: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, str]], List[str]]:
    """
    Parses the full markdown guide into:
    1. points: List of structured bargaining points (BP-01 ~ BP-XX)
    2. cheatsheet: List of rows in the Cheatsheet table
    3. structural_sections: Detected H1/H2 header titles
    """
    # Extract headers
    headers = re.findall(r'^#{1,3}\s+(.+)$', content, re.MULTILINE)

    # 1. Parse Detailed Bargaining Points Cards
    # Split content by markdown H3 heading matching: ### 理由 XX
    point_blocks = re.split(r'\n(?=###\s+理由\s*\d+)', content)
    points = []

    for blk in point_blocks:
        m_title = re.match(r'###\s+理由\s*(\d+)[:：]\s*【?(.+?)】?\s*$', blk.strip().split('\n')[0])
        if not m_title:
            continue
        num_str = m_title.group(1).zfill(2)
        title_str = m_title.group(2).strip()

        # Extract 6 standard fields
        m_id = re.search(r'-\s*\*\*編號識別\*\*\s*[:：]\s*(BP-\d+)', blk)
        m_cat = re.search(r'-\s*\*\*分類標籤\*\*\s*[:：]\s*[`"\'「]?([^`"\'\r\n]+)[`"\'」]?', blk)
        m_arg = re.search(r'-\s*\*\*核心論點\*\*\s*[:：]\s*(.*?)(?=\n\s*-\s*\*\*|\Z)', blk, re.DOTALL)
        m_scn = re.search(r'-\s*\*\*適用情境\*\*\s*[:：]\s*(.*?)(?=\n\s*-\s*\*\*|\Z)', blk, re.DOTALL)
        m_dlg = re.search(r'-\s*\*\*實戰話術\*\*\s*[:：]\s*(.*?)(?=\n\s*-\s*\*\*|\Z)', blk, re.DOTALL)
        m_src = re.search(r'-\s*\*\*佐證出處\*\*\s*[:：]\s*(.*?)(?=\n###|\n---|\Z)', blk, re.DOTALL)

        points.append({
            "num": num_str,
            "title": title_str,
            "bp_id": m_id.group(1).strip() if m_id else None,
            "category": m_cat.group(1).strip() if m_cat else None,
            "argument": m_arg.group(1).strip() if m_arg else "",
            "scenario": m_scn.group(1).strip() if m_scn else "",
            "dialogue": m_dlg.group(1).strip() if m_dlg else "",
            "source": m_src.group(1).strip() if m_src else "",
            "raw": blk
        })

    # 2. Parse Cheatsheet table
    # Looking for a markdown table containing: 序號, 分類, 痛點, 數據, 台詞/金句
    cheatsheet_rows = []
    # Match markdown tables
    table_blocks = re.findall(r'(\|[^\n]+\|\r?\n\|[-:\s|]+\|\r?\n(?:\|[^\n]+\|\r?\n?)+)', content)
    for tbl in table_blocks:
        lines = [line.strip() for line in tbl.strip().split('\n') if line.strip()]
        if len(lines) < 3:
            continue
        header = lines[0]
        # Check if header matches Cheatsheet columns
        if ("序號" in header or "ID" in header) and "分類" in header and ("痛點" in header or "切入" in header):
            for row in lines[2:]:
                cells = [c.strip() for c in row.split('|')[1:-1]]
                if len(cells) >= 5:
                    clean_id = re.sub(r'[*`\s]', '', cells[0])
                    if clean_id.startswith('BP-'):
                        cheatsheet_rows.append({
                            "id": clean_id,
                            "category": cells[1],
                            "topic": cells[2],
                            "data": cells[3],
                            "quote": cells[4],
                            "raw": row
                        })

    return points, cheatsheet_rows, headers


def verify_tier1(file_path: str, content: str) -> List[str]:
    """Tier 1: File Existence, UTF-8 Encoding, Size, and Major Structure."""
    errors = []
    if not os.path.exists(file_path):
        return [f"[Tier 1] Target file not found: {file_path}"]

    file_size = os.path.getsize(file_path)
    if file_size < 10240:
        errors.append(f"[Tier 1] File size {file_size} bytes is too small (< 10 KB threshold).")

    # Verify key sections presence
    required_sections = [
        ("大墩人文", "Document Title referencing 大墩人文"),
        ("目錄", "Table of Contents"),
        ("速查", "Cheatsheet / Quick Reference section"),
        ("實戰話術", "Detailed Bargaining Points"),
        ("談判階段", "Negotiation Stages / Tactical Playbook"),
        ("特約", "Contract Protection / Terms"),
        ("佐證", "Appendix / Evidence references")
    ]

    for kw, desc in required_sections:
        if kw not in content:
            errors.append(f"[Tier 1] Missing required architectural section: '{desc}' (keyword: '{kw}')")

    return errors


def verify_tier2(points: List[Dict[str, Any]]) -> List[str]:
    """Tier 2: Bargaining Points Count (>= 50), Uniqueness, and 6 Required Fields."""
    errors = []
    total_count = len(points)
    if total_count < 50:
        errors.append(f"[Tier 2] Total points count {total_count} is less than 50 minimum threshold.")

    ids_seen: Set[str] = set()

    for idx, p in enumerate(points, 1):
        bp_id = p.get('bp_id')
        num = p.get('num', str(idx))

        if not bp_id:
            errors.append(f"[Tier 2] Item #{num} ('{p.get('title')}') is missing '編號識別' (BP-XX).")
            continue

        if not re.match(r'^BP-\d{2,}$', bp_id):
            errors.append(f"[Tier 2] Item #{num} has invalid BP-ID format '{bp_id}' (expected BP-XX).")

        if bp_id in ids_seen:
            errors.append(f"[Tier 2] Duplicate BP-ID detected: '{bp_id}' in item #{num}.")
        ids_seen.add(bp_id)

        # Field: 分類標籤
        cat = p.get('category')
        if not cat:
            errors.append(f"[Tier 2] {bp_id}: Missing '分類標籤'.")
        elif cat not in VALID_CATEGORIES:
            errors.append(f"[Tier 2] {bp_id}: Invalid category '{cat}'. Must be one of: {sorted(list(VALID_CATEGORIES))}")

        # Field: 核心論點
        arg = p.get('argument', '')
        arg_clean = clean_text(arg)
        if len(arg_clean) < 30:
            errors.append(f"[Tier 2] {bp_id}: '核心論點' content too short ({len(arg_clean)} chars < 30 char minimum).")

        # Field: 適用情境
        scn = p.get('scenario', '')
        scn_clean = clean_text(scn)
        if len(scn_clean) < 10:
            errors.append(f"[Tier 2] {bp_id}: '適用情境' content too short ({len(scn_clean)} chars < 10 char minimum).")

        # Field: 實戰話術
        dlg = p.get('dialogue', '')
        dlg_clean = clean_text(dlg)
        if len(dlg_clean) < 50:
            errors.append(f"[Tier 2] {bp_id}: '實戰話術' content too short ({len(dlg_clean)} chars < 50 char minimum).")
        if '「' not in dlg or '」' not in dlg:
            errors.append(f"[Tier 2] {bp_id}: '實戰話術' must contain Taiwanese first-person dialogue in quotation marks 「...」.")

        # Field: 佐證出處
        src = p.get('source', '')
        src_clean = clean_text(src)
        if len(src_clean) < 15:
            errors.append(f"[Tier 2] {bp_id}: '佐證出處' content too short ({len(src_clean)} chars < 15 char minimum).")

    return errors


def verify_tier3(points: List[Dict[str, Any]]) -> Tuple[List[str], Dict[str, Any]]:
    """
    Tier 3: Source Duality Coverage.
    Validates coverage of local Dadun Renwen files and external property experts/YouTubers.
    """
    errors = []
    local_count = 0
    external_count = 0
    expert_citations: Dict[str, int] = {expert: 0 for expert in AUTHORITATIVE_EXPERTS}

    for p in points:
        source_text = p.get('source', '')
        full_text = p.get('argument', '') + " " + source_text

        # Check local project citations
        is_local = any(kw in source_text or kw in full_text for kw in LOCAL_DOC_KEYWORDS)
        if is_local:
            local_count += 1

        # Check external expert citations
        cited_any_external = False
        for exp in AUTHORITATIVE_EXPERTS:
            if exp in source_text or exp in full_text:
                expert_citations[exp] += 1
                cited_any_external = True

        if cited_any_external or "YouTube" in source_text or "專家" in source_text:
            external_count += 1

    unique_experts_cited = [exp for exp, cnt in expert_citations.items() if cnt > 0]
    unique_expert_count = len(unique_experts_cited)

    stats = {
        "local_source_points": local_count,
        "external_source_points": external_count,
        "unique_experts_count": unique_expert_count,
        "unique_experts_cited": unique_experts_cited
    }

    if local_count < 15:
        errors.append(f"[Tier 3] Local document coverage insufficient: {local_count} points cite local files (< 15 threshold).")

    if external_count < 25:
        errors.append(f"[Tier 3] External expert coverage insufficient: {external_count} points cite external experts (< 25 threshold).")

    if unique_expert_count < 10:
        errors.append(
            f"[Tier 3] Insufficient distinct authoritative experts cited: {unique_expert_count} < 10 threshold. "
            f"Detected: {unique_experts_cited}"
        )

    return errors, stats


def verify_tier4(cheatsheet: List[Dict[str, str]], point_ids: Set[str]) -> List[str]:
    """Tier 4: Cheatsheet Table Presence, Row Count, and 100% ID Alignment."""
    errors = []
    sheet_count = len(cheatsheet)
    if sheet_count < 50:
        errors.append(f"[Tier 4] Cheatsheet table rows count {sheet_count} is less than 50 minimum threshold.")

    sheet_ids = set()
    for row in cheatsheet:
        rid = row.get('id', '')
        if rid in sheet_ids:
            errors.append(f"[Tier 4] Duplicate ID in Cheatsheet: '{rid}'.")
        sheet_ids.add(rid)

        # Ensure all columns are non-empty
        for col_name in ['category', 'topic', 'data', 'quote']:
            val = row.get(col_name, '').strip()
            if not val or len(val) < 2:
                errors.append(f"[Tier 4] Cheatsheet row '{rid}' column '{col_name}' is blank or too short.")

    # Check 100% mutual alignment
    missing_in_cheatsheet = point_ids - sheet_ids
    if missing_in_cheatsheet:
        errors.append(f"[Tier 4] Points missing from Cheatsheet table: {sorted(list(missing_in_cheatsheet))}")

    extra_in_cheatsheet = sheet_ids - point_ids
    if extra_in_cheatsheet:
        errors.append(f"[Tier 4] Cheatsheet contains IDs not found in detailed points: {sorted(list(extra_in_cheatsheet))}")

    return errors


def run_verification(target_path: str, verbose: bool = False) -> Tuple[bool, Dict[str, Any]]:
    """Runs full verification suite on the specified target file."""
    report: Dict[str, Any] = {
        "target_path": target_path,
        "tier1_errors": [],
        "tier2_errors": [],
        "tier3_errors": [],
        "tier4_errors": [],
        "stats": {},
        "passed": False
    }

    print("=" * 80)
    print("【大墩人文 50+ 條買房殺價理由全攻略 — 獨立自動化驗證套件】")
    print(f"Target File: {target_path}")
    print("=" * 80)

    # 1. Tier 1 Check
    if not os.path.exists(target_path):
        err = f"[Tier 1 FAIL] Target file not found: {target_path}"
        print(f"\n❌ {err}")
        report["tier1_errors"].append(err)
        return False, report

    try:
        with open(target_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        err = f"[Tier 1 FAIL] Failed to read file with UTF-8 encoding: {e}"
        print(f"\n❌ {err}")
        report["tier1_errors"].append(err)
        return False, report

    t1_errors = verify_tier1(target_path, content)
    report["tier1_errors"] = t1_errors
    if t1_errors:
        print(f"\n❌ [Tier 1] Architecture & Encoding check failed ({len(t1_errors)} errors):")
        for e in t1_errors:
            print(f"   • {e}")
    else:
        print(f"✅ [Tier 1 PASS] File exists, UTF-8 encoded, size={os.path.getsize(target_path):,} bytes, structural sections intact.")

    # Parse points and cheatsheet
    points, cheatsheet, headers = parse_guide(content)
    point_ids = {p['bp_id'] for p in points if p.get('bp_id')}

    report["stats"]["points_parsed"] = len(points)
    report["stats"]["cheatsheet_rows_parsed"] = len(cheatsheet)
    print(f"ℹ️  [INFO] Extracted {len(points)} bargaining point cards, {len(cheatsheet)} Cheatsheet rows.")

    # 2. Tier 2 Check
    t2_errors = verify_tier2(points)
    report["tier2_errors"] = t2_errors
    if t2_errors:
        print(f"\n❌ [Tier 2] Quantity & Field Completeness check failed ({len(t2_errors)} errors):")
        for e in t2_errors[:15]:
            print(f"   • {e}")
        if len(t2_errors) > 15:
            print(f"   ... and {len(t2_errors) - 15} more errors.")
    else:
        print(f"✅ [Tier 2 PASS] Found {len(points)} items (>= 50 threshold). 100% valid 6 fields, no duplicates.")

    # 3. Tier 3 Check
    t3_errors, t3_stats = verify_tier3(points)
    report["tier3_errors"] = t3_errors
    report["stats"]["tier3"] = t3_stats
    print(f"ℹ️  [INFO] Citations: Local={t3_stats['local_source_points']}, External={t3_stats['external_source_points']}, Unique Experts={t3_stats['unique_experts_count']}/10")
    if verbose:
        print(f"     Unique experts cited: {', '.join(t3_stats['unique_experts_cited'])}")
    if t3_errors:
        print(f"\n❌ [Tier 3] Source Duality check failed ({len(t3_errors)} errors):")
        for e in t3_errors:
            print(f"   • {e}")
    else:
        print(f"✅ [Tier 3 PASS] Source Duality verified (Local >= 15, External >= 25, Unique Experts >= 10).")

    # 4. Tier 4 Check
    t4_errors = verify_tier4(cheatsheet, point_ids)
    report["tier4_errors"] = t4_errors
    if t4_errors:
        print(f"\n❌ [Tier 4] Cheatsheet Alignment check failed ({len(t4_errors)} errors):")
        for e in t4_errors[:15]:
            print(f"   • {e}")
        if len(t4_errors) > 15:
            print(f"   ... and {len(t4_errors) - 15} more errors.")
    else:
        print(f"✅ [Tier 4 PASS] Cheatsheet ({len(cheatsheet)} rows) 100% aligned with bargaining cards.")

    total_defects = len(t1_errors) + len(t2_errors) + len(t3_errors) + len(t4_errors)
    report["total_defects"] = total_defects
    report["passed"] = (total_defects == 0)

    print("=" * 80)
    if report["passed"]:
        print("🎉【驗收結論：PASS】全 4 個驗證層級 (Tier 1 ~ 4) 100% 驗收通過！符合專案最高交付標準。")
    else:
        print(f"⚠️ 【驗收結論：FAIL】累計檢測到 {total_defects} 個規格缺陷，請依據上方定位訊息進行修訂。")
    print("=" * 80)

    return report["passed"], report


def generate_mock_guide_markdown(
    num_points: int = 52,
    break_field: Optional[str] = None,
    short_dialogue: bool = False,
    few_experts: bool = False,
    table_mismatch: bool = False
) -> str:
    """Helper to generate a synthesized markdown guide for unit tests."""
    categories_list = list(VALID_CATEGORIES)
    lines = [
        "# 大墩人文屋主議價與殺價理由 50 條全攻略（含現場談判速查表與實戰劇本）",
        "",
        "> 物件：大墩人文 11 樓",
        "",
        "## 目錄",
        "1. 核心哲學",
        "2. 速查專區",
        "3. 實戰詳解 52 條買房殺價理由",
        "4. 五大談判階段實戰演練",
        "5. 簽約前五大特約底線",
        "6. 附錄索引與佐證文檔",
        "",
        "## 【速查專區】現場談判快速檢閱速查表",
        "",
        "| 序號 | 分類 | 談判切入痛點 | 關鍵客觀數據 / 依據 | 現場 5 秒殺價台詞 (破冰金句) |",
        "| :---: | :---: | :--- | :--- | :--- |"
    ]

    # Cheatsheet rows
    for i in range(1, num_points + 1):
        if table_mismatch and i == num_points:
            continue  # Miss the last row to test mismatch
        bp_id = f"BP-{str(i).zfill(2)}"
        cat_short = "屋況" if i <= 10 else "行情"
        lines.append(f"| **{bp_id}** | {cat_short} | 痛點說明測試內容{i} | 數據依據測試內容{i} | 「這是現場5秒台詞{i}」 |")

    lines.append("")
    lines.append("## 【實戰詳解】52 條買房殺價理由與實戰對談話術全攻略")
    lines.append("")

    # Card items
    for i in range(1, num_points + 1):
        bp_id = f"BP-{str(i).zfill(2)}"
        cat = categories_list[(i - 1) % len(categories_list)]

        # Determine expert citation
        if few_experts:
            expert = "35線上賞屋"
        else:
            expert = AUTHORITATIVE_EXPERTS[(i - 1) % len(AUTHORITATIVE_EXPERTS)]

        dlg_text = "太短" if short_dialogue else f"「張大哥您好，我們看過大墩人文非常喜歡，但考量到這項修繕剛性支出高達數十萬元，希望屋主能在總價上給予實質折讓，大家圓滿結緣！第{i}條實戰對談話術完整台詞。」"

        lines.append(f"### 理由 {str(i).zfill(2)}：【核心論點主題標題測試項目第 {i} 條】")
        if break_field != "id":
            lines.append(f"- **編號識別**：{bp_id}")
        if break_field != "category":
            lines.append(f"- **分類標籤**：`{cat}`")
        if break_field != "argument":
            lines.append(f"- **核心論點**：這是經過精確客觀分析的核心論點，針對大墩人文社區提供詳盡事實依據與金額百分比折抵計算第 {i} 條。")
        if break_field != "scenario":
            lines.append(f"- **適用情境**：現場看屋或下斡旋談判時發動。")
        if break_field != "dialogue":
            lines.append(f"- **實戰話術**：{dlg_text}")
        if break_field != "source":
            lines.append(f"- **佐證出處**：")
            lines.append(f"  - 專案內部：引用《大墩人文11樓_老屋翻新基礎統包工程報價單.md》與信義房屋不動產說明書第 {i} 頁。")
            lines.append(f"  - 外部專家：引用房產專家與 YouTube《{expert}》影片解析第 {i} 集。")
        lines.append("")

    lines.extend([
        "## 【實戰演練】五大談判階段攻防劇本與走位教學",
        "談判階段說明...",
        "",
        "## 【合約防線】簽約前絕不可妥協的五大特約底線",
        "特約條款說明...",
        "",
        "## 【附錄索引】大墩人文既有佐證文檔與全網專家資料來源庫",
        "佐證文檔清單..."
    ])

    return "\n".join(lines)


def run_self_tests() -> int:
    """Runs built-in unit tests against synthesized markdown inputs to verify test suite accuracy."""
    print("=" * 80)
    print("【執行 verify_bargaining_points.py 內部單元自我測試 (Self-Test Suite)】")
    print("=" * 80)

    test_cases = [
        ("TC1: Valid 52-point guide should pass 100%", {}, True),
        ("TC2: Missing mandatory field (編號識別) must fail Tier 2", {"break_field": "id"}, False),
        ("TC3: Short dialogue (< 50 chars) must fail Tier 2", {"short_dialogue": True}, False),
        ("TC4: Insufficient unique experts (< 10) must fail Tier 3", {"few_experts": True}, False),
        ("TC5: Cheatsheet table ID mismatch must fail Tier 4", {"table_mismatch": True}, False),
        ("TC6: Total points count < 50 must fail Tier 2 & Tier 4", {"num_points": 30}, False),
    ]

    all_passed = True
    for name, kwargs, expected_pass in test_cases:
        content = generate_mock_guide_markdown(**kwargs)
        with tempfile.NamedTemporaryFile('w', encoding='utf-8', suffix='.md', delete=False) as tf:
            tf.write(content)
            temp_path = tf.name

        try:
            passed, report = run_verification(temp_path, verbose=False)
            if passed == expected_pass:
                print(f"  [PASS] {name}")
            else:
                print(f"  [FAIL] {name}: Expected passed={expected_pass}, got passed={passed}")
                all_passed = False
        finally:
            try:
                os.remove(temp_path)
            except Exception:
                pass

    # TC7: Missing file must fail Tier 1
    non_existent_file = os.path.join(tempfile.gettempdir(), "non_existent_guide_xyz123.md")
    passed, _ = run_verification(non_existent_file, verbose=False)
    if not passed:
        print("  [PASS] TC7: Non-existent file correctly rejected by Tier 1")
    else:
        print("  [FAIL] TC7: Non-existent file was not rejected")
        all_passed = False

    print("=" * 80)
    if all_passed:
        print("🎉 [SELF-TEST SUCCESS] All 7 internal verification tests passed successfully!")
        return 0
    else:
        print("❌ [SELF-TEST FAILURE] One or more internal self-tests failed!")
        return 1


def main():
    parser = argparse.ArgumentParser(
        description="Verify Dadun Renwen 50+ Bargaining Points Guide (Tier 1 ~ Tier 4 Assertions)."
    )
    parser.add_argument(
        "--target", "--file",
        dest="target",
        default=DEFAULT_TARGET_PATH,
        help="Path to the markdown guide file to verify (default: 大墩人文/大墩人文屋主議價與殺價理由50條全攻略.md)"
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Print detailed debugging information during verification."
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output JSON summary report."
    )
    parser.add_argument(
        "--test-self",
        action="store_true",
        help="Run internal unit tests to verify oracle test assertion integrity."
    )

    args = parser.parse_args()

    if args.test_self:
        sys.exit(run_self_tests())

    passed, report = run_verification(args.target, verbose=args.verbose)

    if args.json:
        print("\n--- JSON REPORT ---")
        print(json.dumps(report, ensure_ascii=False, indent=2))

    sys.exit(0 if passed else 1)


if __name__ == '__main__':
    main()

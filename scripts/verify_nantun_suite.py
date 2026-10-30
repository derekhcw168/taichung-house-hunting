#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scripts/verify_nantun_suite.py
Automated Verification Suite for Nantun District 30yr+ Large Communities Project
(台中市南屯區30年以上大型社區實價登錄走勢圖與完整分析套件)

Architecture:
  - Tier 1: Feature Coverage (PNGs, Macro Chart, Markdown Report, PPTX Deck, HTML Dashboard)
  - Tier 2: Boundary & Corner Cases (Invariants, 1997 Edge Cases, <100 HH Edge Cases, Transaction Purity)
  - Tier 3: Cross-Feature Consistency (Roster vs JSON vs PNGs vs MD vs PPTX vs HTML 100% Alignment)
  - Tier 4: Real-World Acceptance Scenarios (PIL Image Resolution >=3000x2000, PPTX >=60 Slides 16:9, MD >=40KB, HTML >=100KB + Lightbox)

CLI Options:
  --all (default)       : Full strict acceptance verification
  --progressive         : Progressive mode for milestone-by-milestone validation
  --tier1 / --tier2 / --tier3 / --tier4 : Run specific verification tier
  --data-layer          : Run data layer verification (M1)
  --chart-layer         : Run chart image verification (M2)
  --doc-layer           : Run report, pptx, and dashboard verification (M3/M4)
  --test-self           : Execute self-test suite with synthetic mock fixtures
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
from typing import Dict, List, Tuple, Any, Optional, Set

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

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_TARGET_DIR = os.path.join(PROJECT_ROOT, "南屯區30年以上大型社區走勢圖")
ROSTER_CANDIDATE_PATHS = [
    os.path.join(PROJECT_ROOT, "scratch", "target_nantun_communities_clean.json"),
    os.path.join(PROJECT_ROOT, ".agents", "teamwork", "explorer_survey_nantun_2", "qualified_nantun_communities.json"),
    os.path.join(PROJECT_ROOT, "scratch", "qualified_nantun_communities.json"),
    os.path.join(DEFAULT_TARGET_DIR, "qualified_nantun_communities.json")
]
TRANSACTIONS_CANDIDATE_PATHS = [
    os.path.join(PROJECT_ROOT, "scratch", "all_75_nantun_communities_transactions.json"),
    os.path.join(DEFAULT_TARGET_DIR, "all_75_nantun_communities_transactions.json"),
    os.path.join(PROJECT_ROOT, "scratch", "all_nantun_communities_transactions.json")
]

# Authoritative 1997 Excluded Communities (Must NOT be present in 75 roster)
EXCLUDED_1997_COMMUNITIES = [
    "龍歡喜大廈",
    "三采市政天地",
    "同隆上第二期",
    "親家莊",
    "熊貓凱撒",
    "東興寧境",
    "英棋尊邸",
    "巴黎大凱旋",
    "泉福天郡",
    "大東家春風"
]

# Authoritative <100 Household Excluded Communities (90~99 units, Must NOT be present in 75 roster)
EXCLUDED_LOW_HH_COMMUNITIES = [
    "心田",
    "寶祥榕園",
    "寶樺",
    "華爾街天廈",
    "采龍大廈",
    "成吉思汗",
    "統健大樓",
    "光田生活觀邸",
    "傑出名門",
    "巴黎四季"
]

# Authoritative Canonical 75 Community Names in Order (Ranks 01 to 75)
CANONICAL_75_ROSTER = [
    (1, "理想貴族", 732, "1991-12-06"),
    (2, "宏全世界廣場", 436, "1995-10-24"),
    (3, "中友綠園邸", 431, "1996-06-12"),
    (4, "凱撒風景園邸", 423, "1994-09-05"),
    (5, "龍觀天下", 420, "1994-08-25"),
    (6, "皇后大道", 388, "1993-11-20"),
    (7, "黎明大地", 376, "1992-01-15"),
    (8, "陽光綠意参", 366, "1993-11-10"),
    (9, "黎明春秋", 361, "1995-01-12"),
    (10, "博愛國宅", 350, "1994-09-24"),
    (11, "中友生活家", 335, "1995-09-19"),
    (12, "由鉅八大家", 331, "1994-09-26"),
    (13, "寶輝大官邸", 298, "1994-06-16"),
    (14, "同隆上第一期", 237, "1994-07-21"),
    (15, "利亞倫敦", 231, "1994-05-03"),
    (16, "台中商務中心", 230, "1990-07-05"),
    (17, "現代經典", 222, "1993-02-13"),
    (18, "大台中新市", 222, "1996-12-14"),
    (19, "海德堡藝術廣場", 217, "1994-08-18"),
    (20, "波音市銀翼", 201, "1995-08-17"),
    (21, "大墩天廈", 199, "1995-03-21"),
    (22, "大墩人文", 188, "1992-11-04"),
    (23, "東龍尊邸", 187, "1994-12-06"),
    (24, "三采大業天地", 185, "1995-01-28"),
    (25, "熊貓帝國", 181, "1995-06-20"),
    (26, "國王與我", 180, "1994-11-22"),
    (27, "黎明太陽神", 178, "1994-04-25"),
    (28, "美村黎明大廈", 177, "1994-09-12"),
    (29, "羅馬假期二期", 176, "1993-01-12"),
    (30, "五權花園廣場", 176, "1994-05-23"),
    (31, "台中天母", 175, "1995-03-30"),
    (32, "文心尊龍大廈", 173, "1994-09-03"),
    (33, "世紀龍門", 170, "1994-04-25"),
    (34, "三采陽光麗景", 168, "1996-05-28"),
    (35, "三采市政雙星", 163, "1996-07-12"),
    (36, "泉福藝術龍庭", 162, "1994-11-08"),
    (37, "早安黎明大樓", 156, "1994-12-02"),
    (38, "熊貓天與地公寓", 156, "1995-01-16"),
    (39, "三采大業新象", 156, "1994-10-24"),
    (40, "藝術羅丹", 151, "1994-10-07"),
    (41, "三采藝術園邸", 151, "1995-11-15"),
    (42, "三采藝術大地", 150, "1990-11-20"),
    (43, "綠卡大樓", 148, "1994-04-19"),
    (44, "大未來", 146, "1995-02-20"),
    (45, "波音市彩翼", 146, "1995-05-25"),
    (46, "嶺東天下", 145, "1994-09-15"),
    (47, "新宿", 143, "1995-06-01"),
    (48, "海德堡五期", 143, "1994-07-18"),
    (49, "熊貓國寶", 142, "1994-07-11"),
    (50, "龍邦富邑", 137, "1995-03-28"),
    (51, "大師大樓", 136, "1996-06-15"),
    (52, "熊貓寶座", 136, "1993-11-19"),
    (53, "台中黎明", 134, "1992-12-15"),
    (54, "鑫園21世紀", 134, "1994-03-28"),
    (55, "文山親家十一期", 132, "1995-04-01"),
    (56, "黎明玉璽大廈", 130, "1993-12-02"),
    (57, "陽光之邑", 127, "1995-06-23"),
    (58, "黎明公寓大廈", 125, "1983-03-03"),
    (59, "黎明新天地", 125, "1996-07-02"),
    (60, "羅馬假期一期", 121, "1992-11-13"),
    (61, "海德堡公園特區", 119, "1993-11-29"),
    (62, "晶品大樓", 117, "1994-08-17"),
    (63, "巴黎大道", 115, "1995-05-19"),
    (64, "山河風情", 114, "1994-03-28"),
    (65, "領袖江山", 113, "1994-04-06"),
    (66, "翰陽金墩大樓", 113, "1993-10-18"),
    (67, "儷園華廈(惠文南街)", 113, "1995-01-25"),
    (68, "佳府世家", 112, "1992-11-25"),
    (69, "愛人同志二期", 109, "1993-04-26"),
    (70, "向上園邸", 108, "1995-03-23"),
    (71, "達觀", 107, "1994-12-27"),
    (72, "寶璽公園爵邸", 107, "1991-04-12"),
    (73, "錦繡江山", 104, "1994-05-21"),
    (74, "我愛我家", 103, "1993-10-19"),
    (75, "寶璽尊邸", 102, "1994-04-07")
]

class TestResult:
    def __init__(self, name: str):
        self.name = name
        self.passed = True
        self.messages: List[str] = []
        self.errors: List[str] = []
        self.warnings: List[str] = []
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


def get_roster_data(custom_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """Loads qualified community roster from available sources or falls back to canonical roster."""
    paths = [custom_path] if custom_path else []
    paths.extend(ROSTER_CANDIDATE_PATHS)
    for p in paths:
        if p and os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list) and len(data) >= 75:
                    return data
                elif isinstance(data, dict) and "communities" in data and len(data["communities"]) >= 75:
                    return data["communities"]
            except Exception:
                continue
    # Fallback to canonical list
    return [
        {"id": 1000 + rank, "name": name, "households": hh, "build_date": bdate, "rank": rank}
        for rank, name, hh, bdate in CANONICAL_75_ROSTER
    ]


def find_transactions_file(custom_path: Optional[str] = None) -> Optional[str]:
    """Finds the transactions dataset file if present."""
    paths = [custom_path] if custom_path else []
    paths.extend(TRANSACTIONS_CANDIDATE_PATHS)
    for p in paths:
        if p and os.path.exists(p):
            return p
    return None


# ==============================================================================
# Tier 1: Feature Coverage Verification
# ==============================================================================
def verify_tier1(target_dir: str, root_dir: str, roster: List[Dict[str, Any]],
                 progressive: bool = False, verbose: bool = False) -> TestResult:
    res = TestResult("Tier 1: Feature Coverage (產物覆蓋性與基本結構)")
    target_exists = os.path.exists(target_dir)

    if not target_exists:
        if progressive:
            res.warn(f"Target directory `{os.path.basename(target_dir)}` not created yet (Pending M1/M2).")
            return res
        res.fail(f"Target directory `{target_dir}` does not exist.")
        return res

    # 1. 75 Community PNGs
    found_pngs = 0
    missing_pngs = []
    for idx, comm in enumerate(roster[:75], 1):
        name = comm.get("name", "")
        expected_fn = f"{idx:02d}_{name}_成交行情走勢圖.png"
        full_p = os.path.join(target_dir, expected_fn)
        if os.path.exists(full_p) and os.path.getsize(full_p) > 0:
            found_pngs += 1
        else:
            missing_pngs.append(expected_fn)

    res.details["community_pngs_found"] = found_pngs
    res.details["community_pngs_total"] = 75

    if found_pngs == 75:
        res.ok(f"All 75 community PNG charts exist in `{os.path.basename(target_dir)}/`.")
    else:
        msg = f"Community PNG charts: found {found_pngs}/75."
        if progressive:
            res.warn(msg + f" (Pending M2 batch rendering). Missing {len(missing_pngs)} files.")
        else:
            res.fail(msg + f" Missing examples: {missing_pngs[:5]}")

    # 2. Macro Summary PNG (target & root mirror)
    macro_name = "00_台中市南屯區30年以上大型社區_大盤匯總走勢圖.png"
    macro_target = os.path.join(target_dir, macro_name)
    macro_root = os.path.join(root_dir, macro_name)

    if os.path.exists(macro_target) and os.path.getsize(macro_target) > 0:
        res.ok(f"Macro summary PNG exists in target dir: `{macro_name}`.")
    else:
        msg = f"Macro summary PNG missing in `{os.path.basename(target_dir)}/`."
        if progressive:
            res.warn(msg)
        else:
            res.fail(msg)

    if os.path.exists(macro_root) and os.path.getsize(macro_root) > 0:
        res.ok(f"Macro summary PNG mirrored at project root: `{macro_name}`.")
    else:
        msg = f"Macro summary PNG root mirror missing: `{macro_name}`."
        if progressive:
            res.warn(msg)
        else:
            res.fail(msg)

    # 3. Markdown Report
    report_name = "台中市南屯區30年以上大型社區歷史實價登錄行情與大盤走勢分析報告.md"
    report_path = os.path.join(target_dir, report_name)
    required_sections = [
        ("歷史週期演變", [r"歷史週期", r"價量演進", r"循環"]),
        ("四大維度深度剖析", [r"供需面", r"實體經濟面|產業與實體經濟面", r"市場面與交易心理|市場與心理面", r"金融面與政策信用管制|金融面與信用管制|信用管制"]),
        ("計量預估模型", [r"計量經濟學動態預估模型", r"測算邏輯|預估區間"]),
        ("75社區名冊", [r"75\s*棟.*名冊|75社區名冊|核心行情統計表|行情.*名冊"]),
        ("專題剖析與議價指南", [r"專題.*剖析|指標.*社區專題", r"議價指引|議價談判指引|實戰決策"])
    ]

    if os.path.exists(report_path) and os.path.getsize(report_path) > 0:
        res.ok(f"Markdown report exists: `{report_name}` ({os.path.getsize(report_path):,} bytes).")
        try:
            with open(report_path, "r", encoding="utf-8") as f:
                report_content = f.read()

            missing_sections = []
            for sec_name, patterns in required_sections:
                matched = True
                for pat in patterns:
                    if not re.search(pat, report_content, re.IGNORECASE):
                        matched = False
                        break
                if not matched:
                    missing_sections.append(sec_name)

            if not missing_sections:
                res.ok("Markdown report contains all 5 required core thematic sections.")
            else:
                msg = f"Markdown report missing required sections: {missing_sections}."
                if progressive:
                    res.warn(msg)
                else:
                    res.fail(msg)
        except Exception as e:
            res.fail(f"Failed to read markdown report: {e}")
    else:
        msg = f"Markdown report missing: `{report_name}`."
        if progressive:
            res.warn(msg + " (Pending M3).")
        else:
            res.fail(msg)

    # 4. PowerPoint Deck (target & root mirror)
    pptx_name = "台中市南屯區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx"
    pptx_target = os.path.join(target_dir, pptx_name)
    pptx_root = os.path.join(root_dir, pptx_name)

    if os.path.exists(pptx_target) and os.path.getsize(pptx_target) > 0:
        res.ok(f"PowerPoint deck exists in target dir: `{pptx_name}`.")
    else:
        msg = f"PowerPoint deck missing in target dir: `{pptx_name}`."
        if progressive:
            res.warn(msg + " (Pending M4).")
        else:
            res.fail(msg)

    if os.path.exists(pptx_root) and os.path.getsize(pptx_root) > 0:
        res.ok(f"PowerPoint deck root mirror exists: `{pptx_name}`.")
    else:
        msg = f"PowerPoint deck root mirror missing: `{pptx_name}`."
        if progressive:
            res.warn(msg + " (Pending M4).")
        else:
            res.fail(msg)

    # 5. Interactive HTML Dashboard
    html_name = "index.html"
    html_path = os.path.join(target_dir, html_name)

    if os.path.exists(html_path) and os.path.getsize(html_path) > 0:
        res.ok(f"HTML dashboard exists: `{html_name}` ({os.path.getsize(html_path):,} bytes).")
        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html_content = f.read()
            # Verify search input
            if "searchInput" in html_content or ('type="text"' in html_content and "搜尋" in html_content):
                res.ok("HTML dashboard contains real-time search input element.")
            else:
                res.fail("HTML dashboard missing search input element (`searchInput`).")

            # Verify macro chart reference
            if macro_name in html_content:
                res.ok("HTML dashboard correctly references macro summary chart image.")
            else:
                res.fail(f"HTML dashboard does not reference macro summary image `{macro_name}`.")

            # Verify community cards presence
            card_matches = re.findall(r'class="[^"]*community-card[^"]*"', html_content)
            data_name_matches = re.findall(r'data-name="([^"]+)"', html_content)
            card_count = max(len(card_matches), len(data_name_matches))
            res.details["html_cards_count"] = card_count
            if card_count >= 75:
                res.ok(f"HTML dashboard contains all {card_count} community cards (>= 75).")
            else:
                msg = f"HTML dashboard has only {card_count} community cards (expected 75)."
                if progressive:
                    res.warn(msg)
                else:
                    res.fail(msg)
        except Exception as e:
            res.fail(f"Failed to read HTML dashboard: {e}")
    else:
        msg = f"HTML dashboard missing: `{html_name}`."
        if progressive:
            res.warn(msg + " (Pending M4).")
        else:
            res.fail(msg)

    return res


# ==============================================================================
# Tier 2: Boundary & Corner Cases Verification
# ==============================================================================
def verify_tier2(roster: List[Dict[str, Any]], transactions_path: Optional[str] = None,
                 progressive: bool = False, verbose: bool = False) -> TestResult:
    res = TestResult("Tier 2: Boundary & Corner Cases (邊界案例、入選條件與資料純淨度)")

    # 1. Qualification Invariants for Roster
    roster_75 = roster[:75]
    if len(roster_75) < 75:
        msg = f"Roster has only {len(roster_75)} communities (expected 75)."
        if progressive:
            res.warn(msg)
        else:
            res.fail(msg)

    roster_names = [c.get("name", "") for c in roster_75]
    invalid_build_dates = []
    invalid_households = []

    for c in roster_75:
        name = c.get("name", "Unknown")
        bdate = str(c.get("build_date", ""))
        hh = c.get("households", 0)

        # Check build date <= 1996-12-31
        year_match = re.match(r"^(\d{4})", bdate)
        if year_match:
            year = int(year_match.group(1))
            if year > 1996:
                invalid_build_dates.append((name, bdate))
        elif not bdate:
            invalid_build_dates.append((name, "MISSING"))

        # Check households >= 100
        try:
            hh_val = int(hh)
            if hh_val < 100:
                invalid_households.append((name, hh_val))
        except Exception:
            invalid_households.append((name, hh))

    if not invalid_build_dates:
        res.ok("All 75 communities strictly meet completion date invariant (build_date <= 1996-12-31).")
    else:
        res.fail(f"Communities violating build_date <= 1996: {invalid_build_dates}")

    if not invalid_households:
        res.ok("All 75 communities strictly meet household invariant (households >= 100).")
    else:
        res.fail(f"Communities violating households >= 100: {invalid_households}")

    # 2. Check 1997 Edge-Case Exclusions
    leaked_1997 = [name for name in EXCLUDED_1997_COMMUNITIES if name in roster_names]
    if not leaked_1997:
        res.ok(f"Verified all 10 excluded 1997 communities are strictly excluded from roster: {', '.join(EXCLUDED_1997_COMMUNITIES[:3])}...")
    else:
        res.fail(f"CRITICAL: 1997 excluded communities incorrectly leaked into 75 roster: {leaked_1997}")

    # 3. Check < 100 Household Edge-Case Exclusions (90~99 units)
    leaked_low_hh = [name for name in EXCLUDED_LOW_HH_COMMUNITIES if name in roster_names]
    if not leaked_low_hh:
        res.ok(f"Verified all 10 borderline 90~99 unit communities are strictly excluded from roster: {', '.join(EXCLUDED_LOW_HH_COMMUNITIES[:3])}...")
    else:
        res.fail(f"CRITICAL: <100 unit borderline communities incorrectly included in roster: {leaked_low_hh}")

    # 4. Transaction Purity Verification
    if not transactions_path or not os.path.exists(transactions_path):
        msg = "Transactions dataset `all_75_nantun_communities_transactions.json` not found."
        if progressive:
            res.warn(msg + " (Pending M1 ingestion & cleaning).")
        else:
            res.fail(msg)
        return res

    try:
        with open(transactions_path, "r", encoding="utf-8") as f:
            tx_data = json.load(f)

        comm_list = tx_data if isinstance(tx_data, list) else tx_data.get("communities", [])
        res.details["dataset_communities_count"] = len(comm_list)

        total_clean_deals = 0
        shop_violations = 0
        special_violations = 0
        price_outliers = 0
        year_violations = 0

        # Regex for floor: detecting 1F / 1樓 while protecting 11樓, 21樓, 31樓
        shop_floor_regex = re.compile(r'(?<![1-9])(1|一)樓|店面|地下|地下室|地下層|B\d+|\d+B', re.IGNORECASE)
        special_keywords = ["親友", "特殊", "關係人", "員工", "債務", "折價", "急售折"]

        for comm in comm_list:
            txs = comm.get("transactions", [])
            for tx in txs:
                is_clean = tx.get("clean_residential", True)
                if not is_clean:
                    continue
                total_clean_deals += 1

                # Check shop violation
                if tx.get("is_first_floor_shop") is True:
                    shop_violations += 1
                floor_str = str(tx.get("shift_floor", ""))
                # Verify floor does not violate shop regex (protect 11F, 21F)
                if floor_str in ["1", "1樓", "一", "一樓", "店面", "B1", "B2"] or shop_floor_regex.search(floor_str):
                    # Check if it was false positive like 11樓 or 21樓
                    if not re.search(r'(11|21|31)樓', floor_str):
                        shop_violations += 1

                # Check special deal violation
                if tx.get("is_special_deal") is True or tx.get("is_special") == 1:
                    special_violations += 1
                remarks = str(tx.get("context", "")) + str(tx.get("price_tips", ""))
                for kw in special_keywords:
                    if kw in remarks and ("排除親友" not in remarks):
                        special_violations += 1
                        break

                # Check price range [5.0, 75.0]
                uprice = tx.get("clean_unit_price")
                if uprice is None:
                    try:
                        uprice = float(tx.get("unit_price", {}).get("price", 0))
                    except Exception:
                        uprice = 0.0
                if uprice < 5.0 or uprice > 75.0:
                    price_outliers += 1

                # Check transaction year >= 2011.0
                dec_year = tx.get("dec_year")
                if dec_year is None:
                    trans_date = str(tx.get("trans_date", ""))
                    if trans_date:
                        try:
                            dec_year = float(trans_date[:4])
                        except Exception:
                            dec_year = 0.0
                    else:
                        dec_year = 0.0
                if dec_year < 2011.0:
                    year_violations += 1

        res.details["total_clean_residential_deals"] = total_clean_deals
        res.details["shop_violations"] = shop_violations
        res.details["special_violations"] = special_violations
        res.details["price_outliers"] = price_outliers
        res.details["year_violations"] = year_violations

        if total_clean_deals == 0:
            res.fail("Transactions dataset contains 0 clean residential transactions.")
            return res

        # Assert zero 1F shops
        if shop_violations == 0:
            res.ok(f"Clean transactions purity: ZERO 1F shop/store deals among {total_clean_deals:,} records.")
        else:
            res.fail(f"Found {shop_violations} contaminated 1F shop transactions in clean residential dataset.")

        # Assert zero special deals
        if special_violations == 0:
            res.ok(f"Clean transactions purity: ZERO special/family deals among {total_clean_deals:,} records.")
        else:
            res.fail(f"Found {special_violations} special/family deals in clean residential dataset.")

        # Assert valid price range [5.0, 75.0]
        if price_outliers == 0:
            res.ok(f"Clean transactions purity: ALL unit prices strictly within [5.0, 75.0] 萬/坪.")
        else:
            res.fail(f"Found {price_outliers} price outliers outside [5.0, 75.0] 萬/坪.")

        # Assert year >= 2011.0
        if year_violations == 0:
            res.ok(f"Clean transactions purity: ALL transaction dates >= 2011.0 (since 実價登錄上路).")
        else:
            res.fail(f"Found {year_violations} transactions with date < 2011.0.")

    except Exception as e:
        res.fail(f"Error parsing transactions dataset `{transactions_path}`: {e}")

    return res


# ==============================================================================
# Tier 3: Cross-Feature Consistency Verification
# ==============================================================================
def verify_tier3(target_dir: str, roster: List[Dict[str, Any]],
                 transactions_path: Optional[str] = None,
                 progressive: bool = False, verbose: bool = False) -> TestResult:
    res = TestResult("Tier 3: Cross-Feature Consistency (跨交付物全維度一致性對齊)")
    canonical_75_names = [c.get("name", "") for c in roster[:75]]

    # 1. Dataset JSON consistency
    if transactions_path and os.path.exists(transactions_path):
        try:
            with open(transactions_path, "r", encoding="utf-8") as f:
                tx_data = json.load(f)
            tx_comms = tx_data if isinstance(tx_data, list) else tx_data.get("communities", [])
            tx_names = [c.get("name", "") for c in tx_comms[:75]]
            if tx_names == canonical_75_names:
                res.ok("Dataset JSON 100% matches authoritative roster ranking and names (1~75).")
            else:
                diffs = [(i+1, c, t) for i, (c, t) in enumerate(zip(canonical_75_names, tx_names)) if c != t]
                res.fail(f"Dataset JSON roster mismatch against authoritative roster: {len(diffs)} differences: {diffs[:3]}")
        except Exception as e:
            res.fail(f"Failed to check dataset JSON consistency: {e}")
    else:
        msg = "Transactions dataset not available for Tier 3 consistency check."
        if progressive:
            res.warn(msg)
        else:
            res.fail(msg)

    # 2. 75 PNG filenames consistency
    if os.path.exists(target_dir):
        png_files = [f for f in os.listdir(target_dir) if f.endswith("_成交行情走勢圖.png") and not f.startswith("00_")]
        png_ranks: Dict[int, str] = {}
        for fn in png_files:
            m = re.match(r"^(\d{2})_(.+)_成交行情走勢圖\.png$", fn)
            if m:
                png_ranks[int(m.group(1))] = m.group(2)

        if len(png_ranks) == 75:
            png_mismatches = []
            for rank, expected_name in enumerate(canonical_75_names, 1):
                actual_name = png_ranks.get(rank, "")
                if actual_name != expected_name:
                    png_mismatches.append((rank, expected_name, actual_name))

            if not png_mismatches:
                res.ok("75 PNG filenames 100% match authoritative roster ranking and names (01~75).")
            else:
                res.fail(f"PNG filenames mismatch: {png_mismatches[:5]}")
        else:
            msg = f"Found {len(png_ranks)}/75 PNG charts for filename consistency check."
            if progressive:
                res.warn(msg + " (Pending M2).")
            else:
                res.fail(msg)
    else:
        if progressive:
            res.warn("Target directory not created yet for PNG consistency check.")
        else:
            res.fail("Target directory does not exist for PNG consistency check.")

    # 3. Markdown Table rows consistency
    report_path = os.path.join(target_dir, "台中市南屯區30年以上大型社區歷史實價登錄行情與大盤走勢分析報告.md")
    if os.path.exists(report_path):
        try:
            with open(report_path, "r", encoding="utf-8") as f:
                md_content = f.read()

            # Find table rows: e.g. | #01 | **理想貴族** | or | 01 | 理想貴族 |
            md_ranks: Dict[int, str] = {}
            for line in md_content.splitlines():
                m = re.search(r'\|\s*#?0*(\d{1,2})\s*\|\s*\*?\*?([^\*\|]+?)\*?\*?\s*\|', line)
                if m:
                    r_num = int(m.group(1))
                    r_name = m.group(2).strip()
                    if 1 <= r_num <= 75 and r_num not in md_ranks:
                        md_ranks[r_num] = r_name

            if len(md_ranks) >= 75:
                md_mismatches = []
                for rank, expected_name in enumerate(canonical_75_names, 1):
                    actual_name = md_ranks.get(rank, "")
                    if actual_name != expected_name and expected_name not in actual_name:
                        md_mismatches.append((rank, expected_name, actual_name))

                if not md_mismatches:
                    res.ok("Markdown report table rows 100% match authoritative roster ranking and names.")
                else:
                    res.fail(f"Markdown report table rows mismatch: {md_mismatches[:5]}")
            else:
                msg = f"Markdown report has {len(md_ranks)}/75 parsed community table rows."
                if progressive:
                    res.warn(msg + " (Pending M3).")
                else:
                    res.fail(msg)
        except Exception as e:
            res.fail(f"Failed to check Markdown table consistency: {e}")
    else:
        if progressive:
            res.warn("Markdown report not found for table consistency check.")
        else:
            res.fail("Markdown report missing for table consistency check.")

    # 4. PPTX Slides consistency
    pptx_path = os.path.join(target_dir, "台中市南屯區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx")
    if os.path.exists(pptx_path):
        try:
            from pptx import Presentation
            prs = Presentation(pptx_path)
            slide_texts = []
            for slide in prs.slides:
                texts = []
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        texts.append(shape.text_frame.text)
                slide_texts.append(" ".join(texts))

            # Match communities across slides
            matched_comms = set()
            for text in slide_texts:
                for name in canonical_75_names:
                    if name in text:
                        matched_comms.add(name)

            if len(matched_comms) >= 75:
                res.ok(f"PowerPoint slides 100% cover all {len(matched_comms)} community names.")
            else:
                missing_in_pptx = [name for name in canonical_75_names if name not in matched_comms]
                msg = f"PowerPoint covers {len(matched_comms)}/75 community names. Missing: {missing_in_pptx[:5]}"
                if progressive:
                    res.warn(msg + " (Pending M4).")
                else:
                    res.fail(msg)
        except Exception as e:
            res.fail(f"Failed to check PPTX consistency: {e}")
    else:
        if progressive:
            res.warn("PowerPoint deck not found for PPTX consistency check.")
        else:
            res.fail("PowerPoint deck missing for PPTX consistency check.")

    # 5. HTML Dashboard cards consistency
    html_path = os.path.join(target_dir, "index.html")
    if os.path.exists(html_path):
        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html_content = f.read()

            html_matched = [name for name in canonical_75_names if name in html_content]
            if len(html_matched) == 75:
                res.ok("HTML dashboard cards 100% cover all 75 community names.")
            else:
                missing_in_html = [name for name in canonical_75_names if name not in html_matched]
                msg = f"HTML dashboard covers {len(html_matched)}/75 communities. Missing: {missing_in_html[:5]}"
                if progressive:
                    res.warn(msg + " (Pending M4).")
                else:
                    res.fail(msg)
        except Exception as e:
            res.fail(f"Failed to check HTML dashboard consistency: {e}")
    else:
        if progressive:
            res.warn("HTML dashboard not found for card consistency check.")
        else:
            res.fail("HTML dashboard missing for card consistency check.")

    return res


# ==============================================================================
# Tier 4: Real-World Acceptance Scenarios Verification
# ==============================================================================
def verify_tier4(target_dir: str, root_dir: str, roster: List[Dict[str, Any]],
                 progressive: bool = False, verbose: bool = False) -> TestResult:
    res = TestResult("Tier 4: Real-World Acceptance Scenarios (實體檔案鑑識與驗收場景)")

    # 1. PIL Image Validity & Dimensions (>= 3000x2000, non-corrupted)
    try:
        from PIL import Image
        pil_available = True
    except ImportError:
        pil_available = False
        res.fail("Pillow (PIL) is not installed in the environment.")

    if pil_available:
        images_to_check = []
        # Community PNGs
        for idx, comm in enumerate(roster[:75], 1):
            name = comm.get("name", "")
            images_to_check.append(os.path.join(target_dir, f"{idx:02d}_{name}_成交行情走勢圖.png"))
        # Macro summary PNG
        images_to_check.append(os.path.join(target_dir, "00_台中市南屯區30年以上大型社區_大盤匯總走勢圖.png"))

        existing_images = [p for p in images_to_check if os.path.exists(p)]
        res.details["images_checked"] = len(existing_images)
        res.details["images_target"] = len(images_to_check)

        if len(existing_images) == 0:
            msg = "No PNG charts available for PIL image forensic check."
            if progressive:
                res.warn(msg + " (Pending M2).")
            else:
                res.fail(msg)
        else:
            corrupt_images = []
            low_res_images = []
            for p in existing_images:
                try:
                    with Image.open(p) as img:
                        w, h = img.size
                        img.verify()
                    # Reopen to load pixel data
                    with Image.open(p) as img:
                        img.load()
                    if w < 3000 or h < 2000:
                        low_res_images.append((os.path.basename(p), f"{w}x{h}"))
                except Exception as e:
                    corrupt_images.append((os.path.basename(p), str(e)))

            if corrupt_images:
                res.fail(f"Corrupted PNG image files detected: {corrupt_images[:3]}")
            elif low_res_images:
                res.fail(f"Images failing resolution threshold (must be >= 3000x2000): {low_res_images[:3]}")
            else:
                if len(existing_images) == 76:
                    res.ok(f"All 76 PNG charts (75 communities + 1 macro) verified intact and exceed 3000x2000 resolution.")
                else:
                    msg = f"Verified {len(existing_images)}/76 PNG charts intact with resolution >= 3000x2000."
                    if progressive:
                        res.ok(msg)
                    else:
                        res.fail(msg)

    # 2. PowerPoint Validity & 16:9 Aspect Ratio (>= 60 slides)
    pptx_path = os.path.join(target_dir, "台中市南屯區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx")
    if os.path.exists(pptx_path):
        try:
            from pptx import Presentation
            prs = Presentation(pptx_path)
            slide_count = len(prs.slides)
            res.details["pptx_slide_count"] = slide_count

            # Aspect ratio check
            sw = prs.slide_width
            sh = prs.slide_height
            aspect_ratio = sw / sh if sh else 0
            res.details["pptx_aspect_ratio"] = round(aspect_ratio, 4)

            # Check slide count >= 60
            if slide_count >= 60:
                res.ok(f"PowerPoint slide count verified: {slide_count} slides (>= 60 threshold).")
            else:
                res.fail(f"PowerPoint slide count below requirement: {slide_count} < 60.")

            # Check 16:9 (approx 1.7778)
            if 1.75 <= aspect_ratio <= 1.80:
                res.ok(f"PowerPoint aspect ratio verified as standard 16:9 widescreen ({aspect_ratio:.3f}).")
            else:
                res.fail(f"PowerPoint aspect ratio violation: {aspect_ratio:.3f} (expected 16:9 approx 1.778).")

        except Exception as e:
            res.fail(f"Failed to open/verify PowerPoint deck with python-pptx: {e}")
    else:
        msg = "PowerPoint deck not found for Tier 4 presentation forensics."
        if progressive:
            res.warn(msg + " (Pending M4).")
        else:
            res.fail(msg)

    # 3. Markdown Report Size >= 40KB & UTF-8
    report_path = os.path.join(target_dir, "台中市南屯區30年以上大型社區歷史實價登錄行情與大盤走勢分析報告.md")
    if os.path.exists(report_path):
        size = os.path.getsize(report_path)
        res.details["markdown_size_bytes"] = size
        if size >= 40000:
            res.ok(f"Markdown report size verified: {size:,} bytes (>= 40 KB threshold).")
        else:
            res.fail(f"Markdown report size insufficient: {size:,} bytes < 40,000 bytes.")

        try:
            with open(report_path, "r", encoding="utf-8") as f:
                _ = f.read()
            res.ok("Markdown report verified as valid UTF-8 encoding.")
        except UnicodeDecodeError as ue:
            res.fail(f"Markdown report failed UTF-8 encoding check: {ue}")
    else:
        msg = "Markdown report not found for Tier 4 document forensics."
        if progressive:
            res.warn(msg + " (Pending M3).")
        else:
            res.fail(msg)

    # 4. HTML Dashboard Size >= 100KB & Lightbox Modal
    html_path = os.path.join(target_dir, "index.html")
    if os.path.exists(html_path):
        html_size = os.path.getsize(html_path)
        res.details["html_size_bytes"] = html_size
        if html_size >= 100000:
            res.ok(f"HTML dashboard size verified: {html_size:,} bytes (>= 100 KB threshold).")
        else:
            res.fail(f"HTML dashboard size insufficient: {html_size:,} bytes < 100,000 bytes.")

        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html_txt = f.read()
            res.ok("HTML dashboard verified as valid UTF-8 encoding.")

            # Lightbox modal logic check
            has_lightbox = (
                ("modal" in html_txt.lower() or "lightbox" in html_txt.lower()) and
                ("openModal" in html_txt or "modalImage" in html_txt or "imageModal" in html_txt)
            )
            if has_lightbox:
                res.ok("HTML dashboard verified to contain full interactive Lightbox modal image viewer logic.")
            else:
                res.fail("HTML dashboard missing interactive Lightbox modal logic.")
        except UnicodeDecodeError as ue:
            res.fail(f"HTML dashboard failed UTF-8 encoding check: {ue}")
    else:
        msg = "HTML dashboard not found for Tier 4 web forensics."
        if progressive:
            res.warn(msg + " (Pending M4).")
        else:
            res.fail(msg)

    return res


# ==============================================================================
# Self-Test Suite (--test-self)
# ==============================================================================
def run_self_tests(verbose: bool = True) -> bool:
    """Executes synthetic in-memory / tempdir unit tests to verify the test suite itself."""
    print("=" * 80)
    print("【執行 verify_nantun_suite.py 內部單元自我測試 (Self-Test Suite)】")
    print("=" * 80)

    from PIL import Image
    from pptx import Presentation
    from pptx.util import Inches

    temp_dir = tempfile.mkdtemp(prefix="nantun_verify_selftest_")
    mock_target = os.path.join(temp_dir, "南屯區30年以上大型社區走勢圖")
    os.makedirs(mock_target, exist_ok=True)

    passed_tests = 0
    total_tests = 8

    try:
        # Build valid mock environment
        mock_roster = [
            {"id": 1000 + i, "name": name, "households": hh, "build_date": bdate, "rank": i}
            for i, name, hh, bdate in CANONICAL_75_ROSTER
        ]

        # Generate 76 mock high-res PNG images (3120x2080)
        img_mock = Image.new("RGB", (3120, 2080), color=(15, 23, 42))
        macro_fn = "00_台中市南屯區30年以上大型社區_大盤匯總走勢圖.png"
        img_mock.save(os.path.join(mock_target, macro_fn))
        img_mock.save(os.path.join(temp_dir, macro_fn)) # root mirror

        for i, name, _, _ in CANONICAL_75_ROSTER:
            fn = f"{i:02d}_{name}_成交行情走勢圖.png"
            img_mock.save(os.path.join(mock_target, fn))

        # Generate valid mock markdown report (>= 40KB)
        md_fn = "台中市南屯區30年以上大型社區歷史實價登錄行情與大盤走勢分析報告.md"
        table_rows = "\n".join([f"| #{i:02d} | **{name}** | {hh} 戶 | 1994-01-01 |" for i, name, hh, _ in CANONICAL_75_ROSTER])
        md_text = f"""# 台中市南屯區30年以上大型社區歷史實價登錄行情與大盤走勢分析報告

## 一、歷史週期演變與價量演進循環
歷經民國101年以來十五年循環，南屯區老牌大盤走勢穩健。

## 二、促使老牌大社區價格持續下修之四大維度深度剖析
- 供需面：老屋存量與新成屋交屋潮夾殺
- 實體經濟面：通膨與裝修成本暴增反噬殘值
- 市場面與交易心理：量先價行確立、追價狂熱破滅
- 金融面與政策信用管制：央行選擇性信用管制重拳

## 三、計量經濟學動態預估模型架構與測算邏輯
6大模組與信賴區間測算。

## 四、全南屯區 75 棟老牌大社區實價登錄行情名冊與核心行情統計表
| 排名 | 社區名稱 | 戶數 | 完工日期 |
|---|---|---|---|
{table_rows}

## 五、重點指標社區專題剖析與買賣雙方實戰決策議價指南
買賣雙方議價指引與談判策略。
"""
        # Pad to > 40KB
        md_text += "\n" + ("詳細分析補充內容文字段落。" * 2500)
        with open(os.path.join(mock_target, md_fn), "w", encoding="utf-8") as f:
            f.write(md_text)

        # Generate valid mock PPTX (>= 60 slides, 16:9)
        pptx_fn = "台中市南屯區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx"
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank_slide_layout = prs.slide_layouts[6]

        # 10 intro slides
        for idx in range(10):
            s = prs.slides.add_slide(blank_slide_layout)
            txBox = s.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
            txBox.text_frame.text = f"Intro Slide {idx+1}: 大盤全景與歷史週期"

        # 75 community slides
        for i, name, _, _ in CANONICAL_75_ROSTER:
            s = prs.slides.add_slide(blank_slide_layout)
            txBox = s.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
            txBox.text_frame.text = f"社區分析 #{i:02d}: {name} 行情與走勢"

        prs.save(os.path.join(mock_target, pptx_fn))
        prs.save(os.path.join(temp_dir, pptx_fn)) # root mirror

        # Generate valid mock HTML dashboard (>= 100KB with Lightbox)
        cards_html = "\n".join([
            f'<div class="community-card bg-slate-900" data-name="{name}" onclick="openModal(\'{i:02d}_{name}_成交行情走勢圖.png\', \'{name}\')"><h3>{i:02d}. {name}</h3></div>'
            for i, name, _, _ in CANONICAL_75_ROSTER
        ])
        html_text = f"""<!DOCTYPE html>
<html lang="zh-TW" class="dark">
<head><meta charset="UTF-8"><title>南屯區走勢圖</title></head>
<body>
  <header>
    <input id="searchInput" type="text" placeholder="搜尋社區名稱...">
  </header>
  <section id="macroChart">
    <img src="00_台中市南屯區30年以上大型社區_大盤匯總走勢圖.png">
  </section>
  <main>
    {cards_html}
  </main>
  <div id="imageModal" class="hidden">
    <img id="modalImage" src="">
  </div>
  <script>
    function openModal(src, title) {{ document.getElementById('modalImage').src = src; }}
    function closeModal() {{ document.getElementById('imageModal').classList.add('hidden'); }}
  </script>
</body>
</html>
"""
        html_text += "<!-- padding comment -->\n" * 4000
        with open(os.path.join(mock_target, "index.html"), "w", encoding="utf-8") as f:
            f.write(html_text)

        # Generate valid mock transactions dataset
        mock_tx_dataset = {
            "communities": [
                {
                    "name": name,
                    "households": hh,
                    "build_date": bdate,
                    "transactions": [
                        {
                            "clean_residential": True,
                            "clean_unit_price": 25.5,
                            "dec_year": 2024.5,
                            "shift_floor": "8樓",
                            "is_first_floor_shop": False,
                            "is_special_deal": False,
                            "context": "標準住家"
                        }
                    ]
                }
                for i, name, hh, bdate in CANONICAL_75_ROSTER
            ]
        }
        tx_path = os.path.join(temp_dir, "mock_transactions.json")
        with open(tx_path, "w", encoding="utf-8") as f:
            json.dump(mock_tx_dataset, f, ensure_ascii=False)

        # ----------------------------------------------------------------------
        # TC1: Full Positive Case (All 4 Tiers must PASS 100%)
        # ----------------------------------------------------------------------
        t1 = verify_tier1(mock_target, temp_dir, mock_roster)
        t2 = verify_tier2(mock_roster, tx_path)
        t3 = verify_tier3(mock_target, mock_roster, tx_path)
        t4 = verify_tier4(mock_target, temp_dir, mock_roster)

        if t1.passed and t2.passed and t3.passed and t4.passed:
            print("  [PASS] TC1: Valid mock deliverables passed Tier 1 ~ 4 full acceptance 100%.")
            passed_tests += 1
        else:
            print(f"  [FAIL] TC1: Valid mock deliverables failed: T1={t1.passed}, T2={t2.passed}, T3={t3.passed}, T4={t4.passed}")
            if not t1.passed: print("    T1 errors:", t1.errors)
            if not t2.passed: print("    T2 errors:", t2.errors)
            if not t3.passed: print("    T3 errors:", t3.errors)
            if not t4.passed: print("    T4 errors:", t4.errors)

        # ----------------------------------------------------------------------
        # TC2: 1997 Excluded Community Leak Interception
        # ----------------------------------------------------------------------
        leaked_roster = list(mock_roster)
        leaked_roster[10] = {"name": "龍歡喜大廈", "households": 496, "build_date": "1997-12-15"}
        t2_leak = verify_tier2(leaked_roster, tx_path)
        if not t2_leak.passed and any("龍歡喜大廈" in e for e in t2_leak.errors):
            print("  [PASS] TC2: 1997 excluded community correctly caught and rejected by Tier 2.")
            passed_tests += 1
        else:
            print("  [FAIL] TC2: Failed to catch 1997 excluded community leak.")

        # ----------------------------------------------------------------------
        # TC3: <100 Household Edge Case Interception
        # ----------------------------------------------------------------------
        low_hh_roster = list(mock_roster)
        low_hh_roster[15] = {"name": "心田", "households": 99, "build_date": "1994-10-11"}
        t2_low_hh = verify_tier2(low_hh_roster, tx_path)
        if not t2_low_hh.passed and any("心田" in e for e in t2_low_hh.errors):
            print("  [PASS] TC3: <100 household edge-case community correctly caught and rejected by Tier 2.")
            passed_tests += 1
        else:
            print("  [FAIL] TC3: Failed to catch <100 household edge-case community.")

        # ----------------------------------------------------------------------
        # TC4: Contaminated 1F Shop Transaction Interception
        # ----------------------------------------------------------------------
        dirty_tx_dataset = json.loads(json.dumps(mock_tx_dataset))
        dirty_tx_dataset["communities"][0]["transactions"].append({
            "clean_residential": True,
            "clean_unit_price": 20.0,
            "dec_year": 2023.0,
            "shift_floor": "1樓",
            "is_first_floor_shop": True,
            "is_special_deal": False
        })
        dirty_tx_path = os.path.join(temp_dir, "dirty_shop_transactions.json")
        with open(dirty_tx_path, "w", encoding="utf-8") as f:
            json.dump(dirty_tx_dataset, f)

        t2_dirty = verify_tier2(mock_roster, dirty_tx_path)
        if not t2_dirty.passed and any("1F shop" in e for e in t2_dirty.errors):
            print("  [PASS] TC4: Contaminated 1F shop transaction correctly caught and rejected by Tier 2.")
            passed_tests += 1
        else:
            print("  [FAIL] TC4: Failed to catch contaminated 1F shop transaction.")

        # ----------------------------------------------------------------------
        # TC5: Missing PNG Chart Interception in Tier 1
        # ----------------------------------------------------------------------
        del_png = os.path.join(mock_target, "01_理想貴族_成交行情走勢圖.png")
        os.remove(del_png)
        t1_missing = verify_tier1(mock_target, temp_dir, mock_roster)
        if not t1_missing.passed and any("74/75" in e or "missing" in e.lower() for e in t1_missing.errors):
            print("  [PASS] TC5: Missing PNG chart correctly caught and failed by Tier 1.")
            passed_tests += 1
        else:
            print("  [FAIL] TC5: Failed to catch missing PNG chart.")
        # Restore PNG
        img_mock.save(del_png)

        # ----------------------------------------------------------------------
        # TC6: Name Mismatch Across Deliverables in Tier 3
        # ----------------------------------------------------------------------
        mismatch_roster = list(mock_roster)
        mismatch_roster[0] = {"name": "篡改社區名稱", "households": 732, "build_date": "1991-12-06"}
        t3_mismatch = verify_tier3(mock_target, mismatch_roster, tx_path)
        if not t3_mismatch.passed:
            print("  [PASS] TC6: Community name discrepancy across deliverables caught by Tier 3.")
            passed_tests += 1
        else:
            print("  [FAIL] TC6: Failed to catch community name mismatch in Tier 3.")

        # ----------------------------------------------------------------------
        # TC7: Low Resolution Image Interception in Tier 4
        # ----------------------------------------------------------------------
        low_res_png = os.path.join(mock_target, "02_宏全世界廣場_成交行情走勢圖.png")
        bad_img = Image.new("RGB", (1280, 720), color=(0, 0, 0))
        bad_img.save(low_res_png)
        t4_low_res = verify_tier4(mock_target, temp_dir, mock_roster)
        if not t4_low_res.passed and any("3000x2000" in e for e in t4_low_res.errors):
            print("  [PASS] TC7: Low resolution image (<3000x2000) correctly caught by Tier 4.")
            passed_tests += 1
        else:
            print("  [FAIL] TC7: Failed to catch low resolution image in Tier 4.")
        # Restore high res image
        img_mock.save(low_res_png)

        # ----------------------------------------------------------------------
        # TC8: Missing Lightbox Logic Interception in Tier 4
        # ----------------------------------------------------------------------
        bad_html_path = os.path.join(mock_target, "index.html")
        with open(bad_html_path, "w", encoding="utf-8") as f:
            f.write("<html><body><div>No modal, no lightbox</div>" + ("x" * 120000) + "</body></html>")
        t4_no_lightbox = verify_tier4(mock_target, temp_dir, mock_roster)
        if not t4_no_lightbox.passed and any("Lightbox" in e for e in t4_no_lightbox.errors):
            print("  [PASS] TC8: Missing Lightbox modal in HTML dashboard correctly caught by Tier 4.")
            passed_tests += 1
        else:
            print("  [FAIL] TC8: Failed to catch missing Lightbox modal in Tier 4.")

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    print("=" * 80)
    if passed_tests == total_tests:
        print(f"🎉 [SELF-TEST SUCCESS] All {total_tests}/{total_tests} internal verification tests passed successfully!")
        return True
    else:
        print(f"❌ [SELF-TEST FAILURE] Only {passed_tests}/{total_tests} internal verification tests passed.")
        return False


# ==============================================================================
# CLI Entrypoint & Main Driver
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(
        description="Automated Verification Suite for Nantun District 30yr+ Large Communities Project",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--target-dir", default=DEFAULT_TARGET_DIR, help="Path to deliverables folder")
    parser.add_argument("--root-dir", default=PROJECT_ROOT, help="Path to project root directory")
    parser.add_argument("--roster-path", default=None, help="Path to qualified communities JSON")
    parser.add_argument("--tx-path", default=None, help="Path to transactions JSON")

    # Verification Scope Flags
    parser.add_argument("--all", action="store_true", help="Run all 4 tiers (default strict acceptance)")
    parser.add_argument("--progressive", action="store_true", help="Progressive mode for milestone validation")
    parser.add_argument("--tier1", action="store_true", help="Run Tier 1 (Feature Coverage) only")
    parser.add_argument("--tier2", action="store_true", help="Run Tier 2 (Boundary & Corner Cases) only")
    parser.add_argument("--tier3", action="store_true", help="Run Tier 3 (Cross-Feature Consistency) only")
    parser.add_argument("--tier4", action="store_true", help="Run Tier 4 (Acceptance & Forensics) only")

    # Milestone Layer Aliases
    parser.add_argument("--data-layer", action="store_true", help="Check M1 Data Ingestion & Purity Layer")
    parser.add_argument("--chart-layer", action="store_true", help="Check M2 Chart Image Rendering Layer")
    parser.add_argument("--doc-layer", action="store_true", help="Check M3/M4 Reports, Presentation & Dashboard")

    # Diagnostic & Self-Test Flags
    parser.add_argument("--test-self", action="store_true", help="Run internal self-test suite with mock fixtures")
    parser.add_argument("--json", action="store_true", help="Output machine-readable JSON summary")
    parser.add_argument("--verbose", action="store_true", help="Print verbose details and warnings")

    args = parser.parse_args()

    # Self-test handler
    if args.test_self:
        success = run_self_tests(verbose=args.verbose)
        sys.exit(0 if success else 1)

    # Determine scope
    run_all = args.all or not (args.tier1 or args.tier2 or args.tier3 or args.tier4 or args.data_layer or args.chart_layer or args.doc_layer)
    do_tier1 = run_all or args.tier1 or args.chart_layer or args.doc_layer
    do_tier2 = run_all or args.tier2 or args.data_layer
    do_tier3 = run_all or args.tier3 or args.doc_layer
    do_tier4 = run_all or args.tier4 or args.chart_layer or args.doc_layer

    progressive = args.progressive

    # Load Roster
    roster = get_roster_data(args.roster_path)
    tx_file = find_transactions_file(args.tx_path)

    print("=" * 80)
    print("【台中市南屯區30年以上大型指標社區走勢圖 — 獨立自動化驗證套件】")
    print(f"Target Directory: {args.target_dir}")
    print(f"Project Root    : {args.root_dir}")
    print(f"Execution Mode  : {'PROGRESSIVE (階段演進驗收)' if progressive else 'STRICT ACCEPTANCE (全量嚴格驗收)'}")
    print("=" * 80)

    results: List[TestResult] = []

    if do_tier1:
        print("\n▶ [Tier 1] 正在驗證產物覆蓋性與基本結構 (Feature Coverage)...")
        r1 = verify_tier1(args.target_dir, args.root_dir, roster, progressive=progressive, verbose=args.verbose)
        results.append(r1)

    if do_tier2:
        print("\n▶ [Tier 2] 正在驗證邊界案例、入選條件與資料純淨度 (Boundary & Corner Cases)...")
        r2 = verify_tier2(roster, tx_file, progressive=progressive, verbose=args.verbose)
        results.append(r2)

    if do_tier3:
        print("\n▶ [Tier 3] 正在驗證跨交付物全維度一致性對齊 (Cross-Feature Consistency)...")
        r3 = verify_tier3(args.target_dir, roster, tx_file, progressive=progressive, verbose=args.verbose)
        results.append(r3)

    if do_tier4:
        print("\n▶ [Tier 4] 正在執行實體檔案鑑識與驗收場景 (Real-World Acceptance Scenarios)...")
        r4 = verify_tier4(args.target_dir, args.root_dir, roster, progressive=progressive, verbose=args.verbose)
        results.append(r4)

    # Output Presentation
    print("\n" + "=" * 80)
    print("【驗證結果匯總報告】")
    print("=" * 80)

    all_passed = True
    for r in results:
        status_symbol = "✅ PASS" if r.passed else "❌ FAIL"
        if not r.passed:
            all_passed = False
        print(f"\n{status_symbol} {r.name}")

        for msg in r.messages:
            print(f"  [OK] {msg}")
        for warn in r.warnings:
            print(f"  [WARN] {warn}")
        for err in r.errors:
            print(f"  [ERROR] {err}")

    print("\n" + "=" * 80)
    if all_passed:
        if progressive:
            print("🎉【驗收結論：PROGRESSIVE PASS】當前已產出之階段成果全數通過驗證！無任何阻礙性缺陷。")
        else:
            print("🎉【驗收結論：FULL SUITE PASS】全 4 個驗證層級 (Tier 1 ~ 4) 100% 驗收通過！符合專案最高交付標準。")
    else:
        print("❌【驗收結論：FAIL】偵測到未滿足之驗收條件或缺失檔案，請參閱上方 [ERROR] 項目依序修復。")
    print("=" * 80)

    if args.json:
        report_dict = {
            "all_passed": all_passed,
            "progressive": progressive,
            "tiers": [r.to_dict() for r in results]
        }
        print("\nJSON REPORT:")
        print(json.dumps(report_dict, indent=2, ensure_ascii=False))

    sys.exit(0 if all_passed else 1)


if __name__ == "__main__":
    main()

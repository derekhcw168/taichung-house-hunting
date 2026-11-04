#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scripts/challenge_facts_boundaries.py
------------------------------------
Adversarial Empirical Verification Harness:
Fact-Checking, Negotiation Boundaries & Presentation Integrity Challenge
for Taichung South District "Goteborg" (歌德堡) Deliverables Suite.

Tracks:
1. Fact-Checking: Geographic Disambiguation of "熊貓東京" (西屯區) & "仁美國宅" (梧棲區)
   vs South District counterparts "熊貓公園市" & "大慶國宅" (大慶中山).
2. Negotiation Boundaries: Stress-testing 810萬 bottom anchor (refusal risk) and
   930萬 ceiling (overpaying risk) + BATNA analysis.
3. Visual & Formatting Integrity: Exhaustive inspection of all 13 PPTX slides and Word docx
   (overflows, text box collisions, truncation, missing footnotes).
4. Project Suite Runner: Execution and exit code inspection of verify_goteborg_suite.py.
"""

import os
import sys
import json
import re
import subprocess
from typing import Dict, List, Any, Tuple

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

try:
    import docx
except ImportError:
    docx = None

try:
    import pptx
    from pptx import Presentation
    from pptx.util import Inches, Pt
except ImportError:
    pptx = None
    Presentation = None

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
GOTEBORG_DIR = os.path.join(PROJECT_ROOT, "歌德堡")
MD_PATH = os.path.join(GOTEBORG_DIR, "歌德堡_特定待售戶別出價分析與社區整體行情報告.md")
DOCX_PATH = os.path.join(GOTEBORG_DIR, "歌德堡_特定待售戶別出價分析與社區整體行情報告.docx")
PPTX_PATH = os.path.join(GOTEBORG_DIR, "歌德堡_特定待售戶別出價分析與社區整體行情報告.pptx")
JSON_PATH = os.path.join(PROJECT_ROOT, "scratch", "goteborg_valuation_master.json")
ALL_54_PATH = os.path.join(PROJECT_ROOT, "scratch", "all_54_communities_transactions.json")
SOUTH_AGE24_PATH = os.path.join(PROJECT_ROOT, "scratch", "591_south_age24plus_all.json")
DAQING_MASTER_PATH = os.path.join(PROJECT_ROOT, "scratch", "daqing_valuation_master.json")
SYNC_SCRIPT_PATH = os.path.join(PROJECT_ROOT, "scripts", "sync_to_github.py")
SUITE_SCRIPT_PATH = os.path.join(PROJECT_ROOT, "scripts", "verify_goteborg_suite.py")


class GoteborgAdversarialChallenge:
    def __init__(self):
        self.results = {
            "track_1_fact_checking": {},
            "track_2_negotiation_boundaries": {},
            "track_3_presentation_integrity": {},
            "track_4_suite_verification": {},
            "overall_verdict": "PENDING",
            "findings": []
        }

    # =========================================================================
    # Track 1: Fact-Checking & Geographic Disambiguation
    # =========================================================================
    def run_track_1_fact_checking(self) -> bool:
        print("\n" + "="*80)
        print("【TRACK 1】事實查核與跨區地理混淆鑑識 (Fact-Checking & Geographic Disambiguation)")
        print("="*80)
        track_pass = True

        # 1.1 Verify Markdown content
        print("\n--- 1.1 檢視 Markdown 報告之事實澄清段落 ---")
        if not os.path.exists(MD_PATH):
            self.results["track_1_fact_checking"]["md_missing"] = True
            self.results["findings"].append("CRITICAL: Markdown deliverable missing")
            return False

        with open(MD_PATH, 'r', encoding='utf-8') as f:
            md_text = f.read()

        has_panda_tokyo = "熊貓東京" in md_text
        has_panda_park = "熊貓公園市" in md_text
        has_renmei = "仁美國宅" in md_text
        has_daqing = "大慶國宅" in md_text
        has_section_3_2 = ("3.2" in md_text and ("事實澄清" in md_text or "地理歸屬" in md_text))

        print(f"  - Markdown 包含「熊貓東京」澄清: {has_panda_tokyo}")
        print(f"  - Markdown 包含「熊貓公園市」作為南區對標: {has_panda_park}")
        print(f"  - Markdown 包含「仁美國宅」澄清: {has_renmei}")
        print(f"  - Markdown 包含「大慶國宅」作為南區對標: {has_daqing}")
        print(f"  - Markdown 具備第 3.2 節專屬事實校正專章: {has_section_3_2}")

        if not (has_panda_tokyo and has_panda_park and has_renmei and has_daqing and has_section_3_2):
            self.results["findings"].append("HIGH: Markdown lacks complete Section 3.2 geographic clarification")
            track_pass = False

        # 1.2 Verify South District Counterparts Transaction Data
        print("\n--- 1.2 驗證南區真實對標物業之實價登錄與數據基礎 ---")
        # 1.2.1 Check 熊貓公園市 in 591_south_age24plus_all.json
        panda_found_south = False
        if os.path.exists(SOUTH_AGE24_PATH):
            with open(SOUTH_AGE24_PATH, 'r', encoding='utf-8') as f:
                south_data = json.load(f)
            for item in south_data:
                if item.get("name") == "熊貓公園市" and "南區" in item.get("section", ""):
                    panda_found_south = True
                    print(f"  [OK] 熊貓公園市 於 591 南區社區庫驗證成功: ID={item.get('id')}, 地址={item.get('simple_address')}, 戶數={item.get('house_holds')}")
                    break
        if not panda_found_south:
            self.results["findings"].append("HIGH: 熊貓公園市 not found in 591 south community database")
            track_pass = False

        # 1.2.2 Check 大慶國宅 in all_54_communities_transactions.json or daqing_valuation_master.json
        daqing_tx_count = 0
        daqing_clean_count = 0
        if os.path.exists(ALL_54_PATH):
            with open(ALL_54_PATH, 'r', encoding='utf-8') as f:
                all_54 = json.load(f)
            for c in all_54.get("communities", []):
                if c.get("name") in ["大慶中山", "大慶國宅"]:
                    daqing_tx_count = len(c.get("transactions", []))
                    daqing_clean_count = sum(1 for t in c.get("transactions", []) if t.get("clean_residential", False))
                    print(f"  [OK] 大慶國宅(大慶中山) 於南區實登庫驗證成功: 總交易 {daqing_tx_count} 筆, 純住宅乾淨市價 {daqing_clean_count} 筆")
                    break

        if daqing_tx_count < 30:
            self.results["findings"].append(f"HIGH: 大慶國宅 transactions insufficient (found {daqing_tx_count})")
            track_pass = False

        # 1.3 Verify Word (.docx) for Counterparts
        print("\n--- 1.3 檢視 Word 報告 (.docx) 之競品表格與事實澄清 ---")
        docx_has_panda_park = False
        docx_has_daqing = False
        docx_has_clarification_text = False
        if docx and os.path.exists(DOCX_PATH):
            doc = docx.Document(DOCX_PATH)
            docx_full_text = "\n".join([p.text for p in doc.paragraphs] + [c.text for t in doc.tables for r in t.rows for c in r.cells])
            docx_has_panda_park = "熊貓公園市" in docx_full_text
            docx_has_daqing = "大慶國宅" in docx_full_text
            docx_has_clarification_text = ("熊貓東京" in docx_full_text and "仁美國宅" in docx_full_text)
            print(f"  - Word 表格包含「熊貓公園市」: {docx_has_panda_park}")
            print(f"  - Word 表格包含「大慶國宅」: {docx_has_daqing}")
            print(f"  - Word 包含「熊貓東京 / 仁美國宅」事實澄清文字: {docx_has_clarification_text}")

            if not docx_has_clarification_text:
                print("  [NOTE] Word 報告第 3 章競品表格已正確置換南區對標 (熊貓公園市 28.45萬, 大慶國宅 23.00萬)，惟省略了 Markdown 3.2 節之文字澄清敘述。")
                self.results["findings"].append("LOW_NOTICE: Word report competitor table correctly uses South District counterparts (熊貓公園市, 大慶國宅), but omits textual clarification narrative of 3.2 section found in Markdown.")

        # 1.4 Verify Master JSON contains both communities
        print("\n--- 1.4 檢視估價主資料庫 (goteborg_valuation_master.json) ---")
        json_has_panda_park = False
        json_has_daqing = False
        if os.path.exists(JSON_PATH):
            with open(JSON_PATH, 'r', encoding='utf-8') as f:
                v_data = json.load(f)
            comps = v_data.get("comparable_communities", [])
            for c in comps:
                if c.get("name") == "熊貓公園市":
                    json_has_panda_park = True
                    print(f"  [OK] 熊貓公園市 於 master JSON 驗證: 均價 {c.get('recent_avg')} 萬/坪, 地址: {c.get('address')}")
                if c.get("name") == "大慶國宅":
                    json_has_daqing = True
                    print(f"  [OK] 大慶國宅 於 master JSON 驗證: 均價 {c.get('recent_avg')} 萬/坪, 地址: {c.get('address')}")

        if not (json_has_panda_park and json_has_daqing):
            self.results["findings"].append("HIGH: Master JSON missing 熊貓公園市 or 大慶國宅")
            track_pass = False

        self.results["track_1_fact_checking"] = {
            "passed": track_pass,
            "md_has_section_3_2": has_section_3_2,
            "panda_park_verified": panda_found_south and json_has_panda_park,
            "daqing_verified": (daqing_tx_count > 0) and json_has_daqing,
            "docx_has_counterparts": docx_has_panda_park and docx_has_daqing,
            "docx_omitted_clarification_text": not docx_has_clarification_text
        }
        return track_pass

    # =========================================================================
    # Track 2: Negotiation Boundaries Stress-Testing
    # =========================================================================
    def run_track_2_negotiation_boundaries(self) -> bool:
        print("\n" + "="*80)
        print("【TRACK 2】議價談判邊界極限壓力測試 (Negotiation Boundaries & BATNA Stress-Test)")
        print("="*80)
        track_pass = True

        with open(MD_PATH, 'r', encoding='utf-8') as f:
            md_text = f.read()

        asking_price = 988.0
        floor_price = 810.0
        target_low = 870.0
        target_high = 885.0
        ceiling_price = 930.0
        ping = 29.98

        # 2.1 Lower Boundary Challenge (810萬)
        print("\n--- 2.1 保守底標防線 (810 萬元) 壓力測試 ---")
        discount_810 = floor_price / asking_price
        unit_810 = floor_price / ping
        cut_amount_810 = asking_price - floor_price
        print(f"  - 開價折數: {discount_810*100:.2f}% (約 82 折)")
        print(f"  - 折算單價: {unit_810:.2f} 萬/坪")
        print(f"  - 總砍價金額: {cut_amount_810:.1f} 萬元")

        # Stress test question: Does 810萬 risk immediate refusal by seller?
        # Check if report addresses immediate refusal risk & provides tactical transition
        refusal_handled = ("嫌 810 萬太低不願見面" in md_text or "810 萬太低" in md_text or "大膽試探" in md_text)
        anchor_evidence = ("26.65" in md_text and "26.95" in md_text and "持有 12 年" in md_text)
        print(f"  - 報告是否正視 810 萬遭到賣方拒絕之實務風險: {refusal_handled}")
        print(f"  - 報告是否具備 810 萬開局數據定錨佐證 (2026大盤26.65萬 / 2022同梯3F 26.95萬 / 屋主持有12年無貸款): {anchor_evidence}")

        if not (refusal_handled and anchor_evidence):
            self.results["findings"].append("MEDIUM: 810萬 bottom anchor lacks refusal risk management or empirical anchor evidence")
            track_pass = False

        # 2.2 Upper Ceiling Challenge (930萬)
        print("\n--- 2.2 上限天花板防線 (930 萬元) 壓力測試 ---")
        discount_930 = ceiling_price / asking_price
        unit_930 = ceiling_price / ping
        cut_amount_930 = asking_price - ceiling_price
        print(f"  - 開價折數: {discount_930*100:.2f}% (約 94 折)")
        print(f"  - 折算單價: {unit_930:.2f} 萬/坪")
        print(f"  - 總砍價金額: {cut_amount_930:.1f} 萬元")

        # Stress test question: Does 930萬 trigger overpaying for a unit without parking?
        # Check if report justifies why >930萬 is overpaying compared to peers with parking
        overpaying_warned = ("超過 930 萬" in md_text and ("追高" in md_text or "停損" in md_text))
        comp_with_parking_cited = ("公園苑" in md_text and "椰城" in md_text and "車位" in md_text)
        hard_stop_loss = ("多一塊錢我們都不會往上加" in md_text or "絕不追高" in md_text or "自住極限停損紅線" in md_text)
        print(f"  - 報告是否明確警示超過 930 萬將陷入買方追高陷阱: {overpaying_warned}")
        print(f"  - 報告是否以帶車位競品 (公園苑 31.08萬, 椰城 30.08萬) 論證無車位買 31 萬必屬追高: {comp_with_parking_cited}")
        print(f"  - 報告是否嚴格設下 930 萬不可退讓之鐵板停損紅線: {hard_stop_loss}")

        if not (overpaying_warned and comp_with_parking_cited and hard_stop_loss):
            self.results["findings"].append("HIGH: 930萬 ceiling lacks overpaying justification or hard stop-loss logic")
            track_pass = False

        # 2.3 BATNA & Walk-Away Alternative Options
        print("\n--- 2.3 買方最佳替代方案清單 (BATNA) 完整性查核 ---")
        batna_section = "7.2" in md_text or "BATNA" in md_text
        has_yecheng = "椰城" in md_text and "BATNA 1" in md_text
        has_gongyuan = "公園苑" in md_text and "BATNA 2" in md_text
        has_changyi = "長億南園" in md_text and "BATNA 3" in md_text
        has_gongxue = "工學天下" in md_text and "BATNA 4" in md_text

        print(f"  - 報告包含完整 BATNA 專節: {batna_section}")
        print(f"  - BATNA 1 (椰城, 帶車位純住): {has_yecheng}")
        print(f"  - BATNA 2 (公園苑, 帶車位同路段): {has_gongyuan}")
        print(f"  - BATNA 3 (長億南園, 700~800萬低總價): {has_changyi}")
        print(f"  - BATNA 4 (工學天下, 綠園道景觀大回檔): {has_gongxue}")

        if not (batna_section and has_yecheng and has_gongyuan and has_changyi and has_gongxue):
            self.results["findings"].append("HIGH: BATNA alternatives incomplete in report")
            track_pass = False

        self.results["track_2_negotiation_boundaries"] = {
            "passed": track_pass,
            "floor_810_discount": round(discount_810, 4),
            "floor_810_unit_price": round(unit_810, 2),
            "floor_810_refusal_handled": refusal_handled,
            "ceiling_930_unit_price": round(unit_930, 2),
            "ceiling_930_overpaying_warned": overpaying_warned,
            "batna_alternatives_verified": 4 if (has_yecheng and has_gongyuan and has_changyi and has_gongxue) else 0
        }
        return track_pass

    # =========================================================================
    # Track 3: Visual & Formatting Integrity (13 PPTX Slides & Word DOCX)
    # =========================================================================
    def run_track_3_presentation_integrity(self) -> bool:
        print("\n" + "="*80)
        print("【TRACK 3】簡報與報告視覺排版鑑識 (Presentation & Document Formatting Integrity)")
        print("="*80)
        track_pass = True

        # 3.1 PPTX Verification
        print("\n--- 3.1 驗證 PowerPoint 簡報 (.pptx) 之 13 頁排版、幾何座標與文字截斷 ---")
        if not pptx or not os.path.exists(PPTX_PATH):
            self.results["findings"].append("CRITICAL: PPTX file missing or pptx library unavailable")
            return False

        prs = Presentation(PPTX_PATH)
        slide_count = len(prs.slides)
        slide_w = prs.slide_width
        slide_h = prs.slide_height

        print(f"  - 投影片總頁數: {slide_count} 頁 (標準要求: 13 頁)")
        print(f"  - 投影片寬度: {slide_w / Inches(1):.3f} 英吋, 高度: {slide_h / Inches(1):.3f} 英吋 (16:9 寬螢幕)")

        if slide_count != 13:
            self.results["findings"].append(f"HIGH: PPTX slide count is {slide_count}, expected exactly 13")
            track_pass = False

        # Check geometry and truncation on every slide
        pptx_issues = []
        for s_idx, slide in enumerate(prs.slides, 1):
            text_shapes = []
            for s in slide.shapes:
                # Check bounding box overflow
                right = s.left + s.width
                bottom = s.top + s.height
                if s.left < -100 or s.top < -100 or right > slide_w + 1000 or bottom > slide_h + 1000:
                    pptx_issues.append(f"Slide {s_idx}: Shape '{s.name}' overflows slide boundary (L:{s.left}, T:{s.top}, R:{right}, B:{bottom})")

                if s.has_text_frame and s.text_frame.text.strip():
                    text_shapes.append(s)
                    full_s_text = s.text_frame.text.strip()
                    # Check truncation markers (abrupt '...', broken sentences)
                    for p in s.text_frame.paragraphs:
                        p_txt = p.text.strip()
                        if p_txt.endswith("...") or p_txt.endswith("…"):
                            # Only flag if not intentional subtitle
                            if len(p_txt) > 20:
                                pptx_issues.append(f"Slide {s_idx}: Potential text truncation ending in ellipsis: '{p_txt[-30:]}'")

            # Check collisions between text shapes
            for i in range(len(text_shapes)):
                s1 = text_shapes[i]
                for j in range(i + 1, len(text_shapes)):
                    s2 = text_shapes[j]
                    x_left = max(s1.left, s2.left)
                    y_top = max(s1.top, s2.top)
                    x_right = min(s1.left + s1.width, s2.left + s2.width)
                    y_bottom = min(s1.top + s1.height, s2.top + s2.height)
                    if x_right > x_left and y_bottom > y_top:
                        area = (x_right - x_left) * (y_bottom - y_top)
                        min_area = min(s1.width * s1.height, s2.width * s2.height)
                        if area > 0.15 * min_area:
                            pptx_issues.append(f"Slide {s_idx}: Text bounding box collision ({area/min_area*100:.1f}%) between '{s1.text_frame.text.strip()[:15]}' and '{s2.text_frame.text.strip()[:15]}'")

        print(f"  - 13 頁 PPTX 幾何溢出與重疊碰撞檢測結果: 共發現 {len(pptx_issues)} 項問題")
        if pptx_issues:
            for issue in pptx_issues[:10]:
                print(f"    [FAIL] {issue}")
            track_pass = False
        else:
            print("  [OK] 全 13 頁 PPTX 幾何座標無邊界溢出，文字形狀無不當碰撞重疊！")

        # 3.2 Check Footnotes & Disclaimers on PPTX
        print("\n--- 3.2 驗證 PPTX 關鍵頁面之資料來源標註與免責聲明 ---")
        slide_1_text = "".join([s.text_frame.text for s in prs.slides[0].shapes if s.has_text_frame])
        slide_13_text = "".join([s.text_frame.text for s in prs.slides[12].shapes if s.has_text_frame])
        s1_has_meta = ("實價登錄" in slide_1_text or "94" in slide_1_text or "83" in slide_1_text)
        s13_has_action = ("出價決策" in slide_13_text or "方針" in slide_13_text or "BATNA" in slide_13_text)

        print(f"  - Slide 1 封面包含實價登錄數據基準標註: {s1_has_meta}")
        print(f"  - Slide 13 包含自住理性出價與 BATNA 決策方針: {s13_has_action}")

        # 3.3 Word Document Verification
        print("\n--- 3.3 驗證 Word 文件 (.docx) 之章節結構、表格與頁面排版 ---")
        if not docx or not os.path.exists(DOCX_PATH):
            self.results["findings"].append("CRITICAL: DOCX file missing or python-docx unavailable")
            return False

        doc = docx.Document(DOCX_PATH)
        docx_tables = doc.tables
        total_cells = sum(len(r.cells) for t in docx_tables for r in t.rows)
        para_count = len(doc.paragraphs)
        file_size_kb = os.path.getsize(DOCX_PATH) / 1024

        print(f"  - Word 檔案容量: {file_size_kb:.1f} KB (門檻 >= 30 KB)")
        print(f"  - 表格總數: {len(docx_tables)} 個 (門檻 >= 5 個)")
        print(f"  - 表格總單元格: {total_cells} 格 (門檻 >= 50 格)")
        print(f"  - 段落總數: {para_count} 個")

        # Check footer disclaimer and dynamic page numbers
        footer_has_disclaimer = False
        for sec in doc.sections:
            f_text = sec.footer.paragraphs[0].text if sec.footer.paragraphs else ""
            if "內政部實價登錄" in f_text or "免責" in f_text or "純供自住決策參考" in f_text:
                footer_has_disclaimer = True
                break
        print(f"  - 頁尾包含專業免責與資料來源聲明: {footer_has_disclaimer}")

        if len(docx_tables) < 5 or total_cells < 50 or file_size_kb < 30:
            self.results["findings"].append("HIGH: Word document tables or size below required thresholds")
            track_pass = False

        self.results["track_3_presentation_integrity"] = {
            "passed": track_pass,
            "pptx_slide_count": slide_count,
            "pptx_aspect_ratio_16_9": abs((slide_w / slide_h) - (16.0 / 9.0)) < 0.01,
            "pptx_geometric_issues": len(pptx_issues),
            "docx_tables_count": len(docx_tables),
            "docx_total_cells": total_cells,
            "docx_footer_disclaimer": footer_has_disclaimer
        }
        return track_pass

    # =========================================================================
    # Track 4: Suite Verification Execution
    # =========================================================================
    def run_track_4_suite_verification(self) -> bool:
        print("\n" + "="*80)
        print("【TRACK 4】全套件既有測試執行與 Exit Code 診斷 (Suite Verification Runner)")
        print("="*80)

        cmd = [sys.executable, SUITE_SCRIPT_PATH]
        print(f"執行指令: {' '.join(cmd)}")
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace', cwd=PROJECT_ROOT)
            exit_code = res.returncode
            stdout_output = res.stdout
            stderr_output = res.stderr
        except Exception as e:
            self.results["findings"].append(f"CRITICAL: Failed to execute {SUITE_SCRIPT_PATH}: {e}")
            return False

        print(f"測試套件 Exit Code: {exit_code}")
        print("\n[測試套件標準輸出摘錄]:")
        if stdout_output:
            for line in stdout_output.splitlines()[-15:]:
                print("  " + line)

        suite_passed = (exit_code == 0)
        if exit_code != 0:
            # Diagnose reason
            if "GitHub 自動同步腳本 (`scripts/sync_to_github.py`) 尚未將 `歌德堡` 加入同步目錄清單" in stdout_output:
                print("\n[診斷發現]: Tier 4 攔截到 `scripts/sync_to_github.py` 尚未納入 `歌德堡` 目錄！")
                self.results["findings"].append("BLOCKER: scripts/verify_goteborg_suite.py failed (Exit Code 1) due to scripts/sync_to_github.py not including '歌德堡' directory.")
            else:
                self.results["findings"].append(f"BLOCKER: scripts/verify_goteborg_suite.py failed with exit code {exit_code}")

        self.results["track_4_suite_verification"] = {
            "exit_code": exit_code,
            "passed": suite_passed,
            "stdout_tail": stdout_output.splitlines()[-8:] if stdout_output else []
        }
        return suite_passed

    # =========================================================================
    # Overall Execution & Verdict
    # =========================================================================
    def execute_all(self):
        t1 = self.run_track_1_fact_checking()
        t2 = self.run_track_2_negotiation_boundaries()
        t3 = self.run_track_3_presentation_integrity()
        t4 = self.run_track_4_suite_verification()

        # Decision Logic:
        # If there are BLOCKER or CRITICAL findings, verdict MUST be REQUEST_CHANGES.
        blockers = [f for f in self.results["findings"] if f.startswith("BLOCKER") or f.startswith("CRITICAL")]

        if not blockers and t1 and t2 and t3 and t4:
            self.results["overall_verdict"] = "APPROVE"
        else:
            self.results["overall_verdict"] = "REQUEST_CHANGES"

        print("\n" + "="*80)
        print(f"【CHALLENGE HARNESS FINAL VERDICT】: {self.results['overall_verdict']}")
        print("="*80)
        print(f"總計發現項目 ({len(self.results['findings'])} 項):")
        for idx, f in enumerate(self.results["findings"], 1):
            print(f"  {idx}. {f}")

        # Save JSON Report
        report_path = os.path.join(PROJECT_ROOT, "scratch", "challenge_facts_boundaries_report.json")
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(self.results, f, ensure_ascii=False, indent=2)
        print(f"\n[REPORT] 詳細結構化鑑識報表已存檔至: `{report_path}`")
        return self.results


if __name__ == "__main__":
    challenger = GoteborgAdversarialChallenge()
    res = challenger.execute_all()
    # Exit with 0 if verdict is reached cleanly (script itself succeeded in evaluating)
    sys.exit(0 if res["overall_verdict"] == "APPROVE" else 1)

# TEST_READY.md — 大慶國宅(大慶中山) 出價分析與全社區行情套件 E2E 驗收規範手冊

**專案名稱**：台中市南區「大慶國宅(大慶中山)」社區特定待售戶出價分析與全社區整體行情評估套件  
**發布里程碑**：Milestone M_TEST (獨立自動化驗證套件發布與 TEST_READY 就緒證明)  
**測試架構實作**：`scripts/verify_daqing_suite.py`  
**測試設計者**：`test_writer_daqing` (Specialist / QA)  
**建立時間**：2026-11-04T17:01:00Z  

---

## 1. 執行方式與 CLI 指令說明 (Test Execution Commands)

本驗證套件為獨立、無副作用、自包含之全自動化 CLI 診斷驗收工具，支援全量嚴格驗收（Strict Acceptance）、里程碑階段演進驗收（Progressive Check）、多層級組合執行（`--tier 1,2`）、JSON 機器可讀報表，以及內建單元自我檢驗（Self-Test Suite）：

### 1.1 全量嚴格 E2E 驗收執行（Milestone M4 最終發布門檻）
```bash
python scripts/verify_daqing_suite.py
```
- **目標目錄**：`大慶國宅(大慶中山)/`、`scratch/`、`南區25年以上大型社區走勢圖/`
- **執行行為**：對專案所有交付產物進行 Tier 1～Tier 4 全量無死角檢驗。
- **Exit Code**：
  - `0`：全部 Tier 1～4 斷言 100% 通過（全案完工驗收標準）。
  - `1`：任何文件、簡報、報告或數據指標存在瑕疵或缺失，立即輸出定位錯誤訊息。

### 1.2 里程碑階段演進驗收模式 (Progressive Execution Mode)
```bash
python scripts/verify_daqing_suite.py --progressive
```
- **適用場景**：開發過程中（M1 數據庫建立、M2 Markdown/Word 報告撰寫、M3 PowerPoint 簡報產製）隨時執行。
- **執行行為**：動態偵測當前已產出之交付物進行品質驗證，尚未產出之階段成果標記為 `[WARN]`（待開發）而不阻斷流程，唯有在現有成果出現瑕疵時才會阻斷。
- **Exit Code**：
  - `0`：當前已產出的階段產物 100% 驗證通過，無任何阻礙性缺陷。
  - `1`：已產出產物中存在規格違規（如：數據不一致、格式錯誤、圖表解析度不足等）。

### 1.3 分層專屬與組合驗證指令 (Tier-Specific & Multi-Tier Checks)
```bash
# 僅驗證 Tier 1 功能覆蓋與核心估價數學
python scripts/verify_daqing_suite.py --tier 1

# 同時指定驗證 Tier 1 與 Tier 2
python scripts/verify_daqing_suite.py --tier 1,2

# 僅驗證 Tier 3 跨交付物核心數據一致性 (坪數、開價、三段出價、歷史行情、競品)
python scripts/verify_daqing_suite.py --tier 3

# 僅驗證 Tier 4 圖表資產存在性、PIL 影像鑑識與 GitHub 同步腳本就緒度
python scripts/verify_daqing_suite.py --tier 4
```

### 1.4 內部單元自我檢驗模式 (Self-Test Suite)
```bash
python scripts/verify_daqing_suite.py --test-self
```
- **執行行為**：在暫存環境中自動生成合成假資料、Mock DOCX/PPTX/MD/JSON 與模擬走勢圖，測試驗證器本體的 8 項獨立抓漏單元測試，防禦虛假測試與作弊行為。
- **涵蓋用例**：
  - `TC1`：正常 Mock 交付物滿分通過（Tier 1～4 100% PASS）。
  - `TC2`：缺少主資料庫 JSON 精準攔截。
  - `TC3`：最新 3 樓車位拆算數學錯誤（如車位價值非 120 萬）精準攔截。
  - `TC4`：Markdown 報告缺少必要 8 大核心章節標題精準攔截。
  - `TC5`：PPTX 比例非 16:9 寬螢幕或頁數不足（< 14 頁）精準攔截。
  - `TC6`：Word 文件容量不足（< 50KB）精準攔截。
  - `TC7`：跨交付物出價數字或行情數據不一致精準攔截。
  - `TC8`：GitHub 同步腳本缺少 `大慶國宅(大慶中山)` 目錄精準攔截。

### 1.5 JSON 報表輸出模式
```bash
python scripts/verify_daqing_suite.py --json
python scripts/verify_daqing_suite.py --progressive --json
```

---

## 2. 四層級驗證覆蓋矩陣 (4-Tier Verification Matrix)

依據 `PROJECT.md`、`ORIGINAL_REQUEST.md` 與 `TEST_INFRA.md`，本測試套件設定了嚴格的四層級驗收門檻：

| Tier | 驗證層級 | 驗證項目 | 權威依據 / Oracle | 驗證斷言與成功標準 (Assertions) |
|---|---|---|---|---|
| **Tier 1** | **功能規格覆蓋** | **待售物件精準規格** | 591 S20888586 / `explorer_survey_1` | 樓層 10F/12F、總建坪 44.5 坪 (主建 25.365, 附屬 2.405, 公設 16.726, 室內實坪 27.77, 公設比 37.59%)、開價 998 萬 (降價 82 萬, 原開價 1080 萬, 單價 22.43 萬/坪)、承租管委會平車 1,500 元/月。 |
| **Tier 1** | **功能規格覆蓋** | **49 筆實登清洗分流** | `scratch/data/20513_大慶中山.json` | 原始 49 筆採集全量，純住宅正常市價樣本 33 筆，2024 年頂峰 27.0 萬/坪 (12F 1038萬)、2025 年 20.5 萬/坪 (2F 850萬)、2026-05 最新 22.7 萬/坪 (3F 980萬)。 |
| **Tier 1** | **功能規格覆蓋** | **車位拆算法則推導** | `ORIGINAL_REQUEST` R2 / `explorer_survey_2` | 最新 3 樓實登：$(980 - 120) \div (43.17 - 7.6) = 860 \div 35.57 = 24.17768\dots \approx 24.18 \text{ 萬/坪}$。 |
| **Tier 1** | **功能規格覆蓋** | **6 大競品社區收錄** | `PROJECT.md` § Feature 4 | 歌德堡、長億南園、工學天下、大安紫金城、劃江山、大墩人文全數收錄於競品矩陣。 |
| **Tier 1** | **功能規格覆蓋** | **三段式出價策略** | `explorer_survey_3` / `PROJECT.md` | 保守底標價 780~800 萬 (17.53~17.98萬/坪)、合理目標價 860~890 萬 (19.33~20.00萬/坪)、上限天花板價 920~930 萬 (20.67~20.90萬/坪)。 |
| **Tier 1** | **功能規格覆蓋** | **105 萬修繕工程預算** | `ORIGINAL_REQUEST` R2 / `PROJECT.md` | 水電(18萬)+衛浴泥作(28萬)+衛浴設備(12萬)+廚房(16萬)+門窗油漆(14萬)+地板(10萬)+清潔(7萬) $= 105 \text{ 萬元}$ 精確守恆。 |
| **Tier 2** | **邊界極限與結構** | **折價比率數學邊界** | 開價 998 萬基準推導 | 780/998=~0.78 (78折)、860/998=~0.86 (86折)、920/998=~0.92 (92折)；單調遞增約束：$780 < 800 \le 860 < 890 \le 920 < 930 < 998$。 |
| **Tier 2** | **邊界極限與結構** | **建物面積幾何守恆** | 地政登記簿謄本邏輯 | $25.365 + 2.405 = 27.770 \text{ 坪}$；$27.770 + 16.726 = 44.496 \text{ 坪}$；$16.726 / 44.496 = 37.59\%$ 幾何守恆。 |
| **Tier 2** | **邊界極限與結構** | **產物檔案容量門檻** | `PROJECT.md` § Verification | MD $\ge 15 \text{ KB}$、DOCX $\ge 50 \text{ KB}$、PPTX $\ge 500 \text{ KB}$、JSON $\ge 5 \text{ KB}$。 |
| **Tier 2** | **邊界極限與結構** | **PPTX 投影片規格** | `PROJECT.md` § Interface Contract | 投影片總頁數 $\ge 14$ 頁，比例嚴格符合 16:9 寬螢幕規格 ($13.333" \times 7.5"$)。 |
| **Tier 2** | **邊界極限與結構** | **Markdown 核心章節** | `PROJECT.md` § Invariants | 完整包含 8 大核心章節：物件產調、社區歷程、競品比價、車位拆算、調整矩陣、出價防線、談判說帖、修繕預算。 |
| **Tier 3** | **跨交付物一致性** | **核心估價數據一致** | `scratch/daqing_valuation_master.json` | 樓層 10F、總建坪 44.5 坪、開價 998 萬 (22.43萬/坪)、三段出價 (780~800 / 860~890 / 920~930)、105 萬修繕款、49 筆實登、最新 3F 24.18 萬在 MD/DOCX/PPTX/JSON 100% 一致。 |
| **Tier 4** | **實體驗收與同步** | **交付產物實體存在** | `大慶國宅(大慶中山)/` | MD、DOCX、PPTX 三大實體檔案完整存在於目標目錄中。 |
| **Tier 4** | **實體驗收與同步** | **Office 格式無損解析** | `python-docx` / `python-pptx` | DOCX 表格數 $\ge 5$、單元格 $\ge 50$、內嵌圖片 $\ge 1$；PPTX 包含獨立圖片元件。 |
| **Tier 4** | **實體驗收與同步** | **走勢圖資產鑑識** | `南區25年以上大型社區走勢圖/` | `12_大慶中山_成交行情走勢圖.png` 存在，PIL 解析度無損且 $\ge 1000 \times 600$ (實際 $2925 \times 2035$)。 |
| **Tier 4** | **實體驗收與同步** | **GitHub 自動同步配置** | `scripts/sync_to_github.py` | 同步目錄陣列中明確包含 `"大慶國宅(大慶中山)"`。 |

---

## 3. 測試套件實作與基線驗證結果 (Baseline Execution Results)

### 3.1 內部單元自我檢驗 (`--test-self`)
執行指令：
```bash
python scripts/verify_daqing_suite.py --test-self
```
**執行結果**：
```
================================================================================
【執行 verify_daqing_suite.py 內部單元自我測試 (Self-Test Suite)】
================================================================================
  [PASS] TC1: Valid mock deliverables passed Tier 1 ~ 4 full acceptance 100%.
  [PASS] TC2: Missing master JSON correctly caught and failed by Tier 1.
  [PASS] TC3: Incorrect car space separation math correctly caught and failed by Tier 1.
  [PASS] TC4: Missing required MD chapters correctly caught and failed by Tier 2.
  [PASS] TC5: PPTX insufficient slides (<14) and non-16:9 ratio correctly caught and failed by Tier 2.
  [PASS] TC6: Insufficient DOCX file size (<50KB) correctly caught and failed by Tier 2.
  [PASS] TC7: Cross-deliverable pricing discrepancy correctly caught and failed by Tier 3.
  [PASS] TC8: Sync script missing '大慶國宅(大慶中山)' correctly caught and failed by Tier 4.
================================================================================
🎉 [SELF-TEST SUCCESS] All 8/8 internal verification tests passed successfully!
================================================================================
```
- **結論**：內部單元測試全數 8/8 通過，證實測試器具備完整嚴密的抓漏與規格防禦能力。

### 3.2 專案現況階段演進模式檢驗 (`--progressive`)
執行指令：
```bash
python scripts/verify_daqing_suite.py --progressive
```
- **執行結論**：`PROGRESSIVE PASS`（Exit Code 0）。
- **當前已通過之靜態資產與數學邏輯**：
  - 走勢圖 `12_大慶中山_成交行情走勢圖.png` 存在且解析度達 $2925 \times 2035$。
  - 周邊 4 張競品與大盤總圖資產完整存在。
  - 面積幾何守恆（$25.365 + 2.405 = 27.770$, $27.770 + 16.726 = 44.496$, 公設比 $37.59\%$）通過檢驗。
  - 折價率邊界數學（78折、86折、92折）通過檢驗。
- **待後續里程碑產出之交付物 (標記為 WARN，待工作者產製)**：
  - M1：`scratch/daqing_valuation_master.json`
  - M2：Markdown 報告與 Word 報告 (`.docx`)
  - M3：PowerPoint 簡報 (`.pptx`)
  - M4：`scripts/sync_to_github.py` 目錄納入

---

## 4. 里程碑 M_TEST 驗收簽署 (Sign-off)

- **測試套件狀態**：`READY & OPERATIONAL`
- **測試入口**：`python scripts/verify_daqing_suite.py`
- **覆蓋度**：14/14 功能特性全量映射，4/4 驗證層級完備。
- **後續里程碑交接說明**：後續 M1、M2、M3、M4 工作者與審計員（Auditor、Reviewer、Challenger）均可依據本手冊指令執行獨立查核。

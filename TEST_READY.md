# TEST_READY.md — E2E 測試套件就緒與全方位驗收規範手冊

**專案名稱**：台中市南屯區30年以上大型指標社區實價登錄走勢圖與完整分析套件  
**發布里程碑**：Milestone M_TEST (獨立自動化 E2E 驗證套件發布與 TEST_READY)  
**測試設計者**：teamwork_test_writer_nantun (Specialist / QA)  
**主要驗證腳本**：`scripts/verify_nantun_suite.py`  
**最後更新時間**：2026-10-29T16:59:00Z  

---

## 1. 執行方式與 CLI 指令說明 (Test Execution Commands)

本驗證套件具備強大、獨立、無副作用之自動化 CLI 診斷能力，支援全量嚴格驗收、里程碑階段演進驗收 (Progressive Check)、各圖層獨立檢驗、詳細除錯輸出、JSON 機器可讀報表，以及內建單元自我檢驗（Self-Test Suite）：

### 1.1 全量嚴格 E2E 驗收執行（Milestone M5 最終發布門檻）
```bash
python scripts/verify_nantun_suite.py
```
- **目標目錄**：`南屯區30年以上大型社區走勢圖/`
- **執行行為**：對專案所有交付產物進行 Tier 1～Tier 4 全量無死角檢驗。
- **Exit Code**：
  - `0`：全部 Tier 1～4 斷言 100% 通過（全案驗收完工）。
  - `1`：任何圖檔、簡報、報告、網頁或數據指標存在瑕疵或缺失，立即輸出定位診斷訊息。

### 1.2 里程碑階段演進驗收模式 (Progressive Execution Mode)
```bash
python scripts/verify_nantun_suite.py --progressive
```
- **適用場景**：開發過程中（M1 數據清洗、M2 繪圖、M3 撰寫報告、M4 產製簡報與網頁）隨時執行。
- **執行行為**：動態偵測當前已產出之交付物進行品質驗證，尚未產出之階段成果標記為 `[WARN]`（待開發）而不阻斷流程，唯有在現有成果出現瑕疵時才會阻斷。
- **Exit Code**：
  - `0`：當前已產出的階段產物 100% 驗證通過，無任何阻礙性缺陷。
  - `1`：已產出產物中存在規格違規（如：1F店面污染、1997年社區混入、解析度不足等）。

### 1.3 分層/分階段專屬驗證指令 (Layer-Specific Checks)
```bash
# 僅驗證 M1 數據清洗層（入選條件、邊界排除、純住宅交易純淨度）
python scripts/verify_nantun_suite.py --data-layer --progressive

# 僅驗證 M2 走勢圖產出層（75張社區走勢圖 + 1張大盤總圖 + 根目錄鏡像 + PIL高解析度鑑識）
python scripts/verify_nantun_suite.py --chart-layer --progressive

# 僅驗證 M3/M4 分析報告、PPTX 簡報與互動 HTML 儀表板
python scripts/verify_nantun_suite.py --doc-layer --progressive

# 依特定 Tier 進行檢驗
python scripts/verify_nantun_suite.py --tier1   # Feature Coverage
python scripts/verify_nantun_suite.py --tier2   # Boundary & Corner Cases
python scripts/verify_nantun_suite.py --tier3   # Cross-Feature Consistency
python scripts/verify_nantun_suite.py --tier4   # Acceptance & Forensics
```

### 1.4 內部單元自我檢驗模式 (Self-Test Suite)
```bash
python scripts/verify_nantun_suite.py --test-self
```
- **執行行為**：在暫存環境中自動生成合成假資料與模擬圖檔，測試驗證器本體的 8 項獨立抓漏單元測試，防禦虛假測試與作弊行為。
- **涵蓋用例**：
  - TC1：正常 75 社區全量交付物滿分通過（Tier 1～4 100% PASS）。
  - TC2：1997 年完工排除社區（如龍歡喜大廈）滲漏精準攔截。
  - TC3：<100 戶邊界社區（如心田 99 戶）滲漏精準攔截。
  - TC4：1 樓店面污染交易精準攔截。
  - TC5：缺少社區走勢圖 PNG 或編號中斷精準攔截。
  - TC6：跨交付物（JSON/PNG/MD/PPTX/HTML）名稱或序號不一致精準攔截。
  - TC7：走勢圖解析度小於 3000x2000 或圖檔損毀精準攔截。
  - TC8：HTML 儀表板缺少即時搜尋或 Lightbox 彈窗邏輯精準攔截。

### 1.5 JSON 報表輸出模式
```bash
python scripts/verify_nantun_suite.py --progressive --json
```

---

## 2. 四層級驗證覆蓋矩陣 (4-Tier Verification Matrix)

依據 `ORIGINAL_REQUEST.md`、`PROJECT.md` 與 `explorer_survey_nantun_2/handoff.md` 之權威規格，本測試套件設定了嚴格的四層級驗收門檻：

| Tier | 驗證層級 | 驗證項目 | 權威資料源 / Oracle | 驗證斷言與成功標準 (Assertions) |
|---|---|---|---|---|
| **Tier 1** | **產物覆蓋性與基本結構** (Feature Coverage) | **75 棟社區獨立走勢圖 PNG** | `PROJECT.md` § Feature 5 | `南屯區30年以上大型社區走勢圖/{排名:02d}_{社區名稱}_成交行情走勢圖.png`，01～75 全數 75 張必須存在且非空。 |
| **Tier 1** | **產物覆蓋性與基本結構** | **全區大盤走勢圖與鏡像** | `PROJECT.md` § Feature 6 | 目標目錄 `00_台中市南屯區30年以上大型社區_大盤匯總走勢圖.png` 與專案根目錄鏡像檔必須同時存在。 |
| **Tier 1** | **產物覆蓋性與基本結構** | **深度歷史分析報告 (Markdown)** | `PROJECT.md` § Feature 8 | `台中市南屯區30年以上大型社區歷史實價登錄行情與大盤走勢分析報告.md` 必須存在，且完整涵蓋 5 大核心章節：歷史週期演變、四大維度深度剖析、計量預估模型、75社區完整名冊、重點社區專題與議價指南。 |
| **Tier 1** | **產物覆蓋性與基本結構** | **PowerPoint 分析簡報檔 (.pptx)** | `PROJECT.md` § Feature 9 | `台中市南屯區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx` 與根目錄鏡像檔必須同時存在。 |
| **Tier 1** | **產物覆蓋性與基本結構** | **互動網頁儀表板 (`index.html`)** | `PROJECT.md` § Feature 10 | `index.html` 必須存在，內含 `searchInput` 搜尋框、大盤總圖引用、以及全數 75 棟社區展示卡片。 |
| **Tier 2** | **邊界案例、入選條件與純淨度** (Boundary & Invariants) | **75 社區入選法定條件** | `explorer_survey_nantun_2` § 4 | 75 棟入選社區之使用執照建築完成日全數 $\le 1996-12-31$（屋齡 30 年以上），且總戶數全數 $\ge 100$ 戶。 |
| **Tier 2** | **邊界案例、入選條件與純淨度** | **1997 年交界排除社區檢驗** | `explorer_survey_nantun_2` § 邊界判定表 | 嚴格確認 10 棟 1997 年（民國 86 年）完工之邊界社區（龍歡喜大廈、三采市政天地、同隆上第二期、親家莊、熊貓凱撒、東興寧境、英棋尊邸、巴黎大凱旋、泉福天郡、大東家春風）**絕對未混入** 75 棟正式名冊。 |
| **Tier 2** | **邊界案例、入選條件與純淨度** | **<100 戶邊界排除社區檢驗** | `explorer_survey_nantun_2` § 邊界判定表 | 嚴格確認 10 棟 90～99 戶邊界社區（心田 99、寶祥榕園 99、寶樺 98、華爾街天廈 98、采龍大廈 97、成吉思汗 95、統健大樓 94、光田生活觀邸 92、傑出名門 90、巴黎四季 90）**絕對未混入** 75 棟正式名冊。 |
| **Tier 2** | **邊界案例、入選條件與純淨度** | **實價登錄純住宅交易純淨度** | `PROJECT.md` § Contract M1 | 在清洗後資料庫 `scratch/all_75_nantun_communities_transactions.json` 中：<br>• 1 樓店面筆數精確為 **0**（`is_first_floor_shop == false`，且無 1F/店面，保護 11F/21F 純住）<br>• 特殊/親友交易筆數精確為 **0**（`is_special == 0`，無親友/特殊備註）<br>• 成交淨單價嚴格落在 **[5.0, 75.0] 萬元/坪** 有效自住區間<br>• 交易年份全數 $\ge 2011.0$（民國 101 年實登全面上路後）。 |
| **Tier 3** | **跨交付物全維度一致性對齊** (Cross Consistency) | **名冊與數據庫排名一致性** | 跨模組介面契約 | `qualified_nantun_communities.json` 與 `all_75_nantun_communities_transactions.json` 之社區名冊、排名（1～75）與戶數排序 **100% 完全相符**。 |
| **Tier 3** | **跨交付物全維度一致性對齊** | **圖檔命名與名冊一致性** | 檔案命名規格 | 75 張個別社區走勢圖之序號（01～75）與社區名稱與名冊 **100% 吻合**。 |
| **Tier 3** | **跨交付物全維度一致性對齊** | **Markdown 報表列一致性** | 分析報告表格 | Markdown 報告中核心行情統計表之 75 筆列序與社區名稱與名冊 **100% 吻合**。 |
| **Tier 3** | **跨交付物全維度一致性對齊** | **PPTX 投影片涵蓋率** | 簡報規格 | PPTX 簡報之獨立社區分析頁面 100% 完整涵蓋名冊之全數 75 棟社區。 |
| **Tier 3** | **跨交付物全維度一致性對齊** | **HTML 儀表板卡片一致性** | 前端資料綁定 | `index.html` 包含之社區卡片資料屬性（`data-name`）100% 完整涵蓋名冊 75 棟社區。 |
| **Tier 4** | **實體檔案鑑識與驗收場景** (Acceptance Forensics) | **圖表高解析度鑑識 (PIL)** | 圖表引擎規格 | 透過 Pillow (PIL) 實體驗收 76 張 PNG 圖檔：<br>• 每一張圖皆可完整加載像素，**零損毀**<br>• 每一張圖解析度寬度 $\ge 3000$ 且高度 $\ge 2000$（標準 3120x2080）。 |
| **Tier 4** | **實體檔案鑑識與驗收場景** | **PPTX 簡報規格鑑識 (python-pptx)** | 簡報交付規格 | 透過 `python-pptx` 實體驗證：<br>• 總頁數 $\ge 60$ 頁（預計包含封面、大盤、成因、模型、總表與 75 棟社區獨立頁面，共約 86 頁）<br>• 投影片比例嚴格符合 **16:9 寬螢幕**（長寬比介於 1.76～1.79 之間）。 |
| **Tier 4** | **實體檔案鑑識與驗收場景** | **Markdown 報告容量與編碼** | 報告深度要求 | 分析報告檔案大小 $\ge 40 \text{ KB}$，且為無 BOM 之標準 UTF-8 編碼。 |
| **Tier 4** | **實體檔案鑑識與驗收場景** | **HTML 儀表板容量與互動邏輯** | 前端功能規格 | `index.html` 檔案大小 $\ge 100 \text{ KB}$，UTF-8 編碼，且必須包含完整 **Lightbox 燈箱全螢幕圖片放大檢視邏輯**。 |

---

## 3. 當前基線測試結果 (M_TEST Baseline Execution Results)

### 3.1 內部單元自我檢驗 (`--test-self`)
執行 `python scripts/verify_nantun_suite.py --test-self`：
```
================================================================================
【執行 verify_nantun_suite.py 內部單元自我測試 (Self-Test Suite)】
================================================================================
  [PASS] TC1: Valid mock deliverables passed Tier 1 ~ 4 full acceptance 100%.
  [PASS] TC2: 1997 excluded community correctly caught and rejected by Tier 2.
  [PASS] TC3: <100 household edge-case community correctly caught and rejected by Tier 2.
  [PASS] TC4: Contaminated 1F shop transaction correctly caught and rejected by Tier 2.
  [PASS] TC5: Missing PNG chart correctly caught and failed by Tier 1.
  [PASS] TC6: Community name discrepancy across deliverables caught by Tier 3.
  [PASS] TC7: Low resolution image (<3000x2000) correctly caught by Tier 4.
  [PASS] TC8: Missing Lightbox modal in HTML dashboard correctly caught by Tier 4.
================================================================================
🎉 [SELF-TEST SUCCESS] All 8/8 internal verification tests passed successfully!
```
- **自我檢驗結論**：8 項內部單元測試全數 100% 通過，證實測試器具備真實鑑別能力，能有效防範偽造數據與作弊行為。

### 3.2 階段演進模式基線檢驗 (`--progressive`)
執行 `python scripts/verify_nantun_suite.py --progressive`：
```
================================================================================
【台中市南屯區30年以上大型指標社區走勢圖 — 獨立自動化驗證套件】
Target Directory: ...\南屯區30年以上大型社區走勢圖
Project Root    : ...
Execution Mode  : PROGRESSIVE (階段演進驗收)
================================================================================

✅ PASS Tier 1: Feature Coverage (產物覆蓋性與基本結構)
  [WARN] Target directory `南屯區30年以上大型社區走勢圖` not created yet (Pending M1/M2).

✅ PASS Tier 2: Boundary & Corner Cases (邊界案例、入選條件與資料純淨度)
  [OK] All 75 communities strictly meet completion date invariant (build_date <= 1996-12-31).
  [OK] All 75 communities strictly meet household invariant (households >= 100).
  [OK] Verified all 10 excluded 1997 communities are strictly excluded from roster.
  [OK] Verified all 10 borderline 90~99 unit communities are strictly excluded from roster.
  [WARN] Transactions dataset `all_75_nantun_communities_transactions.json` not found. (Pending M1 ingestion & cleaning).

✅ PASS Tier 3: Cross-Feature Consistency (跨交付物全維度一致性對齊)
  [WARN] Transactions dataset not available for Tier 3 consistency check.
  ...

✅ PASS Tier 4: Real-World Acceptance Scenarios (實體檔案鑑識與驗收場景)
  [WARN] No PNG charts available for PIL image forensic check. (Pending M2).
  ...

================================================================================
🎉【驗收結論：PROGRESSIVE PASS】當前已產出之階段成果全數通過驗證！無任何阻礙性缺陷。
================================================================================
```
- **Exit Code**：`0`。
- **結論**：名冊 75 棟社區法定條件、1997 排除與 <100 戶排除完全通過驗證，守門員已就位。

### 3.3 全量嚴格驗收基線檢驗（無參數執行）
執行 `python scripts/verify_nantun_suite.py`：
- **Exit Code**：`1`（預期行為）。
- **結論**：在目標資料夾與產物尚未完全由 M1～M4 生成前，驗證器嚴格防禦，精確列出所有待完成交付物，杜絕未完工提前通過。

---

## 4. 下游里程碑 (M1 ~ M5) 驗收與交付準則 (Milestone Exit Criteria)

各里程碑負責代理人在完成該階段產物後，必須執行指定驗收指令確認品質無誤：

### 4.1 Milestone M1（數據採集、五重清洗與結構化儲存）
- **產出物**：`scratch/all_75_nantun_communities_transactions.json`、`scratch/target_nantun_communities_clean.json`
- **驗收指令**：
  ```bash
  python scripts/verify_nantun_suite.py --data-layer --progressive
  ```
- **通過門檻**：75 社區純住宅筆數 $>0$；0 筆 1 樓店面；0 筆特殊交易；單價全在 [5.0, 75.0] 萬/坪；年份 $\ge 2011.0$。

### 4.2 Milestone M2（75 張個別走勢圖與大盤總圖繪製）
- **產出物**：`南屯區30年以上大型社區走勢圖/00_...png`、`01_...png`～`75_...png`、根目錄大盤鏡像圖
- **驗收指令**：
  ```bash
  python scripts/verify_nantun_suite.py --chart-layer --progressive
  ```
- **通過門檻**：76 張圖檔全數存在且非空；PIL 鑑識無損毀；解析度全數 $\ge 3000 \times 2000$。

### 4.3 Milestone M3 & M4（深度分析報告、PPTX 簡報與網頁儀表板）
- **產出物**：Markdown 分析報告、16:9 PPTX 簡報（含鏡像）、深色主題 `index.html`
- **驗收指令**：
  ```bash
  python scripts/verify_nantun_suite.py --doc-layer --progressive
  ```
- **通過門檻**：報告 $\ge 40 \text{ KB}$ 且含 5 大核心章節；簡報 $\ge 60$ 頁且為 16:9 比例；網頁 $\ge 100 \text{ KB}$ 且含搜尋與 Lightbox 邏輯。

### 4.4 Milestone M5（全套件 E2E 最終交付與發布）
- **驗收指令**：
  ```bash
  python scripts/verify_nantun_suite.py
  ```
- **最終驗收門檻**：
  1. **Exit Code 必須為 `0`**。
  2. 終端輸出包含：  
     `🎉【驗收結論：FULL SUITE PASS】全 4 個驗證層級 (Tier 1 ~ 4) 100% 驗收通過！符合專案最高交付標準。`
  3. 執行 `python scripts/sync_to_github.py "<Commit Message>"` 同步至 GitHub `main` 與 `gh-pages` 雙分支。

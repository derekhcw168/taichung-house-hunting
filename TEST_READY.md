# TEST_READY.md — 歌德堡出價分析與全社區行情套件 E2E 驗收規範手冊

**專案名稱**：台中市南區「歌德堡」社區特定待售戶出價分析與全社區整體行情評估套件  
**發布里程碑**：Milestone M_TEST (獨立自動化驗證套件發布與 TEST_READY)  
**測試設計者**：`test_writer_goteborg` (Specialist / QA)  
**主要驗證腳本**：`scripts/verify_goteborg_suite.py`  
**建立時間**：2026-11-04T16:32:00Z  

---

## 1. 執行方式與 CLI 指令說明 (Test Execution Commands)

本驗證套件為獨立、不依賴外部服務、無副作用之自動化 CLI 診斷工具，支援全量嚴格驗收、里程碑階段演進驗收 (Progressive Check)、各 Tier 獨立檢驗、JSON 機器可讀報表，以及內建單元自我檢驗（Self-Test Suite）：

### 1.1 全量嚴格 E2E 驗收執行（Milestone M4 最終發布門檻）
```bash
python scripts/verify_goteborg_suite.py
```
- **目標目錄**：`歌德堡/`、`scratch/`、`南區25年以上大型社區走勢圖/`
- **執行行為**：對專案所有交付產物進行 Tier 1～Tier 4 全量無死角檢驗。
- **Exit Code**：
  - `0`：全部 Tier 1～4 斷言 100% 通過（全案驗收完工）。
  - `1`：任何文件、簡報、報告或數據指標存在瑕疵或缺失，立即輸出精確定位診斷訊息。

### 1.2 里程碑階段演進驗收模式 (Progressive Execution Mode)
```bash
python scripts/verify_goteborg_suite.py --progressive
```
- **適用場景**：開發過程中（M1 數據庫建立、M2 Markdown/Word 報告撰寫、M3 PowerPoint 簡報產製）隨時執行。
- **執行行為**：動態偵測當前已產出之交付物進行品質驗證，尚未產出之階段成果標記為 `[WARN]`（待開發）而不阻斷流程，唯有在現有成果出現瑕疵時才會阻斷。
- **Exit Code**：
  - `0`：當前已產出的階段產物 100% 驗證通過，無任何阻礙性缺陷。
  - `1`：已產出產物中存在規格違規（如：數據不一致、格式錯誤、圖表解析度不足等）。

### 1.3 分層專屬驗證指令 (Tier-Specific Checks)
```bash
# 僅驗證 Tier 1 產物存在性與檔案容量
python scripts/verify_goteborg_suite.py --tier 1

# 僅驗證 Tier 2 文件格式與內部結構完整性 (DOCX 表格/單元格, PPTX 16:9 頁數, MD 章節標題)
python scripts/verify_goteborg_suite.py --tier 2

# 僅驗證 Tier 3 跨交付物核心數據一致性 (坪數、開價、三段出價、歷史行情、競品)
python scripts/verify_goteborg_suite.py --tier 3

# 僅驗證 Tier 4 圖表資產存在性、PIL 影像鑑識與 GitHub 同步腳本就緒度
python scripts/verify_goteborg_suite.py --tier 4
```

### 1.4 內部單元自我檢驗模式 (Self-Test Suite)
```bash
python scripts/verify_goteborg_suite.py --test-self
```
- **執行行為**：在暫存環境中自動生成合成假資料、Mock DOCX/PPTX/MD/JSON 與模擬走勢圖，測試驗證器本體的 8 項獨立抓漏單元測試，防禦虛假測試與作弊行為。
- **涵蓋用例**：
  - `TC1`：正常 Mock 交付物滿分通過（Tier 1～4 100% PASS）。
  - `TC2`：缺少交付物（如缺少 PPTX）精準攔截。
  - `TC3`：DOCX 表格數量不足（< 5 表格）精準攔截。
  - `TC4`：PPTX 比例非 16:9 寬螢幕或頁數不足精準攔截。
  - `TC5`：Markdown 報告缺少必要章節標題精準攔截。
  - `TC6`：跨交付物出價數字或行情數據不一致精準攔截。
  - `TC7`：缺少走勢圖 PNG 或解析度過低精準攔截。
  - `TC8`：GitHub 同步腳本缺少 `歌德堡` 目錄精準攔截。

### 1.5 JSON 報表輸出模式
```bash
python scripts/verify_goteborg_suite.py --json
python scripts/verify_goteborg_suite.py --progressive --json
```

---

## 2. 四層級驗證覆蓋矩陣 (4-Tier Verification Matrix)

依據 `ORIGINAL_REQUEST.md`、`PROJECT.md`、`TEST_INFRA.md` 與調查成果，本測試套件設定了嚴格的四層級驗收門檻：

| Tier | 驗證層級 | 驗證項目 | 權威依據 / Oracle | 驗證斷言與成功標準 (Assertions) |
|---|---|---|---|---|
| **Tier 1** | **產物存在性與檔案容量** | **Markdown 分析報告** | `PROJECT.md` § Deliverables | `歌德堡/歌德堡_特定待售戶別出價分析與社區整體行情報告.md` 存在且容量 $\ge 15 \text{ KB}$。 |
| **Tier 1** | **產物存在性與檔案容量** | **Word 分析報告 (.docx)** | `PROJECT.md` § Deliverables | `歌德堡/歌德堡_特定待售戶別出價分析與社區整體行情報告.docx` 存在且容量 $\ge 30 \text{ KB}$。 |
| **Tier 1** | **產物存在性與檔案容量** | **PowerPoint 簡報 (.pptx)** | `PROJECT.md` § Deliverables | `歌德堡/歌德堡_特定待售戶別出價分析與社區整體行情報告.pptx` 存在且容量 $\ge 300 \text{ KB}$。 |
| **Tier 1** | **產物存在性與檔案容量** | **估價主資料庫 (.json)** | `PROJECT.md` § Engine Layer | `scratch/goteborg_valuation_master.json` 存在且容量 $\ge 5 \text{ KB}$。 |
| **Tier 2** | **格式與結構完整性** | **Word (.docx) 結構與表格** | `ORIGINAL_REQUEST` R3 | 可由 `docx.Document` 無損解析，表格數 $\ge 5$ 個，總單元格 $\ge 50$ 格，段落文字完整且字數充足。 |
| **Tier 2** | **格式與結構完整性** | **PowerPoint (.pptx) 結構** | `ORIGINAL_REQUEST` R3 | 可由 `pptx.Presentation` 無損解析，總頁數 $\ge 10$ 頁，投影片寬高嚴格符合 **16:9 寬螢幕** (13.333" $\times$ 7.5")。 |
| **Tier 2** | **格式與結構完整性** | **Markdown 核心章節架構** | `PROJECT.md` § Contract | 完整涵蓋 6 大核心章節：**物件現況**、**社區歷程**、**競品比價**、**車位拆算**、**出價三道防線**、**談判話術**。 |
| **Tier 3** | **跨交付物數據一致性** | **目標物件產調規格** | `explorer_survey_goteborg_1` | 樓層 **8F (次頂樓)**、總權狀 **29.98 坪** (主建 20.03, 附屬 2.04, 公設 7.91, 公設比 26.38%)、**無汽車車位 (機車位 B1-59)**、開價 **988 萬 (32.96 萬/坪)** 在 MD/DOCX/PPTX/JSON 100% 一致。 |
| **Tier 3** | **跨交付物數據一致性** | **三段式出價策略精算** | `explorer_survey_goteborg_3` | **保守底標價 810 萬 (27.0 萬/坪)**、**合理目標價 870~885 萬 (29.0~29.5 萬/坪)**、**上限天花板價 930 萬 (31.0 萬/坪)** 在所有產物中完全一致。 |
| **Tier 3** | **跨交付物數據一致性** | **社區實價登錄歷史行情** | `explorer_survey_goteborg_2` | 純住宅有效交易 **83 筆** (排除 11 筆非典型/店面)、2024 年頂峰均價 **31.13 萬/坪**、2026 年最新成交均價 **26.65 萬/坪** (回檔 **-14.4%**) 嚴格吻合。 |
| **Tier 3** | **跨交付物數據一致性** | **周邊同型核心競品比價** | `PROJECT.md` § Feature 4 | **公園苑**、**椰城**、**長億南園**、**廣三大時代** 在各分析文檔與資料庫中完整對比且數據一致。 |
| **Tier 4** | **圖表資產與 Git 就緒度** | **走勢圖檔案與解析度** | `PROJECT.md` § Code Layout | `南區25年以上大型社區走勢圖/29_歌德堡_成交行情走勢圖.png` 存在，PIL 影像鑑識無損且解析度 $\ge 1000 \times 600$（實際 2925x2035 超高解析度）。周邊競品總圖均存在。 |
| **Tier 4** | **圖表資產與 Git 就緒度** | **GitHub 自動同步腳本配置** | `PROJECT.md` § Layer 5 / `GEMINI.md` | `scripts/sync_to_github.py` 目錄清單必須包含 `"歌德堡"`，確保成果能自動推播至 `main` 與 `gh-pages`。 |
| **Tier 4** | **圖表資產與 Git 就緒度** | **文件內嵌圖表元件檢驗** | `TEST_INFRA.md` § Coverage | DOCX 具備圖片媒體關聯，PPTX 包含獨立圖片形狀元件，確認無破圖或漏嵌。 |

---

## 3. 當前基線測試結果 (Baseline Execution Results)

### 3.1 內部單元自我檢驗 (`--test-self`)
執行 `python scripts/verify_goteborg_suite.py --test-self`：
```
================================================================================
【執行 verify_goteborg_suite.py 內部單元自我測試 (Self-Test Suite)】
================================================================================
  [PASS] TC1: Valid mock deliverables passed Tier 1 ~ 4 full acceptance 100%.
  [PASS] TC2: Missing PPTX deliverable correctly caught and failed by Tier 1.
  [PASS] TC3: Insufficient DOCX table count correctly caught and failed by Tier 2.
  [PASS] TC4: Non-16:9 PPTX aspect ratio correctly caught and failed by Tier 2.
  [PASS] TC5: Missing required MD section headers correctly caught and failed by Tier 2.
  [PASS] TC6: Valuation pricing tier mismatch correctly caught and failed by Tier 3.
  [PASS] TC7: Missing primary chart image correctly caught and failed by Tier 4.
  [PASS] TC8: Sync script missing '歌德堡' correctly caught and failed by Tier 4.
================================================================================
🎉 [SELF-TEST SUCCESS] All 8/8 internal verification tests passed successfully!
================================================================================
```
- **結論**：自我測試 8/8 全數通過，驗證套件具備強大抓漏與鑑識敏感度。

### 3.2 階段演進模式基線檢驗 (`--progressive`)
執行 `python scripts/verify_goteborg_suite.py --progressive`：
```
================================================================================
【台中市南區歌德堡特定待售戶出價分析與社區行情 — 自動化驗證套件】
Project Root : D:\GoogleDrive\01_個人_家庭與生活\02_家庭_安莉的資料\小寶的買房子資料
Target Dir   : D:\GoogleDrive\01_個人_家庭與生活\02_家庭_安莉的資料\小寶的買房子資料\歌德堡
Execute Mode : PROGRESSIVE (階段演進模式)
Active Tiers : [1, 2, 3, 4]
================================================================================

✅ PASS Tier 1: File Existence & Basic Checks (產物存在性與檔案容量檢驗)
  [WARN] [Markdown 分析報告] 尚未產生：`歌德堡\歌德堡_特定待售戶別出價分析與社區整體行情報告.md` (Progressive Mode 略過阻斷)
  [WARN] [Word 分析報告 (.docx)] 尚未產生：`歌德堡\歌德堡_特定待售戶別出價分析與社區整體行情報告.docx` (Progressive Mode 略過阻斷)
  [WARN] [PowerPoint 簡報 (.pptx)] 尚未產生：`歌德堡\歌德堡_特定待售戶別出價分析與社區整體行情報告.pptx` (Progressive Mode 略過阻斷)

✅ PASS Tier 2: Format & Structural Integrity (檔案格式與內部結構完整性)
  [WARN] Word 文件尚未產生，略過 Tier 2 DOCX 檢驗
  [WARN] PowerPoint 簡報尚未產生，略過 Tier 2 PPTX 檢驗
  [WARN] Markdown 分析報告尚未產生，略過 Tier 2 MD 檢驗

✅ PASS Tier 3: Data Consistency Across Artifacts (跨交付物核心數據一致性)
  [WARN] 所有文檔交付物 (MD/DOCX/PPTX) 均未產生，略過文檔間一致性驗證

✅ PASS Tier 4: Visual Assets & Git Sync Readiness (圖表資產與 GitHub 同步就緒度)
  [WARN] GitHub 自動同步腳本 (`scripts/sync_to_github.py`) 尚未將 `歌德堡` 加入同步目錄清單 (Pending M4)

================================================================================
🎉【驗收結論：PROGRESSIVE PASS】現有已產出產物通過檢驗！(尚有待完成交付物由工作者產製中)
================================================================================
```
- **Exit Code**：`0`。
- **結論**：現有數據資產與走勢圖完全通過驗證，守門員已正確就位。

### 3.3 全量嚴格驗收基線檢驗（無參數執行）
執行 `python scripts/verify_goteborg_suite.py`：
- **Exit Code**：`1`（預期防禦行為）。
- **結論**：在工作者產製出合格產物前，嚴格攔截未完工狀態，防止虛假交付。

---

## 4. 下游里程碑 (M1 ~ M4) 驗收與交付準則 (Milestone Exit Criteria)

### 4.1 Milestone M1（估價主資料庫產製）
- **產出物**：`scratch/goteborg_valuation_master.json`
- **驗收指令**：
  ```bash
  python scripts/verify_goteborg_suite.py --tier 1 --progressive
  python scripts/verify_goteborg_suite.py --tier 3 --progressive
  ```
- **通過門檻**：JSON 檔案 $\ge 5 \text{ KB}$，包含 8F、29.98坪、988萬、810/870-885/930萬、83筆純住、公園苑/椰城/長億南園/廣三大時代。

### 4.2 Milestone M2（Markdown 分析報告與 Word 文件產製）
- **產出物**：`歌德堡/歌德堡_特定待售戶別出價分析與社區整體行情報告.md`、`.docx`
- **驗收指令**：
  ```bash
  python scripts/verify_goteborg_suite.py --tier 1 --progressive
  python scripts/verify_goteborg_suite.py --tier 2 --progressive
  python scripts/verify_goteborg_suite.py --tier 3 --progressive
  ```
- **通過門檻**：MD $\ge 15 \text{ KB}$ 且含 6 大核心章節；DOCX $\ge 30 \text{ KB}$ 且表格數 $\ge 5$、單元格 $\ge 50$、內嵌圖表；數據 100% 一致。

### 4.3 Milestone M3（PowerPoint 簡報檔產製）
- **產出物**：`歌德堡/歌德堡_特定待售戶別出價分析與社區整體行情報告.pptx`
- **驗收指令**：
  ```bash
  python scripts/verify_goteborg_suite.py --tier 1 --progressive
  python scripts/verify_goteborg_suite.py --tier 2 --progressive
  python scripts/verify_goteborg_suite.py --tier 3 --progressive
  ```
- **通過門檻**：PPTX $\ge 300 \text{ KB}$，頁數 $\ge 10$ 頁，16:9 寬螢幕規格，內嵌圖表，數據 100% 一致。

### 4.4 Milestone M4（全套件 E2E 最終交付與 GitHub 雙分支同步）
- **執行指令**：
  ```bash
  # 1. 執行嚴格全量驗證 (Exit Code 必須為 0)
  python scripts/verify_goteborg_suite.py

  # 2. 執行 GitHub 雙分支同步 (依 GEMINI.md 核心規範)
  python scripts/sync_to_github.py "feat(goteborg): 交付歌德堡特定待售戶出價分析與全社區行情評估套件"
  ```
- **最終驗收門檻**：
  1. `python scripts/verify_goteborg_suite.py` 輸出：
     `🎉【驗收結論：FULL SUITE PASS】全 4 個驗證層級 (Tier 1 ~ 4) 100% 驗收通過！符合專案最高交付標準。`
  2. Exit Code 為 `0`。
  3. GitHub `main` 與 `gh-pages` 分支完成同步推播。

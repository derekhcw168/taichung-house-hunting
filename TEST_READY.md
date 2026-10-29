# TEST_READY.md — E2E 測試套件就緒與驗證規範手冊

**專案名稱**：台中市南區老牌大型社區實價登錄趨勢與後續三年預估分析簡報全面修訂  
**發布里程碑**：Milestone M1 (E2E Test Suite Design and Delivery)  
**測試設計者**：test_writer_m1 (Specialist / QA)  
**主要測試腳本**：`scratch/test_presentation_revision_e2e.py`  
**最後更新時間**：2026-10-29T06:05:00Z  

---

## 1. 執行方式與指令 (Test Execution Commands)

本測試套件支援獨立 CLI 豐富格式報表執行與 Python 標準庫 `unittest` 執行：

### 1.1 標準 E2E 驗證執行（CI/CD 與品質驗收）
```bash
python scratch/test_presentation_revision_e2e.py
```
- **執行行為**：對預設目標簡報執行 Tier 1～4 全部檢驗。
- **Exit Code**：全部通過時回傳 `0`；發現任何缺陷時回傳 `1` 並列印詳細診斷資訊。

### 1.2 標準庫 Unittest 執行
```bash
python -m unittest scratch/test_presentation_revision_e2e.py
```
- **執行行為**：執行 11 個獨立測試用例（`PresentationRevisionE2EUnitTest`），相容於標準測試架構。

### 1.3 基線診斷與進度觀察模式（Pre-M2 模式）
```bash
python scratch/test_presentation_revision_e2e.py --allow-pre-m2
```
- **執行行為**：輸出完整的 4 層級檢驗報表，但即使尚有未完成功能亦回傳 `0`，適合在 M2 開發期間比對修復進度。

### 1.4 指定特定簡報路徑測試
```bash
python scratch/test_presentation_revision_e2e.py --target "南區25年以上大型社區走勢圖/台中市南區老牌大型社區實價登錄趨勢與後續三年預估分析.pptx"
```

---

## 2. 四層級測試覆蓋矩陣 (4-Tier Test Coverage Matrix)

| Tier | 測試類別 | 測試項目 | 權威資料源 / Oracle | 驗證斷言與成功標準 |
|---|---|---|---|---|
| **Tier 1** | **Feature Coverage** (核心功能覆蓋) | **簡報目標發布路徑與雙檔同步** | 專案目錄與 `南區25年以上大型社區走勢圖/` | 驗證根目錄與圖表目錄下的 PPTX 檔案皆存在，且 SHA-256 二進位雜湊值 bit-for-bit 100% 完全一致。 |
| **Tier 1** | **Feature Coverage** | **投影片總頁數驗收** | `PROJECT.md` / `ORIGINAL_REQUEST.md` R1 | 簡報投影片總數精確為 **65 頁** (`len(prs.slides) == 65`)。 |
| **Tier 1** | **Feature Coverage** | **54 社區全數獨立投影片與圖表嵌入** | `scratch/all_54_communities_transactions.json` | 驗證 Slide 11～64（共 54 頁）中，每一頁皆包含大於 5KB 之真實圖表圖片（`MSO_SHAPE_TYPE.PICTURE`），且 54 個社區名稱 100% 完整覆蓋、無一遺漏、無任何重複。 |
| **Tier 2** | **Boundary & Corner Cases** (邊界與極端狀況) | **大墩人文 115 零成交量縮急凍處理** | `all_54_communities_transactions.json` | 驗證 Slide 10 表格與註解中包含「大墩人文」，並明確加註 115 年無成交/量縮說明；驗證其最新成交為 114 年 29.9 萬，跌幅精確反映為約 **1.3%**。 |
| **Tier 2** | **Boundary & Corner Cases** | **1F 店面過濾與高樓層純住宅保留** | `all_54_communities_transactions.json` | 驗證底層資料 5,935 筆純住宅交易中無任何 1 樓店面（`is_first_floor_shop == False`），且 494 筆 11 樓以上純住宅（`reclaimed_high_floors_count == 494`）未被店面正規表達式誤殺。 |
| **Tier 2** | **Boundary & Corner Cases** | **版面寬高比與形狀幾何完整性** | PowerPoint 16:9 規範 | 驗證投影片尺寸為 16:9 寬螢幕（$13.333 \times 7.5$ 英吋），且所有文字框、卡片、表格與圖片之 `width > 0`、`height > 0`、坐標無 NaN 或極端越界。 |
| **Tier 3** | **Cross-Feature Combinations** (跨功能與數據一致性) | **指標對照總表 11 欄位結構** | `ORIGINAL_REQUEST.md` R6 | 驗證 Slide 10 表格精確包含 **11 個欄位**，且表頭明確包含新增之「歷史最高點到115最新成交跌幅(%)」與「歷史最高點到115均價跌幅(%)」。 |
| **Tier 3** | **Cross-Feature Combinations** | **十大指標社區名單重組** | `ORIGINAL_REQUEST.md` R3 | 驗證表格 10 檔指標社區中強制收錄「大墩人文」與「工學天下」，其餘 8 檔由全區顯著跌幅社區構成。 |
| **Tier 3** | **Cross-Feature Combinations** | **關鍵跌幅數據與原始交易完全吻合** | `all_54_communities_transactions.json` | 遍歷表格所有指標建案，由底層交易原始歷史最高價、115最新成交價與115年均價動態計算跌幅，斷言表格顯示之百分比數值誤差在 $\pm 0.5\%$ 容許區間內（排除大墩人文之均價為無成交註記）。 |
| **Tier 4** | **Acceptance Criteria & Narrative** (驗收標準與論述精準度) | **禁止用詞全文負向掃描** | `ORIGINAL_REQUEST.md` R4.1, R4.2 | 全簡報 65 頁全文掃描：嚴格禁止出現「室內翻修水電工料成本」；嚴格禁止出現「150～200 萬的裝潢翻修」或「150～200 萬」翻修預算扣抵。 |
| **Tier 4** | **Acceptance Criteria & Narrative** | **必要概念全文正向掃描** | `ORIGINAL_REQUEST.md` R4.1, R4.2, R4.3 | 全簡報 65 頁全文掃描：必須包含「結構機能衰退與公共維護負擔」；必須包含「數十至上百萬」；必須包含「50 年」或「50年」老屋無實質買盤論點。 |
| **Tier 4** | **Acceptance Criteria & Narrative** | **價格預測模型獨立說明頁** | `ORIGINAL_REQUEST.md` R5 | 驗證 Slide 9 獨立整頁專門說明計量模型推估原理，明確呈現「月度時間序列」（36 個月），並詳述 Holt 阻尼平滑、經驗貝式收縮、IQR 支撐帶、動態擴展信賴錐 4 大核心機制。 |

---

## 3. 當前基線測試結果 (M1 Baseline Execution Results)

在 Milestone M1 階段，針對尚未進行 M2 修訂之既有簡報檔案（21 頁版本）執行 `python scratch/test_presentation_revision_e2e.py`，檢驗出 **14 項待修復規格差距**（證實測試套件具備強大且真實之防禦抓漏能力，非虛設測試）：

```
========================================================================
VERIFICATION RESULTS SUMMARY
========================================================================
  [ FAIL ] Tier 1 (Feature Coverage) (2 defects detected)
           • Slide count mismatch: expected exactly 65 slides, got 21
           • Cannot verify 54 community slides: total slide count is 21 (< 65)
  [ FAIL ] Tier 2 (Boundary & Corner Cases) (1 defects detected)
           • Slide 10 Top 10 Benchmark table or notes do not contain '大墩人文'.
  [ FAIL ] Tier 3 (Cross-Feature Combinations) (4 defects detected)
           • Slide 10 Top 10 table column count mismatch: expected 11 columns, got 9
           • Slide 10 table missing '歷史最高點到115最新成交跌幅(%)' column.
           • Slide 10 table missing '歷史最高點到115均價跌幅(%)' column.
           • Slide 10 Top 10 Benchmark table MUST include '大墩人文', but it is missing.
  [ FAIL ] Tier 4 (Acceptance Criteria & Narrative Integrity) (7 defects detected)
           • Slide 6 contains PROHIBITED term '150～200 萬的裝潢翻修'.
           • Slide 6 contains PROHIBITED term '150～200 萬'.
           • Slide 21 contains PROHIBITED term '150～200 萬'.
           • Presentation full text missing MANDATORY concept '結構機能衰退與公共維護負擔'.
           • Presentation full text missing MANDATORY concept '數十至上百萬'.
           • Presentation full text missing MANDATORY concept '50 年老屋買盤'.
           • Slide 9 missing explanation of monthly time series granularity ('月度' / '月份' / '36 個月').
========================================================================
  OVERALL STATUS: [ FAILED ] 14 DEFECTS DETECTED ACROSS TIERS
========================================================================
```

與此同時，底層資料管道之 3 項關鍵數據斷言（54 社區、7,072 筆交易、5,935 筆純住宅、494 筆高樓層找回、16:9 寬螢幕幾何版面）**100% 通過**，證明測試環境與計算 Oracle 基礎穩固可靠。

---

## 4. Milestone M2 / M3 驗收成功判定標準 (Exit Criteria)

當代碼生成代理人（Milestone M2）完成 `scripts/generate_south_district_trends_presentation.py` 之修訂並產出新版 PPTX 後，執行以下指令：

```bash
python scratch/test_presentation_revision_e2e.py
```

**驗收通過標準**：
1. 終端輸出顯示 `OVERALL STATUS: [ SUCCESS ] ALL 4 TIERS PASSED 100%!`。
2. Exit Code 為 `0`。
3. 11 項測試用例全部為 `PASS`，0 項缺陷。

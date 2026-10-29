# TEST_READY.md — E2E 測試套件就緒與驗證規範手冊

**專案名稱**：大墩人文屋主議價與殺價理由 50 條全攻略與實戰談判速查表  
**發布里程碑**：Milestone M1 (獨立自動化驗證套件發布與 TEST_READY)  
**測試設計者**：teamwork_test_writer (Specialist / QA)  
**主要測試腳本**：`scripts/verify_bargaining_points.py`  
**最後更新時間**：2026-10-29T15:45:00Z  

---

## 1. 執行方式與指令 (Test Execution Commands)

本驗證套件具備強大、獨立、無副作用之 CLI 診斷能力，支援標準驗收執行、自定義目標檔案路徑、詳細除錯輸出、JSON 機器可讀報表，以及內建單元自我檢驗（Self-Test Suite）：

### 1.1 標準 E2E 驗收執行（預設目標檔案）
```bash
python scripts/verify_bargaining_points.py
```
- **目標檔案**：`大墩人文/大墩人文屋主議價與殺價理由50條全攻略.md`
- **執行行為**：對目標攻略進行 Tier 1～Tier 4 全量檢驗。
- **Exit Code**：
  - `0`：全部 Tier 1～4 斷言 100% 通過。
  - `1`：目標檔案缺失或發現任何規格缺陷，並印出清晰定位訊息。

### 1.2 內部單元自我檢驗模式（Self-Test Suite）
```bash
python scripts/verify_bargaining_points.py --test-self
```
- **執行行為**：在記憶體中動態生成多組包含正常與邊界瑕疵之模擬資料，驗證 Oracle 斷言本身是否具備真實抓漏能力（防禦虛假測試與作弊）。
- **涵蓋用例**：
  - TC1：正常 52 條全攻略滿分通過。
  - TC2：缺少必要欄位（如編號識別）精準攔截。
  - TC3：實戰話術字數不足或缺少中文對白引號 `「...」` 精準攔截。
  - TC4：外部專家引用未達 10 位獨立專家精準攔截。
  - TC5：速查表 (Cheatsheet) 與卡片編號脫鉤精準攔截。
  - TC6：條數未達 50 條最低門檻精準攔截。
  - TC7：目標檔案不存在時 Tier 1 乾淨回傳失敗。

### 1.3 指定自定義檔案路徑驗證
```bash
python scripts/verify_bargaining_points.py --target "大墩人文/大墩人文屋主議價與殺價理由50條全攻略.md"
```

### 1.4 詳細除錯與 JSON 輸出模式
```bash
# 印出所有引用的權威房產專家與 YouTuber 名單
python scripts/verify_bargaining_points.py --verbose

# 輸出 JSON 格式報告供 CI/CD 或下游審查代理人使用
python scripts/verify_bargaining_points.py --json
```

---

## 2. 四層級驗證覆蓋矩陣 (4-Tier Verification Matrix)

依據 `explorer_survey_bargain_3/analysis.md` 與 `PROJECT.md` 介面合約，本測試套件設定了嚴格的四層級驗收標準：

| Tier | 驗證類別 | 驗證項目 | 權威資料源 / Oracle | 驗證斷言與成功標準 (Assertions) |
|---|---|---|---|---|
| **Tier 1** | **架構與編碼完整性** (Architecture & Encoding) | **目標檔案存在性與檔案大小** | 專案目錄結構規範 | 目標檔案 `大墩人文/大墩人文屋主議價與殺價理由50條全攻略.md` 必須存在，且檔案大小 $\ge 10 \text{ KB}$。 |
| **Tier 1** | **架構與編碼完整性** | **UTF-8 無損編碼檢驗** | 跨平台字元集規範 | 檔案必須為標準 UTF-8 編碼，無 BOM 損毀，無解碼例外。 |
| **Tier 1** | **架構與編碼完整性** | **核心架構章節完備性** | `PROJECT.md` § Architecture | 必須完整包含：大標題、目錄、速查專區、實戰話術、五大談判階段、簽約特約底線、附錄佐證索引。 |
| **Tier 2** | **條數與欄位完整性** (Quantity & Fields) | **殺價理由總條數** | `ORIGINAL_REQUEST.md` R3 | 解析出之有效殺價卡片總數 $\ge 50$ 條（預設規劃為 52 條，BP-01～BP-52）。 |
| **Tier 2** | **條數與欄位完整性** | **編號連續性與唯一性** | `PROJECT.md` § Contracts | 編號必須符合 `BP-\d{2,}` 格式，且全篇**絕對零重複、零缺號**。 |
| **Tier 2** | **條數與欄位完整性** | **六大必要欄位完整性** | `explorer_survey_bargain_3` § 2.1 | 每張卡片必須包含 6 項必要欄位：<br>1. **編號識別** (`BP-XX`)<br>2. **分類標籤** (必須為六大標準分類之一)<br>3. **核心論點** (去除空白字數 $\ge 30$ 字)<br>4. **適用情境** (去除空白字數 $\ge 10$ 字)<br>5. **實戰話術** (去除空白字數 $\ge 50$ 字，且必須包含對話引號 `「` 與 `」`)<br>6. **佐證出處** (去除空白字數 $\ge 15$ 字) |
| **Tier 2** | **條數與欄位完整性** | **六大標準分類合規性** | `PROJECT.md` 分類定義 | 分類標籤必須嚴格符合以下六類之一：<br>• 屋況瑕疵與基礎翻修成本折抵<br>• 社區公設老化與公共管理提撥<br>• 實價登錄行情與銀行估價落差<br>• 地段環境與周邊抗性<br>• 央行信用管制與房貸政策風險<br>• 賣方心理攻防與談判讓步戰術 |
| **Tier 3** | **雙軌佐證來源覆蓋率** (Source Duality) | **大墩人文在地檔案引用數** | `大墩人文/` 資料夾既有檔案 | 引用專案在地真實檔案（信義說明書/盛祥統包報價單/樂居實價登錄/說帖等）之理由數 $\ge 15$ 條。 |
| **Tier 3** | **雙軌佐證來源覆蓋率** | **外部專家與 YouTuber 引用數** | 全網房產知識庫 | 引用外部房產專家、書籍專欄或影音頻道之理由數 $\ge 25$ 條。 |
| **Tier 3** | **雙軌佐證來源覆蓋率** | **獨立權威專家多樣性** | `ORIGINAL_REQUEST.md` R1 | 全文所引用之獨立知名專家/頻道數 $\ge 10$ 位（如：35線上賞屋、地產秘密客、喬王、房產阿北、國際超級房仲、邱愛莉、李奕農、徐佳馨、安娜、好房網、591、理財周刊、PTT等）。 |
| **Tier 4** | **現場速查表對齊驗證** (Cheatsheet Alignment) | **Markdown 表格存在性與欄位結構** | `explorer_survey_bargain_3` § 3.1 | 必須包含 5 欄式標準 Cheatsheet 表格：`序號`、`分類`、`談判切入痛點`、`關鍵客觀數據 / 依據`、`現場 5 秒殺價台詞 (破冰金句)`。 |
| **Tier 4** | **現場速查表對齊驗證** | **速查表格列數門檻** | `ORIGINAL_REQUEST.md` R3 | Cheatsheet 有效資料列數 $\ge 50$ 筆。 |
| **Tier 4** | **現場速查表對齊驗證** | **主清單與速查表 100% 雙向對齊** | 防禦性一致性原則 | 速查表中的所有序號集合與詳細攻略卡片的編號集合 **100% 完全吻合**，無遺漏任何項目，亦無多餘編號。 |
| **Tier 4** | **現場速查表對齊驗證** | **速查表儲存格資料密度** | 現場實戰可用性標準 | 每列之 5 個欄位皆不得為空，台詞欄位必須具備精練具殺傷力之直接口語。 |

---

## 3. 當前基線測試結果 (M1 Baseline Execution Results)

### 3.1 測試套件自我檢驗 (`--test-self`)
執行 `python scripts/verify_bargaining_points.py --test-self`：
```
================================================================================
【執行 verify_bargaining_points.py 內部單元自我測試 (Self-Test Suite)】
================================================================================
  [PASS] TC1: Valid 52-point guide should pass 100%
  [PASS] TC2: Missing mandatory field (編號識別) must fail Tier 2
  [PASS] TC3: Short dialogue (< 50 chars) must fail Tier 2
  [PASS] TC4: Insufficient unique experts (< 10) must fail Tier 3
  [PASS] TC5: Cheatsheet table ID mismatch must fail Tier 4
  [PASS] TC6: Total points count < 50 must fail Tier 2 & Tier 4
  [PASS] TC7: Non-existent file correctly rejected by Tier 1
================================================================================
🎉 [SELF-TEST SUCCESS] All 7 internal verification tests passed successfully!
```
**自我檢驗結論**：7 項內部單元測試全數 PASS，證實測試套件能精準辨識真偽、有效阻絕作弊與規格漏洞。

### 3.2 針對專案目標檔案之基線執行
執行 `python scripts/verify_bargaining_points.py`：
```
================================================================================
【大墩人文 50+ 條買房殺價理由全攻略 — 獨立自動化驗證套件】
Target File: ...\大墩人文\大墩人文屋主議價與殺價理由50條全攻略.md
================================================================================

❌ [Tier 1 FAIL] Target file not found: ...\大墩人文\大墩人文屋主議價與殺價理由50條全攻略.md
```
- **Exit Code**：`1`
- **狀態評估**：符合雙軌測試驅動 (TDD) 開發常態。目標攻略文件尚未由 Milestone M2 產製，驗證腳本即時正確攔截缺失，測試守門員就位完畢。

---

## 4. 下游里程碑 (M2 / M3) 交付與驗收準則 (Exit Criteria)

當 Milestone M2 撰寫工程師完成 `大墩人文/大墩人文屋主議價與殺價理由50條全攻略.md` 後，必須執行：

```bash
python scripts/verify_bargaining_points.py
```

### 驗收成功判定標準：
1. **Exit Code 必須為 `0`**。
2. 終端輸出包含：
   `🎉【驗收結論：PASS】全 4 個驗證層級 (Tier 1 ~ 4) 100% 驗收通過！符合專案最高交付標準。`
3. 任何 `[FAIL]`、欄位遺漏、台詞缺少雙引號、或 Cheatsheet 編號脫鉤將立即導致驗收失敗並阻絕發布。

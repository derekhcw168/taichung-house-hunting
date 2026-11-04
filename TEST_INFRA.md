# E2E Test Infra: 台中市南區「大東家南園9樓」待售物件出價分析與全社區行情評估套件

## 1. Test Philosophy (測試哲學與方法論)

- **Opaque-Box & Requirement-Driven (黑箱與需求驅動)**：測試套件獨立於生產端生成腳本，完全依據 `ORIGINAL_REQUEST.md`、`PROJECT.md` 與 `DISPATCH.md` 規範之公開交付接口，對產出物實施客觀鑑識與斷言，絕無虛假硬編碼或循環自證。
- **Dual-Track Verification (雙軌驗收機制)**：
  - **嚴格驗收軌道 (Strict Acceptance Track)**：無參數執行時為全量強制嚴格驗收，任何未產出檔案、格式瑕疵或數據不一致立即阻斷（Exit Code 1），作為專案最終發布（Milestone M4）之通行證。
  - **階段演進軌道 (Progressive Evolution Track)**：啟用 `--progressive` 旗標時，未完成之里程碑產物以 `[WARN]` 標記而不阻斷，專注驗證已完成之產物品質，支援漸進式敏捷交付。
- **Adversarial Self-Testing (對抗式自測防禦)**：內建 `--test-self` 合成 Mock 夾具，覆蓋 8 種極端缺陷與異常情境（檔案缺失、DOCX 表格不足、PPTX 非 16:9 比例、MD 核心章節缺失、定價數字偏差、走勢圖缺失、Git 同步清單遺漏），確保測試器自身敏銳可靠。
- **4-Tier Hierarchical Methodology (4 層階梯式驗證方法)**：
  - **Tier 1 (產物存在性與檔案容量)**：快速過濾零位元組或容量不足之無效產物。
  - **Tier 2 (檔案格式與內部結構完整性)**：使用專業函式庫（`python-docx`、`python-pptx`）解析底層 XML 結構，驗證表格數、單元格數、段落數、16:9 寬螢幕比例與 Markdown 必要章節。
  - **Tier 3 (跨交付物數據一致性)**：多源交叉比對待售戶坪數、開價、車位拆算淨單價、三段式出價防線、歷史實登成交筆數與核心競品名稱。
  - **Tier 4 (圖表資產與 GitHub 同步就緒度)**：透過 PIL 進行影像鑑識（無損讀取且解析度 $\ge 1000 \times 600$），檢核 `scripts/sync_to_github.py` 目錄同步清單與文件圖表嵌入狀態。

---

## 2. Feature Inventory Coverage (功能清單與測試層級覆蓋矩陣)

| # | Feature (功能項目) | 需求來源 | Tier 1 | Tier 2 | Tier 3 | Tier 4 |
|---|---|---|:---:|:---:|:---:|:---:|
| 1 | 待售物件產權與規格查核 (33.70坪, 主附21.828坪, 9F) | ORIGINAL_REQUEST R1 | ✓ | ✓ | ✓ | - |
| 2 | 開價行情與毛單價核定 (898萬, 26.65萬/坪) | ORIGINAL_REQUEST R1 | ✓ | ✓ | ✓ | - |
| 3 | 大東家南園歷史實登清洗 (77筆登錄, 63/67筆純住宅, 排除14筆) | ORIGINAL_REQUEST R1 | ✓ | ✓ | ✓ | - |
| 4 | 車位拆算解耦 (機械平移車位72~75萬/2.6~3.7坪, 淨單價26.46萬/坪) | ORIGINAL_REQUEST R2 | ✓ | ✓ | ✓ | - |
| 5 | 南區指標競品社區實價對比 (長億南園, 公園苑, 椰城, 大東家華園, 大東家樸園) | ORIGINAL_REQUEST R1 | ✓ | ✓ | ✓ | - |
| 6 | 多維度比價調整矩陣推導 (屋齡31.8年, 9樓採光視野, 一層兩戶) | ORIGINAL_REQUEST R2 | ✓ | ✓ | ✓ | - |
| 7 | 三段式出價策略 (保守805~810萬, 目標845~860萬, 天花板880~890萬) | ORIGINAL_REQUEST R2 | ✓ | ✓ | ✓ | - |
| 8 | 2026 最新實登定錨 (3F 855萬 25.3萬/坪) 與 2025 歷史峰值 (980萬/928萬) | ORIGINAL_REQUEST R2 | ✓ | ✓ | ✓ | - |
| 9 | 結構化主資料庫格式完整性 (`scratch/dadongjia_valuation_master.json`) | ORIGINAL_REQUEST R1/R2 | ✓ | ✓ | ✓ | - |
| 10| Markdown 深度分析報告完整性 (8 大核心章節) | ORIGINAL_REQUEST R3 | ✓ | ✓ | ✓ | - |
| 11| Word (.docx) 排版、表格與圖表整合 (表格 $\ge 5$, 單元格 $\ge 50$) | ORIGINAL_REQUEST R3 | ✓ | ✓ | ✓ | ✓ |
| 12| PowerPoint (.pptx) 16:9 簡報格式與卡片呈現 (頁數 $\ge 10$, 13.333" $\times$ 7.5") | ORIGINAL_REQUEST R3 | ✓ | ✓ | ✓ | ✓ |
| 13| 走勢圖資產無損鑑識 (`42_大東家南園_成交行情走勢圖.png` $\ge 1000 \times 600$) | ORIGINAL_REQUEST R3 | - | - | - | ✓ |
| 14| Git & GitHub 同步腳本完整性 (`sync_to_github.py` 納入 `大東家南園9樓`) | GEMINI.md / R4 | - | - | - | ✓ |

---

## 3. Test Architecture & 4-Tier Breakdown (測試架構與分層標準)

### Tier 1: 檔案存在性與容量門檻 (File Existence & Sizing)
- `大東家南園9樓/大東家南園9樓_特定待售戶別出價分析與社區整體行情報告.md`：$\ge 15 \text{ KB}$ (15,360 bytes)。
- `大東家南園9樓/大東家南園9樓_特定待售戶別出價分析與社區整體行情報告.docx`：$\ge 30 \text{ KB}$ (30,720 bytes)。
- `大東家南園9樓/大東家南園9樓_特定待售戶別出價分析與社區整體行情報告.pptx`：$\ge 300 \text{ KB}$ (307,200 bytes)。
- `scratch/dadongjia_valuation_master.json`：$\ge 5 \text{ KB}$ (5,120 bytes)。

### Tier 2: 結構完整度與格式排版 (Structural & Formatting Integrity)
- **Markdown (.md)**：必須完整包含 8 大核心章節標題：
  1. 物件產調與規格查核
  2. 歷史實登洗淨分析
  3. 同類型競品比價
  4. 車位拆算解耦
  5. 多維度比價調整矩陣
  6. 三段式出價策略精算
  7. 談判議價話術說帖
  8. 簽約避坑與修繕檢核
- **Word (.docx)**：
  - 成功由 `docx.Document` 解析。
  - 表格數量 $\ge 5$ 個。
  - 表格總單元格數 $\ge 50$ 格。
  - 內文段落數 $\ge 10$ 個且內文長度 $\ge 500$ 字元。
- **PowerPoint (.pptx)**：
  - 成功由 `pptx.Presentation` 解析。
  - 投影片總頁數 $\ge 10$ 頁。
  - 比例嚴格符合 **16:9 寬螢幕**（寬度 13.333 英吋，高度 7.5 英吋，誤差容許 $\le 0.15$ 英吋）。

### Tier 3: 數值一致性檢驗 (Numerical Consistency Across Deliverables)
以下關鍵數值在 JSON、Markdown、Word、PowerPoint 四者之間必須 100% 精確吻合：
- **標的樓層**：9 樓（9F / 9/15F，次高樓層/景觀高樓層）。
- **權狀總面積**：33.70 坪（或 33.71 坪，主建物 19.844 坪、附屬建物 1.984 坪、主附合計 21.828 坪）。
- **開價總額**：898 萬元（毛單價 26.65 萬/坪）。
- **車位型態與拆算**：坡道機械下層平移車位（可停休旅車），估值 72~75 萬元，持分 2.6~3.7 坪，拆算後房屋淨單價約 26.46 萬/坪。
- **三段式出價策略**：
  - **保守底標價**：805 ~ 810 萬元（約 90 折，淨單價 23.5 ~ 24.0 萬/坪；或廣義試探區間 755 ~ 810 萬）。
  - **合理目標價**：845 ~ 860 萬元（約 94 ~ 96 折，淨單價 24.9 ~ 25.5 萬/坪，精準對標 2026-01 最新成交 855 萬）。
  - **上限天花板價**：880 ~ 890 萬元（約 98 折，淨單價 25.9 ~ 26.2 萬/坪，停損絕不追高）。
- **實登與市場基線**：
  - 社區登錄總數 77 筆，清洗後純住宅 63 筆（或 67 筆）。
  - 2026 年最新成交為 3 樓總價 855 萬元（單價 25.3 萬/坪）。
- **競品社區覆蓋**：長億南園、公園苑、椰城、大東家華園、大東家樸園。

### Tier 4: 圖表資產無損鑑識與專案規範就緒 (Asset & Git Sync Readiness)
- **主圖表資產**：`南區25年以上大型社區走勢圖/42_大東家南園_成交行情走勢圖.png` 存在，PIL 讀取無損，解析度 $\ge 1000 \times 600$（實際尺寸 2925 $\times$ 2035）。
- **競品走勢圖**：長億南園、公園苑、椰城、大東家樸園、大東家華園及大盤匯總圖均存在。
- **GitHub 同步腳本**：`scripts/sync_to_github.py` 目錄清單必須包含 `"大東家南園9樓"`。
- **文件內嵌圖表**：DOCX 具備圖片媒體關聯，PPTX 包含獨立圖片形狀元件。

---

## 4. Test Runner CLI Usage (測試執行指令)

```bash
# 1. 內部對抗式自測 (Self-Test Suite, 8/8 鑑識單元測試)
python scripts/verify_dadongjia_suite.py --test-self

# 2. 里程碑階段演進驗收模式 (Progressive Mode, 適用 M1~M3 開發中)
python scripts/verify_dadongjia_suite.py --progressive
python scripts/verify_dadongjia_suite.py --progressive --verbose

# 3. 分層專屬檢驗 (Tier 1 ~ Tier 4)
python scripts/verify_dadongjia_suite.py --tier 1
python scripts/verify_dadongjia_suite.py --tier 2
python scripts/verify_dadongjia_suite.py --tier 3
python scripts/verify_dadongjia_suite.py --tier 4

# 4. 全量嚴格驗收 (Strict Mode, 專案最終發布門檻, Exit Code 0 = 通過)
python scripts/verify_dadongjia_suite.py

# 5. JSON 機器可讀報表產出
python scripts/verify_dadongjia_suite.py --json
python scripts/verify_dadongjia_suite.py --progressive --json
```

# E2E Test Infra: 歌德堡社區待售物件估價與全社區行情分析套件

## Test Philosophy
- Opaque-box, requirement-driven, verifiable.
- Methodology: Category-Partition + Boundary Value Analysis + Office Document Integrity + Data Cross-Verification.

## Feature Inventory
| # | Feature | Source (requirement) | Tier 1 | Tier 2 | Tier 3 | Tier 4 |
|---|---------|---------------------|:------:|:------:|:------:|:------:|
| 1 | 待售物件產權與規格查核 | ORIGINAL_REQUEST R1 | 5 | 5 | ✓ | ✓ |
| 2 | 歌德堡實價登錄歷程與清洗 | ORIGINAL_REQUEST R1 | 5 | 5 | ✓ | ✓ |
| 3 | 車位拆算與淨建坪單價 | ORIGINAL_REQUEST R2 | 5 | 5 | ✓ | ✓ |
| 4 | 南區18棟競品社區實價對比 | ORIGINAL_REQUEST R1 | 5 | 5 | ✓ | ✓ |
| 5 | 比價調整矩陣與邏輯推導 | ORIGINAL_REQUEST R2 | 5 | 5 | ✓ | ✓ |
| 6 | 三段式出價策略 (810/870-885/930) | ORIGINAL_REQUEST R2 | 5 | 5 | ✓ | ✓ |
| 7 | 結構化主資料庫格式 | ORIGINAL_REQUEST R1/R2 | 5 | 5 | ✓ | ✓ |
| 8 | Markdown 深度分析報告完整性 | ORIGINAL_REQUEST R3 | 5 | 5 | ✓ | ✓ |
| 9 | Word (.docx) 排版、表格與圖表整合 | ORIGINAL_REQUEST R3 | 5 | 5 | ✓ | ✓ |
| 10| PowerPoint (.pptx) 16:9 簡報格式與卡片呈現 | ORIGINAL_REQUEST R3 | 5 | 5 | ✓ | ✓ |
| 11| 鑑識審計完整性 (無造假/無硬編碼) | Project Pattern / Audit | 5 | 5 | ✓ | ✓ |
| 12| Git & GitHub 同步腳本完整性 | GEMINI.md / R3 | 5 | 5 | ✓ | ✓ |

## Test Architecture
- Test runner: `python scripts/verify_goteborg_suite.py`
- Pass/Fail Semantics: Exits with 0 if all tests pass, non-zero on failure with detailed assertions.
- Deliverables location: `歌德堡/`

## Coverage Thresholds
- Tier 1: 檔案存在性與非空檢驗（.md, .docx, .pptx, master JSON）
- Tier 2: 檔案格式與結構檢驗（Word/PPTX 可正常由 docx/pptx 函式庫解析，頁數/字數/表格數符合標準）
- Tier 3: 數據一致性檢驗（810萬底標、870-885萬目標、930萬天花板、29.98坪、無車位、18棟競品數據在各文件中完全一致）
- Tier 4: 圖表與視覺資產嵌入驗證（走勢圖 PNG 成功嵌入 Word 與 PPTX，無破圖或缺失）

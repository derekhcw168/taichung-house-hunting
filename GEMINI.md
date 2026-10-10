# 專案作業規範與記憶 (Project Guidelines & Memory)

## 核心規則：自動版本控管與 GitHub 即時同步 (Mandatory Git & GitHub Sync)

1. **每次程式碼或資料異動後自動同步至 GitHub**：
   - 凡是程式碼、HTML 網頁、CSS 樣式、JavaScript 功能、Python 腳本或資料庫有任何新增、修改或刪除，**在每次任務完成回報給使用者之前，必須主動執行 `python scripts/sync_to_github.py "<清晰具體的 Commit 說明>"`**。
   - 必須確保：
     - 本機 Git 倉庫保持最新版本。
     - GitHub `main` 分支（全專案原始碼版本控管）已完成推播。
     - GitHub `gh-pages` 分支（線上公開站台）已即時發布生效。

2. **口語化版本控管支援**：
   - 使用者透過自然口語提出任何功能調整、版面修改或資料修正時，AI 助手必須在完成修改後，自動產生適當且易讀的 Git Commit 訊息並完成提交與推送，讓使用者能單純以口語交談完成專業的版本控管。

3. **保護既有資料完整性**：
   - 嚴格保護 PostgreSQL 資料庫與 HTML 的物件資料，不任意更動既有資料庫記錄，維持資料完整性。

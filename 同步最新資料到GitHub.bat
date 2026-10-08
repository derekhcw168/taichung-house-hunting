@echo off
chcp 65001 >nul
echo 正在將 PostgreSQL 最新資料同步並發布到 GitHub Pages...
python scripts\sync_to_github.py
pause

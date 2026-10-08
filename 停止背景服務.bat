@echo off
chcp 65001 >nul
echo 正在停止看屋決策平台背景服務 (Port 8899)...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8899 ^| findstr LISTENING') do (
    taskkill /F /PID %%a >nul 2>&1
    echo 已關閉 PID: %%a
)
echo 背景服務已停止。
timeout /t 3 >nul

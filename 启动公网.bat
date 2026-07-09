@echo off
chcp 65001 >nul
title 采购到位率查询系统 - 公网模式

echo.
echo ============================================
echo   采购到位率查询系统 v1.0 - 公网模式
echo ============================================
echo.

cd /d "%~dp0"

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python
    pause
    exit /b 1
)

:: 安装依赖
echo [1/3] 检查依赖...
pip install flask pandas openpyxl xlrd -q

:: 启动 Flask
echo [2/3] 启动后端服务...
start "采购到位率-后端" python app.py
timeout /t 3 /nobreak >nul

:: 检查 cloudflared
if not exist "cloudflared.exe" (
    echo [注意] 首次使用需下载 cloudflared，请稍候...
    curl -sL -o cloudflared.exe https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe
)

:: 启动公网隧道
echo [3/3] 启动公网隧道...
echo.
echo   公网地址将显示在下方，按 Ctrl+C 停止服务
echo ============================================
echo.

cloudflared.exe tunnel --url http://localhost:5000

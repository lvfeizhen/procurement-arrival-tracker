@echo off
chcp 65001 >nul
title 采购到位率查询系统

echo.
echo ============================================
echo   采购到位率查询系统 v1.0
echo ============================================
echo.

cd /d "%~dp0"

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 安装依赖（首次运行）
echo [1/2] 检查依赖库...
pip install flask pandas openpyxl xlrd -q

:: 启动服务
echo [2/2] 启动服务...
echo.
start "" http://127.0.0.1:5000
python app.py

pause

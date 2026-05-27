@echo off
chcp 65001 >nul
title 通途订单整理工具 - Windows 打包

echo ========================================
echo    通途订单整理工具 - Windows 打包
echo ========================================
echo.

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 请先安装 Python 3
    echo 下载地址: https://www.python.org/downloads/
    echo 安装时请勾选 "Add Python to PATH"
    pause
    exit /b 1
)

echo [1/3] 安装依赖...
pip install flask openpyxl pyinstaller -q

echo [2/3] 正在打包...
pyinstaller --onefile --name "通途订单整理工具" --hidden-import=openpyxl app.py

echo [3/3] 打包完成！
echo.
echo 可执行文件在 dist\通途订单整理工具.exe
echo 将此文件发送给同事即可使用。
echo.
pause

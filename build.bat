@echo off
chcp 65001 > nul  # 切換UTF-8編碼，避免中文亂碼
echo ==============================
echo Excel數據匹配工具 - 一鍵打包腳本
echo Excel Data Matching Tool - One-click Packaging Script
echo ==============================

:: 檢查Python是否安裝
python --version > nul 2>&1
if errorlevel 1 (
    echo 錯誤：未找到Python，請先安裝Python 3.8~3.10並添加到PATH！
    echo Error: Python not found, please install Python 3.8~3.10 and add to PATH first!
    pause
    exit /b 1
)

:: 進入源碼目錄
cd /d "%~dp0src"

:: 安裝/升級依賴
echo.
echo 正在安裝/升級依賴套件...
echo Installing/upgrading dependencies...
pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

:: 清理舊打包文件
echo.
echo 清理舊打包文件...
echo Cleaning old packaging files...
if exist "../dist" rmdir /s /q "../dist"
if exist "../build" rmdir /s /q "../build"
if exist "excel_matcher.spec" del /f /q "excel_matcher.spec"

:: 打包EXE（無圖標版，兼容所有環境）
echo.
echo 開始打包EXE文件...
echo Starting to package EXE file...
pyinstaller -F -w ^
--hidden-import openpyxl ^
--hidden-import xlrd ^
--hidden-import pandas ^
--hidden-import tkinter ^
--clean ^
--name "Excel數據匹配工具" ^
excel_matcher.py

:: 檢查打包結果
echo.
if exist "../dist/Excel數據匹配工具.exe" (
    echo 打包成功！/ Packaging success!
    echo EXE文件路徑 / EXE file path: ../dist/Excel數據匹配工具.exe
) else (
    echo 打包失敗！/ Packaging failed!
    pause
    exit /b 1
)

:: 清理臨時文件
if exist "excel_matcher.spec" del /f /q "excel_matcher.spec"

echo.
echo ==============================
echo 打包完成！/ Packaging completed!
echo 請前往 ../dist/ 目錄查看EXE文件。
echo Please go to ../dist/ directory to check the EXE file.
echo ==============================
pause
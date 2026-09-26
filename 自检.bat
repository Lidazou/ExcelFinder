@echo off
chcp 65001 >nul 2>&1
setlocal
cd /d "%~dp0"

echo ============================================================
echo   ExcelFinder 自检（命令行，不需要图形界面）
echo ============================================================
echo.
echo 用法：把要检索的目录路径当作参数传进来
echo   例： "自检.bat"  D:\报表
echo.

set "TARGET=%~1"
if "%TARGET%"=="" (
    echo [提示] 未指定目录，改为检查本程序所在目录。
    set "TARGET=%~dp0"
)

echo 目标目录: %TARGET%
echo ------------------------------------------------------------
"%~dp0ExcelFinder-cli.exe" "%TARGET%" --query "报表" --limit 10
echo ------------------------------------------------------------
echo.
pause

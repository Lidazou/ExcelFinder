@echo off
chcp 65001 >nul 2>&1
setlocal
cd /d "%~dp0"

echo ============================================================
echo   ExcelFinder  -  本地 Excel 极速检索（U盘便携版）
echo ============================================================
echo.

if not exist "ExcelFinder.exe" (
    echo [错误] 找不到 ExcelFinder.exe
    echo        请确认本文件与 ExcelFinder.exe 在同一个文件夹内。
    echo.
    pause
    exit /b 1
)

rem Data folder lives next to the EXE so the stick is self-contained and
rem moving it between PCs (or changing drive letters) loses nothing.
if not exist "data" mkdir "data" >nul 2>&1

echo 配置与索引目录: %~dp0data
echo 正在启动...
echo.

rem Windows SmartScreen may warn on first run of an unsigned EXE.
rem Right-click - Properties - Unblock, or "More info" - "Run anyway".
start "" "%~dp0ExcelFinder.exe" %*

exit /b 0

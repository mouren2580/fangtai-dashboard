@echo off
chcp 65001 >nul
REM ============================================================
REM  方太西北服务产品看板 —— 家里电脑 一键同步
REM  用法（双击或用命令行）：
REM     sync.bat "2026年8月西北服务产品.xlsx" 2026-08-16
REM  说明：
REM    - 第1个参数 = 最新月报 Excel 完整路径（建议放在本文件夹内，直接写文件名）
REM    - 第2个参数 = 数据截止日（更新日的前一天，YYYY-MM-DD）
REM    - 令牌：把 GitHub 个人令牌(ghp_xxx)放进同目录 .gh_token 文件即可自动读取
REM    - 想先看效果不发布：在下方命令加 --no-deploy
REM ============================================================
setlocal
set "EXCEL=%~1"
set "DATE=%~2"
if "%EXCEL%"=="" set "EXCEL=2026年8月西北服务产品.xlsx"
if "%DATE%"=="" set "DATE=2026-08-16"

python sync.py --excel "%EXCEL%" --date %DATE% --token-file .gh_token
if errorlevel 1 (
  echo.
  echo [失败] 请检查：1) 是否已安装 Python 并加入 PATH  2) 是否已 pip install -r requirements.txt  3) .gh_token 是否已放好
  pause
  exit /b 1
)
echo.
echo [完成] 看板已刷新并发布。Ctrl+Shift+R 硬刷新页面查看。
pause
endlocal

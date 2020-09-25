echo off
echo "ドラッグアンドドロップでmsgファイルをemlファイルに変換します（複数ファイル可）。"
cd /d "%~dp0"
setlocal enabledelayedexpansion
set x=%*
for %%a in (!x!) do (
    python main.py %%a
)
timeout 5
pause

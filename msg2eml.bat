REM ドラッグアンドドロップでmsgファイルをemlファイルに変換します（複数ファイル可）。
REM 名前に半角空白文字が入っているファイルはリネームしてからご利用ください。
@echo off
cd /d "%~dp0"
setlocal enabledelayedexpansion
set x=%*
for %%a in (!x!) do (
    python main.py %%a
)
timeout 10

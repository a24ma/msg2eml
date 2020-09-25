echo off
echo "ドラッグアンドドロップでmsgファイルをemlファイルに変換します（複数ファイル可）。"
cd /d %~dp0
for %%a in (%*) do python main.py "%%a"
timeout 5

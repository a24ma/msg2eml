echo off
echo "�h���b�O�A���h�h���b�v��msg�t�@�C����eml�t�@�C���ɕϊ����܂��i�����t�@�C���j�B"
cd /d "%~dp0"
setlocal enabledelayedexpansion
set x=%*
for %%a in (!x!) do (
    python main.py %%a
)
timeout 5
pause

REM �h���b�O�A���h�h���b�v��msg�t�@�C����eml�t�@�C���ɕϊ����܂��i�����t�@�C���j�B
REM ���O�ɔ��p�󔒕����������Ă���t�@�C���̓��l�[�����Ă��炲���p���������B
@echo off
cd /d "%~dp0"
setlocal enabledelayedexpansion
set x=%*
for %%a in (!x!) do (
    python main.py %%a
)
timeout 10

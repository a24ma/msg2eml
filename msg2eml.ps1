#!/usr/bin/env pwsh

if (Test-Path venv) {
    .\venv\Scripts\activate.ps1
}
py .\main.py

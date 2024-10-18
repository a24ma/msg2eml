#!/usr/bin/env pwsh

Write-Host "exe の作成: $srcRunPy" -ForegroundColor "Green"
pyinstaller --onefile "main.py" `
    --workpath .\.build.tmp `
    --clean --name msg2eml `
    --collect-all msg2eml -p .
if (!$?) {
    Write-Host 'exe の作成に失敗しました。' -ForegroundColor "Red"
    Start-Sleep 5
    exit 1
}
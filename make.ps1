#!/usr/bin/env pwsh

<# ###################################
# DOCUMENT

* Created on: 2024/10/23
* UPdated on: 2024/10/23
* Requirements:
    + choco install xxxxx

## Introduction

TODO: Do something.

################################### #>
#region    <SETTINGS>

$waitSecOnSccess = 3
$waitSecOnFailed = 30
$logSupLevel = 0 # TRACE
$logSupLevel = 1 # DEBUG
$logSupLevel = 2 # INFO
# $logSupLevel = 3 # WARN
# $logSupLevel = 4 # ERROR
# $logSupLevel = 5 # QUIET

#endregion </SETTINGS>
######################################
#region    <UTILITY>

if ($null -eq $logSupLevel) {
    $logSupLevel = 2 # INFO
}

function __p($lv, $msg, $col) { 
    if ($logSupLevel -le $lv) { 
        Write-Host "$msg" -ForegroundColor "$col" 
    } 
}
function logTrace ($msg) { __p 0 "[TRACE] $msg" "magenta" }
function logDebug ($msg) { __p 1 "[DEBUG] $msg" "magenta" } 
function logInfo ($msg) { __p 2 "[INFO] $msg" "blue" }
function logNote ($msg) { __p 2 "[NOTE] $msg" "green" }
function LogWarn ($msg) { __p 3 "[WARN] $msg" "yellow" }
function logError ($msg) { __p 4 "[ERROR] $msg" "red" }
function waitPrompt ($sec) { 
    if ($sec -le 0 -or $logSupLevel -gt 4) { return }
    __p 4 "Press any key (or wait ${sec}s):" "red"
    timeout $sec >$null
}

trap {
    if ($logSupLevel -le 4) { 
        $delim = "---------------------------------------"
        Write-Host $delim -ForegroundColor "red"
        $_
        Write-Host $delim -ForegroundColor "red"
    }
    waitPrompt $waitSecOnFailed
    exit 1
}

#endregion </UTILITY>
######################################
#region    <MAIN>


logInfo "exe を作成します..." -ForegroundColor "Green"
if (Test-Path("msg2eml.exe")) {
    LogWarn "すでに存在する exe を削除します."
    Remove-Item "msg2eml.exe"
}

pyinstaller --onefile "main.py" `
    --workpath .\.build.tmp `
    --clean --name msg2eml `
    --collect-all msg2eml `
    --additional-hooks-dir=.\mat `
    --icon ".\mat\icon.ico" `
    -p . `
    --windowed `
    --distpath=.

if (!$?) {
    throw 'exe の作成に失敗しました.'
}
logNote "完了しました." -ForegroundColor "Green"

#endregion </MAIN>
######################################

waitPrompt $waitSecOnSccess

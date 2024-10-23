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
    deactivate
    waitPrompt $waitSecOnFailed
    exit 1
}

#endregion </UTILITY>
######################################
#region    <MAIN>

logInfo "venv 環境を作成します..."

if (-not(Test-Path("venv"))) {
    py -m venv venv
}
.\venv\Scripts\activate
py -m pip install --upgrade pip
pip install -r conf/requirements.txt

logInfo "exe を作成します..."

nuitka `
    --clean-cache="all" `
    --include-package="msg2eml" `
    --output-filename="msg2eml.exe" `
    --output-dir=".build" `
    --windows-icon-from-ico=".\mat\icon.ico" `
    --windows-console-mode="disable" `
    main.py

if (!$?) {
    throw 'exe の作成に失敗しました.'
}

logInfo "exe (debug 用) を作成します..."

nuitka `
    --include-package="msg2eml" `
    --output-filename="msg2eml_debug.exe" `
    --output-dir=".build" `
    --windows-icon-from-ico=".\mat\icon.ico" `
    main.py

if (!$?) {
    throw 'exe (debug 用) の作成に失敗しました.'
}

logNote "完了しました." -ForegroundColor "Green"

deactivate

<# ###################################

* onefile は可能だが、ウイルス扱いされる。
* splash/company/file/product 系オプションはコンパイルエラーになる。

# nuitka `
#     --clean-cache="all" `
#     --include-package="msg2eml" `
#     --include-package="tkinterdnd2" `
#     --include-package="win32com.client" `
#     --output-filename="msg2eml.exe" `
#     --output-dir=".build" `
#     --onefile `
#     --enable-plugin="tk-inter" `
#     --windows-icon-from-ico=".\mat\icon.ico" `
#     main.py

    # --company-name="a24ma" `
    # --file-version="1.0.0" `
    # --product-name="msg2eml" `
    # --product-version="1.0" `
    # --file-description="Convert *.msg to .eml files by D&D." `

    # --onefile-windows-splash-screen-image=".\mat\bg.png" `
    # --standalone `
    # --windows-console-mode="disable" `

################################### #>

#endregion </MAIN>
######################################

waitPrompt $waitSecOnSccess

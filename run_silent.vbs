Option Explicit

Dim shell, psScript, cmdLine

' CreateObject は大文字小文字を区別しませんが、区切り文字は半角で
Set shell = CreateObject("WScript.Shell")

' 実行したい .ps1 のフルパスを指定
psScript = ".\msg2eml.ps1"

' PowerShell の実行オプションをまとめて書きます
cmdLine = _
    "powershell.exe" & _
    " -NoProfile" & _
    " -NonInteractive" & _
    " -ExecutionPolicy Bypass" & _
    " -WindowStyle Hidden" & _
    " -File """ & psScript & """"

' 0 = ウィンドウ完全非表示, True = 終了待ち
shell.Run cmdLine, 0, True

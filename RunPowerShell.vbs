'Run PowerShell with VBScript!
'1) Create your own PowerShell file "ps.ps1" to execute in the same folder as this file
'2) Run this file in Windows

Set objShell = CreateObject("Wscript.shell")
objShell.run("powershell.exe -noprofile -executionpolicy bypass -file .\ps.ps1")

Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run chr(34) & "C:\MinGW\capture-k\main.exe" & Chr(34), 0
Set WshShell = Nothing
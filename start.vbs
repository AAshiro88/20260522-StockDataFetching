' Silently start the real-time quote monitoring system.

Option Explicit

Dim fso, shell, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = scriptDir

shell.Run """C:\ProgramData\anaconda3\pythonw.exe"" main.py", 0, False
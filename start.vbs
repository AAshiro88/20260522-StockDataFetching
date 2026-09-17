' Silently start the real-time quote monitoring system.
' Runs pythonw.exe directly with a hidden window and does not wait for it to finish.

Option Explicit

Dim fso, shell, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = scriptDir

' 「"C:\ProgramData\anaconda3\pythonw.exe" main.py」
shell.Run """C:\ProgramData\anaconda3\pythonw.exe"" main.py", 0, False
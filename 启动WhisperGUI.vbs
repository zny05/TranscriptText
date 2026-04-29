Set WshShell = CreateObject("WScript.Shell")
Dim scriptDir
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WshShell.Run """C:\Users\HadesCanyon\Downloads\ServUO\ServUO\.venv\Scripts\pythonw.exe"" """ & scriptDir & "\whisper_gui.py""", 0, False

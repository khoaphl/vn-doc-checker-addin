Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
serverDir = FSO.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = serverDir
WshShell.Run "cmd /c node """ & serverDir & "\server.js""", 0, False

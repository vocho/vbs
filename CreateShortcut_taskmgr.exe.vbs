Option Explicit

Dim objWshShell
Dim objFSO
Dim strShortcutFolder
Dim strShortcutPath
Dim strTargetFolder
Dim strTargetPath

Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

strShortcutFolder = objWshShell.SpecialFolders("Startup")
strShortcutPath = objFSO.BuildPath(strShortcutFolder, "taskmgr.exe.lnk")
strTargetPath = "taskmgr.exe"

With objWshShell.CreateShortcut(strShortcutPath)
	.TargetPath = strTargetPath
    .WindowStyle = 7 ' Minimize
	.Save
End With

Set objFSO = Nothing
Set objWshShell = Nothing

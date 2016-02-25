Option Explicit

Dim objWshShell
Dim objFSO
Dim strDesktopPath
Dim strShortcutPath
Dim strTargetFolder
Dim strTargetPath

Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

strDesktopPath  = objWshShell.SpecialFolders("Desktop")
strShortcutPath = objFSO.BuildPath(strDesktopPath, "Internet Explorer.lnk")
strTargetPath   = objWshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE\")

With objWshShell.CreateShortcut(strShortcutPath)
	.TargetPath = strTargetPath
	.Save
End With

Set objFSO = Nothing
Set objWshShell = Nothing

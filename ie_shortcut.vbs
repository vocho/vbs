Option Explicit

Dim strArg			' 引数情報
Dim objWshShell		' WshShell オブジェクト
Dim strShortcutPath	' デスクトップのフォルダ名
Dim strStartupPath	' デスクトップのフォルダ名
Dim objFSO			' FileSystemObject
Dim strBaseName		' ファイルのベース名(拡張子を除いたもの)
Dim objShortcut		' ショートカット情報
Dim objShell		' Shell オブジェクト

For Each strArg In WScript.Arguments
	Set objWshShell = WScript.CreateObject("WScript.Shell")
	If Err.Number = 0 Then
		strStartupPath = objWshShell.SpecialFolders("Startup")
		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
		strBaseName = objFSO.GetBaseName(strArg)
		strShortcutPath = objFSO.BuildPath(strStartupPath, strBaseName & ".lnk")
		Set objShortcut = objWshShell.CreateShortcut(strShortcutPath)
		objShortcut.TargetPath = strArg
		objShortcut.Save
		
		If Err.Number = 0 Then
			Set objShell = CreateObject("Shell.Application")
			objShell.Open strStartupPath
		Else
			WScript.Echo "エラー: " & Err.Description
		End If
	Else
		WScript.Echo "エラー: " & Err.Description
	End If
	Set objShortcut = Nothing
	Set objWshShell = Nothing
Next

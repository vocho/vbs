Set Shell = CreateObject("Shell.Application")
Set objFolder = Shell.BrowseForFolder(0, "フォルダ選択", 11, 0)
If objFolder Is Nothing Then
	WScript.Quit
End If
If Not objFolder.Self.IsFileSystem Then
	WScript.Echo "ファイルシステムではありません"
	WScript.Quit
End If

Set objDelFolder = Shell.NameSpace("::{645FF040-5081-101B-9F08-00AA002F954E}")

Call objDelFolder.MoveHere( objFolder.Self, 0 )
strPath = objFolder.Self.Path

' 移動元が無くなるまで待つ
Do
	Set obj = Shell.NameSpace( strPath )
	If obj Is Nothing Then
		Exit Do
	End If
	Set obj = Nothing
	WScript.Sleep 500
Loop

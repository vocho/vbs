Set Shell = CreateObject("Shell.Application")
Set objFolder = Shell.BrowseForFolder(0, "�t�H���_�I��", 11, 0)
If objFolder Is Nothing Then
	WScript.Quit
End If
If Not objFolder.Self.IsFileSystem Then
	WScript.Echo "�t�@�C���V�X�e���ł͂���܂���"
	WScript.Quit
End If

Set objDelFolder = Shell.NameSpace("::{645FF040-5081-101B-9F08-00AA002F954E}")

Call objDelFolder.MoveHere( objFolder.Self, 0 )
strPath = objFolder.Self.Path

' �ړ����������Ȃ�܂ő҂�
Do
	Set obj = Shell.NameSpace( strPath )
	If obj Is Nothing Then
		Exit Do
	End If
	Set obj = Nothing
	WScript.Sleep 500
Loop

Option Explicit

Dim strArg			' �������
Dim objWshShell		' WshShell �I�u�W�F�N�g
Dim strShortcutPath	' �f�X�N�g�b�v�̃t�H���_��
Dim strStartupPath	' �f�X�N�g�b�v�̃t�H���_��
Dim objFSO			' FileSystemObject
Dim strBaseName		' �t�@�C���̃x�[�X��(�g���q������������)
Dim objShortcut		' �V���[�g�J�b�g���
Dim objShell		' Shell �I�u�W�F�N�g

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
			WScript.Echo "�G���[: " & Err.Description
		End If
	Else
		WScript.Echo "�G���[: " & Err.Description
	End If
	Set objShortcut = Nothing
	Set objWshShell = Nothing
Next

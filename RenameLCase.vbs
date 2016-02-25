Set objShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each strArg In WScript.Arguments
    If Err.Number = 0 Then
    	If objFSO.FolderExists(strArg) Then
    		Call FolderProc(objFSO.GetFolder(strArg))
    	ElseIf objFSO.FileExists(strArg) Then
    		Call FileProc(objFSO.GetFile(strArg))
        Else
            WScript.Echo "Error: " & Err.Description
    	End If
    Else
        WScript.Echo "Error: " & Err.Description
    End If
Next

Function FolderProc(objFolder)
    For Each objFile In objFolder.Files
        Call FileProc(objFile)
    Next
    For Each objSubFolder In objFolder.SubFolders
        Call FolderProc(objSubFolder)
    Next
    Call FileProc(objFolder)
End Function

Function FileProc(objFileOrFolder)
	Set objParentFolder = objShell.NameSpace(objFileOrFolder.ParentFolder.Path)
	Set FolderItem = objParentFolder.Items.Item(objFileOrFolder.Name)
	FolderItem.Name = LCase(objFileOrFolder.Name)
End Function


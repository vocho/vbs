
Set objFSO = CreateObject("Scripting.FileSystemObject")

Call ArgProc

Set objFSO = Nothing

Function ArgProc
    For Each strArg In WScript.Arguments
        If objFSO.FileExists(strArg) Then
        	Call FileProc(objFSO.GetFile(strArg))
        End If
    Next
End Function

Function FileProc(objFile)
    strBaseName = objFSO.GetBaseName(objFile)
    strDestFolder = objFSO.BuildPath(objFile.ParentFolder.Path, strBaseName)
    If Not objFSO.FolderExists(strDestFolder) Then
        objFSO.CreateFolder(strDestFolder)
        objFSO.MoveFile objFile.Path, strDestFolder & "\"
    End If
End Function



Set objFSO = CreateObject("Scripting.FileSystemObject")

Call ArgProc

Set objFSO = Nothing

Function ArgProc
    For Each strArg In WScript.Arguments
        If objFSO.FolderExists(strArg) Then
            Call FolderProc(objFSO.GetFolder(strArg))
        ElseIf objFSO.FileExists(strArg) Then
        	Call FileProc(objFSO.GetFile(strArg))
        End If
    Next
End Function

Function FolderProc(objFolder)
    For Each objFile In objFolder.Files
        Call FileProc(objFile)
    Next
    For Each objSubFolder In objFolder.SubFolders
        Call FolderProc(objSubFolder)
    Next
End Function

Function FileProc(objFile)
    With CreateObject("WScript.Shell")
        .CurrentDirectory = objFile.ParentFolder.Path
        .Exec("C:\Program Files\7-Zip\7z.exe e """ & objFile.Path & """").StdOut.ReadAll
    End With
End Function


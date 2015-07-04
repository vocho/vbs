
Set objFSO = CreateObject("Scripting.FileSystemObject")

strEXE = "notepad.exe"
'strEXE = objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), objFSO.GetBaseName(WScript.ScriptName))

Call ArgProcess

Set objFSO = Nothing

Function ArgProcess
    For Each strArg In WScript.Arguments
        If objFSO.FileExists(strArg) Then
        	Call FileProcess(objFSO.GetFile(strArg))
        ElseIf objFSO.FolderExists(strArg) Then
            Call FolderProcess(objFSO.GetFolder(strArg))
        End If
    Next
End Function

Function FileProcess(objFile)
    If "txt"=objFSO.GetExtensionName(objFile.Name) Then
        With CreateObject("WScript.Shell")
            .CurrentDirectory = objFile.ParentFolder.Path
            .Exec("""" & strEXE & """ """ & objFile.Path & """").StdOut.ReadAll
        End With
    End If
End Function

Function FolderProcess(objFolder)
    For Each objFile In objFolder.Files
        Call FileProcess(objFile)
    Next
    For Each objSubFolder In objFolder.SubFolders
        Call FolderProcess(objSubFolder)
    Next
End Function

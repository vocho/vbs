
For Each strArg In WScript.Arguments
    If Err.Number = 0 Then
        With CreateObject("Scripting.FileSystemObject")
            If .FolderExists(strArg) Then
                Call FolderProc(.GetFolder(strArg))
            End If
        End With
    Else
        WScript.Echo "ÉGÉâÅ[: " & Err.Description
    End If
Next

Function FolderProc(objFolder)
    Call RenameToUnderbar(objFolder)
    For Each objFile In objFolder.Files
        Call RenameToUnderbar(objFile)
    Next
    For Each objSubFolder In objFolder.SubFolders
        Call FolderProc(objSubFolder)
        Call RenameToUnderbar(objSubFolder)
    Next
End Function

Function RenameToUnderbar(objFileOrFolder)
    strNewName = UnicodeToUnderbar(objFileOrFolder.Name)
    If objFileOrFolder.Name<>strNewName Then
        objFileOrFolder.Name = strNewName
    End If
End Function

Function UnicodeToUnderbar(str)
    strNew = ""
    For i = 1 To Len(str)
        strNew = strNew & Replace(Chr(Asc(Mid(str, i, 1))), "?", "_")
    Next
    UnicodeToUnderbar = strNew
End Function

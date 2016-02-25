Sub RecursiveCreateFolder(ByVal strFolder)
    With CreateObject("Scripting.FileSystemObject")
        If strFolder<>"" Then
            If Not .FolderExists(.GetParentFolderName(strFolder)) Then
                Call RecursiveCreateFolder(.GetParentFolderName(strFolder))
            End If
            If Not .FolderExists(strFolder) Then
                .CreateFolder(strFolder)
            End If
        End If
    End With
End Sub

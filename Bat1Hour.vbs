Do
    With CreateObject("Scripting.FileSystemObject")
        For Each objFile In .GetFile(WScript.ScriptFullName).ParentFolder.Files
            If .GetExtensionName(objFile)="bat" And _
               .FileExists(objFile) And _
               .FileExists(.BuildPath(objFile.ParentFolder.Path, .GetBaseName(objFile))) Then
                If Now() > DateAdd("h", 30, objFile.DateCreated) Then
                    With CreateObject("WScript.Shell")
                        .CurrentDirectory = objFile.ParentFolder.Path
                        .Run """" & objFile.Name & """", 1, True
                    End With
                End If
            End If
        Next
    End With
    WScript.Sleep(60*60*1000)
Loop

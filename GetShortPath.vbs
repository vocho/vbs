
For Each strArg In WScript.Arguments
    With CreateObject("Scripting.FileSystemObject")
        If .FolderExists(strArg) Then
            With .GetFolder(strArg)
                WScript.Echo .ShortPath
            End With
        ElseIf .FileExists(strArg) Then
            With .GetFile(strArg)
                WScript.Echo .ShortPath
            End With
        Else
            WScript.Echo ""
        End If
    End With
Next

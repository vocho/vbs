
For Each strArg In WScript.Arguments
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(strArg) Then
            .CreateTextFile .BuildPath(.GetParentFolderName(strArg), .GetBaseName(strArg) + ".txt")
        ElseIf .FolderExists(strArg) Then
            .CreateTextFile .BuildPath(.GetParentFolderName(strArg), .GetFolder(strArg).Name + ".txt")
        End If
    End With
Next


Const ForReading                = 1
Const ForWriting                = 2
Const ForAppending              = 8

Set objWshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each strArg In WScript.Arguments
    strSrcText = ""
    strDstText = ""

    With objFSO.OpenTextFile(strArg, ForReading, False)
        strSrcText = .ReadAll()
        .Close
    End With

    With New RegExp
        .Pattern = "^(\w)"
        .IgnoreCase = True
        .Global = True
        .MultiLine = True
        strDstText = .Replace(strSrcText, "#$1")
    End With

    If strSrcText<>strDstText Then
        With objFSO.OpenTextFile(strArg, ForWriting, False)
            .Write strDstText
            .Close
        End With
    End If
Next

Set objFSO = Nothing
Set objWshShell = Nothing


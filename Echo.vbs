Option Explicit

Dim strArg
Dim objFSO
Dim objFolder

For Each strArg In WScript.Arguments
    If Err.Number = 0 Then
        WScript.Echo strArg
    Else
        WScript.Echo "ÉGÉâÅ[: " & Err.Description
    End If
Next

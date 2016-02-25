
WScript.Echo "IsAdmin: " & CStr(IsAdmin())

RunAsAdmin()

Function IsAdmin()
    IsAdmin = False
    On Error Resume Next
    Call CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    If Err.Number=0 Then
        IsAdmin = True
    End If
End Function

Sub RunAsAdmin()
    On Error Resume Next
    Call CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    If Err.Number<>0 Then
        Dim strArgs, strArg
        strArgs = Chr(34) & WScript.ScriptFullName & Chr(34)
        For Each strArg In WScript.Arguments
            strArgs = strArgs  & " " & Chr(34) & strArg & Chr(34)
        Next
        Call CreateObject("Shell.Application").ShellExecute(WScript.FullName, strArgs, "", "runas", 1)
    End If
End Sub

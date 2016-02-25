
WScript.Echo "IsAdmin: " & CStr(IsAdmin())

Function IsAdmin()
    IsAdmin = False
    On Error Resume Next
    Call CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    If Err.Number=0 Then
        IsAdmin = True
    End If
End Function

Option Explicit 

Function GetPID()
    With CreateObject("WScript.Shell").Exec("MSHTA.EXE -")
        GetPID = GetObject("winmgmts:Win32_Process.Handle=" & .ProcessID).ParentProcessID
        .Terminate
    End With
End Function

MsgBox GetPID()

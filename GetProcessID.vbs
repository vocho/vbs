Option Explicit 

Function GetPID()
    Dim objExec
    Dim objProcess
    Set objExec = CreateObject("WScript.Shell").Exec("MSHTA.EXE -")
    For Each objProcess In GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_Process WHERE ProcessID=" & objExec.ProcessID)
        GetPID = objProcess.ParentProcessID
    Next
    objExec.Terminate
End Function

MsgBox GetPID()


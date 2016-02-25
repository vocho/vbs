
Call RunCScript

Sub RunCScript
    If LCase(CreateObject("Scripting.FileSystemObject").GetBaseName(WScript.FullName))<>"cscript" Then
        Dim strRun, strArg
        strRun = "cscript """ & WScript.ScriptFullName & """"
        For Each strArg In WScript.Arguments
            strRun = strRun & " """ & strArg & """"
        Next
        CreateObject("WScript.Shell").Run strRun
        WScript.Quit
    End If
End Sub



Set objArray2 = TestJSArray()

WScript.Echo objArray2

objArray2.sort()

For Each strItem In objArray2
    WScript.Echo strItem
Next

Function TestJSArray()
    Set objArray = CreateObjectJS("new Array()")
    For Each strItem In objArray
        WScript.Echo strItem
    Next
    objArray.push "F"
    objArray.push("B")
    objArray.push("C")
    objArray.push("E")
    objArray.push("D")
    objArray.push("A")
    Set TestJSArray = objArray
End Function

Function CreateObjectJS(strCode)
    Dim strJS
    Dim strTempFile
    CreateObject("Shell.Application").Windows().Item(0).PutProperty strCode, Nothing
    strJS = "new ActiveXObject('Shell.Application').Windows().Item(0).PutProperty('" & strCode & "', " & strCode & ");" & vbNewLine & _
            "var exec = new ActiveXObject('WScript.Shell').Exec('MSHTA.EXE -');" & vbNewLine & _
            "var wmi_service = GetObject('winmgmts:');" & vbNewLine & _
            "var current_pid = wmi_service.Get('Win32_Process.Handle=' + exec.ProcessID).ParentProcessID;" & vbNewLine & _
            "exec.Terminate();" & vbNewLine & _
            "var parent_pid = wmi_service.Get('Win32_Process.Handle=' + current_pid).ParentProcessID;" & vbNewLine & _
            "while (wmi_service.ExecQuery('SELECT * FROM Win32_Process WHERE ProcessID=' + parent_pid).Count != 0)" & vbNewLine & _
            "    WScript.Sleep(1000);" & vbNewLine & _
            "var fso = new ActiveXObject('Scripting.FileSystemObject');" & vbNewLine & _
            "if (fso.FileExists(WScript.ScriptFullName))" & vbNewLine & _
            "    fso.DeleteFile(WScript.ScriptFullName);" & vbNewLine & _
            ""
    With CreateObject("Scripting.FileSystemObject")
        Do
            strTempFile = .BuildPath(.GetSpecialFolder(2), .GetTempName() & ".js")
        Loop While .FileExists(strTempFile)
        With .OpenTextFile(strTempFile, 2, True) ' ForWriting = 2, ForAppending = 8
            .WriteLine strJS
            .Close
        End With
    End With
    With CreateObject("WScript.Shell").Environment("Process")
        .Item("SysWOW64")     = CreateObject("Scripting.FileSystemObject").BuildPath(.Item("SystemRoot"), "SysWOW64")
        .Item("WScriptName")  = CreateObject("Scripting.FileSystemObject").GetFileName(WScript.FullName)
        .Item("WScriptWOW64") = CreateObject("Scripting.FileSystemObject").BuildPath(.Item("SysWOW64"), .Item("WScriptName"))
        .Item("Run") = .Item("WScriptWOW64") & " """ & strTempFile & """"
        CreateObject("WScript.Shell").Run .Item("Run")
    End With
    Do
        Set CreateObjectJS = CreateObject("Shell.Application").Windows().Item(0).GetProperty(strCode)
    Loop While CreateObjectJS Is Nothing
End Function


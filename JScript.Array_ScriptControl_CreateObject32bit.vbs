




Set objArray2 = TestJSArray()

WScript.Echo objArray2

objArray2.sort()

For Each strItem In objArray2
    WScript.Echo strItem
Next

Function TestJSArray()
    Set objArray = JSArray()
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

Function JSArray()
    With CreateObject32bit("ScriptControl")
	    .Language = "JScript"
	    Set JSArray = .Eval("new Array()")
    End With
End Function

Function CreateObject32bit(strClassName)
    Dim str32
    Dim strTempFile
    CreateObject("Shell.Application").Windows().Item(0).PutProperty strClassName, Nothing
    str32 = "CreateObject(""Shell.Application"").Windows().Item(0).PutProperty """ & strClassName & """, CreateObject(""" & strClassName & """)" & vbNewLine & _
            "With CreateObject(""WScript.Shell"").Exec(""MSHTA.EXE -"")" & vbNewLine & _
            "    lngCurrentPID = GetObject(""winmgmts:"").Get(""Win32_Process.Handle="" & .ProcessID).ParentProcessID" & vbNewLine & _
            "    .Terminate" & vbNewLine & _
            "End With" & vbNewLine & _
            "lngParentPID = GetObject(""winmgmts:"").Get(""Win32_Process.Handle="" & lngCurrentPID).ParentProcessID" & vbNewLine & _
            "Do While GetObject(""winmgmts:"").ExecQuery(""SELECT * FROM Win32_Process WHERE ProcessID="" & lngParentPID).Count<>0" & vbNewLine & _
            "    WScript.Sleep 1000" & vbNewLine & _
            "Loop" & vbNewLine & _
            "With CreateObject(""Scripting.FileSystemObject"")" & vbNewLine & _
            "    If .FileExists(WScript.ScriptFullName) Then .DeleteFile WScript.ScriptFullName" & vbNewLine & _
            "End With" & vbNewLine & _
            ""
    With CreateObject("Scripting.FileSystemObject")
        Do
            strTempFile = .BuildPath(.GetSpecialFolder(2), .GetTempName() & ".vbs")
        Loop While .FileExists(strTempFile)
        With .OpenTextFile(strTempFile, 2, True) ' ForWriting = 2, ForAppending = 8
            .WriteLine str32
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
        Set CreateObject32bit = CreateObject("Shell.Application").Windows().Item(0).GetProperty(strClassName)
    Loop While CreateObject32bit Is Nothing
End Function


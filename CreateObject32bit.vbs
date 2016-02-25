
Call Main

Sub Main
    str  = "abc 123 ‚  ˆŸ"
    str1 = "encodeURI(str):"            & encodeURI(str)
    str2 = "decodeURI(encodeURI(str)):" & decodeURI(encodeURI(str))
    str3 = "Escape(str):"               & Escape(str)
    str4 = "Unescape(Escape(str)):"     & Unescape(Escape(str))
    str5 = "Unescape(encodeURI(str)):"  & Unescape(encodeURI(str)) ' •¶Žš‰»‚¯
'   str6 = "decodeURI(Escape(str)):"    & decodeURI(Escape(str))   ' error
    WScript.Echo(   str  & vbCrLf & _
                    str1 & vbCrLf & _
                    str2 & vbCrLf & _
                    str3 & vbCrLf & _
                    str4 & vbCrLf & _
                    str5 )
End Sub

Function encodeURI(strURI)
	With CreateObject32bit("ScriptControl")
		.Language = "JScript" ' "VBScript", "JavaScript"
		encodeURI = .CodeObject.encodeURI(strURI)
	End With
End Function

Function decodeURI(strURI)
	With CreateObject32bit("ScriptControl")
		.Language = "JScript" ' "VBScript", "JavaScript"
		decodeURI = .CodeObject.decodeURI(strURI)
	End With
End Function

Function CreateObject32bit(strClassName)
    Dim str32
    Dim strTempFile
    CreateObject("Shell.Application").Windows().Item(0).PutProperty strClassName, Nothing
    str32 = "CreateObject(""Shell.Application"").Windows().Item(0).PutProperty """ & strClassName & """, CreateObject(""" & strClassName & """)" & vbNewLine & _
            "Set objExec = CreateObject(""WScript.Shell"").Exec(""MSHTA.EXE -"")" & vbNewLine & _
            "Set objWMIService = GetObject(""winmgmts:"")" & vbNewLine & _
            "lngCurrentPID = objWMIService.Get(""Win32_Process.Handle="" & objExec.ProcessID).ParentProcessID" & vbNewLine & _
            "objExec.Terminate" & vbNewLine & _
            "lngParentPID = objWMIService.Get(""Win32_Process.Handle="" & lngCurrentPID).ParentProcessID" & vbNewLine & _
            "Do While objWMIService.ExecQuery(""SELECT * FROM Win32_Process WHERE ProcessID="" & lngParentPID).Count<>0" & vbNewLine & _
            "    WScript.Sleep 1000" & vbNewLine & _
            "Loop" & vbNewLine & _
            "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbNewLine & _
            "If objFSO.FileExists(WScript.ScriptFullName) Then objFSO.DeleteFile WScript.ScriptFullName" & vbNewLine & _
            ""
    With CreateObject("Scripting.FileSystemObject")
        Do
            strTempFile = .BuildPath(.GetSpecialFolder(2), .GetTempName() & ".vbs") ' Const TemporaryFolder = 2
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


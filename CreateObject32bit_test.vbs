
Call Main

Sub Main
    str  = "abc 123 ‚  ˆŸ"
    str1 = "encodeURI(str):"            & encodeURI(str)
    str2 = "decodeURI(encodeURI(str)):" & decodeURI(encodeURI(str))
    str3 = "Escape(str):"               & Escape(str)
    str4 = "Unescape(Escape(str)):"     & Unescape(Escape(str))
    str5 = "Unescape(encodeURI(str)):"  & Unescape(encodeURI(str)) ' •¶Žš‰»‚¯
'   str6 = "decodeURI(Escape(str)):"    & decodeURI(Escape(str))   ' error
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
    Dim lngCurrentPID
    Dim strKey
    Dim strWS
    Dim strTempFolder
    Dim strTempName
    Dim strTempFile
    With CreateObject("WScript.Shell").Exec("MSHTA.EXE -")
        lngCurrentPID = GetObject("winmgmts:Win32_Process.Handle=" & .ProcessID).ParentProcessID
        .Terminate
    End With
    With CreateObject("Scripting.FileSystemObject")
        Do
            strTempFolder = .GetSpecialFolder(2) ' Const TemporaryFolder = 2
            strTempName = .GetTempName()
            strTempFile = .BuildPath(strTempFolder, strTempName & ".vbs") ' Const TemporaryFolder = 2
        Loop While .FileExists(strTempFile)
        strKey = WScript.ScriptName & "/" & CStr(lngCurrentPID) & "/" & CStr(Timer) & "/" & strTempName
        CreateObject("Shell.Application").Windows().Item(0).PutProperty strKey, Nothing
        strWS = "CreateObject(""Shell.Application"").Windows().Item(0).PutProperty """ & strKey & """, CreateObject(""" & strClassName & """)" & vbNewLine & _
                "Do While GetObject(""winmgmts:"").ExecQuery(""SELECT * FROM Win32_Process WHERE ProcessID=" & CStr(lngCurrentPID) & """).Count>0" & vbNewLine & _
                "    WScript.Sleep 1000" & vbNewLine & _
                "Loop" & vbNewLine & _
                "If CreateObject(""Scripting.FileSystemObject"").FileExists(WScript.ScriptFullName) Then" & vbNewLine & _
                "    CreateObject(""Scripting.FileSystemObject"").DeleteFile WScript.ScriptFullName" & vbNewLine & _
                "End If" & vbNewLine & _
                ""
        With .OpenTextFile(strTempFile, 2, True) ' ForWriting = 2, ForAppending = 8
            .WriteLine strWS
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
        Set CreateObject32bit = CreateObject("Shell.Application").Windows().Item(0).GetProperty(strKey)
    Loop While CreateObject32bit Is Nothing
End Function


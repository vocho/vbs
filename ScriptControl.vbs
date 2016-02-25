' Escape   Function �� encodeURI Method (Windows Scripting - JScript)
' Unescape Function �� decodeURI Method (Windows Scripting - JScript)

Call Run32bit
Call Main

Sub Main
    str  = "abc 123 �� ��"
    str1 = "encodeURI(str):"            & encodeURI(str)
    str2 = "decodeURI(encodeURI(str)):" & decodeURI(encodeURI(str))
    str3 = "Escape(str):"               & Escape(str)
    str4 = "Unescape(Escape(str)):"     & Unescape(Escape(str))
    str5 = "Unescape(encodeURI(str)):"  & Unescape(encodeURI(str)) ' ��������
'   str6 = "decodeURI(Escape(str)):"    & decodeURI(Escape(str))   ' error
    WScript.Echo(   str  & vbCrLf & _
                    str1 & vbCrLf & _
                    str2 & vbCrLf & _
                    str3 & vbCrLf & _
                    str4 & vbCrLf & _
                    str5 )
End Sub

Function encodeURI(strURI)
	With CreateObject("ScriptControl")
		.Language = "JScript" ' "VBScript", "JavaScript"
		encodeURI = .CodeObject.encodeURI(strURI)
	End With
End Function

Function decodeURI(strURI)
	With CreateObject("ScriptControl")
		.Language = "JScript" ' "VBScript", "JavaScript"
		decodeURI = .CodeObject.decodeURI(strURI)
	End With
End Function

Sub Run32bit
    With CreateObject("WScript.Shell").Environment("Process")
        If .Item("PROCESSOR_ARCHITECTURE")="AMD64" Then ' AMD64, x86
            Dim strArg
            .Item("SysWOW64")     = CreateObject("Scripting.FileSystemObject").BuildPath(.Item("SystemRoot"), "SysWOW64")
            .Item("WScriptName")  = CreateObject("Scripting.FileSystemObject").GetFileName(WScript.FullName)
            .Item("WScriptWOW64") = CreateObject("Scripting.FileSystemObject").BuildPath(.Item("SysWOW64"), .Item("WScriptName"))
            .Item("Run") = """" & .Item("WScriptWOW64") & """ """ & WScript.ScriptFullName & """"
            For Each strArg In WScript.Arguments
                .Item("Run") = .Item("Run") & " """ & strArg & """"
            Next
            CreateObject("WScript.Shell").Run .Item("Run")
            WScript.Quit
        End If
    End With
End Sub


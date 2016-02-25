' Escape   Function Å‡ encodeURI Method (Windows Scripting - JScript)
' Unescape Function Å‡ decodeURI Method (Windows Scripting - JScript)

Call Run32bit
Call Main

Sub Main
    str  = "http://abc 123 Ç† àü '\'/"
    str1 = "encodeURI(str):"                                & encodeURI(str)
    str2 = "decodeURI(encodeURI(str)):"                     & decodeURI(encodeURI(str))
    str3 = "Escape(str):"                                   & Escape(str)
    str4 = "Unescape(Escape(str)):"                         & Unescape(Escape(str))
    str5 = "Unescape(encodeURI(str)):"                      & Unescape(encodeURI(str)) ' ï∂éöâªÇØ
    str6 = "encodeURIComponent(str):"                       & encodeURIComponent(str)
    str7 = "decodeURIComponent(encodeURIComponent(str)):"   & decodeURIComponent(encodeURIComponent(str))
'   strx = "decodeURI(Escape(str)):"    & decodeURI(Escape(str))   ' error
    WScript.Echo(   str  & vbCrLf & _
                    str1 & vbCrLf & _
                    str2 & vbCrLf & _
                    str3 & vbCrLf & _
                    str4 & vbCrLf & _
                    str5 & vbCrLf & _
                    str6 & vbCrLf & _
                    str7 )
End Sub

Function encodeURI(strURI)
	With CreateObject("ScriptControl")
		.Language = "JScript"
		encodeURI = .CodeObject.encodeURI(strURI)
	End With
End Function

Function decodeURI(strURI)
	With CreateObject("ScriptControl")
		.Language = "JScript"
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

Function encodeURIComponent(ByVal strURI)
    With CreateObject("htmlfile")
        Call .appendChild(.createElement("a"))
        .lastChild.innerText = strURI
        Call .parentWindow.execScript("document.lastChild.innerText = encodeURIComponent(document.lastChild.innerText);", "JScript")
        encodeURIComponent = .lastChild.innerText
    End With
End Function

Function decodeURIComponent(ByVal strURI)
    With CreateObject("htmlfile")
        Call .appendChild(.createElement("a"))
        .lastChild.innerText = strURI
        Call .parentWindow.execScript("document.lastChild.innerText = decodeURIComponent(document.lastChild.innerText);", "JScript")
        decodeURIComponent = .lastChild.innerText
    End With
End Function


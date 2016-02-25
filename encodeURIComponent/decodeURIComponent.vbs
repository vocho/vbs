Call Run32bit

Call InputBox("decodeURIComponent‚µ‚Ü‚µ‚½", "decodeURIComponent", decodeURIComponent(InputBox("decodeURIComponent‚·‚é•¶Žš‚ð“ü—Í‚µ‚Ä‚­‚¾‚³‚¢", "decodeURIComponent")))

Function decodeURIComponent(ByVal strURI)
    With CreateObject("htmlfile")
        Call .appendChild(.createElement("a"))
        .lastChild.innerText = strURI
        Call .parentWindow.execScript("document.lastChild.innerText = decodeURIComponent(document.lastChild.innerText);", "JScript")
        decodeURIComponent = .lastChild.innerText
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

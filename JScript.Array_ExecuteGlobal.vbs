
Call Run32bit

Set objArray = JSArray()

objArray.push "F"

objArray.push("B")
objArray.push("C")
objArray.push("E")
objArray.push("D")
objArray.push("A")

objArray.sort()
objArray.reverse()

For Each strItem In objArray
    WScript.Echo strItem
Next


Function JSArray
    Dim objSC
    ExecuteGlobal("Set objSC = CreateObject(""ScriptControl"")")
	objSC.Language = "JScript"
	Set JSArray = objSC.Eval("new Array()")
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


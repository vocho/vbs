
WScript.Echo(CreateObject("WScript.Shell").Environment("Process").Item("PROCESSOR_ARCHITECTURE"))
Call RunCScript32bit

Sub RunCScript32bit
    With CreateObject("WScript.Shell").Environment("Process")
        If (.Item("PROCESSOR_ARCHITECTURE")="AMD64") Or (LCase(CreateObject("Scripting.FileSystemObject").GetBaseName(WScript.FullName))<>"cscript") Then ' AMD64, x86
            Dim strArg
            .Item("WScript.Shell.Run") = "cscript.exe"
            .Item("WScript.Shell.Run") = CreateObject("Scripting.FileSystemObject").BuildPath("SysWOW64", .Item("WScript.Shell.Run"))
            .Item("WScript.Shell.Run") = CreateObject("Scripting.FileSystemObject").BuildPath(.Item("SystemRoot"), .Item("WScript.Shell.Run"))
            .Item("WScript.Shell.Run") = """" & .Item("WScript.Shell.Run") & """ """ & WScript.ScriptFullName & """"
            For Each strArg In WScript.Arguments
                .Item("WScript.Shell.Run") = .Item("WScript.Shell.Run") & " """ & strArg & """"
            Next
            CreateObject("WScript.Shell").Run .Item("WScript.Shell.Run")
            WScript.Quit
        End If
    End With
End Sub


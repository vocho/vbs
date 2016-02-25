
WScript.Echo(CreateObject("WScript.Shell").ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%"))
Call Run32bit

Sub Run32bit
    With CreateObject("WScript.Shell")
        If .ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")="AMD64" Then ' AMD64, x86
            Dim strSystemRoot, strCommand, strArg
            strSystemRoot = .ExpandEnvironmentStrings("%SystemRoot%")
            With CreateObject("Scripting.FileSystemObject")
                strCommand = """" & .BuildPath(strSystemRoot, .BuildPath("SysWOW64", .GetFileName(WScript.FullName))) & """" & " " & """" & WScript.ScriptFullName & """"
            End With
            For Each strArg In WScript.Arguments
                strCommand = strCommand  & " " & """" & strArg & """"
            Next
            .Run strCommand
            WScript.Quit
        End If
    End With
End Sub


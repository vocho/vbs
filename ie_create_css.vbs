
Const READYSTATE_UNINITIALIZED  = 0
Const READYSTATE_LOADING        = 1
Const READYSTATE_LOADED         = 2
Const READYSTATE_INTERACTIVE    = 3
Const READYSTATE_COMPLETE       = 4

Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
blnIE_OnQuit = False
strNavigate = "http://www.ugtop.com/i/"

Dim objIE

Call Main

Sub Main
    Set objIE = WScript.CreateObject("InternetExplorer.Application", "IE_")
    With objIE
        .Left = 0
        .Top = 0
        .Width = 700
        .Height = 1050
        .ToolBar = False
        .StatusBar = False
        .Navigate strNavigate
        .Visible = True
        
        Do While .Busy Or .ReadyState<>READYSTATE_COMPLETE
           WScript.Sleep(100)
        Loop
        
        Do Until blnIE_OnQuit
            WScript.Sleep(1000)
        Loop
    End With
    Set objIE = Nothing
End Sub

Sub IE_OnQuit()
    blnIE_OnQuit = True
End Sub

Sub IE_DocumentComplete(pDisp, URL)
    If URL=strNavigate Then
        Set objStyleSheet = objIE.Document.createStyleSheet()
        'objStyleSheet.cssText = "body{background-color:pink;}"
        Call objStyleSheet.addRule("body", "background-color:pink")
        Call objStyleSheet.addRule("font", "color:orange")
        Call objStyleSheet.addRule(".", "color:green")
        Call objStyleSheet.addRule(".", "font-size:64px")
        MsgBox(objStyleSheet.cssText)
    End If
End Sub

Set WshShell = Nothing
Set fso = Nothing

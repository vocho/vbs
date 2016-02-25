
Const READYSTATE_UNINITIALIZED  = 0
Const READYSTATE_LOADING        = 1
Const READYSTATE_LOADED         = 2
Const READYSTATE_INTERACTIVE    = 3
Const READYSTATE_COMPLETE       = 4

Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
blnIE_OnQuit = False
strNavigate = "http://honyaku.yahoo.co.jp/transtext"

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
    End If
End Sub

Sub IE_BeforeNavigate2(pDisp, Url, Flags, TargetFrameName, PostData, Headers, Cancel)
    Const adTypeBinary = 1
    Const adTypeText   = 2
    If LenB(PostData)>0 Then
        strReadText = ""
        With CreateObject("ADODB.Stream")
            .Open
            .Type = adTypeBinary
            .Write(PostData)
            .Position = 0
            .Type = adTypeText
            .Charset = "_autodetect"
            strReadText = .ReadText
            .Close
        End With
        MsgBox(Join(Split(strReadText, "&"),", "))
    End If
End Sub


Sub StrConvxxxx()
    Dim strTargetFilePath ' As String
    Dim lngStart
    Dim lngLen
    
    strTargetFilePath =  "test.txt"
    lngStart = 2
    lngLen   = 6

    Dim bytData
    Dim strText
    Const adTypeBinary = 1
    Const adTypeText   = 2
    
    With CreateObject("ADODB.Stream")
        .Open
        .Type = adTypeBinary
        .LoadFromFile strTargetFilePath
        .Position = lngStart
        bytData = .Read(lngLen)
        .Close
    End With
    
    With CreateObject("ADODB.Stream")
        .Open
        .Type = adTypeBinary
        .Write bytData
        .Position = 0
        .Type = adTypeText
        .CharSet = "shift_jis"
        strText = .ReadText
        .Close
    End With
End Sub



Set WshShell = Nothing
Set fso = Nothing

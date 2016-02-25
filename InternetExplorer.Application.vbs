
With CreateObject("InternetExplorer.Application")
    .Left = 0
    .Top = 0
    .Width = 1500
    .Height = 1000
    .ToolBar = True
    .StatusBar = True
    .Navigate "http://www.yahoo.co.jp/"
    .Visible = True
    
    Do While .Busy Or .ReadyState<>4 ' READYSTATE_COMPLETE
       WScript.Sleep(100)
    Loop
    WScript.Echo(typename(.Document))
    WScript.Echo(typename(.Document.cookie))
    WScript.Echo(.Document.cookie)
    .Quit()
End With

With CreateObject("InternetExplorer.Application")
    .Left = 0
    .Top = 0
    .Width = 1500
    .Height = 1000
    .ToolBar = True
    .StatusBar = True
    .Navigate "http://www.nicovideo.jp/login"
    .Visible = True
    
    Do While .Busy Or .ReadyState<>4 ' READYSTATE_COMPLETE
       WScript.Sleep(100)
    Loop
    
    For Each objForm In .Document.getElementsByTagName("form")
        For Each objInput In objForm.getElementsByTagName("input")
            If objInput.getAttribute("id")="mail" Then
                objInput.focus
                objInput.value = "windki@gmail.com"
            ElseIf objInput.getAttribute("id")="password" Then
                objInput.focus
                objInput.value = "vovvovvovvov"
            End If
        Next
        For Each objInput In objForm.getElementsByTagName("input")
            If objInput.getAttribute("type")="submit" Then
                objInput.focus
                objInput.click
            End If
        Next
    Next
    
    Do While .Busy Or .ReadyState<>4 ' READYSTATE_COMPLETE
       WScript.Sleep(100)
    Loop
    
    .Navigate "http://flapi.nicovideo.jp/api/getflv/1356673087"
    
    Do While .Busy Or .ReadyState<>4 ' READYSTATE_COMPLETE
       WScript.Sleep(100)
    Loop
    
    strText = Unescape(.Document.body.innerText)
    
    With New RegExp
        .Pattern = "\b" & "url=([^&]+)"
        For Each objMatch In .Execute(strText)
            For Each strSubMatch In objMatch.SubMatches
                WScript.Echo(strSubMatch)
                If Right(strSubMatch, Len("low"))="low" Then
                    WScript.Echo("low")
                End If
            Next
        Next
    End With
    
    
    .Quit()
End With

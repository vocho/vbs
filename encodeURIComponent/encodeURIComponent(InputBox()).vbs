
Call InputBox("", "", encodeURIComponent(InputBox("", "", "http://example.com")))

Function encodeURIComponent(strURI)
    With CreateObject("htmlfile")
        Call .appendChild(.createElement("a"))
        .lastChild.innerText = strURI
        Call .parentWindow.execScript("document.lastChild.innerText = encodeURIComponent(document.lastChild.innerText);", "JScript")
        encodeURIComponent = .lastChild.innerText
    End With
End Function

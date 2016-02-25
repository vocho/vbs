With CreateObject("htmlfile")
    MsgBox .parentWindow.clipboardData.getData("text")
    MsgBox .parentWindow.clipboardData.setData("text", WScript.FullName)
    MsgBox .parentWindow.clipboardData.getData("text")
End With
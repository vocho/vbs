strSpecialFolders = ""
For Each strSpecialFolder In CreateObject("WScript.Shell").SpecialFolders
    strSpecialFolders = strSpecialFolders & strSpecialFolder & vbCrLf
Next
WScript.Echo(strSpecialFolders)

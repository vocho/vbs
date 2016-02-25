WScript.Echo(TypeName(CreateObject("WScript.Shell").SpecialFolders))
WScript.Echo(TypeName(CreateObject("htmlfile")))
WScript.Echo(TypeName(WScript.Arguments))
For Each Argument In WScript.Arguments
    WScript.Echo(TypeName(Argument))
Next
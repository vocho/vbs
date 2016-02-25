With CreateObject("Scripting.FileSystemObject")
    strFolder = .GetParentFolderName(WScript.ScriptFullName)
    strFile = .GetBaseName(WScript.ScriptName)
End With
WScript.Sleep(120000)
With CreateObject("WScript.Shell")
    .CurrentDirectory = strFolder
    .Exec(strFile)
End With

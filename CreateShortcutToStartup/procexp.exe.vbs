
With CreateObject("WScript.Shell")
    strSpecialFolder = .SpecialFolders("Startup")
    With CreateObject("Scripting.FileSystemObject")
        strFileName = .GetBaseName(WScript.ScriptName)
        strParentFolder = .GetParentFolderName(WScript.ScriptFullName)
        strFilePath = .BuildPath(strParentFolder, strFileName)
        strShortcutPath = .BuildPath(strSpecialFolder, strFileName & ".lnk")
    End With
    With .CreateShortcut(strShortcutPath)
        .TargetPath = strFilePath
        .WorkingDirectory = strParentFolder
        .WindowStyle = 7
        .Save
    End With
End With
With CreateObject("Shell.Application")
    .Open strSpecialFolder
End With


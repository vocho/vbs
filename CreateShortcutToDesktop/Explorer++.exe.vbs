
With CreateObject("WScript.Shell")
    strSpecialFolder = .SpecialFolders("Desktop")
    With CreateObject("Scripting.FileSystemObject")
        strFileName = .GetBaseName(WScript.ScriptName)
        strParentFolder = .GetParentFolderName(WScript.ScriptFullName)
        strFilePath = .BuildPath(strParentFolder, strFileName)
        strShortcutPath = .BuildPath(strSpecialFolder, strFileName & ".lnk")
    End With
    With .CreateShortcut(strShortcutPath)
        .TargetPath = strFilePath
        .WorkingDirectory = strParentFolder
        .WindowStyle = 1
        .Save
    End With
End With
With CreateObject("Shell.Application")
    .Open strSpecialFolder
End With


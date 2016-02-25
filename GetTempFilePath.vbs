Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

WScript.Echo GetTempFilePath

Function GetTempFilePath
    Const TemporaryFolder = 2
    Dim fso
    Dim strTempFolder
    Dim strTempFile
    Dim strTempFilePath
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Do
        strTempFolder = fso.GetSpecialFolder(TemporaryFolder)
        strTempFile = fso.GetTempName()
        strTempFilePath = fso.BuildPath(strTempFolder, strTempFile)
    Loop While fso.FileExists(strTempFilePath)
    
    GetTempFilePath = strTempFilePath
End Function


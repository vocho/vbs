Const FILE_NAME = 0

For Each strArg In WScript.Arguments
    If Err.Number = 0 Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    	If objFSO.FileExists(strArg) Then
            strVersion = objFSO.GetFileVersion(strArg)
            strFolderName = objFSO.GetParentFolderName(strArg)
            strBaseName = objFSO.GetBaseName(strArg)
            strExtensionName = objFSO.GetExtensionName(strArg)
            strDstPath = objFSO.BuildPath(strFolderName, strBaseName & "_" & strVersion & "." & strExtensionName)
            
            Set objFile = objFSO.GetFile(strArg)
            objFile.Copy(strDstPath)
    	Else
            
    	End If
    Else
        WScript.Echo "ÉGÉâÅ[: " & Err.Description
    End If
Next

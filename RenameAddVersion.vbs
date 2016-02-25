Const FILE_NAME = 0

For Each strArg In WScript.Arguments
    If Err.Number = 0 Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    	If objFSO.FileExists(strArg) Then
            strGetFilePath = Replace(strArg, "\", "\\")
            strGet = "CIM_Datafile.name='" & strGetFilePath & "'"
            Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
            Set objCIM_DataFile = objWMIService.Get(strGet)
            strVersion = Replace(objCIM_DataFile.Version, " ", "_")
            
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

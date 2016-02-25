Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strScriptBaseName = objFSO.GetBaseName(WScript.ScriptFullName)
strOutTextPath = objFSO.BuildPath(strScriptPath, strScriptBaseName & ".txt")

Set objTS = objFSO.CreateTextFile(strOutTextPath, True, False)
For Each strArg In WScript.Arguments
    If Err.Number = 0 Then
    	If objFSO.FolderExists(strArg) Then
    		Set objFolder = objFSO.GetFolder(strArg)
    		Call FolderProc(objFolder)
    	ElseIf objFSO.FileExists(strArg) Then
    		Set objFile = objFSO.GetFile(strArg)
    		Call FileProc(objFile)
        Else
            WScript.Echo "Error: " & Err.Description
    	End If
    Else
        WScript.Echo "Error: " & Err.Description
    End If
Next
objTS.Close()

Function FolderProc(objFolder)
    Call FileProc(objFolder)
    For Each objFile In objFolder.Files
        Call FileProc(objFile)
    Next
    For Each objSubFolder In objFolder.SubFolders
        Call FolderProc(objSubFolder)
    Next
End Function

Function FileProc(objFileOrFolder)
	With objFSO.OpenTextFile(objFileOrFolder.Path, ForReading, TristateFalse)
		objTS.Write(.ReadAll)
		objTS.WriteBlankLines(2)
		.Close
	End With
End Function


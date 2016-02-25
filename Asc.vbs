
Dim fso
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Dim strArg

For Each strArg In WScript.Arguments
    If Err.Number = 0 Then
        code = Asc(fso.GetBaseName(strArg))
        str = Chr(code)
        WScript.Echo code & ", " & str
    Else
        WScript.Echo "ÉGÉâÅ[: " & Err.Description
    End If
Next

Sub U2A(strArg)
    Dim objRE
    Set objRE = New RegExp
    Dim objFolder
    Dim objFile
    Dim aryIgnoreName
    Dim strIgnoreName
    
    aryIgnoreName = Array("\.txt$", "\.gif$", "\.url$", "\.lnk$", "\.html$", "\.htm$", "^Thumbs\.db$", "mimip2p.*\.jpg$", "mimip2p.*\.png$", "^_____padding_file_.*____$", "^99p2p@.*\.gif$", "www\.city9x\.com\.gif", "www\.city9x\.com\.jpg")
    For Each objFile In fso.GetFolder(strArg).Files
        For Each strIgnoreName In aryIgnoreName
            objRE.IgnoreCase = True
            objRE.Pattern = strIgnoreName
            If objRE.Test(objFile.Name) Then
                objFile.Delete(True)
                Exit For
            End If
        Next
    Next
    
    For Each objFolder In fso.GetFolder(strArg).SubFolders
        Call FolderProc(objFolder)
    Next
End Sub

Sub FolderProc(strArg)
    Dim objRE
    Set objRE = New RegExp
    Dim objFolder
    Dim objFile
    Dim aryIgnoreName
    Dim strIgnoreName
    
    aryIgnoreName = Array("\.txt$", "\.gif$", "\.url$", "\.lnk$", "\.html$", "\.htm$", "^Thumbs\.db$", "mimip2p.*\.jpg$", "mimip2p.*\.png$", "^_____padding_file_.*____$", "^99p2p@.*\.gif$", "www\.city9x\.com\.gif", "www\.city9x\.com\.jpg")
    For Each objFile In fso.GetFolder(strArg).Files
        For Each strIgnoreName In aryIgnoreName
            objRE.IgnoreCase = True
            objRE.Pattern = strIgnoreName
            If objRE.Test(objFile.Name) Then
                objFile.Delete(True)
                Exit For
            End If
        Next
    Next
    
    For Each objFolder In fso.GetFolder(strArg).SubFolders
        Call FolderProc(objFolder)
    Next
End Sub

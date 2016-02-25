Const ForReading    = 1
Const ForWriting    = 2
Const ForAppending  = 8

Set fso = CreateObject("Scripting.FileSystemObject")

Call Main

Sub Main
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    
    strHtmlPath = GetHtmlPath()
    strParentFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    
    Set objHtml = xmlDoc.createElement("html")
    Set objBody = xmlDoc.createElement("body")
    
    xmlDoc.appendChild(objHtml)
    objHtml.appendChild(objBody)
    
    Call FolderProc(strParentFolder, xmlDoc, objBody, 0, 0, "")
    Call SaveHtml(xmlDoc.xml, strHtmlPath)
    
    Set xmlDoc = Nothing
    
    Call OpenHtml(strHtmlPath)
End Sub

Function FolderProc(strFolderPath, xmlDoc, objBody, intCount, intTotalCount, strParentTree)
    Dim intChildCount, intChildTotalCount
    
    With fso.GetFolder(strFolderPath)
        intChildCount = 0
        intChildTotalCount = .SubFolders.Count + .Files.Count
        
        If intTotalCount=0 Then
            strSelfTree     = ""
            strChildTree    = ""
        ElseIf intCount=intTotalCount Then
            strSelfTree     = strParentTree & "Ñ§ "
            strChildTree    = strParentTree & "Å@ "
        Else
            strSelfTree     = strParentTree & "Ñ• "
            strChildTree    = strParentTree & "Ñ† "
        End If
        
        Set objTt = xmlDoc.createElement("tt")
        Set objText = xmlDoc.createTextNode(strSelfTree & .Name)
        Set objBr = xmlDoc.createElement("br")
        objBody.appendChild(objTt)
        objTt.appendChild(objText)
        objBody.appendChild(objBr)
        
        For Each objFolder In .SubFolders
            intChildCount = intChildCount + 1
            Call FolderProc(objFolder.Path, xmlDoc, objBody, intChildCount, intChildTotalCount, strChildTree)
        Next
        
        For Each objFile In .Files
            intChildCount = intChildCount + 1
            Call FileProc(objFile.Path, xmlDoc, objBody, intChildCount, intChildTotalCount, strChildTree)
        Next
    End With
End Function

Function FileProc(strFilePath, xmlDoc, objBody, intCount, intTotalCount, strParentTree)
    With fso.GetFile(strFilePath)
        If intCount=intTotalCount Then
            strSelfTree = strParentTree & "Ñ§ "
        Else
            strSelfTree = strParentTree & "Ñ• "
        End If
        
        Set objTt = xmlDoc.createElement("tt")
        Set objText = xmlDoc.createTextNode(strSelfTree & .Name)
        Set objBr = xmlDoc.createElement("br")
        objBody.appendChild(objTt)
        objTt.appendChild(objText)
        objBody.appendChild(objBr)
    End With
End Function

Function GetHtmlPath
    Dim strParentFolder, strBaseName, strFileName
    strParentFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    strBaseName = fso.GetBaseName(WScript.ScriptFullName)
    strFileName = strBaseName & ".html"
    GetHtmlPath = fso.BuildPath(strParentFolder, strFileName)
End Function

Sub WriteText(strText, strTextFilePath)
    With fso.CreateTextFile(strTextFilePath, True, True)
        .Write strText
        .Close
    End With
End Sub

Sub SaveHtml(strXml, strSavePath)
    Set xmlWriter = CreateObject("MSXML2.MXXMLWriter")
    Set xmlReader = CreateObject("MSXML2.SAXXMLReader")
    
    xmlWriter.indent = True
    xmlWriter.standalone = True
'    xmlWriter.omitXMLDeclaration = True
    Set xmlReader.contentHandler = xmlWriter
    Call xmlReader.parse(strXml)
    
    With CreateObject("MSXML2.DOMDocument")
        .loadXML(xmlWriter.output)
        .save(strSavePath)
    End With
    
    Set xmlReader = Nothing
    Set xmlWriter = Nothing
End Sub

Sub OpenHtml(strFilePath)
    With CreateObject("InternetExplorer.Application")
        Call .Navigate(strFilePath)
        .Visible = True
    End With
End Sub

Set fso = Nothing

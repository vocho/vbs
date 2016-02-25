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
    Set objOl = xmlDoc.createElement("ol")
    
    xmlDoc.appendChild(objHtml)
    objHtml.appendChild(objBody)
    objBody.appendChild(objOl)
    Call objOl.setAttribute("style", "list-style:none")
    
    Call FolderProc(strParentFolder, strParentFolder, xmlDoc, objOl, 0, 0, "")
    Call SaveHtml(xmlDoc.xml, strHtmlPath)
    
    Set xmlDoc = Nothing
    
    Call OpenHtml(strHtmlPath)
End Sub

Function FolderProc(strParentFolder, strFolderPath, xmlDoc, objParentElement, intCount, intTotalCount, strParentTree)
    Dim intChildCount, intChildTotalCount
    
    With fso.GetFolder(strFolderPath)
        intChildCount = 0
        intChildTotalCount = .SubFolders.Count + .Files.Count
        
        If intTotalCount=0 Then
            strSelfTree     = " "
            strChildTree    = " "
        ElseIf intCount=intTotalCount Then
            strSelfTree     = strParentTree & "Ñ§ "
            strChildTree    = strParentTree & "Å@ "
        Else
            strSelfTree     = strParentTree & "Ñ• "
            strChildTree    = strParentTree & "Ñ† "
        End If
        
        Set objLi = xmlDoc.createElement("li")
        Set objTreeTt = xmlDoc.createElement("tt")
        Set objTreeText = xmlDoc.createTextNode(strSelfTree)
        Set objNameA = xmlDoc.createElement("a")
        Set objNameText = xmlDoc.createTextNode(.Name)
        objParentElement.appendChild(objLi)
        objLi.appendChild(objTreeTt)
        objTreeTt.appendChild(objTreeText)
        objLi.appendChild(objNameA)
        strHref = Replace(.Path, strParentFolder, "")
        strHref = fso.BuildPath(".", strHref)
        Call objNameA.setAttribute("href", strHref)
        objNameA.appendChild(objNameText)
        
        For Each objFolder In .SubFolders
            intChildCount = intChildCount + 1
            Call FolderProc(strParentFolder, objFolder.Path, xmlDoc, objParentElement, intChildCount, intChildTotalCount, strChildTree)
        Next
        
        For Each objFile In .Files
            intChildCount = intChildCount + 1
            Call FileProc(strParentFolder, objFile.Path, xmlDoc, objParentElement, intChildCount, intChildTotalCount, strChildTree)
        Next
    End With
End Function

Function FileProc(strParentFolder, strFilePath, xmlDoc, objParentElement, intCount, intTotalCount, strParentTree)
    With fso.GetFile(strFilePath)
        If intCount=intTotalCount Then
            strSelfTree = strParentTree & "Ñ§ "
        Else
            strSelfTree = strParentTree & "Ñ• "
        End If
        
        Set objLi = xmlDoc.createElement("li")
        Set objTreeTt = xmlDoc.createElement("tt")
        Set objTreeText = xmlDoc.createTextNode(strSelfTree)
        Set objNameA = xmlDoc.createElement("a")
        Set objNameText = xmlDoc.createTextNode(.Name)
        objParentElement.appendChild(objLi)
        objLi.appendChild(objTreeTt)
        objTreeTt.appendChild(objTreeText)
        objLi.appendChild(objNameA)
        strHref = Replace(.Path, strParentFolder, "")
        strHref = fso.BuildPath(".", strHref)
        Call objNameA.setAttribute("href", strHref)
        objNameA.appendChild(objNameText)
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

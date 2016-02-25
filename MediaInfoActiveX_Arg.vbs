' In order to run this example, put MediaInfo.dll and MediaInfoActiveX.dll
' into your system directory and Example.ogg into the root directory of drive
' C: (i.e. C:\). Use regsvr32.exe which is provided with Windows to register
' MediaInfoActiveX.dll.
'
' Use at own risk, under the same license as MediaInfo itself.
'
' Ingo BrÅEkl, May 2006

Call Run32bit

Sub Run32bit
    With CreateObject("WScript.Shell").Environment("Process")
        If .Item("PROCESSOR_ARCHITECTURE")="AMD64" Then ' AMD64, x86
            Dim strArg
            .Item("SysWOW64")     = CreateObject("Scripting.FileSystemObject").BuildPath(.Item("SystemRoot"), "SysWOW64")
            .Item("WScriptName")  = CreateObject("Scripting.FileSystemObject").GetFileName(WScript.FullName)
            .Item("WScriptWOW64") = CreateObject("Scripting.FileSystemObject").BuildPath(.Item("SysWOW64"), .Item("WScriptName"))
            .Item("Run") = """" & .Item("WScriptWOW64") & """ """ & WScript.ScriptFullName & """"
            For Each strArg In WScript.Arguments
                .Item("Run") = .Item("Run") & " """ & strArg & """"
            Next
            CreateObject("WScript.Shell").Run .Item("Run")
            WScript.Quit
        End If
    End With
End Sub

Const MediaInfo_Stream_General  	= 0
Const MediaInfo_Stream_Video    	= 1
Const MediaInfo_Stream_Audio    	= 2
Const MediaInfo_Stream_Text     	= 3 
Const MediaInfo_Stream_Chapters 	= 4
Const MediaInfo_Stream_Image    	= 5
Const MediaInfo_Stream_Menu     	= 6
Const MediaInfo_Stream_Max     	    = 7

Const MediaInfo_Info_Name 	 		= 0 
Const MediaInfo_Info_Text 			= 1
Const MediaInfo_Info_Measure 		= 2
Const MediaInfo_Info_Options 		= 3
Const MediaInfo_Info_Name_Text 		= 4
Const MediaInfo_Info_Measure_Text 	= 5
Const MediaInfo_Info_Info 			= 6
Const MediaInfo_Info_HowTo 			= 7
Const MediaInfo_Info_Max            = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")

Call ArgProc

Set objFSO = Nothing

Function ArgProc
    For Each strArg In WScript.Arguments
        If objFSO.FolderExists(strArg) Then
            Call FolderProc(objFSO.GetFolder(strArg))
        ElseIf objFSO.FileExists(strArg) Then
        	Call FileProc(objFSO.GetFile(strArg))
        End If
    Next
End Function

Function FolderProc(objFolder)
    For Each objFile In objFolder.Files
        Call FileProc(objFile)
    Next
    For Each objSubFolder In objFolder.SubFolders
        Call FolderProc(objSubFolder)
    Next
End Function

Function FileProc(objFile)
    With CreateObject("MediaInfo.ActiveX")
        lngHandle = .MediaInfo_New()
        .MediaInfo_Open lngHandle, objFile.Path
        
        .MediaInfo_Option lngHandle, "Info_Parameters_CSV", ""
        
        .MediaInfo_Option lngHandle, "Complete", "1"
        strText = .MediaInfo_Inform(lngHandle, 0)
        strPath = objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), objFile.Name & ".Complete=1.txt")
        Call TextSaveToFile(strText, strPath)
        
        .MediaInfo_Option lngHandle, "Inform", "HTML"
        strText = .MediaInfo_Inform(lngHandle, 0)
        strPath = objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), objFile.Name & ".Complete=1.html")
        Call TextSaveToFile(strText, strPath)
        
        .MediaInfo_Option lngHandle, "Inform", "XML"
        strText = .MediaInfo_Inform(lngHandle, 0)
        strPath = objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), objFile.Name & ".Complete=1.xml")
        Call TextSaveToFile(strText, strPath)
        
        .MediaInfo_Option lngHandle, "Inform", "CSV"
        strText = .MediaInfo_Inform(lngHandle, 0)
        strPath = objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), objFile.Name & ".Complete=1.csv")
        Call TextSaveToFile(strText, strPath)
        
        .MediaInfo_Close lngHandle
        .MediaInfo_Delete lngHandle
    End With
End Function

Function TextSaveToFile(strText, strPath)
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText strText, 0
        .SaveToFile strPath, 2
        .Close
    End With
End Function

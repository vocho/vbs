
Const READYSTATE_UNINITIALIZED  = 0
Const READYSTATE_LOADING        = 1
Const READYSTATE_LOADED         = 2
Const READYSTATE_INTERACTIVE    = 3
Const READYSTATE_COMPLETE       = 4

Set WshShell = CreateObject("WScript.Shell")
Set dicIEList = CreateObject("Scripting.Dictionary")

Call Main

Sub Main
    Call ArgumentsProc
    Call MaintenanceProc
End Sub

Sub ArgumentsProc
    For Each strArg In WScript.Arguments
        AddIEList(strArg)
    Next
End Sub

Sub MaintenanceProc
    Do While dicIEList.Count>0
        WScript.Sleep(1000)
    Loop
End Sub

Function AddIEList(strArg)
    intID = dicIEList.Count
    Set dicIEList(intID) = New IE
    dicIEList(intID).intID = intID
    dicIEList(intID).Run
End Function

Function RemoveIEList(intID)
    Set dicIEList(intID) = Nothing
    dicIEList.Remove intID
End Function

Class IE
    Public  intID
    Private objApp
    
    Private Sub Class_Initialize()
    End Sub
    
    Private Sub Class_Terminate()
        Set objApp = Nothing
    End Sub
    
    Public Sub Run()
        Init()
    End Sub
    
    Public Sub DocumentComplete(pDisp, URL)
    End Sub
    
    Public Sub NavigateComplete2(pDisp, URL)
    End Sub
    
    Public Sub ProgressChange(nProgress, nProgressMax)
    End Sub
    
    Public Sub TitleChange(sText)
    End Sub
    
    Private Sub Init()
        DefineEventWrapper("OnQuit()")
        DefineEventWrapper("TitleChange(sText)")
        DefineEventWrapper("DocumentComplete(pDisp, URL)")
        DefineEventWrapper("NavigateComplete2(pDisp, URL)")
        DefineEventWrapper("ProgressChange(nProgress, nProgressMax)")
        
        strPrefix = "IE" & CStr(intID) & "_"
        Set objApp = WScript.CreateObject("InternetExplorer.Application", strPrefix)
        
        With objApp
            .Left = 0
            .Top = 0
            .Width = 700
            .Height = 1050
            .ToolBar = False
            .StatusBar = False
            .Navigate "http://www.yahoo.co.jp"
            .Visible = True
        End With
        
        Do While objApp.Busy Or objApp.ReadyState<>READYSTATE_COMPLETE
           WScript.Sleep(100)
        Loop
        
    End Sub
    
    Private Sub DefineEventWrapper(strEvent)
        strNum      = CStr(intID)
        strWrapper  = "IE" & strNum & "_" & strEvent
        With New RegExp
            .Pattern = "(.+)\((.*)\)"
            With .Execute(strEvent)
                strEventName = .Item(0).SubMatches(0)
                strEventArg  = .Item(0).SubMatches(1)
            End With
        End With
        With New RegExp
            .Pattern = "\S+"
            If .Test(strEventArg) Then
                strCallSub = "IE_" & strEventName & " " & strNum & ", " & strEventArg
            Else
                strCallSub = "IE_" & strEventName & " " & strNum
            End If
        End With
        strExecute  = "Sub " & strWrapper   & vbCrLf _
                    & "    " & strCallSub   & vbCrLf _
                    & "End Sub"             & vbCrLf
        ExecuteGlobal strExecute
    End Sub
End Class

Sub IE_OnQuit(intID)
    RemoveIEList(intID)
End Sub

Sub IE_DocumentComplete(intID, pDisp, URL)
    Call dicIEList(intID).DocumentComplete(pDisp, URL)
End Sub

Sub IE_NavigateComplete2(intID, pDisp, URL)
    Call dicIEList(intID).NavigateComplete2(pDisp, URL)
End Sub

Sub IE_TitleChange(intID, sText)
    Call dicIEList(intID).TitleChange(sText)
End Sub

Sub IE_ProgressChange(intID, nProgress, nProgressMax)
    On Error Resume Next
    Call dicIEList(intID).ProgressChange(nProgress, nProgressMax)
End Sub

Set dicIEList = Nothing
Set WshShell = Nothing

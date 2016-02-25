
Set fso = CreateObject("Scripting.FileSystemObject")

Const UseMilliSecond = True

Call Main

Sub Main
    strSuffix = GetSuffix()
    Call ArgProc(strSuffix)
End Sub

Function GetSuffix
    If UseMilliSecond Then
        With CreateObject("ScriptControl")
            .Language = "JScript"
            With .Eval("new Date()")
                strFullYear     = Mid(.getFullYear()        + 10000 , 2, 4)
                strMonth        = Mid(.getMonth()+1         + 100   , 2, 2)
                strDate         = Mid(.getDate()            + 100   , 2, 2)
                strHours        = Mid(.getHours()           + 100   , 2, 2)
                strMinutes      = Mid(.getMinutes()         + 100   , 2, 2)
                strSeconds      = Mid(.getSeconds()         + 100   , 2, 2)
                strMilliseconds = Mid(.getMilliseconds()    + 1000  , 2, 2)
                strSuffix   = strFullYear _
                            & strMonth _
                            & strDate _
                            & "_" _ 
                            & strHours _
                            & strMinutes _
                            & strSeconds _
                            & strMilliseconds
            End With
        End With
    Else
        strNow      = Now()
        strYear     = Mid(Year(strNow)      + 10000 , 2, 4)
        strMonth    = Mid(Month(strNow)     + 100   , 2, 2)
        strDay      = Mid(Day(strNow)       + 100   , 2, 2)
        strHour     = Mid(Hour(strNow)      + 100   , 2, 2)
        strMinute   = Mid(Minute(strNow)    + 100   , 2, 2)
        strSecond   = Mid(Second(strNow)    + 100   , 2, 2)
        strSuffix   = strYear _
                    & strMonth _
                    & strDay _
                    & "_" _ 
                    & strHour _
                    & strMinute _
                    & strSecond
    End If
    GetSuffix = strSuffix
End Function

Sub ArgProc(strSuffix)
    For Each strArg In WScript.Arguments
    	If fso.FolderExists(strArg) Then
    		strSrc = strArg
    		strDst = strArg & "_" & strSuffix
    		Call fso.CopyFolder(strSrc, strDst)
    	ElseIf fso.FileExists(strArg) Then
            strSrc = strArg
    		strParentFolderName = fso.GetParentFolderName(strArg)
    		strBaseName = fso.GetBaseName(strArg)
    		strExtensionName = fso.GetExtensionName(strArg)
    		strFileName = strBaseName & "_" & strSuffix & "." & strExtensionName
    		strDst = fso.BuildPath(strParentFolderName, strFileName)
    		Call fso.CopyFile(strSrc, strDst)
        Else
            
    	End If
    Next
End Sub

Set fso = Nothing

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'version 14 - made forcestreamtowrite public

Public Enum LogLevel
    eDEBUG = 10
    eINFO = 20
    eWARNING = 30
    eERROR = 40
    eCRITICAL = 50
    eOFF = 60
End Enum

Private Type TLogger
    ImmediateLevel As Long
    LogStreamColl As Collection
    errHandler As clsErrHandler
End Type

Private this As TLogger

Private Const SPACER As String = vbTab

'-----------------------------------------
'       Create
'-----------------------------------------
Public Function Create( _
        ImmediateWindowLevel As LogLevel _
        , LogStreamColl As Collection _
        , errHandler As clsErrHandler _
) As clsLogger

    On Error GoTo errH
    
    this.ImmediateLevel = ImmediateWindowLevel
    Set this.LogStreamColl = LogStreamColl
    Set this.errHandler = errHandler
    
    Set Create = Me
    
exitHere:
    Exit Function
    
errH:
    MsgBox "There was an error with clsLogger.Create"
    GoTo exitHere
    
End Function

Private Sub Class_Terminate()

    Dim stream As clsLogStream
    For Each stream In this.LogStreamColl
        Set stream = Nothing
    Next
    
End Sub

'-----------------------------------------
'       Interface
'-----------------------------------------

Public Sub TimeStamp(level As LogLevel)

    Dim msg As String
        msg = msg & vbNewLine & "---------------------------------------------------------------"
        msg = msg & vbNewLine & ThisWorkbook.FullName
        msg = msg & SPACER & Environ("Username")
        
        
    Call Log(msg, level)

End Sub

Public Sub Dbug(msg)
    
    Call Log("[DBUG] " & msg, level:=eDEBUG)
    
End Sub

Public Sub Info(msg As String)
    
    Call Log("[INFO] " & msg, level:=eINFO)
    
End Sub

Public Sub Warning(msg)

    Call Log("[WARN] " & msg & " | Err# " & Err.Number & ": " & Err.Description, level:=eWARNING)

End Sub

Public Sub Error(Optional msg As String = "")
    
    Dim text As String
    text = "[ERR ] " & msg & " | Err# " & Err.Number & ": " & Err.Description
    
    MsgBox text, vbCritical, "Error"
    Call Log(text, eERROR)
    
End Sub

Public Sub Critical(procName)
    
    Dim msg As String
    msg = "Runtime Error " & Err.Number & ": " & Err.Description & ". "
    If procName <> "" Then msg = msg & "Macro Script: " & procName
    Call Log("[CRIT] " & msg, eCRITICAL)
    
    Call ForceStreamsWriteToFile
    
    'this needs to come last! Otherwise it will send the log
    'before the log queues have been written to file
    Call this.errHandler.SubmitError(msg, GetFnames)
    
exitHere:
    Exit Sub
    
errH:
    MsgBox "There was an error with clsLogger.Critical"
    GoTo exitHere
    
End Sub



'-----------------------------------------
'       Private Subs
'-----------------------------------------

Public Sub Log(msg As String, level As LogLevel)
    
    On Error GoTo errH
    
    msg = Now() & SPACER & msg
    
    If level >= this.ImmediateLevel Then Debug.Print msg
    
    Dim stream As clsLogStream
    For Each stream In this.LogStreamColl
        Call stream.SubmitLog(msg, level)
    Next
    
    Exit Sub
    
errH:
    MsgBox "There was an error with clsLogger.Log"
    Resume Next
    
End Sub

Public Sub ForceStreamsWriteToFile()

    Dim s As clsLogStream
    For Each s In this.LogStreamColl
        s.ForceWriteToFile
    Next

End Sub

Private Function GetFnames() As Collection

    Set GetFnames = New Collection
    
    Dim s As clsLogStream
    For Each s In this.LogStreamColl
        GetFnames.Add s.fname
    Next
    
End Function

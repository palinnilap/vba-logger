VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TErrH
    msg As String
    fnameColl As Collection
    outApp As Object
    OutDisabled As Boolean
End Type

Private this As TErrH


'--------------------------
'      Initialize
'--------------------------


Private Sub Class_Terminate()
    
    Set this.outApp = Nothing
    
End Sub



'--------------------------
'      Interface
'--------------------------
Public Sub SubmitError(msg As String, fnameColl As Collection)

    'usually this will sub will be called from the logger object
    On Error GoTo errH
    
    Set this.fnameColl = fnameColl
    this.msg = msg
        
    Call EmailMsgBox
    
    Exit Sub
    
errH:
    MsgBox "The ErrHandler class crashed."
    
End Sub



'--------------------------
'      Private Subs
'--------------------------

Private Sub EmailMsgBox()

    Dim choice As VbMsgBoxResult
    
    choice = MsgBox( _
                "ERROR. Email Nick the error log files for debugging." _
                    & vbNewLine & vbNewLine & "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -" _
                    & vbNewLine & vbNewLine _
                    & this.msg _
                , vbCritical + vbOKCancel _
                , Title:="Critical Error" _
                )
    
    
    If choice = vbOK Then
        
        Call SendEmail
        
    End If

End Sub

Private Sub SendEmail()

    'PURPOSE: Opens a new outlook application instance (if not already open)
    'to email nick a copy of the error logs for debugging purposes

    On Error Resume Next
    Set this.outApp = CreateObject("Outlook.Application")
    If Err.Number <> 0 Then
        Exit Sub
    End If

    Call ComposeEmail( _
        "sample@sample.com" _
        , HTMLbody:=Now() _
        , subject:="MACRO ERROR REPORT" _
        , fnameColl:=this.fnameColl _
        , send:=True)
        
    Set this.outApp = Nothing

End Sub

Public Sub ComposeEmail( _
    recipient As String _
    , Optional HTMLbody As String _
    , Optional subject As String _
    , Optional fnameColl As Collection _
    , Optional send As Boolean = False)

    Dim Email As Object
    
    Set Email = this.outApp.CreateItem(0)
    
    With Email
    
        '.display  <-- this will make the signature appear, but also slows down creation
        .subject = subject
        .HTMLbody = HTMLbody & .HTMLbody
        .To = recipient
        
        If fnameColl.count > 0 Then
            Dim x As Variant
            For Each x In fnameColl
                .Attachments.Add x
            Next
        End If
        
        If send Then
            .send
        End If
    
    End With
    
End Sub




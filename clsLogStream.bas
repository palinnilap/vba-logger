VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TStream

    fileno As Long
    fname As String
    queue As Collection
    queueLen As Long
    level As LogLevel
    
End Type

Private this As TStream

'------------------------------------------
'           Create
'------------------------------------------

Private Sub Class_Initialize()

    Set this.queue = New Collection

End Sub

Public Sub Create(fname As String, level As LogLevel, queueLen As Long)

    this.fname = fname
    this.level = level
    this.queueLen = queueLen

    this.fileno = FreeFile

End Sub

Private Sub Class_Terminate()

    If this.queue.count > 0 Then
        WriteQueueToFile
    End If
    
End Sub

'------------------------------------------
'           Interface
'------------------------------------------

Public Property Get fname()

    fname = this.fname

End Property

Public Sub SubmitLog(msg, level As LogLevel)

    If level < this.level Then
        Exit Sub
    End If

    this.queue.Add msg
    
    If this.queue.count >= this.queueLen Then
        Call WriteQueueToFile
    End If
    
End Sub

Public Sub ForceWriteToFile()

    Call WriteQueueToFile

End Sub

'------------------------------------------
'           Private Subs
'------------------------------------------

Private Sub WriteQueueToFile()

    Open this.fname For Append As #this.fileno
    Print #this.fileno, QtoMsg
    Close #this.fileno

    Set this.queue = New Collection
    
End Sub

Private Function QtoMsg() As String

    Dim i As Long
    Dim count As Long
    Dim msg As String
    
    count = this.queue.count
    For i = 1 To count
        
        msg = msg & this.queue(i)
        
        If i <> count Then
            msg = msg & vbNewLine
        End If
        
    Next

    QtoMsg = msg
    
End Function


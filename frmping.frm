VERSION 5.00
Begin VB.Form frmping 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ping a IP"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "frmping.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3990
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2048
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   2760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ping"
      Height          =   255
      Left            =   1088
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   128
      TabIndex        =   3
      Top             =   360
      Width           =   3735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ping Continuously"
      Height          =   255
      Left            =   2288
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   368
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "IP:"
      Height          =   255
      Left            =   128
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Check1.Value = 1 Then
Timer1.Enabled = True
Command2.Enabled = True
Exit Sub
Else
End If


   List1.Clear
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Integer
   Dim i As String
   i = Text1.Text
   
   Call Ping(i, ECHO)
   
  'display the results from the ECHO structure
   List1.AddItem "Status: " & vbTab & vbTab & GetStatusCode(ECHO.status)
   List1.AddItem "Address: " & vbTab & vbTab & ECHO.Address
   List1.AddItem "Round Trip Time: " & vbTab & ECHO.RoundTripTime & " ms"
   List1.AddItem "Data Size: " & vbTab & ECHO.DataSize & " bytes"
   
   If Left$(ECHO.Data, 1) <> Chr$(0) Then
      pos = InStr(ECHO.Data, Chr$(0))
      List1.AddItem Left$(ECHO.Data, pos - 1)
   End If

   List1.AddItem "Data Pointer: " & vbTab & ECHO.DataPointer
   List1.AddItem ""
End Sub

Private Sub Command2_Click()
Check1.Value = 0
Timer1.Enabled = False
Command2.Enabled = False
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

Private Sub Timer1_Timer()
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Integer
   Dim i As String
   i = Text1.Text
   
   Call Ping(i, ECHO)
   
  'display the results from the ECHO structure
   List1.AddItem "Status: " & vbTab & vbTab & GetStatusCode(ECHO.status)
   List1.AddItem "Address: " & vbTab & vbTab & ECHO.Address
   List1.AddItem "Round Trip Time: " & vbTab & ECHO.RoundTripTime & " ms"
   List1.AddItem "Data Size: " & vbTab & ECHO.DataSize & " bytes"
   
   If Left$(ECHO.Data, 1) <> Chr$(0) Then
      pos = InStr(ECHO.Data, Chr$(0))
      List1.AddItem Left$(ECHO.Data, pos - 1)
   End If

   List1.AddItem "Data Pointer: " & vbTab & ECHO.DataPointer
   List1.AddItem ""

End Sub

VERSION 5.00
Begin VB.Form frmcomputer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administer Computer"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmcomputer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6255
   Begin VB.CommandButton Command3 
      Caption         =   "Services"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Processes"
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4560
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label11 
      Caption         =   "Processor Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label9 
      Caption         =   "OS Version:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "OS:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Organization:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Owner:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Computer Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmcomputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmprocesses.Show
frmprocesses.Text1.Text = Label2.Caption
End Sub

Private Sub Command3_Click()
frmservices.Show
frmservices.Server.Text = Label2.Caption
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
MousePointer = vbHourglass
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text


Dim computer As IADsComputer
Dim computername As String
Dim computerdomain As String
computerdomain = frmdomainlogin.Combo1.Text
computername = Label2.Caption

If frmdomainlogin.Check1.Value = 1 Then
Set computer = GetObject("WinNT://" & computerdomain & "/" & computername & ",computer")
Else
Set dso = GetObject("WinNT:")
Set computer = dso.OpenDSObject("WinNT://" & computerdomain & "/" & computername & ",computer", username, password, 0)
End If
Dim retval As String
retval = computer.Owner
Label4.Caption = retval

retval = computer.Division
Label6.Caption = retval

retval = computer.OperatingSystem
Label8.Caption = retval

retval = computer.OperatingSystemVersion
Label10.Caption = retval

retval = computer.Processor
Label12.Caption = retval

Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
MousePointer = 0
Timer1.Enabled = False
End Sub

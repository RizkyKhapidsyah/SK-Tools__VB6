VERSION 5.00
Begin VB.Form frmadmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Admin"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmadmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5145
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   3000
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Bulk Administration"
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   1800
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   360
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Administer Selected Computers"
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Administer Selected Group"
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Administer Selected User"
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Computers:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Groups:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Users:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Totals:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Computers"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Groups"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Users"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Command1.Caption = "Administer All Users"
Command2.Caption = "Administer All Groups"
Command4.Enabled = False
Else
Command1.Caption = "Administer Selected User"
Command2.Caption = "Administer Selected Group"
Command4.Enabled = True
End If
End Sub

Private Sub Command1_Click()
If Check1.Value = 0 Then
    If Combo1.Text = "" Then
    MsgBox "Please select a user to administer"
    Exit Sub
    Else
    frmuser.Show
    frmuser.Label7.Caption = Combo1.Text
    DoEvents
    frmuser.Timer1.Enabled = True
    End If
Else
frmbulkuser.Show
frmbulkuser.Timer12.Enabled = True
End If
End Sub

Private Sub Command2_Click()
If Check1.Value = 0 Then
frmgroup.Show
frmgroup.Label3.Caption = Combo2.Text
frmgroup.Timer1.Enabled = True
Else
frmbulkgroup.Show
frmbulkgroup.Timerload.Enabled = True
End If
End Sub

Private Sub Command3_Click()
If Check1.Value = 0 Then
frmOUS.Show
Else
frmBulkOUS.Show
End If
End Sub

Private Sub Command4_Click()
If Combo4.Text = "" Then
MsgBox "Please select a computer to administer"
Exit Sub
Else
frmcomputer.Show
frmcomputer.Label2.Caption = Combo4.Text
frmcomputer.Timer1.Enabled = True
End If
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
frmadmin.Top = 0
frmadmin.Left = 0
End Sub

Private Sub Timer1_Timer()
Label10.Caption = Combo1.ListCount
Label11.Caption = Combo2.ListCount
Label13.Caption = Combo4.ListCount
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
MousePointer = vbHourglass
frmadmin.Combo1.Clear
frmadmin.Combo2.Clear
frmadmin.Combo4.Clear

MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Please Wait Loading Users, Groups and Computers..."

Dim dso As IADsOpenDSObject
username = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim container As IADsContainer
Dim containername As String
containername = frmdomainlogin.Combo1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set container = GetObject("WinNT://" & containername)
Else
Set dso = GetObject("WinNT:")
Set container = dso.OpenDSObject("WinNT://" & DomainName, username, password, 1)
End If

container.Filter = Array("User")
Dim user As IADsUser
For Each user In container
Combo1.AddItem user.Name
Next

container.Filter = Array("Group")
Dim group As IADsGroup
For Each group In container
Combo2.AddItem group.Name
Next

container.Filter = Array("Computer")
Dim computer As IADsComputer
For Each computer In container
Combo4.AddItem computer.Name
Next
DoEvents

Err = 0

DoEvents
MousePointer = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer2.Enabled = False
End Sub

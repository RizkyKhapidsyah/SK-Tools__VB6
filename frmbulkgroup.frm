VERSION 5.00
Begin VB.Form frmbulkgroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bulk Adminstration for Groups"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "frmbulkgroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7515
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6960
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6480
      Top             =   3120
   End
   Begin VB.ListBox List3 
      Height          =   450
      Left            =   6240
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timerload 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Decription on all Groups"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   7335
   End
   Begin VB.Timer timertotal 
      Interval        =   100
      Left            =   120
      Top             =   3120
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5400
      Picture         =   "frmbulkgroup.frx":27A2
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Decription on all Groups"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Groups:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current Groups - Local and Global"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmbulkgroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
List1.ListIndex = 0
List2.Clear
List3.Clear
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If List3.ListCount = List1.ListCount Then
Timer1.Enabled = False
Else
List1.ListIndex = List1.ListIndex + 1
Timer2.Enabled = True
Timer1.Enabled = False
End If
Err = 0
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
MousePointer = vbHourglass
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text


Dim group As IADsGroup
Dim groupname As String
Dim groupdomain As String
groupname = List1.Text
groupdomain = frmdomainlogin.Combo1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set group = GetObject("WinNT://" & groupdomain & "/" & groupname & ",group")
Else
Set dso = GetObject("WinNT:")
Set group = dso.OpenDSObject("WinNT://" & groupdomain & "/" & groupname & ",group", username, password, 1)
End If
Dim retval As String
retval = group.Description
List2.AddItem "Group Name: " & List1.Text
List2.AddItem "Decription: " & retval
List2.AddItem ""

List3.AddItem List1.Text

MousePointer = 0
Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer1.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timerload_Timer()
On Error Resume Next
MousePointer = vbHourglass
List1.Clear
List2.Clear
List3.Clear

MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Please Wait Loading Users, Groups, OUS, and Computers..."

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

container.Filter = Array("Group")
Dim group As IADsGroup
For Each group In container
List1.AddItem group.Name
Next


Err = 0

DoEvents
MousePointer = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timerload.Enabled = False

End Sub

Private Sub timertotal_Timer()
Label2.Caption = "Total Groups: " & List1.ListCount
End Sub

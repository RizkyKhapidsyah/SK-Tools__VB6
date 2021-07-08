VERSION 5.00
Begin VB.Form frmgroupdesc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Description"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmgroupdesc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5280
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4800
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1125
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Description for"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmgroupdesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
MousePointer = vbHourglass

Dim dso As IADsOpenDSObject
username = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim group As IADsGroup
Dim groupname As String
Dim groupdomain As String
groupname = Label2.Caption
groupdomain = frmdomainlogin.Combo1.Text

If Check1.Value = 0 Then
Else
If frmdomainlogin.Check1.Value = 1 Then
Set group = GetObject("WinNT://" & groupdomain & "/" & groupname & ",group")
Else
Set dso = GetObject("WinNT:")
Set group = dso.OpenDSObject("WinNT://" & groupdomain & "/" & groupname & ",group", username, password, 1)
End If
group.Description = Text1.Text
group.SetInfo
End If
MousePointer = 0
Err = 0
Unload Me
End Sub

Private Sub Timer1_Timer()
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
groupname = Label2.Caption
groupdomain = frmdomainlogin.Combo1.Text

If Check1.Value = 0 Then
If frmdomainlogin.Check1.Value = 1 Then
Set group = GetObject("WinNT://" & groupdomain & "/" & groupname & ",group")
Else
Set dso = GetObject("WinNT:")
Set group = dso.OpenDSObject("WinNT://" & groupdomain & "/" & groupname & ",group", username, password, 1)
End If

Dim retval As String
retval = group.Description
Text1.Text = retval
Else
Command1.Caption = "Save && Close"
End If
MousePointer = 0
Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer1.Enabled = False
End Sub

VERSION 5.00
Begin VB.Form frmgroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administer A Group"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "frmgroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6270
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2400
      Top             =   480
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Change Group Description"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View Group Description"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   2175
   End
   Begin VB.ListBox List2 
      Height          =   3570
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove User -->"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<-- Add User"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "All Users"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Members of"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmgroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
If List2.Text = "" Then
MsgBox "Please select a group to add"
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Exit Sub
Else
MousePointer = vbHourglass
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim group As IADsGroup
Dim groupname As String
Dim groupdomain As String
Dim user As IADsUser
Dim username As String
Dim userdomain As String

groupname = Label3.Caption
groupdomain = frmdomainlogin.Combo1.Text
username = List2.Text
userdomain = frmdomainlogin.Combo1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Set group = GetObject("WinNT://" & groupdomain & "/" & groupname & ",group")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & userdomain & "/" & username & ",user", username2, password, 1)
Set group = dso.OpenDSObject("WinNT://" & groupdomain & "/" & groupname & ",group", username2, password, 1)
End If

group.Add (user.ADsPath)
group.SetInfo
List1.AddItem List2.Text

Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
MousePointer = 0
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
If List1.Text = "" Then
MsgBox "Please select a User to remove"
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Exit Sub
Else
MousePointer = vbHourglass

Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim group As IADsGroup
Dim groupname As String
Dim groupdomain As String
Dim user As IADsUser
Dim username As String
Dim userdomain As String

groupname = Label3.Caption
groupdomain = frmdomainlogin.Combo1.Text
username = List1.Text
userdomain = frmdomainlogin.Combo1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Set group = GetObject("WinNT://" & groupdomain & "/" & groupname & ",group")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & userdomain & "/" & username & ",user", username2, password, 1)
Set group = dso.OpenDSObject("WinNT://" & groupdomain & "/" & groupname & ",group", username2, password, 1)
End If

group.Remove (user.ADsPath)

List1.RemoveItem List1.ListIndex

Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
MousePointer = 0
End If
End Sub

Private Sub Command3_Click()
frmgroupdesc.Show
frmgroupdesc.Label2.Caption = Label3.Caption
frmgroupdesc.Timer1.Enabled = True
End Sub

Private Sub Command4_Click()
frmgroupdesc.Show
frmgroupdesc.Label2.Caption = Label3.Caption
frmgroupdesc.Check1.Value = 1
frmgroupdesc.Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
MousePointer = vbHourglass
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim container As IADsContainer
Dim containername As String
containername = frmdomainlogin.Combo1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set container = GetObject("WinNT://" & containername)
Else
Set dso = GetObject("WinNT:")
Set container = dso.OpenDSObject("WinNT://" & DomainName, username2, password, 1)
End If

container.Filter = Array("User")
Dim user As IADsUser
For Each user In container
List2.AddItem user.Name
Next

Dim group As IADsGroup
Dim groupname As String
Dim groupdomain As String

groupname = Label3.Caption
groupdomain = frmdomainlogin.Combo1.Text
Set group = GetObject("WinNT://" & groupdomain & "/" & groupname & ",group")

For Each member In group.Members
List1.AddItem member.Name
Next

Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
MousePointer = 0
Timer1.Enabled = False
End Sub

VERSION 5.00
Begin VB.Form frmusergroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Groups"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   Icon            =   "frmusergroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7020
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5040
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<- Add"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove ->"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   255
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "User:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Members of:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Groups                     (Local && Global)"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
End
Attribute VB_Name = "frmusergroup"
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

groupname = List2.Text
groupdomain = frmdomainlogin.Combo1.Text
username = Label2.Caption
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
MsgBox "Please select a group to remove"
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

groupname = List1.Text
groupdomain = frmdomainlogin.Combo1.Text
username = Label2.Caption
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

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
MousePointer = vbHourglass
List1.Clear
List2.Clear

Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomain As String
Dim group As IADsGroup
Dim container As IADsContainer
Dim containername As String
containername = frmdomainlogin.Combo1.Text
userdomain = frmdomainlogin.Combo1.Text
username = Label2.Caption

If frmdomainlogin.Check1.Value = 1 Then
Set container = GetObject("WinNT://" & containername)
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
Set container = dso.OpenDSObject("WinNT://" & DomainName, username, password, 1)
End If

For Each group In user.Groups
List1.AddItem group.Name
Next
container.Filter = Array("Group")
For Each group In container
List2.AddItem group.Name
Next

Err = 0
MousePointer = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer1.Enabled = False
End Sub

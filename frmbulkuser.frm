VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbulkuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bulk Administration of Users"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   Icon            =   "frmbulkuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   10155
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   3240
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   3240
   End
   Begin VB.OptionButton Option6 
      Height          =   285
      Left            =   3360
      TabIndex        =   32
      Top             =   3600
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Home Directory"
      Height          =   1335
      Left            =   3720
      TabIndex        =   25
      Top             =   3120
      Width           =   6375
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   30
         Top             =   840
         Width           =   3975
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   29
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Connect"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   360
         Width           =   4815
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Local Path:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "To"
         Height          =   255
         Left            =   1920
         TabIndex        =   31
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2520
      Top             =   2760
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   2400
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   2400
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   1440
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   1440
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   6840
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5280
      Top             =   1440
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   1440
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   840
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   840
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5280
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear for all in List"
      Height          =   255
      Left            =   8400
      TabIndex        =   22
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply to all in List"
      Height          =   255
      Left            =   8400
      TabIndex        =   21
      Top             =   120
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      Height          =   285
      Left            =   3360
      TabIndex        =   20
      Top             =   2400
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Must Change Password At Next Logon"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Password Never Expires"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Account Disabled"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3720
      TabIndex        =   17
      Top             =   2760
      Width           =   2295
   End
   Begin VB.OptionButton Option2 
      Height          =   285
      Left            =   3360
      TabIndex        =   16
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   15
      Top             =   1560
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Top             =   960
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3720
      TabIndex        =   12
      Top             =   960
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4920
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   2895
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "List Users By Group"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   2760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove User from List"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8400
      Picture         =   "frmbulkuser.frx":27A2
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label Label9 
      Caption         =   "Done:"
      Height          =   255
      Left            =   6840
      TabIndex        =   24
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "User profile path:"
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Login Script:"
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "100%"
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "50%"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "0%"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "All Users"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmbulkuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
On Error Resume Next
MousePointer = vbHourglass
List1.Clear
If Combo1.Text = "All Users" Then
Combo1.Clear
Timer12.Enabled = True
Label1.Caption = "All Users"
Else
Label1.Caption = "Members of " & Combo1.Text

Dim dso As IADsOpenDSObject
username = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim group As IADsGroup
Dim groupname As String
Dim groupdomain As String

groupname = Combo1.Text
groupdomain = frmdomainlogin.Combo1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set group = GetObject("WinNT://" & groupdomain & "/" & groupname & ",group")
Else
Set dso = GetObject("WinNT:")
Set group = dso.OpenDSObject("WinNT://" & groupdomain & "/" & groupname & ",group", username, password, 1)
End If

For Each member In group.Members
List1.AddItem member.Name
Next
End If
Err = 0
MousePointer = 0
End Sub

Private Sub Command1_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
Err = 0
End Sub

Private Sub Command2_Click()
List1.ListIndex = 0
ProgressBar1.Max = List1.ListCount
ProgressBar1.Value = 0
List2.Clear

DoEvents

If Option1.Value = True Then
    If Text1.Text = "" Then
    MsgBox "You must enter a login script"
    Exit Sub
    Else
    End If
Timer3.Enabled = True
Else
End If

If Option2.Value = True Then
    If Text2.Text = "" Then
    MsgBox "You must enter a profile path"
    Exit Sub
    Else
    End If
Timer7.Enabled = True
Else
End If

If Option3.Value = True Then
Timer11.Enabled = True
Else
End If

If Option6.Value = True Then
Timer14.Enabled = True
Else
End If
End Sub

Private Sub Command3_Click()
List1.ListIndex = 0
ProgressBar1.Max = List1.ListCount
ProgressBar1.Value = 0
List2.Clear
DoEvents
If Option1.Value = True Then
Text1.Text = ""
Timer5.Enabled = True
Else
End If
If Option2.Value = True Then
Text2.Text = ""
Timer9.Enabled = True
Else
End If

If Option6.Value = True Then
Option4.Value = True
Text3.Text = ""
Text4.Text = ""
Combo2.Text = ""
Timer14.Enabled = True
Else
End If
End Sub

Private Sub Form_Load()
frmadmin.Top = 0
frmadmin.Left = 0
Combo2.AddItem "D:"
Combo2.AddItem "E:"
Combo2.AddItem "F:"
Combo2.AddItem "G:"
Combo2.AddItem "H:"
Combo2.AddItem "I:"
Combo2.AddItem "J:"
Combo2.AddItem "K:"
Combo2.AddItem "L:"
Combo2.AddItem "M:"
Combo2.AddItem "N:"
Combo2.AddItem "O:"
Combo2.AddItem "P:"
Combo2.AddItem "Q:"
Combo2.AddItem "R:"
Combo2.AddItem "S:"
Combo2.AddItem "T:"
Combo2.AddItem "U:"
Combo2.AddItem "V:"
Combo2.AddItem "W:"
Combo2.AddItem "X:"
Combo2.AddItem "Y:"
Combo2.AddItem "Z:"
End Sub

Private Sub Option1_Click()
Text1.Enabled = True
Text2.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Command3.Enabled = True
Option4.Enabled = False
Option5.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Combo2.Enabled = False
End Sub

Private Sub Option2_Click()
Text1.Enabled = False
Text2.Enabled = True
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Command3.Enabled = True
Option4.Enabled = False
Option5.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Combo2.Enabled = False
End Sub

Private Sub Option3_Click()
Text1.Enabled = False
Text2.Enabled = False
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Command3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Combo2.Enabled = False
End Sub

Private Sub Option6_Click()
Text1.Enabled = False
Text2.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Command3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Combo2.Enabled = True
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command2_Click
 DoEvents
 End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command2_Click
 DoEvents
 End If
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "Total Users: " & List1.ListCount
End Sub

Private Sub Timer10_Timer()
On Error Resume Next
If List2.ListCount = List1.ListCount Then
Timer10.Enabled = False
Else
List1.ListIndex = List1.ListIndex + 1
Timer11.Enabled = True
Timer10.Enabled = False
End If
Err = 0
End Sub

Private Sub Timer11_Timer()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomain As String
userdomain = frmdomainlogin.Combo1.Text
username = List1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If

Dim passwordexpired As Integer
Dim Flags As Long
Dim newvalue As Boolean

If Check1.Value = 1 Then
user.Put "PasswordExpired", 1
user.SetInfo
Else
user.Put "PasswordExpired", 0
user.SetInfo
End If


If Check2.Value = 1 Then
Flags = user.Get("userflags")
user.Put "userflags", Flags Or &H10000
user.SetInfo
Else
Flags = user.Get("userflags")
user.Put "userflags", Flags Xor &H10000
user.SetInfo
End If

If Check3.Value = 1 Then
newvalue = True
user.AccountDisabled = newvalue
user.SetInfo
Else
newvalue = False
user.AccountDisabled = newvalue
user.SetInfo
End If

List2.AddItem List1.Text
ProgressBar1.Value = ProgressBar1.Value + 1


Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer10.Enabled = True
Timer11.Enabled = False
End Sub

Private Sub Timer12_Timer()
On Error Resume Next
MousePointer = vbHourglass
Combo1.AddItem "All Users"
List1.Clear
List2.Clear
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Please Wait Loading Users..."

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
List1.AddItem user.Name
Next

container.Filter = Array("Group")
Dim group As IADsGroup
For Each group In container
Combo1.AddItem group.Name
Next

DoEvents

Err = 0

MousePointer = 0

DoEvents
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer12.Enabled = False
End Sub

Private Sub Timer13_Timer()
On Error Resume Next
If List2.ListCount = List1.ListCount Then
Timer13.Enabled = False
Else
List1.ListIndex = List1.ListIndex + 1
Timer14.Enabled = True
Timer13.Enabled = False
End If
Err = 0
End Sub

Private Sub Timer14_Timer()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomain As String
userdomain = frmdomainlogin.Combo1.Text
username = List1.Text
If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If

Dim newvalue As String

If Option4.Value = True Then
    If Text3.Text = "" Then
    newvalue = ""
    Call user.Put("HomeDirDrive", "")
    user.HomeDirectory = newvalue
    user.SetInfo
    Else
    newvalue = Text3.Text
    Call user.Put("HomeDirDrive", "")
    user.HomeDirectory = newvalue
    user.SetInfo
    End If
    Else
    End If

If Option5.Value = True Then
    newvalue2 = Combo2.Text
    Call user.Put("HomeDirDrive", newvalue2)
    user.HomeDirectory = Text4.Text
    user.SetInfo
    Else
    End If
    
List2.AddItem List1.Text
ProgressBar1.Value = ProgressBar1.Value + 1

Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer13.Enabled = True
Timer14.Enabled = False

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If List2.ListCount = List1.ListCount Then
Timer2.Enabled = False
Else
List1.ListIndex = List1.ListIndex + 1
Timer3.Enabled = True
Timer2.Enabled = False
End If
Err = 0
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomain As String
userdomain = frmdomainlogin.Combo1.Text
username = List1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If
Dim newvalue As String
newvalue = Text1.Text
user.LoginScript = newvalue
user.SetInfo

List2.AddItem List1.Text
ProgressBar1.Value = ProgressBar1.Value + 1

Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer2.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
If List2.ListCount = List1.ListCount Then
Timer4.Enabled = False
Else
List1.ListIndex = List1.ListIndex + 1
Timer5.Enabled = True
Timer4.Enabled = False
End If
Err = 0
End Sub

Private Sub Timer5_Timer()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomain As String
userdomain = frmdomainlogin.Combo1.Text
username = List1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If

Dim newvalue As String
newvalue = ""
user.LoginScript = newvalue
user.SetInfo

List2.AddItem List1.Text
ProgressBar1.Value = ProgressBar1.Value + 1

Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer4.Enabled = True
Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
On Error Resume Next
If List2.ListCount = List1.ListCount Then
Timer6.Enabled = False
Else
List1.ListIndex = List1.ListIndex + 1
Timer7.Enabled = True
Timer6.Enabled = False
End If
Err = 0
End Sub

Private Sub Timer7_Timer()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomain As String
userdomain = frmdomainlogin.Combo1.Text
username = List1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If

Dim newvalue As String
newvalue = Text2.Text
user.Profile = newvalue
user.SetInfo

List2.AddItem List1.Text
ProgressBar1.Value = ProgressBar1.Value + 1

Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer6.Enabled = True
Timer7.Enabled = False
End Sub

Private Sub Timer8_Timer()
On Error Resume Next
If List2.ListCount = List1.ListCount Then
Timer8.Enabled = False
Else
List1.ListIndex = List1.ListIndex + 1
Timer9.Enabled = True
Timer8.Enabled = False
End If
Err = 0
End Sub

Private Sub Timer9_Timer()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomain As String
userdomain = frmdomainlogin.Combo1.Text
username = List1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If

Dim newvalue As String
newvalue = ""
user.Profile = newvalue
user.SetInfo

List2.AddItem List1.Text
ProgressBar1.Value = ProgressBar1.Value + 1

Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer8.Enabled = True
Timer9.Enabled = False
End Sub

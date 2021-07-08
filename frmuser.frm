VERSION 5.00
Begin VB.Form frmuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administer User"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   Icon            =   "frmuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7185
   Begin VB.CommandButton Command10 
      Caption         =   "Add User"
      Height          =   255
      Left            =   3600
      TabIndex        =   26
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6720
      Top             =   0
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Delete User"
      Height          =   255
      Left            =   3600
      TabIndex        =   25
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   24
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Close"
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save && Close"
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Groups"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Profile"
      Height          =   495
      Left            =   1560
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Account"
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Rename Username"
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "User Must Change Password at Next Logon"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "User Cannot Change Password"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Password Nevers Expires"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   3615
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Account Disabled"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Account Locked Out"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6120
      Picture         =   "frmuser.frx":0BC2
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label Label8 
      Caption         =   "Last Logoff"
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Last Login"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Bad Login Count"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Full Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "UserName:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmusergroup.Show
frmusergroup.Label2.Caption = Label7.Caption
frmusergroup.Timer1.Enabled = True
End Sub

Private Sub Command10_Click()
frmadduser.Show
Unload Me
End Sub

Private Sub Command2_Click()
frmuserprofile.Show
frmuserprofile.Label2.Caption = Label7.Caption
frmuserprofile.Timer2.Enabled = True
End Sub
Private Sub Command5_Click()
frmuseraccount.Show
frmuseraccount.Label2.Caption = Label7.Caption
frmuseraccount.Timer1.Enabled = True
End Sub

Private Sub Command6_Click()
frmrenameuser.Show
frmrenameuser.Text1.Text = frmuser.Label7.Caption
frmrenameuser.Label3.Caption = frmdomainlogin.Combo1.Text
DoEvents
End Sub

Private Sub Command7_Click()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomian As String

userdomain = frmdomainlogin.Combo1.Text
username = Label7.Caption
If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If

Dim passwordexpired As Integer
Dim Flags As Long
Dim newvalue As Boolean
Dim newfullname As String
Dim newdescription As String

If Text1.Text = "" Then
newfullname = ""
user.FullName = newfullname
user.SetInfo
Else
newfullname = Text1.Text
user.FullName = newfullname
user.SetInfo
End If

If Text2.Text = "" Then
newdescription = ""
user.Description = newdescription
user.SetInfo
Else
newdescription = Text2.Text
user.Description = newdescription
user.SetInfo
End If

If Check1.Value = 1 Then
user.Put "PasswordExpired", 1
user.SetInfo
Else
user.Put "PasswordExpired", 0
user.SetInfo
End If


If Check3.Value = 1 Then
Flags = user.Get("userflags")
user.Put "userflags", Flags Or &H10000
user.SetInfo
Else
Flags = user.Get("userflags")
user.Put "userflags", Flags Xor &H10000
user.SetInfo
End If

If Check4.Value = 1 Then
newvalue = True
user.AccountDisabled = newvalue
user.SetInfo
Else
newvalue = False
user.AccountDisabled = newvalue
user.SetInfo
End If

Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Unload Me
End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Command9_Click()
On Error Resume Next
mouepointer = 11

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
Set container = dso.OpenDSObject("WinNT://" & containername, username2, password, 1)

End If

Dim usertoremove As String
usertoremove = Label7.Caption
Call container.Delete("User", usertoremove)

Err = 0
MousePointer = 0
frmadmin.Timer2.Enabled = True
Unload Me
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
MousePointer = vbHourglass
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0

Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomian As String

userdomain = frmdomainlogin.Combo1.Text
username = Label7.Caption

If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If

Dim retval As String
retval = user.FullName
Text1.Text = retval

retval = user.Description
Text2.Text = retval

Dim Flags As Long
Flags = user.Get("userflags")
If (Flags And &H10000) <> 0 Then
Check3.Value = 1
End If

If (Flags And &H10) <> 0 Then
Check5.Value = 1
End If

If (Flags And &H2) <> 0 Then
Check4.Value = 1
End If

If (Flags And &H40) <> 0 Then
Check2.Value = 1
End If

Dim passwordexpired As Integer
passwordexpired = user.Get("passwordexpired")
If passwordexpired = 1 Then
Check1.Value = 1
End If

Dim retval2 As Integer
retval2 = user.BadLoginCount
Text4.Text = retval2

Dim retval3 As Date
retval3 = user.LastLogin
Text5.Text = retval3

Dim retval4 As Date
retval4 = user.LastLogoff
Text7.Text = retval4

Text3.Text = "**********"
Err = 0
MousePointer = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer1.Enabled = False
End Sub

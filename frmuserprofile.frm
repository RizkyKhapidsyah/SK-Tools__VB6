VERSION 5.00
Begin VB.Form frmuserprofile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frmuserprofile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7710
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4320
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Profiles"
      Height          =   1215
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   6375
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "User Profile Path:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Logon Scipt Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Home Directory"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   6375
      Begin VB.OptionButton Option1 
         Caption         =   "Local Path:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   4815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Connect"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.Label Label1 
      Caption         =   "User:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmuserprofile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
MousePointer = vbHourglass
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomain As String
userdomain = frmdomainlogin.Combo1.Text
username = Label2.Caption
If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If

Dim newvalue As String

If Text2.Text = "" Then
newvalue = ""
user.LoginScript = newvalue
user.SetInfo
Else
newvalue = Text2.Text
user.LoginScript = newvalue
user.SetInfo
End If

If Text1.Text = "" Then
newvalue = ""
user.Profile = newvalue
user.SetInfo
Else
newvalue = Text1.Text
user.Profile = newvalue
user.SetInfo
End If

If Option1.Value = True Then
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

If Option2.Value = True Then
    newvalue2 = Combo1.Text
    Call user.Put("HomeDirDrive", newvalue2)
    user.HomeDirectory = Text4.Text
    user.SetInfo
    Else
    End If
    

Err = 0
MousePointer = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Combo1.AddItem "D:"
Combo1.AddItem "E:"
Combo1.AddItem "F:"
Combo1.AddItem "G:"
Combo1.AddItem "H:"
Combo1.AddItem "I:"
Combo1.AddItem "J:"
Combo1.AddItem "K:"
Combo1.AddItem "L:"
Combo1.AddItem "M:"
Combo1.AddItem "N:"
Combo1.AddItem "O:"
Combo1.AddItem "P:"
Combo1.AddItem "Q:"
Combo1.AddItem "R:"
Combo1.AddItem "S:"
Combo1.AddItem "T:"
Combo1.AddItem "U:"
Combo1.AddItem "V:"
Combo1.AddItem "W:"
Combo1.AddItem "X:"
Combo1.AddItem "Y:"
Combo1.AddItem "Z:"
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
MousePointer = vbHourglass
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomain As String
userdomain = frmdomainlogin.Combo1.Text
username = Label2.Caption
If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If

Dim retval As String
retval = user.LoginScript
Text2.Text = retval
DoEvents

retval = user.Profile
Text1.Text = retval
DoEvents

retval = user.Get("homedirdrive")
Combo1.Text = retval

If Combo1.Text = "" Then
retval = user.HomeDirectory
Text3.Text = retval
Else
retval = user.HomeDirectory
Text4.Text = retval
End If

MousePointer = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer2.Enabled = False
End Sub

VERSION 5.00
Begin VB.Form frmuseraccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmuseraccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6555
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Account Type"
      Height          =   1575
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   3615
      Begin VB.OptionButton Option4 
         Caption         =   "Local Account"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Global Account"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "for regular user accounts in this domain"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "for users from untrusted domains"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1200
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account Expires"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "01/01/2000"
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "End of"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Never"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "User:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmuseraccount"
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

Dim date1 As Date
Dim Flags As Long

If Option1.Value = True Then
date1 = #12:00:00 AM#
user.AccountExpirationDate = date1
user.SetInfo
Else
End If

If Option2.Value = True Then
date1 = Text1.Text
user.AccountExpirationDate = date1
user.SetInfo
Else
End If

If Option3.Value = True Then
Flags = user.Get("userflags")
user.Put "userflags", Flags Xor &H100
user.SetInfo
Flags = user.Get("userflags")
user.Put "userflags", Flags Xor &H200
user.SetInfo
Else
End If

If Option4.Value = True Then
Flags = user.Get("userflags")
user.Put "userflags", Flags Xor &H200
user.SetInfo
Flags = user.Get("userflags")
user.Put "userflags", Flags Xor &H100
user.SetInfo
Else
End If

Err = 0
MousePointer = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Unload Me
End Sub

Private Sub Command2_Click()
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

Dim Flags As Long
Flags = user.Get("userflags")
If (Flags And &H100) <> 0 Then
Option4.Value = True
Else
Option3.Value = True
End If

Dim date1 As Date
date1 = user.AccountExpirationDate
Text1.Text = date1

If Text1.Text = "12:00:00 AM" Then
Option1.Value = True
Text1.Text = "01/01/2000"
Else
Option2.Value = True
End If

Err = 0
MousePointer = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Timer1.Enabled = False
End Sub

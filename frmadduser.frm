VERSION 5.00
Begin VB.Form frmadduser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add User Account"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3960
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create User"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "After you have made the new user you will have the option of editing the user info and account properities."
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please enter the New User Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmadduser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
MousePointer = vbHourglass
If Text1.Text = "" Then
MsgBox "Please type in a user name to create"
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Exit Sub
Else
End If

Dim dso As IADsOpenDSObject
username = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim container As IADsContainer
Dim containername As String
Dim user As IADsUser
Dim newuser As String
containername = frmdomainlogin.Combo1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set container = GetObject("WinNT://" & containername)
Else
Set dso = GetObject("WinNT:")
Set container = dso.OpenDSObject("WinNT://" & containername, username, password, 0)
End If

newuser = Text1.Text
Set user = container.Create("User", newuser)
user.SetInfo

Command2.Enabled = True
Command1.Enabled = False
Err = 0
MousePointer = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End Sub

Private Sub Command2_Click()
frmuser.Show
frmuser.Label7.Caption = Text1.Text
frmuser.Timer1.Enabled = True
DoEvents
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

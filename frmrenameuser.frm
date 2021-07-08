VERSION 5.00
Begin VB.Form frmrenameuser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rename User Account"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   3495
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   405
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "New User Account Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Old User Account Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmrenameuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
If Text2.Text = "" Then
MsgBox "You must specify a new account name"
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Exit Sub
Else
End If
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = frmdomainlogin.Combo1.Text

Dim container As IADsContainer
Dim containername As String
Dim oldname As String
Dim user As IADsUser
Dim newuser As IADsUser
Dim newname As String

oldname = Text1.Text
newname = Text2.Text
containername = Label3.Caption

If frmdomainlogin.Check1.Value = 1 Then
Set container = GetObject("WinNT://" & containername)
Set user = GetObject("WinNT://" & containername & "/" & oldname & ",user")
Else
Set dso = GetObject("WinNT:")
Set container = dso.OpenDSObject("WinNT://" & containername, username2, password, 1)
Set user = dso.OpenDSObject("WinNT://" & containername & "/" & oldname & ",user", username2, password, 1)
End If

Set newuser = container.MoveHere(user.ADsPath, newname)
Set user = Nothing
frmuser.Label7.Caption = Text2.Text
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

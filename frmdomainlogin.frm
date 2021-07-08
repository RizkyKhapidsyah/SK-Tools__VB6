VERSION 5.00
Begin VB.Form frmdomainlogin 
   BorderStyle     =   0  'None
   ClientHeight    =   4665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Tip:"
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   3135
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   $"frmdomainlogin.frx":0000
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Computer"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Domain"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Administrator Credentials "
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4215
      Begin VB.CheckBox Check1 
         Caption         =   "Use Current Credentials"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Text            =   "Administrator"
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblAdminPassword 
         Caption         =   "Pass&word:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblAdminName 
         Caption         =   "Use&r Name:"
         Height          =   255
         Left            =   225
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Domain/Computer Name:"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmdomainlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If (Check1.Value = 1) Then
  bCheck = False
 Else
  bCheck = True
 End If
 
 'Now set the controls based on the credential mode
 Text1.Enabled = bCheck
 lblAdminName.Enabled = bCheck
 Text2.Enabled = bCheck
 lblAdminPassword.Enabled = bCheck
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

Private Sub Command1_Click()
frmdomainlogin.Hide
DoEvents
frmadmin.Show
frmadmin.Timer2.Enabled = True
End Sub

Private Sub Form_Load()
Check1.Value = 1
Check1_Click
frmdomainlogin.Combo1.AddItem MDIFrmmain.Winsock1.LocalHostName
Dim namespace As IADsContainer
Dim domain As IADs
 'Loads Combo box1 with all the current domains
Set namespace = GetObject("WinNT:")

For Each domain In namespace
frmdomainlogin.Combo1.AddItem domain.Name
Next
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

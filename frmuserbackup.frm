VERSION 5.00
Begin VB.Form frmuserbackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Migration/Backup"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmuserbackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4680
   Begin VB.CommandButton Command3 
      Caption         =   "View/Clear/Modify Stored User Information"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   4800
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore User Information"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   4440
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Store User Information"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Note:  This doen't save the groups or sercurity settings for the users.  This is utilitiy is for making adding users much easier."
      Height          =   855
      Left            =   1073
      TabIndex        =   14
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label12 
      Caption         =   "Account Type"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "Account Expires"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Home Directory"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Login Script"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Profile Path"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Account Disabled"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Password Never Expires"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "User Must Change Password"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Discription"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmuserbackup.frx":0BC2
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmuserbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmuserdata.Show
End Sub

Private Sub Command2_Click()
frmuserrestoredata.Show
End Sub

Private Sub Command3_Click()
frmuserdataview.Show
End Sub

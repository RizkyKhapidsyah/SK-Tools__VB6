VERSION 5.00
Begin VB.Form frmsplash 
   BorderStyle     =   0  'None
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3960
      Top             =   240
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   0
      Picture         =   "frmsplash.frx":0000
      Top             =   0
      Width           =   4755
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
MDIFrmmain.Show
Timer1.Enabled = False
Unload Me
End Sub

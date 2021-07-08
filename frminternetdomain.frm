VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frminternetdomain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internet Domain Name Lookup"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frminternetdomain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6180
   Begin VB.TextBox txtResponse 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   23
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   360
      Width           =   6135
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H80000014&
      Height          =   270
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "lookup"
      Height          =   285
      Left            =   5160
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5400
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "www."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "frminternetdomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    txtSearch = ""
    txtResponse = ""
End Sub

Private Sub Command4_Click()
   MousePointer = vbHourglass
   MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
   txtResponse = ""
   Winsock1.Close
   Winsock1.LocalPort = 0
   If Right(txtSearch, 3) = ".tr" Then
      Winsock1.Connect "whois.metu.edu.tr", 43
   Else
      Winsock1.Connect "rs.internic.net", 43
   End If
   MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command4_Click
 DoEvents
 End If
End Sub

Private Sub Winsock1_Connect()
    Winsock1.SendData txtSearch & vbCrLf
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String

    On Error Resume Next

    Winsock1.GetData strData
    strData = Replace(strData, Chr$(10), vbCrLf)
    txtResponse = txtResponse & strData
    MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub


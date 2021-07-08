VERSION 5.00
Begin VB.Form frmremotecommand 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Command"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmremotecommand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6540
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4403
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   0
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   83
      TabIndex        =   7
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Do It"
      Height          =   255
      Left            =   3683
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   83
      TabIndex        =   4
      Text            =   "C:\WINNT\Notepad.exe"
      Top             =   840
      Width           =   4815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remote Restart"
      Height          =   255
      Left            =   83
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remote Shut Down"
      Height          =   255
      Left            =   83
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Extra:"
      Height          =   255
      Left            =   83
      TabIndex        =   9
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "NOTE:  You must have the RemoteShutdown privilege to successfully invoke the Reboot or Shut Down method."
      Height          =   855
      Left            =   1883
      TabIndex        =   8
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Please type the computer name:"
      Height          =   255
      Left            =   83
      TabIndex        =   6
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Command to run on remote computer:"
      Height          =   255
      Left            =   83
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frmremotecommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim computername As String
computername = Text2.Text
Set OpSysSet = GetObject("winmgmts:{(Debug,RemoteShutdown)}//" & computername & "/root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")

For Each OpSys In OpSysSet
    Call OpSys.Shutdown
Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End Sub

Private Sub Command3_Click()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim computername As String
computername = Text2.Text
Set OpSysSet = GetObject("winmgmts:{(RemoteShutdown)}//" & computername & "/root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")

For Each OpSys In OpSysSet
    Call OpSys.Reboot
Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim servername As String
Dim Command2 As String
Dim username2 As String
Dim password As String
username2 = txtAdminName
password = txtAdminPassw
Command2 = Text1.Text
servername = Text2.Text

If frmdomainlogin.Check1.Value = 1 Then
Set Process = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & servername & "\root\cimv2:Win32_Process")
Else
Set Process = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & servername & "\root\cimv2:Win32_Process")
End If
result = Process.Create(Text1.Text, Null, Null, processid)

If Err <> 0 Then
    Text3.Text = "Error: " & Err.Description & " 0x" & Hex(Err.Number)
    Else
    Text3.Text = "Successful"
End If
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command5_Click
 DoEvents
 End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command5_Click
 DoEvents
 End If
End Sub


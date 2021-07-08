VERSION 5.00
Begin VB.Form frmresolve 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resolve a Host to a IP"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frmresolve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   6180
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   863
      TabIndex        =   3
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   3863
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   4703
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Do it"
      Height          =   255
      Left            =   143
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Host Name/Computer Name"
      Height          =   255
      Left            =   1103
      TabIndex        =   6
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Successful?"
      Height          =   255
      Left            =   3623
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   4943
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "frmresolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
   hName As Long
   hAliases As Long
   hAddrType As Integer
   hLength As Integer
   hAddrList As Long
End Type

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADescription_Len) As Byte
   szSystemStatus(0 To WSASYS_Status_Len) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpszVendorInfo As Long
End Type

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)


Function HiByte(ByVal wParam As Integer)
   
   HiByte = wParam \ &H100 And &HFF&
   
End Function

Function LoByte(ByVal wParam As Integer)
   
   LoByte = wParam And &HFF&
   
End Function

Sub SocketsInitialize()
   
   Dim WSAD As WSADATA
   Dim iReturn As Integer
   Dim sLowByte As String, sHighByte As String, sMsg As String
   
   iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
   
   If iReturn <> 0 Then
      MsgBox "Winsock.dll is not responding."
      End
   End If
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      sHighByte = Trim$(Str$(HiByte(WSAD.wVersion)))
      sLowByte = Trim$(Str$(LoByte(WSAD.wVersion)))
      sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
      sMsg = sMsg & " is not supported by winsock.dll "
      MsgBox sMsg
      End
   End If
   
   If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
      sMsg = "This application requires a minimum of "
      sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
      MsgBox sMsg
      End
   End If
   
End Sub

Sub SocketsCleanup()
   Dim lReturn As Long
   
   lReturn = WSACleanup()
   
   If lReturn <> 0 Then
      MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
      End
   End If
   
End Sub

Private Sub Command4_Click()
   On Error Resume Next
   MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
   Dim hostent_addr As Long
   Dim host As HOSTENT
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim i As Integer
   Dim ip_address As String
   
If Text1.Text = "" Then
   Else
   hostent_addr = gethostbyname(Text1)

If hostent_addr = 0 Then
      Text11.Text = "NO"
      Text21.Text = "0"
      MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
      Exit Sub
    Else
End If
   
   RtlMoveMemory host, hostent_addr, LenB(host)
   RtlMoveMemory hostip_addr, host.hAddrList, 4
   
   ReDim temp_ip_address(1 To host.hLength)
   RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength
   
   For i = 1 To host.hLength
      ip_address = ip_address & temp_ip_address(i) & "."
   Next
   ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
   
  Text21.Text = ip_address
  Text11.Text = "YES"
End If
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command4_Click
 DoEvents
 End If

End Sub

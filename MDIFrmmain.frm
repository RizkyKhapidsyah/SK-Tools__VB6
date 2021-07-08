VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIFrmmain 
   BackColor       =   &H8000000C&
   Caption         =   "CS Tools 2.0"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   -45
   ClientWidth     =   8940
   Icon            =   "MDIFrmmain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6180
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9499
            Text            =   "Status:"
            TextSave        =   "Status:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "9:32 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "8/31/2000"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu menufile 
      Caption         =   "&File"
      Begin VB.Menu c 
         Caption         =   "-"
      End
      Begin VB.Menu menuselect 
         Caption         =   "Select Domain or Computer"
         Shortcut        =   {F12}
      End
      Begin VB.Menu menuline 
         Caption         =   "-"
      End
      Begin VB.Menu menuexit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
   End
   Begin VB.Menu menutools 
      Caption         =   "&Extra Tools"
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu menuremotec 
         Caption         =   "Remote Command"
         Shortcut        =   {F1}
      End
      Begin VB.Menu menuservices 
         Caption         =   "Services"
         Shortcut        =   {F2}
      End
      Begin VB.Menu menuprocesses 
         Caption         =   "Processes"
         Shortcut        =   {F3}
      End
      Begin VB.Menu menushortcutmaker 
         Caption         =   "Shortcut Maker"
         Shortcut        =   {F4}
      End
      Begin VB.Menu menubulkshortcut 
         Caption         =   "Bulk Shortcut Maker"
         Shortcut        =   {F5}
      End
      Begin VB.Menu menuresolve 
         Caption         =   "Resolve a Host to a IP"
         Shortcut        =   {F6}
      End
      Begin VB.Menu menuping 
         Caption         =   "Ping a IP"
         Shortcut        =   {F7}
      End
      Begin VB.Menu menuinternetdomain 
         Caption         =   "Internet Domain Name Lookup"
         Shortcut        =   {F8}
      End
      Begin VB.Menu menureset 
         Caption         =   "Reset Users Password"
         Shortcut        =   {F9}
      End
      Begin VB.Menu menuprintstat 
         Caption         =   "Printer Status && Queue"
         Shortcut        =   {F11}
      End
      Begin VB.Menu menuusermig 
         Caption         =   "User Migration/Backup"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
   End
   Begin VB.Menu menuhelp 
      Caption         =   "&Help"
      Begin VB.Menu z 
         Caption         =   "-"
      End
      Begin VB.Menu menuabout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu menusupport 
         Caption         =   "Help"
         Shortcut        =   ^H
      End
      Begin VB.Menu menuweb 
         Caption         =   "Web Page"
         Shortcut        =   ^W
      End
      Begin VB.Menu x 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "MDIFrmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub MDIForm_Load()
frmdomainlogin.Show
End Sub

Private Sub menuabout_Click()
frmAbout.Show
End Sub

Private Sub menubulkshortcut_Click()
frmbulkshortcut.Show
End Sub

Private Sub menuexit_Click()
End
End Sub

Private Sub menuinternetdomain_Click()
frminternetdomain.Show
End Sub

Private Sub menuping_Click()
frmping.Show
End Sub

Private Sub menuprintstat_Click()
frmprintstat.Show
End Sub

Private Sub menuprocesses_Click()
frmprocesses.Show
End Sub

Private Sub menuremotec_Click()
frmremotecommand.Show
End Sub

Private Sub menureset_Click()
ResetPwd.Show
End Sub

Private Sub menuresolve_Click()
frmresolve.Show
End Sub

Private Sub menuselect_Click()
frmdomainlogin.Show
End Sub

Private Sub menuservices_Click()
frmservices.Show
End Sub

Private Sub menushortcutmaker_Click()
frmshortcut.Show
End Sub

Private Sub menusupport_Click()
On Error Resume Next
Call ShellExecute(hwnd, "Open", App.path & "\Help\Help.htm", "", App.path, 1)
End Sub

Private Sub menuusermig_Click()
frmuserbackup.Show
End Sub

Private Sub menuweb_Click()
On Error Resume Next
Call ShellExecute(hwnd, "Open", "http://www.croftssoftware.com", "", App.path, 1)

End Sub

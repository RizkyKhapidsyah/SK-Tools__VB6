VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmshortcut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shortcut Maker"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmshortcut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5265
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Desktop"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Startup"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Programs"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Start Menu"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Create ShortCut"
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4680
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Command Line  (What do you want the the shortcut to open?)"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "What do you want to name your shortcut?"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label4 
      Caption         =   "Where would you like to place your Shortcut?"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   5055
   End
End
Attribute VB_Name = "frmshortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
   (lpFileOp As SHFILEOPSTRUCT) As Long

' // Shell File Operations

Const FO_MOVE = &H1
Const FO_COPY = &H2
Const FO_DELETE = &H3
Const FO_RENAME = &H4
Const FOF_MULTIDESTFILES = &H1
Const FOF_CONFIRMMOUSE = &H2
Const FOF_SILENT = &H4                      '  don't create progress/report
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings
                                      '  Must be freed using SHFreeNameMappings
Const FOF_ALLOWUNDO = &H40
Const FOF_FILESONLY = &H80                  '  on *.*, do only files - not directories
Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs

Const PO_DELETE = &H13           '  printer is being deleted
Const PO_RENAME = &H14           '  printer is being renamed
Const PO_PORTCHANGE = &H20       '  port this printer connected to is being changed
                                '  if this id is set, the strings received by
                                '  the copyhook are a doubly-null terminated
                                '  list of strings.  The first is the printer
                                '  name and the second is the printer port.
Const PO_REN_PORT = &H34         '  PO_RENAME and PO_PORTCHANGE at same time.


Private Declare Function fCreateShellLink Lib "VB5STKIT.DLL" (ByVal _
lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
Private Sub Command1_Click()
CD1.ShowOpen
Text1.Text = CD1.FileName
End Sub

Private Sub Command2_Click()
  Dim bi As BROWSEINFO
  Dim idl As ITEMIDLIST
  Dim rtn&, pidl&, path$, pos%
  
  '  the calling app
  bi.hOwner = Me.hwnd
  
 
 
  '  set the banner text
  bi.lpszTitle = "Browsing"
  
  '  set the type of folder to return
  '  play with these option constants to see what can be returned
  bi.ulFlags = BIF_RETURNONLYFSDIRS  'BIF_RETURNFSANCESTORS 'BIF_BROWSEFORPRINTER + BIF_DONTGOBELOWDOMAIN
  
  '  show the browse folder dialog
  pidl& = SHBrowseForFolder(bi)
  
  '  if displaying the return value, get the selected folder
    path$ = Space$(512)
    rtn& = SHGetPathFromIDList(ByVal pidl&, ByVal path$)
    If rtn& Then
      
      '  parce & display the folder selection
      pos% = InStr(path$, Chr$(0))
      Text3.Text = Left(path$, pos - 1)
    Else
      MsgBox "Dialog was cancelled", vbInformation
    End If

End Sub
Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Dim lReturn As Long

If Text3.Text = "" Then
Else
lReturn = fCreateShellLink("..\..\..\..\", _
Text2.Text, Text1.Text, "")
Dim lResult As Long, SHF As SHFILEOPSTRUCT
SHF.hwnd = hwnd
SHF.wFunc = FO_COPY
SHF.pFrom = "C:\" & Text2.Text & ".lnk"
SHF.pTo = Text3.Text
SHF.fFlags = FOF_FILESONLY
lResult = SHFileOperation(SHF)
DoEvents
Kill "C:\" & Text2.Text & ".lnk"
End If

If Check1.Value = 1 Then
lReturn = fCreateShellLink("..\..\Desktop", _
Text2.Text, Text1.Text, "")
End If

If Check2.Value = 1 Then
lReturn = fCreateShellLink("\Startup", Text2.Text, _
Text1.Text, "")
End If

If Check3.Value = 1 Then
lReturn = fCreateShellLink("", Text2.Text, _
Text1.Text, "")
End If

If Check5.Value = 1 Then
lReturn = fCreateShellLink("..\..\Start Menu", _
Text2.Text, Text1.Text, "")
End If

DoEvents
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check5.Value = 0
MsgBox "DONE"

End Sub

Private Sub Form_Load()
CD1.CancelError = False
End Sub



Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Text2.SetFocus
 DoEvents
 End If

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Text3.SetFocus
 DoEvents
 End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command6_Click
 DoEvents
 End If
End Sub


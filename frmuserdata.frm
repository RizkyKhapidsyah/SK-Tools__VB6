VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmuserdata 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Store User Information"
   ClientHeight    =   5430
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8640
   Icon            =   "frmuserdata.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8640
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   40
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   3120
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   2520
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   1920
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   1320
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   6000
      TabIndex        =   39
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   5145
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Store User Information"
      Height          =   255
      Left            =   6000
      TabIndex        =   34
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8160
      Top             =   720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove From List"
      Height          =   255
      Left            =   6000
      TabIndex        =   33
      Top             =   3960
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   6000
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   32
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Users"
      Height          =   285
      Left            =   6720
      TabIndex        =   30
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6000
      TabIndex        =   29
      Top             =   240
      Width           =   2535
   End
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5775
      TabIndex        =   22
      Top             =   3975
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   26
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   25
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1213
         TabIndex        =   24
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   23
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "User Name"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   21
      Top             =   0
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "User Must Change Password"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   19
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Profile Path"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Primary Group"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Password Never Expires"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Login Script"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Home Directory"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Discription"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Account Type"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Account Expires"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Account Disabled"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   330
      Left            =   0
      Top             =   4275
      Visible         =   0   'False
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=User_Migration.mdb;"
      OLEDBString     =   "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=User_Migration.mdb;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmuserdata.frx":0E42
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   $"frmuserdata.frx":0F17
      Height          =   975
      Left            =   240
      TabIndex        =   41
      Top             =   3960
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "100%"
      Height          =   255
      Left            =   8160
      TabIndex        =   38
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "50%"
      Height          =   255
      Left            =   4140
      TabIndex        =   37
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "0%"
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   6000
      TabIndex        =   31
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Server or Domain"
      Height          =   255
      Left            =   6000
      TabIndex        =   28
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "User Name:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "User Must Change Password:"
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Profile Path:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Primary Group:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Password Never Expires:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Login Script:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Home Directory:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Discription:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Account Type:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Account Expires:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Account Disabled:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "frmuserdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
MousePointer = vbHourglass
Dim dso As IADsOpenDSObject
username = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = Combo1.Text

Dim container As IADsContainer
Dim containername As String
containername = Combo1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set container = GetObject("WinNT://" & containername)
Else
Set dso = GetObject("WinNT:")
Set container = dso.OpenDSObject("WinNT://" & DomainName, username, password, 1)
End If

container.Filter = Array("User")
Dim user As IADsUser
For Each user In container
List1.AddItem user.Name
Next
MousePointer = 0
End Sub

Private Sub Command2_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
Err = 0
End Sub

Private Sub Command3_Click()
List1.ListIndex = 0
ProgressBar1.Max = List1.ListCount
ProgressBar1.Value = 0
List2.Clear

DoEvents
Timer3.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next
Combo1.AddItem MDIFrmmain.Winsock1.LocalHostName
Dim namespace As IADsContainer
Dim domain As IADs
 'Loads Combo box1 with all the current domains
Set namespace = GetObject("WinNT:")

For Each domain In namespace
Combo1.AddItem domain.Name
Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error Resume Next
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "Total Users: " & List1.ListCount
If List1.ListCount = 0 Then
Command3.Enabled = False
Else
Command3.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If List2.ListCount = List1.ListCount Then
Timer2.Enabled = False
Else
List1.ListIndex = List1.ListIndex + 1
Timer3.Enabled = True
Timer2.Enabled = False
End If
Err = 0
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
cmdAdd_Click
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
MousePointer = vbHourglass

Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomian As String

userdomain = Combo1.Text
username = List1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If

txtFields(10).Text = List1.Text

retval = user.Description
txtFields(3).Text = retval

Dim Flags As Long
Flags = user.Get("userflags")
If (Flags And &H10000) <> 0 Then
txtFields(6).Text = "Checked"
Else
txtFields(6).Text = "Unchecked"
End If


If (Flags And &H2) <> 0 Then
txtFields(0).Text = "Checked"
Else
txtFields(0).Text = "Unchecked"
End If


Dim passwordexpired As Integer
passwordexpired = user.Get("passwordexpired")
If passwordexpired = 1 Then
txtFields(9).Text = "Checked"
Else
txtFields(9).Text = "Unchecked"
End If

retval = user.LoginScript
txtFields(5).Text = retval
DoEvents

retval = user.Profile
txtFields(8).Text = retval
DoEvents

retval = user.HomeDirectory
txtFields(4).Text = retval

txtFields(7).Text = "N/A"

Flags = user.Get("userflags")
If (Flags And &H100) <> 0 Then
txtFields(2).Text = "Local Account"
Else
txtFields(2).Text = "Global Account"
End If

Dim date1 As Date
date1 = user.AccountExpirationDate
Text1.Text = date1

If Text1.Text = "12:00:00 AM" Then
txtFields(1).Text = "Never"
Else
txtFields(1).Text = Text1.Text
End If

Err = 0
MousePointer = 0
Timer4.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Timer5_Timer()
On Error Resume Next
cmdUpdate_Click

List2.AddItem List1.Text
ProgressBar1.Value = ProgressBar1.Value + 1

Err = 0

Timer2.Enabled = True
Timer5.Enabled = False

End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmuserrestoredata 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore Users"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   Icon            =   "frmuserrestoredata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Set Domain/Server"
      Height          =   255
      Left            =   6000
      TabIndex        =   39
      Top             =   720
      Width           =   2535
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7200
      Top             =   4920
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6600
      Top             =   4920
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   4920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Restore/Update Users"
      Height          =   255
      Left            =   6000
      TabIndex        =   37
      Top             =   4560
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   6000
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   36
      Top             =   1560
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   5400
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6000
      TabIndex        =   30
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Account Disabled"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2179
      TabIndex        =   16
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Account Expires"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2179
      TabIndex        =   15
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Account Type"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2179
      TabIndex        =   14
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Discription"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2179
      TabIndex        =   13
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Home Directory"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2179
      TabIndex        =   12
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Login Script"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2179
      TabIndex        =   11
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Password Never Expires"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   2179
      TabIndex        =   10
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Primary Group"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2179
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Profile Path"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   2179
      TabIndex        =   8
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "User Must Change Password"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   2179
      TabIndex        =   7
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "User Name"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   2179
      TabIndex        =   6
      Top             =   360
      Width           =   3375
   End
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5775
      TabIndex        =   0
      Top             =   4305
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1213
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   330
      Left            =   0
      Top             =   4605
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
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
      RecordSource    =   $"frmuserrestoredata.frx":0BC2
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
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   5640
      TabIndex        =   38
      Top             =   1680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Users that have been Created/Updated"
      Height          =   375
      Left            =   6000
      TabIndex        =   35
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "100%"
      Height          =   255
      Left            =   8520
      TabIndex        =   34
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "50%"
      Height          =   255
      Left            =   4275
      TabIndex        =   33
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "0%"
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Server or Domain to restore to:"
      Height          =   255
      Left            =   6000
      TabIndex        =   29
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblLabels 
      Caption         =   "Account Disabled:"
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   28
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Account Expires:"
      Height          =   255
      Index           =   1
      Left            =   270
      TabIndex        =   27
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Account Type:"
      Height          =   255
      Index           =   2
      Left            =   270
      TabIndex        =   26
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Discription:"
      Height          =   255
      Index           =   3
      Left            =   270
      TabIndex        =   25
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Home Directory:"
      Height          =   255
      Index           =   4
      Left            =   270
      TabIndex        =   24
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Login Script:"
      Height          =   255
      Index           =   5
      Left            =   270
      TabIndex        =   23
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Password Never Expires:"
      Height          =   255
      Index           =   6
      Left            =   270
      TabIndex        =   22
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Primary Group:"
      Height          =   255
      Index           =   7
      Left            =   270
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Profile Path:"
      Height          =   255
      Index           =   8
      Left            =   270
      TabIndex        =   20
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "User Must Change Password:"
      Height          =   375
      Index           =   9
      Left            =   270
      TabIndex        =   19
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "User Name:"
      Height          =   255
      Index           =   10
      Left            =   270
      TabIndex        =   18
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   870
      TabIndex        =   17
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmuserrestoredata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command2_Click
 DoEvents
 End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
datPrimaryRS.Recordset.MoveFirst
ProgressBar1.Max = datPrimaryRS.Recordset.RecordCount
ProgressBar1.Value = 0
Timer3.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
MousePointer = vbHourglass
List1.Clear
List2.Clear

MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."

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
List2.AddItem user.Name
Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
MousePointer = 0
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
Label1.Caption = "Total Entries: " & datPrimaryRS.Recordset.RecordCount
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If List1.ListCount = datPrimaryRS.Recordset.RecordCount Then
Timer2.Enabled = False
Else
datPrimaryRS.Recordset.MoveNext
Timer3.Enabled = True
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
MousePointer = vbHourglass
List2.Text = txtFields(10).Text
If List2.Text = txtFields(10).Text Then
Timer4.Enabled = True
Timer3.Enabled = False
Exit Sub
End If

Dim dso As IADsOpenDSObject
username = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = Combo1.Text

Dim container As IADsContainer
Dim containername As String
Dim user As IADsUser
Dim newuser As String
containername = Combo1.Text

If frmdomainlogin.Check1.Value = 1 Then
Set container = GetObject("WinNT://" & containername)
Else
Set dso = GetObject("WinNT:")
Set container = dso.OpenDSObject("WinNT://" & containername, username, password, 0)
End If

newuser = txtFields(10).Text
Set user = container.Create("User", newuser)
user.SetInfo

Err = 0
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
Dim dso As IADsOpenDSObject
username2 = frmdomainlogin.Text1.Text
password = frmdomainlogin.Text2.Text
DomainName = Combo1.Text

Dim user As IADsUser
Dim username As String
Dim userdomian As String

userdomain = Combo1.Text
username = txtFields(10).Text
If frmdomainlogin.Check1.Value = 1 Then
Set user = GetObject("WinNT://" & userdomain & "/" & username & ",user")
Else
Set dso = GetObject("WinNT:")
Set user = dso.OpenDSObject("WinNT://" & DomainName & "/" & username & ",user", username2, password, 1)
End If

Dim passwordexpired As Integer
Dim Flags As Long
Dim newvalue As Boolean
Dim newfullname As String
Dim newdescription As String

If txtFields(3).Text = "" Then
newdescription = ""
user.Description = newdescription
user.SetInfo
Else
newdescription = txtFields(3).Text
user.Description = newdescription
user.SetInfo
End If

If txtFields(9).Text = "Checked" Then
user.Put "PasswordExpired", 1
user.SetInfo
Else
user.Put "PasswordExpired", 0
user.SetInfo
End If


If txtFields(6).Text = "Checked" Then
Flags = user.Get("userflags")
user.Put "userflags", Flags Or &H10000
user.SetInfo
Else
Flags = user.Get("userflags")
user.Put "userflags", Flags Xor &H10000
user.SetInfo
End If

If txtFields(0).Text = "Checked" Then
newvalue = True
user.AccountDisabled = newvalue
user.SetInfo
Else
newvalue = False
user.AccountDisabled = newvalue
user.SetInfo
End If

Dim newvalue2 As String

If txtFields(5).Text = "" Then
newvalue2 = ""
user.LoginScript = newvalue2
user.SetInfo
Else
newvalue2 = txtFields(5).Text
user.LoginScript = newvalue2
user.SetInfo
End If

If txtFields(8).Text = "" Then
newvalue2 = ""
user.Profile = newvalue2
user.SetInfo
Else
newvalue2 = txtFields(8).Text
user.Profile = newvalue2
user.SetInfo
End If

If txtFields(4).Text = "" Then
newvalue2 = ""
Call user.Put("HomeDirDrive", "")
user.HomeDirectory = newvalue2
user.SetInfo
Else
newvalue2 = txtFields(4).Text
Call user.Put("HomeDirDrive", "")
user.HomeDirectory = newvalue2
user.SetInfo
End If

Dim date1 As Date

If txtFields(1).Text = "Never" Then
date1 = #12:00:00 AM#
user.AccountExpirationDate = date1
user.SetInfo
Else
date1 = txtFields(1).Text
user.AccountExpirationDate = date1
user.SetInfo
End If

If txtFields(2).Text = "Global Account" Then
Flags = user.Get("userflags")
user.Put "userflags", Flags Xor &H100
user.SetInfo
Flags = user.Get("userflags")
user.Put "userflags", Flags Xor &H200
user.SetInfo
Else
End If

If txtFields(2).Text = "Local Account" Then
Flags = user.Get("userflags")
user.Put "userflags", Flags Xor &H200
user.SetInfo
Flags = user.Get("userflags")
user.Put "userflags", Flags Xor &H100
user.SetInfo
Else
End If

Err = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
List1.AddItem txtFields(10).Text
ProgressBar1.Value = ProgressBar1.Value + 1
MousePointer = 0
Timer2.Enabled = True
Timer4.Enabled = False
End Sub

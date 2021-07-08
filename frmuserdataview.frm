VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmuserdataview 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View/Clear/Modify Stored User Information"
   ClientHeight    =   4920
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5820
   Icon            =   "frmuserdataview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5820
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5820
      TabIndex        =   22
      Top             =   4290
      Width           =   5820
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
      Left            =   2179
      TabIndex        =   21
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "User Must Change Password"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   2179
      TabIndex        =   19
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Profile Path"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   2179
      TabIndex        =   17
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Primary Group"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2179
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Password Never Expires"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   2179
      TabIndex        =   13
      Top             =   1440
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
      DataField       =   "Home Directory"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2179
      TabIndex        =   9
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Discription"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2179
      TabIndex        =   7
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Account Type"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2179
      TabIndex        =   5
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Account Expires"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2179
      TabIndex        =   3
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Account Disabled"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2179
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   4590
      Width           =   5820
      _ExtentX        =   10266
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
      RecordSource    =   $"frmuserdataview.frx":0BC2
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   863
      TabIndex        =   28
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lblLabels 
      Caption         =   "User Name:"
      Height          =   255
      Index           =   10
      Left            =   270
      TabIndex        =   20
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "User Must Change Password:"
      Height          =   375
      Index           =   9
      Left            =   270
      TabIndex        =   18
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Profile Path:"
      Height          =   255
      Index           =   8
      Left            =   270
      TabIndex        =   16
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Primary Group:"
      Height          =   255
      Index           =   7
      Left            =   270
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Password Never Expires:"
      Height          =   255
      Index           =   6
      Left            =   270
      TabIndex        =   12
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Login Script:"
      Height          =   255
      Index           =   5
      Left            =   270
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Home Directory:"
      Height          =   255
      Index           =   4
      Left            =   270
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Discription:"
      Height          =   255
      Index           =   3
      Left            =   270
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Account Type:"
      Height          =   255
      Index           =   2
      Left            =   270
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Account Expires:"
      Height          =   255
      Index           =   1
      Left            =   270
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Account Disabled:"
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "frmuserdataview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

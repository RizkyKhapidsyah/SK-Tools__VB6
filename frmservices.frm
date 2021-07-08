VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmservices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Services"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmservices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7695
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   5520
      Top             =   120
   End
   Begin VB.CommandButton Start 
      Caption         =   "Start"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Pause 
      Caption         =   "Pause"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5040
      Top             =   120
   End
   Begin VB.TextBox Server 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton Connect 
      Caption         =   "Connect"
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Service Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Service Description"
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Service State"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Server"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmservices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Locator As SWbemLocator
Public services As SWbemServices
Public TimerCount
Public Item As ListItem

Dim WithEvents eventSink As SWbemSink
Attribute eventSink.VB_VarHelpID = -1
    
Public Sub InitialiseView()
    ListView1.ListItems.Clear
End Sub

Private Sub Server_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Connect_Click
 DoEvents
 End If

End Sub

Private Sub Timer1_Timer()

TimerCount = TimerCount + 1
If TimerCount > 10 Then
    Item.Bold = False
    Timer1.Enabled = False
    TimerCount = 0
Else
    Item.Bold = Not Item.Bold
End If

End Sub

Private Sub eventSink_OnObjectReady(ByVal Object As WbemScripting.ISWbemObject, ByVal AsyncContext As WbemScripting.ISWbemNamedValueSet)

Dim ServiceName
Dim ServiceStatus

ServiceName = Object.TargetInstance.Name
ServiceStatus = Object.TargetInstance.State
Set Item = ListView1.FindItem(ServiceName)
Item.SubItems(2) = ServiceStatus
Item.Bold = True
Timer1.Enabled = True
TimerCount = 0

End Sub

Public Sub LoadView()
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
    Dim Enumerator As SWbemObjectSet
    Dim Object As SWbemObject
    Dim Item As ListItem
    
    On Error Resume Next
        
    SavePointer = Form1.MousePointer
    Form1.MousePointer = vbHourglass
    Form1.Enabled = False
    ListView1.ListItems.Clear
    
    eventSink.Cancel
    
    Set services = Locator.ConnectServer(Server.Text)
    services.ExecNotificationQueryAsync eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = services.ExecQuery("Select * From Win32_Service")
    
    For Each Object In Enumerator
    
        Set Item = ListView1.ListItems.Add(, Object.Name, Object.Name)
        Item.SubItems(1) = Object.Description
        Item.SubItems(2) = Object.State
        
    Next
    
    Form1.Enabled = True
    Form1.MousePointer = SavePointer
    MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End Sub

Private Sub Form_Load()

    Set Locator = New SWbemLocator
    Set eventSink = New SWbemSink
    
    InitialiseView

End Sub

Private Sub Connect_Click()
    
    LoadView

End Sub

Private Sub Exit_Click()

    End
    
End Sub

Private Sub Pause_Click()
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
    Dim ServiceObject As SWbemObject
    Dim ServiceName
    
    On Error Resume Next
    ServiceName = ListView1.SelectedItem.Text
    If Err.Number = 0 Then
    
        Set ServiceObject = services.Get("Win32_Service='" & ServiceName & "'")
        
        ' Note how the CIM method "PauseService" of Win32_Service
        ' is executed as if it were an automation method of SWbemObject
        ServiceObject.PauseService
    End If
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End Sub

Private Sub Start_Click()
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
    Dim ServiceObject As SWbemObject
    Dim ServiceName
    
    On Error Resume Next
    ServiceName = ListView1.SelectedItem.Text
    If Err.Number = 0 Then
    
        ' Note how the CIM method "StartService" of Win32_Service
        ' is executed as if it were an automation method of SWbemObject
        Set ServiceObject = services.Get("Win32_Service='" & ServiceName & "'")
        ServiceObject.StartService
    End If
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End Sub

Private Sub Stop_Click()
MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
    Dim ServiceObject As SWbemObject
    Dim ServiceName
    
    On Error Resume Next
    ServiceName = ListView1.SelectedItem.Text
    If Err.Number = 0 Then
    
        ' Note how the CIM method "StopService" of Win32_Service
        ' is executed as if it were an automation method of SWbemObject
        Set ServiceObject = services.Get("Win32_Service='" & ServiceName & "'")
        ServiceObject.StopService
    End If
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End Sub


Private Sub Timer2_Timer()
Label2.Caption = "Total Services: " & ListView1.ListItems.Count
End Sub

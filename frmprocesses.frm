VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmprocesses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processes"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "frmprocesses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5160
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   3960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Get Processes"
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kill Process"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PID"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Priority"
         Object.Width           =   1457
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Mem Usage"
         Object.Width           =   2602
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Current Processes"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Computer Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmprocesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents sink As SWbemSink
Attribute sink.VB_VarHelpID = -1
Dim services As SWbemServices
Private Sub Command2_Click()
On Error Resume Next
Dim computername As String
Dim processname As String
processname = ListView.SelectedItem
computername = Text1.Text
For Each Process In GetObject("winmgmts:{impersonationLevel=impersonate}!//" & computername).ExecQuery("select * from Win32_Process where Name='" & processname & "'")
Process.Terminate
Next
DoEvents
Command3_Click
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    MDIFrmmain.StatusBar1.Panels(1).Text = "Status: Working..."
If Text1.Text = "" Then
MsgBox "Please enter a Computername"
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
Exit Sub
Else
ListView.ListItems.Clear
MousePointer = vbHourglass
    Dim computername As String
    computername = Text1.Text

    ' Create a sink to recieve the results of the enumeration
    Set sink = New SWbemSink
    
    ' Connect to root\cimv2.
    Set services = GetObject("winmgmts://" & computername)
' Perform the asynchronous enumeration of processes
services.InstancesOfAsync sink, "Win32_process"
MousePointer = 0
MDIFrmmain.StatusBar1.Panels(1).Text = "Status:"
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command3_Click
 DoEvents
 End If

End Sub
Private Sub sink_OnCompleted(ByVal iHResult As WbemScripting.WbemErrorEnum, ByVal objWbemErrorObject As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)
    ' This event handler is called when there are no more instances to
    ' be returned
    MousePointer = vbDefault
    
    If (iHResult <> wbemNoErr) Then
        MsgBox "Error: " & Err.Description & " [0x" & Hex(Err.Number) & "]", vbOKOnly
    End If
    
End Sub

Private Sub sink_OnObjectReady(ByVal Process As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)
    ' This event handler is called once for every process returned by the
    ' enumeration
    
    Key = "Handle:" & Process.Handle
    
    Set Item = ListView.ListItems.Add(, Key, Process.Name)
    Item.SubItems(1) = Process.Handle
    
    If vbNull <> VarType(Process.Priority) Then
        Item.SubItems(2) = Process.Priority
    End If
        
    If vbNull <> VarType(Process.WorkingSetSize) Then
        Item.SubItems(3) = CStr(Process.WorkingSetSize / 1024) + " K"
    End If
    
End Sub

Private Sub Timer1_Timer()
Label3.Caption = "Total Processes: " & ListView.ListItems.Count
End Sub

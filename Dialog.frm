VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scan in:"
   ClientHeight    =   3855
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3168
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   4692
      _ExtentX        =   8281
      _ExtentY        =   5583
      _Version        =   393217
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   3540
      Visible         =   0   'False
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   675
      Left            =   60
      TabIndex        =   3
      Top             =   3180
      Width           =   5925
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 Dim cFunction   As clsFunction
 Dim nodX        As Node
 Dim lChecks     As Long

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  
  FreezeCtl TreeView1
  lChecks = colCodes.Count * colVariables.Count
  Label1.Caption = "Functions:" & colCodes.Count & ", Variables:" & colVariables.Count & ", if You select all nodes - program will iterate:" & lChecks & " times."
  
  Set nodX = TreeView1.Nodes.Add(, , "R", frmAddIn.VBInstance.ActiveVBProject.Name)
  
  For Each cFunction In colCodes
    With cFunction
      If Len(.sName) = 0 Then
        Set nodX = TreeView1.Nodes.Add("R", tvwChild, .sModuleName, .sModuleName)
       Else
        Set nodX = TreeView1.Nodes.Add(.sModuleName, tvwChild, .sModuleName & .sName, .sName)
      End If
    End With
  Next cFunction
  Set cFunction = Nothing
  UnFreezeCtl TreeView1

Exit Sub
  '{{{ Added It!
Generated_trap:
  UnFreezeCtl TreeView1
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Form_Load")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Public Function InitProgress(Optional sText As String, Optional lMin As Long, Optional lMax As Long) As Boolean
  On Error GoTo ErrorHandler
  If lMax = 0 Then
    lMax = 100
  End If
  With Me.pbProgress
    .Max = lMax
    .Min = lMin
    .Value = lMin
  End With
  Me.Label1.Caption = sText
  'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
  Me.Visible = True
  Me.ZOrder 0
  InitProgress = True
ErrorHandler:
  'Exit without setting ReportProgress = True
End Function

':) Ulli's Code Formatter V2.0 (2001-01-23 10:53:27) 1 + 226 = 227 Lines

Public Function ReportProgress(Optional sText As String, Optional vValue As Long) As Boolean

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  On Error GoTo ErrorHandler
  With Me
    .pbProgress.Visible = True
    .pbProgress.Value = vValue
    .Label1.Caption = sText
    .pbProgress.Refresh
    .Label1.Refresh
  End With
  ReportProgress = True
ErrorHandler:
  'Exit without setting ReportProgress = True
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "ReportProgress")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Function

Private Sub CancelButton_Click()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  Unload Me
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "CancelButton_Click")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub OKButton_Click()
 Dim cVariable     As clsVariable
 Dim cVariableMas  As clsVariable 'master variable
 Dim nodX          As Node
 Dim li            As ListItem
 Dim i             As Long
 Dim j             As Long
 Dim cFunction     As clsFunction
 Dim sTempStr      As String
 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  
  'On Error GoTo errOkButton
    
  For Each nodX In TreeView1.Nodes
    With nodX
      If .Key <> "R" Then
        colCodes.ItemByKey(.Key).bScanForUse = .Checked
      End If
    End With
  Next nodX
  Set nodX = Nothing
  
  FreezeCtl TreeView1
  'refresh data in list view

  
  'search data for var use
  With cSearchObject
    .ScanForUse
  End With
  
  If colVariables Is Nothing Then
    Exit Sub
  End If
  
  If colVariableNames Is Nothing Then
    Exit Sub
  End If
  
  frmAddIn.TreeView1.Nodes.Clear
    
  Set nodX = frmAddIn.TreeView1.Nodes.Add(, , "R", frmAddIn.VBInstance.ActiveVBProject.Name)

  For Each cVariable In colVariableNames
    i = i + 1
    With cVariable
      Set nodX = frmAddIn.TreeView1.Nodes.Add("R", tvwChild, "C" & i, .sName)
      'add info about variable use
      If Not .FastCollection Is Nothing Then
        j = 0
        For Each cFunction In .FastCollection
          j = j + 1
          With cFunction
            Set nodX = frmAddIn.TreeView1.Nodes.Add("C" & i, tvwChild, _
                "C" & i & "F" & j, .sModuleName & ":" & .sName & _
                ":" & .lHowMany)
            nodX.EnsureVisible
          End With
        Next cFunction
      Else
        If .FastCollection.Count = 0 Then
          frmHints.Show
          sTempStr = "[" & .sModuleName
          If Len(Trim(.sFunctionName)) > 0 Then
              sTempStr = sTempStr & "." & .sFunctionName
          End If
          sTempStr = sTempStr & "] " & .sName & " is death variable - never used."
          frmHints.AddHint sTempStr
        End If
      End If
      
'      For Each cVariableMas In colVariables
'        If cVariableMas.sName = .sName Then
'           Set cVariableMas.FastCollection = .FastCollection
'        End If
'      Next cVariableMas
'
'      Set cVariableMas = Nothing
      
      Set cFunction = Nothing
      
    End With
  Next cVariable
  
  Set cVariable = Nothing
  
  frmAddIn.TreeView1.Style = tvwTreelinesText ' Style 4.
  frmAddIn.TreeView1.BorderStyle = vbFixedSingle
  frmAddIn.TreeView1.Enabled = True
  
  UnFreezeCtl TreeView1
  
  frmAddIn.cmdReport.Enabled = True
  frmAddIn.cmdFind.Enabled = True
'
'  FreezeCtl frmAddIn.ListView1
  
  
'  If colVariables Is Nothing Then
'    Exit Sub
'  End If
'
'
'  frmAddIn.ListView1.ListItems.Clear
'  frmHints.RemoveAllHints
'
'
'  For Each cVariableMas In colVariables
'
'    Set cVariable = colVariables.Item(i)
'    With cVariable
'      Set li = frmAddIn.ListView1.ListItems.Add(, , .sName)
'      li.SubItems(1) = .sScope
'      li.SubItems(2) = .sType
'      li.SubItems(3) = .sModuleName
'      li.SubItems(4) = .sFunctionName
'      li.SubItems(5) = .bConst
'      li.SubItems(6) = .vValue
'      li.SubItems(7) = .bWithEvents
'      li.SubItems(8) = .bPreserve
'      li.SubItems(9) = .sDimension
'      li.SubItems(10) = .FastCollection.Count
'      li.SubItems(11) = .sDescription
'      li.SubItems(12) = .sComment
'      If Len(Trim(.sDescription & .sComment)) > 0 Then
'        frmHints.Show
'        sTempStr = "[" & .sModuleName
'        If Len(Trim(.sFunctionName)) > 0 Then
'            sTempStr = sTempStr & "." & .sFunctionName
'        End If
'        sTempStr = sTempStr & "] " & .sName & " " & .sDescription & " " & .sComment
'        frmHints.AddHint sTempStr
'      End If
'      If .FastCollection.Count > 0 Then
'        frmHints.Show
'        sTempStr = "[" & .sModuleName
'        If Len(Trim(.sFunctionName)) > 0 Then
'            sTempStr = sTempStr & "." & .sFunctionName
'        End If
'        sTempStr = sTempStr & "] " & .sName & " Death variable - never used."
'        frmHints.AddHint sTempStr
'      End If
'    End With
'  Next cVariableMas
'
'  Set cVariableMas = Nothing
  
'  UnFreezeCtl frmAddIn.ListView1
  
  frmAddIn.cmdReport.Enabled = True
  frmAddIn.cmdFind.Enabled = True
  
  
  Unload Me

  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "OKButton_Click")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub


Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  FreezeCtl TreeView1
  TreeSelectiveCheck Node
  UnFreezeCtl TreeView1
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "TreeView1_NodeCheck")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

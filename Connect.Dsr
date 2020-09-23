VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7935
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   8670
   _ExtentX        =   15293
   _ExtentY        =   13996
   _Version        =   393216
   Description     =   "I wrote this addin for me to help working with multiple variables and constants."
   DisplayName     =   "Var Helper"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Public mcbMenuCommandBar      As Office.CommandBarControl

Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Private mfrmAddIn             As New frmAddIn

Sub Hide()
'{{{ Added It!
On Error GoTo Generated_trap '}}}
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmAddIn.Hide
   
'{{{ Added It!
Err.Clear
Generated_trap:
If Err <> 0 Then
      Select Case ToDoOnError(Err, "Hide")
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
     End Select
End If '}}}
End Sub

Sub Show()
'{{{ Added It!
On Error GoTo Generated_trap '}}}
  
    On Error Resume Next
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
    End If
    
    Set mfrmAddIn.VBInstance = VBInstance
    Set mfrmAddIn.Connect = Me
    FormDisplayed = True
    mfrmAddIn.Show
   
'{{{ Added It!
Err.Clear
Generated_trap:
If Err <> 0 Then
      Select Case ToDoOnError(Err, "Show")
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
     End Select
End If '}}}
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
'{{{ Added It!
On Error GoTo Generated_trap '}}}
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar(sAddinName)
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
'{{{ Added It!
Err.Clear
Generated_trap:
If Err <> 0 Then
      Select Case ToDoOnError(Err, "AddinInstance_OnConnection")
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
     End Select
End If '}}}
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
'{{{ Added It!
On Error GoTo Generated_trap '}}}
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing
'{{{ Added It!
Err.Clear
Generated_trap:
If Err <> 0 Then
      Select Case ToDoOnError(Err, "AddinInstance_OnDisconnection")
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
     End Select
End If '}}}
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
'{{{ Added It!
On Error GoTo Generated_trap '}}}
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
'{{{ Added It!
Err.Clear
Generated_trap:
If Err <> 0 Then
      Select Case ToDoOnError(Err, "IDTExtensibility_OnStartupComplete")
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
     End Select
End If '}}}
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'{{{ Added It!
On Error GoTo Generated_trap '}}}
  With frmAddIn
    Set .VBInstance = VBInstance
    Set .Connect = Me
    .Show
  End With
'{{{ Added It!
Err.Clear
Generated_trap:
If Err <> 0 Then
      Select Case ToDoOnError(Err, "MenuHandler_Click")
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
     End Select
End If '}}}
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
'{{{ Added It!
On Error GoTo Generated_trap '}}}
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:
'{{{ Added It!
Err.Clear
Generated_trap:
If Err <> 0 Then
      Select Case ToDoOnError(Err, "AddToAddInCommandBar")
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
     End Select
End If '}}}
End Function



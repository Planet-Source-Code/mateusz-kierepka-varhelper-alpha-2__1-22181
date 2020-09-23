Attribute VB_Name = "modTreeSelect"
Option Explicit

Private Enum NodeCheck
  nodChecked = True
  nodUnchecked = False
  nodPartial = 1
End Enum

Private Const WM_SETREDRAW  As Long = &HB
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub FreezeCtl(Ctl As Control)
#If NotInIdeDebug Then
  SendMessageLong Ctl.hWnd, WM_SETREDRAW, 0, 0
#End If
End Sub

Private Function TreeCheckSibling(Node As MSComctlLib.Node, ByVal bCheck As Boolean) As Boolean
    
 Dim nodX As Node

  TreeCheckSibling = True
    
  Set nodX = Node.FirstSibling
    
  Do Until nodX Is Nothing
    If nodX.Checked <> bCheck Then
      TreeCheckSibling = False
      Exit Do
    End If
        
    Set nodX = nodX.Next
  Loop
    
End Function

':) Ulli's Code Formatter V2.0 (2001-01-23 10:53:19) 10 + 119 = 129 Lines

Public Sub TreeSelectiveCheck(Node As MSComctlLib.Node)
    
  Node.Bold = False
    
  TreeSetChildren Node, Node.Checked
    
  If TreeCheckSibling(Node, Node.Checked) Then
    TreeSetParents Node, Node.Checked
   Else
    TreeSetParents Node, nodPartial
  End If
    
End Sub

Private Sub TreeSetChildren(Node As MSComctlLib.Node, bCheck As Boolean)

 Dim nodX As Node
        
  If Node.children = 0 Then
    Exit Sub
  End If
    
  Set nodX = Node.Child
  Do Until nodX Is Nothing
    nodX.Bold = False
    nodX.Checked = bCheck
        
    If nodX.children > 0 Then
      TreeSetChildren nodX, bCheck
    End If
        
    Set nodX = nodX.Next
  Loop
    
End Sub

Private Sub TreeSetParents(Node As MSComctlLib.Node, ByVal nCheck As NodeCheck)
    
 Dim nodX As Node

  If (Node.Parent Is Nothing) Then
    Exit Sub
  End If
    
  Set nodX = Node.Parent
  Select Case nCheck
   Case nodChecked
    nodX.Checked = True
    nodX.Bold = False
   Case nodUnchecked
    nodX.Checked = False
    nodX.Bold = False
   Case nodPartial
    nodX.Checked = False
    nodX.Bold = True
  End Select
    
  If nCheck = nodPartial Then
    TreeSetParents nodX, nodPartial
   Else
    If TreeCheckSibling(nodX, nodX.Checked) Then
      TreeSetParents nodX, nodX.Checked
     Else
      TreeSetParents nodX, nodPartial
    End If
  End If
        
End Sub

Public Sub TreeSingleCheck(Tree As TreeView, Node As MSComctlLib.Node)
    
 Dim nodX As Node

  If Node.Checked Then
    For Each nodX In Tree.Nodes
      If nodX.Index <> Node.Index And nodX.Checked Then
        nodX.Checked = False
      End If
    Next nodX
    Set nodX = Nothing
  End If
    
End Sub

Public Sub UnFreezeCtl(Ctl As Control)

  On Error Resume Next    'Gdyby Ctl nie mia³a .Refresh
#If NotInIdeDebug Then
   SendMessageLong Ctl.hWnd, WM_SETREDRAW, 1, 0
   Ctl.Refresh
#End If
  On Error GoTo 0
End Sub

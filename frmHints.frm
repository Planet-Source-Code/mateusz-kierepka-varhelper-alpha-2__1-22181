VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHints 
   Caption         =   "Hints"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvHints 
      Height          =   3225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   5689
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmHints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
On Error Resume Next
    With lvHints
        .Width = Width - 100
        .Height = Height - 100
        .ColumnHeaders(0).Width = .Width
    End With
On Error GoTo 0
End Sub

Public Sub AddHint(ByVal sHint As String)
Dim li    As ListItem
    Set li = lvHints.ListItems.Add(, , sHint)
End Sub

Public Sub RemoveAllHints()
    lvHints.ListItems.Clear
End Sub

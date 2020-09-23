VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddIn 
   Caption         =   "Var helper - Alpha version - demo"
   ClientHeight    =   3795
   ClientLeft      =   2190
   ClientTop       =   1950
   ClientWidth     =   6165
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   3225
      Left            =   6270
      ScaleHeight     =   1404.304
      ScaleMode       =   0  'User
      ScaleWidth      =   1716
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1950
      ScaleHeight     =   285
      ScaleWidth      =   945
      TabIndex        =   5
      Top             =   2940
      Visible         =   0   'False
      Width           =   975
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   390
         TabIndex        =   6
         Top             =   30
         Width           =   135
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3105
      Left            =   2850
      TabIndex        =   4
      Top             =   0
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   5477
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Scope"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Module Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Function Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Is Const"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Const Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Is WithEvents"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Preserve"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Dimension"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Comment"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6105
      TabIndex        =   2
      Top             =   3240
      Width           =   6165
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   6465
         TabIndex        =   3
         Top             =   30
         Width           =   6465
         Begin VB.CommandButton Command1 
            Cancel          =   -1  'True
            Caption         =   "&About"
            Height          =   375
            Left            =   4920
            TabIndex        =   12
            Top             =   30
            Width           =   1115
         End
         Begin VB.CommandButton OKButton 
            Caption         =   "Collect &data"
            Default         =   -1  'True
            Height          =   375
            Left            =   2520
            TabIndex        =   11
            Top             =   30
            Width           =   1115
         End
         Begin VB.CommandButton CancelButton 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   3720
            TabIndex        =   10
            Top             =   30
            Width           =   1115
         End
         Begin VB.CommandButton cmdReport 
            Caption         =   "&Report"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   30
            Width           =   1115
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find use"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   8
            Top             =   30
            Width           =   1115
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   3240
      Left            =   0
      ScaleHeight     =   3240
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3105
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   5477
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.Image imgSplitter 
      BorderStyle     =   1  'Fixed Single
      Height          =   3225
      Left            =   2640
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   165
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VBInstance As VBIDE.VBE
Public Connect As Connect

'Zmienne formy
Private mTPPX As Single
Private mTPPY As Single
Private mnMarginWidth As Single
Private mnMarginHeight As Single

'API: Kursor
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

Private Type ICONINFO
  fIcon As Long
  xHotspot As Long
  yHotspot As Long
  hbmMask As Long
  hbmColor As Long
End Type

Private prevOrder           As Integer
Private mbMoving            As Boolean
Private Const sglSplitLimit As Long = 500

Private Sub GetCursorDimensions(Optional PointerX As Single, Optional PointerY As Single, Optional Left As Single, Optional Top As Single, Optional Right As Single, Optional Bottom As Single)
 Dim ptCursor As POINTAPI
 Dim hCursor As Long
 Dim udtIconInfo As ICONINFO
 Dim nMultiplier As Single
 Dim udtBitmapInfo As BITMAPINFO

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}


  'Odczyt pozycji kursora
  If GetCursorPos(ptCursor) = 0 Then
    Err.Raise 5, , "GetCursorPos failed."
  End If

  'uchwyt do kursora
  hCursor = GetCursor
    
  'Odczyt informacji o ikonie kursora
  If GetIconInfo(hCursor, udtIconInfo) = 0 Then
    Err.Raise 5, , "GetIconInfo failed."
  End If
    
  If udtIconInfo.hbmMask = 0 Then
    Err.Raise 5, , "GetIconInfo zwróci³o nieprawid³owe hbmMask."
  End If
    
  'hbmColor = 0 => kursor jest czarno-bia³y
  If udtIconInfo.hbmColor = 0 Then
    'kursor jest czarno-bia³y => wysokoœæ bitmapy jest 2 razy wiêksza ni¿ wysokoœæ kursora
    nMultiplier = 0.5
   Else
    'kursor jest kolorowy => wysokoœæ bitmapy jest = wysokoœci kursora
    nMultiplier = 1
    'Usuniêcie kolorowej bitmapy zwróconej przez GetIconInfo
    DeleteObject udtIconInfo.hbmColor
  End If
    
  'Zaincjowanie biSize dla nastêpnej linii
  udtBitmapInfo.bmiHeader.biSize = Len(udtBitmapInfo.bmiHeader)
    
  'Odczyt informacji o bitmap-ie
  If GetDIBits(hDC, udtIconInfo.hbmMask, 0, 0, ByVal 0, udtBitmapInfo, 0) = 0 Then
    Err.Raise 5, , "Wywo³anie GetDIBits nieudane."
  End If
    
  'Usuniêcie maski bitmapy zwróconej przez GetIconInfo
  DeleteObject udtIconInfo.hbmMask
    
  'Przeliczenie wysokoœci (patrz wy¿ej)
  udtBitmapInfo.bmiHeader.biHeight = udtBitmapInfo.bmiHeader.biHeight * nMultiplier

  'Obliczenie zwracanych wartoœci
  With ptCursor
    'Zamiana na twipsy
    PointerX = .X * mTPPX
    PointerY = .Y * mTPPY
        
    'Po³o¿enie HotSpot-u kursora w twipsach
    Left = (.X - udtIconInfo.xHotspot) * mTPPX
    Top = (.Y - udtIconInfo.yHotspot) * mTPPY
  End With
    
  'Rozmiar kursora w twipsach
  Right = (udtBitmapInfo.bmiHeader.biWidth * mTPPX) + Left
  Bottom = (udtBitmapInfo.bmiHeader.biHeight * mTPPY) + Top
    
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "GetCursorDimensions")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub SetPosition()
 Dim nCursorLeft As Single
 Dim nCursorTop As Single
 Dim nCursorBottom As Single

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

    
  'Odczyt rozmiaru kursora
  GetCursorDimensions Left:=nCursorLeft, Top:=nCursorTop, Bottom:=nCursorBottom

  With Screen
    'Czy forma mieœci siê w poziomie
    If nCursorLeft + Picture4.Width <= .Width Then
      'Tak to wg kursora
      Picture4.Left = nCursorLeft - Me.Left
     Else
      'Nie to wg krawêdzi ekranu
      Picture4.Left = .Width - Picture4.Width
    End If
    
    'Czy forma mieœci siê w pionie
    If nCursorBottom + Picture4.Height <= .Height Then
      'Tak to ustawiamy pod kursorem
      Picture4.Top = nCursorTop - Me.Top
     Else
      'Nie to utawiamy nad kursorem
      Picture4.Top = nCursorTop - Height
    End If
  End With
    
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "SetPosition")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub SetSize(Optional MarginWidth As Long = 2, Optional MarginHeight As Long = 2)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  'Przeliczniki odczytujemy tylko raz

  mTPPX = Screen.TwipsPerPixelX
  mTPPY = Screen.TwipsPerPixelY

  'Przelicz marginesy
  mnMarginWidth = MarginWidth * mTPPX
  mnMarginHeight = MarginHeight * mTPPY
  'Margines lewy i górny
  Label1.Move mnMarginWidth, mnMarginWidth
  'dolny margines + border formy
  Picture4.Height = Label1.Height + (2 * mnMarginHeight)
  'Szerokoœæ formy z uwzglêdnieniem border
  Picture4.Width = Label1.Left + Label1.Width + mnMarginWidth
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "SetSize")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Public Sub ShowTip(ByVal sText As String)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  'Picture4.Visible = False
  
  Label1.Caption = sText
    
  'Ustaw rozmiar formy
  SetSize 2, 2
    
  'Ustaw po³o¿enie formy
  SetPosition
  Picture4.Refresh
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "ShowTip")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Sub SizeControls(X As Single)
  On Error Resume Next
  
   'set the width
   If X < 1500 Then
    X = 1500
   End If
   If X > (Me.Width - 1500) Then
    X = Me.Width - 1500
   End If
  
   TreeView1.Height = Height - Picture2.Height - 400
   ListView1.Height = TreeView1.Height
   imgSplitter.Height = TreeView1.Height
   imgSplitter.Top = 0
   imgSplitter.Left = X
   TreeView1.Width = X
   Picture1.Width = TreeView1.Width + 10
   ListView1.Width = (Me.Width - X - 350)
   ListView1.Left = X + imgSplitter.Width
   
   Picture3.Left = Me.Width - Picture3.Width
   
  On Error GoTo 0

End Sub

':) Ulli's Code Formatter V2.0 (2001-01-23 10:53:30) 59 + 640 = 699 Lines

Private Sub CancelButton_Click()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  Connect.Hide
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

Private Sub cmdFind_Click()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  frmDialog.Show vbModal, Me
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "cmdFind_Click")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub cmdReport_Click()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  If cSearchObject Is Nothing Then
    MsgBox "Please collect data."
   Else
    Call cSearchObject.SaveXML("VarReport")
  End If
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "cmdReport_Click")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub Command1_Click()
    frmAbout.Show vbModal
End Sub

Private Sub Form_Activate()
On Error Resume Next
   If Not frmHints Is Nothing Then
      With frmHints
        .Width = Width
        .Top = Top + Height
        .Left = Left
      End With
    End If
On Error GoTo 0
End Sub

Private Sub Form_Initialize()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  
  Set cSearchObject = New clsSearch
    
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Form_Initialize")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub Form_LostFocus()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  mbMoving = False
  Picture4.Visible = False
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Form_LostFocus")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  '  If (cHS.SplitterFormMouseUp(X, Y)) Then
  '    Form_Resize
  '  End If
  '  Picture4.Visible = False
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Form_MouseUp")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  Set cSearchObject = Nothing
  
  Call frmScanProgress.RemoveProgress
  Set frmScanProgress = Nothing
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Form_QueryUnload")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub Form_Resize()

  On Error Resume Next
   If Me.Width < 3000 Then
    Me.Width = 3000
   End If
    If Not frmHints Is Nothing Then
      With frmHints
        .Width = Width
        .Top = Top + Height
        .Left = Left
        .lvHints.ColumnHeaders(1).Width = Width - 200
      End With
    End If
   SizeControls imgSplitter.Left
  On Error GoTo 0
End Sub

Private Sub Form_Terminate()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  Set cSearchObject = Nothing
  '  Set cHS = Nothing
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Form_Terminate")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim sngPos As Single
 '{{{ Added It!

  On Error GoTo Generated_trap '}}}



  With imgSplitter
    picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
  
    sngPos = (((.Left + (.Width \ 2)) * 100) / Me.Width)
    
    Picture4.Visible = True
    ShowTip FormatNumber(sngPos, 1) & " %"
  
  End With
  picSplitter.Visible = True
  mbMoving = True
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "imgSplitter_MouseDown")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sngPos As Single
 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
 
  
  If mbMoving Then
    
    sngPos = X + imgSplitter.Left
    If sngPos < sglSplitLimit Then
      picSplitter.Left = sglSplitLimit
     ElseIf sngPos > Me.Width - sglSplitLimit Then
      picSplitter.Left = Me.Width - sglSplitLimit
     Else
      picSplitter.Left = sngPos
    End If
    
    With picSplitter
      sngPos = (((.Left + (.Width \ 2)) * 100) / Me.Width)
      
      Picture4.Visible = True
      ShowTip FormatNumber(sngPos, 1) & " %"
    End With
  End If
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "imgSplitter_MouseMove")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  SizeControls picSplitter.Left
  picSplitter.Visible = False
  Picture4.Visible = False
  mbMoving = False
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "imgSplitter_MouseUp")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub ListView1_Click()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  'Stop
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "ListView1_Click")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

Dim currSortKey As Integer

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
 

  With ListView1
    .SortKey = ColumnHeader.Index - 1
    currSortKey = .SortKey
    
    .SortOrder = Abs(Not .SortOrder = 1)
    .Sorted = True
    
    '    mnuOrder(prevOrder).Checked = False
    '
    '    mnuSortAZ.Checked = .SortOrder = 0
    '    mnuSortZA.Checked = mnuSortAZ.Checked = False
    
    If currSortKey > -1 Then
      '      mnuOrder(currSortKey).Checked = True
      prevOrder% = currSortKey
    End If
  End With
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "ListView1_ColumnClick")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Private Sub OKButton_Click()
 Dim cVariable     As clsVariable
 Dim nodX          As Node
 Dim li            As ListItem
 Dim i             As Long
 Dim j             As Long
 Dim cFunction     As clsFunction
 Dim sTempStr      As String
 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  FreezeCtl ListView1
  
  With cSearchObject
    Set .VBInstance = VBInstance
    .Scan
  End With
  If colVariables Is Nothing Then
    Exit Sub
  End If
  
  TreeView1.Nodes.Clear
  TreeView1.Enabled = False
  
  ListView1.ListItems.Clear
  
  frmHints.RemoveAllHints
  
  For Each cVariable In colVariables
    
    With cVariable
      Set li = ListView1.ListItems.Add(, , .sName)
      li.SubItems(1) = .sScope
      li.SubItems(2) = .sType
      li.SubItems(3) = .sModuleName
      li.SubItems(4) = .sFunctionName
      li.SubItems(5) = .bConst
      li.SubItems(6) = .vValue
      li.SubItems(7) = .bWithEvents
      li.SubItems(8) = .bPreserve
      li.SubItems(9) = .sDimension
      li.SubItems(10) = .sDescription
      li.SubItems(11) = .sComment
      If Len(Trim(.sDescription & .sComment)) > 0 Then
        frmHints.Show
        sTempStr = "[" & .sModuleName
        If Len(Trim(.sFunctionName)) > 0 Then
            sTempStr = sTempStr & "." & .sFunctionName
        End If
        sTempStr = sTempStr & "] " & .sName & " " & .sDescription & " " & .sComment
        frmHints.AddHint sTempStr
      End If
    End With
  Next cVariable

  Set cVariable = Nothing
  
  UnFreezeCtl ListView1
  
  cmdReport.Enabled = True
  cmdFind.Enabled = True

Exit Sub
  '{{{ Added It!
  
  Err.Clear
Generated_trap:
  
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "OKButton_Click")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}
  UnFreezeCtl ListView1
End Sub

Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  If Source = imgSplitter Then
    SizeControls X
  End If
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "TreeView1_DragDrop")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

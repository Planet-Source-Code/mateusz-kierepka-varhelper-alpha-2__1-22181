VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Var Helper"
   ClientHeight    =   4575
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5880
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   3157.746
   ScaleMode       =   0  'User
   ScaleWidth      =   5521.624
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox P2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1770
      Left            =   0
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   395
      TabIndex        =   4
      Top             =   1680
      Width           =   5925
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5460
      Top             =   2940
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4425
      TabIndex        =   0
      Top             =   3675
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4440
      TabIndex        =   1
      Top             =   4125
      Width           =   1245
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4346
      Left            =   0
      ScaleHeight     =   4290
      ScaleWidth      =   5865
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   5925
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "eMail: mkprog@mkprog.own.pl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   1
      Left            =   120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   6
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "www.mkprog.own.pl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   5
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Var Helper Addin - by MKPROG"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   253.544
      X2              =   5478.428
      Y1              =   2412.311
      Y2              =   2412.311
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   267.63
      X2              =   5478.428
      Y1              =   2422.664
      Y2              =   2422.664
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This code is from ProjectBOB Greg Cohen's with little changes

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private thetop As Long
Private p1hgt As Long
Private p1wid As Long
Private theleft As Long
Private Tempstring As String

' Reg Key Security Options...
Const READ_CONTROL              As Long = &H20000
Const KEY_QUERY_VALUE           As Long = &H1
Const KEY_SET_VALUE             As Long = &H2
Const KEY_CREATE_SUB_KEY        As Long = &H4
Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Const KEY_NOTIFY                As Long = &H10
Const KEY_CREATE_LINK           As Long = &H20
Const KEY_ALL_ACCESS            As Long = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Const ERROR_SUCCESS             As Long = 0
Const REG_SZ                    As Long = 1                         ' Unicode nul terminated string
Const REG_DWORD                 As Long = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC         As String = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC         As String = "MSINFO"
Const gREGKEYSYSINFO            As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO            As String = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub



Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


Sub Form_Load()
    With P1
        .AutoRedraw = True
        .Visible = False
        .FontSize = 8
        .ForeColor = vbWhite
        .BackColor = vbBlack
        P2.ForeColor = vbWhite
        P2.BackColor = vbBlack
        .ScaleMode = 3
        ScaleMode = 3
        Open (App.Path & "\readme.txt") For Input As #1
        Line Input #1, Tempstring
        .Height = (Val(Tempstring) * .TextHeight("Test Height")) + 300
        Do Until EOF(1)
            Line Input #1, Tempstring
            PrintText Tempstring
        Loop
        Close #1
        theleft = P2.ScaleLeft
        thetop = P2.ScaleHeight
        p1hgt = .ScaleHeight
        p1wid = .ScaleWidth
        Timer1.Enabled = True
        Timer1.Interval = 100
    End With
End Sub



Private Sub Label2_Click(Index As Integer)
Select Case Index
    Case 1
        OpenURL "mailto:mkprog@mkprog.own.pl"
    Case 0
        OpenURL "http://www.mkprog.own.pl"
End Select
End Sub

Sub Timer1_Timer()
Dim X As Long
Dim Txt As String
    With P2
        X = BitBlt(.hDC, theleft, thetop, p1wid, p1hgt, P1.hDC, 0, 0, &HCC0020)
        thetop = thetop - 1
        If thetop < -p1hgt Then
            Timer1.Enabled = False
            Txt = "Credits Completed"
            CurrentY = .ScaleHeight / 2
            CurrentX = (.ScaleWidth - .TextWidth(Txt)) / 2
            P2.Print Txt
        End If
    End With
End Sub

Sub PrintText(Text As String)
Dim X As Long
Dim Y As Long
    With P1
        .CurrentX = (P1.ScaleWidth / 2) - (P1.TextWidth(Text) / 2)
        .ForeColor = vbBlack
        X = .CurrentX
        Y = .CurrentY

        X = X + 1
        Y = Y + 1
        .CurrentX = X
        .CurrentY = Y

        .ForeColor = &H8080&
        P1.Print Text
        .ForeColor = &HC0FFFF
        .CurrentX = X - 1
        .CurrentY = Y - 1
        P1.Print Text
    End With
End Sub

'*******************************************************************************
' OpenURL (FUNCTION)
'
' PARAMETERS:
' (In)     - sFile    - String  -
' (In/Out) - vArgs    - Variant -
' (In/Out) - vShow    - Variant -
' (In/Out) - vInitDir - Variant -
' (In/Out) - vVerb    - Variant -
' (In/Out) - vhWnd    - Variant -
'
' RETURN VALUE:
' Long -
'
' DESCRIPTION:
' ***Description goes here***
'*******************************************************************************
Function OpenURL(ByVal sFile As String, Optional vArgs As Variant, Optional vShow As Variant, Optional vInitDir As Variant, Optional vVerb As Variant, Optional vhWnd As Variant) As Long
    '// Fill any empty optional arguments
    If IsMissing(vArgs) Then vArgs = vbNullString
    If IsMissing(vShow) Then vShow = vbNormalFocus
    If IsMissing(vInitDir) Then vInitDir = vbNullString
    If IsMissing(vVerb) Then vVerb = vbNullString
    If IsMissing(vhWnd) Then vhWnd = 0
    '// Call the dll
    OpenURL = ShellExecute(Me.hWnd, "Open", sFile, vbNullString, App.Path, vbNormalFocus)
    'MsgBox Shell(sFile)
End Function


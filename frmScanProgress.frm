VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScanProgress 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1080
   ClientLeft      =   408
   ClientTop       =   408
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   405
      Width           =   4515
      _ExtentX        =   7959
      _ExtentY        =   508
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   750
      Width           =   4515
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   4515
   End
End
Attribute VB_Name = "frmScanProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
'
'   Forest Software
'   Arlington, TX
'   http://www.flash.net/~dsfergus
'
'-----------------------------------------------------------------------------
'
'   Module: frmScanProgress.frm
'   Author: Roger Aikin
'
'   This code was graciously provied by Roger aikin who wanted a progress bar
'   and wrote this one before he passed it on to me.  Thanks Roger
'
'   $Header: $
'   $Archive: $
'   $Date: $
'   $Author:$
'   $Modtime:$
'
'-----------------------------------------------------------------------------
'
'   $Log: $
'
'=============================================================================
Option Explicit

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
  Me.lblProgress.Caption = sText
  'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
  Me.Visible = True
  Me.ZOrder 0
  InitProgress = True
ErrorHandler:
  'Exit without setting ReportProgress = True

End Function

Public Function RemoveProgress() As Boolean

  On Error GoTo ErrorHandler
  Unload Me
  RemoveProgress = True
ErrorHandler:
  'Exit without setting ReportProgress = True

End Function

':) Ulli's Code Formatter V2.0 (2001-01-23 10:53:25) 26 + 59 = 85 Lines

Public Function ReportChange(ByVal sText As String) As Boolean

  On Error GoTo ErrorHandler
    
  Me.Label1.Caption = sText
  Me.Label1.Refresh
  ReportChange = True
ErrorHandler:
  'Exit without setting ReportProgress = True

End Function

Public Function ReportProgress(Optional sText As String, Optional vValue As Long) As Boolean

  On Error GoTo ErrorHandler
  With Me
    .Show
    .pbProgress.Value = vValue
    .lblProgress.Caption = sText
    .pbProgress.Refresh
    .lblProgress.Refresh
  End With
  ReportProgress = True
ErrorHandler:
  'Exit without setting ReportProgress = True

End Function

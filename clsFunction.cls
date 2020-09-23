VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsName As String 'local copy
Private mvarlHowMany As Long 'local copy
Private mvarsModuleName As String 'local copy
Private mvarsCode As String 'local copy
Private mvarbScanForUse As Boolean 'local copy

Public Property Get bScanForUse() As Boolean

  bScanForUse = mvarbScanForUse

End Property

Public Property Let bScanForUse(ByVal vData As Boolean)

  mvarbScanForUse = vData

End Property

Public Property Get lHowMany() As Long

  lHowMany = mvarlHowMany

End Property

Public Property Let lHowMany(ByVal vData As Long)

  mvarlHowMany = vData

End Property

Public Property Get sCode() As String

  sCode = mvarsCode

End Property

Public Property Let sCode(ByVal vData As String)

  mvarsCode = vData

End Property

Public Property Get sModuleName() As String

  sModuleName = mvarsModuleName

End Property

Public Property Let sModuleName(ByVal vData As String)

  mvarsModuleName = vData

End Property

Public Property Get sName() As String

  sName = mvarsName

End Property

':) Ulli's Code Formatter V2.0 (2001-01-23 10:53:16) 7 + 62 = 69 Lines

Public Property Let sName(ByVal vData As String)

  mvarsName = vData

End Property

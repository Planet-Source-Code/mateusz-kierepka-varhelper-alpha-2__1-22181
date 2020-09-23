VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This is object which represents Visual Basic Variable"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"FastCollection"
Option Explicit

Private mvarsName           As String
Private mvarsScope          As String
Private mvarsType           As String
Private mvarsFunctionName   As String
Private mvarsModuleName     As String
Private mvarbConst          As Boolean
Private mvarvValue          As Variant
Private mvarsDescription    As String
Private mvarbWithEvents     As Boolean
Private mvarsDimension      As String
Private mvarbPreserve       As Boolean
Private mvarFastCollection  As FastCollection
Private mvarsComment        As String

Private Sub Class_Terminate()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  Set mvarFastCollection = Nothing
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Class_Terminate")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Public Function AddFunction(sName As String, _
       lHowMany As Long, _
       sModuleName As String, _
       Optional sKey As String)

 Dim cFunction As clsFunction

  '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  If mvarFastCollection Is Nothing Then
    Set mvarFastCollection = New FastCollection
  End If

  Set cFunction = New clsFunction
  With cFunction
    .lHowMany = lHowMany
    .sModuleName = sModuleName
    .sName = sName
  End With

  Call mvarFastCollection.Add(cFunction)
  Set cFunction = Nothing
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "AddFunction")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Function


Public Property Get bConst() As Boolean

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  bConst = mvarbConst
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get bConst")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Let bConst(ByVal vData As Boolean)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  mvarbConst = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let bConst")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Get bPreserve() As Boolean

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  'used when retrieving value of a property, on the right side of an assignment.
  'Syntax: Debug.Print X.bPreserve

  bPreserve = mvarbPreserve
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get bPreserve")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Let bPreserve(ByVal vData As Boolean)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  'used when assigning a value to the property, on the left side of an assignment.
  'Syntax: X.bPreserve = 5

  mvarbPreserve = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let bPreserve")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Get bWithEvents() As Boolean

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  bWithEvents = mvarbWithEvents
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get bWithEvents")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Let bWithEvents(ByVal vData As Boolean)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  mvarbWithEvents = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let bWithEvents")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Get FastCollection() As FastCollection

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  If mvarFastCollection Is Nothing Then
    Set mvarFastCollection = New FastCollection
  End If

  Set FastCollection = mvarFastCollection
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get FastCollection")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property
Public Property Set FastCollection(ByVal mNewFastCol As FastCollection)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  Set mvarFastCollection = mNewFastCol
   
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get FastCollection")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Get sComment() As String

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  sComment = mvarsComment
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get sComment")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Let sComment(ByVal vData As String)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  mvarsComment = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let sComment")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Get sDescription() As String

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  sDescription = mvarsDescription
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get sDescription")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Let sDescription(ByVal vData As String)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  mvarsDescription = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let sDescription")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Get sDimension() As String

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  sDimension = mvarsDimension
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get sDimension")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Let sDimension(ByVal vData As String)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  mvarsDimension = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let sDimension")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Get sFunctionName() As String

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  sFunctionName = mvarsFunctionName
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get sFunctionName")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Let sFunctionName(ByVal vData As String)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  mvarsFunctionName = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let sFunctionName")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Get sModuleName() As String

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  sModuleName = mvarsModuleName
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get sModuleName")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Let sModuleName(ByVal vData As String)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  mvarsModuleName = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let sModuleName")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Get sName() As String

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  sName = mvarsName
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get sName")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

':) Ulli's Code Formatter V2.0 (2001-02-02 23:10:28) 15 + 605 = 620 Lines

Public Property Let sName(ByVal vData As String)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  mvarsName = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let sName")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Get sScope() As String

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  sScope = mvarsScope
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get sScope")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Let sScope(ByVal vData As String)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  mvarsScope = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let sScope")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Get sType() As String

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  sType = mvarsType
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get sType")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Let sType(ByVal vData As String)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  mvarsType = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let sType")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

' Return True if a string is a valid name for a Visual Basic variable or
' procedure.
' This function turns to be useful when writing add-ins that generate code.
' The optional arguments specify whether type-declaration characters (%,&,!,#,@,
'  and $)
' should be considered as valid, and whether the dot character is accepted.

Public Function ValidateVBName(ByVal VBName As String, _
       Optional ByVal acceptTypeChars _
       As Boolean, Optional ByVal acceptDots As Boolean) As Boolean
    
 Dim i As Long
    
  ' reject null strings or strings that are too long
  If Len(VBName) = 0 Or Len(VBName) > 255 Then Exit Function
    
  ' check each character
  For i = 1 To Len(VBName)
    Select Case Asc(Mid$(VBName, i, 1))
     Case 65 To 90, 97 To 122
      ' alphabetic chars
     Case 48 To 57, 95
      ' digits and underscore, only good if not firs character
      If i = 1 Then Exit Function
     Case 37, 38, 33, 35, 64, 36
      ' type declaration characters
      If i < Len(VBName) Or Not acceptTypeChars Then Exit Function
     Case 46
      ' dot
      If i = 1 Or Not acceptDots Then Exit Function
     Case Else
      Exit Function
    End Select
  Next i
    
  ' all tests passed
  ValidateVBName = True
    
End Function


Public Property Get vValue() As Variant

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  If IsObject(mvarvValue) Then
    Set vValue = mvarvValue
   Else
    vValue = mvarvValue
  End If
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get vValue")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Let vValue(ByVal vData As Variant)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  mvarvValue = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Let vValue")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Set vValue(ByVal vData As Variant)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  Set mvarvValue = vData
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Set vValue")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

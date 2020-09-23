VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private m_ScanDone        As Boolean
Private m_VBInstance      As VBIDE.VBE

' *******************************************************
' Routine Name : (PRIVATE in CLASS) Sub Class_Initialize
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MKPROG]
' Date Writen  : 02-24-2001 -- 08:43:33
' Inputs       : N/A
' Outputs      : N/A
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Private Sub Class_Initialize()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  Set colVariables = New FastCollection
  
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Class_Initialize")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

' *******************************************************
' Routine Name : (PRIVATE in CLASS) Sub Class_Terminate
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MKPROG]
' Date Writen  : 02-24-2001 -- 08:43:27
' Inputs       : N/A
' Outputs      : N/A
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Private Sub Class_Terminate()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  Set colVariables = Nothing
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

' *******************************************************
' Routine Name : (PRIVATE in CLASS) Sub AddCodes
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MKPROG]
' Date Writen  : 02-24-2001 -- 08:43:11
' Inputs       : N/A
' Outputs      : N/A
' Description  : load code to fast collection
'              :
'              :
' Called By    :
' *******************************************************
Private Sub AddCodes()
 Dim cFunction         As clsFunction
 Dim vbComponentObj    As VBComponent
 Dim vbMemberObj       As Member
 Dim strKey            As String
 Dim lngProgressCnt    As Long
 Dim i                 As Long
 Dim lngProcedureType  As Long
 Dim strProcedureCode  As String
 Dim lngBodyStart      As Long
 Dim lngStartLine      As Long
 Dim lngLineCount      As Long
 Dim lngProcedureLineCount  As Long
 '{{{ Added It!

  On Error GoTo Generated_trap '}}}


  
  Set colCodes = New FastCollection
  
  '
  '   initialize the progres dialog
  
  Call frmScanProgress.InitProgress("Please Wait - collecting data...", 0, VBInstance.ActiveVBProject.VBComponents.Count - 1)
  
  For Each vbComponentObj In VBInstance.ActiveVBProject.VBComponents
    
    Call frmScanProgress.ReportProgress("Scanning " & vbComponentObj.Name & " for data.", lngProgressCnt)
    
    '   DO NOT Scan if this is a related document or a resource file
    '   There is no CodeModule component for this types
    If vbComponentObj.Type <> vbext_ct_ResFile And vbComponentObj.Type <> vbext_ct_RelatedDocument Then
      With vbComponentObj
      
        If .CodeModule.CountOfDeclarationLines > 0 Then
          
          Call AddFunction(.CodeModule.Lines(1, _
               .CodeModule.CountOfDeclarationLines), _
               .Name, _
               "")
        End If
        
        For Each vbMemberObj In vbComponentObj.CodeModule.Members
          '   The member type tells us if this is a function or a variable
          
          Call frmScanProgress.ReportChange("Scanning function: " & vbMemberObj.Name & " for data.")
          
          lngProcedureType = -1
          
          Select Case vbMemberObj.Type
           Case vbext_mt_Method, vbext_mt_Event
            lngProcedureType = vbext_pk_Proc
              
           Case vbext_mt_Property
            '
            '   This all necessary due to the fact that we have to knwo the proper type of Get/Let to use
            '   We are just getting one line now because we are doing this just to test and probe
            On Error Resume Next
             strProcedureCode = vbComponentObj.Lines(vbComponentObj.ProcStartLine(vbMemberObj.Name, vbext_pk_Get), 1)
             If Err.Number = 35 Then ' Sub or Function not defined
               Err.Clear
               strProcedureCode = vbComponentObj.Lines(vbComponentObj.ProcStartLine(vbMemberObj.Name, vbext_pk_Let), 1)
               If Err.Number = 35 Then ' Sub or Function not defined
                 Err.Clear
                 strProcedureCode = vbComponentObj.Lines(vbComponentObj.ProcStartLine(vbMemberObj.Name, vbext_pk_Set), 1)
                 If Err.Number = 0 Then
                   lngProcedureType = vbext_pk_Set
                  Else
                   ' could not determine the type...
                   lngProcedureType = -1
                 End If
                Else
                 lngProcedureType = vbext_pk_Let
               End If
              Else
               lngProcedureType = vbext_pk_Get
             End If
            On Error GoTo 0
          End Select
          '
          '   If we could not determine the type or it is not a type we wnat, skip on outta here
          If lngProcedureType <> -1 Then
            '
            '   We are going to inspect each function as a whole for the presence
            '   of any procedure calls.  To do so we need to get information about
            '   the code for each procedure.
            '
            On Error Resume Next
             '   First we will get the starting lines and line count of the procedure
             lngBodyStart = vbComponentObj.CodeModule.ProcBodyLine(vbMemberObj.Name, lngProcedureType)
             lngStartLine = vbComponentObj.CodeModule.ProcStartLine(vbMemberObj.Name, lngProcedureType)
             lngLineCount = vbComponentObj.CodeModule.ProcCountLines(vbMemberObj.Name, lngProcedureType)
             '
             '   ProcCountLInes returns the number of lines INCLUDING any comments before the actual
             '   start of the code.  We don't want those so we calculate a new line count which only
             '   includes the actual body of code
             lngProcedureLineCount = lngLineCount - (lngBodyStart - lngStartLine)
             '
             '   But, as we look through the code for occurances of functions, the entry point (as MS calls it)
             '   will cause false hits so lets skip the first line
             lngBodyStart = lngBodyStart + 1
             '
             '   And we might as well skip the End Sub line
             lngProcedureLineCount = lngProcedureLineCount - 2
                
             '   Now, let's actually get the code
             '   obj.Lines (startline, count)
            
             Call AddFunction(vbComponentObj.CodeModule.Lines _
                  (lngBodyStart, lngProcedureLineCount), _
                  .Name, vbMemberObj.Name)
             
            On Error GoTo 0
          End If
          
        Next vbMemberObj
        
        Set vbMemberObj = Nothing
      End With
    End If
    lngProgressCnt = lngProgressCnt + 1
  Next vbComponentObj
   
  frmScanProgress.Hide
  
  On Error GoTo 0
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "AddCodes")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub


' *******************************************************
' Routine Name : (PRIVATE in CLASS) Sub AddFunction
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MKPROG]
' Date Writen  : 02-24-2001 -- 08:43:03
' Inputs       : ByRef sCode:String -
'              : ByRef sModuleName:String -
'              : ByRef sFunctionName:String -
' Outputs      : N/A
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Private Sub AddFunction(ByRef sCode As String, _
                        ByRef sModuleName As String, _
                        ByRef sFunctionName As String)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
                       
 Dim cFunction As clsFunction

  Set cFunction = New clsFunction
  With cFunction
    .sCode = sCode
    .sModuleName = sModuleName
    .sName = sFunctionName
  End With
  colCodes.Add cFunction, sModuleName & sFunctionName
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

End Sub


' *******************************************************
' Routine Name : (PRIVATE in CLASS) Function CheckConst
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MKPROG]
' Date Writen  : 02-24-2001 -- 08:43:00
' Inputs       : ByRef sCheckThis:String -
'              : ByVal bModule:Boolean -
'              : ByRef sFunction:String -
'              : ByRef sModuleName:String -
' Outputs      : N/A
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Private Function CheckConst(ByRef sCheckThis As String, _
                            ByVal bModule As Boolean, _
                            ByRef sFunction As String, _
                            ByRef sModuleName As String)

 Dim cVariable    As clsVariable
 Dim sParse       As String
 
 Dim clsRegExp   As RegExp
 Dim clsMatchCol As MatchCollection
 Dim clsMatch    As Match
 Dim strResult   As String
 Dim strTemp     As String
 Dim i           As Long

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

 
  Set cVariable = New clsVariable
  
  Set clsRegExp = New RegExp
   
  strResult = sCheckThis

  '-----------------------------
  'first set regular expresion
  '-----------------------------
  With clsRegExp
    .Global = True
    .IgnoreCase = True
    '-----------------------------
    'comments
    '-----------------------------
    .Pattern = "\'{1}.*"
    Set clsMatchCol = .Execute(strResult)
    
    For Each clsMatch In clsMatchCol
      With clsMatch
        Debug.Print .Value
        cVariable.sComment = Trim$(.Value)
      End With
    Next clsMatch
    
    strResult = .Replace(strResult, " ")
    '-----------------------------
    'remove const keyword
    ' - first check maybe there
    '   are more then 1 const
    '   declaration
    '-----------------------------
    .Pattern = "\s?(const){1}\s"
    Set clsMatchCol = .Execute(strResult)
    
    '-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
    ' ! -> R E C U R E N C Y if more const then 1 <- !
    '-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
    If clsMatchCol.Count > 1 Then
      
      Call CheckConst(Right$(sCheckThis, Len(sCheckThis) - _
           (clsMatchCol.Item(0).FirstIndex + clsMatchCol.Item(0).Length)), bModule, _
           sFunction, sModuleName)
      sCheckThis = Left$(sCheckThis, clsMatchCol.Item(0).FirstIndex)
    End If
    
    If clsMatchCol.Count = 0 Then
      Exit Function
    End If
    
    '-----------------------------
    'WithEvents
    '-----------------------------
    .Pattern = "\s?(withevents){1}\s"
    Set clsMatchCol = .Execute(strResult)
    
    If clsMatchCol.Count > 0 Then
      cVariable.bWithEvents = True
      strResult = .Replace(strResult, " ")
      '-----------------------------
      'too many public
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many WithEvents keyword in const declaration." & vbCrLf
      End If
     Else
      cVariable.bWithEvents = False
    End If
    
    '-----------------------------
    'public with
    '-----------------------------
    .Pattern = "\s?(public){1}\s"
    Set clsMatchCol = .Execute(strResult)
    
    If clsMatchCol.Count > 0 Then
      cVariable.sScope = "public"
      strResult = .Replace(strResult, " ")
      '-----------------------------
      'too many public
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many public keyword in const declaration." & vbCrLf
      End If
    End If
    
    '-----------------------------
    'private
    '-----------------------------
    .Pattern = "\s?(private){1}\s"
    Set clsMatchCol = .Execute(strResult)
    If clsMatchCol.Count > 0 Then
      strResult = .Replace(strResult, " ")
      '-----------------------------
      'too many private
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many Private keyword in const declaration." & vbCrLf
      End If
      '----------------------------------------------------------
      'someone declared const with private and public keyword
      '----------------------------------------------------------
      If Len(cVariable.sScope) > 0 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Private and public keyword in const declaration." & vbCrLf
        cVariable.sScope = "Error"
       Else
        cVariable.sScope = "Private"
      End If
    End If
    '-----------------------------
    'global
    '-----------------------------
    .Pattern = "\s?(global){1}\s"
    Set clsMatchCol = .Execute(strResult)
    If clsMatchCol.Count > 0 Then
      strResult = .Replace(strResult, " ")
      '-----------------------------
      'too many private
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many Global keyword in const declaration." & vbCrLf
      End If
      '----------------------------------------------------------
      'someone declared const with private and public keyword
      '----------------------------------------------------------
      If Len(cVariable.sScope) > 0 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Global and Private or public keyword in const declaration." & vbCrLf
        cVariable.sScope = "Error"
       Else
        cVariable.sScope = "Global"
      End If
      cVariable.sDescription = cVariable.sDescription & _
                               "Global keyword is obslote. You should use public instead." & vbCrLf
    End If
    
    '-----------------------------
    ' Scope declared?
    '-----------------------------
    If Len(cVariable.sScope) = 0 Then
    
      cVariable.sDescription = cVariable.sDescription & _
                               "You should add Private or public keyword for constant to declare scope." & vbCrLf
      
      If bModule Then
        cVariable.sScope = "Not declared - public"
       Else
        cVariable.sScope = "Not declared - Private"
      End If
    End If
    
    '-----------------------------
    'find type of constant
    '-----------------------------
    .Pattern = "\s?(as){1}\s"
    Set clsMatchCol = .Execute(strResult)
    
    If clsMatchCol.Count > 0 Then
      '-----------------------------
      'too many as
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many As keyword in const declaration." & vbCrLf
      End If
      
      '-----------------------------
      'find =
      '-----------------------------
      i = InStr(clsMatchCol.Item(0).FirstIndex + 2, strResult, "=")
      If i > 0 Then
        cVariable.sType = Trim$(Mid$(strResult, _
                          clsMatchCol.Item(0).FirstIndex + 4, _
                          i - (clsMatchCol.Item(0).FirstIndex + 4)))
        
        cVariable.vValue = Right$(strResult, Len(strResult) - i)
        
       Else
        cVariable.sType = Trim$(Right$(strResult, i - _
                          (clsMatchCol.Item(0).FirstIndex + 2)))
        cVariable.sDescription = cVariable.sDescription & _
                                 "Const not declared." & vbCrLf
      End If
      '-------------------------------
      'remove as keyword and find name
      '-------------------------------
      strResult = Trim$(Left$(strResult, clsMatchCol.Item(0).FirstIndex))
      cVariable.sName = Replace$(strResult, "=", "")
     Else
      '-----------------------------
      'without as
      '-----------------------------
      cVariable.sDescription = cVariable.sDescription & "You should specify type of const [As 'Data type'] ." & vbCrLf
      cVariable.sType = "Not declared - Variant"
      '-----------------------------
      ' Find =
      '-----------------------------
      i = InStr(1, strResult, "=")
      If i > 0 Then
        
        cVariable.vValue = Right$(strResult, Len(strResult) - (i + 1))
        
       Else
        cVariable.sDescription = cVariable.sDescription & _
                                 "Const not declared." & vbCrLf
      End If
      '-------------------------------
      'remove as keyword and find name
      '-------------------------------
      strResult = Trim$(Left$(strResult, i))
    End If
    
  End With
  With cVariable
    'remove "
    clsRegExp.Global = True
    clsRegExp.IgnoreCase = True
    clsRegExp.Pattern = """{1}"
    .vValue = clsRegExp.Replace(.vValue, "")
    'remove =
    .vValue = Trim$(Replace$(cVariable.vValue, "=", ""))
    'remove vbcrlf
    .vValue = Trim$(Replace$(cVariable.vValue, vbCrLf, ""))
    'remove =
    .sName = Trim$(Replace$(strResult, "=", ""))
    .sName = Trim$(Replace$(.sName, vbCrLf, ""))
    .sName = Trim$(Replace$(.sName, vbCr, ""))
    .sName = Trim$(Replace$(.sName, vbLf, ""))
    
    .sFunctionName = sFunction
    .sModuleName = sModuleName
    .bConst = True
  End With
  '-----------------------------
  'realese objects
  '-----------------------------
  Set clsRegExp = Nothing
  Set clsMatchCol = Nothing

  colVariables.Add cVariable
  
  Set cVariable = Nothing
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "CheckConst")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Function


' *******************************************************
' Routine Name : (PRIVATE in CLASS) Function CheckType
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MKPROG]
' Date Writen  : 02-24-2001 -- 08:42:54
' Inputs       : ByRef sModuleName:String -
'              : ByRef sFunction:String -
'              : ByRef sCheckThis:String -
' Outputs      : N/A
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Private Function CheckType(ByRef sModuleName As String, _
                           ByRef sFunction As String, _
                           ByRef sCheckThis As String)

 Dim sParse       As String
 Dim bModule      As Boolean
 Dim clsRegExp    As RegExp
 Dim clsMatchCol  As MatchCollection
 Dim cVariable    As clsVariable

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

 
  Set cVariable = New clsVariable
  
  Set clsRegExp = New RegExp
    
  '-----------------------------
  'first set regular expresion
  '-----------------------------
  With clsRegExp
    .Global = True
    .IgnoreCase = True
       
    '-----------------------------
    'Remove not nessesery data
    '-----------------------------
    .Pattern = "\s?(declare|enum|type|event|sub){1}\s+"
    Set clsMatchCol = .Execute(sCheckThis)
  End With
  
  If clsMatchCol.Count = 0 Then

    If Len(sFunction) > 0 Then
      bModule = False
     Else
      bModule = True
    End If
    
    clsRegExp.Pattern = "\s?(const){1}\s"
    Set clsMatchCol = clsRegExp.Execute(sCheckThis)
    If clsMatchCol.Count = 0 Then
      Call CheckVariable(sCheckThis, bModule, sFunction, sModuleName)
     Else
      Call CheckConst(sCheckThis, bModule, sFunction, sModuleName)
    End If

  End If

  Set clsMatchCol = Nothing
  Set clsRegExp = Nothing
  
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "CheckType")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Function


' *******************************************************
' Routine Name : (PRIVATE in CLASS) Function CheckVariable
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MKPROG]
' Date Writen  : 02-24-2001 -- 08:42:50
' Inputs       : ByRef sCheckThis:String -
'              : ByVal bModule:Boolean -
'              : ByRef sFunction:String -
'              : ByRef sModuleName:String -
' Outputs      : N/A
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Private Function CheckVariable(ByRef sCheckThis As String, _
                               ByVal bModule As Boolean, _
                               ByRef sFunction As String, _
                               ByRef sModuleName As String)

 Dim cVariable    As clsVariable
 Dim sParse       As String
 
 Dim clsRegExp   As RegExp
 Dim clsMatchCol As MatchCollection
 Dim clsMatch    As Match
 Dim strResult   As String
 Dim strTemp     As String
 Dim i           As Long
 Dim j           As Long
 Dim k           As Long


 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

 
  Set cVariable = New clsVariable
  
  Set clsRegExp = New RegExp
   
  strResult = sCheckThis
  
  '-----------------------------
  'first set regular expresion
  '-----------------------------
  
  With clsRegExp
    .Global = True
    .IgnoreCase = True
    
    '-----------------------------
    'comments
    '-----------------------------
    .Pattern = "\'{1}.*\n"
    Set clsMatchCol = .Execute(strResult)
    
    For Each clsMatch In clsMatchCol
      With clsMatch
        Debug.Print .Value
        cVariable.sComment = Trim$(.Value)
      End With
    Next clsMatch
    
    strResult = .Replace(strResult, " ")
            
    '-----------------------------
    'WithEvents
    '-----------------------------
    .Pattern = "\s?(withevents){1}\s"
    Set clsMatchCol = .Execute(strResult)
    
    If clsMatchCol.Count > 0 Then
      cVariable.bWithEvents = True
      strResult = .Replace(strResult, " ")
      '-----------------------------
      'too many WithEvents
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many WithEvents keyword in variable declaration." & vbCrLf
      End If
     Else
      cVariable.bWithEvents = False
    End If
    
    '-----------------------------
    'Preserve
    '-----------------------------
    .Pattern = "\s?(preserve){1}\s"
    Set clsMatchCol = .Execute(strResult)
    
    If clsMatchCol.Count > 0 Then
      cVariable.bPreserve = True
      strResult = .Replace(strResult, " ")
      '-----------------------------
      'too many Preserve
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many Preserve keyword in variable declaration." & vbCrLf
      End If
     Else
      cVariable.bPreserve = False
    End If
    
    '-----------------------------
    'public with
    '-----------------------------
    .Pattern = "\s?(public){1}\s"
    Set clsMatchCol = .Execute(strResult)
    
    If clsMatchCol.Count > 0 Then
      cVariable.sScope = "public"
      strResult = .Replace(strResult, " ")
      '-----------------------------
      'too many public
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many public keyword in variable declaration." & vbCrLf
      End If
    End If
    
    '-----------------------------
    'private
    '-----------------------------
    .Pattern = "\s?(private){1}\s"
    Set clsMatchCol = .Execute(strResult)
    If clsMatchCol.Count > 0 Then
      strResult = .Replace(strResult, " ")
      '-----------------------------
      'too many private
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many Private keyword in variable declaration." & vbCrLf
      End If
      '----------------------------------------------------------
      'someone declared variable with private and public keyword
      '----------------------------------------------------------
      If Len(cVariable.sScope) > 0 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Private and public keyword in variable declaration." & vbCrLf
        cVariable.sScope = "Error"
       Else
        cVariable.sScope = "Private"
      End If
    End If
    '-----------------------------
    'global
    '-----------------------------
    .Pattern = "\s?(global){1}\s"
    Set clsMatchCol = .Execute(strResult)
    If clsMatchCol.Count > 0 Then
      strResult = .Replace(strResult, " ")
      '-----------------------------
      'too many Global
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many Global keyword in variable declaration." & vbCrLf
      End If
      '----------------------------------------------------------
      'someone declared variable with Global and public or
      'private keyword
      '----------------------------------------------------------
      If Len(cVariable.sScope) > 0 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Global and Private or public keyword in variable declaration." & vbCrLf
        cVariable.sScope = "Error"
       Else
        cVariable.sScope = "Global=public"
      End If
      cVariable.sDescription = cVariable.sDescription & _
                               "Global keyword is obslote. You should use public instead." & vbCrLf
    End If
    
    '-----------------------------
    'Dim
    '-----------------------------
    .Pattern = "\s?(Dim){1}\s"
    Set clsMatchCol = .Execute(strResult)
    If clsMatchCol.Count = 0 Then
      .Pattern = "\s?(redim){1}\s"
      Set clsMatchCol = .Execute(strResult)
    End If

    If clsMatchCol.Count > 0 Then
      strResult = .Replace(strResult, " ")
      '-----------------------------
      'too many Global
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many Dim keyword in variable declaration." & vbCrLf
      End If
      '----------------------------------------------------------
      'someone declared variable with Dim ,
      'with private, public or global keyword
      '----------------------------------------------------------
      If Len(cVariable.sScope) > 0 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Dim and Private or public keyword in variable declaration." & vbCrLf
        cVariable.sScope = "Error"
       Else
        If bModule Then
          cVariable.sScope = "public"
         Else
          cVariable.sScope = "Private"
        End If
        
      End If
      
    End If
    
    '-----------------------------
    'find other variables
    '-----------------------------
    i = InStr(1, strResult, ", ")
    If i > 0 Then
      j = InStr(1, Left$(strResult, i), "(")
      '-----------------------------
      'find dimension
      '-----------------------------
      If j > 0 Then
        'there are bracket - check for dimension
        k = InStr(1, strResult, ") ")
        If k > 0 Then
          Call CheckVariable(cVariable.sScope & _
               " " & Right$(strResult, _
               (Len(strResult) - k) + 1), _
               bModule, sFunction, sModuleName)
          strResult = Left$(strResult, k)
        End If
       Else
        'no dimension, check other variables
        '-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        ' ! -> R E C U R E N C Y <- !
        '-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

        Call CheckVariable(cVariable.sScope & " " & Right$(strResult, _
             (Len(strResult) - i)), _
             bModule, sFunction, sModuleName)
                    
        strResult = Left$(strResult, i - 1)
      End If
    End If
    '-----------------------------
    'find type of variable
    '-----------------------------
    .Pattern = "\s?(as){1}\s"
    Set clsMatchCol = .Execute(strResult)
    
    If clsMatchCol.Count > 0 Then
      '-----------------------------
      'too many as
      '-----------------------------
      If clsMatchCol.Count > 1 Then
        cVariable.sDescription = cVariable.sDescription & _
                                 "Too many As keyword in variable declaration." & vbCrLf
      End If
      
      cVariable.sType = Trim$(Mid$(strResult, _
                        clsMatchCol.Item(0).FirstIndex + 4, _
                        Len(strResult) - (clsMatchCol.Item(0).FirstIndex + 3)))
      
      '-------------------------------
      'remove as keyword and find name
      '-------------------------------
      strResult = Trim$(Left$(strResult, clsMatchCol.Item(0).FirstIndex))
      
     Else
      '-----------------------------
      'without as
      '-----------------------------
      cVariable.sDescription = cVariable.sDescription & "You should specify type of variable [As 'Data type'] ." & vbCrLf
      cVariable.sType = "Not declared - Variant"
      
    End If
    
    '-----------------------------
    'Find any dimension settings
    '-----------------------------
    i = InStr(1, strResult, "(")
    
    If i > 0 Then
      j = InStr(1, strResult, ")")
      If i + 1 = j Then
        cVariable.sDimension = "Dynamic array"
       Else
        cVariable.sDimension = Trim$(Mid$(strResult, i + 1, j - i - 1))
      End If
      cVariable.sName = Trim$(Left$(strResult, i - 1))
     Else
      cVariable.sName = Trim$(strResult)
    End If
        
  End With
  
  With cVariable
    .sFunctionName = sFunction
    .sModuleName = sModuleName
    .bConst = False
    
    'remove vbcrlf
    .sName = Trim$(Replace$(.sName, vbCrLf, ""))
    .sName = Trim$(Replace$(.sName, vbCr, ""))
    .sName = Trim$(Replace$(.sName, vbLf, ""))
    'remove space
    .sName = Trim$(Replace$(.sName, " ", ""))
    
    'remove vbcrlf
    .sType = Trim$(Replace$(.sType, vbCrLf, ""))
    'remove space
    .sType = Trim$(Replace$(.sType, " ", ""))
    
  End With
  

  colVariables.Add cVariable

  '-----------------------------
  'realese objects
  '-----------------------------
  Set clsRegExp = Nothing
  Set clsMatchCol = Nothing

  Set cVariable = Nothing
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "CheckVariable")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Function


' *******************************************************
' Routine Name : (PRIVATE in CLASS) Function FormatXML
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MKPROG]
' Date Writen  : 02-24-2001 -- 08:42:44
' Inputs       : ptrNode:IXMLDOMNode -
'              : Optional iTabLevel:Integer = 0 -
' Outputs      : String --
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Private Function FormatXML(ptrNode As IXMLDOMNode, Optional iTabLevel As Integer = 0) As String
Dim bMixedTextNode As Boolean
Dim bHasOnlyATextNode As Boolean
Dim i As Integer

    With ptrNode
        Select Case .nodeType
        Case NODE_DOCUMENT, NODE_DOCUMENT_FRAGMENT
            'all child nodes of the document should be at the same indent Level
            'just iterate over them and recurse with 0 indent
            For i = 0 To .childNodes.Length - 1
                FormatXML = FormatXML & FormatXML(.childNodes(i))
            Next i
        
        Case NODE_TEXT 'should render the same way the default IE5 stylesheet does for mixed content
            'figure out if we're in some mixed content
            bMixedTextNode = (.parentNode.childNodes.Length > 1)  'if this text node has any siblings it's in mixed content
            
            'if mixed indent this string
            If bMixedTextNode Then FormatXML = String(iTabLevel, vbTab)
            'we're gonna strip out any tabs and carriage returns from the Text
            FormatXML = FormatXML & Trim(Replace(Replace(.xml, vbCrLf, " "), vbTab, " "))
            'if mixed add cariage return
            If bMixedTextNode Then FormatXML = FormatXML & vbCrLf

        Case NODE_ELEMENT
            If .hasChildNodes Then
                'if the node has only one child and that child is text we won't add carriage return after opening tag
                bHasOnlyATextNode = (.childNodes(0).nodeType = NODE_TEXT) And (.childNodes.Length = 1)
            End If
            
            'open the start tag
            FormatXML = String(iTabLevel, vbTab) & "<" & .nodeName

            'recurse over the attributes
            For i = 0 To .Attributes.Length - 1
                FormatXML = FormatXML + FormatXML(.Attributes(i))
            Next i

            'properly close the start tag based on node's contents
            If Not .hasChildNodes Then       'no child nodes so it's an empty element
                FormatXML = FormatXML & "/>" & vbCrLf
                
            Else
                If bHasOnlyATextNode Then   'has only text for children - don't add carriage return
                    FormatXML = FormatXML & ">"
                Else                            'has child elements - add carriage return
                    FormatXML = FormatXML & ">" & vbCrLf
                End If
                
                'recurse if there's children
                For i = 0 To .childNodes.Length - 1
                    FormatXML = FormatXML & FormatXML(.childNodes(i), iTabLevel + 1)
                Next i
                
                'properly indent and add the end tag
                If Not bHasOnlyATextNode Then FormatXML = FormatXML & String(iTabLevel, vbTab)
                FormatXML = FormatXML & "</" & .nodeName & ">" & vbCrLf
                
            End If
                    
        Case NODE_COMMENT, NODE_CDATA_SECTION
            'if comment is on more than one line don't indent
            If InStr(1, .xml, vbCr) = 0 Then FormatXML = String(iTabLevel, vbTab)
            FormatXML = FormatXML & .xml & vbCrLf
        
        Case NODE_ATTRIBUTE
            'if there are double quotes in the attribute use single quotes to surrond the attr value
            If InStr(1, .Text, Chr(34)) > 0 Then
                FormatXML = " " & .nodeName & "='" & .Text & "'"
            Else
                FormatXML = " " & .nodeName & "=" & Chr(34) & .Text & Chr(34)
            End If
        
        Case NODE_ENTITY
            'and we would never want entites expanded
            
        Case Else
            'all other node types should just return their xml (properly indented)
            'these include - entity refs, pi's, notations, doctypes
            FormatXML = String(iTabLevel, vbTab) & .xml & vbCrLf
            
        End Select
        
    End With

End Function

'Special characters and sequences are used in writing patterns for regular expressions. The following table describes these characters and includes short examples showing how the characters are used.
'Character Description
'\  Marks the next character as special. /n/ matches the character "n". The sequence /\n/ matches a linefeed or newline character.
'^  Matches the beginning of input or line.
'$  Matches the end of input or line.
'*  Matches the preceding character zero or more times. /zo*/ matches either "z" or "zoo."
'+  Matches the preceding character one or more times. /zo+/ matches "zoo" but not "z."
'?  Matches the preceding character zero or one time. /a?ve?/ matches the "ve" in "never."
'. Matches any single character except a newline character.
'(pattern)  Matches pattern and remembers the match. The matched substring can be retrieved from the result Array object elements [1]...[n] or the RegExp object's $1...$9 properties. To match parentheses characters ( ), use "\(" or "\)".
'x|y Matches either x or y. /z|food?/ matches "zoo" or "food."
'{n} n is a nonnegative integer. Matches exactly n times. /o{2}/ does not match the "o" in "Bob," but matches the first two o's in "foooood."
'{n,}  n is a nonnegative integer. Matches at least n times. /o{2,}/ does not match the "o" in "Bob" and matches all the o's in "foooood." /o{1,}/ is equivalent to /o+/.
'{n,m}  m and n are nonnegative integers. Matches at least n and at most m times. /o{1,3}/ matches the first three o's in "fooooood."
'[xyz]  A character set. Matches any one of the enclosed characters. /[abc]/ matches the "a" in "plain."
'[^xyz]  A negative character set. Matches any character not enclosed. /[^abc]/ matches the "p" in "plain."
'\b  Matches a word boundary, such as a space. /ea*r\b/ matches the "er" in "never early."
'\B  Matches a nonword boundary. /ea*r\B/ matches the "ear" in "never early."
'\d  Matches a digit character. Equivalent to [0-9].
'\D  Matches a nondigit character. Equivalent to [^0-9].
'\f  Matches a form-feed character.
'\n  Matches a linefeed character.
'\r  Matches a carriage return character.
'\s  Matches any white space including space, tab, form-feed, and so on. Equivalent to [ \f\n\r\t\v]
'\S  Matches any nonwhite space character. Equivalent to [^ \f\n\r\t\v]
'\t  Matches a tab character.
'\v  Matches a vertical tab character.
'\w  Matches any word character including underscore. Equivalent to [A-Za-z0-9_].
'\W  Matches any nonword character. Equivalent to [^A-Za-z0-9_].
'\num  Matches num, where num is a positive integer. A reference back to remembered matches. \1 matches what is stored in RegExp.$1.
'/n/  Matches n, where n is an octal, hexadecimal, or decimal escape value. Allows embedding of ASCII codes into regular expressions.

':) Ulli's Code Formatter V2.0 (2001-01-23 10:53:15) 4 + 1506 = 1510 Lines

' *******************************************************
' Routine Name : (PUBLIC in CLASS) Sub PrepareForScan
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MKPROG]
' Date Writen  : 02-24-2001 -- 08:42:38
' Inputs       : N/A
' Outputs      : N/A
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Public Sub PrepareForScan()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

 Dim vbComponentObj    As VBComponent
 Dim vbMemberObj       As Member
 Dim strKey            As String
 Dim lngProgressCnt    As Long
 Dim i                 As Long
 Dim lngProcedureType  As Long
 Dim strProcedureCode  As String
 Dim lngBodyStart      As Long
 Dim lngStartLine      As Long
 Dim lngLineCount      As Long
 Dim lngProcedureLineCount  As Long

  '
  '   initialize the progres dialog
  
  Call frmScanProgress.InitProgress("Please Wait...", 0, VBInstance.ActiveVBProject.VBComponents.Count - 1)
   
  'now scan for use variables
  '(\s|\(|\+|\-|\*|\&|\/|\\|^|\=){1}varname(\s|\)|\+|\-|\*|\&|\/|\\|^|\=){0,1}
  For Each vbComponentObj In VBInstance.ActiveVBProject.VBComponents
    Call frmScanProgress.ReportProgress("Adding " & vbComponentObj.Name & " .", lngProgressCnt)
    '   DO NOT Scan if this is a related document or a resource file
    '   There is no CodeModule component for this types
    If vbComponentObj.Type <> vbext_ct_ResFile And vbComponentObj.Type <> vbext_ct_RelatedDocument Then
      With vbComponentObj
      
        For Each vbMemberObj In vbComponentObj.CodeModule.Members
          
          Call frmScanProgress.ReportChange("Scanning function: " & vbMemberObj.Name & " for variable names")
          '   The member type tells us if this is a function or a variable
              
          lngProcedureType = -1
          
          Select Case vbMemberObj.Type
           Case vbext_mt_Method, vbext_mt_Event
            lngProcedureType = vbext_pk_Proc
              
           Case vbext_mt_Property
            '
            '   This all necessary due to the fact that we have to knwo the proper type of Get/Let to use
            '   We are just getting one line now because we are doing this just to test and probe
            On Error Resume Next
             strProcedureCode = vbComponentObj.Lines(vbComponentObj.ProcStartLine(vbMemberObj.Name, vbext_pk_Get), 1)
             If Err.Number = 35 Then ' Sub or Function not defined
               Err.Clear
               strProcedureCode = vbComponentObj.Lines(vbComponentObj.ProcStartLine(vbMemberObj.Name, vbext_pk_Let), 1)
               If Err.Number = 35 Then ' Sub or Function not defined
                 Err.Clear
                 strProcedureCode = vbComponentObj.Lines(vbComponentObj.ProcStartLine(vbMemberObj.Name, vbext_pk_Set), 1)
                 If Err.Number = 0 Then
                   lngProcedureType = vbext_pk_Set
                  Else
                   ' could not determine the type...
                   lngProcedureType = -1
                 End If
                Else
                 lngProcedureType = vbext_pk_Let
               End If
              Else
               lngProcedureType = vbext_pk_Get
             End If
            On Error GoTo 0
          End Select
          '
          '   If we could not determine the type or it is not a type we wnat, skip on outta here
          If lngProcedureType <> -1 Then
            '
            '   We are going to inspect each function as a whole for the presence
            '   of any procedure calls.  To do so we need to get information about
            '   the code for each procedure.
            '
            On Error Resume Next
             '   First we will get the starting lines and line count of the procedure
             lngBodyStart = vbComponentObj.CodeModule.ProcBodyLine(vbMemberObj.Name, lngProcedureType)
             lngStartLine = vbComponentObj.CodeModule.ProcStartLine(vbMemberObj.Name, lngProcedureType)
             lngLineCount = vbComponentObj.CodeModule.ProcCountLines(vbMemberObj.Name, lngProcedureType)
             '
             '   ProcCountLInes returns the number of lines INCLUDING any comments before the actual
             '   start of the code.  We don't want those so we calculate a new line count which only
             '   includes the actual body of code
             lngProcedureLineCount = lngLineCount - (lngBodyStart - lngStartLine)
             '
             '   But, as we look through the code for occurances of functions, the entry point (as MS calls it)
             '   will cause false hits so lets skip the first line
             lngBodyStart = lngBodyStart + 1
             '
             '   And we might as well skip the End Sub line
             lngProcedureLineCount = lngProcedureLineCount - 2
                
             '   Now, let's actually get the code
             '   obj.Lines (startline, count)
             strProcedureCode = Trim$(vbComponentObj.CodeModule.Lines(lngBodyStart, lngProcedureLineCount))
            
             If Len(strProcedureCode) > 0 Then
               ScanVariablesUSE .Name, vbMemberObj.Name, strProcedureCode
             End If
            On Error GoTo 0
          End If
          
        Next vbMemberObj
        
        Set vbMemberObj = Nothing
      End With
    End If
    lngProgressCnt = lngProgressCnt + 1
  Next vbComponentObj
   
  frmScanProgress.Hide
  
  On Error GoTo 0
  
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "PrepareForScan")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

' *******************************************************
' Routine Name : (PUBLIC in CLASS) Sub SaveXML
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MKPROG]
' Date Writen  : 02-24-2001 -- 08:42:28
' Inputs       : ByRef sReportName:String -
' Outputs      : N/A
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Public Sub SaveXML(ByRef sReportName As String)

 Dim gXml      As DOMDocument
 Dim xlist     As IXMLDOMNode
 Dim xnode     As IXMLDOMElement
 Dim xnodeCol  As IXMLDOMElement
 Dim xroot     As IXMLDOMElement

 Dim cVariable        As clsVariable
 Dim cFunction        As clsFunction
 Dim lngProgressCnt   As Long
 Dim sStrReport       As String
 Dim lngFileNumber    As Long
 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  'create xml document
 
  
  Call frmScanProgress.InitProgress("Please Wait...", 0, colVariables.Count + 10)

  Set gXml = New DOMDocument

  With gXml
    .async = False
    
    Set xroot = .createElement("root")
  
    Set .documentElement = xroot
    lngProgressCnt = 0
    For Each cVariable In colVariables
      lngProgressCnt = lngProgressCnt + 1
      Call frmScanProgress.ReportProgress("Saving variable: " & cVariable.sName, lngProgressCnt)
      
      Set xnode = .createElement("Variable")
      xroot.appendChild xnode

      With xnode
      
        .setAttribute "sName", cVariable.sName
        .setAttribute "sScope", cVariable.sScope
        .setAttribute "sType", cVariable.sType
        .setAttribute "sModuleName", cVariable.sModuleName
        .setAttribute "sFunctionName", cVariable.sFunctionName
        .setAttribute "bConst", cVariable.bConst
        .setAttribute "vValue", cVariable.vValue
        .setAttribute "bWithEvents", cVariable.bWithEvents
        .setAttribute "bPreserve", cVariable.bPreserve
        .setAttribute "sDimension", cVariable.sDimension
        .setAttribute "sDimension", cVariable.sDimension
        .setAttribute "sDescription", cVariable.sDescription
      End With
      
      For Each cFunction In cVariable.FastCollection
        Set xnodeCol = .createElement("Functions")
        With xnodeCol
          Call frmScanProgress.ReportChange("Saving use in function: " & cFunction.sName)
          xnode.appendChild xnodeCol
          .setAttribute "sModuleName", cFunction.sModuleName
          .setAttribute "sFunctionName", cFunction.sName
          .setAttribute "lHowMany", cFunction.lHowMany
          
        End With
      Next cFunction
      Set cFunction = Nothing
    Next cVariable
    
    Set cVariable = Nothing
    
    On Error Resume Next
    lngProgressCnt = lngProgressCnt + 1
    Call frmScanProgress.ReportProgress("Saving... " & sReportName, lngProgressCnt)
    Kill App.Path + "\" & sReportName & ".xml"
    .Save App.Path + "\" & sReportName & ".xml"
     
    lngProgressCnt = lngProgressCnt + 1
    Call frmScanProgress.ReportChange("Formating report...")
    
    sStrReport = FormatXML(xroot, 0)
    
    lngFileNumber = FreeFile
    Call frmScanProgress.ReportProgress("Saving... " & sReportName & "f", lngProgressCnt)
    
    Kill App.Path + "\" & sReportName & "f.xml"
    
    Open App.Path + "\" & sReportName & "f.xml" For Output As #lngFileNumber ' Create file name.
       
    Write #lngFileNumber, sStrReport    ' Output text.
    Close #lngFileNumber  ' Close file.



   End With
   
   frmScanProgress.Hide
  
  On Error GoTo 0
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "SaveXML")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

' *******************************************************
' Routine Name : (PUBLIC in CLASS) Sub Scan
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MKPROG]
' Date Writen  : 02-24-2001 -- 08:42:18
' Inputs       : N/A
' Outputs      : N/A
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Public Sub Scan()
 Dim cFunction       As clsFunction
 Dim vObj            As Variant
 Dim lngProgressCnt  As Long

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
  ' ---------------------------------------------------
  '   AddCodes
  ' ---------------------------------------------------
  AddCodes
  ' ---------------------------------------------------
  '   Clear out any current memebers in the collection
  ' ---------------------------------------------------
  Set colVariables = New FastCollection
  
  ' ---------------------------------------------------
  '   initialize the progres dialog
  ' ---------------------------------------------------
  Call frmScanProgress.InitProgress("Please Wait...", 0, colCodes.Count)
  ' ---------------------------------------------------
  ' scan for variables in code
  ' ---------------------------------------------------
  For Each cFunction In colCodes
        
    With cFunction
      ' ---------------------------------------------------
      ' set progress
      ' ---------------------------------------------------
      Call frmScanProgress.ReportProgress("Scanning " & .sModuleName & " for variable names.", lngProgressCnt)
      If Len(.sName) > 0 Then
        Call frmScanProgress.ReportChange("Scanning function: " & .sName & " for variable names.")
       Else
        Call frmScanProgress.ReportChange("")
      End If
      ' ---------------------------------------------------
      ' Scan
      ' ---------------------------------------------------
      ScanVariables .sModuleName, .sName, .sCode
    
    End With
    
    Set cFunction = Nothing
    
    lngProgressCnt = lngProgressCnt + 1
    
  Next cFunction
  Set cFunction = Nothing
   
  frmScanProgress.Hide
  
  On Error GoTo 0
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Scan")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

Public Sub ScanForUse()

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
 Dim cFunction       As clsFunction
 Dim lngProgressCnt  As Long
 Dim cVariable       As clsVariable

  '   initialize the progres dialog
  
  Call frmDialog.InitProgress("Please Wait preparing for scaning...", 0, VBInstance.ActiveVBProject.VBComponents.Count - 1)
   
  'now scan for use variables
  
  Set colVariableNames = New FastCollection
  'remove duplicate variable names
  For Each cVariable In colVariables
    With colVariableNames
      If Not .KeyExists(cVariable.sName) Then
        Call .Add(cVariable, cVariable.sName)
      End If
    End With
  Next cVariable
  Set cVariable = Nothing
  
  For Each cFunction In colCodes
        
    With cFunction
      
      If .bScanForUse Then
        ' ---------------------------------------------------
        ' set progress
        ' ---------------------------------------------------
        Call frmDialog.ReportProgress("Scanning " & .sModuleName & " for variable names.", lngProgressCnt)
        ' ---------------------------------------------------
        ' Scan
        ' ---------------------------------------------------
        ScanVariablesUSE .sModuleName, .sName, .sCode
      End If
    End With
    
    Set cFunction = Nothing
    
    lngProgressCnt = lngProgressCnt + 1
    
  Next cFunction
  Set cFunction = Nothing
  
  On Error GoTo 0
               
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "ScanForUse")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub

' *******************************************************
' Routine Name : (PRIVATE in CLASS) Sub ScanVariables
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MK_PROG]
' Date Writen  : 01-20-2001 -- 14:00:42
' Inputs       : ByVal sModuleName:String -
'              : ByVal sFunction:String -
'              : ByVal sCode:String -
' Outputs      : N/A
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Private Sub ScanVariables(ByVal sModuleName As String, _
                          ByVal sFunction As String, _
                          ByVal sCode As String)
 Dim clsRegExp   As RegExp
 Dim clsMatchCol As MatchCollection
 Dim clsMatch    As Match
 Dim strTemp     As String
 Dim colVars     As FastCollection
 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
                          


  On Error GoTo errH

  Set clsRegExp = New RegExp
  'first remove comments
  With clsRegExp
    .Global = True
    .IgnoreCase = True
    
    '.Pattern = "'.*\n"
    'sCode = .Replace(sCode, vbCrLf)
  
    'next remove new line char
    .Pattern = "\s*_{1}\s*\n"
    sCode = .Replace(sCode, "")

    'replace more then 1 space to 1
    '.Pattern = "( ){2,}"
    'sCode = .Replace(sCode, " ")
    
    'replace other duplicated ascii chars
    '.Pattern = "\r{2,}"
    'sCode = .Replace(sCode, vbCrLf)
  
    '.Pattern = "\n{2,}"
    'sCode = .Replace(sCode, vbCrLf)
  
  End With
  
  '------------------------------------------------------------------------------------------------------------------------------------------------
  'Search for variables global, private, public, static, dim, const
  '------------------------------------------------------------------------------------------------------------------------------------------------
  '(global|private|public|static|dim|const)+\s+\w+.*\n{0,1}
  '------------------------------------------------------------------------------------------------------------------------------------------------
  clsRegExp.Pattern = "\s*(global|private|public|static|dim|redim|const)+\s+\w+.*\n{0,1}"

  Set clsMatchCol = clsRegExp.Execute(sCode)
  
  For Each clsMatch In clsMatchCol
    With clsMatch
      'Debug.Print .Value
      strTemp = strTemp & CheckType(sModuleName, sFunction, .Value)
    End With
  Next clsMatch

  sCode = strTemp

  Set clsMatch = Nothing
  Set clsMatchCol = Nothing

Exit Sub

errH:
  Debug.Print "Err:" & Err.Description
  Stop
  Resume
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "ScanVariables")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub


' *******************************************************
' Routine Name : (PRIVATE in CLASS) Sub ScanVariablesUSE
' Written By   : Mateusz Kierepka
' Programmer   : Mateusz Kierepka [MK_PROG]
' Date Writen  : 01-20-2001 -- 14:00:44
' Inputs       : ByVal sModuleName:String -
'              : ByVal sFunction:String -
'              : ByVal sCode:String -
' Outputs      : N/A
' Description  :
'              :
'              :
' Called By    :
' *******************************************************
Private Sub ScanVariablesUSE(ByVal sModuleName As String, _
                             ByVal sFunction As String, _
                             ByVal sCode As String)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}
                          
 Dim clsRegExp   As RegExp
 Dim clsMatchCol As MatchCollection
 Dim strTemp     As String
 Dim colVars     As FastCollection
 Dim cVariable   As clsVariable
 Dim i           As Long
  On Error GoTo errH

  Set clsRegExp = New RegExp
  'first remove comments
  With clsRegExp
    .Global = True
    .IgnoreCase = True
    .Pattern = "'.*\n"

    sCode = .Replace(sCode, vbCrLf)
  
    'next remove new line char
    .Pattern = "\s*_{1}\s*\n"
    sCode = .Replace(sCode, "")

    'replace more then 1 space to 1
    .Pattern = "( ){2,}"
    sCode = .Replace(sCode, " ")
    
    'replace other duplicated ascii chars
    .Pattern = "\r{2,}"
    sCode = .Replace(sCode, vbCrLf)
  
    .Pattern = "\n{2,}"
    sCode = .Replace(sCode, vbCrLf)
  
  End With
  
  'i = 0
  For Each cVariable In colVariableNames
    With cVariable
      '------------------------------------------------------------------------------------------------------------------------------------------------
      'Search for variables global, private, public, static, dim, const
      '------------------------------------------------------------------------------------------------------------------------------------------------
      '  '(\s|\(|\+|\-|\*|\&|\/|\\|^|\=){1}varname(\s|\)|\+|\-|\*|\&|\/|\\|^|\=){0,1}
      '------------------------------------------------------------------------------------------------------------------------------------------------
      'Debug.Print "var: " & .sName
      clsRegExp.Pattern = "(\s|\(|\+|\-|\*|\&|\/|\\|^|\=){1}" & _
                          .sName & "(\s|\)|\+|\-|\*|\&|\/|\\|^|\=){0,1}"
    
      Set clsMatchCol = clsRegExp.Execute(sCode)
      
      If clsMatchCol.Count > 0 Then
        Call .AddFunction(sFunction, clsMatchCol.Count, sModuleName)
      End If

    End With
    'i = i + 1
    
  Next cVariable
  
  sCode = strTemp

  Set clsMatchCol = Nothing

Exit Sub

errH:
  Debug.Print "Err:" & Err.Description
  'Stop
  Resume Next
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "ScanVariablesUSE")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Sub


Public Property Get VBInstance() As VBIDE.VBE

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  Set VBInstance = m_VBInstance
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Get VBInstance")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

Public Property Set VBInstance(ByRef vNewValue As VBIDE.VBE)

 '{{{ Added It!

  On Error GoTo Generated_trap '}}}

  Set m_VBInstance = vNewValue
  '{{{ Added It!
  Err.Clear
Generated_trap:
  If Err <> 0 Then
    Select Case ToDoOnError(Err, "Set VBInstance")
     Case vbRetry: Resume
     Case vbIgnore: Resume Next
    End Select
  End If '}}}

End Property

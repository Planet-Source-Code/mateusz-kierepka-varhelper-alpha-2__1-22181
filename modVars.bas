Attribute VB_Name = "modVars"
Option Explicit

Public Const sAddinName           As String = "VarHelper"

Public Const LB_GETTOPINDEX       As Long = &H18E
Public Const LB_SETTOPINDEX       As Long = &H197

Public cSearchObject              As clsSearch
Public colCodes                   As FastCollection
Public colVariables               As FastCollection
Public colVariableNames           As FastCollection

Public Function ToDoOnError(inErr As ErrObject, inProcName As String) As Integer

 Dim Msg As String

  Msg = inErr.Source & " caused error """ & inErr.Description & """ (" & inErr.Number & ")" & vbCrLf _
        & "in module clsSearch procedure " & inProcName & ", line " & Erl & "."

  ToDoOnError = MsgBox(Msg, vbAbortRetryIgnore + vbMsgBoxHelpButton + vbCritical, _
                "What to you want to do?", inErr.HelpFile, inErr.HelpContext)

End Function

':) Ulli's Code Formatter V2.0 (2001-01-23 10:53:17) 11 + 14 = 25 Lines


Option Compare Database
Option Explicit

Private Const CurrentModName = "clsUpdateVersionNumber"
Private Const UpdatedModuleName = "mGIT"
Private Const UpdatedConstantName = "vbaGIVersionNumber"

Public Function Update()
        
    Dim PreviousNumberVerion As String
    Dim NewNumberVerion As String
   'Dim RepoNumberVerion As String 'TO DO - pobrac numer wersji z repo
    
    PreviousNumberVerion = ReturnConstValueFromLineInModule(UpdatedModuleName, vbaGIVersionNumber)
   'NewNumberVerion = Format(FormatDateTime(VBA.Now, vbGeneralDate), "YYMMDD.HHmms")
    NewNumberVerion = Left(Format(VBA.Now, "YYMMDD") & "0", 6) & "." & Left(Format(VBA.Now, "HHmms") & "0", 6)
    NewNumberVerion = Chr(34) & NewNumberVerion & Chr(34)
    
    Dim OldDesc$, NewDesc$
    Dim Msg$
    
    OldDesc = PreviousNumberVerion & " <= Old Number"
    NewDesc = NewNumberVerion & " <= New Number"
    Msg = "Czy chcesz zaktualizować Numer wersji projektu ?" & vbLf & vbLf & OldDesc & vbLf & NewDesc
    If VBA.MsgBox(Msg, vbQuestion + vbYesNo, "DEBUG") = vbYes Then
        Debug.Print " --- UPDATE Version Number ---"
        Debug.Print OldDesc
        Debug.Print NewDesc
        Call FindAndReplace(UpdatedModuleName, PreviousNumberVerion, NewNumberVerion)
    End If
End Function

Private Function ReturnConstValueFromLineInModule(strModuleName As String, strSearchConstName As String) As String
 
 Dim mdl As Module
 Dim strLine As String
 Dim strValue As String
 
 Dim lngSLine As Long
 Dim lngSCol As Long, lngECol As Long, lngELine As Long
 Dim bModuleIsOpen As Boolean
 
 ' Open module.
 DoCmd.OpenModule strModuleName
 ' Return reference to Module object.
 Set mdl = Modules(strModuleName)
 
 ' Search for string.
 If mdl.Find(strSearchConstName, lngSLine, lngSCol, lngELine, lngECol) Then
 
    ' Store text of line containing string
    strLine = mdl.Lines(lngSLine, Abs(lngELine - lngSLine) + 1)
    ' get values
    strValue = CStr(Split(strLine, "=")(1))
    ' remove som characters
    strValue = Replace(strValue, "#", "")
    
    ' return values
    ReturnConstValueFromLineInModule = Trim(strValue)
    
 Else
    MsgBox "Const Variable not found. (" & strSearchConstName & ")"
 End If
 
Exit_ReturnConstValueFromLineInModule:
 Exit Function
 
Error_ReturnConstValueFromLineInModule:
 
MsgBox Err & ": " & Err.Description

 Resume Exit_ReturnConstValueFromLineInModule
End Function

Private Function FindAndReplace(strModuleName As String, _
 strSearchText As String, _
 strNewText As String) As Boolean
 Dim mdl As Module
 Dim lngSLine As Long, lngSCol As Long
 Dim lngELine As Long, lngECol As Long
 Dim strLine As String, strNewLine As String
 Dim intChr As Integer, intBefore As Integer, _
 intAfter As Integer
 Dim strLeft As String, strRight As String
 
 ' Open module.
 DoCmd.OpenModule strModuleName
 ' Return reference to Module object.
 Set mdl = Modules(strModuleName)
 
 ' Search for string.
 If mdl.Find(strSearchText, lngSLine, lngSCol, lngELine, _
 lngECol) Then
 ' Store text of line containing string.
 strLine = mdl.Lines(lngSLine, Abs(lngELine - lngSLine) + 1)
 ' Determine length of line.
 intChr = Len(strLine)
 ' Determine number of characters preceding search text.
 intBefore = lngSCol - 1
 ' Determine number of characters following search text.
 intAfter = intChr - CInt(lngECol - 1)
 ' Store characters to left of search text.
 strLeft = Left$(strLine, intBefore)
 ' Store characters to right of search text.
 strRight = Right$(strLine, intAfter)
 ' Construct string with replacement text.
 strNewLine = strLeft & strNewText & strRight
 ' Replace original line.
 mdl.ReplaceLine lngSLine, strNewLine
 FindAndReplace = True
 Else
 MsgBox "Text not found."
 FindAndReplace = False
 End If
 
Exit_FindAndReplace:
 Exit Function
 
Error_FindAndReplace:
 
MsgBox Err & ": " & Err.Description
 FindAndReplace = False
 Resume Exit_FindAndReplace
End Function

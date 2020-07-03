Attribute VB_Name = "mdVb5Comp"
Option Explicit

Public Function Replace(ByVal TextToReplace As String, ByVal OldChr As String, ByVal NewChr As String) As String
    Replace = Join(Split(TextToReplace, OldChr), NewChr)
End Function

Public Function Split(ByVal TextToSplit As String, Optional Delimiter As String = " ") As Variant
    Dim tempStr As String, tempArr() As String
    Dim X As Long
    
    ReDim tempArr(-1 To -1) As String
    If TextToSplit <> "" Then
        If Delimiter <> "" Then
            Do
                X = InStr(TextToSplit, Delimiter)
                If X <> 0 Then
                    tempStr = Left(TextToSplit, X - 1)
                    TextToSplit = Right(TextToSplit, Len(TextToSplit) - X + 1 - Len(Delimiter))
                    If UBound(tempArr) < 0 Then
                        ReDim tempArr(0 To 0) As String
                    Else
                        ReDim Preserve tempArr(UBound(tempArr) + 1) As String
                    End If
                    tempArr(UBound(tempArr)) = tempStr
                    
                End If
            Loop While X <> 0
        End If
        If TextToSplit <> "" Then
            If UBound(tempArr) < 0 Then
                ReDim tempArr(0 To 0) As String
            Else
                ReDim Preserve tempArr(UBound(tempArr) + 1) As String
            End If
            tempArr(UBound(tempArr)) = TextToSplit
        End If
    End If
    Split = tempArr()
End Function

Public Function Join(ByVal SourceArray As Variant, Optional Delimiter As String = " ") As String
    Dim tempStr As String
    Dim A As Integer
    
    If Not IsArray(SourceArray) Then Exit Function
    If UBound(SourceArray) = -1 Then Exit Function
    For A = 0 To UBound(SourceArray)
        If A = UBound(SourceArray) Then
            tempStr = tempStr & SourceArray(A)
        Else
            tempStr = tempStr & SourceArray(A) & Delimiter
        End If
    Next
    
    Join = tempStr
End Function


Public Function InStrRev(ByVal StringCheck As String, _
   StringMatch As String, Optional Start As Long = -1, _
   Optional Compare As VbCompareMethod = vbBinaryCompare) _
   As Long

'********************************************
'PURPOSE: Implements InStrRev functionality in
'VB5, whereby you can begin searching for
'a character sequence within a string from the
'end rather than the beginning

'PARAMETERS:  StringCheck: String expression being searched.

'             StringMatch: String expression being searched for

'             Start (Optional) starting position for each
'             search. If omitted, search begins at the last
'             character position. Starting position is calculated
'             from the beginning of StringCheck

'             Compare (Optional) kind of comparison to use;
'             defaults to vbBinaryCompare
'             Search begins at end of string

'EXAMPLE:     debug.print(InstrRev("www.freevbcode.com", ".")
'             outputs 15, wherease inStr returs 4

'NOTE:        For VB6 and above, use the built-in InStrRev
'             Function, not this
'**************************************************************
If Len(StringMatch) > Len(StringCheck) Then Exit Function
If Start < -1 Or Start = 0 Or Start > Len(StringCheck) _
      Then Exit Function

Dim lStartPoint As Long
Dim lEndPoint As Long
Dim lSearchLength As Long
Dim lCtr As Long
Dim sWkg As String

lSearchLength = Len(StringMatch)
lStartPoint = IIf(Start = -1, Len(StringCheck), Start)
lEndPoint = 1

For lCtr = lStartPoint To lEndPoint Step -1
    sWkg = Mid(StringCheck, lCtr, lSearchLength)
    If StrComp(sWkg, StringMatch, Compare) = 0 Then
        InStrRev = lCtr
        Exit Function
    End If
Next
        
End Function


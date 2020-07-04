Attribute VB_Name = "mdVb5Comp"
Option Explicit

Private Function Replace(ByVal Source As String, ByVal Find As String, ByVal ReplaceStr As String) As String
    Find = preg_replace("[.*+?^${}()/|[\]\\]", Find, "\$&")
    Replace = preg_replace(Find, Source, ReplaceStr)
End Function

Private Function preg_replace(sPattern As String, sText As String, Optional Replace As String) As String
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = sPattern
        preg_replace = .Replace(sText, Replace)
    End With
End Function

Private Function Split(ByVal TextToSplit As String, Optional Delimiter As String = " ") As Variant
    Dim sTemp           As String
    Dim aRetVal()       As String
    Dim lPos            As Long
    
    ReDim aRetVal(-1 To -1) As String
    If TextToSplit <> "" Then
        If Delimiter <> "" Then
            Do
                lPos = InStr(TextToSplit, Delimiter)
                If lPos <> 0 Then
                    sTemp = Left$(TextToSplit, lPos - 1)
                    TextToSplit = Right$(TextToSplit, Len(TextToSplit) - lPos + 1 - Len(Delimiter))
                    If UBound(aRetVal) < 0 Then
                        ReDim aRetVal(0 To 0) As String
                    Else
                        ReDim Preserve aRetVal(UBound(aRetVal) + 1) As String
                    End If
                    aRetVal(UBound(aRetVal)) = sTemp
                End If
            Loop While lPos <> 0
        End If
        If TextToSplit <> "" Then
            If UBound(aRetVal) < 0 Then
                ReDim aRetVal(0 To 0) As String
            Else
                ReDim Preserve aRetVal(UBound(aRetVal) + 1) As String
            End If
            aRetVal(UBound(aRetVal)) = TextToSplit
        End If
    End If
    Split = aRetVal()
End Function

Private Function Join(ByVal SourceArray As Variant, Optional Delimiter As String = " ") As String
    Dim sTemp           As String
    Dim lIdx            As Integer
    
    If Not IsArray(SourceArray) Then Exit Function
    If UBound(SourceArray) = -1 Then Exit Function
    For lIdx = 0 To UBound(SourceArray)
        If lIdx = UBound(SourceArray) Then
            sTemp = sTemp & SourceArray(lIdx)
        Else
            sTemp = sTemp & SourceArray(lIdx) & Delimiter
        End If
    Next
    Join = sTemp
End Function

Private Function InStrRev( _
            ByVal StringCheck As String, _
            StringMatch As String, _
            Optional Start As Long = -1, _
            Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
    Dim lStartPoint     As Long
    Dim lEndPoint       As Long
    Dim lSearchLength   As Long
    Dim lCtr            As Long
    Dim sWkg            As String
    
    If Len(StringMatch) > Len(StringCheck) Then
        Exit Function
    End If
    If Start < -1 Or Start = 0 Or Start > Len(StringCheck) Then
        Exit Function
    End If
    
    
    lSearchLength = Len(StringMatch)
    lStartPoint = IIf(Start = -1, Len(StringCheck), Start)
    lEndPoint = 1
    For lCtr = lStartPoint To lEndPoint Step -1
        sWkg = Mid$(StringCheck, lCtr, lSearchLength)
        If StrComp(sWkg, StringMatch, Compare) = 0 Then
            InStrRev = lCtr
            Exit Function
        End If
    Next
End Function

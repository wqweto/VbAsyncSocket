Attribute VB_Name = "mdGlobals"
Option Explicit

'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private m_oRedimStats               As Object

Public Function DesignDumpArray(baData() As Byte, Optional ByVal lPos As Long, Optional ByVal lSize As Long = -1) As String
    If lSize < 0 Then
        lSize = UBound(baData) + 1 - lPos
    End If
    If lSize > 0 Then
        DesignDumpArray = DesignDumpMemory(VarPtr(baData(lPos)), lSize)
    End If
End Function

Public Function DesignDumpMemory(ByVal lPtr As Long, ByVal lSize As Long) As String
    Dim lIdx            As Long
    Dim sHex            As String
    Dim sChar           As String
    Dim lValue          As Long
    Dim aResult()       As String
    
    ReDim aResult(0 To (lSize + 15) \ 16) As String
    Debug.Assert RedimStats("DesignDumpMemory.aResult", UBound(aResult) + 1)
    For lIdx = 0 To ((lSize + 15) \ 16) * 16
        If lIdx < lSize Then
            If IsBadReadPtr(lPtr, 1) = 0 Then
                Call CopyMemory(lValue, ByVal lPtr, 1)
                sHex = sHex & Right$("0" & Hex$(lValue), 2) & " "
                If lValue >= 32 Then
                    sChar = sChar & Chr$(lValue)
                Else
                    sChar = sChar & "."
                End If
            Else
                sHex = sHex & "?? "
                sChar = sChar & "."
            End If
        Else
            sHex = sHex & "   "
        End If
        If ((lIdx + 1) Mod 4) = 0 Then
            sHex = sHex & " "
        End If
        If ((lIdx + 1) Mod 16) = 0 Then
            aResult(lIdx \ 16) = Right$("000" & Hex$(lIdx - 15), 4) & " - " & sHex & sChar
            sHex = vbNullString
            sChar = vbNullString
        End If
        lPtr = (lPtr Xor &H80000000) + 1 Xor &H80000000
    Next
    DesignDumpMemory = Join(aResult, vbCrLf)
End Function

Public Sub WriteBinaryFile(sFile As String, baBuffer() As Byte)
    Dim nFile           As Integer
    
    nFile = FreeFile
    Open sFile For Binary Access Write Shared As nFile
    If UBound(baBuffer) >= 0 Then
        Put nFile, , baBuffer
    End If
    Close nFile
End Sub

Public Function ToHex(baText() As Byte, Optional Delimiter As String = "-") As String
    Dim aText()         As String
    Dim lIdx            As Long
    
    If LenB(CStr(baText)) <> 0 Then
        ReDim aText(0 To UBound(baText)) As String
        Debug.Assert RedimStats("ToHex.aText", 0)
        For lIdx = 0 To UBound(baText)
            aText(lIdx) = Right$("0" & Hex$(baText(lIdx)), 2)
        Next
        ToHex = Join(aText, Delimiter)
    End If
End Function

Public Function FromHex(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    On Error GoTo QH
    '--- check for hexdump delimiter
    If sText Like "*[!0-9A-Fa-f]*" Then
        ReDim baRetVal(0 To Len(sText) \ 3) As Byte
        Debug.Assert RedimStats("FromHex.baRetVal", UBound(baRetVal) + 1)
        For lIdx = 1 To Len(sText) Step 3
            baRetVal(lIdx \ 3) = "&H" & Mid$(sText, lIdx, 2)
        Next
    ElseIf LenB(sText) <> 0 Then
        ReDim baRetVal(0 To Len(sText) \ 2 - 1) As Byte
        Debug.Assert RedimStats("FromHex.baRetVal", UBound(baRetVal) + 1)
        For lIdx = 1 To Len(sText) Step 2
            baRetVal(lIdx \ 2) = "&H" & Mid$(sText, lIdx, 2)
        Next
    Else
        baRetVal = vbNullString
    End If
    FromHex = baRetVal
QH:
End Function

Public Function ToUtf8Array(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Debug.Assert RedimStats("ToUtf8Array.baRetVal", UBound(baRetVal) + 1)
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
    Else
        baRetVal = vbNullString
    End If
    ToUtf8Array = baRetVal
End Function

Public Function FromUtf8Array(baText() As Byte) As String
    Dim lSize           As Long
    
    If UBound(baText) >= 0 Then
        FromUtf8Array = String$(2 * UBound(baText), 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, StrPtr(FromUtf8Array), Len(FromUtf8Array))
        FromUtf8Array = Left$(FromUtf8Array, lSize)
    End If
End Function

Private Function RedimStats(sFuncName As String, ByVal lSize As Long) As Boolean
    If m_oRedimStats Is Nothing Then
        Set m_oRedimStats = CreateObject("Scripting.Dictionary")
    End If
    If LenB(sFuncName) <> 0 Then
        m_oRedimStats.Item(sFuncName) = m_oRedimStats.Item(sFuncName) + 1
        m_oRedimStats.Item("#" & sFuncName) = m_oRedimStats.Item("#" & sFuncName) + lSize
    End If
    '--- success
    RedimStats = True
End Function

Public Function DesignDumpRedimStats(Optional ByVal Clear As Boolean) As String
    Dim vElem           As Variant
    Dim aText()         As String
    Dim lIdx            As Long
    
    If m_oRedimStats Is Nothing Then
        Exit Function
    End If
    ReDim aText(0 To m_oRedimStats.Count - 1) As String
    For Each vElem In m_oRedimStats.Keys
        aText(lIdx) = vElem & ": " & m_oRedimStats.Item(vElem)
        If Left$(vElem, 1) = "#" Then
            aText(lIdx) = aText(lIdx) & " (avg. " & Format$(m_oRedimStats.Item(vElem) / m_oRedimStats.Item(Mid(vElem, 2)), "0.0") & ")"
        End If
        lIdx = lIdx + 1
    Next
    DesignDumpRedimStats = Join(aText, vbCrLf)
    If Clear Then
        Set m_oRedimStats = Nothing
    End If
End Function

Public Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Public Sub DebugLog(sModule As String, sFunction As String, sText As String, Optional ByVal eType As LogEventTypeConstants = vbLogEventTypeInformation)
    Debug.Print Format$(TimerEx, "0.000") & " " & Switch( _
        eType = vbLogEventTypeError, "[ERROR]", _
        eType = vbLogEventTypeWarning, "[WARN]", _
        True, "[INFO]") & " " & sText & " [" & sModule & "." & sFunction & "]"
End Sub

Public Function ConcatCollection(oCol As Collection, Optional Separator As String = vbCrLf) As String
    Dim lSize           As Long
    Dim vElem           As Variant
    
    For Each vElem In oCol
        lSize = lSize + Len(vElem) + Len(Separator)
    Next
    If lSize > 0 Then
        ConcatCollection = String$(lSize - Len(Separator), 0)
        lSize = 1
        For Each vElem In oCol
            If lSize <= Len(ConcatCollection) Then
                Mid$(ConcatCollection, lSize, Len(vElem) + Len(Separator)) = vElem & Separator
            End If
            lSize = lSize + Len(vElem) + Len(Separator)
        Next
    End If
End Function

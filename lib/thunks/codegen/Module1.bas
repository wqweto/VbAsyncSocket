Attribute VB_Name = "Module1"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private m_baBuffer()            As Byte
Private m_lBuffIdx              As Long

Private Sub pvAppendBuffer(ByVal a01 As Long, ByVal a02 As Long, ByVal a03 As Long, ByVal a04 As Long, ByVal a05 As Long, ByVal a06 As Long, ByVal a07 As Long, ByVal a08 As Long, _
                           ByVal a09 As Long, ByVal a10 As Long, ByVal a11 As Long, ByVal a12 As Long, ByVal a13 As Long, ByVal a14 As Long, ByVal a15 As Long, ByVal a16 As Long)
    #If a02 And a03 And a04 And a05 And a06 And a07 And a08 And a09 And a10 And a11 And a12 And a13 And a14 And a15 And a16 Then '--- touch args
    #End If
    Call CopyMemory(m_baBuffer(m_lBuffIdx), a01, 4 * 16)
    m_lBuffIdx = m_lBuffIdx + 4 * 16
End Sub

Public Sub GetModelData(baBuffer() As Byte)
    ReDim m_baBuffer(0 To 1024) As Byte
    m_lBuffIdx = 0
    '--- begin thunk data
    pvAppendBuffer 1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 12345691, 12345692, 12345693, 12345694, 12345695, 12345696, 12345697, 12345698
    '--- end thunk data
    ReDim baBuffer(0 To 300) As Byte
    Call CopyMemory(baBuffer(0), m_baBuffer(0), UBound(baBuffer) + 1)
    Erase m_baBuffer
End Sub

Public Function GenerModuleDeclares() As String
    Const NUM_PARAMS    As Long = 32
    Dim cOutput         As Collection
    Dim lIdx            As Long
    Dim aStr(0 To NUM_PARAMS - 1) As String
    
    For lIdx = 1 To NUM_PARAMS
        aStr(lIdx - 1) = "a" & Right$("0" & lIdx, 2)
    Next
    Set cOutput = New Collection
    cOutput.Add "Option Explicit"
    cOutput.Add ""
    cOutput.Add "Private Declare Sub CopyMemory Lib ""kernel32"" Alias ""RtlMoveMemory"" (Destination As Any, Source As Any, ByVal Length As Long)"
    cOutput.Add ""
    cOutput.Add "Private m_baBuffer()            As Byte"
    cOutput.Add "Private m_lBuffIdx              As Long"
    cOutput.Add ""
    cOutput.Add "Private Sub pvAppendBuffer(ByVal " & Join(aStr, " As Long, ByVal ") & " As Long)"
    cOutput.Add "    #If " & Join(aStr, " And ") & " Then '--- touch args"
    cOutput.Add "    #End If"
    cOutput.Add "    Call CopyMemory(m_baBuffer(m_lBuffIdx), a01, 4 * " & NUM_PARAMS & ")"
    cOutput.Add "    m_lBuffIdx = m_lBuffIdx + 4 * " & NUM_PARAMS
    cOutput.Add "End Sub"
    cOutput.Add ""
    GenerModuleDeclares = ConcatCollection(cOutput, vbCrLf)
End Function

Public Function GenerThunkData(sThunkStr As String, sName As String) As String
    Const NUM_PARAMS    As Long = 32
    Const LNG_ALIGN     As Long = 4 * NUM_PARAMS
    Dim baThunk()       As Byte
    Dim lSize           As Long
    Dim lAlignedSize    As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim cOutput         As Collection
    Dim lPtr            As Long
    Dim aParams(0 To NUM_PARAMS - 1) As Long
    Dim aStr(0 To NUM_PARAMS - 1) As String
    
    baThunk = FromBase64Array(sThunkStr)
    lSize = UBound(baThunk) + 1
    lAlignedSize = (lSize + LNG_ALIGN - 1) And -LNG_ALIGN
    ReDim Preserve baThunk(0 To lAlignedSize - 1) As Byte
    Set cOutput = New Collection
    cOutput.Add "Private Sub pvGet" & sName & "Data(baBuffer() As Byte)"
    cOutput.Add "    ReDim m_baBuffer(0 To " & lAlignedSize & " - 1) As Byte"
    cOutput.Add "    m_lBuffIdx = 0"
    cOutput.Add "    '--- begin thunk data"
    For lIdx = 0 To lSize - 1 Step LNG_ALIGN
        Call CopyMemory(aParams(0), baThunk(lIdx), LNG_ALIGN)
        For lJdx = 0 To UBound(aParams)
            aStr(lJdx) = "&H" & Hex$(aParams(lJdx)) & "&"
        Next
        cOutput.Add "    pvAppendBuffer " & Join(aStr, ", ")
    Next
    cOutput.Add "    '--- end thunk data"
    cOutput.Add "    ReDim baBuffer(0 To " & lSize & " - 1) As Byte"
    cOutput.Add "    Call CopyMemory(baBuffer(0), m_baBuffer(0), UBound(baBuffer) + 1)"
    cOutput.Add "    Erase m_baBuffer"
    cOutput.Add "End Sub"
    cOutput.Add ""
    GenerThunkData = ConcatCollection(cOutput, vbCrLf)
End Function

Public Function FromBase64Array(sText As String) As Byte()
    With VBA.CreateObject("MSXML2.DOMDocument").createElement("dummy")
        .DataType = "bin.base64"
        .Text = sText
        FromBase64Array = .NodeTypedValue
    End With
End Function

Public Function ConcatCollection(oCol As Collection, Optional Separator As String = "") As String
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


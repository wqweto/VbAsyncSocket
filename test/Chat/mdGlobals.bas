Attribute VB_Name = "mdGlobals"
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const WM_VSCROLL                        As Long = &H115
Private Const SB_BOTTOM                         As Long = 7

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'=========================================================================
' Functions
'=========================================================================

Public Function RtbAppendLine( _
            rchCtl As InkEdit, _
            sText As String, _
            Optional ByVal HorAlign As SelAlignmentConstants = rtfLeft, _
            Optional ByVal ForeColor As OLE_COLOR = vbWindowText)
    rchCtl.SelStart = &H7FFFFFFF
    rchCtl.SelAlignment = HorAlign
    rchCtl.SelColor = ForeColor
    rchCtl.SelText = sText & vbCrLf
    rchCtl.SelStart = &H7FFFFFFF
    Call SendMessage(rchCtl.hWnd, WM_VSCROLL, SB_BOTTOM, ByVal 0&)
End Function


Public Function Printf(ByVal sText As String, ParamArray A() As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    
    For lIdx = UBound(A) To LBound(A) Step -1
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE)))
    Next
    Printf = Replace(sText, ChrW$(LNG_PRIVATE), "%")
End Function

Public Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    On Error GoTo QH
    AssignVariant RetVal, pCol.Item(Index)
    SearchCollection = True
QH:
End Function

Public Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

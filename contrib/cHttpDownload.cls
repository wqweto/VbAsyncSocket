VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' VbAsyncSocket Project (c) 2018-2019 by wqweto@gmail.com
'
' Simple and thin WinSock API wrappers for VB6
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cHttpDownload"

'=========================================================================
' Events
'=========================================================================

Event DownloadProgress(ByVal BytesRead As Double, ByVal BytesTotal As Double)
Event DownloadError(ByVal Number As Long, ByVal Description As String)
Event DownloadComplete(ByVal LocalFileName As String)

'=========================================================================
' API
'=========================================================================

'--- for SHCreateStreamOnFile
Private Const STGM_WRITE                    As Long = 1
Private Const STGM_CREATE                   As Long = &H1000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private Declare Function SHCreateStreamOnFile Lib "shlwapi" Alias "SHCreateStreamOnFileW" (ByVal pszFile As Long, ByVal grfMode As Long, ppstm As IUnknown) As Long
Private Declare Function IStream_Write Lib "shlwapi" (ByVal pstm As IUnknown, pv As Any, ByVal cb As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const HDR_CONTENT_LENGTH    As String = "content-length:"
Private Const HDR_LOCATION          As String = "location:"
Private Const LNG_FLUSH_LIMIT       As Long = 1024& * 1024
'--- errors
Private Const ERR_INVALID_URL       As String = "Invalid URL"
Private Const ERR_HTTP_ONLY         As String = "Http protocol supported only"
Private Const ERR_INVALID_REDIRECT  As String = "Invalid redirect to %1"
Private Const ERR_INVALID_RESPONSE  As String = "Server returned %1"

Private WithEvents m_oSocket    As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_uRemote               As UcsParsedUrl
Private m_sLocalFileName        As String
Private m_pOutputStream         As IUnknown
Private m_dStartDate            As Date
Private m_lCallbackPtr          As Long
Private m_eState                As UcsStateEnum
Private m_baRecvBuffer()        As Byte
Private m_dblBytesRead          As Double
Private m_dblContentLength      As Double

Private Enum UcsStateEnum
    ucsIdle
    ucsWaitHeaders
    ucsWaitBody
End Enum

Private Type UcsParsedUrl
    Protocol        As String
    Host            As String
    Port            As Long
    Path            As String
    User            As String
    Pass            As String
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get URL() As String
    URL = m_uRemote.Protocol & "://" & m_uRemote.Host & m_uRemote.Path
End Property

Property Get LocalFileName() As String
    LocalFileName = m_sLocalFileName
End Property

Property Get StartDate() As Date
    StartDate = m_dStartDate
End Property

Property Get CallbackWeakRef() As Object
    Call vbaObjSetAddref(CallbackWeakRef, m_lCallbackPtr)
End Property

Property Set CallbackWeakRef(oValue As Object)
    m_lCallbackPtr = ObjPtr(oValue)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub DownloadFile(URL As String, LocalFileName As Variant)
    Const FUNC_NAME     As String = "DownloadFile"
    Dim hResult     As Long
    
    On Error GoTo EH
    If Not pvParseUrl(URL, m_uRemote) Then
        On Error GoTo 0
        Err.Raise vbObjectError, , ERR_INVALID_URL
    End If
    If LCase$(m_uRemote.Protocol) <> "http" Then
        On Error GoTo 0
        Err.Raise vbObjectError, , ERR_HTTP_ONLY
    End If
    If IsObject(LocalFileName) Then
        m_sLocalFileName = "[stream]"
        Set m_pOutputStream = LocalFileName
    Else
        m_sLocalFileName = LocalFileName
        hResult = SHCreateStreamOnFile(StrPtr(m_sLocalFileName), STGM_WRITE Or STGM_CREATE, m_pOutputStream)
        If hResult < 0 Then
            On Error GoTo 0
            Err.Raise hResult
        End If
    End If
    m_dStartDate = Now
    m_eState = ucsIdle
    m_dblBytesRead = 0
    m_baRecvBuffer = vbNullString
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketType:=ucsSckStream) Then
        On Error GoTo 0
        Err.Raise vbObjectError, , m_oSocket.GetErrorDescription(m_oSocket.LastError)
    End If
    If Not m_oSocket.Connect(m_uRemote.Host, m_uRemote.Port) Then
        On Error GoTo 0
        Err.Raise vbObjectError, , m_oSocket.GetErrorDescription(m_oSocket.LastError)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Public Sub CancelDownload()
    Const FUNC_NAME     As String = "CancelDownload"
    
    On Error GoTo EH
    Set m_oSocket = Nothing
    Set m_pOutputStream = Nothing
    m_lCallbackPtr = 0
    m_eState = ucsIdle
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'= private ===============================================================

Private Function pvRecvHeaders(baBuffer() As Byte) As Boolean
    Const FUNC_NAME     As String = "pvRecvHeaders"
    Const STR_DELIM     As String = vbCrLf & vbCrLf
    Dim lPos            As Long
    Dim sHeaders        As String
    Dim vSplit          As Variant
    Dim vElem           As Variant
    Dim uRedirect       As UcsParsedUrl
    
    On Error GoTo EH
    sHeaders = m_oSocket.FromTextArray(m_baRecvBuffer, ucsScpAcp) & m_oSocket.FromTextArray(baBuffer, ucsScpAcp)
    lPos = InStr(sHeaders, STR_DELIM)
    If lPos > 0 Then
        vSplit = Split(Left$(sHeaders, lPos), vbCrLf)
        Select Case Mid$(vSplit(0), 10, 3)
        Case "301", "302"
            For Each vElem In vSplit
                If Left$(LCase$(vElem), Len(HDR_LOCATION)) = HDR_LOCATION Then
                    vElem = Trim$(Mid$(vElem, Len(HDR_LOCATION) + 1))
                    If Not pvParseUrl(CStr(vElem), uRedirect) Then
                        pvSetError vbObjectError, Replace(ERR_INVALID_REDIRECT, "%1", vElem)
                        GoTo QH
                    End If
                    uRedirect.User = m_uRemote.User
                    uRedirect.Pass = m_uRemote.Pass
                    Exit For
                End If
            Next
            If LCase$(uRedirect.Protocol) <> "http" Then
                pvSetError vbObjectError, ERR_HTTP_ONLY
                GoTo QH
            End If
            m_uRemote = uRedirect
            m_oSocket.Close_
            If Not m_oSocket.Create(SocketType:=ucsSckStream) Then
                pvSetError m_oSocket.LastError
                GoTo QH
            End If
            If Not m_oSocket.Connect(m_uRemote.Host, m_uRemote.Port) Then
                pvSetError m_oSocket.LastError
                GoTo QH
            End If
        Case "200"
            m_dblContentLength = -1
            For Each vElem In vSplit
                If Left$(LCase$(vElem), Len(HDR_CONTENT_LENGTH)) = HDR_CONTENT_LENGTH Then
                    m_dblContentLength = Val(Mid$(vElem, Len(HDR_CONTENT_LENGTH) + 1))
                End If
            Next
            m_eState = ucsWaitBody
            pvRecvBody m_oSocket.ToTextArray(Mid$(sHeaders, lPos + Len(STR_DELIM)), ucsScpAcp)
        Case Else
            pvSetError vbObjectError, Replace(ERR_INVALID_RESPONSE, "%1", Mid$(vSplit(0), 10))
            GoTo QH
        End Select
    End If
    '--- success
    pvRecvHeaders = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvRecvBody(baBuffer() As Byte, Optional ByVal Flush As Boolean) As Boolean
    Const FUNC_NAME     As String = "pvRecvBody"
    Dim lIdx            As Long
    Dim hResult         As Long
    
    On Error GoTo EH
    If UBound(baBuffer) >= 0 Then
        m_dblBytesRead = m_dblBytesRead + UBound(baBuffer) + 1
        RaiseEvent DownloadProgress(m_dblBytesRead, m_dblContentLength)
        If m_lCallbackPtr <> 0 Then
            Call CallbackWeakRef.DownloadProgress(Me, m_dblBytesRead, m_dblContentLength)
        End If
        lIdx = UBound(m_baRecvBuffer) + 1
        ReDim Preserve m_baRecvBuffer(0 To lIdx + UBound(baBuffer)) As Byte
        Call CopyMemory(m_baRecvBuffer(lIdx), baBuffer(0), UBound(baBuffer) + 1)
    End If
    If UBound(m_baRecvBuffer) >= LNG_FLUSH_LIMIT Or Flush Then
        hResult = IStream_Write(m_pOutputStream, m_baRecvBuffer(0), UBound(m_baRecvBuffer) + 1)
        m_baRecvBuffer = vbNullString
        If hResult < 0 Then
            pvSetError hResult
            GoTo QH
        End If
    End If
    '--- success
    pvRecvBody = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Sub pvSetError(ByVal lNumber As Long, Optional sDescription As String)
    Const FUNC_NAME     As String = "pvSetError"
    
    On Error GoTo EH
    If LenB(sDescription) = 0 Then
        If Not m_oSocket Is Nothing Then
            sDescription = m_oSocket.GetErrorDescription(lNumber)
        Else
            sDescription = "Error " & lNumber
        End If
    End If
    CancelDownload
    RaiseEvent DownloadError(lNumber, sDescription)
    If m_lCallbackPtr <> 0 Then
        Call CallbackWeakRef.DownloadError(Me, lNumber, sDescription)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvParseUrl(sUrl As String, uParsed As UcsParsedUrl) As Boolean
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "^(.*)://(?:(?:([^:]*):)?([^@]*)@)?([A-Za-z0-9\-\.]+)(:[0-9]+)?(.*)$"
        With .Execute(sUrl)
            If .Count > 0 Then
                With .Item(0).SubMatches
                    uParsed.Protocol = .Item(0)
                    uParsed.User = .Item(1)
                    If LenB(uParsed.User) = 0 Then
                        uParsed.User = .Item(2)
                    Else
                        uParsed.Pass = .Item(2)
                    End If
                    uParsed.Host = .Item(3)
                    uParsed.Port = Val(Mid$(.Item(4), 2))
                    If uParsed.Port = 0 Then
                        Select Case LCase$(uParsed.Protocol)
                        Case "https"
                            uParsed.Port = 443
                        Case "socks5"
                            uParsed.Port = 1080
                        Case Else
                            uParsed.Port = 80
                        End Select
                    End If
                    uParsed.Path = .Item(5)
                    If LenB(uParsed.Path) = 0 Then
                        uParsed.Path = "/"
                    End If
                End With
                pvParseUrl = True
            End If
        End With
    End With
End Function

'=========================================================================
' Socket events
'=========================================================================

Private Sub m_oSocket_OnConnect()
    Const FUNC_NAME     As String = "m_oSocket_OnConnect"
    
    On Error GoTo EH
    m_eState = ucsWaitHeaders
    If Not m_oSocket.SendText("GET " & m_uRemote.Path & " HTTP/1.0" & vbCrLf & _
            "Host: " & m_uRemote.Host & vbCrLf & _
            "Accept: */*" & vbCrLf & vbCrLf, CodePage:=ucsScpAcp) Then
        pvSetError m_oSocket.LastError
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oSocket_OnReceive"
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If m_eState = ucsWaitHeaders Then
        If Not m_oSocket.ReceiveArray(baBuffer) Then
            pvSetError m_oSocket.LastError
            GoTo QH
        End If
        If UBound(baBuffer) < 0 Then
            GoTo QH
        End If
        If Not pvRecvHeaders(baBuffer) Then
            GoTo QH
        End If
    End If
    If m_eState = ucsWaitBody Then
        If Not m_oSocket.ReceiveArray(baBuffer) Then
            pvSetError m_oSocket.LastError
            GoTo QH
        End If
        If UBound(baBuffer) < 0 Then
            GoTo QH
        End If
        If Not pvRecvBody(baBuffer) Then
            GoTo QH
        End If
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnClose()
    Const FUNC_NAME     As String = "m_oSocket_OnClose"
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If m_eState = ucsWaitBody Then
        baBuffer = vbNullString
        pvRecvBody baBuffer, Flush:=True
        RaiseEvent DownloadComplete(m_sLocalFileName)
        If m_lCallbackPtr <> 0 Then
            Call CallbackWeakRef.DownloadComplete(Me, m_sLocalFileName)
        End If
    End If
    CancelDownload
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "m_oSocket_OnError"
    
    On Error GoTo EH
    pvSetError ErrorCode
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

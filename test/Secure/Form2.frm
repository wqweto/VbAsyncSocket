VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5592
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11256
   LinkTopic       =   "Form2"
   ScaleHeight     =   5592
   ScaleWidth      =   11256
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Download"
      Default         =   -1  'True
      Height          =   348
      Left            =   8484
      TabIndex        =   3
      Top             =   84
      Width           =   1356
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Left            =   84
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   504
      Width           =   9756
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ReDim Info"
      Height          =   348
      Left            =   9912
      TabIndex        =   1
      Top             =   84
      Visible         =   0   'False
      Width           =   1104
   End
   Begin VB.ComboBox cobUrl 
      Height          =   288
      Left            =   1344
      TabIndex        =   0
      Text            =   "cert-test.sandbox.google.com"
      Top             =   84
      Width           =   7068
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   348
      Left            =   252
      TabIndex        =   4
      Top             =   84
      Width           =   936
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "Form1"

'=========================================================================
' API
'=========================================================================

'--- Windows Messages
Private Const WM_SETREDRAW              As Long = &HB
Private Const EM_SETSEL                 As Long = &HB1
Private Const EM_REPLACESEL             As Long = &HC2
Private Const WM_VSCROLL                As Long = &H115
'--- for WM_VSCROLL
Private Const SB_BOTTOM                 As Long = 7

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'=========================================================================
' Constants and member variables
'=========================================================================

Private WithEvents m_oSocket    As cTlsSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_sServerName           As String
Private WithEvents m_oServerSocket As cTlsSocket
Attribute m_oServerSocket.VB_VarHelpID = -1
Private m_cRequestHandlers      As Collection
Private m_oRootCa               As cTlsSocket

Private Type UcsParsedUrl
    Protocol        As String
    Host            As String
    Port            As Long
    Path            As String
    QueryString     As String
    Anchor          As String
    User            As String
    Pass            As String
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Sub

Private Sub Command2_Click()
    txtResult.Text = DesignDumpRedimStats(Clear:=IsKeyPressed(vbKeyControl))
End Sub

'=========================================================================
' Events
'=========================================================================

Private Sub Form_Load()
    Const PEM_FILES     As String = "eccert.pem|ecprivkey.pem|fullchain2.pem"
    Const PFX_FILE      As String = "eccert.pfx"
    Const PFX_PASSWORD  As String = ""
    Dim vElem           As Variant
    Dim sAddr           As String
    Dim lPort           As Long
    
    On Error GoTo EH
    Debug.Assert pvSetVisible(Command2, True)
    If txtResult.Font.Name = "Arial" Then
        txtResult.Font.Name = "Courier New"
    End If
    For Each vElem In Split("cert-test.sandbox.google.com|tls13.1d.pw|localhost:44330|tls.ctf.network|www.mikestoolbox.org|swifttls.org|tls13.pinterjann.is|rsa8192.badssl.com|rsa4096.badssl.com|rsa2048.badssl.com|ecc384.badssl.com|ecc256.badssl.com|dir.bg|host.bg|bgdev.org|cnn.com|gmail.com|google.com|saas.bg|saas.bg:465", "|")
        cobUrl.AddItem vElem
    Next
    sAddr = GetSetting(App.Title, "Form1", "Url", cobUrl.Text)
    If LenB(sAddr) <> 0 Then
        cobUrl.Text = sAddr
    End If
    Set m_oRootCa = New cTlsSocket
    m_oRootCa.PkiPemImportRootCaCertStore App.Path & "\ca-bundle.pem"
    Set m_oServerSocket = New cTlsSocket
    ChDir App.Path
    If Not m_oServerSocket.PkiPemImportCertificates(Split(PEM_FILES, "|")) Then
        If Not m_oServerSocket.PkiPkcs12ImportCertificates(PFX_FILE, PFX_PASSWORD) Then
'            MsgBox "Error starting TLS server on localhost:10443" & vbCrLf & vbCrLf & "No private key found!", vbExclamation
'            GoTo QH
        End If
    End If
    If Not m_oServerSocket.Create(SocketPort:=10443, SocketAddress:="localhost") Then
        GoTo QH
    End If
    If Not m_oServerSocket.Listen() Then
        GoTo QH
    End If
    Set m_cRequestHandlers = New Collection
    m_oServerSocket.Socket.GetSockName sAddr, lPort
    Debug.Print "Listening on " & sAddr & ":" & lPort
QH:
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
    Resume QH
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        txtResult.Move 0, txtResult.Top, ScaleWidth, ScaleHeight - txtResult.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Form1", "Url", cobUrl.Text
    Set m_cRequestHandlers = Nothing
End Sub

Private Sub Command1_Click()
    Dim uRemote         As UcsParsedUrl
    Dim sResult         As String
    Dim sError          As String
    Dim bKeepDebug      As Boolean
    Dim dblTimer        As Double
    
    On Error GoTo EH
    dblTimer = Timer
    Screen.MousePointer = vbHourglass
    bKeepDebug = IsKeyPressed(vbKeyControl)
    If Not ParseUrl(Trim$(cobUrl.Text), uRemote, DefProtocol:="https") Then
        txtResult.Text = "Error: Invalid URL"
        GoTo QH
    End If
    txtResult.Text = vbNullString
    sResult = HttpsRequest(uRemote, sError)
    If LenB(sError) <> 0 Then
        pvAppendLogText txtResult, "Error: " & sError
        GoTo QH
    End If
    If LenB(sResult) <> 0 Then
        If Not bKeepDebug Then
            txtResult.Text = vbNullString
            pvAppendLogText txtResult, "Received " & Len(sResult) & " bytes in " & Format$(Timer - dblTimer, "0.000") & " sec" & vbCrLf
        End If
        pvAppendLogText txtResult, sResult
        txtResult.SelStart = 0
    End If
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & " [" & Err.Source & "]", vbCritical
    Set m_oSocket = Nothing
End Sub

'=========================================================================
' Methods
'=========================================================================

Private Function HttpsRequest(uRemote As UcsParsedUrl, sError As String) As String
    Const HDR_CONTENT_LENGTH As String = "content-length:"
    Const HDR_TRANSFER_ENCODING As String = "transfer-encoding:"
    Const HDR_CONNECTION As String = "connection:"
    Dim baRecv()        As Byte
    Dim sRequest        As String
    Dim bResult         As Boolean
    Dim vHeaders        As Variant
    Dim lHeaderLength   As Long
    Dim lContentLength  As Long
    Dim sEncoding       As String
    Dim sConnection     As String
    Dim vElem           As Variant
    
    If Not m_oSocket Is Nothing Then
        If m_oSocket.IsClosed Then
            Set m_oSocket = Nothing
        End If
    End If
    pvAppendLogText txtResult, "Connecting to " & uRemote.Host & vbCrLf
    If m_sServerName <> uRemote.Host & ":" & uRemote.Port Or m_oSocket Is Nothing Then
        Set m_oSocket = New cTlsSocket
        If uRemote.Host = "localhost" Then
            m_oSocket.PkiPkcs12ImportCertificates App.Path & "\client1.full.pfx"
        End If
        If Not m_oSocket.SyncConnect(uRemote.Host, uRemote.Port, _
                LocalFeatures:=IIf(pvIsKnownBadCertificate(uRemote.Host), ucsTlsIgnoreServerCertificateErrors, 0), _
                RootCa:=m_oRootCa) Then
            sError = m_oSocket.LastError.Description
            GoTo QH
        End If
        m_sServerName = uRemote.Host & ":" & uRemote.Port
    End If
    '--- send TLS application data and wait for reply
    sRequest = "GET " & uRemote.Path & uRemote.QueryString & " HTTP/1.1" & vbCrLf & _
               "Connection: keep-alive" & vbCrLf & _
               "Host: " & uRemote.Host & vbCrLf & vbCrLf
    If Not m_oSocket.SyncSendArray(StrConv(sRequest, vbFromUnicode)) Then
        sError = m_oSocket.LastError.Description
        GoTo QH
    End If
    lContentLength = -1
    Do
        bResult = m_oSocket.SyncReceiveArray(baRecv, Timeout:=15000)
        If UBound(baRecv) < 0 Then
            If m_oSocket.IsClosed Then
                Set m_oSocket = Nothing
                Exit Do
            End If
            If Not bResult Then
                sError = m_oSocket.LastError.Description
                GoTo QH
            End If
        Else
            HttpsRequest = HttpsRequest & StrConv(baRecv, vbUnicode)
'            Debug.Print "Len(HttpsRequest)=" & Len(HttpsRequest), Timer
        End If
        If IsEmpty(vHeaders) Then
            lHeaderLength = InStr(1, HttpsRequest, vbCrLf & vbCrLf) - 1
            If lHeaderLength > 0 Then
                vHeaders = Split(Left$(HttpsRequest, lHeaderLength), vbCrLf)
                lHeaderLength = lHeaderLength + 4
                For Each vElem In vHeaders
                    If Left$(LCase(vElem), Len(HDR_CONTENT_LENGTH)) = HDR_CONTENT_LENGTH Then
                        lContentLength = Val(Mid$(vElem, Len(HDR_CONTENT_LENGTH) + 1))
                    ElseIf Left$(LCase(vElem), Len(HDR_TRANSFER_ENCODING)) = HDR_TRANSFER_ENCODING Then
                        sEncoding = LCase$(Trim$(Mid$(vElem, Len(HDR_TRANSFER_ENCODING) + 1)))
                    ElseIf Left$(LCase(vElem), Len(HDR_CONNECTION)) = HDR_CONNECTION Then
                        sConnection = LCase$(Trim$(Mid$(vElem, Len(HDR_CONNECTION) + 1)))
                    End If
                Next
            End If
        End If
        If lContentLength >= 0 Then
            If Len(HttpsRequest) >= lHeaderLength + lContentLength Then
                If Len(HttpsRequest) <> lHeaderLength + lContentLength Then
                    Debug.Print "Warning: Received " & Len(HttpsRequest) & " instead of " & lHeaderLength + lContentLength
                End If
                Exit Do
            End If
        ElseIf sEncoding = "chunked" Then
            If Right$(HttpsRequest, 7) = vbCrLf & "0" & vbCrLf & vbCrLf Then
                Exit Do
            End If
        End If
        If Not bResult Then
            sError = m_oSocket.LastError.Description
            GoTo QH
        End If
    Loop
    If Not m_oSocket Is Nothing And sConnection = "close" Then
        m_oSocket.ShutDown
        Set m_oSocket = Nothing
    End If
QH:
    HttpsRequest = Replace(Replace(HttpsRequest, vbCr, vbNullString), vbLf, vbCrLf)
    If LenB(sError) <> 0 Then
        Set m_oSocket = Nothing
    End If
End Function

Private Function pvIsKnownBadCertificate(sHost As String) As Boolean
    Const STR_HOSTS     As String = "mikestoolbox.org|localhost"
    Dim vElem           As Variant
    
    For Each vElem In Split(STR_HOSTS, "|")
        If Right$(LCase$(sHost), Len(vElem)) = vElem Then
            pvIsKnownBadCertificate = True
            Exit For
        End If
    Next
End Function

Private Function ParseUrl(sUrl As String, uParsed As UcsParsedUrl, Optional DefProtocol As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "^(?:(.*)://)?(?:(?:([^:]*):)?([^@]*)@)?([A-Za-z0-9\-\.]+)(:[0-9]+)?(/[^?#]*)?(\?[^#]*)?(#.*)?$"
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
                        Select Case LCase$(IIf(LenB(uParsed.Protocol) = 0, DefProtocol, uParsed.Protocol))
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
                    uParsed.QueryString = .Item(6)
                    uParsed.Anchor = .Item(7)
                End With
                ParseUrl = True
            End If
        End With
    End With
End Function

Private Sub pvAppendLogText(txtLog As TextBox, sValue As String)
    Call SendMessage(txtLog.hWnd, WM_SETREDRAW, 0, ByVal 0)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, 0, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, -1, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_REPLACESEL, 1, ByVal sValue)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, 0, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, -1, ByVal -1)
    Call SendMessage(txtLog.hWnd, WM_SETREDRAW, 1, ByVal 0)
    Call SendMessage(txtLog.hWnd, WM_VSCROLL, SB_BOTTOM, ByVal 0)
End Sub

Public Function IsKeyPressed(ByVal lVirtKey As KeyCodeConstants) As Boolean
    IsKeyPressed = ((GetAsyncKeyState(lVirtKey) And &H8000) = &H8000)
End Function

Private Sub m_oServerSocket_OnAccept()
    Const FUNC_NAME     As String = "m_oServerSocket_OnAccept"
    Dim oSocket         As cTlsSocket
    Dim oHandler        As cRequestHandler
    Dim sKey            As String
    
    On Error GoTo EH
    If Not m_oServerSocket.Accept(oSocket) Then
        GoTo QH
    End If
    Set oHandler = New cRequestHandler
    sKey = "#" & ObjPtr(oHandler)
    If Not oHandler.Init(oSocket, sKey, Me) Then
        GoTo QH
    End If
    m_cRequestHandlers.Add oHandler, "#" & ObjPtr(oHandler)
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Friend Sub frRemoveHandler(sKey As String)
    RemoveCollection m_cRequestHandlers, sKey
End Sub

Private Function pvSetVisible(oCtl As Object, ByVal bValue As Boolean) As Boolean
    oCtl.Visible = bValue
    pvSetVisible = True
End Function

Private Sub m_oSocket_OnResolve(IpAddress As String)
    Debug.Print "m_oSocket_OnResolve, IpAddress=" & IpAddress, Timer
End Sub

Private Sub m_oSocket_OnConnect()
    Debug.Print "m_oSocket_OnConnect", Timer
End Sub

Private Sub m_oSocket_OnReceive()
    Debug.Print "m_oSocket_OnReceive", Timer
End Sub

Private Sub m_oSocket_OnSend()
    Debug.Print "m_oSocket_OnSend", Timer
End Sub

Private Sub m_oSocket_OnClose()
    Debug.Print "m_oSocket_OnClose", Timer
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Debug.Print "m_oSocket_OnError, m_oSocket.LastError=&H" & Hex$(m_oSocket.LastError.Number) & ", " & m_oSocket.LastError.Description, Timer
End Sub

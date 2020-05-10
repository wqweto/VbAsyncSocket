VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4968
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5448
   LinkTopic       =   "Form1"
   ScaleHeight     =   4968
   ScaleWidth      =   5448
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkUseHttps 
      Caption         =   "Use HTTPS"
      Height          =   264
      Left            =   2688
      TabIndex        =   14
      Top             =   4284
      Value           =   1  'Checked
      Width           =   1608
   End
   Begin VB.TextBox txtBandwidth 
      Height          =   288
      Left            =   4452
      TabIndex        =   13
      Text            =   "1024"
      Top             =   3948
      Width           =   852
   End
   Begin VB.CheckBox chkRateLimit 
      Caption         =   "Rate limit (KB/s):"
      Height          =   264
      Left            =   2688
      TabIndex        =   12
      Top             =   3948
      Value           =   1  'Checked
      Width           =   1608
   End
   Begin VB.CommandButton Command12 
      Caption         =   "HttpUpload"
      Height          =   432
      Left            =   168
      TabIndex        =   11
      Top             =   4368
      Width           =   2364
   End
   Begin VB.CommandButton Command11 
      Caption         =   "HttpDownload"
      Height          =   432
      Left            =   168
      TabIndex        =   10
      Top             =   3864
      Width           =   2364
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Sync operations"
      Height          =   432
      Left            =   2688
      TabIndex        =   9
      Top             =   252
      Width           =   2364
   End
   Begin VB.TextBox txtProxy 
      Height          =   288
      Left            =   2688
      TabIndex        =   4
      Text            =   "user:pass@80.252.241.107:1080"
      ToolTipText     =   "SOCKS5 Proxy Address"
      Top             =   2268
      Width           =   2616
   End
   Begin VB.CheckBox chkProxy 
      Caption         =   "Use proxy:"
      Height          =   276
      Left            =   2688
      TabIndex        =   3
      Top             =   1932
      Width           =   1104
   End
   Begin VB.CommandButton Command9 
      Caption         =   "wss:// protocol"
      Height          =   432
      Left            =   168
      TabIndex        =   8
      Top             =   3360
      Width           =   2364
   End
   Begin VB.CommandButton Command8 
      Caption         =   "expired.badssl.com"
      Height          =   432
      Left            =   168
      TabIndex        =   7
      Top             =   2856
      Width           =   2364
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SMTP with STARTTLS"
      Height          =   432
      Left            =   2688
      TabIndex        =   6
      Top             =   1008
      Width           =   2364
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Client Certificate"
      Height          =   432
      Left            =   168
      TabIndex        =   5
      Top             =   2352
      Width           =   2364
   End
   Begin VB.CommandButton Command5 
      Caption         =   "cTlsClient HTTPS"
      Height          =   432
      Left            =   168
      TabIndex        =   2
      Top             =   1848
      Width           =   2364
   End
   Begin VB.CommandButton Command4 
      Caption         =   "cTlsSocket SMTP over SSL"
      Height          =   432
      Left            =   168
      TabIndex        =   1
      Top             =   1008
      Width           =   2364
   End
   Begin VB.CommandButton Command1 
      Caption         =   "cAsyncSocket"
      Height          =   432
      Left            =   168
      TabIndex        =   0
      Top             =   252
      Width           =   2364
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

Private WithEvents m_oSocket As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private WithEvents m_oHttpDownload As cHttpDownload
Attribute m_oHttpDownload.VB_VarHelpID = -1
Private m_oRateLimiter As cRateLimiter
Private m_dblStartTimerEx As Double

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

Private Sub Command1_Click()
    Dim sName As String
    Dim sAddr As String
    Dim lPort As Long
    
    On Error GoTo EH
    Set m_oSocket = New cAsyncSocket
    With m_oSocket
        .GetLocalHost sName, sAddr
        Debug.Print Format$(TimerEx, "0.000"), "GetLocalHost=" & sName & ", " & sAddr
        .Create EventMask:=ucsSfdConnect Or ucsSfdRead
        .Connect "www.bgdev.org", 80
        .GetPeerName sAddr, lPort
        Debug.Print Format$(TimerEx, "0.000"), "GetPeerName=" & sAddr & ":" & lPort
        .GetSockName sAddr, lPort
        Debug.Print Format$(TimerEx, "0.000"), "GetSockName=" & sAddr & ":" & lPort
    End With
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub m_oSocket_OnConnect()
    Debug.Print Format$(TimerEx, "0.000"), "OnConnect"
    m_oSocket.SendArray ToUtf8Array("GET / HTTP/1.0" & vbCrLf & _
        "Host: www.bgdev.org" & vbCrLf & _
        "Connection: close" & vbCrLf & vbCrLf)
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Debug.Print Format$(TimerEx, "0.000"), "OnError, ErrorCode=" & ErrorCode & ", EventMask=" & EventMask & ", Desc=" & m_oSocket.GetErrorDescription(ErrorCode)
End Sub

Private Sub m_oSocket_OnResolve(Address As String)
    Debug.Print Format$(TimerEx, "0.000"), "OnResolve, Address=" & Address
End Sub

Private Sub m_oSocket_OnSend()
    Debug.Print Format$(TimerEx, "0.000"), "OnSend"
End Sub

Private Sub m_oSocket_OnReceive()
    Dim baBuffer()      As Byte
    Dim lBytes          As Long
    
    Debug.Print Format$(TimerEx, "0.000"), "OnReceive"
    lBytes = m_oSocket.AvailableBytes
    If lBytes > 0 Then
        ReDim baBuffer(0 To lBytes - 1) As Byte
    Else
        ReDim baBuffer(0 To 4096 - 1) As Byte
    End If
    lBytes = m_oSocket.Receive(VarPtr(baBuffer(0)), UBound(baBuffer) + 1)
    If lBytes > 0 Then
        ReDim Preserve baBuffer(0 To lBytes - 1) As Byte
        Debug.Print Format$(TimerEx, "0.000"), Replace(FromUtf8Array(baBuffer), vbCrLf, "\n")
    End If
End Sub

Private Sub m_oSocket_OnClose()
    Debug.Print Format$(TimerEx, "0.000"), "OnClose"
End Sub

Private Sub m_oSocket_OnAccept()
    Debug.Print Format$(TimerEx, "0.000"), "OnAccept"
End Sub

Private Sub Command10_Click()
    Const LNG_TIMEOUT   As Long = 5000
    Dim sRequest        As String
    Dim baBuffer()      As Byte
    Dim sResponse       As String
    
    With New cAsyncSocket
        If Not .SyncConnect("bgdev.org", 80, Timeout:=LNG_TIMEOUT) Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "Connected"
        sRequest = "GET / HTTP/1.0" & vbCrLf & _
            "Host: www.bgdev.org" & vbCrLf & _
            "Connection: close" & vbCrLf & vbCrLf
        If Not .SyncSendText(sRequest, Timeout:=LNG_TIMEOUT) Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "->", sRequest
        If Rnd < 0.5 Then
            sResponse = .SyncReceiveText(10000, Timeout:=LNG_TIMEOUT)
            If LenB(sResponse) <> 0 Then
                Debug.Print Format$(TimerEx, "0.000"), "<-", sResponse
                Debug.Print Format$(TimerEx, "0.000"), "Size", Len(sResponse) & " chars"
            Else
                GoTo QH
            End If
        Else
            If Not .SyncReceiveArray(baBuffer, 10000, Timeout:=LNG_TIMEOUT) Then
                If UBound(baBuffer) >= 0 Then
                    Debug.Print Format$(TimerEx, "0.000"), "<-", FromUtf8Array(baBuffer)
                    Debug.Print Format$(TimerEx, "0.000"), "Trimmed", UBound(baBuffer) + 1 & " bytes"
                End If
                GoTo QH
            End If
            Debug.Print Format$(TimerEx, "0.000"), "<-", FromUtf8Array(baBuffer)
        End If
        Exit Sub
QH:
        Debug.Print Format$(TimerEx, "0.000"), "Error", .GetErrorDescription(.LastError)
    End With
End Sub

Private Sub Command4_Click()
    Dim oTlsSocket      As cTlsSocket
    Dim baBuffer()      As Byte

    Screen.MousePointer = vbHourglass
    Debug.Print Format$(TimerEx, "0.000"), "Connect secure socket to port 465"
    Set oTlsSocket = New cTlsSocket
    If Not oTlsSocket.SyncConnect("smtp.gmail.com", 465) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "TLS handshake complete: " & oTlsSocket.RemoteHostName
    If Not oTlsSocket.SyncReceiveArray(baBuffer) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", FromUtf8Array(baBuffer);
    Debug.Assert Left$(FromUtf8Array(baBuffer), 3) = "220"
    Debug.Print Format$(TimerEx, "0.000"), "<-", "QUIT"
    If Not oTlsSocket.SyncSendArray(ToUtf8Array("QUIT" & vbCrLf)) Then
        GoTo QH
    End If
    If Not oTlsSocket.SyncReceiveArray(baBuffer) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", FromUtf8Array(baBuffer);
    Screen.MousePointer = vbDefault
    Exit Sub
QH:
    With oTlsSocket.LastError
        Debug.Print Format$(TimerEx, "0.000"), Hex$(.Number) & ": " & .Description & " at " & .Source
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command5_Click()
    Dim oTlsSocket      As cTlsSocket
    Dim sHeaders        As String
    Dim sResponse       As String
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim sUrl            As String
    Dim sProxy          As String

    Screen.MousePointer = vbHourglass
    sUrl = "https://www.google.com"
    If chkProxy.Value = vbChecked Then
        sProxy = txtProxy.Text
    End If
    Debug.Print Format$(TimerEx, "0.000"), "Open " & sUrl
Repeat:
    Set oTlsSocket = pvInitHttpRequest(sUrl, sProxy)
    If oTlsSocket Is Nothing Then
        GoTo QH
    End If
    sHeaders = vbNullString
    Do
        sResponse = oTlsSocket.SyncReceiveText()
        If LenB(sResponse) = 0 Then
            sHeaders = vbNullString
            Exit Do
        End If
        sHeaders = sHeaders & sResponse
        lIdx = InStr(sHeaders, vbCrLf & vbCrLf)
        If lIdx > 0 Then
            vSplit = Split(Left$(sHeaders, lIdx - 1), vbCrLf)
            Exit Do
        End If
    Loop
    If IsArray(vSplit) Then
        Debug.Print Format$(TimerEx, "0.000"), Join(vSplit, vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
        If Mid$(sHeaders, 10, 3) = "302" Then
            For lIdx = 0 To UBound(vSplit)
                If Left$(vSplit(lIdx), 9) = "Location:" Then
                    sUrl = Trim$(Mid$(vSplit(lIdx), 10))
                    Debug.Print Format$(TimerEx, "0.000"), "Redirect to " & sUrl
                    GoTo Repeat
                End If
            Next
        End If
    End If
    oTlsSocket.Close_
    Debug.Print Format$(TimerEx, "0.000"), "Done"
    Screen.MousePointer = vbDefault
    Exit Sub
QH:
    If Not oTlsSocket Is Nothing Then
        With oTlsSocket.LastError
            Debug.Print Format$(TimerEx, "0.000"), Hex$(.Number) & ": " & .Description & " at " & .Source
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Function pvInitHttpRequest(sUrl As String, Optional sProxyUrl As String, Optional ByVal LocalFeature As UcsTlsLocalFeaturesEnum) As cTlsSocket
    Dim oRetVal         As cTlsSocket
    Dim uRemote         As UcsParsedUrl
    Dim uProxy          As UcsParsedUrl
    Dim baBuffer()      As Byte
    
    If Not pvParseUrl(sUrl, uRemote) Then
        GoTo QH
    End If
    Set oRetVal = New cTlsSocket
    If Not pvParseUrl(sProxyUrl, uProxy, "socks5") Then
        If Not oRetVal.SyncConnect(uRemote.Host, uRemote.Port, UseTls:=False) Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "Connected to " & uRemote.Host & ":" & uRemote.Port
    Else
        If Not oRetVal.SyncConnect(uProxy.Host, uProxy.Port, UseTls:=False) Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "Tunnel to " & uProxy.Host & ":" & uProxy.Port
        If LenB(uProxy.User) <> 0 Then
            If Not oRetVal.SyncSendArray(pvArrayByte(5, 2, 0, 2)) Then
                GoTo QH
            End If
        Else
            If Not oRetVal.SyncSendArray(pvArrayByte(5, 1, 0)) Then
                GoTo QH
            End If
        End If
        If Not oRetVal.SyncReceiveArray(baBuffer) Then
            GoTo QH
        End If
        If UBound(baBuffer) < 1 Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "Proxy auth method chosen: " & baBuffer(1)
        If baBuffer(1) = 2 Then
            oRetVal.SyncSendArray pvArrayByte(1)
            baBuffer = oRetVal.Socket.ToTextArray(uProxy.User, ucsScpUtf8)
            oRetVal.SyncSendArray pvArrayByte(UBound(baBuffer) + 1)
            oRetVal.SyncSendArray baBuffer
            baBuffer = oRetVal.Socket.ToTextArray(uProxy.Pass, ucsScpUtf8)
            oRetVal.SyncSendArray pvArrayByte(UBound(baBuffer) + 1)
            oRetVal.SyncSendArray baBuffer
            If Not oRetVal.SyncReceiveArray(baBuffer) Then
                GoTo QH
            End If
            If UBound(baBuffer) < 1 Then
                GoTo QH
            End If
            Debug.Print Format$(TimerEx, "0.000"), "Authentication result: " & baBuffer(1)
            If baBuffer(1) <> 0 Then
                GoTo QH
            End If
        End If
        oRetVal.SyncSendArray pvArrayByte(5, 1, 0, 3) '--- 5 = version, 1 = TCP stream conn, 0 = reserved, 3 = domain name
        baBuffer = oRetVal.ToTextArray(uRemote.Host, ucsScpUtf8)
        oRetVal.SyncSendArray pvArrayByte(UBound(baBuffer) + 1)
        oRetVal.SyncSendArray baBuffer
        oRetVal.SyncSendArray pvArrayByte(uRemote.Port \ &H100, uRemote.Port And &HFF)
        If Not oRetVal.SyncReceiveArray(baBuffer) Then
            GoTo QH
        End If
        If UBound(baBuffer) < 3 Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "Proxy connection to " & uRemote.Host & ":" & uRemote.Port & " status: " & baBuffer(1)
        If baBuffer(1) <> 0 Then
            GoTo QH
        End If
        If baBuffer(3) = 1 Then
            Debug.Print Format$(TimerEx, "0.000"), "Connection info: " & baBuffer(4) & "." & baBuffer(5) & "." & baBuffer(6) & "." & baBuffer(7) & ":" & baBuffer(8) * 256& + baBuffer(9)
        End If
    End If
    If LCase$(uRemote.Protocol) = "https" Then
        If Not oRetVal.SyncStartTls(uRemote.Host, LocalFeature) Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "TLS handshake complete"
    End If
    If Not oRetVal.SyncSendText("GET " & uRemote.Path & uRemote.QueryString & " HTTP/1.0" & vbCrLf & _
            "Host: " & uRemote.Host & vbCrLf & _
            "Connection: close" & vbCrLf & vbCrLf) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "Request sent"
    Set pvInitHttpRequest = oRetVal
    Exit Function
QH:
    If Not oRetVal Is Nothing Then
        With oRetVal.LastError
            Debug.Print Format$(TimerEx, "0.000"), "&H" & Hex$(.Number) & ": " & .Description & " at " & .Source
        End With
    End If
End Function

Private Function pvParseUrl(sUrl As String, uParsed As UcsParsedUrl, Optional DefProtocol As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "^(?:(?:(.+):)?//)?(?:(?:([^:]*):)?([^@]*)@)?([A-Za-z0-9\-\.]+)(:[0-9]+)?(/[^?#]*)?(\?[^#]*)?(#.*)?$"
        With .Execute(sUrl)
            If .Count > 0 Then
                With .Item(0).SubMatches
                    uParsed.Protocol = IIf(LenB(.Item(0)) = 0, DefProtocol, .Item(0))
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
                    uParsed.QueryString = .Item(6)
                    uParsed.Anchor = .Item(7)
                End With
                pvParseUrl = True
            End If
        End With
    End With
End Function

Private Function pvArrayByte(ParamArray A() As Variant) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    ReDim baRetVal(0 To UBound(A))
    For lIdx = 0 To UBound(A)
        baRetVal(lIdx) = A(lIdx)
    Next
    pvArrayByte = baRetVal
End Function

Private Sub Command6_Click()
    Dim sUrl            As String
    Dim sProxy          As String
    Dim oTlsSocket      As cTlsSocket
    
    sUrl = "https://server.cryptomix.com/secure/"
    If chkProxy.Value = vbChecked Then
        sProxy = txtProxy.Text
    End If
    Debug.Print Format$(TimerEx, "0.000"), "Open " & sUrl
    Set oTlsSocket = pvInitHttpRequest(sUrl, sProxy)
    If oTlsSocket Is Nothing Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), oTlsSocket.SyncReceiveText()
    Debug.Print Format$(TimerEx, "0.000"), oTlsSocket.SyncReceiveText()
    Debug.Print Format$(TimerEx, "0.000"), "Done"
    Exit Sub
QH:
    If Not oTlsSocket Is Nothing Then
        With oTlsSocket.LastError
            Debug.Print Format$(TimerEx, "0.000"), Hex$(.Number) & ": " & .Description & " at " & .Source
        End With
    End If
End Sub

Private Sub Command7_Click()
    Dim oTlsSocket      As cTlsSocket
    Dim sResponse       As String
    Dim sRequest        As String

    Screen.MousePointer = vbHourglass
    Debug.Print Format$(TimerEx, "0.000"), "Connect to port 587"
    Set oTlsSocket = New cTlsSocket
    If Not oTlsSocket.SyncConnect("smtp.gmail.com", 587, UseTls:=False) Then
        GoTo QH
    End If
    sResponse = oTlsSocket.SyncReceiveText()
    If LenB(sResponse) = 0 Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", sResponse;
    sRequest = "HELO " & pvGetExternalIP & vbCrLf
    If Not oTlsSocket.SyncSendText(sRequest) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "<-", sRequest;
    sResponse = oTlsSocket.SyncReceiveText()
    If LenB(sResponse) = 0 Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", sResponse;
    sRequest = "STARTTLS" & vbCrLf
    If Not oTlsSocket.SyncSendText(sRequest) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "<-", sRequest;
    sResponse = oTlsSocket.SyncReceiveText()
    If LenB(sResponse) = 0 Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", sResponse;
    If Not oTlsSocket.SyncStartTls("smtp.gmail.com") Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "TLS handshake complete: " & oTlsSocket.RemoteHostName
    sRequest = "NOOP" & vbCrLf
    If Not oTlsSocket.SyncSendText(sRequest) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "<-", sRequest;
    sResponse = oTlsSocket.SyncReceiveText()
    If LenB(sResponse) = 0 Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", sResponse;
    sRequest = "QUIT" & vbCrLf
    If Not oTlsSocket.SyncSendText(sRequest) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "<-", sRequest;
    sResponse = oTlsSocket.SyncReceiveText()
    If LenB(sResponse) = 0 Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", sResponse
    Screen.MousePointer = vbDefault
    Exit Sub
QH:
    With oTlsSocket.LastError
        Debug.Print Format$(TimerEx, "0.000"), Hex$(.Number) & ": " & .Description & " at " & .Source
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Function pvGetExternalIP() As String
    Dim sResponse     As String
    
    With New cAsyncSocket
        .SyncConnect "ifconfig.co", 80
        .SyncSendText "GET /ip HTTP/1.1" & vbCrLf & "Host: ifconfig.co" & vbCrLf & vbCrLf
        Do
            sResponse = sResponse & .SyncReceiveText()
            If InStr(sResponse, vbCrLf & vbCrLf) > 0 Then
                sResponse = At(Split(At(Split(sResponse, vbCrLf & vbCrLf), 1), vbLf), 0)
                If sResponse Like "*.*.*.*" Then
                    Exit Do
                End If
            End If
            If .LastError <> 0 Then
                .GetSockName sResponse, 0
                Exit Do
            End If
        Loop
    End With
    If sResponse Like "*.*.*.*" Then
        pvGetExternalIP = sResponse
    Else
        pvGetExternalIP = "127.0.0.1"
    End If
End Function

Private Sub Command8_Click()
    Dim sResponse       As String
    Dim sUrl            As String
    Dim sProxy          As String
    
'    sUrl = "https://www.howsmyssl.com/a/check"
    sUrl = "https://expired.badssl.com/"
    If chkProxy.Value = vbChecked Then
        sProxy = txtProxy.Text
    End If
    With pvInitHttpRequest(sUrl, sProxy, ucsTlsIgnoreServerCertificateErrors)
        sResponse = sResponse & .SyncReceiveText(1)
    End With
    Debug.Print Format$(TimerEx, "0.000"), sResponse
End Sub

Private Sub Command9_Click()
    Dim oTlsSocket      As cTlsSocket
    Dim baBuffer()      As Byte
    
    Screen.MousePointer = vbHourglass
    Debug.Print Format$(TimerEx, "0.000"), "Connect secure socket to port 443"
    Set oTlsSocket = New cTlsSocket
    If Not oTlsSocket.SyncConnect("connect-bot.classic.blizzard.com", 443) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "TLS handshake complete: " & oTlsSocket.RemoteHostName
    If Not oTlsSocket.SyncSendText("GET /v1/rpc/chat HTTP/1.1" & vbCrLf & _
                "Host: connect-bot.classic.blizzard.com" & vbCrLf & _
                "Upgrade: websocket" & vbCrLf & _
                "Connection: Upgrade" & vbCrLf & _
                "Sec-WebSocket-Key: x3JJHMbDL1EzLkh9GBhXDw==" & vbCrLf & _
                "Sec-WebSocket-Protocol: chat, superchat" & vbCrLf & _
                "Sec-WebSocket-Version: 13" & vbCrLf & _
                "Origin: http://connect-bot.classic.blizzard.com/v1/rpc/chat" & vbCrLf & vbCrLf) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "<-", "(HTTP request)"
    If Not oTlsSocket.SyncReceiveArray(baBuffer) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", FromUtf8Array(baBuffer)
QH:
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command11_Click()
    Set m_oHttpDownload = New cHttpDownload
    m_oHttpDownload.DownloadFile IIf(chkUseHttps.Value = vbChecked, "https", "http") & "://dl.unicontsoft.com/upload/pix/ss_vbyoga_flex_container.gif", Environ$("TMP") & "\aaa.gif"
End Sub

Private Sub m_oHttpDownload_OperationStart()
    m_dblStartTimerEx = TimerEx
    If chkRateLimit.Value = vbChecked And Val(txtBandwidth.Text) > 0 Then
        Set m_oRateLimiter = New cRateLimiter
        m_oRateLimiter.Init m_oHttpDownload.Socket, Val(txtBandwidth.Text) * 1024
    Else
        Set m_oRateLimiter = Nothing
    End If
End Sub

Private Sub m_oHttpDownload_DownloadProgress(ByVal BytesRead As Double, ByVal BytesTotal As Double)
    Debug.Print Format$(TimerEx, "0.000"), "Downloaded " & BytesRead & " from " & BytesTotal & " @ " & Format$(BytesRead / (TimerEx - m_dblStartTimerEx) / 1024, "0.0") & "KB/s"
    Caption = "Downloaded " & BytesRead & " from " & BytesTotal & " @ " & Format$(BytesRead / (TimerEx - m_dblStartTimerEx) / 1024, "0.0") & "KB/s"
'    If BytesRead > 2000000 Then
'        m_oHttpDownload.CancelOperation
'        Set m_oRateLimiter = Nothing
'    End If
End Sub

Private Sub m_oHttpDownload_DownloadComplete(ByVal LocalFileName As String)
    Debug.Print Format$(TimerEx, "0.000"), "Download to " & LocalFileName & " complete"
    MsgBox "Download to " & LocalFileName & " complete", vbExclamation
End Sub

Private Sub m_oHttpDownload_OperationError(ByVal Number As Long, ByVal Description As String)
    Debug.Print Format$(TimerEx, "0.000"), Hex$(Number) & ": " & Description
    MsgBox Description, vbCritical, TypeName(m_oHttpDownload)
End Sub

Private Sub Command12_Click()
    Set m_oHttpDownload = New cHttpDownload
    m_oHttpDownload.UploadFile IIf(chkUseHttps.Value = vbChecked, "https", "http") & "://www.unicontsoft.com/upload_errors.php?id=deldeldel", Environ$("TMP") & "\aaa.gif"
End Sub

Private Sub m_oHttpDownload_UploadProgress(ByVal BytesWritten As Double, ByVal BytesTotal As Double)
    Debug.Print Format$(TimerEx, "0.000"), "Uploaded " & BytesWritten & " of " & BytesTotal & " @ " & Format$(BytesWritten / (TimerEx - m_dblStartTimerEx) / 1024, "0.0") & "KB/s"
    Caption = "Uploaded " & BytesWritten & " of " & BytesTotal
End Sub

Private Sub m_oHttpDownload_UploadComplete(ByVal LocalFileName As String)
    Debug.Print Format$(TimerEx, "0.000"), "Upload of " & LocalFileName & " complete"
    MsgBox "Upload of " & LocalFileName & " complete", vbExclamation
End Sub

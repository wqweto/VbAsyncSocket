VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5544
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5436
   LinkTopic       =   "Form1"
   ScaleHeight     =   5544
   ScaleWidth      =   5436
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtProxy 
      Height          =   288
      Left            =   2688
      TabIndex        =   8
      Text            =   "user:pass@80.252.241.107:1080"
      ToolTipText     =   "SOCKS5 Proxy Address"
      Top             =   3360
      Width           =   2616
   End
   Begin VB.CheckBox chkProxy 
      Caption         =   "Use proxy:"
      Height          =   276
      Left            =   2688
      TabIndex        =   7
      Top             =   3024
      Width           =   1104
   End
   Begin VB.CommandButton Command9 
      Caption         =   "wss:// protocol"
      Height          =   432
      Left            =   168
      TabIndex        =   12
      Top             =   4452
      Width           =   2364
   End
   Begin VB.CommandButton Command8 
      Caption         =   "expired.badssl.com"
      Height          =   432
      Left            =   168
      TabIndex        =   11
      Top             =   3948
      Width           =   2364
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SMTP with STARTTLS"
      Height          =   432
      Left            =   2688
      TabIndex        =   10
      Top             =   2436
      Width           =   2364
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Client Certificate"
      Height          =   432
      Left            =   168
      TabIndex        =   9
      Top             =   3444
      Width           =   2364
   End
   Begin VB.CommandButton Command5 
      Caption         =   "cTlsClient HTTPS"
      Height          =   432
      Left            =   168
      TabIndex        =   6
      Top             =   2940
      Width           =   2364
   End
   Begin VB.CommandButton Command4 
      Caption         =   "cTlsClient SMTP over SSL"
      Height          =   432
      Left            =   168
      TabIndex        =   5
      Top             =   2436
      Width           =   2364
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Keep-alive"
      Height          =   264
      Left            =   3864
      TabIndex        =   3
      Top             =   1428
      Width           =   1188
   End
   Begin VB.CommandButton Command3 
      Caption         =   "cWinSockRequest 5123/udp"
      Height          =   432
      Left            =   168
      TabIndex        =   4
      Top             =   1596
      Width           =   2364
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Async"
      Height          =   264
      Left            =   2688
      TabIndex        =   2
      Top             =   1428
      Width           =   1188
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cWinSockRequest 80/tcp"
      Height          =   432
      Left            =   168
      TabIndex        =   1
      Top             =   1092
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
Private WithEvents m_oRequest As cWinSockRequest
Attribute m_oRequest.VB_VarHelpID = -1

Private Sub Command1_Click()
    Dim sName As String
    Dim sAddr As String
    Dim lPort As Long
    
    On Error GoTo EH
    Set m_oSocket = New cAsyncSocket
    m_oSocket.GetLocalHost sName, sAddr
    Debug.Print "GetLocalHost=" & sName & ", " & sAddr, Format$(TimerEx, "0.000")
    m_oSocket.Create EventMask:=ucsSfdConnect Or ucsSfdRead
    m_oSocket.Connect "www.bgdev.org", 80
    m_oSocket.GetPeerName sAddr, lPort
    Debug.Print "GetPeerName=" & sAddr & ":" & lPort, Format$(TimerEx, "0.000")
    m_oSocket.GetSockName sAddr, lPort
    Debug.Print "GetSockName=" & sAddr & ":" & lPort, Format$(TimerEx, "0.000")
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub m_oSocket_OnConnect()
'    Dim baBuffer()      As Byte
'    Dim lBytes          As Long
'
'    baBuffer = ToTextArray("GET / HTTP/1.0" & vbCrLf & _
'        "Host: www.bgdev.org" & vbCrLf & _
'        "Connection: close" & vbCrLf & vbCrLf)
'    Do
'        lBytes = lBytes + m_oSocket.SendPtr(VarPtr(baBuffer(lBytes)), UBound(baBuffer) + 1 - lBytes)
'    Loop While lBytes <= UBound(baBuffer)
    Debug.Print "OnConnect", Format$(TimerEx, "0.000")
    m_oSocket.SendArray ToUtf8Array("GET / HTTP/1.0" & vbCrLf & _
        "Host: www.bgdev.org" & vbCrLf & _
        "Connection: close" & vbCrLf & vbCrLf)
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Debug.Print "OnError, ErrorCode=" & ErrorCode & ", EventMask=" & EventMask & ", Desc=" & m_oSocket.GetErrorDescription(ErrorCode), Format$(TimerEx, "0.000"), Format$(TimerEx, "0.000")
End Sub

Private Sub m_oSocket_OnResolve(Address As String)
    Debug.Print "OnResolve, Address=" & Address, Format$(TimerEx, "0.000")
End Sub

Private Sub m_oSocket_OnSend()
    Debug.Print "OnSend", Format$(TimerEx, "0.000")
End Sub

Private Sub m_oSocket_OnReceive()
    Dim baBuffer()      As Byte
    Dim lBytes          As Long
    
    Debug.Print "OnReceive", Format$(TimerEx, "0.000")
    lBytes = m_oSocket.AvailableBytes
    If lBytes > 0 Then
        ReDim baBuffer(0 To lBytes - 1) As Byte
    Else
        ReDim baBuffer(0 To 4096 - 1) As Byte
    End If
    lBytes = m_oSocket.Receive(VarPtr(baBuffer(0)), UBound(baBuffer) + 1)
    If lBytes > 0 Then
        ReDim Preserve baBuffer(0 To lBytes - 1) As Byte
        Debug.Print Replace(FromUtf8Array(baBuffer), vbCrLf, "\n")
    End If
End Sub

Private Sub m_oSocket_OnClose()
    Debug.Print "OnClose", Format$(TimerEx, "0.000")
End Sub

Private Sub m_oSocket_OnAccept()
    Debug.Print "OnAccept", Format$(TimerEx, "0.000")
End Sub

Private Sub Command2_Click()
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    Set m_oRequest = New cWinSockRequest
    m_oRequest.SetTimeouts 0, 5000, 5000, 5000, 50
    m_oRequest.Open_ "vws03:100", Async:=(Check1.Value = vbChecked)
    m_oRequest.Send "GET /product/likyor-trakiyska-roza-0-2-podarachna-kutia-likyor-trakiyska-roza-7305 HTTP/1.1" & vbCrLf & _
        "Host: vws03:100" & vbCrLf & _
        "Connection: " & IIf(Check2.Value = vbChecked, "keep-alive", "close") & vbCrLf & vbCrLf
'    If Check1.Value = vbChecked Then
'        m_oRequest.WaitForResponse 5000
'    End If
'    Debug.Print Replace(m_oRequest.ResponseText, vbCrLf, "\n")
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
    Screen.MousePointer = vbDefault
End Sub

Private Sub m_oRequest_OnReadyStateChange()
    If m_oRequest.ReadyState = ucsRdsCompleted Then
        Debug.Print Replace(m_oRequest.ResponseText, vbCrLf, "\n")
    End If
    Debug.Print "OnReadyStateChange, ReadyState=" & m_oRequest.ReadyState, Format$(TimerEx, "0.000")
End Sub

Private Sub Command3_Click()
    Dim baBuffer()  As Byte
    
    On Error GoTo EH
    Set m_oRequest = New cWinSockRequest
    m_oRequest.SetTimeouts 0, 5000, 5000, 5000, 50
    m_oRequest.Open_ "wqw-pc:5123/udp", Async:=(Check1.Value = vbChecked)
    baBuffer = ToUtf8Array(Chr$(1) & "test")
    m_oRequest.Send baBuffer
'    If Check1.Value = vbChecked Then
'        m_oRequest.WaitForResponse 5000
'    End If
'    Debug.Print m_oRequest.ResponseText
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command4_Click()
    Dim oTlsClient      As cTlsClient
    Dim baBuffer()      As Byte

    Screen.MousePointer = vbHourglass
    Debug.Print Format$(TimerEx, "0.000"), "Connect secure socket to port 465"
    Set oTlsClient = New cTlsClient
    oTlsClient.SetTimeouts 0, 5000, 5000, 5000
    If Not oTlsClient.Connect("smtp.gmail.com", 465, UseTls:=True) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "TLS handshake complete: " & oTlsClient.TlsHostAddress
    If Not oTlsClient.ReadArray(baBuffer) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", FromUtf8Array(baBuffer);
    Debug.Assert Left$(FromUtf8Array(baBuffer), 3) = "220"
    Debug.Print Format$(TimerEx, "0.000"), "<-", "QUIT"
    If Not oTlsClient.WriteArray(ToUtf8Array("QUIT" & vbCrLf)) Then
        GoTo QH
    End If
    If Not oTlsClient.ReadArray(baBuffer) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", FromUtf8Array(baBuffer);
    Screen.MousePointer = vbDefault
    Exit Sub
QH:
    With oTlsClient.LastError
        Debug.Print Hex$(.Number) & ": " & .Description & " at " & .Source
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command5_Click()
    Dim oTlsClient      As cTlsClient
    Dim sHeaders        As String
    Dim sResponse       As String
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim sUrl            As String
    Dim sProxy          As String

    Screen.MousePointer = vbHourglass
    sUrl = "https://www.google.com"
    If chkProxy.Value = vbChecked Then
        sProxy = "socks5://" & txtProxy.Text
    End If
    Debug.Print Format$(TimerEx, "0.000"), "Open " & sUrl
Repeat:
    Set oTlsClient = pvInitHttpRequest(sUrl, sProxy)
    If oTlsClient Is Nothing Then
        GoTo QH
    End If
    sHeaders = vbNullString
    Do
        sResponse = oTlsClient.ReadText()
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
    oTlsClient.Close_
    Debug.Print Format$(TimerEx, "0.000"), "Done"
    Screen.MousePointer = vbDefault
    Exit Sub
QH:
    If Not oTlsClient Is Nothing Then
        With oTlsClient.LastError
            Debug.Print Hex$(.Number) & ": " & .Description & " at " & .Source
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Function pvInitHttpRequest(sUrl As String, Optional sProxyUrl As String) As cTlsClient
    Dim oRetVal         As cTlsClient
    Dim sProto          As String
    Dim sHost           As String
    Dim lPort           As Long
    Dim sPath           As String
    Dim sProxyHost      As String
    Dim lProxyPort      As Long
    Dim sProxyUser      As String
    Dim sProxyPass      As String
    Dim baBuffer()      As Byte
    
    If Not pvParseUrl(sUrl, sProto, sHost, lPort, sPath) Then
        GoTo QH
    End If
    Set oRetVal = New cTlsClient
    oRetVal.SetTimeouts 0, 5000, 5000, 5000
    If Not pvParseUrl(sProxyUrl, vbNullString, sProxyHost, lProxyPort, vbNullString, sProxyUser, sProxyPass) Then
        If Not oRetVal.Connect(sHost, lPort) Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "Connected to " & sHost & ":" & lPort
    Else
        If Not oRetVal.Connect(sProxyHost, lProxyPort) Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "Tunnel to " & sProxyHost & ":" & lProxyPort
        If LenB(sProxyUser) <> 0 Then
            If Not oRetVal.WriteArray(pvToByteArray(5, 2, 0, 2)) Then
                GoTo QH
            End If
        Else
            If Not oRetVal.WriteArray(pvToByteArray(5, 1, 0)) Then
                GoTo QH
            End If
        End If
        If Not oRetVal.ReadArray(baBuffer) Then
            GoTo QH
        End If
        If UBound(baBuffer) < 1 Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "Proxy auth method chosen: " & baBuffer(1)
        If baBuffer(1) = 2 Then
            oRetVal.WriteArray pvToByteArray(1)
            baBuffer = oRetVal.Socket.ToTextArray(sProxyUser, ucsScpUtf8)
            oRetVal.WriteArray pvToByteArray(UBound(baBuffer) + 1)
            oRetVal.WriteArray baBuffer
            baBuffer = oRetVal.Socket.ToTextArray(sProxyPass, ucsScpUtf8)
            oRetVal.WriteArray pvToByteArray(UBound(baBuffer) + 1)
            oRetVal.WriteArray baBuffer
            If Not oRetVal.ReadArray(baBuffer) Then
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
        oRetVal.WriteArray pvToByteArray(5, 1, 0, 3) '--- 5 = version, 1 = TCP stream conn, 0 = reserved, 3 = domain name
        baBuffer = oRetVal.Socket.ToTextArray(sHost, ucsScpUtf8)
        oRetVal.WriteArray pvToByteArray(UBound(baBuffer) + 1)
        oRetVal.WriteArray baBuffer
        oRetVal.WriteArray pvToByteArray(lPort \ &H100, lPort And &HFF)
        If Not oRetVal.ReadArray(baBuffer) Then
            GoTo QH
        End If
        If UBound(baBuffer) < 3 Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "Proxy connection to " & sHost & ":" & lPort & " status: " & baBuffer(1)
        If baBuffer(1) <> 0 Then
            GoTo QH
        End If
        If baBuffer(3) = 1 Then
            Debug.Print Format$(TimerEx, "0.000"), "Connection info: " & baBuffer(4) & "." & baBuffer(5) & "." & baBuffer(6) & "." & baBuffer(7) & ":" & baBuffer(8) * 256& + baBuffer(9)
        End If
    End If
    If LCase$(sProto) = "https" Then
        If Not oRetVal.StartTls(sHost) Then
            GoTo QH
        End If
        Debug.Print Format$(TimerEx, "0.000"), "TLS handshake complete"
    End If
    If Not oRetVal.WriteText("GET " & sPath & " HTTP/1.0" & vbCrLf & _
            "Host: " & sHost & vbCrLf & _
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

Private Function pvParseUrl( _
            sUrl As String, _
            sProto As String, _
            sHost As String, _
            lPort As Long, _
            sPath As String, _
            Optional sUser As String, _
            Optional sPass As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "^(.*)://(?:(?:([^:]*):)?([^@]*)@)?([A-Za-z0-9\-\.]+)(:[0-9]+)?(.*)$"
        With .Execute(sUrl)
            If .Count > 0 Then
                With .Item(0).SubMatches
                    sProto = .Item(0)
                    sUser = .Item(1)
                    If LenB(sUser) = 0 Then
                        sUser = .Item(2)
                    Else
                        sPass = .Item(2)
                    End If
                    sHost = .Item(3)
                    lPort = Val(Mid$(.Item(4), 2))
                    If lPort = 0 Then
                        Select Case LCase$(sProto)
                        Case "https"
                            lPort = 443
                        Case "socks5"
                            lPort = 1080
                        Case Else
                            lPort = 80
                        End Select
                    End If
                    sPath = .Item(5)
                    If LenB(sPath) = 0 Then
                        sPath = "/"
                    End If
                End With
                pvParseUrl = True
            End If
        End With
    End With
End Function

Private Function pvToByteArray(ParamArray A() As Variant) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    ReDim baRetVal(0 To UBound(A))
    For lIdx = 0 To UBound(A)
        baRetVal(lIdx) = A(lIdx)
    Next
    pvToByteArray = baRetVal
End Function

Private Sub Command6_Click()
    Dim sUrl            As String
    Dim sProxy          As String
    Dim oTlsClient      As cTlsClient
    
    sUrl = "https://server.cryptomix.com/secure/"
    If chkProxy.Value = vbChecked Then
        sProxy = "socks5://" & txtProxy.Text
    End If
    Debug.Print Format$(TimerEx, "0.000"), "Open " & sUrl
    Set oTlsClient = pvInitHttpRequest(sUrl, sProxy)
    If oTlsClient Is Nothing Then
        GoTo QH
    End If
    Debug.Print oTlsClient.ReadText()
    Debug.Print oTlsClient.ReadText()
    Debug.Print Format$(TimerEx, "0.000"), "Done"
    Exit Sub
QH:
    If Not oTlsClient Is Nothing Then
        With oTlsClient.LastError
            Debug.Print Hex$(.Number) & ": " & .Description & " at " & .Source
        End With
    End If
End Sub

Private Sub Command7_Click()
    Dim oTlsClient      As cTlsClient
    Dim sResponse       As String
    Dim sRequest        As String

    Screen.MousePointer = vbHourglass
    Debug.Print Format$(TimerEx, "0.000"), "Connect to port 587"
    Set oTlsClient = New cTlsClient
    oTlsClient.SetTimeouts 0, 5000, 5000, 5000
    If Not oTlsClient.Connect("smtp.gmail.com", 587) Then
        GoTo QH
    End If
    sResponse = oTlsClient.ReadText()
    If LenB(sResponse) = 0 Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", sResponse;
    sRequest = "HELO " & pvGetExternalIP & vbCrLf
    If Not oTlsClient.WriteText(sRequest) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "<-", sRequest;
    sResponse = oTlsClient.ReadText()
    If LenB(sResponse) = 0 Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", sResponse;
    sRequest = "STARTTLS" & vbCrLf
    If Not oTlsClient.WriteText(sRequest) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "<-", sRequest;
    sResponse = oTlsClient.ReadText()
    If LenB(sResponse) = 0 Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", sResponse;
    If Not oTlsClient.StartTls("smtp.gmail.com") Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "TLS handshake complete: " & oTlsClient.TlsHostAddress
    sRequest = "QUIT" & vbCrLf
    If Not oTlsClient.WriteText(sRequest) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "<-", sRequest;
    sResponse = oTlsClient.ReadText()
    If LenB(sResponse) = 0 Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", sResponse
    Screen.MousePointer = vbDefault
    Exit Sub
QH:
    With oTlsClient.LastError
        Debug.Print Hex$(.Number) & ": " & .Description & " at " & .Source
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Function pvGetExternalIP() As String
    Dim sResponse     As String
    
    With New cTlsClient
        .Connect "ifconfig.co", 80
        .WriteText "GET /ip HTTP/1.1" & vbCrLf & "Host: ifconfig.co" & vbCrLf & vbCrLf
        Do
            sResponse = sResponse & .ReadText()
            If InStr(sResponse, vbCrLf & vbCrLf) > 0 Then
                sResponse = At(Split(At(Split(sResponse, vbCrLf & vbCrLf), 1), vbLf), 0)
                If sResponse Like "*.*.*.*" Then
                    Exit Do
                End If
            End If
            If .LastError.Number <> 0 Then
                .Socket.GetSockName sResponse, 0
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
        sProxy = "socks5://" & txtProxy.Text
    End If
    With pvInitHttpRequest(sUrl, sProxy)
        DoEvents: DoEvents: DoEvents
        sResponse = sResponse & .ReadText
    End With
    Debug.Print sResponse
End Sub

Private Sub Command9_Click()
    Dim oTlsClient      As cTlsClient
    Dim baBuffer()      As Byte
    
    Screen.MousePointer = vbHourglass
    Debug.Print Format$(TimerEx, "0.000"), "Connect secure socket to port 443"
    Set oTlsClient = New cTlsClient
    oTlsClient.SetTimeouts 0, 5000, 5000, 5000
    If Not oTlsClient.Connect("connect-bot.classic.blizzard.com", 443, UseTls:=True) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "TLS handshake complete: " & oTlsClient.TlsHostAddress
    If Not oTlsClient.WriteText("GET /v1/rpc/chat HTTP/1.1" & vbCrLf & _
                "Host: connect-bot.classic.blizzard.com" & vbCrLf & _
                "Upgrade: websocket" & vbCrLf & _
                "Connection: Upgrade" & vbCrLf & _
                "Sec-WebSocket-Key: x3JJHMbDL1EzLkh9GBhXDw==" & vbCrLf & _
                "Sec-WebSocket-Protocol: chat, superchat" & vbCrLf & _
                "Sec-WebSocket-Version: 13" & vbCrLf & _
                "Origin: http://connect-bot.classic.blizzard.com/v1/rpc/chat" & vbCrLf & vbCrLf) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "->", "(HTTP request)"
    If Not oTlsClient.ReadArray(baBuffer) Then
        GoTo QH
    End If
    Debug.Print Format$(TimerEx, "0.000"), "<-", FromUtf8Array(baBuffer)
QH:
    Screen.MousePointer = vbDefault
End Sub


VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7008
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   7008
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cobVerb 
      Height          =   288
      Left            =   168
      TabIndex        =   3
      Text            =   "GET"
      Top             =   168
      Width           =   852
   End
   Begin MSWinsockLib.Winsock wscSocket 
      Left            =   6888
      Top             =   168
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3708
      Left            =   84
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   672
      Width           =   6480
   End
   Begin VB.ComboBox cobAddress 
      Height          =   288
      Left            =   1092
      TabIndex        =   1
      Text            =   "www.howsmyssl.com/a/check"
      Top             =   168
      Width           =   3792
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   348
      Left            =   4956
      TabIndex        =   0
      Top             =   168
      Width           =   1524
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private m_uCtx              As UcsTlsContext
Private m_sRequest          As String

Private Sub Connect(ByVal sServer As String, ByVal lPort As Long)
    Call TlsInitClient(m_uCtx, sServer)
    wscSocket.Close
    wscSocket.Connect sServer, lPort
End Sub

Private Sub SendData(baData() As Byte)
    Dim baOutput()          As Byte
    Dim lOutputPos          As Long
    
    If Not TlsSend(m_uCtx, baData, UBound(baData) + 1, baOutput, lOutputPos) Then
        OnError TlsGetLastError(m_uCtx), "TlsSend"
    End If
    If lOutputPos > 0 Then
        wscSocket.SendData baOutput
    End If
End Sub

Private Sub wscSocket_Connect()
    Dim baEmpty()           As Byte
    Dim baOutput()          As Byte
    Dim lOutputPos          As Long
    
    On Error GoTo EH
    If Not TlsHandshake(m_uCtx, baEmpty, -1, baOutput, lOutputPos) Then
        OnError TlsGetLastError(m_uCtx), "TlsHandshake"
    End If
    If lOutputPos > 0 Then
        wscSocket.SendData baOutput
    End If
    Exit Sub
EH:
    OnError Err.Description, "wscSocket_Connect"
End Sub

Private Sub wscSocket_DataArrival(ByVal bytesTotal As Long)
    Dim bError              As Boolean
    Dim baEmpty()           As Byte
    Dim baRecv()            As Byte
    Dim baOutput()          As Byte
    Dim lOutputPos          As Long
    Dim baPlainText()       As Byte
    Dim lSize               As Long
    
    On Error GoTo EH
    baEmpty = vbNullString
    Do While Not TlsIsClosed(m_uCtx)
        wscSocket.GetData baRecv
        If UBound(baRecv) < 0 Then
            Exit Do
        End If
        lOutputPos = 0
        If Not TlsIsReady(m_uCtx) Then
            bError = Not TlsHandshake(m_uCtx, baRecv, -1, baOutput, lOutputPos)
            If lOutputPos > 0 Then
                wscSocket.SendData baOutput
            End If
            If bError Then
                OnError TlsGetLastError(m_uCtx), "TlsHandshake"
            End If
            If TlsIsReady(m_uCtx) Then
                OnConnect
            End If
        Else
            bError = Not TlsReceive(m_uCtx, baRecv, -1, baPlainText, lSize, baOutput, lOutputPos)
            If lOutputPos > 0 Then
                wscSocket.SendData baOutput
            End If
            If bError Then
                OnError TlsGetLastError(m_uCtx), "TlsReceive"
            End If
            If lSize > 0 Then
                OnDataArrival lSize, baPlainText
            End If
            If TlsIsClosed(m_uCtx) Then
                OnClose
            End If
        End If
    Loop
    Exit Sub
EH:
    OnError Err.Description, "wscSocket_DataArrival"
End Sub

Private Sub wscSocket_Close()
    If Not TlsIsClosed(m_uCtx) Then
        OnClose
    End If
End Sub

'= callbacks =============================================================

Private Sub OnConnect()
    SendData StrConv(m_sRequest, vbFromUnicode)
End Sub

Private Sub OnDataArrival(ByVal bytesTotal As Long, baData() As Byte)
    Debug.Print "OnDataArrival, bytesTotal=" & bytesTotal, Timer
    pvAppendLogText txtResult, StrConv(baData, vbUnicode)
End Sub

Private Sub OnClose()
    Debug.Print "OnClose", Timer
End Sub

Private Sub OnError(sDescription As String, sSource As String)
    Debug.Print "Critical error: " & sDescription & " in " & sSource, Timer
    pvAppendLogText txtResult, "Critical error: " & sDescription & " in " & sSource & vbCrLf & vbCrLf
End Sub

'= form events ===========================================================

Private Sub cmdConnect_Click()
    Dim vAddr           As Variant
    Dim vPort           As Variant
    Dim sPath           As String
    
    On Error GoTo EH
    txtResult.Text = vbNullString
    '--- parse address and connect
    vAddr = Split(cobAddress.Text, "://", 2)
    If UBound(vAddr) > 0 Then
        vAddr = Split(vAddr(1), ":", 2)
    Else
        vAddr = Split(vAddr(0), ":", 2)
    End If
    If UBound(vAddr) = 0 Then
        vAddr = Split(vAddr(0), "/", 2)
        vPort = Array(443)
        If UBound(vAddr) > 0 Then
            sPath = vAddr(1)
        End If
    Else
        vPort = Split(vAddr(1), "/", 2)
        If UBound(vPort) > 0 Then
            sPath = vPort(1)
        End If
    End If
    Connect vAddr(0), vPort(0)
    '--- construct initial http request based on host and path
    If cobVerb.Text = "POST" Then
        m_sRequest = String(20000, "a")
    End If
    m_sRequest = cobVerb.Text & " /" & sPath & " HTTP/1.1" & vbCrLf & _
               "Connection: keep-alive" & vbCrLf & _
               "Content-Length: " & Len(m_sRequest) & vbCrLf & _
               "Host: " & vAddr(0) & vbCrLf & vbCrLf & _
               m_sRequest
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "cmdConnect_Click"
End Sub

Private Sub Form_Load()
    Dim vElem               As Variant
    
    For Each vElem In Split("GET|POST", "|")
        cobVerb.AddItem vElem
    Next
    For Each vElem In Split("websocket.org|www.howsmyssl.com/a/check|cert-test.sandbox.google.com|tls13.1d.pw|localhost:44330|tls.ctf.network|www.mikestoolbox.org|swifttls.org|tls13.pinterjann.is|rsa8192.badssl.com|rsa4096.badssl.com|rsa2048.badssl.com|ecc384.badssl.com|ecc256.badssl.com|dir.bg|host.bg|bgdev.org|cnn.com|gmail.com|google.com|saas.bg|saas.bg:465|www.cloudflare.com|devblogs.microsoft.com|www.brentozar.com|ayende.com/blog|www.nerds2nerds.com|robert.ocallahan.org|distrowatch.com|server.cryptomix.com/secure/|www.integralblue.com/testhandshake/", "|")
        cobAddress.AddItem vElem
    Next
    With New Form2
        .Show , Me
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        cmdConnect.Left = ScaleWidth - cmdConnect.Width - txtResult.Left
        cobAddress.Width = cmdConnect.Left - cobAddress.Left - txtResult.Left
        txtResult.Width = ScaleWidth - 2 * txtResult.Left
        txtResult.Height = ScaleHeight - txtResult.Top - txtResult.Left
    End If
End Sub

'= helpers ===============================================================

Private Sub pvAppendLogText(txtLog As TextBox, sValue As String)
    Const WM_SETREDRAW              As Long = &HB
    Const EM_SETSEL                 As Long = &HB1
    Const EM_REPLACESEL             As Long = &HC2
    Const WM_VSCROLL                As Long = &H115
    Const SB_BOTTOM                 As Long = 7
    Call SendMessage(txtLog.hWnd, WM_SETREDRAW, 0, ByVal 0)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, 0, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, -1, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_REPLACESEL, 1, ByVal sValue)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, 0, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, -1, ByVal -1)
    Call SendMessage(txtLog.hWnd, WM_SETREDRAW, 1, ByVal 0)
    Call SendMessage(txtLog.hWnd, WM_VSCROLL, SB_BOTTOM, ByVal 0)
End Sub

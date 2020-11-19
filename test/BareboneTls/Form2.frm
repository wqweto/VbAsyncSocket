VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wscSocket 
      Index           =   0
      Left            =   2772
      Top             =   252
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Label labInfo 
      Alignment       =   2  'Center
      Height          =   432
      Left            =   84
      TabIndex        =   0
      Top             =   504
      Width           =   3456
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cCerts            As Collection
Private m_cPrivKey          As Collection
Private m_uCtx()            As UcsTlsContext

Private Sub Listen(sAddr As String, ByVal lPort As Long)
    If Not PkiGenerSelfSignedCertificate(m_cCerts, m_cPrivKey) Then
        Exit Sub
    End If
    ReDim m_uCtx(0 To 0)
    wscSocket(0).Bind lPort, sAddr
    wscSocket(0).Listen
End Sub

Private Sub SendData(Index As Integer, baData() As Byte)
    Dim baOutput()          As Byte
    Dim lOutputPos          As Long
    
    If Not TlsSend(m_uCtx(Index), baData, UBound(baData) + 1, baOutput, lOutputPos) Then
        OnError Index, TlsGetLastError(m_uCtx(Index)), "TlsSend"
    End If
    If lOutputPos > 0 Then
        wscSocket(Index).SendData baOutput
    End If
End Sub

Private Sub wscSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Load wscSocket(wscSocket.UBound + 1)
    wscSocket(wscSocket.UBound).Accept requestID
    ReDim Preserve m_uCtx(0 To wscSocket.UBound)
    Call TlsInitServer(m_uCtx(wscSocket.UBound), wscSocket(wscSocket.UBound).RemoteHostIP, m_cCerts, m_cPrivKey)
End Sub

Private Sub wscSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim bError              As Boolean
    Dim baEmpty()           As Byte
    Dim baRecv()            As Byte
    Dim baOutput()          As Byte
    Dim lOutputPos          As Long
    Dim baPlainText()       As Byte
    Dim lSize               As Long
    
    On Error GoTo EH
    baEmpty = vbNullString
    Do While Not TlsIsClosed(m_uCtx(Index))
        wscSocket(Index).GetData baRecv
        If UBound(baRecv) < 0 Then
            Exit Do
        End If
        lOutputPos = 0
        If Not TlsIsReady(m_uCtx(Index)) Then
            bError = Not TlsHandshake(m_uCtx(Index), baRecv, -1, baOutput, lOutputPos)
            If lOutputPos > 0 Then
                wscSocket(Index).SendData baOutput
            End If
            If bError Then
                OnError Index, TlsGetLastError(m_uCtx(Index)), "TlsHandshake"
            End If
            If TlsIsReady(m_uCtx(Index)) Then
                OnConnect Index
            End If
        Else
            bError = Not TlsReceive(m_uCtx(Index), baRecv, -1, baPlainText, lSize, baOutput, lOutputPos)
            If lOutputPos > 0 Then
                wscSocket(Index).SendData baOutput
            End If
            If bError Then
                OnError Index, TlsGetLastError(m_uCtx(Index)), "TlsReceive"
            End If
            If lSize > 0 Then
                OnDataArrival Index, lSize, baPlainText
            End If
            If TlsIsClosed(m_uCtx(Index)) Then
                OnClose Index
            End If
        End If
    Loop
    Exit Sub
EH:
    OnError Index, Err.Description, "wscSocket_DataArrival"
End Sub

Private Sub wscSocket_Close(Index As Integer)
    If Not TlsIsClosed(m_uCtx(Index)) Then
        OnClose Index
    End If
End Sub

'= callbacks =============================================================

Public Sub OnConnect(Index As Integer)
    labInfo.Caption = "Connection " & Index & " from " & wscSocket(Index).RemoteHostIP & " port " & wscSocket(Index).RemotePort
End Sub

Private Sub OnDataArrival(Index As Integer, ByVal bytesTotal As Long, baData() As Byte)
    Dim sResponse           As String
    
    Debug.Print "OnDataArrival, Index=" & Index & ", bytesTotal=" & bytesTotal, Timer
    Debug.Print StrConv(baData, vbUnicode)
    sResponse = "<html><body>" & Now & "</body></html>"
    sResponse = "HTTP/1.0 200 Ok" & vbCrLf & _
        "Content-Type: text/html; charset=UTF-8" & vbCrLf & _
        "Content-Length: " & Len(sResponse) & vbCrLf & _
        "Connection: Close" & vbCrLf & vbCrLf & _
        sResponse
    SendData Index, StrConv(sResponse, vbFromUnicode)
End Sub

Public Sub OnClose(Index As Integer)
    Debug.Print "OnClose, Index=" & Index, Timer
    Unload wscSocket(Index)
End Sub

Public Sub OnError(Index As Integer, sDescription As String, sSource As String)
    Debug.Print "Critical error(" & Index & "): " & sDescription & " in " & sSource, Timer
End Sub

'= form events ===========================================================

Private Sub Form_Load()
    Listen "0.0.0.0", 10443
    labInfo.Caption = "Listening on " & wscSocket(0).LocalIP & " port " & wscSocket(0).LocalPort
End Sub

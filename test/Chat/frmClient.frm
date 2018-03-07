VERSION 5.00
Object = "{8405D0DF-9FDD-4829-AEAD-8E2B0A18FEA4}#1.0#0"; "Inked.dll"
Begin VB.Form frmClient 
   Caption         =   "Client"
   ClientHeight    =   4572
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5112
   LinkTopic       =   "Form3"
   ScaleHeight     =   4572
   ScaleWidth      =   5112
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   432
      Left            =   3444
      TabIndex        =   1
      Top             =   3864
      Width           =   768
   End
   Begin VB.TextBox txtMsg 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   0
      TabIndex        =   0
      Top             =   3864
      Width           =   3288
   End
   Begin INKEDLibCtl.InkEdit rchChat 
      Height          =   3624
      Left            =   0
      OleObjectBlob   =   "frmClient.frx":0000
      TabIndex        =   2
      Top             =   0
      Width           =   4716
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const DARK_RED              As Long = &H80
Private Const DARK_GREEN            As Long = &H8000&
Private Const DARK_BLUE             As Long = &H800000

Private WithEvents m_oSocket    As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_sUserName             As String
Private m_sHostAddress          As String
Private m_sHostIP               As String
Private m_lHostPort             As Long
Private m_bReconnect            As Boolean

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sAddress As String, ByVal lPort As Long, ByVal eType As UcsAsyncSocketTypeEnum) As Boolean
    On Error GoTo EH
    m_sHostAddress = IIf(LenB(sAddress), sAddress, "localhost")
    m_sHostIP = m_sHostAddress
    m_lHostPort = lPort
    m_sUserName = InputBox("Enter your chat username (leave empty for anonymous)", "Connect to " & m_sHostAddress)
    If StrPtr(m_sUserName) = 0 Then '--- InputBox cancelled
        Exit Function
    End If
    Caption = IIf(LenB(m_sUserName) <> 0, m_sUserName, "Anonymous") & " - " & IIf(eType = ucsSckStream, "TCP", "UDP") & " Chat on " & m_sHostAddress
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketType:=eType) Then
        GoTo QH
    End If
    If Not m_oSocket.Connect(m_sHostAddress, m_lHostPort) Then
        GoTo QH
    End If
    Show
    '--- success
    Init = True
    Exit Function
QH:
    Err.Raise vbObjectError, , Printf("Error %1: %2", m_oSocket.LastError, m_oSocket.GetErrorDescription(m_oSocket.LastError))
EH:
    MsgBox Err.Description, vbCritical
End Function

Property Get pvPeerName() As String
    Dim sType           As String
    Dim sAddr           As String
    Dim lPort           As Long
    
    If Not m_oSocket Is Nothing Then
        sType = IIf(m_oSocket.SockOpt(ucsSsoType) = ucsSckStream, "tcp", "udp")
        If sType = "tcp" And m_oSocket.GetPeerName(sAddr, lPort) Then
            pvPeerName = m_sHostAddress & " (" & sAddr & ") port " & lPort & "/" & sType
        Else
            pvPeerName = m_sHostAddress & " (" & m_sHostIP & ") port " & m_lHostPort & "/" & sType
        End If
    End If
End Property

'=========================================================================
' Control events
'=========================================================================

Private Sub cmdSend_Click()
    Dim eType       As UcsAsyncSocketTypeEnum
    
    On Error GoTo EH
    If m_bReconnect Then
        '--- note: m_bReconnect cleared in OnConnect event
        eType = m_oSocket.SockOpt(ucsSsoType)
        m_oSocket.Close_
        If Not m_oSocket.Create(SocketType:=eType) Then
            GoTo QH
        End If
        If Not m_oSocket.Connect(m_sHostAddress, m_lHostPort) Then
            GoTo QH
        End If
        '--- note: delay-send msg -> OnConnect event
        GoTo QH
    End If
    If LenB(txtMsg.Text) <> 0 Then
        RtbAppendLine rchChat, "[me]: " & txtMsg.Text, ForeColor:=DARK_BLUE
        If Not m_oSocket.SendText(txtMsg.Text, m_sHostIP, m_lHostPort) Then
            RtbAppendLine rchChat, Printf("Error %1 while sending" & vbCrLf & "%2", m_oSocket.LastError, m_oSocket.GetErrorDescription(m_oSocket.LastError)), _
                ForeColor:=DARK_RED
            GoTo QH
        End If
        txtMsg.Text = vbNullString
    End If
    txtMsg.SetFocus
QH:
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "Send"
End Sub

Private Sub m_oSocket_OnResolve(IpAddress As String)
    m_sHostIP = IpAddress
End Sub

Private Sub m_oSocket_OnConnect()
    On Error GoTo EH
    RtbAppendLine rchChat, Printf("Connected to %1", pvPeerName)
    If Not m_oSocket.SendText(Chr$(1) & m_sUserName, m_sHostIP, m_lHostPort) Then
        GoTo QH
    End If
    If m_bReconnect Then
        '--- msg send delayed until connected
        m_bReconnect = False
        DoEvents
        cmdSend.Value = True
    End If
QH:
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "OnConnect"
End Sub

Private Sub m_oSocket_OnReceive()
    Dim sMsg            As String
    
    On Error GoTo EH
    sMsg = m_oSocket.ReceiveText()
    Do While LenB(sMsg) <> 0
        RtbAppendLine rchChat, sMsg, HorAlign:=IIf(Left$(sMsg, 1) = "[", rtfRight, rtfCenter), _
            ForeColor:=IIf(Left$(sMsg, 1) = "[", DARK_GREEN, vbButtonShadow)
        sMsg = m_oSocket.ReceiveText()
    Loop
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "OnReceive"
End Sub

Private Sub m_oSocket_OnClose() '--- raised only for TCP
    On Error GoTo EH
    RtbAppendLine rchChat, "Disconnected from " & m_sHostAddress
    m_bReconnect = True
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "OnClose"
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    On Error GoTo EH
    Select Case EventMask
    Case ucsSfdConnect
        RtbAppendLine rchChat, Printf("Error %1 connecting to %2" & vbCrLf & "%3", ErrorCode, pvPeerName, m_oSocket.GetErrorDescription(ErrorCode)), _
            ForeColor:=DARK_RED
        m_bReconnect = True
    Case Else
        RtbAppendLine rchChat, Printf("Error %1: %2", ErrorCode, m_oSocket.GetErrorDescription(ErrorCode)), _
            ForeColor:=DARK_RED
    End Select
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "OnError"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    '--- note: explicitly send disconnect event (esp. for UDP clients)
    m_oSocket.SendText Chr$(2), m_sHostIP, m_lHostPort
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "Unload"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtMsg.Move 0, ScaleHeight - txtMsg.Height, ScaleWidth - cmdSend.Width - 84
    cmdSend.Move ScaleWidth - cmdSend.Width - 42, txtMsg.Top
    rchChat.Move 0, 0, ScaleWidth, txtMsg.Top - 42
End Sub



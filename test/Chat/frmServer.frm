VERSION 5.00
Object = "{8405D0DF-9FDD-4829-AEAD-8E2B0A18FEA4}#1.0#0"; "Inked.dll"
Begin VB.Form frmServer 
   Caption         =   "Server"
   ClientHeight    =   4032
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5880
   LinkTopic       =   "Form2"
   ScaleHeight     =   4032
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin INKEDLibCtl.InkEdit rchLog 
      Height          =   3456
      Left            =   0
      OleObjectBlob   =   "frmServer.frx":0000
      TabIndex        =   1
      Top             =   504
      Width           =   5388
   End
   Begin VB.Label labClientsCount 
      Caption         =   "0 clients"
      Height          =   348
      Left            =   168
      TabIndex        =   0
      Top             =   168
      Width           =   2532
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oSocket               As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private WithEvents m_oTcpSink   As cAsyncSocket
Attribute m_oTcpSink.VB_VarHelpID = -1
Private WithEvents m_oUdpSink   As cAsyncSocket
Attribute m_oUdpSink.VB_VarHelpID = -1
Private m_cClients              As Collection
Private m_lCount                As Long

'=========================================================================
' Methods
'=========================================================================

Public Function Init(ByVal lListenPort As Long, ByVal eSocketType As UcsAsyncSocketTypeEnum) As Boolean
    Dim sAddr           As String
    Dim lPort           As Long
    
    On Error GoTo EH
    Terminate
    '--- init member vars
    Set m_cClients = New Collection
    Caption = IIf(eSocketType = ucsSckStream, "TCP", "UDP") & " Chat Server"
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketPort:=lListenPort, SocketType:=eSocketType) Then
        GoTo QH
    End If
    If eSocketType = ucsSckStream Then
        Set m_oTcpSink = m_oSocket
        If Not m_oSocket.Listen() Then
            GoTo QH
        End If
    Else
        Set m_oUdpSink = m_oSocket
    End If
    If Not m_oSocket.GetSockName(sAddr, lPort) Then
        GoTo QH
    End If
    RtbAppendLine rchLog, Printf("Listening on %1:%2/%3...", sAddr, lPort, IIf(eSocketType = ucsSckStream, "tcp", "udp"))
    Show
    '--- success
    Init = True
    Exit Function
QH:
    Err.Raise vbObjectError, , Printf("Error %1: %2", m_oSocket.LastError, m_oSocket.GetErrorDescription(m_oSocket.LastError))
EH:
    MsgBox Err.Description, vbCritical, "Init"
End Function

Public Sub Terminate()
    On Error GoTo EH
    Set m_cClients = Nothing
    Set m_oSocket = Nothing
    Set m_oTcpSink = Nothing
    Set m_oUdpSink = Nothing
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical, "Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Terminate
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rchLog.Move 0, rchLog.Top, ScaleWidth, ScaleHeight - rchLog.Top
End Sub

'= common ================================================================

Private Sub pvClientSend(sMsg As String, Optional SkipID As String)
    Dim oClient         As cClientInfo
    
    On Error GoTo EH
    For Each oClient In m_cClients
        If oClient.ID <> SkipID Then
            If Not oClient.Socket Is Nothing Then
                oClient.Socket.SendText sMsg
            Else
                m_oUdpSink.SendText sMsg, oClient.Address, oClient.Port
            End If
        End If
    Next
    Exit Sub
EH:
    RtbAppendLine rchLog, "Critical: " & Err.Description & " (ClientSend)"
End Sub

Private Sub pvClientReceive(baBuffer() As Byte, oClient As cClientInfo, sKey As String)
    Dim sMsg            As String
    Dim sUserName       As String
    
    On Error GoTo EH
    sMsg = m_oSocket.FromTextArray(baBuffer)
    If baBuffer(0) = 1 Then
        If Len(sMsg) > 1 Then
            sUserName = Mid$(sMsg, 2)
        Else
            sUserName = "Anonymous " & oClient.ID
        End If
        If oClient.UserName <> sUserName Then
            oClient.UserName = sUserName
            RtbAppendLine rchLog, Printf("Client %1 set name to %2", oClient.ID, oClient.UserName)
        End If
        pvClientSend Printf("%1 connected from %2", sUserName, oClient.Address & ":" & oClient.Port)
    ElseIf baBuffer(0) = 2 Then
        frTcpClose sKey
    Else
        pvClientSend Printf("[%1]: %2", oClient.UserName, sMsg), oClient.ID
    End If
    Exit Sub
EH:
    RtbAppendLine rchLog, "Critical: " & Err.Description & " (ClientReceive)"
End Sub

'= TCP clients ===========================================================

Private Sub m_oTcpSink_OnAccept()
    Dim oSocket         As cAsyncSocket
    Dim oClient         As cClientInfo
    Dim sKey            As String
    
    On Error GoTo EH
    Set oSocket = New cAsyncSocket
    m_oSocket.Accept oSocket
    m_lCount = m_lCount + 1
    sKey = "#" & m_lCount
    Set oClient = New cClientInfo
    oClient.Init sKey, oSocket, Me
    m_cClients.Add oClient, sKey
    RtbAppendLine rchLog, Printf("Client %1 connected from %2", oClient.ID, oClient.Address & ":" & oClient.Port)
    labClientsCount.Caption = Printf("%1 clients", m_cClients.Count)
    Exit Sub
EH:
    RtbAppendLine rchLog, "Critical: " & Err.Description & " (OnAccept)"
End Sub

Friend Sub frTcpReceive(sKey As String, oSocket As cAsyncSocket)
    Dim baBuffer()      As Byte
    Dim oClient         As cClientInfo
    
    On Error GoTo EH
    If SearchCollection(m_cClients, sKey, RetVal:=oClient) Then
        Do While oSocket.ReceiveArray(baBuffer)
            If UBound(baBuffer) < 0 Then
                Exit Do
            End If
            pvClientReceive baBuffer, oClient, sKey
        Loop
    End If
    Exit Sub
EH:
    RtbAppendLine rchLog, "Critical: " & Err.Description & " (TcpReceive)"
End Sub

Friend Sub frTcpClose(sKey As String)
    Dim oClient         As cClientInfo
    
    On Error GoTo EH
    If SearchCollection(m_cClients, sKey, RetVal:=oClient) Then
        m_cClients.Remove sKey
        pvClientSend Printf("%1 disconnected", oClient.UserName)
        RtbAppendLine rchLog, Printf("Client %1 disconnected", oClient.ID)
        labClientsCount.Caption = Printf("%1 clients", m_cClients.Count)
    End If
    Exit Sub
EH:
    RtbAppendLine rchLog, "Critical: " & Err.Description & " (TcpClose)"
End Sub

Friend Sub frTcpError(sKey As String, oSocket As cAsyncSocket, ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Dim oClient         As cClientInfo
    
    On Error GoTo EH
    If SearchCollection(m_cClients, sKey, RetVal:=oClient) Then
        RtbAppendLine rchLog, Printf("Client %1 error %2 during %3" & vbCrLf & "%4", oClient.ID, ErrorCode, EventMask, _
            oSocket.GetErrorDescription(ErrorCode))
    End If
    Exit Sub
EH:
    RtbAppendLine rchLog, "Critical: " & Err.Description & " (TcpError)"
End Sub

'= UDP clients ===========================================================

Private Sub m_oUdpSink_OnReceive()
    Dim baBuffer()      As Byte
    Dim lPort           As Long
    Dim sAddress        As String
    Dim oClient         As cClientInfo
    Dim sKey            As String
    
    On Error GoTo EH
    If m_oSocket.ReceiveArray(baBuffer, sAddress, lPort) Then
        sKey = sAddress & ":" & lPort
        If Not SearchCollection(m_cClients, sKey, RetVal:=oClient) Then
            Set oClient = New cClientInfo
            m_lCount = m_lCount + 1
            oClient.ID = "#" & m_lCount
            oClient.UserName = "Anonymous " & oClient.ID
            oClient.Address = sAddress
            oClient.Port = lPort
            m_cClients.Add oClient, sKey
        End If
        pvClientReceive baBuffer, oClient, sKey
    End If
    Exit Sub
EH:
    RtbAppendLine rchLog, "Critical: " & Err.Description & " (OnReceive)"
End Sub


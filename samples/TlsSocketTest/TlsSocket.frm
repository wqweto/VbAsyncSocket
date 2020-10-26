VERSION 5.00
Begin VB.Form TlsSocket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test SvTLS"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Listen"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "TlsSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Serversck    As cTlsSocket
Attribute Serversck.VB_VarHelpID = -1
Private LastID                  As Long
Private Requests                As Collection

Private Sub Form_Load()
    Set Requests = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '--- cleanup circular references
    Set Requests = Nothing
End Sub

Private Sub Command1_Click()
    Set Serversck = New cTlsSocket
    If Not Serversck.InitServerTls() Then
        GoTo QH
    End If
    If Not Serversck.Create(SocketPort:=5880, SocketAddress:="localhost") Then
        GoTo QH
    End If
    If Not Serversck.Listen() Then
        GoTo QH
    End If
QH:
End Sub

Private Sub Serversck_OnAccept()
    Dim oSocket As cTlsSocket
    Dim oClient As cClientRequest
    
    If Not Serversck.Accept(oSocket, UseTls:=True) Then
        GoTo QH
    End If
    Debug.Print "New User"
    Set oClient = New cClientRequest
    LastID = LastID + 1
    oClient.ID = LastID
    Set oClient.Socket = oSocket
    Set oClient.Callback = Me
    Requests.Add oClient, "#" & oClient.ID
QH:
End Sub

Private Sub Serversck_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Debug.Print "Critical error: " & Serversck.LastError.Description & " [Serversck_OnError]", Timer
End Sub

Public Sub ClientOnReceive(Client As cClientRequest)
    Dim Svdata() As Byte
    If Not Client.Socket.ReceiveArray(Svdata) Then
        GoTo QH
    End If
    Debug.Print Client.ID, StrConv(Svdata, vbUnicode)
QH:
End Sub

Public Sub ClientOnClose(Client As cClientRequest)
    Requests.Remove "#" & Client.ID
    Debug.Print Client.ID, "Disconnected"
End Sub

Public Sub ClientOnError(Client As cClientRequest, ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    With Client.Socket.LastError
        Debug.Print Client.ID, "Critical error: " & .Description & " in " & .Source & " [ClientOnError]", Timer
    End With
End Sub

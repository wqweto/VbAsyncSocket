VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2952
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2952
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "HTTPS Server"
      Height          =   516
      Left            =   252
      TabIndex        =   3
      Top             =   2268
      Width           =   1524
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HTTPS request"
      Height          =   516
      Left            =   252
      TabIndex        =   2
      Top             =   1596
      Width           =   1524
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HTTP Server"
      Height          =   516
      Left            =   252
      TabIndex        =   1
      Top             =   924
      Width           =   1524
   End
   Begin WinsockTest.ctxWinsock ctxServer 
      Index           =   0
      Left            =   2604
      Top             =   840
      _ExtentX        =   677
      _ExtentY        =   677
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HTTP request"
      Height          =   516
      Left            =   252
      TabIndex        =   0
      Top             =   252
      Width           =   1524
   End
   Begin WinsockTest.ctxWinsock ctxWinsock 
      Left            =   2604
      Top             =   252
      _ExtentX        =   677
      _ExtentY        =   677
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ctxWinsock.Protocol = UcsProtocolConstants.sckTCPProtocol
    ctxWinsock.Connect "bgdev.org", 80
End Sub

Private Sub Command3_Click()
    ctxWinsock.Protocol = UcsProtocolConstants.sckTLSProtocol
    ctxWinsock.Connect "bgdev.org", 443
End Sub

Private Sub Command2_Click()
    ctxServer(0).Close_
    ctxServer(0).Protocol = UcsProtocolConstants.sckTCPProtocol
    ctxServer(0).Bind 8088, "127.0.0.1"
    ctxServer(0).Listen
    Shell "cmd /c start http://localhost:8088/"
End Sub

Private Sub Command4_Click()
    ctxServer(0).Close_
    ctxServer(0).Protocol = UcsProtocolConstants.sckTLSProtocol
    ctxServer(0).Bind 8088, "127.0.0.1"
    ctxServer(0).Listen ' CertSubject:="68b5220077de8bbeaed8e1c2540fec6c16b418a8"
    Shell "cmd /c start https://localhost:8088/"
End Sub

Private Sub ctxWinsock_Connect()
    Dim lIdx            As Long
    
    Debug.Print "Connected to " & ctxWinsock.RemoteHostIP, Timer
    ctxWinsock.SendData "GET / HTTP/1.0" & vbCrLf & _
        "Host: www.bgdev.org" & vbCrLf & _
        "Connection: close" & vbCrLf & vbCrLf
    For lIdx = 1 To 5000
        ctxWinsock.SendData String(1000, "a")
    Next
End Sub

Private Sub ctxWinsock_DataArrival(ByVal bytesTotal As Long)
    Dim sBuffer         As String
    
    Debug.Print "DataArrival", bytesTotal
    ctxWinsock.PeekData sBuffer
    ctxWinsock.GetData sBuffer
    Debug.Print sBuffer;
End Sub

Private Sub ctxWinsock_Error(ByVal Number As Long, Description As String, ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description & " &H" & Hex$(Number) & " [" & Source & "]", vbCritical, "ctxWinsock_Error"
End Sub

Private Sub ctxServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Debug.Print "ctxServer(" & Index & ")_ConnectionRequest, requestID=" & requestID & ", RemoteHostIP=" & ctxServer(Index).RemoteHostIP & ", RemotePort=" & ctxServer(Index).RemotePort, Timer
    Load ctxServer(ctxServer.UBound + 1)
    ctxServer(ctxServer.UBound).Accept requestID
'    Debug.Print "ctxServer(" & ctxServer.UBound & ").Protocol=" & ctxServer(ctxServer.UBound).Protocol
End Sub

Private Sub ctxServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim sRequest            As String
    Dim vSplit              As Variant
    Dim sBody               As String
    
    Debug.Print "ctxServer(" & Index & ")_DataArrival, bytesTotal=" & bytesTotal, Timer
    ctxServer(Index).GetData sRequest
    vSplit = Split(sRequest, vbCrLf)
    If UBound(vSplit) >= 0 Then
        Debug.Print vSplit(0)
        sBody = "<html><body><p>" & Join(vSplit, "</p>" & vbCrLf & "<p>" & Index & ": ") & "</p>" & vbCrLf & _
            "<p>" & Index & ": Current time is " & Now & "</p>" & _
            "<p>" & Index & ": RemoteHostIP is " & ctxServer(Index).RemoteHostIP & "</p>" & vbCrLf & _
            "<p>" & Index & ": RemotePort is " & ctxServer(Index).RemotePort & "</p>" & vbCrLf & _
            "</body></html>" & vbCrLf
        ctxServer(Index).SendData "HTTP/1.1 200 OK" & vbCrLf & _
            "Content-Type: text/html" & vbCrLf & _
            "Content-Length: " & Len(sBody) & vbCrLf & vbCrLf & _
            sBody
    End If
    Debug.Print "ctxServer(" & Index & ")_DataArrival, done", Timer
End Sub

Private Sub ctxServer_CloseEvent(Index As Integer)
    Unload ctxServer(Index)
End Sub

Private Sub ctxServer_Close(Index As Integer)
    ctxServer_CloseEvent Index
End Sub

Private Sub ctxServer_OnServerCertificate(Index As Integer, Socket As Object, Certificates As Object, PrivateKey As Object, Confirmed As Boolean)
    Debug.Print "ctxServer(" & Index & ")_OnServerCertificate, SniRequested=" & Socket.SniRequested
End Sub

Private Sub ctxServer_Error(Index As Integer, ByVal Number As Long, Description As String, ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description & " &H" & Hex$(Number) & " [" & Source & "]", vbCritical, "ctxServer(" & Index & ")_Error"
End Sub


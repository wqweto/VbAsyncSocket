VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "HTTP Server"
      Height          =   600
      Left            =   252
      TabIndex        =   1
      Top             =   924
      Width           =   1524
   End
   Begin Project1.ctxWinsock ctxServer 
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
   Begin Project1.ctxWinsock ctxWinsock 
      Left            =   2604
      Top             =   252
      _ExtentX        =   677
      _ExtentY        =   677
   End
   Begin MSWinsockLib.Winsock ctxWinsock1 
      Left            =   2100
      Top             =   252
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ctxServer1 
      Index           =   0
      Left            =   2100
      Top             =   840
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ctxWinsock.Connect "bgdev.org", 80
End Sub

Private Sub ctxWinsock_Connect()
    Debug.Print "Connected to " & ctxWinsock.RemoteHostIP, Timer
    ctxWinsock.SendData "GET / HTTP/1.0" & vbCrLf & _
        "Host: www.bgdev.org" & vbCrLf & _
        "Connection: close" & vbCrLf & vbCrLf
End Sub

Private Sub ctxWinsock_DataArrival(ByVal bytesTotal As Long)
    Dim sBuffer         As String
    
    Debug.Print "DataArrival", bytesTotal
    ctxWinsock.PeekData sBuffer
    Do
        ctxWinsock.GetData sBuffer, maxLen:=10
        Debug.Print sBuffer;
    Loop While LenB(sBuffer) <> 0
End Sub

Private Sub Command2_Click()
    ctxServer(0).Bind 8088, "127.0.0.1"
    ctxServer(0).Listen
    Shell "cmd /c start http://localhost:8088/"
End Sub

Private Sub ctxServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Debug.Print "ctxServer_ConnectionRequest, requestID=" & requestID & ", RemoteHostIP=" & ctxServer(Index).RemoteHostIP & ", RemotePort=" & ctxServer(Index).RemotePort, Timer
    Load ctxServer(ctxServer.UBound + 1)
    ctxServer(ctxServer.UBound).Accept requestID
End Sub

Private Sub ctxServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim sRequest            As String
    Dim vSplit              As Variant
    Dim sBody               As String
    
    Debug.Print "ctxServer_DataArrival, bytesTotal=" & bytesTotal, Timer
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
    Debug.Print "ctxServer_DataArrival, done", Timer
End Sub

Private Sub ctxServer_CloseEvent(Index As Integer)
    Unload ctxServer(Index)
End Sub

Private Sub ctxServer_Close(Index As Integer)
    ctxServer_CloseEvent Index
End Sub



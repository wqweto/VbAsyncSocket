VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Port Listener"
   ClientHeight    =   5292
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5292
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBody 
      Height          =   1332
      Left            =   3720
      TabIndex        =   13
      Text            =   "Nobody can beat Chuck Norris !!!"
      Top             =   360
      Width           =   3972
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   492
      Left            =   2040
      TabIndex        =   11
      Top             =   1200
      Width           =   1332
   End
   Begin VB.TextBox txtConsole 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   2892
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2280
      Width           =   3372
   End
   Begin WinsockTest_Server.ctxWinsock ctxServer 
      Index           =   0
      Left            =   2760
      Top             =   120
      _ExtentX        =   677
      _ExtentY        =   677
   End
   Begin VB.OptionButton optHTTPS 
      Caption         =   "HTTPS"
      Height          =   252
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   1092
   End
   Begin VB.OptionButton optHTTP 
      Caption         =   "HTTP"
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Value           =   -1  'True
      Width           =   852
   End
   Begin VB.TextBox txtContent 
      Height          =   2892
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2280
      Width           =   4212
   End
   Begin VB.TextBox txtPort 
      Height          =   288
      Left            =   2520
      TabIndex        =   3
      Text            =   "8088"
      Top             =   720
      Width           =   972
   End
   Begin VB.TextBox txtAddress 
      Height          =   288
      Left            =   240
      TabIndex        =   1
      Text            =   "localhost"
      Top             =   720
      Width           =   2172
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   516
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1524
   End
   Begin VB.Label Label5 
      Caption         =   "Answer on GET request"
      Height          =   252
      Left            =   3720
      TabIndex        =   12
      Top             =   120
      Width           =   1932
   End
   Begin VB.Label Label4 
      Caption         =   "Console:"
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "Incoming Data:"
      Height          =   252
      Left            =   3720
      TabIndex        =   5
      Top             =   1920
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      Height          =   252
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2172
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim g_sHost As String
Dim g_iPort As Long

Private Sub cmdListen_Click()
    
    Dim Protocol As UcsProtocolConstants
    
    If (optHTTP.Value) Then
        Protocol = sckTCPProtocol
    Else
        Protocol = sckTLSProtocol
    End If
    
    g_sHost = txtAddress.Text
    g_iPort = CLng(txtPort.Text)
    
    ctxServer(0).Close_
    ctxServer(0).Protocol = Protocol
    ctxServer(0).Bind g_iPort, g_sHost
    'ctxServer(0).Protocol = sckTCPProtocol
    'ctxServer(0).Bind 8088, "localhost"
    ctxServer(0).Listen
    
    cmdListen.Enabled = False
    cmdStop.Enabled = True
    txtAddress.Enabled = False
    txtPort.Enabled = False
    optHTTP.Enabled = False
    optHTTPS.Enabled = False
End Sub

Private Sub cmdStop_Click()
    Dim i&
    For i = 0 To ctxServer.UBound
        If Not (ctxServer(i) Is Nothing) Then
            ctxServer(i).Close_
        End If
    Next
    cmdListen.Enabled = True
    cmdStop.Enabled = False
    txtAddress.Enabled = True
    txtPort.Enabled = True
    optHTTP.Enabled = True
    optHTTPS.Enabled = True
End Sub

Private Sub ctxServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    DebugToConsole "ctxServer_ConnectionRequest, requestID=" & requestID & ", RemoteHostIP=" & ctxServer(Index).RemoteHostIP & ", RemotePort=" & ctxServer(Index).RemotePort
    Load ctxServer(ctxServer.UBound + 1)
    ctxServer(ctxServer.UBound).Protocol = ctxServer(Index).Protocol
    ctxServer(ctxServer.UBound).Accept requestID
End Sub

Private Sub ctxServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim sRequest            As String
    Dim sBody               As String
    Dim sContent            As String
    
    DebugToConsole "ctxServer_DataArrival, bytesTotal=" & bytesTotal
    ctxServer(Index).GetData sRequest
    
    txtContent.Text = txtContent.Text & sRequest & vbCrLf
    
    sBody = txtBody.Text
    sContent = "HTTP/1.1 200 OK" & vbCrLf & _
        "Content-Type: text/plain; charset=windows-1251" & vbCrLf & _
        "Content-Length: " & Len(sBody) & vbCrLf & _
        "Connection: close" & vbCrLf & vbCrLf & _
        sBody & vbCrLf & vbCrLf
    
    ctxServer(Index).SendData sContent
    ctxServer(Index).Close_
End Sub

Private Sub ctxServer_CloseEvent(Index As Integer)
    DebugToConsole "ctxServer_CloseEvent", Index
    Unload ctxServer(Index)
End Sub

Private Sub ctxServer_Close(Index As Integer)
    DebugToConsole "ctxServer_Close", Index
    ctxServer_CloseEvent Index
End Sub

Private Sub ctxServer_Error(Index As Integer, ByVal Number As Long, Description As String, ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    DebugToConsole "Error: " & Description
End Sub

Sub DebugToConsole(ParamArray pa())
    Dim s$: s = Join(pa, " ")
    txtConsole.Text = vbCrLf & txtConsole.Text & vbCrLf & s
    txtConsole.SelStart = Len(txtConsole.Text)
    Debug.Print s
End Sub

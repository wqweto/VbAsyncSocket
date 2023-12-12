VERSION 5.00
Begin VB.Form frmRemaster 
   Caption         =   "frmRemaster"
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
   Begin VB.CommandButton Command1 
      Caption         =   "HTTP request"
      Height          =   516
      Left            =   252
      TabIndex        =   0
      Top             =   252
      Width           =   1524
   End
End
Attribute VB_Name = "frmRemaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oClient    As cTlsRemaster
Attribute m_oClient.VB_VarHelpID = -1
Private WithEvents m_oServer    As cTlsRemaster
Attribute m_oServer.VB_VarHelpID = -1
Private m_cConnPool             As New Collection
Attribute m_cConnPool.VB_VarHelpID = -1

Private Sub Command1_Click()
    Set m_oClient = New cTlsRemaster
    m_oClient.Protocol = UcsProtocolConstants.sckTCPProtocol
    m_oClient.Connect "bgdev.org", 80
End Sub

Private Sub Command3_Click()
    Set m_oClient = New cTlsRemaster
    m_oClient.Protocol = UcsProtocolConstants.sckTLSProtocol
    m_oClient.Connect "bgdev.org", 443
End Sub

Private Sub Command2_Click()
    Set m_oServer = New cTlsRemaster
    m_oServer.Protocol = UcsProtocolConstants.sckTCPProtocol
    m_oServer.Bind 8088, "127.0.0.1"
    m_oServer.Listen
    Shell "cmd /c start http://localhost:8088/"
End Sub

Private Sub Command4_Click()
    Set m_oServer = New cTlsRemaster
    m_oServer.Protocol = UcsProtocolConstants.sckTLSProtocol
    m_oServer.Bind 8088, "127.0.0.1"
    m_oServer.Listen ' CertSubject:="68b5220077de8bbeaed8e1c2540fec6c16b418a8"
    Shell "cmd /c start https://localhost:8088/"
End Sub

Private Sub m_oClient_Connect()
    Dim lIdx            As Long
    
    Debug.Print "Connected to " & m_oClient.RemoteHostIP, Timer
    m_oClient.SendData "GET / HTTP/1.0" & vbCrLf & _
        "Host: www.bgdev.org" & vbCrLf & _
        "Connection: close" & vbCrLf & vbCrLf
    For lIdx = 1 To 5000
        m_oClient.SendData String(1000, "a")
    Next
End Sub

Private Sub m_oClient_DataArrival(ByVal bytesTotal As Long)
    Dim sBuffer         As String
    
    Debug.Print "DataArrival", bytesTotal
    m_oClient.PeekData sBuffer
    m_oClient.GetData sBuffer
    Debug.Print sBuffer;
End Sub

Private Sub m_oServer_ConnectionRequest(ByVal requestID As Long)
    Debug.Print "m_oServer_ConnectionRequest, requestID=" & requestID & ", RemoteHostIP=" & m_oServer.RemoteHostIP & ", RemotePort=" & m_oServer.RemotePort, Timer
    Dim oCallback           As cClientCallback
    
    Set oCallback = New cClientCallback
    m_cConnPool.Add oCallback
    oCallback.Index = m_cConnPool.Count
    Set oCallback.Parent = Me
    Set oCallback.Socket = New cTlsRemaster
    oCallback.Socket.Protocol = m_oServer.Protocol
    oCallback.Socket.Accept requestID
End Sub

Public Sub OnDataArrival(Index As Long, ByVal bytesTotal As Long)
    Dim sRequest            As String
    Dim vSplit              As Variant
    Dim sBody               As String
    Dim oCallback           As cClientCallback
    
    Set oCallback = m_cConnPool.Item(Index)
    Debug.Print "OnDataArrival, Index=" & Index & ", bytesTotal=" & bytesTotal, Timer
    oCallback.Socket.GetData sRequest
    vSplit = Split(sRequest, vbCrLf)
    If UBound(vSplit) >= 0 Then
        Debug.Print vSplit(0)
        sBody = "<html><body><p>" & Join(vSplit, "</p>" & vbCrLf & "<p>" & Index & ": ") & "</p>" & vbCrLf & _
            "<p>" & Index & ": Current time is " & Now & "</p>" & _
            "<p>" & Index & ": RemoteHostIP is " & oCallback.Socket.RemoteHostIP & "</p>" & vbCrLf & _
            "<p>" & Index & ": RemotePort is " & oCallback.Socket.RemotePort & "</p>" & vbCrLf & _
            "</body></html>" & vbCrLf
        oCallback.Socket.SendData "HTTP/1.1 200 OK" & vbCrLf & _
            "Content-Type: text/html" & vbCrLf & _
            "Content-Length: " & Len(sBody) & vbCrLf & vbCrLf & _
            sBody
    End If
    Debug.Print "OnDataArrival, Index=" & Index & ", done", Timer
End Sub
 
Public Sub OnCloseSck(Index As Long)
    m_cConnPool.Remove Index
    If m_cConnPool.Count >= Index Then
        m_cConnPool.Add Nothing, Before:=Index
    End If
    Do While m_cConnPool.Count > 0
        If Not m_cConnPool.Item(m_cConnPool.Count) Is Nothing Then
            Exit Do
        End If
        m_cConnPool.Remove m_cConnPool.Count
    Loop
End Sub

'Private Sub m_oClient_Error(ByVal Number As Long, Description As String, ByVal sCode As UcsErrorConstants, Source As String, HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    MsgBox Description & " &H" & Hex$(Number) & " [" & Source & "]", vbCritical, "m_oClient_Error"
'End Sub

'Private Sub m_oServer_Error(Index As Integer, ByVal Number As Long, Description As String, ByVal sCode As UcsErrorConstants, Source As String, HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    MsgBox Description & " &H" & Hex$(Number) & " [" & Source & "]", vbCritical, "m_oServer(" & Index & ")_Error"
'End Sub


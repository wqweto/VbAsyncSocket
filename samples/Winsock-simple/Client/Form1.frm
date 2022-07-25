VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "HTTP Data Sender"
   ClientHeight    =   3468
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7368
   LinkTopic       =   "Form1"
   ScaleHeight     =   3468
   ScaleWidth      =   7368
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbTemplate 
      Height          =   288
      Left            =   1800
      TabIndex        =   17
      Top             =   360
      Width           =   1692
   End
   Begin VB.TextBox txtValue 
      Height          =   288
      Left            =   720
      TabIndex        =   14
      Text            =   "testValue"
      Top             =   3120
      Width           =   1452
   End
   Begin VB.TextBox txtKey 
      Height          =   288
      Left            =   720
      TabIndex        =   13
      Text            =   "testKey"
      Top             =   2880
      Width           =   1452
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "GET"
      Height          =   492
      Left            =   2280
      TabIndex        =   12
      Top             =   2880
      Width           =   1212
   End
   Begin VB.CommandButton cmdRaw 
      Caption         =   "Raw"
      Height          =   492
      Left            =   480
      TabIndex        =   11
      Top             =   2280
      Width           =   1332
   End
   Begin VB.TextBox txtConsole 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2892
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   480
      Width           =   3372
   End
   Begin VB.OptionButton optHTTPS 
      Caption         =   "HTTPS"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   360
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
      Height          =   492
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   1680
      Width           =   3372
   End
   Begin VB.TextBox txtPort 
      Height          =   288
      Left            =   2520
      TabIndex        =   3
      Text            =   "8088"
      Top             =   960
      Width           =   972
   End
   Begin VB.TextBox txtAddress 
      Height          =   288
      Left            =   240
      TabIndex        =   1
      Text            =   "localhost"
      Top             =   960
      Width           =   2172
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "POST"
      Height          =   516
      Left            =   1920
      TabIndex        =   0
      Top             =   2280
      Width           =   1284
   End
   Begin WinsockTest_Client.ctxWinsock ctxWinsock 
      Left            =   6840
      Top             =   0
      _ExtentX        =   677
      _ExtentY        =   677
   End
   Begin VB.Label Label7 
      Caption         =   "Template"
      Height          =   252
      Left            =   1800
      TabIndex        =   18
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "Value:"
      Height          =   252
      Left            =   120
      TabIndex        =   16
      Top             =   3150
      Width           =   492
   End
   Begin VB.Label Label5 
      Caption         =   "Key:"
      Height          =   252
      Left            =   120
      TabIndex        =   15
      Top             =   2920
      Width           =   492
   End
   Begin VB.Label Label4 
      Caption         =   "Console:"
      Height          =   252
      Left            =   3840
      TabIndex        =   10
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label3 
      Caption         =   "Data to send:"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      Height          =   252
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   720
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

Enum QUERY_TYPE
    QUERY_RAW
    QUERY_POST
    QUERY_GET
End Enum

Dim g_Query As QUERY_TYPE


Private Sub cmdGet_Click()
    g_Query = QUERY_GET
    SendConnect
End Sub

Private Sub cmdPost_Click()
    g_Query = QUERY_POST
    SendConnect
End Sub

Private Sub cmdRaw_Click()
    g_Query = QUERY_RAW
    SendConnect
End Sub

Sub SendConnect()
    g_sHost = txtAddress.Text
    g_iPort = CLng(txtPort.Text)
    
    ctxWinsock.Protocol = IIf(optHTTP.Value, sckTCPProtocol, sckTLSProtocol)
    ctxWinsock.Connect g_sHost, g_iPort, ucsTlsIgnoreServerCertificateErrors
End Sub

Private Sub ctxWinsock_Connect()
    Dim sMessage As String
    Dim sKeyValue As String
    
    sMessage = txtContent.Text
    'key + value?
    If txtKey.Text <> "" And txtValue.Text <> "" Then
        sKeyValue = "?" & txtKey.Text & "=" & txtValue.Text
    ' only key?
    ElseIf txtKey.Text <> "" Then
        sKeyValue = txtKey.Text
    End If
    
    DebugToConsole vbCrLf & "Connected to " & ctxWinsock.RemoteHostIP
    
    Select Case g_Query
    
    Case QUERY_RAW
        ctxWinsock.SendData sMessage
        
    Case QUERY_POST
    
        ctxWinsock.SendData "POST / HTTP/1.0" & vbCrLf & _
            "Host: " & g_sHost & vbCrLf & _
            "Content-lenght: " & LenB(sMessage) & vbCrLf & vbCrLf & _
            sMessage
    
    Case QUERY_GET
    
        ctxWinsock.SendData _
            "GET /" & sKeyValue & " HTTP/1.1" & vbCrLf & _
            "Host: " & g_sHost & ":" & g_iPort & vbCrLf & _
            "Connection: keep-alive" & vbCrLf & _
            "Upgrade-Insecure-Requests: 1" & vbCrLf & _
            "User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64)" & vbCrLf & _
            "Accept: text/html,application/xhtml+xml,application/xm" & vbCrLf & vbCrLf
    End Select
End Sub

Private Sub ctxWinsock_DataArrival(ByVal bytesTotal As Long)
    Dim sBuffer         As String
    
    DebugToConsole "DataArrival", bytesTotal
    ctxWinsock.GetData sBuffer
    DebugToConsole sBuffer
End Sub

Private Sub ctxWinsock_Error(ByVal Number As Long, Description As String, ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    DebugToConsole "Error: " & Description
End Sub

Sub DebugToConsole(ParamArray pa())
    Dim s$: s = Join(pa, " ")
    txtConsole.Text = vbCrLf & txtConsole.Text & vbCrLf & s
    txtConsole.SelStart = Len(txtConsole.Text)
    Debug.Print s
End Sub

Private Sub Form_Load()
    cmbTemplate.AddItem "localhost"
    cmbTemplate.AddItem "GET"
End Sub

Private Sub cmbTemplate_Click()
    Dim sItem As String
    sItem = cmbTemplate.List(cmbTemplate.ListIndex)
    Select Case sItem
        Case "localhost":
            optHTTPS.Value = True
            txtKey.Text = "testKey"
            txtValue.Text = "testValue"
            txtAddress.Text = "localhost"
            txtPort.Text = "8088"
        Case "GET":
            optHTTP.Value = True
            txtKey.Text = "ip"
            txtValue.Text = ""
            txtAddress.Text = "ifconfig.co"
            txtPort.Text = "80"
    End Select
End Sub

Private Sub optHTTP_Click()
    txtPort.Text = 80
End Sub

Private Sub optHTTPS_Click()
    txtPort.Text = 443
End Sub

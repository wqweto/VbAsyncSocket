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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   516
      Left            =   1596
      TabIndex        =   0
      Top             =   420
      Width           =   1524
   End
   Begin Project1.ctxWinsock ctxWinsock1 
      Left            =   756
      Top             =   1260
      _ExtentX        =   677
      _ExtentY        =   677
   End
   Begin MSWinsockLib.Winsock ctxWinsock2 
      Left            =   588
      Top             =   504
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
    ctxWinsock1.Connect "bgdev.org", 80
End Sub

Private Sub ctxWinsock1_Connect()
    Debug.Print "Connected to " & ctxWinsock1.RemoteHostIP, Timer
    ctxWinsock1.SendData "GET / HTTP/1.0" & vbCrLf & _
        "Host: www.bgdev.org" & vbCrLf & _
        "Connection: close" & vbCrLf & vbCrLf
End Sub

Private Sub ctxWinsock1_DataArrival(ByVal bytesTotal As Long)
    Dim sBuffer         As String
    
    Debug.Print "DataArrival", bytesTotal
    ctxWinsock1.PeekData sBuffer
    Do
        Debug.Print sBuffer;
        ctxWinsock1.GetData sBuffer, maxLen:=10
    Loop While LenB(sBuffer) <> 0
End Sub

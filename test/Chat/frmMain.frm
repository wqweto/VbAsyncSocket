VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Sample Chat"
   ClientHeight    =   3312
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4956
   LinkTopic       =   "Form1"
   ScaleHeight     =   3312
   ScaleWidth      =   4956
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkUdp 
      Caption         =   "Use UDP"
      Height          =   192
      Left            =   588
      TabIndex        =   0
      Top             =   252
      Width           =   3204
   End
   Begin VB.TextBox txtServer 
      Height          =   348
      Left            =   504
      TabIndex        =   3
      Top             =   2352
      Width           =   3960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Client"
      Height          =   432
      Left            =   504
      TabIndex        =   2
      Top             =   1512
      Width           =   1356
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Server"
      Height          =   432
      Left            =   504
      TabIndex        =   1
      Top             =   756
      Width           =   1356
   End
   Begin VB.Label Label3 
      Caption         =   "Server IP/Name (leave empty for local)"
      Height          =   264
      Left            =   504
      TabIndex        =   6
      Top             =   2100
      Width           =   3120
   End
   Begin VB.Label Label2 
      Caption         =   "(start single instance or restart)"
      Height          =   264
      Left            =   2016
      TabIndex        =   5
      Top             =   840
      Width           =   2868
   End
   Begin VB.Label Label1 
      Caption         =   "(start multiple clients)"
      Height          =   264
      Left            =   2016
      TabIndex        =   4
      Top             =   1596
      Width           =   2868
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

Private Const DEF_CHAT_PORT     As Long = 5123

Private Sub Command3_Click()
    On Error GoTo EH
    With frmServer
        .Init DEF_CHAT_PORT, IIf(chkUdp.Value = vbChecked, ucsSckDatagram, ucsSckStream)
    End With
EH:
End Sub

Private Sub Command4_Click()
    On Error GoTo EH
    With New frmClient
        .Init txtServer.Text, DEF_CHAT_PORT, IIf(chkUdp.Value = vbChecked, ucsSckDatagram, ucsSckStream)
    End With
EH:
End Sub

Private Sub Form_Load()
    Left = Screen.Width - Width - 1000
End Sub

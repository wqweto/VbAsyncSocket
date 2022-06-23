VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3504
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3504
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "ECDSA"
      Height          =   516
      Left            =   2100
      TabIndex        =   4
      Top             =   924
      Width           =   1608
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ECDH"
      Height          =   516
      Left            =   336
      TabIndex        =   3
      Top             =   924
      Width           =   1608
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Chacha20-Poly1305"
      Height          =   516
      Left            =   3864
      TabIndex        =   2
      Top             =   252
      Width           =   1608
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AES-CBC"
      Height          =   516
      Left            =   2100
      TabIndex        =   1
      Top             =   252
      Width           =   1608
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AES-GCM"
      Height          =   516
      Left            =   336
      TabIndex        =   0
      Top             =   252
      Width           =   1608
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

Private Sub Command1_Click()
    TestCryptoAesGcm JsonParseObject(ReadTextFile(App.Path & "\wycheproof\testvectors\aes_gcm_test.json"))
End Sub

Private Sub Command2_Click()
    TestCryptoAesCbc JsonParseObject(ReadTextFile(App.Path & "\wycheproof\testvectors\aes_cbc_pkcs5_test.json"))
End Sub

Private Sub Command3_Click()
    TestCryptoChacha20 JsonParseObject(ReadTextFile(App.Path & "\wycheproof\testvectors\chacha20_poly1305_test.json"))
End Sub

Private Sub Command4_Click()
    TestCryptoEcdh JsonParseObject(ReadTextFile(App.Path & "\wycheproof\testvectors\ecdh_test.json"))
End Sub

Private Sub Command5_Click()
    TestCryptoEcdsa JsonParseObject(ReadTextFile(App.Path & "\wycheproof\testvectors\ecdsa_test.json"))
End Sub

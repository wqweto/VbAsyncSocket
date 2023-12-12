VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cClientCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Index                As Long
Public Parent               As frmRemaster
Public WithEvents Socket    As cTlsRemaster
Attribute Socket.VB_VarHelpID = -1

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Parent.OnDataArrival Index, bytesTotal
End Sub

Private Sub Socket_CloseSck()
    Parent.OnCloseSck Index
End Sub


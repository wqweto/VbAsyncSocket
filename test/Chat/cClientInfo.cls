VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cClientInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Public ID                       As String
Public UserName                 As String
Public Address                  As String
Public Port                     As Long
Public WithEvents Socket        As cAsyncSocket
Attribute Socket.VB_VarHelpID = -1
Private m_oParent               As frmServer

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sID As String, oSocket As cAsyncSocket, oParent As frmServer)
    ID = sID
    UserName = "Anonymous " & sID
    oSocket.GetPeerName Address, Port
    Set Socket = oSocket
    Set m_oParent = oParent
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub Socket_OnClose()
    m_oParent.frTcpClose ID
End Sub

Private Sub Socket_OnReceive()
    m_oParent.frTcpReceive ID, Socket
End Sub

Private Sub Socket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    m_oParent.frTcpError ID, Socket, ErrorCode, EventMask
End Sub


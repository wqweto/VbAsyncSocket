Attribute VB_Name = "Module2"
Option Explicit

Public Type UcsTlsContext
    Dummy As Long
End Type

Public Property Get TlsIsClosed(uCtx As UcsTlsContext) As Boolean
    
End Property

Public Property Get TlsIsStarted(uCtx As UcsTlsContext) As Boolean
    
End Property

Public Property Get TlsIsReady(uCtx As UcsTlsContext) As Boolean
    
End Property

Public Property Get TlsIsShutdown(uCtx As UcsTlsContext) As Boolean
    
End Property

'=========================================================================
' TLS support
'=========================================================================

Public Function TlsInitClient( _
            uCtx As UcsTlsContext, _
            Optional RemoteHostName As String, _
            Optional ByVal LocalFeatures As Long, _
            Optional ClientCertCallback As Object, _
            Optional AlpnProtocols As String) As Boolean
    
End Function

Public Function TlsInitServer( _
            uCtx As UcsTlsContext, _
            Optional RemoteHostName As String, _
            Optional Certificates As Collection, _
            Optional PrivateKey As Collection, _
            Optional AlpnProtocols As String, _
            Optional ByVal LocalFeatures As Long) As Boolean

End Function

Public Function TlsTerminate(uCtx As UcsTlsContext)

End Function

Public Function TlsHandshake(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baOutput() As Byte, lOutputPos As Long) As Boolean

End Function

Public Function TlsReceive(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baPlainText() As Byte, lPos As Long, baOutput() As Byte, lOutputPos As Long) As Boolean

End Function

Public Function TlsSend(uCtx As UcsTlsContext, baPlainText() As Byte, ByVal lSize As Long, baOutput() As Byte, lOutputPos As Long) As Boolean
    
End Function

Public Function TlsShutdown(uCtx As UcsTlsContext, baOutput() As Byte, lPos As Long) As Boolean
    
End Function

Public Function TlsGetLastError(uCtx As UcsTlsContext, Optional LastErrNumber As Long, Optional LastErrSource As String) As String
    
End Function




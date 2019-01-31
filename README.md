## VbAsyncSocket

Simple and thin WinSock API wrappers for VB6 loosly based on the original [`CAsyncSocket`](https://docs.microsoft.com/en-us/cpp/mfc/reference/casyncsocket-class?view=vs-2017) wrapper in [`MFC`](https://docs.microsoft.com/en-us/cpp/mfc/mfc-and-atl?view=vs-2017).

### Description

Base class `cAsyncSocket` wraps OS non-blocking sockets that can be used to implement various network components in VB6 -- clients and servers -- and supports both async and blocking network communications.

Additionally in `contrib` folder there are helper classes that build-up on `cAsyncSocket` base to provide familiar interfaces for completing common tasks.

For instance class `cWinSockRequest` exposes same-named methods, properties and events as the [OS provided `WinHttpRequest` object](https://msdn.microsoft.com/en-us/library/windows/desktop/aa384106%28v=vs.85%29.aspx) but can be used to easily access generic TCP/UDP servers instead.

Then class `cTlsClient` provides streaming support (familiar `Read`/`Write` methods) for accessing both plain-text TCP servers (e.g. http) and SSL/TLS secured ones (e.g. https) using [OS provided SSPI/Schannel](https://msdn.microsoft.com/en-us/library/windows/desktop/aa374782(v=vs.85).aspx) implementation w/o any additional dependency (e.g. OpenSSL library).

For simple demo TCP/UDP client/server implementation take a look at `test\Chat` folder. Same client and server sides are capable of communicating over reliable TCP channels and fast UDP datagrams. 

For sample clients to various network services (https, SMTP over SSL, etc.) open sample project in `test\Basic` folder.

### Usage

Start by including `src\cAsyncSocket.cls` with companion `src\cAsyncSocketHelperWindow.cls` in your project to have a convenient wrapper to most WinSock API functions. Note that `cAsyncSocketHelperWindow` uses self-subclassing implementation that can be unstable to the `End` button/statement in VBIDE.

Optionally you can add `contrib\cTlsClient.cls` to your project for SSL/TLS secured channels and/or `contrib\cWinSockRequest.cls` for familiar class methods when accessing TCP/UDP network services.

### Sample SMTP with STARTTLS

A working sample with error checking omitted for brevity for accessing smtp.gmail.com over port 587.

At first the communication goes over unencrypted plain-text socket, then later its switched to TLS secured one before issuing last `QUIT` command.

    With New cTlsClient
        .SetTimeouts 0, 5000, 5000, 5000
        .Connect "smtp.gmail.com", 587
        Debug.Print .ReadText();
        .WriteText "HELO 127.0.0.1" & vbCrLf
        Debug.Print .ReadText();
        .WriteText "STARTTLS" & vbCrLf
        Debug.Print .ReadText();
        .StartTls "smtp.gmail.com"
        Debug.Print "TLS handshake complete: " & .TlsHostAddress
        .WriteText "QUIT" & vbCrLf
        Debug.Print .ReadText();
    End With

Produces debug output in `Immediate Window` similar to this:
    
    220 smtp.gmail.com ESMTP l8sm10315047wmf.39 - gsmtp
    250 smtp.gmail.com at your service
    220 2.0.0 Ready to start TLS
    TLS handshake complete: smtp.gmail.com
    221 2.0.0 closing connection l8sm10315047wmf.39 - gsmtp

### ToDo

- Allow client to perform TLS server certificate check
- Allow client to assign client certificate for connection
- Provide UI for end-user to choose suitable certificates from Personal certificate store
- Add wrappers for http and ftp protocols
- Add WinSock control replacement
- Add more samples (incl. `vbcurl.exe` utility)
- Refactor subclassing thunk to use msg queue not to re-enter IDE in debug mode

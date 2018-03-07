## VbAsyncSocket

Simple and thin WinSock API wrappers for VB6

### Description

Base class `cAsyncSocket` wraps OS non-blocking sockets that can be used to implement various network components in VB6 - clients and servers -- and supports both async and blocking network communications.

Additionally in `contrib` folder there are helper classes that build-up on `cAsyncSocket` and provide familiar interfaces for common tasks. Class `cWinSockRequest` exposes same-named methods, properties and events as the OS provided `WinHttpRequest` object but can be used to easily access generic TCP/UDP servers instead, while `cTlsClient` provides streaming support for accessing both plain TCP servers and SSL/TLS secured ones.

For simple TCP/UDP server/client implementation take a look at `test\Chat` folder. For some sample clients to various network services (https, SMTP over SSL, etc.) open sample project in `test\Basic` folder.

### Usage

Just include `src\cAsyncSocket.cls` with companion `src\cAsyncSocketHelperWindow.cls` in your project to have a convenient wrapper to most WinSock API functions. Note that `cAsyncSocketHelperWindow` uses self-subclassing implementation that can be unstable to the `End` button/statement in VBIDE.

### ToDo

- Allow TLS server certificate check
- More samples
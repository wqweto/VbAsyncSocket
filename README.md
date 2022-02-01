## VbAsyncSocket

Simple and thin WinSock API wrappers for VB6 loosly based on the original [`CAsyncSocket`](https://docs.microsoft.com/en-us/cpp/mfc/reference/casyncsocket-class?view=vs-2017) wrapper in [`MFC`](https://docs.microsoft.com/en-us/cpp/mfc/mfc-and-atl?view=vs-2017).

### Description

Base class `cAsyncSocket` wraps OS non-blocking sockets that can be used to implement various network components in VB6 -- clients and servers -- and supports both async and blocking network communications.

Additionally there is a source-compatible `cTlsSocket` class for transparent TLS transport layer encryption with several crypto backend implementations:

1. `mdTlsThunks` is a pure VB6 with ASM thunks implementation for TLS 1.3 and (legacy) TLS 1.2 client-side and server-side support with no dependency on external libraries (like openssl)

2. `mdTlsNative` is a native client-side and server-side TLS support using OS provided SSPI/Schannel library for all available protocol versions.

3. `mdTlsSodium` is a stripped down compact backend with dependency on libsodium for crypto primitives (no ASM thunking used) with a total compiled size of 64KB.

The VB6 with thunks backend implementation auto-detects AES-NI and PCLMULQDQ instruction set availability on client machine and switches to [performance optimized implementation of AES](https://github.com/wqweto/VbAsyncSocket/blob/4b7f4d8bc650688e2b6ad5460c997ed1df26d2e0/lib/thunks/sshaes.c#L100-L240)[-GCM](https://github.com/wqweto/VbAsyncSocket/blob/4b7f4d8bc650688e2b6ad5460c997ed1df26d2e0/lib/thunks/gf128.c#L116-L165) which is even faster that OS native SSPI/Schannel implementation of this cipher suit. The VB6 with thunks backend and native backend support legacy OSes up to NT 4.0 while libsodium DLL is compiled with XP support only.

### Usage

Start by including `src\cAsyncSocket.cls` in your project to have a convenient wrapper of most WinSock API functions.

Optionally you can add `src\cTlsSocket.cls` and `src\mdTlsThunks.bas` pair of source files to your project for TLS secured connections using VB6 with thunks backend or add `src\cTlsSocket.cls` and `src\mdTlsNative.bas` pair of source files for an alternative backend using native OS provided SSPI/Schannel library.

#### WinHttpRequest Replacement Class

Start by including `src\cAsyncSocket.cls`, `src\cTlsSocket.cls` and `src\mdTlsThunks.bas` backend for TLS support (or any other backend) and finally add `contrib\cHttpRequest.cls` for the TLS 1.3 capable source-compatible replacement class.

Notice that the original `Open` method and `Option` property of the `WinHttpRequest` object have been suffixed with an underscore (`_`) in the replacement implementation (a limitation of the VB6 IDE) so some source-code fixes will be required to integrate the replacement `cHttpRequest` class.

#### Sample SMTP with STARTTLS

Here is a working sample with error checking omitted for brevity for accessing smtp.gmail.com over port 587.

At first the communication goes over unencrypted plain-text socket, then later it is switched to TLS secured one before issuing the final `QUIT` command.

    With New cTlsSocket
        .SyncConnect "smtp.gmail.com", 587, UseTls:=False
        Debug.Print .SyncReceiveText();
        .SyncSendText "HELO 127.0.0.1" & vbCrLf
        Debug.Print .SyncReceiveText();
        .SyncSendText "STARTTLS" & vbCrLf
        Debug.Print .SyncReceiveText();
        .SyncStartTls "smtp.gmail.com"
        Debug.Print "TLS handshake complete: " & .RemoteHostName
        .SyncSendText "QUIT" & vbCrLf
        Debug.Print .SyncReceiveText();
    End With

Which produces debug output in `Immediate Window` similar to this:
    
    220 smtp.gmail.com ESMTP c69sm2955334lfg.23 - gsmtp
    250 smtp.gmail.com at your service
    220 2.0.0 Ready to start TLS
    1428790.043 [INFO] Using TLS_AES_128_GCM_SHA256 from smtp.gmail.com [mdTlsThunks.pvTlsParseHandshakeServerHello]
    1428790.057 [INFO] Valid ECDSA_SECP256R1_SHA256 signature [mdTlsThunks.pvTlsSignatureVerify]
    TLS handshake complete: smtp.gmail.com
    221 2.0.0 closing connection c69sm2955334lfg.23 - gsmtp

### Is it any good?

[Yes](https://news.ycombinator.com/item?id=3067434).

### Implemented Cipher Suites

This list includes cipher suites as implemented in the ASM thunks backend while the native backend list depends on the OS version and SSPI/Schannel settings.

Cipher Suite | First&nbsp;In | Selection String | Notes
--|--|--|--
<sub><sup>TLS_AES_128_GCM_SHA256                          </sup></sub>|TLS 1.3|EECDH+AESGCM|AEAD
<sub><sup>TLS_AES_256_GCM_SHA384                          </sup></sub>|TLS 1.3|EECDH+AESGCM|AEAD
<sub><sup>TLS_CHACHA20_POLY1305_SHA256                    </sup></sub>|TLS 1.3|EECDH+AESGCM|AEAD
<sub><sup>TLS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256         </sup></sub>|TLS 1.2|EECDH+AESGCM|AEAD
<sub><sup>TLS_ECDHE_RSA_WITH_AES_128_GCM_SHA256           </sup></sub>|TLS 1.2|EECDH+AESGCM|AEAD
<sub><sup>TLS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384         </sup></sub>|TLS 1.2|EECDH+AESGCM|AEAD
<sub><sup>TLS_ECDHE_RSA_WITH_AES_256_GCM_SHA384           </sup></sub>|TLS 1.2|EECDH+AESGCM|AEAD
<sub><sup>TLS_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256   </sup></sub>|TLS 1.2|EECDH+CHACHA20|AEAD
<sub><sup>TLS_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256     </sup></sub>|TLS 1.2|EECDH+CHACHA20|AEAD
<sub><sup>TLS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA256         </sup></sub>|TLS 1.2|EECDH+AES+SHA256|Weak, Exotic
<sub><sup>TLS_ECDHE_RSA_WITH_AES_128_CBC_SHA256           </sup></sub>|TLS 1.2|EECDH+AES+SHA256|Weak, Exotic
<sub><sup>TLS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA384         </sup></sub>|TLS 1.2|EECDH+AES+SHA384|Weak, Exotic
<sub><sup>TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA384           </sup></sub>|TLS 1.2|EECDH+AES+SHA384|Weak, Exotic
<sub><sup>TLS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA            </sup></sub>|TLSv1|EECDH+AES+SHA1|Weak, HMAC-SHA1
<sub><sup>TLS_ECDHE_RSA_WITH_AES_128_CBC_SHA              </sup></sub>|TLSv1|EECDH+AES+SHA1|Weak, HMAC-SHA1
<sub><sup>TLS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA            </sup></sub>|TLSv1|EECDH+AES+SHA1|Weak, HMAC-SHA1
<sub><sup>TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA              </sup></sub>|TLSv1|EECDH+AES+SHA1|Weak, HMAC-SHA1
<sub><sup>TLS_RSA_WITH_AES_128_GCM_SHA256                 </sup></sub>|TLS 1.2|RSA+AESGCM|Weak, No FS
<sub><sup>TLS_RSA_WITH_AES_256_GCM_SHA384                 </sup></sub>|TLS 1.2|RSA+AESGCM|Weak, No FS
<sub><sup>TLS_RSA_WITH_AES_128_CBC_SHA256                 </sup></sub>|TLS 1.2|RSA+AES+SHA256|Weak, No FS, Exotic
<sub><sup>TLS_RSA_WITH_AES_256_CBC_SHA256                 </sup></sub>|TLS 1.2|RSA+AES+SHA256|Weak, No FS, Exotic
<sub><sup>TLS_RSA_WITH_AES_128_CBC_SHA                    </sup></sub>|SSLv3|RSA+AES+SHA1|Weak, No FS, HMAC-SHA1
<sub><sup>TLS_RSA_WITH_AES_256_CBC_SHA                    </sup></sub>|SSLv3|RSA+AES+SHA1|Weak, No FS, HMAC-SHA1

Note that "exotic" cipher suites are included behind a conditional compilation flag only (off by default).

### ToDo

 - [ ] Allow client to assign client certificate for connection
 - [ ] Provide UI for end-user to choose suitable certificates from Personal certificate store
 - [ ] Add wrappers for http and ftp protocols
 - [x] Add WinSock control replacement
 - [ ] Add more samples (incl. `vbcurl.exe` utility)
 - [x] Refactor subclassing thunk to use msg queue not to re-enter IDE in debug mode

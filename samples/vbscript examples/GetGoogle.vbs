Option Explicit

' This is a simple barebones GET for VBScript
' It works for me in Win7-64.

' I put the HttpRequest.dll from here into SysWOW64 and registered it with regsvr32.exe.

' Be sure to run using the 32-bit cscript, if you have a 64-bit OS
' C:\Windows\SysWOW64\cscript.exe GetGoogle.vbs

Dim URL : URL = "https://www.google.com"

With CreateObject("HttpRequest.cHttpRequest")
  .Open_ "GET", URL
  .Send

  If .Status = 200 Then
    msgbox "RESPONSE : " & .responseText
  Else
    msgbox "Oops, Status : " & .Status
  End If
End With

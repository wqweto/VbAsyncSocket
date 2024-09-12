Option Explicit

' This script sets or unsets someone to be a VIP in a Twitch channel.
' It works for me in Win7-64.
' I put the HttpRequest.dll from here into SysWOW64 and registered it with regsvr32.exe.

' Example:
' C:\Windows\SysWOW64\cscript.exe TwitchVIP.vbs 1 username
' will set username as VIP.  (0 to remove)

Dim clientid, apioauth, broadcasterid
' Note:  These need to be set, and the oAuth needs to permit this action,
' but that's outside the scope of this example

Dim URL, userid
URL = "https://api.twitch.tv/helix/users?login=" & WScript.Arguments(1)

With CreateObject("HttpRequest.cHttpRequest")
  .Open_ "GET",URL
  .setRequestHeader "Client-Id",clientid
  .setRequestHeader "Authorization", "Bearer " & apioauth
  .Send

'   msgbox "RESPONSE : " & .responseText
  If .Status = 200 Then
      Dim y, html : Set html = CreateObject("htmlfile")
      Dim w : Set w = html.parentWindow
      w.execScript "var json=" & .responseText & ";var e=new Enumerator(json.data);", "JScript"
      While Not w.e.atEnd()
          Set y = w.e.item()
          userid=y.id
          w.e.moveNext
      Wend
  End If
End With

With CreateObject("HttpRequest.cHttpRequest")
  URL = "https://api.twitch.tv/helix/channels/vips"
  If WScript.Arguments(0) = 1 Then
    .open_ "POST",URL
  Else
    .open_ "DELETE",URL
  End If
  .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  .setRequestHeader "Authorization", "Bearer " & apioauth
  .setRequestHeader "Client-ID",clientid
  .send "broadcaster_id=" & broadcasterid & "&user_id=" & userid

' msgbox .responseText

End With
'Use a remote computer to ping another remote computer

Option Explicit 
'Change the SourceServer and RemoteServer Strings below to servernames or IP addresses for you.
'wscript.Echo RemotePing("SourceComputer", "DestinationComputer")
wscript.Echo RemotePing(".", WScript.Arguments(0))

Function RemotePing( SourceComputer, DestinationComputer)

 Dim strComputer1, strComputer2
 Dim objWMIService, colItems, objItem

 strComputer1 = SourceComputer
 strComputer2 = DestinationComputer

 On Error Resume Next
  ' error control block
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}//" & strComputer1 & "\root\cimv2")
  Set colItems = objWMIService.ExecQuery ("Select * from Win32_PingStatus " & "Where Address = '" & strComputer2 & "'")
  For Each objItem in colItems
      If objItem.StatusCode = 0 Then
          RemotePing = 1         
      Else
       RemotePing = 0
      End If
  Next
 On Error GoTo 0
End function

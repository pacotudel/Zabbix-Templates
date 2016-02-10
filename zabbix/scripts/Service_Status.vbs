strComputer = "." 
'srvc_name = "WerSvcl"
resp="unknown"
' Revisar argumentos
Set args = WScript.Arguments
Dim srvc_name
srvc_name = args(0)

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery("SELECT State FROM Win32_Service where Name= '" & srvc_name & "' ",,48) 
For Each objItem in colItems 
	resp=objItem.State
Next
Wscript.Echo resp
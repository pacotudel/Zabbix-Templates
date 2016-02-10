' 100 * Free / Total =  Porcentaje Libre
' 100 * (Total - Free) / Total =  Porcentaje Usado
'----------------------------------------------------------------------------------------------
' Script creates mapped  and checks network share for free space.
' Test on windows 7 workstation.

'on error resume next
Dim sPath
Dim drv
Dim oFs
Dim intDiskSize, strDiskSize
Dim objNetwork, strLocalDrive, strRemoteShare

Dim VERSION 
Dim INI_FILE
VERSION="1.0"
INI_FILE="c:\zabbix\scripts\ovt.ini"

sPath = "R:" 'set drive name

' Revisar argumentos
set args = WScript.Arguments
'sTipo = args(0)


select case args.Count
case 0
    Help
case 1
    Help
case 2
    Help
case 3
    Help
case 4
	sNAS = args(0)
	sTipo = args(1)
	sServer = args(2)
	sRuta = args(3)
	
	'sUsuario = "Administrador"
	sUsuario = ReadIni( INI_FILE, sNAS, "USER" )
	'sPassword = "Overtel_10"
	sPassword = ReadIni( INI_FILE, sNAS, "PASS" )
	strRemoteShare = "\\" & sServer & "\" & sRuta

	' Debug de los datos enviados
	'Wscript.Echo "Server:" & sServer & " Ruta: " & sRuta & " sTipo: " & sTipo & " NAS: " & sNAS
	Set objNetwork = WScript.CreateObject("WScript.Network")
	Set oFs = CreateObject("Scripting.FileSystemObject")

	' Section to map the R: network drive
	If (oFs.DriveExists(sPath) = True) Then
		objNetwork.RemoveNetworkDrive sPath, True, True
	End If

	objNetwork.MapNetworkDrive sPath, strRemoteShare, False, sUsuario , sPassword

	Set oFs = CreateObject("Scripting.FileSystemObject")
	Set drv = oFs.GetDrive(oFs.GetDriveName(sPath))
	strDiskSize = Int(drv.FreeSpace / (1024 * 1024 * 1024))
	intDiskSize = CLng(strDiskSize)
	intFreeSize = intDiskSize
	strTotalSize =  Int(drv.TotalSize / (1024 * 1024 * 1024))
	intTotalSize = CLng(strTotalSize)
	PL = (100 * intFreeSize) / intTotalSize
	PU = (100 * (intTotalSize - intFreeSize)) / intTotalSize
	U = (intTotalSize - intFreeSize)
	objNetwork.RemoveNetworkDrive "R:"
	'WScript.Echo("PON:" & sTipo & " - " & args(1))

	If sTipo = "T" Then 'Total
		WScript.Echo(strTotalSize)
	ElseIf sTipo = "L" Then 'Libre
		WScript.Echo(strDiskSize)
	ElseIf sTipo = "U" Then 'Usado
		Wscript.Echo U
	ElseIf sTipo = "PL" Then 'Porcentaje Libre
		'WScript.Echo Round(PL,2)
		'WScript.Echo Round(PL)
		WScript.Echo Int(PL) & "." & (100*round(PL-Int(PL),2))
	ElseIf sTipo = "PU" Then 'Porcentaje Usado
		'WScript.Echo Round(PU,2)
		'WScript.Echo Round(PU)
		WScript.Echo Int(PU) & "." & (100*round(PU-Int(PU),2))
	ElseIf sTipo = "LEC" Then 'Datos en formato CSV
		WScript.Echo ("Dispon: " & strDiskSize & "Gb de un Total de: " & strTotalSize & "Gb (" & Round(PL,2) & "%)" )
	Else
		WScript.Echo("T;" & strTotalSize & ";L;" & strDiskSize & ";U;" & U & ";PL;" & Round(PL,2) & ";PU;" & Round(PU,2))
	End If    

case else
	Help
end select



If (intDiskSize > strTotalSize) Then
    WScript.Echo("-1")
ELSE
    ' WScript.Echo("Free Space: " & strDiskSize & " GB")
	' WScript.Echo(strTotalSize & ";" & strDiskSize & ";" & (strTotalSize-strTotalSize))
End If

If err.number <> 0 Then
    ' WScript.Echo ("Script Check Failed") 
    Wscript.Quit (1001)
Else
    ' WScript.Echo ("Successfully Passed")
    WScript.Quit(0) 
End If
 
objNetwork.RemoveNetworkDrive sPath

Function ReadIni( myFilePath, mySection, myKey )
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If
End Function


Function Help() 
	WScript.Echo "Args: (" & args.count & ")"
	For Each strArg in args
		WScript.Echo " - " & strArg
	Next
	'sNAS = args(0)
	'sTipo = args(1)
	'sServer = args(2)
	'sRuta = args(3)
    WScript.Echo "Ayuda: "
	WScript.Echo " Param1: sNAS"
	WScript.Echo "  Nombre del Nas como aparece en el INI de configuracion"
	WScript.Echo " Param2: sTipo"
	WScript.Echo "  T - Total en Gb, L - Libre en Gb, U - Usado en Gb"
	WScript.Echo "  PL - Porcentaje Libre en Gb, PU - Porcentaje Usado en Gb"
	WScript.Echo "  LEC - Formato para facilitar lectura"
	WScript.Echo "  CSV - Todo separado por comas"
	WScript.Echo " Param3: sServer"
	WScript.Echo "  IP del NAS"
	WScript.Echo " Param4: sRuta"
	WScript.Echo "  Carpeta compartida"
	WScript.Echo "  ----..------------------------------------.----------------------------------"
End Function
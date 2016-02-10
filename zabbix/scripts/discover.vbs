' Revisar argumentos
Set args = WScript.Arguments
Dim fName
Dim VERSION 
Dim INI_FILE
VERSION="1.0"
'INI_FILE="c:\zabbix\scripts\ovt.ini"
'INI_FILE=".\ovt.ini"
includeFile "c:\zabbix\scripts\ini_location.vbs"
'WScript.Echo "INI_FILE= " & INI_FILE

' dim fso: set fso = CreateObject("Scripting.FileSystemObject")
' WScript.Echo fso.GetAbsolutePathName(".")
    

' Revisar argumentos
set args = WScript.Arguments
' Parse args
select case args.Count
case 0
    Help
case 1
    sTipo = args(0)
	If sTipo = "SERVICIO" Then 'Regla de descubrimiento de SERVICIOS
        WScript.Echo DiscoverSectionsOnIniSERVICIOS( INI_FILE, "SERVICIO")
    ElseIf sTipo = "NAS" Then 'Cuenta de carpetas
        WScript.Echo DiscoverSectionsOnIniNAS( INI_FILE, "NAS")
	ElseIf sTipo = "COBIAN" Then 'Cuenta de carpetas
        WScript.Echo DiscoverSectionsOnIniCOBIAN( INI_FILE, "COBIAN")
	ElseIf sTipo = "FICHERO" Then 'Cuenta de carpetas
        'FolderCount(folder)
	ElseIf sTipo = "PING" Then 'Discovery de Ping
        WScript.Echo DiscoverSectionsOnIniPING( INI_FILE, "PING")
	ElseIf sTipo = "SNMP" Then 'Discovery de SNMP
        WScript.Echo DiscoverSectionsOnIniSNMP( INI_FILE, "SNMP")
	ElseIf sTipo = "NSFA" Then 'Discovery de SNMP
        WScript.Echo DiscoverSectionsOnIniNSFA( INI_FILE, "NSFA")
	ElseIf sTipo = "DEBUG" Then 'Cuenta de carpetas
        debug()
	ElseIf sTipo = "VERSION" Then 'Cuenta de carpetas
        WScript.Echo VERSION
	Else
		WScript.Echo("")
	End If
case else
	Help
end select

Function Help() 
    WScript.Echo "Ayuda: "
	WScript.Echo " Param1: SERVICIO"
	WScript.Echo "  Descubrimiento de Servicios"
	WScript.Echo " Param1: NAS"
	WScript.Echo "  Descubrimiento de NAS y carpetas a revisar en estos"
	WScript.Echo " Param1: COBIAN"
	WScript.Echo "  Descubrimiento de copias de cobian, carpeta donde van a parar"
	WScript.Echo " Param1: SNMP"
	WScript.Echo "  Descubrimiento de dispositivos a los que revisar datos por SNMP"
	WScript.Echo " Param1: PING"
	WScript.Echo "  Descubrimiento de equipos a los que hacer ping"
	WScript.Echo " Param1: DEBUG"
	WScript.Echo "  Lee el fichero .\ovt.ini y muestra algunos valores"
	WScript.Echo "  -----------.-----------.-----------------------------------------------------"
End Function

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

Function CountSectionsOnIni( myFilePath, mySection)
	Dim count 
	count=0
	search = "[" & mySection
	With CreateObject("Scripting.FileSystemObject")
	  Set iniFile = .OpenTextFile(myFilePath)
	End With
	While iniFile.AtEndOfStream = False 
		line = Trim(iniFile.ReadLine)
		If Left(line, 1) = "[" And Right(line, 1) = "]" Then
			'WScript.Echo "Seccion: " & line & " --- " & search
			If InStr(1, line, "[" & mySection) = 1 Then count = count + 1
		End If
	WEnd
	
	CountSectionsOnIni = count
End Function

Function DiscoverSectionsOnIniSERVICIOS( myFilePath, mySection)
	Dim blnKeyExists 
	Dim incexA
	'blnKeyExists     = ( ReadIni( myFilePath, mySection, strKey,"SERVICIO" ) <> "" )
	count = (CountSectionsOnIni(myFilePath, mySection))
	' Inicio de JSON
	s = "{" & vbCrLf & vbTab & chr(34) & "data" & chr(34) & ":[" & vbCrLf
	i = 0
	For indexA = 1 to count
		sect = "SERVICIO" & indexA
		'WScript.Echo sect
		nombre = ReadIni( myFilePath, sect, "SERVICIO" )
		descripcion = ReadIni( myFilePath, sect, "DESCRIPCION" )
		Dim WshNetwork
		Set WshNetwork = CreateObject("WScript.Network")
		Set wshShell = CreateObject( "WScript.Shell" )
		strUserDomain = wshShell.ExpandEnvironmentStrings( "%USERDOMAIN%" )
		host = WshNetwork.ComputerName & "_" &strUserDomain
		' Solo nombre
		' s = s & vbTab & " { " & chr(34) & "{#FSSERVICE}" & chr(34) & ":" & chr(34) & nombre & chr(34) & " }"
		' Nombre y descripcion
		' D - Discover
		' S - Service
		' SERVICE
		' DESCRIPTION
		s = s & vbTab & " { " & chr(34) & "{#DSSERVICENAME}" & chr(34) & ":" & chr(34) & nombre & chr(34) & ", " & chr(34) & "{#DSSERVICEDESCRIPTION}" & chr(34) & ":" & chr(34) & descripcion & chr(34) & ", " & chr(34) & "{#DSSERVICEHOST}" & chr(34) & ":" & chr(34) & host & chr(34) & " }"
		i = i + 1
		if  i < count then
			s = s & "," & vbCrLf
		end if
    Next 
	
	s = s & vbCrLf & vbTab & "]" & vbCrLf & "}"
	DiscoverSectionsOnIniSERVICIOS = s
End Function

Function DiscoverSectionsOnIniPING( myFilePath, mySection)
	Dim blnKeyExists 
	Dim incexA
	'blnKeyExists     = ( ReadIni( myFilePath, mySection, strKey,"SERVICIO" ) <> "" )
	count = (CountSectionsOnIni(myFilePath, mySection))
	' Inicio de JSON
	s = "{" & vbCrLf & vbTab & chr(34) & "data" & chr(34) & ":[" & vbCrLf
	i = 0
	For indexA = 1 to count
		sect = "PING" & indexA
		'WScript.Echo sect
		ip = ReadIni( myFilePath, sect, "IP" )
		name = ReadIni( myFilePath, sect, "NAME" )
		' Solo ip
		's = s & vbTab & " { " & chr(34) & "{#FSPING}" & chr(34) & ":" & chr(34) & ip & chr(34) & " }"
		' Nombre y descripcion
		' D - Discover
		' S - Service
		' PINGIP
		' PINGNAME
		s = s & vbTab & " { " & chr(34) & "{#DSPINGIP}" & chr(34) & ":" & chr(34) & ip & chr(34) & ", " & chr(34) & "{#DSPINGNAME}" & chr(34) & ":" & chr(34) & name & chr(34) & " }"
		i = i + 1
		if  i < count then
			s = s & "," & vbCrLf
		end if
    Next 
	
	s = s & vbCrLf & vbTab & "]" & vbCrLf & "}"
	DiscoverSectionsOnIniPING = s
End Function

Function DiscoverSectionsOnIniNAS( myFilePath, mySection)
	Dim blnKeyExists 
	Dim incexA
	count = (CountSectionsOnIni(myFilePath, mySection))
	' Inicio de JSON
	s = "{" & vbCrLf & vbTab & chr(34) & "data" & chr(34) & ":[" & vbCrLf
	i = 0
	For indexA = 1 to count
		sect = "NAS" & indexA
		'WScript.Echo sect
		ip   = ReadIni( myFilePath, sect, "IP" )
		name = ReadIni( myFilePath, sect, "NAME" )
		sf   = ReadIni( myFilePath, sect, "SF" )
		fld  = ReadIni( myFilePath, sect, "FLD" )
		' Los campos user y pass pueden leerse pero mejor que se queden en local y sea el script quien los maneje
		'user  = ReadIni( myFilePath, sect, "USER" )
		'pass  = ReadIni( myFilePath, sect, "PASS" )
		' Nombre y descripcion
		' D - Discover
		' S - Service
		' NASIP
		' NASNAME
		s = s & vbTab & " { " & chr(34) & "{#DSNASIP}" & chr(34) & ":" & chr(34) & ip & chr(34) & ", " & chr(34) & "{#DSNASNAME}" & chr(34) & ":" & chr(34) & name & chr(34) & ", " & chr(34) & "{#DSNASSF}" & chr(34) & ":" & chr(34) & sf & chr(34) & ", " & chr(34) & "{#DSNASNAME2}" & chr(34) & ":" & chr(34) & sect & chr(34) & ", " & chr(34) & "{#DSNASFLD}" & chr(34) & ":" & chr(34) & fld & chr(34) & " }"
		i = i + 1
		if  i < count then
			s = s & "," & vbCrLf
		end if
    Next 
	
	s = s & vbCrLf & vbTab & "]" & vbCrLf & "}"
	DiscoverSectionsOnIniNAS = s
End Function

Function DiscoverSectionsOnIniSNMP( myFilePath, mySection)
	Dim blnKeyExists 
	Dim incexA
	count = (CountSectionsOnIni(myFilePath, mySection))
	' Inicio de JSON
	s = "{" & vbCrLf & vbTab & chr(34) & "data" & chr(34) & ":[" & vbCrLf
	i = 0
	For indexA = 1 to count
		sect = "SNMP" & indexA
		'WScript.Echo sect
		ip   = ReadIni( myFilePath, sect, "IP" )
		name = ReadIni( myFilePath, sect, "NAME" )
		user  = ReadIni( myFilePath, sect, "USER" )
		oid   = ReadIni( myFilePath, sect, "OID" )
		oidn  = ReadIni( myFilePath, sect, "OIDDESCRIPTION" )
		' Nombre y descripcion
		' D - Discover
		' S - Service
		' NASIP
		' NASNAME
		s = s & vbTab & " { " & chr(34) & "{#DSSNMPIP}" & chr(34) & ":" & chr(34) & ip & chr(34) & ", " & chr(34) & "{#DSSNMPNAME}" & chr(34) & ":" & chr(34) & name & chr(34) & ", " & chr(34) & "{#DSSNMPUSER}" & chr(34) & ":" & chr(34) & user & chr(34) & ", " & chr(34) & "{#DSSNMPOID}" & chr(34) & ":" & chr(34) & oid & chr(34) & ", " & chr(34) & "{#DSSNMPOIDDESC}" & chr(34) & ":" & chr(34) & oidn & chr(34) & " }"
		i = i + 1
		if  i < count then
			s = s & "," & vbCrLf
		end if
    Next 
	
	s = s & vbCrLf & vbTab & "]" & vbCrLf & "}"
	DiscoverSectionsOnIniSNMP = s
End Function

Function DiscoverSectionsOnIniCOBIAN( myFilePath, mySection)
	Dim blnKeyExists 
	Dim incexA
	
	count = (CountSectionsOnIni(myFilePath, mySection))
	' Inicio de JSON
	s = "{" & vbCrLf & vbTab & chr(34) & "data" & chr(34) & ":[" & vbCrLf
	i = 0
	For indexA = 1 to count
		sect = "COBIAN" & indexA
		'WScript.Echo sect
		tarea = ReadIni( myFilePath, sect, "TAREA" )
		destino = ReadIni( myFilePath, sect, "DESTINO" )
		Dim WshNetwork
		Set WshNetwork = CreateObject("WScript.Network")
		Set wshShell = CreateObject( "WScript.Shell" )
		strUserDomain = wshShell.ExpandEnvironmentStrings( "%USERDOMAIN%" )
		host = WshNetwork.ComputerName & "_" &strUserDomain
		' Solo nombre
		' s = s & vbTab & " { " & chr(34) & "{#FSSERVICE}" & chr(34) & ":" & chr(34) & nombre & chr(34) & " }"
		' Nombre y descripcion
		' D - Discover
		' S - Service
		' SERVICE
		' DESCRIPTION
		s = s & vbTab & " { " & chr(34) & "{#DSSERVICENAME}" & chr(34) & ":" & chr(34) & tarea & chr(34) & ", " & chr(34) & "{#DSCOBIANTASK}" & chr(34) & ":" & chr(34) & destino & chr(34) & ", " & chr(34) & "{#DSCOBIANDEST}" & chr(34) & ":" & chr(34) & host & chr(34) & " }"
		i = i + 1
		if  i < count then
			s = s & "," & vbCrLf
		end if
    Next 
	
	s = s & vbCrLf & vbTab & "]" & vbCrLf & "}"
	DiscoverSectionsOnIniCOBIAN = s
End Function

Function DiscoverSectionsOnIniNSFA( myFilePath, mySection)
	Dim blnKeyExists 
	Dim incexA
	count = (CountSectionsOnIni(myFilePath, mySection))
	' Inicio de JSON
	s = "{" & vbCrLf & vbTab & chr(34) & "data" & chr(34) & ":[" & vbCrLf
	i = 0
	For indexA = 1 to count
		sect = "NSFA" & indexA
		'WScript.Echo sect
		ip   = ReadIni( myFilePath, sect, "IP" )
		name = ReadIni( myFilePath, sect, "NAME" )
		sf   = ReadIni( myFilePath, sect, "SF" )
		fld  = ReadIni( myFilePath, sect, "FLD" )
		ext  = ReadIni( myFilePath, sect, "EXT" )
		nh   = ReadIni( myFilePath, sect, "NH" )
		oh   = ReadIni( myFilePath, sect, "OH" )
		' Los campos user y pass pueden leerse pero mejor que se queden en local y sea el script quien los maneje
		'user  = ReadIni( myFilePath, sect, "USER" )
		'pass  = ReadIni( myFilePath, sect, "PASS" )
		' Nombre y descripcion
		' D - Discover
		' S - Service
		' NASIP
		' NASNAME
		s = s & vbTab & " { " & chr(34) & "{#DSNSFAIP}" & chr(34) & ":" & chr(34) & ip & chr(34) _ 
		               & ", " & chr(34) & "{#DSNSFANAME}" & chr(34) & ":" & chr(34) & name & chr(34) _ 
		               & ", " & chr(34) & "{#DSNSFASF}" & chr(34) & ":" & chr(34) & sf & chr(34) _
					   & ", " & chr(34) & "{#DSNSFANAME2}" & chr(34) & ":" & chr(34) & sect & chr(34) _
					   & ", " & chr(34) & "{#DSNSFAEXT}" & chr(34) & ":" & chr(34) & ext & chr(34) _
					   & ", " & chr(34) & "{#DSNSFANH}" & chr(34) & ":" & chr(34) & nh & chr(34) _
					   & ", " & chr(34) & "{#DSNSFAOH}" & chr(34) & ":" & chr(34) & oh & chr(34) _
					   & ", " & chr(34) & "{#DSNSFAFLD}" & chr(34) & ":" & chr(34) & fld & chr(34) & " }"
		i = i + 1
		if  i < count then
			s = s & "," & vbCrLf
		end if
    Next 
	
	s = s & vbCrLf & vbTab & "]" & vbCrLf & "}"
	DiscoverSectionsOnIniNSFA = s
End Function

Sub WriteIni( myFilePath, mySection, myKey, myValue )
    ' This subroutine writes a value to an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be written
    ' myValue     [string]  the value to be written (myKey will be
    '                       deleted if myValue is <DELETE_THIS_VALUE>)
    '
    ' Returns:
    ' N/A
    '
    ' CAVEAT:     WriteIni function needs ReadIni function to run
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre, Johan Pol and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim blnInSection, blnKeyExists, blnSectionExists, blnWritten
    Dim intEqualPos
    Dim objFSO, objNewIni, objOrgIni, wshShell
    Dim strFilePath, strFolderPath, strKey, strLeftString
    Dim strLine, strSection, strTempDir, strTempFile, strValue

    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )
    strValue    = Trim( myValue )

    Set objFSO   = CreateObject( "Scripting.FileSystemObject" )
    Set wshShell = CreateObject( "WScript.Shell" )

    strTempDir  = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    strTempFile = objFSO.BuildPath( strTempDir, objFSO.GetTempName )

    Set objOrgIni = objFSO.OpenTextFile( strFilePath, ForReading, True )
    Set objNewIni = objFSO.CreateTextFile( strTempFile, False, False )

    blnInSection     = False
    blnSectionExists = False
    ' Check if the specified key already exists
    blnKeyExists     = ( ReadIni( strFilePath, strSection, strKey ) <> "" )
    blnWritten       = False

    ' Check if path to INI file exists, quit if not
    strFolderPath = Mid( strFilePath, 1, InStrRev( strFilePath, "\" ) )
    If Not objFSO.FolderExists ( strFolderPath ) Then
        WScript.Echo "Error: WriteIni failed, folder path (" _
                   & strFolderPath & ") to ini file " _
                   & strFilePath & " not found!"
        Set objOrgIni = Nothing
        Set objNewIni = Nothing
        Set objFSO    = Nothing
        WScript.Quit 1
    End If

    While objOrgIni.AtEndOfStream = False
        strLine = Trim( objOrgIni.ReadLine )
        If blnWritten = False Then
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                blnSectionExists = True
                blnInSection = True
            ElseIf InStr( strLine, "[" ) = 1 Then
                blnInSection = False
            End If
        End If

        If blnInSection Then
            If blnKeyExists Then
                intEqualPos = InStr( 1, strLine, "=", vbTextCompare )
                If intEqualPos > 0 Then
                    strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                    If LCase( strLeftString ) = LCase( strKey ) Then
                        ' Only write the key if the value isn't empty
                        ' Modification by Johan Pol
                        If strValue <> "<DELETE_THIS_VALUE>" Then
                            objNewIni.WriteLine strKey & "=" & strValue
                        End If
                        blnWritten   = True
                        blnInSection = False
                    End If
                End If
                If Not blnWritten Then
                    objNewIni.WriteLine strLine
                End If
            Else
                objNewIni.WriteLine strLine
                    ' Only write the key if the value isn't empty
                    ' Modification by Johan Pol
                    If strValue <> "<DELETE_THIS_VALUE>" Then
                        objNewIni.WriteLine strKey & "=" & strValue
                    End If
                blnWritten   = True
                blnInSection = False
            End If
        Else
            objNewIni.WriteLine strLine
        End If
    Wend

    If blnSectionExists = False Then ' section doesn't exist
        objNewIni.WriteLine
        objNewIni.WriteLine "[" & strSection & "]"
            ' Only write the key if the value isn't empty
            ' Modification by Johan Pol
            If strValue <> "<DELETE_THIS_VALUE>" Then
                objNewIni.WriteLine strKey & "=" & strValue
            End If
    End If

    objOrgIni.Close
    objNewIni.Close

    ' Delete old INI file
    objFSO.DeleteFile strFilePath, True
    ' Rename new INI file
    objFSO.MoveFile strTempFile, strFilePath

    Set objOrgIni = Nothing
    Set objNewIni = Nothing
    Set objFSO    = Nothing
    Set wshShell  = Nothing
End Sub

Function Debug()
	WScript.Echo "-----------------------------------------------------"
	WScript.Echo "NAS 1"
	WScript.Echo "IP            : " & ReadIni( INI_FILE, "NAS1", "IP" ) & vbCrLf & _
				 "Shared Folder : " & ReadIni( INI_FILE, "NAS1", "SF" ) & vbCrLf & _
				 "Folder        : " & ReadIni( INI_FILE, "NAS1", "FLD" ) & vbCrLf & _
				 "USER          : " & ReadIni( INI_FILE, "NAS1", "USER" ) & vbCrLf & _
				 "PASS          : " & ReadIni( INI_FILE, "NAS1", "PASS" )
				 
	WScript.Echo "NAS 2"
	WScript.Echo "IP            : " & ReadIni( INI_FILE, "NAS2", "IP" ) & vbCrLf & _
				 "Shared Folder : " & ReadIni( INI_FILE, "NAS2", "SF" ) & vbCrLf & _
				 "Folder        : " & ReadIni( INI_FILE, "NAS2", "FLD" ) & vbCrLf & _
				 "USER          : " & ReadIni( INI_FILE, "NAS2", "USER" ) & vbCrLf & _
				 "PASS          : " & ReadIni( INI_FILE, "NAS2", "PASS" )

	WScript.Echo "NAS 3"
	WScript.Echo "IP            : " & ReadIni( INI_FILE, "NAS3", "IP" ) & vbCrLf & _
				 "Shared Folder : " & ReadIni( INI_FILE, "NAS3", "SF" ) & vbCrLf & _
				 "Folder        : " & ReadIni( INI_FILE, "NAS3", "FLD" ) & vbCrLf & _
				 "USER          : " & ReadIni( INI_FILE, "NAS3", "USER" ) & vbCrLf & _
				 "PASS          : " & ReadIni( INI_FILE, "NAS3", "PASS" )
	WScript.Echo "-----------------------------------------------------"
	WScript.Echo "COBIAN 1"
	WScript.Echo "Tarea         : " & ReadIni( INI_FILE, "COBIAN1", "Tarea" ) & vbCrLf & _
				 "Destino       : " & ReadIni( INI_FILE, "COBIAN1", "Destino" )
	WScript.Echo "-----------------------------------------------------"
	WScript.Echo "Servicio 1"
	servicio="SERVICIO1"
	WScript.Echo "Tarea         : " & ReadIni( INI_FILE, servicio, "SERVICIO" ) & vbCrLf & _
				 " Descripcion  : " & ReadIni( INI_FILE, servicio, "DESCRIPCION" )
	WScript.Echo "Servicio 2"
	servicio="SERVICIO2"
	WScript.Echo "Tarea         : " & ReadIni( INI_FILE, servicio, "SERVICIO" ) & vbCrLf & _ 
				 " Descripcion  : " & ReadIni( INI_FILE, servicio, "DESCRIPCION" )
	WScript.Echo "-----------------------------------------------------"
	servicio="PING1"
	WScript.Echo servicio
	WScript.Echo " IP           : " & ReadIni( INI_FILE, servicio, "IP" ) & vbCrLf & _
				 " Nombre       : " & ReadIni( INI_FILE, servicio, "NAME" )
	servicio="PING2"
	WScript.Echo servicio
	WScript.Echo " IP           : " & ReadIni( INI_FILE, servicio, "IP" ) & vbCrLf & _
				 " Nombre       : " & ReadIni( INI_FILE, servicio, "NAME" )
	WScript.Echo "-----------------------------------------------------"
	servicio="SNMP1"
	WScript.Echo servicio
	WScript.Echo " IP           : " & ReadIni( INI_FILE, servicio, "IP" ) & vbCrLf & _
				 " Nombre       : " & ReadIni( INI_FILE, servicio, "NAME" ) & vbCrLf & _
				 " User         : " & ReadIni( INI_FILE, servicio, "USER" ) & vbCrLf & _
				 " OID          : " & ReadIni( INI_FILE, servicio, "OID" ) & vbCrLf & _
				 " OIDDESCRIPT  : " & ReadIni( INI_FILE, servicio, "OIDDESCRIPTION" )
	
	WScript.Echo "--------CUANTOS DE CADA---------------------------------------------"
	WScript.Echo "Servicios     : " & CountSectionsOnIni( INI_FILE, "SERVICIO")
	WScript.Echo "Cobian        : " & CountSectionsOnIni( INI_FILE, "COBIAN")
	WScript.Echo "Nas           : " & CountSectionsOnIni( INI_FILE, "NAS")
	WScript.Echo "Ping IP       : " & CountSectionsOnIni( INI_FILE, "PING")
	WScript.Echo "Snmp          : " & CountSectionsOnIni( INI_FILE, "SNMP")

	WScript.Echo "--------DISCOVERY------------------------------------"
	WScript.Echo "Servicios:"
	WScript.Echo "-----------------------------------------------------"
	WScript.Echo DiscoverSectionsOnIniSERVICIOS( INI_FILE, "SERVICIO")
	WScript.Echo "-----------------------------------------------------"
	WScript.Echo "Ping:"
	WScript.Echo "-----------------------------------------------------"
	WScript.Echo DiscoverSectionsOnIniPING( INI_FILE, "PING")
	WScript.Echo "-----------------------------------------------------"
	WScript.Echo "NAS:"
	WScript.Echo "-----------------------------------------------------"
	WScript.Echo DiscoverSectionsOnIniNAS( INI_FILE, "NAS")
	WScript.Echo "-----------------------------------------------------"
	WScript.Echo "SNMP:"
	WScript.Echo "-----------------------------------------------------"
	WScript.Echo DiscoverSectionsOnIniSNMP( INI_FILE, "SNMP")
	WScript.Echo "-------DONDE ESTA EL INI-----------------------------"
	WScript.Echo "INI_FILE= " & INI_FILE
	WScript.Echo "-----------------------------------------------------"
End Function
'Function startsWith(str As String, prefix As String) As Boolean
 '   startsWith = Left(str, Len(prefix)) = prefix
'End Function


Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub
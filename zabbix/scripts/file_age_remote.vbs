'Config para Zabbix
' parametros para file_age.vbs 
' $1=ip del nas 
' $2=carpeta compartida en el nas 
' $3=(si hace falta, es la carpeta en la que se busca el fichero mas moderno) 
' $4=
'    ON Older Name
'    OH Older Hours
'    NN Newer Name
'    NH Newer Hours
'    AL Todos los datos separados por coma NN,NDateLastModified,NH
'UserParameter=NetworkShare.FileAge.ON[*],cscript.exe c:\zabbix\scripts\file_age.vbs $1 $2 $3 ON //NoLogo
'UserParameter=NetworkShare.FileAge.OH[*],cscript.exe c:\zabbix\scripts\file_age.vbs $1 $2 $3 OH //NoLogo
'UserParameter=NetworkShare.FileAge.NN[*],cscript.exe c:\zabbix\scripts\file_age.vbs $1 $2 $3 NN //NoLogo
'UserParameter=NetworkShare.FileAge.NH[*],cscript.exe c:\zabbix\scripts\file_age.vbs $1 $2 $3 NH //NoLogo
'UserParameter=NetworkShare.FileAge.AL[*],cscript.exe c:\zabbix\scripts\file_age.vbs $1 $2 $3 AL //NoLogo
'
' Ejemplos de uso
'@set NAS=10.0.0.170
'@set SF=Backup2
'@echo -----------------------------------NAS %NAS% - %SF% --------------
'@echo Respuesta al PING
'@C:\zabbix\bin\win64\zabbix_get.exe -s 127.0.0.1 -k PingCheck[%NAS%] 
'@echo TOTAL:
'@C:\zabbix\bin\win64\zabbix_get.exe -s 127.0.0.1 -k NetworkShare.Total[%NAS%,%SF%]
'@echo USED:
'@C:\zabbix\bin\win64\zabbix_get.exe -s 127.0.0.1 -k NetworkShare.Used[%NAS%,%SF%]
'@echo FREE:
'@C:\zabbix\bin\win64\zabbix_get.exe -s 127.0.0.1 -k NetworkShare.Free[%NAS%,%SF%]
'@echo Porcentage USED:
'@C:\zabbix\bin\win64\zabbix_get.exe -s 127.0.0.1 -k NetworkShare.PUsed[%NAS%,%SF%]
'@echo Porcentage FREE:
'@C:\zabbix\bin\win64\zabbix_get.exe -s 127.0.0.1 -k NetworkShare.PFree[%NAS%,%SF%]
'@set FOLDER=a
'@echo Fichero mas antiguo en carpeta : %FOLDER% (Nombre) 
'@C:\zabbix\bin\win64\zabbix_get.exe -s 127.0.0.1 -k NetworkShare.FileAge.ON[%NAS%,%SF%,%FOLDER%,ON]
'@echo Fichero mas antiguo en carpeta : %FOLDER% (Edad en Horas) 
'@C:\zabbix\bin\win64\zabbix_get.exe -s 127.0.0.1 -k NetworkShare.FileAge.OH[%NAS%,%SF%,%FOLDER%,OH]
'@echo Fichero mas moderno en carpeta : %FOLDER% (Nombre) 
'@C:\zabbix\bin\win64\zabbix_get.exe -s 127.0.0.1 -k NetworkShare.FileAge.NN[%NAS%,%SF%,%FOLDER%,ON]
'@echo Fichero mas moderno en carpeta : %FOLDER% (Edad en Horas) 
'@C:\zabbix\bin\win64\zabbix_get.exe -s 127.0.0.1 -k NetworkShare.FileAge.NH[%NAS%,%SF%,%FOLDER%,OH]
'@echo Fichero mas moderno y mas antiguo, todos los datos en carpeta : %FOLDER% 
'@C:\zabbix\bin\win64\zabbix_get.exe -s 127.0.0.1 -k NetworkShare.FileAge.AL[%NAS%,%SF%,%FOLDER%,AL]

Dim sPath
Dim drv
Dim intDiskSize, strDiskSize

Dim VERSION 
Dim INI_FILE
VERSION="1.0"
INI_FILE="c:\zabbix\scripts\ovt.ini"

sPath = "R:" 'set drive name

' Revisar argumentos
set args = WScript.Arguments

select case args.Count
case 0
	Help
case 1
	Help
case 1
	Help
case 2
	Help
case 3
	Help
case 4
	Help
case 5
	CalcFile()
case 6
	CalcFile2(args(5)) 	
case else 
	Help
end select


Function CalcFile()
	sNAS = args(0)
	sTipo = args(1)
	sServer = args(2)
	sRuta = args(3)
	sFolder = args(4)
	Dim oFs
	Dim objNetwork, strLocalDrive, strRemoteShare
	Dim FSO
	Dim sFolder
	Const DateLastModified = 1, DateLastAccessed = 2

	strLocalDrive = "R:" 'set drive name 
	'sUsuario = "Administrador"
	sUsuario = ReadIni( INI_FILE, sNAS, "USER" )
	'sPassword = "Overtel_10"
	sPassword = ReadIni( INI_FILE, sNAS, "PASS" )
	strRemoteShare = "\\" & sServer & "\" & sRuta
	
	Set objNetwork = WScript.CreateObject("WScript.Network")
	Set oFs = CreateObject("Scripting.FileSystemObject")

	' Section to map the R: network drive
	If (oFs.DriveExists(strLocalDrive) = True) Then
		objNetwork.RemoveNetworkDrive strLocalDrive, True, True
	End If

	objNetwork.MapNetworkDrive strLocalDrive, strRemoteShare, False, sUsuario , sPassword

	Set FSO = CreateObject("Scripting.FileSystemObject")
	srcPath = strLocalDrive & "\" & sFolder
	If FSO.FolderExists(srcPath) Then
		Set srcFolder = FSO.GetFolder(srcPath)
	else
		Wscript.Echo "Folder not exist"
		WScript.Quit
	End If

	Dim oldestFile,  newerFile
	For Each objFile In srcFolder.Files
		' Search Older
		If oldestFile = "" Then
			Set oldestFile = objFile
		Else
			If objFile.DateLastModified < oldestFile.DateLastModified Then
				Set oldestFile = objFile
			End If
		End If
		' Search Newer
		If newerFile = "" Then
			Set newerFile = objFile
		Else
			If objFile.DateLastModified >= newerFile.DateLastModified Then
				Set newerFile = objFile
			End If
		End If
	Next

	'Debug()

	'Get name of file if necessary
	Set objFSO2 = CreateObject("Scripting.FileSystemObject")

	If sTipo = "OH" Then 'Older Hours
		WScript.Echo(FileAgeDiff(oldestFile,DateLastModified))
	ElseIf sTipo = "ON" Then 'Older Name
		WScript.Echo(objFSO2.GetFileName(oldestFile))
	ElseIf sTipo = "OS" Then 'Older Name
		If (objFSO2.FileExists(oldestFile)) Then
			WScript.Echo CLNG(objFSO2.GetFile(oldestFile).Size / 1024 )
		else
			WScript.Echo "0"
		end if
	ElseIf sTipo = "NH" Then 'Newer Hours
		WScript.Echo(FileAgeDiff(newerFile,DateLastModified))
	ElseIf sTipo = "NN" Then 'Newer Name
		WScript.Echo(objFSO2.GetFileName(newerFile))
	ElseIf sTipo = "NS" Then 'Older Name
		If (objFSO2.FileExists(newerFile)) Then
			WScript.Echo CLNG(objFSO2.GetFile(newerFile).Size / 1024)
		else
			WScript.Echo "0"
		end if
	ElseIf sTipo = "CSV" Then 'Datos en CSV
		If (objFSO2.FileExists(newerFile)) And (objFSO2.FileExists(oldestFile))   Then
			WScript.Echo(objFSO2.GetFileName(newerFile) & ";" & newerFile.DateLastModified & ";" & FileAgeDiff(newerFile,DateLastModified) & ";" & CLNG(objFSO2.GetFile(newerFile).Size / 1024 ) & ";" & objFSO2.GetFileName(oldestFile) & ";" & oldestFile.DateLastModified & ";" & FileAgeDiff(oldestFile,DateLastModified)) & ";" & CLNG(objFSO2.GetFile(oldestFile).Size / 1024 ) 
		else
			WScript.Echo ""
		end if
	End If
End Function

Function CalcFile2(ext)
	sNAS = args(0)
	sTipo = args(1)
	sServer = args(2)
	sRuta = args(3)
	sFolder = args(4)
	Dim oFs
	Dim objNetwork, strLocalDrive, strRemoteShare
	Dim FSO
	Dim sFolder
	Const DateLastModified = 1, DateLastAccessed = 2

	strLocalDrive = "R:" 'set drive name 
	'sUsuario = "Administrador"
	sUsuario = ReadIni( INI_FILE, sNAS, "USER" )
	'sPassword = "Overtel_10"
	sPassword = ReadIni( INI_FILE, sNAS, "PASS" )
	strRemoteShare = "\\" & sServer & "\" & sRuta
	
	Set objNetwork = WScript.CreateObject("WScript.Network")
	Set oFs = CreateObject("Scripting.FileSystemObject")

	' Section to map the R: network drive
	If (oFs.DriveExists(strLocalDrive) = True) Then
		objNetwork.RemoveNetworkDrive strLocalDrive, True, True
	End If

	objNetwork.MapNetworkDrive strLocalDrive, strRemoteShare, False, sUsuario , sPassword

	Set FSO = CreateObject("Scripting.FileSystemObject")
	srcPath = strLocalDrive & "\" & sFolder
	If FSO.FolderExists(srcPath) Then
		Set srcFolder = FSO.GetFolder(srcPath)
	else
		Wscript.Echo "Folder not exist"
		WScript.Quit
	End If

	Dim oldestFile,  newerFile
	For Each objFile In srcFolder.Files
		If UCase(FSO.GetExtensionName(objFile.name)) = ext Then
			'Wscript.echo ext & " Extension del fichero buscado"
			' Search Older
			If oldestFile = "" Then
				Set oldestFile = objFile
			Else
				If objFile.DateLastModified < oldestFile.DateLastModified Then
					Set oldestFile = objFile
				End If
			End If
			' Search Newer
			If newerFile = "" Then
				Set newerFile = objFile
			Else
				If objFile.DateLastModified >= newerFile.DateLastModified Then
					Set newerFile = objFile
				End If
			End If
		End If
	Next

	'Debug()

	'Get name of file if necessary
	Set objFSO2 = CreateObject("Scripting.FileSystemObject")

	If sTipo = "OH" Then 'Older Hours
		WScript.Echo(FileAgeDiff(oldestFile,DateLastModified))
	ElseIf sTipo = "ON" Then 'Older Name
		WScript.Echo(objFSO2.GetFileName(oldestFile))
	ElseIf sTipo = "OS" Then 'Older Name
		If (objFSO2.FileExists(oldestFile)) Then
			WScript.Echo CLNG(objFSO2.GetFile(oldestFile).Size / 1024 )
		else
			WScript.Echo "0"
		end if
	ElseIf sTipo = "NH" Then 'Newer Hours
		WScript.Echo(FileAgeDiff(newerFile,DateLastModified))
	ElseIf sTipo = "NN" Then 'Newer Name
		WScript.Echo(objFSO2.GetFileName(newerFile))
	ElseIf sTipo = "NS" Then 'Older Name
		If (objFSO2.FileExists(newerFile)) Then
			WScript.Echo CLNG(objFSO2.GetFile(newerFile).Size / 1024)
		else
			WScript.Echo "0"
		end if
	ElseIf sTipo = "CSV" Then 'Datos en CSV
		If (objFSO2.FileExists(newerFile)) And (objFSO2.FileExists(oldestFile))   Then
			WScript.Echo(objFSO2.GetFileName(newerFile) & ";" & newerFile.DateLastModified & ";" & FileAgeDiff(newerFile,DateLastModified) & ";" & CLNG(objFSO2.GetFile(newerFile).Size / 1024 ) & ";" & objFSO2.GetFileName(oldestFile) & ";" & oldestFile.DateLastModified & ";" & FileAgeDiff(oldestFile,DateLastModified)) & ";" & CLNG(objFSO2.GetFile(oldestFile).Size / 1024 ) 
		else
			WScript.Echo ""
		end if
	End If
End Function

Function FileAgeDiff(fName, CompareType)
	'function returns True if the file is older than the fAge (File Age) specified and false if it isn’t
	Dim LastModified, LastAccessed, FSO, DateDifference

	Set FSO = CreateObject("Scripting.FileSystemObject")
	
	If (fso.FileExists(fname)) Then
		LastModified = FSO.GetFile(fname).DateLastModified
		' WScript.Echo ("Last Modified: " & LastModified)
		LastAccessed = FSO.GetFile(fname).DateLastAccessed
		' WScript.Echo ("Last Accessed: " & LastAccessed)
	Else
		LastModified=Now()
		LastAccessed=Now()
	End If

	Select Case CompareType
	Case 1
	DateDifference = DateDiff("h",LastModified, Now())
	Case 2
	DateDifference = DateDiff("h",LastAccessed, Now())
	End Select
	' Return age of file
	FileAgeDiff = DateDifference
End Function

Function Debug()
	Set objFSO2 = CreateObject("Scripting.FileSystemObject")
	WScript.Echo ("Mas antiguo : " & oldestFile & " - " & oldestFile.DateLastModified)
	WScript.Echo ("    Last Modified: " & FileAgeDiff(oldestFile,DateLastModified))
	WScript.Echo ("    Last Accessed: " & FileAgeDiff(oldestFile,DateLastAccessed))
	Wscript.Echo "     Absolute path: " & objFSO2.GetAbsolutePathName(oldestFile)
	Wscript.Echo "     Parent folder: " & objFSO2.GetParentFolderName(oldestFile) 
	Wscript.Echo "     File name: " & objFSO2.GetFileName(oldestFile)
	Wscript.Echo "     Base name: " & objFSO2.GetBaseName(oldestFile)
	Wscript.Echo "     Extension name: " & objFSO2.GetExtensionName(oldestFile)

	WScript.Echo ("Mas moderno : " & newerFile & " - " & newerFile.DateLastModified)
	WScript.Echo ("    Last Modified: " & FileAgeDiff(newerFile,DateLastModified))
	WScript.Echo ("    Last Accessed: " & FileAgeDiff(newerFile,DateLastAccessed))
	Wscript.Echo "     Absolute path: " & objFSO2.GetAbsolutePathName(newerFile)
	Wscript.Echo "     Parent folder: " & objFSO2.GetParentFolderName(newerFile) 
	Wscript.Echo "     File name: " & objFSO2.GetFileName(newerFile)
	Wscript.Echo "     Base name: " & objFSO2.GetBaseName(newerFile)
	Wscript.Echo "     Extension name: " & objFSO2.GetExtensionName(newerFile)

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
	WScript.Echo "  ---..-----------.-----------------------.---------------------------"
End Function
'Config para Zabbix
' parametros para file_age.vbs 
' $1=Nombre carpeta
' $2=es la carpeta en la que se busca el fichero mas moderno
'    ON Older Name
'    OH Older Hours
'    OS Older Size
'    ODM Older Date Last Modified
'    ODC Older Date Created
'    ODA Older Date Last Accessed
'    NN Newer Name
'    NH Newer Hours
'    CSV Todos los datos separados por coma NN,NDateLastModified,NH


'Dim sPath
Dim drv
Dim intDiskSize, strDiskSize

Dim VERSION 
Dim INI_FILE
VERSION="1.0"
INI_FILE="c:\zabbix\scripts\ovt.ini"

'sPath = "R:" 'set drive name
nf = 0

' Revisar argumentos
set args = WScript.Arguments

select case args.Count
case 0
	Help
case 1
	Help
case 2
	' Carpeta como primer argumento y tipo de busqueda como segundo
	CalcFile args(0),args(1)
case 3
	Help
case 4
	Help
case 5
	Help
case else 
	Help
end select


Function CalcFile(sName, sTipo)
	Dim FSO
	Dim sFolder
	Const DateLastModified = 1, DateLastAccessed = 2
	
	sLetra = ReadIni( INI_FILE, sName, "LET" )
	sRuta = ReadIni( INI_FILE, sName, "FLD" )
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	srcPath = sLetra & "/" & sRuta
	If FSO.FolderExists(srcPath) Then
		Set srcFolder = FSO.GetFolder(srcPath)
	else
		Wscript.Echo "FOLDER_NOT_EXIST: " & srcPath
		WScript.Quit
	End If
	Dim oldestFile,  newerFile
	For Each objFile In srcFolder.Files
		'Wscript.Echo "Fichero: " & objFile.Name
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
		nf = nf + 1
	Next

	If nf = 0 Then
		Wscript.Echo "NO_FILES_IN_FOLDER: " & srcPath
		WScript.Quit
	End If
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
	ElseIf sTipo = "ODM" Then 'Older Date Last Modified
		WScript.Echo oldestFile.DateLastModified
	ElseIf sTipo = "ODA" Then 'Older Date Last Accessed
		WScript.Echo oldestFile.DateLastAccessed
	ElseIf sTipo = "ODC" Then 'Older Date Creation
		WScript.Echo oldestFile.Datecreated
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
	ElseIf sTipo = "NDM" Then 'newerFile Date Last Modified
		WScript.Echo newerFile.DateLastModified
	ElseIf sTipo = "NDA" Then 'newerFile Date Last Accessed
		WScript.Echo newerFile.DateLastAccessed
	ElseIf sTipo = "NDC" Then 'newerFile Date Creation
		WScript.Echo newerFile.Datecreated
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
	
    WScript.Echo "Ayuda: "
	WScript.Echo " Param1: Unidad (C:) o (E:)"
	WScript.Echo "  Letra de la unidad de la que se hacen las comprobaciones"
	WScript.Echo " Param2: ruta separada por / users/user/appdata "
	WScript.Echo "  Ruta en la que se busca el fichero"
	WScript.Echo " Param3: sTipo"
	WScript.Echo "  OH - Older Hours (Horas desde ahora del mas antiguo en la carpeta)"
	WScript.Echo "  ON - Older Name (Nombre del mas antiguo en la carpeta)"
	WScript.Echo "  OS - Older Size (Tamaño en kbytes del mas antiguo en la carpeta)"
	WScript.Echo "  ODM - Older Date Last Modified"
	WScript.Echo "  ODA - Older Date Last Accessed"
	WScript.Echo "  ODC - Older Date Creation"
	WScript.Echo "  NH - Newer Hours (Horas desde ahora del mas moderno en la carpeta)"
	WScript.Echo "  NN - Newer Name (Nombre del mas moderno en la carpeta)"
	WScript.Echo "  NS - Newer Size (Tamaño en kbytes del mas moderno en la carpeta)"
	
	WScript.Echo "  ---.-----------.------------.--------------------------."
End Function
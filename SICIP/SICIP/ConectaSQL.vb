Option Strict Off
Option Explicit On
Option Compare Text
Module mdlConectaSQL
	
	Public gstrLogin As String ' Login para entrar al Servidor
	Public gstrPassword As String ' Password para el Login
	Public gstrServidor As String ' Servidor de SQL
	Public gstrBaseDeDatos As String ' Nombre de la Base de Datos de SQL (SIM)
	Public gstrDirectorioIconos As String
	Public gstrCuenta As String
	Public gstrAplicacion As String ' Nombre de la aplicacion activa
	Public gblnFallasEnFoseo As Boolean 'Identificar por donde se realizara la consulta.
	Public gblnManejaVigilancia As Boolean
	Public gintCveBaseSuperior As Short
	
	' Variables para conectarse a Almacen
	Public gstrServidorCentral As String
	Public gstrServidorAlmacen As String ' Servidor de Almacen
	Public gstrBaseDeDatosAlmacen As String ' Nombre de la Base de Datos de SQL
	Public gblnManejaInterfaseCombustible As Boolean
	
	Public gen As RDO.rdoEnvironment
	Public gcn As Object 'Conneccion de SQL, que en su defecto seria la BD
	Public gcnVictoria As Object
	Public gcnTallerCentral As Object
	Public gcnReynosa As Object
	Public gcnTampico As Object
	Public gcnValles As Object
	Public gcnMatamoros As Object
	Public gcnMante As Object
	Public gcnSanLuis As Object
	Public gcnLUMX As Object
	Public gcnATMT As Object
	Public gcnAlmacen As Object
	Public gcnServidor As Object
	Public gstrComputerName As String
	'-----------------------------------------------------------
	'    Variables de configuración que usan los ejecutables
	'-----------------------------------------------------------
	Public gblnTransmiteKardex As Boolean
	Public gblnTransmiteLlegadas As Boolean
	Public gblnTransmiteMovtosAlmacen As Boolean
	Public gblnCentralizaODTS As Boolean
	Public gblnRegistraCausasNoRealizacion As Boolean
	Public gblnValidaLlantaTraccion As Boolean
	Public gblnValidaLlantaDireccion As Boolean
	Public gblnControlLlantasCentral As Boolean
	Public gblnActualizaServidoresRemotos As Boolean
	Public gblnActualizaKmsConOdometro As Boolean
	Public gblnInterfaseERP As Boolean
	Public gintTipoImpuesto As Short
	Public gbytCveDiaSemanaTarjeta As Byte
	
	Dim mstrDirectorioWindows As String 'Ruta de Acceso del directorio de Windows
	Dim mblnEnvOpen As Short 'Sí se efectuo la apertura del ambiente
	Public mblncn As Short 'Sí Se efectuo la apertura de la coneccion
	Public gblnConeccionVictoria As Boolean
	Public gblnConeccionReynosa As Boolean
	Public gblnConeccionTallerCentral As Boolean
	Public gblnConeccionTampico As Boolean
	Public gblnConeccionValles As Boolean
	Public gblnConeccionMatamoros As Boolean
	Public gblnConeccionMante As Boolean
	Public gblnConeccionSanLuis As Boolean
	Public gblnConeccionATMT As Boolean
	Public gblnConeccionLUMX As Boolean
	Public gblnConeccionAlmacen As Boolean
	
	Declare Function GetWindowsDirectory Lib "kernel32"  Alias "GetWindowsDirectoryA"(ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	
	'Para Localizar el Nombre de la Micro
	Declare Function GetComputerName Lib "kernel32"  Alias "GetComputerNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	' Para ubicar los INI's
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnerstring As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	
	Declare Function WNetAddConnection Lib "mpr.dll"  Alias "WNetAddConnectionA"(ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Integer
	
	Declare Function WNetCancelConnection Lib "mpr.dll"  Alias "WNetCancelConnectionA"(ByVal lpszName As String, ByVal bForce As Integer) As Integer
	
	Declare Function WNetGetConnection Lib "mpr.dll"  Alias "WNetGetConnectionA"(ByVal lpszLocalName As String, ByVal lpszRemoteName As String, ByRef cbRemoteName As Integer) As Integer
	
	Public Sub CargaParametrosConfiguracion()
		'************************************************************************
		' Rutina que realiza la carga de parámetros de configuracion del sistema
		' a variables globales.
		'************************************************************************
		
		On Error GoTo Err_CargaParametrosConfiguracion
		
		Dim strSQL As String
		Dim strmsg As String
		Dim rsParametros As RDO.rdoResultset
		
		gblnTransmiteKardex = False
		gblnTransmiteLlegadas = False
		gblnTransmiteMovtosAlmacen = False
		gblnCentralizaODTS = False
		gblnRegistraCausasNoRealizacion = False
		gblnValidaLlantaTraccion = False
		gblnValidaLlantaDireccion = False
		gblnControlLlantasCentral = False
		gblnActualizaServidoresRemotos = False
		gblnActualizaKmsConOdometro = False
		gblnInterfaseERP = False
		gblnFallasEnFoseo = False
		gblnManejaVigilancia = False
		
		strSQL = " select * from Parametros "
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsParametros = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		If rsParametros.rdoColumns("TransmiteKardex").Value = 1 Then gblnTransmiteKardex = True
		If rsParametros.rdoColumns("TransmiteLlegadas").Value = 1 Then gblnTransmiteLlegadas = True
		If rsParametros.rdoColumns("TransmiteMovtosAlmacen").Value = 1 Then gblnTransmiteMovtosAlmacen = True
		If rsParametros.rdoColumns("CentralizaODTS").Value = 1 Then gblnCentralizaODTS = True
		If rsParametros.rdoColumns("RegistraCausasNoRealizacion").Value = 1 Then gblnRegistraCausasNoRealizacion = True
		If rsParametros.rdoColumns("ValidaLlantaTraccion").Value = 1 Then gblnValidaLlantaTraccion = True
		If rsParametros.rdoColumns("ValidaLlantaDireccion").Value = 1 Then gblnValidaLlantaDireccion = True
		If rsParametros.rdoColumns("ControlLlantasCentral").Value = 1 Then gblnControlLlantasCentral = True
		If rsParametros.rdoColumns("ActualizaServidoresRemotos").Value = 1 Then gblnActualizaServidoresRemotos = True
		If rsParametros.rdoColumns("ActualizaKmsConOdometro").Value = 1 Then gblnActualizaKmsConOdometro = True
		If rsParametros.rdoColumns("InterfaseERP").Value = 1 Then gblnInterfaseERP = True
		gbytCveDiaSemanaTarjeta = rsParametros.rdoColumns("CveDiaSemanaTarjeta").Value
		
		rsParametros.Close()
		
		strSQL = " select * from Base where CveBase =" & gintCveBase
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsParametros = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(rsParametros.rdoColumns("InterfaceCombustible").Value) Then gblnManejaInterfaseCombustible = rsParametros.rdoColumns("InterfaceCombustible").Value
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(rsParametros.rdoColumns("CveTipoImpuesto").Value) Then gintTipoImpuesto = rsParametros.rdoColumns("CveTipoImpuesto").Value
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(rsParametros.rdoColumns("FallasEnFoseo").Value) Then gblnFallasEnFoseo = rsParametros.rdoColumns("FallasEnFoseo").Value
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(rsParametros.rdoColumns("ManejaVigilancia").Value) Then gblnManejaVigilancia = rsParametros.rdoColumns("ManejaVigilancia").Value
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(rsParametros.rdoColumns("ManejaVigilancia").Value) Then gintCveBaseSuperior = rsParametros.rdoColumns("CveBaseSuperior").Value
		rsParametros.Close()
		
		Exit Sub
		
Err_CargaParametrosConfiguracion: 
		strmsg = "Ocurrió un error al leer los parámetros de configuración " & Chr(13)
		strmsg = strmsg & "de inicio del sistema. La ejecución se detendrá."
		MsgBox(strmsg)
		CierraConeccion()
		End
		Resume Next
	End Sub
	
	Public Sub ConectaUnidad(ByVal vstrServidor As String, ByVal vstrDirectorio As String, ByVal vstrUnidad As String)
		
		Const ERROR_ACCESS_DENIED As Short = 5
		Const ERROR_ALREADY_ASSIGNED As Short = 85
		Const ERROR_BAD_DEV_TYPE As Short = 66
		Const ERROR_BAD_DEVICE As Short = 1200
		Const ERROR_BAD_NET_NAME As Short = 67
		Const ERROR_BAD_PROFILE As Short = 1206
		Const ERROR_CANNOT_OPEN_PROFILE As Short = 1205
		Const ERROR_DEVICE_ALREADY_REMEMBERED As Short = 1202
		Const ERROR_EXTENDED_ERROR As Short = 1208
		Const ERROR_INVALID_PASSWORD As Short = 86
		Const ERROR_NO_NET_OR_BAD_PATH As Short = 1203
		Const ERROR_NO_NETWORK As Short = 1222
		Const ERROR_MORE_DATA As Short = 234
		Const ERROR_NOT_CONNECTED As Short = 2250
		
		Dim intPaso As Short
		Dim strRutaActual As String
		Dim strRutaNueva As String
		Dim strMensajeError As String
		
		strRutaNueva = "\\" & vstrServidor & "\" & vstrDirectorio
		strRutaActual = New String(Chr(0), 250)
		
		intPaso = WNetGetConnection(vstrUnidad, strRutaActual, 250)
		If ERROR_NOT_CONNECTED = intPaso Then
			intPaso = WNetAddConnection(strRutaNueva, "", vstrUnidad)
		Else
			If UCase(Mid(strRutaActual, 1, InStr(strRutaActual, Chr(0)) - 1)) <> strRutaNueva Then
				intPaso = WNetCancelConnection(vstrUnidad, True)
				intPaso = WNetAddConnection(strRutaNueva, "", vstrUnidad)
			End If
		End If
		
		Select Case intPaso
			Case ERROR_ACCESS_DENIED
				strMensajeError = "Access is denied."
			Case ERROR_ALREADY_ASSIGNED
				strMensajeError = "The device specified in the " & vstrUnidad & " parameter is already connected."
			Case ERROR_BAD_DEV_TYPE
				strMensajeError = "The device type and the resource type do not match."
			Case ERROR_BAD_DEVICE
				strMensajeError = "The specified device name " & vstrUnidad & " is invalid."
			Case ERROR_BAD_NET_NAME
				strMensajeError = "The value specified in the " & strRutaNueva & " parameter is not valid or cannot be located."
			Case ERROR_BAD_PROFILE
				strMensajeError = "The user profile is in an incorrect format."
			Case ERROR_CANNOT_OPEN_PROFILE
				strMensajeError = "The system is unable to open the user profile to process persistent connections."
			Case ERROR_DEVICE_ALREADY_REMEMBERED
				strMensajeError = "An entry for the device specified in " & vstrUnidad & " is already in the user profile."
			Case ERROR_EXTENDED_ERROR
				strMensajeError = "A network-specific error occurred. To get a description of the error, use the WNetGetLastError function."
			Case ERROR_INVALID_PASSWORD
				strMensajeError = "The specified password is invalid."
			Case ERROR_NO_NET_OR_BAD_PATH
				strMensajeError = "The operation cannot be performed because either a network component is not started or the specified name cannot be used."
			Case ERROR_NO_NETWORK
				strMensajeError = "The network is not present."
			Case ERROR_MORE_DATA
				strMensajeError = "More data is available."
			Case ERROR_NOT_CONNECTED
				strMensajeError = "This network connection does not exist."
			Case 0
				strMensajeError = ""
			Case Else
				strMensajeError = ""
		End Select
		
		If strMensajeError <> "" Then
			MsgBox("Ocurrio un Error al Conectar la Unidad " & vstrUnidad & Chr(System.Windows.Forms.Keys.Return) & intPaso & strMensajeError, MsgBoxStyle.Critical, "ConectaUnidad")
			End
		End If
		
	End Sub
	
	Function MsgError(ByRef numError As Short) As String
		Select Case numError
			Case 3
				MsgError = "Return without GoSub"
			Case 5
				MsgError = "Illegal function call"
			Case 6
				MsgError = "overflow"
			Case 7
				MsgError = "Out of memory"
			Case 9
				MsgError = "Subscript out of range"
			Case 10
				MsgError = "Duplicate definition"
			Case 11
				MsgError = "Division by zero"
			Case 13
				MsgError = "Type mismatch"
			Case 14
				MsgError = "Out of string space"
			Case 16
				MsgError = "String formula too complex"
			Case 17
				MsgError = "Can't continue"
			Case 19
				MsgError = "No Resume"
			Case 20
				MsgError = "Resume without error"
			Case 28
				MsgError = "Out of stack space"
			Case 35
				MsgError = "Sub or Function not defined"
			Case 48
				MsgError = "Error in loading DLL"
			Case 49
				MsgError = "Bad DLL calling convention"
			Case 51
				MsgError = "Internal error"
			Case 52
				MsgError = "Bad file name or number"
			Case 53
				MsgError = "File not found"
			Case 54
				MsgError = "Bad file mode"
			Case 55
				MsgError = "File already open"
			Case 57
				MsgError = "Device I/O error"
			Case 58
				MsgError = "File already exists"
			Case 59
				MsgError = "Bad record length"
			Case 61
				MsgError = "Disk full"
			Case 62
				MsgError = "Input past end of file"
			Case 63
				MsgError = "Bad record number"
			Case 64
				MsgError = "Bad file name"
			Case 67
				MsgError = "Too many files"
			Case 68
				MsgError = "El dispositivo no esta disponible"
			Case 70
				MsgError = "Permission denied"
			Case 71
				MsgError = "Disk not ready, Inserte un disquette en el dispositivo"
			Case 74
				MsgError = "Can't rename with different drive"
			Case 75
				MsgError = "Path/File access error"
			Case 76
				MsgError = "Path not found"
			Case 91
				MsgError = "Object variable not Set"
			Case 92
				MsgError = "For loop not initialized"
			Case 93
				MsgError = "Invalid pattern string"
			Case 94
				MsgError = "Invalid use of Null"
			Case 95
				MsgError = "Cannot destroy active form instance"
			Case 260
				MsgError = "No timer available"
			Case 280
				MsgError = "DDE channel not fully closed; awaiting response from foreign application"
			Case 281
				MsgError = "No More DDE channels"
			Case 282
				MsgError = "No foreign application responded to a DDE initiate"
			Case 283
				MsgError = "Multiple applications responded to a DDE initiate"
			Case 284
				MsgError = "DDE channel locked"
			Case 285
				MsgError = "Foreign application won't perform DDE method or operation"
			Case 286
				MsgError = "Timeout while waiting for DDE response"
			Case 287
				MsgError = "User pressed Escape key during DDE operation"
			Case 288
				MsgError = "destination Is busy"
			Case 289
				MsgError = "Data not provided in DDE operation"
			Case 290
				MsgError = "Data in wrong format"
			Case 291
				MsgError = "Foreign application quit"
			Case 292
				MsgError = "DDE conversation closed or changed"
			Case 293
				MsgError = "DDE Method invoked with no channel open"
			Case 294
				MsgError = "Invalid DDE Link format"
			Case 295
				MsgError = "Message queue filled; DDE message lost"
			Case 296
				MsgError = "PasteLink already performed on this control"
			Case 297
				MsgError = "Can't set LinkMode; invalid LinkTopic"
			Case 298
				MsgError = "DDE requires ddeml.dll"
			Case 320
				MsgError = "Can't use character device names in file names: ' '"
			Case 321
				MsgError = "Invalid file format"
			Case 340
				MsgError = "Control array element ' ' doesn't exist"
			Case 341
				MsgError = "Invalid control array index"
			Case 342
				MsgError = "Not enough room to allocate control array ' '"
			Case 343
				MsgError = "Object not an array"
			Case 344
				MsgError = "Must specify index for object array"
			Case 345
				MsgError = "Reached limit: cannot create any more controls for this form"
			Case 360
				MsgError = "Object already loaded"
			Case 361
				MsgError = "Can't load or unload this object"
			Case 362
				MsgError = "Can't unload controls created at design time"
			Case 363
				MsgError = "Custom control ' ' not found"
			Case 364
				MsgError = "Object was unloaded"
			Case 365
				MsgError = "Unable to unload within this context"
			Case 366
				MsgError = "No MDI Form available to load"
			Case 380
				MsgError = "Invalid property value"
			Case 381
				MsgError = "Invalid property array index"
			Case 382
				MsgError = "' ' property cannot be set at run time"
			Case 383
				MsgError = "' ' property is read-only"
			Case 384
				MsgError = "A form can't be moved or sized while minimized or maximized"
			Case 385
				MsgError = "Must specify index when using property array"
			Case 386
				MsgError = "' ' property not available at run time"
			Case 387
				MsgError = "' ' property can't be set on this control"
			Case 388
				MsgError = "Can't set Visible property from a parent menu"
			Case 389
				MsgError = "Invalid key"
			Case 390
				MsgError = "No Defined Value"
			Case 391
				MsgError = "Name not available"
			Case 392
				MsgError = "MDI child forms cannot be hidden"
			Case 393
				MsgError = "' ' property cannot be read at run time"
			Case 394
				MsgError = "' ' property is write-only"
			Case 395
				MsgError = "Can't use separator bar as menu name"
			Case 400
				MsgError = "Form already displayed; can't show modally"
			Case 401
				MsgError = "Can't show non-modal form when modal form is displayed"
			Case 402
				MsgError = "Must close or hide topmost modal form first"
			Case 403
				MsgError = "MDI forms cannot be shown modally"
			Case 404
				MsgError = "MDI child forms cannot be shown modally"
			Case 420
				MsgError = "Invalid object reference"
			Case 421
				MsgError = "Method not applicable for this object"
			Case 422
				MsgError = "property ' ' not found"
			Case 423
				MsgError = "property Or control ' ' not found"
			Case 424
				MsgError = "Object required"
			Case 425
				MsgError = "Invalid object use"
			Case 426
				MsgError = "Only one MDI Form allowed"
			Case 427
				MsgError = "Invalid object type; Menu control required"
			Case 428
				MsgError = "Popup menu must have at least one submenu"
			Case 429
				MsgError = "OLE Automation server cannot create object"
			Case 430
				MsgError = "Class does not support OLE Automation"
			Case 431
				MsgError = "OLE Automation server cannot load file"
			Case 432
				MsgError = "OLE Automation file or object name syntax error"
			Case 433
				MsgError = "OLE Automation object does not exist"
			Case 434
				MsgError = "Access to OLE Automation object denied"
			Case 435
				MsgError = "OLE initialization error"
			Case 436
				MsgError = "OLE Automation method returned unsupported type"
			Case 437
				MsgError = "OLE Automation method did not return a value"
			Case 438
				MsgError = "OLE Automation no such property or method"
			Case 439
				MsgError = "OLE Automation argument type mismatch"
			Case 440
				MsgError = "OLE Automation error."
			Case 441
				MsgError = "Error loading VBOA300.DLL"
			Case 442
				MsgError = "OLE Automation Lbound or Ubound on non Array value"
			Case 443
				MsgError = "OLE Automation Object does not have a default value"
			Case 444
				MsgError = "Method not applicable in this context"
			Case 460
				MsgError = "Invalid Clipboard format"
			Case 461
				MsgError = "Specified format doesn't match format of data"
			Case 480
				MsgError = "Can't create AutoRedraw image"
			Case 481
				MsgError = "Invalid picture"
			Case 482
				MsgError = "Printer error"
			Case 520
				MsgError = "Can't empty Clipboard"
			Case 521
				MsgError = "Can't open Clipboard"
			Case 600
				MsgError = "Set value not allowed on collections"
			Case 601
				MsgError = "Get value not allowed on collections"
			Case 602
				MsgError = "General ODBC error: ' '"
			Case 603
				MsgError = "ODBC - SQLAllocEnv failure"
			Case 604
				MsgError = "ODBC - SQLAllocConnect failure"
			Case 605
				MsgError = "OpenDatabase - invalid connect string"
			Case 606
				MsgError = "ODBC - SQLConnect failure ' '"
			Case 607
				MsgError = "Access attempted on unopened DataBase"
			Case 608
				MsgError = "ODBC - SQLFreeConnect error"
			Case 609
				MsgError = "ODBC - GetDriverFunctions failure"
			Case 610
				MsgError = "ODBC - SQLAllocStmt failure"
			Case 611
				MsgError = "ODBC - SQLTables (TableDefs.Refresh) failure: ' '"
			Case 612
				MsgError = "ODBC - SQLBindCol failure"
			Case 613
				MsgError = "ODBC - SQLFetch failure: ' '"
			Case 614
				MsgError = "ODBC - SQLColumns (Fielrs.Refresh) failure: ' '"
			Case 615
				MsgError = "ODBC - SQLStatistics (Indexes.Refresh) failure: ' '"
			Case 616
				MsgError = "Table exists - append not allowed"
			Case 617
				MsgError = "No fielrs defined - cannot append table"
			Case 618
				MsgError = "ODBC - SQLNumResultCols (CreaterdoResultset) failure: ' '"
			Case 619
				MsgError = "ODBC - SQLDescibeCol (CreaterdoResultset) failure: ' '"
			Case 620
				MsgError = "rdoResultset is open - CreaterdoResultset method not allowed"
			Case 621
				MsgError = "Row-returning SQL is illegal in ExecuteSQL method"
			Case 622
				MsgError = "CommitTrans/Rollback illegal - Transactions not support"
			Case 623
				MsgError = "Name not found in this collection"
			Case 624
				MsgError = "Unable to Build Data Type Table"
			Case 625
				MsgError = "Data type of field ' ' not supported by target database"
			Case 626
				MsgError = "Attempt to Move past EOF"
			Case 627
				MsgError = "rdoResultset is not updatable or Edit method has not been invoked"
			Case 628
				MsgError = "' ' rdoResultset method illegal - no scrollable cursor support"
			Case 629
				MsgError = "Warning:   (ODBC - SQLSetConnectOption failure)"
			Case 630
				MsgError = "Property is read-only"
			Case 631
				MsgError = "Zero rows affected by Update method"
			Case 632
				MsgError = "Update illegal without previous Edit or AddNew method"
			Case 633
				MsgError = "Append illegal - Field is part of a TableDefs collection"
			Case 634
				MsgError = "Property value only valid when Field is part of a rdoResultset"
			Case 635
				MsgError = "Cannot set the property of an object which is part of a Database object"
			Case 636
				MsgError = "Set field value illegal without previous Edit or AddNew method"
			Case 637
				MsgError = "Append illegal - Index is part of a TableDefs collection"
			Case 638
				MsgError = "Access attempted on unopened rdoResultset"
			Case 639
				MsgError = "Field type is illegal"
			Case 640
				MsgError = "Field size illegal for specified Field Type"
			Case 641
				MsgError = "illegal - no current record"
			Case 642
				MsgError = "Reserved parameter must be FALSE"
			Case 643
				MsgError = "Property Not Found"
			Case 644
				MsgError = "ODBC - SQLConfigDataSource error ' '"
			Case 645
				MsgError = "ODBC Driver does not support exclusive access to rdoResultsets"
			Case 646
				MsgError = "GetChunk: Offset/Size argument combination illegal"
			Case 647
				MsgError = "Delete method requires a name argument"
			Case 648
				MsgError = "Data access objects require VBDB300.DLL"
			Case 2420
				MsgError = "Syntax error in number"
			Case 2421
				MsgError = "Syntax error in date"
			Case 2422
				MsgError = "Syntax error in string"
			Case 2423
				MsgError = "Invalid use of '.', '!', or '()'."
			Case 2424
				MsgError = "Unknown name"
			Case 2425
				MsgError = "Unknown function name"
			Case 2426
				MsgError = "Function isn't available in expressions"
			Case 2427
				MsgError = "Object has no value"
			Case 2428
				MsgError = "Invalid arguments used with domain function"
			Case 2429
				MsgError = "In operator without ()"
			Case 2430
				MsgError = "Between operator without And"
			Case 2431
				MsgError = "Syntax error"
			Case 2432
				MsgError = "Syntax error"
			Case 2433
				MsgError = "Syntax error"
			Case 2434
				MsgError = "Syntax error"
			Case 2435
				MsgError = "Extra )"
			Case 2436
				MsgError = "Missing ), ], or"
			Case 2437
				MsgError = "Invalid use of vertical bars"
			Case 2438
				MsgError = "Syntax error"
			Case 2439
				MsgError = "Wrong number of arguments used with function"
			Case 2440
				MsgError = "IIF function without ()"
			Case 2442
				MsgError = "Invalid use of parentheses"
			Case 2443
				MsgError = "Invalid use of Is operator"
			Case 2445
				MsgError = "Expression too complex"
			Case 2446
				MsgError = "Out of memory during calculation"
			Case 2447
				MsgError = "Invalid use of '.', '!', or '()'."
			Case 2448
				MsgError = "Can't set value."
			Case 2449
				MsgError = "Invalid method in expression."
			Case 2450
				MsgError = "Invalid reference to form ' '."
			Case 2451
				MsgError = "Invalid reference to report ' '."
			Case 2452
				MsgError = "Invalid reference to Parent property."
			Case 2453
				MsgError = "Invalid reference to control ' '."
			Case 2454
				MsgError = "Invalid reference to '! '."
			Case 2455
				MsgError = "Invalid reference to property ' '."
			Case 2456
				MsgError = "Invalid form number reference."
			Case 2457
				MsgError = "Invalid report number reference."
			Case 2458
				MsgError = "Invalid control number reference."
			Case 2459
				MsgError = "Can't refer to Parent property in Design view."
			Case 2460
				MsgError = "Can't refer to rdoResultset property in Design view."
			Case 2461
				MsgError = "Invalid section reference."
			Case 2462
				MsgError = "Invalid section number reference."
			Case 2463
				MsgError = "Invalid group level reference."
			Case 2464
				MsgError = "Invalid group level number reference."
			Case 2465
				MsgError = "Invalid reference to field ' '."
			Case 2466
				MsgError = "Invalid reference to rdoResultset property."
			Case 2467
				MsgError = "Object referred to in expression no longer exists."
			Case 2468
				MsgError = "Invalid argument used with DatePart, DateAdd or DateDiff function."
			Case 2469
				MsgError = "1 in validation rule: '|2'."
			Case 2470
				MsgError = "in validation rule."
			Case 2471
				MsgError = "in query."
			Case 2472
				MsgError = "in linked master field."
			Case 2473
				MsgError = "1 in '|2' expression."
			Case 2474
				MsgError = "No control is active."
			Case 2475
				MsgError = "No form is active."
			Case 2476
				MsgError = "No report is active."
			Case 2477
				MsgError = "Invalid subclass ' ' referred to in TypeOf function."
			Case 3000
				MsgError = "Reserved error ( ); there is no message for this error."
			Case 3001
				MsgError = "Invalid argument."
			Case 3002
				MsgError = "Couldn't start session."
			Case 3003
				MsgError = "Couldn't start transaction; too many transactions already nested."
			Case 3004
				MsgError = "Couldn't find database ' '."
			Case 3005
				MsgError = "' ' isn't a valid database name."
			Case 3006
				MsgError = "Database ' ' is exclusively locked."
			Case 3007
				MsgError = "Couldn't open database ' '."
			Case 3008
				MsgError = "TABLE ' ' is exclusively locked."
			Case 3009
				MsgError = "Couldn't lock table ' '; currently in use."
			Case 3010
				MsgError = "TABLE ' ' already exists."
			Case 3011
				MsgError = "Couldn't find object ' '."
			Case 3012
				MsgError = "object ' ' already exists."
			Case 3013
				MsgError = "Couldn't rename installable ISAM file."
			Case 3014
				MsgError = "Can't open any more tables."
			Case 3015
				MsgError = "' ' isn't an index in this table."
			Case 3016
				MsgError = "Field won't fit in record."
			Case 3017
				MsgError = "Field length is too long."
			Case 3018
				MsgError = "Couldn't find field ' '."
			Case 3019
				MsgError = "Operation invalid without a current index."
			Case 3020
				MsgError = "Update without AddNew or Edit."
			Case 3021
				MsgError = "No current record."
			Case 3022
				MsgError = "Can't have duplicate key; index changes were unsuccessful."
			Case 3023
				MsgError = "AddNew or Edit already used."
			Case 3024
				MsgError = "Couldn't find file ' '."
			Case 3025
				MsgError = "Can't open any more files."
			Case 3026
				MsgError = "Not enough space on disk."
			Case 3027
				MsgError = "Couldn't update; database is read-only."
			Case 3028
				MsgError = "Couldn't initialize data access because file 'SYSTEM.MDA' couldn't be opened."
			Case 3029
				MsgError = "Not a valid account name or password."
			Case 3030
				MsgError = "' ' isn't a valid account name."
			Case 3031
				MsgError = "Not a valid password."
			Case 3032
				MsgError = "Can't delete account."
			Case 3033
				MsgError = "No permission for ' '."
			Case 3034
				MsgError = "Commit or Rollback without BeginTrans."
			Case 3035
				MsgError = "Out of memory."
			Case 3036
				MsgError = "Database has reached maximum size."
			Case 3037
				MsgError = "Can't open any more tables or queries."
			Case 3038
				MsgError = "Out of memory."
			Case 3039
				MsgError = "Couldn't create index; too many indexes already defined."
			Case 3040
				MsgError = "Disk I/O error during read."
			Case 3041
				MsgError = "Incompatible database version."
			Case 3042
				MsgError = "Out of MS-DOS file handles."
			Case 3043
				MsgError = "Disk or network error."
			Case 3044
				MsgError = "' ' isn't a valid path."
			Case 3045
				MsgError = "Couldn't use ' '; file already in use."
			Case 3046
				MsgError = "Couldn't save; currently locked by another user."
			Case 3047
				MsgError = "Record is too large."
			Case 3048
				MsgError = "Can't open any more databases."
			Case 3049
				MsgError = "' ' is corrupted or isn't a Microsoft Access database."
			Case 3050
				MsgError = "Couldn't lock file; SHARE.EXE hasn't been loaded."
			Case 3051
				MsgError = "Couldn't open file ' '."
			Case 3052
				MsgError = "MS-DOS file sharing lock count exceeded.  You need to increase the number of locks installed with SHARE.EXE."
			Case 3053
				MsgError = "Too many client tasks."
			Case 3054
				MsgError = "Too many Memo or Long Binary fielrs."
			Case 3055
				MsgError = "Not a valid file name."
			Case 3056
				MsgError = "Couldn't repair this database."
			Case 3057
				MsgError = "Operation not supported on attached tables."
			Case 3058
				MsgError = "Can't have Null value in index."
			Case 3059
				MsgError = "Operation canceled by user."
			Case 3060
				MsgError = "Wrong data type for parameter ' '."
			Case 3061
				MsgError = "1 parameters were expected, but only |2 were supplied."
			Case 3062
				MsgError = "Duplicate output alias ' '."
			Case 3063
				MsgError = "Duplicate output destination ' '."
			Case 3064
				MsgError = "Can't open action query ' '."
			Case 3065
				MsgError = "Can't execute a non-action query."
			Case 3066
				MsgError = "Query must have at least one output field."
			Case 3067
				MsgError = "Query input must contain at least one table or query."
			Case 3068
				MsgError = "Not a valid alias name."
			Case 3069
				MsgError = "Can't have action query ' ' as an input."
			Case 3070
				MsgError = "Can't bind name ' '."
			Case 3071
				MsgError = "Can't evaluate expression."
			Case 3073
				MsgError = "Operation must use an updatable query."
			Case 3074
				MsgError = "Can't repeat table name ' ' in from clause."
			Case 3075
				MsgError = "1 in query expression '|2'."
			Case 3076
				MsgError = "in criteria expression."
			Case 3077
				MsgError = "in expression."
			Case 3078
				MsgError = "Couldn't find input table or query ' '."
			Case 3079
				MsgError = "Ambiguous field reference ' '."
			Case 3080
				MsgError = "Joined table ' ' not listed in from clause."
			Case 3081
				MsgError = "Can't join more than one table with the same name ( )."
			Case 3082
				MsgError = "JOIN operation ' ' refers to a non-joined table."
			Case 3083
				MsgError = "Can't use internal report query."
			Case 3084
				MsgError = "Can't insert into action query."
			Case 3085
				MsgError = "Undefined function ' ' in expression."
			Case 3086
				MsgError = "Couldn't delete from specified tables."
			Case 3087
				MsgError = "Too many expressions in GROUP BY clause."
			Case 3088
				MsgError = "Too many expressions in order BY clause."
			Case 3089
				MsgError = "Too many expressions in DISTINCT output."
			Case 3090
				MsgError = "Resultant table may not have more than one Counter field."
			Case 3091
				MsgError = "HAVING clause ( ) without grouping or aggregation."
			Case 3092
				MsgError = "Can't use HAVING clause in TRANSFORM statement."
			Case 3093
				MsgError = "order BY clause ( ) conflicts with DISTINCT."
			Case 3094
				MsgError = "order BY clause ( ) conflicts with GROUP BY clause."
			Case 3095
				MsgError = "Can't have aggregate function in expression ( )."
			Case 3096
				MsgError = "Can't have aggregate function in where clause ( )."
			Case 3097
				MsgError = "Can't have aggregate function in order BY clause ( )."
			Case 3098
				MsgError = "Can't have aggregate function in GROUP BY clause ( )."
			Case 3099
				MsgError = "Can't have aggregate function in JOIN operation ( )."
			Case 3100
				MsgError = "Can't set field ' ' in join key to Null."
			Case 3101
				MsgError = "Join is broken by value(s) in fielrs ' '."
			Case 3102
				MsgError = "Circular reference caused by ' '."
			Case 3103
				MsgError = "Circular reference caused by alias ' ' in query definition's select list."
			Case 3104
				MsgError = "Can't specify Fixed Column Heading ' ' in a crosstab query more than once."
			Case 3105
				MsgError = "Missing destination field name in select INTO statement ( )."
			Case 3106
				MsgError = "Missing destination field name in UPDATE statement ( )."
			Case 3107
				MsgError = "Couldn't insert; no insert permission for table or query ' '."
			Case 3108
				MsgError = "Couldn't replace; no replace permission for table or query ' '."
			Case 3109
				MsgError = "Couldn't delete; no delete permission for table or query ' '."
			Case 3110
				MsgError = "Couldn't read definitions; no read definitions permission for table or query ' '."
			Case 3111
				MsgError = "Couldn't create; no create permission for table or query ' '."
			Case 3112
				MsgError = "Couldn't read; no read permission for table or query ' '."
			Case 3113
				MsgError = "Can't update ' '; field not updatable."
			Case 3114
				MsgError = "Can't include Memo or Long Binary when you select unique values ( )."
			Case 3115
				MsgError = "Can't have Memo or Long Binary in aggregate argument ( )."
			Case 3116
				MsgError = "Can't have Memo or Long Binary in criteria ( ) for aggregate function."
			Case 3117
				MsgError = "Can't sort on Memo or Long Binary ( )."
			Case 3118
				MsgError = "Can't join on Memo or Long Binary ( )."
			Case 3119
				MsgError = "Can't group on Memo or Long Binary ( )."
			Case 3120
				MsgError = "Can't group on fielrs selected with '*' ( )."
			Case 3121
				MsgError = "Can't group on fielrs selected with '*'."
			Case 3122
				MsgError = "' ' not part of aggregate function or grouping."
			Case 3123
				MsgError = "Can't use '*' in crosstab query."
			Case 3124
				MsgError = "Can't input from internal report query ( )."
			Case 3125
				MsgError = "' ' isn't a valid name."
			Case 3126
				MsgError = "Invalid bracketing of name ' '."
			Case 3127
				MsgError = "INSERT INTO statement contains unknown field name ' '."
			Case 3128
				MsgError = "Must specify tables to delete from."
			Case 3129
				MsgError = "Invalid SQL statement; expected 'DELETE', 'INSERT', 'PROCEDURE', 'select', or 'UPDATE'."
			Case 3130
				MsgError = "Syntax error in DELETE statement."
			Case 3131
				MsgError = "Syntax error in from clause."
			Case 3132
				MsgError = "Syntax error in GROUP BY clause."
			Case 3133
				MsgError = "Syntax error in HAVING clause."
			Case 3134
				MsgError = "Syntax error in INSERT statement."
			Case 3135
				MsgError = "Syntax error in JOIN operation."
			Case 3136
				MsgError = "Syntax error in LEVEL clause."
			Case 3137
				MsgError = "Missing semicolon (;) at end of SQL statement."
			Case 3138
				MsgError = "Syntax error in order BY clause."
			Case 3139
				MsgError = "Syntax error in PARAMETER clause."
			Case 3140
				MsgError = "Syntax error in PROCEDURE clause."
			Case 3141
				MsgError = "Syntax error in select statement."
			Case 3142
				MsgError = "Characters found after end of SQL statement."
			Case 3143
				MsgError = "Syntax error in TRANSFORM statement."
			Case 3144
				MsgError = "Syntax error in UPDATE statement."
			Case 3145
				MsgError = "Syntax error in where clause."
			Case 3146
				MsgError = "ODBC--call failed."
			Case 3147
				MsgError = "ODBC--data buffer overflow."
			Case 3148
				MsgError = "ODBC--connection failed."
			Case 3149
				MsgError = "ODBC--incorrect DLL."
			Case 3150
				MsgError = "ODBC--missing DLL."
			Case 3151
				MsgError = "ODBC--connection to ' ' failed."
			Case 3152
				MsgError = "ODBC--incorrect driver version ' 1'; expected version '|2'."
			Case 3153
				MsgError = "ODBC--incorrect server version ' 1'; expected version '|2'."
			Case 3154
				MsgError = "ODBC - -Couldn't find DLL ' '."
			Case 3155
				MsgError = "ODBC--insert failed."
			Case 3156
				MsgError = "ODBC--delete failed."
			Case 3157
				MsgError = "ODBC--update failed."
			Case 3158
				MsgError = "Couldn't save record; currently locked by another user."
			Case 3159
				MsgError = "Not a valid bookmark."
			Case 3160
				MsgError = "Table isn't open."
			Case 3161
				MsgError = "Couldn't decrypt file."
			Case 3162
				MsgError = "Null is invalid."
			Case 3163
				MsgError = "Couldn't insert or paste; data too long for field."
			Case 3164
				MsgError = "Couldn't update field."
			Case 3165
				MsgError = "Couldn't open .INF file."
			Case 3166
				MsgError = "Missing memo file."
			Case 3167
				MsgError = "Record is deleted."
			Case 3168
				MsgError = "Invalid .INF file."
			Case 3169
				MsgError = "Illegal type in expression."
			Case 3170
				MsgError = "Couldn't find installable ISAM."
			Case 3171
				MsgError = "Couldn't find net path or user name."
			Case 3172
				MsgError = "Couldn't open PARADOX.NET."
			Case 3173
				MsgError = "Couldn't open table 'MSysAccounts' in SYSTEM.MDA."
			Case 3174
				MsgError = "Couldn't open table 'MSysGroups' in SYSTEM.MDA."
			Case 3175
				MsgError = "Date is out of range or is in an invalid format."
			Case 3176
				MsgError = "Couldn't open file ' '."
			Case 3177
				MsgError = "Not a valid table name."
			Case 3178
				MsgError = "Out of memory."
			Case 3179
				MsgError = "Encountered unexpected end of file."
			Case 3180
				MsgError = "Couldn't write to file ' '."
			Case 3181
				MsgError = "Invalid range."
			Case 3182
				MsgError = "Invalid file format."
			Case 3183
				MsgError = "Not enough space on temporary disk."
			Case 3184
				MsgError = "Couldn't execute query; couldn't find linked table."
			Case 3185
				MsgError = "select INTO remote database tried to produce too many fielrs."
			Case 3186
				MsgError = "Couldn't save; currently locked by user ' 2' on machine '|1'."
			Case 3187
				MsgError = "Couldn't read; currently locked by user ' 2' on machine '|1'."
			Case 3188
				MsgError = "Couldn't update; currently locked by another session on this machine."
			Case 3189
				MsgError = "TABLE ' 1' is exclusively locked by user '|3' on machine '|2'."
			Case 3190
				MsgError = "Too many fielrs defined."
			Case 3191
				MsgError = "Can't define field more than once."
			Case 3192
				MsgError = "Couldn't find output table ' '."
			Case 3193
				MsgError = "(unknown)"
			Case 3194
				MsgError = "(unknown)"
			Case 3195
				MsgError = "(expression)"
			Case 3196
				MsgError = "Couldn't use ' '; database already in use."
			Case 3197
				MsgError = "Data has changed; operation stopped."
			Case 3198
				MsgError = "Couldn't start session.  Too many sessions already active."
			Case 3199
				MsgError = "Couldn't find reference."
			Case 3200
				MsgError = "Can't delete or change record. Since related recorrs exist in table ' ', referential integrity rules would be violated."
			Case 3201
				MsgError = "Can't add or change record.  Referential integrity rules require a related record in table ' '."
			Case 3202
				MsgError = "Couldn't save; currently locked by another user."
			Case 3203
				MsgError = "Can't specify subquery in expression ( )."
			Case 3204
				MsgError = "Database already exists."
			Case 3205
				MsgError = "Too many crosstab column headers ( )."
			Case 3206
				MsgError = "Can't create a relationship between a field and itself."
			Case 3207
				MsgError = "Operation not supported on Paradox table with no primary key."
			Case 3208
				MsgError = "Invalid Deleted entry in [dBASE ISAM] section in INI file."
			Case 3209
				MsgError = "Invalid Stats entry in [dBASE ISAM] section in INI file."
			Case 3210
				MsgError = "Connect string too long."
			Case 3211
				MsgError = "Couldn't lock table ' '; currently in use."
			Case 3212
				MsgError = "Couldn't lock table ' 1'; currently in use by user '|3' on machine '|2'."
			Case 3213
				MsgError = "Invalid Date entry in [dBASE ISAM] section in INI file."
			Case 3214
				MsgError = "Invalid Mark entry in [dBASE ISAM] section in INI file."
			Case 3215
				MsgError = "Too many Btrieve tasks."
			Case 3216
				MsgError = "Parameter ' ' specified where a table name is required."
			Case 3217
				MsgError = "Parameter ' ' specified where a database name is required."
			Case 3218
				MsgError = "Couldn't update; currently locked."
			Case 3219
				MsgError = "Can't perform operation; it is illegal."
			Case 3220
				MsgError = "Wrong Paradox sort sequence."
			Case 3221
				MsgError = "Invalid entries in [Btrieve ISAM] section in WIN.INI."
			Case 3222
				MsgError = "Query can't contain a Database parameter."
			Case 3223
				MsgError = "' ' isn't a valid parameter name."
			Case 3224
				MsgError = "Btrieve--data dictionary is corrupted."
			Case 3225
				MsgError = "Encountered record locking deadlock while performing Btrieve operation."
			Case 3226
				MsgError = "Errors encountered while using the Btrieve DLL."
			Case 3227
				MsgError = "Invalid Century entry in [dBASE ISAM] section in INI file."
			Case 3228
				MsgError = "Invalid CollatingSequence entry in [Paradox ISAM] section in INI file."
			Case 3229
				MsgError = "Btrieve - -Can't change field."
			Case 3230
				MsgError = "Out-of-date Paradox lock file."
			Case 3231
				MsgError = "ODBC--field would be too long; data truncated."
			Case 3232
				MsgError = "ODBC - -Couldn't create table."
			Case 3233
				MsgError = "ODBC--incorrect driver version."
			Case 3234
				MsgError = "ODBC--remote query timeout expired."
			Case 3235
				MsgError = "ODBC--data type not supported on server."
			Case 3236
				MsgError = "ODBC--encountered unexpected Null value."
			Case 3237
				MsgError = "ODBC--unexpected type."
			Case 3238
				MsgError = "ODBC--data out of range."
			Case 3239
				MsgError = "Too many active users."
			Case 3240
				MsgError = "Btrieve--missing WBTRCALL.DLL."
			Case 3241
				MsgError = "Btrieve--out of resources."
			Case 3242
				MsgError = "Invalid reference in select statement."
			Case 3243
				MsgError = "None of the import field names match fielrs in the appended table."
			Case 3244
				MsgError = "Can't import password-protected sprearsheet."
			Case 3245
				MsgError = "Couldn't parse field names from first row of import table."
			Case 3246
				MsgError = "Operation not supported in transactions."
			Case 3247
				MsgError = "ODBC--linked table definition has changed."
			Case 3248
				MsgError = "Invalid NetworkAccess entry in INI file."
			Case 3249
				MsgError = "Invalid PageTimeout entry in INI file."
			Case 3250
				MsgError = "Couldn't build key."
			Case 3251
				MsgError = "Feature not available."
			Case 3252
				MsgError = "Illegal reentrancy during query execution."
			Case 3254
				MsgError = "ODBC - -Can't lock all recorrs."
			Case 3255
				MsgError = "ODBC - -Can't change connect string parameter."
			Case 3256
				MsgError = "Index file not found."
			Case 3257
				MsgError = "Syntax error in WITH OWNERACCESS OPTION declaration."
			Case 3258
				MsgError = "Query contains ambiguous (outer) joins."
			Case 3259
				MsgError = "Invalid field data type."
			Case 3260
				MsgError = "Couldn't update; currently locked by user ' 2' on machine '|1'."
			Case 3263
				MsgError = "Invalid database object."
			Case 3264
				MsgError = "No fielrs defined - cannot append table."
			Case 3265
				MsgError = "Name not found in this collection."
			Case 3266
				MsgError = "Append illegal - Field is part of a TableDefs collection."
			Case 3267
				MsgError = "Property value only valid when Field is part of a rdoResultset."
			Case 3268
				MsgError = "Cannot set the property of an object which is part of a Database object."
			Case 3269
				MsgError = "Append illegal - Index is part of a TableDefs collection."
			Case 3270
				MsgError = "Property not found."
			Case 3271
				MsgError = "Invalid property value."
			Case 3272
				MsgError = "Object is not an array."
			Case 3273
				MsgError = "Method not applicable for this object."
			Case 3274
				MsgError = "External table isn't in the expected format."
			Case 3275
				MsgError = "Unexpected error from external database driver ( )."
			Case 3276
				MsgError = "Invalid database ID."
			Case 3277
				MsgError = "Can't have more than 10 fielrs in an index."
			Case 3278
				MsgError = "Database engine has not been initialized."
			Case 3279
				MsgError = "Database engine has already been initialized."
			Case 3280
				MsgError = "Can't delete a field that is part of an index."
			Case 3281
				MsgError = "Can't delete an index that is used in a relationship."
			Case 3282
				MsgError = "Can't perform operation on a nontable."
			Case 3283
				MsgError = "Primary key already exists."
			Case 3284
				MsgError = "Index already exists."
			Case 3285
				MsgError = "Invalid index definition."
			Case 3286
				MsgError = "Invalid type for Memo field."
			Case 3287
				MsgError = "Can't create index on Memo field or Long Binary field."
			Case 3288
				MsgError = "Invalid ODBC driver."
			Case 3289
				MsgError = "Paradox: No primary index."
			Case 3290
				MsgError = "Syntax error."
			Case 3291
				MsgError = "Syntax error in CREATE TABLE statement."
			Case 3292
				MsgError = "Syntax error in CREATE INDEX statement."
			Case 3293
				MsgError = "Syntax error in column definition."
			Case 3294
				MsgError = "Syntax error in ALTER TABLE statement."
			Case 3295
				MsgError = "Syntax error in DROP INDEX statement."
			Case 3296
				MsgError = "Syntax error in DROP statement."
			Case 3297
				MsgError = "Operation not supported in version 1.1"
			Case 3298
				MsgError = "Couldn't import. No recorrs found or all recorrs contained errors."
			Case 3299
				MsgError = "Several tables exist with that name; please specify owner, as in 'owner.table'."
			Case Else
				MsgError = "Error Desconocido"
		End Select
	End Function
	
	Public Sub CierraConeccion()
		
		Dim intIndice As Short ' En vez de utilizar el I en los For-Next
		Dim intIndiceEn As Short ' En vez de utilizar el J en los For-Next
		
		' Para cada una de las connecciones de la aplicación
		For intIndice = gen.rdoConnections.Count - 1 To 0 Step -1
			
			'Cierra todos los rdoResultsets
			For intIndiceEn = gen.rdoConnections.Item(intIndice).rdoResultsets.Count - 1 To 0 Step -1
				gen.rdoConnections.Item(intIndice).rdoResultsets.Item(intIndiceEn).Close()
			Next intIndiceEn
			
		Next intIndice
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mblncn Then gcn.Close() 'Si se abrio la connecion se cierra
		If mblnEnvOpen Then gen.Close() 'Se se abrio el ambiente lo cierra
		mblncn = False
		mblnEnvOpen = False
		
	End Sub
	
	Public Sub AbreConeccion()
		'----------------------------------------------------'
		'     Abre la coneccion con la base de datos         '
		'----------------------------------------------------'
		
		On Error GoTo blnAbreConeccion_Err
		
		
		Dim strError As String ' Ultimo Mensaje de Errores de Rdo's
		Dim intIndice As Short ' En vez de utilizar I para los For-next
		Dim lngLongitudArchivo As Integer ' Longitud del archivo para checar si esta en RED
		
		
		' Busca los Parametros Iniciales del TRANS.INI
		'UPGRADE_WARNING: Couldn't resolve default property of object BuscaDirectorioWindows(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mstrDirectorioWindows = BuscaDirectorioWindows()
		
		' Inicializa el Environment (Ambiente de trabajo)
		RDOrdoEngine_definst.rdoDefaultCursorDriver = RDO.CursorDriverConstants.rdUseOdbc
		gen = RDOrdoEngine_definst.rdoCreateEnvironment("app", "app", "app")
		mblnEnvOpen = True 'Se efectuo la apertura del ambiente
		
		'Abre Conección
AbreConeccion_Apertura: 
		gcn = gen.OpenConnection(gstrServidor, RDO.PromptConstants.rdDriverNoPrompt, False, "DSN=" & gstrServidor & ";" & "UID=" & gstrLogin & ";" & "PWD=" & gstrPassword & ";" & "DATABASE=" & gstrBaseDeDatos & ";")
		mblncn = True 'Se efectuo la apertura de SQL
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.QueryTimeout. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gcn.QueryTimeout = 3600
		
		' Obtiene el nombre de la computadora
		gstrComputerName = New String(Chr(0), 20) ' Asigna 20 Espacios a la variable
		If GetComputerName(gstrComputerName, 256) Then
			gstrComputerName = UCase(Mid(gstrComputerName, 1, InStr(gstrComputerName, Chr(0)) - 1))
			gstrComputerName = TextoSinEspacio(gstrComputerName)
		Else
			gstrComputerName = "NoTiene"
		End If
		
		GoTo blnAbreConeccion_Exit
		
blnAbreConeccion_Err: 
		Dim lngIndice As Integer
		Dim strmsg As String
		
		Select Case Err.Number
			Case 53, 68 'File not found, Device unavailable
				strmsg = "Necesitas primeramente entrar a la Red"
				
			Case 40002
				For lngIndice = 0 To RDOrdoEngine_definst.rdoErrors.Count - 1
					strmsg = strmsg & RDOrdoEngine_definst.rdoErrors(lngIndice).Description & System.Windows.Forms.Keys.Return
				Next lngIndice
			Case Else
				strmsg = Err.Number & " " & ErrorToString()
		End Select
		Err.Clear()
		MsgBox("Error al Conectarse al Servidor de SQL." & Chr(13) & strmsg, MsgBoxStyle.Exclamation, "Abre Coneccion")
		
		If mblnEnvOpen Then gen.Close() 'Si se efectuo la apertura del ambiente la cierra
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If mblncn Then gcn.Close() 'Si se efectuo la apertura de sql lo cierra
		
		'Termina la aplicacion que halla solicitado la conexion
		End
		
blnAbreConeccion_Exit: 
		
	End Sub
	
	Function BuscaParametrosIni(ByRef Encabezado As String, ByRef Variable As String) As String
		'**********************
		'Descripcion : Lee el parametro solicitado por la VARIABLE en el .INI
		'   El GetPrivateProfileString da como salida la longitud del parametro solicitado
		' INPUT PARAMETERS:
		'   Encabezado: Nombre del encabezado en las secciones del INI
		'   Variable: Nombre del parametro o de la Llave
		'
		'**********************
		
		Dim strTemp As String ' Variable de Paso para almacenar la salida
		
		strTemp = New String(Chr(0), 255) 'Asigna 255 Nulos a la Variable
		BuscaParametrosIni = Left(strTemp, GetPrivateProfileString(Encabezado, Variable, "", strTemp, Len(strTemp), gstrArchivoIni)) 'Reasigna la variable a BuscaParametrosIni
		
		If BuscaParametrosIni = "" Then
			MsgBox("Error al buscar en Archivo INI, no se encontró " & Variable, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Parametros Ini")
			End
		End If
		
	End Function
	
	Function BuscaDirectorioWindows() As Object
		'**********************
		'Descripcion : Localiza y guarda la ruta de acceso que tiene el directorio
		'              de Windows
		'**********************
		
		Dim strTemp As String ' Variable de Paso para almacenar el Directorio
		Dim lngLon As Short ' Longitud del nombre del Directorio
		
		strTemp = New String(Chr(0), 145)
		lngLon = GetWindowsDirectory(strTemp, 145)
		strTemp = Left(strTemp, lngLon)
		
		If Right(strTemp, 1) <> "\" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object BuscaDirectorioWindows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BuscaDirectorioWindows = strTemp & "\"
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object BuscaDirectorioWindows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BuscaDirectorioWindows = strTemp
		End If
	End Function
	
	Public Sub AbreConeccionesRemotas()
		'----------------------------------------------------'
		'     Abre la coneccion con la base de datos         '
		'----------------------------------------------------'
		On Error GoTo Err_AbreConeccionesRemotas
		
		Dim rsServidores As RDO.rdoResultset
		Dim strSQL As String
		Dim strError As String
		Dim i As Short
		Dim bytIntento As Byte
		
		gblnConeccionVictoria = False
		gblnConeccionReynosa = False
		gblnConeccionTallerCentral = False
		gblnConeccionTampico = False
		gblnConeccionValles = False
		gblnConeccionMatamoros = False
		gblnConeccionMante = False
		gblnConeccionSanLuis = False
		bytIntento = 0
		
		strSQL = "select distinct Servidor from Base where BaseLocal = 0 and Servidor IS NOT NULL ORDER BY Servidor"
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsServidores = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
AbreConeccion_Apertura: 
		Do While Not rsServidores.EOF
			bytIntento = bytIntento + 1
			Select Case Trim(rsServidores.rdoColumns("Servidor").Value)
				'        Case "TRANSOPER"
				'           Set gcnVictoria = gen.OpenConnection(Trim(rsServidores!Servidor), rdDriverNoPrompt, False, _
				'"DSN=" & Trim(rsServidores!Servidor) & ";" & _
				'"UID=" & LOGIN & ";" & _
				'"PWD=" & PASSWORD & ";" & _
				'"DATABASE=" & gstrBaseDeDatos & ";")
				'          gblnConeccionVictoria = True       'Se efectuo la apertura de SQL
				'         gcnVictoria.QueryTimeout = 3600
				'        bytIntento = 0
				
				'        Case "PENSION_REYNOSA"
				'           Set gcnReynosa = gen.OpenConnection(Trim(rsServidores!Servidor), rdDriverNoPrompt, False, _
				'"DSN=" & Trim(rsServidores!Servidor) & ";" & _
				'"UID=" & LOGIN & ";" & _
				'"PWD=" & PASSWORD & ";" & _
				'"DATABASE=" & gstrBaseDeDatos & ";")
				'          gblnConeccionReynosa = True       'Se efectuo la apertura de SQL
				'         gcnReynosa.QueryTimeout = 3600
				'        bytIntento = 0
				
				Case "TCServer"
					If gstrServidor <> "TCSERVER" Then
						gcnTallerCentral = gen.OpenConnection(Trim(rsServidores.rdoColumns("Servidor").Value), RDO.PromptConstants.rdDriverNoPrompt, False, "DSN=" & Trim(rsServidores.rdoColumns("Servidor").Value) & ";" & "UID=" & LOGIN & ";" & "PWD=" & PASSWORD & ";" & "DATABASE=" & gstrBaseDeDatos & ";")
						gblnConeccionTallerCentral = True 'Se efectuo la apertura de SQL
						'UPGRADE_WARNING: Couldn't resolve default property of object gcnTallerCentral.QueryTimeout. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gcnTallerCentral.QueryTimeout = 3600
						bytIntento = 0
					End If
					
					
					'        Case "NTTAM"
					'           Set gcnTampico = gen.OpenConnection(Trim(rsServidores!Servidor), rdDriverNoPrompt, False, _
					'"DSN=" & Trim(rsServidores!Servidor) & ";" & _
					'"UID=" & LOGIN & ";" & _
					'"PWD=" & PASSWORD & ";" & _
					'"DATABASE=" & gstrBaseDeDatos & ";")
					'          gblnConeccionTampico = True       'Se efectuo la apertura de SQL
					'         gcnTampico.QueryTimeout = 3600
					'        bytIntento = 0
					
					'        Case "NTVAL"
					'           Set gcnValles = gen.OpenConnection(Trim(rsServidores!Servidor), rdDriverNoPrompt, False, _
					'"DSN=" & Trim(rsServidores!Servidor) & ";" & _
					'"UID=" & LOGIN & ";" & _
					'"PWD=" & PASSWORD & ";" & _
					'"DATABASE=" & gstrBaseDeDatos & ";")
					'          gblnConeccionValles = True       'Se efectuo la apertura de SQL
					'         gcnValles.QueryTimeout = 3600
					'        bytIntento = 0
					
					'        Case "NTMAT"
					'           Set gcnMatamoros = gen.OpenConnection(Trim(rsServidores!Servidor), rdDriverNoPrompt, False, _
					'"DSN=" & Trim(rsServidores!Servidor) & ";" & _
					'"UID=" & LOGIN & ";" & _
					'"PWD=" & PASSWORD & ";" & _
					'"DATABASE=" & gstrBaseDeDatos & ";")
					'          gblnConeccionMatamoros = True       'Se efectuo la apertura de SQL
					'         gcnMatamoros.QueryTimeout = 3600
					'        bytIntento = 0
					
					'        Case "NTMAN"
					'           Set gcnMante = gen.OpenConnection(Trim(rsServidores!Servidor), rdDriverNoPrompt, False, _
					'"DSN=" & Trim(rsServidores!Servidor) & ";" & _
					'"UID=" & LOGIN & ";" & _
					'"PWD=" & PASSWORD & ";" & _
					'"DATABASE=" & gstrBaseDeDatos & ";")
					'          gblnConeccionMante = True       'Se efectuo la apertura de SQL
					'         gcnMante.QueryTimeout = 3600
					'        bytIntento = 0
					
					'        Case "NTSLP"
					'           Set gcnSanLuis = gen.OpenConnection(Trim(rsServidores!Servidor), rdDriverNoPrompt, False, _
					'"DSN=" & Trim(rsServidores!Servidor) & ";" & _
					'"UID=" & LOGIN & ";" & _
					'"PWD=" & PASSWORD & ";" & _
					'"DATABASE=" & gstrBaseDeDatos & ";")
					'         gblnConeccionSanLuis = True       'Se efectuo la apertura de SQL
					'        gcnSanLuis.QueryTimeout = 3600
					'       bytIntento = 0
					
				Case "ATMSERVER"
					gcnATMT = gen.OpenConnection(Trim(rsServidores.rdoColumns("Servidor").Value), RDO.PromptConstants.rdDriverNoPrompt, False, "DSN=" & Trim(rsServidores.rdoColumns("Servidor").Value) & ";" & "UID=" & LOGIN & ";" & "PWD=" & PASSWORD & ";" & "DATABASE=" & gstrBaseDeDatos & ";")
					gblnConeccionATMT = True 'Se efectuo la apertura de SQL
					'UPGRADE_WARNING: Couldn't resolve default property of object gcnATMT.QueryTimeout. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gcnATMT.QueryTimeout = 3600
					bytIntento = 0
					
				Case "LUMXSBD"
					gcnLUMX = gen.OpenConnection(Trim(rsServidores.rdoColumns("Servidor").Value), RDO.PromptConstants.rdDriverNoPrompt, False, "DSN=" & Trim(rsServidores.rdoColumns("Servidor").Value) & ";" & "UID=" & LOGIN & ";" & "PWD=" & PASSWORD & ";" & "DATABASE=" & gstrBaseDeDatos & ";")
					gblnConeccionLUMX = True 'Se efectuo la apertura de SQL
					'UPGRADE_WARNING: Couldn't resolve default property of object gcnLUMX.QueryTimeout. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gcnLUMX.QueryTimeout = 3600
					bytIntento = 0
					
			End Select
AbreConeccion_MoveNext: 
			rsServidores.MoveNext()
		Loop 
		rsServidores.Close()
		
		Exit Sub
		
		'"IM002: [Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified"
		'"IM002: [Microsoft][ODBC Driver Manager] El nombre del origen de datos no se encontró y no se especificó ningún controlador predeterminado"
		'"IM002: [Microsoft][Administrador de controladores ODBC] El nombre del origen de datos no se encontró y no se especificó ningún controlador predeterminado"
		'"IM002: [Microsoft][Administrador de controladores ODBC] No se encuentra el nombre del origen de datos y no se especificó ningún controlador predeterminado"
		
Err_AbreConeccionesRemotas: 
		strError = ""
		If Err.Number = 40002 Then
			If (Mid(ErrorToString(), 1, 65) = "IM002: [Microsoft][ODBC Driver Manager] Data source name not foun" Or Mid(ErrorToString(), 1, 65) = "IM002: [Microsoft][ODBC Driver Manager] El nombre del origen de d" Or Mid(ErrorToString(), 1, 65) = "IM002: [Microsoft][Administrador de controladores ODBC] El nombre" Or Mid(ErrorToString(), 1, 65) = "IM002: [Microsoft][Administrador de controladores ODBC] No se enc") And bytIntento = 1 Then
				'If MsgBox("La ruta de Acceso a los Datos " & Trim(rsServidores!Servidor) & " No existe Deseas dar el alta ", vbYesNo + vbQuestion) = vbYes Then _
				'RegistraODBC Trim(rsServidores!Servidor), "SIM", "SQL Server"
				RegistraODBC(Trim(rsServidores.rdoColumns("Servidor").Value), "SIM", "SQL Server")
				RDOrdoEngine_definst.rdoErrors.Clear()
				Err.Clear()
				Resume AbreConeccion_Apertura
			End If
			
			For i = 0 To RDOrdoEngine_definst.rdoErrors.Count - 1
				strError = strError & RDOrdoEngine_definst.rdoErrors(i).Description & vbLf
				If RDOrdoEngine_definst.rdoErrors(i).SQLState = "08001" Then
					MsgBox("Error al Conectarse al Servidor ->" & Trim(rsServidores.rdoColumns("Servidor").Value) & vbLf & strError, MsgBoxStyle.Exclamation, "Abre Conecciones Remotas")
					RDOrdoEngine_definst.rdoErrors.Clear()
					Err.Clear()
					bytIntento = 0
					Resume AbreConeccion_MoveNext
				End If
			Next i
		Else
			strError = Err.Number & " " & ErrorToString()
		End If
		MsgBox("Error SIN IDENTIFICAR EN " & Trim(rsServidores.rdoColumns("Servidor").Value) & Chr(13) & strError, MsgBoxStyle.Exclamation, "Abre Conecciones Remotas")
		
		If mblnEnvOpen Then gen.Close() 'Si se efectuo la apertura del ambiente la cierra
		CierraConeccionesRemotas()
		
		End
		
	End Sub
	
	Public Sub RegistraODBC(ByVal vstrServidor As String, ByVal vstrDBName As String, ByVal vstrDriver As String)
		'Subrutina que registra la llamada a ODBC de un determinado servidor
		'CGSR 23/07/1997
		
		Dim strAttributes As String
		
		strAttributes = "Description=" & vstrServidor & Chr(13) & "Server=" & vstrServidor & Chr(13) & "Database=" & vstrDBName & Chr(13) & "FastConnectOption=Yes" & Chr(13) & "UseProcForPrepare=No" & Chr(13) & "OEMTOANSI=Yes" & Chr(13)
		
		RDOrdoEngine_definst.rdoRegisterDataSource(DSN:=vstrServidor, Driver:=vstrDriver, Silent:=True, Attributes:=strAttributes)
		
	End Sub
	
	
	Public Sub CierraConeccionAlmacen()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnAlmacen.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionAlmacen Then gcnAlmacen.Close()
		
		gblnConeccionAlmacen = False
		
	End Sub
	
	Public Sub CierraConeccionesRemotas()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnVictoria.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionVictoria Then gcnVictoria.Close()
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnTallerCentral.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionTallerCentral Then gcnTallerCentral.Close()
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnReynosa.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionReynosa Then gcnReynosa.Close()
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnTampico.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionTampico Then gcnTampico.Close()
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnValles.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionValles Then gcnValles.Close()
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnMatamoros.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionMatamoros Then gcnMatamoros.Close()
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnMante.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionMante Then gcnMante.Close()
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnSanLuis.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionSanLuis Then gcnSanLuis.Close()
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnATMT.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionATMT Then gcnATMT.Close()
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnLUMX.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionLUMX Then gcnLUMX.Close()
		
		gblnConeccionVictoria = False
		gblnConeccionTallerCentral = False
		gblnConeccionReynosa = False
		gblnConeccionTampico = False
		gblnConeccionValles = False
		gblnConeccionMante = False
		gblnConeccionMatamoros = False
		gblnConeccionSanLuis = False
		gblnConeccionATMT = False
		gblnConeccionLUMX = False
		
	End Sub
	
	Public Sub AbreConeccionAlmacen()
		'-------------------------------------------------------------'
		'     Abre la coneccion con la base de datos de Almacen       '
		'-------------------------------------------------------------'
		On Error GoTo Err_AbreConeccionAlmacen
		
		Dim strSQL As String
		Dim strError As String
		Dim i As Short
		
		gblnConeccionAlmacen = False
		
		gcnAlmacen = gen.OpenConnection(gstrServidorAlmacen, RDO.PromptConstants.rdDriverNoPrompt, False, "DSN=" & gstrServidorAlmacen & ";" & "UID=" & LOGIN & ";" & "PWD=" & PASSWORD & ";" & "DATABASE=" & gstrBaseDeDatosAlmacen & ";")
		gblnConeccionAlmacen = True 'Se efectuo la apertura de SQL
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnAlmacen.QueryTimeout. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gcnAlmacen.QueryTimeout = 3600
		
		Exit Sub
		
Err_AbreConeccionAlmacen: 
		If Err.Number = 40002 Then
			For i = 0 To RDOrdoEngine_definst.rdoErrors.Count - 1
				strError = strError & " , " & RDOrdoEngine_definst.rdoErrors(i).Description
			Next i
			RDOrdoEngine_definst.rdoErrors.Clear()
		Else
			strError = Err.Number & " " & ErrorToString()
		End If
		Err.Clear()
		MsgBox("Error al Conectarse al Servidor " & gstrServidorAlmacen & Chr(13) & strError, MsgBoxStyle.Exclamation, "Abre Coneccion Almacen")
		
		If mblnEnvOpen Then gen.Close() 'Si se efectuo la apertura del ambiente la cierra
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnAlmacen.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionAlmacen Then gcnAlmacen.Close()
		
		End
		
	End Sub
	
	Public Sub AbreConeccionCentral()
		'----------------------------------------------------'
		'     Abre la coneccion con la base de datos         '
		'----------------------------------------------------'
		On Error GoTo Err_AbreConeccionesRemotas
		
		Dim strSQL As String
		Dim strError As String
		Dim i As Short
		
		gblnConeccionVictoria = False
		
		gcnVictoria = gen.OpenConnection(gstrServidorCentral, RDO.PromptConstants.rdDriverNoPrompt, False, "DSN=" & gstrServidorCentral & ";" & "UID=" & LOGIN & ";" & "PWD=" & PASSWORD & ";" & "DATABASE=" & gstrBaseDeDatos & ";")
		gblnConeccionVictoria = True 'Se efectuo la apertura de SQL
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.QueryTimeout. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gcn.QueryTimeout = 3600
		
		Exit Sub
		
Err_AbreConeccionesRemotas: 
		If Err.Number = 40002 Then
			For i = 0 To RDOrdoEngine_definst.rdoErrors.Count - 1
				strError = strError & " , " & RDOrdoEngine_definst.rdoErrors(i).Description
			Next i
		Else
			strError = Err.Number & " " & ErrorToString()
		End If
		MsgBox("Error al Conectarse al Servidor " & gstrServidorCentral & Chr(13) & strError, MsgBoxStyle.Exclamation, "Abre Coneccion Central")
		
		If mblnEnvOpen Then gen.Close() 'Si se efectuo la apertura del ambiente la cierra
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnVictoria.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionVictoria Then gcnVictoria.Close()
		
		End
		
	End Sub
	Public Function ObtieneFechaExe() As String
		'----------------------
		' Obtiene la fecha en que se generó el ejecutable,
		' para mejor referencia al ofrecer soporte técnico
		'----------------------
		Dim strAplicacion As String
		
		'Obtener nombre y ruta de la aplicación
		strAplicacion = My.Application.Info.DirectoryPath
		If Right(strAplicacion, 1) <> "\" Then
			strAplicacion = strAplicacion & "\"
		End If
		
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		strAplicacion = strAplicacion & My.Application.Info.AssemblyName & ".EXE"
		'Verificar que sí exista el archivo, y por tanto obtener los datos
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Dir(strAplicacion) <> "" Then
			' Aplicar el formato para presentarlo en el título de la aplicación
			ObtieneFechaExe = "v" & VB6.Format(FileDateTime(strAplicacion), "yymmdd.hhmm")
		Else
			ObtieneFechaExe = ""
		End If
		
	End Function
End Module
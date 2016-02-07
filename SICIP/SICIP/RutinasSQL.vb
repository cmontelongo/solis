Option Strict Off
Option Explicit On
Option Compare Text
Imports VB = Microsoft.VisualBasic
Module mdlRutinasSQL
	
	'Constantes para el formato de fechas
	'************************************************
	Public Const FECHAMESLETRA As String = "dd / mmm / yyyy"
	Public Const FECHAMMDDYYYY As String = "mm/dd/yyyy"
	Public Const FECHADDMMYY As String = "dd/mm/yy"
	Public Const FECHADDMMYYYY As String = "dd/mm/yyyy"
	Public Const FECHAYYYYMMDD As String = "yyyy-mm-dd"
	Public Const HORAMINUTOS As String = " hh:nn"
	Public Const HORAMINUTOSSEGUNDOS As String = " hh:nn:ss"
	Public Const FECHADIALETRA As String = "ddd, dd / mmm / yy"
	Public Const HORAINICIAL As String = " 0:00"
	Public Const HORAFINAL As String = " 23:59"
	'************************************************
	
	' Cosntantes de resolucion de la pantalla
	Public Const RESOLUCIONANCHO As Short = 600
	Public Const RESOLUCIONALTO As Short = 800
	
	
	'utilizadas para búsquedas de listas y combos
	'************************************************
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	Const CB_ERR As Short = (-1)
	Const WM_USER As Integer = &H400
	Const CB_EncuentraTexto As Integer = &H14C
	'************************************************
	
	Declare Function GetComputerName Lib "kernel32"  Alias "GetComputerNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	
	'Para el momento de limpiar un spread
	'****************************************
	Public Const COLORBLANCO As Integer = &HFFFFFF
	Public Const COLORGRIS As Integer = &HC0C0C0
	Public Const COLORNEGRO As Integer = &H0
	'****************************************
	
	Public gstrDirectorioRpt As String
	Public gvntCveConsulta As Object
	
	Public mintOPCION_TRAN_PENDIENTES As Short
	'Variable que contiene el modo en que debe tomarse la interrupci¢n
	'de una edición o alta en un ABC :
	Public Const mintNO_GRABAR As Short = 1 ' Ignorar la edición o alta.
	Public Const mintPREGUNTA_GRABAR As Short = 2 ' Pregunta de confirmaci¢n para grabar los cambios.
	Public Const mintNO_PREGUNTA_GRABAR As Short = 3 ' Grabar los cambios autom ticamente, sin preguntar.
	
	Public mintOPCION_TRAN_SALIDA As Short 'Variable que contiene el modo en que debe
	'tomarse la interrupcion de una edicion o
	'alta en un ABC al intentar salir de la pantalla.
	
	'*----------------------------------------------------------------------*
	'* Variables que se utilizan para validar impresiones en Operaciones    *
	'*----------------------------------------------------------------------*
	Public gstrValorABuscar As String
	Public gstrValorABuscar2 As String
	Public mlngRenglon As Integer
	Public mlngColumna As Integer
	Public mlngRenglon2 As Integer
	Public mlngColumna2 As Integer
	Public gintPermisoTipo As Short
	
	Public gblnImprimeReporteViajes As Boolean
	Public gblnContinuar As Boolean ' Se usa como resultado de una frm de confirmación.
	Public gblnPasswordOk As Boolean ' Se usa en la .frm que pide password
	
	Public strMensaje As String
	
	Public Const SINDECIMALES As String = "##########0"
	Public Const DOSDECIMALES As String = "##########0.00"
	Public Const DOSDECIMALESCOMAS As String = "##,###,###0.00"
	Public Const TRESDECIMALES As String = "##########0.000"
	
	' Tipos de periodicidades de tareas preventivas
	Public Const PERIODICIDADNINGUNA As Short = 1
	Public Const PERIODICIDADSERVICIO As Short = 2
	Public Const PERIODICIDADTAREAINDIVIDUAL As Short = 3
	Public Const PERIODICIDADSERVICIOCADAXKMS As Short = 4
	
	' Tipo de Manejo de los preventivos
	Public Const PREVENTIVOSINDIVIDUALES As Short = 1
	Public Const PREVENTIVOSAGRUPADOS As Short = 2
	Public Const PREVENTIVOSSERVICIOCADAXKMS As Short = 3
	
	' Constantes de Tipos de Entradas y Salidas
	Public Const ENTRADASNUEVAS As Short = 1
	Public Const ENTRADASCONSIGNACION As Short = 2
	Public Const ENTRADASPORREPARAR As Short = 3
	Public Const ENTRADASTRASPASOS As Short = 4
	Public Const ENTRADASREPARADAS As Short = 5
	Public Const ENTRADASDEVOLUCION As Short = 6
	Public Const ENTRADASGARANTIA As Short = 7
	Public Const SALIDASODT As Short = 21
	Public Const SALIDASTRASPASOS As Short = 22
	Public Const SALIDASPORREPARAR As Short = 23
	Public Const SALIDASDEVOLUCION As Short = 24
	Public Const SALIDASVENTA As Short = 25
	
	' Constantes de Tipos de Movimientos de Almacen
	Public Const MOVIMIENTOENTRADA As Short = 1
	Public Const MOVIMIENTOSALIDA As Short = 2
	
	' Constantes de Tipos de Movimientos Contables
	Public Const MOVIMIENTOCARGO As Short = 1
	Public Const MOVIMIENTOCREDITO As Short = 2
	
	' Constantes de Divisiones de Almacen
	Public Const DIVISIONPROPIO As Short = 1
	Public Const DIVISIONCONSIGNACION As Short = 2
	
	' Estatus de Movimientos de Almacen
	Public Const MOVIMIENTOALMACENCAPTURADO As Short = 1
	Public Const MOVIMIENTOALMACENCONTABILIZADO As Short = 5
	
	' Tipos de manejo de almacen
	Public Const MANEJOALMACENPEPS As Short = 1
	Public Const MANEJOALMACENUEPS As Short = 2
	Public Const MANEJOALMACENPROM As Short = 3
	Public Const gstrTipoSalidaAlmacen As Short = 3
	
	' Tipos de Permisos de Acceso
	Public Const PERMISONOACCESO As Short = 0
	Public Const PERMISOACTUALIZACION As Short = 1
	Public Const PERMISOCONSULTA As Short = 2
	
	' Tipos de Combustible
	Public Const COMBUSTIBLEDIESEL As Short = 1
	Public Const COMBUSTIBLEGASOLINA As Short = 2
	Public Const COMBUSTIBLEGAS As Short = 3
	
	'*-----------------------------------------------------------------------*
	'*          Variables para cambiar resolucion del monitor                *
	'*-----------------------------------------------------------------------*
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function EnumDisplaySettings Lib "user32"  Alias "EnumDisplaySettingsA"(ByVal lpszDeviceName As Integer, ByVal iModeNum As Integer, ByRef lpDevMode As Any) As Boolean
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function ChangeDisplaySettings Lib "user32"  Alias "ChangeDisplaySettingsA"(ByRef lpDevMode As Any, ByVal dwflags As Integer) As Integer
	
	Dim mdvmDispositivoGrafico As DEVMODE
	
	Const CCDEVICENAME As Short = 32
	Const CCFORMNAME As Short = 32
	Const DM_PELSWIDTH As Integer = &H80000
	Const DM_PELSHEIGHT As Integer = &H100000
	Public gsngAncho As Single
	Public gsngAlto As Single
	
	Private Structure DEVMODE
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCDEVICENAME),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=CCDEVICENAME)> Public dmDeviceName() As Char
		Dim dmSpecVersion As Short
		Dim dmDriverVersion As Short
		Dim dmSize As Short
		Dim dmDriverExtra As Short
		Dim dmFields As Integer
		Dim dmOrientation As Short
		Dim dmPaperSize As Short
		Dim dmPaperLength As Short
		Dim dmPaperWidth As Short
		Dim dmScale As Short
		Dim dmCopies As Short
		Dim dmDefaultSource As Short
		Dim dmPrintQuality As Short
		Dim dmColor As Short
		Dim dmDuplex As Short
		Dim dmYResolution As Short
		Dim dmTTOption As Short
		Dim dmCollate As Short
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCFORMNAME),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=CCFORMNAME)> Public dmFormName() As Char
		Dim dmUnusedPadding As Short
		Dim dmBitsPerPel As Short
		Dim dmPelsWidth As Integer
		Dim dmPelsHeight As Integer
		Dim dmDisplayFlags As Integer
		Dim dmDisplayFrequency As Integer
	End Structure
	
	Public gstrEquipo_Operador As String
	'UPGRADE_WARNING: Lower bound of array gintTabPermiso was changed from 0 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public gintTabPermiso(10) As Short ' Guarda el tipo de permiso definido por Tab
	
	Public glngPrintPageCount As Integer
	
	Public gstrArchivoRpt As String
	Public Sub CambiaColorSpread(ByRef rsprSpread As AxFPSpread.AxvaSpread, ByVal vlngLinea As Integer, ByRef vintColumnaInicial As Short, ByRef vintColumnaFinal As Short, ByVal vvntColor As Object, ByRef vintPropiedad As Short)
		'-----------------------------------------------------------------------------------------------
		'   Descripcion:    Procedimiento que cambia el color de las letras de un renglon del spread
		'   Entradas:       1) Nombre del Spread
		'                   2) Renglon en el que se cambiará el color
		'                   3) Columna inicial
		'                   4) Columna Final
		'                   5) Color deseado
		'                   6) Propiedad a la que se quiere aplicar el color:
		'                       1 -> ForeColor , 2 -> BackColor , 3 -> BorderColor
		'-----------------------------------------------------------------------------------------------
		Dim lngColumna As Short
		
		rsprSpread.Row = vlngLinea
		For lngColumna = vintColumnaInicial To vintColumnaFinal
			rsprSpread.Col = lngColumna
			Select Case vintPropiedad
				Case 1
					'UPGRADE_WARNING: Couldn't resolve default property of object vvntColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rsprSpread.ForeColor = System.Drawing.ColorTranslator.FromOle(vvntColor)
				Case 2
					'UPGRADE_WARNING: Couldn't resolve default property of object vvntColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rsprSpread.BackColor = System.Drawing.ColorTranslator.FromOle(vvntColor)
				Case 3
					'UPGRADE_WARNING: Couldn't resolve default property of object rsprSpread.BorderColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object vvntColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rsprSpread.BorderColor = vvntColor
			End Select
		Next 
		
	End Sub
	
	Public Sub InsertaLlantaTrayectoria(ByRef lngCveLlanta As Integer, ByRef intPiso As Short, ByRef intProfundidadInicial As Short, ByRef intProfundidadFinal As Short, ByRef lngKmsInicial As Integer, ByRef lngKmsFinal As Integer, ByRef sngCosto As Single, ByRef strFechaInicioPiso As String)
		
		Dim strSQL As String
		
		On Error GoTo Err_Inserta
		
		strSQL = "insert into LlantaTrayectoria "
		strSQL = strSQL & " (CveLlanta, CveLlantaPiso, ProfundidadInicial, ProfundidadFinal,"
		strSQL = strSQL & "  KmsInicial, KmsFinal,Costo, FechaInicioPiso) "
		strSQL = strSQL & " values (" & lngCveLlanta & ","
		strSQL = strSQL & intPiso & "," & intProfundidadInicial & ","
		strSQL = strSQL & intProfundidadFinal & "," & lngKmsInicial & "," & lngKmsFinal
		strSQL = strSQL & "," & sngCosto & ",'" & strFechaInicioPiso & "'" & ")"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gcn.Execute(strSQL)
		
		
		Exit Sub
		
Err_Inserta: 
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Dim strmsg As String 'String del Error
		Dim lngIndice As Integer 'Indice del Error de RDO
		
		Select Case Err.Number
			Case 40002
				For lngIndice = 0 To RDOrdoEngine_definst.rdoErrors.Count - 1
					strmsg = strmsg & RDOrdoEngine_definst.rdoErrors(lngIndice).Description & Chr(System.Windows.Forms.Keys.Return)
				Next lngIndice
				RDOrdoEngine_definst.rdoErrors.Clear()
				strmsg = strmsg & vbLf & strSQL
			Case Else
				strmsg = Err.Number & " " & ErrorToString()
		End Select
		Err.Clear()
		MsgBox("Error al insertar trayectoria " & strmsg, MsgBoxStyle.Critical, "InsertaLlantaTrayectoria")
		
	End Sub
	
	Public Sub InsertaLlantaHistorial(ByRef lngCveLlanta As Integer, ByRef intCveLlantaEstatus As Short, ByRef sngCosto As Single, ByRef lngCveUnidad As Integer, ByRef intPosicion As Short)
		
		Dim strSQL As String
		Dim rsQuery As RDO.rdoResultset
		Dim strFechaHoy As String
		
		strFechaHoy = ObtieneFechaHora(1)
		
		strSQL = "select * from Llanta where CveLlanta = " & lngCveLlanta
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		If Not rsQuery.EOF Then
			strSQL = "insert into LlantaHistorial "
			strSQL = strSQL & " (CveLlanta, Fecha, CveLlantaEstatus, CveLlantaPiso,"
			strSQL = strSQL & "  Profundidad, KmsAcumulados, Costo, CveUnidad,"
			strSQL = strSQL & "  Posicion) "
			strSQL = strSQL & " values (" & lngCveLlanta & ","
			strSQL = strSQL & " '" & VB6.Format(strFechaHoy, FECHAMMDDYYYY & HORAMINUTOS) & "',"
			strSQL = strSQL & intCveLlantaEstatus & "," & rsQuery.rdoColumns("CveLlantaPiso").Value & ","
			strSQL = strSQL & rsQuery.rdoColumns("Profundidad").Value & "," & rsQuery.rdoColumns("KmsAcumulados").Value & "," & sngCosto & ","
			strSQL = strSQL & lngCveUnidad & "," & intPosicion & ")"
			
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcn.Execute(strSQL)
		End If
		rsQuery.Close()
		
	End Sub
	
	Sub ActualizaInventario(ByRef CveAlmacen As Short, ByRef CveDivision As Short, ByRef CveRefaccion As Integer, ByRef FechaMovimiento As Date, ByRef TipoMovimiento As Short, ByRef NumFactura As Short, ByRef PrecioUnitario As Single, ByRef Cantidad As Single)
		'--------------------------------------------------------------------------------
		'  Actualiza el inventario en base a datos de una entrada o salida
		'  de Almacen
		'
		'  Se reciben como parámetros:
		'       CveAlmacen  -> # de Almacen del movto.
		'       CveDivision  -> Indica si es propio o consignación
		'       CveRefaccion -> # de la refaccion
		'       FechaMovimiento  -> Fecha del movimiento
		'       Tipo Movimiento  -> 1 = Entrada     2 = Salida
		'       NumFactura       -> # de Factura en el caso de una entrada
		'       PrecioUnitario   -> Precio unitario de la refaccion
		'       Cantidad         -> Cantidad de refacciones
		'--------------------------------------------------------------------------------
		On Error GoTo err_ActualizaInventario
		
		Dim rsQuery As RDO.rdoResultset
		Dim strSQL As String
		
		Select Case TipoMovimiento
			Case MOVIMIENTOENTRADA
				strSQL = "select * from Inventario where CveAlmacen = " & CveAlmacen
				strSQL = strSQL & " and CveDivision = " & CveDivision
				strSQL = strSQL & " and CveRefaccion = " & CveRefaccion
				'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
				If Not rsQuery.EOF Then
					strSQL = " Update Inventario set EntradaCantidad = EntradaCantidad + " & Cantidad
					strSQL = strSQL & " , ExistenciaActual = ExistenciaActual + " & Cantidad
					strSQL = strSQL & " , FechaUltEntrada = '" & VB6.Format(FechaMovimiento, FECHAYYYYMMDD) & "'"
					strSQL = strSQL & " Where CveAlmacen = " & CveAlmacen
					strSQL = strSQL & " and CveDivision = " & CveDivision
					strSQL = strSQL & " and CveRefaccion = " & CveRefaccion
					'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gcn.Execute(strSQL)
				Else
					strSQL = " Insert into Inventario (CveAlmacen, CveDivision, CveRefaccion,"
					strSQL = strSQL & "SaldoInicialCantidad, EntradaCantidad, SalidaCantidad,"
					strSQL = strSQL & "ExistenciaActual, FechaUltEntrada)"
					strSQL = strSQL & " values (" & CveAlmacen & "," & CveDivision & ","
					strSQL = strSQL & CveRefaccion & ",0," & Cantidad & ",0," & Cantidad & ",'"
					strSQL = strSQL & VB6.Format(FechaMovimiento, FECHAYYYYMMDD) & "')"
					'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gcn.Execute(strSQL)
				End If
				
			Case MOVIMIENTOSALIDA
				strSQL = " Update Inventario set SalidaCantidad = SalidaCantidad + " & Cantidad
				strSQL = strSQL & " , ExistenciaActual = ExistenciaActual - " & Cantidad
				strSQL = strSQL & " , FechaUltSalida = '" & VB6.Format(FechaMovimiento, FECHAYYYYMMDD) & "'"
				strSQL = strSQL & " Where CveAlmacen = " & CveAlmacen
				strSQL = strSQL & " and CveDivision = " & CveDivision
				strSQL = strSQL & " and CveRefaccion = " & CveRefaccion
				'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gcn.Execute(strSQL)
				
		End Select
		
		Exit Sub
		
err_ActualizaInventario: 
		MsgBox("Error al Actualizar Inventario" & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	
	Public Function NombreMes(ByVal vbytMes As Byte) As String
		
		Select Case vbytMes
			Case 1
				NombreMes = "Enero"
			Case 2
				NombreMes = "Febrero"
			Case 3
				NombreMes = "Marzo"
			Case 4
				NombreMes = "Abril"
			Case 5
				NombreMes = "Mayo"
			Case 6
				NombreMes = "Junio"
			Case 7
				NombreMes = "Julio"
			Case 8
				NombreMes = "Agosto"
			Case 9
				NombreMes = "Septiembre"
			Case 10
				NombreMes = "Octubre"
			Case 11
				NombreMes = "Noviembre"
			Case 12
				NombreMes = "Diciembre"
			Case Else
				NombreMes = "****"
				
		End Select
		
	End Function
	
	Sub Posicionaselector(ByRef Llave As Object, ByRef Control As System.Windows.Forms.Control)
		'*************************************
		'Posiciona la LLAVE en el Control del combobox
		'Recibe como valor la llave cuando la llave es numerica
		' En caso de que la llave sea alfanumerica recibe el AbsolutePosition del rdoResultset
		'*************************************
		
		Dim blnLoEncontro As Boolean 'Bandera para saber si encontro la llave
		Dim intIndice1 As Short
		
		blnLoEncontro = False
		'UPGRADE_WARNING: Couldn't resolve default property of object Control.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For intIndice1 = 0 To Control.ListCount - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Control.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Control.ItemData(intIndice1) = Llave Then
				blnLoEncontro = True
				Exit For
			End If
		Next intIndice1
		
		If blnLoEncontro Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Control.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Control.ListIndex = intIndice1
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Control.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Control.ListIndex = -1
		End If
		
	End Sub
	
	Public Sub SeleccionaCampo(ByVal vtxtCampo As System.Windows.Forms.TextBox)
		
		vtxtCampo.SelectionStart = 0
		vtxtCampo.SelectionLength = Len(vtxtCampo.Text)
		
	End Sub
	
	Public Sub VerificaPreventivos(ByRef lngCveUnidad As Integer, ByRef intTipoPreventivos As Short)
		'*******************************************************************************************
		'  Selecciona las Tareas que ya estan vencidas, ya se por Fecha o por Kilometraje
		'  Obtiene la periodicidad de Dias y de Kilometros de la Tabla TareaPeriodicidad
		'  Obtiene el Kilometraje Acumulado de la Tabla de Unidad
		'  Obtiene la Fecha de Ocurrencia de la Tarea utilizando la siguiente Formula :
		'    (((Periodicidad - Tolerancia) + FechaOcurrencia) - el Dia de Hoy) <= 0 , Si es menor aparecera en el Programa
		'  Obtiene el Kilometraje de la Tarea utilizando la siguiente Formula :
		'    ((KmsAcumulados + (Periodicidad - Tolerancia)) - KmsAcum de la Unidad) <= 0 , Si es menor aparecera en el Programa
		'  Donde la Tarea sea la Ultima que se haya efectuado y este registrada en el Kardex.
		'  Donde la Periodicida sea Diferente a 0
		'
		'  Parametros:
		'        lngCveUnidad .- # de Unidad a la que se desean revisar los preventivos
		'        intTipoPreventivos .- Me indica si los preventivos son individuales o agrupados
		'*******************************************************************************************
		On Error GoTo Err_VerificaPreventivos
		
		Dim strSQL As String
		Dim strFormato As String
		Dim rsKms As RDO.rdoResultset
		Dim lngKmsPorVuelta As Object
		
		Select Case intTipoPreventivos
			
			Case PREVENTIVOSINDIVIDUALES
				strSQL = "select U.CveUnidad,T.CveTarea,T.Nombre,UK.KmsAcumulados,UK.FechaOcurrencia, "
				strSQL = strSQL & " (UK.KmsAcumulados+TP.PeriodicidadKms  - U.KmsAcumulados)  ToleranciaKms, "
				strSQL = strSQL & " (UK.KmsAcumulados+TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados)  VencidoKms, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias ,UK.FechaOcurrencia)) ToleranciaDias, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias + TP.ToleranciaDias,UK.FechaOcurrencia)) VencidoDias, 0 CveSubTarea "
				strSQL = strSQL & " from Unidad U (NOLOCK) , UnidadKardex UK (NOLOCK) , Tarea T (NOLOCK), TareaPeriodicidad TP (NOLOCK) "
				strSQL = strSQL & " where U.CveUnidad = UK.CveUnidad "
				strSQL = strSQL & " and U.CveUnidad = " & lngCveUnidad
				strSQL = strSQL & " and UK.CveTarea = T.CveTarea "
				strSQL = strSQL & " and T.CveTarea = TP.CveTarea "
				strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo "
				strSQL = strSQL & " and UK.FechaOcurrencia= (select MAX(FechaOcurrencia) Ocurrencia from UnidadKardex EK (NOLOCK) " & " where EK.CveUnidad = UK.CveUnidad and EK.CveTarea  = UK.CveTarea)" & " and T.CveTareaTipo = " & TAREAPREVENTIVO & " and ( ((UK.KmsAcumulados+(TP.PeriodicidadKms-TP.ToleranciaKms))-U.KmsAcumulados) <= 0 OR " & " (TP.PeriodicidadDias > 0 and datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias - TP.ToleranciaDias,UK.FechaOcurrencia)) <= 0)) " & " and (TP.PeriodicidadDias <> 0 OR TP.PeriodicidadKms <> 0) "
				strSQL = strSQL & " and T.CveTarea in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADTAREAINDIVIDUAL & ")"
				
				' Caso de tareas individuales que aun no estan en Kardex
				strSQL = strSQL & " Union "
				strSQL = strSQL & " SELECT  U.CveUnidad, T.CveTarea,T.Nombre, 0 KmsAcumulados,"
				strSQL = strSQL & " GETDATE() FechaOcurrencia, (TP.PeriodicidadKms  - U.KmsAcumulados)  ToleranciaKms, "
				strSQL = strSQL & " (TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados)  VencidoKms, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias ,GETDATE())) ToleranciaDias, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias + TP.ToleranciaDias,GETDATE())) VencidoDias, 0 CveSubTarea "
				strSQL = strSQL & " FROM    Unidad U (NOLOCK), Tarea T (NOLOCK),TareaPeriodicidad TP(NOLOCK) "
				strSQL = strSQL & " WHERE U.CveUnidad = " & lngCveUnidad & "  and T.CveTarea = TP.CveTarea "
				strSQL = strSQL & " and T.CveTarea in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADTAREAINDIVIDUAL & ") "
				strSQL = strSQL & " AND T.CveTarea NOT IN(SELECT CveTarea FROM UnidadKardex (NOLOCK) WHERE CveUnidad = U.CveUnidad) "
				strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo and T.CveTareaTipo = 1 "
				strSQL = strSQL & " and ( ((TP.PeriodicidadKms-TP.ToleranciaKms)-U.KmsAcumulados) <= 0 OR "
				strSQL = strSQL & " (TP.PeriodicidadDias > 0 and datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias - TP.ToleranciaDias,GETDATE())) <= 0)) "
				strSQL = strSQL & " and (TP.PeriodicidadDias <> 0 OR TP.PeriodicidadKms <> 0) "
				
			Case PREVENTIVOSSERVICIOCADAXKMS
				strSQL = "select U.CveUnidad,T.CveTarea,T.Nombre,UK.KmsAcumulados,UK.FechaOcurrencia, "
				strSQL = strSQL & " (UK.KmsAcumulados+TP.PeriodicidadKms  - U.KmsAcumulados)  ToleranciaKms, "
				strSQL = strSQL & " (UK.KmsAcumulados+TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados)  VencidoKms, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias ,UK.FechaOcurrencia)) ToleranciaDias, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias + TP.ToleranciaDias,UK.FechaOcurrencia)) VencidoDias "
				strSQL = strSQL & " from Unidad U (NOLOCK) , UnidadKardex UK (NOLOCK) , Tarea T (NOLOCK), TareaPeriodicidad TP (NOLOCK) "
				strSQL = strSQL & " where U.CveUnidad = UK.CveUnidad "
				strSQL = strSQL & " and U.CveUnidad = " & lngCveUnidad
				strSQL = strSQL & " and UK.CveTarea = T.CveTarea "
				strSQL = strSQL & " and T.CveTarea = TP.CveTarea "
				strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo "
				strSQL = strSQL & " and UK.FechaOcurrencia= (select MAX(FechaOcurrencia) Ocurrencia from UnidadKardex EK (NOLOCK) " & " where EK.CveUnidad = UK.CveUnidad and EK.CveTarea  = UK.CveTarea)" & " and T.CveTareaTipo = " & TAREAPREVENTIVO & " and ( ((UK.KmsAcumulados+(TP.PeriodicidadKms-TP.ToleranciaKms))-U.KmsAcumulados) <= 0 OR " & " (TP.PeriodicidadDias > 0 and datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias - TP.ToleranciaDias,UK.FechaOcurrencia)) <= 0)) " & " and (TP.PeriodicidadDias <> 0 OR TP.PeriodicidadKms <> 0) "
				strSQL = strSQL & " and T.CveTarea in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADSERVICIOCADAXKMS & ")"
				
				' Caso de tareas individuales que aun no estan en Kardex
				strSQL = strSQL & " Union "
				strSQL = strSQL & " SELECT  U.CveUnidad, T.CveTarea,T.Nombre, 0 KmsAcumulados,"
				strSQL = strSQL & " GETDATE() FechaOcurrencia, (TP.PeriodicidadKms  - U.KmsAcumulados)  ToleranciaKms, "
				strSQL = strSQL & " (TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados)  VencidoKms, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias ,GETDATE())) ToleranciaDias, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias + TP.ToleranciaDias,GETDATE())) VencidoDias "
				strSQL = strSQL & " FROM    Unidad U (NOLOCK), Tarea T (NOLOCK),TareaPeriodicidad TP(NOLOCK) "
				strSQL = strSQL & " WHERE U.CveUnidad = " & lngCveUnidad & "  and T.CveTarea = TP.CveTarea "
				strSQL = strSQL & " and T.CveTarea in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADSERVICIOCADAXKMS & ") "
				strSQL = strSQL & " AND T.CveTarea NOT IN(SELECT CveTarea FROM UnidadKardex (NOLOCK) WHERE CveUnidad = U.CveUnidad) "
				strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo and T.CveTareaTipo = 1 "
				strSQL = strSQL & " and ( ((TP.PeriodicidadKms-TP.ToleranciaKms)-U.KmsAcumulados) <= 0 OR "
				strSQL = strSQL & " (TP.PeriodicidadDias > 0 and datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias - TP.ToleranciaDias,GETDATE())) <= 0)) "
				strSQL = strSQL & " and (TP.PeriodicidadDias <> 0 OR TP.PeriodicidadKms <> 0) "
				
				'strSQL = "EXEC sp_SIMServicioPendiente " & lngCveUnidad
				
			Case PREVENTIVOSAGRUPADOS
				strSQL = "select KmsAcumulados from Unidad where CveUnidad = " & lngCveUnidad
				'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsKms = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
				If rsKms.rdoColumns("KmsAcumulados").Value < 1000000 Then
					strSQL = "select U.CveUnidad,T.CveTarea,T.Nombre,U.KmsAcumulados , "
					strSQL = strSQL & " (TP.PeriodicidadKms  - U.KmsAcumulados)  ToleranciaKms, "
					strSQL = strSQL & " (TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados)  VencidoKms "
					strSQL = strSQL & " from Unidad U (NOLOCK) ,  Tarea T (NOLOCK) , TareaPeriodicidad TP (NOLOCK)"
					strSQL = strSQL & " Where U.CveUnidad = " & lngCveUnidad
					strSQL = strSQL & " and T.CveTarea = TP.CveTarea "
					strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo "
					strSQL = strSQL & " and T.CveTarea  in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADSERVICIO & ")"
					strSQL = strSQL & " and T.CveTarea not in (select CveTarea from UnidadKardex UK (NOLOCK) "
					strSQL = strSQL & " where UK.CveTarea = T.CveTarea and UK.CveUnidad = U.CveUnidad )"
					strSQL = strSQL & " and (TP.PeriodicidadKms - TP.ToleranciaKms) - U.KmsAcumulados <= 0 "
					strSQL = strSQL & " and TP.PeriodicidadKms > 0 "
				Else
					If rsKms.rdoColumns("KmsAcumulados").Value < 2000000 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngKmsPorVuelta = 1000000
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngKmsPorVuelta = 2000000
					End If
					strSQL = "select U.CveUnidad,T.CveTarea,T.Nombre,U.KmsAcumulados , "
					'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSQL = strSQL & " (TP.PeriodicidadKms  - U.KmsAcumulados + " & lngKmsPorVuelta & ")  ToleranciaKms, "
					'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSQL = strSQL & " (TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados + " & lngKmsPorVuelta & " )  VencidoKms "
					strSQL = strSQL & " from Unidad U (NOLOCK) ,  Tarea T (NOLOCK), TareaPeriodicidad TP (NOLOCK) "
					strSQL = strSQL & " Where U.CveUnidad = " & lngCveUnidad
					strSQL = strSQL & " and T.CveTarea = TP.CveTarea "
					strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo "
					strSQL = strSQL & " and T.CveTarea  in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADSERVICIO & ")"
					strSQL = strSQL & " and T.CveTarea not in (select CveTarea from UnidadKardex UK (NOLOCK) "
					strSQL = strSQL & " where UK.CveTarea = T.CveTarea and UK.CveUnidad = U.CveUnidad "
					'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSQL = strSQL & " and UK.KmsAcumulados > " & lngKmsPorVuelta & ") "
					'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSQL = strSQL & " and (TP.PeriodicidadKms - TP.ToleranciaKms) - U.KmsAcumulados + " & lngKmsPorVuelta & " <= 0 "
					strSQL = strSQL & " and TP.PeriodicidadKms > 0 "
				End If
				rsKms.Close()
				
		End Select
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		grsPreventivos = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		Exit Sub
		
Err_VerificaPreventivos: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Verificar Preventivos " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	
	Public Sub VerificaPreventivosSybase(ByRef lngCveUnidad As Integer, ByRef intTipoPreventivos As Short)
		'*******************************************************************************************
		'  Selecciona las Tareas que ya estan vencidas, ya se por Fecha o por Kilometraje
		'  Obtiene la periodicidad de Dias y de Kilometros de la Tabla TareaPeriodicidad
		'  Obtiene el Kilometraje Acumulado de la Tabla de Unidad
		'  Obtiene la Fecha de Ocurrencia de la Tarea utilizando la siguiente Formula :
		'    (((Periodicidad - Tolerancia) + FechaOcurrencia) - el Dia de Hoy) <= 0 , Si es menor aparecera en el Programa
		'  Obtiene el Kilometraje de la Tarea utilizando la siguiente Formula :
		'    ((KmsAcumulados + (Periodicidad - Tolerancia)) - KmsAcum de la Unidad) <= 0 , Si es menor aparecera en el Programa
		'  Donde la Tarea sea la Ultima que se haya efectuado y este registrada en el Kardex.
		'  Donde la Periodicida sea Diferente a 0
		'
		'  Parametros:
		'        lngCveUnidad .- # de Unidad a la que se desean revisar los preventivos
		'        intTipoPreventivos .- Me indica si los preventivos son individuales o agrupados
		'
		'  ADAPTADO PARA SINTAXIS DE SYBASE
		'*******************************************************************************************
		On Error GoTo Err_VerificaPreventivos
		
		Dim strSQL As String
		Dim strFormato As String
		Dim rsKms As RDO.rdoResultset
		Dim lngKmsPorVuelta As Object
		
		Select Case intTipoPreventivos
			
			Case PREVENTIVOSINDIVIDUALES
				strSQL = "select U.CveUnidad,T.CveTarea,T.Nombre,UK.KmsAcumulados,UK.FechaOcurrencia, "
				strSQL = strSQL & " (UK.KmsAcumulados+TP.PeriodicidadKms  - U.KmsAcumulados)  ToleranciaKms, "
				strSQL = strSQL & " (UK.KmsAcumulados+TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados)  VencidoKms, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias ,UK.FechaOcurrencia)) ToleranciaDias, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias + TP.ToleranciaDias,UK.FechaOcurrencia)) VencidoDias "
				strSQL = strSQL & " from Unidad U  , UnidadKardex UK  , Tarea T , TareaPeriodicidad TP "
				strSQL = strSQL & " where U.CveUnidad = UK.CveUnidad "
				strSQL = strSQL & " and U.CveUnidad = " & lngCveUnidad
				strSQL = strSQL & " and UK.CveTarea = T.CveTarea "
				strSQL = strSQL & " and T.CveTarea = TP.CveTarea "
				strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo "
				strSQL = strSQL & " and UK.FechaOcurrencia= (select MAX(FechaOcurrencia) Ocurrencia from UnidadKardex EK " & " where EK.CveUnidad = UK.CveUnidad and EK.CveTarea  = UK.CveTarea)" & " and T.CveTareaTipo = " & TAREAPREVENTIVO & " and ( ((UK.KmsAcumulados+(TP.PeriodicidadKms-TP.ToleranciaKms))-U.KmsAcumulados) <= 0 OR " & " (TP.PeriodicidadDias > 0 and datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias - TP.ToleranciaDias,UK.FechaOcurrencia)) <= 0)) " & " and (TP.PeriodicidadDias <> 0 OR TP.PeriodicidadKms <> 0) "
				strSQL = strSQL & " and T.CveTarea in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADTAREAINDIVIDUAL & ")"
				
				' Caso de tareas individuales que aun no estan en Kardex
				strSQL = strSQL & " Union "
				strSQL = strSQL & " SELECT  U.CveUnidad, T.CveTarea,T.Nombre, 0 KmsAcumulados,"
				strSQL = strSQL & " GETDATE() FechaOcurrencia, (TP.PeriodicidadKms  - U.KmsAcumulados)  ToleranciaKms, "
				strSQL = strSQL & " (TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados)  VencidoKms, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias ,GETDATE())) ToleranciaDias, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias + TP.ToleranciaDias,GETDATE())) VencidoDias "
				strSQL = strSQL & " FROM    Unidad U , Tarea T ,TareaPeriodicidad TP "
				strSQL = strSQL & " WHERE U.CveUnidad = " & lngCveUnidad & "  and T.CveTarea = TP.CveTarea "
				strSQL = strSQL & " and T.CveTarea in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADTAREAINDIVIDUAL & ") "
				strSQL = strSQL & " AND T.CveTarea NOT IN(SELECT CveTarea FROM UnidadKardex  WHERE CveUnidad = U.CveUnidad) "
				strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo and T.CveTareaTipo = 1 "
				strSQL = strSQL & " and ( ((TP.PeriodicidadKms-TP.ToleranciaKms)-U.KmsAcumulados) <= 0 OR "
				strSQL = strSQL & " (TP.PeriodicidadDias > 0 and datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias - TP.ToleranciaDias,GETDATE())) <= 0)) "
				strSQL = strSQL & " and (TP.PeriodicidadDias <> 0 OR TP.PeriodicidadKms <> 0) "
				
				
			Case PREVENTIVOSSERVICIOCADAXKMS
				strSQL = "select U.CveUnidad,T.CveTarea,T.Nombre,UK.KmsAcumulados,UK.FechaOcurrencia, "
				strSQL = strSQL & " (UK.KmsAcumulados+TP.PeriodicidadKms  - U.KmsAcumulados)  ToleranciaKms, "
				strSQL = strSQL & " (UK.KmsAcumulados+TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados)  VencidoKms, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias ,UK.FechaOcurrencia)) ToleranciaDias, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias + TP.ToleranciaDias,UK.FechaOcurrencia)) VencidoDias "
				strSQL = strSQL & " from Unidad U  , UnidadKardex UK  , Tarea T , TareaPeriodicidad TP  "
				strSQL = strSQL & " where U.CveUnidad = UK.CveUnidad "
				strSQL = strSQL & " and U.CveUnidad = " & lngCveUnidad
				strSQL = strSQL & " and UK.CveTarea = T.CveTarea "
				strSQL = strSQL & " and T.CveTarea = TP.CveTarea "
				strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo "
				strSQL = strSQL & " and UK.FechaOcurrencia= (select MAX(FechaOcurrencia) Ocurrencia from UnidadKardex EK  " & " where EK.CveUnidad = UK.CveUnidad and EK.CveTarea  = UK.CveTarea)" & " and T.CveTareaTipo = " & TAREAPREVENTIVO & " and ( ((UK.KmsAcumulados+(TP.PeriodicidadKms-TP.ToleranciaKms))-U.KmsAcumulados) <= 0 OR " & " (TP.PeriodicidadDias > 0 and datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias - TP.ToleranciaDias,UK.FechaOcurrencia)) <= 0)) " & " and (TP.PeriodicidadDias <> 0 OR TP.PeriodicidadKms <> 0) "
				strSQL = strSQL & " and T.CveTarea in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADSERVICIOCADAXKMS & ")"
				
				' Caso de tareas individuales que aun no estan en Kardex
				strSQL = strSQL & " Union "
				strSQL = strSQL & " SELECT  U.CveUnidad, T.CveTarea,T.Nombre, 0 KmsAcumulados,"
				strSQL = strSQL & " GETDATE() FechaOcurrencia, (TP.PeriodicidadKms  - U.KmsAcumulados)  ToleranciaKms, "
				strSQL = strSQL & " (TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados)  VencidoKms, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias ,GETDATE())) ToleranciaDias, "
				strSQL = strSQL & " datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias + TP.ToleranciaDias,GETDATE())) VencidoDias "
				strSQL = strSQL & " FROM    Unidad U , Tarea T ,TareaPeriodicidad TP "
				strSQL = strSQL & " WHERE U.CveUnidad = " & lngCveUnidad & "  and T.CveTarea = TP.CveTarea "
				strSQL = strSQL & " and T.CveTarea in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADSERVICIOCADAXKMS & ") "
				strSQL = strSQL & " AND T.CveTarea NOT IN(SELECT CveTarea FROM UnidadKardex  WHERE CveUnidad = U.CveUnidad) "
				strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo and T.CveTareaTipo = 1 "
				strSQL = strSQL & " and ( ((TP.PeriodicidadKms-TP.ToleranciaKms)-U.KmsAcumulados) <= 0 OR "
				strSQL = strSQL & " (TP.PeriodicidadDias > 0 and datediff(day,getdate(),dateadd(day,TP.PeriodicidadDias - TP.ToleranciaDias,GETDATE())) <= 0)) "
				strSQL = strSQL & " and (TP.PeriodicidadDias <> 0 OR TP.PeriodicidadKms <> 0) "
				
				
			Case PREVENTIVOSAGRUPADOS
				strSQL = "select KmsAcumulados from Unidad where CveUnidad = " & lngCveUnidad
				'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsKms = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
				If rsKms.rdoColumns("KmsAcumulados").Value < 1000000 Then
					strSQL = "select U.CveUnidad,T.CveTarea,T.Nombre,U.KmsAcumulados , "
					strSQL = strSQL & " (TP.PeriodicidadKms  - U.KmsAcumulados)  ToleranciaKms, "
					strSQL = strSQL & " (TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados)  VencidoKms "
					strSQL = strSQL & " from Unidad U  ,  Tarea T  , TareaPeriodicidad TP "
					strSQL = strSQL & " Where U.CveUnidad = " & lngCveUnidad
					strSQL = strSQL & " and T.CveTarea = TP.CveTarea "
					strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo "
					strSQL = strSQL & " and T.CveTarea  in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADSERVICIO & ")"
					strSQL = strSQL & " and T.CveTarea not in (select CveTarea from UnidadKardex UK  "
					strSQL = strSQL & " where UK.CveTarea = T.CveTarea and UK.CveUnidad = U.CveUnidad )"
					strSQL = strSQL & " and (TP.PeriodicidadKms - TP.ToleranciaKms) - U.KmsAcumulados <= 0 "
					strSQL = strSQL & " and TP.PeriodicidadKms > 0 "
				Else
					If rsKms.rdoColumns("KmsAcumulados").Value < 2000000 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngKmsPorVuelta = 1000000
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngKmsPorVuelta = 2000000
					End If
					strSQL = "select U.CveUnidad,T.CveTarea,T.Nombre,U.KmsAcumulados , "
					'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSQL = strSQL & " (TP.PeriodicidadKms  - U.KmsAcumulados + " & lngKmsPorVuelta & ")  ToleranciaKms, "
					'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSQL = strSQL & " (TP.PeriodicidadKms + TP.ToleranciaKms - U.KmsAcumulados + " & lngKmsPorVuelta & " )  VencidoKms "
					strSQL = strSQL & " from Unidad U  ,  Tarea T , TareaPeriodicidad TP  "
					strSQL = strSQL & " Where U.CveUnidad = " & lngCveUnidad
					strSQL = strSQL & " and T.CveTarea = TP.CveTarea "
					strSQL = strSQL & " and U.CveUnidadTipo = TP.CveUnidadTipo "
					strSQL = strSQL & " and T.CveTarea  in (select CveTarea from Tarea where Baja = 0 AND CvePeriodicidadTipo = " & PERIODICIDADSERVICIO & ")"
					strSQL = strSQL & " and T.CveTarea not in (select CveTarea from UnidadKardex UK  "
					strSQL = strSQL & " where UK.CveTarea = T.CveTarea and UK.CveUnidad = U.CveUnidad "
					'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSQL = strSQL & " and UK.KmsAcumulados > " & lngKmsPorVuelta & ") "
					'UPGRADE_WARNING: Couldn't resolve default property of object lngKmsPorVuelta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strSQL = strSQL & " and (TP.PeriodicidadKms - TP.ToleranciaKms) - U.KmsAcumulados + " & lngKmsPorVuelta & " <= 0 "
					strSQL = strSQL & " and TP.PeriodicidadKms > 0 "
				End If
				rsKms.Close()
				
		End Select
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		grsPreventivos = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		Exit Sub
		
Err_VerificaPreventivos: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Verificar Preventivos " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	
	Public Function ListaMultiselect(ByRef Lista As System.Windows.Forms.ListBox, ByRef Texto As String) As String
		'-----------------------------------------------------------------------
		'Esta rutina concatena todos los elementos del arreglo itemdata para   -
		'formar una expresión lógica OR
		'-----------------------------------------------------------------------
		Dim i As Short
		Dim blnPrimero As Boolean
		
		ListaMultiselect = ""
		If Lista.SelectedItems.Count > 0 Then
			For i = 0 To Lista.Items.Count - 1
				If Not blnPrimero Then
					If Lista.GetSelected(i) Then
						ListaMultiselect = "(" & Texto & " = " & VB6.GetItemData(Lista, i)
						blnPrimero = True
					End If
				Else
					If Lista.GetSelected(i) Then
						ListaMultiselect = ListaMultiselect & " OR " & Texto & " = " & VB6.GetItemData(Lista, i)
						blnPrimero = True
					End If
				End If
			Next i
			ListaMultiselect = ListaMultiselect & ")"
		Else
			Beep()
			MsgBox("Seleccione por lo menos un elemento de la lista")
			Exit Function
		End If
		
	End Function
	
	Public Function ListaMultiSelectReporteSQL(ByRef Lista As System.Windows.Forms.ListBox, ByRef Texto As String) As String
		'-----------------------------------------------------------------------
		'Esta rutina concatena todos los elementos del arreglo itemdata para   -
		'formar una expresión lógica OR
		'-----------------------------------------------------------------------
		Dim i As Short
		Dim blnPrimero As Boolean
		
		ListaMultiSelectReporteSQL = ""
		If Lista.SelectedItems.Count > 0 Then
			For i = 0 To Lista.Items.Count - 1
				If Not blnPrimero Then
					If Lista.GetSelected(i) Then
						ListaMultiSelectReporteSQL = Texto & " = " & VB6.GetItemData(Lista, i)
						blnPrimero = True
					End If
				Else
					If Lista.GetSelected(i) Then
						ListaMultiSelectReporteSQL = ListaMultiSelectReporteSQL & " OR " & Texto & " = " & VB6.GetItemData(Lista, i)
						blnPrimero = True
					End If
				End If
			Next i
		Else
			Beep()
			MsgBox("Seleccione por lo menos un elemento de la lista")
			Exit Function
		End If
		
	End Function
	
	Public Function ObtieneCuenta(ByRef strCuenta As String) As Object
		
		Dim intPosicion As Short
		
		intPosicion = InStr(1, strCuenta, "/")
		ObtieneCuenta = Trim(Mid(strCuenta, 1, intPosicion - 1))
		
	End Function
	
	Public Function ObtieneCuentaDeLista(ByRef Lista As System.Windows.Forms.ListBox, ByRef Texto As String) As String
		'-----------------------------------------------------------------------
		'Esta rutina concatena todos los elementos del arreglo itemdata para   -
		'formar una expresión lógica OR
		'-----------------------------------------------------------------------
		Dim i As Short
		Dim blnPrimero As Boolean
		Dim strCuenta As String
		Dim intPosicion As Short
		
		ObtieneCuentaDeLista = ""
		If Lista.SelectedItems.Count > 0 Then
			For i = 0 To Lista.Items.Count - 1
				If Not blnPrimero Then
					If Lista.GetSelected(i) Then
						intPosicion = InStr(1, Lista.Text, "/")
						strCuenta = Trim(Mid(Lista.Text, 1, intPosicion - 1))
						ObtieneCuentaDeLista = "(" & Texto & "=" & strCuenta
						blnPrimero = True
					End If
				Else
					If Lista.GetSelected(i) Then
						intPosicion = InStr(1, Lista.Text, "/")
						strCuenta = Trim(Mid(Lista.Text, 1, intPosicion - 1))
						ObtieneCuentaDeLista = ObtieneCuentaDeLista & " OR " & Texto & "=" & strCuenta
						blnPrimero = True
					End If
				End If
			Next i
			ObtieneCuentaDeLista = ObtieneCuentaDeLista & ")"
		Else
			Beep()
			MsgBox("Seleccione por lo menos un elemento de la lista")
			Exit Function
		End If
		
	End Function
	
	Public Sub LlenaSelectorCuenta(ByRef strSQL As String, ByRef NombreControl As System.Windows.Forms.Control)
		'*----------------------------------------------------------------------*
		'* Entrada: - Un string de SQL "select ",                               *
		'*          - el nombre del COMBO                                       *
		'*                                                                      *
		'*----------------------------------------------------------------------*
		
		On Error GoTo LlenaVariosselectores_Error
		
		Dim ctrselector As System.Windows.Forms.Control ' Control que se actualizara con los datos del rsfLlena
		Dim dsfLlenaControl As RDO.rdoResultset ' Recordset de los datos a insertar en combos ó Listbox
		Dim bytIndice As Byte ' Indice del Arreglo de Strings
		Dim blnAbierto As Boolean ' Bandera del resultado, para saber si se abrio con exito el RS
		Dim blnEncontro As Boolean ' Para saber si el Combo Existe en la forma
		Dim i As Byte
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dsfLlenaControl = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
		blnAbierto = True ' Se abrio OK el ResultSet
		
		bytIndice = 0
		
		While Not dsfLlenaControl.EOF
			blnEncontro = False
			'UPGRADE_WARNING: Couldn't resolve default property of object NombreControl.Clear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			NombreControl.Clear() 'Para limpiar los combos
			Do Until dsfLlenaControl.EOF
				'UPGRADE_WARNING: Couldn't resolve default property of object NombreControl.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NombreControl.AddItem(Trim(dsfLlenaControl.rdoColumns.Item(1).Value)) ' Agrega la descripcion
				dsfLlenaControl.MoveNext()
			Loop 
			'UPGRADE_WARNING: Couldn't resolve default property of object NombreControl.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If NombreControl.ListCount > 0 Then ' Posiciona en el primer indice el selector
				'UPGRADE_WARNING: Couldn't resolve default property of object NombreControl.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NombreControl.ListIndex = -1 ' El -1 es para que no genere un Click
				NombreControl.Refresh()
			End If
		End While
		
		GoTo LlenaVariosselectores_Exit
		
LlenaVariosselectores_Error: 
		If Err.Number = 344 Then
			MsgBox("Debes especificar el combo como un array", MsgBoxStyle.Exclamation)
		Else
			MsgBox("Error : " & ErrorToString(Err.Number))
		End If
		
		Beep()
		Resume LlenaVariosselectores_Exit
		
LlenaVariosselectores_Exit: 
		If blnAbierto Then
			dsfLlenaControl.Close()
		End If
		
	End Sub
	
	Private Function EliminaSimbolos(ByVal vstrTexto As String) As Double
		'Elimina caracteres no numericos
		
		Dim localidad As Integer
		
		localidad = InStr(vstrTexto, "%")
		Do Until localidad = 0
			vstrTexto = Mid(vstrTexto, 1, localidad - 1) & Mid(vstrTexto, localidad + 1)
			localidad = InStr(vstrTexto, "%")
		Loop 
		
		localidad = InStr(vstrTexto, "$")
		Do Until localidad = 0
			vstrTexto = Mid(vstrTexto, 1, localidad - 1) & Mid(vstrTexto, localidad + 1)
			localidad = InStr(vstrTexto, "$")
		Loop 
		
		localidad = InStr(vstrTexto, ",")
		Do Until localidad = 0
			vstrTexto = Mid(vstrTexto, 1, localidad - 1) & Mid(vstrTexto, localidad + 1)
			localidad = InStr(vstrTexto, ",")
		Loop 
		EliminaSimbolos = Val(vstrTexto)
		
	End Function
	
	Sub ObtieneResolucion(ByRef rsngAncho As Single, ByRef rsngAlto As Single)
		
		Dim blnFunciona As Boolean
		Dim lngContadorGraficos As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object lngContadorGraficos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngContadorGraficos = 0
		blnFunciona = True
		Do Until (blnFunciona = False)
			'UPGRADE_WARNING: Couldn't resolve default property of object mdvmDispositivoGrafico. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object lngContadorGraficos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			blnFunciona = EnumDisplaySettings(0, lngContadorGraficos, mdvmDispositivoGrafico)
			'UPGRADE_WARNING: Couldn't resolve default property of object lngContadorGraficos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lngContadorGraficos = lngContadorGraficos + 1
		Loop 
		
		mdvmDispositivoGrafico.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
		
		rsngAncho = mdvmDispositivoGrafico.dmPelsWidth
		rsngAlto = mdvmDispositivoGrafico.dmPelsHeight
		
	End Sub
	
	Sub CambiaResolucion(ByVal vsngAncho As Single, ByVal vsngAlto As Single)
		
		Dim blnFunciona As Boolean
		Dim lngContadorGraficos As Integer
		Dim blnFuncionaCambio As Boolean
		
		lngContadorGraficos = 0
		blnFunciona = True
		Do Until (blnFunciona = False)
			'UPGRADE_WARNING: Couldn't resolve default property of object mdvmDispositivoGrafico. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			blnFunciona = EnumDisplaySettings(0, lngContadorGraficos, mdvmDispositivoGrafico)
			lngContadorGraficos = lngContadorGraficos + 1
		Loop 
		
		mdvmDispositivoGrafico.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
		
		mdvmDispositivoGrafico.dmPelsWidth = vsngAncho
		mdvmDispositivoGrafico.dmPelsHeight = vsngAlto
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mdvmDispositivoGrafico. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		blnFuncionaCambio = ChangeDisplaySettings(mdvmDispositivoGrafico, 0)
	End Sub
	
	Sub SelListaEliminar(ByRef Lista As System.Windows.Forms.Control)
		'**************************************************************************
		' Rutina que elimina los renglones seleccionados de la lista.
		' Entrada .-
		'   Lista .- List sobre el que se trabaja.
		'**************************************************************************
		Dim intRen As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Lista.SelCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Lista.SelCount > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Lista.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For intRen = 0 To Lista.ListCount - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Lista.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If intRen < Lista.ListCount Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Lista.Selected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Lista.Selected(intRen) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Lista.RemoveItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Lista.RemoveItem(intRen)
						intRen = -1
					End If
				End If
			Next 
		End If
		
	End Sub
	
	Sub CopiarCombo(ByRef rcboOrigen As System.Windows.Forms.ComboBox, ByRef rcboDestino As System.Windows.Forms.ComboBox)
		'**************************************************************************
		' Rutina que copia los renglones seleccionados en la ListaOrigen a la
		' ListaDestino.
		' Entrada .-
		'   ListaOrigen.
		'   ListaDestino.
		'**************************************************************************
		Dim intRen As Short
		
		If rcboOrigen.Items.Count > 0 Then
			For intRen = 0 To rcboOrigen.Items.Count - 1
				rcboDestino.Items.Add(New VB6.ListBoxItem(VB6.GetItemString(rcboOrigen, intRen), VB6.GetItemData(rcboOrigen, intRen)))
			Next 
		End If
		
	End Sub
	Sub SelListaCopiar(ByRef rlstOrigen As System.Windows.Forms.ListBox, ByRef rlstDestino As System.Windows.Forms.ListBox)
		'**************************************************************************
		' Rutina que copia los renglones seleccionados en la ListaOrigen a la
		' ListaDestino.
		' Entrada .-
		'   ListaOrigen.
		'   ListaDestino.
		'**************************************************************************
		Dim intRen As Short
		
		If rlstOrigen.SelectedItems.Count > 0 Then
			For intRen = 0 To rlstOrigen.Items.Count - 1
				If rlstOrigen.GetSelected(intRen) Then
					rlstDestino.Items.Add(New VB6.ListBoxItem(VB6.GetItemString(rlstOrigen, intRen), VB6.GetItemData(rlstOrigen, intRen)))
				End If
			Next 
		End If
		
	End Sub
	Public Sub LlenaVariosSelectores(ByRef strSQL As String, ByRef strCombos As Object, ByRef frmForma As System.Windows.Forms.Form)
		' ---------------------------------------
		' Entrada: Un string de SQL "select .... select ..", el nombre de los combos ó
		'       Listas en un arreglo de strings "ARRAY("COMBO1","COMBO2")" y el nombre de
		'       la Forma en donde se aplicara la Sub.
		' Salida:  Llena varios combos y listbox con los datos de las tablas especificadas
		'          en el estatuto sql
		'------------------------------------------
		
		On Error GoTo LlenaVariosselectores_Error
		
		Dim ctrselector As System.Windows.Forms.Control 'Control que se actualizara con los datos del rsfLlena
		Dim rsfLlenaControl As RDO.rdoResultset 'rdoResultset de los datos a insertar en combos ó Listbox
		Dim bytIndice As Byte 'Indice del Arreglo de Strings
		Dim blnAbierto As Boolean 'Bandera del resultado, para saber si se abrio con exito el RS
		Dim blnEncontro As Boolean 'Para saber si el Combo Existe en la forma
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.StillExecuting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Do Until Not gcn.StillExecuting
			
		Loop 
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsfLlenaControl = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly)
		blnAbierto = True ' Se abrio OK el rdoResultset
		
		bytIndice = 1
		Do 
			blnEncontro = False
			For	Each ctrselector In frmForma.Controls ' Ciclo para localizar los controles recibidos
				'UPGRADE_WARNING: Couldn't resolve default property of object strCombos(bytIndice). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ctrselector.Name = strCombos(bytIndice) Then 'Si encontro el Control
					blnEncontro = True
					'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.Clear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctrselector.Clear() 'Para limpiar los combos
					ctrselector.Refresh()
					Exit For
				End If
			Next ctrselector
			
			While Not rsfLlenaControl.EOF And blnEncontro
				Do Until rsfLlenaControl.EOF
					'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctrselector.AddItem(Trim(rsfLlenaControl.rdoColumns.Item(1).Value)) ' Agrega la descripcion
					'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.NewIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctrselector.ItemData(ctrselector.NewIndex) = rsfLlenaControl.rdoColumns.Item(0).Value
					rsfLlenaControl.MoveNext()
				Loop 
				'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ctrselector.ListCount > 0 Then ' Posiciona en el primer indice el selector
					'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctrselector.ListIndex = -1 ' El -1 es para que no genere un Click
				End If
				ctrselector.Refresh()
			End While
			If Not blnEncontro Then
				Beep()
				'UPGRADE_WARNING: Couldn't resolve default property of object strCombos(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MsgBox("El selector " & strCombos(bytIndice) & " No Existe", MsgBoxStyle.Exclamation)
				GoTo LlenaVariosselectores_Exit
			End If
			bytIndice = bytIndice + 1
		Loop While rsfLlenaControl.MoreResults ' si existen mas rdoResultsets dentro del rsfLlena
		
		GoTo LlenaVariosselectores_Exit
		
LlenaVariosselectores_Error: 
		Dim strmsg As String 'String del Error
		Dim lngIndice As Integer 'Indice del Error de RDO
		
		Select Case Err.Number
			'Case 9
			'    strmsg = "No se ha especificado en las Declaraciones de la Forma " & frmForma.Name & "." & _
			''        "El estandar 'OPTION BASE 1' ó el strSQL no concuerda con la cantidad de combos que se desean llenar."
			Case 344
				strmsg = "El combo se debe especificar como un array."
			Case 40002
				For lngIndice = 0 To RDOrdoEngine_definst.rdoErrors.Count - 1
					strmsg = strmsg & RDOrdoEngine_definst.rdoErrors(lngIndice).Description & Chr(System.Windows.Forms.Keys.Return)
				Next lngIndice
				RDOrdoEngine_definst.rdoErrors.Clear()
			Case Else
				strmsg = Err.Number & " " & ErrorToString()
		End Select
		Err.Clear()
		
		MsgBox("Ocurrió un error al Llenar Varios selectores :" & Chr(System.Windows.Forms.Keys.Return) & strmsg, MsgBoxStyle.Exclamation)
		Resume LlenaVariosselectores_Exit
		Resume Next
LlenaVariosselectores_Exit: 
		If blnAbierto Then
			rsfLlenaControl.Close()
			blnAbierto = False
		End If
		
	End Sub
	
	Public Function DateOfFirstDayofWeek(ByRef intCurrentDayofWeek As Short, ByRef WhichDate As Date) As Date
		
		On Error GoTo error_handler
		
		DateOfFirstDayofWeek = DateAdd(Microsoft.VisualBasic.DateInterval.Day, intCurrentDayofWeek * (-1) + 1, WhichDate)
		
		Exit Function
error_handler: 
		DateOfFirstDayofWeek = Today
		
	End Function
	Public Sub LlenaSelectoresAlmacen(ByRef strSQL As String, ByRef strCombos As Object, ByRef frmForma As System.Windows.Forms.Form)
		' ---------------------------------------
		' Entrada: Un string de SQL "select .... select ..", el nombre de los combos ó
		'       Listas en un arreglo de strings "ARRAY("COMBO1","COMBO2")" y el nombre de
		'       la Forma en donde se aplicara la Sub.
		' Salida:  Llena varios combos y listbox con los datos de las tablas especificadas
		'          en el estatuto sql
		'------------------------------------------
		
		On Error GoTo LlenaVariosselectores_Error
		
		Dim ctrselector As System.Windows.Forms.Control 'Control que se actualizara con los datos del rsfLlena
		Dim rsfLlenaControl As RDO.rdoResultset 'rdoResultset de los datos a insertar en combos ó Listbox
		Dim bytIndice As Byte 'Indice del Arreglo de Strings
		Dim blnAbierto As Boolean 'Bandera del resultado, para saber si se abrio con exito el RS
		Dim blnEncontro As Boolean 'Para saber si el Combo Existe en la forma
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnAlmacen.StillExecuting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Do Until Not gcnAlmacen.StillExecuting
			
		Loop 
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnAlmacen.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsfLlenaControl = gcnAlmacen.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly)
		blnAbierto = True ' Se abrio OK el rdoResultset
		
		bytIndice = 1
		Do 
			blnEncontro = False
			For	Each ctrselector In frmForma.Controls ' Ciclo para localizar los controles recibidos
				'UPGRADE_WARNING: Couldn't resolve default property of object strCombos(bytIndice). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ctrselector.Name = strCombos(bytIndice) Then 'Si encontro el Control
					blnEncontro = True
					'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.Clear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctrselector.Clear() 'Para limpiar los combos
					ctrselector.Refresh()
					Exit For
				End If
			Next ctrselector
			
			While Not rsfLlenaControl.EOF And blnEncontro
				Do Until rsfLlenaControl.EOF
					'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctrselector.AddItem(Trim(rsfLlenaControl.rdoColumns.Item(1).Value)) ' Agrega la descripcion
					'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.NewIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctrselector.ItemData(ctrselector.NewIndex) = rsfLlenaControl.rdoColumns.Item(0).Value
					rsfLlenaControl.MoveNext()
				Loop 
				'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.ListCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ctrselector.ListCount > 0 Then ' Posiciona en el primer indice el selector
					'UPGRADE_WARNING: Couldn't resolve default property of object ctrselector.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctrselector.ListIndex = -1 ' El -1 es para que no genere un Click
				End If
				ctrselector.Refresh()
			End While
			If Not blnEncontro Then
				Beep()
				'UPGRADE_WARNING: Couldn't resolve default property of object strCombos(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MsgBox("El selector " & strCombos(bytIndice) & " No Existe", MsgBoxStyle.Exclamation)
				GoTo LlenaVariosselectores_Exit
			End If
			bytIndice = bytIndice + 1
		Loop While rsfLlenaControl.MoreResults ' si existen mas rdoResultsets dentro del rsfLlena
		
		GoTo LlenaVariosselectores_Exit
		
LlenaVariosselectores_Error: 
		Dim strmsg As String 'String del Error
		Dim lngIndice As Integer 'Indice del Error de RDO
		
		Select Case Err.Number
			'Case 9
			'    strmsg = "No se ha especificado en las Declaraciones de la Forma " & frmForma.Name & "." & _
			''        "El estandar 'OPTION BASE 1' ó el strSQL no concuerda con la cantidad de combos que se desean llenar."
			Case 344
				strmsg = "El combo se debe especificar como un array."
			Case 40002
				For lngIndice = 0 To RDOrdoEngine_definst.rdoErrors.Count - 1
					strmsg = strmsg & RDOrdoEngine_definst.rdoErrors(lngIndice).Description & Chr(System.Windows.Forms.Keys.Return)
				Next lngIndice
				RDOrdoEngine_definst.rdoErrors.Clear()
			Case Else
				strmsg = Err.Number & " " & ErrorToString()
		End Select
		Err.Clear()
		
		MsgBox("Ocurrió un error al Llenar Varios selectores :" & Chr(System.Windows.Forms.Keys.Return) & strmsg, MsgBoxStyle.Exclamation)
		Resume LlenaVariosselectores_Exit
		
LlenaVariosselectores_Exit: 
		If blnAbierto Then rsfLlenaControl.Close()
		
	End Sub
	
	Public Function ObtieneConsecutivo(ByVal vstrNomTabla As String, ByVal vstrNomCampo As String, Optional ByVal vstrWhere As String = "") As Integer
		'**********************************************************************
		' Función que obtiene el número consecutivo de una tabla.
		' El Consecutivo debe estar compuesto por un prefijo y un sufijo,
		' ambos numéricos.
		' Entrada .-
		'   vstrNomTabla .- Nombre de la tabla sobre la cual se va a buscar
		'                   el siguiente consecutivo.
		'   vstrNomCampo .- Nombre del campo sobre la cual se va a buscar
		'                   el siguiente consecutivo.
		'   vstrWhere.-     Condición para la búsqueda
		' Salida .-
		'   0 : Error de ejecución.
		'   Mayor a 0 : Consecutivo.
		'**********************************************************************
		
		Dim blnCreado As Boolean
		Dim rsTemp As RDO.rdoResultset
		Dim strSQL As String
		
		On Error GoTo Er_ObtieneConsComp
		
		blnCreado = False
		ObtieneConsecutivo = 0
		
		'Determinación del Formato Deseado.
		strSQL = "select MAX(" & vstrNomCampo & ") Siguiente " & "from " & vstrNomTabla & " " & vstrWhere
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsTemp = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
		
		blnCreado = True
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not rsTemp.EOF And Not IsDbNull(rsTemp.rdoColumns(0).Value) Then
			ObtieneConsecutivo = rsTemp.rdoColumns(0).Value + 1
		Else
			ObtieneConsecutivo = 1
		End If
		
Ex_ObtieneConsComp: 
		If blnCreado Then rsTemp.Close()
		Exit Function
		
Er_ObtieneConsComp: 
		MsgBox("Error al Obtener Consecutivo" & ErrorToString(), MsgBoxStyle.Critical)
		
	End Function
	
	Public Function ListaMultiselectCuenta(ByRef Lista As System.Windows.Forms.ListBox, ByRef Texto As String) As String
		'-----------------------------------------------------------------------
		'Esta rutina concatena todos los elementos del arreglo itemdata para   -
		'formar una expresión lógica OR
		'-----------------------------------------------------------------------
		Dim i As Short
		Dim blnPrimero As Boolean
		
		ListaMultiselectCuenta = ""
		If Lista.SelectedItems.Count > 0 Then
			For i = 0 To Lista.Items.Count - 1
				If Not blnPrimero Then
					If Lista.GetSelected(i) Then
						ListaMultiselectCuenta = "(" & Texto & "=" & gCuentas(i).Cuenta
						blnPrimero = True
					End If
				Else
					If Lista.GetSelected(i) Then
						ListaMultiselectCuenta = ListaMultiselectCuenta & " OR " & Texto & "=" & gCuentas(i).Cuenta
						blnPrimero = True
					End If
				End If
			Next i
			ListaMultiselectCuenta = ListaMultiselectCuenta & ")"
		Else
			Beep()
			MsgBox("Seleccione por lo menos un elemento de la lista")
			Exit Function
		End If
		
	End Function
	
	Sub LimpiaBloque(ByRef ctlSpread As System.Windows.Forms.Control, ByVal vintRen As Short, ByVal vintCol As Short, ByVal vintRen2 As Short, ByVal vintCol2 As Short)
		'*************************************************************************
		' Rutina que limpia los DATOS de de las celdas tipo TEXTO un spread
		' en el rango dado.
		' Entrada :
		'   ctlSpread .- Spread
		'   vintRen .- Renglón inicio
		'   vintCol .- Columna inicio
		'   vintRen2 .- Renglon final
		'   vintCol2 .- Columna final
		'*************************************************************************
		Dim i As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.ClearRange. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlSpread.ClearRange(vintCol, vintRen, vintCol2, vintRen2, True)
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.SetActiveCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlSpread.SetActiveCell(vintCol, vintRen)
		
		For i = vintRen To vintRen2
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ctlSpread.Col = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ctlSpread.Row = i
			
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.CellType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If ctlSpread.CellType = FPSpread.CellTypeConstants.CellTypePicture Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.DeleteRows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ctlSpread.DeleteRows(i, 1)
			End If
			
		Next i
	End Sub
	Sub LimpiaBloqueVAnt(ByRef ctlSpread As System.Windows.Forms.Control, ByVal vintRen As Short, ByVal vintCol As Short, ByVal vintRen2 As Short, ByVal vintCol2 As Short)
		'*************************************************************************
		' Rutina que limpia los DATOS de de las celdas tipo TEXTO un spread
		' en el rango dado.
		' Entrada :
		'   ctlSpread .- Spread
		'   vintRen .- Renglón inicio
		'   vintCol .- Columna inicio
		'   vintRen2 .- Renglon final
		'   vintCol2 .- Columna final
		'*************************************************************************
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlSpread.Row = vintRen
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlSpread.Col = vintCol
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.Row2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlSpread.Row2 = vintRen2
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.Col2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlSpread.Col2 = vintCol2
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.BlockMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlSpread.BlockMode = True
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlSpread.Action = SS_ACTION_CLEAR
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.BlockMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlSpread.BlockMode = False
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlSpread.SetActiveCell. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlSpread.SetActiveCell(vintCol, vintRen)
		
	End Sub
	
	
	Public Function ValidaLlaveDuplicadaBase(ByRef strNombreTabla As String, ByRef strNombreCampo As String, ByRef strValorABuscar As String, ByRef strMensaje As String, ByRef intBase As Object) As Object
		'-----------------------------------------------------------------------------------
		' Valida si la llave con la que se insertara no existe
		'
		'   Parametros de Entrada:
		'       strNombreTabla .- Nombre de la tabla en la cual se buscara
		'       strNombreCampo .- Nombre del campo en el cual se buscara
		'       strValorABuscar .- Valor Buscado
		'       strMensaje .- Mensaje en caso de ya existir la llave
		'       intBase .- Base a localizar
		'-----------------------------------------------------------------------------------
		Dim rsExiste As RDO.rdoResultset
		Dim strSQL As String
		
		If IsNumeric(strValorABuscar) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object intBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strSQL = "select " & strNombreCampo & " from " & strNombreTabla & " where " & strNombreCampo & " = " & strValorABuscar & "  AND CveBase = " & intBase
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsExiste = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenStatic)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object intBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strSQL = "select " & strNombreCampo & " from " & strNombreTabla & " where " & strNombreCampo & " = '" & strValorABuscar & "'" & "  AND CveBase = " & intBase
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsExiste = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenStatic)
		End If
		
		If rsExiste.EOF Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ValidaLlaveDuplicadaBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ValidaLlaveDuplicadaBase = True
		Else
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			MsgBox(strMensaje & " ya existe", MsgBoxStyle.Exclamation)
			'UPGRADE_WARNING: Couldn't resolve default property of object ValidaLlaveDuplicadaBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ValidaLlaveDuplicadaBase = False
		End If
		rsExiste.Close()
		
	End Function
	
	Public Sub DatosNoDuplicados(ByRef strSQL As String, ByRef NomSpread As System.Windows.Forms.Control, ByVal Col As Integer, ByVal vstrCampo As String, ByVal vblnNumerico As Boolean)
		'**********************************************************************
		'Procedimiento que sirve para validar que no se
		'repitan los datos en un spread
		'**********************************************************************
		Dim i As Short
		Dim intCont As Short
		Dim strCondicion As String
		
		If InStr(1, strSQL, "WHERE") = 0 Then
			strSQL = strSQL & " WHERE 1 = 1"
			
		End If
		
		intCont = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object NomSpread.DataRowCnt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 1 To NomSpread.DataRowCnt
			'UPGRADE_WARNING: Couldn't resolve default property of object NomSpread.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			NomSpread.Row = i
			'UPGRADE_WARNING: Couldn't resolve default property of object NomSpread.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			NomSpread.Col = Col
			If NomSpread.Text <> "" Then
				If vblnNumerico Then
					strSQL = strSQL & " And " & vstrCampo & "<>" & Trim(NomSpread.Text)
				Else
					strSQL = strSQL & " And " & vstrCampo & "<>'" & Trim(NomSpread.Text) & "'"
				End If
			End If
		Next 
		
	End Sub
	Public Function ValidaLlaveDuplicada(ByRef strNombreTabla As String, ByRef strNombreCampo As String, ByRef strValorABuscar As String, ByRef strMensaje As String, ByVal vblnNumerico As Boolean) As Object
		'-----------------------------------------------------------------------------------
		' Valida si la llave con la que se insertara no existe
		'
		'   Parametros de Entrada:
		'       strNombreTabla .- Nombre de la tabla en la cual se buscara
		'       strNombreCampo .- Nombre del campo en el cual se buscara
		'       strValorABuscar .- Valor Buscado
		'       strMensaje .- Mensaje en caso de ya existir la llave
		'-----------------------------------------------------------------------------------
		Dim rsExiste As RDO.rdoResultset
		Dim strSQL As String
		
		If IsNumeric(strValorABuscar) And vblnNumerico Then
			strSQL = "select " & strNombreCampo & " from " & strNombreTabla & " where " & strNombreCampo & " = " & strValorABuscar
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsExiste = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenStatic)
		Else
			strSQL = "select " & strNombreCampo & " from " & strNombreTabla & " where " & strNombreCampo & " = '" & strValorABuscar & "'"
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsExiste = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenStatic)
		End If
		
		If rsExiste.EOF Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ValidaLlaveDuplicada. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ValidaLlaveDuplicada = True
		Else
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			MsgBox(strMensaje & " ya existe", MsgBoxStyle.Exclamation)
			'UPGRADE_WARNING: Couldn't resolve default property of object ValidaLlaveDuplicada. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ValidaLlaveDuplicada = False
		End If
		rsExiste.Close()
		
	End Function
	Public Function MsgComunes(ByVal vnMsg As Short) As Short
		'************************************************************************
		' Funcion que despliega un mensaje propio de la aplicacion
		' Entrada :
		'   vnMsg .- Numero de mensaje a desplegar
		' Salida  :
		'   Eleccion del usuario en el msgbox
		'************************************************************************
		
		Dim sTitulo As String
		Dim sMsg As String
		Dim nIcono As Short
		
		sTitulo = ""
		Select Case vnMsg
			
			Case 100 'Preguntas
				sMsg = "¿ Desea eliminar el registro ?"
				nIcono = MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Question
				
			Case 101
				sMsg = "¿ Desea actualizar el registro ?"
				nIcono = MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Question
				
			Case 103
				sMsg = "¿ Desea cerrar el registro ?"
				nIcono = MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Question
				
			Case 110
				sMsg = "Este proceso afectará el Inventario conforme al" & Chr(System.Windows.Forms.Keys.Return) & "Levantamiento Físico." & Chr(System.Windows.Forms.Keys.Return) & Chr(System.Windows.Forms.Keys.Return) & "Estas seguro de querer proceder ...?"
				nIcono = MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation
				
			Case 200 'Informativos
				sMsg = "¡ El registro ya existe !"
				nIcono = MsgBoxStyle.OKOnly + MsgBoxStyle.Information
				
			Case 202
				sMsg = "El registro ha sido modificado por otro usuario." & Chr(System.Windows.Forms.Keys.Return) & "La información más reciente se desplegará en" & Chr(System.Windows.Forms.Keys.Return) & "pantalla y perderá sus cambios..."
				nIcono = MsgBoxStyle.OKOnly + MsgBoxStyle.Information
				
			Case 204
				sMsg = "Se ha presentado un error de concurrencia en el " & Chr(System.Windows.Forms.Keys.Return) & "proceso de alta. Se consultará nuevamente la " & Chr(System.Windows.Forms.Keys.Return) & "información y perderá el último movimimento..."
				nIcono = MsgBoxStyle.OKOnly + MsgBoxStyle.Information
				
		End Select
		MsgComunes = MsgBox(sMsg, nIcono, sTitulo)
		
	End Function
	
	Public Function GrabarTransPendiente(ByRef blnEdicion_Logica As Boolean, ByRef blnAlta_Logica As Boolean) As Boolean
		'*************************************************************************
		' Función que determina si hay eventos pendientes (alta o edición) que
		' deban grabarse o no.
		' Entrada :
		'   blnEdicion_Logica .- Bandera que indica si hay captura en proceso
		'                        que representaría una edición en BD.
		'   blnAlta_Logica    .- Bandera que indica si hay captura en proceso
		'                        que representaría una alta en BD.
		' Salida :
		'   True  .- Hay que grabar la transacción
		'   False .- No hay que grabar la transacción
		'*************************************************************************
		Dim intRet As Short
		
		If (blnEdicion_Logica Or blnAlta_Logica) Then
			intRet = MsgComunes(101)
			If intRet = MsgBoxResult.Yes Then
				GrabarTransPendiente = True
			Else
				GrabarTransPendiente = False
			End If
		End If
		
	End Function
	
	Public Sub DespliegaUbicacionRegistro(ByRef rsrdoResultset As RDO.rdoResultset, ByRef objEtiqueta As Object)
		'**********************************************************************
		' Despliega en el Label del Data Control el número del Registro actual
		' Entrada:
		'       rsrdoResultset .- rdoResultset sobre el que estamos trabajando
		'       objEtiqueta .- Nombre del control sobre el que desplegaremos
		'                      el mensaje
		' Salida o Resultado:
		'       Despliega en el control que le indicamos el # de registro actual
		'**********************************************************************
		
		On Error GoTo Err_Actualiza
		
		If rsrdoResultset.RowCount > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object objEtiqueta.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			objEtiqueta.Text = rsrdoResultset.AbsolutePosition & "/" & rsrdoResultset.RowCount
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object objEtiqueta.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			objEtiqueta.Text = "0/0"
		End If
		
		Exit Sub
Err_Actualiza: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Dim strmsg As String 'String del Error
		Dim lngIndice As Integer 'Indice del Error de RDO
		
		Select Case Err.Number
			Case 40002
				For lngIndice = 0 To RDOrdoEngine_definst.rdoErrors.Count - 1
					strmsg = strmsg & RDOrdoEngine_definst.rdoErrors(lngIndice).Description & Chr(System.Windows.Forms.Keys.Return)
				Next lngIndice
				RDOrdoEngine_definst.rdoErrors.Clear()
			Case Else
				strmsg = Err.Number & " " & ErrorToString()
				Err.Clear()
		End Select
		
		MsgBox("Error al Desplegar Ubicacion de Registro " & strmsg, MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "DespliegaUbicacionRegistro")
		
	End Sub
	
	Public Sub ToolBar_EstadoBrowse(ByRef ObjToolBar As AxComctlLib.AxToolbar)
		'************************************************************
		' Rutina para dejar el ToolBar de un ABC en el estado Browse,
		' es decir, inhabilitados el grabar y el cancelar
		'************************************************************
		' Habilita todos Los Botones
		ToolBotones_Estado(ObjToolBar, True)
		
		'Deshabilita los botones Actualizar y Cancelar
		ToolBoton_Estado(ObjToolBar, "Actualizar", False)
		ToolBoton_Estado(ObjToolBar, "Cancelar", False)
		
	End Sub
	
	Public Sub ToolBotones_Estado(ByRef ObjToolBar As AxComctlLib.AxToolbar, ByVal bEstado As Boolean)
		'*******************************************************************
		' Descripción : Rutina para poner todos los  botones de un ToolBar en
		'               el estado que se requiera
		' Entrada :
		'       ObjToolBar .- Nombre del Toolbar
		'       bEstado .- Estado deseado
		'*******************************************************************
		Dim i As Byte
		For i = 1 To ObjToolBar.Buttons.Count
			ObjToolBar.Buttons(i).Enabled = bEstado
		Next i
	End Sub
	
	Public Sub ToolBoton_Estado(ByRef ObjToolBar As AxComctlLib.AxToolbar, ByVal sKey As String, ByVal bEnabled As Boolean)
		'*******************************************************************
		' Descripción : Rutina para poner un  botón de un ToolBar en
		'               el estado que se requiera
		' Entrada :
		'       ObjToolBar .- Nombre del ToolBar
		'       sKey .- Botón deseado
		'       bEnabled .- Estado que se le desea poner
		'*******************************************************************
		Dim i As Short
		Dim blnEncontrado As Boolean
		For i = 1 To ObjToolBar.Buttons.Count
			If ObjToolBar.Buttons(i).Key = sKey Then
				blnEncontrado = True
				Exit For
			End If
		Next i
		
		If blnEncontrado Then
			ObjToolBar.Buttons(i).Enabled = bEnabled
		End If
		
	End Sub
	
	Public Sub ToolBar_EstadoCambio(ByRef ObjToolBar As AxComctlLib.AxToolbar)
		'*************************************************************
		' Rutina para poner todos los botones del Toolbar de un ABC
		' en el estado de cambio
		'*************************************************************
		' DesHabilita todos Los Botones
		ToolBotones_Estado(ObjToolBar, False)
		
		'Habilita los botones Actualizar y Cancelar
		ToolBoton_Estado(ObjToolBar, "Actualizar", True)
		ToolBoton_Estado(ObjToolBar, "Cancelar", True)
		'ToolBoton_Estado ObjToolBar, "AgregarTareas", false
		ToolBoton_Estado(ObjToolBar, "Imprimir", True)
		ToolBoton_Estado(ObjToolBar, "Ayuda", True)
		ToolBoton_Estado(ObjToolBar, "Salir", True)
		
	End Sub
	
	Public Function ObtieneFechaHora(ByRef FormatoFechaHora As Short) As String
		'*******************************************************
		' Descripción   : Obtiene la Fecha y la hora Actual del Servidor
		' Entrada       : Se le pasa como parámetro el formato en que se
		'                 desea la fecha
		' Salida        : entrada = 1   => Fecha Hora
		'                 entrada = 2   => Fecha
		'                 entrada = 3   => Hora
		'*******************************************************
		
		Dim rsFecha As RDO.rdoResultset
		Dim blnAbrioRS As Boolean
		blnAbrioRS = False
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsFecha = gcn.OpenResultset("select getdate()", RDO.ResultsetTypeConstants.rdOpenForwardOnly)
		blnAbrioRS = True
		
		Select Case FormatoFechaHora
			Case 1 ' Fecha Hora
				ObtieneFechaHora = VB6.Format(rsFecha.rdoColumns(0).Value, FECHADDMMYYYY & HORAMINUTOS)
			Case 2 ' Fecha
				ObtieneFechaHora = VB6.Format(rsFecha.rdoColumns(0).Value, FECHADDMMYYYY)
			Case 3 ' Hora
				ObtieneFechaHora = VB6.Format(rsFecha.rdoColumns(0).Value, HORAMINUTOS)
		End Select
		
		'cierra resulset local
		If blnAbrioRS Then rsFecha.Close()
		
	End Function
	Public Function ConcatenaLista(ByVal vstrTabla As String, ByVal vlstControl As System.Windows.Forms.ListBox, ByVal vintDato As Short, Optional ByVal vvarLongitud As Object = Nothing) As String
		'Esta rutina concatena todos los elementos del arreglo para
		'vlstControl:     Nombre de la lista
		'intDato:   1 - itemdata    2 - descripción
		
		Dim i As Short
		Dim blnTodos As Boolean
		Dim intLongi As Short
		
		'Inicializacion
		ConcatenaLista = ""
		
		If vlstControl.SelectedItems.Count > 0 Then
			'Para lista sin separadores
			intLongi = 0
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vvarLongitud. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not IsNothing(vvarLongitud) Then intLongi = vvarLongitud
			
			blnTodos = False
			For i = 0 To vlstControl.Items.Count - 1
				ConcatenaLista = ConcatenaLista & ""
				If vlstControl.GetSelected(i) Then
					If intLongi = 0 Then 'si desea separadores. . .
						If ConcatenaLista <> "" Then ConcatenaLista = ConcatenaLista & ","
					End If
					
					Select Case vintDato
						Case 1
							ConcatenaLista = ConcatenaLista & VB6.GetItemData(vlstControl, i) 'ItemData
						Case 2 'ListData
							If intLongi = 0 Then ConcatenaLista = ConcatenaLista & "'" & VB6.GetItemString(vlstControl, i) & "'" Else ConcatenaLista = ConcatenaLista & Left(VB6.GetItemString(vlstControl, i) & Space(intLongi), intLongi)
					End Select
					If VB6.GetItemString(vlstControl, i) = "<Todos>" Then blnTodos = True
				End If
			Next i
			If blnTodos Then ConcatenaLista = "<TODOS>"
		End If
		
		If Len(ConcatenaLista) = 0 Then
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			MsgBox("Capture o corrija los siguientes datos:" & Chr(10) & Chr(10) & vstrTabla, MsgBoxStyle.Exclamation, "Aviso del Sistema")
			vlstControl.Focus()
		End If
		
	End Function
	
	
	
	
	Function UltimaUnidadLlanta(ByRef CveLlanta As Integer, ByRef CveLlantaPiso As Short) As Integer
		'--------------------------------------------------------------------------------
		'  Consulta la ultima unidad donde una llanta estuvo montada en el piso
		'   especificado.
		'
		'  Se reciben como parámetros:
		'       CveLlanta  -> # de Llanta
		'       CveLlantaPiso  -> Piso en que se encuentra la llanta
		'
		'  Regresa como resultado:
		'       Cve de Unidad donde estuvo montada la ultima vez.
		'--------------------------------------------------------------------------------
		On Error GoTo err_UltimaUnidadLlanta
		
		Dim rsQuery As RDO.rdoResultset
		Dim strSQL As String
		
		strSQL = "select CveUnidad from LlantaHistorial Where CveLlanta = " & CveLlanta
		strSQL = strSQL & " and CveLlantaPiso = " & CveLlantaPiso
		strSQL = strSQL & " and Fecha = "
		strSQL = strSQL & "(Select max(Fecha) from LlantaHistorial Where CveLlanta = " & CveLlanta
		strSQL = strSQL & " and CveLlantaPiso = " & CveLlantaPiso & ")"
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
		If Not rsQuery.EOF Then
			UltimaUnidadLlanta = rsQuery.rdoColumns("CveUnidad").Value
		Else
			UltimaUnidadLlanta = 0
		End If
		rsQuery.Close()
		
		Exit Function
		
err_UltimaUnidadLlanta: 
		MsgBox("Error al Consultar Ultima Unidad donde se montó una Llanta" & ErrorToString(), MsgBoxStyle.Critical)
		
	End Function
	
	Sub ActualizaInventarioCapa(ByRef CveAlmacen As Short, ByRef CveDivision As Short, ByRef CveRefaccion As Integer, ByRef FechaMovimiento As Date, ByRef TipoMovimiento As Short, ByRef NumFactura As Short, ByRef PrecioUnitario As Single, ByRef Cantidad As Single)
		'--------------------------------------------------------------------------------
		'  Actualiza las capas de inventario en base a datos de una entrada o salida
		'  de Almacen
		'
		'  Se reciben como parámetros:
		'       CveAlmacen  -> # de Almacen del movto.
		'       CveDivision  -> Indica si es propio o consignación
		'       CveRefaccion -> # de la refaccion
		'       FechaMovimiento  -> Fecha del movimiento
		'       Tipo Movimiento  -> 1 = Entrada     2 = Salida
		'       NumFactura       -> # de Factura en el caso de una entrada
		'       PrecioUnitario   -> Precio unitario de la refaccion
		'       Cantidad         -> Cantidad de refacciones
		'--------------------------------------------------------------------------------
		On Error GoTo err_ActualizaInventarioCapa
		
		Dim rsQuery As RDO.rdoResultset
		Dim strSQL As String
		Dim sngCantidadSurtida As Single
		Dim sngFaltante As Single
		Dim sngExistencia As Single
		Dim sngPrecio As Single
		Dim strOrden As String
		
		sngCantidadSurtida = 0
		
		Select Case TipoMovimiento
			Case MOVIMIENTOENTRADA
				If gstrTipoSalidaAlmacen = MANEJOALMACENPROM Then
					strSQL = "select * from InventarioCapa where CveAlmacen = " & CveAlmacen
					strSQL = strSQL & " and CveDivision = " & CveDivision
					strSQL = strSQL & " and CveRefaccion = " & CveRefaccion
					'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
					If rsQuery.EOF Then
						strSQL = " Insert into InventarioCapa (CveAlmacen, CveDivision, CveRefaccion,"
						strSQL = strSQL & " NumFactura, ExistenciaInicial, ExistenciaActual,"
						strSQL = strSQL & " PrecioUnitario, PrecioRevaluado, FechaMovimiento)"
						strSQL = strSQL & " values (" & CveAlmacen & "," & CveDivision & ","
						strSQL = strSQL & CveRefaccion & "," & NumFactura & "," & Cantidad & ","
						strSQL = strSQL & Cantidad & "," & PrecioUnitario & "," & PrecioUnitario & ","
						strSQL = strSQL & "'" & VB6.Format(FechaMovimiento, FECHAYYYYMMDD) & "')"
						'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gcn.Execute(strSQL)
					Else
						sngExistencia = rsQuery.rdoColumns("ExistenciaActual").Value + Cantidad
						sngPrecio = ((rsQuery.rdoColumns("ExistenciaActual").Value * rsQuery.rdoColumns("PrecioUnitario").Value) + (Cantidad * PrecioUnitario)) / sngExistencia
						strSQL = " Update InventarioCapa set ExistenciaActual = " & sngExistencia & ","
						strSQL = strSQL & " PrecioUnitario = " & sngPrecio & ","
						strSQL = strSQL & " PrecioRevaluado = " & sngPrecio
						strSQL = strSQL & " Where CveAlmacen = " & CveAlmacen
						strSQL = strSQL & " And CveDivision = " & CveDivision
						strSQL = strSQL & " And CveRefaccion = " & CveRefaccion
						'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gcn.Execute(strSQL)
					End If
				Else
					strSQL = "select * from InventarioCapa where CveAlmacen = " & CveAlmacen
					strSQL = strSQL & " and CveDivision = " & CveDivision
					strSQL = strSQL & " and CveRefaccion = " & CveRefaccion
					strSQL = strSQL & " and NumFactura = " & NumFactura
					'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
					If rsQuery.EOF Then
						strSQL = " Insert into InventarioCapa (CveAlmacen, CveDivision, CveRefaccion,"
						strSQL = strSQL & " NumFactura, ExistenciaInicial, ExistenciaActual,"
						strSQL = strSQL & " PrecioUnitario, PrecioRevaluado, FechaMovimiento)"
						strSQL = strSQL & " values (" & CveAlmacen & "," & CveDivision & ","
						strSQL = strSQL & CveRefaccion & "," & NumFactura & "," & Cantidad & ","
						strSQL = strSQL & Cantidad & "," & PrecioUnitario & "," & PrecioUnitario & ","
						strSQL = strSQL & "'" & VB6.Format(FechaMovimiento, FECHAYYYYMMDD) & "')"
						'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gcn.Execute(strSQL)
					End If
				End If
				
			Case MOVIMIENTOSALIDA
				If gstrTipoSalidaAlmacen = MANEJOALMACENUEPS Then
					strOrden = "DESC"
				Else
					strOrden = "ASC"
				End If
				
				strSQL = "select * from InventarioCapa where CveAlmacen = " & CveAlmacen
				strSQL = strSQL & " and CveDivision = " & CveDivision
				strSQL = strSQL & " and CveRefaccion = " & CveRefaccion
				strSQL = strSQL & " Order By FechaMovimiento " & strOrden
				'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
				Do While sngCantidadSurtida < Cantidad
					sngFaltante = Cantidad - sngCantidadSurtida
					If sngFaltante <= rsQuery.rdoColumns("ExistenciaActual").Value Then
						strSQL = " Update InventarioCapa set ExistenciaActual = ExistenciaActual - " & sngFaltante
						strSQL = strSQL & " Where CveAlmacen = " & CveAlmacen
						strSQL = strSQL & " and CveDivision = " & CveDivision
						strSQL = strSQL & " and CveRefaccion = " & CveRefaccion
						strSQL = strSQL & " and NumFactura = " & rsQuery.rdoColumns("NumFactura").Value
						'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gcn.Execute(strSQL)
						sngCantidadSurtida = sngCantidadSurtida + sngFaltante
					Else
						strSQL = " Update InventarioCapa set ExistenciaActual = 0 "
						strSQL = strSQL & " Where CveAlmacen = " & CveAlmacen
						strSQL = strSQL & " and CveDivision = " & CveDivision
						strSQL = strSQL & " and CveRefaccion = " & CveRefaccion
						strSQL = strSQL & " and NumFactura = " & rsQuery.rdoColumns("NumFactura").Value
						'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gcn.Execute(strSQL)
						
						sngCantidadSurtida = sngCantidadSurtida + rsQuery.rdoColumns("ExistenciaActual").Value
					End If
					
					rsQuery.MoveNext()
				Loop 
		End Select
		
		Exit Sub
		
err_ActualizaInventarioCapa: 
		MsgBox("Error al Actualizar Capas de Inventario" & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	
	Sub ActualizaRefaccionesXODT(ByRef CveODT As Integer, ByRef CveTarea As Short, ByRef CveRefaccion As Integer, ByRef CveMovtoAlmacen As Integer, ByRef Cantidad As Single, ByRef PrecioUnitario As Single, ByRef FechaMovimiento As Date, ByRef TipoMovimiento As Short)
		'--------------------------------------------------------------------------------
		'       Actualiza en ODT´s  las refacciones que se cargaron o devolvieron
		'
		'  Se reciben como parámetros:
		'       CveODT          -> # de ODT
		'       CveTarea        -> # de Tarea
		'       CveRefaccion    -> # de Refaccion
		'       CveMovtoAlmacen -> # de Salida del Almacen
		'       Cantidad        -> Cantidad de refacciones
		'       PrecioUnitario  -> Precio unitario de la refaccion
		'       FechaMovimiento -> Fecha del movimiento
		'       Tipo Movimiento -> 1 = Entrada     2 = Salida
		'--------------------------------------------------------------------------------
		On Error GoTo err_ActualizaRefaccionesXODT
		
		Dim rsQuery As RDO.rdoResultset
		Dim strSQL As String
		Dim sngCosto As Single
		
		sngCosto = Cantidad * PrecioUnitario
		
		Select Case TipoMovimiento
			Case MOVIMIENTOSALIDA ' Salida de Almacen cargada a una ODT
				strSQL = "select * from ODT where CveODT = " & CveODT
				'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
				If Not rsQuery.EOF Then
					strSQL = " Insert into ODTDetalleRefaccion (CveODT, CveTarea, CveRefaccion,"
					strSQL = strSQL & "CveMovimientoAlmacen, Cantidad, Costo, FechaMovimiento)"
					strSQL = strSQL & " values (" & CveODT & "," & CveTarea & ","
					strSQL = strSQL & CveRefaccion & "," & CveMovtoAlmacen & "," & Cantidad & ","
					strSQL = strSQL & sngCosto & ",'"
					strSQL = strSQL & VB6.Format(FechaMovimiento, FECHAYYYYMMDD) & "')"
					'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gcn.Execute(strSQL)
				Else
					MsgBox("Error al Actualizar Refacciones X ODT , # de ODT no existe", MsgBoxStyle.Critical)
					rsQuery.Close()
					End
				End If
				rsQuery.Close()
				
			Case MOVIMIENTOENTRADA ' Devolución a Almacen desde una ODT
				strSQL = "select * from ODT where CveODT = " & CveODT
				'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
				If Not rsQuery.EOF Then
					strSQL = " Delete from ODTDetalleRefaccion Where CveODT = " & CveODT
					strSQL = strSQL & " And CveTarea = " & CveTarea
					strSQL = strSQL & " And CveRefaccion = " & CveRefaccion
					'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gcn.Execute(strSQL)
				Else
					MsgBox("Error al Actualizar Refacciones X ODT , # de ODT no existe", MsgBoxStyle.Critical)
					rsQuery.Close()
					End
				End If
				rsQuery.Close()
				
		End Select
		
		Exit Sub
		
err_ActualizaRefaccionesXODT: 
		MsgBox("Error al Actualizar Refacciones X ODT" & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	
	Sub ActualizaPolizaDetalle(ByRef CvePoliza As Integer, ByRef FechaPoliza As String, ByRef Cuenta As Integer, ByRef Subcuenta As Integer, ByRef Subsubcuenta As Integer, ByRef TipoMovimiento As Short, ByRef Valor As Single, ByRef Analisis As Integer, ByRef Concepto As String, ByRef TipoMovimientoAlmacen As Short)
		'--------------------------------------------------------------------------------
		'       Actualiza tablas de Poliza y PolizaDetalle (Afectaciones Contables)
		'
		'  Se reciben como parámetros:
		'       CvePoliza       -> Consecutivo de Poliza
		'       FechaPoliza     -> Fecha de la Póliza
		'       Cuenta          -> # de Cuenta
		'       Subcuenta       -> # de Subcuenta
		'       Subsubcuenta    -> # de Subsubcuenta
		'       TipoMovimiento  -> Cargo -> 1  ,  Credito -> 2
		'       Valor           -> Monto del Movimiento
		'       Analisis        -> # de Analisis , sirve para conciliación de partidas
		'       Concepto        -> Concepto del movimiento
		'       TipoMovimientoAlmacen  -> Indica el tipo de entrada o salida
		'--------------------------------------------------------------------------------
		On Error GoTo err_ActualizaPolizaDetalle
		
		Dim rsQuery As RDO.rdoResultset
		Dim strSQL As String
		
		strSQL = "select * from Poliza Where CvePoliza = " & CvePoliza
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
		If rsQuery.EOF Then
			strSQL = " Insert into Poliza (CvePoliza, Fecha, CveTipoMovimientoAlmacen) "
			strSQL = strSQL & " values (" & CvePoliza & ",'" & FechaPoliza & "',"
			strSQL = strSQL & TipoMovimientoAlmacen & ")"
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcn.Execute(strSQL)
		End If
		rsQuery.Close()
		
		strSQL = " Insert into PolizaDetalle (CvePoliza, Fecha, Cuenta,"
		strSQL = strSQL & "Subcuenta, Subsubcuenta, TipoMovimiento, Valor, Analisis, Concepto)"
		strSQL = strSQL & " values (" & CvePoliza & ",'" & FechaPoliza & "',"
		strSQL = strSQL & Cuenta & "," & Subcuenta & "," & Subsubcuenta & ","
		strSQL = strSQL & TipoMovimiento & "," & Valor & "," & Analisis & ",'"
		strSQL = strSQL & Concepto & "')"
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gcn.Execute(strSQL)
		
		Exit Sub
		
err_ActualizaPolizaDetalle: 
		MsgBox("Error al Actualizar Poliza Detalle " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	
	Sub CentrarForma(ByRef NombreForma As System.Windows.Forms.Form)
		'*********************************************************
		'  Rutina para centrar una forma en la pantalla
		'*********************************************************
		
		NombreForma.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(NombreForma.Width)) / 2)
		NombreForma.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(NombreForma.Height)) / 2)
		
	End Sub
	
	
	' Return the number of instances of a form
	' that are currently loaded
	
	Function ExisteForma(ByVal frmName As String) As Integer
		Dim frm As System.Windows.Forms.Form
		For	Each frm In My.Application.OpenForms
			If StrComp(frm.Name, frmName, CompareMethod.Text) = 0 Then
				ExisteForma = ExisteForma + 1
			End If
		Next frm
	End Function
	Public Sub LimpiaCelda(ByRef ctlNombreSpread As System.Windows.Forms.Control, ByRef renglon As Object, ByRef columna As Object, ByRef renglon2 As Object, ByRef columna2 As Object)
		'***************
		'Este procedimiento borra por celda o por bloque de celdas. Es necesario
		'declarar variables de row,col,row2 y col2 para cuando es un bloque de celdas
		'en el evento de blockselected, y estas variables deben declararse en el
		'Gral del módulo y en el evento asignarle los row y cols correspondientes
		'************
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.MultiSelCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.IsBlockSelected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If ctlNombreSpread.IsBlockSelected Or ctlNombreSpread.MultiSelCount Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object renglon. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ctlNombreSpread.Row = renglon
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object columna. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ctlNombreSpread.Col = columna
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.Row2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object renglon2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ctlNombreSpread.Row2 = renglon2
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.Col2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object columna2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ctlNombreSpread.Col2 = columna2
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.BlockMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ctlNombreSpread.BlockMode = True
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.ActiveRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ctlNombreSpread.Row = ctlNombreSpread.ActiveRow
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.ActiveCol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ctlNombreSpread.Col = ctlNombreSpread.ActiveCol
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlNombreSpread.Action = SS_ACTION_CLEAR
		ctlNombreSpread.BackColor = System.Drawing.ColorTranslator.FromOle(COLORBLANCO)
		'UPGRADE_WARNING: Couldn't resolve default property of object ctlNombreSpread.BlockMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ctlNombreSpread.BlockMode = False
		
	End Sub
	
	Public Sub LlenaComboSpread(ByRef NomSpread As System.Windows.Forms.Control, ByRef strSQL As String, ByRef columna As Integer, ByRef renglon As Integer)
		
		Dim rsCombo As RDO.rdoResultset
		Dim strTexto As String
		
		On Error GoTo Err_LlenaCombo
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NomSpread.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NomSpread.Col = columna
		'UPGRADE_WARNING: Couldn't resolve default property of object NomSpread.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NomSpread.Row = renglon
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsCombo = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
		Do While Not rsCombo.EOF
			strTexto = strTexto & Trim(rsCombo.rdoColumns(0).Value) & Chr(9)
			rsCombo.MoveNext()
		Loop 
		rsCombo.Close()
		'UPGRADE_NOTE: Object rsCombo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCombo = Nothing
		
		'UPGRADE_WARNING: Couldn't resolve default property of object NomSpread.CellType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NomSpread.CellType = SS_CELL_TYPE_COMBOBOX ' Define cell as type COMBOBOX
		'UPGRADE_WARNING: Couldn't resolve default property of object NomSpread.TypeComboBoxList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NomSpread.TypeComboBoxList = strTexto
		'UPGRADE_WARNING: Couldn't resolve default property of object NomSpread.TypeComboBoxEditable. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		NomSpread.TypeComboBoxEditable = False
		
		Exit Sub
Err_LlenaCombo: 
		Dim lngIndice As Integer
		Dim strmsg As String
		
		If Err.Number = 40002 Then
			For lngIndice = 0 To RDOrdoEngine_definst.rdoErrors.Count - 1
				strmsg = strmsg & RDOrdoEngine_definst.rdoErrors(lngIndice).Description & Chr(13)
			Next lngIndice
			RDOrdoEngine_definst.rdoErrors.Clear()
		Else
			strmsg = Err.Number & " " & ErrorToString()
		End If
		Err.Clear()
		
		MsgBox("Ocurrió un error al cargar combo en Spread:" & NomSpread.Name & Chr(13) & strmsg, MsgBoxStyle.Exclamation, "LlenaComboSpread")
		
		
		
		Exit Sub
		
	End Sub
	
	Public Sub ActivaDesactivaCeldas(ByRef Spread As System.Windows.Forms.Control, ByRef intCol1 As Short, ByRef intReng1 As Short, ByRef intCol2 As Short, ByRef intReng2 As Short, ByRef blnActiva As Boolean)
		'******************************************************************
		'  Rutina que permite bloquear o desbloquear celdas de un spread
		'
		'   Parametros:
		'       Spread.- Nombre del Spread
		'       intCol1.- Columna inicial
		'       intReng1.- Renglon inicial
		'       intCol2.- Columna final
		'       intReng2.- Renglon Final
		'       blnActiva.- True => indica que se quiere activar
		'                   False => indica que se quiere desactivar
		'
		'******************************************************************
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Spread.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Spread.Col = intCol1
		'UPGRADE_WARNING: Couldn't resolve default property of object Spread.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Spread.Row = intReng1
		'UPGRADE_WARNING: Couldn't resolve default property of object Spread.Col2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Spread.Col2 = intCol2
		'UPGRADE_WARNING: Couldn't resolve default property of object Spread.Row2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Spread.Row2 = intReng2
		'UPGRADE_WARNING: Couldn't resolve default property of object Spread.BlockMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Spread.BlockMode = True
		If blnActiva Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Spread.Lock. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Spread.Lock = False
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Spread.Lock. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Spread.Lock = True
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Spread.BlockMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Spread.BlockMode = False
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Spread.Protect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Spread.Protect = True
		
	End Sub
	
	Public Function ConvierteHoras(ByRef Minutos As Integer) As String
		'*******************************
		'
		'  Procedimiento que sirve para convertir las horas en minutos totales
		'
		'******************************
		
		ConvierteHoras = Int(Minutos / 60) & ":" & VB6.Format(((Minutos / 60) - Int(Minutos / 60)) * 60, "00")
		
	End Function
	
	Public Function TextoSinReturn(ByVal vstrTexto As String) As String
		'*----------------------------------------------------------------------*
		'* Rutina que Elimina los return's insertados en un texto               *
		'* Input .- Texto con return's                                          *
		'* Output .- Texto Sin Return's                                         *
		'*----------------------------------------------------------------------*
		
		Dim bytPosicion As Byte
		
		vstrTexto = Trim(vstrTexto)
		
		Do Until InStr(vstrTexto, Chr(System.Windows.Forms.Keys.Return)) = 0
			bytPosicion = InStr(vstrTexto, Chr(System.Windows.Forms.Keys.Return))
			vstrTexto = Mid(vstrTexto, 1, bytPosicion - 1) & " " & Mid(vstrTexto, bytPosicion + 1, Len(vstrTexto))
		Loop 
		
		Do Until InStr(vstrTexto, Chr(34)) = 0 '      Comilla "
			bytPosicion = InStr(vstrTexto, Chr(34))
			vstrTexto = Mid(vstrTexto, 1, bytPosicion - 1) & " " & Mid(vstrTexto, bytPosicion + 1, Len(vstrTexto))
		Loop 
		
		Do Until InStr(vstrTexto, Chr(10)) = 0 '      Comilla "
			bytPosicion = InStr(vstrTexto, Chr(10))
			vstrTexto = Mid(vstrTexto, 1, bytPosicion - 1) & " " & Mid(vstrTexto, bytPosicion + 1, Len(vstrTexto))
		Loop 
		
		TextoSinReturn = Trim(vstrTexto)
		
	End Function
	
	Public Function CVTexto(ByRef Texto As Object) As String
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(Texto) Then
			CVTexto = ""
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Texto. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CVTexto = Trim(Texto)
		End If
		
	End Function
	Public Sub BuscaEnCombo(ByRef selector As System.Windows.Forms.Control, ByRef KeyAscii As Short)
		'***************************************************************
		'Procedimiento que sirve para posicionarse en un string determinado
		'al teclear una letra.
		'Recibe como parámetros el nombre del Combo y
		'                       el keyascii
		'Ejemplo: BuscaEnCombo cboCombo1,keyascii
		'Esto se utiliza en el keypress de cada combo que se utilice en la forma
		'****************************************************************
		Dim CB As Integer
		Dim BuscaTexto As String
		
		If (KeyAscii < 32 Or KeyAscii > 127) And KeyAscii <> 241 And KeyAscii <> 209 And KeyAscii <> 225 And KeyAscii <> 233 And KeyAscii <> 237 And KeyAscii <> 243 And KeyAscii <> 250 And KeyAscii <> 193 And KeyAscii <> 201 And KeyAscii <> 205 And KeyAscii <> 211 And KeyAscii <> 218 And KeyAscii <> 220 And KeyAscii <> 252 Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object selector.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If selector.SelLength = 0 Then
			BuscaTexto = selector.Text & Chr(KeyAscii)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object selector.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			BuscaTexto = Left(selector.Text, selector.SelStart) & Chr(KeyAscii)
		End If
		
		CB = SendMessage(selector.Handle.ToInt32, CB_EncuentraTexto, -1, BuscaTexto)
		
		If CB <> CB_ERR Then
			'UPGRADE_WARNING: Couldn't resolve default property of object selector.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			selector.ListIndex = CB
			'UPGRADE_WARNING: Couldn't resolve default property of object selector.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			selector.SelStart = Len(BuscaTexto)
			'UPGRADE_WARNING: Couldn't resolve default property of object selector.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object selector.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			selector.SelLength = Len(selector.Text) - selector.SelStart
		End If
		KeyAscii = 0
		
	End Sub
	Public Function TraduceNumero(ByVal vdblNumero As Double) As String
		
		Dim intCientos As Short
		Dim intMiles As Short
		Dim intMillones As Short
		Dim strCientos As String
		Dim strMiles As String
		Dim strMillones As String
		Dim strCentavos As String
		Dim lngNumero As Integer
		Dim sngCentavos As Single
		
		
		If vdblNumero = 0 Then
			TraduceNumero = "Cero"
		Else
			'Calcula los cientos
			intCientos = EliminaSimbolos(Mid(VB6.Format(Fix(vdblNumero), "000000000000"), 10, 3))
			strCientos = TraduceCentenas(intCientos) & TraduceDecenas(intCientos) & TraduceUnidades(intCientos)
			
			'Calcula los miles
			intMiles = EliminaSimbolos(Mid(VB6.Format(Fix(vdblNumero), "000000000000"), 7, 3))
			strMiles = TraduceCentenas(intMiles) & TraduceDecenas(intMiles) & TraduceUnidades(intMiles)
			If Trim(strMiles) <> "" Then strMiles = strMiles & "Mil"
			
			'Calcula los millones
			intMillones = EliminaSimbolos(Mid(VB6.Format(Fix(vdblNumero), "000000000000"), 4, 3))
			strMillones = TraduceCentenas(intMillones) & TraduceDecenas(intMillones) & TraduceUnidades(intMillones)
			Select Case intMillones
				Case 0
					'No debe de agregar nada
				Case 1
					strMillones = strMillones & "Millón "
				Case Else
					strMillones = strMillones & "Millones "
			End Select
			
			' Calcula los centavos
			lngNumero = Int(vdblNumero)
			sngCentavos = vdblNumero - lngNumero
			If sngCentavos = 0 Then
				strCentavos = "00"
			Else
				If Len(Str(sngCentavos)) = 4 Then
					strCentavos = Right(CStr(sngCentavos), 2)
				Else
					strCentavos = Right(CStr(sngCentavos), 1) & "0"
				End If
			End If
			
			TraduceNumero = UCase(Trim(strMillones & strMiles & strCientos & " PESOS M.N.  " & strCentavos & " /100 "))
		End If
		
	End Function
	Private Function TraduceUnidades(ByVal vintNumero As Short) As String
		'Es Private debido a que sola una instruccion la utiliza
		
		Dim intDecena As Short
		Dim intUnidad As Short
		
		intDecena = CShort(Mid(VB6.Format(vintNumero, "000"), 2, 1))
		intUnidad = vintNumero - Fix(vintNumero / 10) * 10
		
		If intDecena > 2 Or intDecena = 0 Then 'Si el numero es mayor a 30 unidades
			Select Case intUnidad
				Case 1
					TraduceUnidades = "Un "
				Case 2
					TraduceUnidades = "Dos "
				Case 3
					TraduceUnidades = "Tres "
				Case 4
					TraduceUnidades = "Cuatro "
				Case 5
					TraduceUnidades = "Cinco "
				Case 6
					TraduceUnidades = "Seis "
				Case 7
					TraduceUnidades = "Siete "
				Case 8
					TraduceUnidades = "Ocho "
				Case 9
					TraduceUnidades = "Nueve "
			End Select
		End If
		
	End Function
	
	Private Function TraduceDecenas(ByVal vintNumero As Short) As String
		
		Dim intDecena As Short
		Dim intUnidad As Short
		
		intDecena = CShort(Mid(VB6.Format(vintNumero, "000"), 2, 1))
		intUnidad = CShort(Mid(VB6.Format(vintNumero, "000"), 3, 1))
		
		Select Case intDecena
			Case 1
				Select Case intUnidad
					Case 0
						TraduceDecenas = "Diez "
					Case 1
						TraduceDecenas = "Once "
					Case 2
						TraduceDecenas = "Doce "
					Case 3
						TraduceDecenas = "Trece "
					Case 4
						TraduceDecenas = "Catorce "
					Case 5
						TraduceDecenas = "Quince "
					Case 6
						TraduceDecenas = "Dieciseis "
					Case 7
						TraduceDecenas = "Diecisiete "
					Case 8
						TraduceDecenas = "Dieciocho "
					Case 9
						TraduceDecenas = "Diecinueve "
				End Select
			Case 2
				Select Case intUnidad
					Case 0
						TraduceDecenas = "Veinte "
					Case 1
						TraduceDecenas = "Veintiuno "
					Case 2
						TraduceDecenas = "Veintidos "
					Case 3
						TraduceDecenas = "VeintiTres "
					Case 4
						TraduceDecenas = "Veinticuatro "
					Case 5
						TraduceDecenas = "Veinticinco "
					Case 6
						TraduceDecenas = "Veintiseis "
					Case 7
						TraduceDecenas = "Veintisiete "
					Case 8
						TraduceDecenas = "Veintiocho "
					Case 9
						TraduceDecenas = "Veintinueve "
				End Select
			Case 3
				TraduceDecenas = "Treinta "
			Case 4
				TraduceDecenas = "Cuarenta "
			Case 5
				TraduceDecenas = "Cincuenta "
			Case 6
				TraduceDecenas = "Sesenta "
			Case 7
				TraduceDecenas = "Setenta "
			Case 8
				TraduceDecenas = "Ochenta "
			Case 9
				TraduceDecenas = "Noventa "
		End Select
		If intUnidad > 0 And intDecena > 2 Then TraduceDecenas = TraduceDecenas & " y "
		
	End Function
	
	Private Function TraduceCentenas(ByVal vintNumero As Short) As String
		
		Dim intUnidad As Short
		Dim intDecenasUnidades As Short
		
		intUnidad = CShort(Mid(VB6.Format(vintNumero, "000"), 1, 1))
		intDecenasUnidades = CShort(Mid(VB6.Format(vintNumero, "000"), 2, 2))
		
		Select Case intUnidad
			Case 1
				If intDecenasUnidades = 0 Then
					TraduceCentenas = " Cien "
				Else
					TraduceCentenas = " Ciento "
				End If
			Case 2
				TraduceCentenas = " Doscientos "
			Case 3
				TraduceCentenas = " Trescientos "
			Case 4
				TraduceCentenas = " Cuatrocientos "
			Case 5
				TraduceCentenas = " Quinientos "
			Case 6
				TraduceCentenas = " Seiscientos "
			Case 7
				TraduceCentenas = " Setecientos "
			Case 8
				TraduceCentenas = " Ochocientos "
			Case 9
				TraduceCentenas = " Novecientos "
		End Select
		
	End Function
	
	Public Sub DireccionaImpresora(ByVal vstrImpresoraRequerida As String)
		Dim Printer As New Printer
		'*********************************************************************
		'*   Procedimiento para definir una impresora como default
		'*
		'*      Parametros:
		'*         vstrImpresoraRequerida.- Nombre de la impresora que
		'*                                  se desea sea el default
		'*********************************************************************
		
		Dim Impresora As Printer
		Dim EncontroImpresora As Boolean
		
		'Identifica la impresora a utilizar
		EncontroImpresora = False
		For	Each Impresora In Printers
			If Impresora.DeviceName = vstrImpresoraRequerida Then
				Printer = Impresora
				EncontroImpresora = True
				Exit For
			End If
		Next Impresora
		
		If EncontroImpresora = False Then
			MsgBox("No se encuentra la impresora requerida" & vstrImpresoraRequerida & "  ", MsgBoxStyle.Exclamation)
		End If
		
	End Sub
	
	
	Public Function FormateaFecha(ByRef strFecha As String, ByRef TipoFormato As Short) As Object
		
		Dim strMes As String
		Dim strDia As String
		Dim strAnio As String
		
		strDia = CStr(VB.Day(CDate(strFecha)))
		strMes = CStr(Month(CDate(strFecha)))
		strAnio = CStr(Year(CDate(strFecha)))
		
		If TipoFormato = 1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object FormateaFecha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FormateaFecha = strDia & "/" & strMes & "/" & strAnio
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object FormateaFecha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FormateaFecha = strMes & "/" & strDia & "/" & strAnio
		End If
		
	End Function
	
	Public Function ListaMultiselectReporte(ByRef Lista As System.Windows.Forms.ListBox, ByRef Texto As String) As String
		'-----------------------------------------------------------------------
		'Esta rutina concatena todos los elementos del arreglo itemdata para   -
		'formar una expresión lógica OR, El formato en que deja el string es   -
		' especial para el selectionformula de Crystal                         -
		'-----------------------------------------------------------------------
		Dim i As Short
		Dim blnPrimero As Boolean
		
		ListaMultiselectReporte = ""
		If Lista.SelectedItems.Count > 0 Then
			For i = 0 To Lista.Items.Count - 1
				If Not blnPrimero Then
					If Lista.GetSelected(i) Then
						ListaMultiselectReporte = "({" & Texto & "} =" & VB6.GetItemData(Lista, i)
						blnPrimero = True
					End If
				Else
					If Lista.GetSelected(i) Then
						ListaMultiselectReporte = ListaMultiselectReporte & " OR {" & Texto & "} =" & VB6.GetItemData(Lista, i)
						blnPrimero = True
					End If
				End If
			Next i
			ListaMultiselectReporte = ListaMultiselectReporte & ")"
		Else
			Beep()
			MsgBox("Seleccione por lo menos un elemento de la lista")
			Exit Function
		End If
		
	End Function
	
	Public Function ConcatenaSeleccion(ByRef Lista As System.Windows.Forms.ListBox, ByRef Texto As String) As String
		'-----------------------------------------------------------------------
		'Esta rutina concatena todos los elementos del arreglo itemdata para   -
		'formar una expresión lógica para el "IN" de un query
		'-----------------------------------------------------------------------
		Dim i As Short
		Dim blnPrimero As Boolean
		
		ConcatenaSeleccion = ""
		If Lista.SelectedItems.Count > 0 Then
			For i = 0 To Lista.Items.Count - 1
				If Not blnPrimero Then
					If Lista.GetSelected(i) Then
						ConcatenaSeleccion = Texto & " in (" & VB6.GetItemData(Lista, i)
						blnPrimero = True
					End If
				Else
					If Lista.GetSelected(i) Then
						ConcatenaSeleccion = ConcatenaSeleccion & "," & VB6.GetItemData(Lista, i)
						blnPrimero = True
					End If
				End If
			Next i
			ConcatenaSeleccion = ConcatenaSeleccion & ")"
		Else
			Beep()
			MsgBox("Seleccione por lo menos un elemento de la lista")
			Exit Function
		End If
		
	End Function
	
	Public Sub ActualizaVidaUtil(ByRef TipoMovimiento As Short, ByRef CveODT As Integer, ByRef CveRefaccion As Integer, ByRef FechaMovimiento As Date, ByRef CvePosicion As Short)
		'--------------------------------------------------------------------------------
		'       Actualiza en Tabla REFACCIONTRAYECTORIA el seguimiento a Vida Util
		'
		'  Se reciben como parámetros:
		'       Tipo Movimiento -> 1 = Montaje     2 = Retiro
		'       CveODT          -> # de ODT
		'       CveRefaccion    -> # de Refaccion
		'       FechaMovimiento -> Fecha del movimiento
		'       CvePosicion     -> Posicion en que se instalará la refaccion
		'--------------------------------------------------------------------------------
		On Error GoTo err_ActualizaVidaUtil
		
		Dim rsQuery As RDO.rdoResultset
		Dim strSQL As String
		Dim lngCveUnidad As Integer
		Dim lngKmsAcumulados As Integer
		
		
		Select Case TipoMovimiento
			Case MOVIMIENTOMONTAJE ' Salida de Almacen, inicia Vida Util de Refaccion
				strSQL = "select CveUnidad, KmsAcumulados from Unidad where CveUnidad = "
				strSQL = strSQL & "(select CveUnidad from ODT Where CveODT = " & CveODT & ")"
				'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
				If rsQuery.EOF Then
					MsgBox("Error al Actualizar Vida Util , # de ODT no existe", MsgBoxStyle.Critical)
					rsQuery.Close()
					End
				End If
				
				strSQL = " Insert into RefaccionTrayectoria (CveRefaccion, CveUnidad,"
				strSQL = strSQL & "CvePosicion, FechaMontaje, KmsMontaje, "
				strSQL = strSQL & "FechaRetiro, KmsRetiro)"
				strSQL = strSQL & " values (" & CveRefaccion & "," & rsQuery.rdoColumns("CveUnidad").Value & ","
				strSQL = strSQL & CvePosicion & ",'" & VB6.Format(FechaMovimiento, FECHAYYYYMMDD & HORAMINUTOS) & "',"
				strSQL = strSQL & rsQuery.rdoColumns("KmsAcumulados").Value & ",'" & VB6.Format(FechaMovimiento, FECHAYYYYMMDD) & "',"
				strSQL = strSQL & "0)"
				'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gcn.Execute(strSQL)
				rsQuery.Close()
				
				
			Case MOVIMIENTORETIRO ' Cuando se retira una refaccion, fin de vida util
				strSQL = "select CveUnidad, KmsAcumulados from Unidad where CveUnidad = "
				strSQL = strSQL & "(select CveUnidad from ODT Where CveODT = " & CveODT & ")"
				'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
				If rsQuery.EOF Then
					MsgBox("Error al Actualizar Vida Util , # de ODT no existe", MsgBoxStyle.Critical)
					rsQuery.Close()
					End
				End If
				
				strSQL = "Update RefaccionTrayectoria set KmsRetiro = " & rsQuery.rdoColumns("KmsAcumulados").Value & ","
				strSQL = strSQL & " FechaRetiro = '" & VB6.Format(FechaMovimiento, FECHAYYYYMMDD) & "'"
				strSQL = strSQL & " Where CveUnidad = " & rsQuery.rdoColumns("CveUnidad").Value
				strSQL = strSQL & " And CveRefaccion = " & CveRefaccion
				strSQL = strSQL & " And CvePosicion = " & CvePosicion
				strSQL = strSQL & " And KmsRetiro = 0 " ' Indica que esta montada
				'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gcn.Execute(strSQL)
				rsQuery.Close()
				
		End Select
		
		Exit Sub
		
err_ActualizaVidaUtil: 
		MsgBox("Error al Actualizar Vida Util" & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	Public Function ObtieneClave(ByVal vstrTabla As String, ByVal vstrCampo As String, ByVal vstrNombre As String) As Object
		'***************************
		'Procedimiento que sirve para
		'seleccionar la clave del destino
		'recibe com parámetro la columna para tomar el dato en que se encuentre
		'CGF
		'************************
		Dim rsDestino As RDO.rdoResultset
		
		On Error GoTo Err_cveDestino
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsDestino = gcn.OpenResultset("SELECT " & vstrCampo & " Campo " & "FROM " & vstrTabla & " (NOLOCK) " & "WHERE NombreCorto = '" & vstrNombre & "' or " & "Nombre = '" & vstrNombre & "'")
		If rsDestino.EOF Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ObtieneClave. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ObtieneClave = 0
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object ObtieneClave. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ObtieneClave = rsDestino.rdoColumns("Campo").Value
		End If
		
		rsDestino.Close()
		Exit Function
		
Err_cveDestino: 
		
	End Function
	Public Function TextoSinEspacio(ByVal vstrTexto As String) As String
		'*----------------------------------------------------------------------*
		'* Rutina que Elimina los Espacios insertados en un texto               *
		'* Hecha por CGSR                                                       *
		'* Input .- Texto con Espacios                                          *
		'* Output .- Texto Sin Espacios                                         *
		'* Fecha : 19 de Mayo de 1997                                           *
		'*----------------------------------------------------------------------*
		
		Dim bytPosicion As Byte
		Dim intPosicion As Short
		
		vstrTexto = Trim(vstrTexto)
		
		Do Until InStr(vstrTexto, Chr(System.Windows.Forms.Keys.Space)) = 0
			intPosicion = InStr(vstrTexto, Chr(System.Windows.Forms.Keys.Space))
			vstrTexto = Mid(vstrTexto, 1, intPosicion - 1) & Mid(vstrTexto, intPosicion + 1, Len(vstrTexto))
		Loop 
		
		TextoSinEspacio = Trim(vstrTexto)
		
	End Function
End Module
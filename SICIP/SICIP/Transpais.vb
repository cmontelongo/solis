Option Strict Off
Option Explicit On
Option Compare Text
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Module mdlTranspais
	
	Public gintTotalFormasActualizador As Short
	Public gblnForma As Boolean
	Public gblnTabla As Boolean
	Public gintNumElemento As Short
	
	' Variables para guardar quien abre y cierra las ODT´s
	Public gstrLoginAbrioODT As String
	Public gstrLoginCerroODT As String
	
	Public gstrEmpresa As String
	Public gstrNombreEmpresa As String
	Public gstrBase As String
	Public gintCveBase As Short
	Public gstrNombreImpresoraDefault As String
	Public gstrNombreImpresoraRol As String
	Public gstrTitulo As String
	
	Public gstrTablaEscogida As String
	
	Public glngCveODT As Integer ' Cve de ODT activa
	Public gstrSQL As String ' String de SQL, se pasa de una forma a otra
	Public glngCveUnidad As Integer ' Unidad Actual, se pasa de una forma a otra
	Public gblnAbrirODT As Boolean ' Indica si se va a abrir una ODT o no
	Public gblnCaptarNota As Boolean ' Indica si se va a captar una nota de diesel o no
	Public gblnCerrarODT As Boolean
	Public gblnAgrega As Boolean
	Public gblnPassword As Boolean
	Public gblnActualizaRemotoEnLinea As Boolean ' Indica si lo servidores remotos se actualizan en linea
	Public gsngPorcentajeRefaccionesOpcionales As Single ' Usada en desplegado de ppto , se pasa desde forma de filtro
	Public gsngPorcentajeCorrectivos As Single ' Usada en desplegado de ppto , se pasa desde forma de filtro
	
	Public gintCveUnidad As Short
	Public gintCveTarea As Short
	Public gintTipoReporte As Short
	Public gintCveCliente As Short
	Public gintMes As Short
	Public gintYear As Short
	Public gstrFiltro As String
	Public gstrOrden As String
	Public gsngCostoTallerExterno As Single
	Public gcurImpuesto As Decimal
	Public gcurExencionSubtotal As Decimal
	Public gcurImpuestoGravado As Decimal
	Public gcurImpuestoExento As Decimal
	
	Public grsPreventivos As RDO.rdoResultset ' Recorset que contiene el programa de preventivos de una unidad dada.
	Public Const gblnPreventivosIndividuales As Boolean = False
	
	' Tipos de Tareas
	Public Const TAREAPREVENTIVO As Short = 1
	Public Const TAREACORRECTIVO As Short = 2
	Public Const RAZONPREVENTIVO As Short = 1
	
	' Constantes del estatus de las ODT's
	Public Const ESTATUSABIERTA As Short = 1
	Public Const ESTATUSCERRADA As Short = 2
	Public Const ESTATUSFACTURADA As Short = 3
	
	' Orden en los querys
	Public Const ORDENASCENDENTE As Short = 1
	Public Const ORDENDESCENDENTE As Short = 2
	
	' Tipos de Componentes
	Public Const TIPOCOMPONENTEMOTOR As Short = 1
	Public Const TIPOCOMPONENTETRANSMISION As Short = 2
	Public Const TIPOCOMPONENTEDIFERENCIAL As Short = 3
	
	' Tareas PreDefinidas
	Public Const TAREAFOSEADO As Short = 94
	Public Const TAREALAVADO As Short = 95
	Public Const TAREALIMPIEZA As Short = 96
	Public Const TAREACARGACOMBUSTIBLE As Short = 97
	
	Public Const RAZONREPARACIONEXCEPCION As Short = 17
	
	Public Const PROVEEDOREXTERNO As Short = 1
	
	Public Const LUGARTALLERPROPIO As Short = 1
	
	' Cuenta y password de SQL para accesar el SIM
	Public Const LOGIN As String = "SIM"
	Public Const PASSWORD As String = "SIM"
	
	' Datos para accesar el Sistema de Almacen
	Public Const SERVIDORALMACEN As String = "ALPHAGTP"
	Public Const LOGINALMACEN As String = "ALMACEN"
	Public Const PASSWORDALMACEN As String = "ALMACEN"
	
	' Constantes de las Bases, la Base Central es la Cve de la base Victoria
	Public Const BASECENTRAL As Short = 1
	Public Const BASEVICTORIA As Short = 1
	Public Const BASETAMPICO As Short = 2
	Public Const BASEVALLES As Short = 3
	Public Const BASEMANTE As Short = 4
	Public Const BASEREYNOSA As Short = 16
	Public Const BASEMATAMOROS As Short = 6
	Public Const BASEVICTORIASUR As Short = 8
	Public Const BASETAMPICO1 As Short = 9
	Public Const BASESANLUIS As Short = 12
	Public Const BASELUMX As Short = 11
	Public Const BASEATMT As Short = 21
	Public Const BASECVR As Short = 18
	Public Const BASETALLERCENTRAL As Short = 17
	
	' Rubros de Presupuestos
	Public Const RUBROREFACCIONES As Short = 1
	Public Const RUBROACEITE As Short = 2
	Public Const RUBRODIESEL As Short = 3
	Public Const RUBROLLANTAS As Short = 4
	
	' Constantes de llantas
	Public Const LLANTAESTATUSMONTADA As Short = 1
	Public Const LLANTAESTATUSALMACEN As Short = 2
	Public Const LLANTAESTATUSREPARAR As Short = 3
	Public Const LLANTAESTATUSRECAPEAR As Short = 4
	Public Const LLANTAESTATUSBAJA As Short = 5
	
	Public Const LLANTAPISOORIGINAL As Short = 0
	Public Const LLANTAPISOPRIMERRENOVADO As Short = 1
	Public Const LLANTAPISOSEGUNDORENOVADO As Short = 2
	Public Const LLANTAPISOTERCERRENOVADO As Short = 3
	
	Public Const LLANTAUSODIRECCION As Short = 1
	Public Const LLANTAUSOTRACCION As Short = 2
	Public Const LLANTAUSOCUALQUIERPOSICION As Short = 3
	
	Public Const POSICIONLLANTA1 As Short = 1
	Public Const POSICIONLLANTA2 As Short = 2
	Public Const POSICIONLLANTA3 As Short = 3
	Public Const POSICIONLLANTA4 As Short = 4
	Public Const POSICIONLLANTA5 As Short = 5
	Public Const POSICIONLLANTA6 As Short = 6
	Public Const POSICIONLLANTA7 As Short = 7
	Public Const POSICIONLLANTA8 As Short = 8
	Public Const POSICIONLLANTA9 As Short = 9
	Public Const POSICIONLLANTA10 As Short = 10
	Public Const POSICIONREFACCION As Short = 0
	
	
	' Tipos de Movimientos de piezas controladas
	Public Const MOVIMIENTOMONTAJE As Short = 1
	Public Const MOVIMIENTORETIRO As Short = 2
	
	' Tipo de Manejo de los preventivos
	Public Const CLIENTETRANSPAIS As Short = 1
	Public Const CLIENTEOPERADORA As Short = 3
	
	' Reportes de Carros Tirados
	Public Const REPORTEBASE As Short = 1
	Public Const REPORTEFLOTILLA As Short = 2
	Public Const REPORTESISTEMA As Short = 3
	
	' Secciones de los Check-List de Foseo
	Public Const SECCIONINICIAL As Short = 1
	Public Const SECCIONFINAL As Short = 8
	
	' Origen de los Costos de refacciones
	Public Const ORIGENALMACEN As Short = 1
	Public Const ORIGENCYPSA As Short = 2
	Public Const ORIGENCARGOSDIRECTOS As Short = 3
	
	Public Const STATUSENTALLERCENTRAL As Short = 11
	Public Const STATUSDIAGNOSTICO As Short = 12
	Public Const STATUSENESPERADEMTTO As Short = 13
	Public Const STATUSREALIZANDOMTTO As Short = 14
	Public Const STATUSMTTOFINALIZADO As Short = 15
	Public Const STATUSPATIO As Short = 3
	
	' Se utilizan para pasar los preventivos de una forma a otra
	Public gintNumTareas As Short
	'UPGRADE_WARNING: Lower bound of array gTareas was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public gTareas(500) As Tareas
	Public Structure Tareas
		Dim CveTarea As Short
		Dim Nombre As String
		Dim Comentarios As String
		Dim CveTareaPadre As Short
	End Structure
	
	' Se utilizan para tener en vector las cuentas
	'UPGRADE_WARNING: Lower bound of array gCuentas was changed from 0 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public gCuentas(50) As Cuentas
	Public Structure Cuentas
		Dim Indice As Short
		Dim Cuenta As Integer
		Dim Nombre As String
	End Structure
	
	' Variables de Permisos
	Public gblnPermisoActualizar As Boolean
	
	Public gblnTieneAplicacionPadre As Boolean
	
	Public gintCveAplicacion As Short
	Public Function strEncripta(ByVal vstrLlave As String) As String
		'**********************
		'Descripcion : Desencripta la palabra recibida para dejarla leible
		'   Tiene como algoritmo la 1er. posicion le resta 1,
		'                        La 2a. Posicion le resta 2,
		'                        La 3er. Posicion le resta 3, etc.
		'   Para encriptar una palabra se debe de seguir la siguiente logica
		'       La 1er letra sacar el ASC y restarle 1
		'       La 2da Letra Sacar el ASC y Restarle 2
		'       La 3er Letra Sacar el ASC y restarle 3, Etc.
		' INPUT PARAMETERS:
		'   vstrllave: Palabra en String
		' OUTPUT : La palabra Encriptada
		'
		'cgsr 31/Dic/96
		'**********************
		
		Dim strTemp As String ' Variable de Paso para almacenar la salida
		Dim intPosicion As Short 'Posicion del cursor o de la letra a investigar
		
		For intPosicion = 1 To Len(vstrLlave)
			strTemp = strTemp & Chr(Asc(Mid(vstrLlave, intPosicion, 1)) - intPosicion)
		Next intPosicion
		
		strEncripta = strTemp
		
	End Function
	Public Function strDesEncripta(ByVal vstrLlave As String) As String
		'**********************
		'Descripcion : Desencripta la palabra recibida para dejarla leible
		'   Tiene como algoritmo la 1er. posicion le suma 1,
		'                        La 2a. Posicion le suma 2,
		'                        La 3er. Posicion le suma 3, etc.
		'   Para encriptar una palabra se debe de seguir la siguiente logica
		'       La 1er letra sacar el ASC y restarle 1
		'       La 2da Letra Sacar el ASC y Restarle 2
		'       La 3er Letra Sacar el ASC y restarle 3, Etc.
		' INPUT PARAMETERS:
		'   vstrllave: Palabra Encriptada en String
		' OUTPUT : La palabra desencriptada
		'
		'cgsr 31/Dic/96
		'**********************
		
		Dim strTemp As String ' Variable de Paso para almacenar la salida
		Dim intPosicion As Short 'Posicion del cursor o de la letra a investigar
		
		For intPosicion = 1 To Len(vstrLlave)
			strTemp = strTemp & Chr(Asc(Mid(vstrLlave, intPosicion, 1)) + intPosicion)
		Next intPosicion
		
		strDesEncripta = strTemp
		
	End Function
	
	Public Sub ActualizaServidoresRemotos(ByRef strQuery As String)
		
		On Error GoTo err_ActualizaServidoresRemotos
		
		Dim strServidor As String
		
		'strServidor = "Transoper"
		'If gblnConeccionVictoria Then gcnVictoria.Execute strQuery
		
		'strServidor = "PENSION_REYNOSA"
		'If gblnConeccionReynosa Then gcnReynosa.Execute strQuery
		
		strServidor = "TCServer"
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnTallerCentral.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionTallerCentral And gstrServidor <> "TCSERVER" Then gcnTallerCentral.Execute(strQuery)
		
		'strServidor = "NTTAM"
		'If gblnConeccionTampico Then gcnTampico.Execute strQuery
		
		'strServidor = "NTVAL"
		'If gblnConeccionValles Then gcnValles.Execute strQuery
		
		'strServidor = "NTMAT"
		'If gblnConeccionMatamoros Then gcnMatamoros.Execute strQuery
		
		'strServidor = "NTMAN"
		'If gblnConeccionMante Then gcnMante.Execute strQuery
		
		'strServidor = "NTSLP"
		'If gblnConeccionSanLuis Then gcnSanLuis.Execute strQuery
		
		strServidor = "LUMXSBD"
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnLUMX.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionLUMX Then gcnLUMX.Execute(strQuery)
		
		strServidor = "ATMSERVER"
		'UPGRADE_WARNING: Couldn't resolve default property of object gcnATMT.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gblnConeccionATMT Then gcnATMT.Execute(strQuery)
		
		Exit Sub
		
err_ActualizaServidoresRemotos: 
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
		End Select
		Err.Clear()
		MsgBox(" Error al Actualizar Servidores Remotos -> " & strServidor & vbLf & strmsg, MsgBoxStyle.Critical, "Actualiza Servidores Remotos")
		Resume Next
		
	End Sub
	Public Function ObtieneConsecutivoADT() As Object
		
		Dim blnCreado As Boolean
		Dim rsTemp As RDO.rdoResultset
		Dim lngBase As Integer
		Dim lngMiles As Integer
		Dim lngBaseLimSup As Integer
		Dim lngBaseLimInf As Integer
		Dim vintPrefijo As Short
		Dim vintCerosSufijo As Short
		Dim strSQL As String
		
		On Error GoTo err_ObtieneConsecutivoODT
		
		
		'FORMATO de BAMMCCCC. (Base, Año, Mes, Consecutivo)
		vintCerosSufijo = 5
		strSQL = "SELECT DATEPART(yy,GETDATE()) Year, datepart(mm,getdate()) Mes "
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsTemp = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
		blnCreado = True
		If rsTemp.EOF Then Exit Function
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(rsTemp.rdoColumns("Year").Value) Or IsDbNull(rsTemp.rdoColumns("Mes").Value) Then Exit Function
		
		'Obtengo Prefijo de la Empresa.
		vintPrefijo = CShort(VB6.Format(gintCveBase, "0"))
		vintPrefijo = CShort(vintPrefijo & Right(rsTemp.rdoColumns("Year").Value, 2))
		vintPrefijo = Val(CStr(vintPrefijo))
		rsTemp.Close()
		
		lngMiles = 10 ^ vintCerosSufijo
		lngBase = vintPrefijo * lngMiles
		lngBaseLimSup = lngBase + lngMiles
		lngBaseLimInf = lngBase - 1
		strSQL = "SELECT MAX(CveADT) FROM  ADT  WHERE CveADT > " & Str(lngBaseLimInf) & " AND CveADT < " & Str(lngBaseLimSup)
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsTemp = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
		
		If Not rsTemp.EOF Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(rsTemp.rdoColumns(0).Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ObtieneConsecutivoADT. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ObtieneConsecutivoADT = lngBase + 1
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object ObtieneConsecutivoADT. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ObtieneConsecutivoADT = rsTemp.rdoColumns(0).Value + 1
			End If
		Else
			MsgBox("Error al obtener consecutivo")
		End If
		rsTemp.Close()
		
		Exit Function
		
err_ObtieneConsecutivoODT: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Obtener Consecutivo de ADT " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Function
	
	Public Function ObtieneConsecutivoODT() As Integer
		
		Dim blnCreado As Boolean
		Dim rsTemp As RDO.rdoResultset
		Dim lngBase As Integer
		Dim lngMiles As Integer
		Dim lngBaseLimSup As Integer
		Dim lngBaseLimInf As Integer
		Dim vintPrefijo As Short
		Dim vintCerosSufijo As Short
		Dim strSQL As String
		
		On Error GoTo err_ObtieneConsecutivoODT
		
		
		'FORMATO de BAMMCCCC. (Base, Año, Mes, Consecutivo)
		vintCerosSufijo = 4
		strSQL = "SELECT DATEPART(yy,GETDATE()) Year, datepart(mm,getdate()) Mes "
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsTemp = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
		blnCreado = True
		If rsTemp.EOF Then Exit Function
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If IsDbNull(rsTemp.rdoColumns("Year").Value) Or IsDbNull(rsTemp.rdoColumns("Mes").Value) Then Exit Function
		
		'Obtengo Prefijo de la Empresa.
		vintPrefijo = CShort(VB6.Format(gintCveBase, "0"))
		vintPrefijo = CShort(vintPrefijo & Right(rsTemp.rdoColumns("Year").Value, 1))
		vintPrefijo = CShort(vintPrefijo & VB6.Format(rsTemp.rdoColumns("Mes").Value, "00"))
		vintPrefijo = Val(CStr(vintPrefijo))
		rsTemp.Close()
		
		lngMiles = 10 ^ vintCerosSufijo
		lngBase = vintPrefijo * lngMiles
		lngBaseLimSup = lngBase + lngMiles
		lngBaseLimInf = lngBase - 1
		strSQL = "SELECT MAX(CveODT) FROM  ODT  WHERE CveODT > " & Str(lngBaseLimInf) & " AND CveODT < " & Str(lngBaseLimSup)
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsTemp = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenForwardOnly, RDO.LockTypeConstants.rdConcurReadOnly)
		
		If Not rsTemp.EOF Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(rsTemp.rdoColumns(0).Value) Then
				ObtieneConsecutivoODT = lngBase + 1
			Else
				ObtieneConsecutivoODT = rsTemp.rdoColumns(0).Value + 1
			End If
		Else
			MsgBox("Error al obtener consecutivo")
		End If
		rsTemp.Close()
		
		Exit Function
		
err_ObtieneConsecutivoODT: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Obtener Consecutivo de ODT " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Function
	
	
	Public Sub CierreMesIndicadores(ByRef intAnio As Short, ByRef intMes As Short)
		
		Dim strSQL As String
		Dim i As Short
		Dim rsKmsLitros As RDO.rdoResultset
		Dim rsUnidades As RDO.rdoResultset
		Dim rsTotales As RDO.rdoResultset
		Dim rsConsumo As RDO.rdoResultset
		Dim lngKms As Integer
		Dim lngLitros As Integer
		Dim sngRefacciones As Single
		Dim sngLlantas As Single
		Dim sngAceites As Single
		
		
		On Error GoTo err_CreaTemporal
		
		' Depura la información existente
		strSQL = " delete from IndicadoresAnual where CveAnio = " & intAnio
		strSQL = strSQL & " and CveMes = " & intMes
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gcn.Execute(strSQL)
		
		
		strSQL = "Select * from Unidad (NOLOCK) order by CveUnidad "
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsUnidades = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		Do While Not rsUnidades.EOF
			
			'--------------------------------------
			'      Obtiene los Kms y litros
			'--------------------------------------
			strSQL = "Select isnull(sum(LL.KmsRecorridos),0) TotalKms ,"
			strSQL = strSQL & " isnull(sum(LL.NumLitros),0) TotalLitros from Llegada LL (NOLOCK)"
			strSQL = strSQL & " where LL.CveUnidad = " & rsUnidades.rdoColumns("CveUnidad").Value
			strSQL = strSQL & " and DATEPART(yy,LL.FechaLlegada) = " & intAnio
			strSQL = strSQL & " and DATEPART(mm,LL.FechaLlegada) = " & intMes
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsKmsLitros = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
			lngKms = rsKmsLitros.rdoColumns("TotalKms").Value
			lngLitros = rsKmsLitros.rdoColumns("TotalLitros").Value
			rsKmsLitros.Close()
			
			'-----------------------------------------------------------
			'        Obtiene el Consumo de refacciones
			'-----------------------------------------------------------
			strSQL = "Select isnull(sum(UC.Monto),0) TotalConsumo "
			strSQL = strSQL & " from UnidadConsumos UC (NOLOCK) "
			strSQL = strSQL & " where UC.CveUnidad = " & rsUnidades.rdoColumns("CveUnidad").Value
			strSQL = strSQL & " and DATEPART(yy,UC.Fecha) = " & intAnio
			strSQL = strSQL & " and DATEPART(mm,UC.Fecha) = " & intMes
			strSQL = strSQL & " and UC.Rubro = " & RUBROREFACCIONES
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsConsumo = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
			sngRefacciones = rsConsumo.rdoColumns("TotalConsumo").Value
			rsConsumo.Close()
			
			'-----------------------------------------------------------
			'          Obtiene el Consumo de Aceite
			'-----------------------------------------------------------
			strSQL = "Select isnull(sum(UC.Monto),0) TotalConsumo "
			strSQL = strSQL & " from UnidadConsumos UC (NOLOCK) "
			strSQL = strSQL & " where UC.CveUnidad = " & rsUnidades.rdoColumns("CveUnidad").Value
			strSQL = strSQL & " and DATEPART(yy,UC.Fecha) = " & intAnio
			strSQL = strSQL & " and DATEPART(mm,UC.Fecha) = " & intMes
			strSQL = strSQL & " and UC.Rubro = " & RUBROACEITE
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsConsumo = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
			sngAceites = rsConsumo.rdoColumns("TotalConsumo").Value
			rsConsumo.Close()
			
			'-----------------------------------------------------------
			'            Obtiene el Consumo de Llantas
			'-----------------------------------------------------------
			strSQL = "Select isnull(sum(UC.Monto),0) TotalConsumo "
			strSQL = strSQL & " from UnidadConsumos UC (NOLOCK) "
			strSQL = strSQL & " where UC.CveUnidad = " & rsUnidades.rdoColumns("CveUnidad").Value
			strSQL = strSQL & " and DATEPART(yy,UC.Fecha) = " & intAnio
			strSQL = strSQL & " and DATEPART(mm,UC.Fecha) = " & intMes
			strSQL = strSQL & " and UC.Rubro = " & RUBROLLANTAS
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsConsumo = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
			sngLlantas = rsConsumo.rdoColumns("TotalConsumo").Value
			rsConsumo.Close()
			
			
			' Inserta en tabla de Indicadores
			strSQL = "insert into IndicadoresAnual "
			strSQL = strSQL & " (CveAnio, CveMes, CveUnidad, Kms, "
			strSQL = strSQL & " Litros, Refacciones, Llantas, Aceites )"
			strSQL = strSQL & " values (" & intAnio & "," & intMes & "," & rsUnidades.rdoColumns("CveUnidad").Value
			strSQL = strSQL & "," & lngKms & "," & lngLitros & "," & sngRefacciones & ","
			strSQL = strSQL & sngLlantas & "," & sngAceites & ")"
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcn.Execute(strSQL)
			
			rsUnidades.MoveNext()
		Loop 
		rsUnidades.Close()
		
		Exit Sub
		
err_CreaTemporal: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error en Cierre Mensual de Indicadores " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	
	
	Public Function ObtieneNombreCliente(ByRef gintCveCliente As Short) As Object
		'*******************************************************
		' Descripción   : Obtiene el Nombre del Cliente
		' Entrada       : Se le pasa como parámetro el # del cliente
		' Salida        : Nombre del cliente
		'*******************************************************
		
		On Error GoTo Err_ObtieneNombreCliente
		
		Dim rsCliente As RDO.rdoResultset
		Dim strSQL As String
		
		
		strSQL = "select NombreCorto from Cliente Where CveCliente = " & gintCveCliente
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsCliente = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		If Not rsCliente.EOF Then
			ObtieneNombreCliente = Trim(rsCliente.rdoColumns("NombreCorto").Value)
		Else
			MsgBox("error al obtener el Nombre del Cliente ")
			End
		End If
		
		rsCliente.Close()
		Exit Function
		
Err_ObtieneNombreCliente: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Obtener Nombre del Cliente " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Function
	
	Public Function ObtieneBase(ByRef strServidor As String) As Short
		'*******************************************************
		' Descripción   : Obtiene la cve de la base
		' Entrada       : Se le pasa como parámetro el servidor donde se correrá la aplicacion
		'                 desea la fecha
		' Salida        : el # de la base
		'*******************************************************
		
		On Error GoTo Err_ObtieneBase
		
		Dim rsBase As RDO.rdoResultset
		Dim strSQL As String
		
		
		strSQL = "select CveBase from Base Where upper(Servidor) = '" & UCase(strServidor) & "'"
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsBase = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		If Not rsBase.EOF Then
			ObtieneBase = rsBase.rdoColumns("CveBase").Value
		Else
			MsgBox("error al obtener la base, Servidor no existe ")
			End
		End If
		
		rsBase.Close()
		Exit Function
		
Err_ObtieneBase: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Obtener base local " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Function
	
	Public Sub EstableceEncabezadosCalendario(ByRef rspr As AxFPSpread.AxvaSpread)
		
		Dim strDiasCortos As String
		Dim strDiasLargos As String
		Dim strMesCorto As String
		Dim strMesLargo As String
		Dim strChar As String
		
		strChar = Chr(9)
		strDiasCortos = "L" & strChar & "M" & strChar & "MI" & strChar & "J" & strChar & "V" & strChar & "S" & strChar & "D"
		strDiasLargos = "Lunes" & strChar & "Martes" & strChar & "Miercoles" & Chr(201) & strChar & "Jueves" & strChar & "Viernes" & strChar & "Sabado" & Chr(193) & "B" & strChar & "Domingo"
		strMesCorto = "Ene" & strChar & "Feb" & strChar & "Mar" & strChar & "Abr" & strChar & "May" & strChar & "Jun" & strChar & "Jul" & strChar & "Ago" & strChar & "Sep" & strChar & "Ooct" & strChar & "Nov" & strChar & "Dic"
		strMesLargo = "Enero" & strChar & "Febrero" & strChar & "Marzo" & strChar & "Abril" & strChar & "Mayo" & strChar & "Junio" & strChar & "Julio" & strChar & "Agosto" & strChar & "Septiembre" & strChar & "Octubre" & strChar & "Noviembre" & strChar & "Diciembre"
		
		Call rspr.SetCalTextOverride(strDiasCortos, strDiasLargos, strMesCorto, strMesLargo, "Aceptar", "Cancelar")
		
		
	End Sub
	
	Public Sub ActualizaKmsPartes(ByVal vlngCveUnidad As Integer, ByVal vlngKms As Integer, ByVal vintOperacion As Object)
		'*************************************************
		'Subrutina de Actualizacion de Kilometros a las llantas de los equipos
		'Recibe    vlngCveUnidad .- # de la Unidad
		'          vlngKms.- Kms a acumular
		'          vintOperacion .-  1 = Suma    2 = Resta
		'*************************************************
		
		On Error GoTo Err_ActualizaKmsLlantas
		
		Dim strSQL As String
		Dim rsQuery As RDO.rdoResultset
		Dim strQuery As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object vintOperacion. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If vintOperacion = 1 Then
			'Acumula kms.
			strSQL = "Update P "
			strSQL = strSQL & "Set P.KmsAcumulados = P.KmsAcumulados + " & vlngKms & ", "
			strSQL = strSQL & "    P.KmsPeriodo = P.KmsPeriodo + " & vlngKms & " "
			strSQL = strSQL & "from ParteKardex PK join Parte P ON P.CveParte = PK.CveParte "
			strSQL = strSQL & "Where PK.FechaBaja Is Null"
			strSQL = strSQL & "  and PK.CveTipoMovimiento = 1"
			strSQL = strSQL & "  AND PK.CveUnidad = " & vlngCveUnidad & "  and PK.CvePosicion <> " & POSICIONREFACCION
		Else
			'Resta kms.
			strSQL = "Update P "
			strSQL = strSQL & "Set P.KmsAcumulados = P.KmsAcumulados - " & vlngKms & ", "
			strSQL = strSQL & "    P.KmsPeriodo = P.KmsPeriodo - " & vlngKms & " "
			strSQL = strSQL & "from ParteKardex PK join Parte P ON P.CveParte = PK.CveParte "
			strSQL = strSQL & "Where PK.FechaBaja Is Null"
			strSQL = strSQL & "  and PK.CveTipoMovimiento = 1"
			strSQL = strSQL & "  AND PK.CveUnidad = " & vlngCveUnidad & "  and PK.CvePosicion <> " & POSICIONREFACCION
		End If
		
		If gblnActualizaServidoresRemotos Then
			strQuery = "select CveBase from Unidad " & "where CveUnidad = " & vlngCveUnidad
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsQuery = gcn.OpenResultset(strQuery, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
			If Not rsQuery.EOF Then
				Select Case rsQuery.rdoColumns("CveBase").Value
					Case BASELUMX
						'UPGRADE_WARNING: Couldn't resolve default property of object gcnLUMX.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If gblnConeccionLUMX Then gcnLUMX.Execute(strSQL)
					Case BASEATMT
						'UPGRADE_WARNING: Couldn't resolve default property of object gcnATMT.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If gblnConeccionATMT Then gcnATMT.Execute(strSQL)
					Case Else
						'UPGRADE_WARNING: Couldn't resolve default property of object gcnTallerCentral.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If gblnConeccionTallerCentral Then gcnTallerCentral.Execute(strSQL)
				End Select
			End If
			rsQuery.Close()
		End If
		
		' Acumula en el servidor de Taller Central (TCServer)
		'If gblnConeccionTallerCentral And gstrServidor = "TCSERVER" Then
		'    gcnTallerCentral.Execute strSQL
		'Else
		'    gcn.Execute strSQL
		'End If
		
		Exit Sub
		
Err_ActualizaKmsLlantas: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al acumular kms a las Partes " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	
	Public Sub ActualizaKmsLlantas(ByVal vlngCveUnidad As Integer, ByVal vlngKms As Integer, ByVal vintOperacion As Object)
		'*************************************************
		'Subrutina de Actualizacion de Kilometros a las llantas de los equipos
		'Recibe    vlngCveUnidad .- # de la Unidad
		'          vlngKms.- Kms a acumular
		'          vintOperacion .-  1 = Suma    2 = Resta
		'*************************************************
		
		On Error GoTo Err_ActualizaKmsLlantas
		
		Dim rsQuery As RDO.rdoResultset
		Dim strSQL As String
		Dim strQuery As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object vintOperacion. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If vintOperacion = 1 Then
			'Acumula kms.
			strSQL = "UPDATE Llanta set KmsAcumulados = KmsAcumulados +  "
			strSQL = strSQL & vlngKms & " Where CveLlanta in (SELECT DISTINCT CveLlanta FROM UnidadLlanta "
			strSQL = strSQL & " Where CveUnidad = " & vlngCveUnidad & " and Posicion <> " & POSICIONREFACCION & ")"
		Else
			'Resta kms.
			strSQL = "UPDATE Llanta set KmsAcumulados = KmsAcumulados -  "
			strSQL = strSQL & vlngKms & " Where CveLlanta in (SELECT DISTINCT CveLlanta FROM UnidadLlanta "
			strSQL = strSQL & " Where CveUnidad = " & vlngCveUnidad & " and Posicion <> " & POSICIONREFACCION & ")"
		End If
		
		If gblnActualizaServidoresRemotos Then
			strQuery = "select CveBase from Unidad " & "where CveUnidad = " & vlngCveUnidad
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsQuery = gcn.OpenResultset(strQuery, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
			If Not rsQuery.EOF Then
				Select Case rsQuery.rdoColumns("CveBase").Value
					Case BASELUMX
						'UPGRADE_WARNING: Couldn't resolve default property of object gcnLUMX.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If gblnConeccionLUMX Then gcnLUMX.Execute(strSQL)
					Case BASEATMT
						'UPGRADE_WARNING: Couldn't resolve default property of object gcnATMT.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If gblnConeccionATMT Then gcnATMT.Execute(strSQL)
					Case Else
						'UPGRADE_WARNING: Couldn't resolve default property of object gcnTallerCentral.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If gblnConeccionTallerCentral Then gcnTallerCentral.Execute(strSQL)
				End Select
			End If
			rsQuery.Close()
		End If
		
		'If gblnConeccionTallerCentral And gstrServidor = "TCSERVER" Then
		'    gcnTallerCentral.Execute strSQL
		'Else
		'    gcn.Execute strSQL
		'End If
		
		Exit Sub
		
Err_ActualizaKmsLlantas: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al acumular kms a las llantas " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	
	
	Public Sub ImprimeTareasPorRealizar(ByRef lngCveUnidad As Integer)
		Dim Printer As New Printer
		'************************************************************************
		'  Procedimiento para Imprimir el reporte de Tareas por realizar a una
		'  unidad cuando llega al ROL
		'
		'  Recibe como parametros:
		'           lngCveUnidad .- # de la unidad
		'************************************************************************
		
		Dim strSQL As String
		Dim rsTareasDefault As RDO.rdoResultset
		Dim rsQuery As RDO.rdoResultset
		Dim rsPredictivos As RDO.rdoResultset
		Dim rsDatosLlegada As RDO.rdoResultset
		Dim rsFallas As RDO.rdoResultset
		Dim lngKmsAcum As Integer
		Dim lngNumLitros As Integer
		Dim lngKmsRendimiento As Integer
		Dim lngCveLlegada As Integer
		Dim sngRendimiento As Single
		
		
		On Error GoTo Err_ImprimeComprobanteViaje
		
		' Obtiene datos de la unidad
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsQuery = gcn.OpenResultset("select KmsAcumulados from Unidad where CveUnidad =" & lngCveUnidad)
		lngKmsAcum = rsQuery.rdoColumns("KmsAcumulados").Value
		rsQuery.Close()
		
		' Obtiene datos de la unidad
		strSQL = " Select max(CveLlegada) as Maxima from Llegada where CveUnidad = " & lngCveUnidad
		strSQL = strSQL & " and CveBase = " & gintCveBase
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(rsQuery.rdoColumns("Maxima").Value) Then
			lngCveLlegada = rsQuery.rdoColumns("Maxima").Value
		Else
			lngCveLlegada = 0
		End If
		rsQuery.Close()
		
		' Obtiene los kms y litros de la llegada
		strSQL = "select NumLitros,KmsRendimiento from Llegada where CveLlegada = " & lngCveLlegada
		strSQL = strSQL & " and CveBase = " & gintCveBase
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsDatosLlegada = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		If Not rsDatosLlegada.EOF Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(rsDatosLlegada.rdoColumns("NumLitros").Value) Then
				lngNumLitros = rsDatosLlegada.rdoColumns("NumLitros").Value
			Else
				lngNumLitros = 0
			End If
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not IsDbNull(rsDatosLlegada.rdoColumns("KmsRendimiento").Value) Then
				lngKmsRendimiento = rsDatosLlegada.rdoColumns("KmsRendimiento").Value
			Else
				lngKmsRendimiento = 0
			End If
			If lngNumLitros = 0 Or lngKmsRendimiento = 0 Then
				sngRendimiento = 0
			Else
				sngRendimiento = CSng(VB6.Format(lngKmsRendimiento / lngNumLitros, DOSDECIMALES))
			End If
		Else
			lngNumLitros = 0
			lngKmsRendimiento = 0
			sngRendimiento = 0
		End If
		rsDatosLlegada.Close()
		
		'Imprime Encabezados
		Printer.FontName = "Courier New"
		Printer.Print()
		Printer.Print()
		Printer.Print()
		Printer.Print()
		Printer.FontSize = 12
		Printer.FontUnderline = True
		Printer.FontBold = True
		Printer.Print(TAB(27), "TAREAS PENDIENTES POR UNIDAD")
		Printer.FontUnderline = False
		Printer.FontBold = False
		Printer.FontSize = 10
		Printer.Print()
		Printer.Print(TAB(65), VB6.Format(ObtieneFechaHora(1), FECHADDMMYYYY & HORAMINUTOS))
		Printer.Print()
		
		Printer.FontBold = True
		Printer.Print(TAB(20), "Unidad:", TAB(28), lngCveUnidad, TAB(53), "Kms Acumulados: " & lngKmsAcum)
		Printer.Print(TAB(10), "Kms. Rend. :" & lngKmsRendimiento, TAB(35), "Litros: " & lngNumLitros, TAB(55), "Rendimiento:" & sngRendimiento)
		Printer.Print()
		Printer.FontUnderline = True
		Printer.Print(TAB(10), "Preventivos")
		Printer.FontBold = False
		Printer.Print()
		Printer.Print(TAB(12), "Tarea", TAB(30), "Kms. Ultimo Mtto.", TAB(55), "Fecha ultimo Mtto.")
		Printer.FontUnderline = False
		
		'----------------------------------------------
		'      Carga las Tareas por default
		'----------------------------------------------
		strSQL = "select T.CveTarea,T.Nombre "
		strSQL = strSQL & " from Tarea T"
		strSQL = strSQL & " where T.Baja = 0 AND T.CveTarea in (" & TAREAFOSEADO & "," & TAREACARGACOMBUSTIBLE & ")"
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsTareasDefault = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		Do While Not rsTareasDefault.EOF
			Printer.Print(TAB(8), rsTareasDefault.rdoColumns("Nombre"))
			rsTareasDefault.MoveNext()
		Loop 
		rsTareasDefault.Close()
		
		
		'-------------------------------------------
		'   Carga los preventivos individuales por realizar
		'-------------------------------------------
		VerificaPreventivos(lngCveUnidad, PREVENTIVOSINDIVIDUALES)
		
		Do While Not grsPreventivos.EOF
			Printer.Print(TAB(8), grsPreventivos.rdoColumns("Nombre"), TAB(36), grsPreventivos.rdoColumns("KmsAcumulados"), TAB(62), grsPreventivos.rdoColumns("FechaOcurrencia"))
			grsPreventivos.MoveNext()
		Loop 
		grsPreventivos.Close()
		
		'-------------------------------------------
		'   Carga los preventivos agrupados por realizar
		'-------------------------------------------
		VerificaPreventivos(lngCveUnidad, PREVENTIVOSAGRUPADOS)
		
		Do While Not grsPreventivos.EOF
			Printer.Print(TAB(8), grsPreventivos.rdoColumns("Nombre"))
			grsPreventivos.MoveNext()
		Loop 
		grsPreventivos.Close()
		
		
		'--------------------------------------
		'    Encabezado de los predictivos
		'--------------------------------------
		Printer.Print()
		Printer.Print()
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print(TAB(10), "Predictivos")
		Printer.FontBold = False
		Printer.Print()
		Printer.Print(TAB(10), "Tarea", TAB(30), "Kms. Pronostico", TAB(55), "Fecha Pronostico")
		Printer.FontUnderline = False
		
		'---------------------------------------
		'  Carga los predictivos por realizar
		'---------------------------------------
		strSQL = "select U.CveUnidad,T.CveTarea,T.Nombre, UP.KmsPronostico,UP.FechaPronostico ,(UP.KmsPronostico - U.KmsAcumulados) as VencidoKms "
		strSQL = strSQL & " from Tarea T, UnidadPredictivo UP , Unidad U "
		strSQL = strSQL & " where T.Baja = 0 AND UP.CveUnidad = " & lngCveUnidad
		strSQL = strSQL & " and UP.CveUnidad = U.CveUnidad "
		strSQL = strSQL & " and UP.CveTarea = T.CveTarea "
		strSQL = strSQL & " and (UP.KmsPronostico <= U.KmsAcumulados or "
		strSQL = strSQL & " UP.FechaPronostico <= '" & VB6.Format(ObtieneFechaHora(2), FECHAMMDDYYYY) & "')"
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsPredictivos = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		Do While Not rsPredictivos.EOF
			Printer.Print(TAB(8), rsPredictivos.rdoColumns("Nombre"), TAB(36), rsPredictivos.rdoColumns("KmsPronostico"), TAB(62), rsPredictivos.rdoColumns("FechaPronostico"))
			rsPredictivos.MoveNext()
		Loop 
		rsPredictivos.Close()
		
		
		'--------------------------------------------
		'    Encabezado de las fallas en Camino
		'--------------------------------------------
		Printer.Print()
		Printer.Print()
		Printer.FontBold = True
		Printer.FontUnderline = True
		Printer.Print(TAB(10), "Fallas en Camino Reportadas")
		Printer.FontBold = False
		Printer.FontUnderline = False
		Printer.Print()
		
		strSQL = " Select * from LlegadaDetalle where CveLlegada = " & lngCveLlegada
		strSQL = strSQL & " and CveBase = " & gintCveBase & " order by NumRenglon "
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsFallas = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		Do While Not rsFallas.EOF
			Printer.Print(TAB(4), Trim(rsFallas.rdoColumns("FallaEnCamino").Value))
			rsFallas.MoveNext()
		Loop 
		rsFallas.Close()
		
		
		' Termina el documento
		Printer.EndDoc()
		
		Exit Sub
		
		
Err_ImprimeComprobanteViaje: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Imprimir Tareas por Realizar " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	Public Function VerificaExistenciaDiesel(ByVal vsngLitro As Single) As Boolean
		'**************************************************************************
		'Funcion que verifica la existencia de Diesel
		'Entrada .-
		'   vsngLitro .- Cantidad de Litros que se descontaran
		'Salida .-
		'   True .- Si Existe disponibilidad de existencia sobre facturas del diesel
		'   false .- No hay existencia de Disel segun facturas
		'****************************************
		On Error GoTo Err_VerificaExistenciaDiesel
		
		Dim rsSaldo As RDO.rdoResultset
		Dim sngLitrosDisponibles As Single
		Dim strSQL As String
		
		
		'Abre rdoResultset para cargar la cantidad de litros disponibles segun la capa de existencia
		strSQL = "select sum(SaldoActualLitros) as Saldo from  CombustibleExistencia "
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsSaldo = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not rsSaldo.EOF Or Not IsDbNull(rsSaldo.rdoColumns("Saldo").Value) Then sngLitrosDisponibles = rsSaldo.rdoColumns("Saldo").Value
		rsSaldo.Close()
		
		If vsngLitro <= sngLitrosDisponibles Then
			VerificaExistenciaDiesel = True
		Else
			VerificaExistenciaDiesel = False
		End If
		
		Exit Function
		
		
		
Err_VerificaExistenciaDiesel: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Verificar Existencia de Diesel " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Function
	
	Public Sub DescuentaLitrosDiesel(ByVal vsngLitro As Single)
		'**************************************************************************
		' Subrutina que descuenta los litros a cada factura que se surtieron segun la
		' nota, esto se hace por medio de PEPS
		'
		' Entrada .-
		'    vsngLitro .- Cantidad de Litros que se descontaran
		'
		'****************************************
		On Error GoTo Err_DescuentaLitros
		
		Dim rsSaldoActual As RDO.rdoResultset
		Dim sngLitros As Single ' Litros por Surtir
		Dim sngSaldoFactura As Single
		Dim strmsg As String
		Dim lngIndice As Integer
		Dim strSQL As String
		
		strSQL = "select * from CombustibleExistencia where SaldoActualLitros > 0 order by FechaEntrada asc"
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsSaldoActual = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		sngLitros = vsngLitro
		Do Until rsSaldoActual.EOF
			'Toma los valores que tiene actualmente la factura
			sngSaldoFactura = rsSaldoActual.rdoColumns("SaldoActualLitros").Value
			
			rsSaldoActual.Edit()
			If sngLitros < sngSaldoFactura Then
				rsSaldoActual.rdoColumns("SaldoActualLitros").Value = sngSaldoFactura - sngLitros
				sngLitros = 0
			Else
				sngLitros = sngLitros - sngSaldoFactura
				rsSaldoActual.rdoColumns("SaldoActualLitros").Value = 0
			End If
			rsSaldoActual.Update()
			
			If sngLitros = 0 Then Exit Do
			rsSaldoActual.MoveNext()
		Loop 
		rsSaldoActual.Close()
		
		Exit Sub
		
Err_DescuentaLitros: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Descontar Litros de Diesel " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Sub
	
	Public Sub CargaParametrosTranspais()
		'************************************************************************
		' Rutina que realiza la carga de parámetros generales del sistema
		' a variables globales.
		'************************************************************************
		
		On Error GoTo Err_CargaParametros
		
		Dim strmsg As String
		
		' Lee los parametros del archivo .ini
		gstrBaseDeDatos = BuscaParametrosIni("Datos Generales", "BaseDeDatos")
		gstrDirectorioRpt = BuscaParametrosIni("Datos Generales", "DirReportes")
		gstrDirectorioIconos = BuscaParametrosIni("Datos Generales", "DirIconos")
		gstrNombreImpresoraDefault = BuscaParametrosIni("Datos Generales", "NombreImpresoraDefault")
		gstrBase = BuscaParametrosIni("Datos Generales", "Base")
		gstrNombreEmpresa = BuscaParametrosIni("Datos Generales", "NombreEmpresa")
		gstrServidorCentral = BuscaParametrosIni("Datos Generales", "ServidorCentral")
		gstrServidorAlmacen = BuscaParametrosIni("Datos Generales", "ServidorAlmacen")
		
		Exit Sub
		
		
Err_CargaParametros: 
		strmsg = "Ocurrió un error al leer los parámetros" & Chr(13)
		strmsg = strmsg & "de inicio del sistema. La ejecución se detendrá."
		MsgBox(strmsg)
		CierraConeccion()
		End
		
	End Sub
	
	
	Public Sub InicializaVectorTareas()
		
		Dim i As Short
		
		For i = 1 To 500
			gTareas(i).CveTarea = 0
			gTareas(i).Nombre = ""
		Next i
		
		gintNumTareas = 0
		
	End Sub
	
	Public Sub ActualizaServidoresRemotosConfirmando(ByRef strQueryUpdate As String, ByRef strQueryInsert As String)
		'-------------------------------------------------------------------------------
		'   Esta rutina actualiza en servidores remotos, tratando de hacer un update,
		'   pero si no se hace el update porque no existe el registro, intenta un insert
		'   del registro deseado
		'-------------------------------------------------------------------------------
		On Error GoTo err_ActualizaServidoresRemotosConfirmando
		
		' Actualiza Victoria
		If gblnConeccionVictoria Then
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnVictoria.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcnVictoria.Execute(strQueryUpdate)
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnVictoria.RowsAffected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnVictoria.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If gcnVictoria.RowsAffected = 0 And Len(strQueryInsert) > 0 Then gcnVictoria.Execute(strQueryInsert)
		End If
		
		' Actualiza Reynosa
		If gblnConeccionReynosa Then
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnReynosa.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcnReynosa.Execute(strQueryUpdate)
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnReynosa.RowsAffected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnReynosa.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If gcnReynosa.RowsAffected = 0 And Len(strQueryInsert) > 0 Then gcnReynosa.Execute(strQueryInsert)
		End If
		
		' Actualiza Taller Central
		If gblnConeccionTallerCentral Then
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnTallerCentral.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcnTallerCentral.Execute(strQueryUpdate)
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnTallerCentral.RowsAffected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnTallerCentral.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If gcnTallerCentral.RowsAffected = 0 And Len(strQueryInsert) > 0 Then gcnTallerCentral.Execute(strQueryInsert)
		End If
		
		' Actualiza Tampico
		If gblnConeccionTampico Then
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnTampico.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcnTampico.Execute(strQueryUpdate)
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnTampico.RowsAffected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnTampico.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If gcnTampico.RowsAffected = 0 And Len(strQueryInsert) > 0 Then gcnTampico.Execute(strQueryInsert)
		End If
		
		
		' Actualiza Valles
		If gblnConeccionValles Then
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnValles.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcnValles.Execute(strQueryUpdate)
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnValles.RowsAffected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnValles.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If gcnValles.RowsAffected = 0 And Len(strQueryInsert) > 0 Then gcnValles.Execute(strQueryInsert)
		End If
		
		
		' Actualiza Matamoros
		If gblnConeccionMatamoros Then
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnMatamoros.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcnMatamoros.Execute(strQueryUpdate)
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnMatamoros.RowsAffected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnMatamoros.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If gcnMatamoros.RowsAffected = 0 And Len(strQueryInsert) > 0 Then gcnMatamoros.Execute(strQueryInsert)
		End If
		
		' Actualiza Mante
		If gblnConeccionMante Then
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnMante.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcnMante.Execute(strQueryUpdate)
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnMante.RowsAffected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnMante.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If gcnMante.RowsAffected = 0 And Len(strQueryInsert) > 0 Then gcnMante.Execute(strQueryInsert)
		End If
		
		' Actualiza San Luis Potosí
		If gblnConeccionSanLuis Then
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnSanLuis.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcnSanLuis.Execute(strQueryUpdate)
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnSanLuis.RowsAffected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnSanLuis.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If gcnSanLuis.RowsAffected = 0 And Len(strQueryInsert) > 0 Then gcnSanLuis.Execute(strQueryInsert)
		End If
		
		If gblnConeccionLUMX Then
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnLUMX.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcnLUMX.Execute(strQueryUpdate)
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnLUMX.RowsAffected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnLUMX.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If gcnLUMX.RowsAffected = 0 And Len(strQueryInsert) > 0 Then gcnLUMX.Execute(strQueryInsert)
		End If
		
		If gblnConeccionATMT Then
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnATMT.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gcnATMT.Execute(strQueryUpdate)
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnATMT.RowsAffected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gcnATMT.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If gcnATMT.RowsAffected = 0 And Len(strQueryInsert) > 0 Then gcnATMT.Execute(strQueryInsert)
		End If
		
		
		Exit Sub
		
err_ActualizaServidoresRemotosConfirmando: 
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
		End Select
		Err.Clear()
		MsgBox("Error al Actualizar Servidores Remotos Confirmando " & vbLf & strmsg, MsgBoxStyle.Critical, "mdlTranspais")
		
		End
		Resume Next
	End Sub
	Public Function VerificaPresupuestoRubro(ByRef intRubro As Short, ByRef intUnidadTipo As Short, ByRef intCveCliente As Short, ByRef intYear As Short, ByRef intMes As Short) As Object
		'*******************************************************************************
		'   Funcion para obtener el presupuesto de un rubro, mes y tipo de unidad
		'   determinados
		'      Entrada:
		'               intRubro -> Rubro deseado
		'               intUnidadTipo -> Tipo de Unidad o Flotilla
		'               intMes -> Mes del presupuesto
		'*******************************************************************************
		
		Dim strSQL As String
		Dim rsPresupuesto As RDO.rdoResultset
		
		strSQL = "Select * from PresupuestoRubro where CveRubro = " & intRubro
		strSQL = strSQL & " and CveUnidadTipo = " & intUnidadTipo
		strSQL = strSQL & " and Anio = " & intYear
		strSQL = strSQL & " and CveCliente = " & intCveCliente
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsPresupuesto = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		If Not rsPresupuesto.EOF Then
			Select Case intMes
				Case 1
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("Ene").Value
				Case 2
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("Feb").Value
				Case 3
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("Mar").Value
				Case 4
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("Abr").Value
				Case 5
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("May").Value
				Case 6
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("Jun").Value
				Case 7
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("Jul").Value
				Case 8
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("Ago").Value
				Case 9
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("Sep").Value
				Case 10
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("Oct").Value
				Case 11
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("Nov").Value
				Case 12
					'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					VerificaPresupuestoRubro = rsPresupuesto.rdoColumns("Dic").Value
			End Select
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object VerificaPresupuestoRubro. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			VerificaPresupuestoRubro = 0
		End If
		rsPresupuesto.Close()
		
	End Function
	
	Public Function ChecaRellenoAceite(ByRef intCveUnidad As Short) As Single
		'************************************************************************
		'  Rutina para checar el factor de consumo de relleno de aceite de motor
		'  por cada 1,000 kms. recorridos de una unidad dada.
		'  Considera los Kms y rellenos de aceite de los ultimos 30 dias
		'
		'      Entrada:
		'         intCveUnidad .-  # de Unidad a verificar
		'************************************************************************
		
		Dim rsKms As RDO.rdoResultset
		Dim rsLitros As RDO.rdoResultset
		Dim lngKms As Integer
		Dim lngLitros As Integer
		Dim strFechaHoy As String
		Dim strFechaInicio As String
		Dim strFechaFin As String
		Dim strSQL As String
		
		strFechaHoy = ObtieneFechaHora(1)
		
		strFechaFin = strFechaHoy
		strFechaInicio = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -30, CDate(strFechaFin)))
		
		' Obtiene los Kms acumulados en el periodo para la unidad
		strSQL = "Select sum(KmsRecorridos) TotalKms from Llegada where CveUnidad = " & intCveUnidad
		strSQL = strSQL & " and FechaLlegada >= '" & VB6.Format(strFechaInicio, FECHAMMDDYYYY & HORAMINUTOS) & "' "
		strSQL = strSQL & " and FechaLlegada <= '" & VB6.Format(strFechaFin, FECHAMMDDYYYY & HORAMINUTOS) & "' "
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsKms = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(rsKms.rdoColumns("TotalKms").Value) Then
			lngKms = rsKms.rdoColumns("TotalKms").Value
		Else
			lngKms = 0
		End If
		rsKms.Close()
		
		' Obtiene los litros de aceite de relleno en el periodo
		strSQL = "Select sum(ND.LitrosRelleno) TotalLitros from NotaDetalle ND, Nota N "
		strSQL = strSQL & " where N.CveUnidad = " & intCveUnidad
		strSQL = strSQL & " and N.Fecha >= '" & VB6.Format(strFechaInicio, FECHAMMDDYYYY & HORAMINUTOS) & "' "
		strSQL = strSQL & " and N.Fecha <= '" & VB6.Format(strFechaFin, FECHAMMDDYYYY & HORAMINUTOS) & "' "
		strSQL = strSQL & " and N.CveNota = ND.CveNota  "
		strSQL = strSQL & " and ND.CveComponente = " & TIPOCOMPONENTEMOTOR
		
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsLitros = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(rsLitros.rdoColumns("TotalLitros").Value) Then
			lngLitros = rsLitros.rdoColumns("TotalLitros").Value
		Else
			lngLitros = 0
		End If
		rsLitros.Close()
		
		If lngKms > 0 Then
			ChecaRellenoAceite = (lngLitros * 1000) / lngKms
		Else
			ChecaRellenoAceite = 0
		End If
		
		
	End Function
	
	Public Function CalculaCostoRefacciones(ByRef intCveTarea As Short, ByRef intCveUnidad As Short) As Single
		'*********************************************************************
		'   Rutina para calcular el costo de refacciones de una Tarea para
		'   determinado tipo de unidad
		'*********************************************************************
		On Error GoTo err_CalculaCostoRefacciones
		
		Dim strSQL As String
		Dim rsRefacciones As RDO.rdoResultset
		
		CalculaCostoRefacciones = 0
		
		' Calcula el costo de refacciones para la tarea y la unidad requerida
		strSQL = "select sum(SUM_015.SUM01513_PRECIO * TR.Cantidad) as Total "
		strSQL = strSQL & " from  TareaRefaccion TR, Unidad U, SUM_015 SUM_015 "
		strSQL = strSQL & " Where U.CveUnidad = " & intCveUnidad
		strSQL = strSQL & "  and TR.CveTarea = " & intCveTarea
		strSQL = strSQL & "  and TR.CveUnidadTipo = U.CveUnidadTipo "
		strSQL = strSQL & "  and TR.CveRefaccion = SUM_015.SUM00411_CLAVE "
		strSQL = strSQL & "  and TR.Requerida = 1 "
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsRefacciones = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(rsRefacciones.rdoColumns("Total").Value) Then CalculaCostoRefacciones = CSng(VB6.Format(rsRefacciones.rdoColumns("Total").Value, "###,##0.00"))
		rsRefacciones.Close()
		
		Exit Function
		
err_CalculaCostoRefacciones: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Calcular Costo de Refacciones " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Function
	
	Public Function VerificaProveedorExterno(ByRef intCveProveedor As Short) As Object
		'******************************************************************
		'   Rutina que verifica si el proveedor de la ODT es Externo
		'   propio o no
		'
		'   Proveedor Propio = 0     ;   Proveedor Externo = 1
		'******************************************************************
		
		Dim strSQL As String
		Dim rsQuery As RDO.rdoResultset
		
		' Verifica si es proveedor externo
		strSQL = "select ProveedorExterno from Proveedor where CveProveedor = " & intCveProveedor
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		If Not rsQuery.EOF Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(rsQuery.rdoColumns("PROVEEDOREXTERNO").Value) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object VerificaProveedorExterno. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				VerificaProveedorExterno = 0
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object VerificaProveedorExterno. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				VerificaProveedorExterno = rsQuery.rdoColumns("PROVEEDOREXTERNO").Value
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object VerificaProveedorExterno. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			VerificaProveedorExterno = 0
		End If
		
		rsQuery.Close()
		
		
	End Function
	
	Public Function VerificaTallerPropio(ByRef intCveLugarReparacion As Short) As Object
		'******************************************************************
		'   Rutina que verifica si el lugar de reparacion es de Taller
		'   propio o no
		'
		'   Taller Propio = 1     ;   Taller Externo = 0
		'******************************************************************
		
		Dim strSQL As String
		Dim rsQuery As RDO.rdoResultset
		
		' Verifica que sea Taller Propio
		strSQL = "select TallerPropio from LugarReparacion where CveLugarReparacion = " & intCveLugarReparacion
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		If Not rsQuery.EOF Then
			'UPGRADE_WARNING: Couldn't resolve default property of object VerificaTallerPropio. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			VerificaTallerPropio = rsQuery.rdoColumns("TallerPropio").Value
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object VerificaTallerPropio. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			VerificaTallerPropio = 0
		End If
		
		rsQuery.Close()
		
		
	End Function
	
	Public Sub CargaVectorCuentas()
		
		Dim i As Short
		Dim strSQL As String
		Dim rsCuenta As RDO.rdoResultset
		
		strSQL = " select * from Cuenta order by CveCuenta "
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsCuenta = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		
		i = 0
		Do While Not rsCuenta.EOF
			gCuentas(i).Indice = i
			gCuentas(i).Cuenta = rsCuenta.rdoColumns("CveCuenta").Value
			rsCuenta.MoveNext()
			i = i + 1
		Loop 
		rsCuenta.Close()
		
	End Sub
	
	Public Function VerificaPoliticaOperador(ByRef intCveOperador As Short, ByRef intCveBase As Short) As Boolean
		'*******************************************************
		' Descripción   : Verifica si la asignacion del Operador a cierta ruta cumple con la política
		'
		' Entrada       : 1) Cve. del Operador
		'                 2) Cve de la  Base de la Corrida
		'
		' Salida        : Verdadero o Falso segun sea el caso
		'*******************************************************
		
		On Error GoTo Err_VerificaPoliticaOperador
		
		Dim rsQuery As RDO.rdoResultset
		Dim strSQL As String
		Dim intCveUnidadTipoAsignada As Short
		
		VerificaPoliticaOperador = False
		
		' Verifica si la categoría del operador concuerda con el definido para la ruta
		strSQL = "SELECT CveEmpresa, CveBase, CveOperadorCategoria FROM Operador WHERE CveOperador = " & intCveOperador
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		If rsQuery.EOF Then
			MsgBox("Se realizo una operacion no valida al buscar Operador: " & intCveOperador & Chr(System.Windows.Forms.Keys.Return) & "No existe en la Base de Datos", MsgBoxStyle.Information, "VerificaOperador")
			VerificaPoliticaOperador = False
			rsQuery.Close()
			Exit Function
		Else
			If rsQuery.rdoColumns("CveBase").Value = intCveBase Then
				VerificaPoliticaOperador = True
				rsQuery.Close()
				Exit Function
			End If
		End If
		rsQuery.Close()
		
		Exit Function
		
Err_VerificaPoliticaOperador: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Verificar Política del Operador " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Function
	
	Public Function VerificaPoliticaUnidad(ByRef intCveUnidad As Short, ByRef intCveUnidadTipo As Short) As Boolean
		'*******************************************************
		' Descripción   : Verifica si la asignacion de una unidad a cierta ruta cumple con la política
		'
		' Entrada       : 1) Cve. de Unidad Asignada
		'                 2) Cve. del tipo de unidad definido para la ruta
		'
		' Salida        : Verdadero o Falso segun sea el caso
		'*******************************************************
		
		On Error GoTo Err_VerificaPoliticaUnidad
		
		Dim rsQuery As RDO.rdoResultset
		Dim strSQL As String
		Dim intCveUnidadTipoAsignada As Short
		
		VerificaPoliticaUnidad = False
		
		' Verifica si el tipo de la unidad seleccionada coincide con el tipo de unidad definida
		strSQL = "SELECT CveUnidadTipo FROM Unidad WHERE CveUnidad = " & intCveUnidad
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		If rsQuery.EOF Then
			MsgBox("Se realizo una operacion no valida al buscar unidad: " & intCveUnidad & Chr(System.Windows.Forms.Keys.Return) & "No existe en la Base de Datos", MsgBoxStyle.Information, "VerificaUnidad")
			VerificaPoliticaUnidad = False
			rsQuery.Close()
			Exit Function
		Else
			If rsQuery.rdoColumns("CveUnidadTipo").Value = intCveUnidadTipo Then
				VerificaPoliticaUnidad = True
				rsQuery.Close()
				Exit Function
			Else
				intCveUnidadTipoAsignada = rsQuery.rdoColumns("CveUnidadTipo").Value
			End If
		End If
		rsQuery.Close()
		
		'Verifica si el tipo de la unidad seleccionada coincide con alguno de los tipos de unidades alternativas
		strSQL = "SELECT * FROM UnidadTipoEquivalencia WHERE CveUnidadTipo = " & intCveUnidadTipo
		strSQL = strSQL & " AND CveUnidadTipoEquivalencia = " & intCveUnidadTipoAsignada
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsQuery = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		If Not rsQuery.EOF Then
			VerificaPoliticaUnidad = True
			rsQuery.Close()
			Exit Function
		End If
		rsQuery.Close()
		
		Exit Function
		
Err_VerificaPoliticaUnidad: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(" Error al Verificar Política de Unidad " & ErrorToString(), MsgBoxStyle.Critical)
		
	End Function
End Module
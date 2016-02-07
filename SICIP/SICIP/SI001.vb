Option Strict Off
Option Explicit On
Option Compare Text
Friend Class frmLogin
	Inherits System.Windows.Forms.Form
	
	Const APLICACION As Short = 2
	
	Dim blnPermiso As Boolean
	Private Sub cboBase_Click()
		Dim frmODT As Object
		Dim cboBase As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cboBase.ListIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object cboBase.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gintCveBase = cboBase.ItemData(cboBase.ListIndex)
		
		'------------------------------------------------------
		'    Obtiene parámetros de configuración globales
		'------------------------------------------------------
		CargaParametrosConfiguracion()
		
		'---------------------------------
		'   Inicia la forma Principal
		'---------------------------------
		Me.Hide()
		'UPGRADE_WARNING: Couldn't resolve default property of object frmODT.Show. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmODT.Show()
		
	End Sub
	
	'UPGRADE_WARNING: Event cboArticulo.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboArticulo_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboArticulo.SelectedIndexChanged
		
		Dim strSQL As String
		Dim rsdetalle As RDO.rdoResultset
		
		If cboArticulo.SelectedIndex >= 0 Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.MaxRows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.MaxRows = 0
			
			'Llena los combos
			strSQL = "SELECT Notas from Articulo WHERE CveArticulo =  " & VB6.GetItemData(cboArticulo, cboArticulo.SelectedIndex)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsdetalle = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
			If rsdetalle.EOF Then
				MsgBox("No existe Informacion", MsgBoxStyle.Exclamation, "ButtonClick")
			Else
				txtDescripcion.Text = rsdetalle.rdoColumns("Notas").Value
			End If
			rsdetalle.Close()
			
			strSQL = "select AMD.Nombre,AM.CantidadRequerida,UM.NombreCorto,ISNULL(AMD.KgPorM2,0) KgPorM2,AMD.KgPorM2 * AM.CantidadRequerida Peso " & ",ISNULL(AMD.PrecioLista,D.PrecioLista) PrecioLista, (AMD.KgPorM2 * AM.CantidadRequerida) * AMD.PrecioLista Importe " & ",AM.NumRenglon " & "from Articulo A " & " JOIN ArticuloManufactura AM ON A.CveArticulo = AM.CveArticulo" & " JOIN Articulo AMD ON AMD.CveArticulo = AM.CveArticuloDetalle" & " LEFT JOIN UnidadMedida UM ON UM.CveUnidadMedida = AMD.CveUnidadMedidaCotizacion" & " LEFT JOIN (SELECT AD.CveArticulo,SUM(ADS.PrecioLista*AD.CantidadRequerida) PrecioLista " & " FROM ArticuloDetalle AD " & " JOIN Articulo ADS on ADS.CveArticulo = AD.CveArticuloDetalle " & " GROUP BY AD.CveArticulo) AS D ON D.CveArticulo = AMD.CveArticulo " & "Where A.cvearticulo = " & VB6.GetItemData(cboArticulo, cboArticulo.SelectedIndex) & " order by AM.NumRenglon"
			
			
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.EditModePermanent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.EditModePermanent = True
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.MaxRows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.Row = sprInsumos.MaxRows
			
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsdetalle = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
			
			' Llena el spread
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.ReDraw. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.ReDraw = False
			Do Until rsdetalle.EOF
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.MaxRows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.MaxRows = sprInsumos.MaxRows + 1
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.MaxRows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Row = sprInsumos.MaxRows
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MakeFloatCell(2, 2, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0)
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MakeFloatCell(4, 5, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0)
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MakeFloatCell(6, 7, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", True, True, 2, 0)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 1 'A
				sprInsumos.Text = rsdetalle.rdoColumns("Nombre").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 2 'B
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Value = rsdetalle.rdoColumns("CantidadRequerida").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 3 'C
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(rsdetalle.rdoColumns("NombreCorto").Value) Then sprInsumos.Text = rsdetalle.rdoColumns("NombreCorto").Value
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 4 'D
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(rsdetalle.rdoColumns("KgPorM2").Value) Then sprInsumos.Text = rsdetalle.rdoColumns("KgPorM2").Value
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 5 'E
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Formula. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Formula = "B" & sprInsumos.Row & " * D" & sprInsumos.Row
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 6 'F
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(rsdetalle.rdoColumns("PrecioLista").Value) Then sprInsumos.Text = rsdetalle.rdoColumns("PrecioLista").Value
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 7 'G
				If rsdetalle.rdoColumns("KgPorM2").Value = 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Formula. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sprInsumos.Formula = "B" & sprInsumos.Row & " * F" & sprInsumos.Row
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Formula. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sprInsumos.Formula = "E" & sprInsumos.Row & " * F" & sprInsumos.Row
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft
				
				rsdetalle.MoveNext()
			Loop 
			rsdetalle.Close()
			
			
			
			
		End If
		
		
		
		
		
	End Sub
	
	
	Private Sub cmdAgregar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAgregar.Click
		
		Dim strSQL As String
		Dim rsdetalle As RDO.rdoResultset
		
		If cboArticulos.SelectedIndex >= 0 Then
			
			'Always have the spreadsheet in edit mode
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.EditModePermanent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.EditModePermanent = True
			
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.MaxRows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.MaxRows = sprInsumos.MaxRows + 1
			
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.MaxRows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.Row = sprInsumos.MaxRows
			
			strSQL = "select CveArticulo,A.Nombre,UM.NombreCorto,KgPorM2,CostoMonedaNacional " & "from Articulo A " & "JOIN UnidadMedida UM ON UM.CveUnidadMedida = A.CveUnidadMedidaCotizacion " & "Where A.CveArticulo = " & VB6.GetItemData(cboArticulos, cboArticulos.SelectedIndex)
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsdetalle = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
			
			' Llena el spread
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.ReDraw. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.ReDraw = False
			Do Until rsdetalle.EOF
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MakeFloatCell(2, 2, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0)
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MakeFloatCell(4, 5, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0)
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MakeFloatCell(6, 7, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", True, True, 2, 0)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 1 'A
				sprInsumos.Text = rsdetalle.rdoColumns("Nombre").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 3 'C
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
				sprInsumos.Text = rsdetalle.rdoColumns("NombreCorto").Value
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 4 'D
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
				sprInsumos.Text = rsdetalle.rdoColumns("KgPorM2").Value
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 5 'E
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Formula. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Formula = "B" & sprInsumos.Row & " * D" & sprInsumos.Row
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 6 'F
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(rsdetalle.rdoColumns("CostoMonedaNacional").Value) Then sprInsumos.Text = rsdetalle.rdoColumns("CostoMonedaNacional").Value
				
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Col = 7 'G
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Formula. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.Formula = "E" & sprInsumos.Row & " * F" & sprInsumos.Row
				'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft
				
				rsdetalle.MoveNext()
			Loop 
			rsdetalle.Close()
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.ReDraw. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.ReDraw = True
			
			cmdAgregar.Visible = False
			cboArticulos.Visible = False
			
			txtBuscar.Visible = True
			cmdBuscarMecanico.Visible = True
			
		End If
	End Sub
	
	
	
	Sub MakeFloatCell(ByRef Col As Integer, ByRef col2 As Integer, ByRef Row As Integer, ByRef row2 As Integer, ByRef floatmin As String, ByRef floatmax As String, ByRef floatmoney As Boolean, ByRef floatsep As Boolean, ByRef decplaces As Short, ByRef fpvalue As Double)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.Col = Col
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.col2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.col2 = col2
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.Row = Row
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.row2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.row2 = row2
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.BlockMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.BlockMode = True
		'Define cells as type FLOAT
		If floatmoney Then
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.CellType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.CellType = FPSpread.CellTypeConstants.CellTypeCurrency
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeCurrencyShowSymbol. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.TypeCurrencyShowSymbol = True
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeCurrencyDecPlaces. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.TypeCurrencyDecPlaces = decplaces
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeCurrencyShowSep. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.TypeCurrencyShowSep = floatsep
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeCurrencyMin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.TypeCurrencyMin = floatmin
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeCurrencyMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.TypeCurrencyMax = floatmax
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.CellType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.CellType = FPSpread.CellTypeConstants.CellTypeNumber
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeNumberDecPlaces. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.TypeNumberDecPlaces = decplaces
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeNumberShowSep. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.TypeNumberShowSep = floatsep
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeNumberMin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.TypeNumberMin = floatmin
			'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeNumberMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sprInsumos.TypeNumberMax = floatmax
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.Value = fpvalue
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.BlockMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.BlockMode = False
		
	End Sub
	Private Sub cmdBuscarMecanico_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBuscarMecanico.Click
		Dim strSQL As String
		
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		strSQL = "SELECT CveArticulo,Nombre FROM Articulo WHERE Activo=1 AND CveTipoRecurso = 1 AND Nombre like '%" & Replace(txtBuscar.Text, " ", "%") & "%' ORDER BY Nombre"
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		LlenaVariosSelectores(strSQL, New Object(){"cboArticulos"}, Me)
		If cboArticulos.Items.Count > 0 Then
			cboArticulos.Visible = True
			txtBuscar.Visible = False
			cmdBuscarMecanico.Visible = False
			cmdAgregar.Visible = True
			txtBuscar.Text = ""
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	
	Public Sub frmLogin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim strSQL As String
		
		CentrarForma(Me)
		'txtServidor =
		'Se asignan Variables de Cuenta y Password
		gstrLogin = "SICIP"
		gstrPassword = "SICIP"
		gstrServidor = "NAUTILIUS"
		gstrBaseDeDatos = "SICIP"
		
		'CargaParametrosTranspais
		AbreConeccion()
		
		strSQL = "SELECT CveArticulo,Nombre FROM Articulo WHERE Activo = 1 AND EsManufacturado = 1"
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		LlenaVariosSelectores(strSQL, New Object(){"cboArticulo"}, Me)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.MaxRows. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.MaxRows = 0
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.Row = -1000
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.Col = 1
		sprInsumos.Font = VB6.FontChangeBold(sprInsumos.Font, True)
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
		sprInsumos.Text = "Materiales"
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.ColWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.ColWidth(1) = 24
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.Col = 2
		sprInsumos.Font = VB6.FontChangeBold(sprInsumos.Font, True)
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
		sprInsumos.Text = "Cant"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.Col = 3
		sprInsumos.Font = VB6.FontChangeBold(sprInsumos.Font, True)
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
		sprInsumos.Text = "UN"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.Col = 4
		sprInsumos.Font = VB6.FontChangeBold(sprInsumos.Font, True)
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
		sprInsumos.Text = "kg/m/pza"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.Col = 5
		sprInsumos.Font = VB6.FontChangeBold(sprInsumos.Font, True)
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
		sprInsumos.Text = "Peso"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.Col = 6
		sprInsumos.Font = VB6.FontChangeBold(sprInsumos.Font, True)
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
		sprInsumos.Text = "$/UN/kg"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.Col = 7
		sprInsumos.Font = VB6.FontChangeBold(sprInsumos.Font, True)
		'UPGRADE_WARNING: Couldn't resolve default property of object sprInsumos.TypeHAlign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sprInsumos.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter
		sprInsumos.Text = "TOTAL"
		
	End Sub
	Private Sub txtCuenta_KeyPress(ByRef KeyAscii As Short)
		Dim txtPassword As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object txtPassword.SetFocus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If KeyAscii = System.Windows.Forms.Keys.Return Then txtPassword.SetFocus()
	End Sub
	Private Sub txtPassword_KeyPress(ByRef KeyAscii As Short)
		Dim cboBase As Object
		Dim txtPassword As Object
		Dim txtCuenta As Object
		
		Dim strSQL As String
		Dim rsPermiso As RDO.rdoResultset
		Dim rsPassword As RDO.rdoResultset
		
		If KeyAscii = System.Windows.Forms.Keys.Return Then
			'    txtServidor.SetFocus
			
			'UPGRADE_WARNING: Couldn't resolve default property of object txtCuenta.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gstrLogin = UCase(txtCuenta.Text)
			
			' Verifica si tiene acceso a este modulo
			strSQL = "select * from Usuario where CveUsuario = '" & gstrLogin & "'"
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsPassword = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
			If rsPassword.EOF Then
				MsgBox(" Cuenta no existe ")
				rsPassword.Close()
				End
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object txtPassword.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If UCase(Trim(rsPassword.rdoColumns("PASSWORD").Value)) <> UCase(Trim(txtPassword.Text)) Then
				MsgBox(" Password es incorrecto ")
				rsPassword.Close()
				End
			End If
			
			strSQL = "select * from UsuarioAplicacion where CveUsuario = '" & gstrLogin
			strSQL = strSQL & "' and CveAplicacion = " & APLICACION
			'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsPermiso = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
			If rsPermiso.EOF Then
				MsgBox("No se tiene acceso a este Módulo de SIM")
				rsPermiso.Close()
				End
			End If
			rsPermiso.Close()
			
			'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			LlenaVariosSelectores("SELECT B.CveBase,B.Nombre FROM Base B, UsuarioBase UB " & "WHERE B.CveBase = UB.CveBase" & "  AND UB.CveUsuario = '" & gstrLogin & "' " & "ORDER BY B.Nombre", New Object(){"cboBase"}, Me)
			'UPGRADE_WARNING: Couldn't resolve default property of object cboBase.SetFocus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			cboBase.SetFocus()
			
		End If
		
	End Sub
	Private Sub txtPassword_LostFocus()
		Dim cboBase As Object
		Dim txtPassword As Object
		Dim txtCuenta As Object
		Dim strSQL As String
		Dim rsPermiso As RDO.rdoResultset
		Dim rsPassword As RDO.rdoResultset
		
		'UPGRADE_WARNING: Couldn't resolve default property of object txtCuenta.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gstrLogin = UCase(txtCuenta.Text)
		
		' Verifica si tiene acceso a este modulo
		strSQL = "select * from Usuario where CveUsuario = '" & gstrLogin & "'"
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsPassword = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		If rsPassword.EOF Then
			MsgBox(" Cuenta no existe ")
			rsPassword.Close()
			End
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object txtPassword.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If UCase(Trim(rsPassword.rdoColumns("PASSWORD").Value)) <> UCase(Trim(txtPassword.Text)) Then
			MsgBox(" Password es incorrecto ")
			rsPassword.Close()
			'UPGRADE_WARNING: Couldn't resolve default property of object txtPassword.SetFocus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			txtPassword.SetFocus()
			Exit Sub
		End If
		
		strSQL = "select * from UsuarioAplicacion where CveUsuario = '" & gstrLogin
		strSQL = strSQL & "' and CveAplicacion = " & APLICACION
		'UPGRADE_WARNING: Couldn't resolve default property of object gcn.OpenResultset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsPermiso = gcn.OpenResultset(strSQL, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		If rsPermiso.EOF Then
			MsgBox("No se tiene acceso a este Módulo de SIM")
			rsPermiso.Close()
			End
		End If
		rsPermiso.Close()
		
		'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		LlenaVariosSelectores("SELECT B.CveBase,B.Nombre FROM Base B, UsuarioBase UB " & "WHERE B.CveBase = UB.CveBase" & "  AND UB.CveUsuario = '" & gstrLogin & "' " & "ORDER BY B.Nombre", New Object(){"cboBase"}, Me)
		'UPGRADE_WARNING: Couldn't resolve default property of object cboBase.SetFocus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cboBase.SetFocus()
		
	End Sub
	Sub frmLogin_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'*** Code added by VB HelpWriter ***
		'*** Subroutine added by VB HelpWriter ***
		'QuitHelp
		'***********************************
	End Sub
	
	Private Sub txtBuscar_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBuscar.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = System.Windows.Forms.Keys.Return Then cmdBuscarMecanico_Click(cmdBuscarMecanico, New System.EventArgs())
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class
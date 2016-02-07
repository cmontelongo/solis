Option Strict Off

Public Class frmArticulo

    Dim mrs As New ADODB.Recordset
    Dim mblnEdicion As Boolean
    Dim mblnAlta As Boolean
    Dim mblnLlena As Boolean
    Dim mblnRenglonSeleccionado As Boolean
    Dim mblnGeneroCodigoArticulo As Boolean
    Dim marrEstatusSuspension() As Boolean
    Dim marrEstatusActivo() As Boolean
    Dim mvntMarca As VariantType
    Dim mlngLlave As Long

    Private Sub frmArticulo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim strSQL As String

        gstrLogin = "SICIP"
        gstrPassword = "SICIP"
        gstrServidor = "NAUTILIUS" 'BuscaParametrosIni("Datos Generales", "Servidor")

        AbreConeccion()

        Try
            mblnLlena = True
            strSQL = "Select CveFamilia,Nombre from Familia ORDER BY Nombre"
            LlenaSelector(cboFamilia, strSQL)
            strSQL = "Select CveUnidadMedida,Nombre from UnidadMedida ORDER BY Nombre"
            LlenaSelector(cboUnidadMedidaCompra, strSQL)
            strSQL = "Select CveUnidadMedida,Nombre from UnidadMedida ORDER BY Nombre"
            LlenaSelector(cboUnidadMedidaCotizacion, strSQL)
            strSQL = "Select CveUnidadMedida,Nombre from UnidadMedida ORDER BY Nombre"
            LlenaSelector(cboUnidadMedidaInv, strSQL)
            strSQL = "Select CveTipoRecurso,Nombre from TipoRecurso ORDER BY Nombre"
            LlenaSelector(cboTipoRecurso, strSQL)
            strSQL = "Select CveMoneda,Nombre from Moneda ORDER BY Nombre"
            LlenaSelector(cboMoneda, strSQL)

            LlenaSelectorEstatus()

            rdbTodos.PerformClick()

            ToolBar_EstadoBrowse()
            mblnLlena = False
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try

    End Sub
    Private Sub lstRegistros_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstRegistros.Click

        PosicionaRegistro(lstRegistros.SelectedValue.ToString)
        CargaControlesdeResultset()

    End Sub
    Private Sub CargaControlesdeResultset()
        '******************************
        'Despliega los Datos del Registro en la Pantalla
        'carga los controles con la información obtenida de la db en el recordset
        '******************************

        Dim rs As ADODB.Recordset
        Dim strSQL As String

        Me.Cursor = Cursors.WaitCursor

        mblnLlena = True
        InicializaCampos()

        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        rs.LockType = ADODB.LockTypeEnum.adLockOptimistic

        strSQL = "SELECT * FROM Articulo WHERE CveArticulo = "
        rs.Open(strSQL, gcn)

        If rs.RecordCount <> 0 Then
            txtArticulo.Text = rs.Fields("CveArticulo").Value
            If Not IsDBNull(rs.Fields("Codigo").Value) Then
                btnGeneraCodigo.Enabled = False
                txtCodigo.Text = rs.Fields("Codigo").Value
            End If
            txtNombre.Text = rs.Fields("Nombre").Value
            If IsDBNull(rs.Fields("NombreCorto").Value) Then
                txtNombreCorto.Text = Mid(rs.Fields("Nombre").Value, 1, 35)
            Else
                txtNombreCorto.Text = rs.Fields("NombreCorto").Value
            End If

            If IsDBNull(rs.Fields("FactorConversion").Value) Then
                txtFactor.Text = 1
            Else
                txtFactor.Text = rs.Fields("FactorConversion").Value
            End If

            cboFamilia.SelectedValue = Trim(Str(rs.Fields("CveFamilia").Value))
            cboUnidadMedidaInv.SelectedValue = Trim(Str(rs.Fields("CveUnidadMedidaInventario").Value))
            cboUnidadMedidaCompra.SelectedValue = Trim(Str(rs.Fields("CveUnidadMedidaCompra").Value))
            chkEsAlmacenable.CheckState = CheckState.Unchecked
            If rs.Fields("EsAlmacenable").Value Then
                chkEsAlmacenable.CheckState = CheckState.Checked
            End If
            chkRequiereArmado.CheckState = CheckState.Unchecked
            If rs.Fields("RequiereArmado").Value Then
                chkRequiereArmado.CheckState = CheckState.Checked
            End If
            chkManufacturado.CheckState = CheckState.Unchecked
            If rs.Fields("EsManufacturado").Value Then
                chkManufacturado.CheckState = CheckState.Checked
            End If
            cboArticuloEstatus.SelectedValue = Trim(Str(rs.Fields("CveArticuloEstatus").Value))
            txtFechaSuspension.Enabled = marrEstatusSuspension(cboArticuloEstatus.SelectedValue)
            txtCausaSuspension.Enabled = marrEstatusSuspension(cboArticuloEstatus.SelectedValue)

            cboTipoRecurso.SelectedValue = Trim(Str(rs.Fields("CveTipoRecurso").Value))
            cboMoneda.SelectedValue = Trim(Str(rs.Fields("CveMoneda").Value))
            cboUnidadMedidaCotizacion.SelectedValue = Trim(Str(rs.Fields("CveUnidadMedidaCotizacion").Value))
            If Not IsDBNull(rs.Fields("PrecioCompra").Value) Then txtPrecioCompra.Text = rs.Fields("PrecioCompra").Value
            If Not IsDBNull(rs.Fields("PrecioLista").Value) Then txtPrecioLista.Text = rs.Fields("PrecioLista").Value
            If Not IsDBNull(rs.Fields("FechaUltimoPrecioCompra").Value) Then txtFechaCompra.Text = rs.Fields("FechaUltimoPrecioCompra").Value
            If Not IsDBNull(rs.Fields("FechaUltimoPrecioLista").Value) Then txtFechaLista.Text = rs.Fields("FechaUltimoPrecioLista").Value
            If Not IsDBNull(rs.Fields("KgPorM2").Value) Then txtKGxM2.Text = rs.Fields("KgPorM2").Value

            'Activo()
            'FechaBaja()
            'Notas()
            'CveUsuarioCreador()
            'FechaAlta()
            'CveUsuarioModifico()
            'FechaModificacion()
            'CalculoPrecioLista()

            CargaDetalle()

        End If

        rs.Close()
        rs = Nothing

            '        Dim lngIndiceTemporal As Long
            '       Dim strSQL As String
            '
            '       If rsODT.RowCount <> 0 Then
            'txtCveBase.Text = rsODT!Base
            '       txtCvePersonal.Text = rsODT!CvePersonal
            '       txtNombre.Text = rsODT!Nombre
            '       If Not IsNull(rsODT!Ubicacion) Then txtCveUbicacion.Text = rsODT!Ubicacion
            '        If Not IsNull(rsODT!Departamento) Then txtCveDepartamento.Text = rsODT!Departamento
            '      If Not IsNull(rsODT!Puesto) Then txtCvePuesto.Text = rsODT!Puesto
            '       If Not IsNull(rsODT!NombreJefe) Then txtSuperior.Text = rsODT!NombreJefe
            '     chkActivo.Value = Abs(rsODT!Activo)
            '     If Not IsNull(rsODT!FechaIngreso) Then txtFechaIngreso.Text = Format(rsODT!FechaIngreso, FECHADDMMYY)
            '     If Not IsNull(rsODT!FechaBaja) Then txtFechaBaja.Text = Format(rsODT!FechaBaja, FECHADDMMYY)
            '    If Not IsNull(rsODT!CorreoElectronico) Then txtEmail.Text = rsODT!CorreoElectronico
            '    End If

            '    mblnEdicion = False
            '    DespliegaUbicacionRegistro(rsODT, txtNumRegistros)

            ' Despliega las Tareas de la orden
            '   DespliegaAmonestaciones()

            '   lstPersonal.Enabled = True
        mblnLlena = False
        Me.Cursor = Cursors.Default
            ' Exit Sub

            'Err_Carga:
            '      Screen.MousePointer = vbDefault
            'MsgBox "Error al Cargar Controles con el rdoResultset " & Error, vbCritical
            'mblnEdicion = False
            '  Exit Sub
            '  Resume Next
    End Sub
    Private Sub CargaDetalle()

        Dim rsDetalle As New ADODB.Recordset()
        Dim intTipoRecursoAnterior As Integer
        Dim lngRenglon As Long
        Dim strSQL As String
        Dim x As Boolean

        rsDetalle.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsDetalle.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        rsDetalle.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic
        Try
            intTipoRecursoAnterior = 0
            LimpiaBloque(sprManufactura, 1, 1, sprManufactura.MaxRows, sprManufactura.MaxCols)
            sprManufactura.MaxRows = 0
            sprManufactura.EditModePermanent = True
            sprManufactura.ReDraw = False

            strSQL = "SELECT AM.CveArticuloDetalle,A.Nombre,AM.CantidadRequerida,A.CveTipoRecurso, TR.Nombre NombreTipoRecurso " & _
                "FROM ArticuloManufactura AM " & _
                    "JOIN Articulo A ON AM.CveArticuloDetalle = A.CveArticulo " & _
                    "JOIN TipoRecurso TR ON A.CveTipoRecurso = TR.CveTipoRecurso " & _
                "WHERE AM.CveArticulo = " & txtArticulo.Text & " ORDER BY AM.NumRenglon"
            rsDetalle.Open(strSQL, gcn)

            Do Until rsDetalle.EOF
                sprManufactura.MaxRows = sprManufactura.MaxRows + 1
                lngRenglon = sprManufactura.MaxRows
                sprManufactura.Row = lngRenglon

                If intTipoRecursoAnterior <> rsDetalle.Fields("CveTipoRecurso").Value Then
                    With sprManufactura
                        .Row = lngRenglon
                        .Col = 1
                        .BackColor = Color.Violet
                        .ForeColor = Color.White
                        .FontBold = True
                        .Text = rsDetalle.Fields("NombreTipoRecurso").Value
                    End With

                    x = sprManufactura.AddCellSpan(1, sprManufactura.Row, 2, 1)
                    sprManufactura.TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignCenter

                    intTipoRecursoAnterior = rsDetalle.Fields("CveTipoRecurso").Value
                    sprManufactura.MaxRows = sprManufactura.MaxRows + 1
                    sprManufactura.Row = sprManufactura.MaxRows
                End If

                sprManufactura.Col = 1
                sprManufactura.Text = rsDetalle.Fields("Nombre").Value

                MakeFloatCell(2, 2, sprManufactura.Row, sprManufactura.Row, "-99999", "99999", False, True, 2, 0, sprManufactura)
                sprManufactura.Col = 2
                sprManufactura.Text = rsDetalle.Fields("CantidadRequerida").Value

                sprManufactura.Col = 3
                sprManufactura.Text = rsDetalle.Fields("CveArticuloDetalle").Value

                rsDetalle.MoveNext()
            Loop

            sprManufactura.EditModePermanent = False
            sprManufactura.ReDraw = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Metodo CargarDetalle", MessageBoxButtons.OK)
        Finally
            rsDetalle.Close()
        End Try
    End Sub
    Private Sub InicializaCampos()
        Me.Cursor = Cursors.WaitCursor

        'limpia controles para proxima captura
        txtArticulo.Text = ""
        txtCodigo.Text = ""
        txtNombre.Text = ""
        txtNombreCorto.Text = ""
        txtFactor.Text = 1
        cboFamilia.SelectedValue = -1
        cboUnidadMedidaInv.SelectedValue = -1
        cboUnidadMedidaCompra.SelectedValue = -1
        chkEsAlmacenable.CheckState = CheckState.Unchecked
        chkRequiereArmado.CheckState = CheckState.Unchecked
        chkManufacturado.CheckState = CheckState.Unchecked
        cboArticuloEstatus.SelectedValue = "1"
        cboTipoRecurso.SelectedValue = -1
        cboMoneda.SelectedValue = -1
        cboUnidadMedidaCotizacion.SelectedValue = -1
        txtPrecioCompra.Text = ""
        txtPrecioLista.Text = ""
        txtFechaCompra.Text = ""
        txtFechaLista.Text = ""
        txtKGxM2.Text = "0"
        mblnGeneroCodigoArticulo = False
        btnGeneraCodigo.Enabled = True
        txtFechaSuspension.Enabled = False
        txtCausaSuspension.Enabled = False

        LimpiaBloque(sprManufactura, 1, 1, sprManufactura.MaxRows, sprManufactura.MaxCols)
        sprManufactura.MaxRows = 0

        Me.Cursor = Cursors.Default
    End Sub
    Private Sub BuscaClave(ByVal vstrClave As String)

        Dim lngLlave As Long

        Me.Cursor = Cursors.WaitCursor

        lngLlave = txtArticulo.Text

        PosicionaRegistro(vstrClave)
        If mrs.EOF Then
            PosicionaRegistro(lngLlave)
            Me.Cursor = Cursors.Default
            Exit Sub
        Else
            CargaControlesdeResultset()
            'ToolBar_EstadoBrowse(tlbODT)
            lstRegistros.Enabled = True
        End If

        mblnEdicion = False
        Me.Cursor = Cursors.Default
        Exit Sub
    End Sub
    Public Sub PosicionaRegistro(ByVal vntValorABuscar As Object)
        '--------------------------------------------------------------------
        '   Rutina para posicionar un rdoResultset o rdoResultset                 '
        '   en determindado valor de la llave                               '
        '       Entrada.-                                                   '
        '                vntValorABuscar ->  Valor a Buscar                 '
        '-------------------------------------------------------------------'

        Dim blnExiste As Boolean

        blnExiste = False
        If mrs.RecordCount > 0 Then
            mrs.MoveFirst()
            Do Until mrs.EOF
                If Str(mrs.Fields("CveArticulo").Value) = Str(vntValorABuscar) Then
                    blnExiste = True
                    Exit Sub
                End If
                mrs.MoveNext()
            Loop
        End If
    End Sub
    Private Function ValidaCampos() As Boolean

        ValidaCampos = False

        If Len(txtNombre.Text) = 0 Then
            Me.Cursor = Cursors.Default
            MsgBox("Debes especificar una descripcion para el Articulo", MsgBoxStyle.Exclamation, "Valida Campos")
            txtNombre.Focus()
            Exit Function
        End If

        If Len(txtNombreCorto.Text) = 0 Then
            Me.Cursor = Cursors.Default
            MsgBox("Debes especificar una descripcion Corta para el Articulo", MsgBoxStyle.Exclamation, "Valida Campos")
            txtNombreCorto.Focus()
            Exit Function
        End If

        If cboFamilia.SelectedValue Is Nothing Then
            Me.Cursor = Cursors.Default
            MsgBox("Debes especificar una familia para el Articulo", MsgBoxStyle.Exclamation, "Valida Campos")
            cboFamilia.Focus()
            Exit Function
        End If

        If Not IsNumeric(txtFactor.Text) Then
            Me.Cursor = Cursors.Default
            MsgBox("Debes especificar un valor para el factor de inventarios", MsgBoxStyle.Exclamation, "Valida Campos")
            txtFactor.Focus()
            Exit Function
        End If

        If cboUnidadMedidaCompra.SelectedValue Is Nothing Then
            Me.Cursor = Cursors.Default
            MsgBox("Debes especificar una Unidad de medida para las compras.", MsgBoxStyle.Exclamation, "Valida Campos")
            cboUnidadMedidaCompra.Focus()
            Exit Function
        End If


        If cboUnidadMedidaInv.SelectedValue Is Nothing Then
            Me.Cursor = Cursors.Default
            MsgBox("Debes especificar una Unidad de medida para el inventario.", MsgBoxStyle.Exclamation, "Valida Campos")
            cboUnidadMedidaInv.Focus()
            Exit Function
        End If

        If cboArticuloEstatus.SelectedValue Is Nothing Then
            Me.Cursor = Cursors.Default
            MsgBox("Debes especificar un status para el articulo.", MsgBoxStyle.Exclamation, "Valida Campos")
            cboArticuloEstatus.Focus()
            Exit Function
        End If

        If cboMoneda.SelectedValue Is Nothing Then
            Me.Cursor = Cursors.Default
            MsgBox("Debes especificar una moneda para el articulo.", MsgBoxStyle.Exclamation, "Valida Campos")
            cboMoneda.Focus()
            Exit Function
        End If

        If txtCausaSuspension.Enabled = True And Len(txtCausaSuspension.Text) = 0 Then
            Me.Cursor = Cursors.Default
            MsgBox("Debes especificar un motivo para suspender el articulo.", MsgBoxStyle.Exclamation, "Valida Campos")
            txtCausaSuspension.Focus()
            Exit Function
        End If
        ValidaCampos = True

    End Function
    Private Sub tabPrincipal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabPrincipal.Click

        If InStr(Me.Text, "-") > 0 Then Me.Text = Trim(Mid(Me.Text, 1, InStr(Me.Text, "-") - 1))
        If tabPrincipal.SelectedIndex <> 0 Then Me.Text = Me.Text & " - " & txtNombre.Text

    End Sub
    Private Sub BuscaArticulo()
        Dim strSQL As String
        Dim strCondicion As String
        Dim lngRenglon As Long

        Me.Cursor = Cursors.WaitCursor

        strCondicion = txtArticulo.Text
        For lngRenglon = 1 To sprManufactura.MaxRows
            If Len(strCondicion) > 0 And Len(sprManufactura.Text) > 0 Then strCondicion = strCondicion & ","
            sprManufactura.Col = 3
            sprManufactura.Row = lngRenglon
            If Len(sprManufactura.Text) > 0 Then strCondicion = strCondicion & sprManufactura.Text
        Next
        If Len(strCondicion) > 0 Then strCondicion = "AND CveArticulo NOT IN(" & strCondicion & ")"

        strSQL = "SELECT CveArticulo,Nombre FROM Articulo WHERE Activo = 1 AND Nombre like '%" & Replace(txtBuscar.Text, " ", "%") & "%' " & strCondicion & " ORDER BY Nombre"
        LlenaSelector(cboArticulo, strSQL)
        If cboArticulo.Items.Count > 0 Then
            cboArticulo.Visible = True
            txtBuscar.Visible = False
            btnBuscar.Visible = False
            btnLimpiar.Visible = True
            btnAgregar.Visible = True
            txtBuscar.Text = ""
        End If
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub btnBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBuscar.Click
        BuscaArticulo()
    End Sub

    Private Sub txtBuscar_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBuscar.KeyPress
        If Asc(e.KeyChar) = System.Windows.Forms.Keys.Return Then BuscaArticulo()
    End Sub

    Private Sub btnLimpiar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLimpiar.Click
        btnAgregar.Visible = False
        btnLimpiar.Visible = False
        cboArticulo.Visible = False

        txtBuscar.Visible = True
        btnBuscar.Visible = True
    End Sub

    Private Sub btnAgregar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAgregar.Click

        Dim strSQL As String
        Dim rsdetalle As New ADODB.Recordset
        Dim lngRenglon As Long
        Dim blnEncontroSeccion As Boolean
        Dim strSeccion As String

        If Len(cboArticulo.Text) >= 0 Then

            'Always have the spreadsheet in edit mode
            sprManufactura.EditModePermanent = True
            sprManufactura.ReDraw = False

            strSQL = "SELECT A.CveArticulo,A.Nombre,A.CveTipoRecurso, ISNULL(TR.Nombre,'') NombreTipoRecurso " & _
                "FROM Articulo A " & _
                    "LEFT JOIN TipoRecurso TR ON A.CveTipoRecurso = TR.CveTipoRecurso " & _
                "WHERE A.CveArticulo = " & cboArticulo.SelectedValue
            rsdetalle.Open(strSQL, gcn)

            Do Until rsdetalle.EOF
                sprManufactura.MaxRows = sprManufactura.MaxRows + 1
                sprManufactura.Row = sprManufactura.MaxRows

                For lngRenglon = 1 To sprManufactura.MaxRows
                    sprManufactura.Col = 1
                    sprManufactura.Row = lngRenglon
                    strSeccion = sprManufactura.Text

                    sprManufactura.Col = 3
                    If sprManufactura.Text = "" Then

                        If strSeccion = rsdetalle.Fields("NombreTipoRecurso").Value Then
                            blnEncontroSeccion = True
                        Else
                            If blnEncontroSeccion Then 'Si encontro la seccion anteriormente agrega el espacio
                                sprManufactura.InsertRows(lngRenglon, 1)
                                Exit For
                            End If
                        End If

                    End If

                Next lngRenglon
                sprManufactura.Row = lngRenglon

                sprManufactura.Col = 1
                sprManufactura.Text = rsdetalle.Fields("Nombre").Value

                MakeFloatCell(2, 2, sprManufactura.Row, sprManufactura.Row, "-99999", "99999", False, True, 2, 0, sprManufactura)
                sprManufactura.Col = 2
                sprManufactura.Text = 0

                sprManufactura.Col = 3
                sprManufactura.Text = rsdetalle.Fields("CveArticulo").Value

                rsdetalle.MoveNext()
            Loop
            rsdetalle.Close()
            rsdetalle = Nothing

            sprManufactura.EditModePermanent = False
            sprManufactura.ReDraw = True

            btnAgregar.Visible = False
            cboArticulo.Visible = False
            btnLimpiar.Visible = False

            txtBuscar.Visible = True
            btnBuscar.Visible = True

        End If
    End Sub

    Private Sub btnArriba_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnArriba.Click
        Dim lngRenglonAMover As Long
        If sprManufactura.SelBlockRow = sprManufactura.SelBlockRow2 Then
            lngRenglonAMover = sprManufactura.SelBlockRow
        End If
        sprManufactura.MaxRows = sprManufactura.MaxRows + 1
        sprManufactura.InsertRows(lngRenglonAMover - 1, 1)
        sprManufactura.MoveRowRange(lngRenglonAMover + 1, lngRenglonAMover + 1, lngRenglonAMover - 1)
        sprManufactura.DeleteRows(lngRenglonAMover + 1, 1)
        sprManufactura.MaxRows = sprManufactura.MaxRows - 1
    End Sub

    Private Sub sprManufactura_BlockSelected(ByVal sender As Object, ByVal e As AxFPSpread._DSpreadEvents_BlockSelectedEvent) Handles sprManufactura.BlockSelected
        mblnRenglonSeleccionado = True
    End Sub

    Private Sub btnGeneraCodigo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGeneraCodigo.Click

        Dim rsDetalle As New ADODB.Recordset()
        Dim strSQL As String
        Dim intCveFamilia As Integer
        Dim intConsecutivo As Integer

        rsDetalle.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsDetalle.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        rsDetalle.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic

        Try
            If cboFamilia.SelectedValue Is Nothing Then
                MsgBox("Es necesario especificar una familia para el articulo", MsgBoxStyle.Critical, "Genera Codigo")
                cboFamilia.Focus()
                Exit Sub
            End If

            intCveFamilia = cboFamilia.SelectedValue
            strSQL = "select F.CveFamilia,F.IdFamilia,MAX(CONVERT(smallint,SUBSTRING(A.Codigo,LEN(F.IdFamilia)+2,LEN(A.Codigo) - LEN(F.IdFamilia)))) Maximo " & _
                "FROM Familia F " & _
                     "LEFT JOIN Articulo A ON F.IdFamilia = SUBSTRING(A.Codigo,1,LEN(F.IdFamilia)) " & _
                "WHERE F.CveFamilia = " & intCveFamilia & _
                " GROUP BY F.CveFamilia,F.IdFamilia"

            rsDetalle.Open(strSQL, gcn)
            If IsDBNull(rsDetalle.Fields("Maximo").Value) Then
                intConsecutivo = 1
            Else
                intConsecutivo = rsDetalle.Fields("Maximo").Value + 1
            End If

            txtCodigo.Text = rsDetalle.Fields("IdFamilia").Value & "-" & Strings.Right("0000" & intConsecutivo, 5)

            mblnGeneroCodigoArticulo = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Metodo btnGeneraCodigo", MessageBoxButtons.OK)
        Finally
            rsDetalle.Close()
        End Try

    End Sub
    Private Sub ActualizaInterfaz(ByVal vlngCveArticulo As Long, ByVal vblnNuevo As Boolean)

        Dim strConnection As String
        Dim rsConsulta As New ADODB.Recordset
        Dim strSQL As String
        Dim strUpdate As String
        Dim strInsertD_LISPRE As String = ""
        Dim strinsertART_ALMA As String = ""
        Dim strArchivo As String

        strArchivo = BuscaParametrosIni("Interfaz " & BuscaParametrosIni("Datos Generales", "Interfaz"), "Directorio") & "\articulo.DBF"

        strConnection = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=" & _
                System.IO.Path.GetDirectoryName(strArchivo) & ";"

        rsConsulta.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsConsulta.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
        rsConsulta.LockType = ADODB.LockTypeEnum.adLockOptimistic

        strSQL = "SELECT F.Codigo ART_FAMILI" & _
                ",UM.Codigo ART_UNDINV " & _
                ",F.Codigo2 ART_SUBFAM " & _
                ",F.Codigo3 ART_TIPO " & _
                "FROM Articulo A " & _
                 "JOIN Familia F ON F.CveFamilia = A.CveFamilia " & _
                 "JOIN UnidadMedida UM ON UM.CveUnidadMedida = A.CveUnidadMedidaInventario " & _
            "WHERE A.CveArticulo = " & vlngCveArticulo

        Using dbConn As New System.Data.Odbc.OdbcConnection(strConnection)
            Try

                rsConsulta.Open(strSQL, gcn)

                If (vblnNuevo And Len(txtCodigo.Text) = 0) Then
                    If mblnGeneroCodigoArticulo Then
                        strUpdate = "INSERT " & System.IO.Path.GetFileNameWithoutExtension(strArchivo) & _
                            "(ART_CLAVE,ART_DESC,ART_DESC2,ART_BARCOD,ART_REFER,ART_COSTO" & _
                             ",ART_COSREP,ART_UNDINV,ART_TIPO,ART_DIAS,ART_ENSAMB" & _
                             ",ART_FAMILI,ART_SUBFAM,ART_MARCA" & _
                             ",ART_ALMVTA,ART_ALMVTP,ART_RETCVE,ART_RETVAL,ART_RETUND,ART_LISPRE,ART_ALMUSA" & _
                             ",ART_IVA,ART_PRECIO,ART_PRECIP,ART_CARGO,ART_FACVTA) " & _
                             "VALUES('" & txtCodigo.Text & "','" & txtNombreCorto.Text & "','" & txtNombre.Text & "','',''," & CDbl(Replace(txtPrecioCompra.Text, ",", "")) & _
                                    ",0,'" & rsConsulta.Fields("ART_UNDINV").Value & "','" & rsConsulta.Fields("ART_TIPO").Value & "',0,''" & _
                                    ",'" & rsConsulta.Fields("ART_FAMILI").Value & "','" & rsConsulta.Fields("ART_SUBFAM").Value & "','SIN'" & _
                                    ",'A-CON','A-CON','',0,0,'SIN        / LISTA BASE',''" & _
                                    ",16,0,0,0,1)"

                        strInsertD_LISPRE = "INSERT D_LISPRE (LIS_DESC,LIS_MARCA,LIS_ARTIC,LIS_PRECIO,LIS_PRECI2) " & _
                            "VALUES('LISTA BASE','','" & txtCodigo.Text & "'," & CDbl(Replace(txtPrecioCompra.Text, ",", "")) & ",0)"

                        ',ART_OCOM,ART_INVMIN,ART_INVMAX,ART_REORD,ART_COSTO
                        strinsertART_ALMA = "INSERT ART_ALMA (ART_ALMA,ART_CLAVE,ART_INVE) " & _
                            "VALUES('A-CON','','" & txtCodigo.Text & "',0)"
                    Else
                        strUpdate = ""
                    End If
                Else
                    strUpdate = "update " & System.IO.Path.GetFileNameWithoutExtension(strArchivo) & _
                            " SET ART_DESC='" & txtNombreCorto.Text & "'" & _
                                ",ART_DESC2='" & txtNombre.Text & "'" & _
                                ",ART_FAMILI ='" & rsConsulta.Fields("ART_FAMILI").Value & "'" & _
                                ",ART_SUBFAM='" & rsConsulta.Fields("ART_SUBFAM").Value & "'" & _
                                ",ART_UNDINV ='" & rsConsulta.Fields("ART_UNDINV").Value & "'" & _
                                ",ART_COSTO=" & CDbl(Replace(txtPrecioCompra.Text, ",", "")) & _
                                ",ART_TIPO='" & rsConsulta.Fields("ART_TIPO").Value & "' " & _
                        "WHERE ART_CLAVE = '" & txtCodigo.Text & "'"
                End If

                rsConsulta.Close()

                If strUpdate <> "" Then
                    dbConn.Open()

                    Dim cmdInstruccion As New System.Data.Odbc.OdbcCommand(strUpdate, dbConn)

                    cmdInstruccion.ExecuteNonQuery()

                    cmdInstruccion = Nothing


                    If strInsertD_LISPRE <> "" Then
                        Dim cmdInstruccionD_LISPRE As New System.Data.Odbc.OdbcCommand(strInsertD_LISPRE, dbConn)

                        cmdInstruccionD_LISPRE.ExecuteNonQuery()

                        cmdInstruccionD_LISPRE = Nothing

                        Dim cmdInstruccionART_ALMA As New System.Data.Odbc.OdbcCommand(strinsertART_ALMA, dbConn)

                        cmdInstruccionART_ALMA.ExecuteNonQuery()

                        cmdInstruccionART_ALMA = Nothing
                    End If
                    dbConn.Close()
                End If

            Catch ex As Exception
                MessageBox.Show("Error al abrir la base de datos" & vbCrLf & ex.Message)

                Exit Sub
            End Try
        End Using
    End Sub
    Private Sub ToolBar_EstadoCambio()
        '*************************************************************
        ' Rutina para poner todos los botones del Toolbar de un ABC
        ' en el estado de cambio
        '*************************************************************
        ' DesHabilita todos Los Botones
        ToolBotones_Estado(tsToolBar, False)

        'Habilita los botones Actualizar y Cancelar
        ToolBoton_Estado(tsToolBar, "tsbActualizar", True)
        ToolBoton_Estado(tsToolBar, "tsbCancelar", True)

    End Sub
    Private Sub ToolBar_EstadoBrowse()
        '************************************************************
        ' Rutina para dejar el ToolBar de un ABC en el estado Browse,
        ' es decir, inhabilitados el grabar y el cancelar
        '************************************************************
        ' Habilita todos Los Botones
        ToolBotones_Estado(tsToolBar, True)

        'Deshabilita los botones Actualizar y Cancelar
        ToolBoton_Estado(tsToolBar, "tsbActualizar", False)
        ToolBoton_Estado(tsToolBar, "tsbCancelar", False)

    End Sub
    Private Sub Agrega()

        InicializaCampos()

        mblnEdicion = False
        mblnAlta = True

        txtNombre.Focus()

        ToolBar_EstadoCambio()
    End Sub

    Private Sub tsbNuevo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbNuevo.Click
        Me.Cursor = Cursors.WaitCursor
        Agrega()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub LlenaSelectorEstatus()

        Dim rsfLlenaControl As New ADODB.Recordset()
        Dim Lista As New ArrayList

        rsfLlenaControl.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsfLlenaControl.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        rsfLlenaControl.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic

        rsfLlenaControl.Open("SELECT * FROM ArticuloEstatus ORDER BY Nombre", gcn)

        Try
            ReDim marrEstatusActivo(rsfLlenaControl.RecordCount + 1)
            ReDim marrEstatusSuspension(rsfLlenaControl.RecordCount + 1)
            Do Until rsfLlenaControl.EOF
                Lista.Add(New clsSelector(rsfLlenaControl.Fields.Item("CveArticuloEstatus").Value, rsfLlenaControl.Fields.Item("Nombre").Value))
                marrEstatusActivo(rsfLlenaControl.Fields.Item("CveArticuloEstatus").Value) = rsfLlenaControl.Fields.Item("DeshabilitaArticulo").Value
                marrEstatusSuspension(rsfLlenaControl.Fields.Item("CveArticuloEstatus").Value) = rsfLlenaControl.Fields.Item("HabilitaSuspension").Value
                rsfLlenaControl.MoveNext()
            Loop

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Metodo LlenaSelectorEstatus", MessageBoxButtons.OK)
        Finally
            rsfLlenaControl.Close()
        End Try


        With cboArticuloEstatus
            .DataSource = Lista
            .ValueMember = "strClave"
            .DisplayMember = "strTexto"
        End With

    End Sub

    Private Sub cboArticuloEstatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboArticuloEstatus.SelectedIndexChanged

        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()

        txtFechaSuspension.Enabled = marrEstatusSuspension(cboArticuloEstatus.SelectedValue)
        txtCausaSuspension.Enabled = marrEstatusSuspension(cboArticuloEstatus.SelectedValue)

    End Sub


    Private Sub txtNombre_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNombre.TextChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub txtNombreCorto_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNombreCorto.TextChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub cboFamilia_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFamilia.SelectedIndexChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub cboUnidadMedidaCompra_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboUnidadMedidaCompra.SelectedIndexChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub txtFactor_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFactor.TextChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub cboUnidadMedidaInv_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboUnidadMedidaInv.SelectedIndexChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub chkEsAlmacenable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEsAlmacenable.CheckedChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub chkRequiereArmado_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRequiereArmado.CheckedChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub chkManufacturado_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkManufacturado.CheckedChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub dtpFechaSuspension_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub txtCausaSuspension_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCausaSuspension.TextChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub cboMoneda_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMoneda.SelectedIndexChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub txtPrecioCompra_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrecioCompra.TextChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub txtPrecioLista_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrecioLista.TextChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub cboUnidadMedidaCotizacion_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboUnidadMedidaCotizacion.SelectedIndexChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub cboTipoRecurso_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTipoRecurso.SelectedIndexChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub txtKGxM2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKGxM2.TextChanged
        If mblnAlta Or mblnLlena Then Exit Sub
        mblnEdicion = True
        ToolBar_EstadoCambio()
    End Sub

    Private Sub tsbActualizar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbActualizar.Click
        If ValidaCampos() Then
            Actualiza()
        End If
    End Sub
    Private Sub Actualiza()

        Dim rsfLlenaControl As New ADODB.Recordset()
        Dim strSQL As String

        rsfLlenaControl.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsfLlenaControl.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        rsfLlenaControl.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic

        Me.Cursor = Cursors.WaitCursor
        Try
            strSQL = "EXEC Articulo_PROCESO_Actualiza " & _
                                    "@CveArticulo = "
            If Len(Trim(txtArticulo.Text)) = 0 Then
                strSQL = strSQL & "NULL"
            Else
                strSQL = strSQL & txtArticulo.Text
            End If
            strSQL = strSQL & ",@Nombre ='" & txtNombre.Text & "'" & _
                        ",@NombreCorto ='" & txtNombreCorto.Text & "'" & _
                        ",@CveFamilia =" & cboFamilia.SelectedValue & _
                        ",@EsAlmacenable =" & chkEsAlmacenable.CheckState & _
                        ",@RequiereArmado =" & chkRequiereArmado.CheckState & _
                        ",@EsManufacturado =" & chkManufacturado.CheckState & _
                        ",@CveArticuloEstatus =" & cboArticuloEstatus.SelectedValue & _
                        ",@CausaSuspension ='" & Trim(txtCausaSuspension.Text) & "'" & _
                        ",@CveUnidadMedidaInventario =" & cboUnidadMedidaInv.SelectedValue & _
                        ",@CveUnidadMedidaCompra =" & cboUnidadMedidaCompra.SelectedValue & _
                        ",@CveUsuario =" & "'SICIP'" & _
                        ",@CveMoneda =" & cboMoneda.SelectedValue & _
                        ",@PrecioCompra =" & Val(txtPrecioCompra.Text) & _
                        ",@PrecioLista =" & Val(txtPrecioLista.Text) & _
                        ",@FactorConversion =" & Val(txtFactor.Text) & _
                        ",@Codigo ="
            If Len(Trim(txtCodigo.Text)) = 0 Then
                strSQL = strSQL & "NULL"
            Else
                strSQL = strSQL & "'" & Trim(txtCodigo.Text) & "'"
            End If
            strSQL = strSQL & ",@CveTipoRecurso ="
            If cboTipoRecurso.SelectedValue Is Nothing Then
                strSQL = strSQL & "NULL"
            Else
                strSQL = strSQL & cboTipoRecurso.SelectedValue
            End If
            strSQL = strSQL & ",@KgPorM2 ="
            If Len(Trim(txtKGxM2.Text)) = 0 Then
                strSQL = strSQL & "0"
            Else
                strSQL = strSQL & txtKGxM2.Text
            End If
            strSQL = strSQL & ",@CveUnidadMedidaCotizacion ="
            If cboUnidadMedidaCotizacion.SelectedValue Is Nothing Then
                strSQL = strSQL & "NULL"
            Else
                strSQL = strSQL & cboUnidadMedidaCotizacion.SelectedValue
            End If

            rsfLlenaControl.Open(strSQL, gcn)
            If rsfLlenaControl.EOF Then
                MsgBox("Ocurrio un error al actualizar el articulo.", MsgBoxStyle.Exclamation, "Actualiza")
                Exit Sub
            End If

            If Len(Trim(txtArticulo.Text)) = 0 Then txtArticulo.Text = rsfLlenaControl.Fields("CveArticulo").Value

            ActualizaInterfaz(txtArticulo.Text, rsfLlenaControl.Fields("EsNuevo").Value)

            rsfLlenaControl.Close()

            rdbTodos.PerformClick()
            PosicionaRegistro(txtArticulo.Text)

        Catch ex As Exception
            Dim errLoop As ADODB.Error
            Dim errs1 As ADODB.Errors
            Dim strmsg As String
            Dim lngIndice As Long
            lngIndice = 1

            ' Proceso
            strmsg = "VB Error # " & Str(Err.Number)
            strmsg = strmsg & vbCrLf & "   Generado por " & Err.Source
            strmsg = strmsg & vbCrLf & "   Descripcion  " & Err.Description

            ' Enumera la coleccion de errores y despliega las propiedades de cada uno de los errores.
            errs1 = gcn.Errors
            For Each errLoop In errs1
                With errLoop
                    strmsg = strmsg & vbCrLf & "Error #" & lngIndice & ":"
                    strmsg = strmsg & vbCrLf & "   ADO Error   #" & .Number
                    strmsg = strmsg & vbCrLf & "   Descripcion  " & .Description
                    strmsg = strmsg & vbCrLf & "   Fuente       " & .Source
                    lngIndice = lngIndice + 1
                End With
            Next

            MsgBox(strmsg, MsgBoxStyle.Critical, "Actualiza")

        Finally
            ToolBar_EstadoBrowse()
            mblnAlta = False
            mblnEdicion = False
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub rdbTodos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdbTodos.Click
        Dim strSQL As String
        strSQL = "Select CveArticulo,Nombre from Articulo ORDER BY Nombre"
        LlenaSelector(lstRegistros, strSQL)
        lstRegistros.SelectedIndex = 0
        PosicionaRegistro(lstRegistros.SelectedValue.ToString)
        CargaControlesdeResultset()
    End Sub

    Private Sub rdbSinCodigo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdbSinCodigo.Click
        Dim strSQL As String
        strSQL = "Select CveArticulo,Nombre from Articulo WHERE Codigo IS NULL AND Activo = 1 ORDER BY Nombre"
        LlenaSelector(lstRegistros, strSQL)
        lstRegistros.SelectedIndex = 0
        PosicionaRegistro(lstRegistros.SelectedValue.ToString)
        CargaControlesdeResultset()
    End Sub

    Private Sub rdbBaja_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdbBaja.Click
        Dim strSQL As String
        strSQL = "Select CveArticulo,Nombre from Articulo WHERE Activo = 0 ORDER BY Nombre"
        LlenaSelector(lstRegistros, strSQL)
        lstRegistros.SelectedIndex = 0
        PosicionaRegistro(lstRegistros.SelectedValue.ToString)
        CargaControlesdeResultset()
    End Sub

End Class

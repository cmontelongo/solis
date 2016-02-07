<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmArticulo
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmArticulo))
        Me.gpbArticulo = New System.Windows.Forms.GroupBox
        Me.rdbBaja = New System.Windows.Forms.RadioButton
        Me.rdbSinCodigo = New System.Windows.Forms.RadioButton
        Me.rdbTodos = New System.Windows.Forms.RadioButton
        Me.lstRegistros = New System.Windows.Forms.ListBox
        Me.txtArticulo = New System.Windows.Forms.TextBox
        Me.tabPrincipal = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.txtFechaSuspension = New System.Windows.Forms.TextBox
        Me.lblFactor = New System.Windows.Forms.Label
        Me.txtFactor = New System.Windows.Forms.TextBox
        Me.sprProveedor = New AxFPSpread.AxvaSpread
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.grpCotizacion = New System.Windows.Forms.GroupBox
        Me.lblKgxM2 = New System.Windows.Forms.Label
        Me.txtKGxM2 = New System.Windows.Forms.TextBox
        Me.lblTipoRecurso = New System.Windows.Forms.Label
        Me.cboTipoRecurso = New System.Windows.Forms.ComboBox
        Me.lblUnidadCotiza = New System.Windows.Forms.Label
        Me.cboUnidadMedidaCotizacion = New System.Windows.Forms.ComboBox
        Me.lblNombreCorto = New System.Windows.Forms.Label
        Me.txtNombreCorto = New System.Windows.Forms.TextBox
        Me.btnGeneraCodigo = New System.Windows.Forms.Button
        Me.txtPrecioCompra = New System.Windows.Forms.TextBox
        Me.txtFechaLista = New System.Windows.Forms.TextBox
        Me.txtFechaCompra = New System.Windows.Forms.TextBox
        Me.lblPrecioLista = New System.Windows.Forms.Label
        Me.lblPrecioCompra = New System.Windows.Forms.Label
        Me.txtPrecioLista = New System.Windows.Forms.TextBox
        Me.lblMoneda = New System.Windows.Forms.Label
        Me.cboMoneda = New System.Windows.Forms.ComboBox
        Me.lblCodigo = New System.Windows.Forms.Label
        Me.txtCodigo = New System.Windows.Forms.TextBox
        Me.lblCausaSuspension = New System.Windows.Forms.Label
        Me.lblFechaSuspension = New System.Windows.Forms.Label
        Me.txtCausaSuspension = New System.Windows.Forms.TextBox
        Me.lblArticuloEstatus = New System.Windows.Forms.Label
        Me.cboArticuloEstatus = New System.Windows.Forms.ComboBox
        Me.chkManufacturado = New System.Windows.Forms.CheckBox
        Me.chkRequiereArmado = New System.Windows.Forms.CheckBox
        Me.chkEsAlmacenable = New System.Windows.Forms.CheckBox
        Me.lblUnidadVenta = New System.Windows.Forms.Label
        Me.lblUnidadCompra = New System.Windows.Forms.Label
        Me.lblFamilia = New System.Windows.Forms.Label
        Me.lblEtiqueta = New System.Windows.Forms.Label
        Me.cboUnidadMedidaCompra = New System.Windows.Forms.ComboBox
        Me.cboUnidadMedidaInv = New System.Windows.Forms.ComboBox
        Me.cboFamilia = New System.Windows.Forms.ComboBox
        Me.txtNombre = New System.Windows.Forms.TextBox
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.btnAbajo = New System.Windows.Forms.Button
        Me.btnArriba = New System.Windows.Forms.Button
        Me.btnLimpiar = New System.Windows.Forms.Button
        Me.btnAgregar = New System.Windows.Forms.Button
        Me.btnBuscar = New System.Windows.Forms.Button
        Me.txtBuscar = New System.Windows.Forms.TextBox
        Me.sprManufactura = New AxFPSpread.AxvaSpread
        Me.cboArticulo = New System.Windows.Forms.ComboBox
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.tsToolBar = New System.Windows.Forms.ToolStrip
        Me.tsbNuevo = New System.Windows.Forms.ToolStripButton
        Me.tsbActualizar = New System.Windows.Forms.ToolStripButton
        Me.tsbBorrar = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.tsbCancelar = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton
        Me.gpbArticulo.SuspendLayout()
        Me.tabPrincipal.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.sprProveedor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCotizacion.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.sprManufactura, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tsToolBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'gpbArticulo
        '
        Me.gpbArticulo.Controls.Add(Me.rdbBaja)
        Me.gpbArticulo.Controls.Add(Me.rdbSinCodigo)
        Me.gpbArticulo.Controls.Add(Me.rdbTodos)
        Me.gpbArticulo.Controls.Add(Me.lstRegistros)
        Me.gpbArticulo.Location = New System.Drawing.Point(12, 43)
        Me.gpbArticulo.Name = "gpbArticulo"
        Me.gpbArticulo.Size = New System.Drawing.Size(232, 489)
        Me.gpbArticulo.TabIndex = 4
        Me.gpbArticulo.TabStop = False
        Me.gpbArticulo.Text = "Listado de Artículos"
        '
        'rdbBaja
        '
        Me.rdbBaja.AutoSize = True
        Me.rdbBaja.Location = New System.Drawing.Point(13, 35)
        Me.rdbBaja.Name = "rdbBaja"
        Me.rdbBaja.Size = New System.Drawing.Size(46, 17)
        Me.rdbBaja.TabIndex = 4
        Me.rdbBaja.TabStop = True
        Me.rdbBaja.Text = "Baja"
        Me.rdbBaja.UseVisualStyleBackColor = True
        '
        'rdbSinCodigo
        '
        Me.rdbSinCodigo.AutoSize = True
        Me.rdbSinCodigo.Location = New System.Drawing.Point(86, 16)
        Me.rdbSinCodigo.Name = "rdbSinCodigo"
        Me.rdbSinCodigo.Size = New System.Drawing.Size(76, 17)
        Me.rdbSinCodigo.TabIndex = 3
        Me.rdbSinCodigo.TabStop = True
        Me.rdbSinCodigo.Text = "Sin Codigo"
        Me.rdbSinCodigo.UseVisualStyleBackColor = True
        '
        'rdbTodos
        '
        Me.rdbTodos.AutoSize = True
        Me.rdbTodos.Location = New System.Drawing.Point(13, 16)
        Me.rdbTodos.Name = "rdbTodos"
        Me.rdbTodos.Size = New System.Drawing.Size(55, 17)
        Me.rdbTodos.TabIndex = 2
        Me.rdbTodos.TabStop = True
        Me.rdbTodos.Text = "Todos"
        Me.rdbTodos.UseVisualStyleBackColor = True
        '
        'lstRegistros
        '
        Me.lstRegistros.FormattingEnabled = True
        Me.lstRegistros.Location = New System.Drawing.Point(6, 58)
        Me.lstRegistros.Name = "lstRegistros"
        Me.lstRegistros.Size = New System.Drawing.Size(216, 407)
        Me.lstRegistros.TabIndex = 1
        '
        'txtArticulo
        '
        Me.txtArticulo.Location = New System.Drawing.Point(774, 59)
        Me.txtArticulo.Name = "txtArticulo"
        Me.txtArticulo.Size = New System.Drawing.Size(82, 20)
        Me.txtArticulo.TabIndex = 7
        '
        'tabPrincipal
        '
        Me.tabPrincipal.Controls.Add(Me.TabPage1)
        Me.tabPrincipal.Controls.Add(Me.TabPage3)
        Me.tabPrincipal.Controls.Add(Me.TabPage4)
        Me.tabPrincipal.Controls.Add(Me.TabPage5)
        Me.tabPrincipal.Location = New System.Drawing.Point(250, 43)
        Me.tabPrincipal.Name = "tabPrincipal"
        Me.tabPrincipal.SelectedIndex = 0
        Me.tabPrincipal.Size = New System.Drawing.Size(630, 481)
        Me.tabPrincipal.TabIndex = 24
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.Transparent
        Me.TabPage1.Controls.Add(Me.txtFechaSuspension)
        Me.TabPage1.Controls.Add(Me.lblFactor)
        Me.TabPage1.Controls.Add(Me.txtFactor)
        Me.TabPage1.Controls.Add(Me.sprProveedor)
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Controls.Add(Me.grpCotizacion)
        Me.TabPage1.Controls.Add(Me.lblNombreCorto)
        Me.TabPage1.Controls.Add(Me.txtNombreCorto)
        Me.TabPage1.Controls.Add(Me.btnGeneraCodigo)
        Me.TabPage1.Controls.Add(Me.txtPrecioCompra)
        Me.TabPage1.Controls.Add(Me.txtFechaLista)
        Me.TabPage1.Controls.Add(Me.txtFechaCompra)
        Me.TabPage1.Controls.Add(Me.lblPrecioLista)
        Me.TabPage1.Controls.Add(Me.lblPrecioCompra)
        Me.TabPage1.Controls.Add(Me.txtPrecioLista)
        Me.TabPage1.Controls.Add(Me.lblMoneda)
        Me.TabPage1.Controls.Add(Me.cboMoneda)
        Me.TabPage1.Controls.Add(Me.lblCodigo)
        Me.TabPage1.Controls.Add(Me.txtCodigo)
        Me.TabPage1.Controls.Add(Me.lblCausaSuspension)
        Me.TabPage1.Controls.Add(Me.lblFechaSuspension)
        Me.TabPage1.Controls.Add(Me.txtCausaSuspension)
        Me.TabPage1.Controls.Add(Me.lblArticuloEstatus)
        Me.TabPage1.Controls.Add(Me.cboArticuloEstatus)
        Me.TabPage1.Controls.Add(Me.chkManufacturado)
        Me.TabPage1.Controls.Add(Me.chkRequiereArmado)
        Me.TabPage1.Controls.Add(Me.chkEsAlmacenable)
        Me.TabPage1.Controls.Add(Me.lblUnidadVenta)
        Me.TabPage1.Controls.Add(Me.lblUnidadCompra)
        Me.TabPage1.Controls.Add(Me.lblFamilia)
        Me.TabPage1.Controls.Add(Me.lblEtiqueta)
        Me.TabPage1.Controls.Add(Me.cboUnidadMedidaCompra)
        Me.TabPage1.Controls.Add(Me.cboUnidadMedidaInv)
        Me.TabPage1.Controls.Add(Me.cboFamilia)
        Me.TabPage1.Controls.Add(Me.txtNombre)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(622, 455)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "General"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'txtFechaSuspension
        '
        Me.txtFechaSuspension.Location = New System.Drawing.Point(128, 215)
        Me.txtFechaSuspension.Name = "txtFechaSuspension"
        Me.txtFechaSuspension.Size = New System.Drawing.Size(111, 20)
        Me.txtFechaSuspension.TabIndex = 56
        '
        'lblFactor
        '
        Me.lblFactor.AutoSize = True
        Me.lblFactor.Location = New System.Drawing.Point(150, 113)
        Me.lblFactor.Name = "lblFactor"
        Me.lblFactor.Size = New System.Drawing.Size(40, 13)
        Me.lblFactor.TabIndex = 61
        Me.lblFactor.Text = "Factor:"
        '
        'txtFactor
        '
        Me.txtFactor.Location = New System.Drawing.Point(141, 129)
        Me.txtFactor.Name = "txtFactor"
        Me.txtFactor.Size = New System.Drawing.Size(52, 20)
        Me.txtFactor.TabIndex = 8
        Me.txtFactor.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'sprProveedor
        '
        Me.sprProveedor.Location = New System.Drawing.Point(336, 83)
        Me.sprProveedor.Name = "sprProveedor"
        Me.sprProveedor.OcxState = CType(resources.GetObject("sprProveedor.OcxState"), System.Windows.Forms.AxHost.State)
        Me.sprProveedor.Size = New System.Drawing.Size(254, 144)
        Me.sprProveedor.TabIndex = 59
        '
        'GroupBox2
        '
        Me.GroupBox2.Location = New System.Drawing.Point(294, 316)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(316, 101)
        Me.GroupBox2.TabIndex = 58
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "GroupBox2"
        '
        'grpCotizacion
        '
        Me.grpCotizacion.Controls.Add(Me.lblKgxM2)
        Me.grpCotizacion.Controls.Add(Me.txtKGxM2)
        Me.grpCotizacion.Controls.Add(Me.lblTipoRecurso)
        Me.grpCotizacion.Controls.Add(Me.cboTipoRecurso)
        Me.grpCotizacion.Controls.Add(Me.lblUnidadCotiza)
        Me.grpCotizacion.Controls.Add(Me.cboUnidadMedidaCotizacion)
        Me.grpCotizacion.Location = New System.Drawing.Point(11, 316)
        Me.grpCotizacion.Name = "grpCotizacion"
        Me.grpCotizacion.Size = New System.Drawing.Size(249, 101)
        Me.grpCotizacion.TabIndex = 57
        Me.grpCotizacion.TabStop = False
        Me.grpCotizacion.Text = "Fuente para Cotizaciones"
        '
        'lblKgxM2
        '
        Me.lblKgxM2.AutoSize = True
        Me.lblKgxM2.Location = New System.Drawing.Point(54, 75)
        Me.lblKgxM2.Name = "lblKgxM2"
        Me.lblKgxM2.Size = New System.Drawing.Size(56, 13)
        Me.lblKgxM2.TabIndex = 57
        Me.lblKgxM2.Text = "Kg por M²:"
        '
        'txtKGxM2
        '
        Me.txtKGxM2.Location = New System.Drawing.Point(119, 75)
        Me.txtKGxM2.Name = "txtKGxM2"
        Me.txtKGxM2.Size = New System.Drawing.Size(71, 20)
        Me.txtKGxM2.TabIndex = 19
        Me.txtKGxM2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblTipoRecurso
        '
        Me.lblTipoRecurso.AutoSize = True
        Me.lblTipoRecurso.Location = New System.Drawing.Point(21, 46)
        Me.lblTipoRecurso.Name = "lblTipoRecurso"
        Me.lblTipoRecurso.Size = New System.Drawing.Size(89, 13)
        Me.lblTipoRecurso.TabIndex = 55
        Me.lblTipoRecurso.Text = "Tipo de Recurso:"
        '
        'cboTipoRecurso
        '
        Me.cboTipoRecurso.FormattingEnabled = True
        Me.cboTipoRecurso.Location = New System.Drawing.Point(119, 46)
        Me.cboTipoRecurso.Name = "cboTipoRecurso"
        Me.cboTipoRecurso.Size = New System.Drawing.Size(103, 21)
        Me.cboTipoRecurso.TabIndex = 18
        '
        'lblUnidadCotiza
        '
        Me.lblUnidadCotiza.AutoSize = True
        Me.lblUnidadCotiza.Location = New System.Drawing.Point(2, 21)
        Me.lblUnidadCotiza.Name = "lblUnidadCotiza"
        Me.lblUnidadCotiza.Size = New System.Drawing.Size(111, 13)
        Me.lblUnidadCotiza.TabIndex = 53
        Me.lblUnidadCotiza.Text = "Unidad de Cotización:"
        '
        'cboUnidadMedidaCotizacion
        '
        Me.cboUnidadMedidaCotizacion.FormattingEnabled = True
        Me.cboUnidadMedidaCotizacion.Location = New System.Drawing.Point(119, 18)
        Me.cboUnidadMedidaCotizacion.Name = "cboUnidadMedidaCotizacion"
        Me.cboUnidadMedidaCotizacion.Size = New System.Drawing.Size(103, 21)
        Me.cboUnidadMedidaCotizacion.TabIndex = 17
        '
        'lblNombreCorto
        '
        Me.lblNombreCorto.AutoSize = True
        Me.lblNombreCorto.Location = New System.Drawing.Point(8, 55)
        Me.lblNombreCorto.Name = "lblNombreCorto"
        Me.lblNombreCorto.Size = New System.Drawing.Size(66, 13)
        Me.lblNombreCorto.TabIndex = 56
        Me.lblNombreCorto.Text = "Desc. Corta:"
        '
        'txtNombreCorto
        '
        Me.txtNombreCorto.Location = New System.Drawing.Point(106, 52)
        Me.txtNombreCorto.Name = "txtNombreCorto"
        Me.txtNombreCorto.Size = New System.Drawing.Size(287, 20)
        Me.txtNombreCorto.TabIndex = 4
        '
        'btnGeneraCodigo
        '
        Me.btnGeneraCodigo.Location = New System.Drawing.Point(582, 6)
        Me.btnGeneraCodigo.Name = "btnGeneraCodigo"
        Me.btnGeneraCodigo.Size = New System.Drawing.Size(20, 19)
        Me.btnGeneraCodigo.TabIndex = 54
        Me.btnGeneraCodigo.UseVisualStyleBackColor = True
        '
        'txtPrecioCompra
        '
        Me.txtPrecioCompra.Location = New System.Drawing.Point(391, 265)
        Me.txtPrecioCompra.Name = "txtPrecioCompra"
        Me.txtPrecioCompra.Size = New System.Drawing.Size(71, 20)
        Me.txtPrecioCompra.TabIndex = 15
        Me.txtPrecioCompra.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtFechaLista
        '
        Me.txtFechaLista.Enabled = False
        Me.txtFechaLista.Location = New System.Drawing.Point(477, 293)
        Me.txtFechaLista.Name = "txtFechaLista"
        Me.txtFechaLista.Size = New System.Drawing.Size(94, 20)
        Me.txtFechaLista.TabIndex = 46
        Me.txtFechaLista.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtFechaCompra
        '
        Me.txtFechaCompra.Enabled = False
        Me.txtFechaCompra.Location = New System.Drawing.Point(477, 267)
        Me.txtFechaCompra.Name = "txtFechaCompra"
        Me.txtFechaCompra.Size = New System.Drawing.Size(94, 20)
        Me.txtFechaCompra.TabIndex = 45
        Me.txtFechaCompra.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblPrecioLista
        '
        Me.lblPrecioLista.AutoSize = True
        Me.lblPrecioLista.Location = New System.Drawing.Point(307, 296)
        Me.lblPrecioLista.Name = "lblPrecioLista"
        Me.lblPrecioLista.Size = New System.Drawing.Size(61, 13)
        Me.lblPrecioLista.TabIndex = 44
        Me.lblPrecioLista.Text = "Precio lista:"
        '
        'lblPrecioCompra
        '
        Me.lblPrecioCompra.AutoSize = True
        Me.lblPrecioCompra.Location = New System.Drawing.Point(307, 270)
        Me.lblPrecioCompra.Name = "lblPrecioCompra"
        Me.lblPrecioCompra.Size = New System.Drawing.Size(78, 13)
        Me.lblPrecioCompra.TabIndex = 43
        Me.lblPrecioCompra.Text = "Precio compra:"
        '
        'txtPrecioLista
        '
        Me.txtPrecioLista.Location = New System.Drawing.Point(391, 293)
        Me.txtPrecioLista.Name = "txtPrecioLista"
        Me.txtPrecioLista.Size = New System.Drawing.Size(71, 20)
        Me.txtPrecioLista.TabIndex = 16
        Me.txtPrecioLista.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblMoneda
        '
        Me.lblMoneda.AutoSize = True
        Me.lblMoneda.Location = New System.Drawing.Point(307, 242)
        Me.lblMoneda.Name = "lblMoneda"
        Me.lblMoneda.Size = New System.Drawing.Size(49, 13)
        Me.lblMoneda.TabIndex = 40
        Me.lblMoneda.Text = "Moneda:"
        '
        'cboMoneda
        '
        Me.cboMoneda.FormattingEnabled = True
        Me.cboMoneda.Location = New System.Drawing.Point(391, 238)
        Me.cboMoneda.Name = "cboMoneda"
        Me.cboMoneda.Size = New System.Drawing.Size(180, 21)
        Me.cboMoneda.TabIndex = 14
        '
        'lblCodigo
        '
        Me.lblCodigo.AutoSize = True
        Me.lblCodigo.Location = New System.Drawing.Point(449, 9)
        Me.lblCodigo.Name = "lblCodigo"
        Me.lblCodigo.Size = New System.Drawing.Size(43, 13)
        Me.lblCodigo.TabIndex = 37
        Me.lblCodigo.Text = "Codigo:"
        '
        'txtCodigo
        '
        Me.txtCodigo.Enabled = False
        Me.txtCodigo.Location = New System.Drawing.Point(493, 6)
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.Size = New System.Drawing.Size(83, 20)
        Me.txtCodigo.TabIndex = 36
        Me.txtCodigo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblCausaSuspension
        '
        Me.lblCausaSuspension.AutoSize = True
        Me.lblCausaSuspension.Location = New System.Drawing.Point(10, 238)
        Me.lblCausaSuspension.Name = "lblCausaSuspension"
        Me.lblCausaSuspension.Size = New System.Drawing.Size(122, 13)
        Me.lblCausaSuspension.TabIndex = 35
        Me.lblCausaSuspension.Text = "Causa de la suspensión:"
        '
        'lblFechaSuspension
        '
        Me.lblFechaSuspension.AutoSize = True
        Me.lblFechaSuspension.Location = New System.Drawing.Point(10, 218)
        Me.lblFechaSuspension.Name = "lblFechaSuspension"
        Me.lblFechaSuspension.Size = New System.Drawing.Size(113, 13)
        Me.lblFechaSuspension.TabIndex = 34
        Me.lblFechaSuspension.Text = "Fecha de Suspensión:"
        '
        'txtCausaSuspension
        '
        Me.txtCausaSuspension.Location = New System.Drawing.Point(11, 254)
        Me.txtCausaSuspension.Name = "txtCausaSuspension"
        Me.txtCausaSuspension.Size = New System.Drawing.Size(228, 20)
        Me.txtCausaSuspension.TabIndex = 13
        '
        'lblArticuloEstatus
        '
        Me.lblArticuloEstatus.AutoSize = True
        Me.lblArticuloEstatus.Location = New System.Drawing.Point(431, 55)
        Me.lblArticuloEstatus.Name = "lblArticuloEstatus"
        Me.lblArticuloEstatus.Size = New System.Drawing.Size(45, 13)
        Me.lblArticuloEstatus.TabIndex = 31
        Me.lblArticuloEstatus.Text = "Estatus:"
        '
        'cboArticuloEstatus
        '
        Me.cboArticuloEstatus.FormattingEnabled = True
        Me.cboArticuloEstatus.Location = New System.Drawing.Point(477, 52)
        Me.cboArticuloEstatus.Name = "cboArticuloEstatus"
        Me.cboArticuloEstatus.Size = New System.Drawing.Size(139, 21)
        Me.cboArticuloEstatus.TabIndex = 5
        '
        'chkManufacturado
        '
        Me.chkManufacturado.AutoSize = True
        Me.chkManufacturado.Location = New System.Drawing.Point(136, 191)
        Me.chkManufacturado.Name = "chkManufacturado"
        Me.chkManufacturado.Size = New System.Drawing.Size(132, 17)
        Me.chkManufacturado.TabIndex = 12
        Me.chkManufacturado.Text = "Requiere Manufactura"
        Me.chkManufacturado.UseVisualStyleBackColor = True
        '
        'chkRequiereArmado
        '
        Me.chkRequiereArmado.AutoSize = True
        Me.chkRequiereArmado.Location = New System.Drawing.Point(136, 168)
        Me.chkRequiereArmado.Name = "chkRequiereArmado"
        Me.chkRequiereArmado.Size = New System.Drawing.Size(108, 17)
        Me.chkRequiereArmado.TabIndex = 11
        Me.chkRequiereArmado.Text = "Requiere Armado"
        Me.chkRequiereArmado.UseVisualStyleBackColor = True
        '
        'chkEsAlmacenable
        '
        Me.chkEsAlmacenable.AutoSize = True
        Me.chkEsAlmacenable.Location = New System.Drawing.Point(13, 168)
        Me.chkEsAlmacenable.Name = "chkEsAlmacenable"
        Me.chkEsAlmacenable.Size = New System.Drawing.Size(87, 17)
        Me.chkEsAlmacenable.TabIndex = 10
        Me.chkEsAlmacenable.Text = "Almacenable"
        Me.chkEsAlmacenable.UseVisualStyleBackColor = True
        '
        'lblUnidadVenta
        '
        Me.lblUnidadVenta.AutoSize = True
        Me.lblUnidadVenta.Location = New System.Drawing.Point(196, 113)
        Me.lblUnidadVenta.Name = "lblUnidadVenta"
        Me.lblUnidadVenta.Size = New System.Drawing.Size(109, 13)
        Me.lblUnidadVenta.TabIndex = 26
        Me.lblUnidadVenta.Text = "Unidad de Inventario:"
        '
        'lblUnidadCompra
        '
        Me.lblUnidadCompra.AutoSize = True
        Me.lblUnidadCompra.Location = New System.Drawing.Point(10, 113)
        Me.lblUnidadCompra.Name = "lblUnidadCompra"
        Me.lblUnidadCompra.Size = New System.Drawing.Size(98, 13)
        Me.lblUnidadCompra.TabIndex = 25
        Me.lblUnidadCompra.Text = "Unidad de Compra:"
        '
        'lblFamilia
        '
        Me.lblFamilia.AutoSize = True
        Me.lblFamilia.Location = New System.Drawing.Point(8, 83)
        Me.lblFamilia.Name = "lblFamilia"
        Me.lblFamilia.Size = New System.Drawing.Size(97, 13)
        Me.lblFamilia.TabIndex = 24
        Me.lblFamilia.Text = "Familia del Articulo:"
        '
        'lblEtiqueta
        '
        Me.lblEtiqueta.AutoSize = True
        Me.lblEtiqueta.Location = New System.Drawing.Point(8, 9)
        Me.lblEtiqueta.Name = "lblEtiqueta"
        Me.lblEtiqueta.Size = New System.Drawing.Size(47, 13)
        Me.lblEtiqueta.TabIndex = 23
        Me.lblEtiqueta.Text = "Nombre:"
        '
        'cboUnidadMedidaCompra
        '
        Me.cboUnidadMedidaCompra.FormattingEnabled = True
        Me.cboUnidadMedidaCompra.Location = New System.Drawing.Point(11, 129)
        Me.cboUnidadMedidaCompra.Name = "cboUnidadMedidaCompra"
        Me.cboUnidadMedidaCompra.Size = New System.Drawing.Size(103, 21)
        Me.cboUnidadMedidaCompra.TabIndex = 7
        '
        'cboUnidadMedidaInv
        '
        Me.cboUnidadMedidaInv.FormattingEnabled = True
        Me.cboUnidadMedidaInv.Location = New System.Drawing.Point(199, 129)
        Me.cboUnidadMedidaInv.Name = "cboUnidadMedidaInv"
        Me.cboUnidadMedidaInv.Size = New System.Drawing.Size(103, 21)
        Me.cboUnidadMedidaInv.TabIndex = 9
        '
        'cboFamilia
        '
        Me.cboFamilia.FormattingEnabled = True
        Me.cboFamilia.Location = New System.Drawing.Point(106, 80)
        Me.cboFamilia.Name = "cboFamilia"
        Me.cboFamilia.Size = New System.Drawing.Size(180, 21)
        Me.cboFamilia.TabIndex = 6
        '
        'txtNombre
        '
        Me.txtNombre.Location = New System.Drawing.Point(106, 6)
        Me.txtNombre.Multiline = True
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.Size = New System.Drawing.Size(332, 40)
        Me.txtNombre.TabIndex = 3
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.btnAbajo)
        Me.TabPage3.Controls.Add(Me.btnArriba)
        Me.TabPage3.Controls.Add(Me.btnLimpiar)
        Me.TabPage3.Controls.Add(Me.btnAgregar)
        Me.TabPage3.Controls.Add(Me.btnBuscar)
        Me.TabPage3.Controls.Add(Me.txtBuscar)
        Me.TabPage3.Controls.Add(Me.sprManufactura)
        Me.TabPage3.Controls.Add(Me.cboArticulo)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(622, 455)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Arts. Base"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'btnAbajo
        '
        Me.btnAbajo.Image = Global.WindowsApplication1.My.Resources.Resources.down
        Me.btnAbajo.Location = New System.Drawing.Point(520, 222)
        Me.btnAbajo.Name = "btnAbajo"
        Me.btnAbajo.Size = New System.Drawing.Size(32, 32)
        Me.btnAbajo.TabIndex = 26
        Me.btnAbajo.UseVisualStyleBackColor = True
        '
        'btnArriba
        '
        Me.btnArriba.Image = Global.WindowsApplication1.My.Resources.Resources.up
        Me.btnArriba.Location = New System.Drawing.Point(520, 171)
        Me.btnArriba.Name = "btnArriba"
        Me.btnArriba.Size = New System.Drawing.Size(32, 32)
        Me.btnArriba.TabIndex = 25
        Me.btnArriba.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Image = Global.WindowsApplication1.My.Resources.Resources.delete
        Me.btnLimpiar.Location = New System.Drawing.Point(501, 13)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(32, 32)
        Me.btnLimpiar.TabIndex = 24
        Me.btnLimpiar.UseVisualStyleBackColor = True
        Me.btnLimpiar.Visible = False
        '
        'btnAgregar
        '
        Me.btnAgregar.Image = Global.WindowsApplication1.My.Resources.Resources.add
        Me.btnAgregar.Location = New System.Drawing.Point(463, 13)
        Me.btnAgregar.Name = "btnAgregar"
        Me.btnAgregar.Size = New System.Drawing.Size(32, 32)
        Me.btnAgregar.TabIndex = 23
        Me.btnAgregar.UseVisualStyleBackColor = True
        Me.btnAgregar.Visible = False
        '
        'btnBuscar
        '
        Me.btnBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnBuscar.Image = Global.WindowsApplication1.My.Resources.Resources.find
        Me.btnBuscar.Location = New System.Drawing.Point(425, 13)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(32, 32)
        Me.btnBuscar.TabIndex = 21
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'txtBuscar
        '
        Me.txtBuscar.Location = New System.Drawing.Point(87, 19)
        Me.txtBuscar.Name = "txtBuscar"
        Me.txtBuscar.Size = New System.Drawing.Size(332, 20)
        Me.txtBuscar.TabIndex = 20
        '
        'sprManufactura
        '
        Me.sprManufactura.Location = New System.Drawing.Point(74, 57)
        Me.sprManufactura.Name = "sprManufactura"
        Me.sprManufactura.OcxState = CType(resources.GetObject("sprManufactura.OcxState"), System.Windows.Forms.AxHost.State)
        Me.sprManufactura.Size = New System.Drawing.Size(429, 355)
        Me.sprManufactura.TabIndex = 0
        '
        'cboArticulo
        '
        Me.cboArticulo.FormattingEnabled = True
        Me.cboArticulo.Location = New System.Drawing.Point(87, 18)
        Me.cboArticulo.Name = "cboArticulo"
        Me.cboArticulo.Size = New System.Drawing.Size(358, 21)
        Me.cboArticulo.TabIndex = 22
        Me.cboArticulo.Visible = False
        '
        'TabPage4
        '
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(622, 455)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Componentes"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'TabPage5
        '
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Size = New System.Drawing.Size(622, 455)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Inventario"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'tsToolBar
        '
        Me.tsToolBar.AllowItemReorder = True
        Me.tsToolBar.Dock = System.Windows.Forms.DockStyle.None
        Me.tsToolBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbNuevo, Me.tsbActualizar, Me.tsbBorrar, Me.ToolStripSeparator1, Me.tsbCancelar, Me.ToolStripSeparator2, Me.ToolStripButton1})
        Me.tsToolBar.Location = New System.Drawing.Point(9, 9)
        Me.tsToolBar.Name = "tsToolBar"
        Me.tsToolBar.Size = New System.Drawing.Size(170, 25)
        Me.tsToolBar.Stretch = True
        Me.tsToolBar.TabIndex = 25
        Me.tsToolBar.Text = "toolStrip1"
        '
        'tsbNuevo
        '
        Me.tsbNuevo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbNuevo.Image = CType(resources.GetObject("tsbNuevo.Image"), System.Drawing.Image)
        Me.tsbNuevo.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbNuevo.Name = "tsbNuevo"
        Me.tsbNuevo.Size = New System.Drawing.Size(23, 22)
        Me.tsbNuevo.Text = "&Nuevo"
        '
        'tsbActualizar
        '
        Me.tsbActualizar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbActualizar.Image = CType(resources.GetObject("tsbActualizar.Image"), System.Drawing.Image)
        Me.tsbActualizar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbActualizar.Name = "tsbActualizar"
        Me.tsbActualizar.Size = New System.Drawing.Size(23, 22)
        Me.tsbActualizar.Text = "&Actualizar"
        '
        'tsbBorrar
        '
        Me.tsbBorrar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbBorrar.Image = CType(resources.GetObject("tsbBorrar.Image"), System.Drawing.Image)
        Me.tsbBorrar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbBorrar.Name = "tsbBorrar"
        Me.tsbBorrar.Size = New System.Drawing.Size(23, 22)
        Me.tsbBorrar.Text = "&Borrar"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'tsbCancelar
        '
        Me.tsbCancelar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbCancelar.Image = CType(resources.GetObject("tsbCancelar.Image"), System.Drawing.Image)
        Me.tsbCancelar.Name = "tsbCancelar"
        Me.tsbCancelar.Size = New System.Drawing.Size(23, 22)
        Me.tsbCancelar.Text = "&Cancelar"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "ToolStripButton1"
        '
        'frmArticulo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(892, 536)
        Me.Controls.Add(Me.tsToolBar)
        Me.Controls.Add(Me.tabPrincipal)
        Me.Controls.Add(Me.txtArticulo)
        Me.Controls.Add(Me.gpbArticulo)
        Me.Name = "frmArticulo"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Artículo"
        Me.gpbArticulo.ResumeLayout(False)
        Me.gpbArticulo.PerformLayout()
        Me.tabPrincipal.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.sprProveedor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCotizacion.ResumeLayout(False)
        Me.grpCotizacion.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.sprManufactura, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tsToolBar.ResumeLayout(False)
        Me.tsToolBar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gpbArticulo As System.Windows.Forms.GroupBox
    Friend WithEvents lstRegistros As System.Windows.Forms.ListBox
    Friend WithEvents txtArticulo As System.Windows.Forms.TextBox
    Friend WithEvents tabPrincipal As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents lblUnidadVenta As System.Windows.Forms.Label
    Friend WithEvents lblUnidadCompra As System.Windows.Forms.Label
    Friend WithEvents lblFamilia As System.Windows.Forms.Label
    Friend WithEvents lblEtiqueta As System.Windows.Forms.Label
    Friend WithEvents cboUnidadMedidaCompra As System.Windows.Forms.ComboBox
    Friend WithEvents cboUnidadMedidaInv As System.Windows.Forms.ComboBox
    Friend WithEvents cboFamilia As System.Windows.Forms.ComboBox
    Friend WithEvents txtNombre As System.Windows.Forms.TextBox
    Friend WithEvents chkEsAlmacenable As System.Windows.Forms.CheckBox
    Friend WithEvents lblArticuloEstatus As System.Windows.Forms.Label
    Friend WithEvents cboArticuloEstatus As System.Windows.Forms.ComboBox
    Friend WithEvents chkManufacturado As System.Windows.Forms.CheckBox
    Friend WithEvents chkRequiereArmado As System.Windows.Forms.CheckBox
    Friend WithEvents lblCausaSuspension As System.Windows.Forms.Label
    Friend WithEvents lblFechaSuspension As System.Windows.Forms.Label
    Friend WithEvents txtCausaSuspension As System.Windows.Forms.TextBox
    Friend WithEvents lblCodigo As System.Windows.Forms.Label
    Friend WithEvents txtCodigo As System.Windows.Forms.TextBox
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents lblMoneda As System.Windows.Forms.Label
    Friend WithEvents cboMoneda As System.Windows.Forms.ComboBox
    Friend WithEvents txtFechaLista As System.Windows.Forms.TextBox
    Friend WithEvents txtFechaCompra As System.Windows.Forms.TextBox
    Friend WithEvents lblPrecioLista As System.Windows.Forms.Label
    Friend WithEvents lblPrecioCompra As System.Windows.Forms.Label
    Friend WithEvents txtPrecioLista As System.Windows.Forms.TextBox
    Friend WithEvents tsToolBar As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbNuevo As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbActualizar As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents txtPrecioCompra As System.Windows.Forms.TextBox
    Friend WithEvents tsbCancelar As System.Windows.Forms.ToolStripButton
    Friend WithEvents sprManufactura As AxFPSpread.AxvaSpread
    Friend WithEvents btnBuscar As System.Windows.Forms.Button
    Friend WithEvents txtBuscar As System.Windows.Forms.TextBox
    Friend WithEvents btnAgregar As System.Windows.Forms.Button
    Friend WithEvents cboArticulo As System.Windows.Forms.ComboBox
    Friend WithEvents btnLimpiar As System.Windows.Forms.Button
    Friend WithEvents btnAbajo As System.Windows.Forms.Button
    Friend WithEvents btnArriba As System.Windows.Forms.Button
    Friend WithEvents btnGeneraCodigo As System.Windows.Forms.Button
    Friend WithEvents lblNombreCorto As System.Windows.Forms.Label
    Friend WithEvents txtNombreCorto As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents grpCotizacion As System.Windows.Forms.GroupBox
    Friend WithEvents lblUnidadCotiza As System.Windows.Forms.Label
    Friend WithEvents cboUnidadMedidaCotizacion As System.Windows.Forms.ComboBox
    Friend WithEvents lblFactor As System.Windows.Forms.Label
    Friend WithEvents txtFactor As System.Windows.Forms.TextBox
    Friend WithEvents sprProveedor As AxFPSpread.AxvaSpread
    Friend WithEvents lblKgxM2 As System.Windows.Forms.Label
    Friend WithEvents txtKGxM2 As System.Windows.Forms.TextBox
    Friend WithEvents lblTipoRecurso As System.Windows.Forms.Label
    Friend WithEvents cboTipoRecurso As System.Windows.Forms.ComboBox
    Friend WithEvents tsbBorrar As System.Windows.Forms.ToolStripButton
    Friend WithEvents txtFechaSuspension As System.Windows.Forms.TextBox
    Friend WithEvents rdbBaja As System.Windows.Forms.RadioButton
    Friend WithEvents rdbSinCodigo As System.Windows.Forms.RadioButton
    Friend WithEvents rdbTodos As System.Windows.Forms.RadioButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton

End Class

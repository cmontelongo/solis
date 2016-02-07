<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLogin
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmdAgregar As System.Windows.Forms.Button
	Public WithEvents cmdBuscarMecanico As System.Windows.Forms.Button
	Public WithEvents txtBuscar As System.Windows.Forms.TextBox
	Public WithEvents sprInsumos As vaSpread
	Public WithEvents txtDescripcion As System.Windows.Forms.TextBox
	Public WithEvents cboArticulo As System.Windows.Forms.ComboBox
	Public WithEvents cboArticulos As System.Windows.Forms.ComboBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLogin))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdAgregar = New System.Windows.Forms.Button
		Me.cmdBuscarMecanico = New System.Windows.Forms.Button
		Me.txtBuscar = New System.Windows.Forms.TextBox
		Me.sprInsumos = New vaSpread
		Me.txtDescripcion = New System.Windows.Forms.TextBox
		Me.cboArticulo = New System.Windows.Forms.ComboBox
		Me.cboArticulos = New System.Windows.Forms.ComboBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.sprInsumos, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Clave de Acceso"
		Me.ClientSize = New System.Drawing.Size(761, 443)
		Me.Location = New System.Drawing.Point(142, 74)
		Me.Icon = CType(resources.GetObject("frmLogin.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Tag = "1"
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmLogin"
		Me.cmdAgregar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdAgregar.Size = New System.Drawing.Size(33, 33)
		Me.cmdAgregar.Location = New System.Drawing.Point(592, 184)
		Me.cmdAgregar.Image = CType(resources.GetObject("cmdAgregar.Image"), System.Drawing.Image)
		Me.cmdAgregar.TabIndex = 5
		Me.cmdAgregar.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAgregar.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAgregar.CausesValidation = True
		Me.cmdAgregar.Enabled = True
		Me.cmdAgregar.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAgregar.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAgregar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAgregar.TabStop = True
		Me.cmdAgregar.Name = "cmdAgregar"
		Me.cmdBuscarMecanico.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdBuscarMecanico.Size = New System.Drawing.Size(21, 21)
		Me.cmdBuscarMecanico.Location = New System.Drawing.Point(560, 192)
		Me.cmdBuscarMecanico.Image = CType(resources.GetObject("cmdBuscarMecanico.Image"), System.Drawing.Image)
		Me.cmdBuscarMecanico.TabIndex = 4
		Me.cmdBuscarMecanico.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBuscarMecanico.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBuscarMecanico.CausesValidation = True
		Me.cmdBuscarMecanico.Enabled = True
		Me.cmdBuscarMecanico.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBuscarMecanico.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBuscarMecanico.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBuscarMecanico.TabStop = True
		Me.cmdBuscarMecanico.Name = "cmdBuscarMecanico"
		Me.txtBuscar.AutoSize = False
		Me.txtBuscar.Size = New System.Drawing.Size(529, 19)
		Me.txtBuscar.Location = New System.Drawing.Point(24, 192)
		Me.txtBuscar.TabIndex = 3
		Me.txtBuscar.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtBuscar.AcceptsReturn = True
		Me.txtBuscar.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtBuscar.BackColor = System.Drawing.SystemColors.Window
		Me.txtBuscar.CausesValidation = True
		Me.txtBuscar.Enabled = True
		Me.txtBuscar.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtBuscar.HideSelection = True
		Me.txtBuscar.ReadOnly = False
		Me.txtBuscar.Maxlength = 0
		Me.txtBuscar.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBuscar.MultiLine = False
		Me.txtBuscar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBuscar.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtBuscar.TabStop = True
		Me.txtBuscar.Visible = True
		Me.txtBuscar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtBuscar.Name = "txtBuscar"
		sprInsumos.OcxState = CType(resources.GetObject("sprInsumos.OcxState"), System.Windows.Forms.AxHost.State)
		Me.sprInsumos.Size = New System.Drawing.Size(705, 145)
		Me.sprInsumos.Location = New System.Drawing.Point(24, 224)
		Me.sprInsumos.TabIndex = 2
		Me.sprInsumos.Name = "sprInsumos"
		Me.txtDescripcion.AutoSize = False
		Me.txtDescripcion.Size = New System.Drawing.Size(713, 105)
		Me.txtDescripcion.Location = New System.Drawing.Point(24, 72)
		Me.txtDescripcion.MultiLine = True
		Me.txtDescripcion.TabIndex = 1
		Me.txtDescripcion.Text = "Text1"
		Me.txtDescripcion.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDescripcion.AcceptsReturn = True
		Me.txtDescripcion.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Window
		Me.txtDescripcion.CausesValidation = True
		Me.txtDescripcion.Enabled = True
		Me.txtDescripcion.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDescripcion.HideSelection = True
		Me.txtDescripcion.ReadOnly = False
		Me.txtDescripcion.Maxlength = 0
		Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDescripcion.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDescripcion.TabStop = True
		Me.txtDescripcion.Visible = True
		Me.txtDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDescripcion.Name = "txtDescripcion"
		Me.cboArticulo.Size = New System.Drawing.Size(713, 21)
		Me.cboArticulo.Location = New System.Drawing.Point(24, 32)
		Me.cboArticulo.TabIndex = 0
		Me.cboArticulo.Text = "Combo1"
		Me.cboArticulo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboArticulo.BackColor = System.Drawing.SystemColors.Window
		Me.cboArticulo.CausesValidation = True
		Me.cboArticulo.Enabled = True
		Me.cboArticulo.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboArticulo.IntegralHeight = True
		Me.cboArticulo.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboArticulo.Sorted = False
		Me.cboArticulo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboArticulo.TabStop = True
		Me.cboArticulo.Visible = True
		Me.cboArticulo.Name = "cboArticulo"
		Me.cboArticulos.Size = New System.Drawing.Size(561, 21)
		Me.cboArticulos.Location = New System.Drawing.Point(24, 192)
		Me.cboArticulos.TabIndex = 6
		Me.cboArticulos.Text = "Combo1"
		Me.cboArticulos.Visible = False
		Me.cboArticulos.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboArticulos.BackColor = System.Drawing.SystemColors.Window
		Me.cboArticulos.CausesValidation = True
		Me.cboArticulos.Enabled = True
		Me.cboArticulos.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboArticulos.IntegralHeight = True
		Me.cboArticulos.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboArticulos.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboArticulos.Sorted = False
		Me.cboArticulos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboArticulos.TabStop = True
		Me.cboArticulos.Name = "cboArticulos"
		Me.Controls.Add(cmdAgregar)
		Me.Controls.Add(cmdBuscarMecanico)
		Me.Controls.Add(txtBuscar)
		Me.Controls.Add(sprInsumos)
		Me.Controls.Add(txtDescripcion)
		Me.Controls.Add(cboArticulo)
		Me.Controls.Add(cboArticulos)
		CType(Me.sprInsumos, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class
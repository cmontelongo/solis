<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRelaciona
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRelaciona))
        Me.sprCoincidencia = New AxFPSpread.AxvaSpread
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.LEMISORNOMBRE = New System.Windows.Forms.Label
        Me.LEMISORRFC = New System.Windows.Forms.Label
        Me.Label36 = New System.Windows.Forms.Label
        Me.Label37 = New System.Windows.Forms.Label
        Me.btnRelaciona = New System.Windows.Forms.Button
        Me.btnAlta = New System.Windows.Forms.Button
        Me.fraElementos = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnBuscar = New System.Windows.Forms.Button
        Me.txtBuscar = New System.Windows.Forms.TextBox
        Me.sprRelaciona = New AxFPSpread.AxvaSpread
        CType(Me.sprCoincidencia, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.fraElementos.SuspendLayout()
        CType(Me.sprRelaciona, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'sprCoincidencia
        '
        Me.sprCoincidencia.Location = New System.Drawing.Point(34, 80)
        Me.sprCoincidencia.Name = "sprCoincidencia"
        Me.sprCoincidencia.OcxState = CType(resources.GetObject("sprCoincidencia.OcxState"), System.Windows.Forms.AxHost.State)
        Me.sprCoincidencia.Size = New System.Drawing.Size(327, 346)
        Me.sprCoincidencia.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.LEMISORNOMBRE)
        Me.GroupBox2.Controls.Add(Me.LEMISORRFC)
        Me.GroupBox2.Controls.Add(Me.Label36)
        Me.GroupBox2.Controls.Add(Me.Label37)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 7)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(447, 69)
        Me.GroupBox2.TabIndex = 26
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Información original"
        '
        'LEMISORNOMBRE
        '
        Me.LEMISORNOMBRE.AutoSize = True
        Me.LEMISORNOMBRE.ForeColor = System.Drawing.Color.Maroon
        Me.LEMISORNOMBRE.Location = New System.Drawing.Point(52, 37)
        Me.LEMISORNOMBRE.Name = "LEMISORNOMBRE"
        Me.LEMISORNOMBRE.Size = New System.Drawing.Size(44, 13)
        Me.LEMISORNOMBRE.TabIndex = 8
        Me.LEMISORNOMBRE.Text = "Nombre"
        '
        'LEMISORRFC
        '
        Me.LEMISORRFC.AutoSize = True
        Me.LEMISORRFC.ForeColor = System.Drawing.Color.Maroon
        Me.LEMISORRFC.Location = New System.Drawing.Point(52, 16)
        Me.LEMISORRFC.Name = "LEMISORRFC"
        Me.LEMISORRFC.Size = New System.Drawing.Size(28, 13)
        Me.LEMISORRFC.TabIndex = 7
        Me.LEMISORRFC.Text = "RFC"
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(5, 37)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(47, 13)
        Me.Label36.TabIndex = 5
        Me.Label36.Text = "Nombre:"
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Location = New System.Drawing.Point(21, 16)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(31, 13)
        Me.Label37.TabIndex = 6
        Me.Label37.Text = "RFC:"
        '
        'btnRelaciona
        '
        Me.btnRelaciona.Enabled = False
        Me.btnRelaciona.Location = New System.Drawing.Point(480, 25)
        Me.btnRelaciona.Name = "btnRelaciona"
        Me.btnRelaciona.Size = New System.Drawing.Size(115, 36)
        Me.btnRelaciona.TabIndex = 28
        Me.btnRelaciona.Text = "Relacionar"
        Me.btnRelaciona.UseVisualStyleBackColor = True
        '
        'btnAlta
        '
        Me.btnAlta.Location = New System.Drawing.Point(601, 25)
        Me.btnAlta.Name = "btnAlta"
        Me.btnAlta.Size = New System.Drawing.Size(115, 36)
        Me.btnAlta.TabIndex = 29
        Me.btnAlta.Text = "Alta Nueva"
        Me.btnAlta.UseVisualStyleBackColor = True
        '
        'fraElementos
        '
        Me.fraElementos.Controls.Add(Me.Label1)
        Me.fraElementos.Controls.Add(Me.btnBuscar)
        Me.fraElementos.Controls.Add(Me.txtBuscar)
        Me.fraElementos.Location = New System.Drawing.Point(12, 82)
        Me.fraElementos.Name = "fraElementos"
        Me.fraElementos.Size = New System.Drawing.Size(726, 364)
        Me.fraElementos.TabIndex = 33
        Me.fraElementos.TabStop = False
        Me.fraElementos.Text = "Registros probables para relacionar:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 13)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "Buscar"
        '
        'btnBuscar
        '
        Me.btnBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnBuscar.Image = Global.WindowsApplication1.My.Resources.Resources.find
        Me.btnBuscar.Location = New System.Drawing.Point(469, 13)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(32, 32)
        Me.btnBuscar.TabIndex = 34
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'txtBuscar
        '
        Me.txtBuscar.Location = New System.Drawing.Point(57, 19)
        Me.txtBuscar.Name = "txtBuscar"
        Me.txtBuscar.Size = New System.Drawing.Size(406, 20)
        Me.txtBuscar.TabIndex = 33
        '
        'sprRelaciona
        '
        Me.sprRelaciona.Location = New System.Drawing.Point(18, 131)
        Me.sprRelaciona.Name = "sprRelaciona"
        Me.sprRelaciona.OcxState = CType(resources.GetObject("sprRelaciona.OcxState"), System.Windows.Forms.AxHost.State)
        Me.sprRelaciona.Size = New System.Drawing.Size(714, 305)
        Me.sprRelaciona.TabIndex = 27
        '
        'frmRelaciona
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(748, 454)
        Me.Controls.Add(Me.sprRelaciona)
        Me.Controls.Add(Me.fraElementos)
        Me.Controls.Add(Me.btnAlta)
        Me.Controls.Add(Me.btnRelaciona)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "frmRelaciona"
        Me.Text = "Interfaz de información"
        CType(Me.sprCoincidencia, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.fraElementos.ResumeLayout(False)
        Me.fraElementos.PerformLayout()
        CType(Me.sprRelaciona, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents sprCoincidencia As AxFPSpread.AxvaSpread
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents LEMISORNOMBRE As System.Windows.Forms.Label
    Friend WithEvents LEMISORRFC As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents btnRelaciona As System.Windows.Forms.Button
    Friend WithEvents btnAlta As System.Windows.Forms.Button
    Friend WithEvents fraElementos As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnBuscar As System.Windows.Forms.Button
    Friend WithEvents txtBuscar As System.Windows.Forms.TextBox
    Friend WithEvents sprRelaciona As AxFPSpread.AxvaSpread
End Class

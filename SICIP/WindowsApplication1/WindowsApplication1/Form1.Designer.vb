<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.LRECEPTORNOMBRE = New System.Windows.Forms.Label
        Me.LRECEPTORRFC = New System.Windows.Forms.Label
        Me.Label40 = New System.Windows.Forms.Label
        Me.Label41 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.LEMISORNOMBRE = New System.Windows.Forms.Label
        Me.LEMISORRFC = New System.Windows.Forms.Label
        Me.Label36 = New System.Windows.Forms.Label
        Me.Label37 = New System.Windows.Forms.Label
        Me.LTOTAL = New System.Windows.Forms.Label
        Me.Label33 = New System.Windows.Forms.Label
        Me.LCOMPROBANTE = New System.Windows.Forms.Label
        Me.Label31 = New System.Windows.Forms.Label
        Me.LSUBTOTAL = New System.Windows.Forms.Label
        Me.Label29 = New System.Windows.Forms.Label
        Me.LAPROBACION = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.LCERTIFICADO = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.LMOTIVO = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.LMETODOPAGO = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.LFORMAPAGO = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.LCONDICIONES = New System.Windows.Forms.Label
        Me.LDESCUENTO = New System.Windows.Forms.Label
        Me.LANIOAPROB = New System.Windows.Forms.Label
        Me.LIMPTRASLADADOS = New System.Windows.Forms.Label
        Me.LIMPRETENIDOS = New System.Windows.Forms.Label
        Me.LSERIE = New System.Windows.Forms.Label
        Me.LFOLIO = New System.Windows.Forms.Label
        Me.LFECHA = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.lblArchivo = New System.Windows.Forms.Label
        Me.prgProceso = New System.Windows.Forms.ProgressBar
        Me.optSeleccionMasiva = New System.Windows.Forms.RadioButton
        Me.optSeleccion = New System.Windows.Forms.RadioButton
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lstArchivos = New System.Windows.Forms.ListBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.LTOTAL)
        Me.GroupBox1.Controls.Add(Me.Label33)
        Me.GroupBox1.Controls.Add(Me.LCOMPROBANTE)
        Me.GroupBox1.Controls.Add(Me.Label31)
        Me.GroupBox1.Controls.Add(Me.LSUBTOTAL)
        Me.GroupBox1.Controls.Add(Me.Label29)
        Me.GroupBox1.Controls.Add(Me.LAPROBACION)
        Me.GroupBox1.Controls.Add(Me.Label27)
        Me.GroupBox1.Controls.Add(Me.LCERTIFICADO)
        Me.GroupBox1.Controls.Add(Me.Label25)
        Me.GroupBox1.Controls.Add(Me.LMOTIVO)
        Me.GroupBox1.Controls.Add(Me.Label23)
        Me.GroupBox1.Controls.Add(Me.LMETODOPAGO)
        Me.GroupBox1.Controls.Add(Me.Label21)
        Me.GroupBox1.Controls.Add(Me.LFORMAPAGO)
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.LCONDICIONES)
        Me.GroupBox1.Controls.Add(Me.LDESCUENTO)
        Me.GroupBox1.Controls.Add(Me.LANIOAPROB)
        Me.GroupBox1.Controls.Add(Me.LIMPTRASLADADOS)
        Me.GroupBox1.Controls.Add(Me.LIMPRETENIDOS)
        Me.GroupBox1.Controls.Add(Me.LSERIE)
        Me.GroupBox1.Controls.Add(Me.LFOLIO)
        Me.GroupBox1.Controls.Add(Me.LFECHA)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(17, 84)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(597, 358)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "CFD"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.LRECEPTORNOMBRE)
        Me.GroupBox3.Controls.Add(Me.LRECEPTORRFC)
        Me.GroupBox3.Controls.Add(Me.Label40)
        Me.GroupBox3.Controls.Add(Me.Label41)
        Me.GroupBox3.Location = New System.Drawing.Point(328, 19)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(262, 54)
        Me.GroupBox3.TabIndex = 25
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Receptor"
        '
        'LRECEPTORNOMBRE
        '
        Me.LRECEPTORNOMBRE.AutoSize = True
        Me.LRECEPTORNOMBRE.ForeColor = System.Drawing.Color.Maroon
        Me.LRECEPTORNOMBRE.Location = New System.Drawing.Point(51, 35)
        Me.LRECEPTORNOMBRE.Name = "LRECEPTORNOMBRE"
        Me.LRECEPTORNOMBRE.Size = New System.Drawing.Size(44, 13)
        Me.LRECEPTORNOMBRE.TabIndex = 12
        Me.LRECEPTORNOMBRE.Text = "Nombre"
        '
        'LRECEPTORRFC
        '
        Me.LRECEPTORRFC.AutoSize = True
        Me.LRECEPTORRFC.ForeColor = System.Drawing.Color.Maroon
        Me.LRECEPTORRFC.Location = New System.Drawing.Point(51, 14)
        Me.LRECEPTORRFC.Name = "LRECEPTORRFC"
        Me.LRECEPTORRFC.Size = New System.Drawing.Size(28, 13)
        Me.LRECEPTORRFC.TabIndex = 11
        Me.LRECEPTORRFC.Text = "RFC"
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Location = New System.Drawing.Point(4, 35)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(47, 13)
        Me.Label40.TabIndex = 9
        Me.Label40.Text = "Nombre:"
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(20, 14)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(31, 13)
        Me.Label41.TabIndex = 10
        Me.Label41.Text = "RFC:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.LEMISORNOMBRE)
        Me.GroupBox2.Controls.Add(Me.LEMISORRFC)
        Me.GroupBox2.Controls.Add(Me.Label36)
        Me.GroupBox2.Controls.Add(Me.Label37)
        Me.GroupBox2.Location = New System.Drawing.Point(13, 19)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(309, 54)
        Me.GroupBox2.TabIndex = 25
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Emisor"
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
        'LTOTAL
        '
        Me.LTOTAL.AutoSize = True
        Me.LTOTAL.ForeColor = System.Drawing.Color.Maroon
        Me.LTOTAL.Location = New System.Drawing.Point(135, 338)
        Me.LTOTAL.Name = "LTOTAL"
        Me.LTOTAL.Size = New System.Drawing.Size(31, 13)
        Me.LTOTAL.TabIndex = 24
        Me.LTOTAL.Text = "Total"
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(95, 338)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(34, 13)
        Me.Label33.TabIndex = 23
        Me.Label33.Text = "Total:"
        '
        'LCOMPROBANTE
        '
        Me.LCOMPROBANTE.AutoSize = True
        Me.LCOMPROBANTE.ForeColor = System.Drawing.Color.Maroon
        Me.LCOMPROBANTE.Location = New System.Drawing.Point(135, 192)
        Me.LCOMPROBANTE.Name = "LCOMPROBANTE"
        Me.LCOMPROBANTE.Size = New System.Drawing.Size(109, 13)
        Me.LCOMPROBANTE.TabIndex = 22
        Me.LCOMPROBANTE.Text = "Tipo de Comprobante"
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(17, 192)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(112, 13)
        Me.Label31.TabIndex = 21
        Me.Label31.Text = "Tipo de Comprobante:"
        '
        'LSUBTOTAL
        '
        Me.LSUBTOTAL.AutoSize = True
        Me.LSUBTOTAL.ForeColor = System.Drawing.Color.Maroon
        Me.LSUBTOTAL.Location = New System.Drawing.Point(135, 317)
        Me.LSUBTOTAL.Name = "LSUBTOTAL"
        Me.LSUBTOTAL.Size = New System.Drawing.Size(50, 13)
        Me.LSUBTOTAL.TabIndex = 20
        Me.LSUBTOTAL.Text = "SubTotal"
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(76, 317)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(53, 13)
        Me.Label29.TabIndex = 19
        Me.Label29.Text = "SubTotal:"
        '
        'LAPROBACION
        '
        Me.LAPROBACION.AutoSize = True
        Me.LAPROBACION.ForeColor = System.Drawing.Color.Maroon
        Me.LAPROBACION.Location = New System.Drawing.Point(135, 150)
        Me.LAPROBACION.Name = "LAPROBACION"
        Me.LAPROBACION.Size = New System.Drawing.Size(76, 13)
        Me.LAPROBACION.TabIndex = 18
        Me.LAPROBACION.Text = "N° Aprobación"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(50, 150)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(79, 13)
        Me.Label27.TabIndex = 17
        Me.Label27.Text = "N° Aprobación:"
        '
        'LCERTIFICADO
        '
        Me.LCERTIFICADO.AutoSize = True
        Me.LCERTIFICADO.ForeColor = System.Drawing.Color.Maroon
        Me.LCERTIFICADO.Location = New System.Drawing.Point(135, 129)
        Me.LCERTIFICADO.Name = "LCERTIFICADO"
        Me.LCERTIFICADO.Size = New System.Drawing.Size(72, 13)
        Me.LCERTIFICADO.TabIndex = 16
        Me.LCERTIFICADO.Text = "N° Certificado"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(54, 129)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(75, 13)
        Me.Label25.TabIndex = 15
        Me.Label25.Text = "N° Certificado:"
        '
        'LMOTIVO
        '
        Me.LMOTIVO.AutoSize = True
        Me.LMOTIVO.ForeColor = System.Drawing.Color.Maroon
        Me.LMOTIVO.Location = New System.Drawing.Point(433, 254)
        Me.LMOTIVO.Name = "LMOTIVO"
        Me.LMOTIVO.Size = New System.Drawing.Size(97, 13)
        Me.LMOTIVO.TabIndex = 14
        Me.LMOTIVO.Text = "Motivo  Descuento"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(327, 254)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(100, 13)
        Me.Label23.TabIndex = 13
        Me.Label23.Text = "Motivo  Descuento:"
        '
        'LMETODOPAGO
        '
        Me.LMETODOPAGO.AutoSize = True
        Me.LMETODOPAGO.ForeColor = System.Drawing.Color.Maroon
        Me.LMETODOPAGO.Location = New System.Drawing.Point(135, 254)
        Me.LMETODOPAGO.Name = "LMETODOPAGO"
        Me.LMETODOPAGO.Size = New System.Drawing.Size(86, 13)
        Me.LMETODOPAGO.TabIndex = 12
        Me.LMETODOPAGO.Text = "Metodo de Pago"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(40, 254)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(89, 13)
        Me.Label21.TabIndex = 11
        Me.Label21.Text = "Metodo de Pago:"
        '
        'LFORMAPAGO
        '
        Me.LFORMAPAGO.AutoSize = True
        Me.LFORMAPAGO.ForeColor = System.Drawing.Color.Maroon
        Me.LFORMAPAGO.Location = New System.Drawing.Point(135, 233)
        Me.LFORMAPAGO.Name = "LFORMAPAGO"
        Me.LFORMAPAGO.Size = New System.Drawing.Size(79, 13)
        Me.LFORMAPAGO.TabIndex = 10
        Me.LFORMAPAGO.Text = "Forma de Pago"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(47, 233)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(82, 13)
        Me.Label19.TabIndex = 9
        Me.Label19.Text = "Forma de Pago:"
        '
        'LCONDICIONES
        '
        Me.LCONDICIONES.AutoSize = True
        Me.LCONDICIONES.ForeColor = System.Drawing.Color.Maroon
        Me.LCONDICIONES.Location = New System.Drawing.Point(135, 213)
        Me.LCONDICIONES.Name = "LCONDICIONES"
        Me.LCONDICIONES.Size = New System.Drawing.Size(108, 13)
        Me.LCONDICIONES.TabIndex = 6
        Me.LCONDICIONES.Text = "Condiciones de Pago"
        '
        'LDESCUENTO
        '
        Me.LDESCUENTO.AutoSize = True
        Me.LDESCUENTO.ForeColor = System.Drawing.Color.Maroon
        Me.LDESCUENTO.Location = New System.Drawing.Point(433, 233)
        Me.LDESCUENTO.Name = "LDESCUENTO"
        Me.LDESCUENTO.Size = New System.Drawing.Size(59, 13)
        Me.LDESCUENTO.TabIndex = 5
        Me.LDESCUENTO.Text = "Descuento"
        '
        'LANIOAPROB
        '
        Me.LANIOAPROB.AutoSize = True
        Me.LANIOAPROB.ForeColor = System.Drawing.Color.Maroon
        Me.LANIOAPROB.Location = New System.Drawing.Point(135, 171)
        Me.LANIOAPROB.Name = "LANIOAPROB"
        Me.LANIOAPROB.Size = New System.Drawing.Size(98, 13)
        Me.LANIOAPROB.TabIndex = 8
        Me.LANIOAPROB.Text = "Año de Aprobación"
        '
        'LIMPTRASLADADOS
        '
        Me.LIMPTRASLADADOS.AutoSize = True
        Me.LIMPTRASLADADOS.ForeColor = System.Drawing.Color.Maroon
        Me.LIMPTRASLADADOS.Location = New System.Drawing.Point(135, 296)
        Me.LIMPTRASLADADOS.Name = "LIMPTRASLADADOS"
        Me.LIMPTRASLADADOS.Size = New System.Drawing.Size(116, 13)
        Me.LIMPTRASLADADOS.TabIndex = 7
        Me.LIMPTRASLADADOS.Text = "Impuestos Trasladados"
        '
        'LIMPRETENIDOS
        '
        Me.LIMPRETENIDOS.AutoSize = True
        Me.LIMPRETENIDOS.ForeColor = System.Drawing.Color.Maroon
        Me.LIMPRETENIDOS.Location = New System.Drawing.Point(135, 275)
        Me.LIMPRETENIDOS.Name = "LIMPRETENIDOS"
        Me.LIMPRETENIDOS.Size = New System.Drawing.Size(106, 13)
        Me.LIMPRETENIDOS.TabIndex = 2
        Me.LIMPRETENIDOS.Text = "Impuestos Retenidos"
        '
        'LSERIE
        '
        Me.LSERIE.AutoSize = True
        Me.LSERIE.ForeColor = System.Drawing.Color.Maroon
        Me.LSERIE.Location = New System.Drawing.Point(233, 107)
        Me.LSERIE.Name = "LSERIE"
        Me.LSERIE.Size = New System.Drawing.Size(31, 13)
        Me.LSERIE.TabIndex = 1
        Me.LSERIE.Text = "Serie"
        '
        'LFOLIO
        '
        Me.LFOLIO.AutoSize = True
        Me.LFOLIO.ForeColor = System.Drawing.Color.Maroon
        Me.LFOLIO.Location = New System.Drawing.Point(136, 107)
        Me.LFOLIO.Name = "LFOLIO"
        Me.LFOLIO.Size = New System.Drawing.Size(29, 13)
        Me.LFOLIO.TabIndex = 4
        Me.LFOLIO.Text = "Folio"
        '
        'LFECHA
        '
        Me.LFECHA.AutoSize = True
        Me.LFECHA.ForeColor = System.Drawing.Color.Maroon
        Me.LFECHA.Location = New System.Drawing.Point(136, 89)
        Me.LFECHA.Name = "LFECHA"
        Me.LFECHA.Size = New System.Drawing.Size(37, 13)
        Me.LFECHA.TabIndex = 3
        Me.LFECHA.Text = "Fecha"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(18, 213)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(111, 13)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Condiciones de Pago:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(365, 233)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(62, 13)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Descuento:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(28, 171)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(101, 13)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Año de Aprobación:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(10, 296)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(119, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Impuestos Trasladados:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(20, 275)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(109, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Impuestos Retenidos:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(98, 107)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Folio:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(90, 89)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Fecha:"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.lblArchivo)
        Me.GroupBox4.Controls.Add(Me.prgProceso)
        Me.GroupBox4.Controls.Add(Me.optSeleccionMasiva)
        Me.GroupBox4.Controls.Add(Me.optSeleccion)
        Me.GroupBox4.Controls.Add(Me.Button1)
        Me.GroupBox4.Controls.Add(Me.TextBox1)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Location = New System.Drawing.Point(17, 3)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(598, 75)
        Me.GroupBox4.TabIndex = 4
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "GroupBox4"
        '
        'lblArchivo
        '
        Me.lblArchivo.AutoSize = True
        Me.lblArchivo.Location = New System.Drawing.Point(242, 33)
        Me.lblArchivo.Name = "lblArchivo"
        Me.lblArchivo.Size = New System.Drawing.Size(39, 13)
        Me.lblArchivo.TabIndex = 9
        Me.lblArchivo.Text = "Label4"
        '
        'prgProceso
        '
        Me.prgProceso.Location = New System.Drawing.Point(239, 16)
        Me.prgProceso.Name = "prgProceso"
        Me.prgProceso.Size = New System.Drawing.Size(239, 13)
        Me.prgProceso.TabIndex = 8
        '
        'optSeleccionMasiva
        '
        Me.optSeleccionMasiva.AutoSize = True
        Me.optSeleccionMasiva.Location = New System.Drawing.Point(126, 21)
        Me.optSeleccionMasiva.Name = "optSeleccionMasiva"
        Me.optSeleccionMasiva.Size = New System.Drawing.Size(90, 17)
        Me.optSeleccionMasiva.TabIndex = 7
        Me.optSeleccionMasiva.TabStop = True
        Me.optSeleccionMasiva.Text = "Carga Masiva"
        Me.optSeleccionMasiva.UseVisualStyleBackColor = True
        '
        'optSeleccion
        '
        Me.optSeleccion.AutoSize = True
        Me.optSeleccion.Checked = True
        Me.optSeleccion.Location = New System.Drawing.Point(19, 21)
        Me.optSeleccion.Name = "optSeleccion"
        Me.optSeleccion.Size = New System.Drawing.Size(91, 17)
        Me.optSeleccion.TabIndex = 6
        Me.optSeleccion.TabStop = True
        Me.optSeleccion.Text = "Carga Manual"
        Me.optSeleccion.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(445, 50)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(145, 23)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Selecciona XML"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(55, 52)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(384, 20)
        Me.TextBox1.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 55)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "XML:"
        '
        'lstArchivos
        '
        Me.lstArchivos.FormattingEnabled = True
        Me.lstArchivos.Location = New System.Drawing.Point(620, 112)
        Me.lstArchivos.Name = "lstArchivos"
        Me.lstArchivos.Size = New System.Drawing.Size(48, 264)
        Me.lstArchivos.TabIndex = 5
        Me.lstArchivos.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(627, 454)
        Me.Controls.Add(Me.lstArchivos)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form1"
        Me.Text = " CFD Extraer Informacion XML"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents LCERTIFICADO As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents LMOTIVO As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents LMETODOPAGO As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents LFORMAPAGO As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents LCONDICIONES As System.Windows.Forms.Label
    Friend WithEvents LDESCUENTO As System.Windows.Forms.Label
    Friend WithEvents LANIOAPROB As System.Windows.Forms.Label
    Friend WithEvents LIMPTRASLADADOS As System.Windows.Forms.Label
    Friend WithEvents LIMPRETENIDOS As System.Windows.Forms.Label
    Friend WithEvents LSERIE As System.Windows.Forms.Label
    Friend WithEvents LFOLIO As System.Windows.Forms.Label
    Friend WithEvents LFECHA As System.Windows.Forms.Label
    Friend WithEvents LTOTAL As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents LCOMPROBANTE As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents LSUBTOTAL As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents LAPROBACION As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents LEMISORNOMBRE As System.Windows.Forms.Label
    Friend WithEvents LEMISORRFC As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents LRECEPTORNOMBRE As System.Windows.Forms.Label
    Friend WithEvents LRECEPTORRFC As System.Windows.Forms.Label
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents optSeleccionMasiva As System.Windows.Forms.RadioButton
    Friend WithEvents optSeleccion As System.Windows.Forms.RadioButton
    Friend WithEvents lblArchivo As System.Windows.Forms.Label
    Friend WithEvents prgProceso As System.Windows.Forms.ProgressBar
    Friend WithEvents lstArchivos As System.Windows.Forms.ListBox

End Class

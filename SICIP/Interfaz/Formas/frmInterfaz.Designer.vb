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
        Me.btnExaminar = New System.Windows.Forms.Button
        Me.btnAbrir = New System.Windows.Forms.Button
        Me.txtFic = New System.Windows.Forms.TextBox
        Me.txtSelect = New System.Windows.Forms.TextBox
        Me.dgvDiarios = New System.Windows.Forms.DataGridView
        CType(Me.dgvDiarios, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExaminar
        '
        Me.btnExaminar.Location = New System.Drawing.Point(371, 54)
        Me.btnExaminar.Name = "btnExaminar"
        Me.btnExaminar.Size = New System.Drawing.Size(92, 34)
        Me.btnExaminar.TabIndex = 0
        Me.btnExaminar.Text = "Examinar"
        Me.btnExaminar.UseVisualStyleBackColor = True
        '
        'btnAbrir
        '
        Me.btnAbrir.Location = New System.Drawing.Point(371, 94)
        Me.btnAbrir.Name = "btnAbrir"
        Me.btnAbrir.Size = New System.Drawing.Size(92, 32)
        Me.btnAbrir.TabIndex = 1
        Me.btnAbrir.Text = "Button2"
        Me.btnAbrir.UseVisualStyleBackColor = True
        '
        'txtFic
        '
        Me.txtFic.Location = New System.Drawing.Point(67, 61)
        Me.txtFic.Name = "txtFic"
        Me.txtFic.Size = New System.Drawing.Size(116, 20)
        Me.txtFic.TabIndex = 2
        '
        'txtSelect
        '
        Me.txtSelect.Location = New System.Drawing.Point(67, 110)
        Me.txtSelect.Name = "txtSelect"
        Me.txtSelect.Size = New System.Drawing.Size(116, 20)
        Me.txtSelect.TabIndex = 3
        '
        'dgvDiarios
        '
        Me.dgvDiarios.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDiarios.Location = New System.Drawing.Point(12, 168)
        Me.dgvDiarios.Name = "dgvDiarios"
        Me.dgvDiarios.Size = New System.Drawing.Size(483, 239)
        Me.dgvDiarios.TabIndex = 4
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(507, 431)
        Me.Controls.Add(Me.dgvDiarios)
        Me.Controls.Add(Me.txtSelect)
        Me.Controls.Add(Me.txtFic)
        Me.Controls.Add(Me.btnAbrir)
        Me.Controls.Add(Me.btnExaminar)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.dgvDiarios, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnExaminar As System.Windows.Forms.Button
    Friend WithEvents btnAbrir As System.Windows.Forms.Button
    Friend WithEvents txtFic As System.Windows.Forms.TextBox
    Friend WithEvents txtSelect As System.Windows.Forms.TextBox
    Friend WithEvents dgvDiarios As System.Windows.Forms.DataGridView

End Class

Imports System
Imports System.Windows.Forms
Imports System.Data

Public Class Form1

    Dim mrs As New ADODB.Recordset

    Private Sub btnExaminar_Click(ByVal sender As Object, _
                                  ByVal e As EventArgs) _
                                  Handles btnExaminar.Click
        Dim oFD As New OpenFileDialog
        With oFD
            .Filter = "Ficheros DBF (*.dbf)|*.dbf|Todos (*.*)|*.*"
            .FileName = txtFic.Text
            If .ShowDialog = DialogResult.OK Then
                txtFic.Text = .FileName
                ' El nombre del fichero
                txtSelect.Text = System.IO.Path.GetFileNameWithoutExtension(txtFic.Text)
                btnAbrir_Click(Nothing, Nothing)
            End If
        End With
    End Sub

    Private Sub btnAbrir_Click(ByVal sender As Object, _
                               ByVal e As EventArgs) _
                               Handles btnAbrir.Click
        Dim sBase As String = txtFic.Text
        Dim sSelect As String = "SELECT * FROM " & txtSelect.Text
        Dim sConn As String

        sConn = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=" & _
                System.IO.Path.GetDirectoryName(sBase) & ";"

        Using dbConn As New System.Data.Odbc.OdbcConnection(sConn)
            Try
                dbConn.Open()

                Dim da As New System.Data.Odbc.OdbcDataAdapter(sSelect, dbConn)
                Dim dt As New DataTable

                da.Fill(dt)

                dgvDiarios.DataSource = dt

                dbConn.Close()

            Catch ex As Exception
                MessageBox.Show("Error al abrir la base de datos" & vbCrLf & ex.Message)
                Exit Sub
            End Try
        End Using

    End Sub

    Private Sub Form1_DragDrop(ByVal sender As Object, _
                               ByVal e As DragEventArgs) _
                               Handles Me.DragDrop, txtFic.DragDrop
        ' Drag & Drop, aceptar el primer fichero
        If e.Data.GetDataPresent("FileDrop") Then
            txtFic.Text = CType(e.Data.GetData("FileDrop", True), String())(0)
            txtSelect.Text = System.IO.Path.GetFileNameWithoutExtension(txtFic.Text)
        End If
    End Sub

    Private Sub Form1_DragEnter(ByVal sender As Object, _
                                ByVal e As DragEventArgs) _
                                Handles Me.DragEnter, txtFic.DragEnter
        ' Drag & Drop, comprobar con DataFormats
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub


    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sBase As String
        Dim sSelect As String
        Dim sConn As String
        Dim rsConsulta As New ADODB.Recordset
        Dim strSQL As String

        AbreConeccion()

        mrs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        mrs.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        mrs.LockType = ADODB.LockTypeEnum.adLockOptimistic


        rsConsulta.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsConsulta.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
        rsConsulta.LockType = ADODB.LockTypeEnum.adLockOptimistic

        txtFic.Text = "E:\BS\bdtoriz\articulo.DBF"
        sBase = txtFic.Text

        sSelect = "SELECT * FROM " & System.IO.Path.GetFileNameWithoutExtension(txtFic.Text)

        sConn = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=" & _
                System.IO.Path.GetDirectoryName(sBase) & ";"

        Using dbConn As New System.Data.Odbc.OdbcConnection(sConn)
            Try

                Dim da As New System.Data.Odbc.OdbcCommand(sSelect, dbConn)

                dbConn.Open()

                gcn.Execute("DELETE dbo.ExtBSArticulo")

                Dim reader As System.Data.Odbc.OdbcDataReader = da.ExecuteReader()

                strSQL = "SELECT * FROM ExtBSArticulo"
                rsConsulta.Open(strSQL, gcn)


                While reader.Read()
                    rsConsulta.AddNew()
                    rsConsulta.Fields(0).Value = reader(0).ToString   'ART_CLAVE	varchar(16) NOT NULL
                    rsConsulta.Fields(1).Value = reader(1).ToString   ',ART_DESC   varchar(35) NULL
                    rsConsulta.Fields(2).Value = reader(2).ToString   ',ART_DESC2  varchar(60) NULL
                    rsConsulta.Fields(3).Value = reader(3).ToString   ',ART_BARCOD varchar(16) NULL 
                    rsConsulta.Fields(4).Value = reader(4).ToString   ',ART_REFER  varchar(16) NULL
                    rsConsulta.Fields(5).Value = CDbl(reader(5).ToString)   ',ART_COSTO  REAL NULL
                    rsConsulta.Fields(6).Value = CDbl(reader(6).ToString)  ',ART_COSREP REAL NULL
                    rsConsulta.Fields(7).Value = reader(7).ToString   ',ART_UNDINV  varchar(5) NULL
                    rsConsulta.Fields(8).Value = reader(8).ToString   ',ART_TIPO   varchar(1) NULL
                    rsConsulta.Fields(9).Value = CDbl(reader(9).ToString)   ',ART_DIAS REAL NULL
                    rsConsulta.Fields(10).Value = reader(10).ToString   ',ART_ENSAMB  varchar(2) NULL
                    rsConsulta.Fields(11).Value = reader(11).ToString   ',ART_FAMILI  varchar(15) NULL
                    rsConsulta.Fields(12).Value = reader(12).ToString   ',ART_SUBFAM  varchar(16) NULL
                    rsConsulta.Fields(13).Value = reader(13).ToString   ',ART_MARCA  varchar(13) NULL
                    rsConsulta.Fields(14).Value = reader(14).ToString   ',ART_ALMVTA  varchar(5) NULL
                    rsConsulta.Fields(15).Value = reader(15).ToString   ',ART_ALMVTP  varchar(5) NULL
                    rsConsulta.Fields(16).Value = reader(16).ToString   ',ART_RETCVE  varchar(16) NULL
                    rsConsulta.Fields(17).Value = CDbl(reader(17).ToString)   ',ART_RETVAL REAL NULL
                    rsConsulta.Fields(18).Value = CDbl(reader(18).ToString)   ',ART_RETUND REAL NULL
                    rsConsulta.Fields(19).Value = reader(19).ToString   ',ART_LISPRE  varchar(28) NULL
                    rsConsulta.Fields(20).Value = reader(20).ToString   ',ART_ALMUSA  varchar(5) NULL
                    rsConsulta.Fields(21).Value = reader(21).ToString   ',ART_IVA  varchar(5) NULL
                    rsConsulta.Fields(22).Value = CDbl(reader(22).ToString)   ',ART_PRECIO REAL NULL
                    rsConsulta.Fields(23).Value = CDbl(reader(23).ToString)  ',ART_PRECIP REAL NULL
                    rsConsulta.Fields(24).Value = CDbl(reader(24).ToString)  ',ART_CARGO REAL NULL
                    rsConsulta.Fields(25).Value = CDbl(reader(25).ToString)  ',ART_FACVTA REAL NULL
                    rsConsulta.Update()
                End While

                reader.Close()
                rsConsulta.Close()
                dbConn.Close()

            Catch ex As Exception
                MessageBox.Show("Error al abrir la base de datos" & vbCrLf & ex.Message)
                Exit Sub
            End Try
        End Using

    End Sub
End Class

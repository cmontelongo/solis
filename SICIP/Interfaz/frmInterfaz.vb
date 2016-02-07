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

                ActualizaProveedor()

                dbConn.Close()

            Catch ex As Exception
                MessageBox.Show("Error al abrir la base de datos" & vbCrLf & ex.Message)
                Exit Sub
            End Try
        End Using

    End Sub
    Private Sub ActualizaProveedor()
        Dim sBase As String
        Dim sSelect As String
        Dim sConn As String
        Dim rsConsulta As New ADODB.Recordset
        Dim strSQL As String

        mrs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        mrs.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        mrs.LockType = ADODB.LockTypeEnum.adLockOptimistic

        rsConsulta.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsConsulta.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
        rsConsulta.LockType = ADODB.LockTypeEnum.adLockOptimistic

        txtFic.Text = "E:\BS\bdtoriz\proveed.DBF"
        sBase = txtFic.Text

        sSelect = "SELECT * FROM " & System.IO.Path.GetFileNameWithoutExtension(txtFic.Text)

        sConn = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=" & _
                System.IO.Path.GetDirectoryName(sBase) & ";"

        Using dbConn As New System.Data.Odbc.OdbcConnection(sConn)
            Try

                Dim da As New System.Data.Odbc.OdbcCommand(sSelect, dbConn)

                dbConn.Open()

                gcn.Execute("DELETE dbo.ExtBSProveed")

                Dim reader As System.Data.Odbc.OdbcDataReader = da.ExecuteReader()

                strSQL = "SELECT * FROM ExtBSProveed"
                rsConsulta.Open(strSQL, gcn)
                While reader.Read()
                    rsConsulta.AddNew()
                    rsConsulta.Fields(0).Value = reader(0).ToString   'PRV_CODIGO VARCHAR(10) NOT NULL,
                    rsConsulta.Fields(1).Value = reader(1).ToString   'PRV_NOMBRE VARCHAR(60) null,
                    rsConsulta.Fields(2).Value = reader(2).ToString   'PRV_RFC VARCHAR(15) null,
                    rsConsulta.Fields(3).Value = reader(3).ToString   'PRV_DIRECC VARCHAR(40) null,
                    rsConsulta.Fields(4).Value = reader(4).ToString   'PRV_COLON VARCHAR(40) null,
                    rsConsulta.Fields(5).Value = reader(5).ToString   'PRV_CIUDAD VARCHAR(5) null,
                    rsConsulta.Fields(6).Value = reader(6).ToString  'PRV_ESTADO VARCHAR(5) null,
                    rsConsulta.Fields(7).Value = reader(7).ToString   'PRV_TIPO VARCHAR(10) null,
                    rsConsulta.Fields(8).Value = reader(8).ToString   'PRV_PAIS VARCHAR(1) null,
                    rsConsulta.Fields(9).Value = CDbl(reader(9).ToString)   'PRV_CODPOS real null,
                    rsConsulta.Fields(10).Value = reader(10).ToString   'PRV_LADA VARCHAR(3) null,
                    rsConsulta.Fields(11).Value = reader(11).ToString   'PRV_TELEF1 VARCHAR(8) null,
                    rsConsulta.Fields(12).Value = reader(12).ToString   'PRV_TELEF2 VARCHAR(8) null,
                    rsConsulta.Fields(13).Value = reader(13).ToString   'PRV_FAX VARCHAR(8) null,
                    rsConsulta.Fields(14).Value = CDbl(reader(14).ToString)  'PRV_PLAZO real null,
                    rsConsulta.Fields(15).Value = CDbl(reader(15).ToString)  'PRV_LIMITE real null,
                    rsConsulta.Fields(16).Value = reader(16).ToString   'PRV_CNOM1 VARCHAR(20) null,
                    rsConsulta.Fields(17).Value = reader(17).ToString   'PRV_CDEPT1 VARCHAR(20) null,
                    rsConsulta.Fields(18).Value = reader(18).ToString   'PRV_TEL1 VARCHAR(10) null,
                    rsConsulta.Fields(19).Value = reader(19).ToString   'PRV_CNOM2 VARCHAR(20) null,
                    rsConsulta.Fields(20).Value = reader(20).ToString   'PRV_CDEPT2 VARCHAR(20) null,
                    rsConsulta.Fields(21).Value = reader(21).ToString   'PRV_TEL2 VARCHAR(10) null,
                    rsConsulta.Fields(22).Value = reader(22).ToString   'PRV_BLQCOM VARCHAR(1) null,
                    rsConsulta.Fields(23).Value = reader(23).ToString  'PRV_SOLAUT VARCHAR(1) null,
                    rsConsulta.Fields(24).Value = reader(24).ToString  'PRV_CTACON VARCHAR(12) null,
                    rsConsulta.Fields(25).Value = reader(25).ToString  'PRV_CTACSG VARCHAR(12) null,
                    rsConsulta.Fields(26).Value = reader(26).ToString   'PRV_SALDO VARCHAR(12) null,
                    rsConsulta.Fields(27).Value = reader(27).ToString   'PRV_SALDOL VARCHAR(12) null,
                    rsConsulta.Fields(28).Value = CDbl(reader(28).ToString)   'PRV_IMPTO real null,
                    rsConsulta.Fields(29).Value = reader(29).ToString   'PRV_VIA VARCHAR(20) null,
                    rsConsulta.Fields(30).Value = reader(30).ToString   'PRV_PLAZOL VARCHAR(5) null,
                    rsConsulta.Fields(31).Value = CDbl(reader(31).ToString)  'PRV_CASCO real null,
                    rsConsulta.Fields(32).Value = reader(32).ToString   'PRV_MONEDA VARCHAR(10) null,
                    rsConsulta.Fields(33).Value = reader(33).ToString   'PRV_CTAGTO VARCHAR(12) null)
                    rsConsulta.Update()

                End While

                reader.Close()
                rsConsulta.Close()

            Catch ex As Exception
                MessageBox.Show("Error al abrir la base de datos" & vbCrLf & ex.Message)
                Exit Sub
            End Try
        End Using
    End Sub
End Class

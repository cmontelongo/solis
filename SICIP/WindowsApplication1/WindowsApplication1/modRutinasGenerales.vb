Module modRutinasGenerales

    Public Class clsSelector
        Private strCveElemento As String
        Private strTextoElemento As String

        Public Sub New(ByVal strClave As String, ByVal strTexto As String)
            strCveElemento = strClave
            strTextoElemento = strTexto
        End Sub

        Public ReadOnly Property strClave() As String
            Get
                Return strCveElemento
            End Get
        End Property

        Public ReadOnly Property strTexto() As String
            Get
                Return strTextoElemento
            End Get
        End Property
    End Class

    Public Sub LlenaSelector(ByRef objSelector As Object, ByVal strSQL As String)

        Dim rsfLlenaControl As New ADODB.Recordset()
        Dim Lista As New ArrayList

        rsfLlenaControl.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsfLlenaControl.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        rsfLlenaControl.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic

        rsfLlenaControl.Open(strSQL, gcn)

        Try

            Do Until rsfLlenaControl.EOF
                Lista.Add(New clsSelector(rsfLlenaControl.Fields.Item(0).Value, rsfLlenaControl.Fields.Item(1).Value))
                rsfLlenaControl.MoveNext()
            Loop

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Metodo CargarComboDesdeSql", MessageBoxButtons.OK)
        Finally
            rsfLlenaControl.Close()
        End Try


        With objSelector
            .DataSource = Lista
            .ValueMember = "strClave"
            .DisplayMember = "strTexto"
        End With
    End Sub

    Public Sub ToolBotones_Estado(ByVal objToolBar As System.Windows.Forms.ToolStrip, ByVal vblnEstado As Boolean)
        '*******************************************************************
        ' Descripción : Rutina para poner todos los  botones de un ToolBar en
        '               el estado que se requiera
        ' Entrada :
        '       ObjToolBar .- Nombre del Toolbar
        '       bEstado .- Estado deseado
        '*******************************************************************
        Dim i As Byte
        For i = 1 To objToolBar.Items.Count
            objToolBar.Items(i - 1).Enabled = vblnEstado
        Next i

    End Sub
    Public Sub ToolBoton_Estado(ByVal objToolBar As System.Windows.Forms.ToolStrip, ByVal vstrKey As String, ByVal vblnEstado As Boolean)
        '*******************************************************************
        ' Descripción : Rutina para poner un  botón de un ToolBar en
        '               el estado que se requiera
        ' Entrada :
        '       objToolBar .- Nombre del ToolBar
        '       vstrKey .- Botón deseado
        '       vblnEstado .- Estado que se le desea poner
        '*******************************************************************
        Dim i As Integer
        Dim blnEncontrado As Boolean

        For i = 1 To objToolBar.Items.Count
            If objToolBar.Items(i - 1).Name = vstrKey Then
                blnEncontrado = True
                Exit For
            End If
        Next i

        If blnEncontrado Then
            objToolBar.Items(i - 1).Enabled = vblnEstado
        End If

    End Sub
End Module

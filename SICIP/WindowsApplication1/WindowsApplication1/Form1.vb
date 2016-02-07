Option Strict Off

Imports System
Imports System.IO

Public Class Form1

    Dim mrs As New ADODB.Recordset
    Dim frmNew As New frmRelaciona()

    Private Structure TInfoCFD

        Dim Sello As String
        Dim Serie As String
        Dim Folio As Long
        Dim Fecha As DateTime
        Dim RFCEmisor As String
        '--------------------------
        Dim noAprobacion As Integer
        Dim anoAprobacion As Integer
        Dim formaDePago As String
        Dim noCertificado As String
        Dim condicionesDePago As String
        Dim subTotal As Double
        Dim descuento As Double
        Dim motivoDescuento As String
        Dim total As Double
        Dim metodoDePago As String
        Dim tipoDeComprobante As String
        Dim Emisor_RFC As String
        Dim Emisor_Nombre As String
        Dim Receptor_RFC As String
        Dim Receptor_Nombre As String
        Dim totalImpuestosTrasladados As Decimal
        Dim totalImpuestosRetenidos As Decimal
    End Structure

    Private Sub ExtraerInfoCFD(ByVal vFile As String)
        Dim xmlDoc As New System.Xml.XmlDocument
        Dim Comprobante As System.Xml.XmlNode
        Dim Impuestos As System.Xml.XmlNode
        Dim nsp As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
        Dim vsTIR As Double = 0
        Dim vsTIT As Double = 0
        Dim pID As String = ""
        Dim pName As String = ""
        Dim pPrice As String = ""
        Dim strSQL As String
        Dim strHora As String
        Dim lngRenglon As Long
        Dim strCodigoProducto As String
        Dim lngCveArticulo As Long
        Dim curValorUnitario As Decimal
        Dim curImporte As Decimal
        Dim intCveMoneda As Integer
        Dim Info As TInfoCFD
        Dim rsConsulta As New ADODB.Recordset
        Dim lngCveProveedor As Long


        rsConsulta.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsConsulta.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        rsConsulta.LockType = ADODB.LockTypeEnum.adLockOptimistic

        Try
            xmlDoc.Load(vFile)
            nsp.AddNamespace("cfdi", "http://www.sat.gob.mx/cfd/3")

            Comprobante = xmlDoc.ChildNodes(1)

            If Comprobante Is Nothing Then
                MsgBox("No se puede evaluar XML " & vFile, MsgBoxStyle.Critical, "Extraer Info XML")
                Exit Sub
            End If

            With Info
                .Fecha = Comprobante.Attributes.GetNamedItem("fecha").Value

                If Not Comprobante.Attributes.GetNamedItem("folio") Is Nothing Then
                    .Folio = Comprobante.Attributes.GetNamedItem("folio").Value
                Else
                    .Folio = 0
                End If
                .Sello = Comprobante.Attributes.GetNamedItem("sello").Value
                If Not Comprobante.Attributes.GetNamedItem("serie") Is Nothing Then
                    .Serie = Comprobante.Attributes.GetNamedItem("serie").Value
                Else
                    .Serie = ""
                End If

                If Not Comprobante("cfdi:Emisor").Attributes.GetNamedItem("rfc") Is Nothing Then
                    .RFCEmisor = Comprobante("cfdi:Emisor").Attributes.GetNamedItem("rfc").Value
                Else
                    .RFCEmisor = ""
                End If
                If Not Comprobante("cfdi:Emisor").Attributes.GetNamedItem("rfc") Is Nothing Then
                    .Emisor_RFC = Comprobante("cfdi:Emisor").Attributes.GetNamedItem("rfc").Value
                Else
                    .Emisor_RFC = ""
                End If
                If Not Comprobante("cfdi:Emisor").Attributes.GetNamedItem("nombre") Is Nothing Then
                    .Emisor_Nombre = Comprobante("cfdi:Emisor").Attributes.GetNamedItem("nombre").Value
                Else
                    .Emisor_Nombre = ""
                End If
                If Not Comprobante("cfdi:Receptor").Attributes.GetNamedItem("rfc") Is Nothing Then
                    .Receptor_RFC = Comprobante("cfdi:Receptor").Attributes.GetNamedItem("rfc").Value
                Else
                    .Receptor_RFC = ""
                End If
                If Not Comprobante("cfdi:Receptor").Attributes.GetNamedItem("nombre") Is Nothing Then
                    .Receptor_Nombre = Comprobante("cfdi:Receptor").Attributes.GetNamedItem("nombre").Value
                Else
                    .Receptor_Nombre = ""
                End If
                If Not Comprobante("cfdi:Impuestos").Attributes.GetNamedItem("totalImpuestosRetenidos") Is Nothing Then
                    .totalImpuestosRetenidos = Comprobante("cfdi:Impuestos").Attributes.GetNamedItem("totalImpuestosRetenidos").Value
                Else
                    If Not Comprobante("cfdi:Impuestos") Is Nothing Then
                        Impuestos = Comprobante("cfdi:Impuestos")
                        For Each vRegIR As System.Xml.XmlElement In Impuestos.ChildNodes
                            If vRegIR.Name = "cfdi:Retenciones" Then vsTIR = vsTIR + vRegIR.FirstChild.Attributes("importe").Value
                        Next
                        .totalImpuestosRetenidos = vsTIR
                    Else
                        .totalImpuestosRetenidos = 0
                    End If
                End If

                If Not Comprobante("cfdi:Impuestos").Attributes.GetNamedItem("totalImpuestosTrasladados") Is Nothing Then
                    .totalImpuestosTrasladados = Comprobante("cfdi:Impuestos").Attributes.GetNamedItem("totalImpuestosTrasladados").Value
                Else
                    If Not Comprobante("cfdi:Impuestos") Is Nothing Then
                        Impuestos = Comprobante("cfdi:Impuestos")
                        For Each vRegIT As System.Xml.XmlElement In Impuestos.ChildNodes
                            If vRegIT.Name = "Traslados" Then vsTIT = vsTIT + vRegIT.FirstChild.Attributes("importe").Value
                        Next
                        .totalImpuestosTrasladados = vsTIT
                    Else
                        .totalImpuestosTrasladados = 0
                    End If
                End If

                If Not Comprobante.Attributes.GetNamedItem("anoAprobacion") Is Nothing Then
                    .anoAprobacion = Comprobante.Attributes.GetNamedItem("anoAprobacion").Value
                Else
                    .anoAprobacion = 0
                End If
                If Not Comprobante.Attributes.GetNamedItem("condicionesDePago") Is Nothing Then
                    .condicionesDePago = Comprobante.Attributes.GetNamedItem("condicionesDePago").Value
                Else
                    .condicionesDePago = ""
                End If
                If Not Comprobante.Attributes.GetNamedItem("descuento") Is Nothing Then
                    .descuento = Comprobante.Attributes.GetNamedItem("descuento").Value
                Else
                    .descuento = 0
                End If
                If Not Comprobante.Attributes.GetNamedItem("formaDePago") Is Nothing Then
                    .formaDePago = Comprobante.Attributes.GetNamedItem("formaDePago").Value
                Else
                    .formaDePago = ""
                End If
                If Not Comprobante.Attributes.GetNamedItem("metodoDePago") Is Nothing Then
                    .metodoDePago = Comprobante.Attributes.GetNamedItem("metodoDePago").Value
                Else
                    .metodoDePago = ""
                End If
                If Not Comprobante.Attributes.GetNamedItem("motivoDescuento") Is Nothing Then
                    .motivoDescuento = Comprobante.Attributes.GetNamedItem("motivoDescuento").Value
                Else
                    .motivoDescuento = ""
                End If
                If Not Comprobante.Attributes.GetNamedItem("noCertificado") Is Nothing Then
                    .noCertificado = Comprobante.Attributes.GetNamedItem("noCertificado").Value
                Else
                    .noCertificado = ""
                End If
                If Not Comprobante.Attributes.GetNamedItem("noAprobacion") Is Nothing Then
                    .noAprobacion = Comprobante.Attributes.GetNamedItem("noAprobacion").Value
                Else
                    .noAprobacion = 0
                End If
                If Not Comprobante.Attributes.GetNamedItem("subTotal") Is Nothing Then
                    .subTotal = Comprobante.Attributes.GetNamedItem("subTotal").Value
                Else
                    .subTotal = 0
                End If
                If Not Comprobante.Attributes.GetNamedItem("tipoDeComprobante") Is Nothing Then
                    .tipoDeComprobante = Comprobante.Attributes.GetNamedItem("tipoDeComprobante").Value
                Else
                    .tipoDeComprobante = ""
                End If
                If Not Comprobante.Attributes.GetNamedItem("total") Is Nothing Then
                    .total = Comprobante.Attributes.GetNamedItem("total").Value
                Else
                    .total = 0
                End If


            End With

            If Not Comprobante.Attributes.GetNamedItem("Moneda") Is Nothing Then
                strSQL = "SELECT CveMoneda FROM Moneda where Diversidad like '%-" & Trim(Comprobante.Attributes.GetNamedItem("Moneda").Value) & "-%'"
                rsConsulta.Open(strSQL, gcn)
                If rsConsulta.EOF Then

                    If Not Comprobante.Attributes.GetNamedItem("TipoCambio") Is Nothing Then
                        If Val(Comprobante.Attributes.GetNamedItem("TipoCambio").Value) <> 1 Then
                            MsgBox("aguas con la moneda=" & Comprobante.Attributes.GetNamedItem("Moneda").Value)
                            Exit Sub
                        Else
                            MsgBox("aguas con la moneda=" & Comprobante.Attributes.GetNamedItem("Moneda").Value)
                            MsgBox("aguas con la moneda")
                            intCveMoneda = 1
                        End If
                    Else
                        MsgBox("aguas con la moneda=" & Comprobante.Attributes.GetNamedItem("Moneda").Value)
                        Exit Sub
                    End If
                Else
                    intCveMoneda = rsConsulta.Fields("CveMoneda").Value
                End If
                rsConsulta.Close()
            Else
                intCveMoneda = 1
            End If

            modABC.pintCveMoneda = intCveMoneda


            strSQL = "SELECT * FROM Proveedor where REPLACE(REPLACE(RFC,' ',''),'-','') ='" & Info.RFCEmisor & "'"
            rsConsulta.Open(strSQL, gcn)
            If rsConsulta.EOF Then
                MsgBox("El RFC del Proveedor No existe")
                modABC.plngCveProveedor = 0
                modABC.pblnProveedor = True
                modABC.pstrNombre = Info.Emisor_Nombre
                modABC.pstrFacturaXML = vFile
                modABC.pstrRFC = Info.Emisor_RFC

                frmNew.ShowDialog()

                lngCveProveedor = modABC.plngCveProveedor

                If modABC.plngCveProveedor = 0 Then Exit Sub
            Else
                lngCveProveedor = rsConsulta.Fields("CveProveedor").Value
            End If
            rsConsulta.Close()

            strSQL = "SELECT * FROM ProveedorFactura where CveProveedor = " & lngCveProveedor & _
                " and Folio = " & Info.Folio & " AND Serie = '" & Info.Serie & "'"
            rsConsulta.Open(strSQL, gcn)
            If rsConsulta.EOF Then
                '
                If Mid(Info.Fecha, 21, 1) = "a" Then
                    If Mid(Info.Fecha, 12, 2) = "12" Then
                        strHora = "00" & Mid(Info.Fecha, 14, 6)
                    Else
                        strHora = Mid(Info.Fecha, 12, 8)
                    End If
                Else
                    If Mid(Info.Fecha, 12, 2) = "12" Then
                        strHora = Mid(Info.Fecha, 12, 8)
                    Else
                        strHora = (Val(Mid(Info.Fecha, 12, 2)) + 12) & Mid(Info.Fecha, 14, 6)
                    End If
                End If

                strSQL = "INSERT INTO ProveedorFactura (CveProveedor,Folio,Serie,Fecha,Certificado,Aprobacion,Anio" & _
                            ",TipoComprobante,CondicionesPago,Descuento,MotivoDescuento,FormaPago,MetodoPago,ImpuestosRetenidos" & _
                            ",ImpuestosTrasladados,Subtotal,Total,FechaLectura,Archivo) " & _
                            "SELECT " & lngCveProveedor & "," & Info.Folio & ",'" & Info.Serie & "','" & Format(Info.Fecha, "yyyy-MM-dd") & " " & strHora & "','" & _
                            Info.noCertificado & "','" & Info.noAprobacion & "'," & Info.anoAprobacion & "," & _
                            "'" & Info.tipoDeComprobante & "','" & Info.condicionesDePago & "'," & Info.descuento & ",'" & _
                            Info.motivoDescuento & "','" & Info.formaDePago & "','" & Info.metodoDePago & "'," & _
                            Info.totalImpuestosRetenidos & "," & Info.totalImpuestosTrasladados & "," & Info.subTotal & "," & Info.total & _
                            ",GETDATE(),'" & vFile & "'"
                gcn.Execute(strSQL)

                lngRenglon = 1
                For Each Element As System.Xml.XmlElement In xmlDoc.SelectNodes("//cfdi:Comprobante/cfdi:Conceptos/*", nsp)
                    pID = Element.Attributes.GetNamedItem("cantidad").Value
                    pName = Element.Attributes.GetNamedItem("unidad").Value
                    pPrice = Element.Attributes.GetNamedItem("descripcion").Value
                    curValorUnitario = Element.Attributes.GetNamedItem("valorUnitario").Value
                    curImporte = Element.Attributes.GetNamedItem("importe").Value

                    If Not Element.Attributes.GetNamedItem("noIdentificacion") Is Nothing Then
                        strCodigoProducto = "'" & Element.Attributes.GetNamedItem("noIdentificacion").Value & "'"
                    Else
                        strCodigoProducto = "'+" & Mid(Replace(Element.Attributes.GetNamedItem("descripcion").Value, "'", ""), 1, 19) & "'"
                    End If

                    rsConsulta.Close()

                    strSQL = "SELECT * FROM ArticuloProveedor where Codigo =" & strCodigoProducto & " AND CveProveedor = " & lngCveProveedor
                    rsConsulta.Open(strSQL, gcn)
                    If rsConsulta.EOF Then
                        MsgBox("El Codigo de Producto del Proveedor No existe")
                        modABC.plngCveArticulo = 0
                        modABC.pblnProveedor = False
                        modABC.pstrNombre = pPrice
                        modABC.pstrFacturaXML = vFile

                        modABC.pstrCodigoArticulo = strCodigoProducto
                        modABC.plngCveProveedor = lngCveProveedor
                        modABC.pcurValorUnitario = curValorUnitario
                        If pName = "1" Then
                            modABC.pstrUnidad = "Pieza"
                        Else
                            modABC.pstrUnidad = pName
                        End If

                        frmNew.ShowDialog()

                        lngCveArticulo = modABC.plngCveArticulo

                        If lngCveArticulo = 0 Then strCodigoProducto = "NULL"

                    Else
                        lngCveArticulo = rsConsulta.Fields("CveArticulo").Value
                        gcn.Execute("UPDATE Articulo SET  PrecioCompra=" & curValorUnitario & _
                                            ",PrecioLista=" & curValorUnitario & _
                                            ",FechaUltimoPrecioCompra ='" & Format(Info.Fecha, "yyyy-MM-dd") & " " & strHora & "' " & _
                                    "WHERE CveArticulo = " & lngCveArticulo & _
                                    " AND FechaUltimoPrecioCompra <'" & Format(Info.Fecha, "yyyy-MM-dd") & " " & strHora & "'")
                    End If

                    rsConsulta.Close()
                    rsConsulta.Open("SELECT * from ProveedorFacturaDetalle where 1= 0", gcn)
                    rsConsulta.AddNew()
                    rsConsulta.Fields("CveProveedor").Value = lngCveProveedor
                    rsConsulta.Fields("Folio").Value = Info.Folio
                    rsConsulta.Fields("Serie").Value = Info.Serie
                    rsConsulta.Fields("NumRenglon").Value = lngRenglon
                    rsConsulta.Fields("Cantidad").Value = pID
                    rsConsulta.Fields("Unidad").Value = Trim(pName)
                    If strCodigoProducto <> "NULL" Then
                        rsConsulta.Fields("Codigo").Value = Mid(strCodigoProducto, 2, Len(strCodigoProducto) - 2)
                    End If
                    rsConsulta.Fields("descripcion").Value = pPrice
                    rsConsulta.Fields("ValorUnitario").Value = curValorUnitario
                    rsConsulta.Fields("Importe").Value = curImporte
                    rsConsulta.Fields("CveArticulo").Value = lngCveArticulo
                    rsConsulta.Update()

                    'strSQL = "INSERT INTO ProveedorFacturaDetalle (CveProveedor,Folio,Serie,NumRenglon,Cantidad" & _
                    '            ",Unidad,Codigo,descripcion" & _
                    '       ",ValorUnitario,Importe,CveArticulo) " & _
                    '      "SELECT " & lngCveProveedor & "," & Info.Folio & ",'" & Info.Serie & "'," & lngRenglon & "," & pID & _
                    '     ",'" & pName & "'," & strCodigoProducto & ",'" & pPrice & "'," & _
                    '    curValorUnitario & "," & curImporte & "," & lngCveArticulo
                    'gcn.Execute(strSQL)

                    lngRenglon = lngRenglon + 1
                Next

            End If

            rsConsulta.Close()
        Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. " & vFile & vbLf & "Original error: " & Ex.Message)

            '<cfdi:Concepto cantidad="40.000" unidad="KGS" noIdentificacion="EI51" descripcion="SOLD INFRA 7018 1/8" valorUnitario="33.85" importe="1354.00"/>
            '<cfdi:Concepto cantidad="10.00" unidad="Pieza" descripcion="GANO-1 1/2 Garruchas para noria de zinc" valorUnitario="38.793103" importe="387.931030"/>

        Finally


        End Try
        nsp = Nothing
        xmlDoc = Nothing
        Comprobante = Nothing
        Impuestos = Nothing
        rsConsulta = Nothing

        LFECHA.Text = Info.Fecha
        LFOLIO.Text = Info.Folio
        LSERIE.Text = Info.Serie
        LEMISORRFC.Text = Info.RFCEmisor
        LEMISORNOMBRE.Text = Info.Emisor_Nombre
        LRECEPTORNOMBRE.Text = Info.Receptor_Nombre
        LRECEPTORRFC.Text = Info.Receptor_RFC
        LRECEPTORNOMBRE.Text = Info.Receptor_Nombre
        LIMPRETENIDOS.Text = Info.totalImpuestosRetenidos
        LIMPTRASLADADOS.Text = Info.totalImpuestosTrasladados
        LANIOAPROB.Text = Info.anoAprobacion
        LCONDICIONES.Text = Info.condicionesDePago
        LDESCUENTO.Text = Info.descuento
        LFORMAPAGO.Text = Info.formaDePago
        LMETODOPAGO.Text = Info.metodoDePago
        LMOTIVO.Text = Info.motivoDescuento
        LCERTIFICADO.Text = Info.noCertificado
        LAPROBACION.Text = Info.noAprobacion
        LSUBTOTAL.Text = Info.subTotal
        LCOMPROBANTE.Text = Info.tipoDeComprobante
        LTOTAL.Text = Info.total
    End Sub
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        AbreConeccion()

        mrs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        mrs.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        mrs.LockType = ADODB.LockTypeEnum.adLockOptimistic

    End Sub

    Private Sub optSeleccion_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSeleccion.CheckedChanged
        Label1.Text = "XML"
        Button1.Text = "Selecciona XML"
        prgProceso.Visible = False
        lblArchivo.Visible = False
    End Sub
    Private Sub optSeleccionMasiva_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSeleccionMasiva.CheckedChanged
        Label1.Text = "Dir."
        Button1.Text = "Selecciona Directorio"

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim lngPosicion As Long

        OpenFileDialog1.Filter = "Archivos XML (*.xml)|*.xml"
        OpenFileDialog1.FilterIndex = 1
        OpenFileDialog1.RestoreDirectory = True

        If optSeleccion.Checked = True Then

            If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Try
                    TextBox1.Text = OpenFileDialog1.FileName
                    ExtraerInfoCFD(TextBox1.Text)

                Catch Ex As Exception
                    MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
                Finally

                End Try
            End If
        Else
            If FolderBrowserDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                TextBox1.Text = FolderBrowserDialog1.SelectedPath
                prgProceso.Visible = True
                lblArchivo.Visible = True
                Try

                    Dim dirDirectorio As New DirectoryInfo(TextBox1.Text)
                    ' Get a reference to each file in that directory.
                    Dim filArchivos As FileInfo() = dirDirectorio.GetFiles()
                    ' Display the names of the files.
                    Dim filNombre As FileInfo

                    lstArchivos.Items.Clear()
                    For Each filNombre In filArchivos
                        lstArchivos.Items.Add(filNombre.Name)
                    Next filNombre

                    prgProceso.Maximum = lstArchivos.Items.Count
                    prgProceso.Value = 0
                    lngPosicion = 1
                    For Each filNombre In filArchivos
                        If filNombre.Extension = ".xml" Then
                            lblArchivo.Text = filNombre.Name
                            ExtraerInfoCFD(TextBox1.Text & "\" & lblArchivo.Text)
                        Else
                            lblArchivo.Text = ""
                        End If
                        prgProceso.Value = lngPosicion
                        lngPosicion = lngPosicion + 1
                    Next filNombre

                    MsgBox("Carga de Archivos realizada con exito", MsgBoxStyle.OkOnly, "Carga XML")
                    prgProceso.Visible = False
                    lblArchivo.Visible = False

                Catch Ex As Exception
                    MessageBox.Show("Problemas para extraer informacion: " & Ex.Message)
                Finally

                End Try

            End If
        End If
    End Sub


End Class

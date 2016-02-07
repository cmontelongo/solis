Option Explicit On
Option Compare Text

Public Class frmRelaciona

    Private Sub frmRelaciona_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim strSQL As String
        Dim rsfLlenaControl As New ADODB.Recordset()
        Dim lngRenglon As Long
        Dim intPosicionPrimerEspacio As Integer
        Dim intPosicionSegundoEspacio As Integer
        Dim intposicionTercerEspacio As Integer
        Dim strNombreDepurado As String

        btnRelaciona.Enabled = False
        btnAlta.Enabled = True

        LEMISORNOMBRE.Text = modABC.pstrNombre
        LEMISORRFC.Text = ""
        txtBuscar.Text = ""

        rsfLlenaControl.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsfLlenaControl.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        rsfLlenaControl.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic

        strNombreDepurado = Replace(modABC.pstrNombre, "'", "")

        intPosicionPrimerEspacio = InStr(strNombreDepurado, " ")
        intPosicionSegundoEspacio = InStr(Mid(strNombreDepurado, intPosicionPrimerEspacio + 1, 17), " ") + intPosicionPrimerEspacio
        intposicionTercerEspacio = InStr(Mid(strNombreDepurado, intPosicionSegundoEspacio + 1, 17), " ")
        If intposicionTercerEspacio = 0 Then
            intposicionTercerEspacio = intPosicionSegundoEspacio
            intPosicionSegundoEspacio = intPosicionPrimerEspacio
        End If

        If modABC.pblnProveedor Then
            Label37.Text = "RFC:"
            LEMISORRFC.Text = modABC.pstrRFC
            If intPosicionPrimerEspacio = 0 Then

                strSQL = "SELECT CveProveedor,Nombre FROM Proveedor WHERE Nombre like '%" & strNombreDepurado & "%'  ORDER BY Nombre"
            Else
                strSQL = "SELECT CveProveedor,Nombre FROM Proveedor WHERE Nombre like '%" & Mid(strNombreDepurado, 1, 8) & "%' or " & _
                    "Nombre like '%" & Mid(strNombreDepurado, 12, 9) & "%' or " & _
                    "Nombre like '%" & Mid(strNombreDepurado, intPosicionSegundoEspacio + 1, intposicionTercerEspacio - 1) & "%' or " & _
                    "Nombre like '%" & Mid(strNombreDepurado, 1, 5) & "%' ORDER BY Nombre"
            End If
        Else
            Label37.Text = "COD:"
            LEMISORRFC.Text = modABC.pstrCodigoArticulo
            If intPosicionPrimerEspacio = 0 Then
                strSQL = "SELECT CveArticulo,Nombre FROM Articulo WHERE Nombre like '%" & strNombreDepurado & "%' OR Nombre LIKE '%" & Replace(LEMISORRFC.Text, "'", "") & "%'  ORDER BY Nombre"
            Else
                strSQL = "SELECT CveArticulo,Nombre FROM Articulo WHERE Nombre like '%" & Mid(strNombreDepurado, 1, 8) & "%' or " & _
                    "Nombre like '%" & Mid(strNombreDepurado, 12, 9) & "%' or " & _
                    "Nombre like '%" & Mid(strNombreDepurado, intPosicionSegundoEspacio + 1, intposicionTercerEspacio - 1) & "%' OR " & _
                    "Nombre LIKE '%" & Replace(LEMISORRFC.Text, "'", "") & "%' or " & _
                    "Nombre like '%" & Mid(strNombreDepurado, 1, 5) & "%' ORDER BY Nombre"
            End If
        End If
        rsfLlenaControl.Open(strSQL, gcn)

        LimpiaBloque(sprRelaciona, 1, 1, sprRelaciona.MaxRows, sprRelaciona.MaxCols)
        sprRelaciona.MaxRows = 0
        sprRelaciona.EditModePermanent = True
        sprRelaciona.ReDraw = False

        Do Until rsfLlenaControl.EOF
            sprRelaciona.MaxRows = sprRelaciona.MaxRows + 1
            lngRenglon = sprRelaciona.MaxRows
            sprRelaciona.Row = lngRenglon

            sprRelaciona.Col = 1
            sprRelaciona.Text = rsfLlenaControl.Fields(0).Value

            sprRelaciona.Col = 2
            sprRelaciona.Text = rsfLlenaControl.Fields("Nombre").Value

            rsfLlenaControl.MoveNext()
        Loop
        rsfLlenaControl.Close()

        sprRelaciona.EditModePermanent = False
        sprRelaciona.ReDraw = True
        rsfLlenaControl = Nothing

    End Sub

    Private Sub btnRelaciona_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRelaciona.Click

        Dim lngRenglon As Long

        For lngRenglon = 1 To sprRelaciona.MaxRows
            sprRelaciona.Row = lngRenglon
            If sprRelaciona.SelBlockRow = lngRenglon Then
                sprRelaciona.Col = 1
                If modABC.pblnProveedor Then
                    modABC.plngCveProveedor = sprRelaciona.Text
                    gcn.Execute("UPDATE Proveedor SET RFC = '" & modABC.pstrRFC & "' where CveProveedor = " & modABC.plngCveProveedor)
                Else
                    modABC.plngCveArticulo = sprRelaciona.Text
                    'gcn.Execute("DELETE ArticuloProveedor WHERE CveArticulo = " & modABC.plngCveArticulo & " AND CveProveedor=" & modABC.plngCveProveedor)
                    gcn.Execute("INSERT INTO ArticuloProveedor (CveArticulo,CveProveedor,Codigo) VALUES(" & modABC.plngCveArticulo & "," & modABC.plngCveProveedor & "," & modABC.pstrCodigoArticulo & ")")

                End If
                Me.Close()
                Exit Sub
            End If
        Next lngRenglon
        
    End Sub

    Private Sub sprRelaciona_BlockSelected(ByVal sender As Object, ByVal e As AxFPSpread._DSpreadEvents_BlockSelectedEvent) Handles sprRelaciona.BlockSelected
        btnRelaciona.Enabled = True
    End Sub


    Private Sub sprRelaciona_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpread._DSpreadEvents_ClickEvent) Handles sprRelaciona.ClickEvent
        btnRelaciona.Enabled = True
        btnRelaciona.Visible = True
    End Sub

    Private Sub btnAlta_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAlta.Click

        Dim xmlDoc As New System.Xml.XmlDocument
        Dim Comprobante As System.Xml.XmlNode
        Dim nodDomicilioFiscal As System.Xml.XmlNodeList
        Dim nsp As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)

        Dim strSQL As String
        Dim rsfLlenaControl As New ADODB.Recordset()
        Dim rskTabla As New ADODB.Recordset()
        Dim intCveUnidadMedida As Integer
        Dim lngCveArticulo As Long
        Dim strNombre As String
        Dim strRFC As String
        Dim strDireccion As String
        Dim strColonia As String
        Dim lngCveCiudad As Long
        Dim strCodigoPostal As String

        rsfLlenaControl.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsfLlenaControl.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
        rsfLlenaControl.LockType = ADODB.LockTypeEnum.adLockReadOnly

        rskTabla.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rskTabla.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
        rskTabla.LockType = ADODB.LockTypeEnum.adLockOptimistic

        If modABC.pblnProveedor Then
            xmlDoc.Load(modABC.pstrFacturaXML)
            nsp.AddNamespace("cfdi", "http://www.sat.gob.mx/cfd/3")

            Comprobante = xmlDoc.ChildNodes(1)
            nodDomicilioFiscal = xmlDoc.SelectNodes("//cfdi:Comprobante/cfdi:Emisor/*", nsp)

            If Not Comprobante("cfdi:Emisor").Attributes.GetNamedItem("nombre") Is Nothing Then
                strNombre = StrConv(Comprobante("cfdi:Emisor").Attributes.GetNamedItem("nombre").Value, VbStrConv.ProperCase)
            Else
                strNombre = ""
            End If

            If Not Comprobante("cfdi:Emisor").Attributes.GetNamedItem("rfc") Is Nothing Then
                strRFC = Comprobante("cfdi:Emisor").Attributes.GetNamedItem("rfc").Value
            Else
                strRFC = ""
            End If
            strDireccion = ""
            strCodigoPostal = "0"
            strColonia = ""

            For Each Element As System.Xml.XmlElement In xmlDoc.SelectNodes("//cfdi:Comprobante/cfdi:Emisor/*", nsp)
                If Element.Name = "cfdi:DomicilioFiscal" Then
                    strDireccion = Element.Attributes.GetNamedItem("calle").Value
                    If Not Element.Attributes.GetNamedItem("noExterior") Is Nothing Then
                        strDireccion = strDireccion & " " & Element.Attributes.GetNamedItem("noExterior").Value
                    End If
                    If Not Element.Attributes.GetNamedItem("noInterior") Is Nothing Then
                        strDireccion = strDireccion & "-" & Element.Attributes.GetNamedItem("noInterior").Value
                    End If
                    strDireccion = StrConv(strDireccion, VbStrConv.ProperCase)
                    If Not Element.Attributes.GetNamedItem("colonia") Is Nothing Then
                        strColonia = StrConv(Element.Attributes.GetNamedItem("colonia").Value, VbStrConv.ProperCase)
                    End If

                    strCodigoPostal = Element.Attributes.GetNamedItem("codigoPostal").Value

                    If Not Element.Attributes.GetNamedItem("municipio") Is Nothing Then
                        strSQL = "SELECT CveCiudad FROM Ciudad where Diversidad Like '%" & Element.Attributes.GetNamedItem("municipio").Value & "%'"
                        rsfLlenaControl.Open(strSQL, gcn)
                        If rsfLlenaControl.EOF Then
                            lngCveCiudad = 0
                        Else
                            lngCveCiudad = rsfLlenaControl.Fields("CveCiudad").Value
                        End If
                        rsfLlenaControl.Close()
                    Else
                        lngCveCiudad = 0
                    End If
                End If
            Next

            strSQL = "SELECT MAX(CveProveedor) CveProveedor FROM Proveedor"
            rsfLlenaControl.Open(strSQL, gcn)
            If rsfLlenaControl.EOF Then
                modABC.plngCveProveedor = 1
            Else
                modABC.plngCveProveedor = rsfLlenaControl("CveProveedor").Value + 1
            End If
            rsfLlenaControl.Close()

            strSQL = "INSERT INTO Proveedor (CveProveedor,Nombre,RFC,Direccion,Colonia,CveCiudad,CodigoPostal,CveMoneda) " & _
                "SELECT " & modABC.plngCveProveedor & ",'" & strNombre & "','" & strRFC & "','" & strDireccion & "'" & _
                        ",'" & strColonia & "'," & lngCveCiudad & "," & strCodigoPostal & "," & modABC.pintCveMoneda
            gcn.Execute(strSQL)
        Else
            strSQL = "SELECT CveUnidadMedida FROM UnidadMedida where Diversidad LIKE '%-" & Trim(modABC.pstrUnidad) & "-%'"
            rsfLlenaControl.Open(strSQL, gcn)
            If rsfLlenaControl.EOF Then
                MsgBox("no encontro unidadmedida")
            Else
                intCveUnidadMedida = rsfLlenaControl("CveUnidadMedida").Value
            End If
            rsfLlenaControl.Close()

            strSQL = "SELECT MAX(CveArticulo) CveArticulo FROM Articulo"
            rsfLlenaControl.Open(strSQL, gcn)
            If rsfLlenaControl.EOF Then
                lngCveArticulo = 1
            Else
                lngCveArticulo = rsfLlenaControl("CveArticulo").Value + 1
            End If
            rsfLlenaControl.Close()

            modABC.plngCveArticulo = lngCveArticulo

            rskTabla.Open("SELECT * from Articulo where 1= 0", gcn)
            rskTabla.AddNew()
            rskTabla.Fields("CveArticulo").Value = lngCveArticulo
            rskTabla.Fields("Nombre").Value = Mid(StrConv(modABC.pstrNombre, VbStrConv.ProperCase), 1, 100)
            rskTabla.Fields("NombreCorto").Value = StrConv(Mid(modABC.pstrNombre, 1, 50), VbStrConv.ProperCase)
            rskTabla.Fields("Activo").Value = 1
            rskTabla.Fields("CveFamilia").Value = 0
            rskTabla.Fields("EsAlmacenable").Value = 0
            rskTabla.Fields("RequiereArmado").Value = 0
            rskTabla.Fields("EsManufacturado").Value = 0
            rskTabla.Fields("CveArticuloEstatus").Value = 1
            rskTabla.Fields("CveUnidadMedidaInventario").Value = intCveUnidadMedida
            rskTabla.Fields("CveUnidadMedidaCompra").Value = intCveUnidadMedida
            rskTabla.Fields("Notas").Value = ""
            rskTabla.Fields("CveUsuarioCreador").Value = "SICIP"
            rskTabla.Fields("FechaAlta").Value = Now
            rskTabla.Fields("CveMoneda").Value = modABC.pintCveMoneda
            rskTabla.Fields("CveUnidadMedidaCotizacion").Value = intCveUnidadMedida
            rskTabla.Fields("PrecioCompra").Value = modABC.pcurValorUnitario
            rskTabla.Fields("PrecioLista").Value = modABC.pcurValorUnitario
            rskTabla.Fields("FechaUltimoPrecioCompra").Value = Now
            rskTabla.Fields("FechaUltimoPrecioLista").Value = Now
            rskTabla.Update()
            rskTabla.Close()

            'gcn.Execute("INSERT Articulo (CveArticulo,Nombre,NombreCorto,Activo,CveFamilia,EsAlmacenable,RequiereArmado,EsManufacturado" & _
            '           ",CveArticuloEstatus,CveUnidadMedidaVenta,CveUnidadMedidaCompra,Notas,CveUsuarioCreador,FechaAlta" & _
            '          ",CveMoneda,CveUnidadMedidaCotizacion,PrecioCompra,PrecioLista" & _
            '         ",FechaUltimoPrecioCompra,FechaUltimoPrecioLista) " & _
            '        "SELECT " & lngCveArticulo & ",'" & StrConv(modABC.pstrNombre, VbStrConv.ProperCase) & "','" & StrConv(Mid(modABC.pstrNombre, 1, 50), VbStrConv.ProperCase) & "',1,0,0,0,0" & _
            '           ",1," & intCveUnidadMedida & "," & intCveUnidadMedida & ",'','SICIP',getdate()" & _
            '          "," & modABC.pintCveMoneda & "," & intCveUnidadMedida & "," & modABC.pcurValorUnitario & "," & modABC.pcurValorUnitario & _
            '         ",GETDATE(),GETDATE()")
            gcn.Execute("INSERT ArticuloProveedor (CveArticulo,CveProveedor,Codigo) " & _
                        "SELECT " & lngCveArticulo & "," & modABC.plngCveProveedor & "," & modABC.pstrCodigoArticulo)

        End If
        Me.Close()
        Exit Sub

    End Sub
    Private Sub Buscar()
        Dim rsfLlenaControl As New ADODB.Recordset()
        Dim lngRenglon As Long
        Dim strSQL As String

        rsfLlenaControl.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsfLlenaControl.CursorType = ADODB.CursorTypeEnum.adOpenStatic
        rsfLlenaControl.LockType = ADODB.LockTypeEnum.adLockBatchOptimistic

        If modABC.pblnProveedor Then
            strSQL = "SELECT CveProveedor,Nombre FROM Proveedor WHERE Nombre like '%" & Replace(txtBuscar.Text, " ", "%") & "%'  ORDER BY Nombre"
        Else
            strSQL = "SELECT CveArticulo,Nombre FROM Articulo WHERE Nombre like '%" & Replace(txtBuscar.Text, " ", "%") & "%'  ORDER BY Nombre"
        End If
        rsfLlenaControl.Open(strSQL, gcn)

        LimpiaBloque(sprRelaciona, 1, 1, sprRelaciona.MaxRows, sprRelaciona.MaxCols)
        sprRelaciona.MaxRows = 0
        sprRelaciona.EditModePermanent = True
        sprRelaciona.ReDraw = False

        Do Until rsfLlenaControl.EOF
            sprRelaciona.MaxRows = sprRelaciona.MaxRows + 1
            lngRenglon = sprRelaciona.MaxRows
            sprRelaciona.Row = lngRenglon

            sprRelaciona.Col = 1
            sprRelaciona.Text = rsfLlenaControl.Fields(0).Value

            sprRelaciona.Col = 2
            sprRelaciona.Text = rsfLlenaControl.Fields("Nombre").Value

            rsfLlenaControl.MoveNext()
        Loop
        rsfLlenaControl.Close()

        sprRelaciona.EditModePermanent = False
        sprRelaciona.ReDraw = True
        rsfLlenaControl = Nothing
    End Sub
    Private Sub btnBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Private Sub txtBuscar_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBuscar.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Buscar()
        End If
    End Sub
End Class
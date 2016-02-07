VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmLogin 
   Caption         =   "Clave de Acceso"
   ClientHeight    =   6645
   ClientLeft      =   2130
   ClientTop       =   1110
   ClientWidth     =   11415
   HelpContextID   =   10
   Icon            =   "SI001.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6645
   ScaleWidth      =   11415
   Tag             =   "1"
   Begin VB.CommandButton cmdAgregar 
      Height          =   495
      Left            =   8880
      Picture         =   "SI001.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdBuscarMecanico 
      Height          =   315
      Left            =   8400
      Picture         =   "SI001.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   315
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   7935
   End
   Begin FPSpread.vaSpread sprInsumos 
      Height          =   2175
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   10575
      _Version        =   393216
      _ExtentX        =   18653
      _ExtentY        =   3836
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "SI001.frx":0AAB
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   1575
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "SI001.frx":0C87
      Top             =   1080
      Width           =   10695
   End
   Begin VB.ComboBox cboArticulo 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   10695
   End
   Begin VB.ComboBox cboArticulos 
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   8415
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Const APLICACION = 2

Dim blnPermiso As Boolean
Private Sub cboBase_Click()

gintCveBase = cboBase.ItemData(cboBase.ListIndex)

'------------------------------------------------------
'    Obtiene parámetros de configuración globales
'------------------------------------------------------
CargaParametrosConfiguracion

'---------------------------------
'   Inicia la forma Principal
'---------------------------------
Me.Hide
frmODT.Show

End Sub

Private Sub cboArticulo_Click()

Dim strSQL As String
Dim rsDetalle As rdoResultset

If cboArticulo.ListIndex >= 0 Then

    sprInsumos.MaxRows = 0

    'Llena los combos
    strSQL = "SELECT Notas from Articulo WHERE CveArticulo =  " & cboArticulo.ItemData(cboArticulo.ListIndex)
    
    Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    If rsDetalle.EOF Then
        MsgBox "No existe Informacion", vbExclamation, "ButtonClick"
    Else
        txtDescripcion.Text = rsDetalle!Notas
    End If
    rsDetalle.Close

    strSQL = "select AMD.Nombre,AM.CantidadRequerida,UM.NombreCorto,ISNULL(AMD.KgPorM2,0) KgPorM2,AMD.KgPorM2 * AM.CantidadRequerida Peso " & _
                    ",ISNULL(AMD.PrecioLista,D.PrecioLista) PrecioLista, (AMD.KgPorM2 * AM.CantidadRequerida) * AMD.PrecioLista Importe " & _
                    ",AM.NumRenglon " & _
        "from Articulo A " & _
            " JOIN ArticuloManufactura AM ON A.CveArticulo = AM.CveArticulo" & _
            " JOIN Articulo AMD ON AMD.CveArticulo = AM.CveArticuloDetalle" & _
            " LEFT JOIN UnidadMedida UM ON UM.CveUnidadMedida = AMD.CveUnidadMedidaCotizacion" & _
            " LEFT JOIN (SELECT AD.CveArticulo,SUM(ADS.PrecioLista*AD.CantidadRequerida) PrecioLista " & _
                        " FROM ArticuloDetalle AD " & _
                            " JOIN Articulo ADS on ADS.CveArticulo = AD.CveArticuloDetalle " & _
                            " GROUP BY AD.CveArticulo) AS D ON D.CveArticulo = AMD.CveArticulo " & _
        "Where A.cvearticulo = " & cboArticulo.ItemData(cboArticulo.ListIndex) & _
        " order by AM.NumRenglon"
        
        
    sprInsumos.EditModePermanent = True
    sprInsumos.Row = sprInsumos.MaxRows
    
    Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    
    ' Llena el spread
    sprInsumos.ReDraw = False
    Do Until rsDetalle.EOF
        sprInsumos.MaxRows = sprInsumos.MaxRows + 1
    
        sprInsumos.Row = sprInsumos.MaxRows
        
        MakeFloatCell 2, 2, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0
        MakeFloatCell 4, 5, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0
        MakeFloatCell 6, 7, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", True, True, 2, 0
    
        sprInsumos.Col = 1 'A
        sprInsumos.Text = rsDetalle!Nombre
        sprInsumos.TypeHAlign = TypeHAlignLeft

        sprInsumos.Col = 2 'B
        sprInsumos.Value = rsDetalle!CantidadRequerida
        sprInsumos.TypeHAlign = TypeHAlignLeft

        sprInsumos.Col = 3 'C
        sprInsumos.TypeHAlign = TypeHAlignCenter
        If Not IsNull(rsDetalle!NombreCorto) Then sprInsumos.Text = rsDetalle!NombreCorto
    
        sprInsumos.Col = 4 'D
        sprInsumos.TypeHAlign = TypeHAlignCenter
        If Not IsNull(rsDetalle!KgPorM2) Then sprInsumos.Text = rsDetalle!KgPorM2
    
        sprInsumos.Col = 5 'E
        sprInsumos.Formula = "B" & sprInsumos.Row & " * D" & sprInsumos.Row
        sprInsumos.TypeHAlign = TypeHAlignLeft
    
        sprInsumos.Col = 6 'F
        sprInsumos.TypeHAlign = TypeHAlignCenter
        If Not IsNull(rsDetalle!PrecioLista) Then sprInsumos.Text = rsDetalle!PrecioLista
    
        sprInsumos.Col = 7 'G
        If rsDetalle!KgPorM2 = 0 Then
            sprInsumos.Formula = "B" & sprInsumos.Row & " * F" & sprInsumos.Row
        Else
            sprInsumos.Formula = "E" & sprInsumos.Row & " * F" & sprInsumos.Row
        End If
        sprInsumos.TypeHAlign = TypeHAlignLeft
    
        rsDetalle.MoveNext
    Loop
    rsDetalle.Close
        
        
        

End If





End Sub


Private Sub cmdAgregar_Click()

Dim strSQL As String
Dim rsDetalle As rdoResultset

If cboArticulos.ListIndex >= 0 Then
    
    'Always have the spreadsheet in edit mode
    sprInsumos.EditModePermanent = True

    sprInsumos.MaxRows = sprInsumos.MaxRows + 1
    
    sprInsumos.Row = sprInsumos.MaxRows
    
    strSQL = "select CveArticulo,A.Nombre,UM.NombreCorto,KgPorM2,CostoMonedaNacional " & _
        "from Articulo A " & _
            "JOIN UnidadMedida UM ON UM.CveUnidadMedida = A.CveUnidadMedidaCotizacion " & _
        "Where A.CveArticulo = " & cboArticulos.ItemData(cboArticulos.ListIndex)
    Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    
    ' Llena el spread
    sprInsumos.ReDraw = False
    Do Until rsDetalle.EOF
    
        MakeFloatCell 2, 2, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0
        MakeFloatCell 4, 5, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0
        MakeFloatCell 6, 7, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", True, True, 2, 0
    
        sprInsumos.Col = 1 'A
        sprInsumos.Text = rsDetalle!Nombre
        sprInsumos.TypeHAlign = TypeHAlignLeft
            
        sprInsumos.Col = 3 'C
        sprInsumos.TypeHAlign = TypeHAlignCenter
        sprInsumos.Text = rsDetalle!NombreCorto
    
        sprInsumos.Col = 4 'D
        sprInsumos.TypeHAlign = TypeHAlignCenter
        sprInsumos.Text = rsDetalle!KgPorM2
    
        sprInsumos.Col = 5 'E
        sprInsumos.Formula = "B" & sprInsumos.Row & " * D" & sprInsumos.Row
        sprInsumos.TypeHAlign = TypeHAlignLeft
    
        sprInsumos.Col = 6 'F
        sprInsumos.TypeHAlign = TypeHAlignCenter
        If Not IsNull(rsDetalle!CostoMonedaNacional) Then sprInsumos.Text = rsDetalle!CostoMonedaNacional
    
        sprInsumos.Col = 7 'G
        sprInsumos.Formula = "E" & sprInsumos.Row & " * F" & sprInsumos.Row
        sprInsumos.TypeHAlign = TypeHAlignLeft
    
        rsDetalle.MoveNext
    Loop
    rsDetalle.Close
    sprInsumos.ReDraw = True
    
    cmdAgregar.Visible = False
    cboArticulos.Visible = False
    
    txtBuscar.Visible = True
    cmdBuscarMecanico.Visible = True
    
End If
End Sub



Sub MakeFloatCell(Col As Long, col2 As Long, Row As Long, row2 As Long, floatmin As String, _
    floatmax As String, floatmoney As Boolean, floatsep As Boolean, decplaces As Integer, fpvalue As Double)
    
    sprInsumos.Col = Col
    sprInsumos.col2 = col2
    sprInsumos.Row = Row
    sprInsumos.row2 = row2
    sprInsumos.BlockMode = True
    'Define cells as type FLOAT
    If floatmoney Then
        sprInsumos.CellType = CellTypeCurrency
        sprInsumos.TypeCurrencyShowSymbol = True
        sprInsumos.TypeCurrencyDecPlaces = decplaces
        sprInsumos.TypeCurrencyShowSep = floatsep
        sprInsumos.TypeCurrencyMin = floatmin
        sprInsumos.TypeCurrencyMax = floatmax
    Else
        sprInsumos.CellType = CellTypeNumber
        sprInsumos.TypeNumberDecPlaces = decplaces
        sprInsumos.TypeNumberShowSep = floatsep
        sprInsumos.TypeNumberMin = floatmin
        sprInsumos.TypeNumberMax = floatmax
    End If
    sprInsumos.Value = fpvalue
    sprInsumos.BlockMode = False
    
End Sub
Private Sub cmdBuscarMecanico_Click()
Dim strSQL As String


Screen.MousePointer = vbHourglass

strSQL = "SELECT CveArticulo,Nombre FROM Articulo WHERE Activo=1 AND CveTipoRecurso = 1 AND Nombre like '%" & Replace(txtBuscar.Text, " ", "%") & "%' ORDER BY Nombre"
LlenaVariosSelectores strSQL, Array("cboArticulos"), Me
If cboArticulos.ListCount > 0 Then
    cboArticulos.Visible = True
    txtBuscar.Visible = False
    cmdBuscarMecanico.Visible = False
    cmdAgregar.Visible = True
    txtBuscar.Text = ""
End If
Screen.MousePointer = vbDefault
End Sub


Public Sub Form_Load()

Dim strSQL As String

CentrarForma Me
'txtServidor =
'Se asignan Variables de Cuenta y Password
gstrLogin = "SICIP"
gstrPassword = "SICIP"
gstrServidor = "NAUTILIUS"
gstrBaseDeDatos = "SICIP"

'CargaParametrosTranspais
AbreConeccion

strSQL = "SELECT CveArticulo,Nombre FROM Articulo WHERE Activo = 1 AND EsManufacturado = 1"
LlenaVariosSelectores strSQL, Array("cboArticulo"), Me

sprInsumos.MaxRows = 0

sprInsumos.Row = -1000

sprInsumos.Col = 1
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "Materiales"
sprInsumos.ColWidth(1) = 24

sprInsumos.Col = 2
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "Cant"

sprInsumos.Col = 3
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "UN"

sprInsumos.Col = 4
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "kg/m/pza"

sprInsumos.Col = 5
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "Peso"

sprInsumos.Col = 6
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "$/UN/kg"

sprInsumos.Col = 7
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "TOTAL"

End Sub
Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtPassword.SetFocus
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)

Dim strSQL As String
Dim rsPermiso As rdoResultset
Dim rsPassword As rdoResultset

If KeyAscii = vbKeyReturn Then
'    txtServidor.SetFocus

    gstrLogin = UCase(txtCuenta.Text)
    
    ' Verifica si tiene acceso a este modulo
    strSQL = "select * from Usuario where CveUsuario = '" & gstrLogin & "'"
    Set rsPassword = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    If rsPassword.EOF Then
        MsgBox " Cuenta no existe "
        rsPassword.Close
        End
    End If
    If UCase(Trim(rsPassword!PASSWORD)) <> UCase(Trim(txtPassword.Text)) Then
        MsgBox " Password es incorrecto "
        rsPassword.Close
        End
    End If

    strSQL = "select * from UsuarioAplicacion where CveUsuario = '" & gstrLogin
    strSQL = strSQL & "' and CveAplicacion = " & APLICACION
    Set rsPermiso = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    If rsPermiso.EOF Then
        MsgBox "No se tiene acceso a este Módulo de SIM"
        rsPermiso.Close
        End
    End If
    rsPermiso.Close
    
    LlenaVariosSelectores "SELECT B.CveBase,B.Nombre FROM Base B, UsuarioBase UB " & _
                        "WHERE B.CveBase = UB.CveBase" & _
                        "  AND UB.CveUsuario = '" & gstrLogin & "' " & _
                        "ORDER BY B.Nombre", Array("cboBase"), Me
    cboBase.SetFocus
    
End If

End Sub
Private Sub txtPassword_LostFocus()
Dim strSQL As String
Dim rsPermiso As rdoResultset
Dim rsPassword As rdoResultset

gstrLogin = UCase(txtCuenta.Text)

' Verifica si tiene acceso a este modulo
strSQL = "select * from Usuario where CveUsuario = '" & gstrLogin & "'"
Set rsPassword = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
If rsPassword.EOF Then
    MsgBox " Cuenta no existe "
    rsPassword.Close
    End
End If
If UCase(Trim(rsPassword!PASSWORD)) <> UCase(Trim(txtPassword.Text)) Then
    MsgBox " Password es incorrecto "
    rsPassword.Close
    txtPassword.SetFocus
    Exit Sub
End If

strSQL = "select * from UsuarioAplicacion where CveUsuario = '" & gstrLogin
strSQL = strSQL & "' and CveAplicacion = " & APLICACION
Set rsPermiso = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
If rsPermiso.EOF Then
    MsgBox "No se tiene acceso a este Módulo de SIM"
    rsPermiso.Close
    End
End If
rsPermiso.Close

LlenaVariosSelectores "SELECT B.CveBase,B.Nombre FROM Base B, UsuarioBase UB " & _
                    "WHERE B.CveBase = UB.CveBase" & _
                    "  AND UB.CveUsuario = '" & gstrLogin & "' " & _
                    "ORDER BY B.Nombre", Array("cboBase"), Me
cboBase.SetFocus

End Sub
Sub Form_Unload(Cancel As Integer)
'*** Code added by VB HelpWriter ***
'*** Subroutine added by VB HelpWriter ***
    'QuitHelp
'***********************************
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdBuscarMecanico_Click
End Sub



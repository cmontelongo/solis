VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmPartidas 
   Caption         =   "Clave de Acceso"
   ClientHeight    =   9600
   ClientLeft      =   4365
   ClientTop       =   225
   ClientWidth     =   11415
   HelpContextID   =   10
   Icon            =   "SI003.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9600
   ScaleWidth      =   11415
   Tag             =   "1"
   Begin VB.TextBox txtTotalRecurso 
      Height          =   285
      Index           =   4
      Left            =   6240
      TabIndex        =   16
      Top             =   9120
      Width           =   1455
   End
   Begin VB.TextBox txtTotalRecurso 
      Height          =   285
      Index           =   3
      Left            =   6240
      TabIndex        =   15
      Top             =   8880
      Width           =   1455
   End
   Begin VB.TextBox txtTotalRecurso 
      Height          =   285
      Index           =   2
      Left            =   6240
      TabIndex        =   14
      Top             =   8640
      Width           =   1455
   End
   Begin VB.TextBox txtTotalRecurso 
      Height          =   285
      Index           =   1
      Left            =   6240
      TabIndex        =   13
      Top             =   8400
      Width           =   1455
   End
   Begin VB.ComboBox cboDescr4 
      Height          =   315
      Left            =   8880
      TabIndex        =   12
      Top             =   720
      Width           =   2175
   End
   Begin VB.ComboBox cboDescr3 
      Height          =   315
      Left            =   6120
      TabIndex        =   11
      Top             =   720
      Width           =   2535
   End
   Begin VB.ComboBox cboDescr2 
      Height          =   315
      Left            =   3240
      TabIndex        =   10
      Top             =   720
      Width           =   2775
   End
   Begin VB.ComboBox cboDescr1 
      Height          =   315
      Left            =   360
      TabIndex        =   9
      Top             =   720
      Width           =   2775
   End
   Begin VB.PictureBox pnlToolbar 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11355
      TabIndex        =   7
      Top             =   0
      Width           =   11415
      Begin ComctlLib.Toolbar tlbCampaña 
         Height          =   405
         Left            =   360
         TabIndex        =   8
         Top             =   15
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   714
         ButtonWidth     =   609
         ButtonHeight    =   609
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "imgIconos"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   16
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   4
               Object.Width           =   320
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Agregar"
               Description     =   "Agregar Registro"
               Object.ToolTipText     =   "Registro Nuevo"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Actualizar"
               Description     =   "Grabar Registro"
               Object.ToolTipText     =   "Grabar Registro"
               Object.Tag             =   ""
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   4
               Object.Width           =   500
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Cancelar"
               Description     =   "Cancelar Acción"
               Object.ToolTipText     =   "Cancelar Acción"
               Object.Tag             =   ""
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   4
               Object.Width           =   500
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Buscar"
               Description     =   "Buscar Registro"
               Object.ToolTipText     =   "Buscar Registro"
               Object.Tag             =   ""
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Enabled         =   0   'False
               Key             =   ""
               Object.Tag             =   ""
               Style           =   4
               Object.Width           =   500
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   4
               Object.Width           =   500
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Enabled         =   0   'False
               Key             =   ""
               Object.Tag             =   ""
               Style           =   4
               Object.Width           =   500
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Enabled         =   0   'False
               Key             =   ""
               Object.Tag             =   ""
               Style           =   4
               Object.Width           =   500
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   4
               Object.Width           =   500
            EndProperty
            BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Imprimir"
               Description     =   "Imprimir"
               Object.ToolTipText     =   "Imprimir"
               Object.Tag             =   ""
               ImageIndex      =   11
            EndProperty
            BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   4
               Object.Width           =   500
            EndProperty
            BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   4
               Object.Width           =   500
            EndProperty
            BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Salir"
               Description     =   "Salir"
               Object.ToolTipText     =   "Salir"
               Object.Tag             =   ""
               ImageIndex      =   13
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdAgregar 
      Height          =   495
      Left            =   8880
      Picture         =   "SI003.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdBuscarMecanico 
      Height          =   315
      Left            =   8400
      Picture         =   "SI003.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   315
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   7935
   End
   Begin FPSpread.vaSpread sprInsumos 
      Height          =   5535
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   10575
      _Version        =   393216
      _ExtentX        =   18653
      _ExtentY        =   9763
      _StockProps     =   64
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "SI003.frx":0AAB
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   855
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "SI003.frx":0D08
      Top             =   1320
      Width           =   10695
   End
   Begin VB.ComboBox cboArticulo 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
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
   Begin ComctlLib.ImageList imgIconos 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":0D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":1028
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":1342
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":165C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":1976
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":1C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":1FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":22C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":25DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":28F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":2C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":2F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":3246
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":3560
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":387A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI003.frx":3B94
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPartidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Const COLCONFIRMA = 1
Const COLNOMBRE = 2
Const COLCANTREQUERIDA = 3
Const COLUNIDADMEDIDA = 4
Const COLKGPORM2 = 5
Const COLPESOTOTAL = 6
Const COLPRECIOLISTA = 7
Const COLPRECIOTOTAL = 8
Const COLARTICULO = 9
Const COLFECHAPRECIO = 10
Const COLTIPORECURSO = 11
Const COLFORMULA = 12

Const COLMAXIMA = 12

Dim blnPermiso As Boolean
Dim mblnLlena As Boolean
Private Sub CargaArticulo(ByVal vlngCveArticulo As Long)
Dim strSQL As String
Dim rsDetalle As rdoResultset
Dim intTipoRecursoAnterior As Integer
Dim x As Boolean

If vlngCveArticulo > 0 Then

    mblnLlena = True
    
    sprInsumos.MaxRows = 0

    'Llena los combos
    strSQL = "SELECT Notas from Articulo WHERE CveArticulo =  " & vlngCveArticulo
    
    Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    If rsDetalle.EOF Then
        MsgBox "No existe Informacion", vbExclamation, "ButtonClick"
    Else
        txtDescripcion.Text = rsDetalle!Notas
    End If
    rsDetalle.Close

    strSQL = "select A.Nombre,AM.CantidadRequerida,UM.NombreCorto,ISNULL(A.KgPorM2,0) KgPorM2,A.KgPorM2 * AM.CantidadRequerida Peso " & _
                    ",ISNULL(A.PrecioLista,D.PrecioLista) PrecioLista, (A.KgPorM2 * AM.CantidadRequerida) * A.PrecioLista Importe " & _
                    ",AM.NumRenglon,A.CveArticulo,A.CveTipoRecurso, TR.Nombre NombreTipoRecurso,A.CalculoPrecioLista " & _
                    ",ISNULL(A.FechaUltimoPrecioLista,D.FechaUltimoPrecioLista) FechaUltimoPrecioLista " & _
        "from ArticuloManufactura AM " & _
            " JOIN Articulo A ON A.CveArticulo = AM.CveArticuloDetalle" & _
            " JOIN TipoRecurso TR ON A.CveTipoRecurso = TR.CveTipoRecurso " & _
            " LEFT JOIN UnidadMedida UM ON UM.CveUnidadMedida = A.CveUnidadMedidaCotizacion" & _
            " LEFT JOIN (SELECT AD.CveArticulo,SUM(ADS.PrecioLista*AD.CantidadRequerida) PrecioLista, MIN(ADS.FechaUltimoPrecioLista) FechaUltimoPrecioLista " & _
                        " FROM ArticuloDetalle AD " & _
                            " JOIN Articulo ADS on ADS.CveArticulo = AD.CveArticuloDetalle " & _
                            " GROUP BY AD.CveArticulo) AS D ON D.CveArticulo = A.CveArticulo " & _
        "Where AM.CveArticulo = " & vlngCveArticulo & _
        " order by AM.NumRenglon"
        
        
    sprInsumos.EditModePermanent = True
    sprInsumos.Row = sprInsumos.MaxRows
    
    Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    
    ' Llena el spread
    sprInsumos.ReDraw = False
    intTipoRecursoAnterior = 0

    Do Until rsDetalle.EOF
        sprInsumos.MaxRows = sprInsumos.MaxRows + 1
    
        sprInsumos.Row = sprInsumos.MaxRows
        
        If intTipoRecursoAnterior <> rsDetalle!CveTipoRecurso Then

            With sprInsumos
                .Row = sprInsumos.Row
                .Col = COLCONFIRMA
                .CellType = CellTypeEdit
                .BackColor = vbYellow
                .ForeColor = vbBlack
                .FontBold = True
                .Text = rsDetalle!NombreTipoRecurso
            End With
            
            sprInsumos.SetText COLTIPORECURSO, sprInsumos.Row, rsDetalle!CveTipoRecurso

            x = sprInsumos.AddCellSpan(COLCONFIRMA, sprInsumos.Row, COLPRECIOTOTAL, 1)
            sprInsumos.TypeHAlign = TypeHAlignCenter

            intTipoRecursoAnterior = rsDetalle!CveTipoRecurso
            sprInsumos.MaxRows = sprInsumos.MaxRows + 1
            sprInsumos.Row = sprInsumos.MaxRows
        End If

        MakeFloatCell COLCANTREQUERIDA, COLCANTREQUERIDA, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0
        MakeFloatCell COLKGPORM2, COLPESOTOTAL, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0
        MakeFloatCell COLPRECIOLISTA, COLPRECIOTOTAL, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", True, True, 2, 0
    
        sprInsumos.Col = COLCONFIRMA
        ' Define cell type as check box
        sprInsumos.CellType = CellTypeCheckBox
        sprInsumos.TypeCheckCenter = True
        ' Make it a three state check box
        sprInsumos.TypeCheckType = TypeCheckTypeNormal
        sprInsumos.BackColor = &HC0C0C0

        sprInsumos.SetCellBorder COLCONFIRMA, sprInsumos.Row, COLCONFIRMA, sprInsumos.Row, 15, 8421504, CellBorderStyleSolid

        sprInsumos.Col = COLNOMBRE 'B
        sprInsumos.Text = rsDetalle!Nombre
        sprInsumos.TypeHAlign = TypeHAlignLeft

        sprInsumos.Col = COLCANTREQUERIDA 'C
        sprInsumos.Value = rsDetalle!CantidadRequerida
        sprInsumos.TypeHAlign = TypeHAlignRight

        sprInsumos.Col = COLUNIDADMEDIDA 'D
        sprInsumos.TypeHAlign = TypeHAlignCenter
        If Not IsNull(rsDetalle!NombreCorto) Then sprInsumos.Text = rsDetalle!NombreCorto
    
        If cboDescr1.Visible = False Then
            sprInsumos.Col = COLKGPORM2 'E
            sprInsumos.ColHidden = False
            sprInsumos.TypeHAlign = TypeHAlignCenter
            If Not IsNull(rsDetalle!KgPorM2) Then sprInsumos.Text = rsDetalle!KgPorM2
        
            sprInsumos.Col = COLPESOTOTAL 'F COLCANTREQUERIDA x COLKGPORM2
            sprInsumos.ColHidden = False
            sprInsumos.Formula = "C" & sprInsumos.Row & " * E" & sprInsumos.Row
            sprInsumos.TypeHAlign = TypeHAlignLeft
        Else
            sprInsumos.Col = COLKGPORM2 'E
            sprInsumos.ColHidden = True
        
            sprInsumos.Col = COLPESOTOTAL 'F
            sprInsumos.ColHidden = True
        End If
        sprInsumos.Col = COLPRECIOLISTA 'G
        sprInsumos.TypeHAlign = TypeHAlignCenter
        If Not IsNull(rsDetalle!PrecioLista) Then sprInsumos.Text = rsDetalle!PrecioLista
    
        sprInsumos.Col = COLPRECIOTOTAL 'H
        sprInsumos.TypeHAlign = TypeHAlignRight
        If rsDetalle!KgPorM2 = 0 Then
            sprInsumos.Formula = "C" & sprInsumos.Row & " * G" & sprInsumos.Row 'COLCANTREQUERIDA x COLPRECIOLISTA
        Else
            sprInsumos.Formula = "F" & sprInsumos.Row & " * G" & sprInsumos.Row 'COLPESOTOTAL x COLPRECIOLISTA
        End If
    
        sprInsumos.Col = COLFECHAPRECIO
        If Not IsNull(rsDetalle!FechaUltimoPrecioLista) Then sprInsumos.Text = rsDetalle!FechaUltimoPrecioLista
        
        sprInsumos.Col = COLARTICULO 'I
        sprInsumos.Text = rsDetalle!CveArticulo
    
        sprInsumos.Col = COLFORMULA ' K
        If Not IsNull(rsDetalle!CalculoPrecioLista) Then sprInsumos.Text = rsDetalle!CalculoPrecioLista
    
        rsDetalle.MoveNext
    Loop
    rsDetalle.Close
      mblnLlena = False
End If

End Sub

Private Sub cboArticulo_Click()

If cboArticulo.ListIndex >= 0 Then CargaArticulo cboArticulo.ItemData(cboArticulo.ListIndex)

End Sub
Private Sub cboDescr1_Click()
Dim strSQL As String
If cboDescr1.ListIndex >= 0 Then
    strSQL = "SELECT DISTINCT -1,Descripcion2 FROM ArticuloLinea WHERE Descripcion1 = '" & cboDescr1.Text & "' AND Descripcion2 != '' ORDER BY Descripcion2"
    LlenaVariosSelectores strSQL, Array("cboDescr2"), Me
    
    If cboDescr2.ListCount = 0 Then
        cboDescr1.Visible = False
        cboDescr2.Visible = False
        cboDescr3.Visible = False
        cboDescr4.Visible = False
        cboArticulo.Visible = True
    End If
End If
End Sub
Private Sub cboDescr2_Click()
Dim strSQL As String
If cboDescr2.ListIndex >= 0 Then
    strSQL = "SELECT DISTINCT -1,Descripcion3 FROM ArticuloLinea WHERE Descripcion1 ='" & cboDescr1.Text & "' AND Descripcion2 = '" & cboDescr2.Text & "' ORDER BY Descripcion3"
    LlenaVariosSelectores strSQL, Array("cboDescr3"), Me
End If
End Sub

Private Sub cboDescr3_Click()
Dim strSQL As String
If cboDescr3.ListIndex >= 0 Then
    strSQL = "SELECT DISTINCT CveArticulo,Descripcion4 FROM ArticuloLinea WHERE Descripcion1 ='" & cboDescr1.Text & "' AND Descripcion2 = '" & cboDescr2.Text & "' AND Descripcion3 = '" & cboDescr3.Text & "' ORDER BY Descripcion4"
    LlenaVariosSelectores strSQL, Array("cboDescr4"), Me
End If
End Sub

Private Sub cboDescr4_Click()
If cboDescr4.ListIndex >= 0 Then CargaArticulo cboDescr4.ItemData(cboDescr4.ListIndex)

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
    
        MakeFloatCell COLCANTREQUERIDA, COLCANTREQUERIDA, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0
        MakeFloatCell COLUNIDADMEDIDA, COLKGPORM2, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", False, True, 2, 0
        MakeFloatCell COLPRECIOLISTA, COLPRECIOLISTA, sprInsumos.Row, sprInsumos.Row, "-99999", "99999", True, True, 2, 0
    
        sprInsumos.Col = COLNOMBRE 'A
        sprInsumos.Text = rsDetalle!Nombre
        sprInsumos.TypeHAlign = TypeHAlignLeft
            
        sprInsumos.Col = COLUNIDADMEDIDA 'C
        sprInsumos.TypeHAlign = TypeHAlignCenter
        sprInsumos.Text = rsDetalle!NombreCorto
    
        sprInsumos.Col = COLKGPORM2 'D
        sprInsumos.TypeHAlign = TypeHAlignCenter
        sprInsumos.Text = rsDetalle!KgPorM2
    
        sprInsumos.Col = COLPESOTOTAL 'E
        sprInsumos.Formula = "B" & sprInsumos.Row & " * D" & sprInsumos.Row
        sprInsumos.TypeHAlign = TypeHAlignLeft
    
        sprInsumos.Col = COLPRECIOLISTA 'F
        sprInsumos.TypeHAlign = TypeHAlignCenter
        If Not IsNull(rsDetalle!CostoMonedaNacional) Then sprInsumos.Text = rsDetalle!CostoMonedaNacional
    
        sprInsumos.Col = COLPRECIOTOTAL 'G
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
Dim Style As Variant
    Dim Color As Variant


CentrarForma Me
'txtServidor =
'Se asignan Variables de Cuenta y Password

strSQL = "SELECT DISTINCT -1,Descripcion1 FROM ArticuloLinea"
LlenaVariosSelectores strSQL, Array("cboDescr1"), Me

strSQL = "SELECT CveArticulo,Nombre FROM Articulo WHERE Activo = 1 AND EsManufacturado = 1"
LlenaVariosSelectores strSQL, Array("cboArticulo"), Me

sprInsumos.MaxRows = 0
sprInsumos.MaxCols = COLMAXIMA

sprInsumos.ColWidth(COLCONFIRMA) = 4
sprInsumos.ColWidth(COLCANTREQUERIDA) = 8
sprInsumos.ColWidth(COLUNIDADMEDIDA) = 8
sprInsumos.ColWidth(COLKGPORM2) = 8
sprInsumos.ColWidth(COLPESOTOTAL) = 8
sprInsumos.ColWidth(COLPRECIOLISTA) = 8
sprInsumos.ColWidth(COLPRECIOTOTAL) = 8
sprInsumos.ColWidth(COLARTICULO) = 8

sprInsumos.Row = -1000

sprInsumos.Col = COLNOMBRE
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "Materiales"
sprInsumos.ColWidth(COLNOMBRE) = 24

sprInsumos.Col = COLCANTREQUERIDA
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "Cant"

sprInsumos.Col = COLUNIDADMEDIDA
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "UN"

sprInsumos.Col = COLKGPORM2
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "kg/m/pza"

sprInsumos.Col = COLPESOTOTAL
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "Peso"

sprInsumos.Col = COLPRECIOLISTA
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "$/UN/kg"

sprInsumos.Col = COLPRECIOTOTAL
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "TOTAL"

sprInsumos.Col = COLFECHAPRECIO
sprInsumos.FontBold = True
sprInsumos.TypeHAlign = TypeHAlignCenter
sprInsumos.Text = "Fecha Precio"

sprInsumos.Col = COLARTICULO
sprInsumos.ColHidden = True

sprInsumos.Col = COLTIPORECURSO
sprInsumos.ColHidden = True

sprInsumos.Col = COLFORMULA
sprInsumos.ColHidden = True

End Sub
Private Sub sprInsumos_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

Dim i As Integer
Dim dblCantidad(1 To 4) As Double
Dim intTipoRecursoAnterior As Integer
Dim intTipoRecursoActual As Integer
Dim varValor As Variant
Dim x As Boolean
Dim strColumna As String
Dim strAccion As String
Dim strRecurso As String

intTipoRecursoAnterior = 0
For i = 1 To sprInsumos.MaxRows
    x = sprInsumos.GetText(COLTIPORECURSO, i, varValor)
    If Val(varValor) > 0 Then intTipoRecursoActual = varValor
    If intTipoRecursoAnterior <> intTipoRecursoActual And intTipoRecursoAnterior > 0 Then
        txtTotalRecurso(intTipoRecursoAnterior).Text = dblCantidad(intTipoRecursoAnterior)
    
    End If
    intTipoRecursoAnterior = intTipoRecursoActual
    x = sprInsumos.GetText(COLCONFIRMA, i, varValor)
    If Val(varValor) > 0 Then
        x = sprInsumos.GetText(COLPRECIOTOTAL, i, varValor)
        dblCantidad(intTipoRecursoAnterior) = dblCantidad(intTipoRecursoAnterior) + Val(varValor)
    End If
Next i
txtTotalRecurso(intTipoRecursoAnterior).Text = dblCantidad(intTipoRecursoAnterior)

For i = 1 To sprInsumos.MaxRows
    x = sprInsumos.GetText(COLFORMULA, i, varValor)
    If Len(varValor) > 0 Then
        sprInsumos.Row = i
        Select Case Trim(Mid(varValor, 1, InStr(1, varValor, " ")))
            Case "COLPRECIOLISTA"
                sprInsumos.Col = COLPRECIOLISTA
                sprInsumos.Value = dblCantidad(Val(Right(varValor, 1)))
        
        End Select
        
    
    End If
Next i

End Sub


Private Sub tlbCampaña_ButtonClick(ByVal Button As ComctlLib.Button)

Dim strSQL As String
Dim vntResponde
Dim lngCveODT As Long
Dim lngCveUnidad As Long
Dim rsConsulta As rdoResultset
Dim intCveUnidad As Integer
Dim lngCveOperador As Long

On Error GoTo err_tlbODT_ButtonClick

Screen.MousePointer = vbHourglass

Select Case Button.Key
   
        
   Case Is = "Agregar"

   Case Is = "Actualizar"

       
            ActualizaDetalle
       
   Case Is = "Borrar"
      
   Case Is = "Cancelar"
   
   Case Is = "Buscar"
        
   Case Is = "Cerrar"

   Case Is = "Salir"
        Screen.MousePointer = vbDefault
        Unload Me


End Select

Screen.MousePointer = vbDefault

Exit Sub

err_tlbODT_ButtonClick:
Screen.MousePointer = vbDefault
Dim strmsg          As String       'String del Error
Dim lngIndice       As Long         'Indice del Error de RDO

Screen.MousePointer = vbDefault

Select Case Err
    Case 40002
        For lngIndice = 0 To rdoErrors.Count - 1
            strmsg = strmsg & rdoErrors(lngIndice).Description & Chr(vbKeyReturn)
        Next lngIndice
        rdoErrors.Clear
    
    Case Else
        strmsg = Err & " " & Error
        Err.Clear
End Select

MsgBox "Error en elemento del ToolBar " & strmsg, vbCritical, "tlbODT_ButtonClick"
Resume Next

End Sub

Public Sub ActualizaDetalle()

On Error GoTo Err_ActualizaDetalle

Dim strSQL As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim i As Long
Dim curPrecioUnitario As Currency
Dim lngArticuloPrincipal As Long
Dim lngArticuloDetalle As Long
Dim lngRenglon As Long
Dim strUnidadMedida As String
Dim dblCantidad As Double
Dim curPrecio As Currency
Dim varValor As Variant
Dim x As Boolean

If cboDescr4.ListIndex >= 0 Then
    lngArticuloPrincipal = cboDescr4.ItemData(cboDescr4.ListIndex)
Else
    lngArticuloPrincipal = cboArticulo.ItemData(cboArticulo.ListIndex)
End If

For i = 1 To sprInsumos.MaxRows
    x = sprInsumos.GetText(COLCONFIRMA, i, varValor)
    If Val(varValor) > 0 Then
        x = sprInsumos.GetText(COLPRECIOTOTAL, i, varValor)
        curPrecioUnitario = curPrecioUnitario + Val(varValor)
    End If
Next i

strSQL = "'<P CveCotizacion=""" & gstrCveCotizacion & """ CveArticulo=""" & lngArticuloPrincipal & """ Cantidad=""" & 1 & """ PrecioUnitario=""" & curPrecioUnitario & """>"
strSQL2 = strSQL
strSQL3 = strSQL
curPrecioUnitario = 0
lngRenglon = 0

For i = 1 To sprInsumos.DataRowCnt
    sprInsumos.Row = i
    x = sprInsumos.GetText(COLCONFIRMA, i, varValor)
    If Val(varValor) > 0 Then
        lngRenglon = lngRenglon + 1
        sprInsumos.Col = COLARTICULO
        lngArticuloDetalle = Val(sprInsumos.Text)
        
        sprInsumos.Col = COLUNIDADMEDIDA
        strUnidadMedida = sprInsumos.Text
        
        x = sprInsumos.GetText(COLCANTREQUERIDA, i, varValor)
        dblCantidad = Val(varValor)
        
        x = sprInsumos.GetText(COLPRECIOTOTAL, i, varValor)
        curPrecioUnitario = Val(varValor)

        If Len(strSQL) > 7800 Then
            If Len(strSQL2) > 7800 Then
                strSQL3 = strSQL3 & "<D R=""" & lngRenglon & """ A=""" & lngArticuloDetalle & """ UM=""" & strUnidadMedida & """ C=""" & dblCantidad & """ PU=""" & curPrecio & """/>"
            Else
                strSQL2 = strSQL2 & "<D R=""" & lngRenglon & """ A=""" & lngArticuloDetalle & """ UM=""" & strUnidadMedida & """ C=""" & dblCantidad & """ PU=""" & curPrecio & """/>"
            End If
        Else
            strSQL = strSQL & "<D R=""" & lngRenglon & """ A=""" & lngArticuloDetalle & """ UM=""" & strUnidadMedida & """ C=""" & dblCantidad & """ PU=""" & curPrecio & """/>"
        End If
    End If
Next i
strSQL = strSQL & "</P>'"
strSQL2 = strSQL2 & "</P>'"
strSQL3 = strSQL3 & "</P>'"

gcn.Execute "EXEC CotizacionArticulo_PROCESO_Update @XML=" & strSQL & ",@Depura = 0,@XML2=" & strSQL2 & ",@XML3=" & strSQL3

Unload Me

Exit Sub

Err_ActualizaDetalle:
Screen.MousePointer = vbDefault
Dim strmsg          As String       'String del Error
Dim lngIndice       As Long         'Indice del Error de RDO

Select Case Err
    Case 40002
        For lngIndice = 0 To rdoErrors.Count - 1
            strmsg = strmsg & rdoErrors(lngIndice).Description & Chr(vbKeyReturn)
        Next lngIndice
        rdoErrors.Clear
    Case Else
        strmsg = Err & " " & Error
        Err.Clear
End Select
  
MsgBox "Error al actualizar el detalle " & strmsg, vbExclamation + vbOKOnly, "ActualizaDetalle"
'mblnEdicion = False
Resume Next

End Sub
Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdBuscarMecanico_Click
End Sub

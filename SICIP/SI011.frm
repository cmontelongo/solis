VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDetalleOT 
   Caption         =   "Estimación de Avance"
   ClientHeight    =   7545
   ClientLeft      =   6780
   ClientTop       =   660
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   13500
   Begin VB.Frame fraObservaciones 
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   9480
      TabIndex        =   39
      Top             =   4200
      Width           =   3975
      Begin VB.TextBox txtObservaciones 
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame fraPresupuesto 
      Height          =   2175
      Left            =   7560
      TabIndex        =   24
      Top             =   1080
      Width           =   4695
      Begin VB.TextBox txtMontoPresupuesto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         TabIndex        =   31
         Top             =   240
         Width           =   1600
      End
      Begin VB.TextBox txtFondoGarantia 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         TabIndex        =   30
         Top             =   600
         Width           =   1600
      End
      Begin VB.TextBox txtMontoAnticipo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         TabIndex        =   29
         Top             =   960
         Width           =   1600
      End
      Begin VB.TextBox txtProcentajeFondoGarantia 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3600
         TabIndex        =   28
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtPagosAcumulados 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         TabIndex        =   27
         Top             =   1320
         Width           =   1600
      End
      Begin VB.TextBox txtMontoEstimacion 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         TabIndex        =   26
         Top             =   1680
         Width           =   1600
      End
      Begin VB.TextBox txtPorcentajeAnticipo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3600
         TabIndex        =   25
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Asignado :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Fondo Garantia :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Anticipo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Pagos Acumulados :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Estimacion :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   4320
         TabIndex        =   33
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   4320
         TabIndex        =   32
         Top             =   960
         Width           =   255
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Height          =   420
      Left            =   240
      TabIndex        =   23
      Top             =   120
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "imlIconos"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Guardar información"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   5400
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imlIconos 
      Left            =   12120
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SI011.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SI011.frx":08DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle de estimaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   19
      Top             =   4200
      Width           =   9255
      Begin FPSpread.vaSpread sprDetalle 
         Height          =   2895
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   9015
         _Version        =   393216
         _ExtentX        =   15901
         _ExtentY        =   5106
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
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         SpreadDesigner  =   "SI011.frx":0BF4
      End
   End
   Begin VB.TextBox txtProveedor 
      Height          =   315
      Left            =   1440
      TabIndex        =   15
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Frame fraProgreso 
      Caption         =   "Progreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   7215
      Begin VB.TextBox txtProgresoOriginal 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5640
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtProgreso 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5640
         TabIndex        =   1
         Top             =   450
         Width           =   1095
      End
      Begin MSComCtl2.UpDown updProgreso 
         Height          =   495
         Left            =   5041
         TabIndex        =   14
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         BuddyControl    =   "prbProgreso"
         BuddyDispid     =   196630
         OrigLeft        =   5280
         OrigTop         =   360
         OrigRight       =   5535
         OrigBottom      =   855
         Increment       =   5
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   5
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.ProgressBar prbProgreso 
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   6840
         TabIndex        =   22
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.ComboBox cboEstatus 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Top             =   2640
      Width           =   5055
   End
   Begin VB.Frame fraFechas 
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6975
      Begin VB.TextBox txtFin 
         Height          =   315
         Left            =   4680
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtEstimadoEntrega 
         Height          =   315
         Left            =   2400
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtInicio 
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   315
         Left            =   4680
         TabIndex        =   6
         Top             =   480
         Width           =   1730
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   42276
      End
      Begin MSComCtl2.DTPicker dtpEstimadoEntrega 
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Top             =   480
         Width           =   1730
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   42276
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1730
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   42276
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrega Real :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Estimado Entrega :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label lblEtiqueta 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado Actual :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblEtiqueta 
      BackStyle       =   0  'Transparent
      Caption         =   "Contratista :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblEtiqueta 
      BackStyle       =   0  'Transparent
      Caption         =   "DescPartida :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11895
   End
End
Attribute VB_Name = "frmDetalleOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Public pbytPartida As Byte



Private Sub DespliegaDetalle()

Dim intRenglon As Integer
Dim rsDetalle As rdoResultset
Dim strSQL As String

On Error GoTo Err_DespliegaDetalle

sprDetalle.MaxRows = 0
sprDetalle.MaxCols = 5
'
sprDetalle.Row = -1000

sprDetalle.Col = 1
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "Fecha"
sprDetalle.ColWidth(1) = 10

sprDetalle.Col = 2
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "% Avance"
sprDetalle.ColWidth(2) = 7

sprDetalle.Col = 3
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "Montos pagados"
sprDetalle.ColWidth(3) = 10

sprDetalle.Col = 4
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "Autorizó"
sprDetalle.ColWidth(4) = 8

sprDetalle.Col = 5
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "Observaciones"
sprDetalle.ColWidth(5) = 35

' Limpia el spread
'LimpiaBloque sprTareas, 1, 1, sprTareas.MaxRows, sprTareas.MaxCols
sprDetalle.MaxRows = 0

' Crea el rdoResultset de la tabla de ODTDetalle
strSQL = "SELECT NumEstimacion,Progreso,Observaciones,FechaEstimacion,MontoPago,CveUsuarioAutorizaEstimacion " & _
    "FROM OTArticuloEstimacion " & _
    "WHERE CveOT =" & glngCveOT & " AND NumPartida = " & pbytPartida & _
    " ORDER BY NumEstimacion"

Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)

sprDetalle.MaxRows = rsDetalle.RowCount
' Llena el spread de Tareas
intRenglon = 1
sprDetalle.ReDraw = False
Do Until rsDetalle.EOF

    sprDetalle.Row = intRenglon
    
    sprDetalle.Col = 1
    sprDetalle.Text = Format(rsDetalle!FechaEstimacion, "DD/MM/YYYY")
    sprDetalle.TypeHAlign = TypeHAlignCenter
    
    sprDetalle.Col = 2
    sprDetalle.Text = rsDetalle!Progreso
    sprDetalle.TypeHAlign = TypeHAlignCenter
    
    sprDetalle.Col = 3
    sprDetalle.TypeHAlign = TypeHAlignRight
    sprDetalle.CellType = CellTypeCurrency
    sprDetalle.TypeCurrencyDecimal = "."
    sprDetalle.TypeCurrencyDecPlaces = 2
    sprDetalle.TypeCurrencySeparator = ","
    sprDetalle.TypeCurrencyShowSep = True
    sprDetalle.TypeCurrencyShowSymbol = True
    sprDetalle.TypeCurrencySymbol = "$"
    sprDetalle.Text = rsDetalle!MontoPago
    
    sprDetalle.Col = 4
    sprDetalle.TypeHAlign = TypeHAlignCenter
    sprDetalle.Text = Trim(rsDetalle!CveUsuarioAutorizaEstimacion)
    
    sprDetalle.Col = 5
    sprDetalle.Text = rsDetalle!Observaciones
    
    rsDetalle.MoveNext
    intRenglon = intRenglon + 1
Loop
rsDetalle.Close
sprDetalle.ReDraw = True

Exit Sub

Err_DespliegaDetalle:
  Screen.MousePointer = vbDefault
  MsgBox "Error al Desplegar Detalle de las estimaciones  " & Error, vbCritical
  Exit Sub
Resume Next

End Sub
Private Sub dtpEstimadoEntrega_Change()
txtEstimadoEntrega.Visible = False
End Sub
Private Sub dtpFin_Change()
txtFin.Visible = False
End Sub
Private Sub dtpInicio_Change()
txtInicio.Visible = False
End Sub
Private Sub Form_Load()
'---------------------------------------------------------------------
'          Rutina para llenar el spread de  Tareas                   -
'---------------------------------------------------------------------
      
Dim intRenglon As Integer
Dim rsDetalle As rdoResultset
Dim rsMecanicos As rdoResultset
Dim rsNombre As rdoResultset
Dim strSQL As String
Dim strNombre As String
Dim X As Boolean
Dim bytProgreso As Byte
On Error GoTo Err_DespliegaDetalle

txtMontoPresupuesto.Text = 0
txtMontoAnticipo.Text = 0
txtFondoGarantia.Text = 0
txtMontoEstimacion.Text = 0
txtProgreso.Text = "0"
LlenaVariosSelectores "SELECT CveOTArticuloEstatus,Nombre FROM OTArticuloEstatus order by Nombre", Array("cboEstatus"), Me

' Crea el rdoResultset de la tabla de ODTDetalle
strSQL = "SELECT OTA.CveOT,OTA.NumPartida,A.Nombre,OTA.FechaInicio,OTA.FechaFin,OTA.FechaEstimadaFin" & _
    ",OTA.CveProveedor,OTA.Observaciones,OTA.Progreso,P.Nombre NomProveedor,OTA.CveOTArticuloEstatus " & _
    ",OTA.MontoPresupuesto,OTA.MontoAnticipo,OTA.PorcentajeAnticipo,OTA.MontoFondoGarantia" & _
    ",OTA.PorcentajeFondoGarantia,OTA.MontoPenaDiaria,ISNULL(EST.MontoPago,0) PagosAcumulados,ISNULL(P.CveProveedorTipo,4) CveProveedorTipo " & _
    "FROM OTArticulo OTA WITH (NOLOCK) " & _
        "JOIN Articulo A WITH (NOLOCK) ON A.CveArticulo = OTA.CveArticulo " & _
        "LEFT JOIN Proveedor P WITH (NOLOCK) ON P.CveProveedor = OTA.CveProveedor " & _
        "LEFT JOIN (SELECT CveOT,NumPartida,SUM(MontoPago) MontoPago " & _
                   "FROM OTArticuloEstimacion WITH (NOLOCK) " & _
                   "GROUP BY CveOT,NumPartida) AS EST ON EST.CveOT = OTA.CveOT and EST.NumPartida = OTA.NumPartida " & _
    "WHERE OTA.CveOT = " & glngCveOT & _
    "  AND OTA.NumPartida = " & pbytPartida

Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
If Not rsDetalle.EOF Then
    lblEtiqueta(0).Caption = rsDetalle!Nombre
    If Not IsNull(rsDetalle!FechaInicio) Then
        dtpInicio.Value = rsDetalle!FechaInicio
        txtInicio.Visible = False
    Else
        txtInicio.Visible = True
    End If
    If Not IsNull(rsDetalle!FechaEstimadaFin) Then
        dtpEstimadoEntrega.Value = rsDetalle!FechaEstimadaFin
        txtEstimadoEntrega.Visible = False
    Else
        txtEstimadoEntrega.Visible = True
    End If
    If Not IsNull(rsDetalle!FechaFin) Then
        dtpFin.Value = rsDetalle!FechaFin
        txtFin.Visible = False
    Else
        txtFin.Visible = True
    End If
    If Not IsNull(rsDetalle!NomProveedor) Then txtProveedor.Text = rsDetalle!NomProveedor
    If Not IsNull(rsDetalle!CveOTArticuloEstatus) Then Posicionaselector rsDetalle!CveOTArticuloEstatus, cboEstatus
    If Not IsNull(rsDetalle!Progreso) Then
        txtProgresoOriginal.Text = rsDetalle!Progreso
        txtProgreso.Text = rsDetalle!Progreso
        prbProgreso.Value = rsDetalle!Progreso
    End If
    If Not IsNull(rsDetalle!Observaciones) Then txtObservaciones.Text = rsDetalle!Observaciones
    
    If rsDetalle!CveProveedorTipo <> 4 Then
        'Proveedor al que se le paga
        fraPresupuesto.Visible = True
    Else
        fraPresupuesto.Visible = False
    End If
    
    If IsNull(rsDetalle!MontoPresupuesto) Then
        txtMontoPresupuesto.Text = "0.00"
    Else
        txtMontoPresupuesto.Text = Format(rsDetalle!MontoPresupuesto, "###,##0.00")
    End If
    If IsNull(rsDetalle!MontoAnticipo) Then
        txtMontoAnticipo.Text = "0.00"
    Else
        txtMontoAnticipo.Text = Format(rsDetalle!MontoAnticipo, "###,##0.00")
    End If
    If IsNull(rsDetalle!PorcentajeAnticipo) Then
        txtPorcentajeAnticipo.Text = 0
    Else
        txtPorcentajeAnticipo.Text = rsDetalle!PorcentajeAnticipo
    End If
    If IsNull(rsDetalle!MontoFondoGarantia) Then
        txtFondoGarantia.Text = "0.00"
    Else
        txtFondoGarantia.Text = Format(rsDetalle!MontoFondoGarantia, "###,##0.00")
    End If
    If IsNull(rsDetalle!PorcentajeFondoGarantia) Then
        txtProcentajeFondoGarantia.Text = 0
    Else
        txtProcentajeFondoGarantia.Text = rsDetalle!PorcentajeFondoGarantia
    End If
    If IsNull(rsDetalle!PagosAcumulados) Then
        txtPagosAcumulados.Text = "0.00"
    Else
        txtPagosAcumulados.Text = Format(rsDetalle!PagosAcumulados, "###,##0.00")
    End If

End If
rsDetalle.Close

DespliegaDetalle

Exit Sub

Err_DespliegaDetalle:
  Screen.MousePointer = vbDefault
  MsgBox "Error al Desplegar Detalle de ODT  " & Error, vbCritical
  Exit Sub
Resume Next

End Sub
Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
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

   Case Is = "Actualizar"
        If ValidaCampos() Then
            If Actualiza() Then
            End If
        End If

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

MsgBox "Error en elemento del ToolBar " & strmsg, vbCritical, "tlbBarraHerramientas_ButtonClick"
Resume Next
End Sub
Private Function ValidaCampos()
  
ValidaCampos = False

If cboEstatus.ListIndex = -1 Then
    Screen.MousePointer = vbDefault
    MsgBox "Debes seleccionar un estado del avance de la partida", vbExclamation
    cboEstatus.SetFocus
    Exit Function
End If
        
If Val(txtProgresoOriginal.Text) = Val(txtProgreso.Text) Then
    Screen.MousePointer = vbDefault
    MsgBox "Debes especificar un progreso para generar una estimacion", vbExclamation
    txtProgreso.SetFocus
    Exit Function
End If
                         
If Val(txtMontoEstimacion.Text) <= 0 And fraPresupuesto.Visible Then
    Screen.MousePointer = vbDefault
    MsgBox "Debe existir un valor para el pago de la estimacion", vbExclamation
    txtMontoEstimacion.SetFocus
    Exit Function
End If

ValidaCampos = True

End Function
Private Function Actualiza() As Boolean
'*****************************************************
'  Procedimiento para actualizar o insertar registros
'*****************************************************
On Error GoTo Err_Actualiza
        
Dim strSQL As String

        
Screen.MousePointer = vbHourglass
Actualiza = False
    
strSQL = "EXEC OTArticulo_PROCESO_ActualizaEstimacion " & _
    "@CveOT = " & glngCveOT & _
    ",@NumPartida = " & pbytPartida & _
    ",@Observaciones = '" & txtObservaciones.Text & "' " & _
    ",@Progreso = " & CInt(Val(txtProgreso.Text) - Val(txtProgresoOriginal.Text)) & _
    ",@CveOTArticuloEstatus = " & cboEstatus.ItemData(cboEstatus.ListIndex) & _
    ",@CveUsuarioAutorizaEstimacion = 'SICIP' " & _
    ",@MontoPago = " & Val(Replace(txtMontoEstimacion.Text, ",", ""))

gcn.Execute strSQL

MsgBox "Actualizacion realizada con Exito", vbOKOnly, "Actualiza"

Exit_Actualiza:
    Actualiza = True
    Screen.MousePointer = vbDefault
    Exit Function

Err_Actualiza:
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
  
MsgBox "Error al Actualizar " & strmsg, vbExclamation + vbOKOnly, "Actualiza"
Exit Function
Resume Next
End Function
Private Sub txtProgreso_Change()
ActualizaMontoEstimacion
End Sub
Private Sub ActualizaMontoEstimacion()
If CLng(Val(txtProgreso.Text)) >= 0 And CLng(Val(txtProgreso.Text)) <= 100 Then
    
    If Val(txtProgreso.Text) <= Val(txtProgresoOriginal.Text) Then
        txtMontoEstimacion.Text = 0
        txtProgreso.Text = txtProgresoOriginal.Text
    Else
        'txtMontoEstimacion.Text = Round(((Val(txtMontoPresupuesto.Text) - (Val(txtMontoAnticipo.Text) + Val(txtFondoGarantia.Text))) * (CInt(Val(txtProgreso.Text) - Val(txtProgresoOriginal.Text)) / 100)) + 0.001, 2)
        txtMontoEstimacion.Text = Round(((Val(SinComas(txtMontoPresupuesto.Text)) - (Val(SinComas(txtMontoAnticipo.Text)) + Val(SinComas(txtFondoGarantia.Text)) + Val(SinComas(txtPagosAcumulados.Text)))) * (CInt(Val(txtProgreso.Text)) / 100)) + 0.001, 2)
    End If
    prbProgreso.Value = CInt(Val(txtProgreso.Text))
    txtMontoEstimacion.Text = Format(txtMontoEstimacion.Text, "###,##0.00")
Else
    MsgBox "Valor no es valido para este rubro", vbCritical + vbOKOnly, "txtProgreso.change"
    txtProgreso.Text = txtProgresoOriginal.Text
    Exit Sub
End If
End Sub

Private Sub txtProgreso_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then
   ' If CLng(txtProgreso.Text) >= 0 And CLng(txtProgreso.Text) <= 100 Then
   '     prbProgreso.Value = CLng(txtProgreso.Text)
   '     'txtMontoEstimacion.Text = (txtMontoPresupuesto.Text - (txtMontoAnticipo.Text + txtFondoGarantia.Text)) * (txtProgreso.Text) / 100
   '     txtMontoEstimacion.Text = Round(((Val(SinComas(txtMontoPresupuesto.Text)) - (Val(SinComas(txtMontoAnticipo.Text)) + Val(SinComas(txtFondoGarantia.Text)) + Val(SinComas(txtPagosAcumulados.Text)))) * (CInt(Val(txtProgreso.Text)) / 100)) + 0.001, 2)
   '     txtMontoEstimacion.Text = Format(txtMontoEstimacion.Text, "###,##0.00")
   ' Else'

   '     MsgBox "Valor no es valido para este rubro", vbCritical + vbOKOnly, "txtProgreso.change"
   '     txtProgreso.Text = txtProgresoOriginal.Text
   '     Exit Sub

   'End If
    ActualizaMontoEstimacion
End If
End Sub
Private Sub updProgreso_DownClick()
If updProgreso.Value < Val(txtProgresoOriginal.Text) Then
    updProgreso.Value = Val(txtProgresoOriginal.Text)
    Exit Sub
End If
txtProgreso.Text = updProgreso.Value
End Sub
Private Sub updProgreso_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Solucionar incidente cuando el mouse se quede presionado mucho tiempo.
If updProgreso.Value < Val(txtProgresoOriginal.Text) Then
    updProgreso.Value = Val(txtProgresoOriginal.Text)
    Exit Sub
End If
txtProgreso.Text = updProgreso.Value
End Sub
Private Sub updProgreso_UpClick()
txtProgreso.Text = updProgreso.Value
End Sub

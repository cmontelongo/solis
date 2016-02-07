VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmOT 
   Caption         =   "Ordenes de Trabajo"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   15840
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   15840
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   12480
      TabIndex        =   25
      Top             =   1440
      Width           =   855
   End
   Begin MSComctlLib.StatusBar staEstatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   6960
      Width           =   15840
      _ExtentX        =   27940
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   19711
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "26/10/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "11:48 a.m."
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraListado 
      Caption         =   "OT's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
      Begin MSComctlLib.TreeView treOT 
         Height          =   4815
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   8493
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle de la OT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   2880
      TabIndex        =   1
      Top             =   2880
      Width           =   12855
      Begin FPSpread.vaSpread sprDetalle 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   12615
         _Version        =   393216
         _ExtentX        =   22251
         _ExtentY        =   5530
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
         OperationMode   =   2
         RowHeaderDisplay=   0
         SelectBlockOptions=   2
         SpreadDesigner  =   "SI012.frx":0000
      End
   End
   Begin VB.Frame fraOT 
      Caption         =   "Datos generales de la OT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.TextBox txtContacto 
         Height          =   315
         Left            =   1440
         TabIndex        =   23
         Top             =   1080
         Width           =   6855
      End
      Begin VB.TextBox txtFechaEstimFin 
         Height          =   315
         Left            =   6480
         TabIndex        =   21
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtOT 
         Height          =   285
         Left            =   6480
         TabIndex        =   19
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtFechaInicio 
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtEstatus 
         Height          =   315
         Left            =   6720
         TabIndex        =   17
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Left            =   6720
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtObra 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtCotizacion 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cboTiempoEntrega 
         Height          =   315
         Left            =   2640
         TabIndex        =   12
         Top             =   1800
         Width           =   5535
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto :"
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
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Estimada de Entrega :"
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
         Left            =   3360
         TabIndex        =   20
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio :"
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
         Index           =   6
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Condiciones de Entrega :"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus :"
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
         Left            =   5760
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Obra :"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha :"
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
         Left            =   5880
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
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
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblEtiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cotizacion :"
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
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Dim mblnEdicion As Boolean
Dim mdatUltimaHoraEjecucion As Date
Dim mblnLlena As Boolean

Const COLUMNAPARTIDA = 1
Const COLUMNADESCRIPCION = 2
Const COLUMNAFECHAINICIO = 3
Const COLUMNAFECHAFIN = 4
Const COLUMNAPROVEEDOR = 5
Const COLUMNAOBSERVACIONES = 6
Const COLUMNAPROGRESO = 7


Private Sub Command1_Click()
frmAsignaciondeTrabajos.Show vbModal
End Sub
Private Sub Form_Load()

On Error GoTo Err_Form_Load

Dim X As Boolean
'Dim intRet As Integer
Dim strSQL As String
Dim strTexto As String
'Dim strFechaLimite As String
'Dim rdBase As rdoResultset
Dim bytColumna As Byte
  
Screen.MousePointer = vbHourglass

'***********************
If App.PrevInstance Then ' checa que no haya otro exe ejecutándose
    strTexto = frmOT.Caption
    frmOT.Caption = ""
    MsgBox App.EXEName & "  Actualmente ejecutándose"
    AppActivate strTexto
    End
End If
'**********************

staEstatusBar.Panels(1).Text = gstrProducto

mdatUltimaHoraEjecucion = DateTime.Now

' Obtiene la resolucion actual a la cambia
'ObtieneResolucion gsngAncho, gsngAlto
'CambiaResolucion RESOLUCIONANCHO, RESOLUCIONALTO

strSQL = "SELECT CveTiempoEntrega,Nombre from TiempoEntrega WHERE Activo = 1 ORDER BY Nombre "

LlenaVariosSelectores strSQL, Array("cboTiempoEntrega"), Me

ActualizaTree


sprDetalle.MaxRows = 0
sprDetalle.MaxCols = COLUMNAPROGRESO + 9
'
sprDetalle.Row = -1000

sprDetalle.Col = COLUMNAPARTIDA
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "N°"
sprDetalle.ColWidth(COLUMNAPARTIDA) = 5

sprDetalle.Col = COLUMNADESCRIPCION
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "Descripcion"
sprDetalle.ColWidth(COLUMNADESCRIPCION) = 28

sprDetalle.Col = COLUMNAFECHAINICIO
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "Fec Inicio"
sprDetalle.ColWidth(COLUMNAFECHAINICIO) = 14

sprDetalle.Col = COLUMNAFECHAFIN
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "Fec Fin"
sprDetalle.ColWidth(COLUMNAFECHAFIN) = 14

sprDetalle.Col = COLUMNAPROVEEDOR
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "Proveedor"
sprDetalle.ColWidth(COLUMNAPROVEEDOR) = 28

sprDetalle.Col = COLUMNAOBSERVACIONES
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "Observaciones"
sprDetalle.ColWidth(COLUMNAOBSERVACIONES) = 28

sprDetalle.Col = COLUMNAPROGRESO
sprDetalle.FontBold = True
sprDetalle.TypeHAlign = TypeHAlignCenter
sprDetalle.Text = "Progreso"

X = sprDetalle.AddCellSpan(COLUMNAPROGRESO, -1000, 10, 1)

For bytColumna = COLUMNAPROGRESO To COLUMNAPROGRESO + 9
    sprDetalle.ColWidth(bytColumna) = 1
Next bytColumna

'
'' Carga controles del rdoResultset
''CargaControlesdeResultset
'
' Despliega el nombre del servidor
'strSQL = "select Nombre from Base Where CveBase = " & gintCveBase
'Set rdBase = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
'staEstatusBar.Panels(2).Text = "Ubicación: " & rdBase!Nombre & "    Versión:" & App.Major & "." & App.Minor & "." & App.Revision
'rdBase.Close
'
'' Rutina para preparar toolbar
'ToolBar_EstadoBrowse tlbODT

Screen.MousePointer = vbDefault

Exit_Form_Load:
    Screen.MousePointer = vbDefault
    Exit Sub
 
Err_Form_Load:
  If Err = 53 Then Resume Next          ' No encuentra algun icono
  Screen.MousePointer = vbDefault
  MsgBox "Error en el Load " & Error, vbCritical
  Unload Me
  Exit Sub
Resume Next
End Sub
Private Sub ActualizaTree()
Dim rsConsulta As rdoResultset
Dim nodx As Node
Dim strSQL As String

treOT.Nodes.Clear

Set nodx = treOT.Nodes.Add(, , "OT", "Ordenes de Trabajo")
nodx.EnsureVisible

strSQL = "SELECT CveOT,NumCotizacion FROM vw_OT WHERE CveOTEstatus = 1"
Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
Set nodx = treOT.Nodes.Add("OT", tvwChild, "PA", "Pendientes por Analizar (" & rsConsulta.RowCount & ")")
Do Until rsConsulta.EOF
    Set nodx = treOT.Nodes.Add("PA", tvwChild, "O-" & CStr(rsConsulta!CveOT), rsConsulta!NumCotizacion)
    nodx.EnsureVisible
    rsConsulta.MoveNext
Loop
rsConsulta.Close

strSQL = "SELECT CveOT,NumCotizacion FROM vw_OT WHERE CveOTEstatus = 2"
Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
Set nodx = treOT.Nodes.Add("OT", tvwChild, "PS", "Pendientes Por Suministro (" & rsConsulta.RowCount & ")")
Do Until rsConsulta.EOF
    Set nodx = treOT.Nodes.Add("PS", tvwChild, "O-" & CStr(rsConsulta!CveOT), rsConsulta!NumCotizacion)
    rsConsulta.MoveNext
Loop
rsConsulta.Close

strSQL = "SELECT CveOT,NumCotizacion FROM vw_OT WHERE CveOTEstatus = 3"
Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
Set nodx = treOT.Nodes.Add("OT", tvwChild, "PP", "Pendiente por Programar (" & rsConsulta.RowCount & ")")
Do Until rsConsulta.EOF
    Set nodx = treOT.Nodes.Add("PP", tvwChild, "O-" & CStr(rsConsulta!CveOT), rsConsulta!NumCotizacion)
    rsConsulta.MoveNext
Loop
rsConsulta.Close

strSQL = "SELECT CveOT,NumCotizacion FROM vw_OT WHERE CveOTEstatus = 4"
Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
Set nodx = treOT.Nodes.Add("OT", tvwChild, "ENP", "En Proceso (" & rsConsulta.RowCount & ")")
Do Until rsConsulta.EOF
    Set nodx = treOT.Nodes.Add("ENP", tvwChild, "O-" & CStr(rsConsulta!CveOT), rsConsulta!NumCotizacion)
    rsConsulta.MoveNext
Loop
rsConsulta.Close

strSQL = "SELECT CveOT,NumCotizacion FROM vw_OT WHERE CveOTEstatus = 5"
Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
Set nodx = treOT.Nodes.Add("OT", tvwChild, "TE", "Terminadas (" & rsConsulta.RowCount & ")")
Do Until rsConsulta.EOF
    Set nodx = treOT.Nodes.Add("TE", tvwChild, "O-" & CStr(rsConsulta!CveOT), rsConsulta!NumCotizacion)
    rsConsulta.MoveNext
Loop
rsConsulta.Close

treOT.Style = tvwTreelinesPlusMinusPictureText
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub

Private Sub sprDetalle_Click(ByVal Col As Long, ByVal Row As Long)

Dim X As Boolean
Dim varValor As Variant

X = sprDetalle.GetText(1, Row, varValor)

frmDetalleOT.pbytPartida = varValor
frmDetalleOT.Show vbModal

End Sub


Private Sub treOT_NodeClick(ByVal Node As MSComctlLib.Node)
If IsNumeric(Right(Node.Key, 1)) Then
    glngCveOT = Mid(Node.Key, InStr(1, Node.Key, "-") + 1, 40)
    CargaControlesdeResultset
End If
End Sub
Private Sub CargaControlesdeResultset()
'******************************
'Despliega los Datos del Registro en la Pantalla
'carga los controles con la información obtenida de la db en el rdoResultset
'******************************

On Error GoTo Err_Carga
Screen.MousePointer = vbHourglass
Dim lngIndiceTemporal As Long
Dim strSQL  As String
Dim rs As rdoResultset

strSQL = "SELECT * FROM vw_OT WHERE CveOT=" & glngCveOT
Set rs = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
If Not rs.EOF Then
    mblnLlena = True
    
    gstrCveCotizacion = rs!NumCotizacion
    txtFecha.Text = Format(rs!FechaOT, FECHADDMMYYYY)
    txtOT.Text = glngCveOT
    If IsNull(rs!FechaEstimadaEntrega) Then
        txtFechaEstimFin.Text = ""
    Else
        txtFechaEstimFin.Text = Format(rs!FechaEstimadaEntrega, FECHADDMMYYYY)
    End If
    If IsNull(rs!FechaInicio) Then
        txtFechaInicio.Text = ""
    Else
        txtFechaInicio.Text = Format(rs!FechaInicio, FECHADDMMYYYY)
    End If
    txtCliente.Text = rs!NomCliente
    txtEstatus.Text = rs!NomOTEstatus
    txtContacto.Text = rs!NombreRepresentante
    txtObra.Text = rs!NomObra
    Posicionaselector rs!CveTiempoEntrega, cboTiempoEntrega
    

'    Select Case rs!CveCotizacionEstatus
'
'        Case 1 'En Proceso
'            ToolBoton_Estado tlbODT, "Agregar", True
'            ToolBoton_Estado tlbODT, "Actualizar", False
'            ToolBoton_Estado tlbODT, "Borrar", True
'            ToolBoton_Estado tlbODT, "Cancelar", False
'            ToolBoton_Estado tlbODT, "Autoriza", True
'            ToolBoton_Estado tlbODT, "Envio", False
'            ToolBoton_Estado tlbODT, "Recibe", False
'            ToolBoton_Estado tlbODT, "OT", False
'            ToolBoton_Estado tlbODT, "Compra", False
'        Case 2 'Pendiente por Autorizar
'            ToolBoton_Estado tlbODT, "Agregar", True
'            ToolBoton_Estado tlbODT, "Actualizar", False
'            ToolBoton_Estado tlbODT, "Borrar", True
'            ToolBoton_Estado tlbODT, "Cancelar", False
'            ToolBoton_Estado tlbODT, "Autoriza", True
'            ToolBoton_Estado tlbODT, "Envio", False
'            ToolBoton_Estado tlbODT, "Recibe", False
'            ToolBoton_Estado tlbODT, "OT", False
'            ToolBoton_Estado tlbODT, "Compra", False
'        Case 3 'Autorizada a
'            ToolBoton_Estado tlbODT, "Agregar", True
'            ToolBoton_Estado tlbODT, "Actualizar", False
'            ToolBoton_Estado tlbODT, "Borrar", True
'            ToolBoton_Estado tlbODT, "Cancelar", False
'            ToolBoton_Estado tlbODT, "Autoriza", False
'            ToolBoton_Estado tlbODT, "Envio", True
'            ToolBoton_Estado tlbODT, "Recibe", False
'            ToolBoton_Estado tlbODT, "OT", False
'            ToolBoton_Estado tlbODT, "Compra", False
'        Case 4 'Enviada a Cliente
'            ToolBoton_Estado tlbODT, "Agregar", True
'            ToolBoton_Estado tlbODT, "Actualizar", False
'            ToolBoton_Estado tlbODT, "Borrar", True
'            ToolBoton_Estado tlbODT, "Cancelar", False
'            ToolBoton_Estado tlbODT, "Autoriza", False
'            ToolBoton_Estado tlbODT, "Envio", False
'            ToolBoton_Estado tlbODT, "Recibe", True
'            ToolBoton_Estado tlbODT, "OT", False
'            ToolBoton_Estado tlbODT, "Compra", False
'
'        Case 5 'Autorizada por Cliente
'            ToolBoton_Estado tlbODT, "Agregar", True
'            ToolBoton_Estado tlbODT, "Actualizar", False
'            ToolBoton_Estado tlbODT, "Borrar", True
'            ToolBoton_Estado tlbODT, "Cancelar", False
'            ToolBoton_Estado tlbODT, "Autoriza", False
'            ToolBoton_Estado tlbODT, "Envio", False
'            ToolBoton_Estado tlbODT, "Recibe", False
'            ToolBoton_Estado tlbODT, "OT", True
'            ToolBoton_Estado tlbODT, "Compra", True
'
'    End Select
'    ToolBoton_Estado tlbODT, "Actualizar", False

Else
    ' Inicializacion Para rdoResultset vacio
    'InicializaCampos
End If
rs.Close

mblnEdicion = False

' Despliega las Tareas de la orden
DespliegaDetalle
ActualizaTree

Screen.MousePointer = vbDefault
Exit Sub

Err_Carga:
  Screen.MousePointer = vbDefault
  MsgBox "Error al Cargar Controles con el rdoResultset " & Error, vbCritical
  mblnEdicion = False
  Exit Sub
  Resume Next
End Sub
Public Sub DespliegaDetalle()
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



' Limpia el spread
'LimpiaBloque sprTareas, 1, 1, sprTareas.MaxRows, sprTareas.MaxCols
sprDetalle.MaxRows = 0

' Crea el rdoResultset de la tabla de ODTDetalle
strSQL = "SELECT OTA.CveOT,OTA.NumPartida,A.Nombre,OTA.FechaInicio,OTA.FechaFin,OTA.FechaEstimadaFin,OTA.CveProveedor,OTA.Observaciones,OTA.Progreso,P.Nombre NomProveedor " & _
    "FROM OTArticulo OTA WITH (NOLOCK) " & _
        "JOIN Articulo A WITH (NOLOCK) ON A.CveArticulo = OTA.CveArticulo " & _
        "LEFT JOIN Proveedor P WITH (NOLOCK) ON P.CveProveedor = OTA.CveProveedor " & _
    "WHERE OTA.CveOT = " & glngCveOT & _
    "ORDER BY NumPartida"

Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)

sprDetalle.MaxRows = rsDetalle.RowCount
' Llena el spread de Tareas
intRenglon = 1
sprDetalle.ReDraw = False
Do Until rsDetalle.EOF

    sprDetalle.Row = intRenglon
    sprDetalle.RowHeight(intRenglon) = 10.5
    
    sprDetalle.Col = COLUMNAPARTIDA
    sprDetalle.Text = rsDetalle!NumPartida
    sprDetalle.TypeHAlign = TypeHAlignCenter
    
    sprDetalle.Col = COLUMNADESCRIPCION
    sprDetalle.Text = rsDetalle!Nombre
    sprDetalle.Col = COLUMNAFECHAINICIO
    sprDetalle.TypeHAlign = TypeHAlignCenter
    If Not IsNull(rsDetalle!FechaInicio) Then sprDetalle.Text = rsDetalle!FechaInicio
    sprDetalle.Col = COLUMNAFECHAFIN
    sprDetalle.TypeHAlign = TypeHAlignCenter
    If Not IsNull(rsDetalle!FechaFin) Then sprDetalle.Text = rsDetalle!FechaFin
    sprDetalle.Col = COLUMNAPROVEEDOR
    If Not IsNull(rsDetalle!NomProveedor) Then sprDetalle.Text = rsDetalle!NomProveedor
    sprDetalle.Col = COLUMNAOBSERVACIONES
    If Not IsNull(rsDetalle!Observaciones) Then sprDetalle.Text = rsDetalle!Observaciones
    
    sprDetalle.Col = COLUMNAPROGRESO
    sprDetalle.TypeHAlign = TypeHAlignLeft

    If IsNull(rsDetalle!Progreso) Then
        sprDetalle.BackColor = vbWhite
        sprDetalle.ForeColor = vbBlack
        bytProgreso = 0
    Else
        bytProgreso = rsDetalle!Progreso
        If bytProgreso = 0 Then
            sprDetalle.BackColor = vbWhite
            sprDetalle.ForeColor = vbBlack
        Else
            sprDetalle.TypeHAlign = TypeHAlignCenter
            sprDetalle.BackColor = &HFF0000
            sprDetalle.ForeColor = vbWhite
        End If
    End If
    sprDetalle.Text = bytProgreso & "%"
    X = sprDetalle.AddCellSpan(COLUMNAPROGRESO, intRenglon, Int(bytProgreso / 10), 1)
    X = sprDetalle.AddCellSpan(COLUMNAPROGRESO + Int(bytProgreso / 10), intRenglon, 10 - Int(bytProgreso / 10), 1)

    rsDetalle.MoveNext
    intRenglon = intRenglon + 1
Loop
rsDetalle.Close
sprDetalle.ReDraw = True

Exit Sub

Err_DespliegaDetalle:
  Screen.MousePointer = vbDefault
  MsgBox "Error al Desplegar Detalle de ODT  " & Error, vbCritical
  mblnEdicion = False
  Exit Sub
Resume Next
End Sub


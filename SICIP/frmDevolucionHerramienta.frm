VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDevolucionHerramienta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolución de Herramienta"
   ClientHeight    =   6585
   ClientLeft      =   1635
   ClientTop       =   1530
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   6480
      TabIndex        =   10
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Guardar"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Frame fraPartidas 
      Caption         =   "Partidas"
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   8895
      Begin FPSpread.vaSpread sprPartidas 
         Height          =   3495
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   8535
         _Version        =   393216
         _ExtentX        =   15049
         _ExtentY        =   6159
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
         SpreadDesigner  =   "frmDevolucionHerramienta.frx":0000
      End
   End
   Begin VB.Frame fraCotizacion 
      Caption         =   "Generales"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.TextBox txtVale 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1500
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   7092
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   42306
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Nombre:"
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
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Vale:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDevolucionHerramienta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gstrVale As String
Dim gstrUsuario As String
Dim gstrObservaciones As String


Private Sub cargarDatos()
'******************************
'Despliega los Datos del Registro en la Pantalla
'carga los controles con la información obtenida de la db en el rdoResultset
'******************************

On Error GoTo Err_Carga
Screen.MousePointer = vbHourglass
Dim lngIndiceTemporal As Long
Dim strSQL  As String
Dim rs As rdoResultset

strSQL = "SELECT * FROM ValeHerramienta WHERE CveValeHerramienta=" & gstrVale
Set rs = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
If Not rs.EOF Then
    mblnLlena = True
    
    txtVale.Text = rs!CveValeHerramienta
    txtNombre.Text = rs!Nombre
    dtpFecha.Value = rs!Fecha
    
    dtpFecha.Enabled = False
    txtNombre.Enabled = False
    txtVale.Enabled = False
  
    gstrUsuario = rs!NombreAutoriza
    gstrObservaciones = rs!Observaciones
End If
rs.Close

mblnEdicion = False

' Despliega las Tareas de la orden
DespliegaDetalle
Screen.MousePointer = vbDefault
Exit Sub

Err_Carga:
  Screen.MousePointer = vbDefault
  MsgBox "Error al Cargar Controles con el rdoResultset " & Error, vbCritical
  mblnEdicion = False
  Exit Sub
  Resume Next

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    CargardoResultsetDeControles
    Me.Hide
End Sub

Private Sub Form_Load()
     gstrVale = frmValeHerramienta.strValePorCapturar
    sprPartidas.MaxRows = 0
    sprPartidas.MaxCols = 5

    sprPartidas.ColWidth(1) = 8
    sprPartidas.ColWidth(2) = 24
    sprPartidas.ColWidth(3) = 10
    sprPartidas.ColWidth(4) = 10

    sprPartidas.Row = -1000

    sprPartidas.Col = 1
    sprPartidas.FontBold = True
    sprPartidas.TypeHAlign = TypeHAlignCenter
    sprPartidas.Text = "Cve"

    sprPartidas.Col = 2
    sprPartidas.FontBold = True
    sprPartidas.TypeHAlign = TypeHAlignCenter
    sprPartidas.Text = "Descripcion"

    sprPartidas.Col = 3
    sprPartidas.FontBold = True
    sprPartidas.TypeHAlign = TypeHAlignCenter
    sprPartidas.Text = "Cantidad Por Regresar"

    sprPartidas.Col = 4
    sprPartidas.FontBold = True
    sprPartidas.TypeHAlign = TypeHAlignCenter
    sprPartidas.Text = "Cantidad a Regresar"

    sprPartidas.Col = 5
    sprPartidas.FontBold = True
    sprPartidas.TypeHAlign = TypeHAlignCenter
    sprPartidas.Text = "Código Artíulo"

     cargarDatos
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'frmValeHerramienta.strValePorCapturar = "Saliendo"
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

On Error GoTo Err_DespliegaDetalle

' Limpia el spread
'LimpiaBloque sprTareas, 1, 1, sprTareas.MaxRows, sprTareas.MaxCols
sprPartidas.MaxRows = 0

' Crea el rdoResultset de la tabla de ODTDetalle
strSQL = "select VA.CveValeHerramienta,VA.CveArticulo,A.Nombre,A.Codigo,VA.Cantidad, DV.CantidadRegresada, (VA.Cantidad-ISNULL(DV.CantidadRegresada,0)) PorRegresar " & _
    "from ValeHerramientaDetalle VA " & _
    "  JOIN Articulo A ON A.CveArticulo = VA.CveArticulo " & _
    " LEFT JOIN (SELECT VD.CveValeHerramienta,VDD.CveArticulo,SUM(VDD.Cantidad) CantidadRegresada " & _
                "FROM DevolucionHerramienta VD " & _
                    "JOIN DevolucionHerramientaDetalle VDD ON VDD.CveDevolucionHerramienta = VD.CveDevolucionHerramienta " & _
                "group by VD.CveValeHerramienta,VDD.CveArticulo) DV ON VA.CveValeHerramienta = DV.CveValeHerramienta AND DV.CveArticulo = VA.CveArticulo " & _
    "WHERE VA.CveValeHerramienta = " & glngCveCotizacion & _
    "ORDER BY A.Nombre"

Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)

sprPartidas.MaxRows = rsDetalle.RowCount
' Llena el spread de Tareas
intRenglon = 1
sprPartidas.ReDraw = False
Do Until rsDetalle.EOF

    sprPartidas.Row = intRenglon
    
    sprPartidas.Col = 1
    sprPartidas.Text = rsDetalle!Codigo
    sprPartidas.TypeHAlign = TypeHAlignLeft
    ProtegeCelda sprPartidas, sprPartidas.Row, 1, True

    sprPartidas.Col = 2
    sprPartidas.Text = rsDetalle!Nombre
    sprPartidas.TypeHAlign = TypeHAlignLeft
    sprPartidas.Text = rsDetalle!Nombre
    ProtegeCelda sprPartidas, sprPartidas.Row, 2, True

    'sprPartidas.Col = 3
    'sprPartidas.CellType = CellTypeNumber
    'sprPartidas.TypeNumberDecPlaces = 0
    'sprPartidas.Value = rsDetalle!Cantidad
    'ProtegeCelda sprPartidas, sprPartidas.Row, 3, True

    sprPartidas.Col = 3
    sprPartidas.CellType = CellTypeNumber
    sprPartidas.TypeNumberDecPlaces = 0
    If IsNull(rsDetalle!PorRegresar) Then
        sprPartidas.Value = 0
    Else
        sprPartidas.Value = rsDetalle!PorRegresar
    End If
    ProtegeCelda sprPartidas, sprPartidas.Row, 3, True

    sprPartidas.Col = 4
    sprPartidas.CellType = CellTypeNumber
    sprPartidas.TypeNumberDecPlaces = 0
    sprPartidas.Value = 0
    ProtegeCelda sprPartidas, sprPartidas.Row, 4, False

    sprPartidas.Col = 5
    sprPartidas.Text = rsDetalle!CveArticulo
    ProtegeCelda sprPartidas, sprPartidas.Row, 5, True
    
    rsDetalle.MoveNext
    intRenglon = intRenglon + 1
Loop
rsDetalle.Close
sprPartidas.ReDraw = True
mblnCambioSprTareas = False


Exit Sub

Err_DespliegaDetalle:
  Screen.MousePointer = vbDefault
  MsgBox "Error al Desplegar Detalle de ODT  " & Error, vbCritical
  mblnEdicion = False
  Exit Sub
Resume Next
End Sub

Private Sub CargardoResultsetDeControles()

Dim strSQL As String
Dim strNumFactura As String
Dim i As Integer
Dim strVale As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim lngCveArticulo As Long
Dim intCantidad As Integer
Dim intRegresado As Integer



On Error GoTo Err_CargaRSet

    Screen.MousePointer = vbHourglass

    If txtVale.Text = "" Then
        strVale = "NULL"
    Else
        strVale = txtVale.Text
    End If
        
    strSQL = "'<O Nombre=""" & txtNombre.Text & """ >"
    strSQL2 = strSQL
    strSQL3 = strSQL

    For i = 1 To sprPartidas.DataRowCnt
        sprPartidas.Row = i
        sprPartidas.Col = 5
        lngCveArticulo = Val(sprPartidas.Text)
        sprPartidas.Col = 3
        intCantidad = Val(sprPartidas.Text)
        sprPartidas.Col = 4
        intRegresado = Val(sprPartidas.Text)

        If Len(strSQL) > 7800 Then
            If Len(strSQL2) > 7800 Then
                strSQL3 = strSQL3 & "<D A=""" & lngCveArticulo & """ C=""" & intRegresado & """>"
            Else
                strSQL2 = strSQL2 & "<D A=""" & lngCveArticulo & """ C=""" & intRegresado & """>"
            End If
        Else
            strSQL = strSQL & "<D A=""" & lngCveArticulo & """ C=""" & intRegresado & """/>"
        End If
    Next i
    strSQL = strSQL & "</O>'"
    strSQL2 = strSQL2 & "</O>'"
    strSQL3 = strSQL3 & "</O>'"

'    gcn.Execute "EXEC ValeHerramienta_PROCESO_ActualizaBeta @ValeHerramienta=" & strVale & _
        ",@Fecha='" & Format(dtpFecha.Value, "YYYY-MM-DD") & "'" & _
        ",@Nombre ='" & txtNombre.Text & "'" & _
        ",@CveUsuario='" & gstrUsuario & "'" & _
        ",@Observaciones='" & gstrObservaciones & "'" & _
        ",@XML=" & strSQL & ",@XML2=" & strSQL2 & ",@XML3=" & strSQL3

    gcn.Execute "Exec DevolucionHerramienta_PROCESO_ActualizaBeta" & _
     " @ValeHerramienta=" & strVale & _
     ",@CveUsuario='" & gstrUsuario & "'" & _
     ",@Observaciones=''" & _
    ",@XML=" & strSQL & ",@XML2=" & strSQL2 & ",@XML3=" & strSQL3
    
    Screen.MousePointer = vbDefault
    Exit Sub

Err_CargaRSet:
    Screen.MousePointer = vbDefault
    MsgBox "Error al Cargar rdoResultset de Controles" & Error, vbCritical
    Exit Sub

End Sub


VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmDevolucionHerramienta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolucion de Herramientas"
   ClientHeight    =   5850
   ClientLeft      =   1590
   ClientTop       =   1650
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10485
   Begin VB.TextBox txtObservaciones 
      Height          =   915
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   4800
      Width           =   3975
   End
   Begin VB.TextBox txtUsuario 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Text            =   "MIGUEL"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdDevolver 
      Caption         =   "Devolver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame fraPartidas 
      Caption         =   "Partidas"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin FPSpread.vaSpread sprPartidas 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7815
         _Version        =   393216
         _ExtentX        =   13785
         _ExtentY        =   7435
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
         SpreadDesigner  =   "SI032.frx":0000
      End
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Observaciones:"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Usuario:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   975
   End
End
Attribute VB_Name = "frmDevolucionHerramienta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Private Sub DespliegaDetalle()

Dim intRenglon As Integer
Dim rsDetalle As rdoResultset
Dim rsMecanicos As rdoResultset
Dim rsNombre As rdoResultset
Dim strSQL As String
Dim strNombre As String
Dim x As Boolean

On Error GoTo Err_DespliegaDetalle

' Crea el rdoResultset de la tabla de ODTDetalle
strSQL = "select VA.CveValeHerramienta,VA.CveArticulo,A.Nombre,A.Codigo,VA.Cantidad,VA.Cantidad - ISNULL(DV.CantidadRegresada,0) CantidadPendiente " & _
    "from ValeHerramientaDetalle VA " & _
    "  JOIN Articulo A ON A.CveArticulo = VA.CveArticulo " & _
    " LEFT JOIN (SELECT VD.CveValeHerramienta,VDD.CveArticulo,SUM(VDD.Cantidad) CantidadRegresada " & _
                "FROM DevolucionHerramienta VD " & _
                    "JOIN DevolucionHerramientaDetalle VDD ON VDD.CveDevolucionHerramienta = VD.CveDevolucionHerramienta " & _
                "group by VD.CveValeHerramienta,VDD.CveArticulo) DV ON VA.CveValeHerramienta = DV.CveValeHerramienta AND DV.CveArticulo = VA.CveArticulo " & _
    "WHERE VA.PendienteEntrega = 1 AND VA.CveValeHerramienta = " & glngCveCotizacion & _
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
    
    sprPartidas.Col = 3
    sprPartidas.CellType = CellTypeNumber
    sprPartidas.TypeNumberDecPlaces = 0
    sprPartidas.Value = rsDetalle!CantidadPendiente
    ProtegeCelda sprPartidas, sprPartidas.Row, 3, True
    
    sprPartidas.Col = 4
    sprPartidas.CellType = CellTypeNumber
    sprPartidas.TypeNumberDecPlaces = 0
    sprPartidas.Value = 0
    
    sprPartidas.Col = 5
    sprPartidas.Text = rsDetalle!CveArticulo
    ProtegeCelda sprPartidas, sprPartidas.Row, 5, True
    
    rsDetalle.MoveNext
    intRenglon = intRenglon + 1
Loop
rsDetalle.Close
sprPartidas.ReDraw = True

Screen.MousePointer = vbDefault
Exit Sub

Err_DespliegaDetalle:
  Screen.MousePointer = vbDefault
  MsgBox "Error al Desplegar Detalle de ODT  " & Error, vbCritical
  Exit Sub
Resume Next

End Sub

Private Sub cmdCancelar_Click()

        Screen.MousePointer = vbDefault
        Unload Me
        
End Sub

Private Sub cmdDevolver_Click()

Dim lngValor As Long
Dim lngRenglon As Long
Dim blnCumplio As Boolean
Dim blnExiste As Boolean
Dim strSQL As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim lngCveArticulo  As Long
Dim intCantidad As Integer
                
        If Len(txtUsuario.Text) = 0 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes proporcionar un usuario", vbExclamation
          txtUsuario.SetFocus
          Exit Sub
        End If
                                                            
        blnExiste = False
       For lngRenglon = 1 To sprPartidas.DataRowCnt
            blnCumplio = sprPartidas.GetInteger(4, lngRenglon, lngValor)
            If lngValor > 0 Then
                blnExiste = True
                Exit For
            End If
        Next lngRenglon
        If Not blnExiste Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes especificar en la partida " & lngRenglon & " la cantidad que se esta regresando", vbExclamation
          sprPartidas.SetFocus
          Exit Sub
        End If
                
        strSQL = "'<O Usuario=""" & txtUsuario.Text & """ Obs=""" & txtObservaciones.Text & """>"
        strSQL2 = strSQL
        strSQL3 = strSQL
        
    For lngRenglon = 1 To sprPartidas.DataRowCnt
        sprPartidas.Row = lngRenglon
        sprPartidas.Col = 5
        lngCveArticulo = Val(sprPartidas.Text)
    
        sprPartidas.Col = 4
        intCantidad = Val(sprPartidas.Text)
    
        If Len(strSQL) > 7800 Then
            If Len(strSQL2) > 7800 Then
                strSQL3 = strSQL3 & "<D A=""" & lngCveArticulo & """ C=""" & intCantidad & """>"
            Else
                strSQL2 = strSQL2 & "<D A=""" & lngCveArticulo & """ C=""" & intCantidad & """>"
            End If
        Else
            strSQL = strSQL & "<D A=""" & lngCveArticulo & """ C=""" & intCantidad & """/>"
        End If

    Next lngRenglon
    strSQL = strSQL & "</O>'"
    strSQL2 = strSQL2 & "</O>'"
    strSQL3 = strSQL3 & "</O>'"

    gcn.Execute "EXEC DevolucionHerramienta_PROCESO_ActualizaBeta @ValeHerramienta=" & glngCveCotizacion & _
        ",@CveUsuario='" & txtUsuario.Text & "'" & _
        ",@Observaciones='" & txtObservaciones & "'" & _
        ",@XML=" & strSQL & ",@XML2=" & strSQL2 & ",@XML3=" & strSQL3
        
MsgBox "Movimiento realizado con Exito", vbInformation, "Devolucion Herramienta"
DespliegaDetalle
Screen.MousePointer = vbDefault
Exit Sub

Err_CargaRSet:
    Screen.MousePointer = vbDefault
    MsgBox "Error al Cargar rdoResultset de Controles" & Error, vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
'---------------------------------------------------------------------
'          Rutina para llenar el spread de  Tareas                   -
'---------------------------------------------------------------------
      


sprPartidas.MaxRows = 0
sprPartidas.MaxCols = 5

sprPartidas.ColWidth(1) = 8
sprPartidas.ColWidth(2) = 24
sprPartidas.ColWidth(3) = 4
sprPartidas.ColWidth(4) = 4

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
sprPartidas.Text = "Cant Pend"

sprPartidas.Col = 4
sprPartidas.FontBold = True
sprPartidas.TypeHAlign = TypeHAlignCenter
sprPartidas.Text = "Cant Reg"

DespliegaDetalle

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDevolucionHerramienta = Nothing
End Sub



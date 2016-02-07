VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmAsignaciondeTrabajos 
   Caption         =   "Asignación de Trabajos"
   ClientHeight    =   7785
   ClientLeft      =   420
   ClientTop       =   1335
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   9630
   Begin VB.Frame fraAsignacion 
      Caption         =   "Contratista o Cuadrilla para Asignar Trabajos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   9375
      Begin VB.ComboBox cboProveedor 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9135
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
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   9375
      Begin FPSpread.vaSpread sprDetalle 
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9135
         _Version        =   393216
         _ExtentX        =   16113
         _ExtentY        =   10186
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
         SelectBlockOptions=   4
         SpreadDesigner  =   "SI013.frx":0000
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Height          =   420
      Left            =   120
      TabIndex        =   4
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
      Left            =   6960
      Top             =   0
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
            Picture         =   "SI013.frx":01DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SI013.frx":0AB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgUncheck 
      Height          =   720
      Left            =   7560
      Picture         =   "SI013.frx":0DD2
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgCheck 
      Height          =   720
      Left            =   8400
      Picture         =   "SI013.frx":2914
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmAsignaciondeTrabajos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Const COLUMNASELECCIONAR = 1
Const COLUMNAPARTIDA = 2
Const COLUMNADESCRIPCION = 3
Const COLUMNAPROGRESO = 4
Private Sub Contrato()

Dim objWord As Object
Dim objDocumento As Object
Dim blnPersonaMoral As Boolean

Set objWord = CreateObject("Word.Application")

blnPersonaMoral = False

With objWord
    Set objDocumento = .Documents.Add("e:\sicip\ContratoObra.dotx")
        With objWord.ActiveDocument
            If blnPersonaMoral Then
                .Bookmarks("NombreProveedor1").Range.Text = "EMPRESA MORAL, SA"
                .Bookmarks("NombreProveedor").Range.Text = "EMPRESA MORAL, SA"
                .Bookmarks("ApoderadoExpresa").Range.Text = "Declara a través, de su expresado apoderado legal lo siguiente"
                .Bookmarks("DeclaracionInciso1").Range.Text = "Moral legalmente constituida y registrada bajo Razón " & _
                        "Social " & "RICARDO JAVIER SALINAS OVIEDO" & ", declara que es que es una sociedad anónima debidamente " & _
                        "constituida de conformidad con las normas aplicables de la Ley General de Sociedades Mercantiles, " & _
                        "que se encuentra representada en este acto por " & "aqui va el nombre del representante" & " quien cuenta con " & _
                        "todas las facultades necesarias para ejercer ese derecho, las cuales no le han sido revocadas o limitadas."
                .Bookmarks("ProveedorFirma").Range.Text = "RICARDO JAVIER SALINAS OVIEDO"
                .Bookmarks("TituloFirmaProveedor").Range.Text = "Representante Regional"
                
            Else
                .Bookmarks("NombreProveedor1").Range.Text = "RICARDO JAVIER SALINAS OVIEDO"
                .Bookmarks("NombreProveedor").Range.Text = " "
                .Bookmarks("DeclaracionInciso1").Range.Text = "Física mayor de edad, con plena capacidad legal para cumplir derechos y obligaciones."
                .Bookmarks("ProveedorFirma").Range.Text = "RICARDO JAVIER SALINAS OVIEDO"
                .Bookmarks("TituloFirmaProveedor").Range.Text = " "
                
            End If
            .Bookmarks("RFCProveedor").Range.Text = "SAOR610409G93"
            .Bookmarks("DomicilioProveedor").Range.Text = "Ave. Portal No. 225 Col. Portal del Huajuco Monterrey, N.L. CP 64989"
            .Bookmarks("IMSSProveedor").Range.Text = "D5011350109"
            .Bookmarks("Nave").Range.Text = "KRAEM"
            .Bookmarks("ObraDomicilio").Range.Text = "PROLONGACION AV. ALBORADA LOTE 6B, PARQUE INDUSTRIAL FINSA GUADALUPE AEROPUERTO, GUADALUPE N.L."
            .Bookmarks("TrabajosARealizar").Range.Text = "la cimentación de un muro bajo y obra civil para la cimentación de poste y riel"
            .Bookmarks("Monto").Range.Text = "455,980.00 pesos (Cuatrocientos cincuenta y cinco mil novecientos ochenta pesos 00/100 M.N.)"
            .Bookmarks("FinObra").Range.Text = "30 de Noviembre del 2015"
            .Bookmarks("AnticipoPorcentaje").Range.Text = "0.00%"
            .Bookmarks("FondoGarantiaPorcentaje").Range.Text = "10.00%"
            .Bookmarks("MontoPorPena").Range.Text = "0.00 (Cero pesos 00/100 m.n.)"
            .Bookmarks("FechaContrato").Range.Text = "20 de Agosto del 2015"
            .Bookmarks("ProveedorTestigo").Range.Text = "Aqui va el nombre del testigo"
            .Bookmarks("NumContrato").Range.Text = "001-RS"
            .Bookmarks("NombreProyecto").Range.Text = "SEOHAN"
            .Bookmarks("UbicacionProyecto").Range.Text = "AV. PROLONGACION RUIZ CORTINEZ LOTE NO. 002 MANZANA NO. 243 ENTRE AV. ALBORADA"
            .Bookmarks("MotivoContrato").Range.Text = "MURO BAJO, OBRA CIVIL PARA CIMENTACION POSTE Y RIEL."
            .Bookmarks("FechaDocumento").Range.Text = "20 de Agosto del 2015"
        End With
    End With
objDocumento.SaveAs FileName:="W:\TestWordDoc.doc"

objDocumento.Close False
objWord.Quit False

Set objDocumento = Nothing
Set objWord = Nothing

End Sub

Private Sub DespliegaDetalle()

Dim strSQL As String
Dim intRenglon As Integer
Dim rsDetalle As rdoResultset

' Crea el rdoResultset de la tabla de ODTDetalle
strSQL = "SELECT OTA.CveOT,OTA.NumPartida,A.Nombre,OTA.FechaInicio,OTA.FechaFin,OTA.FechaEstimadaFin,OTA.CveProveedor,OTA.Observaciones,OTA.Progreso,P.Nombre NomProveedor,OTA.CveOTArticuloEstatus " & _
    "FROM OTArticulo OTA WITH (NOLOCK) " & _
        "JOIN Articulo A WITH (NOLOCK) ON A.CveArticulo = OTA.CveArticulo " & _
        "LEFT JOIN Proveedor P WITH (NOLOCK) ON P.CveProveedor = OTA.CveProveedor " & _
    "WHERE OTA.CveProveedor IS NULL AND OTA.CveOT = " & glngCveOT & _
    "  ORDER BY OTA.NumPartida"
    
Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
sprDetalle.MaxRows = rsDetalle.RowCount
' Llena el spread de Tareas
intRenglon = 1
sprDetalle.ReDraw = False
Do Until rsDetalle.EOF

    sprDetalle.Row = intRenglon
    sprDetalle.RowHeight(intRenglon) = 38
    
    sprDetalle.Col = COLUMNASELECCIONAR
    ' Define cell type as check box
    sprDetalle.CellType = CellTypeCheckBox
    ' Center the check box within the cell
    sprDetalle.TypeCheckCenter = True
    ' Make it a three state check box
    sprDetalle.TypeCheckType = TypeCheckTypeNormal
    ' Define the pictures used for each state of the check box
    ' Picture for False state
    sprDetalle.TypeCheckPicture(0) = imgUncheck.Picture
    ' Picture for True state
    sprDetalle.TypeCheckPicture(1) = imgCheck.Picture
    sprDetalle.ColWidth(COLUMNASELECCIONAR) = 7.25

    sprDetalle.Col = COLUMNAPARTIDA
    sprDetalle.Text = rsDetalle!NumPartida
    sprDetalle.TypeHAlign = TypeHAlignCenter
    sprDetalle.TypeVAlign = TypeVAlignCenter
    sprDetalle.ColWidth(COLUMNAPARTIDA) = 5
    
    sprDetalle.Col = COLUMNADESCRIPCION
    sprDetalle.Text = rsDetalle!Nombre
    sprDetalle.TypeEditMultiLine = True
    sprDetalle.ColWidth(COLUMNADESCRIPCION) = 38
    
'    sprDetalle.Col = COLUMNAPROGRESO
'    sprDetalle.BackColor = &HFF0000
'    sprDetalle.ForeColor = vbWhite
'    sprDetalle.TypeHAlign = TypeHAlignCenter
'    sprDetalle.TypeVAlign = TypeVAlignCenter
'    bytProgreso = 0
'    If Not IsNull(rsDetalle!Progreso) Then
'        bytProgreso = rsDetalle!Progreso
'    Else
'        sprDetalle.BackColor = vbWhite
'        sprDetalle.ForeColor = vbBlack
'    End If
'    sprDetalle.Text = bytProgreso & "%"
'    X = sprDetalle.AddCellSpan(COLUMNAPROGRESO, intRenglon, Int(bytProgreso / 10), 1)
'    X = sprDetalle.AddCellSpan(COLUMNAPROGRESO + Int(bytProgreso / 10), intRenglon, 10 - Int(bytProgreso / 10), 1)

    rsDetalle.MoveNext
    intRenglon = intRenglon + 1
Loop
rsDetalle.Close
sprDetalle.ReDraw = True

End Sub

Private Sub Form_Load()
'---------------------------------------------------------------------
'          Rutina para llenar el spread de  Tareas                   -
'---------------------------------------------------------------------
      
Dim bytColumna As Byte

On Error GoTo Err_DespliegaDetalle

sprDetalle.MaxRows = 0
sprDetalle.MaxCols = COLUMNAPROGRESO + 9

For bytColumna = COLUMNAPROGRESO To COLUMNAPROGRESO + 9
    sprDetalle.ColWidth(bytColumna) = 2
Next bytColumna

LlenaVariosSelectores "SELECT CveProveedor,Nombre FROM Proveedor WHERE CveProveedorTipo in(3,4) order by Nombre", Array("cboProveedor"), Me

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
            Actualiza
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

Private Sub Actualiza()
'*****************************************************
'  Procedimiento para actualizar o insertar registros
'*****************************************************
On Error GoTo Err_Actualiza
        
Dim strSQL As String
Dim lngRenglon As Long
Dim lngValor As Long
Dim blnCumplio As Boolean
Dim varValor As Variant

Screen.MousePointer = vbHourglass

For lngRenglon = 1 To sprDetalle.DataRowCnt
    blnCumplio = sprDetalle.GetInteger(1, lngRenglon, lngValor)
    If Abs(lngValor) = 1 Then
        blnCumplio = sprDetalle.GetText(COLUMNAPARTIDA, lngRenglon, varValor)
    
        gcn.Execute "UPDATE OTArticulo SET CveProveedor = " & cboProveedor.ItemData(cboProveedor.ListIndex) & _
            " WHERE CveOT = " & glngCveOT & _
             " AND NumPartida=" & varValor
    End If
Next lngRenglon

MsgBox "Actualizacion realizada con Exito", vbOKOnly, "Actualiza"
DespliegaDetalle

Exit_Actualiza:
    Screen.MousePointer = vbDefault
    Exit Sub

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
Exit Sub
Resume Next
End Sub

Private Function ValidaCampos()

Dim lngRenglon As Long
Dim blnExiste As Boolean
Dim blnCumplio As Boolean
Dim lngValor As Long
  
ValidaCampos = False
blnExiste = False

For lngRenglon = 1 To sprDetalle.DataRowCnt
    blnCumplio = sprDetalle.GetInteger(1, lngRenglon, lngValor)
    If Abs(lngValor) = 1 Then
        blnExiste = True
        Exit For
    End If
Next lngRenglon

If Not blnExiste Then
    Screen.MousePointer = vbDefault
    MsgBox "Debes seleccionar al menos una partida para asignar.", vbExclamation
    sprDetalle.SetFocus
    Exit Function
End If

If cboProveedor.ListIndex = -1 Then
    Screen.MousePointer = vbDefault
    MsgBox "Debes seleccionar una opcion de la lista.", vbExclamation
    cboProveedor.SetFocus
    Exit Function
End If

ValidaCampos = True

End Function

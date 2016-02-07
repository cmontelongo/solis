VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptVales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   915
   ClientTop       =   1560
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9330
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
      Height          =   2895
      Left            =   3480
      TabIndex        =   8
      Top             =   2880
      Width           =   3015
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16711681
         CurrentDate     =   42306
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16711681
         CurrentDate     =   42306
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Hasta:"
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
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Desde:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
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
      Left            =   6720
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   6720
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame fraArticulos 
      Caption         =   "Articulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   3135
      Begin VB.OptionButton optTodosArticulos 
         Caption         =   "Todos"
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
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optSelectivoArticulos 
         Caption         =   "Selectivo"
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
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   1245
      End
      Begin VB.ListBox lstArticulos 
         Height          =   2010
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame fraEstatus 
      Caption         =   "Estatus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton optTodosEstatus 
         Caption         =   "Todos"
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
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optSelectivoEstatus 
         Caption         =   "Selectivo"
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
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   1245
      End
      Begin VB.ListBox lstEstatus 
         Height          =   1620
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.Frame fraSolicitante 
      Caption         =   "Solicitante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.OptionButton optTodosSolicitantes 
         Caption         =   "Todos"
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
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optSelectivoSolicitantes 
         Caption         =   "Selectivo"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   1245
      End
      Begin VB.ListBox lstSolicitantes 
         Height          =   1620
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmRptVales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Dim FechaInicio As Date
Dim FechaFinal As Date
Private Sub cmdAceptar_Click()

'****************************************************************************
'   Procedimiento que filtra la consulta
'****************************************************************************

Dim blnselectivoSolicitantes As Boolean
Dim blnselectivoEstatus As Boolean
Dim blnselectivoArticulos As Boolean
Dim strLista As String
Dim strSQL As String
Dim i As Integer
Dim strFechas As String

Screen.MousePointer = vbHourglass

' Valida que sean fechas correctas
If Not IsDate(dtpInicio.Value) Then Exit Sub
If Not IsDate(dtpFinal.Value) Then Exit Sub

'-----------------------------------------------------------------------------------------------
' Valida que si es selectivo tenga al menos uno seleccionado en UnidadTipo en caso de ser selectivo
'-----------------------------------------------------------------------------------------------
strLista = ""
blnselectivoSolicitantes = False
If optSelectivoSolicitantes Then
    For i = 0 To lstSolicitantes.ListCount - 1
        If lstSolicitantes.Selected(i) Then
            blnselectivoSolicitantes = True
            Exit For
        End If
    Next
    If Not blnselectivoSolicitantes Then Exit Sub
End If

' Prepara los criterios de seleccion por Familia
If optSelectivoSolicitantes Then
    strLista = strLista & " And " & ListaMultiselect(lstSolicitantes, "Nombre")
End If

'---------------------------------------------------------------------------------------------
' Valida que si es selectivo tenga al menos uno seleccionado en Razon en caso de ser selectivo
'---------------------------------------------------------------------------------------------
blnselectivoEstatus = False
If optSelectivoEstatus Then
    For i = 0 To lstEstatus.ListCount - 1
        If lstEstatus.Selected(i) Then
            blnselectivoEstatus = True
            Exit For
        End If
    Next
    If Not blnselectivoEstatus Then Exit Sub
End If

' Prepara los criterios de seleccion por Razon
If optSelectivoEstatus Then
    strLista = strLista & " And " & ListaMultiselect(lstEstatus, "CveValeHerramientaEstatus")
End If

'-----------------------------------------------------------------------------------------------
' Valida que si es selectivo tenga al menos uno seleccionado en Familia Almacen en caso de ser selectivo
'-----------------------------------------------------------------------------------------------
blnselectivoArticulos = False
If optSelectivoArticulos Then
    For i = 0 To lstArticulos.ListCount - 1
        If lstArticulos.Selected(i) Then
            blnselectivoArticulos = True
            Exit For
        End If
    Next
    If Not blnselectivoArticulos Then Exit Sub
End If

' Prepara los criterios de seleccion por Familia
If optSelectivoArticulos Then
    strLista = strLista & " And " & ListaMultiselect(lstArticulos, "CveArticulo")
End If

gstrArchivoRpt = gstrDirectorioRpt & "SI006.rpt"

strSQL = "SELECT vw_ValeHerramienta.CveValeherramienta, vw_ValeHerramienta.Fecha, vw_ValeHerramienta.Nombre, vw_ValeHerramienta.Observaciones" & _
    ",vw_ValeHerramienta.NombreAutoriza, vw_ValeHerramienta.ValeHerramientaEstatus, vw_ValeHerramienta.ArticuloNombre, vw_ValeHerramienta.Cantidad" & _
    ",vw_ValeHerramienta.CantidadDevuelta, vw_ValeHerramienta.Codigo " & _
 "FROM  SICIP.dbo.vw_ValeHerramienta vw_ValeHerramienta "
strSQL = strSQL & " WHERE vw_ValeHerramienta.Fecha >= '" & Format(dtpInicio.Value, FECHAMMDDYYYY) & "' "
strSQL = strSQL & " AND vw_ValeHerramienta.Fecha <= '" & Format(dtpFinal.Value, FECHAMMDDYYYY & HORAMINUTOS) & "' "
strSQL = strSQL & strLista
strSQL = strSQL & " ORDER BY vw_ValeHerramienta.CveValeherramienta"

gstrSQL = strSQL
Unload Me
Screen.MousePointer = vbDefault


End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmRptVales = Nothing

End Sub
Private Sub Form_Load()
Dim strSQL As String

CentrarForma Me
FechaInicio = "01/" & Month(Now) & "/" & Year(Now)
FechaFinal = Now
dtpInicio.Value = CStr(Format(FechaInicio, FECHADDMMYY))
dtpFinal.Value = CStr(Format(FechaFinal, FECHADDMMYY & HORAFINAL))

strSQL = "select DISTINCT 1,Nombre from ValeHerramienta ORDER BY Nombre "
strSQL = strSQL & "select CveValeHerramientaEstatus, Nombre  from ValeHerramientaEstatus order by Nombre "
strSQL = strSQL & "SELECT A.CveArticulo,A.Nombre " & _
    "FROM Articulo AS A JOIN Familia F ON A.CveFamilia = F.CveFamilia " & _
    "WHERE F.CveRama = 2 ORDER BY A.Nombre"
LlenaVariosSelectores strSQL, Array("lstSolicitantes", "lstEstatus", "lstArticulos"), Me

optTodosSolicitantes.Value = 1
optTodosEstatus.Value = 1
optTodosArticulos.Value = 1

gstrSQL = ""

Screen.MousePointer = vbDefault

End Sub

Private Sub optSelectivoArticulos_Click()
lstArticulos.Enabled = True
End Sub

Private Sub optSelectivoEstatus_Click()
lstEstatus.Enabled = True
End Sub

Private Sub optSelectivoSolicitantes_Click()
lstSolicitantes.Enabled = True
End Sub

Private Sub optTodosArticulos_Click()
Dim i As Integer

lstArticulos.Enabled = False

If optTodosArticulos.Value Then
   For i = 0 To lstArticulos.ListCount - 1
       lstArticulos.Selected(i) = False
    Next
End If
End Sub

Private Sub optTodosEstatus_Click()
Dim i As Integer

lstEstatus.Enabled = False

If optTodosEstatus.Value Then
   For i = 0 To lstEstatus.ListCount - 1
       lstEstatus.Selected(i) = False
    Next
End If
End Sub

Private Sub optTodosSolicitantes_Click()

Dim i As Integer

lstSolicitantes.Enabled = False

If optTodosSolicitantes.Value Then
   For i = 0 To lstSolicitantes.ListCount - 1
       lstSolicitantes.Selected(i) = False
    Next
End If

End Sub



VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form frmReporte 
   Caption         =   "Form1"
   ClientHeight    =   6984
   ClientLeft      =   11376
   ClientTop       =   1152
   ClientWidth     =   6552
   LinkTopic       =   "Form1"
   ScaleHeight     =   6984
   ScaleWidth      =   6552
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      lastProp        =   600
      _cx             =   10231
      _cy             =   12347
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

'cgml 20160123
Private crApp As New CRAXDDRT.Application 'CRAXDRT.Application 'Objeto que representa una instancia del programa Crystal Reports.
Private crReport As New CRAXDDRT.Report 'CRAXDRT.Report 'Objeto que representa el reporte que deseamos abrir.
 
Private mflgContinuar As Boolean 'Variable booleana que nos indica si hubo error al tratar de abrir el archivo RPT.
Private mstrParametro1 As String 'Variable de cadena que almacenará el valor que se le pasará al Parametro1 del reporte.
Private mlngParametro2 As Long 'Variable numérica que almacenará el valor que se le pasará al Parametro2 del reporte.

Public mstrNombreArchivo As String
Public mstrSQL As String
Private Sub Form_Load()
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
 
    On Error GoTo ErrHandler
 
    'Abrir el reporte
    Screen.MousePointer = vbHourglass
   
    mflgContinuar = True
    Set crReport = crApp.OpenReport(mstrNombreArchivo, 1)
 
    ' Parametros del reporte
    Set crParamDefs = crReport.ParameterFields
 
    For Each crParamDef In crParamDefs
        Select Case crParamDef.ParameterFieldName
            Case "Parametro1"
                crParamDef.AddCurrentValue (mstrParametro1)
       
            Case "Parametro2"
                crParamDef.AddCurrentValue (mlngParametro2)
               
        End Select
 
    Next
 
    'CRViewer1
    If Len(mstrSQL) > 0 Then crReport.SQLQueryString = mstrSQL
    CRViewer1.ReportSource = crReport
    CRViewer1.DisplayGroupTree = False
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
 
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
    Exit Sub
 
ErrHandler:
    If Err.Number = -2147206461 Then
        MsgBox "El archivo de reporte no se encuentra, restáurelo de los discos de instalación", _
            vbCritical + vbOKOnly
    Else
        MsgBox Err.Description, vbCritical + vbOKOnly
    End If
 
    mflgContinuar = False
    Screen.MousePointer = vbDefault
   
End Sub
Private Sub Form_Activate()
    If Not mflgContinuar Then Unload Me
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set crReport = Nothing
    Set crApp = Nothing
   
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth

End Sub
Public Sub PasarParametros(sParam1 As String, lParam2 As Long)
    mstrParametro1 = sParam1
    mlngParametro2 = lParam2
   
End Sub

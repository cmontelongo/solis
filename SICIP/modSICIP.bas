Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text
Option Base 1

Global gstrCveCotizacion As String
Global glngCveCotizacion As Long
Global glngCveRequisicion As Long
Global glngCveOT As Long
Global gstrNombreEmpresa As String


Public Function SinComas(ByVal vstrValor As String) As String

SinComas = Replace(vstrValor, ",", "")
End Function
Public Sub CargaParametros()
'************************************************************************
' Rutina que realiza la carga de parámetros generales del sistema
' a variables globales.
'************************************************************************

On Error GoTo Err_CargaParametros

Dim strmsg As String

' Lee los parametros del archivo .ini
gstrDirectorioRpt = BuscaParametrosIni("Datos Generales", "DirReportes")
gstrNombreEmpresa = BuscaParametrosIni("Datos Generales", "NombreEmpresa")
gstrServidorCentral = BuscaParametrosIni("Datos Generales", "Servidor")



Exit Sub


Err_CargaParametros:
   strmsg = "Ocurrió un error al leer los parámetros" & Chr$(13)
   strmsg = strmsg & "de inicio del sistema. La ejecución se detendrá."
   MsgBox strmsg
   CierraConeccion
   End

End Sub


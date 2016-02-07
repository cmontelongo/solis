Option Explicit On
Option Compare Text

Imports System.Runtime.InteropServices
Imports System.Text

Public Module modConecta

    Public gcn As New ADODB.Connection()
    Public gstrLogin As String              ' Login para entrar al Servidor
    Public gstrPassword As String           ' Password para el Login
    Public gstrServidor As String           ' Servidor de SQL
    Public gstrBaseDeDatos As String        ' Nombre de la Base de Datos de SQL

    Dim mstrDirectorioWindows As String

    ' Para ubicar los INI's
    'Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    '   "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    '                                ByVal lpKeyName As String, _
    '                                ByVal lpDefault As String, _
    '                                ByVal lpReturnerstring As String, _
    '                                ByVal nSize As Integer, _
    '                                ByVal lpFileName As String) As Integer

    Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String, _
               ByVal lpKeyName As String, _
               ByVal lpDefault As String, _
               ByVal lpReturnedString As StringBuilder, _
               ByVal nSize As Integer, _
               ByVal lpFileName As String) As Integer

    Public Sub AbreConeccion()
        '----------------------------------------------------'
        '     Abre la coneccion con la base de datos         '
        '----------------------------------------------------'

        ' Busca los Parametros Iniciales del TRANS.INI
        'mstrDirectorioWindows = BuscaDirectorioWindows()

        Try
            'Abre Conección
            gcn.ConnectionString = "provider=sqloledb;" & _
                        "server=" & gstrServidor & ";" & _
                        "database=SICIP;" & _
                        "uid=" & gstrLogin & ";" & _
                        "Password=" & gstrPassword
            gcn.Open()

            ' Obtiene el nombre de la computadora
            'gstrComputerName = String(20, Chr(0)) ' Asigna 20 Espacios a la variable
            'If GetComputerName(gstrComputerName, 256) Then
            ' gstrComputerName = UCase(Mid$(gstrComputerName, 1, InStr(gstrComputerName, Chr(0)) - 1))
            ' gstrComputerName = TextoSinEspacio(gstrComputerName)
            ' Else
            ' gstrComputerName = "NoTiene"
            ' End If

        Catch ex As Exception

            Dim errLoop As ADODB.Error
            Dim errs1 As ADODB.Errors
            Dim strmsg As String
            Dim lngIndice As Long
            lngIndice = 1

            ' Proceso
            strmsg = "VB Error # " & Str(Err.Number)
            strmsg = strmsg & vbCrLf & "   Generado por " & Err.Source
            strmsg = strmsg & vbCrLf & "   Descripcion  " & Err.Description

            ' Enumera la coleccion de errores y despliega las propiedades de cada uno de los errores.
            Errs1 = gcn.Errors
            For Each errLoop In Errs1
                With errLoop
                    strmsg = strmsg & vbCrLf & "Error #" & lngIndice & ":"
                    strmsg = strmsg & vbCrLf & "   ADO Error   #" & .Number
                    strmsg = strmsg & vbCrLf & "   Descripcion  " & .Description
                    strmsg = strmsg & vbCrLf & "   Fuente       " & .Source
                    lngIndice = lngIndice + 1
                End With
            Next

            MsgBox(strmsg, MsgBoxStyle.Critical, "Abre Coneccion")

            'Si se efectuo la apertura de sql lo cierra
            If gcn.State = ADODB.ObjectStateEnum.adStateOpen Then gcn.Close()

            gcn = Nothing

            'Termina la aplicacion que halla solicitado la conexion
            End

        End Try

    End Sub
    Public Function BuscaParametrosIni(ByVal Encabezado As String, ByVal Variable As String) As String
        '**********************
        'Descripcion : Lee el parametro solicitado por la VARIABLE en el .INI
        '   El GetPrivateProfileString da como salida la longitud del parametro solicitado
        ' INPUT PARAMETERS:
        '   Encabezado: Nombre del encabezado en las secciones del INI
        '   Variable: Nombre del parametro o de la Llave
        '
        '**********************

        Dim strArchivo As String
        Dim intResultado As Integer
        Dim sb As StringBuilder

        strArchivo = gstrArchivoIni
        If Dir(strArchivo) = "" Then
            strArchivo = "E" & Mid(gstrArchivoIni, 2, 30)
        End If

        sb = New StringBuilder(300)
        intResultado = GetPrivateProfileString(Encabezado, Variable, "", sb, sb.Capacity, strArchivo)
        BuscaParametrosIni = sb.ToString()

        If BuscaParametrosIni = "" Then
            MsgBox("Error al buscar en Archivo INI, no se encontró " & Variable, MsgBoxStyle.Critical, "Parametros Ini")
            End
        End If

    End Function

End Module

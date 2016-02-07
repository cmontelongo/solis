Option Explicit On
Option Compare Text



Public Module modConecta

    Public gcn As New ADODB.Connection()

    Dim mstrDirectorioWindows As String

    Public Sub AbreConeccion()
        '----------------------------------------------------'
        '     Abre la coneccion con la base de datos         '
        '----------------------------------------------------'

        'On Error GoTo blnAbreConeccion_Err


        'Dim strError As String          ' Ultimo Mensaje de Errores de Rdo's
        'Dim intIndice As Integer        ' En vez de utilizar I para los For-next
        'Dim lngLongitudArchivo As Long  ' Longitud del archivo para checar si esta en RED


        ' Busca los Parametros Iniciales del TRANS.INI
        'mstrDirectorioWindows = BuscaDirectorioWindows()

        ' Inicializa el Environment (Ambiente de trabajo)
        'rdoEngine.rdoDefaultCursorDriver = rdUseOdbc
        'gen = rdoEngine.rdoCreateEnvironment("app", "app", "app")
        'mblnEnvOpen = True      'Se efectuo la apertura del ambiente

        'Abre Conección

        gcn.ConnectionString = "provider=sqloledb;server=NAUTILIUS;database=SICIP;uid=SICIP;Password=SICIP"
        gcn.Open()

        ' Obtiene el nombre de la computadora
        'gstrComputerName = String(20, Chr(0)) ' Asigna 20 Espacios a la variable
        'If GetComputerName(gstrComputerName, 256) Then
        ' gstrComputerName = UCase(Mid$(gstrComputerName, 1, InStr(gstrComputerName, Chr(0)) - 1))
        ' gstrComputerName = TextoSinEspacio(gstrComputerName)
        ' Else
        ' gstrComputerName = "NoTiene"
        ' End If

        'GoTo blnAbreConeccion_Exit

        'blnAbreConeccion_Err:
        '        Dim lngIndice As Long
        '       Dim strmsg As String

        '      Select Case Err()
        '         Case 53, 68 'File not found, Device unavailable
        '    strmsg = "Necesitas primeramente entrar a la Red"

        '       Case 40002
        '  For lngIndice = 0 To rdoErrors.Count - 1
        'strmsg = strmsg & rdoErrors(lngIndice).Description & vbKeyReturn
        'Next lngIndice
        '    Case Else
        'strmsg = Err & " " & Error
        'End Select
        'Err.Clear()
        'MsgBox("Error al Conectarse al Servidor de SQL." & Chr$(13) & _
        '    strmsg, vbExclamation, "Abre Coneccion")

        'If mblnEnvOpen Then gen.Close() 'Si se efectuo la apertura del ambiente la cierra
        ' If mblncn Then gcn.Close() 'Si se efectuo la apertura de sql lo cierra

        'Termina la aplicacion que halla solicitado la conexion
        'End

        'blnAbreConeccion_Exit:

    End Sub

End Module

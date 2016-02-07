VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Clave de Acceso"
   ClientHeight    =   1656
   ClientLeft      =   1008
   ClientTop       =   3960
   ClientWidth     =   3072
   HelpContextID   =   10
   Icon            =   "SI040.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1656
   ScaleWidth      =   3072
   Tag             =   "1"
   Begin VB.ComboBox cboBase 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtCuenta 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      WhatsThisHelpID =   190
      Width           =   1635
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      WhatsThisHelpID =   190
      Width           =   1635
   End
   Begin VB.Label lblServidor 
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1230
      WhatsThisHelpID =   190
      Width           =   765
   End
   Begin VB.Label lblCuenta 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   300
      WhatsThisHelpID =   190
      Width           =   645
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      WhatsThisHelpID =   190
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Const APLICACION = 3

Dim blnPermiso As Boolean
Private Sub cboBase_Click()

gintCveBase = cboBase.ItemData(cboBase.ListIndex)

'------------------------------------------------------
'    Obtiene parámetros de configuración globales
'------------------------------------------------------
CargaParametrosConfiguracion

'---------------------------------
'   Inicia la forma Principal
'---------------------------------
Me.Hide
frmValeHerramienta.Show

End Sub
Public Sub Form_Load()

CentrarForma Me
'txtServidor =
'Se asignan Variables de Cuenta y Password
gstrLogin = "SICIP"
gstrPassword = "MTA1971"
'gstrServidor = "NAUTILIUS"
gstrBaseDeDatos = "SICIP"
gstrServidor = BuscaParametrosIni("Datos Generales", "Servidor")

CargaParametros
AbreConeccion

'---------------------------------
'   Inicia la forma Principal
'---------------------------------
Me.Hide
frmValeHerramienta.Show

End Sub
Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtPassword.SetFocus
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)

Dim strSQL As String
Dim rsPermiso As rdoResultset
Dim rsPassword As rdoResultset

If KeyAscii = vbKeyReturn Then
'    txtServidor.SetFocus

    gstrLogin = UCase(txtCuenta.Text)
    
    ' Verifica si tiene acceso a este modulo
    strSQL = "select * from Usuario where CveUsuario = '" & gstrLogin & "'"
    Set rsPassword = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    If rsPassword.EOF Then
        MsgBox " Cuenta no existe "
        rsPassword.Close
        End
    End If
    If UCase(Trim(rsPassword!PASSWORD)) <> UCase(Trim(txtPassword.Text)) Then
        MsgBox " Password es incorrecto "
        rsPassword.Close
        End
    End If

    strSQL = "select * from UsuarioAplicacion where CveUsuario = '" & gstrLogin
    strSQL = strSQL & "' and CveAplicacion = " & APLICACION
    Set rsPermiso = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    If rsPermiso.EOF Then
        MsgBox "No se tiene acceso a este Módulo de SIM"
        rsPermiso.Close
        End
    End If
    rsPermiso.Close
    
    LlenaVariosSelectores "SELECT B.CveBase,B.Nombre FROM Base B, UsuarioBase UB " & _
                        "WHERE B.CveBase = UB.CveBase" & _
                        "  AND UB.CveUsuario = '" & gstrLogin & "' " & _
                        "ORDER BY B.Nombre", Array("cboBase"), Me
    cboBase.SetFocus
    
End If

End Sub
Private Sub txtPassword_LostFocus()
Dim strSQL As String
Dim rsPermiso As rdoResultset
Dim rsPassword As rdoResultset

gstrLogin = UCase(txtCuenta.Text)

' Verifica si tiene acceso a este modulo
strSQL = "select * from Usuario where CveUsuario = '" & gstrLogin & "'"
Set rsPassword = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
If rsPassword.EOF Then
    MsgBox " Cuenta no existe "
    rsPassword.Close
    End
End If
If UCase(Trim(rsPassword!PASSWORD)) <> UCase(Trim(txtPassword.Text)) Then
    MsgBox " Password es incorrecto "
    rsPassword.Close
    txtPassword.SetFocus
    Exit Sub
End If

strSQL = "select * from UsuarioAplicacion where CveUsuario = '" & gstrLogin
strSQL = strSQL & "' and CveAplicacion = " & APLICACION
Set rsPermiso = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
If rsPermiso.EOF Then
    MsgBox "No se tiene acceso a este Módulo de SIM"
    rsPermiso.Close
    End
End If
rsPermiso.Close

LlenaVariosSelectores "SELECT B.CveBase,B.Nombre FROM Base B, UsuarioBase UB " & _
                    "WHERE B.CveBase = UB.CveBase" & _
                    "  AND UB.CveUsuario = '" & gstrLogin & "' " & _
                    "ORDER BY B.Nombre", Array("cboBase"), Me
cboBase.SetFocus

End Sub

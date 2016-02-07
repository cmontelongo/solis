Option Strict Off
Option Explicit On
Module modAPIs
	
	
	' Declaraciones del Api
	Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Integer, ByVal nIDEvent As Integer) As Integer
	
	Private Declare Function SendMessageLongRef Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer
	
	Private Declare Function FindWindow Lib "user32"  Alias "FindWindowA"(ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
	
	Private Declare Function FindWindowEx Lib "user32"  Alias "FindWindowExA"(ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
	
	Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Integer, ByVal nIDEvent As Integer, ByVal uElapse As Integer, ByVal lpTimerFunc As Integer) As Integer
	
	
	Private m_ASC As Integer
	Sub inputbox_Password(ByRef El_Form As System.Windows.Forms.Form, ByRef Caracter As String)
		
		m_ASC = Asc(Caracter)
		
		'UPGRADE_WARNING: Add a delegate for AddressOf TimerProc Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
		Call SetTimer(El_Form.Handle.ToInt32, &H5000, 100, AddressOf TimerProc)
		
	End Sub
	Private Sub TimerProc(ByVal hwnd As Integer, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTime As Integer)
		
		Dim Handle_InputBox As Integer
		
		'Captura el handle del textBox del InputBox
		Handle_InputBox = FindWindowEx(FindWindow("#32770", My.Application.Info.Title), 0, "Edit", "")
		
		'Le establece el PasswordChar
		Call SendMessageLongRef(Handle_InputBox, &HCC, m_ASC, 0)
		'Finaliza el Timer
		Call KillTimer(hwnd, idEvent)
		
	End Sub
End Module
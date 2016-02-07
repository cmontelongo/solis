Option Strict Off
Option Explicit On
Module ContextIDs
	'=====================================================================
	'=====================================================================
	'
	'This source code contains the following routines:
	'  o SetAppHelp() 'Called in the main Form_Load event to register your
	'                 'program with WINHELP.EXE
	'  o QuitHelp()    'Deregisters your program with WINHELP.EXE. Should
	'                  'be called in your main Form_Unload event
	'  o ShowHelpTopic(Topicnum) 'Brings up context sensitive help based on
	'                  'any of the following CONTEXT Irs
	'  o ShowContents  'Displays the startup topic
	'  o HelpWindowSize(x,y,dx,dy) ' Position help window in a screen
	'                              ' independent manner
	'  o SearchHelp()  'Brings up the windows help KEYWORD SEARCH dialog box
	'***********************************************************************
	'
	'=====================================================================
	'List of Context Irs for <ODT>
	'=====================================================================
	Public Const Hlp_ODT As Short = 10 'Main Help Window
	Public Const Hlp_Agregar_Tareas1 As Short = 40 'Main Help Window
	Public Const Hlp_Asignar_Mxcanicos As Short = 50 'Main Help Window
	Public Const Hlp_Abrir_una As Short = 60 'Main Help Window
	Public Const Hlp_Consulta_de As Short = 70 'Main Help Window
	Public Const Hlp_Consulta_de1 As Short = 80 'Main Help Window
	Public Const Hlp_Cerrar_la As Short = 90 'Main Help Window
	Public Const Hlp_Reportes As Short = 100 'Main Help Window
	Public Const Hlp_Agregar_Tareas2 As Short = 110 'Main Help Window
	Public Const Hlp_Salir As Short = 120 'Main Help Window
	Public Const GLOS_ODT As Short = 130
	Public Const GLOS_Toolbar As Short = 140
	Public Const GLOS_Tarea As Short = 150
	Public Const GLOS_Filtrar As Short = 160
	Public Const GLOS_Razon_de_Reparacixxf3n As Short = 170
	'=====================================================================
	'
	'
	'  Help engine section.
	
	' Commanrs to pass WinHelp()
	Public Const HELP_CONTEXT As Integer = &H1 '  Display topic in ulTopic
	Public Const HELP_QUIT As Integer = &H2 '  Terminate help
	Public Const HELP_FINDER As Integer = &HB '  Display Contents tab
	Public Const HELP_INDEX As Integer = &H3 '  Display index
	Public Const HELP_HELPONHELP As Integer = &H4 '  Display help on using help
	Public Const HELP_SETINDEX As Integer = &H5 '  Set the current Index for multi index help
	Public Const HELP_KEY As Integer = &H101 '  Display topic for keyword in offabData
	Public Const HELP_MULTIKEY As Integer = &H201
	Public Const HELP_CONTENTS As Integer = &H3 ' Display Help for a particular topic
	Public Const HELP_SETCONTENTS As Integer = &H5 ' Display Help contents topic
	Public Const HELP_CONTEXTPOPUP As Integer = &H8 ' Display Help topic in popup window
	Public Const HELP_FORCEFILE As Integer = &H9 ' Ensure correct Help file is displayed
	Public Const HELP_COMMAND As Integer = &H102 ' Execute Help macro
	Public Const HELP_PARTIALKEY As Integer = &H105 ' Display topic found in keyword list
	Public Const HELP_SETWINPOS As Integer = &H203 ' Display and position Help window
	
	
	Structure HELPWININFO
		Dim wStructSize As Integer
		Dim X As Integer
		Dim Y As Integer
		Dim dX As Integer
		Dim dY As Integer
		Dim wMax As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(2),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=2)> Public rgChMember() As Char
	End Structure
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function WinHelp Lib "User32.dll"  Alias "WinHelpA"(ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Any) As Integer
	'UPGRADE_WARNING: Structure HELPWININFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Declare Function WinHelpByInfo Lib "User32.dll"  Alias "WinHelpA"(ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByRef dwData As HELPWININFO) As Integer
	Declare Function WinHelpByStr Lib "User32.dll"  Alias "WinHelpA"(ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As String) As Integer
	Declare Function WinHelpByNum Lib "User32.dll"  Alias "WinHelpA"(ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Integer) As Integer
	Dim m_hWndMainWindow As Integer ' hWnd to tell WINHELP the helpfile owner
	
	
	Dim MainWindowInfo As HELPWININFO
	Sub SetAppHelp(ByVal hWndMainWindow As Object)
		'=====================================================================
		'To use these subroutines to access WINHELP, you need to add
		'at least this one subroutine call to your code
		'     o  In the Form_Load event of your main Form enter:
		'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
		'         (If you are not interested in keyword searching or context
		'         sensitive help, this is the only call you need to make!)
		'=====================================================================
		'UPGRADE_WARNING: Couldn't resolve default property of object hWndMainWindow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_hWndMainWindow = hWndMainWindow
		If Right(Trim(My.Application.Info.DirectoryPath), 1) = "\" Then
			'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			App.HelpFile = My.Application.Info.DirectoryPath & "ODT.HLP"
		Else
			'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			App.HelpFile = My.Application.Info.DirectoryPath & "\ODT.HLP"
		End If
#If Win32 Then
		MainWindowInfo.wStructSize = 26
#Else
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression Else did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		MainWindowInfo.wStructSize = 14
#End If
		MainWindowInfo.X = 256
		MainWindowInfo.Y = 256
		MainWindowInfo.dX = 512
		MainWindowInfo.dY = 512
		MainWindowInfo.rgChMember = Chr(0) & Chr(0)
	End Sub
	Sub QuitHelp()
		Dim Result As Object
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_QUIT, Chr(0) & Chr(0) & Chr(0) & Chr(0))
	End Sub
	Sub ShowHelpTopic(ByVal ContextID As Integer)
		'=====================================================================
		'  FOR CONTEXT SENSITIVE HELP IN RESPONSE TO A COMMAND BUTTON ...
		'=====================================================================
		'     o   For 'Help button' controls, you can call:
		'         Call ShowHelpTopic(<any Hlpxxx entry above>)
		'=====================================================================
		'  TO ADD FORM LEVEL CONTEXT SENSITIVE HELP...
		'=====================================================================
		'     o  For FORM level context sensetive help, you should set each
		'        Me.HelpContext=<any Hlp_xxx entry above>
		'
		Dim Result As Object
		
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXT, CInt(ContextID))
		
	End Sub
	Sub ShowHelpTopic2(ByVal ContextID As Integer)
		'=====================================================================
		'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 2 ...
		'=====================================================================
		'     o   For 'Help button' controls, you can call:
		'         Call ShowHelpTopic2(<any Hlpxxx entry above>)
		'
		Dim Result As Object
		
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile & ">HlpWnd02", HELP_CONTEXT, CInt(ContextID))
		
	End Sub
	Sub ShowHelpTopic3(ByVal ContextID As Integer)
		'=====================================================================
		'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 3 ...
		'=====================================================================
		'     o   For 'Help button' controls, you can call:
		'         Call ShowHelpTopic3(<any Hlpxxx entry above>)
		'
		Dim Result As Object
		
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile & ">HlpWnd03", HELP_CONTEXT, CInt(ContextID))
		
	End Sub
	Sub ShowGlossary()
		Dim Result As Object
		
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXT, CInt(64000))
		
	End Sub
	Sub ShowPopupHelp(ByVal ContextID As Integer)
		'=====================================================================
		'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
		'=====================================================================
		Dim Result As Object
		
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTEXTPOPUP, CInt(ContextID))
		
	End Sub
	Sub DoHelpMacro(ByVal MacroString As String)
		'=====================================================================
		'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
		'=====================================================================
		Dim Result As Object
		
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelpByStr(m_hWndMainWindow, App.HelpFile, HELP_COMMAND, (MacroString))
		
	End Sub
	Sub ShowHelpContents()
		'=====================================================================
		'  DISPLAY HELP STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
		'=====================================================================
		'
		Dim Result As Object
		
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTENTS, CInt(0))
		
	End Sub
	Sub ShowContentsTab()
		'=====================================================================
		'  DISPLAY Contents tab (*.CNT)
		'=====================================================================
		'
		Dim Result As Object
		
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_FINDER, CInt(0))
		
	End Sub
	Sub ShowHelpOnHelp()
		'=====================================================================
		'  DISPLAY HELP for WINHELP.EXE  ...
		'=====================================================================
		'
		Dim Result As Object
		
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_HELPONHELP, CInt(0))
		
	End Sub
	
	Sub SearchHelp()
		'=====================================================================
		'  TO ADD KEYWORD SEARCH CAPABILITY...
		'=====================================================================
		'     o   In your Help|Search menu selection, simply enter:
		'         Call SearchHelp() 'To invoke helpfile keyword search dialog
		'
		Dim Result As Object
		
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_PARTIALKEY, "")
		
	End Sub
	
	Sub SearchHelpKeyWord(ByRef Argument As String)
		'=====================================================================
		'  TO ADD KEYWORD SEARCH CAPABILITY...
		'=====================================================================
		'     o   In your Help|Search menu selection, simply enter:
		'         Call SearchHelp() 'To invoke helpfile keyword search dialog
		'
		Dim Result As Object
		
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_PARTIALKEY, Trim(Argument))
		
	End Sub
	Sub HelpWindowSize(ByRef X As Short, ByRef Y As Short, ByRef wx As Short, ByRef wy As Short)
		'=====================================================================
		'  TO SET THE SIZE AND POSITION OF THE MAIN HELP WINDOW...
		'=====================================================================
		'     o   Call HelpWindowSize(x, y, dx, dy), where:
		'             x = 1-1024 (position from left edge of screen)
		'             y = 1-1024 (position from top of screen)
		'             dx= 1-1024 (width)
		'             dy= 1-1024 (height)
		'
		Dim Result As Object
		MainWindowInfo.X = X
		MainWindowInfo.Y = Y
		MainWindowInfo.dX = wx
		MainWindowInfo.dY = wy
		'UPGRADE_ISSUE: App property App.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Result = WinHelpByInfo(m_hWndMainWindow, App.HelpFile, HELP_SETWINPOS, MainWindowInfo)
	End Sub
End Module
VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "Tab32x20.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmOT 
   Caption         =   "Módulo de Proyectos"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17385
   HelpContextID   =   10
   Icon            =   "SI002.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9615
   ScaleWidth      =   17385
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tlbODT 
      Height          =   420
      Left            =   60
      TabIndex        =   39
      Top             =   60
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgIconos"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   22
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Refrescar"
            Object.ToolTipText     =   "Refrescar"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   240
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Agregar"
            Object.ToolTipText     =   "Agregar Cotizacion"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Guardar Cotizacion"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar Cotizacion"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   250
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar Modificaciones"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   250
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Autoriza"
            Object.ToolTipText     =   "Solicitar Autorizacion"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   100
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Envio"
            Object.ToolTipText     =   "Enviar a Cliente"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Recibe"
            Object.ToolTipText     =   "Autorizacion del Cliente"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "OT"
            Object.ToolTipText     =   "Generar OT"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Compra"
            Object.ToolTipText     =   "Solicitar Requisicion"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   350
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   200
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TabproLib.vaTabPro TabPrincipal 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   17235
      _Version        =   131072
      _ExtentX        =   30401
      _ExtentY        =   15901
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      ForeColor       =   16777215
      TabHeight       =   500
      TabCount        =   2
      AlignTextH      =   1
      AlignTextV      =   1
      ThreeD          =   -1  'True
      ShowFocusRect   =   0   'False
      MarginLeft      =   150
      MarginRight     =   150
      ApplyTo         =   2
      GrayAreaColor   =   -2147483633
      TabSeparator    =   6
      OffsetFromClientTop=   -1  'True
      ShowEarMark     =   -1  'True
      BookShowMetalSpine=   -1  'True
      BookRingShowHole=   -1  'True
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   405
      MouseIcon       =   "SI002.frx":030A
      ThreeDOuterWidthActive=   2
      DrawFocusRect   =   1
      TabCaption      =   "SI002.frx":0326
      Begin VB.Frame fraPartidas 
         Caption         =   "Partidas"
         Height          =   4455
         Left            =   3600
         TabIndex        =   54
         Top             =   4320
         Width           =   10575
         Begin VB.CommandButton Command1 
            Caption         =   "+"
            Height          =   375
            Left            =   10080
            TabIndex        =   57
            Top             =   600
            Width           =   375
         End
         Begin FPSpread.vaSpread sprPartidas 
            Height          =   3735
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   9855
            _Version        =   393216
            _ExtentX        =   17383
            _ExtentY        =   6588
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
            SpreadDesigner  =   "SI002.frx":057D
         End
      End
      Begin VB.Frame frmCotizacion 
         Caption         =   "Cotizacion"
         Height          =   3615
         Left            =   3600
         TabIndex        =   46
         Top             =   720
         Width           =   13455
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   315
            Left            =   6840
            TabIndex        =   80
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtRutaAcceso 
            Height          =   315
            Left            =   1680
            TabIndex        =   78
            Top             =   2880
            Width           =   3855
         End
         Begin VB.ComboBox cboCotizacionTipo 
            Height          =   315
            Left            =   1560
            TabIndex        =   75
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Height          =   555
            Left            =   5880
            Picture         =   "SI002.frx":0759
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   840
            Width           =   555
         End
         Begin VB.ComboBox cboObra 
            Height          =   315
            Left            =   1560
            TabIndex        =   72
            Top             =   960
            Width           =   4215
         End
         Begin VB.TextBox txtUsuario 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7440
            TabIndex        =   71
            Text            =   "AIBARRA"
            Top             =   3000
            Width           =   1095
         End
         Begin VB.ComboBox cboFormaPago 
            Height          =   315
            Left            =   1680
            TabIndex        =   69
            Top             =   2520
            Width           =   3855
         End
         Begin VB.ComboBox cboTiempoEntrega 
            Height          =   315
            Left            =   1680
            TabIndex        =   67
            Top             =   2160
            Width           =   3855
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1680
            TabIndex        =   65
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox txtNombreRepresentante 
            Height          =   315
            Left            =   1560
            TabIndex        =   63
            Top             =   600
            Width           =   3855
         End
         Begin VB.ComboBox cboCotizacionEstatus 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4200
            TabIndex        =   61
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtFechaRecepcion 
            Height          =   315
            Left            =   7440
            TabIndex        =   53
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox txtFechaEnvio 
            Height          =   315
            Left            =   7440
            TabIndex        =   52
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox txtFechaCotizacion 
            Height          =   315
            Left            =   7440
            TabIndex        =   51
            Top             =   1680
            Width           =   1095
         End
         Begin VB.ComboBox cboCliente 
            Height          =   315
            Left            =   1560
            TabIndex        =   49
            Top             =   240
            Width           =   5175
         End
         Begin VB.TextBox txtNumCotizacion 
            Height          =   315
            Left            =   4200
            TabIndex        =   48
            Top             =   1320
            Width           =   1575
         End
         Begin FPSpread.vaSpread sprAreas 
            Height          =   3255
            Left            =   9000
            TabIndex        =   77
            Top             =   240
            Width           =   4335
            _Version        =   393216
            _ExtentX        =   7646
            _ExtentY        =   5741
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
            MaxCols         =   2
            RowHeaderDisplay=   0
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "SI002.frx":1023
            VisibleCols     =   2
            ScrollBarTrack  =   3
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Directorio:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   79
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Tipo Cotizacion:"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   76
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Obra:"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   73
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Forma de Pago:"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   70
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Tiempo de Entrega:"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   68
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Moneda:"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   66
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Representante:"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   64
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Estado:"
            Height          =   255
            Index           =   7
            Left            =   3360
            TabIndex        =   62
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fecha Recepcion:"
            Height          =   255
            Index           =   6
            Left            =   6000
            TabIndex        =   60
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fecha envio:"
            Height          =   255
            Index           =   5
            Left            =   6000
            TabIndex        =   59
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fecha:"
            Height          =   255
            Index           =   4
            Left            =   6000
            TabIndex        =   58
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Usuario:"
            Height          =   255
            Index           =   3
            Left            =   6000
            TabIndex        =   56
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Cliente:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Cotizacion:"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   47
            Top             =   1320
            Width           =   1095
         End
      End
      Begin ComctlLib.TreeView treeview1 
         Height          =   7875
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   13891
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "iml16"
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame fraTareas 
         Caption         =   "Tareas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   60
         TabIndex        =   37
         Top             =   3720
         Visible         =   0   'False
         Width           =   16965
         Begin FPSpread.vaSpread sprTareas 
            Height          =   4215
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   16815
            _Version        =   393216
            _ExtentX        =   29660
            _ExtentY        =   7435
            _StockProps     =   64
            AllowMultiBlocks=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   9
            RowHeaderDisplay=   0
            ScrollBarMaxAlign=   0   'False
            SelectBlockOptions=   2
            SpreadDesigner  =   "SI002.frx":1286
            UserResize      =   0
            VisibleCols     =   9
            VisibleRows     =   500
         End
      End
      Begin VB.Frame fraImpresiones 
         Enabled         =   0   'False
         Height          =   7095
         Left            =   -25934
         TabIndex        =   4
         Top             =   -22934
         Width           =   8895
         Begin VB.OptionButton optReporte 
            Caption         =   "Reportes de Campañas"
            Height          =   195
            Index           =   35
            Left            =   4800
            TabIndex        =   44
            Top             =   6480
            Width           =   2775
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Reimpresión Reporte del Operador"
            Height          =   195
            Index           =   34
            Left            =   4800
            TabIndex        =   43
            Top             =   6120
            Width           =   2775
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Foseos y Fallas Reportadas"
            Height          =   195
            Index           =   33
            Left            =   4800
            TabIndex        =   42
            Top             =   5760
            Width           =   2775
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Tabulado de Unidades en Taller"
            Height          =   195
            Index           =   25
            Left            =   4800
            TabIndex        =   41
            Top             =   5400
            Width           =   2775
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Consulta Informacion"
            Height          =   195
            Index           =   32
            Left            =   4800
            TabIndex        =   40
            Top             =   5040
            Width           =   2535
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Rendimiento Vida Util"
            Height          =   195
            Index           =   31
            Left            =   4800
            TabIndex        =   36
            Top             =   4680
            Width           =   2535
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Reporte Diario de Operación"
            Height          =   195
            Index           =   30
            Left            =   4800
            TabIndex        =   35
            Top             =   4320
            Width           =   2535
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Horas Hombre x Unidad"
            Height          =   195
            Index           =   29
            Left            =   4800
            TabIndex        =   34
            Top             =   3960
            Width           =   2535
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "ODT´s de Taller Externo"
            Height          =   195
            Index           =   28
            Left            =   4800
            TabIndex        =   33
            Top             =   3240
            Width           =   2415
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Mantenimientos Vencidos"
            Height          =   255
            Index           =   27
            Left            =   960
            TabIndex        =   32
            Top             =   2520
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Tabulado de Refacciones por Unidad"
            Height          =   255
            Index           =   26
            Left            =   960
            TabIndex        =   31
            Top             =   4680
            Width           =   3375
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Kms. por Operador"
            Height          =   195
            Index           =   24
            Left            =   960
            TabIndex        =   30
            Top             =   6480
            Width           =   2655
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Fallas en Camino"
            Height          =   195
            Index           =   23
            Left            =   4800
            TabIndex        =   29
            Top             =   2880
            Width           =   1815
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Kardex de Foseo por Unidad"
            Height          =   255
            Index           =   22
            Left            =   960
            TabIndex        =   28
            Top             =   2160
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Kardex por Tarea"
            Height          =   255
            Index           =   21
            Left            =   960
            TabIndex        =   27
            Top             =   1800
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Asignaciones por Mecánico"
            Height          =   195
            Index           =   20
            Left            =   4800
            TabIndex        =   25
            Top             =   2520
            Width           =   2775
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Rendimiento por Ruta"
            Height          =   195
            Index           =   19
            Left            =   4800
            TabIndex        =   24
            Top             =   2160
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Indicadores"
            Height          =   195
            Index           =   18
            Left            =   4800
            TabIndex        =   23
            Top             =   3600
            Width           =   1815
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Unidades en Operación"
            Height          =   195
            Index           =   17
            Left            =   4800
            TabIndex        =   22
            Top             =   1800
            Width           =   3615
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Costo de Refacciones / Km."
            Height          =   255
            Index           =   16
            Left            =   960
            TabIndex        =   21
            Top             =   5040
            Width           =   3615
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Consumo de Refacciones por Flotilla"
            Height          =   255
            Index           =   15
            Left            =   960
            TabIndex        =   20
            Top             =   4320
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Informacion de ODT's"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   19
            Top             =   360
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Bitácora de Llegadas"
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   18
            Top             =   720
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Rendimiento de Mecánicos"
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   17
            Top             =   1080
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Pronóstico de Mantenimientos"
            Height          =   255
            Index           =   5
            Left            =   960
            TabIndex        =   16
            Top             =   2880
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Kardex por Unidad"
            Height          =   255
            Index           =   4
            Left            =   960
            TabIndex        =   15
            Top             =   1440
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Rendimiento por Unidad"
            Height          =   195
            Index           =   6
            Left            =   960
            TabIndex        =   14
            Top             =   5760
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Estadística de Consumos por Refacción"
            Height          =   255
            Index           =   7
            Left            =   960
            TabIndex        =   13
            Top             =   3600
            Width           =   3495
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Consumo de Refacciones por Unidad"
            Height          =   255
            Index           =   8
            Left            =   960
            TabIndex        =   12
            Top             =   3960
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Catálogo de Refacciones"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   11
            Top             =   5400
            Width           =   3015
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Kms. Recorridos por Unidad"
            Height          =   195
            Index           =   9
            Left            =   960
            TabIndex        =   10
            Top             =   6120
            Width           =   2655
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Eficiencias por Area y por Tarea"
            Height          =   195
            Index           =   10
            Left            =   4800
            TabIndex        =   9
            Top             =   360
            Width           =   2655
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Tareas Realizadas por Razon"
            Height          =   195
            Index           =   11
            Left            =   4800
            TabIndex        =   8
            Top             =   720
            Width           =   2655
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Pronóstico de Consumo de Refacciones"
            Height          =   255
            Index           =   12
            Left            =   960
            TabIndex        =   7
            Top             =   3240
            Width           =   3495
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Causas de No Realización de Tareas"
            Height          =   195
            Index           =   13
            Left            =   4800
            TabIndex        =   6
            Top             =   1080
            Width           =   3615
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Disponibilidad de Cajones"
            Height          =   195
            Index           =   14
            Left            =   4800
            TabIndex        =   5
            Top             =   1440
            Width           =   3615
         End
      End
      Begin VB.Label lblLetrero 
         BackStyle       =   0  'Transparent
         Caption         =   "Información Adicional"
         Enabled         =   0   'False
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
         Left            =   -24209
         TabIndex        =   3
         Top             =   -16469
         Width           =   1845
      End
      Begin VB.Label lblLetrero 
         BackStyle       =   0  'Transparent
         Caption         =   "Estadísticas"
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   -20879
         TabIndex        =   2
         Top             =   -16469
         Width           =   1260
      End
      Begin VB.Label lblLetrero 
         BackStyle       =   0  'Transparent
         Caption         =   "Operación del Taller"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   -17894
         TabIndex        =   1
         Top             =   -16454
         Width           =   1845
      End
   End
   Begin ComctlLib.StatusBar staEstatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   26
      Top             =   9315
      Width           =   17385
      _ExtentX        =   30665
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Text            =   "SIM"
            TextSave        =   "SIM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   22437
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "28/10/2015"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "03:51 p.m."
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport rptReporte 
      Left            =   14520
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   7200
      WindowHeight    =   6900
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Image imgMas 
      Height          =   135
      Left            =   13680
      Picture         =   "SI002.frx":24D2
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgMenos 
      Height          =   135
      Left            =   13440
      Picture         =   "SI002.frx":2610
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin ComctlLib.ImageList imgIconos 
      Left            =   13800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   43
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":274E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":2A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":2D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":309C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":33B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":36D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":39EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":3D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":401E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":4338
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":4652
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":496C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":4C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":4FA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":52BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":55D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":58EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":5C08
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":5F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":623C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":6556
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":6870
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":6B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":6EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":71BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":74D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":77F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":7B0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":7E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":8140
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":845A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":8774
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":8A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":8DA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":90C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":93DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":96F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":9A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":9D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":A044
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":A35E
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":AB78
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI002.frx":AE92
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1


' Constantes del Tab
Const TABUNIDADES = 0
Const TABREPORTES = 1

' Constantes de las prioridades
Const PRIORIDADALTA = 1
Const PRIORIDADMEDIA = 2
Const PRIORIDADBAJA = 3

' Constantes del Spread
Const COLUMNAARTICULO = 1
Const COLUMNACANTIDAD = 2
Const COLUMNAUNIDADMEDIDA = 3
Const COLUMNAPRECIO = 4

'Orden de Reparación Unidades.
Dim rsODT  As rdoResultset

'Banderas de acción
Dim mblnAlta        As Boolean
Dim mblnLlena       As Boolean
Dim mblnEdicion     As Boolean

'Variables de uso general
Dim mdatFechaInicio As Date
Dim mdatFechaEquipo As Date
Dim mlngIdPCs As Long 'Identificador de los Pc's Opcionales
Dim mvntMarca         As Variant
Dim mintIndice As Integer

Dim mlngCveTarea As Long
Dim mblnCambioSprTareas As Boolean
Dim mdatUltimaHoraEjecucion As Date

Public blnReabrioODT As Boolean

Dim mstrArray(400, 3) As String










Private Sub ActualizaTree()
Dim rsConsulta As rdoResultset
Dim nodx As Node
Dim strSQL As String

treeview1.Nodes.Clear

Set nodx = treeview1.Nodes.Add(, , "CT", "Cotizaciones")
nodx.EnsureVisible

Set nodx = treeview1.Nodes.Add("CT", tvwChild, "PC", "Pendientes por Cotizar")

strSQL = "SELECT CveCotizacion,NumCotizacion FROM Cotizacion WHERE CveCotizacionEstatus = 1"
Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
Set nodx = treeview1.Nodes.Add("CT", tvwChild, "EP", "En Proceso (" & rsConsulta.RowCount & ")")
Do Until rsConsulta.EOF
    Set nodx = treeview1.Nodes.Add("EP", tvwChild, "C-" & CStr(rsConsulta!CveCotizacion), rsConsulta!NumCotizacion)
    nodx.EnsureVisible
    rsConsulta.MoveNext
Loop
rsConsulta.Close

strSQL = "SELECT CveCotizacion,NumCotizacion FROM Cotizacion WHERE CveCotizacionEstatus = 2"
Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
Set nodx = treeview1.Nodes.Add("CT", tvwChild, "PA", "Pendientes Por Autorizar (" & rsConsulta.RowCount & ")")
Do Until rsConsulta.EOF
    Set nodx = treeview1.Nodes.Add("PA", tvwChild, "C-" & CStr(rsConsulta!CveCotizacion), rsConsulta!NumCotizacion)
    rsConsulta.MoveNext
Loop
rsConsulta.Close

strSQL = "SELECT CveCotizacion,NumCotizacion FROM Cotizacion WHERE CveCotizacionEstatus = 3"
Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
Set nodx = treeview1.Nodes.Add("CT", tvwChild, "AU", "Autorizadas (" & rsConsulta.RowCount & ")")
Do Until rsConsulta.EOF
    Set nodx = treeview1.Nodes.Add("AU", tvwChild, "C-" & CStr(rsConsulta!CveCotizacion), rsConsulta!NumCotizacion)
    rsConsulta.MoveNext
Loop
rsConsulta.Close

strSQL = "SELECT CveCotizacion,NumCotizacion FROM Cotizacion WHERE CveCotizacionEstatus = 4"
Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
Set nodx = treeview1.Nodes.Add("CT", tvwChild, "EC", "Enviadas al Cliente (" & rsConsulta.RowCount & ")")
Do Until rsConsulta.EOF
    Set nodx = treeview1.Nodes.Add("EC", tvwChild, "C-" & CStr(rsConsulta!CveCotizacion), rsConsulta!NumCotizacion)
    rsConsulta.MoveNext
Loop
rsConsulta.Close

strSQL = "SELECT CveCotizacion,NumCotizacion FROM Cotizacion WHERE CveCotizacionEstatus = 5"
Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
Set nodx = treeview1.Nodes.Add("CT", tvwChild, "AC", "Autorizadas por el Cliente (" & rsConsulta.RowCount & ")")
Do Until rsConsulta.EOF
    Set nodx = treeview1.Nodes.Add("AC", tvwChild, "C-" & CStr(rsConsulta!CveCotizacion), rsConsulta!NumCotizacion)
    rsConsulta.MoveNext
Loop
rsConsulta.Close

Set nodx = treeview1.Nodes.Add(, , "OT", "Ordenes de Trabajo")

nodx.EnsureVisible





treeview1.Style = tvwTreelinesPlusMinusPictureText
End Sub

Public Sub DespliegaDetalle()
'---------------------------------------------------------------------
'          Rutina para llenar el spread de  Tareas                   -
'---------------------------------------------------------------------
      
Dim intRenglon As Integer
Dim rsDetalle As rdoResultset
Dim rsMecanicos As rdoResultset
Dim rsNombre As rdoResultset
Dim strSQL As String
Dim strNombre As String
Dim x As Boolean

On Error GoTo Err_DespliegaDetalle

Erase mstrArray

' Limpia el spread
'LimpiaBloque sprTareas, 1, 1, sprTareas.MaxRows, sprTareas.MaxCols
sprPartidas.MaxRows = 0

' Crea el rdoResultset de la tabla de ODTDetalle
strSQL = "select A.Nombre,UM.Nombre UnidadMedida, CA.Cantidad,CA.PrecioUnitario " & _
    "from CotizacionArticulo CA " & _
    "   JOIN Articulo A ON A.CveArticulo = CA.CveArticulo " & _
    "   JOIN UnidadMedida UM ON UM.CveUnidadMedida = CA.CveUnidadMedida " & _
    "WHERE CA.CveCotizacion = " & glngCveCotizacion & _
    "ORDER BY NumPartida"

Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)

sprPartidas.MaxRows = rsDetalle.RowCount
' Llena el spread de Tareas
intRenglon = 1
sprPartidas.ReDraw = False
Do Until rsDetalle.EOF

    sprPartidas.Row = intRenglon
    sprPartidas.RowHeight(intRenglon) = 10.5
    
    sprPartidas.Col = COLUMNAARTICULO
    sprPartidas.Text = rsDetalle!Nombre
    sprPartidas.Col = COLUMNACANTIDAD
    sprPartidas.Text = rsDetalle!Cantidad
    sprPartidas.Col = COLUMNAUNIDADMEDIDA
    sprPartidas.Text = rsDetalle!UnidadMedida
    
    MakeFloatCell COLUMNAPRECIO, COLUMNAPRECIO, sprPartidas.Row, sprPartidas.Row, "-99999", "99999", True, True, 2, 0
    sprPartidas.Col = COLUMNAPRECIO
    sprPartidas.Text = rsDetalle!PrecioUnitario
    rsDetalle.MoveNext
    intRenglon = intRenglon + 1
Loop
rsDetalle.Close
sprPartidas.ReDraw = True
mblnCambioSprTareas = False

'-----------------------------------------
sprAreas.MaxRows = 0

' Crea el rdoResultset de la tabla de ODTDetalle
strSQL = "select A.Nombre NombreArea,P.Nombre NombrePersonal " & _
    "from OTPersonalArea CPA WITH(NOLOCK) " & _
    "   JOIN Area A WITH(NOLOCK) ON A.CveArea = CPA.CveArea " & _
    "   JOIN Personal P WITH(NOLOCK) ON P.CvePersonal = CPA.CvePersonal " & _
    "   JOIN CotizacionOT CO WITH(NOLOCK) ON CO.CveOT = CPA.CveOT " & _
    "WHERE CO.CveCotizacion = " & glngCveCotizacion & _
    "ORDER BY CPA.NumPosicion,CPA.CveArea"

Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)

sprAreas.MaxRows = rsDetalle.RowCount + 1
' Llena el spread de Tareas
intRenglon = 1
sprAreas.ReDraw = False
Do Until rsDetalle.EOF

    sprAreas.Row = intRenglon
    sprAreas.RowHeight(intRenglon) = 10.5
    
    sprAreas.Col = 1
    sprAreas.Text = rsDetalle!NombreArea
    sprAreas.Col = 2
    sprAreas.Text = rsDetalle!NombrePersonal
    
    rsDetalle.MoveNext
    intRenglon = intRenglon + 1
Loop
rsDetalle.Close
sprAreas.ReDraw = True
mblnCambioSprTareas = False

Exit Sub

Err_DespliegaDetalle:
  Screen.MousePointer = vbDefault
  MsgBox "Error al Desplegar Detalle de ODT  " & Error, vbCritical
  mblnEdicion = False
  Exit Sub
Resume Next
End Sub
Sub MakeFloatCell(Col As Long, col2 As Long, Row As Long, row2 As Long, floatmin As String, _
    floatmax As String, floatmoney As Boolean, floatsep As Boolean, decplaces As Integer, fpvalue As Double)
    
    sprPartidas.Col = Col
    sprPartidas.col2 = col2
    sprPartidas.Row = Row
    sprPartidas.row2 = row2
    sprPartidas.BlockMode = True
    'Define cells as type FLOAT
    If floatmoney Then
        sprPartidas.CellType = CellTypeCurrency
        sprPartidas.TypeCurrencyShowSymbol = True
        sprPartidas.TypeCurrencyDecPlaces = decplaces
        sprPartidas.TypeCurrencyShowSep = floatsep
        sprPartidas.TypeCurrencyMin = floatmin
        sprPartidas.TypeCurrencyMax = floatmax
    Else
        sprPartidas.CellType = CellTypeNumber
        sprPartidas.TypeNumberDecPlaces = decplaces
        sprPartidas.TypeNumberShowSep = floatsep
        sprPartidas.TypeNumberMin = floatmin
        sprPartidas.TypeNumberMax = floatmax
    End If
    sprPartidas.Value = fpvalue
    sprPartidas.BlockMode = False
    
End Sub

Public Sub PosicionaRegistro(vntValorABuscar As Variant)
'--------------------------------------------------------------------
'   Rutina para posicionar un rdoResultset o rdoResultset                 '
'   en determindado valor de la llave                               '
'       Entrada.-                                                   '
'                vntValorABuscar ->  Valor a Buscar                 '
'-------------------------------------------------------------------'

Dim blnExiste As Boolean

blnExiste = False
If rsODT.RowCount > 0 Then
    rsODT.MoveFirst
    Do While Not rsODT.EOF
        If rsODT!CveODT = Val(vntValorABuscar) Then
            blnExiste = True
            Exit Sub
        End If
        rsODT.MoveNext
    Loop
End If
End Sub

Private Function ValidaCampos()
  
'valida los campos en cuanto a la captura y tipo de dato, posicionandose
'en el control correspondiente despues del error.
Dim blnTemp As Integer
Dim strMsgErrValidacion As String
Dim sngTotalSpread As Single
Dim sngImporteTotal As Single
Dim i As Integer
Dim rsConsulta As rdoResultset
Dim rsCuenta As rdoResultset
Dim strSQL As String
Dim sngSaldo As Single
Dim lngCveTarea As Long
Dim strRazon As String
Dim strTarea As String

ValidaCampos = False

Select Case TabPrincipal.ActiveTab
    Case TABUNIDADES

        If cboCliente.ListIndex = -1 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes seleccionar un cliente para continuar con el proceso", vbExclamation
          cboCliente.SetFocus
          Exit Function
        End If
                
        If Len(txtNombreRepresentante.Text) = 0 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes especificar un representante", vbExclamation
          txtNombreRepresentante.SetFocus
          Exit Function
        End If
                                 
        If cboObra.ListIndex = -1 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes seleccionar una obra", vbExclamation
          cboObra.SetFocus
          Exit Function
        End If
        
        If cboMoneda.ListIndex = -1 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes especificar una moneda", vbExclamation
          cboMoneda.SetFocus
          Exit Function
        End If
        
        If txtNumCotizacion = "" Then
          Screen.MousePointer = vbDefault
          MsgBox "Debe existir un numero de cotizacion", vbExclamation
          txtNumCotizacion.SetFocus
          Exit Function
        End If
                                                    
        If cboTiempoEntrega.ListIndex = -1 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes seleccionar una obra", vbExclamation
          cboTiempoEntrega.SetFocus
          Exit Function
        End If
        
        If cboFormaPago.ListIndex = -1 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes especificar una moneda", vbExclamation
          cboFormaPago.SetFocus
          Exit Function
        End If
                      
        If cboCotizacionTipo.ListIndex = -1 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes especificar un tipo de cotizacion", vbExclamation
          cboCotizacionTipo.SetFocus
          Exit Function
        End If
                                                    
End Select

ValidaCampos = True

End Function

Private Sub CargardoResultsetDeControles()

Dim strSQL As String
Dim strNumFactura As String
Dim i As Integer

On Error GoTo Err_CargaRSet

Screen.MousePointer = vbHourglass

Select Case TabPrincipal.ActiveTab
    Case TABUNIDADES
'        If mblnAlta Then

'        End If

        gstrCveCotizacion = txtNumCotizacion.Text
        
        strSQL = "EXEC Cotizacion_PROCESO_Actualiza " & _
            "@CveCotizacion = NULL" & _
            ",@NumCotizacion ='" & gstrCveCotizacion & "'" & _
            ",@CveCliente= " & cboCliente.ItemData(cboCliente.ListIndex) & _
            ",@CveUsuarioAtiende='" & txtUsuario.Text & "'" & _
            ",@FechaCotizacion='" & Format(txtFechaCotizacion.Text, FECHAYYYYMMDD) & "'"
        If Len(txtFechaEnvio.Text) > 0 Then
            strSQL = strSQL & ",@FechaEnvio='" & Format(txtFechaEnvio.Text, FECHAYYYYMMDD) & "'"
        Else
            strSQL = strSQL & ",@FechaEnvio=NULL"
        End If
        If Len(txtFechaRecepcion.Text) > 0 Then
            strSQL = strSQL & ",@FechaRecepcion='" & Format(txtFechaRecepcion.Text, FECHAYYYYMMDD) & "'"
        Else
            strSQL = strSQL & ",@FechaRecepcion=NULL"
        End If
        strSQL = strSQL & ",@Comentarios=''" & _
            ",@CveCotizacionEstatus=1" & _
            ",@NombreRepresentante='" & txtNombreRepresentante.Text & "'" & _
            ",@CveMoneda= " & cboMoneda.ItemData(cboMoneda.ListIndex) & _
            ",@CveTiempoEntrega= " & cboTiempoEntrega.ItemData(cboTiempoEntrega.ListIndex) & _
            ",@CveFormaPago= " & cboFormaPago.ItemData(cboFormaPago.ListIndex) & _
            ",@CveObra= " & cboObra.ItemData(cboObra.ListIndex) & _
            ",@CveCotizacionTipo=" & cboCotizacionTipo.ItemData(cboCotizacionTipo.ListIndex) & _
            ",@RutaAcceso='" & txtRutaAcceso.Text & "'"
        
        gcn.Execute strSQL
                                                          
End Select

Screen.MousePointer = vbDefault
Exit Sub

Err_CargaRSet:
    Screen.MousePointer = vbDefault
    MsgBox "Error al Cargar rdoResultset de Controles" & Error, vbCritical
    Exit Sub

End Sub
Private Sub Borra()
'******************************************************************
'  Rutina para borrar el registro donde estamos posicionados
'******************************************************************

Dim strSQL As String
Dim rsSubproducto As rdoResultset
Dim rsUltimoDeposito As rdoResultset
Dim strCvePedido As String
Dim sngMonto As Single
Dim i As Integer
Dim rsDetalle As rdoResultset
Dim rsPago As rdoResultset
Dim sngPago As Single


On Error GoTo Err_Borra:

Screen.MousePointer = vbHourglass

If MsgComunes(100) = vbNo Then Exit Sub         'Confirma Borrado

Select Case TabPrincipal.ActiveTab
    Case TABUNIDADES

        gcn.Execute "sp_SIMCancelaIntelisis " & gintCveBase & "," & Val(txtCveODT.Text) & ",'O','C','" & gstrLogin & "'"
        
        gcn.Execute "Delete from ODT where CveODT = " & Val(txtCveODT.Text)
        
        strSQL = "exec Sp_SIMActualizaEficiencia @CveUnidad = " & txtUnidad.Text & "," & _
                                            "@intEstatus = 0, " & _
                                            "@lngLlegadaODT = 0," & _
                                            "@chrTipoEvento = 'Quitar'"
        If gintCveBase = BASETALLERCENTRAL Then gcn.Execute strSQL
        'rsODT.Delete
        rsODT.Requery
        
        ' Borra las Tareas de la tabla de TareasNoRealizadas
        strSQL = "Delete from TareasNoRealizadas where CveODT = " & Val(txtCveODT.Text) & " "
        strSQL = strSQL & "Delete from UnidadKardex where CveODT = " & Val(txtCveODT.Text)
        gcn.Execute strSQL

        If rsODT.EOF Then
            CargaControlesdeResultset
            ToolBar_EstadoBrowse tlbODT
            ToolBoton_Estado tlbODT, "Facturar", True
            lstODT.Enabled = True
            GoTo Salir
        End If
        rsODT.MoveLast
        MuevePrimerReng

End Select

Salir:
    Screen.MousePointer = vbDefault
    Exit Sub

Err_Borra:
    Screen.MousePointer = vbDefault
    MsgBox "Error al Borrar " & Error, vbCritical

End Sub

Private Sub CargaControlesdeResultset()
'******************************
'Despliega los Datos del Registro en la Pantalla
'carga los controles con la información obtenida de la db en el rdoResultset
'******************************

On Error GoTo Err_Carga
Screen.MousePointer = vbHourglass
Dim lngIndiceTemporal As Long
Dim strSQL  As String
Dim rs As rdoResultset

strSQL = "SELECT * FROM Cotizacion WHERE CveCotizacion=" & glngCveCotizacion
Set rs = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
If Not rs.EOF Then
    mblnLlena = True
    Posicionaselector rs!CveCliente, cboCliente
    gstrCveCotizacion = rs!NumCotizacion
    txtUsuario.Text = rs!CveUsuarioAtiende
    txtFechaCotizacion.Text = Format(rs!FechaCotizacion, FECHADDMMYYYY)
    txtNumCotizacion.Text = gstrCveCotizacion
    If IsNull(rs!FechaEnvio) Then
        txtFechaEnvio.Text = ""
    Else
        txtFechaEnvio.Text = Format(rs!FechaEnvio, FECHADDMMYYYY)
    End If
    If IsNull(rs!FechaRecepcion) Then
        txtFechaRecepcion.Text = ""
    Else
        txtFechaRecepcion.Text = Format(rs!FechaRecepcion, FECHADDMMYYYY)
    End If
    Posicionaselector rs!CveCotizacionEstatus, cboCotizacionEstatus
    txtNombreRepresentante.Text = rs!NombreRepresentante
    Posicionaselector rs!CveTiempoEntrega, cboTiempoEntrega
    Posicionaselector rs!CveMoneda, cboMoneda
    Posicionaselector rs!CveFormaPago, cboFormaPago
    LlenaVariosSelectores "SELECT CveObra,Nombre FROM Obra WHERE Activo = 1 AND CveCliente =" & rs!CveCliente, Array("cboObra"), Me
    Posicionaselector rs!CveObra, cboObra
    txtRutaAcceso.Text = rs!RutaAcceso

    Select Case rs!CveCotizacionEstatus
                
        Case 1 'En Proceso
            ToolBoton_Estado tlbODT, "Agregar", True
            ToolBoton_Estado tlbODT, "Actualizar", False
            ToolBoton_Estado tlbODT, "Borrar", True
            ToolBoton_Estado tlbODT, "Cancelar", False
            ToolBoton_Estado tlbODT, "Autoriza", True
            ToolBoton_Estado tlbODT, "Envio", False
            ToolBoton_Estado tlbODT, "Recibe", False
            ToolBoton_Estado tlbODT, "OT", False
            ToolBoton_Estado tlbODT, "Compra", False
        Case 2 'Pendiente por Autorizar
            ToolBoton_Estado tlbODT, "Agregar", True
            ToolBoton_Estado tlbODT, "Actualizar", False
            ToolBoton_Estado tlbODT, "Borrar", True
            ToolBoton_Estado tlbODT, "Cancelar", False
            ToolBoton_Estado tlbODT, "Autoriza", True
            ToolBoton_Estado tlbODT, "Envio", False
            ToolBoton_Estado tlbODT, "Recibe", False
            ToolBoton_Estado tlbODT, "OT", False
            ToolBoton_Estado tlbODT, "Compra", False
        Case 3 'Autorizada a
            ToolBoton_Estado tlbODT, "Agregar", True
            ToolBoton_Estado tlbODT, "Actualizar", False
            ToolBoton_Estado tlbODT, "Borrar", True
            ToolBoton_Estado tlbODT, "Cancelar", False
            ToolBoton_Estado tlbODT, "Autoriza", False
            ToolBoton_Estado tlbODT, "Envio", True
            ToolBoton_Estado tlbODT, "Recibe", False
            ToolBoton_Estado tlbODT, "OT", False
            ToolBoton_Estado tlbODT, "Compra", False
        Case 4 'Enviada a Cliente
            ToolBoton_Estado tlbODT, "Agregar", True
            ToolBoton_Estado tlbODT, "Actualizar", False
            ToolBoton_Estado tlbODT, "Borrar", True
            ToolBoton_Estado tlbODT, "Cancelar", False
            ToolBoton_Estado tlbODT, "Autoriza", False
            ToolBoton_Estado tlbODT, "Envio", False
            ToolBoton_Estado tlbODT, "Recibe", True
            ToolBoton_Estado tlbODT, "OT", False
            ToolBoton_Estado tlbODT, "Compra", False
            
        Case 5 'Autorizada por Cliente
            ToolBoton_Estado tlbODT, "Agregar", True
            ToolBoton_Estado tlbODT, "Actualizar", False
            ToolBoton_Estado tlbODT, "Borrar", True
            ToolBoton_Estado tlbODT, "Cancelar", False
            ToolBoton_Estado tlbODT, "Autoriza", False
            ToolBoton_Estado tlbODT, "Envio", False
            ToolBoton_Estado tlbODT, "Recibe", False
            ToolBoton_Estado tlbODT, "OT", True
            ToolBoton_Estado tlbODT, "Compra", True
    
    End Select
    ToolBoton_Estado tlbODT, "Actualizar", False

Else
    ' Inicializacion Para rdoResultset vacio
    InicializaCampos
End If
rs.Close

mblnEdicion = False

' Despliega las Tareas de la orden
DespliegaDetalle
ActualizaTree

Screen.MousePointer = vbDefault
Exit Sub

Err_Carga:
  Screen.MousePointer = vbDefault
  MsgBox "Error al Cargar Controles con el rdoResultset " & Error, vbCritical
  mblnEdicion = False
  Exit Sub
  Resume Next
End Sub
Private Function Actualiza() As Boolean
'*****************************************************
'  Procedimiento para actualizar o insertar registros
'*****************************************************
On Error GoTo Err_Actualiza
        
Dim bytSeccion As Byte
Dim rs As rdoResultset
Dim lngCveUnidad As Long
Dim strSQL As String
Dim bytLoop As Byte
        
Screen.MousePointer = vbHourglass
Actualiza = False
    
bytLoop = 0
bytSeccion = 1

    'carga rdoResultset con datos de los controles y graba
    If mblnEdicion Or mblnAlta Then
        CargardoResultsetDeControles
    End If
    
    ToolBar_EstadoBrowse tlbODT

    mblnAlta = False
    mblnEdicion = False
    
Exit_Actualiza:
    Actualiza = True
    Screen.MousePointer = vbDefault
    Exit Function
    
    
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
  
MsgBox "Error al Actualizar " & strmsg & " (" & bytSeccion & ")", vbExclamation + vbOKOnly, "Actualiza"
Exit Function
Resume Next
End Function
Private Sub Agrega()
'********************************************************************
'  Rutina que prepara la pantalla para agregar un registro
'********************************************************************

On Error GoTo Err_Agrega

Dim rsQuery As rdoResultset
Dim strSQL As String
Screen.MousePointer = vbHourglass
  
'If GrabarTransPendiente(mblnEdicion, mblnAlta) Then
'    If Not ValidaCampos() Then GoTo Exit_Agrega
'    If Not Actualiza() Then GoTo Exit_Agrega
'End If
InicializaCampos   ' Limpia los controles
Posicionaselector 1, cboCotizacionEstatus
txtFechaCotizacion.Text = Format(Now(), FECHADDMMYYYY)

mblnEdicion = False
mblnAlta = True
ToolBar_EstadoCambio tlbODT
cboCliente.SetFocus

Exit_Agrega:
    Screen.MousePointer = vbDefault
    Exit Sub
  
Err_Agrega:
    
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
  
MsgBox "Error al Agregar " & strmsg, vbExclamation + vbOKOnly, "Agrega"

End Sub
Private Function ObtieneConsecutivo(ByVal vlngCliente As Long) As Long

Dim blnCreado       As Boolean
Dim rsTemp          As rdoResultset
Dim lngBase         As Long
Dim lngMiles        As Long
Dim lngBaseLimSup   As Long
Dim lngBaseLimInf   As Long
Dim vintPrefijo     As Integer
Dim vintCerosSufijo As Integer
Dim strSQL          As String

On Error GoTo err_ObtieneConsecutivoODT


'FORMATO de BAMMCCCC. (Año, Cliente, Consecutivo)
vintCerosSufijo = 4
strSQL = "SELECT DATEPART(yy,GETDATE()) Year, datepart(mm,getdate()) Mes "
Set rsTemp = gcn.OpenResultset(strSQL, rdOpenForwardOnly, rdConcurReadOnly)
blnCreado = True
If rsTemp.EOF Then Exit Function
If IsNull(rsTemp!Year) Or IsNull(rsTemp!Mes) Then Exit Function

'Obtengo Prefijo de la Empresa.
vintPrefijo = Format$(gintCveBase, "0")
vintPrefijo = vintPrefijo & Right$(rsTemp!Year, 1)
vintPrefijo = vintPrefijo & Format$(rsTemp!Mes, "00")
vintPrefijo = Val(vintPrefijo)
rsTemp.Close

lngMiles = 10 ^ vintCerosSufijo
lngBase = vintPrefijo * lngMiles
lngBaseLimSup = lngBase + lngMiles
lngBaseLimInf = lngBase - 1
strSQL = "SELECT MAX(CveODT) FROM  ODT  WHERE CveODT > " & Str$(lngBaseLimInf) & _
               " AND CveODT < " & Str$(lngBaseLimSup)
Set rsTemp = gcn.OpenResultset(strSQL, rdOpenForwardOnly, rdConcurReadOnly)

If Not rsTemp.EOF Then
    If IsNull(rsTemp.rdoColumns(0)) Then
        ObtieneConsecutivoODT = lngBase + 1
    Else
        ObtieneConsecutivoODT = rsTemp.rdoColumns(0) + 1
    End If
Else
    MsgBox "Error al obtener consecutivo"
End If
rsTemp.Close

Exit Function

err_ObtieneConsecutivoODT:
    Screen.MousePointer = vbDefault
    MsgBox " Error al Obtener Consecutivo de ODT " & Error, vbCritical
    
End Function



Private Sub Buscar()
    

frmBuscaRegistro.Show 1
If gstrValorABuscar = "" Then Exit Sub
PosicionaRegistro (gstrValorABuscar)

If rsODT.EOF Then
    PosicionaRegistro (txtCveODT.Text)
    Exit Sub
Else
    CargaControlesdeResultset
    ToolBar_EstadoBrowse tlbODT
End If
    
mblnEdicion = False
Screen.MousePointer = vbDefault
Exit Sub

Err_Busca:
    Screen.MousePointer = vbDefault
    MsgBox "Error al Buscar " & Error, vbCritical

End Sub


Private Sub Cancela()
'************************************************************
' Rutina para Cancelar una operación de modificación o alta
'************************************************************
On Error GoTo Err_Cancela

Screen.MousePointer = vbHourglass
 
Select Case TabPrincipal.ActiveTab
    Case TABUNIDADES
        If mblnAlta And rsODT.RowCount <> 0 Then rsODT.Bookmark = mvntMarca
        CargaControlesdeResultset
        ToolBar_EstadoBrowse tlbODT
        lstODT.Enabled = True
        
        
End Select

mblnAlta = False
mblnEdicion = False


Screen.MousePointer = vbDefault
Exit Sub

Err_Cancela:
    Screen.MousePointer = vbDefault
    MsgBox "Error al Actualizar " & Error, vbCritical

End Sub


Private Sub InicializaCampos()
    Screen.MousePointer = vbHourglass
   
    'limpia controles para proxima captura
    
    Select Case TabPrincipal.ActiveTab
        Case TABUNIDADES
            cboCliente.ListIndex = -1
            txtNombreRepresentante.Text = ""
            cboObra.ListIndex = -1
            txtNumCotizacion.Text = ""
            cboCotizacionEstatus.ListIndex = -1
            cboMoneda.ListIndex = -1
            cboTiempoEntrega.ListIndex = -1
            cboFormaPago.ListIndex = -1
            txtFechaCotizacion.Text = ""
            txtFechaEnvio.Text = ""
            txtFechaRecepcion.Text = ""
                        
            LimpiaBloque sprPartidas, 1, 1, sprPartidas.MaxRows, sprPartidas.MaxCols
                                                
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub cboCliente_Click()

Dim rsConsulta As rdoResultset
Dim strSQL As String

If mblnLlena Then Exit Sub

mblnEdicion = True
ToolBar_EstadoCambio tlbODT

If cboCliente.ListIndex >= 0 Then
    
    strSQL = "EXEC Cotizacion_PROCESO_Consecutivo " & cboCliente.ItemData(cboCliente.ListIndex)
    Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    txtNombreRepresentante.Text = rsConsulta!NombreContacto
    txtNumCotizacion.Text = rsConsulta!NumCotizacion
    rsConsulta.Close
    
    LlenaVariosSelectores "SELECT CveObra,Nombre FROM Obra WHERE Activo = 1 AND CveCliente =" & cboCliente.ItemData(cboCliente.ListIndex), Array("cboObra"), Me
End If


End Sub

Private Sub cboCliente_KeyPress(KeyAscii As Integer)
BuscaEnCombo cboCliente, KeyAscii
End Sub
Private Sub cboEstatus_Click()

If mblnAlta Then Exit Sub
mblnEdicion = True
ToolBar_EstadoCambio tlbODT
lstODT.Enabled = False
End Sub


Private Sub cboEstatus_KeyPress(KeyAscii As Integer)

BuscaEnCombo cboEstatus, KeyAscii

End Sub


Private Sub cboProveedor_KeyPress(KeyAscii As Integer)

BuscaEnCombo cboProveedor, KeyAscii

End Sub


Private Sub cboRazonReparacion_Click()

Dim i As Integer

For i = 1 To sprTareas.DataRowCnt
    sprTareas.Row = i
    sprTareas.Col = COLUMNARAZON
    sprTareas.Text = cboRazonReparacion.Text
Next i

If mblnAlta Then Exit Sub
mblnEdicion = True
ToolBar_EstadoCambio tlbODT
lstODT.Enabled = False
End Sub
Private Sub cboRazonReparacion_KeyPress(KeyAscii As Integer)

BuscaEnCombo cboRazonReparacion, KeyAscii

End Sub
Private Sub chkCargodirecto_Click()
If mblnAlta Then Exit Sub
mblnEdicion = True
ToolBar_EstadoCambio tlbODT
lstODT.Enabled = False
End Sub

Private Sub cmdMueveAnt_Click()
    
MuevePrevioReng

End Sub

Private Sub cmdMueveFin_Click()
    
MueveUltimoReng

End Sub

Private Sub cmdMueveInicio_Click()
    
MuevePrimerReng

End Sub

Private Sub cmdMueveSig_Click()
    
MueveSigReng

End Sub








Private Sub Command1_Click()

If ValidaCampos() Then
    If Actualiza() Then
        frmPartidas.Show vbModal
        ActualizaTree
        DespliegaDetalle
    End If
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 112 Then
    ShowHelpTopic (10)
    KeyCode = 0
End If

End Sub
Private Sub Form_Load()

On Error GoTo Err_Form_Load

Dim intRet As Integer
Dim strSQL As String
Dim strTexto As String
Dim strFechaLimite As String
Dim rdBase As rdoResultset
  
Screen.MousePointer = vbHourglass

'***********************
If App.PrevInstance Then ' checa que no haya otro exe ejecutándose
    strTexto = frmOT.Caption
    frmOT.Caption = ""
    MsgBox App.EXEName & "  Actualmente ejecutándose"
    AppActivate strTexto
    End
End If
'**********************

staEstatusBar.Panels(1).Text = gstrProducto

mdatUltimaHoraEjecucion = DateTime.Now

' Obtiene la resolucion actual a la cambia
'ObtieneResolucion gsngAncho, gsngAlto
'CambiaResolucion RESOLUCIONANCHO, RESOLUCIONALTO

strSQL = "SELECT CveCliente,Nombre from Cliente WHERE Activo = 1 ORDER BY Nombre " & _
    "SELECT CveCotizacionEstatus,Nombre from CotizacionEstatus " & _
    "SELECT CveMoneda,Nombre FROM Moneda " & _
    "SELECT CveTiempoEntrega, Nombre FROM TiempoEntrega " & _
    "SELECT CveFormaPago,Nombre FROM FormaPago " & _
    "SELECT CveCotizacionTipo,Nombre FROM CotizacionTipo "

LlenaVariosSelectores strSQL, Array("cboCliente", "cboCotizacionEstatus", "cboMoneda", _
                "cboTiempoEntrega", "cboFormaPago", "cboCotizacionTipo"), Me

ActualizaTree

'5   Autorizada por Cliente  AC

sprPartidas.MaxRows = 0
sprPartidas.MaxCols = COLUMNAPRECIO

sprPartidas.ColWidth(COLUMNACANTIDAD) = 4
sprPartidas.ColWidth(COLUMNAUNIDADMEDIDA) = 4
sprPartidas.ColWidth(COLUMNAPRECIO) = 8

sprPartidas.Row = -1000

sprPartidas.Col = COLUMNAARTICULO
sprPartidas.FontBold = True
sprPartidas.TypeHAlign = TypeHAlignCenter
sprPartidas.Text = "Descripcion"
sprPartidas.ColWidth(COLUMNAARTICULO) = 28

sprPartidas.Col = COLUMNACANTIDAD
sprPartidas.FontBold = True
sprPartidas.TypeHAlign = TypeHAlignCenter
sprPartidas.Text = "Cant"

sprPartidas.Col = COLUMNAUNIDADMEDIDA
sprPartidas.FontBold = True
sprPartidas.TypeHAlign = TypeHAlignCenter
sprPartidas.Text = "UN"

sprPartidas.Col = COLUMNAPRECIO
sprPartidas.FontBold = True
sprPartidas.TypeHAlign = TypeHAlignCenter
sprPartidas.Text = "Precio Unitario"


'Se agregan iconos a las pestañas de cada uno de los tabs
TabPrincipal.Tab = TABUNIDADES
'TabPrincipal.TabPicture = LoadPicture(gstrDirectorioIconos & "Neo.ico")
TabPrincipal.Tab = TABREPORTES
'TabPrincipal.TabPicture = LoadPicture(gstrDirectorioIconos & "Chart5.ico")


'sprAreas.ColWidth(COLUMNACANTIDAD) = 4
'sprAreas.ColWidth(COLUMNAUNIDADMEDIDA) = 4
'sprAreas.ColWidth(COLUMNAPRECIO) = 8

sprAreas.Row = -1000

sprAreas.Col = 1
sprAreas.FontBold = True
sprAreas.TypeHAlign = TypeHAlignCenter
sprAreas.Text = "Area"
sprAreas.ColWidth(COLUMNAARTICULO) = 8


sprAreas.Col = 2
sprAreas.FontBold = True
sprAreas.TypeHAlign = TypeHAlignCenter
sprAreas.Text = "Coordinadores"











' Carga controles del rdoResultset
'CargaControlesdeResultset

' Despliega el nombre del servidor
'strSQL = "select Nombre from Base Where CveBase = " & gintCveBase
'Set rdBase = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
'staEstatusBar.Panels(2).Text = "Ubicación: " & rdBase!Nombre & "    Versión:" & App.Major & "." & App.Minor & "." & App.Revision
'rdBase.Close

' Rutina para preparar toolbar
ToolBar_EstadoBrowse tlbODT

Screen.MousePointer = vbDefault

Exit_Form_Load:
    Screen.MousePointer = vbDefault
    Exit Sub
 
Err_Form_Load:
  If Err = 53 Then Resume Next          ' No encuentra algun icono
  Screen.MousePointer = vbDefault
  MsgBox "Error en el Load " & Error, vbCritical
  Unload Me
  Exit Sub
Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    

    Set frmOT = Nothing
    CierraConeccion
    End
End Sub

Private Sub lstODT_Click()
BuscaClave lstODT.ItemData(lstODT.ListIndex)
End Sub

Private Sub sprAreas_DblClick(ByVal Col As Long, ByVal Row As Long)

Dim strSQL As String
Dim blnReturn As Boolean
Dim varValor As Variant

On Error GoTo err_sprTareas_DblClick

'If Row = 0 Or (varValor = "0" Or varValor = "1") Then Exit Sub
If Row = 0 Then Exit Sub

sprAreas.Col = Col
sprAreas.Row = Row

Screen.MousePointer = vbHourglass

Select Case Col
    Case 1
        strSQL = "SELECT Nombre FROM Area ORDER BY Nombre "
        
        ' Pone el combo en el spread
        LlenaComboSpread sprAreas, strSQL, Col, Row
        sprAreas.MaxRows = sprAreas.MaxRows + 1

    Case 2
        strSQL = "SELECT Nombre FROM Personal WHERE Activo = 1 ORDER BY Nombre "
        
        ' Pone el combo en el spread
        LlenaComboSpread sprAreas, strSQL, Col, Row

End Select

Screen.MousePointer = vbDefault

Exit Sub
err_sprTareas_DblClick:
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

MsgBox "Error en elemento del spread" & strmsg, vbCritical, "err_sprAreas_DblClick"
Exit Sub

End Sub


Private Sub sprTareas_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

mlngRenglon = BlockRow
mlngColumna = BlockCol
mlngRenglon2 = BlockRow2
mlngColumna2 = BlockCol2

End Sub



Private Sub sprTareas_Change(ByVal Col As Long, ByVal Row As Long)

mblnCambioSprTareas = True
    
sprTareas.Row = Row
sprTareas.Col = Col
If Col = COLUMNARAZON Then
    sprTareas.Text = Trim$(sprTareas.Text)
    If sprTareas.CellType = SS_CELL_TYPE_COMBOBOX Then
        ' Define cell as type EDIT
        sprTareas.CellType = SS_CELL_TYPE_EDIT
    End If
End If

If mblnAlta Then Exit Sub
ToolBar_EstadoCambio tlbODT
lstODT.Enabled = False
End Sub

Private Sub sprTareas_Click(ByVal Col As Long, ByVal Row As Long)

Dim bolTienePicture As Boolean
Dim blnReturn As Boolean
Dim varValor As Variant

mlngRenglon = Row
mlngColumna = Col
mlngRenglon2 = Row
mlngColumna2 = Col

If Col = 1 Then
    blnReturn = sprTareas.GetText(COLUMNARAMIFICACION, Row, varValor)
    If varValor = "0" Or varValor = "1" Then
        MuestraEscondeRenglones 1, Row
    End If
End If

End Sub


Private Sub sprTareas_KeyDown(KeyCode As Integer, Shift As Integer)

Dim intCveTarea As Integer
Dim intCveTareaPadre As Integer
Dim strSQL As String
Dim strComentarios As String
Dim i As Integer
Dim rsConsulta As rdoResultset

On Error GoTo err_sprTareas_KeyDown

If KeyCode = vbKeyDelete Then
    Screen.MousePointer = vbHourglass
    If cboEstatus.ItemData(cboEstatus.ListIndex) > ESTATUSABIERTA Then
        Screen.MousePointer = vbDefault
        MsgBox "No se puede modificar la orden, pues ya esta cerrada"
        Exit Sub
    End If
    

    For i = mlngRenglon To mlngRenglon2
        sprTareas.Row = i
        sprTareas.Col = COLUMNACVETAREAPADRE
        intCveTareaPadre = Val(sprTareas.Text)
        sprTareas.Col = COLUMNACVETAREA
        intCveTarea = Val(sprTareas.Text)
        sprTareas.Col = COLUMNACOMENTARIOS
        strComentarios = CVTexto(sprTareas.Text)
        
        strSQL = "select * from TarjetaAccion TA WHERE CveODT  =  " & gstrCveCotizacion & _
                    "  AND CveTarea = " & intCveTarea
        Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
        If rsConsulta.RowCount > 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "No puede eliminarse una tarea que tiene acciones de Operario.", vbCritical + vbOKOnly, "sprTareas"
            rsConsulta.Close
            Exit Sub
        End If
        rsConsulta.Close

        strSQL = "select ID " & _
            "from ExtIntelisisInv INV " & _
            "WHERE INV.Referencia = 'ODT " & gstrCveCotizacion & "'" & _
                    "  AND LTRIM(RTRIM(INV.CveTarea)) = '" & intCveTarea & "'"
        Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
        If rsConsulta.RowCount > 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "No puede eliminarse una tarea que tiene Movimientos de Salida en Intelisis.", vbCritical + vbOKOnly, "sprTareas"
            rsConsulta.Close
            Exit Sub
        End If
        rsConsulta.Close

    Next i

    LimpiaBloque sprTareas, mlngRenglon, mlngColumna, mlngRenglon2, mlngColumna2
    mblnCambioSprTareas = True
    
    ' Si la tarea es parte de un servicio la agrega como predictivo
    If intCveTareaPadre <> 0 Then
        strSQL = "Insert Into UnidadPredictivo "
        strSQL = strSQL & " (CveUnidad,CveTarea,KmsRegistro,FechaRegistro,KmsPronostico,"
        strSQL = strSQL & "  FechaPronostico,Comentarios) "
        strSQL = strSQL & " Values(" & Val(txtUnidad.Text) & "," & intCveTarea & ","
        strSQL = strSQL & Val(txtKmsAcumulados.Text) & ",'" & Format(Now, FECHAYYYYMMDD) & "',"
        strSQL = strSQL & Val(txtKmsAcumulados.Text) & ",'" & Format(Now, FECHAYYYYMMDD) & "','"
        strSQL = strSQL & strComentarios & "')"
        ' gcn.Execute strSQL
    End If

    Screen.MousePointer = vbDefault
    If mblnAlta Then Exit Sub
    ToolBar_EstadoCambio tlbODT
End If
Exit Sub
err_sprTareas_KeyDown:
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

MsgBox "Error en elemento del spread" & strmsg, vbCritical, "err_sprTareas_KeyDown"
Exit Sub

End Sub
Private Sub sprTareas_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
Dim ret As Long
    
On Error GoTo err_sprTareas_TextTipFetch
    
    'ret = sprTareas.GetRowItemData(Row)
If Col = 3 And Row <> 0 Then
    'If ret = 0 Then Exit Sub
    
    If Row <= 400 Then
        TipText = mstrArray(Row, Col)   'Set the text tip text
        MultiLine = 1
        TipWidth = 3500
        ShowTip = True
    Else
        TipText = ""
        ShowTip = False
    End If
End If

Exit Sub
err_sprTareas_TextTipFetch:
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

MsgBox "Error en elemento del spread" & strmsg, vbCritical, "err_sprTareas_TextTipFetch"
Exit Sub
Resume Next
End Sub

Private Sub tlbODT_ButtonClick(ByVal Button As Button)

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
   
   Case Is = "Refrescar"
        rsODT.Requery
        rsODT.MoveLast
        CargaControlesdeResultset
        ToolBar_EstadoBrowse tlbODT
        mblnAlta = False
        mblnEdicion = False
        
   Case Is = "Agregar"
        Agrega

   Case Is = "Actualizar"
'        If cboEstatus.ItemData(cboEstatus.ListIndex) > ESTATUSABIERTA Then
'            MsgBox "La ODT ya esta cerrada, no se puede modificar"
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
        If ValidaCampos() Then
            If Actualiza() Then
            End If
        End If
      
   Case Is = "Borrar"
 '       If cboEstatus.ItemData(cboEstatus.ListIndex) > ESTATUSABIERTA Then
 '           MsgBox "La ODT ya esta cerrada, no se puede borrar"
 '           Screen.MousePointer = vbDefault
 '           Exit Sub
 '       End If
 '       Borra
      
   Case Is = "Cancelar"
 '       Cancela
   
   Case Is = "Autoriza"
    gcn.Execute "UPDATE Cotizacion SET CveCotizacionEstatus = 2 WHERE CveCotizacion =" & glngCveCotizacion
    CargaControlesdeResultset

   Case Is = "Envio"
    gcn.Execute "UPDATE Cotizacion SET CveCotizacionEstatus = 4,FechaEnvio = getdate() WHERE CveCotizacion =" & glngCveCotizacion
    CargaControlesdeResultset
   
   Case Is = "Recibe"
    Select Case MsgBox("La cotizacion fue aceptada por el Cliente", vbYesNoCancel + vbQuestion, "Revision del cliente")
        Case vbYes
            gcn.Execute "UPDATE Cotizacion SET CveCotizacionEstatus = 5,FechaRecepcion=GETDATE() WHERE CveCotizacion =" & glngCveCotizacion
        Case vbNo
            gcn.Execute "UPDATE Cotizacion SET CveCotizacionEstatus = 1,FechaRecepcion=GETDATE() WHERE CveCotizacion =" & glngCveCotizacion
    End Select
    CargaControlesdeResultset
   Case Is = "OT"

   Case Is = "Compra"
        glngCveRequisicion = 0
        frmRequisicion.Show vbModal
        If glngCveRequisicion > 0 Then
             Dim frmRep As New frmReporte
            
            frmRep.mstrNombreArchivo = "E:\SICIP\SI004.RPT"
            frmRep.mstrSQL = "SELECT vw_RequisionDetalle_Proveedor.NombreArticulo, vw_RequisionDetalle_Proveedor.Codigo" & _
                    ", vw_RequisionDetalle_Proveedor.CantidadRequerida, vw_RequisionDetalle_Proveedor.CodigoArticuloProveedor" & _
                    ", vw_RequisionDetalle_Proveedor.NombreProveedor, vw_Cotizacion.NomCliente, vw_Cotizacion.NomObra" & _
                    ", vw_Cotizacion.NumCotizacion, vw_RequisionDetalle_Proveedor.FechaAlta " & _
            "FROM   SICIP.dbo.vw_Cotizacion vw_Cotizacion " & _
                    "INNER JOIN SICIP.dbo.vw_RequisionDetalle_Proveedor vw_RequisionDetalle_Proveedor ON vw_Cotizacion.CveCotizacion=vw_RequisionDetalle_Proveedor.CveCotizacion " & _
            "WHERE vw_RequisionDetalle_Proveedor.CveRequisicion = " & glngCveRequisicion & _
            " ORDER BY vw_RequisionDetalle_Proveedor.NombreArticulo"
            
    '         frmRep.PasarParametros "", 0
             frmRep.Show vbModal
            
             Set frmRep = Nothing
        End If
         
   Case Is = "Imprimir"
        Imprimir

   Case Is = "Salir"
        Screen.MousePointer = vbDefault
        Unload Me
        End

End Select
ActualizaTree
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

MsgBox "Error en elemento del ToolBar " & strmsg, vbCritical, "tlbODT_ButtonClick"
Resume Next
End Sub
Public Sub ImprimirNotaCargo(ByVal vlngCveODT As Long, ByVal vstrCosto As String)

On Error GoTo err_ImprimirNotaCargo

If vstrCosto = "0" Or vstrCosto = "" Then
    MsgBox "No existen cargos para la ODT", vbExclamation, "ImprimirNotaCargo"
    Exit Sub
End If

Screen.MousePointer = vbHourglass

 'Conección del reporte con la base de datos
rptReporte.Connect = "DSN = " & gstrServidor & ";" & _
                     "UID = " & LOGIN & ";" & _
                     "PWD = " & PASSWORD & ";" & _
                     "DSQ = " & gstrBaseDeDatos
'Pasa parámetros al reporte
rptReporte.Destination = crptToWindow
rptReporte.Formulas(1) = "Empresa= """ & gstrNombreEmpresa & """"
rptReporte.WindowTitle = "Nota de Cargo"
rptReporte.ReportFileName = gstrDirectorioRpt & "Tp059.rpt"
If Dir(rptReporte.ReportFileName) = "" Then
    MsgBox "Reporte " & rptReporte.ReportFileName & ", No existe en directorio del SIM", vbCritical, "ImprimirNotaCargo"
    Exit Sub
End If
rptReporte.SQLQuery = "SELECT " & _
                        "vw_ODTDetalleRefaccion.CveODT, vw_ODTDetalleRefaccion.CveUnidad, vw_ODTDetalleRefaccion.CveTarea, vw_ODTDetalleRefaccion.Nombre" & _
                        ",vw_ODTDetalleRefaccion.IdArticulo, vw_ODTDetalleRefaccion.Descripcion, vw_ODTDetalleRefaccion.Cantidad, vw_ODTDetalleRefaccion.Costo" & _
                        ",vw_ODTDetalleRefaccion.NombreSupervisor, vw_ODTDetalleRefaccion.Signante " & _
                    "FROM " & gstrBaseDeDatos & ".dbo.vw_ODTDetalleRefaccion vw_ODTDetalleRefaccion " & _
                    "WHERE vw_ODTDetalleRefaccion.CveODT = " & vlngCveODT & Chr(10) & Chr(13) & _
                    " ORDER BY vw_ODTDetalleRefaccion.CveTarea ASC"
' Manda el reporte
rptReporte.Action = 1

Screen.MousePointer = vbDefault
Exit Sub

err_ImprimirNotaCargo:
Dim lngIndice As Long
Dim strmsg As String

Select Case Err
    Case 40002
        For lngIndice = 0 To rdoErrors.Count - 1
            strmsg = strmsg & rdoErrors(lngIndice).Description & Chr(13)
        Next lngIndice
    Case Else
        strmsg = Err & " " & Error
End Select

Screen.MousePointer = vbDefault
MsgBox "Error al imprimir la nota de Cargo." & Chr$(13) & _
     strmsg, vbCritical, "ImprimirNotaCargo"
Exit Sub
End Sub

Private Sub treeview1_NodeClick(ByVal Node As ComctlLib.Node)

If IsNumeric(Right(Node.Key, 1)) Then
    glngCveCotizacion = Mid(Node.Key, InStr(1, Node.Key, "-") + 1, 40)
    CargaControlesdeResultset
End If

End Sub


Private Sub txtCveODT_Change()

' Asigna a una variable global el no. de orden actual
If txtCveODT.Text <> "" Then gstrCveCotizacion = txtCveODT.Text

End Sub

Private Sub txtFechaInicio_Change()

If mblnAlta Then Exit Sub
mblnEdicion = True
ToolBar_EstadoCambio tlbODT
lstODT.Enabled = False
End Sub


Private Sub txtFechaPrometida_Change()
        
If mblnAlta Then Exit Sub
mblnEdicion = True
ToolBar_EstadoCambio tlbODT
lstODT.Enabled = False
End Sub

Sub Imprimir()

    Dim frmRep As New frmReporte
   
    frmRep.mstrNombreArchivo = "E:\SICIP\SI001.RPT"
    frmRep.mstrSQL = "SELECT vw_CotizacionArticulo.NumPartida, vw_Cotizacion.NomCliente, vw_Cotizacion.NombreRepresentante" & _
            ", vw_Cotizacion.NomObra, vw_Cotizacion.FechaCotizacion, vw_Cotizacion.NomCotizacionTipo, vw_Cotizacion.NumCotizacion" & _
            ", vw_CotizacionArticulo.NomArticulo, vw_CotizacionArticulo.NomUnidadMedida, vw_CotizacionArticulo.PrecioUnitario" & _
            ", vw_CotizacionArticulo.Cantidad, vw_CotizacionArticuloDetalle.NomArticulo, vw_CotizacionArticuloDetalle.NomUnidadMedida" & _
            ", vw_Cotizacion.NomMonedaCorto, vw_Cotizacion.NomTiempoEntrega, vw_Cotizacion.NomFormaPago, Usuario.Nombre " & _
    "FROM   ((SICIP.dbo.vw_Cotizacion vw_Cotizacion " & _
        "INNER JOIN SICIP.dbo.vw_CotizacionArticulo vw_CotizacionArticulo ON vw_Cotizacion.CveCotizacion=vw_CotizacionArticulo.CveCotizacion) " & _
            "INNER JOIN SICIP.dbo.Usuario Usuario ON vw_Cotizacion.CveUsuarioAtiende=Usuario.CveUsuario) " & _
                "INNER JOIN SICIP.dbo.vw_CotizacionArticuloDetalle vw_CotizacionArticuloDetalle ON (vw_CotizacionArticulo.CveCotizacion=vw_CotizacionArticuloDetalle.CveCotizacion) AND (vw_CotizacionArticulo.NumPartida=vw_CotizacionArticuloDetalle.NumPartida) " & _
    "WHERE vw_Cotizacion.CveCotizacion = " & glngCveCotizacion & _
            "ORDER BY vw_CotizacionArticulo.NumPartida "
    frmRep.Show vbModal
   
    Set frmRep = Nothing

End Sub

Private Sub txtFechaPrometida_GotFocus()

If txtFechaPrometida.Text = "" Then
    txtFechaPrometida.Text = Format(ObtieneFechaHora(1), FECHADDMMYY & HORAMINUTOS)
End If

End Sub
Private Sub txtKmsAcumulados_Change()

If mblnAlta Then Exit Sub
mblnEdicion = True
ToolBar_EstadoCambio tlbODT
lstODT.Enabled = False
End Sub

Private Sub txtNumCajon_Change()

If mblnAlta Then Exit Sub
mblnEdicion = True
ToolBar_EstadoCambio tlbODT
lstODT.Enabled = False
End Sub

Public Sub ActualizaDetalle()

On Error GoTo Err_ActualizaDetalle

Dim rsRegistro As rdoResultset
Dim rsRenglon As rdoResultset
Dim rsDatosRazon As rdoResultset
Dim intCveTarea As Integer
Dim blnEncontro As Boolean
Dim strSQL As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim i As Integer
Dim intNumRenglon As Integer
Dim strComentarios As String
Dim strRazon As String
Dim intCveTareaPadre As Integer

strSQL = "'<O CveODT=""" & gstrCveCotizacion & """>"
strSQL2 = "'<O CveODT=""" & gstrCveCotizacion & """>"
strSQL3 = "'<O CveODT=""" & gstrCveCotizacion & """>"
'-----------------------------------------------
' Inserta las tareas que no estan dadas de alta
'-----------------------------------------------
For i = 1 To sprTareas.DataRowCnt
    sprTareas.Row = i
    sprTareas.Col = COLUMNACVETAREA
    intCveTarea = Val(sprTareas.Text)
    sprTareas.Col = COLUMNACOMENTARIOS
    If Len(CVTexto(sprTareas.Text)) > 0 And InStr(1, CVTexto(sprTareas.Text), ">>") > 0 Then
        strComentarios = Mid(CVTexto(sprTareas.Text), 1, InStr(1, CVTexto(sprTareas.Text), ">>") - 1)
    Else
        strComentarios = CVTexto(sprTareas.Text)
    End If
    sprTareas.Col = COLUMNARAZON
    strRazon = CVTexto(sprTareas.Text)
    sprTareas.Col = COLUMNACVETAREAPADRE
    intCveTareaPadre = Val(sprTareas.Text)
    
    If Len(strSQL) > 7800 Then
        If Len(strSQL2) > 7800 Then
            strSQL3 = strSQL3 & "<D T=""" & intCveTarea & """ C=""" & Trim(strComentarios) & """ M=""" & strRazon & """ TP=""" & intCveTareaPadre & """/>"
        Else
            strSQL2 = strSQL2 & "<D T=""" & intCveTarea & """ C=""" & Trim(strComentarios) & """ M=""" & strRazon & """ TP=""" & intCveTareaPadre & """/>"
        End If
    Else
        strSQL = strSQL & "<D T=""" & intCveTarea & """ C=""" & Trim(strComentarios) & """ M=""" & strRazon & """ TP=""" & intCveTareaPadre & """/>"
    End If

Next i
strSQL = strSQL & "</O>'"
strSQL2 = strSQL2 & "</O>'"
strSQL3 = strSQL3 & "</O>'"

gcn.Execute "EXEC ODTDetalle_PROCESO_Update @XML=" & strSQL & ",@Depura = 0,@CveUsuario='" & gstrLogin & "',@XML2=" & strSQL2 & ",@XML3=" & strSQL3

' Refresca el spread de Tareas
DespliegaDetalle

Exit Sub

Err_ActualizaDetalle:
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
  
MsgBox "Error al actualizar el detalle " & strmsg, vbExclamation + vbOKOnly, "ActualizaDetalle"
mblnEdicion = False
Resume Next

End Sub

Public Sub DespliegaTareasAgregadas()
'---------------------------------------------------------------------
'   Rutina para llenar el spread de  Tareas con los preventivos      -
'   que le tocan a la unidad. El rdoResultset se crea en una rutina  -
'   anterior                                                         -
'---------------------------------------------------------------------
Dim i As Integer
Dim intTareasExistentes As Integer
Dim rsExiste As rdoResultset
Dim rsRenglon As rdoResultset
Dim rsDatosRazon As rdoResultset
Dim blnExisteODT As Boolean
Dim intNumRenglon As Integer
Dim strSQL As String
Dim strXML As String

On Error GoTo Err_Despliega

' Por si no trae tareas el vector
If gintNumTareas = 0 Then Exit Sub

' Define apartir de cual renglon se desplegaran
intTareasExistentes = sprTareas.DataRowCnt

' Verifica si ya esta grabada la ODT
strSQL = "select CveODT from ODT where CveODT = " & gstrCveCotizacion
Set rsExiste = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
If rsExiste.EOF Then
    blnExisteODT = False
Else
    blnExisteODT = True
    strSQL = "select Max(NumRenglon) As Maximo from ODTDetalle where CveODT = " & gstrCveCotizacion
    Set rsRenglon = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    If IsNull(rsRenglon!Maximo) Then
        intNumRenglon = 1
    Else
        intNumRenglon = rsRenglon!Maximo + 1
    End If
    rsRenglon.Close
End If
rsExiste.Close



strXML = "<T CveODT=""" & gstrCveCotizacion & """>"
For i = 1 To gintNumTareas

    strXML = strXML & "<TS CveTarea=""" & gTareas(i).CveTarea & """ Orden=""" & intNumRenglon & """/>"

Next
strXML = strXML & "</T>"

Dim lngPosicionPadre As Long
Dim strMotivoReparacion As String
Dim intMotivoReparacion As Integer

Set rsDatosRazon = gcn.OpenResultset("select CveMotivoReparacion,Nombre from MotivoReparacion where CveMotivoReparacion = " & cboRazonReparacion.ItemData(cboRazonReparacion.ListIndex), rdOpenKeyset, rdConcurRowVer)
If rsDatosRazon.EOF Then
    strMotivoReparacion = ""
    intMotivoReparacion = 0
Else
    strMotivoReparacion = rsDatosRazon!Nombre
    intMotivoReparacion = rsDatosRazon!CveMotivoReparacion
End If
rsDatosRazon.Close

strSQL = "EXEC ODTDetalle_PROCESO_Select_bck " & gstrCveCotizacion & ",'" & strXML & "'"
'strSQL = "EXEC ODTDetalle_PROCESO_Select " & gstrCveCotizacion & ",'" & strXML & "'"
Set rsExiste = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
i = 1

sprTareas.MaxRows = sprTareas.MaxRows + rsExiste.RowCount

Do Until rsExiste.EOF
    sprTareas.Row = intTareasExistentes + i
    If rsExiste!TotalHijas > 0 Then
        sprTareas.Col = COLUMNACOLAPSADOR
        lngPosicionPadre = sprTareas.Row
        
        sprTareas.Col = -1
        sprTareas.BackColor = &HC0C0C0      'Gray
        sprTareas.ForeColor = &HC00000     'Blue
        sprTareas.FontBold = True
        
        
        sprTareas.Col = COLUMNACOLAPSADOR
        sprTareas.CellType = CellTypePicture
        sprTareas.TypePictCenter = True
        sprTareas.TypePictMaintainScale = False
        sprTareas.TypePictStretch = False
        sprTareas.Col = COLUMNACOLAPSADOR
        sprTareas.TypePictPicture = imgMenos.Picture  'LoadPicture("C:\Program Files\Spread60\Samples\ActiveX\VB6\Demo Overview\images\minus.bmp")
        
        'Add picture state values
        sprTareas.Col = COLUMNARAMIFICACION
        sprTareas.Text = "0"
        
        sprTareas.SetRowItemData lngPosicionPadre, rsExiste!TotalHijas
        
        sprTareas.SetCellBorder 1, sprTareas.Row, sprTareas.MaxCols, 1, SS_BORDER_TYPE_LEFT + SS_BORDER_TYPE_RIGHT + SS_BORDER_TYPE_TOP + SS_BORDER_TYPE_BOTTOM, &H808080, CellBorderStyleSolid
    
    End If
    
    sprTareas.Row = intTareasExistentes + i
    sprTareas.Col = COLUMNACVETAREA
    sprTareas.Text = rsExiste!CveTarea
    sprTareas.Col = COLUMNATAREA
    sprTareas.Text = rsExiste!Nombre
    'sprTareas.Col = COLUMNACOMENTARIOS
    'sprTareas.Text = gTareas(i).Comentarios
    sprTareas.Col = COLUMNACVETAREAPADRE
    sprTareas.Text = rsExiste!CveTareaPadre
    sprTareas.Col = COLUMNARAZON
    sprTareas.Text = strMotivoReparacion

    If blnExisteODT Then
        strSQL = "Insert Into ODTDetalle "
        strSQL = strSQL & " (CveODT,CveTarea,NumRenglon,CostoRefacciones,CostoManoDeObra,CostoTallerExterno,PrecioRefacciones,PrecioManoDeObra,PrecioTallerExterno,Comentarios,CveMotivoReparacion,CveTareaPadre)"
        strSQL = strSQL & " Values(" & gstrCveCotizacion & "," & rsExiste!CveTarea & "," & intNumRenglon & ",0,0,0,0,0,0,''," & intMotivoReparacion & "," & rsExiste!CveTareaPadre & ")"
        gcn.Execute strSQL
        intNumRenglon = intNumRenglon + 1
    End If

    i = i + 1
    rsExiste.MoveNext
Loop
rsExiste.Close

sprTareas.ReDraw = True

If blnExisteODT Then
    '---------------------------------------------------
    '   Rutinas para avisar al ERP que ODT se abrió
    '---------------------------------------------------
    If gblnInterfaseERP Then
        gcn.Execute "EXEC sp_SIMEnviaERP " & _
                "@CveBase=" & gintCveBase & _
                ",@CveLlegadaODT =" & gstrCveCotizacion & _
                ",@Tipo='N'"
    End If
Else
    mblnCambioSprTareas = True
End If

Exit Sub

Err_Despliega:
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
  
MsgBox "Error al Desplegar Tareas Agregadas " & strmsg, vbExclamation + vbOKOnly, "DespliegaTareasAgregadas"
mblnEdicion = False
Resume Next
End Sub
Sub MuestraEscondeRenglones(Col As Long, Row As Long)
'Collapse or uncollape the specified rows
Dim i As Integer
Dim collapsetype As Integer

sprTareas.Row = Row
sprTareas.Col = COLUMNARAMIFICACION

If sprTareas.Text = "0" Then
    collapsetype = 0  'collape/hide rows : minus picture
    sprTareas.Col = COLUMNACOLAPSADOR
    sprTareas.TypePictPicture = imgMas.Picture
    sprTareas.Col = COLUMNARAMIFICACION
    sprTareas.Text = "1"
Else
    collapsetype = 1  'uncollapse / show rows: plus picture
    sprTareas.Col = COLUMNACOLAPSADOR
    sprTareas.TypePictPicture = imgMenos.Picture
    sprTareas.Col = COLUMNARAMIFICACION
    sprTareas.Text = "0"
End If

sprTareas.ReDraw = False
For i = 1 To sprTareas.GetRowItemData(Row)
    sprTareas.Row = sprTareas.Row + 1
    If collapsetype = 0 Then
        sprTareas.RowHidden = True
    Else
        sprTareas.RowHidden = False
    End If
Next i
sprTareas.ReDraw = True
     
End Sub


Private Sub txtUnidad_Change()

Dim strSQL As String
Dim rsUnidad As rdoResultset

If txtUnidad.Text <> "" Then
    strSQL = "select * from Unidad where CveUnidad = " & txtUnidad.Text
    Set rsUnidad = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    If Not rsUnidad.EOF Then txtKmsAcumulados = rsUnidad!KmsAcumulados
    rsUnidad.Close
End If

End Sub
Private Sub BuscaClave(ByVal vstrClave As String)
    
    Dim lngLlave As Long

    Screen.MousePointer = vbHourglass

    lngLlave = txtCveODT
    
    PosicionaRegistro (vstrClave)
    If rsODT.EOF Then
        PosicionaRegistro (lngLlave)
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        CargaControlesdeResultset
        ToolBar_EstadoBrowse tlbODT
        lstODT.Enabled = True
    End If
        
    mblnEdicion = False
    Screen.MousePointer = vbDefault
    Exit Sub

Err_Busca:
    Screen.MousePointer = vbDefault
    MsgBox "Error al Buscar " & Error, vbCritical

End Sub

Public Sub CerrarODT()

On Error GoTo Err_CerrarODT

Dim strSQL As String
Dim strLista As String
Dim rsDetalle As rdoResultset
Dim rsMecanico As rdoResultset
Dim rsCostos As rdoResultset
Dim rsDatosODT As rdoResultset
Dim rsCliente As rdoResultset
Dim rsKardex As rdoResultset
Dim rsUnidad As rdoResultset
Dim lngCveMecanico As Long
Dim sngCostoTallerExterno As Single
Dim sngPrecioTallerExterno As Single
Dim blnERP As Boolean

' Borra en Unidad Kardex por si acaso ya se hubiera intentado cerrar y el proceso
' haya quedado inconcluso
strSQL = "delete from UnidadKardex where CveUnidad = " & txtUnidad.Text
strSQL = strSQL & " and CveODT = " & txtCveODT.Text
strSQL = strSQL & " and CveBase = " & gintCveBase
gcn.Execute strSQL

' Obtiene costos total de Taller externo
If gsngCostoTallerExterno > 0 Then
    sngCostoTallerExterno = Format(gsngCostoTallerExterno, DOSDECIMALES)
    Set rsCliente = gcn.OpenResultset( _
        "select C.* from Cliente C JOIN Unidad U ON U.CveCliente = C.CveCliente " & _
        "where U.CveUnidad = " & txtUnidad.Text, rdOpenKeyset, rdConcurRowVer)
    If rsCliente.EOF Then
        sngPrecioTallerExterno = 0
    Else
        sngPrecioTallerExterno = Format(sngCostoTallerExterno * (1 + rsCliente!FactorCostoTallerExterno / 100), DOSDECIMALES)
    End If
    rsCliente.Close
Else
    sngCostoTallerExterno = 0
    sngPrecioTallerExterno = 0
End If

' Refresca el query y reposiciona
rsODT.Requery
PosicionaRegistro (txtCveODT.Text)

' Borra tareas predictivas en dado caso de que se hayan realizado en esta orden
strSQL = "delete from UnidadPredictivo where CveUnidad = " & txtUnidad.Text
strSQL = strSQL & " and CveTarea in (select CveTarea from ODTDetalle where CveODT = " & txtCveODT.Text & ")"
gcn.Execute strSQL

'----------------------------------------------------
'  Actualiza el Kardex con los datos de esta orden
'----------------------------------------------------
strSQL = "select * from ODT where CveODT = " & txtCveODT.Text
Set rsDatosODT = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)

If gblnPreventivosIndividuales Then
    strSQL = "select * from ODTDetalle where CveODT = " & txtCveODT.Text
    strSQL = strSQL & " and CveTarea <> " & TAREACARGACOMBUSTIBLE
    Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    Do While Not rsDetalle.EOF
        strSQL = "select * from ODTDetalleMecanico where CveODT = " & txtCveODT.Text & " and CveTarea = " & rsDetalle!CveTarea
        Set rsMecanico = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
        If rsMecanico.EOF Then
            lngCveMecanico = 0
        Else
            lngCveMecanico = rsMecanico!CveMecanico
        End If
        rsMecanico.Close
        
        strSQL = "insert into UnidadKardex "
        strSQL = strSQL & " (CveUnidad, CveTarea, CveBase, CveODT, CveProveedor, CveMecanico,"
        strSQL = strSQL & "  FechaOcurrencia, KmsAcumulados, PrecioRefacciones, PrecioManoDeObra,"
        strSQL = strSQL & "  PrecioTallerExterno, CveRazonReparacion, Comentarios) "
        strSQL = strSQL & " values (" & txtUnidad.Text & ","
        strSQL = strSQL & rsDetalle!CveTarea & "," & gintCveBase & "," & txtCveODT.Text & ","
        strSQL = strSQL & rsDatosODT!CveProveedor & "," & lngCveMecanico & ",'"
        strSQL = strSQL & Format(rsDatosODT!FechaTerminacion, FECHAMMDDYYYY & HORAMINUTOS) & "'," & rsDatosODT!KmsAcumulados & ","
        strSQL = strSQL & rsDetalle!PrecioRefacciones & "," & rsDetalle!PrecioManoDeObra & ","
        strSQL = strSQL & rsDetalle!PrecioTallerExterno & "," & rsDatosODT!CveRazonReparacion & ",'"
        strSQL = strSQL & rsDetalle!Comentarios & "')"
        gcn.Execute strSQL
        
        '------------------------------------------------
        '   Calcula los costos y precios de la Tarea
        '------------------------------------------------
        strSQL = "select MC.CostoHoraHombre, TP.TiempoEstandarHombre, C.FactorCostoManoDeObra "
        strSQL = strSQL & " from Tarea T, MecanicoCategoria MC, Unidad U, "
        strSQL = strSQL & " TareaPeriodicidad TP, Cliente C ,ODT ODT"
        strSQL = strSQL & " where T.CveTarea = " & rsDetalle!CveTarea
        strSQL = strSQL & " and T.CveTarea = TP.CveTarea "
        strSQL = strSQL & " and T.CveMecanicoCategoria = MC.CveMecanicoCategoria "
        strSQL = strSQL & " and ODT.CveODT = " & rsDetalle!CveODT
        strSQL = strSQL & " and ODT.CveUnidad = U.CveUnidad "
        strSQL = strSQL & " and TP.CveUnidadTipo = U.CveUnidadTipo "
        strSQL = strSQL & " and U.CveCliente = C.CveCliente "
        Set rsCostos = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
        If Not rsCostos.EOF Then
            rsDetalle.Edit
            rsDetalle!CostoManoDeObra = Format(rsCostos!CostoHoraHombre * (rsCostos!TiempoEstandarHombre / 60), DOSDECIMALES)
            rsDetalle!PrecioManoDeObra = Format(rsDetalle!CostoManoDeObra * (1 + rsCostos!FactorCostoManoDeObra / 100), DOSDECIMALES)
            rsDetalle.Update
        End If
        rsCostos.Close
        
        rsDetalle.MoveNext
    Loop
Else
    strSQL = "select CveTarea from Tarea where CveTarea in "
    strSQL = strSQL & "(select distinct CveTareaPadre from ODTDetalle "
    strSQL = strSQL & " where CveODT = " & txtCveODT.Text & " and CveTareaPadre <> 0) "
    strSQL = strSQL & " or CveTarea in "
    strSQL = strSQL & "(select CveTarea from ODTDetalle "
    strSQL = strSQL & " where CveODT = " & txtCveODT.Text & ")"
    
    strSQL = "SELECT DISTINCT CveTarea FROM Tarea "
    'Tareas Padres que pueden no haberse incluido en la ODT
    strSQL = strSQL & "where CveTarea in (SELECT distinct CveTareaPadre " & _
                                        "FROM ODTDetalle " & _
                                        "WHERE CveODT = " & txtCveODT.Text & " and CveTareaPadre <> 0) "
    'Tareas del detalle de la OD
    strSQL = strSQL & "OR CveTarea in (SELECT CveTarea " & _
                                    "FROM ODTDetalle " & _
                                    "WHERE CveODT = " & txtCveODT.Text & ") "
    'Inlcuir servicios hijos que no hayan sido especificados en la ODT
    strSQL = strSQL & "OR CveTarea in (SELECT T.CveTarea " & _
                                        "FROM Tarea T join TareaSubTarea TST on T.CveTarea = TST.CveSubTarea " & _
                                                     "join ODTDetalle ODT ON ODT.CveTarea = TST.CveTarea " & _
                                        "WHERE T.CvePeriodicidadTipo IN(" & PERIODICIDADSERVICIO & "," & PERIODICIDADSERVICIOCADAXKMS & ")" & _
                                        "  AND ODT.CveODT = " & txtCveODT.Text & ")"
    
    
    
    Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    Do While Not rsDetalle.EOF
        strSQL = "select * from ODTDetalleMecanico where CveODT = " & txtCveODT.Text & " and CveTarea = " & rsDetalle!CveTarea
        Set rsMecanico = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
        If rsMecanico.EOF Then
            lngCveMecanico = 0
        Else
            lngCveMecanico = rsMecanico!CveMecanico
        End If
        rsMecanico.Close
        
        strSQL = "insert into UnidadKardex "
        strSQL = strSQL & " (CveUnidad, CveTarea, CveBase, CveODT, CveProveedor, CveMecanico,"
        strSQL = strSQL & "  FechaOcurrencia, KmsAcumulados, PrecioRefacciones, PrecioManoDeObra,"
        strSQL = strSQL & "  PrecioTallerExterno, CveRazonReparacion, Comentarios) "
        strSQL = strSQL & " values (" & txtUnidad.Text & ","
        strSQL = strSQL & rsDetalle!CveTarea & "," & gintCveBase & "," & txtCveODT.Text & ","
        strSQL = strSQL & rsDatosODT!CveProveedor & "," & lngCveMecanico & ",'"
        strSQL = strSQL & Format(rsDatosODT!FechaTerminacion, FECHAMMDDYYYY & HORAMINUTOS) & "'," & rsDatosODT!KmsAcumulados & ","
        strSQL = strSQL & 0 & "," & 0 & ","
        strSQL = strSQL & 0 & "," & rsDatosODT!CveRazonReparacion & ","
        strSQL = strSQL & "' ')"
        gcn.Execute strSQL
        
            
        rsDetalle.MoveNext
    Loop
End If

rsDetalle.Close
rsDatosODT.Close

' Actualiza datos de la orden
rsODT.Edit
rsODT!CveODTEstatus = ESTATUSCERRADA
rsODT!CostoTallerExterno = sngCostoTallerExterno
If gcurImpuestoExento > 0 Then
    rsODT!PrecioTallerExterno = Format(gcurExencionSubtotal, DOSDECIMALES)
    rsODT!IvaPrecioTotal = gcurImpuestoGravado
    rsODT!ImpuestoExento = gcurImpuestoExento
Else
    rsODT!PrecioTallerExterno = Format(sngPrecioTallerExterno, DOSDECIMALES)
    rsODT!IvaPrecioTotal = Format((sngPrecioTallerExterno * ((gcurImpuesto + 100) / 100)) - sngPrecioTallerExterno, DOSDECIMALES)
    rsODT!ImpuestoExento = 0
End If
rsODT.Update
rsODT.Requery

Set rsUnidad = gcn.OpenResultset("SELECT CveUnidad FROM Unidad " & _
                                 "where CveUnidad = " & txtUnidad.Text & _
                                 " and CveUnidadEstatus= " & STATUSREALIZANDOMTTO, rdOpenKeyset, rdConcurRowVer)
strSQL = "exec Sp_SIMActualizaEficiencia @CveUnidad = " & txtUnidad.Text & "," & _
                                    "@intEstatus = " & STATUSMTTOFINALIZADO & ", " & _
                                    "@lngLlegadaODT = 0," & _
                                    "@chrTipoEvento = 'Agregar'"
If gintCveBase = BASETALLERCENTRAL And Not rsUnidad.EOF Then gcn.Execute strSQL
rsUnidad.Close

'----------------------------------------------
' Rutinas para interfases a ERP
'----------------------------------------------
If gblnInterfaseERP Then
    blnERP = True
    gcn.Execute "EXEC sp_SIMEnviaERP " & _
                "@CveBase=" & gintCveBase & _
                ",@CveLlegadaODT =" & gstrCveCotizacion & _
                ",@Tipo='C'"

    If gsngCostoTallerExterno > 0 Then
            gcn.Execute "EXEC sp_SIMEnviaERP " & _
                        "@CveBase=" & gintCveBase & _
                        ",@CveLlegadaODT =" & gstrCveCotizacion & _
                        ",@Tipo='T'"
    End If
    blnERP = False
End If

PosicionaRegistro (txtCveODT.Text)
If rsODT.EOF And rsODT.RowCount > 0 Then rsODT.MoveLast

' Refresca la pantalla
CargaControlesdeResultset

' Rutina para preparar toolbar
ToolBar_EstadoBrowse tlbODT

Exit Sub

Err_CerrarODT:
  Screen.MousePointer = vbDefault
  MsgBox "Error al Cerrar la ODT " & Error, vbInformation + vbOKOnly, "CerrarODT"
  If blnERP Then
    Err.Clear
    Resume Next
  End If
  mblnEdicion = False
  Exit Sub

End Sub

Public Function ValidaCierreODT()

On Error GoTo err_ValidaCierre

Dim strSQL As String
Dim rsTareas As rdoResultset
Dim rsMecanicos As rdoResultset

ValidaCierreODT = True

' Verifica el estatus de la ODT
If cboEstatus.ItemData(cboEstatus.ListIndex) > ESTATUSABIERTA Then
    MsgBox "La ODT ya esta cerrada"
    Screen.MousePointer = vbDefault
    ValidaCierreODT = False
    Exit Function
End If

If VerificaProveedorExterno(rsODT!CveProveedor) <> PROVEEDOREXTERNO Then
    strSQL = "select O.CveODT,O.CveTarea " & _
        "from ODTDetalle O " & _
            "LEFT JOIN ODTDetalleMecanico ODTM ON O.CveODT = ODTM.CveODT AND O.CveTarea = ODTM.CveTarea " & _
        "WHERE dbo.ODTDetalle_FUNCION_TotalHijas(O.CveODT,O.CveTarea) = 0" & _
        "  AND ODTM.CveMecanico IS NULL" & _
        "  AND O.CveODT = " & gstrCveCotizacion
    Set rsTareas = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    If Not rsTareas.EOF Then
        MsgBox "Debes asignar Mecánicos a todas las Tareas antes de cerrar la ODT"
        Screen.MousePointer = vbDefault
        rsTareas.Close
        ValidaCierreODT = False
        Exit Function
    End If
    rsTareas.Close
End If

frmActualizaRetrabajo.Tag = gstrCveCotizacion
If frmActualizaRetrabajo.sprTareas.DataRowCnt = 0 Then
    Unload frmActualizaRetrabajo
Else
    MsgBox "Se requiere el dictamen de los retrabajos detectados.", vbExclamation, "Valida CierreODT"
    frmActualizaRetrabajo.Show vbModal
    strSQL = "EXEC ODTDetalle_PROCESO_Valida @TipoValidacion = 3,@CveODT= " & gstrCveCotizacion
    Set rsTareas = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    If Not rsTareas.EOF Then
        MsgBox "Aun se requiere revisar el dictamen de retrabajo detectados.", vbExclamation, "Valida CierreODT"
        Screen.MousePointer = vbDefault
        rsTareas.Close
        ValidaCierreODT = False
        Exit Function
    End If
    rsTareas.Close
End If


Exit Function

err_ValidaCierre:
    Screen.MousePointer = vbDefault
    MsgBox "Error al Validar el cierre de la ODT" & Error, vbCritical
    ValidaCierreODT = False
    
End Function



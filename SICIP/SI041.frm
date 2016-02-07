VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "Tab32x20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmValeHerramienta 
   Caption         =   "Vale de Herramientas"
   ClientHeight    =   9756
   ClientLeft      =   336
   ClientTop       =   456
   ClientWidth     =   12900
   HelpContextID   =   10
   Icon            =   "SI041.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9756
   ScaleWidth      =   12900
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tlbODT 
      Height          =   396
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   13752
      _ExtentX        =   24257
      _ExtentY        =   699
      ButtonWidth     =   635
      ButtonHeight    =   572
      Appearance      =   1
      ImageList       =   "imgIconos"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   17
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   240
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Agregar"
            Object.ToolTipText     =   "Agregar Vale"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Guardar Vale"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar Vale"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   250
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar Modificaciones"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   250
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Autoriza"
            Object.ToolTipText     =   "Devolver Vale"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   100
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   350
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   200
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ComctlLib.StatusBar staEstatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   9456
      Width           =   12900
      _ExtentX        =   22754
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
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   14563
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "30/01/2016"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "11:47 a.m."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabproLib.vaTabPro TabPrincipal 
      Height          =   8775
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   12675
      _Version        =   131072
      _ExtentX        =   22357
      _ExtentY        =   15478
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      ForeColor       =   16777215
      TabHeight       =   500
      TabsPerRow      =   3
      TabCount        =   3
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
      BookCornerGuardWidth=   108
      BookCornerGuardLength=   408
      MouseIcon       =   "SI041.frx":030A
      ThreeDOuterWidthActive=   2
      DrawFocusRect   =   1
      TabCaption      =   "SI041.frx":0326
      Begin VB.Frame fraCotizacion 
         Caption         =   "Generales"
         Height          =   3372
         Left            =   3588
         TabIndex        =   29
         Top             =   720
         Width           =   8892
         Begin VB.TextBox txtNombre 
            Height          =   315
            Left            =   1560
            TabIndex        =   34
            Top             =   600
            Width           =   7092
         End
         Begin VB.TextBox txtUsuario 
            Height          =   315
            Left            =   1560
            TabIndex        =   33
            Text            =   "MIGUEL"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   1392
            Left            =   1560
            MultiLine       =   -1  'True
            TabIndex        =   32
            Top             =   1800
            Width           =   7092
         End
         Begin VB.TextBox txtVale 
            Height          =   315
            Left            =   1560
            TabIndex        =   31
            Top             =   240
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   1560
            TabIndex        =   30
            Top             =   960
            Width           =   1455
            _ExtentX        =   2561
            _ExtentY        =   550
            _Version        =   393216
            Format          =   50855937
            CurrentDate     =   42306
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Vale:"
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
            Index           =   2
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nombre:"
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
            Index           =   8
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Observaciones:"
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
            Index           =   9
            Left            =   120
            TabIndex        =   37
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fecha:"
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
            Index           =   12
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Usuario:"
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
            Index           =   13
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Width           =   975
         End
      End
      Begin VB.Frame fraPartidas 
         Caption         =   "Partidas"
         Height          =   4572
         Left            =   3588
         TabIndex        =   21
         Top             =   4080
         Width           =   8892
         Begin VB.CommandButton cmdCancelar 
            Height          =   315
            Left            =   7560
            Picture         =   "SI041.frx":05A2
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   360
            Width           =   315
         End
         Begin VB.ComboBox cboArticulos 
            Height          =   288
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   7452
         End
         Begin VB.CommandButton cmdAgregar 
            Height          =   495
            Left            =   8160
            Picture         =   "SI041.frx":095F
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscarMecanico 
            Height          =   315
            Left            =   7560
            Picture         =   "SI041.frx":0DA1
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   360
            Width           =   315
         End
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   7452
         End
         Begin FPSpread.vaSpread sprPartidas 
            Height          =   3492
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   8532
            _Version        =   393216
            _ExtentX        =   15049
            _ExtentY        =   6159
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpreadDesigner  =   "SI041.frx":0FC8
         End
      End
      Begin VB.Frame fraSalida 
         Caption         =   "Generales"
         Enabled         =   0   'False
         Height          =   8052
         Left            =   -24491
         TabIndex        =   3
         Top             =   -20651
         Width           =   12372
         Begin VB.TextBox txtValeAlmacenNumero 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   4
            Top             =   240
            Width           =   1452
         End
         Begin VB.TextBox txtValeAlmacenObservaciones 
            Enabled         =   0   'False
            Height          =   1275
            Left            =   1560
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   1800
            Width           =   10572
         End
         Begin VB.TextBox txtValeAlmacenUsuario 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   7
            Text            =   "MIGUEL"
            Top             =   1320
            Width           =   1452
         End
         Begin VB.TextBox txtValeAlmacenNombre 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   5
            Top             =   600
            Width           =   10572
         End
         Begin VB.Frame fraSalidas 
            Caption         =   "Movimientos de Salidas"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   972
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   3240
            Width           =   12096
            Begin VB.OptionButton optSalida 
               Caption         =   "Consumos internos"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   12
               Tag             =   "30"
               Top             =   600
               Width           =   2475
            End
            Begin VB.OptionButton optSalida 
               Caption         =   "Consumos Proyectos"
               Height          =   255
               Index           =   3
               Left            =   3360
               TabIndex        =   16
               Tag             =   "23"
               Top             =   600
               Width           =   2490
            End
            Begin VB.OptionButton optSalida 
               Caption         =   "Traspaso de Almacén"
               Height          =   255
               Index           =   2
               Left            =   3360
               TabIndex        =   14
               Tag             =   "20"
               Top             =   240
               Width           =   2610
            End
            Begin VB.OptionButton optSalida 
               Caption         =   "Ventas de mostrador"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   10
               Tag             =   "21"
               Top             =   200
               Width           =   2295
            End
            Begin VB.ComboBox cboProyectos 
               Height          =   288
               Index           =   0
               Left            =   6000
               TabIndex        =   20
               Text            =   "Seleccionar Proyecto"
               Top             =   600
               Visible         =   0   'False
               Width           =   3132
            End
            Begin VB.ComboBox cboAlmacen 
               Height          =   288
               Index           =   0
               Left            =   6000
               TabIndex        =   18
               Text            =   "cboAlmacen"
               Top             =   240
               Visible         =   0   'False
               Width           =   3132
            End
         End
         Begin MSComCtl2.DTPicker dtpFechaSalida 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   6
            Top             =   960
            Width           =   1455
            _ExtentX        =   2561
            _ExtentY        =   550
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   50855937
            CurrentDate     =   42306
         End
         Begin FPSpread.vaSpread sprValeAlmacenDetalle 
            Height          =   3492
            Left            =   120
            TabIndex        =   22
            Top             =   4440
            Width           =   12132
            _Version        =   393216
            _ExtentX        =   21399
            _ExtentY        =   6159
            _StockProps     =   64
            Enabled         =   0   'False
            ArrowsExitEditMode=   -1  'True
            ColHeaderDisplay=   1
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   7
            MaxRows         =   200
            ProcessTab      =   -1  'True
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "SI041.frx":11D8
            UserResize      =   1
            VisibleCols     =   6
            VisibleRows     =   10
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Usuario:"
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
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fecha:"
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
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Observaciones:"
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
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nombre:"
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
            Index           =   4
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Vale:"
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
            Index           =   5
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
      End
      Begin ComctlLib.TreeView treeview1 
         Height          =   7872
         Left            =   108
         TabIndex        =   40
         Top             =   720
         Width           =   3312
         _ExtentX        =   5842
         _ExtentY        =   13885
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblLetrero 
         BackStyle       =   0  'Transparent
         Caption         =   "Operación del Taller"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   -17894
         TabIndex        =   43
         Top             =   -16454
         Width           =   1845
      End
      Begin VB.Label lblLetrero 
         BackStyle       =   0  'Transparent
         Caption         =   "Estadísticas"
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   -20879
         TabIndex        =   42
         Top             =   -16469
         Width           =   1260
      End
      Begin VB.Label lblLetrero 
         BackStyle       =   0  'Transparent
         Caption         =   "Información Adicional"
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   -24209
         TabIndex        =   41
         Top             =   -16469
         Width           =   1845
      End
   End
   Begin VB.Image imgMas 
      Height          =   108
      Left            =   13680
      Picture         =   "SI041.frx":1637
      Top             =   0
      Visible         =   0   'False
      Width           =   108
   End
   Begin VB.Image imgMenos 
      Height          =   108
      Left            =   13440
      Picture         =   "SI041.frx":1775
      Top             =   0
      Visible         =   0   'False
      Width           =   108
   End
   Begin ComctlLib.ImageList imgIconos 
      Left            =   13800
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   43
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":18B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":1BCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":1EE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":2201
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":251B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":2835
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":2B4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":2E69
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":3183
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":349D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":37B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":3AD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":3DEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":4105
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":441F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":4739
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":4A53
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":4D6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":5087
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":53A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":56BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":59D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":5CEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":6009
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":6323
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":663D
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":6957
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":6C71
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":6F8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":72A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":75BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":78D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":7BF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":7F0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":8227
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":8541
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":885B
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":8B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":8E8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":91A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":94C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":9CDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI041.frx":9FF7
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmValeHerramienta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1


' Constantes del Tab
Const TABVALEHERRAMIENTAS = 0
Const TABVALEALMACEN = 1
Const TABREPORTES = 2
Const TABDEFAULT = 1 ' Salida Almacen

' Constantes de las prioridades
Const PRIORIDADALTA = 1
Const PRIORIDADMEDIA = 2
Const PRIORIDADBAJA = 3

' Constantes del Spread
Const COLUMNAARTICULO = 1
Const COLUMNANOMBRE = 2
Const COLUMNACANTIDAD = 3

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
'Dim mbolGuardando As Boolean ' Indica si actualmente se tiene un tab con registros pendientes por guardar
Dim mintTabGuardando As Integer ' Indica el Tab en el que se esta trabajando para guardar registros

Dim mlngCveTarea As Long
Dim mblnCambioSprTareas As Boolean
Dim mdatUltimaHoraEjecucion As Date

Public blnReabrioODT As Boolean

Private Sub ActualizaTree()
Dim rsConsulta As rdoResultset
Dim nodx As Node
Dim strSQL As String
Dim strAnterior As String

treeview1.Nodes.Clear
strAnterior = ""

strSQL = "SELECT Nombre,CveValeHerramienta FROM ValeHerramienta WHERE CveValeHerramientaEstatus in( 1,2) ORDER BY Nombre,CveValeHerramienta"
Set rsConsulta = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
Do Until rsConsulta.EOF

    If strAnterior <> rsConsulta!Nombre Then
        Set nodx = treeview1.Nodes.Add(, , Mid(rsConsulta!Nombre, 2), rsConsulta!Nombre)
        nodx.EnsureVisible
        strAnterior = rsConsulta!Nombre
    End If
    
    Do Until rsConsulta!Nombre <> strAnterior
        Set nodx = treeview1.Nodes.Add(Mid(rsConsulta!Nombre, 2), tvwChild, "V-" & CStr(rsConsulta!CveValeHerramienta), CStr(rsConsulta!CveValeHerramienta))
        nodx.EnsureVisible
        rsConsulta.MoveNext
        If rsConsulta.EOF Then Exit Do
    Loop
    If rsConsulta.EOF Then Exit Do
Loop
rsConsulta.Close

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
Dim X As Boolean

On Error GoTo Err_DespliegaDetalle

' Limpia el spread
'LimpiaBloque sprTareas, 1, 1, sprTareas.MaxRows, sprTareas.MaxCols
sprPartidas.MaxRows = 0

' Crea el rdoResultset de la tabla de ODTDetalle
strSQL = "select VA.CveValeHerramienta,VA.CveArticulo,A.Nombre,A.Codigo,VA.Cantidad,DV.CantidadRegresada " & _
    "from ValeHerramientaDetalle VA " & _
    "  JOIN Articulo A ON A.CveArticulo = VA.CveArticulo " & _
    " LEFT JOIN (SELECT VD.CveValeHerramienta,VDD.CveArticulo,SUM(VDD.Cantidad) CantidadRegresada " & _
                "FROM DevolucionHerramienta VD " & _
                    "JOIN DevolucionHerramientaDetalle VDD ON VDD.CveDevolucionHerramienta = VD.CveDevolucionHerramienta " & _
                "group by VD.CveValeHerramienta,VDD.CveArticulo) DV ON VA.CveValeHerramienta = DV.CveValeHerramienta AND DV.CveArticulo = VA.CveArticulo " & _
    "WHERE VA.CveValeHerramienta = " & glngCveCotizacion & _
    "ORDER BY A.Nombre"

Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)

sprPartidas.MaxRows = rsDetalle.RowCount
' Llena el spread de Tareas
intRenglon = 1
sprPartidas.ReDraw = False
Do Until rsDetalle.EOF

    sprPartidas.Row = intRenglon
    
    sprPartidas.Col = 1
    sprPartidas.Text = rsDetalle!Codigo
    sprPartidas.TypeHAlign = TypeHAlignLeft
    ProtegeCelda sprPartidas, sprPartidas.Row, 1, True
    
    
    sprPartidas.Col = 2
    sprPartidas.Text = rsDetalle!Nombre
    sprPartidas.TypeHAlign = TypeHAlignLeft
    sprPartidas.Text = rsDetalle!Nombre
    ProtegeCelda sprPartidas, sprPartidas.Row, 2, True
    
    sprPartidas.Col = 3
    sprPartidas.CellType = CellTypeNumber
    sprPartidas.TypeNumberDecPlaces = 0
    sprPartidas.Value = rsDetalle!Cantidad
    ProtegeCelda sprPartidas, sprPartidas.Row, 3, True

    sprPartidas.Col = 4
    sprPartidas.CellType = CellTypeNumber
    sprPartidas.TypeNumberDecPlaces = 0
    If IsNull(rsDetalle!CantidadRegresada) Then
        sprPartidas.Value = 0
    Else
        sprPartidas.Value = rsDetalle!CantidadRegresada
    End If
    ProtegeCelda sprPartidas, sprPartidas.Row, 4, True

    sprPartidas.Col = 5
    sprPartidas.Text = rsDetalle!CveArticulo
    ProtegeCelda sprPartidas, sprPartidas.Row, 5, True
    
    rsDetalle.MoveNext
    intRenglon = intRenglon + 1
Loop
rsDetalle.Close
sprPartidas.ReDraw = True
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
Dim lngRenglon As Long
Dim lngValor As Long
Dim blnExiste As Boolean
Dim blnCumplio As Boolean
Dim i As Integer
Dim intMovimientoSalida As Integer
Dim intCboSeleccionado As Integer

ValidaCampos = False

Select Case TabPrincipal.ActiveTab
    Case TABVALEHERRAMIENTAS

        If Len(txtNombre.Text) = 0 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes proporcionar un nombre", vbExclamation
          txtNombre.SetFocus
          Exit Function
        End If
                
        If Len(txtUsuario.Text) = 0 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes proporcionar un usuario", vbExclamation
          txtUsuario.SetFocus
          Exit Function
        End If
                                 
        If sprPartidas.DataRowCnt = 0 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes especificar cuales herramientas se estan prestando", vbExclamation
          txtBuscar.SetFocus
          Exit Function
        End If
                           
        blnExiste = False
       For lngRenglon = 1 To sprPartidas.DataRowCnt
            blnCumplio = sprPartidas.GetInteger(3, lngRenglon, lngValor)
            If lngValor = 0 Then
                blnExiste = True
                Exit For
            End If
        Next lngRenglon
        If blnExiste Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes especificar en la partida " & lngRenglon & " la cantidad que se esta prestando", vbExclamation
          sprPartidas.SetFocus
          Exit Function
        End If
    Case TABVALEALMACEN
        
        If Len(txtValeAlmacenNumero) = 0 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes proporcionar un folio", vbExclamation
          txtValeAlmacenNumero.SetFocus
          Exit Function
        End If
        
        If Len(txtValeAlmacenNombre) = 0 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes proporcionar un nombre", vbExclamation
          txtValeAlmacenNombre.SetFocus
          Exit Function
        End If
        
        If Len(txtValeAlmacenUsuario) = 0 Then
          Screen.MousePointer = vbDefault
          MsgBox "Debes proporcionar un usuario", vbExclamation
          txtValeAlmacenUsuario.SetFocus
          Exit Function
        End If

        For i = 0 To optSalida.Count - 1
            If optSalida(i).Value Then
                intMovimientoSalida = i
            End If
        Next i
        
        If intMovimientoSalida >= 0 Then
            If intMovimientoSalida = 2 And cboAlmacen(0).ListIndex < 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Debes seleccionar un Almacen", vbExclamation
                optSalida(0).SetFocus
                Exit Function
            ElseIf intMovimientoSalida = 3 And cboProyectos(0).ListIndex < 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Debes seleccionar un Proyecto", vbExclamation
                optSalida(0).SetFocus
                Exit Function
            End If
        Else
          Screen.MousePointer = vbDefault
          MsgBox "Debes seleccionar un movimiento de salida de almacen", vbExclamation
          optSalida(0).SetFocus
          Exit Function
        End If

End Select

ValidaCampos = True

End Function

Private Sub CargardoResultsetDeControles()

Dim strSQL As String
Dim strNumFactura As String
Dim i As Integer
Dim strVale As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim lngCveArticulo As Long
Dim intCantidad As Integer
Dim strNoFab As String
Dim strNumParte As String
Dim strDescipcion As String
Dim curPrecio As Currency
Dim strCveParte As String
Dim intMovimientoSalida As Integer
Dim intCveObra As Integer
Dim intCveAlmacen As Integer



On Error GoTo Err_CargaRSet

Screen.MousePointer = vbHourglass

Select Case TabPrincipal.ActiveTab
    Case TABVALEHERRAMIENTAS
        If txtVale.Text = "" Then
            strVale = "NULL"
        Else
            strVale = txtVale.Text
        End If
        
        strSQL = "EXEC ValeHerramienta_PROCESO_ActualizaBeta " & _
            "@ValeHerramienta = NULL" & _
            ",@Nombre ='" & txtNombre & "'" & _
            ",@CveUsuario='" & txtUsuario.Text & "'" & _
            ",@Observaciones='" & txtObservaciones & "'"
                
        strSQL = "'<O Nombre=""" & txtNombre.Text & """ Usuario=""" & txtUsuario.Text & """ Obs=""" & txtObservaciones.Text & """>"
        strSQL2 = strSQL
        strSQL3 = strSQL

        For i = 1 To sprPartidas.DataRowCnt
            sprPartidas.Row = i
            sprPartidas.Col = 5
            lngCveArticulo = Val(sprPartidas.Text)
            sprPartidas.Col = 3
            intCantidad = Val(sprPartidas.Text)

            If Len(strSQL) > 7800 Then
                If Len(strSQL2) > 7800 Then
                    strSQL3 = strSQL3 & "<D A=""" & lngCveArticulo & """ C=""" & intCantidad & """>"
                Else
                    strSQL2 = strSQL2 & "<D A=""" & lngCveArticulo & """ C=""" & intCantidad & """>"
                End If
            Else
                strSQL = strSQL & "<D A=""" & lngCveArticulo & """ C=""" & intCantidad & """/>"
            End If
        Next i
        strSQL = strSQL & "</O>'"
        strSQL2 = strSQL2 & "</O>'"
        strSQL3 = strSQL3 & "</O>'"

        gcn.Execute "EXEC ValeHerramienta_PROCESO_ActualizaBeta @ValeHerramienta=" & strVale & _
            ",@Fecha='" & Format(dtpFecha.Value, "YYYY-MM-DD") & "'" & _
            ",@Nombre ='" & txtNombre & "'" & _
            ",@CveUsuario='" & txtUsuario.Text & "'" & _
            ",@Observaciones='" & txtObservaciones & "'" & _
            ",@XML=" & strSQL & ",@XML2=" & strSQL2 & ",@XML3=" & strSQL3
    
    Case TABVALEALMACEN
        If txtValeAlmacenNumero.Text = "" Then
            strVale = "NULL"
        Else
            strVale = txtValeAlmacenNumero.Text
        End If

        intMovimientoSalida = -1
        For i = 0 To optSalida.Count - 1
            If optSalida(i).Value Then
                intMovimientoSalida = i
            End If
        Next i
        intCveObra = cboProyectos(0).ListIndex
        intCveAlmacen = cboAlmacen(0).ListIndex
        
        strSQL = "'<O Nombre=""" & txtValeAlmacenNombre.Text & """ Usuario=""" & txtValeAlmacenUsuario.Text & """ Obs=""" & txtValeAlmacenObservaciones.Text & """>"
        strSQL2 = strSQL
        strSQL3 = strSQL

        For i = 1 To sprValeAlmacenDetalle.DataRowCnt
            sprValeAlmacenDetalle.Row = i

            sprValeAlmacenDetalle.Col = 1
            strNoFab = sprValeAlmacenDetalle.Text

            sprValeAlmacenDetalle.Col = 2
            strNumParte = sprValeAlmacenDetalle.Text

            sprValeAlmacenDetalle.Col = 3
            strDescipcion = sprValeAlmacenDetalle.Text

            sprValeAlmacenDetalle.Col = 4
            intCantidad = Val(sprValeAlmacenDetalle.Text)

            sprValeAlmacenDetalle.Col = 5
            curPrecio = sprValeAlmacenDetalle.Text

            sprValeAlmacenDetalle.Col = 6
            strCveParte = sprValeAlmacenDetalle.Text
            
            If Len(strSQL) > 7800 Then
                If Len(strSQL2) > 7800 Then
                    strSQL3 = strSQL3 & "<D F=""" & strNoFab & """ N=""" & strNumParte & """ D=""" & strDescipcion & """ C=""" & intCantidad & """ P=""" & curPrecio & """ CP=""" & strCveParte & """>"
                Else
                    strSQL2 = strSQL2 & "<D F=""" & strNoFab & """ N=""" & strNumParte & """ D=""" & strDescipcion & """ C=""" & intCantidad & """ P=""" & curPrecio & """ CP=""" & strCveParte & """>"
                End If
            Else
                strSQL = strSQL & "<D F=""" & strNoFab & """ N=""" & strNumParte & """ D=""" & strDescipcion & """ C=""" & intCantidad & """ P=""" & curPrecio & """ CP=""" & strCveParte & """/>"
            End If
        Next i
        strSQL = strSQL & "</O>'"
        strSQL2 = strSQL2 & "</O>'"
        strSQL3 = strSQL3 & "</O>'"

        gcn.Execute "EXEC ValeAlmacen_PROCESO_Actualiza @ValeAlmacen=" & strVale & _
            ",@Fecha='" & Format(dtpFechaSalida.Value, "YYYY-MM-DD") & "'" & _
            ",@Nombre ='" & txtValeAlmacenNombre & "'" & _
            ",@CveUsuario='" & txtUsuario.Text & "'" & _
            ",@Observaciones='" & txtObservaciones & "'" & _
            ",@MovimientoSalida='" & intMovimientoSalida & "'" & _
            ",@CveObra='" & intCveObra & "'" & _
            ",@CveAlmacen='" & intCveAlmacen & "'" & _
            ",@XML=" & strSQL & ",@XML2=" & strSQL2 & ",@XML3=" & strSQL3

        InicializaCampos False
End Select

Screen.MousePointer = vbDefault
Exit Sub

Err_CargaRSet:
    Screen.MousePointer = vbDefault
    MsgBox "Error al Cargar rdoResultset de Controles" & Error, vbCritical
    Exit Sub

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

fraCotizacion.Visible = True
fraPartidas.Visible = True
txtBuscar.Visible = False
cmdBuscarMecanico.Visible = False
cmdAgregar.Visible = False
cboArticulos.Visible = False
cmdCancelar.Visible = False

strSQL = "SELECT * FROM ValeHerramienta WHERE CveValeHerramienta=" & glngCveCotizacion
Set rs = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
If Not rs.EOF Then
    mblnLlena = True
    
    txtVale.Text = rs!CveValeHerramienta
    txtNombre.Text = rs!Nombre
    txtUsuario.Text = rs!NombreAutoriza
    txtObservaciones.Text = rs!Observaciones
    dtpFecha.Value = rs!Fecha
    
    dtpFecha.Enabled = False
    txtUsuario.Enabled = False
    txtVale.Enabled = False
    
    txtNombre.Enabled = True
    txtObservaciones.Enabled = True
    
    ToolBoton_Estado tlbODT, "Actualizar", False

Else
    ' Inicializacion Para rdoResultset vacio
    InicializaCampos True
End If
rs.Close

mblnEdicion = False

' Despliega las Tareas de la orden
DespliegaDetalle
ActualizaTree
ToolBar_EstadoBrowse tlbODT
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
    InicializaCampos True   ' Limpia los controles
    
    mblnEdicion = False
    mblnAlta = True
    ToolBar_EstadoCambio tlbODT
    'txtVale.SetFocus

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

Private Sub Cancela()
'********************************************************************
'  Rutina que prepara la pantalla para cancelar accion
'********************************************************************

    On Error GoTo Err_Cancela
    
    Screen.MousePointer = vbHourglass
      
    InicializaTab TabPrincipal.ActiveTab
    InicializaCampos False
    
    mblnEdicion = False
    mblnAlta = False
    ToolBar_EstadoBrowse tlbODT

Exit_Cancela:
    Screen.MousePointer = vbDefault
    Exit Sub
  
Err_Cancela:
    
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

Private Sub NuevoRegistro(tabNumero As Byte)
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
InicializaCampos True   ' Limpia los controles

mblnEdicion = False
mblnAlta = True
ToolBar_EstadoCambio tlbODT
txtVale.SetFocus

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

Private Sub InicializaCampos(bolEnabled As Boolean)
    Dim intActual As Integer
    intActual = 0
    
    Screen.MousePointer = vbHourglass
   
    'limpia controles para proxima captura
    
    Select Case TabPrincipal.ActiveTab
        Case TABVALEHERRAMIENTAS
            fraCotizacion.Visible = True
            fraPartidas.Visible = True
            txtVale.Enabled = True
            txtNombre.Enabled = True
            dtpFecha.Enabled = True
            txtUsuario.Enabled = True
            txtObservaciones.Enabled = True
            
            txtVale.Text = ""
            txtNombre.Text = ""
            dtpFecha.Value = Format(Now, "DD/MM/YYYY")
            txtUsuario.Text = "MIGUEL"
            txtObservaciones.Text = ""
                       
            'LimpiaBloque sprPartidas, 1, 1, sprPartidas.MaxRows, sprPartidas.MaxCols
            sprPartidas.MaxRows = 0
            
            txtBuscar.Visible = True
            cmdBuscarMecanico.Visible = True

        Case TABVALEALMACEN
            txtValeAlmacenNumero.Text = ""
            txtValeAlmacenNumero.Enabled = bolEnabled
            txtValeAlmacenNombre.Text = ""
            txtValeAlmacenNombre.Enabled = bolEnabled
            dtpFechaSalida.Value = DateValue(Now)
            dtpFechaSalida.Enabled = bolEnabled
            txtValeAlmacenUsuario.Text = "MIGUEL"
            txtValeAlmacenUsuario.Enabled = bolEnabled
            txtValeAlmacenObservaciones.Text = ""
            txtValeAlmacenObservaciones.Enabled = bolEnabled

            Do Until intActual >= optSalida.Count
                optSalida(intActual).Value = False
                intActual = intActual + 1
            Loop

            cboAlmacen(0).Visible = False
            cboAlmacen(0).Clear
            cboAlmacen(0).Text = "Seleccionar Almacén"

            cboProyectos(0).Visible = False
            cboProyectos(0).Clear
            cboProyectos(0).Text = "Seleccionar Proyecto"
            
            fraSalidas(0).Enabled = bolEnabled
            sprValeAlmacenDetalle.ClearRange 1, 1, 6, 200, False

            sprValeAlmacenDetalle.Enabled = bolEnabled
            
            If txtValeAlmacenNumero.Enabled Then
                txtValeAlmacenNumero.SetFocus
            End If
    
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdBuscarMecanico_Click()
Dim strSQL As String
Dim strCondicion As String
Dim i As Integer
Dim lngCveArticulo As Long

Screen.MousePointer = vbHourglass

strCondicion = ""
For i = 1 To sprPartidas.DataRowCnt
    sprPartidas.Row = i
    sprPartidas.Col = 5
    lngCveArticulo = Val(sprPartidas.Text)

    If Len(strCondicion) > 0 Then strCondicion = strCondicion & ","
    strCondicion = strCondicion & lngCveArticulo
Next i

If Len(strCondicion) > 0 Then strCondicion = "AND CveArticulo NOT IN(" & strCondicion & ")"


strSQL = "SELECT A.CveArticulo,A.Nombre + ISNULL('  ('+A.Codigo+')','') Nombre " & _
    "FROM Articulo AS A JOIN Familia F ON A.CveFamilia = F.CveFamilia " & _
    "WHERE A.Activo=1 AND (F.CveFamilia in(9,4) OR F.CveRama = 2) " & _
     " AND (A.Codigo like '%" & Replace(txtBuscar.Text, " ", "%") & "%' OR A.Nombre like '%" & Replace(txtBuscar.Text, " ", "%") & "%') " & _
     strCondicion & " ORDER BY A.Nombre"
LlenaVariosSelectores strSQL, Array("cboArticulos"), Me
If cboArticulos.ListCount > 0 Then
    cboArticulos.Visible = True
    txtBuscar.Visible = False
    cmdBuscarMecanico.Visible = False
    cmdAgregar.Visible = True
    cmdCancelar.Visible = True
    txtBuscar.Text = ""
    
    If cboArticulos.ListCount = 1 Then cboArticulos.ListIndex = 0
End If
Screen.MousePointer = vbDefault
End Sub
Private Sub cmdAgregar_Click()

Dim strSQL As String
Dim rsDetalle As rdoResultset

If cboArticulos.ListIndex >= 0 Then
    
    'Always have the spreadsheet in edit mode
    sprPartidas.EditModePermanent = True

    sprPartidas.MaxRows = sprPartidas.MaxRows + 1
    
    sprPartidas.Row = sprPartidas.MaxRows
    
    strSQL = "select CveArticulo,A.Nombre,A.Codigo " & _
        "from Articulo A " & _
        "Where A.CveArticulo = " & cboArticulos.ItemData(cboArticulos.ListIndex)
    Set rsDetalle = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
    
    ' Llena el spread
    sprPartidas.ReDraw = False
    Do Until rsDetalle.EOF
    
        'MakeFloatCell COLCANTREQUERIDA, COLCANTREQUERIDA, SPRPARTIDAS.Row, SPRPARTIDAS.Row, "-99999", "99999", False, True, 2, 0
        'MakeFloatCell COLUNIDADMEDIDA, COLKGPORM2, SPRPARTIDAS.Row, SPRPARTIDAS.Row, "-99999", "99999", False, True, 2, 0
        'MakeFloatCell COLPRECIOLISTA, COLPRECIOLISTA, SPRPARTIDAS.Row, SPRPARTIDAS.Row, "-99999", "99999", True, True, 2, 0
    
        sprPartidas.Col = 1 'A
        sprPartidas.Text = rsDetalle!Codigo
        sprPartidas.TypeHAlign = TypeHAlignLeft
        ProtegeCelda sprPartidas, sprPartidas.Row, 1, True

        sprPartidas.Col = 2 'C
        sprPartidas.TypeHAlign = TypeHAlignLeft
        sprPartidas.Text = rsDetalle!Nombre
        ProtegeCelda sprPartidas, sprPartidas.Row, 2, True
    
        sprPartidas.Col = 3 'C
        sprPartidas.CellType = CellTypeNumber
        sprPartidas.TypeNumberDecPlaces = 0
        ProtegeCelda sprPartidas, sprPartidas.Row, 3, False
        
        sprPartidas.Col = 4 'C
        sprPartidas.CellType = CellTypeNumber
        sprPartidas.TypeNumberDecPlaces = 0
        ProtegeCelda sprPartidas, sprPartidas.Row, 4, True
    
        sprPartidas.Col = 5 'D
        sprPartidas.Text = rsDetalle!CveArticulo
        ProtegeCelda sprPartidas, sprPartidas.Row, 5, True
        
        rsDetalle.MoveNext
    Loop
    rsDetalle.Close
    sprPartidas.ReDraw = True
    
    cmdAgregar.Visible = False
    cmdCancelar.Visible = False
    cboArticulos.Visible = False
    
    txtBuscar.Visible = True
    cmdBuscarMecanico.Visible = True
    
End If
End Sub

Private Sub cmdCancelar_Click()
    cmdAgregar.Visible = False
    cmdCancelar.Visible = False
    cboArticulos.Visible = False
    
    txtBuscar.Visible = True
    cmdBuscarMecanico.Visible = True
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
    strTexto = frmValeHerramienta.Caption
    frmValeHerramienta.Caption = ""
    MsgBox App.EXEName & "  Actualmente ejecutándose"
    AppActivate strTexto
    End
End If
'**********************

staEstatusBar.Panels(1).Text = gstrProducto

mdatUltimaHoraEjecucion = DateTime.Now

ActualizaTree

'5   Autorizada por Cliente  AC

sprPartidas.MaxRows = 0
sprPartidas.MaxCols = 5

sprPartidas.ColWidth(1) = 8
sprPartidas.ColWidth(2) = 24
sprPartidas.ColWidth(3) = 4
sprPartidas.ColWidth(4) = 4

sprPartidas.Row = -1000

sprPartidas.Col = 1
sprPartidas.FontBold = True
sprPartidas.TypeHAlign = TypeHAlignCenter
sprPartidas.Text = "Cve"

sprPartidas.Col = 2
sprPartidas.FontBold = True
sprPartidas.TypeHAlign = TypeHAlignCenter
sprPartidas.Text = "Descripcion"

sprPartidas.Col = 3
sprPartidas.FontBold = True
sprPartidas.TypeHAlign = TypeHAlignCenter
sprPartidas.Text = "Cant Prest"

sprPartidas.Col = 4
sprPartidas.FontBold = True
sprPartidas.TypeHAlign = TypeHAlignCenter
sprPartidas.Text = "Cant Reg"

' Carga controles del rdoResultset
'CargaControlesdeResultset

' Despliega el nombre del servidor
'strSQL = "select Nombre from Base Where CveBase = " & gintCveBase
'Set rdBase = gcn.OpenResultset(strSQL, rdOpenKeyset, rdConcurRowVer)
'staEstatusBar.Panels(2).Text = "Ubicación: " & rdBase!Nombre & "    Versión:" & App.Major & "." & App.Minor & "." & App.Revision
'rdBase.Close

' Mostrar el Tab Salida Almacen
TabPrincipal.ActiveTab = TABDEFAULT
InicializaTab TABDEFAULT


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
    Set frmValeHerramienta = Nothing
    CierraConeccion
    End
End Sub

Private Sub optSalida_Click(Index As Integer)
    cboAlmacen(0).Visible = False
    cboAlmacen(0).Clear
    cboAlmacen(0).Text = "Seleccionar Almacén"
    
    cboProyectos(0).Visible = False
    cboProyectos(0).Clear
    cboProyectos(0).Text = "Seleccionar Proyecto"
    
    If Index = 2 Then
        AbreConeccionSalida
        LlenaComboSalida "SELECT CveAlmacen,Nombre FROM Almacen", cboAlmacen(0)
        CierraConeccionSalida
        cboAlmacen(0).Visible = True
    End If
    If Index = 3 Then
        AbreConeccionSalida
        LlenaComboSalida "SELECT CveObra,Nombre FROM Obra", cboProyectos(0)
        CierraConeccionSalida
        cboProyectos(0).Visible = True
    End If
End Sub

Private Sub InicializaTab(TabToActivate As Integer)
    Select Case TabToActivate
        Case Is = 0
            dtpFecha.Value = DateValue(Now)
        Case Is = 1
            dtpFechaSalida = DateValue(Now)
    End Select
End Sub

Private Sub TabPrincipal_TabShown(ActiveTab As Integer)
    Dim intActive As Integer
    intActive = TabPrincipal.ActiveTab

    If Not (mblnEdicion Or mblnAlta) Then
        InicializaTab ActiveTab
        mintTabGuardando = TabPrincipal.ActiveTab
    ElseIf ActiveTab <> mintTabGuardando Then
        TabPrincipal.ActiveTab = mintTabGuardando
        'InicializaTab intActive
    End If

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
        mintTabGuardando = TabPrincipal.ActiveTab

   Case Is = "Actualizar"
        mintTabGuardando = TabPrincipal.ActiveTab
'        If cboEstatus.ItemData(cboEstatus.ListIndex) > ESTATUSABIERTA Then
'            MsgBox "La ODT ya esta cerrada, no se puede modificar"
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
        If ValidaCampos() Then
            If Actualiza() Then
                If TabPrincipal.ActiveTab = 0 Then
                    ActualizaTree
                    'fraCotizacion.Visible = False
                    'fraPartidas.Visible = False
                End If
            
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
        Cancela

   Case Is = "Autoriza"
        'frmDevolucionHerramienta.Show vbModal 'cgml 20160124
        fraCotizacion.Visible = False
        fraPartidas.Visible = False

         
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
Private Sub treeview1_NodeClick(ByVal Node As ComctlLib.Node)

If IsNumeric(Right(Node.Key, 1)) Then
    glngCveCotizacion = Mid(Node.Key, InStr(1, Node.Key, "-") + 1, 40)
    CargaControlesdeResultset
End If

End Sub


Sub Imprimir()

   Dim frmRep As New frmReporte
'


    frmRptVales.Show vbModal

    If Len(gstrSQL) > 0 Then
        frmRep.mstrNombreArchivo = gstrArchivoRpt
        frmRep.mstrSQL = gstrSQL
    frmRep.Show vbModal

    Set frmRep = Nothing

    End If
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdBuscarMecanico_Click
End Sub
Private Sub txtNombre_Change()
If mblnAlta Then Exit Sub
mblnEdicion = True
ToolBar_EstadoCambio tlbODT
End Sub


Private Sub txtObservaciones_Change()
If mblnAlta Then Exit Sub
mblnEdicion = True
ToolBar_EstadoCambio tlbODT
End Sub



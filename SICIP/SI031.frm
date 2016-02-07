VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "Tab32x20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmValeHerramienta 
   Caption         =   "Vale de Herramientas"
   ClientHeight    =   3495
   ClientLeft      =   330
   ClientTop       =   450
   ClientWidth     =   13890
   HelpContextID   =   10
   Icon            =   "SI031.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3495
   ScaleWidth      =   13890
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tlbODT 
      Height          =   420
      Left            =   60
      TabIndex        =   11
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
   Begin TabproLib.vaTabPro TabPrincipal 
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   12675
      _Version        =   131072
      _ExtentX        =   22357
      _ExtentY        =   15478
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
      MouseIcon       =   "SI031.frx":030A
      ThreeDOuterWidthActive=   2
      DrawFocusRect   =   1
      TabCaption      =   "SI031.frx":0326
      Begin VB.Frame fraPartidas 
         Caption         =   "Partidas"
         Height          =   4575
         Left            =   3600
         TabIndex        =   15
         Top             =   4080
         Visible         =   0   'False
         Width           =   8175
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   6615
         End
         Begin VB.CommandButton cmdBuscarMecanico 
            Height          =   315
            Left            =   6720
            Picture         =   "SI031.frx":051E
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   360
            Width           =   315
         End
         Begin VB.CommandButton cmdAgregar 
            Height          =   495
            Left            =   7200
            Picture         =   "SI031.frx":0745
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin FPSpread.vaSpread sprPartidas 
            Height          =   3495
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   7815
            _Version        =   393216
            _ExtentX        =   13785
            _ExtentY        =   6165
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
            SpreadDesigner  =   "SI031.frx":0B87
         End
         Begin VB.ComboBox cboArticulos 
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   6615
         End
         Begin VB.CommandButton cmdCancelar 
            Height          =   315
            Left            =   6720
            Picture         =   "SI031.frx":0D4B
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Frame fraCotizacion 
         Caption         =   "Cotizacion"
         Height          =   3375
         Left            =   3600
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   7335
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16711681
            CurrentDate     =   42306
         End
         Begin VB.TextBox txtVale 
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   1275
            Left            =   1560
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   1800
            Width           =   5415
         End
         Begin VB.TextBox txtUsuario 
            Height          =   315
            Left            =   1560
            TabIndex        =   4
            Text            =   "MIGUEL"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Usuario:"
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
            Index           =   13
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fecha:"
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
            Index           =   12
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Observaciones:"
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
            Index           =   9
            Left            =   120
            TabIndex        =   18
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nombre:"
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
            Index           =   8
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Vale:"
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
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
      End
      Begin ComctlLib.TreeView treeview1 
         Height          =   7755
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   13679
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   -16454
         Width           =   1845
      End
   End
   Begin ComctlLib.StatusBar staEstatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   3195
      Width           =   13890
      _ExtentX        =   24500
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
            Object.Width           =   16272
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "09/11/2015"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "11:30 a.m."
            Key             =   ""
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
   Begin VB.Image imgMas 
      Height          =   135
      Left            =   13680
      Picture         =   "SI031.frx":1108
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgMenos 
      Height          =   135
      Left            =   13440
      Picture         =   "SI031.frx":1246
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
            Picture         =   "SI031.frx":1384
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":169E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":19B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":1CD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":1FEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":2306
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":2620
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":293A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":2C54
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":2F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":3288
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":35A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":38BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":3BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":3EF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":420A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":4524
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":483E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":4B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":4E72
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":518C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":54A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":57C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":5ADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":5DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":610E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":6428
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":6742
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":6A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":6D76
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":7090
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":73AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":76C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":79DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":7CF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":8012
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":832C
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":8646
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":8960
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":8C7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":8F94
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":97AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SI031.frx":9AC8
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
Const TABUNIDADES = 0
Const TABREPORTES = 1

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
Dim x As Boolean

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
'Dim blnTemp As Integer
'Dim strMsgErrValidacion As String
'Dim sngTotalSpread As Single
'Dim sngImporteTotal As Single
'Dim i As Integer
'Dim rsConsulta As rdoResultset
'Dim rsCuenta As rdoResultset
'Dim strSQL As String
'Dim sngSaldo As Single
'Dim lngCveTarea As Long
'Dim strRazon As String
'Dim strTarea As String
Dim lngRenglon As Long
Dim lngValor As Long
Dim blnExiste As Boolean
Dim blnCumplio As Boolean

ValidaCampos = False

Select Case TabPrincipal.ActiveTab
    Case TABUNIDADES

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

On Error GoTo Err_CargaRSet

Screen.MousePointer = vbHourglass

Select Case TabPrincipal.ActiveTab
    Case TABUNIDADES
        
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
    InicializaCampos
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
InicializaCampos   ' Limpia los controles

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
Private Sub InicializaCampos()
    Screen.MousePointer = vbHourglass
   
    'limpia controles para proxima captura
    
    Select Case TabPrincipal.ActiveTab
        Case TABUNIDADES
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
                ActualizaTree
                fraCotizacion.Visible = False
                fraPartidas.Visible = False
            
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
        frmDevolucionHerramienta.Show vbModal
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



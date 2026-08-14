VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmDocu 
   Caption         =   "Consulta de Operaciones"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   ControlBox      =   0   'False
   ForeColor       =   &H00404000&
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.Frame Fracruze 
      Caption         =   "Cruze de Documento(Solo Ventas):"
      Height          =   2175
      Left            =   5880
      TabIndex        =   107
      Top             =   3960
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton CmdCruze_escapar 
         Caption         =   "Escapar"
         Height          =   330
         Left            =   1080
         TabIndex        =   118
         Top             =   1680
         Width           =   2085
      End
      Begin VB.CommandButton CmdCruze_Int 
         Caption         =   "Intercambiar"
         Height          =   855
         Left            =   3600
         Picture         =   "FrmDocu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cFB2 
         ForeColor       =   &H00404000&
         Height          =   315
         ItemData        =   "FrmDocu.frx":E0BE
         Left            =   960
         List            =   "FrmDocu.frx":E0CB
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   1320
         Width           =   510
      End
      Begin VB.ComboBox cFB1 
         ForeColor       =   &H00404000&
         Height          =   315
         ItemData        =   "FrmDocu.frx":E0D8
         Left            =   960
         List            =   "FrmDocu.frx":E0E5
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   840
         Width           =   510
      End
      Begin VB.TextBox Tn2 
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   2160
         TabIndex        =   114
         Text            =   "0"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Tn1 
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   2160
         TabIndex        =   113
         Text            =   "0"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Ts2 
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   1560
         TabIndex        =   112
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Ts1 
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   1560
         TabIndex        =   111
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin MSMask.MaskEdBox tfcruze 
         Height          =   285
         Left            =   960
         TabIndex        =   119
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lcruze 
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   120
         Top             =   480
         Width           =   855
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   3240
         X2              =   3480
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   3480
         X2              =   3480
         Y1              =   960
         Y2              =   1440
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   3240
         X2              =   3480
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3120
         X2              =   3360
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3360
         X2              =   3360
         Y1              =   1080
         Y2              =   1560
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3120
         X2              =   3360
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lcruze 
         Caption         =   "Doc. Nro 2:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   110
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lcruze 
         Caption         =   "Doc. Nro 1:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   109
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lcruze 
         Caption         =   "Documento a Intercambiar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   108
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CheckBox chefg 
      BackColor       =   &H00808000&
      Caption         =   "Forzar Guia R."
      Height          =   495
      Left            =   10560
      TabIndex        =   102
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton DOCANEXO 
      Height          =   555
      Left            =   10590
      TabIndex        =   94
      Top             =   1470
      Width           =   1155
   End
   Begin VB.Frame fracambio 
      Caption         =   "Cambio de Numeros :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1380
      Left            =   120
      TabIndex        =   82
      Top             =   3960
      Visible         =   0   'False
      Width           =   5745
      Begin VB.CommandButton cmdesc 
         Caption         =   "Escapar"
         Height          =   330
         Left            =   4260
         TabIndex        =   92
         Top             =   750
         Width           =   1365
      End
      Begin VB.CommandButton cmdcambiar 
         Caption         =   "Cambiar"
         Height          =   330
         Left            =   4260
         TabIndex        =   91
         Top             =   300
         Width           =   1350
      End
      Begin VB.TextBox txtnrodos 
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   3225
         TabIndex        =   90
         Text            =   "000000"
         Top             =   780
         Width           =   840
      End
      Begin VB.TextBox txtnrouno 
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   1605
         TabIndex        =   88
         Text            =   "000000"
         Top             =   795
         Width           =   795
      End
      Begin VB.TextBox txtcserie 
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   2985
         TabIndex        =   86
         Text            =   "000"
         Top             =   300
         Width           =   540
      End
      Begin VB.ComboBox fbg 
         ForeColor       =   &H00000080&
         Height          =   315
         ItemData        =   "FrmDocu.frx":E0F2
         Left            =   1095
         List            =   "FrmDocu.frx":E108
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   285
         Width           =   1140
      End
      Begin VB.Label lblcambio 
         Caption         =   "por el Nro :"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   3
         Left            =   2415
         TabIndex        =   89
         Top             =   825
         Width           =   945
      End
      Begin VB.Label lblcambio 
         Caption         =   "Cambio el Numero :"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   87
         Top             =   795
         Width           =   1410
      End
      Begin VB.Label lblcambio 
         Caption         =   "Serie :"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   2400
         TabIndex        =   85
         Top             =   315
         Width           =   540
      End
      Begin VB.Label lblcambio 
         Caption         =   "Cambiar:"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   315
         TabIndex        =   84
         Top             =   330
         Width           =   675
      End
   End
   Begin VB.Frame FRADIRE 
      Caption         =   "Direccion de Entrega de Mercaderia :"
      Height          =   735
      Left            =   120
      TabIndex        =   74
      Top             =   3240
      Visible         =   0   'False
      Width           =   8055
      Begin VB.ComboBox TxtZonaTrabajo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   360
         WhatsThisHelpID =   5
         Width           =   1935
      End
      Begin VB.ComboBox TxtSubZonaTrabajo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   360
         WhatsThisHelpID =   6
         Width           =   1935
      End
      Begin VB.TextBox txtnum 
         Height          =   285
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   76
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtdire 
         Height          =   285
         Left            =   240
         MaxLength       =   30
         TabIndex        =   75
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.CheckBox imp 
      BackColor       =   &H00808000&
      Caption         =   "&Directo a Impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   10560
      TabIndex        =   73
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Ce&rrar"
      Height          =   750
      Left            =   10560
      Picture         =   "FrmDocu.frx":E15B
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   10560
      Picture         =   "FrmDocu.frx":569E5
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Documento "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   6495
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   10335
      Begin VB.CommandButton CmdcruZe 
         Caption         =   "Cruzar Docum."
         Height          =   495
         Left            =   9240
         TabIndex        =   121
         Top             =   4800
         Width           =   975
      End
      Begin VB.CheckBox chelotes 
         Caption         =   "&Ver Lotes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6720
         TabIndex        =   93
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdcambios 
         Caption         =   "Cambiar Orden N. Documen."
         Height          =   750
         Left            =   9225
         TabIndex        =   81
         Top             =   5640
         Width           =   960
      End
      Begin VB.TextBox NUMERO 
         Height          =   285
         Left            =   4560
         TabIndex        =   68
         Top             =   5760
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox sin_valor 
         Caption         =   "&Guia Sin Valor"
         Height          =   315
         Left            =   6720
         TabIndex        =   67
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox FECHA_PART 
         Height          =   285
         Left            =   9120
         TabIndex        =   65
         Top             =   4440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox TRANS 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   5760
         Width           =   4335
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac2 
         Height          =   2535
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4471
         _Version        =   393216
         ForeColor       =   4210752
         ForeColorFixed  =   4210752
         BackColorBkg    =   8421376
         Enabled         =   -1  'True
         HighLight       =   2
         AllowUserResizing=   1
      End
      Begin VB.TextBox tguia 
         Height          =   285
         Left            =   9120
         TabIndex        =   59
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox cherela 
         Caption         =   "Guia Rem."
         Height          =   255
         Left            =   9120
         TabIndex        =   58
         Top             =   3465
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Tag             =   "119"
         Top             =   4800
         Width           =   9015
         Begin VB.Label lblabono 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   8160
            TabIndex        =   106
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lbletiqaj 
            Caption         =   "Abonos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   1
            Left            =   8280
            TabIndex        =   105
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblcargo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6240
            TabIndex        =   104
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lbletiqaj 
            Caption         =   "Cargos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   0
            Left            =   6360
            TabIndex        =   103
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lbldesc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5160
            TabIndex        =   96
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label LCODART 
            Caption         =   "Dsct/Valor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   0
            Left            =   5160
            TabIndex        =   97
            Tag             =   "9999"
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label d_flete 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4080
            TabIndex        =   33
            ToolTipText     =   "Doble Click para modificar..."
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label d_neto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7080
            TabIndex        =   32
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label d_descto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2040
            TabIndex        =   31
            Top             =   360
            Width           =   765
         End
         Begin VB.Label d_impto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3000
            TabIndex        =   30
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label d_gastos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1200
            TabIndex        =   29
            Top             =   360
            Width           =   765
         End
         Begin VB.Label d_subtotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label lblflete 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   4080
            TabIndex        =   20
            Tag             =   "9999"
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label LCODART 
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   2
            Left            =   7320
            TabIndex        =   19
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label LCODART 
            Caption         =   "Impto."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   4
            Left            =   3120
            TabIndex        =   18
            Tag             =   "9999"
            Top             =   120
            Width           =   525
         End
         Begin VB.Label LCODART 
            Caption         =   "Descto.(%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   17
            Tag             =   "9999"
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label LCODART 
            Caption         =   "Gastos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   16
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label LCODART 
            Caption         =   "Subtotal:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   15
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.CheckBox chetrans 
         Caption         =   "Imprimir Datos del Transportista en Guia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   5520
         Visible         =   0   'False
         Width           =   4215
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   165
         Left            =   3720
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label d_ruc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   280
         Left            =   5400
         TabIndex        =   128
         Top             =   840
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label d_codven 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         TabIndex        =   127
         Top             =   840
         Width           =   525
      End
      Begin VB.Label d_nomven 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7560
         TabIndex        =   126
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label FCONT 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fec. Contable :"
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label d_fecha_can 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4080
         TabIndex        =   124
         Top             =   850
         Width           =   1185
      End
      Begin VB.Label lblmotivo 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8160
         TabIndex        =   100
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label LBLEXTORNO 
         Caption         =   "DOCUMENTO EXTORNADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblhora 
         Alignment       =   2  'Center
         Caption         =   "00:00AM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   99
         Top             =   285
         Width           =   1140
      End
      Begin VB.Label l_fecha_can 
         AutoSize        =   -1  'True
         Caption         =   "F. Almacen :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   2880
         TabIndex        =   98
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lbltipvta 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   6825
         TabIndex        =   80
         Top             =   450
         Width           =   615
      End
      Begin VB.Label d_tipvta 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7530
         TabIndex        =   79
         Top             =   465
         Width           =   2685
      End
      Begin VB.Label lbldocu 
         Caption         =   "Fecha Part."
         Height          =   255
         Index           =   2
         Left            =   9120
         TabIndex        =   66
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label d_dias 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9480
         TabIndex        =   13
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label d_newvcto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9120
         TabIndex        =   38
         Top             =   3045
         Width           =   1125
      End
      Begin VB.Label lbldocu 
         Caption         =   "Dias Cred."
         Height          =   255
         Index           =   3
         Left            =   9360
         TabIndex        =   23
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label d_fecha_compra 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   61
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label d_fecha 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/2001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label l_fecha_compra 
         AutoSize        =   -1  'True
         Caption         =   "Fec. Emisión :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label d_moneda 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6600
         TabIndex        =   46
         Top             =   5760
         Width           =   375
      End
      Begin VB.Label txtdocu 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6720
         TabIndex        =   47
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label d_dire 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   40
         Top             =   1920
         Width           =   4380
         WordWrap        =   -1  'True
      End
      Begin VB.Label d_usuario 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   55
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lbldireccion 
         Caption         =   "Dir.Entrega:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblfac 
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   225
         Left            =   5640
         TabIndex        =   48
         Top             =   1935
         Width           =   975
      End
      Begin VB.Label lbldomicilio 
         Caption         =   "Domicilio :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label d_domicilio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   1110
         TabIndex        =   53
         Top             =   1560
         Width           =   5505
         WordWrap        =   -1  'True
      End
      Begin VB.Label d_efectivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   52
         Tag             =   "9999"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblEfectivo 
         Caption         =   "Total Efectivo ."
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
         Left            =   240
         TabIndex        =   51
         Tag             =   "9999"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblsaldo 
         Caption         =   "Saldo Actual ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   36
         Tag             =   "9999"
         Top             =   5760
         Width           =   1275
      End
      Begin VB.Label lblcheque 
         Caption         =   "Total Cheque ."
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
         Left            =   2280
         TabIndex        =   50
         Tag             =   "9999"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label d_cheque 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         TabIndex        =   49
         Tag             =   "9999"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label LBLRUC 
         Caption         =   "R.U.C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   43
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbldocu 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Vcto."
         Height          =   195
         Index           =   1
         Left            =   9120
         TabIndex        =   39
         Top             =   2805
         Width           =   900
      End
      Begin VB.Label d_saldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   7080
         TabIndex        =   37
         Tag             =   "9999"
         Top             =   5760
         Width           =   1485
      End
      Begin VB.Label lblcondicion 
         Alignment       =   1  'Right Justify
         Caption         =   "Condición:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   6420
         TabIndex        =   25
         Top             =   135
         Width           =   1020
      End
      Begin VB.Label d_fechaV 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9120
         TabIndex        =   21
         Top             =   2520
         Width           =   1125
      End
      Begin VB.Label d_condicion 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7530
         TabIndex        =   12
         Top             =   135
         Width           =   2685
      End
      Begin VB.Label d_nomclie 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   280
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label d_Codclie 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lblpersona 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbldocu 
         Caption         =   "Fecha de Vcto."
         Height          =   255
         Index           =   5
         Left            =   9120
         TabIndex        =   24
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lbldocu 
         Caption         =   "Fec. Proceso :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label d_mensaje 
         Alignment       =   2  'Center
         Caption         =   "DOCUMENTO EN BLANCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   44
         Top             =   600
         Visible         =   0   'False
         Width           =   3255
      End
   End
   Begin VB.TextBox tmoneda 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
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
      Height          =   285
      Left            =   6210
      TabIndex        =   57
      Text            =   "S/."
      Top             =   4920
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   -120
      Width           =   10335
      Begin VB.TextBox txtSerie 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   6000
         MaxLength       =   3
         TabIndex        =   123
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtvend 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3045
         TabIndex        =   122
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "&Sig"
         Height          =   735
         Left            =   9240
         Picture         =   "FrmDocu.frx":6780E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdAnterior 
         Caption         =   "&Ant."
         Height          =   735
         Left            =   8400
         Picture         =   "FrmDocu.frx":6B218
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   200
         Width           =   735
      End
      Begin VB.ComboBox TIPMOV 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   345
         Width           =   2655
      End
      Begin VB.TextBox txtNumfac 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   6840
         MaxLength       =   9
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbFBG 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   345
         Width           =   2055
      End
      Begin VB.Label lblvend 
         Caption         =   "Vend."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   3000
         TabIndex        =   63
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblNumfac 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   6240
         TabIndex        =   9
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000080&
         Height          =   855
         Left            =   5880
         TabIndex        =   56
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lbldocu 
         Alignment       =   2  'Center
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lbldocu 
         Alignment       =   2  'Center
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   8
         Top             =   120
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   240
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbldataant 
      BackStyle       =   0  'Transparent
      Caption         =   "Data anterior"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   10560
      TabIndex        =   101
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblfondo 
      BackColor       =   &H000000FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   10500
      TabIndex        =   95
      Top             =   1410
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "OSBusiness"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   10440
      TabIndex        =   69
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Height          =   7335
      Index           =   5
      Left            =   10440
      TabIndex        =   70
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmDocu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DOC102 As String
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const VK_RETURN = &HD
Dim temp_NUMSER_C  As String
Dim temp_NUMFAC_C  As String
Dim loc_flag_espera As String * 1
Dim pstransporte As rdoQuery
Dim transporte As rdoResultset
Dim PSFAR As rdoQuery
Dim far_r As rdoResultset
Dim PSFAR_CONSUL As rdoQuery
Dim far_consul As rdoResultset
Dim PSCAR_CONSUL As rdoQuery
Dim car_consul As rdoResultset
Dim wflag_docu As String * 1
Dim temporal
Dim tempo_serie
Dim LOC_TIPMOV As Integer
Dim LOC_NUMFAC_FIN As Currency
Dim WGUIA_RELA As String
Dim SERIE_GUIA As String * 3
Dim PS_VE2 As rdoQuery
Dim VE2_LLAVE As rdoResultset
Dim SIN_CODART As Integer
Dim PS_TRA As rdoQuery
Dim llave_trans As rdoResultset
Dim LOC_ARROZ As String * 1
Dim LOC_NUMPED As Currency
Dim loc_acceso_descto As String * 1
'variables agregadas
Dim rs As rdoResultset
Dim PS As rdoQuery
Dim loc_acceso_cruzados As String * 1
Dim Doc_EL As String

Private Sub cherela_Click()
 If tguia.Visible Then tguia.SetFocus
 If cherela.Value = 1 Then
   If Val(WGUIA_RELA) <> 0 Then
      tguia.Text = "G/." + Trim(SERIE_GUIA) + " - " + WGUIA_RELA
   End If
 End If
' If LK_EMP = "HER" Then
 '   If cherela.Value = 1 Then
 '      tguia.Text = "G/." + Trim(usu_llave!USU_SERIE_G) + " - "
 '      tguia.SetFocus
 '      tguia.SelStart = Len(tguia.Text)
 '   Else
 '     tguia.Text = ""
 '   End If
 'End If
 NUMERO.Visible = True
End Sub

Private Sub chetrans_Click()
If chetrans.Value = 0 Then
  TRANS.ListIndex = -1
  TRANS.Enabled = False
Else
  TRANS.Enabled = True
End If
If TRANS.Visible Then
 If TRANS.ListCount <> 0 Then
   TRANS.ListIndex = TRANS.ListIndex
 End If
  If TRANS.Enabled Then TRANS.SetFocus
End If
End Sub

Private Sub cmbFBG_Click()
EVENTO_CLICK
End Sub

Private Sub cmbFBG_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
 Exit Sub
End If
If temporal = "X" Then
 Exit Sub
End If
'SQ_OPER = 1
'PUB_CODCIA = LK_CODCIA
'LEER_PAR_LLAVE
'If par_llave.EOF Then
'  Exit Sub
'End If
If LOC_TIPMOV <> 3 Then pu_cp = " "
If Left(cmbFBG.Text, 1) = "P" And PUB_TIPMOV <> 10 Then
 txtNumfac.Text = Nulo_Valor0(par_llave!par_planilla)
 txtNumfac.SetFocus
 txtnumfac_KeyPress 13
 Exit Sub
End If
If LOC_TIPMOV = 30 Then
 Dim PSTEMP_MAYOR As rdoQuery
 Dim temp_mayor  As rdoResultset
 Dim wser
 Dim wnumfac As Currency
 pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ?  and ped_tipmov=20  ORDER BY  PED_NUMFAC DESC "
 Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
 PSTEMP_MAYOR.rdoParameters(0) = " "
 PSTEMP_MAYOR.MaxRows = 1
 Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
 PSTEMP_MAYOR(0) = LK_CODCIA
 temp_mayor.Requery
 If temp_mayor.EOF Then
    wser = 0
    wnumfac = 0
 Else
    wser = Nulo_Valors(temp_mayor!PED_NUMSER)
    wnumfac = Nulo_Valor0(temp_mayor!PED_NUMFAC)
 End If
 pu_cp = "P"
 txtSerie.Text = wser
 txtNumfac.Text = wnumfac
 txtNumfac.SetFocus
 txtnumfac_KeyPress 13
 Exit Sub
End If

If LK_FLAG_FACTURACION = "V" Then
   If ven_llave.EOF Then GoTo dale
   Select Case Left(cmbFBG.Text, 1)
   Case "G"
       txtSerie.Text = ven_llave!VEM_SERIE_G
   Case "B"
       txtSerie.Text = ven_llave!vem_serie_b
   Case "F"
       txtSerie.Text = ven_llave!vem_serie_f
   Case "P"
       txtSerie.Text = ven_llave!VEM_SERIE_P
   End Select
dale:
ElseIf LK_FLAG_FACTURACION = "A" And LOC_TIPMOV <> 93 Then
 If Left(cmbFBG.Text, 1) = "F" Then
  txtSerie.Text = par_llave!PAR_F_SERIE
 ElseIf Left(cmbFBG.Text, 1) = "B" Then
  txtSerie.Text = par_llave!PAR_B_SERIE
 ElseIf Left(cmbFBG.Text, 1) = "G" Then
  txtSerie.Text = par_llave!PAR_G_SERIE
 ElseIf Left(cmbFBG.Text, 1) = "N" Then
  txtSerie.Text = par_llave!PAR_SERIE_NCRE
 ElseIf Left(cmbFBG.Text, 1) = "D" Then
  txtSerie.Text = par_llave!PAR_SERIE_NDEB
 End If
ElseIf LK_FLAG_FACTURACION = "U" Then
   If Left(cmbFBG.Text, 1) = "F" Then
     txtSerie.Text = usu_llave!USU_SERIE_F
   ElseIf Left(cmbFBG.Text, 1) = "B" Then
     txtSerie.Text = usu_llave!USU_SERIE_B
   End If
End If
If LOC_TIPMOV = 10 Then
 pu_cp = "C"
 PU_TIPMOV = 10
 PU_NUMSER = Val(txtSerie.Text)
 PU_FBG = Left(cmbFBG.Text, 1)
ElseIf LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
 PU_TIPMOV = LOC_TIPMOV
 PU_NUMSER = Val(txtSerie.Text)
 PU_FBG = Left(cmbFBG.Text, 1)
 pu_cp = "C"
 If Right(Trim(cmbFBG.Text), 1) = "P" Then
   txtSerie.Text = "0"
   PU_NUMSER = 0
   PU_FBG = "C"
   pu_cp = "P"
 End If
 
ElseIf LOC_TIPMOV = 20 Or LOC_TIPMOV = 99 Then
 pu_cp = "P"
 PU_TIPMOV = LOC_TIPMOV
 PU_NUMSER = 0
 PU_FBG = " "
 If LOC_TIPMOV = 20 And Left(cmbFBG.Text, 1) <> "K" Then
    If Left(cmbFBG.Text, 1) = "F" Then
      pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER_C = ? AND FAR_NUMFAC_C = ?  AND FAR_ESTADO2 <> 'L' ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
    Else
      pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMGUIA = ?   AND FAR_ESTADO2 <> 'L' ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
    End If
    Set PSFAR = CN.CreateQuery("", pub_cadena)
    If Left(cmbFBG.Text, 1) = "F" Then
        PSFAR.rdoParameters(0) = 0
        PSFAR.rdoParameters(1) = 0
        PSFAR.rdoParameters(2) = 0
        PSFAR.rdoParameters(3) = 0
    Else
        PSFAR.rdoParameters(0) = 0
        PSFAR.rdoParameters(1) = 0
        PSFAR.rdoParameters(2) = 0
    End If
    Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurValues)
 End If
 If LOC_TIPMOV = 99 Then
  PU_FBG = Left(cmbFBG.Text, 1)
 End If
Else
'' txtserie.Text = "0"
 PU_TIPMOV = LOC_TIPMOV
 PU_NUMSER = 0
 PU_FBG = " "
End If
PU_NUMSER = Val(txtSerie.Text)
pu_codcia = LK_CODCIA
Screen.MousePointer = 11
LEER_FAR_CONSUL
If Not far_consul.EOF Then
 txtNumfac.Text = far_consul!far_numfac
Else
 txtNumfac.Text = "0"
End If
 Screen.MousePointer = 0
txtNumfac.SetFocus
txtnumfac_KeyPress 13
If LK_EMP = "HER" And Val(txtSerie.Text) = 0 And LOC_TIPMOV = 10 Then
   txtSerie.Locked = False
   txtSerie.Text = ""
   txtSerie.SetFocus
End If
End Sub

Private Sub CmdAnterior_Click()
Dim tempo
If LOC_TIPMOV = 0 Then Exit Sub
tempo = Val(txtNumfac.Text)
If LOC_TIPMOV = 10 Then
 If Trim(txtSerie.Text) = "" Then
  Exit Sub
 End If
End If
If Val(txtNumfac.Text) <= 0 Then
 LIMPIA_DOCU
 grid_fac2.Clear
 Exit Sub
End If
txtNumfac.Text = Val(txtNumfac.Text) - 1
If LOC_TIPMOV = 96 Or LOC_TIPMOV = 30 Then ' PLANILLA
 txtnumfac_KeyPress 13
 Exit Sub
End If
wflag_docu = ""
loc_flag_espera = "A"
LLENA_CONSULTA
loc_flag_espera = ""
If wflag_docu = "A" Then
  If Trim(d_fecha.Caption) <> "" Then LIMPIA_DOCU
  d_mensaje.Visible = True
  'CmdAnterior.Enabled = False
  'Beep
  'Beep
Else
  d_mensaje.Visible = False
  CmdAnterior.Enabled = True
End If
Azul txtNumfac, txtNumfac
LOC_NUMFAC_FIN = Val(txtNumfac.Text)

End Sub

Private Sub cmdcambiar_Click()
Dim kFBG  As String
Dim kSERIE  As String
Dim kNUMERO1  As String
Dim kNUMERO2  As String
cmdesc.Enabled = False
cmdcambiar.Enabled = False
On Error GoTo sale
kFBG = Left(fbg.Text, 1)
kSERIE = txtcserie.Text
kNUMERO1 = Val(txtnrouno.Text)
kNUMERO2 = Val(txtnrodos.Text)

If Left(Trim(fbg.Text), 4) = "ENVI" Then
pub_cadena = "UPDATE FACART SET FAR_NUMFAC = " & kNUMERO2 & " WHERE FAR_CODCIA = '" & LK_CODCIA & "' AND  FAR_TIPMOV = " & LOC_TIPMOV & " AND FAR_FBG = ' ' AND  FAR_NUMSER = " & kSERIE & " AND FAR_NUMFAC = " & kNUMERO1
CN.Execute pub_cadena, rdExecDirect

pub_cadena = "UPDATE ALLOG SET ALL_NUMFAC = " & kNUMERO2 & " WHERE ALL_CODCIA = '" & LK_CODCIA & "' AND ALL_TIPMOV = " & LOC_TIPMOV & " AND ALL_FBG = ' ' AND ALL_NUMSER = " & kSERIE & " AND ALL_NUMFAC = " & kNUMERO1
CN.Execute pub_cadena, rdExecDirect

Else
pub_cadena = "UPDATE FACART SET FAR_NUMFAC = " & kNUMERO2 & " WHERE FAR_CODCIA = '" & LK_CODCIA & "' AND  FAR_TIPMOV = " & LOC_TIPMOV & " AND FAR_FBG = '" & kFBG & "' AND  FAR_NUMSER = " & kSERIE & " AND FAR_NUMFAC = " & kNUMERO1
CN.Execute pub_cadena, rdExecDirect

pub_cadena = "UPDATE ALLOG SET ALL_NUMFAC = " & kNUMERO2 & " WHERE ALL_CODCIA = '" & LK_CODCIA & "' AND ALL_TIPMOV = " & LOC_TIPMOV & " AND ALL_FBG = '" & kFBG & "' AND ALL_NUMSER = " & kSERIE & " AND ALL_NUMFAC = " & kNUMERO1
CN.Execute pub_cadena, rdExecDirect


pub_cadena = "UPDATE CARTERA SET CAR_NUMFAC = " & kNUMERO2 & " WHERE CAR_CODCIA = '" & LK_CODCIA & "' AND  CAR_TIPMOV = " & LOC_TIPMOV & " AND CAR_FBG = '" & kFBG & "' AND CAR_NUMSER = " & kSERIE & " AND CAR_NUMFAC = " & kNUMERO1
CN.Execute pub_cadena, rdExecDirect

pub_cadena = "UPDATE CARACU SET CAA_NUMFAC = " & kNUMERO2 & " WHERE CAA_CODCIA = '" & LK_CODCIA & "' AND CAA_TIPMOV = " & LOC_TIPMOV & " AND CAA_FBG = '" & kFBG & "' AND CAA_NUMSER = " & kSERIE & " AND CAA_NUMFAC = " & kNUMERO1
CN.Execute pub_cadena, rdExecDirect
End If
txtnrouno.Text = ""
txtnrodos.Text = ""
cmdesc.Enabled = True
cmdcambiar.Enabled = True
MsgBox "Ok. Nro Cambiado. ", 48, Pub_Titulo
'cmdesc_Click

Exit Sub
sale:
If Err.Number = 40002 Then
  MsgBox "Nro NO se Actualizo,  Verificar " & Chr(13) & "IMPORTANTE:" & Chr(13) & "El Numero de Documento a Reemplazar debe estar en Blanco !!!", 48, Pub_Titulo
Else
  MsgBox "Nro NO se Actualizo,  Verificar " & Chr(13) & Err.Description, 48, Pub_Titulo
End If
Azul txtnrodos, txtnrodos
cmdesc.Enabled = True
cmdcambiar.Enabled = True
End Sub

Private Sub cmdcambios_Click()
  fracambio.Visible = True
  fbg.SetFocus
  fbg.ListIndex = 0
End Sub

Private Sub cmdCerrar_Click()
Unload frmdocu
End Sub

Private Sub CmdcruZe_Click()
Fracruze.Visible = True
Azul2 tfcruze, tfcruze
End Sub

Private Sub CmdCruze_escapar_Click()
Fracruze.Visible = False
End Sub

Private Sub CmdCruze_Int_Click()
Dim ID_SEC_DOC1 As Integer
Dim ID_SEC_DOC2 As Integer
Screen.MousePointer = 11
' Validaciones:
If IsDate(tfcruze.Text) = False Then
  MsgBox "La Fecha Ingresada es Incorrecta. ", 48, Pub_Titulo
  tfcruze.SetFocus
  GoTo fincruze
End If
If Trim(cFB1.Text) = Trim(cFB2.Text) Then
Else
  MsgBox "Tipo de Documento debe ser Igual.", 48, Pub_Titulo
  GoTo fincruze
End If

'==========================================================

pub_cadena = "SELECT FAR_CODCIA, FAR_NUM_LOTE FROM FACART WHERE FAR_CODCIA = '" & LK_CODCIA & "' AND FAR_TIPMOV = 10 AND FAR_FECHA = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND FAR_FBG = '" & Trim(cFB1.Text) & "' AND FAR_NUMSER = " & Val(Ts1.Text) & " AND FAR_NUMFAC = " & Val(Tn1.Text)
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If X.EOF Then
  MsgBox "Documento Nro 1 , No Existe. ", 48, Pub_Titulo
  GoTo fincruze
End If
ID_SEC_DOC1 = X!FAR_NUM_LOTE

pub_cadena = "SELECT FAR_CODCIA, FAR_NUM_LOTE FROM FACART WHERE FAR_CODCIA = '" & LK_CODCIA & "' AND FAR_TIPMOV = 10 AND FAR_FECHA = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND FAR_FBG = '" & Trim(cFB2.Text) & "' AND FAR_NUMSER = " & Val(Ts2.Text) & " AND FAR_NUMFAC = " & Val(Tn2.Text)
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If X.EOF Then
  MsgBox "Documento Nro 2 , No Existe. ", 48, Pub_Titulo
  GoTo fincruze
End If
ID_SEC_DOC2 = X!FAR_NUM_LOTE
'-----------valia saldo original--------
If ID_SEC_DOC1 = 2 Or ID_SEC_DOC1 = 4 Then
 pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CODCIA = '" & LK_CODCIA & "' AND CAR_TIPMOV = 10 AND CAR_CODTRA = 2401 AND CAR_IMPORTE <> 0 AND CAR_FECHA_INGR = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND CAR_FBG = '" & Trim(cFB1.Text) & "' AND CAR_NUMSER = " & Val(Ts1.Text) & " AND CAR_NUMFAC = " & Val(Tn1.Text)
 Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
 If Not X.EOF Then
   If Val(X!car_importe) <> Val(X!CAR_IMP_INI) Then
       MsgBox "El Documento Nro 1, Tiene Amortizaciones o esta Cancelado, No procede el Cambio, tiene que tener saldo original. ", 48, Pub_Titulo
       GoTo fincruze
   End If
 Else
   MsgBox "El Documento Nro 1, Tiene Amortizaciones o esta Cancelado, No procede el Cambio, tiene que tener saldo original. ", 48, Pub_Titulo
   GoTo fincruze
 End If
End If

If ID_SEC_DOC2 = 2 Or ID_SEC_DOC2 = 4 Then
 pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CODCIA = '" & LK_CODCIA & "' AND CAR_TIPMOV = 10 AND CAR_CODTRA = 2401 AND CAR_IMPORTE <> 0 AND CAR_FECHA_INGR = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND CAR_FBG = '" & Trim(cFB2.Text) & "' AND CAR_NUMSER = " & Val(Ts2.Text) & " AND CAR_NUMFAC = " & Val(Tn2.Text)
 Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
 If Not X.EOF Then
   If Val(X!car_importe) <> Val(X!CAR_IMP_INI) Then
       MsgBox "El Documento Nro 2, Tiene Amortizaciones o esta Cancelado, No procede el Cambio, tiene que tener saldo original. ", 48, Pub_Titulo
       GoTo fincruze
   End If
 Else
   MsgBox "El Documento Nro 2, Tiene Amortizaciones o esta Cancelado, No procede el Cambio, tiene que tener saldo original. ", 48, Pub_Titulo
   GoTo fincruze
 End If
 
End If
pub_mensaje = "Documentos Validados Hacer en Cruze de Documento   ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   GoTo fincruze
End If

'---------------------------------------
On Error GoTo ESCAPA
CN.Execute "Begin Transaction", rdExecDirect
pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)

'=================MOVER EL DOC 1 A UN ESPACIO VACIO=========
pub_cadena = "UPDATE FACART SET FAR_CODCIA = 'XX' WHERE FAR_CODCIA = '" & LK_CODCIA & "' AND FAR_TIPMOV = 10 AND FAR_FECHA = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND FAR_FBG = '" & Trim(cFB1.Text) & "' AND FAR_NUMSER = " & Val(Ts1.Text) & " AND FAR_NUMFAC = " & Val(Tn1.Text)
CN.Execute pub_cadena

pub_cadena = "UPDATE ALLOG SET ALL_CODCIA = 'XX' WHERE ALL_CODCIA = '" & LK_CODCIA & "' AND ALL_TIPMOV = 10 AND ALL_CODTRA = 2401 AND ALL_FECHA_DIA = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND ALL_FBG = '" & Trim(cFB1.Text) & "' AND ALL_NUMSER = " & Val(Ts1.Text) & " AND ALL_NUMFAC = " & Val(Tn1.Text)
CN.Execute pub_cadena

pub_cadena = "UPDATE CARTERA SET CAR_CODCIA = 'XX' WHERE CAR_CODCIA = '" & LK_CODCIA & "' AND CAR_TIPMOV = 10 AND CAR_CODTRA = 2401 AND CAR_FECHA_INGR = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND CAR_FBG = '" & Trim(cFB1.Text) & "' AND CAR_NUMSER = " & Val(Ts1.Text) & " AND CAR_NUMFAC = " & Val(Tn1.Text)
CN.Execute pub_cadena

pub_cadena = "UPDATE CARACU SET CAA_CODCIA = 'XX'  WHERE CAA_CODCIA = '" & LK_CODCIA & "' AND CAA_TIPMOV = 10 AND CAA_CODTRA = 2401 AND CAA_FECHA = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND CAA_FBG = '" & Trim(cFB1.Text) & "' AND CAA_NUMSER = " & Val(Ts1.Text) & " AND CAA_NUMFAC = " & Val(Tn1.Text)
CN.Execute pub_cadena

'===========================================================

'=================MOVER EL DOC 2 A UN ESPACIO DEL DOC 1 ====
pub_cadena = "UPDATE FACART SET FAR_FBG = '" & Trim(cFB1.Text) & "' , FAR_NUMSER = " & Val(Ts1.Text) & " , FAR_NUMFAC = " & Val(Tn1.Text) & "  WHERE FAR_CODCIA = '" & LK_CODCIA & "' AND FAR_TIPMOV = 10 AND FAR_FECHA = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND FAR_FBG = '" & Trim(cFB2.Text) & "' AND FAR_NUMSER = " & Val(Ts2.Text) & " AND FAR_NUMFAC = " & Val(Tn2.Text)
CN.Execute pub_cadena

pub_cadena = "UPDATE ALLOG SET ALL_FBG = '" & Trim(cFB1.Text) & "' , ALL_NUMSER = " & Val(Ts1.Text) & " , ALL_NUMFAC = " & Val(Tn1.Text) & " WHERE ALL_CODCIA = '" & LK_CODCIA & "' AND ALL_TIPMOV = 10 AND ALL_CODTRA = 2401 AND ALL_FECHA_DIA = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND ALL_FBG = '" & Trim(cFB2.Text) & "' AND ALL_NUMSER = " & Val(Ts2.Text) & " AND ALL_NUMFAC = " & Val(Tn2.Text)
CN.Execute pub_cadena

pub_cadena = "UPDATE CARTERA SET CAR_FBG = '" & Trim(cFB1.Text) & "' , CAR_NUMSER = " & Val(Ts1.Text) & " , CAR_NUMFAC = " & Val(Tn1.Text) & " WHERE CAR_CODCIA = '" & LK_CODCIA & "' AND CAR_TIPMOV = 10 AND CAR_CODTRA = 2401 AND CAR_FECHA_INGR = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND CAR_FBG = '" & Trim(cFB2.Text) & "' AND CAR_NUMSER = " & Val(Ts2.Text) & " AND CAR_NUMFAC = " & Val(Tn2.Text)
CN.Execute pub_cadena

pub_cadena = "UPDATE CARACU SET CAA_FBG = '" & Trim(cFB1.Text) & "' , CAA_NUMSER = " & Val(Ts1.Text) & " , CAA_NUMFAC = " & Val(Tn1.Text) & "  WHERE CAA_CODCIA = '" & LK_CODCIA & "' AND CAA_TIPMOV = 10 AND CAA_CODTRA = 2401 AND CAA_FECHA = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND CAA_FBG = '" & Trim(cFB2.Text) & "' AND CAA_NUMSER = " & Val(Ts2.Text) & " AND CAA_NUMFAC = " & Val(Tn2.Text)
CN.Execute pub_cadena

'===========================================================
pub_cadena = "UPDATE FACART SET FAR_CODCIA='" & LK_CODCIA & "', FAR_FBG = '" & Trim(cFB2.Text) & "' , FAR_NUMSER = " & Val(Ts2.Text) & " , FAR_NUMFAC = " & Val(Tn2.Text) & "  WHERE FAR_CODCIA = 'XX' AND FAR_TIPMOV = 10 AND FAR_FECHA = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND FAR_FBG = '" & Trim(cFB1.Text) & "' AND FAR_NUMSER = " & Val(Ts1.Text) & " AND FAR_NUMFAC = " & Val(Tn1.Text)
CN.Execute pub_cadena

pub_cadena = "UPDATE ALLOG SET ALL_CODCIA='" & LK_CODCIA & "', ALL_FBG = '" & Trim(cFB2.Text) & "' , ALL_NUMSER = " & Val(Ts2.Text) & " , ALL_NUMFAC = " & Val(Tn2.Text) & " WHERE ALL_CODCIA = 'XX' AND ALL_TIPMOV = 10 AND ALL_CODTRA = 2401 AND ALL_FECHA_DIA = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND ALL_FBG = '" & Trim(cFB1.Text) & "' AND ALL_NUMSER = " & Val(Ts1.Text) & " AND ALL_NUMFAC = " & Val(Tn1.Text)
CN.Execute pub_cadena

pub_cadena = "UPDATE CARTERA SET CAR_CODCIA='" & LK_CODCIA & "', CAR_FBG = '" & Trim(cFB2.Text) & "' , CAR_NUMSER = " & Val(Ts2.Text) & " , CAR_NUMFAC = " & Val(Tn2.Text) & " WHERE CAR_CODCIA = 'XX' AND CAR_TIPMOV = 10 AND CAR_CODTRA = 2401 AND CAR_FECHA_INGR = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND CAR_FBG = '" & Trim(cFB1.Text) & "' AND CAR_NUMSER = " & Val(Ts1.Text) & " AND CAR_NUMFAC = " & Val(Tn1.Text)
CN.Execute pub_cadena

pub_cadena = "UPDATE CARACU SET CAA_CODCIA='" & LK_CODCIA & "', CAA_FBG = '" & Trim(cFB2.Text) & "' , CAA_NUMSER = " & Val(Ts2.Text) & " , CAA_NUMFAC = " & Val(Tn2.Text) & "  WHERE CAA_CODCIA = 'XX' AND CAA_TIPMOV = 10 AND CAA_CODTRA = 2401 AND CAA_FECHA = '" & Format(tfcruze.Text, "dd/mm/yyyy") & "' AND CAA_FBG = '" & Trim(cFB1.Text) & "' AND CAA_NUMSER = " & Val(Ts1.Text) & " AND CAA_NUMFAC = " & Val(Tn1.Text)
CN.Execute pub_cadena

con_llave.Close
CN.Execute "Commit Transaction", rdExecDirect
Screen.MousePointer = 0
MsgBox "Los Documentos se han cambiado", 48, Pub_Titulo
Tn1.Text = ""
Tn2.Text = ""
Tn1.SetFocus
Exit Sub
fincruze:
Screen.MousePointer = 0
Exit Sub
ESCAPA:
con_llave.Close
CN.Execute "Rollback Transaction", rdExecDirect
MsgBox Err.Description & "Intantar Nuevamente", 48, Pub_Titulo
Screen.MousePointer = 0

End Sub

Private Sub cmdesc_Click()

fracambio.Visible = False
End Sub

Private Sub cmdImp_Click()
If loc_flag_espera = "A" Then
 MsgBox "Espere ....!!!", 48, Pub_Titulo
 Exit Sub
End If
If LOC_TIPMOV = 0 Or Trim(d_fecha.Caption) = "" Then
Exit Sub
End If
If frmdocu.LBLEXTORNO.Visible Then
  MsgBox "Impresión No Procede...", 48, Pub_Titulo
  Exit Sub
End If
If LOC_TIPMOV = 30 Then
  GoSub ORDEN
  Exit Sub
End If
If chetrans.Value = 1 Then
 If TRANS.Text = "" Then
    MsgBox "Datos del Transportista: Seleccione uno de la lista", 48, Pub_Titulo
    Exit Sub
 End If
End If
SIN_CODART = 0
If LOC_TIPMOV = 35 Or LOC_TIPMOV = 75 Or LOC_TIPMOV = 99 Or LOC_TIPMOV = 102 Or LOC_TIPMOV = 3 Or LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Or LOC_TIPMOV = 93 Or LOC_TIPMOV = 20 Or LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Or LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
  If grid_fac2.TextMatrix(1, 1) = "" Then SIN_CODART = 1
  fila = REP_CONSUL
 Exit Sub
End If

If LOC_TIPMOV <> 96 Then
 Exit Sub
End If
If LOC_TIPMOV = 10 Then
 If chetrans.Value = 1 And Val(Right(TRANS.Text, 3)) = 0 Then
      MsgBox "Estro  Activada a opciro n de Transportista. Seleccione un Transportista ?", 48, Pub_Titulo
     Exit Sub
 End If
End If
Dim i, j
Dim wranF
Dim LETRAS(24) As String * 1


Dim xl As Object
On Error GoTo FINTODO
Screen.MousePointer = 11
PB.Visible = True
pbMax = 6
pbMin = 0
pbValue = 0
pbValue = pbValue + 1
GoSub WEXCEL
pbValue = pbValue + 1
pub_cadena = ""
xl.Cells(4, 1) = "PLANILLA : " & Trim(txtSerie.Text) & " - " & Trim(txtNumfac.Text)
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
xl.Cells(2, 1) = "PLANILLA DE COBRANZA"
pbValue = pbValue + 1
For i = 0 To grid_fac2.Rows - 1
  For j = 0 To 14
     If grid_fac2.TextMatrix(i, j) = "" Then
       xl.Cells(i + 7, j + 1) = " "
     Else
       xl.Cells(i + 7, j + 1) = grid_fac2.TextMatrix(i, j)
     End If
  Next j
Next i
pbValue = pbValue + 1
GoSub LETRAS
pbValue = pbValue + 1
wranF = "A" & i + 8 & ":D" & i + 8
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
xl.Cells(i + 1 + 7, 1) = "Total Cheque.  ="
xl.Cells(i + 1 + 7, 2) = "'" & d_cheque.Caption
xl.Cells(i + 2 + 7, 1) = "Total Efectivo.="
xl.Cells(i + 2 + 7, 2) = "'" & d_efectivo.Caption
xl.Cells(i + 3 + 7, 1) = "Total Planilla.="
xl.Cells(i + 3 + 7, 2) = "'" & d_saldo.Caption

wranF = "A8:" & "O8"
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
xl.Cells(2, 1) = "PLANILA DE COBRANZA"
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
pbValue = pbValue + 1
xl.DisplayAlerts = False
xl.Worksheets(1).Protect PUB_CLAVE
xl.Application.Visible = True
Set xl = Nothing
PB.Visible = False
Screen.MousePointer = 0
Exit Sub

WEXCEL:
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\CONSPLA.xls", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

LETRAS:

LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
Return

ORDEN:
Dim wser As String * 3
Dim WSRUTA As String
Dim wRuta As String
Dim rmoneda As String * 1
wRuta = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\"
 frmdocu.Reportes.Destination = crptToWindow  '= crptToPrinter
 frmdocu.Reportes.WindowLeft = 2
 frmdocu.Reportes.WindowTop = 70
 frmdocu.Reportes.WindowWidth = 635
 frmdocu.Reportes.WindowHeight = 390
 frmdocu.Reportes.Formulas(1) = ""
 PUB_NETO = Val(frmdocu.d_neto.Caption)
 PUB_FECHA = frmdocu.d_fecha.Caption
 PU_NUMSER = Val((frmdocu.txtSerie.Text))

If Left(d_moneda.Caption, 3) = "US$" Then
   rmoneda = "D"
Else
   rmoneda = "S"
End If

PU_NUMFAC = Val((frmdocu.txtNumfac.Text))
frmdocu.Reportes.Formulas(1) = "SON_EFECTIVO=  'SON: " & CONVER_LETRAS(PUB_NETO, rmoneda) & "'"
frmdocu.Reportes.WindowTitle = "ORDEN DE COMPRA  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
frmdocu.Reportes.ReportFileName = wRuta + "ORDEN.RPT"
frmdocu.Reportes.Connect = PUB_ODBC
wser = PU_NUMSER
pub_cadena = "{PEDIDOS.PED_ESTADO} = 'N' AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND {PEDIDOS.PED_NUMSER}= '" & wser & "' AND {PEDIDOS.PED_NUMFAC} = " & PU_NUMFAC
frmdocu.Reportes.SelectionFormula = pub_cadena
frmdocu.Reportes.WindowTitle = frmdocu.Reportes.WindowTitle & " Archivo: " & Trim(frmdocu.Reportes.ReportFileName)
On Error GoTo accion
frmdocu.Reportes.Action = 1
On Error GoTo 0
Return
FINTODO:
accion:
 MsgBox Err.Description, 48, Pub_Titulo
 MsgBox "Reintente Nuevamente ..", 48, Pub_Titulo
End Sub

Private Sub cmdSiguiente_Click()
Dim tempo
If LOC_TIPMOV = 0 Then Exit Sub
tempo = Val(txtNumfac.Text)
If LOC_TIPMOV = 10 Then
 If Trim(txtSerie.Text) = "" Then
  Exit Sub
 End If
End If
If Val(txtNumfac.Text) < 0 Then
  Exit Sub
End If
txtNumfac.Text = Val(txtNumfac.Text) + 1
If LOC_TIPMOV = 96 Or LOC_TIPMOV = 30 Then ' PLANILLA
 txtnumfac_KeyPress 13
 Exit Sub
End If
wflag_docu = ""
loc_flag_espera = "A"
LLENA_CONSULTA
loc_flag_espera = ""
If wflag_docu = "A" Then
  If Trim(d_fecha.Caption) <> "" Then LIMPIA_DOCU
  d_mensaje.Visible = True
Else
  d_mensaje.Visible = False
  cmdSiguiente.Enabled = True
End If
Azul txtNumfac, txtNumfac
LOC_NUMFAC_FIN = Val(txtNumfac.Text)
End Sub

Private Sub Command1_Click()

End Sub

Private Sub d_Codclie_DblClick()
' Rutina para cambios de codigos de proveedores
'========

Dim wfecha
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Then
Else
   Exit Sub
End If
If LOC_TIPMOV <> 20 Then Exit Sub

wfecha = InputBox("Cambio de Codigo de Proveedor: " & Trim(d_nomclie.Caption), "Cambio de Codigo ", d_Codclie.Caption)
If wfecha = "" Then Exit Sub

SQ_OPER = 1
pu_codclie = Val(wfecha)
PU__CP = "P"
LEER_CLI_LLAVE
If cli_llave.EOF Then
   MsgBox "Codigo de Proveedor no Existe", 48, Pub_Titulo
   Exit Sub
End If
pub_mensaje = "Cambiar por el codigo : " & Chr(13) & cli_llave!cli_codclie & " - " & Trim(cli_llave!CLI_NOMBRE) & Chr(13) & "   ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
On Error GoTo sale_act

Screen.MousePointer = 11
pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)


SQ_OPER = 3
pu_codcia = LK_CODCIA
pu_codclie = Val(wfecha)
pu_cp = "P"
PUB_TIPDOC = "FA"
PUB_SERDOC = 0
LEER_CAR_LLAVE
If car_menor.EOF Then
   PUB_NUMDOC = 1
Else
   PUB_NUMDOC = car_menor!car_NUMDOC + 1
End If


pub_cadena = "SELECT ALL_CODCLIE, ALL_NUMDOC  FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_TIPMOV = ? AND ALL_FECHA_DIA = ? AND ALL_NUMSER = ? AND ALL_NUMFAC = ? "
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = 0
ALLUSU(1) = 0
ALLUSU(2) = 0
ALLUSU(3) = 0
ALLUSU(4) = 0
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
all_codusu.Requery
Do Until all_codusu.EOF
 all_codusu.Edit
 all_codusu!ALL_CODCLIE = Val(wfecha)
 all_codusu!ALL_NUMDOC = PUB_NUMDOC
 all_codusu.Update
 all_codusu.MoveNext
 Loop

pub_cadena = "SELECT FAR_CODCLIE, FAR_NUMDOC FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? "
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text
far_codusu.Requery
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!far_codclie = Val(wfecha)
 far_codusu!FAR_NUMDOC = PUB_NUMDOC
 far_codusu.Update
 far_codusu.MoveNext
Loop


pub_cadena = "SELECT CAR_CODCLIE, CAR_NUMDOC FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_TIPMOV = ? AND CAR_FECHA_INGR = ? AND CAR_NUMSER = ? AND CAR_NUMFAC = ? "
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
Do Until all_codusu.EOF
 all_codusu.Edit
 all_codusu!CAR_codclie = Val(wfecha)
 all_codusu!car_NUMDOC = PUB_NUMDOC
 all_codusu.Update
 all_codusu.MoveNext
Loop

pub_cadena = "SELECT CAA_CODCLIE, CAA_NUMDOC FROM CARACU WHERE CAA_CODCIA = ? AND CAA_TIPMOV = ? AND CAA_FECHA = ? AND CAA_NUMSER = ? AND CAA_NUMFAC = ? "
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
Do Until all_codusu.EOF
 all_codusu.Edit
 all_codusu!CAA_CODCLIE = Val(wfecha)
 all_codusu!CAA_NUMDOC = PUB_NUMDOC
 all_codusu.Update
 all_codusu.MoveNext
Loop
con_llave.Close
Screen.MousePointer = 0
MsgBox "Ok Cambiado.", 48


Exit Sub
sale_act:
con_llave.Close
MsgBox Err.Description, 48, Pub_Titulo
  
End Sub

Private Sub d_codven_DblClick()
Dim loc_acceso As String
On Error GoTo sale

If LOC_TIPMOV = 10 Or LOC_TIPMOV = 20 Or LOC_TIPMOV = 97 Then
Else
 Exit Sub
End If
If Trim(d_codven.Caption) = "" And LOC_TIPMOV <> 20 And LOC_TIPMOV <> 97 Then Exit Sub

loc_acceso = ""
For fila = 1 To lk_OTROS_Count
   If Val(lk_OTROS(fila)) = 34 Then loc_acceso = "A"
Next fila
If loc_acceso = "A" Then
Else
Exit Sub
End If

Dim FARUSU As rdoQuery
Dim far_codusu As rdoResultset
Dim ALLUSU As rdoQuery
Dim all_codusu As rdoResultset
wcodusu = InputBox("Ingrese el Nuevo Vendedor", "Cambio Vendedor", d_codven.Caption)
If wcodusu = "" Then Exit Sub
pu_codcia = LK_CODCIA
PUB_CODVEN = Val(wcodusu)
LEER_VEN_LLAVE
If ven_llave.EOF Then
   MsgBox "Vendedor No Existe", 48, Pub_Titulo
   Exit Sub
End If
  

pub_cadena = "SELECT ALL_CODVEN FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_TIPMOV = ? AND ALL_FECHA_DIA = ? AND ALL_NUMSER = ? AND ALL_NUMFAC = ? "
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)

Do Until all_codusu.EOF
' Print all_codusu!ALL_FECHA_DIA
' Print all_codusu!ALL_FLAG_SO
 all_codusu.Edit
 all_codusu!ALL_CODVEN = wcodusu
 all_codusu.Update
 all_codusu.MoveNext
 Loop
pub_cadena = "SELECT FAR_CODVEN FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? "
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)

far_codusu.Requery
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!FAR_CODVEN = wcodusu
 far_codusu.Update
 far_codusu.MoveNext
Loop

pub_cadena = "SELECT CAR_CODCLIE , CAR_SERDOC,CAR_NUMDOC, CAR_CODVEN, CAR_COBRADOR FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_TIPMOV = ? AND CAR_FECHA_INGR = ? AND CAR_NUMSER = ? AND CAR_NUMFAC = ? "
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
Do Until all_codusu.EOF
    all_codusu.Edit
    all_codusu!CAR_codven = wcodusu
    all_codusu!CAR_COBRADOR = wcodusu
    all_codusu.Update
    pu_codclie = all_codusu!CAR_codclie
    PUB_SERDOC = all_codusu!car_serdoc
    PUB_NUMDOC = all_codusu!car_NUMDOC
    
    pub_cadena = "SELECT CAA_CODVEN FROM CARACU WHERE CAA_CODCIA = '" & LK_CODCIA & "' AND CAA_CODCLIE = '" & pu_codclie & "'  AND CAA_SERDOC = " & PUB_SERDOC & " AND CAA_NUMDOC = " & PUB_NUMDOC
    Set FARUSU = CN.CreateQuery("", pub_cadena)
    Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
    Do Until far_codusu.EOF
    far_codusu.Edit
    far_codusu!CAA_CODVEN = wcodusu
    far_codusu.Update
    far_codusu.MoveNext
    Loop
 all_codusu.MoveNext
Loop

pub_cadena = "SELECT PED_CODVEN FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_FECHA = ? AND PED_NUMSER = ? AND PED_NUMFAC = ?  "
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = 201
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = 100
ALLUSU(4) = LOC_NUMPED
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
Do Until all_codusu.EOF
' Print all_codusu!ALL_FECHA_DIA
' Print all_codusu!ALL_FLAG_SO
 all_codusu.Edit
 all_codusu!ped_codven = wcodusu
 all_codusu.Update
 all_codusu.MoveNext
Loop


MsgBox "Cambio efectuado.", 48, Pub_Titulo
txtnumfac_KeyPress 13
Exit Sub
sale:
MsgBox Err.Description & " Intente Nuevamente...", 48, Pub_Titulo
End Sub

Private Sub d_descto_DblClick()
Dim TEXTO_DESC As String
Dim SUMA_VALOR As Currency
Dim cap_valor
Dim wcanti  As Currency
Dim wpeso  As Currency

If LOC_TIPMOV <> 20 Then Exit Sub
  If Format(d_fecha.Caption, "dd/mm/yyyy") <> Format(LK_FECHA_DIA, "dd/mm/yyyy") Then
     MsgBox "Documento Revisado.", 48, Pub_Titulo
     LK_ACCESO_REPORT = ""
     Load frmclave2
     Screen.MousePointer = 0
     frmclave2.Show 1
     If LK_ACCESO_REPORT <> "A" Then
        Exit Sub
     End If
  End If

 cap_valor = InputBox("Modificación de Descto de Mercaderia  en valor en Nuevos Soles S/.= " & Chr(13) & "el valor afecta a costo promedio mas no al documento.", " Descto(%)", d_descto.Caption)
 If cap_valor = "" Then Exit Sub
 If Val(cap_valor) = 0 Then
  pub_mensaje = "Valor S/. 0.00( para el Descto... desea continuar... "
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta <> vbYes Then
   Exit Sub
  End If
 End If
 If Val(d_neto.Caption) = 0 Then Exit Sub
 fila = 1
 SUMA_VALOR = 0
 Do Until fila = grid_fac2.Rows
     grid_fac2.Row = fila
     grid_fac2.COL = 6
     If grid_fac2.CellBackColor <> vbYellow Then
        SUMA_VALOR = SUMA_VALOR + Val(grid_fac2.TextMatrix(fila, 5))
     End If
     fila = fila + 1
 Loop
If SUMA_VALOR <> 0 Then
 cap_valor = Format((cap_valor / SUMA_VALOR) * 100, "0.00000")
 End If
 wcanti = 0
 wpeso = 0
 fila = 1
 Do Until fila = grid_fac2.Rows
     grid_fac2.Row = fila
     grid_fac2.COL = 6
     If grid_fac2.CellBackColor <> vbYellow Then
       grid_fac2.TextMatrix(fila, 6) = Format(Val(grid_fac2.TextMatrix(fila, 5)) * (Val(cap_valor) / 100), "0.0000")
       ww_desc = ww_desc + Val(grid_fac2.TextMatrix(fila, 6))
     Else
       'grid_fac2.TextMatrix(fila, 6) = "0"
     End If
     
     fila = fila + 1
 Loop
pub_mensaje = "Chequear los datos del calculo. Total de Descto en " & d_moneda.Caption & " = " & Format(ww_desc, "0.00") & " - Confirmar la modificaciro n ? "
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   txtnumfac_KeyPress 13
   Exit Sub
End If
   


pub_cadena = "SELECT FAR_PORDESCTOS , FAR_TOT_DESCTO, FAR_DESCTO FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ?  AND FAR_ESTADO <> 'E' AND far_estado2 <> 'L' ORDER BY FAR_NUMSER, FAR_NUMFAC, FAR_NUMSEC "
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text
far_codusu.Requery
fila = 1
SUMA_VALOR = 0
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!FAR_DESCTO = redondea(Val(grid_fac2.TextMatrix(fila, 6)))
 far_codusu!far_PORDESCTOS = redondea(Val(grid_fac2.TextMatrix(fila, 6)))
 far_codusu!FAR_TOT_DESCTO = redondea(cap_valor)
 SUMA_VALOR = SUMA_VALOR + Val(far_codusu!far_PORDESCTOS)
 far_codusu.Update
 far_codusu.MoveNext
 fila = fila + 1
Loop
TEXTO_DESC = "Desc. = " & Format(SUMA_VALOR, "##,##0.00")
pub_cadena = "SELECT CAR_CONCEPTO FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_TIPMOV = ? AND CAR_FECHA_INGR = ? AND CAR_NUMSER = ? AND CAR_NUMFAC = ? ORDER BY CAR_NUMFAC"
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = Val(txtSerie.Text)
FARUSU(4) = txtNumfac.Text
far_codusu.Requery
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!car_concepto = TEXTO_DESC
 far_codusu.Update
 far_codusu.MoveNext
Loop



MsgBox "Ok. Descto  Modificado." & Chr(13) & "Todos los articulos del documento deben ser costeados nuevamemte.", 48, Pub_Titulo
txtnumfac_KeyPress 13
 

End Sub

Private Sub d_dire_DblClick()
FRADIRE.Visible = True
Azul txtdire, txtdire
End Sub

Private Sub d_fecha_can_DblClick()
Dim wfecha
Exit Sub
wfecha = InputBox("Modificar F. de Almacen ", "Fecha de Almacen ", Format(d_fecha_can, "dd/mm/yyyy"))
If wfecha = "" Then Exit Sub
If Not IsDate(wfecha) Then
 MsgBox "Fecha Incorrecta.., No Procede ", 48, Pub_Titulo
 Exit Sub
End If
pub_cadena = "SELECT ALL_FECHA_CAN  FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_TIPMOV = ? AND ALL_FECHA_DIA = ? AND ALL_NUMSER = ? AND ALL_NUMFAC = ? ORDER BY ALL_NUMFAC"
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = 0
ALLUSU(1) = 0
ALLUSU(2) = 0
ALLUSU(3) = 0
ALLUSU(4) = 0
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
all_codusu.Requery
Do Until all_codusu.EOF
 all_codusu.Edit
 all_codusu!ALL_FECHA_CAN = Format(wfecha, "dd/mm/yyyy")
 all_codusu.Update
 all_codusu.MoveNext
 Loop

pub_cadena = "SELECT FAR_FECHA_CAN FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? ORDER BY FAR_NUMFAC"
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text
far_codusu.Requery
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!FAR_fecha_can = Format(wfecha, "dd/mm/yyyy")
 far_codusu.Update
 far_codusu.MoveNext
Loop


d_fecha_can.Caption = Format(wfecha, "dd/mm/yyyy")


MsgBox "Ok Cambiado.", 48

End Sub

Private Sub d_fecha_compra_DblClick()
Dim wfecha
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
   Exit Sub
End If
wfecha = InputBox("Modificar Fecha Contable ", "Fecha Contable", Format(d_fecha_compra, "dd/mm/yyyy"))
If wfecha = "" Then Exit Sub
If Not IsDate(wfecha) Then
 MsgBox "Fecha Incorrecta.., No Procede ", 48, Pub_Titulo
 Exit Sub
End If
pub_cadena = "SELECT ALL_FECHA_SUNAT  FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_TIPMOV = ? AND ALL_FECHA_DIA = ? AND ALL_NUMSER = ? AND ALL_NUMFAC = ? ORDER BY ALL_NUMFAC"
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = 0
ALLUSU(1) = 0
ALLUSU(2) = 0
ALLUSU(3) = 0
ALLUSU(4) = 0
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
all_codusu.Requery
Do Until all_codusu.EOF
 all_codusu.Edit
 all_codusu!ALL_FECHA_SUNAT = Format(wfecha, "dd/mm/yyyy")
 all_codusu.Update
 all_codusu.MoveNext
 Loop

pub_cadena = "SELECT FAR_FECHA_COMPRA FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? ORDER BY FAR_NUMFAC"
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text
far_codusu.Requery
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!FAR_fecha_compra = Format(wfecha, "dd/mm/yyyy")
 far_codusu.Update
 far_codusu.MoveNext
Loop


FCONT.Caption = "Fec. Contable : " & Format(wfecha, "dd/mm/yyyy")


pub_cadena = "SELECT CAR_FECHA_SUNAT FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_TIPMOV = ? AND CAR_FECHA_INGR = ? AND CAR_NUMSER = ? AND CAR_NUMFAC = ? "
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
Do Until all_codusu.EOF
' Print all_codusu!ALL_FECHA_DIA
' Print all_codusu!ALL_FLAG_SO
 all_codusu.Edit
 all_codusu!car_fecha_sunat = Format(wfecha, "dd/mm/yyyy")
 all_codusu.Update
 all_codusu.MoveNext
Loop
If LOC_TIPMOV = 97 Then
pub_cadena = "SELECT CAA_FECHA_COBRO FROM CARACU WHERE CAA_CODCIA = ? AND CAA_TIPMOV = ? AND CAA_FECHA = ? AND CAA_NUMSER = ? AND CAA_NUMFAC = ? "
Else
pub_cadena = "SELECT CAA_FECHA_COBRO FROM CARACU WHERE CAA_CODCIA = ? AND CAA_TIPMOV = ? AND CAA_FECHA = ? AND CAA_NUMSER = ? AND CAA_NUMFAC = ? AND CAA_SIGNO_CAR = 1 "
End If
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
Do Until all_codusu.EOF
 all_codusu.Edit
 all_codusu!CAA_FECHA_COBRO = Format(wfecha, "dd/mm/yyyy")
 all_codusu.Update
 all_codusu.MoveNext
Loop

MsgBox "Ok Cambiado.", 48
End Sub

Private Sub d_flete_DblClick()
Dim cap_valor
Dim wcanti  As Currency
Dim wpeso  As Currency

If LOC_TIPMOV <> 20 Then Exit Sub
   If Format(d_fecha.Caption, "dd/mm/yyyy") <> Format(LK_FECHA_DIA, "dd/mm/yyyy") Then
     MsgBox "Documento Revisado.", 48, Pub_Titulo
     LK_ACCESO_REPORT = ""
     Load frmclave2
     Screen.MousePointer = 0
     frmclave2.Show 1
     If LK_ACCESO_REPORT <> "A" Then
        Exit Sub
     End If
  End If


 cap_valor = InputBox("Modificación de Flete en Mercaderia  S/. = ", "Fletes en S/.", d_flete.Caption)
 If cap_valor = "" Then Exit Sub
 If Val(cap_valor) = 0 Then
  pub_mensaje = "Valor 0.00 para el Flete... desea continuar... "
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta <> vbYes Then
   Exit Sub
  End If
 End If
 
 wcanti = 0
 wpeso = 0
 fila = 1
 Do Until fila = grid_fac2.Rows
   grid_fac2.Row = fila
   grid_fac2.COL = 7
   If grid_fac2.CellBackColor <> vbBlue Then
    wcanti = wcanti + Val(grid_fac2.TextMatrix(fila, 2))
    wpeso = wpeso + Val(grid_fac2.TextMatrix(fila, 5))
   End If
   fila = fila + 1
 Loop
   

pub_mensaje = "Flete x Cantidad (Si), x Valor (No)...."
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   ww_desc = 0
   fila = 1
   Do Until fila = grid_fac2.Rows
    grid_fac2.Row = fila
    grid_fac2.COL = 7
    If grid_fac2.CellBackColor <> vbBlue Then
      If Pub_Respuesta = vbYes Then
        If wcanti <> 0 Then
            grid_fac2.TextMatrix(fila, 7) = Format((Val(cap_valor) / wcanti) * Val(grid_fac2.TextMatrix(fila, 2)), "0.00")
        End If
      Else
        If wpeso <> 0 Then
            grid_fac2.TextMatrix(fila, 7) = Format((Val(cap_valor) / wpeso) * grid_fac2.TextMatrix(fila, 5), "0.00")
        End If
      End If
     ww_desc = ww_desc + Val(grid_fac2.TextMatrix(fila, 7))
    End If
     fila = fila + 1
   Loop

pub_mensaje = "Chequear los datos del calculo. Total de Flete : " & Format(ww_desc, "0.00") & " - Confirmar la modificaciro n ? "
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   txtnumfac_KeyPress 13
   Exit Sub
End If
   

'Exit Sub


pub_cadena = "SELECT FAR_TOT_FLETE, FAR_FLETE FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ?  AND FAR_ESTADO <> 'E' AND FAR_ESTADO2 <>  'L' ORDER BY FAR_NUMSER, FAR_NUMFAC, FAR_NUMSEC "
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text
far_codusu.Requery
fila = 1
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!FAR_FLETE = Val(grid_fac2.TextMatrix(fila, 7))
 far_codusu!FAR_TOT_FLETE = ww_desc
 far_codusu.Update
 far_codusu.MoveNext
 fila = fila + 1
Loop

MsgBox "Ok. Fletes  Modificado." & Chr(13) & "Todos los articulos del documento deben ser costeados nuevamemte.", 48, Pub_Titulo
txtnumfac_KeyPress 13
 
End Sub

Private Sub d_usuario_Click()
Dim loc_acceso As String

If Trim(d_usuario.Caption) = "" Then Exit Sub
loc_acceso = ""
For fila = 1 To lk_OTROS_Count
   If Val(lk_OTROS(fila)) = 33 Then loc_acceso = "A"
Next fila
If loc_acceso = "A" Then
Else
Exit Sub
End If

Dim FARUSU As rdoQuery
Dim far_codusu As rdoResultset
Dim ALLUSU As rdoQuery
Dim all_codusu As rdoResultset
wcodusu = InputBox("Ingrese Usuario", "Cambio de Usuario", d_usuario.Caption)
If wcodusu = "" Then Exit Sub
wcodusu = UCase(wcodusu)
PSUSU_LLAVE(0) = UCase(wcodusu)
usu_llave.Requery
If usu_llave.EOF Then
   MsgBox "Usuario no Existe", 48, Pub_Titulo
   Exit Sub
End If
  

pub_cadena = "SELECT ALL_CODUSU FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_TIPMOV = ? AND ALL_FECHA_DIA = ? AND ALL_NUMSER = ? AND ALL_NUMFAC = ? ORDER BY ALL_NUMFAC"
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = 0
ALLUSU(1) = 0
ALLUSU(2) = 0
ALLUSU(3) = 0
ALLUSU(4) = 0
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text

all_codusu.Requery
Do Until all_codusu.EOF
' Print all_codusu!ALL_FECHA_DIA
' Print all_codusu!ALL_FLAG_SO
 all_codusu.Edit
 all_codusu!all_codusu = wcodusu
 all_codusu.Update
 all_codusu.MoveNext
 Loop
pub_cadena = "SELECT FAR_CODUSU FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? ORDER BY FAR_NUMFAC"
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text

far_codusu.Requery
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!far_codusu = wcodusu
 far_codusu.Update
 far_codusu.MoveNext
Loop

MsgBox "Cambio efectuado.", 48, Pub_Titulo

End Sub

Private Sub DOCANEXO_Click()

If loc_acceso_descto <> "A" Then
  MsgBox "NO tiene Acceso a esta Opción...", 48, Pub_Titulo
  Exit Sub
End If

Load FrmControl
FrmControl.d_codigo.Caption = d_Codclie.Caption
FrmControl.d_nombre.Caption = d_nomclie.Caption
FrmControl.d_fecha.Caption = d_fecha_compra.Caption
FrmControl.d_serie.Caption = temp_NUMSER_C
FrmControl.d_numero.Caption = temp_NUMFAC_C
FrmControl.d_numser.Caption = txtSerie.Text
FrmControl.d_numfac.Caption = txtNumfac.Text
FrmControl.d_nombre.Caption = d_nomclie.Caption
FrmControl.flag_signo.Caption = "-1"

FrmControl.Show 1


End Sub

Private Sub fbg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtcserie, txtcserie
End Sub

Private Sub FCONT_DblClick()
Dim wfecha
Exit Sub
wfecha = InputBox("Modificar Fecha Contable ", "Fecha Contable", "")
If wfecha = "" Then Exit Sub
If Not IsDate(wfecha) Then
 MsgBox "Fecha Incorrecta.., No Procede ", 48, Pub_Titulo
 Exit Sub
End If
pub_cadena = "SELECT ALL_FECHA_PRO  FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_TIPMOV = ? AND ALL_FECHA_DIA = ? AND ALL_NUMSER = ? AND ALL_NUMFAC = ? ORDER BY ALL_NUMFAC"
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = 0
ALLUSU(1) = 0
ALLUSU(2) = 0
ALLUSU(3) = 0
ALLUSU(4) = 0
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
all_codusu.Requery
Do Until all_codusu.EOF
 all_codusu.Edit
 all_codusu!ALL_FECHA_PRO = Format(wfecha, "dd/mm/yyyy")
 all_codusu.Update
 all_codusu.MoveNext
 Loop

pub_cadena = "SELECT FAR_FECHA_PRO FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? ORDER BY FAR_NUMFAC"
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text
far_codusu.Requery
Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!FAR_fecha_pro = Format(wfecha, "dd/mm/yyyy")
 far_codusu.Update
 far_codusu.MoveNext
Loop


FCONT.Caption = "Fec. Contable : " & Format(wfecha, "dd/mm/yyyy")

End Sub

Private Sub Form_Activate()
If Trim(TIPMOV.Text) = "" Then
  TIPMOV.SetFocus
End If
End Sub

Private Sub Form_Load()
    CenterMe frmdocu
    frmdocu.Width = 12120
    frmdocu.Height = 8500
Dim SQL As String
Unload FORMGEN
'Unload FORM_GRIFO
'*********************************************************
pu_cp = ""
SQL = "SELECT DIRCOMP FROM DIRCLI " & _
"WHERE CODCIA=? AND DIRCLI=? AND CODCLI=? AND CP=?"
Set PS = CN.CreateQuery("", SQL)
  PS.rdoParameters(0) = " "
  PS.rdoParameters(1) = 0
  PS.rdoParameters(2) = 0
  PS.rdoParameters(3) = " "
  Set rs = PS.OpenResultset(rdOpenKeyset, rdConcurValues)
'*********************************************************
loc_flag_espera = ""
LOC_ARROZ = ""
pub_cadena = "SELECT * FROM TRANSPORTE WHERE TRN_KEY = ? ORDER BY TRN_KEY"
Set PS_TRA = CN.CreateQuery("", pub_cadena)
PS_TRA(0) = 0
Set llave_trans = PS_TRA.OpenResultset(rdOpenKeyset, rdConcurValues)

LOC_NUMFAC_FIN = 0
pub_cadena = "SELECT * FROM TRANSPORTE WHERE TRN_KEY >= ? ORDER BY TRN_NOMBRE"
Set pstransporte = CN.CreateQuery("", pub_cadena)
pstransporte.rdoParameters(0) = 0
Set transporte = pstransporte.OpenResultset(rdOpenKeyset, rdConcurValues)
pstransporte(0) = 0
transporte.Requery
TRANS.Clear
Do Until transporte.EOF
    TRANS.AddItem Trim(transporte!TRN_NOMBRE) & String(80, " ") & transporte!TRN_KEY
    transporte.MoveNext
Loop
TRANS.Enabled = False
pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC = ?  AND FAR_ESTADO2 <> 'L'  ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
Set PSFAR = CN.CreateQuery("", pub_cadena)
PSFAR.rdoParameters(0) = 0
PSFAR.rdoParameters(1) = 0
PSFAR.rdoParameters(2) = 0
PSFAR.rdoParameters(3) = 0
PSFAR.rdoParameters(4) = 0
Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurValues)
'SELE_DOCU
pub_cadena = "SELECT FAR_numfac FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? AND FAR_ESTADO <> 'E' AND FAR_CP = ?  ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_FBG , FAR_NUMSER, FAR_NUMFAC DESC"
Set PSFAR_CONSUL = CN.CreateQuery("", pub_cadena)
PSFAR_CONSUL.rdoParameters(0) = 0
PSFAR_CONSUL.rdoParameters(1) = " "
PSFAR_CONSUL.rdoParameters(2) = " "
PSFAR_CONSUL.rdoParameters(3) = 0
PSFAR_CONSUL.rdoParameters(4) = " "
PSFAR_CONSUL.MaxRows = 1
Set far_consul = PSFAR_CONSUL.OpenResultset(rdOpenKeyset, rdConcurValues) '.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ?  AND CAR_FBG = ? AND CAR_NUMSER = ? AND CAR_NUMFAC = ? AND CAR_IMPORTE <> 0 ORDER BY CAR_CODCIA, CAR_CODCLIE"
Set PSCAR_CONSUL = CN.CreateQuery("", pub_cadena)
PSCAR_CONSUL.rdoParameters(0) = " "
PSCAR_CONSUL.rdoParameters(1) = " "
PSCAR_CONSUL.rdoParameters(2) = 0
PSCAR_CONSUL.rdoParameters(3) = " "
PSCAR_CONSUL.rdoParameters(4) = 0
PSCAR_CONSUL.rdoParameters(5) = 0
Set car_consul = PSCAR_CONSUL.OpenResultset(rdOpenKeyset, rdConcurValues)
PUB_CODCIA = "00"
LLENADOS TIPMOV, 4
LOC_TIPMOV = 0
temporal = "X"
cmbFBG.Clear
lblflete.Caption = "Ajuste"
lblNumfac.Caption = "Nro  de Doc."
temporal = ""
tempo_serie = ""
WGUIA_RELA = ""
SERIE_GUIA = ""
If LK_FLAG_GRIFO = "A" Then
End If
If LK_FLAG_FACTURACION = "V" Then
 txtvend.Visible = True
 lblvend.Visible = True
Else
 txtvend.Visible = False
 lblvend.Visible = False
End If
lblvend.Caption = "Vend."
If LK_FLAG_GRIFO = "A" Then
  lblvend.Caption = "Isla"
End If

LLENA_ZONA TxtZonaTrabajo, 20
LLENA_ZONA TxtSubZonaTrabajo, 35
loc_acceso_descto = ""
loc_acceso_cruzados = ""
For fila = 1 To lk_OTROS_Count
   If Val(lk_OTROS(fila)) = 16 Then 'modificar Venta
     loc_acceso_descto = "A"
   End If
   If Val(lk_OTROS(fila)) = 24 Then 'Cambio de doc. cruzados
     loc_acceso_cruzados = "A"
   End If

Next fila

cmdcambios.Visible = False
If LK_CODUSU = "SUPERVISOR" Or LK_CODUSU = "ADMIN" Or loc_acceso_cruzados = "A" Then
  cmdcambios.Visible = True
End If


End Sub

Public Sub LLENA_CONSULTA()
Dim WS_SUM_DESCT As Currency
Dim subtotal2 As Currency
'*** VERFICA SI HAY DATOS , O ESTAN CORRECTOS
'On Error GoTo SALE
Dim conteo_cantidad As Currency
Dim wnumser_c As String
Dim wnumfac_c As String
Dim WFECHA_CONT As Date
Dim WFECHA_COMPRA As Date
Dim WFECHA_COMPRA_CAN As Date
Dim wwsigno_car  As Integer
Dim wwdias  As Integer
Dim DESC_GRIFO As Currency
Dim WISLA
Dim WRES
Dim ws_serie As Integer
Dim WS_CUENTA As Integer
Dim suma_subtotal As Currency

If LOC_TIPMOV = 10 Then NUMERO.Text = 1
If LOC_TIPMOV = 100 Then NUMERO.Text = 5
If LOC_TIPMOV = 3 Then NUMERO.Text = 7

DESC_GRIFO = 0
WS_FLETE = 0
ws_serie = 0
ws_serie = Val(txtSerie.Text)
CmdAnterior.Enabled = False
'DoEvents
cmdSiguiente.Enabled = False
cmdImp.Enabled = False
'DoEvents
cherela.Visible = False
tguia.Visible = False
chetrans.Visible = False
TRANS.Visible = False
l_fecha_compra.Visible = False
d_fecha_compra.Visible = False
d_fecha_can.Visible = False
l_fecha_can.Visible = False
FECHA_PART.Visible = False
l_fecha_compra.Visible = True
d_fecha_compra.Visible = True
d_fecha_can.Visible = True
l_fecha_can.Visible = True
lblcargo.Caption = ""
lblabono.Caption = ""
If LOC_TIPMOV = 20 Then
    lbletiqaj(0).Caption = "Cargo"
    lbletiqaj(1).Caption = "Abono"
ElseIf LOC_TIPMOV = 97 Then
    lbletiqaj(0).Caption = "Abono"
    lbletiqaj(1).Caption = "Cargo"
Else
    lbletiqaj(0).Caption = ""
    lbletiqaj(1).Caption = ""
End If
If LOC_TIPMOV = 100 Or LOC_TIPMOV = 3 Then

 l_fecha_compra.Visible = True
 d_fecha_compra.Visible = True
 FECHA_PART.Visible = True
 
 tguia.Visible = True
             
 cherela.Visible = True
 chetrans.Visible = True
 tguia.ToolTipText = "F/B/G = @GUIA y  G.Rem. = @NUNDOC"
 TRANS.Visible = True
End If

If LOC_TIPMOV = 10 Then

 l_fecha_compra.Visible = True
 d_fecha_compra.Visible = True
 FECHA_PART.Visible = True
 
 tguia.Visible = True

 cherela.Visible = True
 chetrans.Visible = True
 tguia.ToolTipText = "F/B/G = @GUIA y  G.Rem. = @NUNDOC"
 TRANS.Visible = True
 PSFAR.rdoParameters(0) = LOC_TIPMOV
 PSFAR.rdoParameters(2) = ws_serie
 PSFAR.rdoParameters(3) = Left(cmbFBG.Text, 1)
ElseIf LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
 PSFAR.rdoParameters(0) = LOC_TIPMOV
 PSFAR.rdoParameters(2) = ws_serie
 PSFAR.rdoParameters(3) = Left(cmbFBG.Text, 1)
 If LOC_TIPMOV = 97 Then  'ALAN DIJO QUE SALIA ERROR ???
  PSFAR.rdoParameters(5) = Right(cmbFBG.Text, 1)
 End If
ElseIf LOC_TIPMOV = 20 Or LOC_TIPMOV = 99 Then
 If LOC_TIPMOV = 20 And Left(cmbFBG.Text, 1) <> "K" Then
   If Left(cmbFBG.Text, 1) = "F" Then
     PSFAR.rdoParameters(0) = LOC_TIPMOV
     PSFAR.rdoParameters(1) = LK_CODCIA
     PSFAR.rdoParameters(2) = txtSerie.Text
     PSFAR.rdoParameters(3) = Val(txtNumfac.Text)
   Else
     PSFAR.rdoParameters(0) = LOC_TIPMOV
     PSFAR.rdoParameters(1) = LK_CODCIA
     PSFAR.rdoParameters(2) = txtNumfac.Text
   End If
 Else
   PSFAR.rdoParameters(0) = LOC_TIPMOV
   PSFAR.rdoParameters(2) = ws_serie
   If LOC_TIPMOV = 99 Then
      PSFAR.rdoParameters(3) = Left(cmbFBG.Text, 1)
   Else
     PSFAR.rdoParameters(3) = " "
   End If
 End If
Else
 PSFAR.rdoParameters(0) = LOC_TIPMOV
 PSFAR.rdoParameters(2) = ws_serie
 PSFAR.rdoParameters(3) = ""
End If
If LOC_TIPMOV = 20 And Left(cmbFBG.Text, 1) <> "K" Then
Else
 PSFAR.rdoParameters(1) = LK_CODCIA
 PSFAR.rdoParameters(4) = Val(txtNumfac.Text)
End If
far_r.Requery
If far_r.EOF Then
   wflag_docu = "A"
   CmdAnterior.Enabled = True
   cmdSiguiente.Enabled = True
   cmdImp.Enabled = True
   Exit Sub
Else
  If LK_FLAG_SOS = "A" And far_r!FAR_FLAG_SO <> "A" Then
    wflag_docu = "A"
    CmdAnterior.Enabled = True
    cmdSiguiente.Enabled = True
    cmdImp.Enabled = True
    Exit Sub
  End If
End If
tempo_serie = Trim(txtSerie.Text)
pbValue = 0
pbMin = 0
pbMax = far_r.RowCount + 3
PB.Visible = True
'DoEvents
LIMPIA_DOCU
'MsgBox far_r!far_NUMOPER
'MsgBox far_r!far_CONCEPTO
wflag_activo = 0
WS_CUENTA = far_r.RowCount
fila = 0
Do Until far_r.EOF
 If Nulo_Valors(far_r!far_estado) <> "E" Then
   wflag_activo = 1
   GoTo muestra
   Exit Do
 End If
 far_r.MoveNext
Loop
far_r.MoveLast
fila = far_r!far_numsec
Do Until far_r.BOF
  wfecha_anu = far_r!far_fecha
  If fila <> far_r!far_numsec Then
    far_r.MoveNext
    GoTo muestra
  End If
  fila = fila - 1
  far_r.MovePrevious
Loop
far_r.MoveFirst

muestra:
'd_fecha.Caption = Format(far_r!FAR_fecha, "dd/mm/yyyy")
'If Nulo_Valors(far_r!far_estado) = "E" Then
'  If wfecha_anu <> LK_FECHA_DIA Then
'    LBLEXTORNO.Caption = "DOCUMENTO A N U L A D O"
'    d_fecha.Caption = Format(wfecha_anu, "dd/mm/yyyy")
'  Else
'    If far_r!FAR_fecha = LK_FECHA_DIA Then
'      LBLEXTORNO.Caption = "DOCUMENTO EXTORNADO"
'    End If
'  End If
'  LBLEXTORNO.Visible = True
'Else
'  LBLEXTORNO.Visible = False
'End If
DOC102 = ""

SQ_OPER = 1
pu_codcia = LK_CODCIA
If LOC_TIPMOV = 93 Or LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Or LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Then
 lbldomicilio.Caption = ""
 lbldireccion.Caption = "Concepto :"
 'QUITADO 30/11/2001
 'd_dire.Caption = Trim(far_r!far_concepto)
 'd_dire.ToolTipText = Trim(d_dire.Caption)
  If LOC_TIPMOV = 5 Then
   tguia.Visible = True
   cherela.Visible = True
   chetrans.Visible = True
   tguia.ToolTipText = "F/B/G = @GUIA y  G.Rem. = @NUNDOC"
   TRANS.Visible = True
 End If
Else
 lbldireccion.Caption = "Dirección Entrega:"
 lbldomicilio.Caption = "Domicilio:"
 
End If
If LOC_TIPMOV = 102 Or OC_TIPMOV = 100 Or LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
  
  txtdocu.Visible = True
  pu_cp = "C"
 lbldomicilio.Caption = "Domicilio:"
 PU_FBG = far_r!far_fbg
 WGUIA_RELA = far_r!far_numguia
 SERIE_GUIA = far_r!far_serguia
 Doc_EL = Nulo_Valors(far_r!FAR_DOC_ELECTRONICO)
 
 If LK_EMP <> "HER" Then
   txtdocu.Caption = "G/ " & far_r!far_serguia & " - " & far_r!far_numguia & " - Nro .O/C :" & Trim(Nulo_Valors(far_r!FAR_OC))
 Else
   txtdocu.Caption = "G/ " & far_r!far_serguia & " - " & far_r!far_numguia
 End If
 If Val(WGUIA_RELA) <> 0 Then
   cherela.Value = 1
 End If
 If LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
     txtdocu.Caption = Trim(far_r!far_concepto)  ' serguia & " - " & far_r!far_numguia
     pu_cp = Right(cmbFBG.Text, 1)
     lblmotivo.Visible = True
     If (LOC_TIPMOV = 97 And pu_cp = "P") Then
        lblcargo.Caption = Val(far_r!far_limcre_ant)
        lblabono.Caption = Val(far_r!far_limcre_act)
     End If
 End If
 
 
 lblsaldo.Visible = True
 d_saldo.Visible = True
 d_moneda.Visible = True
ElseIf LOC_TIPMOV = 20 Or LOC_TIPMOV = 99 Or LOC_TIPMOV = 3 Then
 l_fecha_compra.Visible = True
 d_fecha_compra.Visible = True
 d_fecha_can.Visible = True
 l_fecha_can.Visible = True
 lblsaldo.Visible = True
 d_saldo.Visible = True
 
 d_moneda.Visible = True
 'lblsaldo.Visible = False
 'd_saldo.Visible = False
 If LOC_TIPMOV = 20 Then
  lblcargo.Caption = Val(far_r!far_limcre_ant)
  lblabono.Caption = Val(far_r!far_limcre_act)
 End If

 txtdocu.Visible = True
 lblfac.Visible = True
 If Val(far_r!FAR_NUMFAC_C) = 0 Then
   'txtdocu.Caption = " "
   txtdocu.Caption = "G/ " & far_r!far_serguia & " - " & far_r!far_numguia
 Else
   txtdocu.Caption = "F/ " & far_r!FAR_NUMSER_C & " - " & far_r!FAR_NUMFAC_C
 End If
 temp_NUMSER_C = far_r!FAR_NUMSER_C
 temp_NUMFAC_C = far_r!FAR_NUMFAC_C
 If LOC_TIPMOV = 99 Then
    PUB_CODCIA = "00"
    PUB_TIPREG = 50
    PUB_NUMTAB = far_r!far_cod_sunat
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
      If far_r!FAR_NUMFAC_C <> 0 Then
        txtdocu.Caption = Trim(Str(far_r!far_cod_sunat)) & "-" & Trim(Left(tab_llave!TAB_NOMLARGO, 20)) & " / " & far_r!FAR_NUMSER_C & " - " & far_r!FAR_NUMFAC_C
      Else
        txtdocu.Caption = Trim(tab_llave!TAB_NOMLARGO) & " / " & far_r!far_numguia
      End If
    End If
 End If
        
 pu_cp = far_r!FAR_cp
 'pu_cp = "C"
 'If Right(Trim(cmbFBG.Text), 1) = "P" Then
 '  pu_cp = "C"
 'End If
 
 PU_FBG = " "
Else
  GoTo PASACLI
End If

pu_codclie = far_r!far_codclie
LEER_CLI_LLAVE
If cli_llave.EOF Then
  MsgBox "Rgistro de Cliente not :" & far_r!far_codclie, 48, Pub_Titulo
  GoTo sale
End If
pbValue = pbValue + 1
d_Codclie.Caption = Trim(cli_llave!cli_codclie)
If Trim(cli_llave!cli_codclie) = 1 Then
   d_nomclie.Caption = Trim(far_r!far_cliente)
Else
  d_nomclie.Caption = Trim(cli_llave!CLI_NOMBRE)
End If


SQ_OPER = 1
PUB_CODCIA = "00"
PUB_NUMTAB = cli_llave!CLI_LUGAR_TRAB
PUB_TIPREG = 25
LEER_TAB_LLAVE
WLUGAR = ""
If Not tab_llave.EOF Then
WLUGAR = Trim(tab_llave!TAB_NOMLARGO)
End If
PUB_NUMTAB = cli_llave!CLI_LUGAR_CASA
LEER_TAB_LLAVE
WLUGAR1 = ""
If Not tab_llave.EOF Then
WLUGAR1 = Trim(tab_llave!TAB_NOMLARGO)
End If

PUB_NUMTAB = Nulo_Valor0(cli_llave!cli_TRAB_ZONA)
PUB_TIPREG = 20
LEER_TAB_LLAVE
WZONA = ""
If Not tab_llave.EOF Then
WZONA = Trim(tab_llave!TAB_NOMLARGO)
End If
PUB_NUMTAB = Nulo_Valor0(cli_llave!CLI_CASA_ZONA)
LEER_TAB_LLAVE
WZONA1 = ""
If Not tab_llave.EOF Then
WZONA1 = Trim(tab_llave!TAB_NOMLARGO)
End If

PUB_NUMTAB = cli_llave!cli_TRAB_SUBZONA
PUB_TIPREG = 35
LEER_TAB_LLAVE
WSUBZONA = ""
If Not tab_llave.EOF Then
WSUBZONA = Trim(tab_llave!TAB_NOMLARGO)
End If
PUB_NUMTAB = cli_llave!cli_zona_new
LEER_TAB_LLAVE
WSUBZONA1 = ""
If Not tab_llave.EOF Then
WSUBZONA1 = Trim(tab_llave!TAB_NOMLARGO)
End If
'QUITADO 30/11/2001
'd_dire.Caption = Trim(WLUGAR) + " " + Trim(cli_llave!CLI_TRAB_DIREC) + " # " + Trim(cli_llave!CLI_TRAB_NUM) & "  " & WZONA & "  " & WSUBZONA
txtdire.Text = Trim(cli_llave!CLI_TRAB_DIREC)
txtnum.Text = Trim(cli_llave!CLI_TRAB_NUM)
ASIGNA_INT TxtZonaTrabajo, cli_llave!cli_TRAB_ZONA
ASIGNA_INT TxtSubZonaTrabajo, cli_llave!cli_TRAB_SUBZONA

d_domicilio.Caption = Trim(WLUGAR1) + " " + Trim(cli_llave!CLI_CASA_DIREC) + " # " + Trim(cli_llave!CLI_CASA_NUM) & "  " & WZONA1 & "  " & WSUBZONA1

If Left(cmbFBG.Text, 1) = "F" Then
 d_ruc.Caption = Trim(Nulo_Valors(cli_llave!cli_ruc_esposo))
 LBLRUC.Visible = True
Else
 LBLRUC.Visible = False
 d_ruc.Caption = ""
End If

PASACLI:
If LK_FLAG_GRIFO = "A" And far_r!FAR_tipmov = 20 Then
 d_nomven.Caption = "Turno: " & Format(far_r!far_turno, "00") & Chr(13) & "Nro .Carga: " & far_r!FAR_PEDFAC
End If
If LOC_TIPMOV = 20 Then GoTo a_vendedor
If far_r!FAR_cp <> "P" Or LOC_TIPMOV = 97 Then
If LOC_TIPMOV = 102 Or LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
a_vendedor:
  If LK_FLAG_GRIFO <> "A" Then
     If txtvend.Visible Then
       txtvend.Text = Nulo_Valor0(far_r!FAR_CODVEN)
     End If
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_CODVEN = Nulo_Valor0(far_r!FAR_CODVEN)
     LEER_VEN_LLAVE
     If ven_llave.EOF Then
       MsgBox "Verificar Vendedor.", 48, Pub_Titulo
     Else
      d_codven.Caption = Nulo_Valor0(Trim(far_r!FAR_CODVEN))
      d_nomven.Caption = Trim(ven_llave!VEM_NOMBRE)
     End If
   Else
     If txtvend.Visible Then
       txtvend.Text = far_r!far_ISLA
     End If
     SQ_OPER = 1
     PUB_CODCIA = xCODCIA
     PUB_CODVEN = Nulo_Valor0(far_r!far_ISLA)
     LEER_VEN_LLAVE
     WISLA = " "
     If Not ven_llave.EOF Then
         WISLA = ven_llave!VEM_NOMBRE
     End If
     SQ_OPER = 1
     PUB_TIPREG = 2103
     PUB_CODCIA = LK_CODCIA
     PUB_NUMTAB = Val(far_r!FAR_CODVEN)
     LEER_TAB_LLAVE
     WRES = " "
     If Not tab_llave.EOF Then
        WRES = Trim(tab_llave!TAB_NOMLARGO)
     End If
     d_codven.Caption = "-"
     d_nomven.Caption = WISLA & " / " & WRES & " - " & "Turno: " & Format(far_r!far_turno, "00")
 End If
End If
End If

d_usuario.Caption = Trim(Nulo_Valors(far_r!far_codusu))

SQ_OPER = 1
PUB_TIPREG = 65
PUB_CODCIA = LK_CODCIA
PUB_NUMTAB = Val(far_r!far_turno)
LEER_TAB_LLAVE
WRES = " "
If Not tab_llave.EOF Then
   WRES = Trim(tab_llave!TAB_NOMLARGO)
End If
d_tipvta.Caption = Trim(WRES)
pu_cp = far_r!FAR_cp
pu_codcia = LK_CODCIA
PU_NUMSER = far_r!far_numser
PU_NUMFAC = far_r!far_numfac
LEER_CAR_CONSUL
If Not car_consul.EOF Then
   d_saldo.Caption = Format(car_consul!car_importe, "##,##0.00")
   d_newvcto.Caption = Format(car_consul!car_fecha_vcto, "dd/mm/yyyy")
End If

If LOC_TIPMOV = 102 Then
  SQ_OPER = 8
  PU_TIPMOV = 10
  pu_codcia = LK_CODCIA
  PU_FBG = "F"
  PU_FBG2 = "B"
  LEER_FAR_LLAVE
  If far_menor4.EOF Then
    DOC102 = ""
  Else
    DOC102 = far_menor4!far_fbg & "/ " & Format(far_menor4!far_numser, "000") & "-" & far_menor4!far_numfac
  End If
  txtdocu.Visible = True
  txtdocu.Caption = " Doc. Relacionado. = " & DOC102
End If

If LOC_TIPMOV = 100 Then
  d_dire.Caption = Nulo_Valors(far_r!far_concepto)
  d_nomclie.Caption = ""
  PSPAR_MULTI(0) = Nulo_Valors(far_r!far_otra_cia)
  par_multi.Requery
  If par_multi.EOF Then
  Else
  d_nomclie.Caption = par_multi!par_nombre
  End If
  
End If


pbValue = pbValue + 1
wwsigno_car = Nulo_Valor0(far_r!far_signo_car)
wwdias = Nulo_Valor0(far_r!FAR_DIAS)
FECHA_PART.Text = Format(far_r!FAR_fecha_compra, "dd/mm/yyyy")

' configurar despues acv
 If wwsigno_car = 1 And wwdias <> 0 Then
  If LOC_TIPMOV = 10 Then
     d_condicion.Caption = far_r!far_subtra ' far_r!FAR_secuencia ' "COMPRA AL CREDITO"
     ' d_condicion.Caption = "VENTA AL CREDITO"
  ElseIf LOC_TIPMOV = 20 Or LOC_TIPMOV = 99 Then
     d_condicion.Caption = far_r!far_subtra ' far_r!FAR_secuencia ' "COMPRA AL CREDITO"
     d_fecha_compra.Caption = Format(far_r!FAR_fecha_compra, "dd/mm/yyyy")
  Else
     d_condicion.Caption = Left(TIPMOV.Text, 40)
  End If
  d_dias.Caption = Val(far_r!FAR_DIAS)
  d_fechaV.Caption = Format(DateAdd("d", Val(far_r!FAR_DIAS), far_r!FAR_fecha_compra), "dd/mm/yyyy")
 Else
 If LOC_TIPMOV = 10 Then
    d_condicion.Caption = far_r!far_subtra ' far_r!FAR_secuencia ' "COMPRA AL CREDITO"
     'd_condicion.Caption = "VENTA AL CONTADO"
     frmdocu.d_fechaV.Caption = frmdocu.d_fecha.Caption
 ElseIf LOC_TIPMOV = 20 Then
   d_condicion.Caption = far_r!far_subtra ' "COMPRA AL CONTADO"
 Else
   d_condicion.Caption = Left(TIPMOV.Text, 40)
   If LOC_TIPMOV = 99 Or LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Then
    d_condicion.Caption = Nulo_Valors(far_r!far_subtra)
   End If
 End If
End If

If LK_FLAG_GRIFO = "A" And LOC_TIPMOV = 10 Then
  If Val(Nulo_Valor0(far_r!far_signo_car)) <> 1 Then
     d_condicion.Caption = "VENTA AL CONTADO"
  Else
     d_condicion.Caption = "VENTA AL CREDITO"
  End If
End If
'lblven.Caption = "Vendedor"
If Nulo_Valors(far_r!FAR_MONEDA) = "D" Then
  d_moneda.Caption = "$."
  tmoneda.Text = "US$."
'  lblven.Visible = True
  If LOC_TIPMOV = 20 Then
'   lblven.Caption = "T. Cambio:"
   d_codven.Caption = Val(Nulo_Valor0(far_r!FAR_tipo_cambio))
  End If
Else
  d_moneda.Caption = "S/."
  tmoneda.Text = "S/."
End If
d_fecha_compra.Caption = Format(far_r!FAR_fecha_compra, "dd/mm/yyyy")
d_fecha.Caption = Format(far_r!far_fecha, "dd/mm/yyyy")
wfecha_anu = far_r!far_fecha
If Nulo_Valors(far_r!far_estado) = "E" Then
  If wfecha_anu <> LK_FECHA_DIA Then
    LBLEXTORNO.Caption = "DOCUMENTO A N U L A D O"
    d_fecha.Caption = Format(wfecha_anu, "dd/mm/yyyy")
  Else
    If far_r!far_fecha = LK_FECHA_DIA Then
      LBLEXTORNO.Caption = "DOCUMENTO EXTORNADO"
    End If
  End If
  LBLEXTORNO.Visible = True
Else
  LBLEXTORNO.Visible = False
End If

cabe_grid
pbValue = pbValue + 1
fila = 0
WS_BRUTO = 0
SUB_CANT = 0
subtotal = 0
PUB_DESCTO = 0
grid_fac2.Rows = 1
fila = 0
suma_subtotal = 0
LOC_ARROZ = ""
'**********************
'comienza llenado



  PS(0) = LK_CODCIA
  PS(1) = far_r!far_key_dircli
  PS(2) = far_r!far_codclie
  PS(3) = far_r!FAR_cp
  rs.Requery
  If rs.EOF Then
   'Exit Sub
  End If
  Do Until rs.EOF
   d_dire = rs!DIRCOMP
   rs.MoveNext
  Loop
 'termina llenar direccciro n
 WS_SUM_DESCT = 0
subtotal2 = 0
conteo_cantidad = 0
lblhora.Caption = Trim(far_r!far_hora)
If LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
    SQ_OPER = 1
    PUB_TIPREG = 48
    PUB_CODCIA = LK_CODCIA
    PUB_NUMTAB = Val(Nulo_Valor0(far_r!far_mortal))
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
     lblmotivo.Caption = "Motivo: " & Trim(tab_llave!TAB_NOMLARGO)
    Else
     lblmotivo.Caption = "S/Motivo "
    End If
End If
Do Until far_r.EOF
   If LOC_TIPMOV = 20 And Val(far_r!far_signo_arm) = 1 And far_r!far_estado = "E" Then GoTo NADA
   If (LOC_TIPMOV = 97 Or LOC_TIPMOV = 10) And Val(far_r!far_codart) = 0 Then
     grid_fac2.Rows = grid_fac2.Rows + 3
     fila = fila + 2
     SQ_OPER = 1
     PUB_FECHA = far_r!far_fecha
     pu_codcia = LK_CODCIA
     LEER_ALL_LLAVE
     Do Until all_llave.EOF
        If far_r!FAR_NUMOPER = all_llave!ALL_NUMOPER Then Exit Do
        all_llave.MoveNext
     Loop
     
     
'     If Not all_llave.EOF Then grid_fac2.TextMatrix(fila, 0) = all_llave!ALL_CONCEPTO
'     fila = fila + 1
     grid_fac2.TextMatrix(fila, 0) = far_r!far_concepto
     grid_fac2.ColWidth(0) = 6500
     GoTo pasa
   End If
   If LOC_TIPMOV = 98 And Val(far_r!far_codart) = 0 Then
     grid_fac2.Rows = grid_fac2.Rows + 2
     fila = fila + 2
     grid_fac2.TextMatrix(fila, 0) = far_r!far_concepto
     grid_fac2.ColWidth(0) = 6500
     GoTo pasa
   End If
   If LOC_TIPMOV = 99 And Val(far_r!far_codart) = 0 Then
     grid_fac2.Rows = grid_fac2.Rows + 1
     grid_fac2.TextMatrix(grid_fac2.Rows - 1, 0) = "Concepto: " & Trim(far_r!far_concepto)
     grid_fac2.ColWidth(0) = 6500
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_FECHA = far_r!far_fecha
     LEER_ALL_LLAVE
     Do Until all_llave.EOF
      If all_llave!ALL_NUMOPER = far_r!FAR_NUMOPER Then
        Exit Do
      End If
      all_llave.MoveNext
     Loop
     If all_llave.EOF Then
       MsgBox "Verificar...DATOS ", 48, Pub_Titulo
       Exit Sub
     End If
     If all_llave!ALL_SIGNO_CCM <> 0 Then
      SQ_OPER = 1
      PUB_CODBAN = all_llave!all_codban
      pu_codcia = LK_CODCIA
      LEER_CCM_LLAVE
      grid_fac2.Rows = grid_fac2.Rows + 1
      grid_fac2.TextMatrix(grid_fac2.Rows - 1, 0) = "Banco: " & Trim(ccm_llave!CCM_NOMBRE) & " Ch/. " & all_llave!ALL_CHESER & " " & all_llave!all_chenum
     End If
     ''If all_llave!ALL_SIGNO_CCM <> 0 Then
      grid_fac2.Rows = grid_fac2.Rows + 1
      grid_fac2.TextMatrix(grid_fac2.Rows - 1, 0) = "Relación Conable: "
      grid_fac2.Rows = grid_fac2.Rows + 1
      grid_fac2.TextMatrix(grid_fac2.Rows - 1, 0) = "'" & all_llave!ALL_CTAG1 & " = " & Format(all_llave!ALL_IMPG1, "#,#00.00")
      grid_fac2.Rows = grid_fac2.Rows + 1
      If all_llave!ALL_IMPG2 <> 0 Then
       grid_fac2.TextMatrix(grid_fac2.Rows - 1, 0) = "'" & all_llave!ALL_CTAG2 & " = " & Format(all_llave!ALL_IMPG2, "#,#00.00")
      End If
      grid_fac2.Rows = grid_fac2.Rows + 1
      grid_fac2.TextMatrix(grid_fac2.Rows - 1, 0) = "' Fec. Cancelación: " & all_llave!ALL_FECHA_CAN
      
      grid_fac2.Rows = grid_fac2.Rows + 1
      grid_fac2.TextMatrix(grid_fac2.Rows - 1, 0) = "' Fec.Contable: " & all_llave!ALL_FECHA_PRO

     ''End If
      
     GoTo pasa
   End If
   pbValue = pbValue + 1
   grid_fac2.Rows = grid_fac2.Rows + 1
   fila = fila + 1
   PUB_KEY = far_r!far_codart
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
      MsgBox "Error Grave en arti..."
   End If
   If far_r!FAR_tipmov = 20 Then
     If far_r!FAR_COSTEO <> "A" Or far_r!FAR_COSTEO_REAL <> "A" Then
       grid_fac2.Row = fila
       grid_fac2.COL = 1
       grid_fac2.CellBackColor = QBColor(7)
     End If
   End If
   grid_fac2.TextMatrix(fila, 0) = Trim(art_LLAVE!art_nombre)
   grid_fac2.TextMatrix(fila, 1) = Trim(art_LLAVE!ART_alterno)
   grid_fac2.TextMatrix(fila, 2) = far_r!FAR_cantidad / far_r!FAR_equiv
   grid_fac2.TextMatrix(fila, 3) = far_r!far_descri
   grid_fac2.TextMatrix(fila, 4) = far_r!FAR_PRECIO
   subtotal = Format(far_r!FAR_PRECIO * (far_r!FAR_cantidad / far_r!FAR_equiv), "0.00")
   subtotal2 = subtotal2 + subtotal
   conteo_cantidad = conteo_cantidad + Format((far_r!FAR_cantidad / far_r!FAR_equiv), "0.000")
   If LK_FLAG_GRIFO = "A" Then
       grid_fac2.TextMatrix(fila, 6) = Format(far_r!FAR_DESCTO, "0.00")
       DESC_GRIFO = DESC_GRIFO + far_r!FAR_DESCTO
   Else
    If LOC_TIPMOV = 20 Then
       grid_fac2.TextMatrix(fila, 6) = Trim(far_r!far_PORDESCTOS)
       WS_SUM_DESCT = WS_SUM_DESCT + Val(Trim(far_r!far_PORDESCTOS))
    Else
      If LOC_TIPMOV = 3 Then
        grid_fac2.TextMatrix(0, 6) = "Ref."
        grid_fac2.TextMatrix(fila, 6) = Trim(far_r!far_PORDESCTOS)
      Else
        grid_fac2.TextMatrix(fila, 6) = Trim(far_r!far_PORDESCTOS) & "%" ' Format(far_r!FAR_PORDESCTO1, "0.00") & "%"
      End If
    End If
   End If
   If LOC_TIPMOV = 102 Then
     grid_fac2.TextMatrix(fila, 6) = far_r!far_signo_arm
   End If
   grid_fac2.TextMatrix(fila, 7) = Format(far_r!FAR_FLETE, "0.00")
   grid_fac2.TextMatrix(fila, 8) = Format(far_r!far_PESO * far_r!FAR_cantidad, "0.00")
   grid_fac2.TextMatrix(fila, 10) = far_r!far_codart
   grid_fac2.TextMatrix(fila, 12) = far_r!far_numsec
   
   subtotal = redondea(subtotal)
   suma_subtotal = suma_subtotal + subtotal
   grid_fac2.TextMatrix(fila, 5) = subtotal
   SUB_CANT = SUB_CANT + (far_r!FAR_cantidad / far_r!FAR_equiv)
pasa:
   LOC_NUMPED = Val(far_r!FAR_NUMFAC_C)
   WS_BRUTO = far_r!far_bruto
   WS_DESCTO = far_r!FAR_TOT_DESCTO
   WS_IMPTO = far_r!far_impto
   WS_GASTOS = far_r!FAR_GASTOS
 '  If LK_EMP = "HER" And far_r!FAR_TOT_FLETE <> 0 Then
   WS_FLETE = Nulo_Valor0(far_r!FAR_TOT_FLETE)
   WFECHA_COMPRA = far_r!FAR_fecha_compra
   WFECHA_COMPRA_CAN = far_r!FAR_fecha_can
   lblhora.Caption = Trim(far_r!far_hora)
   wnumfac_c = Nulo_Valor0(far_r!FAR_NUMFAC_C)
   If LK_EMP <> "HER" And Not IsNull(far_r!FAR_fecha_pro) Then
   WFECHA_CONT = far_r!FAR_fecha_pro
   End If
   
   If far_r!FAR_EX_IGV = "E" Then LOC_ARROZ = "A"
 ' End If
NADA:
   far_r.MoveNext
Loop
 d_fecha_compra.Caption = Format(WFECHA_COMPRA, "dd/mm/yyyy")
 d_fecha_can.Caption = Format(WFECHA_COMPRA_CAN, "dd/mm/yyyy")
 FCONT.Caption = "Fec. Contable : " & Format(WFECHA_CONT, "dd/mm/yy")
 If Val(wnumfac_c) <> 0 And LOC_TIPMOV = 10 Then
  txtdocu.Caption = txtdocu.Caption + " /  Nro. Ped. = " & wnumfac_c
 End If
If Trim(LBLEXTORNO.Caption) = "" Then
   d_subtotal.Caption = WS_BRUTO '- WS_IMPTO + WS_DESCTO
   d_descto.Caption = WS_DESCTO
   WS_DESCTO = 0
   If LK_FLAG_GRIFO = "A" Then d_descto.Caption = Format(DESC_GRIFO, "0.00")
   d_flete.Caption = WS_FLETE
   d_impto.Caption = WS_IMPTO
   d_gastos.Caption = WS_GASTOS
   If WS_BRUTO = WS_DESCTO + WS_GASTOS Then
      d_neto.Caption = "0.00"
   Else
     If LK_EMP = "3AA" Then
      d_neto.Caption = Format(WS_BRUTO + WS_IMPTO, "0.00")
     Else
      If LOC_TIPMOV = 20 Or (LOC_TIPMOV = 97 And pu_cp = "P") Then
        d_neto.Caption = Format(WS_BRUTO + WS_IMPTO - WS_DESCTO + WS_GASTOS + Val(lblcargo.Caption) - Val(lblabono.Caption), "0.00")
      Else
        d_neto.Caption = Format(WS_BRUTO + WS_IMPTO - WS_DESCTO + WS_GASTOS, "0.00")
      End If
     End If
   End If
   If LK_EMP = "HER" And LOC_TIPMOV = 10 Then d_neto.Caption = Format(suma_subtotal, "0.00")
   If LOC_TIPMOV = 93 Then
       d_neto.Caption = Format(suma_subtotal, "0.00")
   End If
End If
   
lbldesc.Caption = ""
If LOC_TIPMOV = 20 Then
   
   lbldesc.Caption = WS_SUM_DESCT
   Dim relcompra_llave As rdoResultset
   pub_cadena = "select  * from relcompra where rel_codcia = '" & LK_CODCIA & "'  and rel_cp = 'P' " & _
   " and rel_codpro = " & Trim(d_Codclie.Caption) & " and REL_NUMSER   = " & Val(txtSerie.Text) & " and REL_NUMFAC =  " & Val(txtNumfac.Text) & " order by REL_LIQUIDO "
   Set relcompra_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   relcompra_llave.Requery
   lblfondo.Visible = False
   If relcompra_llave.EOF Then
      DOCANEXO.Caption = ""
      GoTo listo_anexos
   End If
   Do Until relcompra_llave.EOF
     If relcompra_llave!rel_LIQUIDO = 0 Then
        DOCANEXO.Caption = "Anexos Pendientes"
        lblfondo.Visible = True
        GoTo listo_anexos
        Exit Do
     End If
     relcompra_llave.MoveNext
   Loop
   DOCANEXO.Caption = "Hist. Anexos"
listo_anexos:
End If
   If LK_CODUSU = "ADMIN" And LOC_TIPMOV = 20 And subtotal2 <> Val(d_neto.Caption) Then
     MsgBox subtotal2
   End If
   CmdAnterior.Enabled = True
   cmdSiguiente.Enabled = True
   cmdImp.Enabled = True
   PB.Visible = False
   If cherela.Visible And LK_EMP <> "HER" Then cherela_Click
   LOC_NUMFAC_FIN = Val(txtNumfac.Text)
   If LK_CODUSU = "ADMIN" And (LOC_TIPMOV = 6 Or LOC_TIPMOV = 5) Then
'     MsgBox "CAntidad :  " & conteo_cantidad, 48, Pub_Titulo
   End If
   
Exit Sub
sale:
 MsgBox Err.Description, 48, Pub_Titulo
Resume Next
 LIMPIA_DOCU
End Sub
Public Sub cabe_grid()
If LOC_TIPMOV = 96 Then
   grid_fac2.Clear
   grid_fac2.Cols = 15
   grid_fac2.TextMatrix(0, 0) = "Cliente"
   grid_fac2.TextMatrix(0, 1) = "Codigo"
   grid_fac2.TextMatrix(0, 2) = "Vend"
   grid_fac2.TextMatrix(0, 3) = "F/LET"
   grid_fac2.TextMatrix(0, 4) = "F-B-G"
   grid_fac2.TextMatrix(0, 5) = "Serie"
   grid_fac2.TextMatrix(0, 6) = "N.Docum."
   grid_fac2.TextMatrix(0, 7) = "Guia   "
   grid_fac2.TextMatrix(0, 8) = "Saldo    "
   grid_fac2.TextMatrix(0, 9) = "Efectivo"
   grid_fac2.TextMatrix(0, 10) = "Nueva Fecha"
   grid_fac2.TextMatrix(0, 11) = "Nro . Cheque"
   grid_fac2.TextMatrix(0, 12) = "Importe"
   grid_fac2.TextMatrix(0, 13) = "Banco"
   grid_fac2.TextMatrix(0, 14) = "Fec.Cobrar"
   
   grid_fac2.ColWidth(0) = 2000
   grid_fac2.ColWidth(1) = 700
   grid_fac2.ColWidth(2) = 400
   grid_fac2.ColWidth(3) = 500
   grid_fac2.ColWidth(4) = 400
   grid_fac2.ColWidth(5) = 500
   grid_fac2.ColWidth(6) = 800
   grid_fac2.ColWidth(7) = 1
   grid_fac2.ColWidth(8) = 1100
   grid_fac2.ColWidth(9) = 1100
   grid_fac2.ColWidth(10) = 1200
   grid_fac2.ColWidth(11) = 1200
   grid_fac2.ColWidth(12) = 1200
   grid_fac2.ColWidth(13) = 1200
 
Else
   grid_fac2.Clear
   grid_fac2.Cols = 13
   grid_fac2.TextMatrix(0, 0) = "Descripción"
   grid_fac2.TextMatrix(0, 1) = "Codigo"
   grid_fac2.TextMatrix(0, 2) = "Cantidad"
   grid_fac2.TextMatrix(0, 3) = "Unidad"
   grid_fac2.TextMatrix(0, 4) = "Precio"
   grid_fac2.TextMatrix(0, 5) = "Subtotal"
   grid_fac2.TextMatrix(0, 6) = "Descto"
   grid_fac2.TextMatrix(0, 7) = "Flete"
   grid_fac2.TextMatrix(0, 8) = "Peso"
   
   grid_fac2.RowHeight(0) = 385
   grid_fac2.ColWidth(0) = 2500
   grid_fac2.ColWidth(1) = 900
   grid_fac2.ColWidth(2) = 800
   grid_fac2.ColWidth(3) = 900
   grid_fac2.ColWidth(4) = 800
   grid_fac2.ColWidth(5) = 900
   grid_fac2.ColWidth(6) = 700
   grid_fac2.ColWidth(7) = 800
   grid_fac2.ColWidth(8) = 800
   grid_fac2.ColWidth(9) = 0
   If LOC_TIPMOV = 20 Then
    grid_fac2.ColWidth(9) = 400
    grid_fac2.TextMatrix(0, 9) = "Flag" '  flag de descto
   End If
   grid_fac2.ColWidth(12) = 0
   
End If
End Sub



Private Sub grid_fac2_DblClick()

' Abrir Arti.frm al hacer doble clic en el nombre del articulo (columna 0) o codigo (columna 10)
If grid_fac2.Row > 0 And (grid_fac2.COL = 0 Or grid_fac2.COL = 10) Then
   If Val(grid_fac2.TextMatrix(grid_fac2.Row, 10)) > 0 Then
      frmARTI.Show vbModeless
      DoEvents
      frmARTI.Txt_key.Text = Trim(Str(Val(grid_fac2.TextMatrix(grid_fac2.Row, 10))))
      ' Usuario debe presionar Enter para buscar el articulo
      Exit Sub
   End If
End If

If LOC_TIPMOV <> 20 Then Exit Sub
If grid_fac2.COL = 1 Then
    LK_ACCESO_REPORT = ""
    Load frmclave2
    Screen.MousePointer = 0
    frmclave2.Show 1
    If LK_ACCESO_REPORT <> "A" Then
        Exit Sub
    End If

    pub_cadena = "SELECT FAR_COSTEO, FAR_COSTEO_REAL  FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? AND FAR_CODART = " & grid_fac2.TextMatrix(grid_fac2.Row, 10) & " AND FAR_NUMSEC = " & grid_fac2.TextMatrix(grid_fac2.Row, 12) & " ORDER BY FAR_NUMFAC"
    Set FARUSU = CN.CreateQuery("", pub_cadena)
    FARUSU(0) = 0
    FARUSU(1) = 0
    FARUSU(2) = 0
    FARUSU(3) = 0
    FARUSU(4) = 0
    Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
    FARUSU(0) = LK_CODCIA
    FARUSU(1) = LOC_TIPMOV
    FARUSU(2) = d_fecha.Caption
    FARUSU(3) = txtSerie.Text
    FARUSU(4) = txtNumfac.Text
    far_codusu.Requery
    Do Until far_codusu.EOF
        far_codusu.Edit
        If grid_fac2.CellBackColor = QBColor(7) Then
         far_codusu!FAR_COSTEO = "A"
         far_codusu!FAR_COSTEO_REAL = "A"
        Else
         far_codusu!FAR_COSTEO = ""
         far_codusu!FAR_COSTEO_REAL = ""
        End If
        far_codusu.Update
        far_codusu.MoveNext
    Loop
    
    MsgBox "OK. Cambiado", 48, Pub_Titulo
    txtnumfac_KeyPress 13
    
End If



If grid_fac2.COL = 6 Then
  If grid_fac2.CellBackColor = vbYellow Then
    grid_fac2.CellBackColor = vbBlack
 Else
    grid_fac2.CellBackColor = vbYellow
  End If
End If
If grid_fac2.COL = 7 Then
  If grid_fac2.CellBackColor = vbBlue Then
    grid_fac2.CellBackColor = vbBlack
 Else
    grid_fac2.CellBackColor = vbBlue
  End If
End If


End Sub


Private Sub lblabono_DblClick()
Dim wNuevoTotal As Currency
Dim wNuevoIGV  As Currency
Dim wNuevoBruto As Currency
If LOC_TIPMOV = 97 And pu_cp = "P" Then GoTo pasaok

If LOC_TIPMOV <> 20 Then Exit Sub
If Val(lblcargo.Caption) <> 0 Then
  MsgBox "Verificar, Casilla de Cargo tiene valor", 48, Pub_Titulo
  Exit Sub
End If
pasaok:
Dim wser As String
wser = InputBox("Ingrese Abono para Ajustar La Compra de Mercaderia : ", "FACTURA DE COMPRA -Modificar", lblabono.Caption)
If wser = "" Then Exit Sub

pub_cadena = "SELECT FAR_BRUTO , FAR_IMPTO, FAR_LIMCRE_ACT FROM FACART WHERE FAR_CODCIA = " & LK_CODCIA & " AND FAR_TIPMOV = " & LOC_TIPMOV & " AND FAR_FECHA = '" & Format(d_fecha.Caption, "dd/mm/yyyy") & "' AND FAR_NUMSER = " & Val(txtSerie.Text) & " AND FAR_NUMFAC = " & Val(txtNumfac.Text) & "  AND FAR_ESTADO <> 'E' AND FAR_ESTADO2 <>  'L'"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If X.EOF Then
  MsgBox "Registro No Existe.", 48, Pub_Titulo
  Exit Sub
End If
wNuevoTotal = (Val(X!far_bruto) + Val(X!far_impto) - Val(X!far_limcre_act))
wNuevoTotal = wNuevoTotal + Val(wser)
wNuevoBruto = Format(Val(wNuevoTotal) / (1 + (LK_IGV / 100)), "0.00")
wNuevoIGV = Val(wNuevoTotal) - Val(wNuevoBruto)
Do Until X.EOF
 X.Edit
 X!far_limcre_act = Val(wser)
 X!far_impto = Val(wNuevoIGV)
 X!far_bruto = Val(wNuevoBruto)
 X.Update
 X.MoveNext
Loop

MsgBox "Ok. Abono Actulizado.", 48, Pub_Titulo
txtnumfac_KeyPress 13


End Sub

Private Sub lblcargo_DblClick()
Dim wNuevoTotal As Currency
Dim wNuevoIGV  As Currency
Dim wNuevoBruto As Currency

If LOC_TIPMOV = 97 And pu_cp = "P" Then GoTo pasaok
If LOC_TIPMOV <> 20 Then Exit Sub
pasaok:
If Val(lblabono.Caption) <> 0 Then
  MsgBox "Verificar, Casilla de Abono tiene valor", 48, Pub_Titulo
  Exit Sub
End If
Dim wser As String
wser = InputBox("Ingrese Cargo para Ajustar La Compra de Mercaderia : ", "FACTURA DE COMPRA -Modificar", lblcargo.Caption)
If wser = "" Then Exit Sub

pub_cadena = "SELECT FAR_BRUTO , FAR_IMPTO, FAR_LIMCRE_ANT FROM FACART WHERE FAR_CODCIA = " & LK_CODCIA & " AND FAR_TIPMOV = " & LOC_TIPMOV & " AND FAR_FECHA = '" & Format(d_fecha.Caption, "dd/mm/yyyy") & "' AND FAR_NUMSER = " & Val(txtSerie.Text) & " AND FAR_NUMFAC = " & Val(txtNumfac.Text) & "  AND FAR_ESTADO <> 'E' AND FAR_ESTADO2 <>  'L'"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If X.EOF Then
  MsgBox "Registro No Existe.", 48, Pub_Titulo
  Exit Sub
End If
'wNuevoTotal = 5560
wNuevoTotal = (Val(X!far_bruto) + Val(X!far_impto) + Val(X!far_limcre_ant))
wNuevoTotal = wNuevoTotal - Val(wser)
wNuevoBruto = Format(Val(wNuevoTotal) / (1 + (LK_IGV / 100)), "0.00")
wNuevoIGV = Val(wNuevoTotal) - Val(wNuevoBruto)
Do Until X.EOF
 X.Edit
 X!far_limcre_ant = Val(wser)
 X!far_impto = Val(wNuevoIGV)
 X!far_bruto = Val(wNuevoBruto)
 X.Update
 X.MoveNext
Loop

MsgBox "Ok. Cargo Actulizado.", 48, Pub_Titulo
txtnumfac_KeyPress 13

End Sub

Private Sub lbldataant_Click()
Dim wser As String
Dim WNUM  As String
wser = InputBox("Ingrese Nro. de Serie : ", "FACTURA DE COMPRA -Modificar Anexo", Format(wser, "0"))
If wser = "" Then Exit Sub
wser = Format(wser, "0")
WNUM = InputBox("Ingrese Nro. de Documento : ", "FACTURA DE COMPRA -Modificar Anexo", Format(WNUM, "0"))
If WNUM = "" Then Exit Sub
txtSerie.Text = wser
txtNumfac.Text = WNUM

pub_cadena = "select  * from cartera where car_tipmov = 20 and car_codcia = '" & LK_CODCIA & "' and car_numser = " & wser & " and car_numfac = " & WNUM & " "
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not X.EOF Then
    temp_NUMSER_C = X!car_numser_c
    temp_NUMFAC_C = X!car_NUMFAC_C
    d_Codclie.Caption = X!CAR_codclie
    DOCANEXO_Click

End If
  

End Sub

Private Sub TIPMOV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If txtvend.Visible Then
   txtvend.SetFocus
  Else
   cmbFBG.SetFocus
 End If
End If
End Sub

Private Sub TIPMOV_LostFocus()
EVENTO_LOSFOCUS_TIPMOV
End Sub


Private Sub TRANS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then
PUB_TIPREG = -10
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "Mantenimiento de Transportistas"
FrmDatArti.Show 1
transporte.Requery
TRANS.Clear
Do Until transporte.EOF
    TRANS.AddItem Trim(transporte!TRN_NOMBRE) & String(80, " ") & transporte!TRN_KEY
    transporte.MoveNext
Loop
TRANS.SetFocus
DoEvents

End If
End Sub


Private Sub txtcserie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtnrouno, txtnrouno
End Sub

Private Sub txtdire_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  FRADIRE.Visible = False
  Exit Sub
End If
If KeyAscii = 13 Then
 Azul txtnum, txtnum
End If
End Sub

Private Sub txtdocu_DblClick()
Dim WNUM As String
Dim wser As String
Dim WNUMG As String
Dim wserg As String

Dim wfecha
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
   Exit Sub
End If

pub_cadena = "SELECT FAR_NUMSER_C, FAR_NUMFAC_C , FAR_SERGUIA, FAR_NUMGUIA FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? ORDER BY FAR_NUMFAC"
Set FARUSU = CN.CreateQuery("", pub_cadena)
FARUSU(0) = 0
FARUSU(1) = 0
FARUSU(2) = 0
FARUSU(3) = 0
FARUSU(4) = 0
Set far_codusu = FARUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
FARUSU(0) = LK_CODCIA
FARUSU(1) = LOC_TIPMOV
FARUSU(2) = d_fecha.Caption
FARUSU(3) = txtSerie.Text
FARUSU(4) = txtNumfac.Text
far_codusu.Requery
If far_codusu.EOF Then
  MsgBox "No Existe Numero Interno", 48, Pub_Titulo
  Exit Sub
End If
wser = far_codusu!FAR_NUMSER_C
wser = InputBox("Ingrese Nro. de Serie : ", "FACTURA DE COMPRA -Modificar el Documento", Format(wser, "0"))
If wser = "" Then Exit Sub
wser = Format(wser, "0")

WNUM = far_codusu!FAR_NUMFAC_C
WNUM = InputBox("Ingrese Nro. de Documento : ", "FACTURA DE COMPRA -Modificar el Documento", Format(WNUM, "0"))
If WNUM = "" Then Exit Sub

wserg = far_codusu!far_serguia
wserg = InputBox("Ingrese Nro. de Serie : ", "GUIA DE COMPRA -Modificar el Documento", Format(wserg, "0"))
If wserg = "" Then Exit Sub
wserg = Format(wserg, "0")

WNUMG = far_codusu!far_numguia
WNUMG = InputBox("Ingrese Nro. de Documento : ", "GUIA DE COMPRA -Modificar el Documento", Format(WNUMG, "0"))
If WNUMG = "" Then Exit Sub


Do Until far_codusu.EOF
 far_codusu.Edit
 far_codusu!FAR_NUMSER_C = wser
 far_codusu!FAR_NUMFAC_C = WNUM
 far_codusu!far_serguia = wserg
 far_codusu!far_numguia = WNUMG
 far_codusu.Update
 far_codusu.MoveNext
Loop



pub_cadena = "SELECT ALL_NUMSER_C,ALL_NUMFAC_C   FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_TIPMOV = ? AND ALL_FECHA_DIA = ? AND ALL_NUMSER = ? AND ALL_NUMFAC = ? ORDER BY ALL_NUMFAC"
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = 0
ALLUSU(1) = 0
ALLUSU(2) = 0
ALLUSU(3) = 0
ALLUSU(4) = 0
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
all_codusu.Requery
Do Until all_codusu.EOF
 all_codusu.Edit
 all_codusu!all_numser_c = wser
 all_codusu!all_numfac_c = WNUM
 all_codusu.Update
 all_codusu.MoveNext
 Loop


'FCONT.Caption = "Fec. Contable : " & Format(wfecha, "dd/mm/yyyy")

txtdocu.Caption = "F/ " & wser & " - " & WNUM
 

pub_cadena = "SELECT CAR_SERGUIA, CAR_NUMGUIA,CAR_NUMSER_C, CAR_NUMFAC_C FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_TIPMOV = ? AND CAR_FECHA_INGR = ? AND CAR_NUMSER = ? AND CAR_NUMFAC = ? "
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
Do Until all_codusu.EOF
  all_codusu.Edit
  all_codusu!car_numser_c = wser
  all_codusu!car_NUMFAC_C = WNUM
  all_codusu!car_serguia = wserg
  all_codusu!car_NUMGUIA = WNUMG
  all_codusu.Update
 all_codusu.MoveNext
Loop

pub_cadena = "SELECT CAA_NUMGUIA , CAA_SERGUIA,CAA_NUMSER_C, CAA_NUMFAC_C FROM CARACU WHERE CAA_CODCIA = ? AND CAA_TIPMOV = ? AND CAA_FECHA = ? AND CAA_NUMSER = ? AND CAA_NUMFAC = ?  "
Set ALLUSU = CN.CreateQuery("", pub_cadena)
ALLUSU(0) = LK_CODCIA
ALLUSU(1) = LOC_TIPMOV
ALLUSU(2) = d_fecha.Caption
ALLUSU(3) = txtSerie.Text
ALLUSU(4) = txtNumfac.Text
Set all_codusu = ALLUSU.OpenResultset(rdOpenKeyset, rdConcurValues)
Do Until all_codusu.EOF
 all_codusu.Edit
 all_codusu!CAA_NUMSER_C = wser
 all_codusu!CAA_NUMFAC_C = WNUM
 all_codusu!caa_serguia = wserg
 all_codusu!caa_numguia = WNUMG
 all_codusu.Update
 all_codusu.MoveNext
Loop

MsgBox "Ok Cambiado.", 48

End Sub

Private Sub txtnrodos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cmdcambiar.SetFocus
End If
End Sub

Private Sub txtnrouno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtnrodos, txtnrodos
End Sub

Private Sub txtnum_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 27 Then
  FRADIRE.Visible = False
  Exit Sub
End If
If KeyAscii = 13 Then
 TxtZonaTrabajo.SetFocus
 SendKeysSeguro VK_UP, True
End If


End Sub

Private Sub txtNumfac_GotFocus()
temporal = txtNumfac.Text
End Sub

Private Sub txtnumfac_KeyPress(KeyAscii As Integer)
'On Error GoTo SALE_X
Dim wven As Integer
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  If Val(txtSerie.Text) <= 0 Then
   'Exit Sub
  End If
  If Val(txtNumfac.Text) <= 0 Then
   LIMPIA_DOCU
   grid_fac2.Clear
   Exit Sub
  End If
  wflag_docu = ""
  If LOC_TIPMOV = 30 Then
'    txtserie.Locked = True
'    txtserie.Text = 0
    GoSub ORDENES
    Exit Sub
  End If
  If Left(cmbFBG.Text, 1) = "P" And LOC_TIPMOV <> 10 Then
    'txtSerie.Locked = True
    txtSerie.Text = 0
    GoSub PLANILLA
    Exit Sub
  Else
   
    loc_flag_espera = "A"
    LLENA_CONSULTA
    loc_flag_espera = ""
  End If
  If Trim(wflag_docu) = "" Then
    temporal = txtNumfac.Text
  Else
    txtSerie.Text = tempo_serie
    txtNumfac.Text = temporal
    Azul txtNumfac, txtNumfac
  End If
End If
Exit Sub
PLANILLA:
Dim PS_REP01 As rdoQuery
Dim llave_rep01  As rdoResultset
Dim ws_ingresos As Currency
Dim ws_salidas As Currency
Dim val_ingresos As Currency
Dim val_salidas As Currency
Dim acu_val_ingresos As Currency
Dim acu_val_salidas As Currency
Dim WS_CHEQUE As Currency
d_cheque.Visible = True
d_efectivo.Visible = True
pub_cadena = "SELECT CAA_NUMDOC,CAA_CP, CAA_SERDOC,CAA_SALDO_CAR, CAA_FECHA, CAA_TIPDOC, CAA_CODVEN,CAA_FECHA_VCTO, CAA_CONCEPTO , CAA_CODCLIE,CAA_FECHA_VCTO,CAA_IMPORTE, CAA_SALDO, CAA_NUMSER ,CAA_NUMFAC , CAA_FECHA, CAA_FBG FROM CARACU WHERE CAA_CODCIA = ? AND CAA_NUMPLAN = ? AND CAA_ESTADO <> 'E' ORDER BY CAA_CODCLIE, CAA_FECHA,CAA_NUM_OPER, CAA_SALDO_CAR"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = " "
PS_REP01(1) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

PS_REP01(0) = LK_CODCIA
PS_REP01(1) = Val(txtNumfac.Text)
llave_rep01.Requery
d_mensaje.Visible = False
If llave_rep01.EOF = True Then
  grid_fac2.Clear
  LIMPIA_DOCU
  d_mensaje.Visible = True
  GoTo CANCELA
End If
cabe_grid
d_fecha.Caption = Format(llave_rep01!CAA_FECHA, "dd/mm/yyyy")
wven = 0
f1 = 0
WS_BRUTO = 0
grid_fac2.Rows = 1
PB.Visible = True
pbMax = llave_rep01.RowCount
pbMin = 0
pbValue = 0
Do Until llave_rep01.EOF
   pbValue = pbValue + 1
   f1 = f1 + 1
   grid_fac2.Rows = grid_fac2.Rows + 1
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = llave_rep01!CAA_CODCLIE
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
      MsgBox "Error en Codigo de cliente, NO EXISTE ...", 48, Pub_Titulo
      GoTo CANCELA
   End If
    grid_fac2.TextMatrix(f1, 0) = Trim(cli_llave!CLI_NOMBRE)
    grid_fac2.TextMatrix(f1, 1) = Trim(cli_llave!cli_codclie)
    grid_fac2.TextMatrix(f1, 2) = Trim(llave_rep01!CAA_CODVEN)
   If Trim(llave_rep01!CAA_TIPDOC) = "FA" Or Trim(llave_rep01!CAA_TIPDOC) = "CC" Then
    grid_fac2.TextMatrix(f1, 3) = Trim(llave_rep01!CAA_TIPDOC)
    If llave_rep01!CAA_FBG = "F" Then
      grid_fac2.TextMatrix(f1, 4) = "FAC."
    ElseIf llave_rep01!CAA_FBG = "B" Then
      grid_fac2.TextMatrix(f1, 4) = "BOL."
    ElseIf llave_rep01!CAA_FBG = "G" Then
      grid_fac2.TextMatrix(f1, 4) = "GUIA"
    End If
    grid_fac2.TextMatrix(f1, 5) = llave_rep01!CAA_NUMSER
    grid_fac2.TextMatrix(f1, 6) = llave_rep01!CAA_NUMFAC
    grid_fac2.TextMatrix(f1, 7) = "" 'llave_rep01!CAA_NUMFAC
    grid_fac2.TextMatrix(f1, 8) = Format(llave_rep01!caa_SALDO_car + Val(llave_rep01!CAA_IMPORTE * -1), "0.00;(0.00)")
    grid_fac2.TextMatrix(f1, 9) = Format(Val(llave_rep01!CAA_IMPORTE * -1), "0.00;(0.00)")
    WS_BRUTO = WS_BRUTO + Val(llave_rep01!CAA_IMPORTE * -1)
    grid_fac2.TextMatrix(f1, 10) = llave_rep01!CAA_FECHA_VCTO
  Else
    grid_fac2.TextMatrix(f1, 3) = Trim(llave_rep01!CAA_TIPDOC)
    grid_fac2.TextMatrix(f1, 4) = Trim(llave_rep01!CAA_TIPDOC)
    SQ_OPER = 1
    pu_cp = llave_rep01!CAA_CP
    pu_codclie = cli_llave!cli_codclie
    pu_codcia = LK_CODCIA
    PUB_SERDOC = llave_rep01!CAA_SERDOC
    PUB_NUMDOC = llave_rep01!CAA_NUMDOC
    PUB_TIPDOC = Trim(llave_rep01!CAA_TIPDOC)
    LEER_CAR_LLAVE
    If Not car_llave.EOF Then grid_fac2.TextMatrix(f1, 11) = car_llave!car_num_cheque
    grid_fac2.TextMatrix(f1, 12) = llave_rep01!CAA_IMPORTE * -1
    grid_fac2.TextMatrix(f1, 13) = llave_rep01!caa_concepto
    grid_fac2.TextMatrix(f1, 14) = llave_rep01!CAA_FECHA_VCTO
    WS_CHEQUE = WS_CHEQUE + Val(llave_rep01!CAA_IMPORTE * -1)
  End If
  If llave_rep01!CAA_CODVEN <> 0 Then wven = llave_rep01!CAA_CODVEN
  llave_rep01.MoveNext
Loop
frmdocu.d_efectivo.Caption = Format(WS_BRUTO, "0.00;(0.00)")
frmdocu.d_cheque.Caption = Format(WS_CHEQUE, "0.00;(0.00)")
frmdocu.d_saldo.Caption = Format(WS_CHEQUE + WS_BRUTO, "0.00;(0.00)")
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODVEN = wven
LEER_VEN_LLAVE
d_domicilio.Caption = ""
If Not ven_llave.EOF Then d_domicilio.Caption = Format(wven, "00") + "  " + Trim(ven_llave!VEM_NOMBRE)

lblcheque.Visible = True
lblEfectivo.Visible = True
CmdAnterior.Enabled = True
cmdSiguiente.Enabled = True
PB.Visible = False
TRANS.Visible = False

CANCELA:
Return

ORDENES:
'Dim PS_REP01 As rdoQuery
'Dim llave_rep01  As rdoResultset
'Dim ws_ingresos As Currency
'Dim ws_salidas As Currency
'Dim val_ingresos As Currency
'Dim val_salidas As Currency
'Dim acu_val_ingresos As Currency
'Dim acu_val_salidas As Currency
'Dim WS_CHEQUE As Currency
d_cheque.Visible = True
d_efectivo.Visible = True
pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_NUMSER = ? AND PED_NUMFAC = ? ORDER BY PED_FECHA,PED_NUMSER, PED_NUMFAC,PED_NUMSEC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = " "
PS_REP01(1) = 0
PS_REP01(2) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

PS_REP01(0) = LK_CODCIA
PS_REP01(1) = txtSerie.Text
PS_REP01(2) = Val(txtNumfac.Text)

llave_rep01.Requery
d_mensaje.Visible = False
If llave_rep01.EOF = True Then
  grid_fac2.Clear
  LIMPIA_DOCU
  d_mensaje.Visible = True
  GoTo CANCELA
End If
cabe_grid
d_fecha.Caption = Format(llave_rep01!PED_FECHA, "dd/mm/yyyy")
wven = 0
f1 = 0
WS_BRUTO = 0
grid_fac2.Rows = 1
PB.Visible = True
pbMax = llave_rep01.RowCount
pbMin = 0
pbValue = 0
pu_codcia = LK_CODCIA
SQ_OPER = 1
pu_cp = "P"
pu_codclie = llave_rep01!PED_CODCLIE
LEER_CLI_LLAVE
If cli_llave.EOF Then
  MsgBox "Rgistro de PROVEEDOR not :" & llave_rep01!PED_CODCLIE, 48, Pub_Titulo
  GoTo CANCELA
End If
pbValue = pbValue + 1
d_condicion.Caption = "Orden de Compra"
d_Codclie.Caption = Trim(cli_llave!cli_codclie)
d_nomclie.Caption = Trim(cli_llave!CLI_NOMBRE)
d_domicilio.Caption = Trim(cli_llave!CLI_CASA_DIREC)
d_ruc.Caption = Trim(cli_llave!cli_ruc_esposo)
d_dire.Caption = "      AGENCIA: " & Trim(Nulo_Valors(par_llave!PAR_AGE_EMP))
If Nulo_Valors(cli_llave!CLI_MONEDA) = "D" Then
d_moneda.Caption = "$."
Else
d_moneda.Caption = "S/."
End If
d_usuario.Caption = llave_rep01!PED_CODUSU
d_moneda.Visible = True
If Nulo_Valor0(llave_rep01!PED_MONEDA) = "D" Then
  d_moneda.Caption = "US$"
Else
  d_moneda.Caption = "S/."
End If
cabe_grid
fila = 0
WS_BRUTO = 0
SUB_CANT = 0
subtotal = 0
PUB_DESCTO = 0
grid_fac2.Rows = 1
fila = 0
Do Until llave_rep01.EOF
'   pbValue = pbValue + 1
   grid_fac2.Rows = grid_fac2.Rows + 1
   fila = fila + 1
   PUB_KEY = llave_rep01!PED_codart
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
      MsgBox "Error Grave en arti..."
   End If
   grid_fac2.TextMatrix(fila, 0) = Trim(art_LLAVE!art_nombre)
   grid_fac2.TextMatrix(fila, 1) = Trim(art_LLAVE!ART_alterno)
   grid_fac2.TextMatrix(fila, 2) = llave_rep01!PED_CANTIDAD / llave_rep01!PED_EQUIV
   grid_fac2.TextMatrix(fila, 3) = llave_rep01!PED_UNIDAD
   grid_fac2.TextMatrix(fila, 4) = Format(llave_rep01!PED_PRECIO, "0.00")
   subtotal = Format(llave_rep01!PED_PRECIO * (llave_rep01!PED_CANTIDAD / llave_rep01!PED_EQUIV), "0.00")
   subtotal = redondea(subtotal)
   grid_fac2.TextMatrix(fila, 5) = subtotal
   SUB_CANT = SUB_CANT + (llave_rep01!PED_CANTIDAD / llave_rep01!PED_EQUIV)
pasa:
   WS_BRUTO = llave_rep01!PED_BRUTO
   WS_DESCTO = 0
   WS_IMPTO = llave_rep01!PED_IGV
   WS_GASTOS = 0
   WS_FLETE = 0
   llave_rep01.MoveNext
Loop
   d_subtotal.Caption = WS_BRUTO '- WS_IMPTO + WS_DESCTO
   d_descto.Caption = WS_DESCTO
   d_flete.Caption = WS_FLETE
   d_impto.Caption = WS_IMPTO
   d_gastos.Caption = WS_GASTOS
   d_neto.Caption = Format(WS_BRUTO + WS_IMPTO - WS_DESCTO + WS_GASTOS, "0.00")
   CmdAnterior.Enabled = True
   cmdSiguiente.Enabled = True
   PB.Visible = False
Exit Sub

Return

SALE_X:
MsgBox Err.Description
End Sub

Public Sub LIMPIA_DOCU()
lblhora.Caption = ""
lblmotivo.Caption = ""
lblmotivo.Visible = False
lblfondo.Visible = False
d_tipvta.Caption = ""
LBLEXTORNO.Caption = ""
d_Codclie.Caption = ""
d_nomclie.Caption = ""
d_codven.Caption = ""
d_nomven.Caption = ""
d_dias.Caption = ""
d_subtotal.Caption = "0.00"
d_gastos.Caption = "0.00"
d_descto.Caption = "0.00"
d_impto.Caption = "0.00"
d_neto.Caption = "0.00"
d_flete.Caption = "0.00"
d_fechaV.Caption = ""
d_newvcto.Caption = ""
d_saldo.Caption = ""
grid_fac2.Clear
d_dire.Caption = ""
d_ruc.Caption = ""
d_condicion.Caption = ""
d_fecha.Caption = ""
LBLEXTORNO.Visible = False
d_moneda.Caption = ""
d_mensaje.Visible = False
txtdocu.Caption = ""
txtdocu.Visible = False
lblfac.Visible = False
lblcheque.Visible = False
d_mensaje.Visible = False
d_efectivo.Visible = False
d_cheque.Visible = False
d_efectivo.Caption = ""
d_cheque.Caption = ""
lblEfectivo.Visible = False
d_domicilio.Caption = ""
d_usuario.Caption = ""
d_fecha_compra.Caption = ""
FECHA_PART.Text = ""

WGUIA_RELA = ""
LOC_ARROZ = ""
' ICA
If LK_EMP <> "HER" Then
   tguia.Text = ""
End If
Doc_EL = ""
End Sub

Public Sub SELE_DOCU()
'Set PSFAR = CN.CreateQuery("", pub_cadena)
'Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

End Sub

Public Sub LEER_FAR_CONSUL()
PSFAR_CONSUL.rdoParameters(0) = PU_TIPMOV
PSFAR_CONSUL.rdoParameters(1) = pu_codcia
PSFAR_CONSUL.rdoParameters(2) = PU_FBG
PSFAR_CONSUL.rdoParameters(3) = PU_NUMSER
PSFAR_CONSUL.rdoParameters(4) = pu_cp
pub_cadena = "SELECT TOP 1 FAR_numfac FROM facart WHERE FAR_TIPMOV = " & PU_TIPMOV & " AND FAR_CODCIA = '" & pu_codcia & "' AND FAR_FBG = '" & PU_FBG & "' AND FAR_NUMSER = " & PU_NUMSER & " AND FAR_ESTADO <> 'E' AND FAR_CP = '" & pu_cp & "'  ORDER BY  FAR_NUMFAC DESC"
Set far_consul = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
'far_consul.Requery


End Sub
Public Sub LEER_CAR_CONSUL()
PSCAR_CONSUL.rdoParameters(0) = pu_codcia
PSCAR_CONSUL.rdoParameters(1) = pu_cp
PSCAR_CONSUL.rdoParameters(2) = pu_codclie
PSCAR_CONSUL.rdoParameters(3) = PU_FBG
PSCAR_CONSUL.rdoParameters(4) = PU_NUMSER
PSCAR_CONSUL.rdoParameters(5) = PU_NUMFAC
car_consul.Requery
End Sub

Private Sub txtserie_GotFocus()
txtSerie.Text = Trim(txtSerie.Text)
End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 If txtNumfac.Enabled Then
  txtNumfac.SetFocus
 End If
End If
End Sub
Public Sub LLENADOS(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        If tab_mayor!tab_codart = 1 Then cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        tab_mayor.MoveNext
    Loop
End Sub

Public Function REP_CONSUL() As Integer
Dim WMONEDA As String * 1
Dim Xx As String * 1
Dim wser As String * 3
Dim WSRUTA As String
Dim indice As Integer
Dim wM As Integer
Dim llave_rep01 As rdoResultset
Dim PS_REP01 As rdoQuery
Dim i As Integer
Dim VALOR
Dim loc_xl As Object
Dim loc_codtra As Integer
Dim wRuta As String
Dim WSNUMDOC As String
Dim numero_device As Integer
If LOC_TIPMOV = 10 Then
    SQ_OPER = 2
    PUB_CODCIA = LK_CODCIA
    PUB_CODVEN = Val(d_codven.Caption)
    LEER_PAR_LLAVE
    If pac_llave.EOF Then
       MsgBox "No se ha definido archivos de Impresión", 48, Pub_Titulo
       Exit Function
    End If
End If

'If LK_EMP = "HER" Then
'  wRuta = "C:\ADMIN\STANDAR\"
'Else
If LK_EMP_PTO = "A" Then
  'wRuta = PUB_RUTA_OTRO & "PTOVTA\"
  wRuta = LK_RUTA_DOC & "PTOVTA\"
Else
  'wRuta = PUB_RUTA_OTRO
  wRuta = LK_RUTA_DOC
End If
If Trim(d_moneda.Caption) = "S/." Then
 WMONEDA = "S"
Else
 WMONEDA = "D"
End If

'End If
  

    frmdocu.Reportes.Connect = PUB_ODBC
    If frmdocu.imp.Value = 1 Then
      frmdocu.Reportes.Destination = crptToPrinter
    Else
      frmdocu.Reportes.Destination = crptToWindow  '= crptToPrinter
    End If
    frmdocu.Reportes.WindowLeft = 2
    frmdocu.Reportes.WindowTop = 70
    frmdocu.Reportes.WindowWidth = 635
    frmdocu.Reportes.WindowHeight = 390
    frmdocu.Reportes.Formulas(1) = ""
    frmdocu.d_neto.Refresh
    PUB_NETO = Val(frmdocu.d_neto.Caption)
    PUB_FECHA = frmdocu.d_fecha.Caption
    PU_NUMSER = Val((frmdocu.txtSerie.Text))
    PU_NUMFAC = Val((frmdocu.txtNumfac.Text))
    If LK_EMP = "PIU" Then
       frmdocu.Reportes.Formulas(1) = "SON=  ' " & CONVER_LETRAS(PUB_NETO, WMONEDA) & "'"
    Else
       frmdocu.Reportes.Formulas(1) = "SON=  'SON: " & CONVER_LETRAS(PUB_NETO, WMONEDA) & "'"
    End If
    If PUB_NETO <> Val(frmdocu.d_neto.Caption) Then
      MsgBox "Espere....!!!", 48, Pub_Titulo
      Exit Function
    End If
    LOC_NUMFAC_FIN = PU_NUMFAC
    Reportes.Formulas(8) = ""
    Reportes.Formulas(9) = ""
    Reportes.Formulas(10) = ""
    Reportes.Formulas(11) = ""
    Reportes.Formulas(12) = ""
    Reportes.Formulas(13) = ""
    Reportes.Formulas(14) = ""
    Reportes.Formulas(15) = ""
    Reportes.Formulas(16) = ""
    Reportes.Formulas(17) = ""
    Reportes.Formulas(18) = ""
    Reportes.Formulas(19) = ""
    Reportes.Formulas(20) = ""
    Reportes.Formulas(21) = ""
    Reportes.Formulas(22) = ""
    Reportes.Formulas(23) = ""
    Reportes.Formulas(24) = ""
    Reportes.Formulas(25) = ""
    Reportes.Formulas(26) = ""
    Reportes.Formulas(27) = ""
    Reportes.Formulas(28) = ""

    
    
    If LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Then
        frmdocu.Reportes.WindowTitle = "GUIA DE REMISION  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        If LOC_TIPMOV = 3 Or LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Then
          frmdocu.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_GUIA)  ' "FACGUIA.RPT"
        Else
          frmdocu.Reportes.ReportFileName = wRuta + "GUIAR3.RPT"
        End If
        pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = " & LOC_TIPMOV & " AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND  ( {FACART.FAR_NUMFAC} >= " & PU_NUMFAC & " and {FACART.FAR_NUMFAC} <= " & LOC_NUMFAC_FIN & ") AND {FACART.FAR_NUMSER} = '" & Trim(txtSerie.Text) & "'  "
        'Debug.Print pub_cadena
        frmdocu.Reportes.Formulas(1) = ""
        If chetrans.Value = 1 Then
             PS_TRA(0) = Val(Right(TRANS.Text, 3))
             llave_trans.Requery
             Reportes.Formulas(12) = "TRN_NOMBRE    =  '" & llave_trans!TRN_NOMBRE & "'"
             Reportes.Formulas(13) = "TRN_DIRECCION =  '" & llave_trans!TRN_DIRECCION & "'"
             Reportes.Formulas(14) = "TRN_RUC       =  '" & llave_trans!TRN_RUC & "'"
             Reportes.Formulas(15) = "TRN_DNI       =  '" & llave_trans!TRN_DNI & "'"
             Reportes.Formulas(16) = "TRN_PLACA     =  '" & llave_trans!TRN_PLACA & "'"
             Reportes.Formulas(16) = "TRN_CHOFER    =  '" & llave_trans!TRN_CHOFER & "'"
        End If
        If LOC_TIPMOV = 100 Or LOC_TIPMOV = 3 Then GoTo PASA_OP
        GoTo pasa_todo
    End If
    If LOC_TIPMOV = 20 Then
        frmdocu.Reportes.WindowTitle = "KARDEX Nro  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        pub_mensaje = "Inventario Valorado (Si), Inventario en Unidades (No) "
        Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
        If Pub_Respuesta = vbYes Then
            frmdocu.Reportes.ReportFileName = wRuta + "NOTAING.RPT"
        Else
            frmdocu.Reportes.ReportFileName = wRuta + "NOTAINV.RPT"
        End If
        wser = PU_NUMSER
        pub_cadena = "{FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_NUMSER}= '" & wser & "'  AND {FACART.FAR_NUMFAC} = " & PU_NUMFAC
        frmdocu.Reportes.Formulas(1) = ""
        Reportes.Formulas(1) = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
        GoTo pasa_todo
    End If
    If LOC_TIPMOV = 75 Or LOC_TIPMOV = 93 Or LOC_TIPMOV = 102 Then
        wser = PU_NUMSER
        frmdocu.Reportes.WindowTitle = "KARDEX Nro  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        If LOC_TIPMOV = 93 Then
        frmdocu.Reportes.ReportFileName = wRuta + "CAMBP.RPT"
        End If
        If LOC_TIPMOV = 75 Then
          frmdocu.Reportes.ReportFileName = wRuta + "DEF01.RPT"
        End If
        If LOC_TIPMOV = 102 Then
          frmdocu.Reportes.ReportFileName = wRuta + "CAMBPRO.RPT"
          Reportes.Formulas(12) = "DOCUMENTO=  '" & DOC102 & "'"
        End If
        pub_cadena = "{FACART.FAR_TIPMOV} = " & LOC_TIPMOV & " AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_NUMSER}= '" & wser & "'  AND {FACART.FAR_NUMFAC} = " & PU_NUMFAC
        frmdocu.Reportes.Formulas(1) = ""
        GoTo pasa_todo
    End If
    If LOC_TIPMOV = 99 Then
        wser = PU_NUMSER
        frmdocu.Reportes.WindowTitle = "KARDEX Nro  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        frmdocu.Reportes.ReportFileName = wRuta + "VOCCM.RPT"
        pub_cadena = "{ALLOG.ALL_CODCIA} = '" & LK_CODCIA & "' AND {ALLOG.ALL_NUMSER}= '" & wser & "'  AND {ALLOG.ALL_NUMFAC} = " & PU_NUMFAC
        frmdocu.Reportes.Formulas(1) = ""
        GoTo pasa_todo
    End If
    If LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
      If Left(frmdocu.cmbFBG.Text, 1) = "B" Then
        frmdocu.Reportes.WindowTitle = "BOLETA  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & " al " & Format(LOC_NUMFAC_FIN, "0000000")
        If LK_EMP = "PIU" Or LK_EMP = "PAR" Then
          frmdocu.Reportes.WindowTitle = "BOLETA  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        End If
        If Doc_EL = "A" Then
            frmdocu.Reportes.ReportFileName = wRuta + "ELCLIBOL.RPT"
        Else
            frmdocu.Reportes.ReportFileName = wRuta + "CLIBOL.RPT"
        End If
        If SIN_CODART = 1 Then frmdocu.Reportes.ReportFileName = wRuta + "CLIBOL99.RPT"
        If LOC_ARROZ = "A" Then frmdocu.Reportes.ReportFileName = wRuta + "CLIBOLIE.RPT"
      ElseIf Left(frmdocu.cmbFBG.Text, 1) = "F" Then
        frmdocu.Reportes.WindowTitle = "FACTURA : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & " al " & Format(LOC_NUMFAC_FIN, "0000000")
        If LK_EMP = "PIU" Or LK_EMP = "PAR" Then
          frmdocu.Reportes.WindowTitle = "FACTURA : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        End If
        If Doc_EL = "A" Then
            frmdocu.Reportes.ReportFileName = wRuta + "ELCLIFAC.RPT"
        Else
            frmdocu.Reportes.ReportFileName = wRuta + "CLIFAC.RPT"
        End If
        
        If LOC_ARROZ = "A" Then frmdocu.Reportes.ReportFileName = wRuta + "CLIFACIE.RPT"
      ElseIf Left(frmdocu.cmbFBG.Text, 1) = "G" Or Left(frmdocu.cmbFBG.Text, 1) = "P" Then
        frmdocu.Reportes.WindowTitle = " GUIA   : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & " al " & Format(LOC_NUMFAC_FIN, "0000000")
        frmdocu.Reportes.ReportFileName = wRuta + "CLIGUIA.RPT"
      ElseIf Left(frmdocu.cmbFBG.Text, 1) = "N" Then
        frmdocu.Reportes.WindowTitle = " N. CREDITO  : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & " al " & Format(LOC_NUMFAC_FIN, "0000000")
        If Trim(grid_fac2.TextMatrix(1, 1)) = "" Then
         frmdocu.Reportes.ReportFileName = wRuta + "NCREDV.RPT"
        Else
         frmdocu.Reportes.ReportFileName = wRuta + "NCRED.RPT"
        End If
      ElseIf Left(frmdocu.cmbFBG.Text, 1) = "C" Then
        frmdocu.Reportes.WindowTitle = " N. CREDITO  : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & " al " & Format(LOC_NUMFAC_FIN, "0000000")
        If Trim(grid_fac2.TextMatrix(1, 1)) = "" Then
         frmdocu.Reportes.ReportFileName = wRuta + "NCREDVP.RPT"
        Else
         frmdocu.Reportes.ReportFileName = wRuta + "NCREDP.RPT"
        End If
      ElseIf Left(frmdocu.cmbFBG.Text, 1) = "D" Then
        frmdocu.Reportes.WindowTitle = " N. DEBITO  : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
        If Trim(grid_fac2.TextMatrix(1, 1)) = "" Then
         frmdocu.Reportes.ReportFileName = wRuta + "NDEBV.RPT"
        Else
         frmdocu.Reportes.ReportFileName = wRuta + "NDEB.RPT"
        End If
      End If
    End If
    wser = PU_NUMSER
    If Left(frmdocu.cmbFBG.Text, 1) = "N" Then
       pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 97 AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FBG} = '" & Left(frmdocu.cmbFBG.Text, 1) & "' AND {FACART.FAR_NUMSER}= '" & wser & "' AND ( {FACART.FAR_NUMFAC} >= " & PU_NUMFAC & " and {FACART.FAR_NUMFAC} <= " & LOC_NUMFAC_FIN & ")"
    ElseIf Left(frmdocu.cmbFBG.Text, 1) = "C" Then
       pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 97 AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FBG} = '" & Left(frmdocu.cmbFBG.Text, 1) & "' AND {FACART.FAR_NUMSER}= '" & wser & "' AND ( {FACART.FAR_NUMFAC} >= " & PU_NUMFAC & " and {FACART.FAR_NUMFAC} <= " & LOC_NUMFAC_FIN & ")"
    ElseIf Left(frmdocu.cmbFBG.Text, 1) = "D" Then
       pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 98 AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FBG} = '" & Left(frmdocu.cmbFBG.Text, 1) & "' AND {FACART.FAR_NUMSER}= '" & wser & "' AND ( {FACART.FAR_NUMFAC} >= " & PU_NUMFAC & " and {FACART.FAR_NUMFAC} <= " & LOC_NUMFAC_FIN & ")"
    Else
       pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FBG} = '" & Left(frmdocu.cmbFBG.Text, 1) & "' AND {FACART.FAR_NUMSER}= '" & wser & "' AND ( {FACART.FAR_NUMFAC} >= " & PU_NUMFAC & " and {FACART.FAR_NUMFAC} <= " & LOC_NUMFAC_FIN & ")"
     End If
    Reportes.Formulas(12) = ""
    Reportes.Formulas(13) = ""
    Reportes.Formulas(14) = ""
    Reportes.Formulas(15) = ""
    Reportes.Formulas(16) = ""
     ' ICA
    If cherela.Value = 1 And LK_EMP = "HER" Then
       Reportes.Formulas(10) = "GUIA=  '" & Trim(tguia.Text) & "'"
    End If
    'Debug.Print pub_cadena
    If LOC_TIPMOV = 3 Or LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Or LOC_TIPMOV = 35 Then
     pub_cadena = "{FACART.FAR_TIPMOV} = " & LOC_TIPMOV & " AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_NUMSER}= '" & wser & "'  AND {FACART.FAR_NUMFAC} = " & PU_NUMFAC
     frmdocu.Reportes.Formulas(1) = ""
     frmdocu.Reportes.SelectionFormula = pub_cadena
     If LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Then
      If chelotes.Value = 1 Then
        frmdocu.Reportes.ReportFileName = wRuta + "TRANSFLT.RPT"
      Else
        frmdocu.Reportes.ReportFileName = wRuta + "TRANSF.RPT"
      End If
     Else
      frmdocu.Reportes.ReportFileName = wRuta + "devl.RPT"
     End If
     If LOC_TIPMOV = 35 Then
       frmdocu.Reportes.ReportFileName = wRuta + "consig.RPT"
     End If
     GoTo LISTOACTION
    End If
    'If Left(frmdocu.cmbFBG.Text, 1) = "G" Then GoTo PASA_OP
    If LK_EMP = "PIU" Or LK_EMP = "PAR" Or LK_EMP = "HER" Then
       Reportes.Formulas(17) = ""
       Reportes.Formulas(18) = ""
       Reportes.Formulas(19) = ""
       Reportes.Formulas(20) = ""
       pub_mensaje = "Desea Imprimir la " & Trim(frmdocu.Reportes.WindowTitle) & "   ¿Desea Continuar... ?"
       Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
       If Pub_Respuesta = vbYes Then
         frmdocu.Reportes.SelectionFormula = pub_cadena
         Reportes.Formulas(10) = ""
         Reportes.Formulas(11) = ""
         Reportes.Formulas(12) = ""
         Reportes.Formulas(13) = ""
         Reportes.Formulas(14) = ""
         Reportes.Formulas(15) = ""
         Reportes.Formulas(16) = ""
           Reportes.Formulas(21) = ""
           Reportes.Formulas(22) = ""
           Reportes.Formulas(23) = ""
           Reportes.Formulas(24) = ""
           Reportes.Formulas(25) = ""
           Reportes.Formulas(26) = ""
           Reportes.Formulas(27) = ""
         If cherela.Value = 1 Then
            Reportes.Formulas(10) = "GUIA=  '" & Trim(tguia.Text) & "'"
         End If
LISTOACTION:
         On Error GoTo accion
         Reportes.Formulas(11) = ""
         frmdocu.Reportes.WindowTitle = frmdocu.Reportes.WindowTitle & " Archivo: " & Trim(frmdocu.Reportes.ReportFileName)
         If chefg.Value = 1 Then
             Reportes.Formulas(11) = "FECHA_PARTIDA=  '" & FECHA_PART.Text & "'"
            frmdocu.Reportes.ReportFileName = wRuta + "FACGUIAAL.RPT" 'Trim(pac_llave!PAC_ARCHI_GUIA) ' "FACGUIA.RPT"
         End If
         frmdocu.Reportes.Action = 1

         
         Exit Function
         On Error GoTo 0
       End If
PASA_OP:
       If LOC_TIPMOV = 10 Or LOC_TIPMOV = 100 Or LOC_TIPMOV = 3 Then
         WSNUMDOC = Left(frmdocu.Reportes.WindowTitle, 23)
         If LK_EMP = "PAR" Then
           WSNUMDOC = Right(Trim(frmdocu.Reportes.WindowTitle), 15)
         End If
         If LK_EMP = "HER" Then
              WSNUMDOC = Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000") & ""
         End If
         frmdocu.Reportes.WindowTitle = "GUIA DE VENTA  " & Trim(frmdocu.Reportes.WindowTitle)
         If sin_valor.Value = 1 Then
            frmdocu.Reportes.ReportFileName = wRuta + "FACGUIA2.RPT"
         Else
            frmdocu.Reportes.ReportFileName = wRuta + "FACGUIA.RPT" 'Trim(pac_llave!PAC_ARCHI_GUIA) ' "FACGUIA.RPT"
         End If
         
         If LOC_TIPMOV = 100 Then
            If sin_valor.Value = 1 Then
            frmdocu.Reportes.ReportFileName = wRuta + "GUIAR2.RPT"
            Else
            frmdocu.Reportes.ReportFileName = wRuta + "GUIAR.RPT"
         End If
         End If
         
         pub_mensaje = "Desea Imprimir la " & Trim(frmdocu.Reportes.WindowTitle) & "   ro  Continuar... ?"
         Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
         If Pub_Respuesta = vbYes Then
           frmdocu.Reportes.SelectionFormula = pub_cadena
           Reportes.Formulas(21) = "FECHA_EMI=  '" & d_fecha_compra.Caption & "'"
           Reportes.Formulas(10) = ""
           Reportes.Formulas(11) = "FECHA_PARTIDA=  '" & FECHA_PART.Text & "'"
           Reportes.Formulas(12) = ""
           Reportes.Formulas(13) = ""
           Reportes.Formulas(14) = ""
           Reportes.Formulas(15) = ""
           Reportes.Formulas(16) = ""
           If cherela.Value = 1 And LOC_TIPMOV = 10 Then
             Reportes.Formulas(10) = "NUMDOC=  '" & WSNUMDOC & "'"
           End If
           Reportes.Formulas(22) = ""
           Reportes.Formulas(23) = ""
           Reportes.Formulas(24) = ""
           Reportes.Formulas(25) = ""
           Reportes.Formulas(26) = ""
           Reportes.Formulas(27) = ""
           Reportes.Formulas(28) = ""
   Select Case NUMERO.Text
   Case "1"
         Reportes.Formulas(22) = "VENT=  'X'"
   Case "2"
         Reportes.Formulas(23) = "VENT2=  'X'"
   Case "3"
         Reportes.Formulas(24) = "VENT3=  'X'"
   Case "4"
         Reportes.Formulas(25) = "VENT4=  'X'"
   Case "5"
         Reportes.Formulas(26) = "VENT5=  'X'"
   Case "6"
         Reportes.Formulas(27) = "VENT6=  'X'"
   Case "7"
         Reportes.Formulas(28) = "VENT7=  'X'"
   End Select

           
           If chetrans.Value = 1 Then
             PS_TRA(0) = Val(Right(TRANS.Text, 3))
             llave_trans.Requery
             Reportes.Formulas(12) = "TRN_NOMBRE    =  '" & llave_trans!TRN_NOMBRE & "'"
             Reportes.Formulas(13) = "TRN_DIRECCION =  '" & llave_trans!TRN_DIRECCION & "'"
             Reportes.Formulas(14) = "TRN_RUC       =  '" & llave_trans!TRN_RUC & "'"
             Reportes.Formulas(15) = "TRN_DNI       =  '" & llave_trans!TRN_DNI & "'"
             Reportes.Formulas(16) = "TRN_PLACA     =  '" & llave_trans!TRN_PLACA & "'"
             If LK_EMP = "PIU" Or LK_EMP = "HER" Then
              Reportes.Formulas(17) = "TRN_CHOFER     =  '" & llave_trans!TRN_CHOFER & "'"
              Reportes.Formulas(18) = "TRN_DIR_CHOFER =  '" & llave_trans!TRN_DIR_CHOFER & "'"
              Reportes.Formulas(19) = "TRN_BREVETE    =  '" & llave_trans!TRN_BREVETE & "'"
              Reportes.Formulas(20) = "TRN_DNI_CHOFER =  '" & llave_trans!TRN_DNI_CHOFER & "'"
             End If
           End If
           frmdocu.Reportes.WindowTitle = frmdocu.Reportes.WindowTitle & " Archivo: " & Trim(frmdocu.Reportes.ReportFileName)
           If cherela.Value = 1 And LOC_TIPMOV = 10 And Left(frmdocu.cmbFBG.Text, 1) = "G" Then
              frmdocu.Reportes.ReportFileName = wRuta + "FACGUIAG.RPT"
           End If
           'Debug.Print frmdocu.Reportes.WindowTitle
           On Error GoTo accion
           frmdocu.Reportes.Action = 1
           On Error GoTo 0
         End If
       End If
    Else
pasa_todo:
       frmdocu.Reportes.SelectionFormula = pub_cadena
''       pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 5 AND {FACART.FAR_CODCIA} = '01' AND  ( {FACART.FAR_NUMFAC} >= 7 and {FACART.FAR_NUMFAC} <= 7) AND {FACART.FAR_NUMSER} = ''"
       frmdocu.Reportes.WindowTitle = frmdocu.Reportes.WindowTitle & " Archivo: " & Trim(frmdocu.Reportes.ReportFileName)
       On Error GoTo accion
       frmdocu.Reportes.Action = 1
       On Error GoTo 0
       If cherela.Value = 1 Then
          GoTo PASA_OP
       End If
    End If
Exit Function
accion:
'Debug.Print pub_cadena
 MsgBox Err.Description
 MsgBox "Intente Nuevamente, la impresion de Modo manual", 48, Pub_Titulo
  Exit Function
Resume Next
End Function

Private Sub TxtSubZonaTrabajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  FRADIRE.Visible = False
  Exit Sub
End If

If KeyAscii = 13 Then
 SQ_OPER = 1
 pu_cp = "C"
 pu_codcia = LK_CODCIA
 pu_codclie = Val(d_Codclie.Caption)
 LEER_CLI_LLAVE
 If cli_llave.EOF Then
   MsgBox "Cleinte no Existe NO Procede... ", 48, Pub_Titulo
   Exit Sub
 End If
 cli_llave.Edit
 cli_llave!CLI_TRAB_DIREC = Trim(txtdire.Text)
 cli_llave!CLI_TRAB_NUM = Val(txtnum.Text)
 cli_llave!cli_TRAB_ZONA = Val(Right(TxtZonaTrabajo.Text, 4))
 cli_llave!cli_TRAB_SUBZONA = Val(Right(TxtSubZonaTrabajo.Text, 4))
 cli_llave.Update
 FRADIRE.Visible = False
 txtnumfac_KeyPress 13
End If



End Sub

Private Sub txtvend_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  If Val(txtvend.Text) <> 0 Then
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODVEN = Nulo_Valor0(txtvend.Text)
    LEER_VEN_LLAVE
    If ven_llave.EOF Then
      MsgBox "Verificar Vendedor.", 48, Pub_Titulo
      Azul txtvend, txtvend
      Exit Sub
    End If
    cmbFBG.SetFocus
    Else
        cmbFBG.SetFocus
        Exit Sub
    End If
End If

End Sub
Public Sub ASIGNA_INT(WCONTROL As ComboBox, txt As Integer)
For fila = 0 To WCONTROL.ListCount - 1
    If Val(Trim(Right(WCONTROL.List(fila), 3))) = txt Then
        WCONTROL.ListIndex = fila
        Exit Sub
    End If
Next fila
End Sub

Public Sub LLENA_ZONA(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Private Sub TxtZonaTrabajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  FRADIRE.Visible = False
  Exit Sub
End If
If KeyAscii = 13 Then
 TxtSubZonaTrabajo.SetFocus
  SendKeysSeguro VK_UP, True
End If


End Sub

Public Sub EVENTO_CLICK()
If temporal = "X" Then
 Exit Sub
End If
If LK_FLAG_FACTURACION = "V" And LOPC_TIPMOV = 10 Then
 If Val(txtvend.Text) <> 0 Then
 If ven_llave.EOF Then
   MsgBox "Digite un Vendedor", 48, Pub_Titulo
   Exit Sub
 End If
 Else
   Exit Sub
 End If
End If

If LOC_TIPMOV = 97 And (Right(cmbFBG.Text, 1) = "P" Or Right(cmbFBG.Text, 1) = "C") Then
  pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC = ? AND FAR_CP = ?  AND FAR_ESTADO2 <> 'L'  ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
ElseIf LOC_TIPMOV = 99 Then
  pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC = ?  AND FAR_ESTADO2 <> 'L'  ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_NUMFAC, FAR_NUMSEC"
Else
  pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC = ?   AND FAR_ESTADO2 <> 'L' ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
End If
Set PSFAR = CN.CreateQuery("", pub_cadena)
 PSFAR.rdoParameters(0) = 0
 PSFAR.rdoParameters(1) = " "
 PSFAR.rdoParameters(2) = 0
 PSFAR.rdoParameters(3) = " "
 PSFAR.rdoParameters(4) = 0
 If LOC_TIPMOV = 97 And (Right(cmbFBG.Text, 1) = "P" Or Right(cmbFBG.Text, 1) = "C") Then
  PSFAR.rdoParameters(5) = " "
 End If
Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurValues)

LIMPIA_DOCU
grid_fac2.Rows = 2
txtNumfac.Text = ""
txtSerie.Text = ""
'SQ_OPER = 1
'PUB_CODCIA = LK_CODCIA
'LEER_PAR_LLAVE
'If par_llave.EOF Then
'  Exit Sub
'End If
txtSerie.Locked = False
lblsaldo.Caption = "Saldo Actual "
lbldomicilio.Caption = "Domicilio :"
lbldocu(1).Visible = True
lbldocu(3).Visible = True
lbldocu(5).Visible = True
d_dias.Visible = True
d_fechaV.Visible = True
d_newvcto.Visible = True

If Left(cmbFBG.Text, 1) = "P" Then
 lblcheque.Visible = True
 d_cheque.Visible = True
 d_efectivo.Visible = True
 lblEfectivo.Visible = True
 lblsaldo.Caption = "Total Planilla="
 'txtSerie.Locked = True
 lbldomicilio.Caption = "Vendedor :"
 txtSerie.Text = 0
 lbldocu(1).Visible = False
 lbldocu(3).Visible = False
 lbldocu(5).Visible = False
 d_dias.Visible = False
 d_fechaV.Visible = False
 d_newvcto.Visible = False
 Exit Sub
End If
If LK_FLAG_FACTURACION = "V" And LOPC_TIPMOV = 10 Then
   Select Case Left(cmbFBG.Text, 1)
   Case "G"
       txtSerie.Text = ven_llave!VEM_SERIE_G
   Case "B"
       txtSerie.Text = ven_llave!vem_serie_b
   Case "F"
       txtSerie.Text = ven_llave!vem_serie_f
   Case "P"
       txtSerie.Text = ven_llave!VEM_SERIE_P
   End Select
ElseIf LK_FLAG_FACTURACION = "A" Then
   If Left(cmbFBG.Text, 1) = "F" Then
     txtSerie.Text = par_llave!PAR_F_SERIE
   ElseIf Left(cmbFBG.Text, 1) = "B" Then
     txtSerie.Text = par_llave!PAR_B_SERIE
   ElseIf Left(cmbFBG.Text, 1) = "G" Then
     txtSerie.Text = par_llave!PAR_G_SERIE
   ElseIf Left(cmbFBG.Text, 1) = "P" Then
     txtSerie.Text = par_llave!PAR_P_SERIE
   End If
ElseIf LK_FLAG_FACTURACION = "U" Then
   If Left(cmbFBG.Text, 1) = "F" Then
     txtSerie.Text = usu_llave!USU_SERIE_F
   ElseIf Left(cmbFBG.Text, 1) = "B" Then
     txtSerie.Text = usu_llave!USU_SERIE_B
   End If
End If
If Left(cmbFBG.Text, 1) = "N" Or Left(cmbFBG.Text, 1) = "D" Then
     If Left(cmbFBG.Text, 1) = "D" Then
        txtSerie.Text = par_llave!PAR_SERIE_NDEB
     Else
        txtSerie.Text = par_llave!PAR_SERIE_NCRE
     End If
    If LK_FLAG_FACTURACION = "U" Then
      If Left(cmbFBG.Text, 1) = "D" Then
        txtSerie.Text = usu_llave!USU_SERIE_ND
      Else
        txtSerie.Text = usu_llave!USU_SERIE_NC
      End If
    End If
    lbldocu(1).Visible = False
    lbldocu(3).Visible = False
    lbldocu(5).Visible = False
    d_dias.Visible = False
    d_fechaV.Visible = False
    d_newvcto.Visible = False
End If
If LOC_TIPMOV = 10 Then
 PU_TIPMOV = 10
 PU_NUMSER = Val(txtSerie.Text)
 PU_FBG = Left(cmbFBG.Text, 1)
 pu_cp = "C"
ElseIf LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
 PU_TIPMOV = LOC_TIPMOV
 If Left(cmbFBG.Text, 1) = "C" Then
   pu_cp = "P"
   txtSerie.Text = 5
 End If
 
 PU_NUMSER = Val(txtSerie.Text)
 
 PU_FBG = Left(cmbFBG.Text, 1)
ElseIf LOC_TIPMOV = 20 Or LOC_TIPMOV = 99 Then
 PU_TIPMOV = LOC_TIPMOV
 pu_cp = "P"
 If LOC_TIPMOV = 20 Then
  txtSerie.Text = Val(par_llave!PAR_SER_KARDEX)
 End If
 PU_NUMSER = Val(txtSerie.Text)
 If LOC_TIPMOV = 99 Then
   PU_FBG = "K"
 Else
   PU_FBG = " "
 End If

Else
If LOC_TIPMOV = 5 Then
   PSCNT_LLAVE.rdoParameters(2) = 1
ElseIf LOC_TIPMOV = 6 Then
   PSCNT_LLAVE.rdoParameters(2) = 0
End If

If LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Then
   PSCNT_LLAVE.rdoParameters(0) = LK_CODCIA
   If LOC_TIPMOV = 5 Then
     txtSerie.Text = 4
     PSCNT_LLAVE.rdoParameters(1) = 2406
   Else
   txtSerie.Text = 1
     PSCNT_LLAVE.rdoParameters(1) = 2403
   End If
    PSCNT_LLAVE.rdoParameters(2) = 0
   cnt_llave.Requery
   If Not cnt_llave.EOF Then txtSerie.Text = Nulo_Valor0(cnt_llave!cnt_serie)
Else
 txtSerie.Text = "0"
End If
If LOC_TIPMOV = 3 Then pu_cp = "P"
If LOC_TIPMOV = 102 Then pu_cp = "C"
If LOC_TIPMOV = 100 Or LOC_TIPMOV = 101 Or LOC_TIPMOV = 93 Then
pu_cp = " "
If LOC_TIPMOV = 100 Then pu_cp = "C"
 txtSerie.Text = "1"
End If

 PU_TIPMOV = LOC_TIPMOV
 PU_NUMSER = Val(txtSerie.Text)
 PU_FBG = " "

End If
If LOC_TIPMOV <> 10 Then
  Screen.MousePointer = 11
  pu_codcia = LK_CODCIA
  LEER_FAR_CONSUL
  If Not far_consul.EOF Then
   txtNumfac.Text = far_consul!far_numfac
  Else
   txtNumfac.Text = ""
  End If
   Screen.MousePointer = 0
  Azul txtNumfac, txtNumfac
Else
  pu_codcia = LK_CODCIA
  txtNumfac.Text = ""
  Azul txtSerie, txtSerie
End If

If Trim(d_fecha.Caption) = "" Then txtnumfac_KeyPress 13

Exit Sub



End Sub

Public Sub EVENTO_LOSFOCUS_TIPMOV()
cmbFBG.Clear
If Trim(TIPMOV.Text) = "" Then
 LOC_TIPMOV = 0
Else
 LOC_TIPMOV = Val(Trim(Right(TIPMOV.Text, 4)))
 cmdImp.Enabled = False
 lblpersona.Visible = True
 d_Codclie.Visible = True
 LBLRUC.Visible = True
 'lblven.Visible = True
 d_codven.Visible = True
 d_condicion.Visible = True
 lblcondicion.Visible = True
 lbldireccion.Visible = True
 If LOC_TIPMOV = 101 Or LOC_TIPMOV = 93 Or LOC_TIPMOV = 20 Or LOC_TIPMOV = 5 Or LOC_TIPMOV = 6 Or LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
   cmdImp.Enabled = True
 ElseIf LOC_TIPMOV = 96 Then
   cmdImp.Enabled = True
   lblpersona.Visible = False
   d_Codclie.Visible = False
   LBLRUC.Visible = False
'   lblven.Visible = False
   d_codven.Visible = False
   d_condicion.Visible = False
   lblcondicion.Visible = False
   lbldireccion.Visible = False
 End If
 If temporal = "X" Then
  Exit Sub
 End If
 cmbFBG.Clear
 If LOC_TIPMOV = 10 Or LOC_TIPMOV = 97 Or LOC_TIPMOV = 98 Then
  If LOC_TIPMOV = 10 Then
    lblflete.Caption = "Flete"
  Else
    lblflete.Caption = "Ajuste"
  End If
  lblNumfac.Caption = "Nro  de Doc."
  lblpersona.Caption = "Cliente :"
'  lblven.Visible = True
  If LOC_TIPMOV = 97 Then
    cmbFBG.AddItem "N.C. Clientes              C"
    cmbFBG.AddItem "C.N.C. Proveedor             P"
  ElseIf LOC_TIPMOV = 98 Then
    cmbFBG.AddItem "Debito Clientes            C"
    cmbFBG.AddItem "A.Debito Proveedor           P"
  Else
   cmbFBG.AddItem "F = Facturas"
   cmbFBG.AddItem "B = Boletas"
   If LK_FLAG_GRIFO <> "A" Then
    cmbFBG.AddItem "G = Guias"
    cmbFBG.AddItem "P = Pedidos"
   Else
    cmbFBG.AddItem "P = Auto.C."
    cmbFBG.AddItem "G = O/D.Cont."
    cmbFBG.AddItem "C = O/D.Cred"
   End If
  End If
 ElseIf LOC_TIPMOV = 20 Then
'  lblven.Visible = False
  lblpersona.Caption = "Proveedor :"
  lblNumfac.Caption = "Nro de Kardex"
  lblflete.Caption = "Flete"
  cmbFBG.AddItem "K = Kardex"
  cmbFBG.AddItem "F = Facturas"
  cmbFBG.AddItem "G = Guias"
  cmbFBG.ListIndex = 0
 ElseIf LOC_TIPMOV = 99 Or LOC_TIPMOV = 30 Then
  lblpersona.Caption = "Proveedor :"
  lblNumfac.Caption = "Nro de Kardex"
  lblflete.Caption = "Flete"
  cmbFBG.AddItem "K = Kardex"
  If LOC_TIPMOV = 30 Then cmdImp.Enabled = True
 ElseIf LOC_TIPMOV = 96 Then
  lblNumfac.Caption = "Panilla"
  lblflete.Caption = ""
  cmbFBG.AddItem "P = Planillas"
 Else
 If LOC_TIPMOV = 102 Then
 Else
  d_codven.Visible = False
  LBLRUC.Visible = False
  d_Codclie.Visible = False
  lblpersona.Caption = ""
 End If
  lblNumfac.Caption = "Guia"
  lblflete.Caption = ""
  cmbFBG.AddItem "G = Guia"
 End If
 If LOC_TIPMOV = 10 And LK_FLAG_FACTURACION = "V" Then
     txtvend.Visible = True
     lblvend.Visible = True
 Else
     txtvend.Visible = False
     lblvend.Visible = False
 End If
' temporal = "X"
 'cmbFBG.ListIndex = 0
' temporal = ""
 If txtvend.Visible Then
   txtvend.SetFocus
 Else
   If cmbFBG.Visible And cmbFBG.Enabled Then cmbFBG.SetFocus
   If Trim(d_fecha.Caption) = "" Then cmbFBG_KeyPress 13
 End If
 
End If

End Sub

Public Sub MUESTA_DOCUMENTO()
    loc_flag_espera = "A"
    LLENA_CONSULTA
    loc_flag_espera = ""
    txtNumfac.SetFocus
End Sub

Private Sub lblbarraos_DblClick()
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
        Me.Move (Screen.Width - (Label3(5).Left + Label3(5).Width + 240)) / 2, _
                (Screen.Height - (Label3(5).Height + 540)) / 2, _
                Label3(5).Left + Label3(5).Width + 240, _
                Label3(5).Height + 540
    Else
        Me.WindowState = vbMaximized
    End If
End Sub

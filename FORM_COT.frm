VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FORM_COT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Pedidos."
   ClientHeight    =   4890
   ClientLeft      =   1500
   ClientTop       =   1140
   ClientWidth     =   6600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FORM_COT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4890
   ScaleWidth      =   6600
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "&Pedidos Pendientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   93
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame frapedido 
      Caption         =   "Pedidos Generados "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      TabIndex        =   86
      Top             =   1800
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CheckBox cheped 
         Caption         =   "Solo Pendientes ï¿½ Fac\Bol Anulados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   92
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton cmdmos 
         Caption         =   "Mos&trar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   91
         Top             =   240
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid gridf 
         Height          =   2775
         Left            =   120
         TabIndex        =   87
         Top             =   840
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4895
         _Version        =   393216
         Cols            =   5
         FixedRows       =   2
         BackColorBkg    =   4210688
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
      Begin MSMask.MaskEdBox fechaf 
         Height          =   285
         Left            =   960
         TabIndex        =   88
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin VB.Label lblped 
         AutoSize        =   -1  'True
         Caption         =   "Espere... Procesando ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   94
         Top             =   2280
         Width           =   1710
      End
      Begin VB.Label lblped 
         Caption         =   "Opciones :  [Enter] = Mostar el Pedido   -    [DEL] =Anular Pedidos  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   90
         Top             =   600
         Width           =   8415
      End
      Begin VB.Label lblped 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   89
         Top             =   240
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GPED 
      Height          =   1995
      Left            =   9480
      TabIndex        =   85
      Top             =   3900
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   3519
      _Version        =   393216
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   0
      BackColorBkg    =   8421376
      HighLight       =   0
      AllowUserResizing=   1
      BorderStyle     =   0
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
   Begin MSComctlLib.ListView ListView2 
      Height          =   495
      Left            =   7470
      TabIndex        =   68
      Top             =   7350
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   495
      Left            =   5670
      TabIndex        =   25
      Top             =   7335
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton c_condi 
      Caption         =   "Condiciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2085
      TabIndex        =   61
      Top             =   7650
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame f1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   23
      Top             =   30
      Width           =   11775
      Begin VB.ComboBox cmdtipo 
         Height          =   315
         ItemData        =   "FORM_COT.frx":0442
         Left            =   8430
         List            =   "FORM_COT.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   2745
      End
      Begin VB.TextBox i_dias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8040
         TabIndex        =   4
         Top             =   855
         Width           =   495
      End
      Begin VB.ComboBox i_condi 
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
         ItemData        =   "FORM_COT.frx":0456
         Left            =   120
         List            =   "FORM_COT.frx":0458
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   375
         Width           =   2730
      End
      Begin VB.ComboBox i_destino 
         Height          =   315
         ItemData        =   "FORM_COT.frx":045A
         Left            =   1620
         List            =   "FORM_COT.frx":0464
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   5490
      End
      Begin VB.ComboBox i_fbg 
         Height          =   315
         ItemData        =   "FORM_COT.frx":047D
         Left            =   8040
         List            =   "FORM_COT.frx":0487
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1260
         Width           =   1110
      End
      Begin VB.TextBox Txt_key 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4245
         MaxLength       =   10
         TabIndex        =   1
         Top             =   375
         Width           =   795
      End
      Begin VB.ComboBox moneda 
         Height          =   315
         ItemData        =   "FORM_COT.frx":0491
         Left            =   3120
         List            =   "FORM_COT.frx":049B
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtruc 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8625
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   855
         Width           =   1515
      End
      Begin VB.TextBox txtcli 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Top             =   870
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8460
         TabIndex        =   80
         Top             =   120
         Width           =   2010
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "F5 = Ingreso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   210
         Index           =   17
         Left            =   10680
         TabIndex        =   77
         Tag             =   "9999"
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "F3 = Condiï¿½n"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   210
         Index           =   15
         Left            =   10680
         TabIndex        =   71
         Tag             =   "9999"
         Top             =   1035
         Width           =   945
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "F4 = Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   210
         Index           =   14
         Left            =   10680
         TabIndex        =   70
         Tag             =   "9999"
         Top             =   1245
         Width           =   1005
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         Caption         =   "F2 = Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   210
         Index           =   13
         Left            =   10680
         TabIndex        =   69
         Tag             =   "9999"
         Top             =   825
         Width           =   870
      End
      Begin VB.Label lcodart 
         Caption         =   "Dias Cred."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   7170
         TabIndex        =   58
         Tag             =   "9999"
         Top             =   885
         Width           =   810
      End
      Begin VB.Label lcodart 
         Caption         =   "Condición Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   150
         TabIndex        =   57
         Tag             =   "9999"
         Top             =   120
         Width           =   1995
      End
      Begin VB.Label lcodart 
         Caption         =   "   Destino Almacen :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   75
         TabIndex        =   56
         Tag             =   "9999"
         Top             =   1305
         Width           =   1470
      End
      Begin VB.Label lblven 
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
         Height          =   270
         Left            =   5160
         TabIndex        =   55
         Top             =   360
         Width           =   3150
      End
      Begin VB.Label lcodart 
         Caption         =   "Fact./Bolet."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   7200
         TabIndex        =   54
         Tag             =   "9999"
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4185
         TabIndex        =   53
         Top             =   150
         Width           =   915
      End
      Begin VB.Label lblcli 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3060
         TabIndex        =   42
         Top             =   855
         Width           =   4020
      End
      Begin VB.Label lcodart 
         Caption         =   "Moneda : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3150
         TabIndex        =   36
         Tag             =   "9999"
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lcodart 
         Caption         =   "R.U.C. ï¿½ D.N.I. :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   8700
         TabIndex        =   34
         Tag             =   "9999"
         Top             =   660
         Width           =   1110
      End
      Begin VB.Label lcodart 
         Caption         =   "Codigo / Raz. Soc. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   24
         Tag             =   "9999"
         Top             =   855
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Height          =   5325
      Left            =   9405
      TabIndex        =   60
      Top             =   1770
      Width           =   2385
      Begin VB.CommandButton siguiente 
         Caption         =   "Sigui..>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1485
         TabIndex        =   15
         Top             =   4740
         Width           =   675
      End
      Begin VB.CommandButton anterior 
         Caption         =   "<Ant."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   735
         TabIndex        =   14
         Top             =   4755
         Width           =   690
      End
      Begin VB.TextBox txtdoc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1230
         TabIndex        =   17
         Top             =   4395
         Width           =   1005
      End
      Begin VB.TextBox tserie 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   705
         TabIndex        =   16
         Top             =   4380
         Width           =   465
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   285
         Left            =   930
         TabIndex        =   72
         Top             =   1320
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin VB.Label lcodart 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         Caption         =   "Pedidos Pendientes x Vend."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   20
         Left            =   165
         TabIndex        =   84
         Tag             =   "9999"
         Top             =   1920
         Width           =   2040
      End
      Begin VB.Label lblvisita 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   945
         TabIndex        =   83
         Tag             =   "9999"
         Top             =   135
         Width           =   1290
      End
      Begin VB.Label lcodart 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         Caption         =   "Dia Visita:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   19
         Left            =   60
         TabIndex        =   82
         Tag             =   "9999"
         Top             =   150
         Width           =   705
      End
      Begin VB.Label lbluser 
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
         Height          =   255
         Left            =   945
         TabIndex        =   79
         Tag             =   "9999"
         Top             =   1650
         Width           =   1275
      End
      Begin VB.Label lcodart 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   18
         Left            =   255
         TabIndex        =   78
         Tag             =   "9999"
         Top             =   1665
         Width           =   600
      End
      Begin VB.Label lcodart 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   16
         Left            =   195
         TabIndex        =   74
         Tag             =   "9999"
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label lcodart 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         Caption         =   "Disponib :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   73
         Tag             =   "9999"
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label lblDeuda 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   945
         TabIndex        =   67
         Tag             =   "9999"
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label lcodart 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         Caption         =   "Deuda  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   12
         Left            =   285
         TabIndex        =   66
         Tag             =   "9999"
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lbldisp 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   945
         TabIndex        =   65
         Tag             =   "9999"
         Top             =   1035
         Width           =   1290
      End
      Begin VB.Label lblcred 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   945
         TabIndex        =   64
         Tag             =   "9999"
         Top             =   435
         Width           =   1290
      End
      Begin VB.Label lcodart 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         Caption         =   "Lim. Cred. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   75
         TabIndex        =   63
         Tag             =   "9999"
         Top             =   435
         Width           =   795
      End
      Begin VB.Label lcodart 
         BackColor       =   &H80000001&
         Caption         =   "Nro de Pedido:"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   5
         Left            =   870
         TabIndex        =   62
         Tag             =   "9999"
         Top             =   4155
         Width           =   1290
      End
   End
   Begin VB.TextBox txtatte 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6165
      TabIndex        =   59
      Top             =   8295
      Width           =   5175
   End
   Begin VB.Frame condi 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2910
      TabIndex        =   45
      Top             =   8115
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   52
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox oferta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   51
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox tiempo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   49
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox forma 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   47
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Validez de la Oferta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   795
         Width           =   1695
      End
      Begin VB.Label c_entrega 
         Caption         =   "Tiempo de Entrega:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label c_forma 
         Caption         =   "Forma de Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   46
         Top             =   165
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdimp 
      Caption         =   "&Anular Ped."
      Height          =   465
      Left            =   5205
      Picture         =   "FORM_COT.frx":04B4
      TabIndex        =   12
      Top             =   6465
      Width           =   1275
   End
   Begin VB.CommandButton cmdconsulta 
      Caption         =   "Co&nsultas"
      Height          =   465
      Left            =   1920
      Picture         =   "FORM_COT.frx":05FE
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdIngreso 
      Caption         =   "&Ingreso"
      Height          =   465
      Left            =   420
      Picture         =   "FORM_COT.frx":0748
      TabIndex        =   8
      Top             =   6495
      Width           =   1215
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "Ce&rrar"
      Height          =   465
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   6450
      Width           =   1260
   End
   Begin VB.CommandButton cancelar 
      Caption         =   "&Cancelar"
      Height          =   465
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   6480
      Width           =   1260
   End
   Begin VB.Frame ESTADO 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   5325
      Left            =   75
      TabIndex        =   18
      Tag             =   "100"
      Top             =   1770
      Width           =   9300
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   180
         Left            =   345
         TabIndex        =   81
         Top             =   4500
         Visible         =   0   'False
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.ComboBox PRECIOS 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox UNIDAD 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txttotal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6660
         TabIndex        =   28
         Top             =   4200
         Width           =   1320
      End
      Begin VB.TextBox txtigv 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3765
         TabIndex        =   27
         Top             =   4185
         Width           =   1230
      End
      Begin VB.TextBox txtvalorv 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1365
         TabIndex        =   26
         Top             =   4185
         Width           =   1200
      End
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   16776960
         BorderStyle     =   0
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"FORM_COT.frx":0892
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2640
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   3420
         Left            =   120
         TabIndex        =   6
         Tag             =   "9999"
         Top             =   330
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   6033
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         BackColor       =   16777215
         FocusRect       =   2
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
      End
      Begin VB.Label LBLSIT 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3900
         TabIndex        =   76
         Top             =   150
         Width           =   5040
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Condición :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2820
         TabIndex        =   75
         Top             =   135
         Width           =   975
      End
      Begin VB.Label nomarti 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   975
         TabIndex        =   38
         Top             =   3825
         Width           =   5175
      End
      Begin VB.Label unid 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6255
         TabIndex        =   41
         Top             =   3825
         Width           =   975
      End
      Begin VB.Label stock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         Height          =   255
         Left            =   7335
         TabIndex        =   40
         Top             =   3825
         Width           =   1095
      End
      Begin VB.Label lblpro 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   135
         TabIndex        =   39
         Top             =   3825
         Width           =   975
      End
      Begin VB.Label i_moneda 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         Height          =   195
         Left            =   6285
         TabIndex        =   37
         Top             =   4185
         Width           =   285
      End
      Begin VB.Label Label3 
         Caption         =   "Total :"
         Height          =   195
         Index           =   3
         Left            =   5325
         TabIndex        =   31
         Tag             =   "9999"
         Top             =   4185
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "I.G.V. :"
         Height          =   195
         Index           =   2
         Left            =   2925
         TabIndex        =   30
         Tag             =   "9999"
         Top             =   4185
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Valor Venta:"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   29
         Tag             =   "9999"
         Top             =   4185
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Detalle del Pedido :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Tag             =   "9999"
         Top             =   120
         Width           =   2160
      End
      Begin VB.Label momen 
         Caption         =   "Un Momento ..."
         Height          =   375
         Left            =   3360
         TabIndex        =   20
         Top             =   1560
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   360
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lcodart 
      Caption         =   "Serie"
      Height          =   255
      Index           =   3
      Left            =   7440
      TabIndex        =   44
      Tag             =   "9999"
      Top             =   240
      Width           =   525
   End
   Begin VB.Label lcodart 
      Caption         =   "Nro. Doc"
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   43
      Tag             =   "9999"
      Top             =   240
      Width           =   1125
   End
End
Attribute VB_Name = "FORM_COT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim peddoc_llave As rdoResultset
Dim PSDOC_LLAVE As rdoQuery
Dim PSFAR_TRANS As rdoQuery
Dim FAR_TRANS As rdoResultset

Dim VAR_ACTIVAR As Integer
Dim WCOD_ORIGINAL As Currency
Dim WPASA As Boolean
Dim WSELE As String * 1
Dim llave1
Dim loc_key
Dim fila As Integer
Dim ws_bruto_d, ws_bruto_h As Currency
Dim SUM_D As Currency
Dim SUM_H As Currency
Dim PSTEMP_LLAVE As rdoQuery
Dim temp_llave As rdoResultset
Dim WMODO As String * 1
Dim LOC_ITEM As Integer
Dim cop_llave As rdoResultset
Dim PSCOP_LLAVE As rdoQuery
Dim LOC_CANCELA As Integer
Dim PSTEMP_MAYOR As rdoQuery
Dim temp_mayor As rdoResultset
Dim temporal
Dim wfila_act As Integer
Dim loc_ini As Integer
Dim loc_fin  As Integer
Dim Wsec As Integer
Dim PSLOC_WARTI As rdoQuery
Dim llave_sum_arti   As rdoResultset
Dim PRE_ETIQUETA(5) As String * 20
Dim LOC_TIPMOV As Integer

Option Explicit

Private Sub anterior_Click()
If Val(txtdoc.Text) <= 0 Then Exit Sub
 txtdoc.Text = Val(txtdoc.Text) - 1
 PUB_NUMSER = Val(tserie.Text)
 PUB_NUMFAC = Val(txtdoc.Text)
 LLENA_DOCU
End Sub

Private Sub c_condi_Click()
If condi.Visible Then
 condi.Visible = False
Else
 condi.Visible = True
 forma.SetFocus
 
End If
End Sub

Private Sub cancelar_Click()
cmdimp.Visible = False
WMODO = ""
cmdIngreso.Caption = "&Ingreso"
f1.Enabled = False
ESTADO.Enabled = False
pb.Visible = False
fila = 0
SUM_D = 0
SUM_H = 0
LIMPIA_DATOS
CABE_MAN
f1.Enabled = False
cmdIngreso.Enabled = True
tserie.Enabled = True
txtdoc.Enabled = True
anterior.Enabled = True
siguiente.Enabled = True
cmdconsulta.Enabled = True
GPED.Rows = 2
'grid_fac.SetFocus

End Sub

Private Sub cheped_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  frapedido.Visible = False
  Command2.SetFocus
  Exit Sub
End If

End Sub

Private Sub cmdconsulta_Click()
cmdimp.Visible = True
cmdIngreso.Enabled = False
tserie.Locked = False
txtdoc.Locked = False
tserie.Enabled = True
txtdoc.Enabled = True
'siguiente.Enabled = True
'anterior.Enabled = True
f1.Enabled = True
tserie.Text = "100"
tserie.Locked = True
Azul txtdoc, txtdoc

End Sub

Private Sub cmdimp_Click()
Dim WS_CANTIDAD As Currency
'Call REP_CONSUL
llave_sum_arti.MoveFirst
If llave_sum_arti.EOF Then
 MsgBox "No Procede..Intente Nuevamente.", 48, Pub_Titulo
 Exit Sub
End If
If Trim(Left(LBLSIT.Caption, 2)) = "01" Then
  MsgBox "Pedido esta Procesado... no procede " & Chr(13) & "Anula el Documento para poder Anular el Pedido.", 48, Pub_Titulo
  Exit Sub
End If
If Val(Left(LBLSIT.Caption, 2)) = 99 Then
  MsgBox "Pedido Anulado... no procede.", 48, Pub_Titulo
  Exit Sub
End If

pub_mensaje = "Anular el pedidos...   ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
pub_signo_arm = -1
If UCase(cmdIngreso.Caption) = "&GRABAR" Then
    ' REGRESA EL STOCK AL ANTERIOR
    '******************
    llave_sum_arti.MoveFirst
    Do Until llave_sum_arti.EOF
       SQ_OPER = 1
       PUB_CODART = llave_sum_arti!PED_CODART
       pu_codcia = LK_CODCIA
       LEER_ARM_LLAVE
       WS_CANTIDAD = Val(llave_sum_arti!PED_cantidad) * Val(llave_sum_arti!PED_EQUIV)
       arm_llave.Edit
       arm_llave!arm_stock = arm_llave!arm_stock + WS_CANTIDAD
       arm_llave!ARM_INGRESOS = arm_llave!ARM_INGRESOS + WS_CANTIDAD
       arm_llave.Update
       llave_sum_arti.MoveNext
    Loop
    If Trim(txtdoc.Text) <> 0 Then
      pub_signo_arm = -1
      pub_cadena = "UPDATE PEDIDOS SET PED_SITUACION = 'E' WHERE PED_TIPMOV = 201 AND PED_CODCIA = '" & LK_CODCIA & "' AND PED_NUMSER = " & Trim(tserie.Text) & " AND PED_NUMFAC = " & Trim(txtdoc.Text)
      CN.Execute pub_cadena, rdExecDirect
    End If
End If
'******************
cancelar_Click
End Sub

Private Sub cmdIngreso_Click()
cmdIngreso.Enabled = False
Dim WS_CANTIDAD As Currency
Dim wau As Currency
Dim WMO As String
Dim RES_DEUDA As Currency
Dim wsumadol As Currency
Dim WTC As Currency
Dim ws_tot_debe, ws_tot_haber As Currency
Dim er As rdoError
Dim pub_mensaje As String
Const ingre = 2
Const MODIF = 1
Dim N As Integer
Dim LOC_SALDO_CAR As Currency
Dim FLAG As Boolean
Dim pub_mensaje_err As String
Dim WS_NRO_MOV, ws_nro_voucher As Long
Dim w_dh  As String

If Left(cmdIngreso.Caption, 2) = "&G" Then
 If Val(Left(LBLSIT.Caption, 2)) = 1 Then
   cmdIngreso.Enabled = True
   MsgBox "Pedido esta Procesado... no procede.", 48, Pub_Titulo
   Exit Sub
 End If
 If Val(Left(LBLSIT.Caption, 2)) = 99 Then
   cmdIngreso.Enabled = True
   MsgBox "Pedido Anulado... no procede.", 48, Pub_Titulo
   Exit Sub
 End If
If Trim(txtcli.Text) = "" Then
   cmdIngreso.Enabled = True
  MsgBox "Nombre del Cliente ", 48, Pub_Titulo
  txtcli.SetFocus
  Exit Sub
End If
If Val(txttotal.Text) <= 0 Then
   cmdIngreso.Enabled = True
  MsgBox "Ingrese Datos ", 48, Pub_Titulo
  grid_fac.SetFocus
  Exit Sub
End If

If Nulo_Valor0(SUT_LLAVE!sut_TIPDOC) = "FA" And Nulo_Valor0(SUT_LLAVE!SUT_SIGNO_CAR) = 1 Then
   If Val(i_dias.Text) = 0 Then
    cmdIngreso.Enabled = True
    MsgBox "Falta Dias de Credito", 48, Pub_Titulo
    Azul i_dias, i_dias
    Exit Sub
   End If
Else
   i_dias.Text = "0"
End If



If grid_fac.Rows = 3 Then
 If grid_fac.TextMatrix(2, 0) = "" Then
   cmdIngreso.Enabled = True
   MsgBox "Ingrese Datos de Productos ", 48, Pub_Titulo
   grid_fac.SetFocus
   Exit Sub
 End If
End If
pu_codcia = LK_CODCIA
pu_cp = "C"
pu_codclie = Val(txtcli.Text)
LEER_CLI_LLAVE
On Error GoTo 0
If cli_llave.EOF Then
  cmdIngreso.Enabled = True
  Azul txtcli, txtcli
  MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
  txtcli.SetFocus
  Exit Sub
End If
If Trim(i_destino.Text) = "" Then
  cmdIngreso.Enabled = True
  MsgBox "Definir la Direccion de Almacen ...", 48, Pub_Titulo
  i_destino.SetFocus
  Exit Sub
End If
suma_grid ' calcula para chequear
If Trim(i_fbg.Text) = "F" Then
 If Trim(cli_llave!cli_ruc_esposo) <> "" Then
      If Len(Trim(cli_llave!cli_ruc_esposo)) <> LK_DIG_RUC Then
         cmdIngreso.Enabled = True
         MsgBox "R.U.C. del Cliente no es de " & LK_DIG_RUC & " Digitos", 48, Pub_Titulo
         Azul txtcli, txtcli
         Exit Sub
      End If
 Else
    cmdIngreso.Enabled = True
    MsgBox "Cliente No Tiene R.U.C. . ", 48, Pub_Titulo
    grid_fac.SetFocus
    ' Azul txtcli, txtcli
    Exit Sub
 End If
Else
  If Val(txttotal.Text) > Val(GEN!GEN_UIT) And Trim(cli_llave!cli_RUC_ESPOSA) = "" Then
     cmdIngreso.Enabled = True
     MsgBox "Cliente supero el limite de la U.I.T. - Declararlo con su D.N.I . ", 48, Pub_Titulo
     Azul txtcli, txtcli
     Exit Sub
   End If
End If
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE
If par_llave!par_flag_cierre = 9 Then
   MsgBox "!!! Compañia a Cerredo Operaciones ... ", 48, Pub_Titulo
   Exit Sub
End If
If par_llave!PAR_FECHA_DIA <> LK_FECHA_DIA Then
   MsgBox "LA FECHA DE LA COMPAñIA HA CAMBIADO NO COINCIDE CON LA ACTUAL , OTRO USUARIO A CERRADO EL DIA!!! SALGA Y REINICIE SU SISTEMA...", 48, Pub_Titulo
   End
   Exit Sub
End If

'chequeo el limite de credito
pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)

' Actualizar parar el Numero de Documento Correlativo
If UCase(cmdIngreso.Caption) = "&GRABAR" Then
Else
llena_numfac
End If
If Not cli_llave.EOF And Nulo_Valors(par_llave!par_flag_cred) <> "A" And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 Then
   If SUT_LLAVE!SUT_SIGNO_CAR = 1 Then
      pu_codcia = LK_CODCIA
      pub_deuda = CAR_TOT_CPX2("C", pu_codcia, cli_llave!cli_codclie)
      If PUB_FLAG_VENCIDO = 1 And LK_FLAG_LIMITE <> "A" And LK_FLAG_LIMITE <> "C" Then
          cmdIngreso.Enabled = True
          con_llave.Close
          MsgBox "CLIENTE TIENE OBLIGACIONES VENCIDAS ... ", 48, Pub_Titulo
          Exit Sub
    End If
    PUB_CAL_INI = LK_FECHA_DIA
    PUB_CAL_FIN = LK_FECHA_DIA
    pu_codcia = LK_CODCIA
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 1
    LEER_CAL_LLAVE
    WTC = 0
    If Not cal_llave.EOF Then
      WTC = Nulo_Valor0(cal_llave!cal_tipo_cambio)
    End If
    If par_llave!PAR_MONEDA_FAC = "S" Then
       WTC = 1
    Else
     If WTC = 0 Then
      con_llave.Close
      cmdIngreso.Enabled = True
      MsgBox "Venta falta parametros ...INGRESE TIPO DE CAMBIO DEL DIA", 48, Pub_Titulo
      Exit Sub
     End If
    End If
    If Trim(Left(moneda.Text, 1)) = "S" Then
     wsumadol = Val(Nulo_Valor0(cli_llave!cli_limcre)) + Val(redondea((Nulo_Valor0(cli_llave!cli_limcre2) * WTC)))
     RES_DEUDA = pub_deuda
     WMO = "S/."
    Else
     wsumadol = Val(redondea(Nulo_Valor0(cli_llave!cli_limcre) / WTC)) + Val(redondea(Val(Nulo_Valor0(cli_llave!cli_limcre2))))
     RES_DEUDA = redondea(Val(pub_deuda / WTC))
     WMO = "US$."
    End If
If (RES_DEUDA + Val(txttotal.Text)) > wsumadol And LK_FLAG_LIMITE <> "B" And LK_FLAG_LIMITE <> "C" Then
   con_llave.Close
   cmdIngreso.Enabled = True
   MsgBox "LIMITE DE CREDITO EXCEDIDO ...SALDO POR ATENDER : " & WMO & " " & Format(wsumadol - RES_DEUDA, "0.00") & Chr(13) & "*** Venta No Procede ***", 48, Pub_Titulo
   Azul txtcli, txtcli
   Exit Sub
End If
   End If
End If

Barra.Visible = False
wau = 0 ' incrementar alguna merma opcional
For fila = 2 To grid_fac.Rows - 1
    PUB_CANTIDAD = Val(grid_fac.TextMatrix(fila, 2))
    PUB_CODART = Val(grid_fac.TextMatrix(fila, 10))
    If Trim(LK_USU_STOCK) = "" And PUB_CODART <> 0 And PUB_CANTIDAD <> 0 Then
        pu_codcia = LK_CODCIA
        SQ_OPER = 1
        LEER_ARM_LLAVE
        PUB_CANTIDAD = Val(grid_fac.TextMatrix(fila, 2)) * Val(grid_fac.TextMatrix(fila, 12))
        If PUB_CANTIDAD > (Val(arm_llave!arm_stock) + wau) And Nulo_Valors(art_LLAVE!art_flag_stock) <> "A" Then
           cmdIngreso.Enabled = True
           con_llave.Close
           MsgBox "La Cantidad en Unidades Minimas es Mayor que el Stock Actual ..." & Chr(13) & "Existencia : " & Trim(grid_fac.TextMatrix(fila, 0)) & " : " & Format(arm_llave!arm_stock, "0.00") & " ......Salida Requerida : " & PUB_CANTIDAD, 48, Pub_Titulo
           grid_fac.SetFocus
           GoTo fin
        End If
     End If
  
 If grid_fac.TextMatrix(fila, 1) <> "" Then
  If Val(grid_fac.TextMatrix(fila, 14)) = 0 Then ' verificar el Numero de precios seleccionado
    cmdIngreso.Enabled = True
    con_llave.Close
    MsgBox "Por favor Verificar el Producto :" & Chr(13) & fila - 1 & " - " & grid_fac.TextMatrix(fila, 0) & Chr(13) & "Intente digitarlo nuevamente. ", 48, Pub_Titulo
    grid_fac.SetFocus
    GoTo fin
  End If
  
  If Val(grid_fac.TextMatrix(fila, 2)) <= 0 Then
    con_llave.Close
    MsgBox "Verificar, cantidad en cero o menor. - " & grid_fac.TextMatrix(fila, 1) & " : " & grid_fac.TextMatrix(fila, 0), 48, Pub_Titulo
    grid_fac.SetFocus
    GoTo fin
  End If
  If Val(grid_fac.TextMatrix(fila, 4)) = 0 Then
  '  con_llave.Close
 ''   MsgBox "Verificar hay algun precio en 0 .", 48, Pub_Titulo
    grid_fac.SetFocus
 ''   GoTo fin
  End If
End If
Next fila
Screen.MousePointer = 11

DoEvents
Barra.Visible = True
DoEvents
Barra.Min = 0
Barra.max = fila
Barra.Value = 0
exito = True
Barra.Value = 1
CN.Execute "Begin Transaction", rdExecDirect
On Error GoTo grave_error
GoSub ACT1
'Call REP_CONSUL
CN.Execute "Commit Transaction", rdExecDirect
On Error GoTo 0
If con_llave Is Nothing Then
Else
  con_llave.Close
End If
fila = 1
SUM_D = 0
SUM_H = 0
CABE_MAN
LIMPIA_DATOS
fila = 0
'cancelar.SetFocus
CABE_MAN
PUB_CODCIA = LK_CODCIA
PUB_CODVEN = Val(Txt_key.Text)
WMO = PED_PENDIENTES()
Barra.Visible = False
f1.Enabled = False
cmdIngreso.Caption = "&Ingreso"
cmdIngreso.Enabled = True
cmdIngreso.SetFocus
GoTo fin

ACT1:
pub_signo_arm = -1
If UCase(cmdIngreso.Caption) = "&GRABAR" Then
    ' REGRESA EL STOCK AL ANTERIOR
    '******************
    llave_sum_arti.MoveFirst
    Do Until llave_sum_arti.EOF
       SQ_OPER = 1
       PUB_CODART = llave_sum_arti!PED_CODART
       pu_codcia = LK_CODCIA
       LEER_ARM_LLAVE
       WS_CANTIDAD = Val(llave_sum_arti!PED_cantidad) * Val(llave_sum_arti!PED_EQUIV)
       arm_llave.Edit
       arm_llave!arm_stock = arm_llave!arm_stock + WS_CANTIDAD
       arm_llave!ARM_INGRESOS = arm_llave!ARM_INGRESOS + WS_CANTIDAD
       arm_llave.Update
       llave_sum_arti.MoveNext
    Loop
    If Trim(txtdoc.Text) <> 0 Then
      pub_signo_arm = -1
      pub_cadena = "DELETE PEDIDOS WHERE PED_TIPMOV = 201 AND PED_CODCIA = '" & LK_CODCIA & "' AND PED_NUMSER = " & Trim(tserie.Text) & " AND PED_NUMFAC = " & Trim(txtdoc.Text)
      CN.Execute pub_cadena, rdExecDirect
    End If

End If
'******************
fila = 1
FLAG = False
WS_NRO_MOV = 0
fila = 2
Do While FLAG = False
   If Trim(grid_fac.TextMatrix(fila, 1)) = "" Then GoTo pasa
    ' grabo todo
   WS_CANTIDAD = Val(grid_fac.TextMatrix(fila, 2)) * Val(grid_fac.TextMatrix(fila, 12))
   temp_llave.AddNew
   temp_llave!PED_CODCIA = LK_CODCIA
   temp_llave!PED_fecha = LK_FECHA_DIA
   temp_llave!PED_NUMSER = Trim(tserie.Text)
   temp_llave!PED_NUMFAC = Val(txtdoc.Text)
   temp_llave!PED_NUMSEC = WS_NRO_MOV
   temp_llave!PED_cantidad = WS_CANTIDAD
   temp_llave!PED_PRECIO = Val(grid_fac.TextMatrix(fila, 4))
   temp_llave!PED_CODUSU = LK_CODUSU
   temp_llave!PED_IGV = Val(txtigv.Text)
   temp_llave!PED_BRUTO = Val(txtvalorv.Text)
   temp_llave!PED_ESTADO = "N"
   temp_llave!PED_CODUSU = LK_CODUSU
   temp_llave!PED_CODART = Val(grid_fac.TextMatrix(fila, 10))
   temp_llave!PED_UNIDAD = Trim(grid_fac.TextMatrix(fila, 3))
   temp_llave!PED_EQUIV = Val(grid_fac.TextMatrix(fila, 12))
   temp_llave!PED_NOMCLIE = Trim(FORM_COT.lblcli.Caption)
   temp_llave!PED_RUCCLIE = Trim(txtruc.Text) ' Trim(fbg.Text)
   temp_llave!PED_CODCLIE = Val(txtcli.Text)
   temp_llave!PED_TIPMOV = 201
   temp_llave!PED_HORA = Format(Now, "hh:mm:ss AMPM")
   temp_llave!PED_descto = Val(grid_fac.TextMatrix(fila, 5))
   temp_llave!PED_MONEDA = Left(Trim(moneda.Text), 1)
   temp_llave!PED_CONTACTO = txtatte.Text
   temp_llave!PED_FORMA = Trim(forma.Text)
   temp_llave!PED_TIEMPO = Trim(tiempo.Text)
   temp_llave!PED_OFERTA = Trim(oferta.Text)
   temp_llave!PED_SUBTOTAL = Val(grid_fac.TextMatrix(fila, 6))
   temp_llave!PED_CONDI = Val(Left(i_condi.Text, 2))
   temp_llave!ped_DIAS = Val(i_dias.Text)
   temp_llave!PED_codven = Val(Txt_key.Text)
   temp_llave!ped_DIRCLI = Val(Right(i_destino.Text, 8))
   temp_llave!ped_FBG = Trim(i_fbg.Text)
   temp_llave!PED_NUMPRE = Val(grid_fac.TextMatrix(fila, 14))
   temp_llave!PED_NUM_UNIDAD = Val(grid_fac.TextMatrix(fila, 11))
   temp_llave!PED_TIPVTA = Val(Right(cmdtipo.Text, 8))
   temp_llave!PED_descto_pre = Val(grid_fac.TextMatrix(fila, 20))
   temp_llave!PED_DIA_VISITA = Val(cli_llave!CLI_DIA_VISITA)
   
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_CODART = Val(grid_fac.TextMatrix(fila, 10))
   LEER_ARM_LLAVE
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_SECUEN = Val(grid_fac.TextMatrix(fila, 18))
   LEER_PRE_LLAVE
   arm_llave.Edit
   arm_llave!arm_stock = Val(arm_llave!arm_stock) + WS_CANTIDAD * pub_signo_arm
   If pub_signo_arm = -1 Then
       arm_llave!ARM_SALIDAS = Val(arm_llave!ARM_SALIDAS) + Val(WS_CANTIDAD)
   ElseIf pub_signo_arm = 1 Then
       arm_llave!ARM_INGRESOS = Val(arm_llave!ARM_INGRESOS) + Val(WS_CANTIDAD)
   End If
   arm_llave.Update
   
   temp_llave.Update
pasa:
   fila = fila + 1
   WS_NRO_MOV = WS_NRO_MOV + 1
   If fila >= FORM_COT.grid_fac.Rows Then
      FLAG = True
   End If
  
Loop

Return
Screen.MousePointer = 1
Exit Sub
End If
' cuando pulsa Ingreso
Dim wser As String
Dim wnumfac As String

cmdIngreso.Caption = "&Grabar / Imprimir"

f1.Enabled = True
ESTADO.Enabled = True
LIMPIA_DATOS
lbluser.Caption = LK_CODUSU
CABE_MAN
WMODO = "I"
llena_numfac
grid_fac.Rows = grid_fac.Rows + 1
grid_fac.RowHeight(grid_fac.Rows - 1) = 285
grid_fac.TextMatrix(grid_fac.Rows - 1, 0) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 2) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 3) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 6) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 7) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 9) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 11) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 12) = ""

grid_fac.TextMatrix(grid_fac.Rows - 1, 14) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 15) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 16) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 17) = ""

If i_condi.ListCount > 0 And f1.Enabled = True Then
  i_condi.SetFocus
  SendKeysSeguro VK_UP, True
End If
If moneda.ListCount > 0 And moneda.ListIndex = -1 And f1.Enabled = True Then moneda.ListIndex = 0

cmdIngreso.Enabled = True
tserie.Enabled = False
txtdoc.Enabled = False
anterior.Enabled = False
siguiente.Enabled = False
cmdconsulta.Enabled = False

Exit Sub

Error_fatal:
    pub_mensaje = "Se ha producido un error " & "al abrir la conexiï¿½n:" & Err & " - " & Error & vbCr
    For Each er In rdoErrors
        pub_mensaje = pub_mensaje & er.Description & ":" & er.Number & vbCr
        MsgBox pub_mensaje
    Next er
    CN.Execute "Rollback Transaction", rdExecDirect
'    Resume AbandonCn
Exit Sub

errorr:
 MsgBox pub_mensaje_err, 48, Pub_Titulo
fin:
cmdIngreso.Enabled = True

Screen.MousePointer = 0
Exit Sub
SALE:
If Err.Number = 6 Then
  MsgBox "Verficar Importe.", 48, Pub_Titulo
  If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR
  FORM_COT.Barra.Visible = False
  Screen.MousePointer = 0
  grid_fac.SetFocus
Else
  MsgBox Err.Description, 48, Pub_Titulo
End If
Exit Sub

grave_error:
' Destruir el Begin Transaccion
 If con_llave Is Nothing Then
 Else
   con_llave.Close
 End If
 CN.Execute "Rollback Transaction", rdExecDirect
 MsgBox Err.Description & Chr(13) & "Reiniciar su Sistema "
 End
Exit Sub
End Sub





Private Sub cmdmos_Click()
Dim PSLOC_PEDIDOS As rdoQuery
Dim loc_pedidos   As rdoResultset
If Not IsDate(fechaf.Text) Then
   MsgBox "Fecha No Procede ", 48, Pub_Titulo
   Azul2 fechaf, fechaf
   Exit Sub
End If

pub_cadena = "SELECT DISTINCT PED_TIPMOV, PED_NUMSER, PED_NUMFAC, PED_CODCLIE, PED_CODVEN, PED_FECHA , PED_CONDI, PED_CODUSU, PED_NOMCLIE, PED_IGV,PED_BRUTO, PED_FBG, ped_situacion , PED_TIPVTA " & _
 "FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = ? " & _
 "AND (PED_FECHA >= ? AND PED_FECHA <= ?) AND PED_FBG IN (?,?) ORDER BY ped_fecha, ped_numfac "
Set PSLOC_PEDIDOS = CN.CreateQuery("", pub_cadena)
PSLOC_PEDIDOS(0) = LK_CODCIA
PSLOC_PEDIDOS(1) = 201
PSLOC_PEDIDOS(2) = fechaf.Text
PSLOC_PEDIDOS(3) = fechaf.Text
PSLOC_PEDIDOS(4) = "F"
PSLOC_PEDIDOS(5) = "B"
Set loc_pedidos = PSLOC_PEDIDOS.OpenResultset(rdOpenKeyset, rdConcurValues)

With gridf
.Rows = 2
.Cols = 5
.Clear
.TextMatrix(0, 0) = "Nro"
.TextMatrix(0, 1) = "CodVen"
.TextMatrix(0, 2) = "Cliente"
.TextMatrix(0, 3) = "Importe"
.TextMatrix(0, 4) = "Estado"
.ColWidth(0) = 1000
.ColWidth(1) = 900
.ColWidth(2) = 2100
.ColWidth(3) = 1100
.ColWidth(4) = 800
.Visible = False
fila = 0
Do Until loc_pedidos.EOF
  If cheped.Value = 1 Then
    If Trim(loc_pedidos!ped_situacion) <> "" Then GoTo sigue_ped
  End If
  fila = fila + 1
  .Rows = .Rows + 1
 .TextMatrix(.Rows - 1, 0) = Format(loc_pedidos!PED_NUMFAC, "0000000")
 .TextMatrix(.Rows - 1, 1) = Format(loc_pedidos!PED_codven, "000")
 .TextMatrix(.Rows - 1, 2) = Trim(loc_pedidos!PED_NOMCLIE)
 .TextMatrix(.Rows - 1, 3) = Format(Val(loc_pedidos!PED_IGV) + Val(loc_pedidos!PED_BRUTO), "#,##0.00")
 If loc_pedidos!ped_situacion = "E" Then
 .TextMatrix(.Rows - 1, 4) = "Anulado"
 ElseIf loc_pedidos!ped_situacion = "P" Then
 .TextMatrix(.Rows - 1, 4) = "Procesado"
 Else
 .TextMatrix(.Rows - 1, 4) = "Pendiente"
 End If
   
sigue_ped:
   loc_pedidos.MoveNext
Loop
.Visible = True
If .Rows <= 2 Then
  Azul2 fechaf, fechaf
Else
 .TextMatrix(1, 0) = Format(fila, "000") & " PED."
 .SetFocus
End If
End With


End Sub

Private Sub cmdmos_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  frapedido.Visible = False
  Command2.SetFocus
  Exit Sub
End If

End Sub

Private Sub cmdtipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Azul txtcli, txtcli
End If
End Sub

Private Sub cmdtipo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then
    PUB_TIPREG = 65
    PUB_CODCIA = LK_CODCIA
    Load FrmDatArti
    FrmDatArti.Caption = Trim(Left(cmdtipo.Text, 30)) & " TAB_TIPREG = " & PUB_TIPREG
    FrmDatArti.Show 1
    PUB_TIPREG = 65
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cmdtipo.Clear
    Do Until tab_mayor.EOF
       cmdtipo.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
       tab_mayor.MoveNext
    Loop
    If cmdtipo.ListCount > 0 Then cmdtipo.ListIndex = 0
    cmdtipo.SetFocus
    SendKeysSeguro VK_UP, True
    DoEvents
End If

End Sub

Private Sub Command1_Click()
condi.Visible = False
c_condi.SetFocus
End Sub

Private Sub Command2_Click()
frapedido.Visible = True
cheped.Value = 1
fechaf.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
Azul2 fechaf, fechaf

End Sub

Private Sub fechaf_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  frapedido.Visible = False
  Command2.SetFocus
  Exit Sub
End If
If KeyAscii = 13 Then
  cmdmos.SetFocus
End If
End Sub

Private Sub Form_Activate()
If i_condi.ListCount > 0 And i_condi.ListIndex = -1 Then
 i_condi.ListIndex = 0
End If
If moneda.ListCount > 0 And moneda.ListIndex = -1 And f1.Enabled = True Then moneda.ListIndex = 0
If txtcli.Enabled Then Azul txtcli, txtcli
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
  If Left(cmdIngreso.Caption, 2) = "&G" Then
     cmdIngreso_Click
  End If
ElseIf KeyCode = 114 Then
 If i_condi.Enabled Then
   i_condi.SetFocus
   SendKeysSeguro VK_UP, True
  End If
ElseIf KeyCode = 115 Then
  cancelar_Click
ElseIf KeyCode = 116 Then
  If Left(cmdIngreso.Caption, 2) = "&G" Then
  Else
    cmdIngreso_Click
  End If
End If
End Sub

Private Sub Form_Load()
    LogInfo "FORM_COT: Inicio Form_Load"
    CenterMe FORM_COT
'On Error GoTo SALE

Wsec = 0
LOC_CANCELA = 0
fila = 0
wfila_act = 0
WSELE = ""
Dim ws_indice As Integer
Dim cade
WMODO = ""
Dim PSPRO_V As rdoQuery
Dim PRO_V As rdoResultset


pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_NUMSER = ? and PED_NUMFAC = ?  ORDER BY PED_NUMSEC"
Set PSLOC_WARTI = CN.CreateQuery("", pub_cadena)
PSLOC_WARTI(0) = 0
PSLOC_WARTI(1) = 0
PSLOC_WARTI(2) = 0
PSLOC_WARTI(3) = 0
Set llave_sum_arti = PSLOC_WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT PED_NUMFAC FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = 201  ORDER BY  PED_NUMFAC DESC "
Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
PSTEMP_MAYOR(0) = LK_CODCIA
PSTEMP_MAYOR.MaxRows = 1
Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PEDIDOS WHERE  PED_TIPMOV = 201 ORDER BY PED_CODCIA"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
temp_llave.Requery

pub_cadena = "SELECT DISTINCT PED_TIPMOV, PED_NUMSER, PED_NUMFAC, PED_CODCLIE, PED_CODVEN, PED_FECHA , PED_CODUSU, PED_NOMCLIE, PED_IGV, PED_BRUTO FROM " & _
 " PEDIDOS WHERE PED_CODCIA = ? AND PED_FECHA = ? AND PED_TIPMOV = 201 AND PED_CODVEN = ? AND PED_SITUACION <> 'E' "
Set PSDOC_LLAVE = CN.CreateQuery("", pub_cadena)
PSDOC_LLAVE(0) = 0
PSDOC_LLAVE(1) = Date
PSDOC_LLAVE(2) = 0
Set peddoc_llave = PSDOC_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)



PUB_TIPREG = 65
PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_TAB_LLAVE
cmdtipo.Clear
Do Until tab_mayor.EOF
   cmdtipo.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
   tab_mayor.MoveNext
Loop
If cmdtipo.ListCount > 0 Then cmdtipo.ListIndex = 0
   
fila = 0
DoEvents
LIMPIA_DATOS
CABE_MAN
SQ_OPER = 2
PUB_TIPREG = 45
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
fila = 0
Do Until tab_mayor.EOF
PRE_ETIQUETA(fila) = Trim(tab_mayor!tab_nomlargo)
fila = fila + 1
tab_mayor.MoveNext
Loop
cmdimp.Visible = False
txtruc.MaxLength = LK_DIG_RUC

carga_venta

Exit Sub
SALE:
MsgBox "Depurar: " & Err.Description, 48, Pub_Titulo
Resume Next
End Sub



Private Sub forma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tiempo.SetFocus
End If
End Sub

Private Sub GPED_Click()
GPED.Left = 6210
GPED.Width = 5505
End Sub

Private Sub GPED_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  If Txt_key.Enabled Then Azul Txt_key, Txt_key
End If
End Sub

Private Sub GPED_LostFocus()
GPED.Left = 9480
GPED.Width = 2265
End Sub

Private Sub GPED_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
GPED.Left = 9480
GPED.Width = 2265
End Sub

Private Sub grid_fac_EnterCell()
TEXTOVAR.Visible = False
TEXTOVAR.Text = Trim(grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL))
TEXTOVAR.Left = grid_fac.Left + grid_fac.CellLeft
TEXTOVAR.Width = grid_fac.CellWidth
TEXTOVAR.Height = grid_fac.CellHeight
TEXTOVAR.Top = grid_fac.Top + grid_fac.CellTop '- 1560 '480
If grid_fac.COL = 1 Then
 If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
  TEXTOVAR.Visible = True
  TEXTOVAR.SetFocus
 End If
End If
If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) <> "" Then
 If Val(grid_fac.TextMatrix(grid_fac.Row, 12)) <> 0 Then
   stock.Caption = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 15)) / Val(grid_fac.TextMatrix(grid_fac.Row, 12)), "0.00")
 End If
 unid.Caption = grid_fac.TextMatrix(grid_fac.Row, 16)
 nomarti.Caption = grid_fac.TextMatrix(grid_fac.Row, 0)
Else
 stock.Caption = ""
 unid.Caption = ""
 nomarti.Caption = ""
End If

End Sub

Private Sub grid_fac_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Dim wprecios As String * 12
Static CONS
Dim wactivo As Integer
Dim antpre As Integer

If KeyAscii <> 13 Then Exit Sub
If grid_fac.Rows <= 1 Then Exit Sub
'If grid_fac.COL = 1 Then Exit Sub
If grid_fac.COL >= 6 Then Exit Sub

If grid_fac.COL = 2 Then
 If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
    grid_fac.SetFocus
    Exit Sub
 End If
End If
If grid_fac.COL = 3 Then
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
     grid_fac.SetFocus
     Exit Sub
    End If
    UNIDAD.Left = grid_fac.Left + grid_fac.CellLeft
    UNIDAD.Width = grid_fac.CellWidth
    UNIDAD.Top = grid_fac.Top + grid_fac.CellTop
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
    LEER_PRE_LLAVE
    UNIDAD.Clear
    UNIDAD.Visible = True
    wactivo = 0
    Do Until pre_mayor.EOF
     UNIDAD.AddItem Trim(pre_mayor!pre_UNIDAD) & String(30, " ") & pre_mayor!pre_secuencia
     If pre_mayor!PRE_FLAG_UNIDAD <> "A" Then
       wactivo = pre_mayor.AbsolutePosition - 1
     End If
     pre_mayor.MoveNext
    Loop
    On Error GoTo pasa
    UNIDAD.ListIndex = wactivo
    grid_fac.TextMatrix(grid_fac.Row, 4) = ""
    grid_fac.TextMatrix(grid_fac.Row, 13) = wactivo
    On Error GoTo 0
    unidad_KeyPress 13
    ' UNIDAD.Visible = True
    ' UNIDAD.SetFocus
    'SendKeysSeguro VK_UP, True
     Exit Sub
End If
If grid_fac.COL = 4 Then
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
      grid_fac.SetFocus
      Exit Sub
    End If
    antpre = Val(grid_fac.TextMatrix(grid_fac.Row, 14))
    PRECIOS.Left = grid_fac.Left + grid_fac.CellLeft
    PRECIOS.Width = grid_fac.CellWidth + 900
    PRECIOS.Top = grid_fac.Top + grid_fac.CellTop

    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
    PUB_SECUEN = Val(Trim(Right(UNIDAD.Text, 3)))
    grid_fac.TextMatrix(grid_fac.Row, 18) = PUB_SECUEN
    LEER_PRE_LLAVE
    PRECIOS.Clear
    PRECIOS.Visible = True
    Do Until pre_llave.EOF
      If Left(moneda.Text, 1) = "S" Then
          wprecios = Format(pre_llave!PRE_PRE1, "0.0000")
          If Val(grid_fac.TextMatrix(grid_fac.Row, 22)) = 1 Then
             wprecios = "0.00"
             PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          Else
            If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          End If
          wprecios = Format(pre_llave!PRE_PRE2, "0.0000")
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(1), 8) & "= " & wprecios & String(60, " ") & "2"
          wprecios = Format(pre_llave!PRE_PRE3, "0.0000")
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(2), 8) & "= " & wprecios & String(60, " ") & "3"
          wprecios = Format(pre_llave!PRE_PRE4, "0.0000")
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(3), 8) & "= " & wprecios & String(60, " ") & "4"
          wprecios = Format(pre_llave!PRE_PRE5, "0.0000")
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(4), 8) & "= " & wprecios & String(60, " ") & "5"
       Else
          wprecios = pre_llave!pre_pre11
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_llave!PRE_PRE22
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "2"
          wprecios = pre_llave!PRE_PRE33
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "3"
          wprecios = pre_llave!PRE_PRE44
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "4"
          wprecios = pre_llave!PRE_PRE55
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "5"
       End If
     pre_llave.MoveNext
    Loop
    On Error GoTo pasa
    If PRECIOS.ListCount <= 0 Then
      PRECIOS.Visible = False
      MsgBox "Definir precios....", 48, Pub_Titulo
      grid_fac.COL = 1
      grid_fac.SetFocus
      Exit Sub
    End If
    ' antpre
    'For a = 0 To PRECIOS.ListCount - 1
    
    If antpre - 1 > 0 And antpre - 1 < PRECIOS.ListCount Then
     PRECIOS.ListIndex = antpre - 1
    Else
     If PRECIOS.ListCount > 0 Then PRECIOS.ListIndex = 0
    End If
    On Error GoTo 0
    
    PRECIOS.Visible = True
    PRECIOS.SetFocus
    SendKeysSeguro VK_UP, True
     Exit Sub
End If
If grid_fac.COL = 5 Then
 If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
    grid_fac.SetFocus
    Exit Sub
 End If
End If


'    TEXTOVAR.Left = grid_fac.Left + grid_fac.CellLeft
'    TEXTOVAR.Width = grid_fac.CellWidth
'    TEXTOVAR.Height = grid_fac.CellHeight
'    TEXTOVAR.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 1200 '480
'    TEXTOVAR.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
'    wfila_act = grid_fac.Row
    TEXTOVAR.Visible = True
    Azul3 TEXTOVAR, TEXTOVAR
    TEXTOVAR.SetFocus
Exit Sub
pasa:
Resume Next
End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean

If KeyCode = 46 Then
If grid_fac.Rows <= 2 Then Exit Sub
If grid_fac.Rows <= 3 Then
    pub_mensaje = MsgBox("Quitar el Producto para la Orden de Compra ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
    If pub_mensaje = vbNo Then
      grid_fac.SetFocus
      Exit Sub
    End If
    CABE_MAN
Else
   pub_mensaje = MsgBox("Quitar el Producto para la Orden de Compra ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     grid_fac.SetFocus
     Exit Sub
   Else
   '  grid_fac.RowHeight(grid_fac.Row) = 1
   grid_fac.RemoveItem (grid_fac.Row)
   grid_fac.Row = grid_fac.Row
   grid_fac.Refresh
   suma_grid
   grid_fac.SetFocus
   End If
End If
End If
'grid_fac.SetFocus
Exit Sub



End Sub



Private Sub gridf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
  MUESTRA_PED
  frapedido.Visible = False
  cmdimp_Click
  frapedido.Visible = True
  
  gridf.SetFocus
  Exit Sub
End If
If KeyCode = 13 Then
  MUESTRA_PED
  frapedido.Visible = False
End If

End Sub

Private Sub gridf_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  frapedido.Visible = False
  Command2.SetFocus
  Exit Sub
End If

End Sub

Private Sub i_condi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 moneda.SetFocus
 SendKeysSeguro VK_UP, True
End If

End Sub

Private Sub i_condi_LostFocus()
PUB_CODTRA = 2401
PUB_SECUENCIA = Trim(Left(i_condi.Text, 2))
SQ_OPER = 1
LEER_SUT_LLAVE
pub_signo_car = SUT_LLAVE!SUT_SIGNO_CAR
End Sub

Private Sub i_destino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 i_fbg.SetFocus
 SendKeysSeguro VK_UP, True
End If

End Sub

Private Sub i_dias_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If UCase(cmdIngreso.Caption) <> "&GRABAR" Then
   If i_destino.Enabled = True Then
     i_destino.SetFocus
     SendKeysSeguro VK_UP, True
   End If
End If
End If
End Sub

Private Sub i_fbg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 grid_fac_EnterCell
 TEXTOVAR.Visible = False
 
 If Val(Left(i_condi.Text, 2)) = 1 Then
    pub_mensaje = "La Condicion es Contado Cash. !!! ...   ¿Desea Continuar... ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       i_condi.SetFocus
       Exit Sub
    End If
 End If
 grid_fac.SetFocus
End If

End Sub

Private Sub ListView1_DblClick()
 loc_key = ListView1.SelectedItem.Index
 TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 textovar_KeyPress 13
End Sub

Private Sub ListView1_GotFocus()
'If loc_key <> 0 Then
' Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
' ListView1.ListItems.Item(loc_key).Selected = True
' ListView1.ListItems.Item(loc_key).EnsureVisible
'End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
'If loc_key <> 0 Then
' loc_key = ListView1.SelectedItem.Index
' TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
'End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView1.Visible = False
 TEXTOVAR.Text = ""
 TEXTOVAR.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
ListView1_DblClick
End Sub

Private Sub ListView1_LostFocus()
ListView1.Visible = False
End Sub

Private Sub moneda_Click()
If Not cmdIngreso.Enabled Then Exit Sub
For fila = 2 To grid_fac.Rows - 1
     PUB_CODART = Val(grid_fac.TextMatrix(fila, 10))
     If PUB_CODART > 0 Then
       pu_codcia = LK_CODCIA
       PUB_SECUENCIA = Val(grid_fac.TextMatrix(fila, 11))
       SQ_OPER = 1
       LEER_PRE_LLAVE
       If Left(moneda.Text, 1) = "S" Then
          If Val(grid_fac.TextMatrix(fila, 14)) = 1 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE1)
          If Val(grid_fac.TextMatrix(fila, 14)) = 2 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE2)
          If Val(grid_fac.TextMatrix(fila, 14)) = 3 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE3)
          If Val(grid_fac.TextMatrix(fila, 14)) = 4 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE4)
          If Val(grid_fac.TextMatrix(fila, 14)) = 5 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE5)
       Else
          If Val(grid_fac.TextMatrix(fila, 14)) = 1 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!pre_pre11)
          If Val(grid_fac.TextMatrix(fila, 14)) = 2 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE22)
          If Val(grid_fac.TextMatrix(fila, 14)) = 3 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE33)
          If Val(grid_fac.TextMatrix(fila, 14)) = 4 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE44)
          If Val(grid_fac.TextMatrix(fila, 14)) = 5 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE55)
       End If
       grid_fac.TextMatrix(fila, 4) = redondea(Val(grid_fac.TextMatrix(fila, 13)) * (100 - Val(grid_fac.TextMatrix(fila, 5))) / 100)
     End If
Next fila
If grid_fac.Rows <> 2 Then suma_grid


If Left(moneda.Text, 1) = "S" Then
 i_moneda.Caption = "S/."
 grid_fac.TextMatrix(1, 4) = "S/."
Else
 i_moneda.Caption = "US$."
 grid_fac.TextMatrix(1, 4) = "US$."
End If
End Sub

Private Sub moneda_GotFocus()
If moneda.ListCount = 1 Then moneda_KeyPress 13
End Sub

Private Sub moneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul Txt_key, Txt_key
End If
End Sub


Private Sub PRECIOS_GotFocus()
On Error GoTo SALE
'grid_fac.TextMatrix(grid_fac.Row, 13) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.00")

'grid_fac.TextMatrix(grid_fac.Row, 14) = Val(Right(PRECIOS.Text, 6))
SALE:
Exit Sub
End Sub

Private Sub PRECIOS_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
  PRECIOS.Visible = False
  grid_fac.SetFocus
  grid_fac.COL = grid_fac.COL - 1
  grid_fac.SetFocus
End If

End Sub

Private Sub PRECIOS_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim wpor As Currency

If KeyAscii = 27 Then
 PRECIOS.Visible = False
 grid_fac.SetFocus
End If
If KeyAscii <> 13 Then Exit Sub
'SQ_OPER = 1
'pu_codcia = LK_CODCIA
'PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
'PUB_SECUEN = Val(Right(UNIDAD.Text, 4))
'LEER_PRE_LLAVE

grid_fac.TextMatrix(grid_fac.Row, 4) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.0000")
grid_fac.TextMatrix(grid_fac.Row, 13) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.0000")
grid_fac.TextMatrix(grid_fac.Row, 14) = Val(Right(PRECIOS.Text, 6))


pre_llave.Requery
wpor = 0
If pre_llave.EOF Then GoTo SALE

If Val(Right(PRECIOS.Text, 3)) = 1 Then
    wpor = pre_llave!PRE_POR1
ElseIf Val(Right(PRECIOS.Text, 3)) = 2 Then
    wpor = pre_llave!PRE_POR2
ElseIf Val(Right(PRECIOS.Text, 3)) = 3 Then
    wpor = pre_llave!PRE_POR3
ElseIf Val(Right(PRECIOS.Text, 3)) = 4 Then
    wpor = pre_llave!PRE_POR4
ElseIf Val(Right(PRECIOS.Text, 3)) = 5 Then
    wpor = pre_llave!PRE_POR5
End If
grid_fac.TextMatrix(grid_fac.Row, 20) = wpor ' SOLO DESCTO DEL NUMERO DE PRECIO
If wpor = 0 Then
 grid_fac.TextMatrix(grid_fac.Row, 21) = ""
Else
 grid_fac.TextMatrix(grid_fac.Row, 21) = wpor ' mUESTRA LOS PORC,
End If


SALE:
PRECIOS.Visible = False
suma_grid
TEXTOVAR.Visible = False
If Trim(grid_fac.TextMatrix(grid_fac.Rows - 1, 1)) <> "" Then
  grid_fac.Rows = grid_fac.Rows + 1
  grid_fac.RowHeight(grid_fac.Rows - 1) = 285
  grid_fac.Row = grid_fac.Rows - 1
Else
 If grid_fac.Row < grid_fac.Rows - 1 Then
    grid_fac.Row = grid_fac.Row + 1
 End If
End If
grid_fac.COL = 1
If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
 TEXTOVAR.Visible = True
 TEXTOVAR.SetFocus
Else
  grid_fac.SetFocus
End If


Exit Sub
grid_fac.COL = 5
grid_fac_KeyPress 13


End Sub

Private Sub PRECIOS_KeyUp(KeyCode As Integer, Shift As Integer)
Dim ww As String
Dim wpre As Currency
If KeyCode = 45 Then
'seIf grid_fac.COL = 6 Then
ww = InputBox("Digite Precios :", "Ingreso de Precio", "0")
wpre = Val(ww)
grid_fac.TextMatrix(grid_fac.Row, 4) = wpre
If Val(grid_fac.TextMatrix(grid_fac.Row, 13)) <> 0 Then
  grid_fac.TextMatrix(grid_fac.Row, 5) = redondea((Val(grid_fac.TextMatrix(grid_fac.Row, 13)) - wpre) * 100 / Val(Val(grid_fac.TextMatrix(grid_fac.Row, 13))))
End If
PRECIOS.Visible = False
suma_grid

grid_fac.SetFocus






End If
End Sub

Private Sub PRECIOS_LostFocus()
PRECIOS.Visible = False
End Sub

Private Sub salir_Click()
Unload FORM_COT
End Sub


Public Sub LIMPIA_DATOS()
i_dias.Text = ""
lblvisita.Caption = ""
lbluser.Caption = ""
LBLSIT.Caption = ""
grid_fac.Enabled = True
lblcli.Caption = ""
txtatte.Text = ""

txtcli.Text = ""
txtruc.Text = ""
grid_fac.Clear

txtigv.Text = ""
txtvalorv.Text = ""
txttotal.Text = ""
TEXTOVAR.Visible = False
stock.Caption = ""
unid.Caption = ""
nomarti.Caption = ""
oferta.Text = ""
forma.Text = ""
tiempo.Text = ""
i_destino.Clear
'tserie.Text = ""
'txtdoc.Text = ""
End Sub

Public Sub CABE_MAN()
grid_fac.Cols = 23
grid_fac.Rows = 2
grid_fac.Clear
fila = 0
grid_fac.ColWidth(0) = 2400 ' nombre arti
grid_fac.ColWidth(1) = 900 ' codigo arti
grid_fac.ColWidth(2) = 800 ' cantidad
grid_fac.ColWidth(3) = 800 ' unidad
grid_fac.ColWidth(4) = 800 ' precio
grid_fac.ColWidth(5) = 700 ' decto.
grid_fac.ColWidth(6) = 900 ' sub total
grid_fac.ColWidth(7) = 700 ' peso
grid_fac.ColWidth(8) = 0
grid_fac.ColWidth(9) = 0
grid_fac.ColWidth(10) = 0 '  COD ORIGINAL
grid_fac.ColWidth(11) = 0 '  PRE_SECUENCIA
grid_fac.ColWidth(12) = 0 '  PRE_EQUIV
grid_fac.ColWidth(13) = 0 '  PRE_PRECIO COLOCADO
grid_fac.ColWidth(14) = 0 '  numero de PRE_PRECIO
grid_fac.ColWidth(15) = 0 '  numero de arm_stock
grid_fac.ColWidth(16) = 0 '  numero de pre_unidad
grid_fac.ColWidth(17) = 0
grid_fac.ColWidth(18) = 0 ' NUMERO DE SECUENCIA EN PRECIOS
grid_fac.ColWidth(19) = 0 ' flag para limite de credito
grid_fac.ColWidth(20) = 0 ' el Descto por el Precio de Venta
grid_fac.ColWidth(21) = 900 ' Total de Descuento en (%)
grid_fac.ColWidth(22) = 0 ' Flag Identifica el DEscuento de BoNifi(%)

grid_fac.TextMatrix(0, 0) = "Articulo"
grid_fac.TextMatrix(0, 1) = "Codigo"
grid_fac.TextMatrix(0, 2) = "Cantidad"
grid_fac.TextMatrix(0, 3) = "Unidad"
grid_fac.TextMatrix(0, 4) = "Precios"
grid_fac.TextMatrix(0, 5) = "Dscto"
grid_fac.TextMatrix(1, 5) = "  (%)"
grid_fac.TextMatrix(0, 6) = "Sub Total"
grid_fac.TextMatrix(0, 7) = "Peso(Kg)"
grid_fac.TextMatrix(0, 8) = ""
grid_fac.TextMatrix(0, 9) = ""
grid_fac.TextMatrix(0, 21) = "(%)Desct."
grid_fac.RowHeight(1) = 320

End Sub
Public Sub suma_grid()
'On Error GoTo SALE
Dim WC_Item As Integer
Dim RES_DEUDA As Currency
Dim wsumadol As Currency
Dim WSIMBOL As String
Dim WTC As Currency
Dim WF As Integer
WF = 2
Dim fx As Integer
Dim wcantid As Currency
Dim wpeso As Currency
fx = 1
SUM_H = 0
SUM_D = 0
wcantid = 0
WC_Item = 0
Do While fx = 1
    'If Left(grid_fac.TextMatrix(WF, 0), 1) <> "T" Then
      If Val(grid_fac.TextMatrix(WF, 1)) <> 0 Then
      WC_Item = WC_Item + 1
      End If
      SUM_D = SUM_D + Val(grid_fac.TextMatrix(WF, 4))
      SUM_H = SUM_H + Val(Val(grid_fac.TextMatrix(WF, 2)) * Val(grid_fac.TextMatrix(WF, 4)))
      wcantid = wcantid + Val(grid_fac.TextMatrix(WF, 2))
      wpeso = wpeso + Val(grid_fac.TextMatrix(WF, 7))
      grid_fac.TextMatrix(WF, 6) = Format(Val(grid_fac.TextMatrix(WF, 2)) * Val(grid_fac.TextMatrix(WF, 4)), "0.00")
    'End If
    WF = WF + 1
    If WF = grid_fac.Rows Then
        fx = 0
    Else
        If Trim(grid_fac.TextMatrix(WF, 0)) = "" Then fx = 0
    End If
Loop
   fila = WF - 1
   grid_fac.TextMatrix(1, 0) = "Totales =   " & Format(WC_Item, "00") & " item(s)"
   
   grid_fac.TextMatrix(1, 6) = Format(SUM_H, "####0.00")
   grid_fac.TextMatrix(1, 2) = Format(wcantid, "####0.00")
   grid_fac.TextMatrix(1, 7) = Format(wpeso, "####0.00")
   
   txttotal.Text = Format(SUM_H, "#####0.00")
   WS_NETO = SUM_H
   txtvalorv.Text = Format((SUM_H / ((100 + LK_IGV) / 100)), "####0.00")
   txtigv.Text = Format(SUM_H - Val(txtvalorv.Text), "#####0.00")

If cli_llave.EOF Then Exit Sub
If SUT_LLAVE.EOF Then Exit Sub
If SUT_LLAVE!SUT_SIGNO_CAR <> 1 Then Exit Sub
If Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 1 Then Exit Sub
WTC = 1
If Left(moneda.Text, 1) = "D" Then
 PUB_CAL_INI = LK_FECHA_DIA
 PUB_CAL_FIN = LK_FECHA_DIA
 pu_codcia = LK_CODCIA
 PUB_CODCIA = LK_CODCIA
 SQ_OPER = 1
 LEER_CAL_LLAVE
 WTC = 0
 WSIMBOL = ""
 If Not cal_llave.EOF Then
   WTC = Nulo_Valor0(cal_llave!cal_tipo_cambio)
 End If
 If WTC = 0 Then
   MsgBox "Ingresar el Tipo de Cambio", 48, Pub_Titulo
   GoTo PA
 End If
End If
If Left(moneda.Text, 1) = "S" Then
  WSIMBOL = "S/."
  wsumadol = Val(Nulo_Valor0(cli_llave!cli_limcre))
  RES_DEUDA = pub_deuda
Else
 WSIMBOL = "US$."
 wsumadol = Val(redondea(Nulo_Valor0(cli_llave!cli_limcre) / WTC)) + Val(redondea(Val(Nulo_Valor0(cli_llave!cli_limcre2))))
 RES_DEUDA = redondea(Val(pub_deuda / WTC))
End If
lblcred.Caption = Format(Nulo_Valor0(cli_llave!cli_limcre), "#,##0.00")
lblDeuda.Caption = Format(RES_DEUDA + WS_NETO, "#,##0.00")
lbldisp.Caption = Format(wsumadol - RES_DEUDA - WS_NETO, "##,###,###.00")

PA:

If LK_FLAG_EXED = "A" And grid_fac.TextMatrix(grid_fac.Row, 19) <> "A" Then
   If Trim(Left(moneda.Text, 1)) = "S" Then
     If (RES_DEUDA + WS_NETO) > Nulo_Valor0(cli_llave!cli_limcre) Then
         MsgBox "El Monto del Pedido supera el Credito... ", 48, Pub_Titulo
         grid_fac.TextMatrix(grid_fac.Row, 19) = "A"
         If TEXTOVAR.Visible Then TEXTOVAR.SetFocus
     End If
   Else
     If (RES_DEUDA + WS_NETO) > Nulo_Valor0(cli_llave!cli_limcre2) Then
         'If ws_bruto_bak <> WS_BRUTO Then MsgBox "Credito Excedido... "
      End If
   End If
End If
 
  
Exit Sub
SALE:
cancelar_Click
'MsgBox "Verficar Importe.", 48, Pub_Titulo
'Resume Next
'If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR
End Sub
Public Sub suma_subtotal()
If WMODO = "I" Then Exit Sub

Dim WF As Integer
Dim WFIN As Integer
Dim WINI As Integer

Dim fx As Integer
Exit Sub
End Sub

Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As RichTextBox, wsKeyAscii As Integer)
  Static VALOR
  Dim car As String
 ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If

End Sub

Public Sub CABE_ING()
grid_fac.Cols = 7
grid_fac.Rows = 3
grid_fac.Clear
grid_fac.MergeCells = 4
grid_fac.MergeCol(0) = True
grid_fac.MergeCol(1) = True
grid_fac.MergeCol(2) = True
grid_fac.MergeCol(3) = True
grid_fac.MergeCol(4) = False
grid_fac.MergeCol(5) = False
grid_fac.MergeRow(2) = False
grid_fac.RowHeight(0) = 285
grid_fac.RowHeight(1) = 285
grid_fac.RowHeight(2) = 285

fila = 0
grid_fac.ColWidth(0) = 400
grid_fac.ColWidth(1) = 1400
grid_fac.ColWidth(2) = 2500
grid_fac.ColWidth(3) = 0
grid_fac.ColWidth(4) = 1500
grid_fac.ColWidth(5) = 1500
grid_fac.ColWidth(6) = 0

grid_fac.TextMatrix(0, 0) = "Item"
grid_fac.TextMatrix(0, 1) = "Cuenta"
grid_fac.TextMatrix(0, 2) = "Descripcion"
grid_fac.TextMatrix(0, 3) = "Glosa"
grid_fac.TextMatrix(0, 4) = "Debe"
grid_fac.TextMatrix(0, 5) = "Haber"
grid_fac.TextMatrix(1, 0) = "Item"
grid_fac.TextMatrix(1, 1) = "Cuenta"
grid_fac.TextMatrix(1, 2) = "Descripcion"
grid_fac.TextMatrix(1, 3) = "Glosa"

'grid_fac.MergeCol
'grid_fac.MergeRow(2) = True



End Sub

Private Sub siguiente_Click()
 txtdoc.Text = Val(txtdoc.Text) + 1
 PUB_NUMSER = Val(tserie.Text)
 PUB_NUMFAC = Val(txtdoc.Text)
 LLENA_DOCU

End Sub

Private Sub textovar_Change()
If Not TEXTOVAR.Visible Then Exit Sub
If grid_fac.COL = 5 Then
 grid_fac.TextMatrix(grid_fac.Row, 4) = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 13)) * (100 - Val(TEXTOVAR.Text)) / 100)
 If Val(TEXTOVAR.Text) = 0 Then
   grid_fac.TextMatrix(grid_fac.Row, 21) = Trim(grid_fac.TextMatrix(grid_fac.Row, 20))
 Else
   grid_fac.TextMatrix(grid_fac.Row, 21) = Trim(grid_fac.TextMatrix(grid_fac.Row, 20)) & "+" & Trim(TEXTOVAR.Text)
 End If
End If
If grid_fac.COL = 1 Then
   If Trim(i_fbg.Text) = "F" Then
     If Nulo_Valor0(par_llave!par_fac_lines) < grid_fac.Row - 1 Then
          TEXTOVAR.Text = ""
          MsgBox "Tope de Lineas Alcanzado...", 48, Pub_Titulo
          TEXTOVAR.SetFocus
          Exit Sub
     End If
   ElseIf Trim(i_fbg.Text) = "B" Then
      If Nulo_Valor0(par_llave!par_BOL_lines) < grid_fac.Row - 1 Then
          TEXTOVAR.Text = ""
         MsgBox "Tope de Lineas Alcanzado...", 48, Pub_Titulo
         TEXTOVAR.SetFocus
         Exit Sub
      End If
   End If
linp:
    grid_fac.TextMatrix(grid_fac.Row, 0) = ""
    grid_fac.TextMatrix(grid_fac.Row, 0) = ""
    grid_fac.TextMatrix(grid_fac.Row, 1) = ""
    grid_fac.TextMatrix(grid_fac.Row, 2) = ""
    grid_fac.TextMatrix(grid_fac.Row, 3) = ""
    grid_fac.TextMatrix(grid_fac.Row, 4) = ""
    grid_fac.TextMatrix(grid_fac.Row, 5) = ""
    grid_fac.TextMatrix(grid_fac.Row, 6) = ""
    grid_fac.TextMatrix(grid_fac.Row, 7) = ""
    grid_fac.TextMatrix(grid_fac.Row, 8) = ""
    grid_fac.TextMatrix(grid_fac.Row, 9) = ""
    grid_fac.TextMatrix(grid_fac.Row, 10) = ""
    grid_fac.TextMatrix(grid_fac.Row, 11) = ""
    grid_fac.TextMatrix(grid_fac.Row, 12) = ""
    grid_fac.TextMatrix(grid_fac.Row, 14) = ""
    grid_fac.TextMatrix(grid_fac.Row, 15) = ""
    grid_fac.TextMatrix(grid_fac.Row, 16) = ""
    grid_fac.TextMatrix(grid_fac.Row, 17) = ""
    grid_fac.Text = TEXTOVAR.Text
    stock.Caption = ""
    unid.Caption = ""
    nomarti.Caption = ""
    suma_grid
Else
 If grid_fac.COL = 2 Then
  grid_fac.Text = TEXTOVAR.Text
 Else
  grid_fac.Text = Format(TEXTOVAR.Text, "0.00")
 End If
 suma_grid
 suma_subtotal
End If
End Sub

Private Sub TEXTOVAR_GotFocus()
'temporal = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
End Sub

Private Sub textovar_KeyDown(KeyCode As Integer, Shift As Integer)

' busca arti
If Not ListView1.Visible Then
If KeyCode = 40 Then  ' flecha abajo
  If grid_fac.Row = grid_fac.Rows - 1 Then Exit Sub
  If Trim(grid_fac.Text) <> "" Then Exit Sub
  grid_fac.Row = grid_fac.Row + 1
  grid_fac.SetFocus
  Exit Sub
End If
If KeyCode = 38 Then
 If Trim(grid_fac.Text) <> "" Then Exit Sub
 grid_fac.Row = grid_fac.Row - 1
 grid_fac.SetFocus
 Exit Sub
End If
If KeyCode = 39 Then
If Trim(grid_fac.Text) <> "" Then Exit Sub
 grid_fac.COL = grid_fac.COL + 1
 grid_fac.SetFocus
 Exit Sub
End If
If KeyCode = 37 Then
 If Trim(grid_fac.Text) <> "" Then Exit Sub
 grid_fac.COL = grid_fac.COL - 1
 grid_fac.SetFocus
 Exit Sub
End If

End If

If grid_fac.COL <> 1 Then Exit Sub

Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And TEXTOVAR.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
'  LISTVIEW1.Visible = False
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
'  KeyCode = 0
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  TEXTOVAR.SelStart = Len(TEXTOVAR.Text)
  DoEvents
fin:

End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
'SOLO_DECIMAL TEXTOVAR, KeyAscii
If KeyAscii = 27 Then
  If TEXTOVAR.Text = "" Then
    TEXTOVAR.Visible = False
    grid_fac.SetFocus
    Exit Sub
  End If
  TEXTOVAR.Text = ""
  ListView1.Visible = False
  If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
   'MDIForm1.Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda nombre de articulo."
   'LK_FLAG_ALTERNO = " "
   'MDIForm1.Toolbar1.Buttons.Item(12).Image = 18 '********** tenr en cuent
   'DoEvents
  Else
    MDIForm1.Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda sub-codigo de articulo."
    LK_FLAG_ALTERNO = "A"
    MDIForm1.Toolbar1.Buttons.Item(12).Image = 5
  End If
  Exit Sub
End If
If grid_fac.COL = 2 Or grid_fac.COL = 4 Or grid_fac.COL = 5 Then Consistencias grid_fac, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then Exit Sub

If grid_fac.COL = 2 Then
 If Trim(TEXTOVAR.Text) = "" Then Exit Sub
 TEXTOVAR.Visible = False
 grid_fac.COL = 3
 If Trim(grid_fac.Text) <> "" Then
   grid_fac.SetFocus
   Exit Sub
 End If
 grid_fac_KeyPress 13
 Exit Sub
End If
If grid_fac.COL = 6 Then
' grid_fac.TextMatrix(grid_fac.Row, 6) = textovar.Text
' suma_grid
' textovar.Visible = False
' grid_fac.SetFocus
' Exit Sub
End If
If grid_fac.COL = 5 Then
 TEXTOVAR.Visible = False
 If Trim(grid_fac.TextMatrix(grid_fac.Rows - 1, 1)) <> "" Then
   grid_fac.Rows = grid_fac.Rows + 1
   grid_fac.RowHeight(grid_fac.Rows - 1) = 285
   grid_fac.Row = grid_fac.Rows - 1
 Else
  If grid_fac.Row < grid_fac.Rows - 1 Then
     grid_fac.Row = grid_fac.Row + 1
  End If
 End If
 grid_fac.COL = 1
 TEXTOVAR.Visible = True
 TEXTOVAR.SetFocus
 Exit Sub
End If




If grid_fac.COL <> 1 Then Exit Sub

Dim VALOR As String
Dim tf As Integer
Dim I, car
Dim itmFound As MSComctlLib.ListItem
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView1.Visible = False
 TEXTOVAR.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" And Val(TEXTOVAR.Text) = 0 Then
  PUB_KEY = 0
Else
  GoTo IR_ALTERNO
 On Error GoTo mucho
 PUB_KEY = Val(TEXTOVAR.Text)
 On Error GoTo 0
 If Len(TEXTOVAR.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(TEXTOVAR.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = TEXTOVAR.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul3 TEXTOVAR, TEXTOVAR
       GoTo fin
    End If
    If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul3 TEXTOVAR, TEXTOVAR
       GoTo fin
    End If
    If VERIFICA_REPET = 1 Then
      TEXTOVAR.Text = ""
      grid_fac.TextMatrix(grid_fac.Row, 16) = ""
      grid_fac.TextMatrix(grid_fac.Row, 15) = ""
      grid_fac.TextMatrix(grid_fac.Row, 12) = ""
      grid_fac.TextMatrix(grid_fac.Row, 11) = ""
      grid_fac.TextMatrix(grid_fac.Row, 0) = ""
      grid_fac.TextMatrix(grid_fac.Row, 10) = ""
      MsgBox "Producto esta Repetido en la Lista.", 48, Pub_Titulo
      Azul3 TEXTOVAR, TEXTOVAR
      GoTo fin
    End If

    WCOD_ORIGINAL = art_LLAVE!ART_KEY
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = WCOD_ORIGINAL
    LEER_ARM_LLAVE
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_SECUEN = 0
    LEER_PRE_LLAVE
    grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_UNIDAD
    grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!arm_stock
    grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!PRE_EQUIV
    grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
    grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE!art_nombre
    grid_fac.TextMatrix(grid_fac.Row, 10) = art_LLAVE!ART_KEY
    grid_fac.TextMatrix(grid_fac.Row, 22) = art_LLAVE!ART_orden
    ListView1.Visible = False
    TEXTOVAR.Visible = False
    grid_fac.COL = 2
    If Trim(grid_fac.Text) <> "" Then
      grid_fac.SetFocus
      Exit Sub
    End If
    TEXTOVAR.Visible = True
    TEXTOVAR.SetFocus
    Exit Sub
Else
  If ListView1.Visible = False And VAR_ACTIVAR <> 99 And TEXTOVAR.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = TEXTOVAR.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul3 TEXTOVAR, TEXTOVAR
       Exit Sub
     End If
     If art_llave_alt!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul3 TEXTOVAR, TEXTOVAR
       GoTo fin
     End If
     SQ_OPER = 1
     PUB_KEY = art_llave_alt!ART_KEY
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul3 TEXTOVAR, TEXTOVAR
       GoTo fin
     End If
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_CODART = WCOD_ORIGINAL
     LEER_ARM_LLAVE
     If Not arm_llave.EOF Then
     If Val(arm_llave!arm_stock) <= 0 Then
        MsgBox "No hay Stock en Producto.", 48, Pub_Titulo
        ListView1.Visible = False
        GoTo IR_arti_can
     End If
     
     End If
    
     If VERIFICA_REPET = 1 Then
      MsgBox "Producto esta Repetido en la Lista.", 48, Pub_Titulo
IR_arti_can:
      grid_fac.TextMatrix(grid_fac.Row, 16) = ""
      grid_fac.TextMatrix(grid_fac.Row, 15) = ""
      grid_fac.TextMatrix(grid_fac.Row, 12) = ""
      grid_fac.TextMatrix(grid_fac.Row, 11) = ""
      grid_fac.TextMatrix(grid_fac.Row, 0) = ""
      grid_fac.TextMatrix(grid_fac.Row, 10) = ""
      ListView1.Visible = False
      If ListView1.Visible Then
        Azul3 TEXTOVAR, TEXTOVAR
        Exit Sub
      Else
        TEXTOVAR.Text = ""
        Azul3 TEXTOVAR, TEXTOVAR
      End If
      GoTo fin
     End If
 
     ListView1.Visible = False
     
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_SECUEN = 0
     LEER_PRE_LLAVE
     grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_UNIDAD
     grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!arm_stock
     grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!PRE_EQUIV
     grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
    
     grid_fac.TextMatrix(grid_fac.Row, 0) = art_llave_alt!art_nombre
     grid_fac.TextMatrix(grid_fac.Row, 10) = art_llave_alt!ART_KEY
     grid_fac.TextMatrix(grid_fac.Row, 22) = art_LLAVE!ART_orden
     TEXTOVAR.Visible = False
     ListView1.Visible = False
     grid_fac.COL = 2
     If Trim(grid_fac.Text) <> "" Then
       grid_fac.SetFocus
       Exit Sub
     End If
     ' ************
     ' REGRESA A CODIGO ALTERNO
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        'MDIForm1.Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda nombre de articulo."
        'LK_FLAG_ALTERNO = " "
        'MDIForm1.Toolbar1.Buttons.Item(12).Image = 18 '********** tenr en cuent
      'DoEvents
      Else
        MDIForm1.Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda sub-codigo de articulo."
        LK_FLAG_ALTERNO = "A"
        MDIForm1.Toolbar1.Buttons.Item(12).Image = 5
      End If
    

     TEXTOVAR.Visible = True
     Azul3 TEXTOVAR, TEXTOVAR
     Exit Sub
  Else
    If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
    If Trim(UCase(TEXTOVAR.Text)) = Left(VALOR, Len(Trim(TEXTOVAR.Text))) And Len(Trim(TEXTOVAR.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" And Val(TEXTOVAR.Text) <> 0 Then
        TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       'TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
       TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(3))
       GoTo IR_ALTERNO
      Else
       TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       PUB_KEY = TEXTOVAR.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul3 TEXTOVAR, TEXTOVAR
        Exit Sub
      End If
      If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul3 TEXTOVAR, TEXTOVAR
       GoTo fin
      End If
      WCOD_ORIGINAL = art_LLAVE!ART_KEY
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      PUB_CODART = WCOD_ORIGINAL
      LEER_ARM_LLAVE
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      PUB_SECUEN = 0
      LEER_PRE_LLAVE
      grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_UNIDAD
      grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!arm_stock
      grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!PRE_EQUIV
      grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
      ListView1.Visible = False
      grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE!art_nombre
      grid_fac.TextMatrix(grid_fac.Row, 10) = art_LLAVE!ART_KEY
      grid_fac.TextMatrix(grid_fac.Row, 22) = art_LLAVE!ART_orden
      grid_fac.COL = 2
      If Trim(grid_fac.Text) <> "" Then
        grid_fac.SetFocus
        Exit Sub
      End If
      TEXTOVAR.Visible = True
      TEXTOVAR.SetFocus
     
      Exit Sub
    Else
      Exit Sub
    End If
    
  End If
End If
dale:
ListView1.Visible = False
fin:
mucho:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul3 TEXTOVAR, TEXTOVAR

Exit Sub
End Sub

Private Sub textovar_KeyUp(KeyCode As Integer, Shift As Integer)
If grid_fac.COL <> 1 Then Exit Sub
' busca arti
Dim var
If KeyCode = 13 Then Exit Sub
If Val(TEXTOVAR.Text) = 0 Then
  GoTo BNombre
Else
  GoTo BCodigo
End If
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
BCodigo:
  If Len(TEXTOVAR.Text) = 0 Or Trim(TEXTOVAR.Text) = "" Then
    ListView1.Visible = False
    Exit Sub
  End If
  If TEXTOVAR.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf TEXTOVAR.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(TEXTOVAR.Text, 1) = "*" Then
   TEXTOVAR.Text = Mid(TEXTOVAR.Text, 2, Len(TEXTOVAR.Text))
   TEXTOVAR.SelStart = Len(TEXTOVAR.Text)
  End If
Else
BNombre:
 If Len(TEXTOVAR.Text) = 0 Or IsNumeric(TEXTOVAR.Text) = True Then
   ListView1.Visible = False
   Exit Sub
 End If
End If
If Val(TEXTOVAR.Text) = 0 And ListView1.Visible = False And KeyCode <> 13 Then 'Or Len(TEXTOVAR.Text) = 1 Then
    var = Asc(TEXTOVAR.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If

    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" And Val(TEXTOVAR.Text) <> 0 Then
      numarchi = 3
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND  ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_ALTERNO BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND  (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
    End If
   ' If Len(TEXTOVAR.text) > 1 And ListView1.ListItems.count = 0 Then
   ' Else
     PROC_LISVIEW ListView1
     loc_key = 1
   ' End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(TEXTOVAR.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If




End Sub

Private Sub textovar_LostFocus()
'TEXTOVAR.Visible = False
'If TEXTOVAR.Visible Then
'   TEXTOVAR.Visible = False
'   grid_fac.Row = wfila_act
'   grid_fac.SetFocus
   Exit Sub
'End If

End Sub

Public Sub LLENADOS(cont As ListBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    cont.AddItem " "
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Private Sub tiempo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   oferta.SetFocus
End If

End Sub

Private Sub Txt_key_Change()
If Txt_key.Text = "" Then
 GPED.Rows = 2
 lblven.Caption = ""
End If
End Sub

Private Sub txt_key_GotFocus()
 Azul Txt_key, Txt_key
End Sub
Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not ListView2.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And Txt_key.Text = "" Then
  loc_key = 1
  Set ListView2.SelectedItem = ListView2.ListItems(loc_key)
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  Txt_key.Text = Trim(ListView2.ListItems.Item(loc_key).Text) & " "
  DoEvents
  Txt_key.SelStart = Len(Txt_key.Text)
  DoEvents
fin:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim WMO As String
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As MSComctlLib.ListItem
If KeyAscii = 27 Then
 Txt_key.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
pu_codclie = Val(Txt_key.Text)
If Len(Txt_key.Text) = 0 Or Txt_key.Locked Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(Txt_key.Text) = True Then
   loc_key = 0
   On Error GoTo mucho
   PUB_CODVEN = Val(Txt_key.Text)
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_VEN_LLAVE
   On Error GoTo 0
   If ven_llave.EOF Then
     Azul Txt_key, Txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Txt_key.SetFocus
     GoTo fin
   End If
   lblven.Caption = Trim(ven_llave!VEM_NOMBRE)
   ListView1.Visible = False
   PUB_CODCIA = LK_CODCIA
   PUB_CODVEN = Val(Txt_key.Text)
   WMO = PED_PENDIENTES()
   cmdtipo.SetFocus
   SendKeysSeguro VK_UP, True
   Screen.MousePointer = 0
   Exit Sub
Else
   If loc_key > ListView2.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   VALOR = UCase(ListView2.ListItems.Item(loc_key).Text)
   If Trim(UCase(Txt_key.Text)) = Left(VALOR, Len(Trim(Txt_key.Text))) Then
   Else
      Exit Sub
   End If
   Txt_key.Text = Trim(ListView2.ListItems.Item(loc_key).SubItems(1))
   PUB_CODVEN = Val(Txt_key.Text)
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_VEN_LLAVE
   On Error GoTo 0
   If ven_llave.EOF Then
     Azul Txt_key, Txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Txt_key.SetFocus
     GoTo fin
   End If
   lblven.Caption = Trim(ven_llave!VEM_NOMBRE)
   PUB_CODCIA = LK_CODCIA
   PUB_CODVEN = Val(Txt_key.Text)
   WMO = PED_PENDIENTES()
   ListView2.Visible = False
   cmdtipo.SetFocus
   SendKeysSeguro VK_UP, True
End If
dale:
mucho:
ListView1.Visible = False
fin:
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(Txt_key.Text) = 0 Or Txt_key.Locked = True Or IsNumeric(Txt_key.Text) = True Then
   ListView2.Visible = False
   Exit Sub
End If
If ListView2.Visible = False And KeyCode <> 13 Or Len(Txt_key.Text) = 1 Then
    var = Asc(Txt_key.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 9
    archi = "SELECT * FROM VEMAEST WHERE  VEM_CODCIA = '" & LK_CODCIA & "' AND VEM_NOMBRE BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY VEM_NOMBRE"
    PROC_LISVIEW ListView2
    loc_key = 1
    If ListView2.Visible = False Then
        loc_key = 0
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If ListView2.Visible Then
  Set itmFound = ListView2.FindItem(LTrim(Txt_key.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView2.ListItems.count Then
      ListView2.ListItems.Item(ListView2.ListItems.count).EnsureVisible
   Else
     ListView2.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If
End Sub

Private Sub txtatte_KeyPress(KeyAscii As Integer)
On Error GoTo SALE
If KeyAscii = 13 Then
 grid_fac.Row = 2
 grid_fac.COL = 2
 grid_fac.COL = 1
 TEXTOVAR.Visible = True
 TEXTOVAR.SetFocus
End If
Exit Sub
SALE:
End Sub

Private Sub txtcli_Change()
 If txtcli.Text = "" Then
   lblcli.Caption = ""
   i_destino.Clear
 End If
End Sub

Private Sub txtcli_LostFocus()
Dim WMO As String
Dim RES_DEUDA As Currency
Dim wsumadol As Currency
Dim WTC As Currency
If Val(txtcli.Text) = 0 Then Exit Sub

pub_cadena = "SELECT * FROM DIRCLI WHERE CODCIA=? AND CODCLI=? AND CP=?"
Set PSFAR_TRANS = CN.CreateQuery("", pub_cadena)
PSFAR_TRANS.rdoParameters(0) = LK_CODCIA
PSFAR_TRANS.rdoParameters(1) = Val(txtcli.Text)
PSFAR_TRANS.rdoParameters(2) = "C"
Set FAR_TRANS = PSFAR_TRANS.OpenResultset(rdOpenKeyset, rdConcurValues)
i_destino.Clear
Do Until FAR_TRANS.EOF
  i_destino.AddItem Trim(FAR_TRANS!DIRCOMP) & String(80, " ") & Trim(FAR_TRANS!DIRCLI)
  FAR_TRANS.MoveNext
Loop
If UCase(cmdIngreso.Caption) = "&GRABAR" Or cmdIngreso.Enabled = False Then GoTo fin
If i_destino.ListCount > 0 Then i_destino.ListIndex = 0
If Not cli_llave.EOF Then
    PUB_CODCIA = LK_CODCIA
    PUB_CODCLIE = Val(txtcli.Text)
    'If WMO <> "" Then MsgBox WMO, 48, Pub_Titulo
    For fila = 0 To i_fbg.ListCount - 1
        i_fbg.ListIndex = fila
        If Trim(i_fbg.Text) = Trim(cli_llave!CLI_TIPO) Then Exit For
    Next fila
    
End If

If Not cli_llave.EOF Then
   SQ_OPER = 1
   PUB_CODCIA = "00"
   PUB_TIPREG = 66
   PUB_NUMTAB = Val(cli_llave!CLI_DIA_VISITA)
   LEER_TAB_LLAVE
   If tab_llave.EOF Then
      MsgBox "No Tiene Dia de Visita, Verificar.", 48, Pub_Titulo
   Else
      lblvisita.Caption = Trim(tab_llave!tab_nomlargo)
   End If

   If LK_FLAG_EXED = "A" Then
     If Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) <> 1 And pub_signo_car <> 0 And Nulo_Valor0(cli_llave!cli_limcre) = 0 And Nulo_Valor0(cli_llave!cli_limcre2) = 0 Then
         MsgBox "Cliente No tiene Limite de Credito.", 48, Pub_Titulo
     End If
   End If
   If SUT_LLAVE!SUT_SIGNO_CAR = 1 Then
      pu_codcia = LK_CODCIA
      pub_deuda = CAR_TOT_CPX2("C", pu_codcia, cli_llave!cli_codclie)
      If PUB_FLAG_VENCIDO = 1 Or PUB_FLAG_VENCIDO_VISTA = 1 Then
        If PUB_FLAG_VENCIDO_VISTA = 1 Then
          MsgBox "OJO... Tiene Documentos Pendientes.", 48, Pub_Titulo
        Else
          MsgBox "Cliente Tiene Obligaciones Vencidas. << Moroso>>.", 48
        End If
         MsgBox pub_mensaje, 48, Pub_Titulo
      End If
      If Nulo_Valors(cli_llave!CLI_TIPO_BLOQ1) = "1" Then
         MsgBox "Cliente con Credito Bloqueado ... (No procede su Venta al Credito)", 48, Pub_Titulo
         txtcli.Text = ""
         Azul txtcli, txtcli
         Exit Sub
      End If
'      If PUB_FLAG_DOC > Nulo_Valor0(cli_llave!CLI_AUTO1) Then
'         MsgBox "Cliente alcanzo el tope de Documentos " & Chr(13) & "Emiitidos : " & PUB_FLAG_DOC & Chr(13) & "Autorizados : " & Trim(Nulo_Valor0(cli_llave!CLI_AUTO1)) & Chr(13) & " No procede la Venta", 48, Pub_Titulo
'         i_codcli.Text = ""
'         Exit Sub
'      End If
'   End If
End If
End If


' ver si usuario tiene acceso

If Not cli_llave.EOF And Nulo_Valors(par_llave!par_flag_cred) <> "A" And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 Then
   If SUT_LLAVE!SUT_SIGNO_CAR = 1 Then
      pu_codcia = LK_CODCIA
      pub_deuda = CAR_TOT_CPX2("C", pu_codcia, cli_llave!cli_codclie)
      If PUB_FLAG_VENCIDO = 1 And LK_FLAG_LIMITE <> "A" And LK_FLAG_LIMITE <> "C" Then
          MsgBox "CLIENTE TIENE OBLIGACIONES VENCIDAS ... ", 48, Pub_Titulo
          Exit Sub
    End If
    PUB_CAL_INI = LK_FECHA_DIA
    PUB_CAL_FIN = LK_FECHA_DIA
    pu_codcia = LK_CODCIA
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 1
    LEER_CAL_LLAVE
    WTC = 0
    If Not cal_llave.EOF Then
      WTC = Nulo_Valor0(cal_llave!cal_tipo_cambio)
    End If
    If par_llave!PAR_MONEDA_FAC = "S" Then
       WTC = 1
    Else
      If WTC = 0 Then
       MsgBox "Venta falta parametros ...INGRESE TIPO DE CAMBIO DEL DIA", 48, Pub_Titulo
        Exit Sub
      End If
    End If
    If Trim(Left(moneda.Text, 1)) = "S" Then
     wsumadol = Val(Nulo_Valor0(cli_llave!cli_limcre)) + Val(redondea((Nulo_Valor0(cli_llave!cli_limcre2) * WTC)))
     RES_DEUDA = pub_deuda
     WMO = "S/."
    Else
     wsumadol = Val(redondea(Nulo_Valor0(cli_llave!cli_limcre) / WTC)) + Val(redondea(Val(Nulo_Valor0(cli_llave!cli_limcre2))))
     RES_DEUDA = redondea(Val(pub_deuda / WTC))
     WMO = "US$."
    End If
If (RES_DEUDA + Val(txttotal.Text)) > wsumadol And LK_FLAG_LIMITE <> "B" And LK_FLAG_LIMITE <> "C" Then
   MsgBox "LIMITE DE CREDITO EXCEDIDO ...SALDO POR ATENDER : " & WMO & " " & Format(wsumadol - RES_DEUDA, "0.00") & Chr(13) & "*** Venta No Procede ***", 48, Pub_Titulo
 '  txtcli.Text = ""
'   Azul txtcli, txtcli
   Exit Sub
End If
   End If
End If

fin:
End Sub

Private Sub txtdoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 PUB_NUMSER = Val(tserie.Text)
 PUB_NUMFAC = Val(txtdoc.Text)
 LLENA_DOCU
End If
End Sub

Private Sub UNIDAD_KeyDown(KeyCode As Integer, Shift As Integer)
If UNIDAD.Visible = False Then Exit Sub
If KeyCode = 37 Then
  UNIDAD.Visible = False
  grid_fac.SetFocus
  grid_fac.COL = grid_fac.COL - 1
  grid_fac.SetFocus
End If
End Sub

Private Sub unidad_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
 UNIDAD.Visible = False
 grid_fac.SetFocus
End If


If KeyAscii <> 13 Then Exit Sub


SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
PUB_SECUEN = Val(Right(UNIDAD.Text, 4))
LEER_PRE_LLAVE
If pre_llave.EOF Then Exit Sub
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
LEER_ARM_LLAVE
If redondea(Val(arm_llave!arm_stock) / Val(pre_llave!PRE_EQUIV)) < Val(grid_fac.TextMatrix(grid_fac.Row, 2)) Then
    MsgBox "El Stock Actual es : " & Format(Val(arm_llave!arm_stock) / Val(pre_llave!PRE_EQUIV), "0.00") & " " & Trim(pre_llave!pre_UNIDAD) & "  -  La Cantidad Faltante es  : " & Format(Val(grid_fac.TextMatrix(grid_fac.Row, 2)) - Val(arm_llave!arm_stock), "0.00"), 48, Pub_Titulo
    If Trim(LK_USU_STOCK) <> "A" Then
        grid_fac.TextMatrix(grid_fac.Row, 16) = ""
        grid_fac.TextMatrix(grid_fac.Row, 15) = ""
        grid_fac.TextMatrix(grid_fac.Row, 12) = ""
        grid_fac.TextMatrix(grid_fac.Row, 11) = ""
        grid_fac.TextMatrix(grid_fac.Row, 0) = ""
        grid_fac.TextMatrix(grid_fac.Row, 1) = ""
        grid_fac.TextMatrix(grid_fac.Row, 2) = ""
        grid_fac.TextMatrix(grid_fac.Row, 10) = ""
        grid_fac.TextMatrix(grid_fac.Row, 3) = ""
        grid_fac.TextMatrix(grid_fac.Row, 4) = "" 'Format(Val(grid_fac.TextMatrix(grid_fac.Row, 11)) / Val(grid_fac.TextMatrix(grid_fac.Row, 17)), "0.00")
        grid_fac.TextMatrix(grid_fac.Row, 7) = ""
        grid_fac.TextMatrix(grid_fac.Row, 6) = ""
        grid_fac.TextMatrix(grid_fac.Row, 11) = ""
        grid_fac.TextMatrix(grid_fac.Row, 12) = ""
        grid_fac.TextMatrix(grid_fac.Row, 16) = ""
        suma_grid
        UNIDAD.Visible = False
        PRECIOS.Visible = False
        grid_fac.COL = 1
        Exit Sub
    End If
    
End If

grid_fac.TextMatrix(grid_fac.Row, 3) = Trim(Left(UNIDAD.Text, 12))
grid_fac.TextMatrix(grid_fac.Row, 4) = "" 'Format(Val(grid_fac.TextMatrix(grid_fac.Row, 11)) / Val(grid_fac.TextMatrix(grid_fac.Row, 17)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 7) = redondea(Nulo_Valor0(pre_llave!pre_PESO) * Val(grid_fac.TextMatrix(grid_fac.Row, 2)))
grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!PRE_EQUIV
grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_UNIDAD
stock.Caption = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 15)) / Val(grid_fac.TextMatrix(grid_fac.Row, 12)), "0.00")
unid.Caption = grid_fac.TextMatrix(grid_fac.Row, 16)
nomarti.Caption = grid_fac.TextMatrix(grid_fac.Row, 0)

UNIDAD.Visible = False
suma_grid
grid_fac.COL = 4
grid_fac_KeyPress 13

End Sub
Public Function REP_CONSUL() As Integer
Dim WMONEDA As String * 1
Dim wser As String * 3
Dim WSRUTA As String
Dim indice As Integer
Dim wm As Integer
Dim llave_rep01 As rdoResultset
Dim PS_REP01 As rdoQuery
Dim I As Integer
Dim VALOR
Dim loc_xl As Object
Dim loc_codtra As Integer
Dim wRuta As String
Dim WSNUMDOC As String
Dim numero_device As Integer
'If LK_EMP = "HER" Then
'  wRuta = "C:\ADMIN\STANDAR\"
'Else
LOC_TIPMOV = 201
If LK_EMP_PTO = "A" Then
  wRuta = PUB_RUTA_OTRO & "PTOVTA\"
Else
  wRuta = PUB_RUTA_OTRO
End If
If Left(moneda.Text, 1) = "S" Then
 WMONEDA = "S"
Else
 WMONEDA = "D"
End If

'End If
  If Trim(Nulo_Valors(par_llave!PAR_DEVICE_FBG)) <> "" Then
     'numero_device = 0
     'Reportes.PrinterName = Printers(numero_device).DeviceName
     'Reportes.PrinterDriver = Printers(numero_device).DriverName '"RASDD.DLL"
     'Reportes.PrinterPort = Printers(numero_device).Port
  End If

    FORM_COT.Reportes.Connect = PUB_ODBC
    FORM_COT.Reportes.Destination = crptToWindow  '= crptToPrinter
    FORM_COT.Reportes.WindowLeft = 2
    FORM_COT.Reportes.WindowTop = 70
    FORM_COT.Reportes.WindowWidth = 635
    FORM_COT.Reportes.WindowHeight = 390
    FORM_COT.Reportes.Formulas(1) = ""
    PUB_NETO = Val(txttotal.Text)
    PU_NUMSER = Val((tserie.Text))
    PU_NUMFAC = Val((txtdoc.Text))
    FORM_COT.Reportes.Formulas(1) = ""
    FORM_COT.Reportes.Formulas(1) = "SON_EFECTIVO=  ' " & CONVER_LETRAS(PUB_NETO, WMONEDA) & "'"
    FORM_COT.Reportes.WindowTitle = "GUIA DE COTIZACION  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "00000000")
    FORM_COT.Reportes.ReportFileName = wRuta + "COTI.RPT"
    pub_cadena = "{PEDIDOS.PED_TIPMOV} = " & LOC_TIPMOV & " AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND  {PEDIDOS.PED_NUMSER} = '" & PU_NUMSER & "' AND {PEDIDOS.PED_NUMFAC} = " & PU_NUMFAC
    FORM_COT.Reportes.SelectionFormula = pub_cadena
    On Error GoTo accion
    FORM_COT.Reportes.Action = 1
    On Error GoTo 0
Exit Function
accion:
 MsgBox Err.Description
 MsgBox "Intente Nuevamente, la impresion de Modo manual", 48, Pub_Titulo
 Exit Function

End Function

Private Sub txtcli_GotFocus()
 Azul txtcli, txtcli
End Sub

Private Sub txtcli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txtcli.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  txtcli.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  txtcli.SelStart = Len(txtcli.Text)
fin:

End Sub
Private Sub txtcli_KeyPress(KeyAscii As Integer)
Dim var As String
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As MSComctlLib.ListItem
On Error GoTo SALCODI

If KeyAscii = 27 Then
 txtcli.Text = ""
 lblcli.Caption = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
If Left(Trim(txtcli.Text), 1) = "+" Then GoTo buscar
On Error GoTo CODI_ERR

If IsNumeric(txtcli.Text) = True Then
   If Len(Trim(txtcli.Text)) = LK_DIG_RUC Then ' LONG DEL RUC
        pu_cp = "C"
        PUB_RUC = Trim(txtcli.Text)
        SQ_OPER = 4
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_ruc.EOF Then
           MsgBox "R.U.C. No Existe ", 48, Pub_Titulo
           Exit Sub
        End If
        txtcli.Text = cli_ruc!cli_codclie
   End If
   SQ_OPER = 1
   On Error GoTo mucho
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = Val(txtcli.Text)
   LEER_CLI_LLAVE
   On Error GoTo 0
   If cli_llave.EOF Then
     Azul txtcli, txtcli
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     txtcli.SetFocus
     GoTo fin
   End If
   ListView1.Visible = False
   txtcli.Text = cli_llave!cli_codclie
   FORM_COT.lblcli.Caption = cli_llave!CLI_NOMBRE
   If Trim(cli_llave!cli_ruc_esposo) <> "" Then
     txtruc.Text = cli_llave!cli_ruc_esposo
   Else
     txtruc.Text = cli_llave!cli_RUC_ESPOSA
   End If
   GoTo salta_dir
   Screen.MousePointer = 0
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(txtcli.Text)) = Left(VALOR, Len(Trim(txtcli.Text))) Then
   Else
      Exit Sub
   End If
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
   LEER_CLI_LLAVE
   On Error GoTo 0
   If cli_llave.EOF Then
     Azul txtcli, txtcli
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     txtcli.SetFocus
     GoTo fin
   End If
   ListView1.Visible = False
   txtcli.Text = cli_llave!cli_codclie
   FORM_COT.lblcli.Caption = cli_llave!CLI_NOMBRE
   If Trim(cli_llave!cli_ruc_esposo) <> "" Then
     txtruc.Text = cli_llave!cli_ruc_esposo
   Else
     txtruc.Text = cli_llave!cli_RUC_ESPOSA
   End If
   GoTo salta_dir
End If
dale:
ListView1.Visible = False
fin:
mucho:
CODI_ERR:
Exit Sub
SALCODI:
MsgBox Err.Description & " Intente Nuevamente ", 48, Pub_Titulo
Unload frmCLI
Exit Sub
salta_dir:
If KeyAscii = 13 Then
If Not cli_llave.EOF Then
    If Trim(cli_llave!cli_ruc_esposo) <> "" Then
      If Len(Trim(cli_llave!cli_ruc_esposo)) <> LK_DIG_RUC Then
         MsgBox "R.U.C. del Cliente no es de " & LK_DIG_RUC & " Digitos", 48, Pub_Titulo
         Azul txtcli, txtcli
         Exit Sub
      End If
    End If
End If


 If Left(Trim(Right(Trim(i_condi.Text), 3)), 1) = "1" And Trim(Right(Trim(i_condi.Text), 2)) = "FA" Then
    i_dias.Text = Val(Nulo_Valor0(cli_llave!CLI_AUTO1))
 Else
    i_dias.Text = ""
 End If
 If Left(Trim(Right(Trim(i_condi.Text), 3)), 1) = "1" And Trim(Right(Trim(i_condi.Text), 2)) = "FA" Then
    i_dias.Locked = False
    Azul i_dias, i_dias
 Else
    i_dias.Text = ""
    i_dias.Locked = True
    i_dias_KeyPress 13
 End If
End If
Exit Sub
buscar:
If Left(txtcli.Text, 2) = "++" Then
 var = Mid(txtcli.Text, 3, Len(txtcli.Text))
 numarchi = alta_vista_nombre(ListView1, var, "C", "D")
Else
 var = Mid(txtcli.Text, 2, Len(txtcli.Text))
 numarchi = alta_vista_nombre(ListView1, var, "C")
End If
If numarchi = 0 Then
  ListView1.Visible = False
  MsgBox "Alta Vista: No Existe .. Esta descripcion..", 48, Pub_Titulo
Else
  ListView1.Visible = True
  txtcli.SetFocus
End If
loc_key = 1
Exit Sub


End Sub

Private Sub txtcli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim NADA
Dim var
If Len(txtcli.Text) = 0 Or IsNumeric(txtcli.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(txtcli.Text) = 1 Then
    If txtcli.Text = "" Then txtcli.Text = " "
    var = Asc(txtcli.Text)
    var = var + 1
    NADA = var
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 1
    'archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM FROM CLIENTES WHERE  CLI_CP = 'C' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txtcli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = 'C' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txtcli.Text & "' AND  '" & var & "' ORDER BY CLI_NOMBRE"
    PROC_LISVIEW ListView1, 3000
    loc_key = 1
    If NADA = 33 Or NADA = 91 Then
      If ListView1.Visible = False Then
        loc_key = 0
        MsgBox "No existe Datos ...", 48, Pub_Titulo
        txtcli.Text = ""
      End If
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(txtcli.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If
End Sub


Public Sub LLENA_DOCU()
Dim MON As String
LIMPIA_DATOS
CABE_MAN

PSLOC_WARTI(0) = LK_CODCIA
PSLOC_WARTI(1) = 201
PSLOC_WARTI(2) = PUB_NUMSER
PSLOC_WARTI(3) = PUB_NUMFAC
llave_sum_arti.Requery
If llave_sum_arti.EOF Then
  tserie.Text = PUB_NUMSER
  txtdoc.Text = PUB_NUMFAC
  MsgBox "No Existe Pedidos de Venta.", 48, Pub_Titulo
  Azul txtdoc, txtdoc
  Exit Sub
End If

For fila = 0 To i_condi.ListCount - 1
  If Val(Trim(Left(i_condi.List(fila), 2))) = Val(llave_sum_arti!PED_CONDI) Then
   i_condi.ListIndex = fila
  End If
Next fila
i_condi_LostFocus
txtigv.Text = llave_sum_arti!PED_IGV
txtvalorv.Text = llave_sum_arti!PED_BRUTO
txttotal.Text = Format(llave_sum_arti!PED_IGV + llave_sum_arti!PED_BRUTO, "0.00")
FORM_COT.lblcli.Caption = llave_sum_arti!PED_NOMCLIE
txtruc.Text = llave_sum_arti!PED_RUCCLIE
txtcli.Text = llave_sum_arti!PED_CODCLIE
Txt_key.Text = llave_sum_arti!PED_codven
i_dias.Text = llave_sum_arti!ped_DIAS
txtfecha.Text = Format(llave_sum_arti!PED_fecha, "dd/mm/yyyy")
If Trim(llave_sum_arti!ped_situacion) = "P" Then
 LBLSIT.Caption = "01 P E D I D O   P R O C E S A D O - " & Trim(llave_sum_arti!PED_FORMA)
 LBLSIT.ForeColor = vbBlue
ElseIf Trim(llave_sum_arti!ped_situacion) = "" Then
 LBLSIT.Caption = "02 P E D I D O   P E N D I E N T E "
 LBLSIT.ForeColor = vbWhite
ElseIf Trim(llave_sum_arti!ped_situacion) = "E" Then
 LBLSIT.Caption = "99 P E D I D O   A N U L A DO "
 LBLSIT.ForeColor = vbRed

End If
If Trim(llave_sum_arti!ped_FBG) = "F" Then
 i_fbg.ListIndex = 0
Else
 i_fbg.ListIndex = 1
End If

If llave_sum_arti!PED_MONEDA = "S" Then
 moneda.ListIndex = 0
 i_moneda.Caption = "S/."
 grid_fac.TextMatrix(1, 4) = "S/."
Else
 moneda.ListIndex = 1
 i_moneda.Caption = "US$."
 grid_fac.TextMatrix(1, 4) = "US$."
End If
txtcli_KeyPress 13
txtcli_LostFocus
SQ_OPER = 1
PUB_CODCIA = "00"
PUB_TIPREG = 66
PUB_NUMTAB = Val(llave_sum_arti!PED_DIA_VISITA)
LEER_TAB_LLAVE
If tab_llave.EOF Then
  lblvisita.Caption = ""
Else
   lblvisita.Caption = Trim(tab_llave!tab_nomlargo)
End If
  

For fila = 0 To i_destino.ListCount - 1
  If Val(Trim(Right(i_destino.List(fila), 8))) = Val(llave_sum_arti!ped_DIRCLI) Then
   i_destino.ListIndex = fila
  End If
Next fila

For fila = 0 To cmdtipo.ListCount - 1
  If Val(Trim(Right(cmdtipo.List(fila), 8))) = Val(llave_sum_arti!PED_TIPVTA) Then
   cmdtipo.ListIndex = fila
  End If
Next fila

tserie.Text = PUB_NUMSER
txtdoc.Text = PUB_NUMFAC
txtatte.Text = Nulo_Valors(llave_sum_arti!PED_CONTACTO)

lbluser.Caption = Trim(llave_sum_arti!PED_CODUSU)

fila = 2
Do Until llave_sum_arti.EOF
   SQ_OPER = 1
   PUB_KEY = llave_sum_arti!PED_CODART
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   grid_fac.Rows = grid_fac.Rows + 1
   grid_fac.RowHeight(grid_fac.Rows - 1) = 285
   grid_fac.TextMatrix(fila, 1) = Trim(art_LLAVE!art_alterno)
   grid_fac.TextMatrix(fila, 0) = art_LLAVE!art_nombre
   
   grid_fac.TextMatrix(fila, 11) = llave_sum_arti!PED_NUM_UNIDAD
   grid_fac.TextMatrix(fila, 14) = llave_sum_arti!PED_NUMPRE
   
   grid_fac.TextMatrix(fila, 2) = Format(llave_sum_arti!PED_cantidad, "0.00")
   grid_fac.TextMatrix(fila, 4) = Format(llave_sum_arti!PED_PRECIO, "0.0000")
   grid_fac.TextMatrix(fila, 10) = llave_sum_arti!PED_CODART
   grid_fac.TextMatrix(fila, 3) = llave_sum_arti!PED_UNIDAD
   grid_fac.TextMatrix(fila, 12) = llave_sum_arti!PED_EQUIV
   grid_fac.TextMatrix(fila, 5) = llave_sum_arti!PED_descto
   grid_fac.TextMatrix(fila, 20) = llave_sum_arti!PED_descto_pre ' SOLO DESCTO DEL NUMERO DE PRECIO
   grid_fac.TextMatrix(fila, 21) = Format(llave_sum_arti!PED_descto_pre, "#0.0#") & "+" & Format(llave_sum_arti!PED_descto, "#0.0#") ' mUESTRA LOS PORC,

   fila = fila + 1
   llave_sum_arti.MoveNext
Loop
suma_grid
grid_fac.Enabled = True

ESTADO.Enabled = True
Azul txtdoc, txtdoc
cmdIngreso.Caption = "&Grabar"
cmdIngreso.Enabled = True
tserie.Enabled = False
txtdoc.Enabled = False
End Sub

Public Sub carga_venta()
SQ_OPER = 2
PUB_CODTRA = 2401
LEER_SUT_LLAVE
i_condi.Clear
Do Until SUT_MAYOR.EOF
 i_condi.AddItem Format(SUT_MAYOR!SUT_SECUENCIA, "00") & ".-" & SUT_MAYOR!sut_descripcion & String(180, " ") & SUT_MAYOR!SUT_SIGNO_CAR & SUT_MAYOR!sut_TIPDOC
 SUT_MAYOR.MoveNext
Loop
moneda.Clear
If LK_MONEDA = "S" Then
   moneda.AddItem "S = S/."
ElseIf LK_MONEDA = "D" Then
   moneda.AddItem "D = US$"
Else
   moneda.AddItem "S = S/."
   moneda.AddItem "D = US$"
End If
txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")

End Sub
Public Function VERIFICA_REPET() As Integer
Dim I As Integer
VERIFICA_REPET = 0
For I = 2 To grid_fac.Row - 1
If grid_fac.TextMatrix(I, 10) = art_LLAVE!ART_KEY Then
'  MsgBox "Ojo ...Articulo ya existe en lista", 48, Pub_Titulo
  VERIFICA_REPET = 1
End If
Next I
End Function

Private Sub UNIDAD_LostFocus()
UNIDAD.Visible = False
End Sub

Public Sub llena_numfac()
Dim wser As Integer
Dim wnumfac As Currency
PSTEMP_MAYOR(0) = LK_CODCIA
temp_mayor.Requery
If temp_mayor.EOF Then
 wser = 100
 wnumfac = 1
Else
 wser = 100
 wnumfac = Val(Nulo_Valor0(temp_mayor!PED_NUMFAC)) + 1
End If
tserie.Text = wser
txtdoc.Text = wnumfac


End Sub

Public Function PED_PENDIENTES() As String
Dim XXCUENTA As Integer

Dim TOT As Currency
Dim cade As String
PSDOC_LLAVE(0) = PUB_CODCIA
PSDOC_LLAVE(1) = LK_FECHA_DIA
PSDOC_LLAVE(2) = PUB_CODVEN
peddoc_llave.Requery
PUB_IMPORTE = 0
cade = ""
TOT = 0
GPED.Clear
GPED.TextMatrix(0, 0) = "Fec."
GPED.TextMatrix(0, 1) = "Doc."
GPED.TextMatrix(0, 2) = "Cliente"
GPED.TextMatrix(0, 3) = "Codven"
GPED.TextMatrix(0, 4) = "Importe"
GPED.TextMatrix(0, 5) = "Usuario"
GPED.ColWidth(0) = 600
GPED.ColWidth(1) = 700
GPED.ColWidth(2) = 1300
GPED.ColWidth(3) = 700
GPED.ColWidth(4) = 900


GPED.Rows = 2
XXCUENTA = 0
Do Until peddoc_llave.EOF
  XXCUENTA = XXCUENTA + 1
  GPED.Rows = GPED.Rows + 1
  PUB_IMPORTE = Val(peddoc_llave!PED_BRUTO) + Val(peddoc_llave!PED_IGV)
  GPED.TextMatrix(GPED.Rows - 1, 0) = Format(peddoc_llave!PED_fecha, "dd/mm")
  GPED.TextMatrix(GPED.Rows - 1, 1) = Format(peddoc_llave!PED_NUMFAC, "00000")
  GPED.TextMatrix(GPED.Rows - 1, 2) = Trim(peddoc_llave!PED_NOMCLIE)
  GPED.TextMatrix(GPED.Rows - 1, 3) = Format(peddoc_llave!PED_codven, "00")
  GPED.TextMatrix(GPED.Rows - 1, 4) = Format(PUB_IMPORTE, "0.00")
  GPED.TextMatrix(GPED.Rows - 1, 5) = peddoc_llave!PED_CODUSU
TOT = TOT + PUB_IMPORTE
peddoc_llave.MoveNext
Loop
GPED.TextMatrix(1, 0) = "T.Ped."
GPED.TextMatrix(1, 1) = "= " & Format(XXCUENTA, "00")
GPED.TextMatrix(1, 3) = "T.Mont."
GPED.TextMatrix(1, 4) = Format(TOT, "0.00")

PED_PENDIENTES = cade
End Function

Public Sub MUESTRA_PED()
cancelar_Click
cmdconsulta_Click
txtdoc.Text = Val(gridf.TextMatrix(gridf.Row, 0))
txtdoc_KeyPress 13
End Sub

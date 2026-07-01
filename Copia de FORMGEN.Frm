VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FORMGEN 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Operaciones en linea."
   ClientHeight    =   6495
   ClientLeft      =   75
   ClientTop       =   1380
   ClientWidth     =   9480
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
   Icon            =   "FORMGEN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6495
   ScaleWidth      =   9480
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid msgrid 
      Height          =   4935
      Left            =   6630
      TabIndex        =   267
      Top             =   1410
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   8705
      _Version        =   393216
   End
   Begin VB.Frame fralotes 
      Caption         =   "Informacion de Lote"
      Height          =   5535
      Left            =   4290
      TabIndex        =   222
      Tag             =   "9999"
      Top             =   1035
      Visible         =   0   'False
      Width           =   5700
      Begin MSFlexGridLib.MSFlexGrid det_lot 
         Height          =   4020
         Left            =   5040
         TabIndex        =   224
         Tag             =   "9999"
         Top             =   1305
         Visible         =   0   'False
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   7091
         _Version        =   393216
         Rows            =   1
         Cols            =   12
         FixedCols       =   0
         AllowUserResizing=   3
      End
      Begin VB.TextBox textolote 
         Height          =   315
         Left            =   3345
         TabIndex        =   234
         Top             =   4845
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdltaceptar 
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
         Height          =   345
         Left            =   645
         TabIndex        =   230
         Tag             =   "9999"
         Top             =   5085
         Width           =   1455
      End
      Begin VB.CommandButton cmdltcancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2190
         TabIndex        =   229
         Tag             =   "9999"
         Top             =   5070
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid gridlt 
         Height          =   3645
         Left            =   90
         TabIndex        =   223
         Tag             =   "9999"
         Top             =   1365
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   6429
         _Version        =   393216
         Cols            =   4
         FixedRows       =   2
         FixedCols       =   0
         AllowUserResizing=   3
         Appearance      =   0
      End
      Begin VB.Label lblt 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   3525
         TabIndex        =   236
         Tag             =   "9999"
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblltunidad 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   4335
         TabIndex        =   233
         Tag             =   "9999"
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lblt 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad Solicitada :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   60
         TabIndex        =   232
         Tag             =   "9999"
         Top             =   960
         Width           =   1785
      End
      Begin VB.Label lblltcantidad 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   1860
         TabIndex        =   231
         Tag             =   "9999"
         Top             =   960
         Width           =   1590
      End
      Begin VB.Label lblt 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   60
         TabIndex        =   228
         Tag             =   "9999"
         Top             =   570
         Width           =   1170
      End
      Begin VB.Label lblltnom 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   1260
         TabIndex        =   227
         Tag             =   "9999"
         Top             =   570
         Width           =   4350
      End
      Begin VB.Label lblt 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   226
         Tag             =   "9999"
         Top             =   255
         Width           =   1170
      End
      Begin VB.Label lblltcod 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   1275
         TabIndex        =   225
         Tag             =   "9999"
         Top             =   240
         Width           =   960
      End
   End
   Begin ComctlLib.ListView LV_TRA 
      Height          =   375
      Left            =   4800
      TabIndex        =   177
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView LV_ART 
      Height          =   375
      Left            =   8400
      TabIndex        =   107
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView LV_CLI 
      Height          =   375
      Left            =   7560
      TabIndex        =   106
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView LV_CCM 
      Height          =   375
      Left            =   6720
      TabIndex        =   108
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView LV_VEN 
      Height          =   375
      Left            =   5880
      TabIndex        =   111
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CmdEsp 
      Caption         =   "Descuento Especial"
      Height          =   495
      Left            =   3480
      TabIndex        =   269
      Tag             =   "199"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdtemp 
      BackColor       =   &H0080FFFF&
      Caption         =   "Temporales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   266
      Tag             =   "0"
      Top             =   2490
      Width           =   1260
   End
   Begin VB.CommandButton cmddd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cronograma de Pagos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   265
      Tag             =   "0"
      Top             =   1560
      Width           =   1320
   End
   Begin VB.Frame fraPedidos 
      Caption         =   "Relación de Pedidos :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   -15
      TabIndex        =   207
      Tag             =   "9999"
      Top             =   465
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CommandButton cmdchequea 
         Caption         =   "Che&quear Pend."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         TabIndex        =   221
         Tag             =   "9999"
         Top             =   1080
         Width           =   1605
      End
      Begin VB.ListBox lstcanal 
         Height          =   510
         Left            =   3150
         Style           =   1  'Checkbox
         TabIndex        =   219
         Tag             =   "9999"
         Top             =   1020
         Width           =   2550
      End
      Begin VB.CommandButton cmdcanped 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   7740
         TabIndex        =   216
         Tag             =   "9999"
         Top             =   5550
         Width           =   1605
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "Mandar a Impresora"
         Height          =   195
         Left            =   165
         TabIndex        =   205
         Tag             =   "9999"
         Top             =   5565
         Value           =   1  'Checked
         Width           =   2130
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar Documentos"
         Height          =   315
         Left            =   2295
         TabIndex        =   206
         Tag             =   "9999"
         Top             =   5550
         Width           =   3495
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Mostrar"
         Height          =   780
         Left            =   9180
         TabIndex        =   203
         Tag             =   "9999"
         Top             =   495
         Width           =   900
      End
      Begin VB.TextBox txtNumDocFinBol 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   7995
         Locked          =   -1  'True
         TabIndex        =   202
         Tag             =   "9999"
         Top             =   645
         Width           =   840
      End
      Begin VB.TextBox txtSerieDocBolFin 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   7470
         Locked          =   -1  'True
         TabIndex        =   201
         Tag             =   "9999"
         Top             =   645
         Width           =   465
      End
      Begin VB.TextBox txtNumDocIniBol 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6510
         TabIndex        =   200
         Tag             =   "9999"
         Top             =   645
         Width           =   840
      End
      Begin VB.TextBox txtSerieDocIniBol 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5970
         TabIndex        =   199
         Tag             =   "9999"
         Top             =   645
         Width           =   465
      End
      Begin VB.TextBox txtNumDocFacFin 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   4905
         Locked          =   -1  'True
         TabIndex        =   197
         Tag             =   "9999"
         Top             =   645
         Width           =   855
      End
      Begin VB.TextBox txtSerieDocFacFin 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   196
         Tag             =   "9999"
         Top             =   645
         Width           =   465
      End
      Begin VB.TextBox txtNumDocIniFac 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3375
         TabIndex        =   195
         Tag             =   "9999"
         Top             =   645
         Width           =   855
      End
      Begin VB.TextBox txtSerieDocIniFac 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2805
         TabIndex        =   194
         Tag             =   "9999"
         Top             =   645
         Width           =   465
      End
      Begin VB.CheckBox chkBoletas 
         Caption         =   "Seleccionar Boletas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5865
         TabIndex        =   198
         Tag             =   "9999"
         Top             =   195
         Width           =   2340
      End
      Begin VB.CheckBox chkFacturas 
         Caption         =   "Seleccionar Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2655
         TabIndex        =   193
         Tag             =   "9999"
         Top             =   225
         Width           =   2040
      End
      Begin VB.TextBox txtCodVen 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1185
         TabIndex        =   192
         Tag             =   "9999"
         Top             =   1020
         Width           =   555
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_detalle 
         Height          =   3945
         Left            =   120
         TabIndex        =   204
         Tag             =   "9999"
         Top             =   1560
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   6959
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
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
      Begin MSMask.MaskEdBox txtCampo2 
         Height          =   285
         Left            =   1215
         TabIndex        =   191
         Tag             =   "9999"
         Top             =   630
         Width           =   1320
         _ExtentX        =   2328
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
      Begin MSMask.MaskEdBox txtCampo1 
         Height          =   285
         Left            =   1200
         TabIndex        =   190
         Tag             =   "9999"
         Top             =   255
         Width           =   1260
         _ExtentX        =   2223
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
      Begin VB.Label lblped 
         Caption         =   "Canal de Distrib.:"
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
         Index           =   5
         Left            =   2355
         TabIndex        =   220
         Tag             =   "9999"
         Top             =   1020
         Width           =   1020
      End
      Begin VB.Label lblped 
         Caption         =   "Porcesando Información..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   9
         Left            =   2085
         TabIndex        =   217
         Tag             =   "9999"
         Top             =   2790
         Width           =   2745
      End
      Begin VB.Label i_nombreven 
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
         Left            =   135
         TabIndex        =   215
         Top             =   1275
         Width           =   2490
      End
      Begin VB.Label lblped 
         Caption         =   "Nº Doc. Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   8
         Left            =   7485
         TabIndex        =   214
         Tag             =   "9999"
         Top             =   450
         Width           =   1590
      End
      Begin VB.Label lblped 
         Caption         =   "Nº Doc. Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   5985
         TabIndex        =   213
         Tag             =   "9999"
         Top             =   450
         Width           =   1590
      End
      Begin VB.Label lblped 
         Caption         =   "Nº Doc. Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   4350
         TabIndex        =   212
         Tag             =   "9999"
         Top             =   435
         Width           =   1590
      End
      Begin VB.Label lblped 
         Caption         =   "Nº Doc. Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   2805
         TabIndex        =   211
         Tag             =   "9999"
         Top             =   435
         Width           =   1590
      End
      Begin VB.Label lblped 
         Caption         =   "Cod. Vend :"
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
         Left            =   195
         TabIndex        =   210
         Tag             =   "9999"
         Top             =   975
         Width           =   885
      End
      Begin VB.Label lblped 
         Caption         =   "Fecha Final  :"
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
         Left            =   120
         TabIndex        =   209
         Tag             =   "9999"
         Top             =   615
         Width           =   990
      End
      Begin VB.Label lblped 
         Caption         =   "Fecha Inicial :"
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
         Left            =   105
         TabIndex        =   208
         Tag             =   "9999"
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Frame fraprecios 
      BackColor       =   &H00808080&
      Caption         =   "Lista de Precios"
      ForeColor       =   &H00404040&
      Height          =   1320
      Left            =   60
      TabIndex        =   237
      Tag             =   "0"
      Top             =   6990
      Visible         =   0   'False
      Width           =   11805
      Begin VB.Label lbllote 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Left            =   2520
         TabIndex        =   270
         Tag             =   "0"
         Top             =   120
         Width           =   5685
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reserva"
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
         Height          =   270
         Left            =   10740
         TabIndex        =   264
         Tag             =   "0"
         Top             =   180
         Width           =   810
      End
      Begin VB.Label lpstocku3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   10590
         TabIndex        =   263
         Tag             =   "0"
         Top             =   825
         Width           =   1140
      End
      Begin VB.Label lpstock3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   10590
         TabIndex        =   262
         Tag             =   "0"
         Top             =   435
         Width           =   1155
      End
      Begin VB.Label lbltitulostock2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock2"
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
         Height          =   270
         Left            =   9450
         TabIndex        =   261
         Tag             =   "0"
         Top             =   195
         Width           =   810
      End
      Begin VB.Label lbltitulostock 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock1"
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
         Height          =   270
         Left            =   8280
         TabIndex        =   260
         Tag             =   "0"
         Top             =   195
         Width           =   810
      End
      Begin VB.Label lpstock2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   9330
         TabIndex        =   259
         Tag             =   "0"
         Top             =   435
         Width           =   1155
      End
      Begin VB.Label lpstocku2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   9360
         TabIndex        =   258
         Tag             =   "0"
         Top             =   825
         Width           =   1140
      End
      Begin VB.Label lblp 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio 6:"
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
         Height          =   255
         Index           =   9
         Left            =   6090
         TabIndex        =   257
         Tag             =   "0"
         Top             =   630
         Width           =   645
      End
      Begin VB.Label lpp6 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   6000
         TabIndex        =   256
         Tag             =   "0"
         Top             =   870
         Width           =   885
      End
      Begin VB.Label lblp 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Unid = "
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
         Height          =   270
         Index           =   8
         Left            =   6720
         TabIndex        =   255
         Tag             =   "0"
         Top             =   930
         Width           =   1290
      End
      Begin VB.Label lpstocku 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   8115
         TabIndex        =   254
         Tag             =   "0"
         Top             =   825
         Width           =   1140
      End
      Begin VB.Label lpp5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   4830
         TabIndex        =   253
         Tag             =   "0"
         Top             =   870
         Width           =   885
      End
      Begin VB.Label lblp 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio 4 :"
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
         Height          =   255
         Index           =   7
         Left            =   4920
         TabIndex        =   252
         Tag             =   "0"
         Top             =   630
         Width           =   765
      End
      Begin VB.Label lpp4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   3705
         TabIndex        =   251
         Tag             =   "0"
         Top             =   870
         Width           =   885
      End
      Begin VB.Label lblp 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio 3 :"
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
         Height          =   270
         Index           =   6
         Left            =   3810
         TabIndex        =   250
         Tag             =   "0"
         Top             =   630
         Width           =   765
      End
      Begin VB.Label lpp3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   2520
         TabIndex        =   249
         Tag             =   "0"
         Top             =   870
         Width           =   885
      End
      Begin VB.Label lblp 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio 2 :"
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
         Height          =   270
         Index           =   5
         Left            =   2580
         TabIndex        =   248
         Tag             =   "0"
         Top             =   630
         Width           =   765
      End
      Begin VB.Label lpp2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   1365
         TabIndex        =   247
         Tag             =   "0"
         Top             =   870
         Width           =   885
      End
      Begin VB.Label lblp 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio 1 :"
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
         Height          =   270
         Index           =   4
         Left            =   1470
         TabIndex        =   246
         Tag             =   "0"
         Top             =   630
         Width           =   765
      End
      Begin VB.Label lpp1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   210
         TabIndex        =   245
         Tag             =   "0"
         Top             =   870
         Width           =   885
      End
      Begin VB.Label lblp 
         BackStyle       =   0  'Transparent
         Caption         =   "P.Cred. :"
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
         Height          =   225
         Index           =   3
         Left            =   300
         TabIndex        =   244
         Tag             =   "0"
         Top             =   630
         Width           =   765
      End
      Begin VB.Label lpunidad 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "UND"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   6000
         TabIndex        =   243
         Tag             =   "0"
         Top             =   300
         Width           =   1110
      End
      Begin VB.Label lblp 
         BackStyle       =   0  'Transparent
         Caption         =   "Unid."
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
         Height          =   225
         Index           =   2
         Left            =   5430
         TabIndex        =   242
         Tag             =   "0"
         Top             =   285
         Width           =   645
      End
      Begin VB.Label lpstock 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   8085
         TabIndex        =   241
         Tag             =   "0"
         Top             =   435
         Width           =   1155
      End
      Begin VB.Label lblp 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Cajas = "
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
         Height          =   270
         Index           =   1
         Left            =   6720
         TabIndex        =   240
         Tag             =   "0"
         Top             =   555
         Width           =   1290
      End
      Begin VB.Label lpnombre 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto :"
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   1005
         TabIndex        =   239
         Tag             =   "0"
         Top             =   360
         Width           =   5640
      End
      Begin VB.Label lblp 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto :"
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
         Height          =   270
         Index           =   0
         Left            =   150
         TabIndex        =   238
         Tag             =   "0"
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame Fracli 
      Caption         =   "Datos:"
      ForeColor       =   &H00808000&
      Height          =   975
      Left            =   -75
      TabIndex        =   178
      Tag             =   "9999"
      Top             =   7740
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox t_doc 
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
         Left            =   5760
         TabIndex        =   184
         Tag             =   "9999"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_direc 
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
         Left            =   1080
         TabIndex        =   183
         Tag             =   "9999"
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox t_nombre 
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
         Left            =   1080
         TabIndex        =   182
         Tag             =   "9999"
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Lnom 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Ident. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   2
         Left            =   4800
         TabIndex        =   181
         Tag             =   "9999"
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Lnom 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   180
         Tag             =   "9999"
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Lnom 
         AutoSize        =   -1  'True
         Caption         =   "Nombre: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   179
         Tag             =   "9999"
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.ComboBox cmdtipo 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   7875
      Style           =   2  'Dropdown List
      TabIndex        =   218
      Tag             =   "206"
      Top             =   7695
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdpedidos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Facturación por Pedidos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10455
      Picture         =   "FORMGEN.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   189
      Tag             =   "0"
      Top             =   825
      Width           =   1320
   End
   Begin VB.ComboBox i_situacion 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      ItemData        =   "FORMGEN.frx":0884
      Left            =   7875
      List            =   "FORMGEN.frx":088E
      Style           =   2  'Dropdown List
      TabIndex        =   136
      Tag             =   "133"
      Top             =   7935
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.ComboBox i_condi 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   7875
      Style           =   2  'Dropdown List
      TabIndex        =   188
      Tag             =   "301"
      Top             =   8295
      Visible         =   0   'False
      Width           =   2950
   End
   Begin VB.ComboBox cmbserie 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   186
      Tag             =   "9999"
      Top             =   6470
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox cheser 
      BackColor       =   &H00808000&
      Caption         =   "Cambiar Serie."
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
      Height          =   435
      Left            =   5880
      TabIndex        =   187
      Tag             =   "9999"
      Top             =   6420
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox i_dircli 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   255
      Style           =   2  'Dropdown List
      TabIndex        =   185
      Tag             =   "300"
      Top             =   8160
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.CommandButton Boton_Recepcion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rece&pcion Artic."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   1800
      Picture         =   "FORMGEN.frx":0898
      Style           =   1  'Graphical
      TabIndex        =   122
      Tag             =   "158"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Boton_Letras 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Doc. Pendientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   240
      Picture         =   "FORMGEN.frx":0CDA
      Style           =   1  'Graphical
      TabIndex        =   65
      Tag             =   "151"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton transferImport 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Proceso Automatico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10455
      Picture         =   "FORMGEN.frx":111C
      Style           =   1  'Graphical
      TabIndex        =   176
      Tag             =   "0"
      Top             =   75
      Width           =   1320
   End
   Begin VB.CommandButton Diario 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Diari&o"
      Height          =   600
      Left            =   10440
      Picture         =   "FORMGEN.frx":155E
      Style           =   1  'Graphical
      TabIndex        =   115
      Tag             =   "153"
      Top             =   2910
      Width           =   1335
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   195
      Left            =   0
      TabIndex        =   79
      Tag             =   "0"
      Top             =   6960
      Visible         =   0   'False
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   0
      Min             =   77
      Max             =   91
   End
   Begin VB.Frame frafondo 
      Height          =   10
      Left            =   0
      TabIndex        =   170
      Tag             =   "9999"
      Top             =   480
      Visible         =   0   'False
      Width           =   10215
   End
   Begin VB.TextBox i_camal 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8880
      MaxLength       =   255
      TabIndex        =   95
      Tag             =   "97"
      Text            =   "i_camal"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5280
      Picture         =   "FORMGEN.frx":19A0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   163
      Tag             =   "9999"
      Top             =   6480
      Width           =   480
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_all 
      Height          =   495
      Left            =   5880
      TabIndex        =   76
      Tag             =   "0"
      Top             =   720
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393216
      Rows            =   50
      Cols            =   15
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      BackColorBkg    =   8421376
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
      FormatString    =   $"FORMGEN.frx":1DE2
   End
   Begin MSMask.MaskEdBox i_fecha_can 
      Height          =   285
      Left            =   5280
      TabIndex        =   159
      Tag             =   "200"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   128
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox i_fecha_oper 
      Height          =   285
      Left            =   2520
      TabIndex        =   147
      Tag             =   "109"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   128
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox i_fecha_compra 
      Height          =   285
      Left            =   2160
      TabIndex        =   35
      Tag             =   "36"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   128
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Frame fragas 
      Caption         =   "Descripcion de Gasto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   5520
      TabIndex        =   150
      Tag             =   "9999"
      Top             =   960
      Visible         =   0   'False
      Width           =   4092
      Begin VB.TextBox cta_gas2 
         BackColor       =   &H00E0E0E0&
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
         Height          =   288
         Left            =   840
         MaxLength       =   12
         TabIndex        =   154
         Tag             =   "9999"
         Top             =   600
         Width           =   852
      End
      Begin VB.TextBox imp_gas2 
         BackColor       =   &H00E0E0E0&
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
         Height          =   288
         Left            =   2640
         TabIndex        =   153
         Tag             =   "9999"
         Top             =   600
         Width           =   1092
      End
      Begin VB.TextBox imp_gas1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   288
         Left            =   2640
         TabIndex        =   152
         Tag             =   "9999"
         Top             =   240
         Width           =   1092
      End
      Begin VB.TextBox cta_gas1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   288
         Left            =   840
         MaxLength       =   12
         TabIndex        =   151
         Tag             =   "9999"
         Top             =   240
         Width           =   852
      End
      Begin VB.Label lblgastos 
         Caption         =   "Imp. 2:"
         ForeColor       =   &H00808000&
         Height          =   252
         Index           =   3
         Left            =   1920
         TabIndex        =   158
         Tag             =   "9999"
         Top             =   600
         Width           =   732
      End
      Begin VB.Label lblgastos 
         Caption         =   "Imp. 1:"
         ForeColor       =   &H00808000&
         Height          =   252
         Index           =   2
         Left            =   1920
         TabIndex        =   157
         Tag             =   "9999"
         Top             =   240
         Width           =   732
      End
      Begin VB.Label lblgastos 
         Caption         =   "Cta 2:"
         ForeColor       =   &H00808000&
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   156
         Tag             =   "9999"
         Top             =   600
         Width           =   612
      End
      Begin VB.Label lblgastos 
         Caption         =   "Cta 1:"
         ForeColor       =   &H00808000&
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   155
         Tag             =   "9999"
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.TextBox NJ 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   400
      Left            =   600
      TabIndex        =   149
      Tag             =   "99"
      Text            =   "NJ"
      Top             =   3720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox i_turno 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   7320
      TabIndex        =   148
      Tag             =   "63"
      Text            =   "i_precio"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_serguia 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   -120
      TabIndex        =   146
      Tag             =   "64"
      Text            =   "i_serguia"
      Top             =   2640
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CheckBox i_cambio 
      BackColor       =   &H00808000&
      Caption         =   "Cambiar Nro."
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
      Height          =   255
      Left            =   9000
      TabIndex        =   145
      Tag             =   "0"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fcomun 
      Caption         =   "Frame3"
      Height          =   1095
      Left            =   4800
      TabIndex        =   143
      Tag             =   "0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   5055
      Begin MSFlexGridLib.MSFlexGrid grid_unid 
         Height          =   975
         Left            =   120
         TabIndex        =   144
         Tag             =   "0"
         Top             =   120
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   1720
         _Version        =   393216
         Rows            =   50
         Cols            =   11
         BackColor       =   16777215
         BackColorFixed  =   12632256
         ForeColorFixed  =   64
         BackColorBkg    =   8421376
         FocusRect       =   2
         HighLight       =   2
         AllowUserResizing=   1
      End
   End
   Begin VB.TextBox i_tasav 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2160
      TabIndex        =   142
      Tag             =   "77"
      Text            =   "i_tasav"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_moneda 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   140
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox i_codban2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2160
      TabIndex        =   137
      Tag             =   "70"
      Text            =   "i_codban"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Guia Ant."
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   5280
      TabIndex        =   135
      Tag             =   "125"
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grid_trans 
      Height          =   495
      Left            =   5400
      TabIndex        =   134
      Tag             =   "0"
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393216
      Rows            =   50
      Cols            =   15
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
      FormatString    =   $"FORMGEN.frx":1E6E
   End
   Begin VB.ComboBox i_cias 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "FORMGEN.frx":1EF5
      Left            =   3840
      List            =   "FORMGEN.frx":1EF7
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   133
      Tag             =   "13"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox i_def 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   6480
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "9999"
      Top             =   25
      Width           =   3720
   End
   Begin VB.ComboBox i_bancos 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      ItemData        =   "FORMGEN.frx":1EF9
      Left            =   4560
      List            =   "FORMGEN.frx":1EFB
      Style           =   2  'Dropdown List
      TabIndex        =   126
      Tag             =   "0"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1000
   End
   Begin RichTextLib.RichTextBox textovarl 
      Height          =   375
      Left            =   6600
      TabIndex        =   125
      Tag             =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   14737632
      MultiLine       =   0   'False
      TextRTF         =   $"FORMGEN.frx":1EFD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox textovar_canje 
      Height          =   375
      Left            =   2280
      TabIndex        =   124
      Tag             =   "0"
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   14737632
      MultiLine       =   0   'False
      MousePointer    =   1
      TextRTF         =   $"FORMGEN.frx":1FCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox i_tipdoc 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      ItemData        =   "FORMGEN.frx":208E
      Left            =   1080
      List            =   "FORMGEN.frx":2090
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Tag             =   "16"
      Top             =   840
      Visible         =   0   'False
      Width           =   765
   End
   Begin MSFlexGridLib.MSFlexGrid grid_canje 
      Height          =   495
      Left            =   8040
      TabIndex        =   78
      Tag             =   "105"
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393216
      Rows            =   3
      Cols            =   7
      FixedRows       =   2
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   3
   End
   Begin VB.ComboBox i_is 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      ItemData        =   "FORMGEN.frx":2092
      Left            =   1080
      List            =   "FORMGEN.frx":209C
      Style           =   2  'Dropdown List
      TabIndex        =   123
      Tag             =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox LisTransa 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5040
      TabIndex        =   53
      Tag             =   "0"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_dias 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   117
      Tag             =   "43"
      Text            =   "i_dias"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Tipo_Cambio 
      BackColor       =   &H00808000&
      Caption         =   "&T/ Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   112
      Tag             =   "167"
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox i_unidades 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   110
      Tag             =   "0"
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox i_precios 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   2295
      Style           =   2  'Dropdown List
      TabIndex        =   109
      Tag             =   "0"
      Top             =   7335
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox i_numfac 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   7920
      TabIndex        =   105
      TabStop         =   0   'False
      Tag             =   "26"
      Text            =   "i_numfac"
      Top             =   6435
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.ComboBox i_mortal 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   104
      Tag             =   "0"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox textovar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   360
      TabIndex        =   103
      Tag             =   "0"
      Text            =   "textovar"
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_nompro 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3240
      MaxLength       =   255
      TabIndex        =   88
      Tag             =   "96"
      Text            =   "i_nompro"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox i_num_cheque 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6600
      MaxLength       =   255
      TabIndex        =   87
      Tag             =   "95"
      Text            =   "i_num_cheque"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox i_nombre_banco 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   7680
      MaxLength       =   255
      TabIndex        =   86
      Tag             =   "94"
      Text            =   "i_nombre_banco"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox i_num_lote 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   5400
      TabIndex        =   85
      Tag             =   "91"
      Text            =   "I_num_lote"
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_TEXTONCRE 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   5280
      TabIndex        =   84
      Tag             =   "90"
      Text            =   "i_textoncre"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox i_numser_r 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   600
      TabIndex        =   83
      TabStop         =   0   'False
      Tag             =   "88"
      Text            =   "i_numser_r"
      Top             =   600
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox i_serdoc 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3840
      TabIndex        =   82
      Tag             =   "86"
      Text            =   "i_serdoc"
      Top             =   3000
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox i_numser_c 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3240
      TabIndex        =   81
      Tag             =   "84"
      Text            =   "i_numser_c"
      Top             =   3000
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox i_numfac_c 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1080
      TabIndex        =   80
      Tag             =   "85"
      Text            =   "i_numfac_c"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.ComboBox i_fbg 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      ItemData        =   "FORMGEN.frx":20BC
      Left            =   2280
      List            =   "FORMGEN.frx":20C9
      Style           =   2  'Dropdown List
      TabIndex        =   73
      Tag             =   "14"
      Top             =   2880
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.ComboBox i_ds 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      ItemData        =   "FORMGEN.frx":20D6
      Left            =   1080
      List            =   "FORMGEN.frx":20E0
      Style           =   2  'Dropdown List
      TabIndex        =   75
      Tag             =   "76"
      Top             =   720
      Visible         =   0   'False
      Width           =   765
   End
   Begin MSFlexGridLib.MSFlexGrid grid_autorizacion 
      Height          =   495
      Left            =   7320
      TabIndex        =   71
      Tag             =   "0"
      Top             =   720
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393216
      Rows            =   50
      Cols            =   15
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      BackColorBkg    =   8421376
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
      FormatString    =   " |#  |Usuario|Articulo|Precio|Factura|Cantidad|Tasa%|Dias|Cash|Est.|Codtra|   Fecha     |  hora        |"
   End
   Begin VB.CommandButton boton_autorizacion 
      Caption         =   "&Autoriz."
      Height          =   375
      Left            =   4905
      TabIndex        =   74
      Tag             =   "150"
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_num_oper 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   7680
      TabIndex        =   70
      Tag             =   "73"
      Text            =   "I_num_oper"
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_num_ini 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1200
      TabIndex        =   68
      Tag             =   "66"
      Text            =   "i_num_ini"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_cheser 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   0
      MaxLength       =   3
      TabIndex        =   67
      Tag             =   "52"
      Text            =   "i_cheser"
      Top             =   4080
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ComboBox i_cant_cheq 
      Height          =   315
      ItemData        =   "FORMGEN.frx":20EA
      Left            =   6360
      List            =   "FORMGEN.frx":20EC
      TabIndex        =   66
      Tag             =   "65"
      Text            =   "i_cant_cheq"
      Top             =   1560
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.TextBox i_tipo_bloq_ant4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   4200
      TabIndex        =   63
      Tag             =   "62"
      Text            =   "i_tipo_bloq_ant"
      Top             =   720
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox i_tipo_bloq_ant3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3600
      TabIndex        =   62
      Tag             =   "61"
      Text            =   "i_tipo_bloq_ant"
      Top             =   480
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox i_tipo_bloq_ant2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2880
      TabIndex        =   61
      Tag             =   "60"
      Text            =   "i_tipo_bloq_ant"
      Top             =   480
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.ComboBox i_tipo_bloq_act4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      ItemData        =   "FORMGEN.frx":20EE
      Left            =   600
      List            =   "FORMGEN.frx":20F0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   60
      Tag             =   "59"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ComboBox i_tipo_bloq_act3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      ItemData        =   "FORMGEN.frx":20F2
      Left            =   615
      List            =   "FORMGEN.frx":20F4
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   59
      Tag             =   "58"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ComboBox i_tipo_bloq_act2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      ItemData        =   "FORMGEN.frx":20F6
      Left            =   600
      List            =   "FORMGEN.frx":20F8
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   58
      Tag             =   "57"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ComboBox i_tipo_bloq_act1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      ItemData        =   "FORMGEN.frx":20FA
      Left            =   600
      List            =   "FORMGEN.frx":20FC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Tag             =   "56"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox i_tipo_bloq_ant1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   56
      Tag             =   "55"
      Text            =   "i_tipo_bloq_ant"
      Top             =   480
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox i_limcre_ant 
      Enabled         =   0   'False
      Height          =   285
      HideSelection   =   0   'False
      Left            =   6240
      MaxLength       =   14
      TabIndex        =   55
      Tag             =   "53"
      Text            =   "i_limcre_ant"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox i_limcre 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      HideSelection   =   0   'False
      Left            =   5400
      MaxLength       =   14
      TabIndex        =   54
      Tag             =   "54"
      Text            =   "i_limcre"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_cantidad 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2160
      TabIndex        =   52
      Tag             =   "51"
      Text            =   "i_cantidad"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_precio 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2160
      TabIndex        =   51
      Tag             =   "35"
      Text            =   "i_precio"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_codart 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1080
      TabIndex        =   50
      Tag             =   "5"
      Text            =   "i_codart"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_numser 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   7440
      TabIndex        =   48
      Tag             =   "25"
      Text            =   "i_numser"
      Top             =   6435
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox i_impto2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   47
      Tag             =   "114"
      Text            =   "i_impto2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.ComboBox i_codsunat 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "FORMGEN.frx":20FE
      Left            =   3120
      List            =   "FORMGEN.frx":2100
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Tag             =   "12"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "Ce&rrar"
      Height          =   625
      Left            =   10440
      Picture         =   "FORMGEN.frx":2102
      Style           =   1  'Graphical
      TabIndex        =   114
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancelar"
      Height          =   625
      Left            =   10440
      Picture         =   "FORMGEN.frx":224C
      Style           =   1  'Graphical
      TabIndex        =   113
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5100
      Width           =   1335
   End
   Begin VB.CommandButton grabar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Grabar"
      Height          =   625
      Left            =   10440
      Picture         =   "FORMGEN.frx":268E
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   3945
      Width           =   1335
   End
   Begin VB.ComboBox i_CODUSU 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Tag             =   "50"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox i_situacion_ant 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   405
      Left            =   6120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   43
      Tag             =   "134"
      Text            =   "FORMGEN.frx":2AD0
      Top             =   4440
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.TextBox i_bruto2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   3120
      MaxLength       =   12
      TabIndex        =   41
      Tag             =   "113"
      Text            =   "i_bruto2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox i_numplan 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   600
      TabIndex        =   34
      Tag             =   "37"
      Text            =   "i_diasV"
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_gastos_fijos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   33
      Tag             =   "10"
      Text            =   "FORMGEN.frx":2AE2
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox i_importe_amort 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   32
      Tag             =   "6"
      Text            =   $"FORMGEN.frx":2AF3
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox i_pordescto2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   -120
      TabIndex        =   31
      Tag             =   "34"
      Text            =   "i_pordescto2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox i_numguia 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3120
      TabIndex        =   30
      Tag             =   "18"
      Text            =   "i_numguia"
      Top             =   720
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_numdoc_r 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1080
      TabIndex        =   29
      Tag             =   "27"
      Text            =   "numdoc_r"
      Top             =   480
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_pordescto1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   -120
      TabIndex        =   23
      Tag             =   "33"
      Text            =   "i_pordescto1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox i_chenum 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1080
      MaxLength       =   8
      TabIndex        =   19
      Tag             =   "21"
      Text            =   "i_numche"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_codban 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Tag             =   "19"
      Text            =   "i_codban"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox i_numdoc 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Tag             =   "42"
      Text            =   "i_numdoc"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TRANS 
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
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   840
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "9999"
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox i_importe 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      HideSelection   =   0   'False
      Left            =   -120
      MaxLength       =   14
      TabIndex        =   7
      Tag             =   "7"
      Text            =   $"FORMGEN.frx":2B06
      Top             =   1560
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox i_codcli 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   -120
      TabIndex        =   6
      Tag             =   "4"
      Text            =   "i_codcli"
      Top             =   1200
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox i_codven 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   -120
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "41"
      Text            =   "i_codven"
      Top             =   840
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Item de Articulos"
      ForeColor       =   &H00808000&
      Height          =   2325
      Left            =   0
      TabIndex        =   21
      Tag             =   "100"
      Top             =   3960
      Visible         =   0   'False
      Width           =   10215
      Begin VB.TextBox i_cant 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   3840
         TabIndex        =   119
         Tag             =   "9999"
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox i_autorizacion 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   4920
         TabIndex        =   64
         TabStop         =   0   'False
         Tag             =   "9999"
         Text            =   "I_AUTORIZACION"
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame2 
         Height          =   760
         Left            =   45
         TabIndex        =   96
         Tag             =   "119"
         Top             =   1500
         Width           =   10215
         Begin VB.TextBox i_neto 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   300
            Left            =   7320
            Locked          =   -1  'True
            TabIndex        =   94
            TabStop         =   0   'False
            Tag             =   "9999"
            Text            =   "i_neto"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox i_flete 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   5520
            TabIndex        =   93
            TabStop         =   0   'False
            Tag             =   "9999"
            ToolTipText     =   "Doble Click para Porrrateo de Flete"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox i_impto 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   4080
            TabIndex        =   92
            TabStop         =   0   'False
            Tag             =   "9999"
            Text            =   "i_impto"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox i_descto 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   91
            TabStop         =   0   'False
            Tag             =   "9999"
            Text            =   "i_descto"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox i_gastos 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   1440
            MaxLength       =   25
            TabIndex        =   90
            TabStop         =   0   'False
            Tag             =   "9999"
            Text            =   "i_gastos"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox i_subtotal 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   89
            TabStop         =   0   'False
            Tag             =   "9999"
            Text            =   "i_subtotal"
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "EFECTIVO TOTAL"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   8760
            TabIndex        =   169
            Tag             =   "9999"
            Top             =   160
            Width           =   1335
         End
         Begin VB.Label LCODART 
            Caption         =   "SUB-TOTAL"
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   102
            Tag             =   "9999"
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label LCODART 
            Alignment       =   2  'Center
            Caption         =   "GASTO"
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   6
            Left            =   1560
            TabIndex        =   101
            Tag             =   "9999"
            Top             =   120
            Width           =   885
         End
         Begin VB.Label LCODART 
            Alignment       =   2  'Center
            Caption         =   "DESCTO.(%)"
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   5
            Left            =   2760
            TabIndex        =   100
            Tag             =   "9999"
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label LCODART 
            Caption         =   "IMPTO"
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   4
            Left            =   4320
            TabIndex        =   99
            Tag             =   "9999"
            Top             =   120
            Width           =   645
         End
         Begin VB.Label LCODART 
            Alignment       =   2  'Center
            Caption         =   "TOTAL"
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   2
            Left            =   7320
            TabIndex        =   98
            Tag             =   "9999"
            Top             =   120
            Width           =   885
         End
         Begin VB.Label LCODART 
            Alignment       =   2  'Center
            Caption         =   "FLETE"
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   1
            Left            =   5760
            TabIndex        =   97
            Tag             =   "9999"
            Top             =   120
            Width           =   765
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   1095
         Left            =   0
         TabIndex        =   69
         Tag             =   "9999"
         Top             =   195
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1931
         _Version        =   393216
         Rows            =   3
         Cols            =   9
         FixedRows       =   2
         BackColorBkg    =   8421376
         GridLines       =   3
         AllowUserResizing=   3
      End
      Begin VB.Label label_precio 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   3120
         TabIndex        =   120
         Tag             =   "9999"
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label label_nomart 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   3120
         TabIndex        =   118
         Tag             =   "9999"
         Top             =   600
         Width           =   765
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gridl 
      Height          =   495
      Left            =   8160
      TabIndex        =   72
      Tag             =   "166"
      Top             =   720
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      _Version        =   393216
      Rows            =   50
      Cols            =   16
      FixedRows       =   2
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      BackColorBkg    =   8421376
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin VB.TextBox i_fecha_vcto 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   360
      TabIndex        =   42
      Tag             =   "24"
      Text            =   "i_fecha_vcto"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1245
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   8760
      Tag             =   "9999"
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox i_concepto 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   6360
      MaxLength       =   255
      TabIndex        =   77
      Tag             =   "20"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estados del Cliente"
      Height          =   2175
      Left            =   6840
      TabIndex        =   127
      Tag             =   "44"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
      Begin VB.CheckBox BLOQ 
         Caption         =   "Check4"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   131
         Tag             =   "0"
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox BLOQ 
         Caption         =   "Check3"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   130
         Tag             =   "0"
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox BLOQ 
         Caption         =   "Check2"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   129
         Tag             =   "0"
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox BLOQ 
         Caption         =   "Check1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   128
         Tag             =   "0"
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.CommandButton letche 
      Caption         =   "Ingreso Docs..."
      Height          =   975
      Left            =   8880
      Picture         =   "FORMGEN.frx":2B15
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "1"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grid_liq 
      Height          =   495
      Left            =   8160
      TabIndex        =   141
      Tag             =   "177"
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393216
      Rows            =   50
      Cols            =   15
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   64
      BackColorBkg    =   8421376
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
      FormatString    =   " |Amortizacion  |    Int. Vencido| Int. Adelantado|  Gastos F. |  Total | Dias Venc.| Dias Adelant.| "
   End
   Begin VB.CheckBox cheigv 
      Caption         =   "Incluir I.G.V."
      ForeColor       =   &H00808000&
      Height          =   225
      Left            =   7755
      TabIndex        =   235
      Tag             =   "9999"
      Top             =   1815
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label lblfecalma 
      Caption         =   "Label9"
      Height          =   495
      Left            =   5400
      TabIndex        =   271
      Tag             =   "9999"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "[F8]=Precio Otros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   10470
      TabIndex        =   268
      Tag             =   "9999"
      Top             =   2190
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Tag             =   "9999"
      X1              =   5160
      X2              =   5160
      Y1              =   6420
      Y2              =   6920
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "[F4]=Cancelar Oper."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   1920
      TabIndex        =   175
      Tag             =   "9999"
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "[F3]=Otra Opción."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   1920
      TabIndex        =   174
      Tag             =   "9999"
      Top             =   6420
      Width           =   1275
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "[F5]=Otro codigo."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   3480
      TabIndex        =   173
      Tag             =   "9999"
      Top             =   6420
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "[F6]=Graba Oper."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   3480
      TabIndex        =   172
      Tag             =   "9999"
      Top             =   6720
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "Codigo:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   105
      TabIndex        =   171
      Tag             =   "9999"
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblop 
      BackColor       =   &H00808000&
      Caption         =   "Opción:"
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   5760
      TabIndex        =   166
      Tag             =   "9999"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblOper 
      BackColor       =   &H00808000&
      Caption         =   "Operación:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   164
      Tag             =   "9999"
      Top             =   75
      Width           =   975
   End
   Begin VB.Label BolFac 
      BackColor       =   &H00E0E0E0&
      Caption         =   "bol/fact."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   6960
      TabIndex        =   116
      Tag             =   "9999"
      Top             =   6465
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Height          =   6975
      Index           =   5
      Left            =   10320
      TabIndex        =   160
      Tag             =   "9999"
      Top             =   15
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   162
      Tag             =   "9999"
      Top             =   6360
      Width           =   10335
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
      TabIndex        =   161
      Tag             =   "9999"
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label LMONEDAB 
      Caption         =   "Lmoneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   139
      Tag             =   "0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label i_nomban2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "i_nomban2"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4440
      TabIndex        =   138
      Tag             =   "0"
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label label_grid 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Documentos Llegados <Seleccionar con ENTER>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   6720
      TabIndex        =   132
      Tag             =   "0"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Lmoneda 
      Caption         =   "Lmoneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   121
      Tag             =   "182"
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label i_nomart 
      BackColor       =   &H00C0C0C0&
      Caption         =   "i_nomart"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4320
      TabIndex        =   45
      Tag             =   "0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   17
      Left            =   9120
      TabIndex        =   40
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   16
      Left            =   9120
      TabIndex        =   39
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   15
      Left            =   9240
      TabIndex        =   38
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   14
      Left            =   9240
      TabIndex        =   37
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   13
      Left            =   9240
      TabIndex        =   36
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   12
      Left            =   9240
      TabIndex        =   28
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   11
      Left            =   9240
      TabIndex        =   27
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   10
      Left            =   9240
      TabIndex        =   26
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   9
      Left            =   9240
      TabIndex        =   25
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   8
      Left            =   9240
      TabIndex        =   24
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   7
      Left            =   9240
      TabIndex        =   16
      Tag             =   "0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   6
      Left            =   9255
      TabIndex        =   15
      Tag             =   "0"
      Top             =   2025
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   5
      Left            =   9255
      TabIndex        =   14
      Tag             =   "0"
      Top             =   2025
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   4
      Left            =   9255
      TabIndex        =   13
      Tag             =   "0"
      Top             =   2025
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   3
      Left            =   9255
      TabIndex        =   12
      Tag             =   "0"
      Top             =   2025
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   2
      Left            =   9255
      TabIndex        =   11
      Tag             =   "0"
      Top             =   2025
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   1
      Left            =   9255
      TabIndex        =   10
      Tag             =   "0"
      Top             =   2025
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LABELGEN 
      AutoSize        =   -1  'True
      Caption         =   "LABELGEN"
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   0
      Left            =   9255
      TabIndex        =   9
      Tag             =   "0"
      Top             =   2025
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label i_nomban 
      BackColor       =   &H00C0C0C0&
      Caption         =   "i_nomban"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Tag             =   "0"
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label i_nomven 
      Caption         =   "i_nomven"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Tag             =   "0"
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label i_nomcli 
      AutoSize        =   -1  'True
      Caption         =   "i_nomcli"
      DataField       =   "1234"
      Height          =   195
      Left            =   4320
      TabIndex        =   8
      Tag             =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label nomtra 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Tag             =   "9999"
      Top             =   60
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Height          =   400
      Index           =   1
      Left            =   0
      TabIndex        =   165
      Tag             =   "9999"
      Top             =   0
      Width           =   10270
   End
   Begin VB.Label lblcabe 
      BorderStyle     =   1  'Fixed Single
      Height          =   1770
      Left            =   0
      TabIndex        =   167
      Tag             =   "9999"
      Top             =   435
      Width           =   10280
   End
   Begin VB.Label LBLDET 
      BorderStyle     =   1  'Fixed Single
      Height          =   4095
      Left            =   -30
      TabIndex        =   168
      Tag             =   "9999"
      Top             =   2220
      Width           =   10275
   End
End
Attribute VB_Name = "FORMGEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temporal As String
Dim ws_fecha_anterior As Date
Dim WS_FLAG_FACART As String * 1
Dim ParaLot_count As Integer
Dim ParaLot_codlot(100) As String
Dim ParaLot_lotcant(100) As Currency
Dim ParaLot_fechalot(100) As String
Dim LOC_PEDIDO As Currency
Dim loc_stock As String
Dim loc_ped_llave As rdoResultset
Dim PSLOC_PED_LLAVE As rdoQuery
Dim loc_var_descto As Integer
Dim loc_secuencia As Integer
Dim dias_bloqueo As Integer
Dim loc_mofi_venta As String * 1
Dim PS_CAMBIO As rdoQuery
Dim far_cambio As rdoResultset
Dim WS_IMPORTE_AMORT As Currency
Dim flag_del As Integer
Dim PSGUI_MENOR As rdoQuery
Dim gui_menor As rdoResultset
Dim PSCAR_MENOR2 As rdoQuery
Dim car_menor2 As rdoResultset
Dim FLAG_DIARIO As String * 1
Dim PSALL_REC As rdoQuery
Dim ALL_REC As rdoResultset
Dim WS_ESTADO As Integer
Dim CONTADOR As Integer
Dim ww_codtra_ext As Integer
Dim tab_derL(10) As Integer
Dim tab_izqL(10) As Integer
Dim tab_derecha(12) As Integer
Dim tab_izquierda(12) As Integer
Dim PSFFF_ART As rdoQuery
Dim FFF_ART As rdoResultset

Dim PSFAR_TRANS As rdoQuery
Dim FAR_TRANS As rdoResultset

Dim pub_montodescto As Currency
Dim WS_NUMSEC As Integer
Public textovar_bak As String
Dim flag_salto As Integer
Dim loc_key As Integer
Dim WS_TIPO_CAMBIO  As Currency
Dim wtemporal As String * 1
Dim xl As Object
Dim WS_MONEDA_CLI As String * 1
Dim WS_MONEDA_CCM As String * 1
Dim VAR_ACTIVAR As Integer
Dim PLAZA_FLAG_MANUAL As String * 1
Dim ext_conta As Integer
Dim pub_numfac_ult As Long
Dim WS_NUM_RECIBO As Long
Dim WS_SER_RECIBO As Integer

'ACV
Dim ws_fecha_emision As Date
Dim LOC_MONEDA_DIG As String * 1
Dim flag_repite As String * 1
Dim WW_FECHA_ANT As Date
Dim LOC_IMP_INI As Currency
Dim LOC_DEFINI As Currency

Dim loc_new_llave As rdoResultset
Dim PSNEW_PRECIO As rdoQuery
Dim wlot_falso_flag As String * 1
Dim loc_acceso_des As String * 1
Option Explicit
Public Function LEE_CARACU() As Boolean
  LEE_CARACU = True
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_codclie = PUB_CODCLIE
    pu_cp = PUB_CP
    PUB_FECHA = far_llave!FAR_fecha
    PUB_NUM_OPER = far_llave!FAR_NUMOPER
    LEER_CAA_LLAVE
    If caa_histo.EOF = True Then
       MsgBox "DOCUMENTO EN CARACU NO EXISTE..."
       LEE_CARACU = False
    End If
End Function
Public Function VERIFICA_REPET() As Integer
Dim I As Integer
VERIFICA_REPET = 0
For I = 2 To grid_fac.Row - 1
If grid_fac.TextMatrix(I, 16) = art_LLAVE!ART_KEY Then
  MsgBox "Ojo ...Articulo ya existe en lista", 48, Pub_Titulo
  VERIFICA_REPET = 1
End If
Next I
End Function
Public Sub LIMPIA_PREC()
Dim fil
Dim COL
For fil = 2 To grid_unid.Rows
    For COL = 0 To grid_unid.Cols - 1
       gridl.TextMatrix(fil, COL) = ""
    Next COL
Next fil
End Sub

Public Sub vuelca_datos()
Dim indice1, INDICE2, INDICE3, INDICE4, CONTADOR
FORMGEN.i_limcre_ant.Text = Nulo_Valor0(cli_llave!cli_limcre)
FORMGEN.i_limcre.Text = Nulo_Valor0(cli_llave!cli_limcre)
FORMGEN.i_dias.Text = Nulo_Valor0(cli_llave!cli_dias_cred)
 
Exit Sub
' NO VALE LO DE ABAJO
If FORMGEN.i_tipo_bloq_act1.Visible = True Or FORMGEN.i_limcre.Visible = True Then
Else
   GoTo fin
End If
FORMGEN.i_tipo_bloq_ant1.Text = Nulo_Valors(cli_llave!CLI_TIPO_BLOQ1)
FORMGEN.i_tipo_bloq_ant2.Text = Nulo_Valors(cli_llave!CLI_TIPO_BLOQ2)
FORMGEN.i_tipo_bloq_ant3.Text = Nulo_Valors(cli_llave!CLI_TIPO_BLOQ3)
FORMGEN.i_tipo_bloq_ant4.Text = Nulo_Valors(cli_llave!CLI_TIPO_BLOQ4)

indice1 = 99
INDICE2 = 99
INDICE3 = 99
INDICE4 = 99
CONTADOR = -1
FORMGEN.i_tipo_bloq_act1.Clear
FORMGEN.i_tipo_bloq_act2.Clear
FORMGEN.i_tipo_bloq_act3.Clear
FORMGEN.i_tipo_bloq_act4.Clear

PUB_TIPREG = 7
SQ_OPER = 2
PUB_CODCIA = "00"
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
  FORMGEN.i_tipo_bloq_act1.AddItem Nulo_Valors(tab_mayor!tab_nomcorto)
  FORMGEN.i_tipo_bloq_act2.AddItem Nulo_Valors(tab_mayor!tab_nomcorto)
  FORMGEN.i_tipo_bloq_act3.AddItem Nulo_Valors(tab_mayor!tab_nomcorto)
  FORMGEN.i_tipo_bloq_act4.AddItem Nulo_Valors(tab_mayor!tab_nomcorto)
  CONTADOR = CONTADOR + 1
  If cli_llave!CLI_TIPO_BLOQ1 = Left(tab_mayor!tab_nomcorto, 1) Then
     indice1 = CONTADOR
  End If
  If cli_llave!CLI_TIPO_BLOQ2 = Left(tab_mayor!tab_nomcorto, 1) Then
     INDICE2 = CONTADOR
  End If
  If cli_llave!CLI_TIPO_BLOQ3 = Left(tab_mayor!tab_nomcorto, 1) Then
     INDICE3 = CONTADOR
  End If
  If cli_llave!CLI_TIPO_BLOQ4 = Left(tab_mayor!tab_nomcorto, 1) Then
     INDICE4 = CONTADOR
  End If

  tab_mayor.MoveNext
Loop

If indice1 <> 99 Then
FORMGEN.i_tipo_bloq_act1.ListIndex = indice1
End If
If INDICE2 <> 99 Then
FORMGEN.i_tipo_bloq_act2.ListIndex = INDICE2
End If
If INDICE3 <> 99 Then
FORMGEN.i_tipo_bloq_act3.ListIndex = INDICE3
End If
If INDICE4 <> 99 Then
FORMGEN.i_tipo_bloq_act4.ListIndex = INDICE4
End If
fin:
End Sub

Public Sub VERIFICA_VENTA()

End Sub
Public Sub verifica_precios()
Dim ws_pre_vta As Currency
Dim tp As Integer
Dim wcont As Integer
Dim descto_cont(20) As Currency
Dim cade_descto As String
Dim WS_DESCTO As Currency
Dim WS_PRECIO As Currency
Dim ws_int As Currency
Dim ww_num As Integer
On Error GoTo SALE
If art_LLAVE!ART_KEY = 0 Then Exit Sub
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_CODART = art_LLAVE!ART_KEY
  LEER_ARM_LLAVE

  pu_codcia = LK_CODCIA
  PUB_CODART = art_LLAVE!ART_KEY
  LEER_PRE_LLAVE
  WS_PRECIO = Val(pre_llave!PRE_PRE1)
  
If LK_EMP = "3AA" Then
'  grid_fac.TextMatrix(grid_fac.Row, 32) = WS_PRECIO
End If
grid_fac.TextMatrix(grid_fac.Row, 16) = PUB_CODART
grid_fac.TextMatrix(grid_fac.Row, 33) = PUB_CODART
grid_fac.TextMatrix(grid_fac.Row, 37) = Val(pre_llave!pre_PESO)
grid_fac.TextMatrix(grid_fac.Row, 43) = Val(Nulo_Valor0(pre_llave!PRE_LITRO))
i_precios.Clear

If LK_CODTRA = 2409 Then  ' TAMBIEN PARA LAS VENTAS CON VENTAS GUIA SOLO PASAN
   grid_fac.TextMatrix(grid_fac.Row, 6) = Format(arm_llave!arm_cospro * pre_llave!PRE_EQUIV, "0.0000")
   Exit Sub
End If
If LK_CODTRA = 2412 Then
    SQ_OPER = 1
    PU_TIPMOV = 10
    pu_codcia = LK_CODCIA
    PU_NUMSER = i_numser_c.Text
    PU_FBG = Trim(i_fbg.Text)
    PU_NUMFAC = Val(i_numfac_c.Text)
    LEER_FAR_LLAVE
    If far_llave.EOF Then
      MsgBox "Nro. Documento No Existe", 48, Pub_Titulo
      Exit Sub
    End If
    ws_pre_vta = 0
    Do Until far_llave.EOF
     If far_llave!far_estado <> "E" Then
        If Val(far_llave!far_codart) = PUB_CODART Then
           ws_pre_vta = Val(far_llave!FAR_PRECIO)
           Exit Do
        End If
     End If
    far_llave.MoveNext
    Loop
   grid_fac.TextMatrix(grid_fac.Row, 6) = ws_pre_vta
End If


If (Val(Nulo_Valors(SUT_LLAVE!SUT_PRECIO)) = 4 And PUB_TIPMOV = 10 And Val(Nulo_Valors(SUT_LLAVE!SUT_TIPMOV_REF)) = 0) Or LK_EMP = "3AA" Then
'   If FORMGEN.i_ds.Text = "D" Then
'      i_precios.AddItem Left(tab_precioss(2), 7) & "  " & Format(pre_llave!pre_pre11, "0.00") & String(20, " ") & "1"
'      grid_fac.TextMatrix(grid_fac.Row, 32) = pre_llave!PRE_PRE55
'      grid_fac.TextMatrix(grid_fac.Row, 35) = pre_llave!pre_pre11
'   Else
'      i_precios.AddItem Left(tab_precioss(1), 7) & "  " & Format(pre_llave!PRE_PRE1, "0.00") & String(20, " ") & "1"
'      grid_fac.TextMatrix(grid_fac.Row, 32) = pre_llave!PRE_PRE5
'      grid_fac.TextMatrix(grid_fac.Row, 35) = pre_llave!PRE_PRE1
'   End If

    'Exit Sub
End If

If (Val(Nulo_Valors(SUT_LLAVE!SUT_TIPMOV_REF)) <> 0 And PUB_TIPMOV = 9999) Then
  SQ_OPER = 2
  PUB_CODCIA = "00"
  PUB_TIPREG = 600
  LEER_TAB_LLAVE
  tp = 0
  Do Until tab_mayor.EOF
   If tab_mayor!tab_codclie = 1 Then
        SQ_OPER = 1
        PUB_CODCIA = LK_CODCIA
        PUB_TIPDES = tab_mayor!tab_numtab
        PUB_LISDES = cli_llave!CLI_SUBGRUPO
        PUB_CODART = PUB_CODART
        LEER_DESC_LLAVE
        If Not desc_llave.EOF Then
           tp = tp + 1
           descto_cont(tp) = Nulo_Valor0(desc_llave!cld_desto1)
        End If
   End If
   tab_mayor.MoveNext
  Loop
  
   SQ_OPER = 1
  '  loc_var_descto = cli_llave!CLI_SUBGRUPO 'SE CAMBIO FREDY
   PUB_CODCIA = LK_CODCIA
   PUB_TIPDES = SUT_LLAVE!SUT_TIPMOV_REF
   PUB_LISDES = loc_var_descto  ' se cambiocli_llave!CLI_SUBGRUPO ' CAMBIAR
   PUB_CODART = PUB_CODART
   LEER_DESC_LLAVE
   WS_DESCTO = 0
   If Not desc_llave.EOF Then
     WS_DESCTO = Nulo_Valor0(desc_llave!cld_desto1)
   End If
   If WS_DESCTO <> 0 Then
      tp = tp + 1
      descto_cont(tp) = WS_DESCTO
   End If
   cade_descto = ""
   If FORMGEN.i_ds.Text = "D" Then
      WS_PRECIO = pre_llave!pre_pre11
      For wcont = 1 To tp
         WS_DESCTO = descto_cont(wcont)
         cade_descto = cade_descto + Format(WS_DESCTO, "0.0#") & "+"
         WS_PRECIO = Format(WS_PRECIO * (100 - WS_DESCTO) / 100, "0.0000")
        'redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 35)) * (100 - WW_DESC1) / 100)
      Next wcont
      If Trim(cade_descto) <> "" Then
        cade_descto = Mid(cade_descto, 1, Len(cade_descto) - 1)
      End If
      grid_fac.TextMatrix(grid_fac.Row, 42) = cade_descto
      grid_fac.TextMatrix(grid_fac.Row, 10) = cade_descto
      'i_precios.AddItem Left(tab_precioss(2), 7) & "  " & Format(pre_llave!PRE_PRE11, "0.00") & String(20, " ") & "1"
      i_precios.AddItem Left(tab_precioss(2), 7) & "  " & Format(WS_PRECIO, "0.0000") & String(20, " ") & "1"
      grid_fac.TextMatrix(grid_fac.Row, 32) = pre_llave!PRE_PRE55
      grid_fac.TextMatrix(grid_fac.Row, 35) = pre_llave!pre_pre11
   Else
      WS_PRECIO = pre_llave!PRE_PRE1
      For wcont = 1 To tp
         WS_DESCTO = descto_cont(wcont)
         cade_descto = cade_descto + Format(WS_DESCTO, "0.0#") & "+"
         WS_PRECIO = Format((WS_PRECIO * (100 - WS_DESCTO) / 100), "0.0000")
        'redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 35)) * (100 - WW_DESC1) / 100)
      Next wcont
      If Trim(cade_descto) <> "" Then
        cade_descto = Mid(cade_descto, 1, Len(cade_descto) - 1)
      End If
      grid_fac.TextMatrix(grid_fac.Row, 42) = cade_descto
      grid_fac.TextMatrix(grid_fac.Row, 10) = cade_descto
      i_precios.AddItem Left(tab_precioss(2), 7) & "  " & Format(WS_PRECIO, "0.0000") & String(20, " ") & "1"
            ' i_precios.AddItem Left(tab_precioss(1), 7) & "  " & Format(pre_llave!PRE_PRE1, "0.00") & String(20, " ") & "1"
      grid_fac.TextMatrix(grid_fac.Row, 32) = pre_llave!PRE_PRE5
      grid_fac.TextMatrix(grid_fac.Row, 35) = pre_llave!PRE_PRE1
   End If

   Exit Sub
End If

If Not cli_llave.EOF Then
   ww_num = Val(Nulo_Valors(cli_llave!CLI_nucleo))
   If ww_num > 0 And ww_num < 6 Then
      i_precios.AddItem Left(tab_precioss(ww_num), 7) & "  " & Format(pre_llave.rdoColumns(8 + ww_num), "0.00") & String(20, " ") & ww_num
      Exit Sub
   End If
End If

If Nulo_Valors(SUT_LLAVE!SUT_PRECIO) = "1" Then
   grid_fac.TextMatrix(grid_fac.Row, 6) = Format(arm_llave!arm_cospro * pre_llave!PRE_EQUIV, "0.0000")
End If

If Nulo_Valors(SUT_LLAVE!SUT_PRECIO) <> "3" Then Exit Sub

If FORMGEN.i_ds.Text = "S" Then
   If cli_llave.EOF Then
     If LK_CODTRA = 2401 And (Val(i_codcli.Text) = 606 Or Trim(i_fbg.Text) = "G") Then
      i_precios.AddItem "Cos.Prom." & "  " & Format(arm_llave!arm_cospro * pre_llave!PRE_EQUIV, "0.000") & String(20, " ") & "0"
      GoTo sigue_por_costo_prom
     End If
   Else
    If LK_CODTRA = 2401 And (Val(i_codcli.Text) = 606 Or Trim(i_fbg.Text) = "G" Or Trim(cli_llave!cli_ruc_esposo) = "20103040406") Then
      i_precios.AddItem "Cos.Prom." & "  " & Format(arm_llave!arm_cospro * pre_llave!PRE_EQUIV, "0.000") & String(20, " ") & "0"
      GoTo sigue_por_costo_prom
    End If
    End If
    If (Left(SUT_LLAVE!SUT_PRECIOS, 1) = "1" Or (LK_PRECIO <> "C" And LK_PRECIO <> "B")) And Val(pre_llave!PRE_PRE1) <> 0 Then
        i_precios.AddItem Left(tab_precioss(1), 7) & "  " & Format(pre_llave!PRE_PRE1, "0.000") & String(20, " ") & "1"
    End If
    If (Mid(SUT_LLAVE!SUT_PRECIOS, 2, 1) = "1" Or (LK_PRECIO <> "C" And LK_PRECIO <> "B")) And Val(pre_llave!PRE_PRE2) <> 0 Then
         i_precios.AddItem Left(tab_precioss(2), 7) & "  " & Format(pre_llave!PRE_PRE2, "0.000") & String(20, " ") & "2"
    End If
    If (Mid(SUT_LLAVE!SUT_PRECIOS, 3, 1) = "1" Or (LK_PRECIO <> "C" And LK_PRECIO <> "B")) And Val(pre_llave!PRE_PRE3) <> 0 Then
       i_precios.AddItem Left(tab_precioss(3), 7) & "  " & Format(pre_llave!PRE_PRE3, "0.000") & String(20, " ") & "3"
    End If
    If (Mid(SUT_LLAVE!SUT_PRECIOS, 4, 1) = "1" Or (LK_PRECIO <> "C" And LK_PRECIO <> "B")) And Val(pre_llave!PRE_PRE4) <> 0 Then
       i_precios.AddItem Left(tab_precioss(4), 7) & "  " & Format(pre_llave!PRE_PRE4, "0.000") & String(20, " ") & "4"
    End If
    If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Or LK_CODUSU = "OPER20" Then
      If (Mid(SUT_LLAVE!SUT_PRECIOS, 5, 1) = "1" Or (LK_PRECIO <> "C" And LK_PRECIO <> "B")) And Val(pre_llave!PRE_PRE5) <> 0 Then
         i_precios.AddItem Left(tab_precioss(5), 7) & "  " & Format(pre_llave!PRE_PRE5, "0.000") & String(20, " ") & "5"
      End If
      i_precios.AddItem "Lista.6" & "  " & Format(pre_llave!PRE_PRE6, "0.000") & String(20, " ") & "6"
      If Val(pre_llave!PRE_PRE7) <> 0 Then i_precios.AddItem "Lista.7" & "  " & Format(pre_llave!PRE_PRE7, "0.000") & String(20, " ") & "7"
      If Val(pre_llave!PRE_PRE8) <> 0 Then i_precios.AddItem "Lista.8" & "  " & Format(pre_llave!PRE_PRE8, "0.000") & String(20, " ") & "8"
      If Val(pre_llave!PRE_PRE9) <> 0 Then i_precios.AddItem "Lista.9" & "  " & Format(pre_llave!PRE_PRE9, "0.000") & String(20, " ") & "9"
    End If
sigue_por_costo_prom:
Else
    If (Left(SUT_LLAVE!SUT_PRECIOS, 1) = "1" Or LK_PRECIO <> "C") And Val(pre_llave!pre_pre11) <> 0 Then
        i_precios.AddItem Left(tab_precioss(1), 7) & "  " & Format(pre_llave!pre_pre11, "0.000") & String(20, " ") & "1"
    End If
    If (Mid(SUT_LLAVE!SUT_PRECIOS, 2, 1) = "1" Or LK_PRECIO <> "C") And Val(pre_llave!PRE_PRE22) <> 0 Then
         i_precios.AddItem Left(tab_precioss(2), 7) & "  " & Format(pre_llave!PRE_PRE22, "0.000") & String(20, " ") & "2"
    End If
    If (Mid(SUT_LLAVE!SUT_PRECIOS, 3, 1) = "1" Or LK_PRECIO <> "C") And Val(pre_llave!PRE_PRE33) <> 0 Then
       i_precios.AddItem Left(tab_precioss(3), 7) & "  " & Format(pre_llave!PRE_PRE33, "0.000") & String(20, " ") & "3"
    End If
    If (Mid(SUT_LLAVE!SUT_PRECIOS, 4, 1) = "1" Or LK_PRECIO <> "C") And Val(pre_llave!PRE_PRE44) <> 0 Then
       i_precios.AddItem Left(tab_precioss(4), 7) & "  " & Format(pre_llave!PRE_PRE44, "0.000") & String(20, " ") & "4"
    End If
    If (Mid(SUT_LLAVE!SUT_PRECIOS, 5, 1) = "1" Or LK_PRECIO <> "C") And Val(pre_llave!PRE_PRE55) <> 0 Then
       i_precios.AddItem Left(tab_precioss(5), 7) & "  " & Format(pre_llave!PRE_PRE55, "0.000") & String(20, " ") & "5"
    End If
End If


Exit Sub
SALE:
  MsgBox "Verificar : " & Err.Description
Resume Next
End Sub
Public Sub LLENA_NUM_RECIBO()
If WS_ESTADO = 99 Or WS_ESTADO = 98 Or WS_ESTADO = 97 Then
 pu_codcia = LK_CODCIA
 PUB_TIPMOV = 0
 PU_NUMSER = Val(i_numser.Text)
 If WS_ESTADO = 98 Then
  PSALL_LLAVE2.rdoParameters(1) = 2720
  PSALL_LLAVE2.rdoParameters(2) = 5714
  PSALL_LLAVE2.rdoParameters(3) = 0
  PSALL_LLAVE2.rdoParameters(4) = 0
  PSALL_LLAVE2.rdoParameters(6) = 1
 ElseIf WS_ESTADO = 97 Then
  PSALL_LLAVE2.rdoParameters(1) = 5318
  PSALL_LLAVE2.rdoParameters(2) = -1
  PSALL_LLAVE2.rdoParameters(3) = -1
  PSALL_LLAVE2.rdoParameters(4) = -1
  PSALL_LLAVE2.rdoParameters(6) = -1
 Else
  PSALL_LLAVE2.rdoParameters(1) = 5714
  PSALL_LLAVE2.rdoParameters(2) = 2735
  PSALL_LLAVE2.rdoParameters(3) = 2738
  PSALL_LLAVE2.rdoParameters(4) = 0
  PSALL_LLAVE2.rdoParameters(6) = -1
 End If
 SQ_OPER = 3
 LEER_ALL_LLAVE
 If all_llave2.EOF Then
   i_numfac.Text = 1
 Else
   If pub_signo_ccm = 0 Then
     i_numfac.Text = 0 'all_llave2!all_numfac
     PUB_NUMFAC = 0 'all_llave2!all_numfac
   Else
     If LK_CODTRA = 5318 And pub_signo_ccm = 1 Then
     Else
      i_numfac.Text = all_llave2!all_numfac + 1
      PUB_NUMFAC = all_llave2!all_numfac + 1
     End If
   End If
 End If
 Exit Sub
End If
   If WS_ESTADO <> 2 Then Exit Sub
   If LK_EMP = "HER" Then Exit Sub

   PSCAR_MENOR2(0) = LK_CODCIA
   PSCAR_MENOR2(1) = Val(i_numser.Text)
   PSCAR_MENOR2.MaxRows = 1
   Set car_menor2 = PSCAR_MENOR2.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
   If car_menor2.EOF Then
      i_numfac.Text = 1
   Else
      i_numfac.Text = car_menor2!CAA_RECIBO + 1
   End If

End Sub
Public Sub LLENA_RECIBO_ALL()
   If WS_ESTADO <> 3 Then Exit Sub

   PSALL_REC(0) = LK_CODCIA
   PSALL_REC(1) = LK_CODTRA
   PSALL_REC(2) = PUB_SECUENCIA
   PSALL_REC(3) = Val(i_numser.Text)
   PSALL_REC.MaxRows = 1
   ALL_REC.Requery
   If ALL_REC.EOF Then
      i_numfac.Text = 1
   Else
      i_numfac.Text = ALL_REC!ALL_NUM_RECIBO + 1
   End If

End Sub

Public Sub verifica_unidades()
Dim cont As Integer
i_unidades.Clear
Dim longit As String * 15
Dim ww As String * 15
ww = "Unidades"
cont = -1
SQ_OPER = 2
PUB_CODART = art_LLAVE!ART_KEY
LEER_PRE_LLAVE
Do Until pre_mayor.EOF
   cont = cont + 1
   longit = pre_mayor!pre_UNIDAD
   i_unidades.AddItem longit & "       " & pre_mayor!pre_secuencia & "       " & pre_mayor!PRE_EQUIV
   If pre_mayor!PRE_FLAG_UNIDAD = "A" Then i_unidades.ListIndex = cont
   pre_mayor.MoveNext
Loop
If i_unidades.ListCount = 1 Then i_unidades.ListIndex = 0
If PUB_TIPMOV = 10 Then grid_fac.TextMatrix(grid_fac.Row, 10) = ""
If PUB_TIPMOV = 20 And i_unidades.ListCount <> 0 Then i_unidades.ListIndex = i_unidades.ListCount - 1

End Sub
Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional ConsisVal, Optional ConsisCol)
  Static VALOR
  Dim car As String
  If ConsisVal = 2 Then ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
'      Flag_Consis = "A"
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
  ElseIf ConsisVal = 1 Then ' NUMEROS ENTEROS
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And wsKeyAscii <> 32 Then
          wsKeyAscii = 0
          Beep
        End If
      End If
  End If

End Sub

Public Sub llena_numfac()
Dim FLAG As Integer
Dim CONTADOR As Integer

PUB_FBG = FORMGEN.i_fbg.Text

If FORMGEN.i_fbg.Visible = False Then
   PUB_FBG = ""
   GoTo SALTA
End If

If PUB_TIPMOV = 10 Then
If LK_FLAG_FACTURACION = "V" Then
   Select Case PUB_FBG
   Case "G"
       FORMGEN.i_numser.Text = ven_llave!VEM_SERIE_G
   Case "B"
       FORMGEN.i_numser.Text = ven_llave!VEM_SERIE_B
   Case "F"
       FORMGEN.i_numser.Text = ven_llave!VEM_SERIE_F
   Case "P"
       FORMGEN.i_numser.Text = ven_llave!VEM_SERIE_P
   End Select
ElseIf LK_FLAG_FACTURACION = "A" Then
   Select Case PUB_FBG
   Case "G"
       FORMGEN.i_numser.Text = par_llave!PAR_G_SERIE
   Case "B"
       FORMGEN.i_numser.Text = par_llave!PAR_B_SERIE
   Case "F"
       FORMGEN.i_numser.Text = par_llave!PAR_F_SERIE
   End Select
ElseIf LK_FLAG_FACTURACION = "U" Then
   Select Case PUB_FBG
   Case "B"
       FORMGEN.i_numser.Text = usu_llave!USU_SERIE_B
   Case "F"
       FORMGEN.i_numser.Text = usu_llave!USU_SERIE_F
   End Select
End If
End If


SALTA:
If PUB_TIPMOV = 20 Then PUB_FBG = ""

If PUB_TIPMOV = 97 And LK_CODTRA <> 2410 Then
   PUB_FBG = "N"
   If LK_FLAG_FACTURACION = "V" Then
     If i_fbg.Text = "F" Then
        FORMGEN.i_numser.Text = Nulo_Valor0(ven_llave!VEM_SERIE_N)
     Else
       FORMGEN.i_numser.Text = "1"
     End If
   Else
      FORMGEN.i_numser.Text = par_llave!PAR_SERIE_NCRE
   End If
End If
If PUB_TIPMOV = 97 And PUB_CP = "P" Then
   PUB_FBG = "C"
   FORMGEN.i_numser.Text = 0
End If

If PUB_TIPMOV = 98 Then
   PUB_FBG = "D"
   If LK_FLAG_FACTURACION = "V" Then
      FORMGEN.i_numser.Text = Nulo_Valor0(ven_llave!VEM_SERIE_D)
   Else
     FORMGEN.i_numser.Text = Nulo_Valor0(par_llave!PAR_SERIE_NDEB)
   End If
End If
If PUB_TIPMOV = 98 And PUB_CP = "P" Then
   PUB_FBG = "A"
   FORMGEN.i_numser.Text = 0
End If

If LK_FLAG_FACTURACION = "U" Then
   Select Case PUB_FBG
   Case "N"
       FORMGEN.i_numser.Text = usu_llave!USU_SERIE_NC
   Case "D"
       FORMGEN.i_numser.Text = usu_llave!USU_SERIE_ND
   End Select
End If

If LK_CODTRA = 1122 Or LK_CODTRA = 1111 Then
   PUB_FBG = "X"
  FORMGEN.i_numser.Text = ""
End If

If PUB_TIPMOV = 99 Then PUB_FBG = "K"

SQ_OPER = 3
PU_TIPMOV = PUB_TIPMOV
pu_codcia = LK_CODCIA
PU_NUMSER = Val(FORMGEN.i_numser.Text)
PU_FBG = PUB_FBG
LEER_FAR_LLAVE
PU_NUMFAC = 1
If Not far_menor.EOF Then
   PU_NUMFAC = Val(far_menor!far_numfac) + 1
End If
If PUB_TIPMOV = 99 Then GoTo fin

 If FORMGEN.i_fbg.Visible = False Then
   If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
   Else
      GoTo fin
   End If
End If

If LK_FLAG_FACTURACION = "V" Then
   If Not ven_llave.EOF Then
   If (ven_llave!VEM_FLAG_G = "A" Or ven_llave!VEM_FLAG_B = "A" Or ven_llave!VEM_FLAG_F = "A" Or ven_llave!VEM_FLAG_P = "A" Or ven_llave!VEM_FLAG_R = "A") Then
   Select Case PUB_FBG
   Case "G"
        If ven_llave!VEM_FLAG_G = "A" Then
           PU_NUMFAC = ven_llave!VEM_NUMFAC_G_INI
        End If
   Case "B"
        If ven_llave!VEM_FLAG_B = "A" Then
           PU_NUMFAC = ven_llave!VEM_NUMFAC_B_INI
        End If
   Case "F"
        If ven_llave!VEM_FLAG_F = "A" Then
           PU_NUMFAC = ven_llave!VEM_NUMFAC_F_INI
        End If
   Case "P"
        If ven_llave!VEM_FLAG_P = "A" Then
           PU_NUMFAC = ven_llave!VEM_NUMFAC_P_INI
        End If
   End Select
 End If
 End If
 End If
 
If LK_FLAG_FACTURACION = "A" And (par_llave!PAR_FLAG_G = "A" Or par_llave!PAR_FLAG_B = "A" Or par_llave!PAR_FLAG_F = "A") Then
   Select Case PUB_FBG
   Case "G"
       If par_llave!PAR_FLAG_G = "A" Then PU_NUMFAC = par_llave!PAR_G_INICIAL
   Case "B"
       If par_llave!PAR_FLAG_B = "A" Then PU_NUMFAC = par_llave!PAR_B_INICIAL
   Case "F"
       If par_llave!PAR_FLAG_F = "A" Then PU_NUMFAC = par_llave!PAR_F_INICIAL
   End Select
End If

NOTACRE:
If PUB_TIPMOV = 97 And PUB_CP = "C" Then
 If par_llave!PAR_FLAG_NCRE = "A" Then
   FORMGEN.i_numser.Text = par_llave!PAR_SERIE_NCRE
   PU_NUMFAC = par_llave!PAR_numfac_ncre
   PUB_FBG = "N"
 End If
End If

If PUB_TIPMOV = 98 Then
 If par_llave!PAR_FLAG_NDEB = "A" Then
  ' FORMGEN.i_numser.Text = par_llave!PAR_SERIE_NDEB
  ' PU_NUMFAC = par_llave!PAR_numfac_nDEB
  ' PUB_FBG = "D"
 End If
End If
      
fin:

PUB_NUMSER = Val(i_numser.Text)
i_numfac.Text = PU_NUMFAC
PUB_NUMFAC = Val(PU_NUMFAC)
pub_numfac_ult = PUB_NUMFAC
If Trim(PUB_FBG) = "" Then
   BolFac.Caption = "Guia: "
Else
   BolFac.Caption = PUB_FBG & "/. "
End If

End Sub


Public Sub LLENA_CAMPOS2()
FORMGEN.lmoneda.Visible = False
FORMGEN.lmoneda.Visible = False
Dim CONTA
Dim ws_codcia As String * 2
Dim wxtabindex
Dim enlace
Dim pos1, CARAC
Dim cias As String
Dim indice
FORMGEN.i_TEXTONCRE.Visible = False
FORMGEN.i_numfac.Text = ""
FORMGEN.i_numser.Text = ""

nn = 2
m_ind = 0
wxtabindex = 2
Do Until Val(tra_llave(nn)) = 0 Or nn = 62

m_ind = m_ind + 1
If FORMGEN.LABELGEN(m_ind) <> "0" Then
   FORMGEN.LABELGEN(m_ind).Visible = True
   If tra_llave(nn + 2) > 1200 And tra_llave(nn + 3) > 7000 Then
      FORMGEN.LABELGEN(m_ind).Top = tra_llave(nn + 2) - 200
      FORMGEN.LABELGEN(m_ind).Left = tra_llave(nn + 3)
   Else
      FORMGEN.LABELGEN(m_ind).Top = tra_llave(nn + 2)
      FORMGEN.LABELGEN(m_ind).Left = tra_llave(nn + 3) - 1000
   End If
   FORMGEN.LABELGEN(m_ind).Caption = tra_llave(nn + 1)
End If

indice = TABLA_TAG(tra_llave(nn))
If IsEmpty(indice) Then
   GoTo enddo
End If


   If TypeOf FORMGEN.Controls(indice) Is MSFlexGrid Then
      FORMGEN.LABELGEN(m_ind).Caption = ""
      pos1 = InStr(1, tra_llave(nn + 1), ",", 1)
      pos1 = pos1 - 1
      If pos1 = -1 Then
         pos1 = 0
      End If
      CARAC = Left(tra_llave(nn + 1), pos1)
      FORMGEN.Controls(indice).Width = Val(CARAC)
      pos1 = pos1 + 2
      CARAC = Mid(tra_llave(nn + 1), pos1, 10)
      DoEvents
      FORMGEN.Controls(indice).Height = Val(CARAC)
    End If

    
If TypeOf FORMGEN.Controls(indice) Is label Then
    enlace = 1
Else
    enlace = 0
End If
    FORMGEN.Controls(indice).Top = tra_llave(nn + 2)
    FORMGEN.Controls(indice).Left = tra_llave(nn + 3) + 1300 * enlace
    FORMGEN.Controls(indice).Visible = True
    If TypeOf FORMGEN.Controls(indice) Is MSFlexGrid Then
       FORMGEN.Controls(indice).Clear
    End If
   
    
    FORMGEN.Controls(indice).Enabled = True
    If tra_llave(106 + m_ind) = -1 Then
       FORMGEN.Controls(indice).Enabled = False
    End If
    If tra_llave(106 + m_ind) = -2 Then
       FORMGEN.Controls(indice).Locked = True
    End If
    
    
    
    
    If FORMGEN.Controls(indice).Tag = 4 Then
       FORMGEN.i_nomCLI.Top = tra_llave(nn + 2)
       FORMGEN.i_nomCLI.Left = tra_llave(nn + 3) + 1300 * 1
       FORMGEN.i_nomCLI.Visible = True
       FORMGEN.i_nomCLI.Caption = ""
       FORMGEN.i_nomCLI.Width = 3000
    End If
    If FORMGEN.Controls(indice).Tag = 41 Then
       FORMGEN.i_nomven.Top = tra_llave(nn + 2)
       FORMGEN.i_nomven.Left = tra_llave(nn + 3) + 1300 * 1
       FORMGEN.i_nomven.Visible = True
       FORMGEN.i_nomven.Caption = ""
       FORMGEN.i_nomven.Width = 3000
    End If
    If FORMGEN.Controls(indice).Tag = 7 Then
       FORMGEN.LMONEDAB.Top = tra_llave(nn + 2)
       FORMGEN.LMONEDAB.Left = tra_llave(nn + 3) - 300
       FORMGEN.LMONEDAB.Visible = True
       FORMGEN.LMONEDAB.Caption = ""
    End If
    If FORMGEN.Controls(indice).Tag = 6 Then
       FORMGEN.lmoneda.Top = tra_llave(nn + 2)
       FORMGEN.lmoneda.Left = tra_llave(nn + 3) - 300
       FORMGEN.lmoneda.Visible = True
       FORMGEN.lmoneda.Caption = ""
    End If
    
    If FORMGEN.Controls(indice).Tag = 19 Then
       FORMGEN.i_nomban.Top = tra_llave(nn + 2)
       FORMGEN.i_nomban.Left = tra_llave(nn + 3) + 1300 * 1
       FORMGEN.i_nomban.Visible = True
       FORMGEN.i_nomban.Caption = ""
       FORMGEN.i_nomban.Width = 4500
    End If
    If FORMGEN.Controls(indice).Tag = 70 Then
       FORMGEN.i_nomban2.Top = tra_llave(nn + 2)
       FORMGEN.i_nomban2.Left = tra_llave(nn + 3) + 1300 * 1
       FORMGEN.i_nomban2.Visible = True
       FORMGEN.i_nomban2.Caption = ""
       FORMGEN.i_nomban2.Width = 4500
    End If
    If FORMGEN.Controls(indice).Tag = 5 Then
       FORMGEN.i_nomart.Top = tra_llave(nn + 2)
       FORMGEN.i_nomart.Left = tra_llave(nn + 3) + 1300 * 1
       FORMGEN.i_nomart.Visible = True
       FORMGEN.i_nomart.Caption = ""
       FORMGEN.i_nomart.Width = 3000
    End If
    
    If FORMGEN.Controls(indice).Tag = 200 Then
       FORMGEN.lblfecalma.Top = tra_llave(nn + 2)
       FORMGEN.lblfecalma.Left = tra_llave(nn + 3) - 1300 * 1
       FORMGEN.lblfecalma.Visible = True
       FORMGEN.lblfecalma.Caption = "F. Almacen"
       'FORMGEN.lblcabe.Width = 3000
       
    End If
    
    If enlace = 0 Then
       FORMGEN.Controls(indice).WhatsThisHelpID = m_ind
       wxtabindex = wxtabindex + 1
'       MsgBox FORMGEN.Controls(indice).Name
       FORMGEN.Controls(indice).TabIndex = wxtabindex
       If Trim(Nulo_Valors(tra_llave(106 + m_ind))) = Null Then
          tab_avanza(m_ind) = 0
       Else
       tab_avanza(m_ind) = tra_llave(106 + m_ind)
       End If
    Else
       FORMGEN.Controls(indice).Width = 3200
    End If
    If TypeOf FORMGEN.Controls(indice) Is TextBox Then
        FORMGEN.Controls(indice).Text = ""
    End If
enddo:
nn = nn + 4
Loop

'BolFac.Top = grabar.Top
BolFac.Left = 7160 'BolFac.Left + 200
i_numfac.Left = BolFac.Left + 800
i_numfac.Top = BolFac.Top
i_numser.Left = BolFac.Left + 400
i_numser.Top = BolFac.Top
'i_cambio.Top = grabar.Top
'i_cambio.Left = i_numfac.Left + 700

 ' AGREGE PARA LA TRANS 2402 EN EL FORGEN
If LK_CODTRA = 2401 Then
   i_fecha_compra.Top = grabar.Top
   i_fecha_compra.Left = i_cambio.Left + 900
End If
i_fecha_compra.Text = Format(LK_FECHA_DIA, "dd/mm/yy")

BolFac.Visible = True
i_numfac.Visible = True
i_numser.Visible = True


FORMGEN.grid_canje.Visible = False
i_ds.Clear
If LK_CODTRA = 2401 Or LK_CODTRA = 2410 Or LK_CODTRA = 2412 Then
   If LK_MONEDA = "S" Then
      i_ds.AddItem "S"
   ElseIf LK_MONEDA = "D" Then
      i_ds.AddItem "D"
   Else
      i_ds.AddItem "S"
      i_ds.AddItem "D"
   End If
Else
   If LK_MONEDA = "S" Then
      i_ds.AddItem "S"
   ElseIf LK_MONEDA = "D" Then
      i_ds.AddItem "D"
   Else
      i_ds.AddItem "S"
      i_ds.AddItem "D"
   End If
End If
i_ds.ListIndex = 0

If FORMGEN.Frame4.Visible = True Then
   FORMGEN.grid_fac.Clear
   FORMGEN.i_moneda.Visible = True
   fila = 0
End If

If FORMGEN.i_situacion.Visible = True Then
'FORMGEN.i_situacion.Clear
'PUB_TIPREG = 133
'PUB_CODCIA = "00"
'SQ_OPER = 2
'LEER_TAB_LLAVE
'If tab_mayor.EOF Then
   'MsgBox "LLenar situacion de letras en TIPREG=133 "
'End If
'Do Until tab_mayor.EOF
  'DoEvents
 ' FORMGEN.i_situacion.AddItem Nulo_Valors(tab_mayor!tab_NOMLARGO)
 ' tab_mayor.MoveNext
'Loop
End If





   FORMGEN.i_subtotal.Text = ""
   FORMGEN.i_gastos.Text = ""
   FORMGEN.i_descto.Text = ""
   FORMGEN.i_impto.Text = ""
   FORMGEN.i_neto.Text = ""
   FORMGEN.i_fecha_compra.Text = Format(LK_FECHA_DIA, "dd/mm/yy")



If FORMGEN.i_CODUSU.Visible = True Then
   MUESTRA_USUario
'   FORMGEN.i_CODUSU.ListIndex = 0
End If

If FORMGEN.Frame1.Visible = True Then
   FORMGEN.Frame1.Width = 4000
   FORMGEN.Frame1.Height = 3000
   FORMGEN.BLOQ(0).Width = 3500
   FORMGEN.BLOQ(1).Width = 3500
   FORMGEN.BLOQ(2).Width = 3500
   FORMGEN.BLOQ(3).Width = 3500
   CONTA = 0
   PUB_TIPREG = 199
   PUB_CODCIA = LK_CODCIA
   SQ_OPER = 2
   LEER_TAB_LLAVE
   If tab_mayor.EOF Then MsgBox "Falta llenar en Tablas del sistema los tipos de bloqueos , Tipo : " & PUB_TIPREG
   Do Until tab_mayor.EOF Or CONTA = 4
        FORMGEN.BLOQ(CONTA).Caption = Trim(tab_mayor!tab_nomlargo)
        FORMGEN.BLOQ(CONTA).Visible = True
        CONTA = CONTA + 1
        tab_mayor.MoveNext
   Loop
End If

If LK_CODTRA = 2406 Or LK_CODTRA = 2401 Or LK_CODTRA = 2003 Then
   'FORMGEN.i_image_llave.Visible = True
Else
   'FORMGEN.i_image_llave.Visible = False
End If

If FORMGEN.i_cias.Visible = True Then
   If Trim(par_llave!PAR_CIAS) <> "" Then
      cias = par_llave!PAR_CIAS
      nn = 1
    For m_ind = 1 To 15
        ws_codcia = Mid(cias, nn, 2)
        If Trim(ws_codcia) = "" Then Exit For
           SQ_OPER = 1
           PUB_CODCIA = ws_codcia
           LEER_PAR_LLAVE
           FORMGEN.i_cias.AddItem par_llave!PAR_NOMBRE & "               " & par_llave!PAR_CODCIA
        nn = nn + 2
    Next m_ind
   End If
   SQ_OPER = 1
   PUB_CODCIA = LK_CODCIA
   LEER_PAR_LLAVE
End If



If FORMGEN.i_tipdoc.Visible = True Then
FORMGEN.i_tipdoc.Clear
PUB_TIPREG = 8
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
If tab_mayor.EOF Then
   MsgBox "LLenar tabla de tipo de documentos...LE, CH,ETC"
End If
Do Until tab_mayor.EOF
  DoEvents
  FORMGEN.i_tipdoc.AddItem Nulo_Valors(tab_mayor!tab_nomcorto)
  tab_mayor.MoveNext
Loop
If FORMGEN.i_tipdoc.ListCount > 0 Then FORMGEN.i_tipdoc.ListIndex = 0


End If


If FORMGEN.i_cias.Visible Then
 If FORMGEN.i_cias.ListCount = 0 Then
   MsgBox " Debe Configurar la Compañia de Destino..", 48, Pub_Titulo
 Else
   FORMGEN.i_cias.ListIndex = 0
 End If
End If
If LK_CODTRA = 1111 Then FORMGEN.i_fecha_vcto.Text = LK_FECHA_DIA


Inicio_de_Todo

fin:


End Sub
Public Sub llena_numGUIA()
Dim FLAG As Integer
Dim CONTADOR As Integer



If LK_FLAG_FACTURACION = "V" Then
  If ven_llave.EOF Then
    MsgBox "Falta Indicar el Vendedor..."
    Exit Sub
  Else
   If ven_llave!VEM_CODVEN <> Val(i_codven.Text) Then
      MsgBox "Falta Indicar el Vendedor..."
      Exit Sub
   End If
  End If
   i_serguia.Text = ven_llave!VEM_SERIE_R
Else
   i_serguia.Text = Nulo_Valor0(par_llave!PAR_R_SERIE)
End If

If LK_CODTRA = 1122 Or LK_CODTRA = 1111 Then
  i_serguia.Text = ""
End If

PUB_SERGUIA = Val(i_serguia.Text)
SQ_OPER = 7
pu_codcia = LK_CODCIA
LEER_FAR_LLAVE
PU_NUMFAC = 1
If Not far_guiam.EOF Then
   PU_NUMFAC = Val(far_guiam!far_numguia) + 1
End If


If LK_FLAG_FACTURACION = "V" Then
   If Not ven_llave.EOF Then
   If ven_llave!VEM_FLAG_R = "A" Then
        If ven_llave!VEM_FLAG_R = "A" Then PU_NUMFAC = ven_llave!VEM_NUMFAC_R_INI
   End If
   End If
End If
 
If LK_FLAG_FACTURACION = "A" And par_llave!PAR_FLAG_R = "A" Then
   PU_NUMFAC = par_llave!PAR_R_INICIAL
End If

      
fin:
i_numguia.Text = PU_NUMFAC


End Sub

Private Sub Consistencias_rich(wsGrid As MSFlexGrid, wsTexto As RichTextBox, wsKeyAscii As Integer, Optional ConsisVal, Optional ConsisCol)
  Static VALOR
  Dim car As String
  If ConsisVal = 2 Then ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
'      Flag_Consis = "A"
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
  ElseIf ConsisVal = 1 Then ' NUMEROS ENTEROS
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And wsKeyAscii <> 32 Then
          wsKeyAscii = 0
          Beep
        End If
      End If
  End If

End Sub
Public Sub PROCESA_GRIDL()
Dim WS_SIGNO As Integer
Dim WS_TIPDOC As String * 2
Dim NUMCAMPO As Integer
Dim TOTAL_DEUDA, TOTAL_GEN  As Variant
TOTAL_DEUDA = 0
TOTAL_GEN = 0
Dim Tit As String
FORMGEN.gridl.Visible = True
FORMGEN.gridl.Clear
pasa_cabeza_L
grid_canje.Clear
pasa_cabeza_canje

If Val(FORMGEN.i_codcli.Text) = 0 Then Exit Sub

If cli_llave.EOF Then Exit Sub


pu_codcia = LK_CODCIA
pu_codclie = Val(FORMGEN.i_codcli.Text)
SQ_OPER = 2
LEER_CAR_LLAVE

fila = 1

FORMGEN.gridl.TextMatrix(2, 15) = cli_llave!CLI_CODCLIE
FORMGEN.gridl.TextMatrix(2, 0) = cli_llave!cli_nombre

If LK_CODTRA = 1455 Then
   If Trim(lmoneda.Caption) = "" Then
      MsgBox "Selecciona Moneda"
      Exit Sub
   End If
End If



If car_mayor.EOF = True Then
   FORMGEN.gridl.Row = 1
   FORMGEN.gridl.COL = 1
   FORMGEN.gridl.Text = "No hay registros"
End If
Do Until car_mayor.EOF
    Pub_Respuesta = vbYes
    If LK_FLAG_SOS = "A" And car_mayor!CAR_FLAG_SO <> "A" Then
      GoTo pasa
    End If
    If LK_CODTRA = 2412 Then
     If Trim(FORMGEN.i_fbg.Text) = Trim(car_mayor!car_FBG) And Val(FORMGEN.i_numser_c) = Val(car_mayor!car_NUMSER) And Val(FORMGEN.i_numfac_c) = Val(car_mayor!car_NUMFAC) Then
     Else
       GoTo pasa
     End If
    End If
    If car_mayor!car_importe = 0 Then
       Pub_Respuesta = vbNo
    End If
                                            
    If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
    Else
    If car_mayor!car_TIPDOC <> PUB_TIPDOC And Trim(PUB_TIPDOC) <> "" Then
       Pub_Respuesta = vbNo
    End If
    End If
    
    
    If Pub_Respuesta = vbYes Then
       fila = fila + 1
       FORMGEN.gridl.Rows = fila + 1
       FORMGEN.grid_canje.Rows = fila + 1
       
       FORMGEN.gridl.RowHeight(fila) = 350
       FORMGEN.gridl.TextMatrix(fila, 15) = cli_llave!CLI_CODCLIE
       FORMGEN.gridl.TextMatrix(fila, 0) = "'" & Format(car_mayor!car_fecha_sunat, "dd/mm/yy") & " " & Trim(cli_llave!cli_nombre)
       If LK_CODTRA = 2725 Then FORMGEN.gridl.TextMatrix(fila, 0) = Trim(Nulo_Valors(car_mayor!car_concepto)) & "'" & Format(car_mayor!car_fecha_sunat, "dd/mm/yy") & " " & Trim(cli_llave!cli_nombre)
       FORMGEN.gridl.TextMatrix(fila, 13) = car_mayor!car_TIPDOC
       FORMGEN.gridl.TextMatrix(fila, 1) = Nulo_Valors(car_mayor!car_FBG)
       If Trim(FORMGEN.gridl.TextMatrix(fila, 1)) = "" Then
          If Nulo_Valors(car_mayor!car_TIPDOC) = "LE" Then
             FORMGEN.gridl.TextMatrix(fila, 1) = Nulo_Valors(car_mayor!car_TIPDOC) & "-" & Nulo_Valors(car_mayor!CAR_SITUACION)
          Else
             FORMGEN.gridl.TextMatrix(fila, 1) = Nulo_Valors(car_mayor!car_TIPDOC)
          End If
       End If
       
       If Nulo_Valors(car_mayor!car_FBG) = "N" Then FORMGEN.gridl.TextMatrix(fila, 1) = "N.Cred." & car_mayor!car_NUMFAC
       If Nulo_Valors(car_mayor!car_FBG) = "D" Then FORMGEN.gridl.TextMatrix(fila, 1) = "N.Deb ." & car_mayor!car_NUMFAC
       If Nulo_Valors(car_mayor!car_TIPDOC) = "CH" Then FORMGEN.gridl.TextMatrix(fila, 1) = "Chq." & car_mayor!car_NUM_CHEQUE
             
       If cli_llave!CLI_CP = "P" Then
          FORMGEN.gridl.TextMatrix(fila, 2) = Nulo_Valor0(car_mayor!car_numser_c)
          FORMGEN.gridl.TextMatrix(fila, 3) = Nulo_Valor0(car_mayor!car_numfac_c)
       Else
          FORMGEN.gridl.TextMatrix(fila, 2) = Nulo_Valor0(car_mayor!car_NUMSER)
          FORMGEN.gridl.TextMatrix(fila, 3) = Nulo_Valor0(car_mayor!car_NUMFAC)
          If car_mayor!CAR_TIPMOV = 0 And Val(Nulo_Valor0(car_mayor!car_NUM_CHEQUE)) <> 0 Then
             FORMGEN.gridl.TextMatrix(fila, 3) = Nulo_Valor0(car_mayor!car_NUM_CHEQUE)
          End If
          If Val(FORMGEN.gridl.TextMatrix(fila, 3)) = 0 Then
             FORMGEN.gridl.TextMatrix(fila, 3) = Nulo_Valor0(car_mayor!car_NUMDOC)
          End If
       End If
       
       If car_mayor!CAR_MONEDA = "S" Then
          FORMGEN.gridl.TextMatrix(fila, 4) = "S/." & car_mayor!car_importe
       Else
          FORMGEN.gridl.TextMatrix(fila, 4) = "  $" & car_mayor!car_importe
       End If
       FORMGEN.gridl.TextMatrix(fila, 5) = ""
       If LK_CODTRA = 1455 Then FORMGEN.gridl.TextMatrix(fila, 5) = car_mayor!car_importe
         
       FORMGEN.gridl.TextMatrix(fila, 6) = Format(car_mayor!car_fecha_vcto, "dd/mm/yy")
       FORMGEN.gridl.TextMatrix(fila, 7) = Nulo_Valor0(car_mayor!CAR_COBRADOR) ' Nulo_Valors(car_mayor!car_concepto)
       FORMGEN.gridl.TextMatrix(fila, 9) = car_mayor!CAR_IMP_INI
       FORMGEN.gridl.TextMatrix(fila, 8) = Format(car_mayor!car_fecha_VCTO_orig, "dd/mm/yy")
       FORMGEN.gridl.TextMatrix(fila, 10) = car_mayor!CAR_cp
       FORMGEN.gridl.TextMatrix(fila, 11) = car_mayor!car_SERDOC
       FORMGEN.gridl.TextMatrix(fila, 12) = car_mayor!car_NUMDOC
       FORMGEN.gridl.TextMatrix(fila, 16) = car_mayor!car_NUMGUIA
       FORMGEN.gridl.TextMatrix(fila, 17) = Nulo_Valor0(car_mayor!CAR_SITUACION)
       FORMGEN.gridl.TextMatrix(fila, 24) = Nulo_Valors(car_mayor!CAR_MONEDA)
       FORMGEN.gridl.TextMatrix(fila, 25) = car_mayor!car_importe
       FORMGEN.gridl.TextMatrix(fila, 26) = car_mayor!car_fecha_vcto
       
       TOTAL_DEUDA = car_mayor!car_importe + TOTAL_DEUDA
    End If
    WS_TIPDOC = car_mayor!car_TIPDOC
pasa:
    car_mayor.MoveNext
Loop
     If LK_CODTRA <> 1455 Then
        FORMGEN.gridl.SetFocus
        FORMGEN.textovarl.Width = 30
        FORMGEN.textovarl.Height = 30
     End If
     
   If FORMGEN.gridl.Visible = False Then FORMGEN.gridl.Visible = True
   If FORMGEN.textovarl.Visible = False Then FORMGEN.textovarl.Visible = True

   
   If LK_CODTRA = 2770 Then
      FORMGEN.grid_canje.ColWidth(0) = 1500
   Else
      FORMGEN.grid_canje.ColWidth(0) = 0
   End If
   
i_bancos.Visible = False
FORMGEN.grid_canje.Rows = gridl.Rows
fila = 2
Do Until fila = gridl.Rows
   FORMGEN.grid_canje.RowHeight(fila) = 350
   fila = fila + 1
Loop
If LK_CODTRA = 2410 Or LK_CODTRA = 2412 Or LK_CODTRA = 2418 Then grid_canje.Visible = False
'Print WS_MONEDA_CLI
calcula_totales2
SendKeys "{UP}"
End Sub



Public Sub procesa_celdas()
TEXTOVAR.MaxLength = 0
If grid_fac.Text = "" Then flag_salto = 0
If flag_salto = 1 Then Exit Sub

If grid_fac.COL > 10 And grid_fac.COL <> 12 Then Exit Sub

If grid_fac.COL = 10 And (PUB_TIPMOV = 93) Then Exit Sub

If grid_fac.COL = 10 And Val(SUT_LLAVE!SUT_DESCTO) <> 0 Then Exit Sub

If grid_fac.COL = 8 And PUB_TIPMOV = 10 Then Exit Sub


If Not SUT_LLAVE.EOF Then
   If grid_fac.COL = 6 And Val(Nulo_Valor0(SUT_LLAVE!SUT_PRECIO)) = 1 Then Exit Sub
End If

calcula_totales
If grid_fac.COL = 9 Then
    TEXTOVAR.Visible = False
    i_mortal.Left = Frame4.Left + grid_fac.CellLeft
    i_mortal.Width = grid_fac.CellWidth
    i_mortal.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
    i_mortal.Visible = True
    TEXTOVAR.Visible = False
    i_unidades.Visible = False
    i_mortal.SetFocus
ElseIf grid_fac.COL = 12 Then
    If LK_CODTRA = 2414 Then Exit Sub
    TEXTOVAR.Visible = False
    i_is.Left = Frame4.Left + grid_fac.CellLeft
    i_is.Width = grid_fac.CellWidth
    i_is.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
    If Val(grid_fac.TextMatrix(grid_fac.Row, 12)) = -1 Then
     i_is.ListIndex = 1
    ElseIf Val(grid_fac.TextMatrix(grid_fac.Row, 12)) = 1 Then
     i_is.ListIndex = 0
    End If
    i_is.Visible = True
    TEXTOVAR.Visible = False
    i_mortal.Visible = False
    i_unidades.Visible = False
    i_is.SetFocus
ElseIf grid_fac.COL = 5 And grid_fac.TextMatrix(grid_fac.Row, 16) <> "" Then    ' AND Val(Nulo_Valor0(SUT_LLAVE!SUT_precio)) = 3 Then
    i_mortal.Visible = False
    i_precios.Visible = False
    i_is.Visible = False
    TEXTOVAR.Visible = False
    i_unidades.Left = Frame4.Left + grid_fac.CellLeft
    i_unidades.Width = grid_fac.CellWidth + 200
    i_unidades.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
    i_unidades.Visible = True
    i_unidades.SetFocus
    SendKeys "%{DOWN}"
ElseIf grid_fac.COL = 6 And Val(Nulo_Valor0(SUT_LLAVE!SUT_PRECIO)) = 3 And grid_fac.TextMatrix(grid_fac.Row, 16) <> "" Then
    i_mortal.Visible = False
    i_is.Visible = False
    i_unidades.Visible = False
    TEXTOVAR.Visible = False
    i_precios.Left = Frame4.Left + grid_fac.CellLeft
    i_precios.Width = grid_fac.CellWidth + 800
    i_precios.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
     ' ICA
    grid_fac.TextMatrix(grid_fac.Row, 10) = ""
    i_precios.Visible = True
    i_precios.SetFocus
ElseIf grid_fac.COL = 1 Or grid_fac.COL = 2 Or grid_fac.COL = 3 Or grid_fac.COL = 4 Or grid_fac.COL = 8 Or grid_fac.COL = 10 Or (grid_fac.COL = 6 And (Val(Nulo_Valor0(SUT_LLAVE!SUT_PRECIO)) = 2 Or Val(Nulo_Valor0(SUT_LLAVE!SUT_PRECIO)) = 4)) Then
     If Val(FORMGEN.i_codcli.Text) <> 0 And Not cli_llave.EOF And grid_fac.COL = 10 And PUB_TIPMOV = 10 And Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) <> "" And grid_fac.Row = grid_fac.Rows - 1 Then
     If cli_llave!CLI_CODCLIE = Val(FORMGEN.i_codcli.Text) Then
        If Left(SUT_LLAVE!SUT_DESCTO, 1) = "P" And Nulo_Valor0(cli_llave!CLI_PORDESCTO) <> 0 Then TEXTOVAR.Text = cli_llave!CLI_PORDESCTO
     End If
     End If
    TEXTOVAR.Left = Frame4.Left + grid_fac.CellLeft
    TEXTOVAR.Width = grid_fac.CellWidth
    TEXTOVAR.Height = grid_fac.CellHeight
    TEXTOVAR.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
    TEXTOVAR.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
    textovar_bak = TEXTOVAR.Text
    TEXTOVAR.DataField = grid_fac.COL
    TEXTOVAR.Tag = grid_fac.Row
    TEXTOVAR.Visible = True
    i_mortal.Visible = False
    i_is.Visible = False
    i_precios.Visible = False
    i_unidades.Visible = False
    Azul TEXTOVAR, TEXTOVAR
ElseIf (grid_fac.COL = 7 And PUB_TIPMOV = 20) Or (SUT_LLAVE!SUT_PRECIO = "2" And PUB_TIPMOV = 10) Or (SUT_LLAVE!SUT_PRECIO = "2" And PUB_TIPMOV = 101) Then
    TEXTOVAR.Left = Frame4.Left + grid_fac.CellLeft
    TEXTOVAR.Width = grid_fac.CellWidth
    TEXTOVAR.Height = grid_fac.CellHeight
    TEXTOVAR.Top = Frame4.Top + grid_fac.Top + grid_fac.CellTop
    TEXTOVAR.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
    textovar_bak = TEXTOVAR.Text
    TEXTOVAR.DataField = grid_fac.COL
    TEXTOVAR.Tag = grid_fac.Row
    TEXTOVAR.Visible = True
    i_mortal.Visible = False
    i_is.Visible = False
    i_precios.Visible = False
    i_unidades.Visible = False
    Azul TEXTOVAR, TEXTOVAR
End If

End Sub

Private Sub procesa_celdas_l()
Exit Sub

If Nulo_Valor0(SUT_LLAVE!SUT_SIGNO_CAR) = 0 Then Exit Sub

If LK_CODTRA <> 2412 And LK_CODTRA <> 2410 Then
If Len(Trim(SUT_LLAVE!sut_TIPDOC)) <> 0 And Trim(gridl.TextMatrix(gridl.Row, 13)) <> "" Then
If Trim(SUT_LLAVE!sut_TIPDOC) <> Trim(gridl.TextMatrix(gridl.Row, 13)) Then
   MsgBox "Tipo de documento no corresponde ..."
   Exit Sub
End If
End If
End If

If gridl.COL = 21 Then
    i_bancos.Left = gridl.Left + gridl.CellLeft
    i_bancos.Width = gridl.CellWidth
    i_bancos.Top = gridl.Top + gridl.CellTop
    textovarl.Width = 30
    textovarl.Height = 30
    i_bancos.Visible = True
    i_bancos_Click
Else
    gridl.RowHeight(gridl.Row) = 350
    textovarl.Left = gridl.Left + gridl.CellLeft
    textovarl.Width = gridl.CellWidth
    textovarl.Height = gridl.CellHeight
    textovarl.Top = Frame4.Top + gridl.Top + gridl.CellTop
    textovarl.Top = gridl.Top + gridl.CellTop
    textovarl.Text = gridl.TextMatrix(gridl.Row, gridl.COL)
    textovar_bak = textovarl.Text
    textovarl.DataField = gridl.COL
    textovarl.Tag = gridl.Row
    If textovarl.Visible = False Then textovarl.Visible = True
    textovarl.SetFocus
End If

    
    

End Sub

Private Sub pasa_cabeza_canje()

   grid_canje.TextMatrix(0, 1) = "Orden"
   grid_canje.TextMatrix(0, 2) = "TipoDoc"
   grid_canje.TextMatrix(0, 3) = "Importe"
   grid_canje.TextMatrix(0, 4) = "Fecha "
   grid_canje.TextMatrix(0, 5) = "Banco"
   grid_canje.TextMatrix(0, 6) = "N. Documento"
   grid_canje.TextMatrix(0, 7) = "Observaciones"

End Sub

Private Sub pasa_cabeza()
   grid_fac.TextMatrix(0, 0) = "Descripción del Articulo"
   grid_fac.TextMatrix(1, 0) = "Fila de Totales ===> "
   grid_fac.TextMatrix(0, 1) = "Codigo"
   grid_fac.TextMatrix(0, 4) = "Cantidad"
   grid_fac.TextMatrix(0, 5) = "Unidad"
   grid_fac.TextMatrix(0, 6) = "Precio"
   grid_fac.TextMatrix(0, 7) = "Subtotal"
   grid_fac.TextMatrix(0, 8) = "Flete"
   grid_fac.TextMatrix(0, 9) = "Causas"
   grid_fac.TextMatrix(0, 10) = "Descto"
   grid_fac.TextMatrix(0, 11) = "Costo"
   grid_fac.TextMatrix(0, 12) = "Ing/Salida"
   grid_fac.TextMatrix(0, 13) = "S t o c k"
   grid_fac.TextMatrix(0, 14) = "Equiv."
   grid_fac.TextMatrix(0, 42) = "Descto"
   grid_fac.TextMatrix(0, 34) = "Peso"


End Sub
Private Sub pasa_cabeza_grid_trans()
If LK_CODTRA = 1405 Then
   grid_trans.TextMatrix(0, 0) = "Serie"
   grid_trans.TextMatrix(0, 1) = "Numero"
   grid_trans.TextMatrix(0, 2) = "Producto"
   grid_trans.TextMatrix(0, 3) = "Unidad"
   grid_trans.TextMatrix(0, 4) = "Cant.Pend."
   grid_trans.TextMatrix(0, 5) = " "
   grid_trans.TextMatrix(0, 6) = " "
Else
   grid_trans.TextMatrix(0, 0) = "N. Documento"
   grid_trans.TextMatrix(0, 1) = "Articulo"
   grid_trans.TextMatrix(0, 2) = "Cantidad"
   grid_trans.TextMatrix(0, 3) = "Procedencia"
   grid_trans.TextMatrix(0, 4) = "Hora "
   grid_trans.TextMatrix(0, 5) = "Estado "
End If
End Sub

Private Sub pasa_cabeza_L()
   gridl.TextMatrix(0, 0) = "Nombre/Razon Social"
   gridl.TextMatrix(0, 1) = "Doc."
   gridl.TextMatrix(0, 2) = "Serie"
   gridl.TextMatrix(0, 3) = "N.Factura"
   gridl.TextMatrix(0, 4) = "Saldo"
   gridl.TextMatrix(0, 5) = "Efectivo"
   gridl.TextMatrix(0, 6) = "Nueva Fecha"
   gridl.TextMatrix(0, 7) = "Cobrador"
   gridl.TextMatrix(0, 8) = "F.V. Orig"
   gridl.TextMatrix(0, 9) = "Saldo Inicial"
   gridl.TextMatrix(0, 10) = "C/P"
   gridl.TextMatrix(0, 11) = "Serdoc"
   gridl.TextMatrix(0, 12) = "NumDoc"
   gridl.TextMatrix(0, 13) = "Tipdoc"
   gridl.TextMatrix(0, 14) = "Codcia"
   gridl.TextMatrix(0, 15) = "codclie"
   gridl.TextMatrix(0, 16) = "N. Guia"
   gridl.TextMatrix(0, 18) = "Vendedor"
   gridl.TextMatrix(0, 19) = "Imp.Cheque"
   gridl.TextMatrix(0, 20) = "Num.Cheque"
   gridl.TextMatrix(0, 21) = "Banco"
   gridl.TextMatrix(0, 22) = "Fecha de cobro"
   gridl.TextMatrix(0, 23) = "Tipo de Documento"
   gridl.TextMatrix(0, 24) = "Moneda"
    
   


End Sub
Private Sub pasa_cabeza_L2()
   gridl.TextMatrix(0, 0) = "      Nombre "
   gridl.TextMatrix(0, 1) = "FBG"
   gridl.TextMatrix(0, 2) = "Serie"
   gridl.TextMatrix(0, 3) = "N.Factura"
   gridl.TextMatrix(0, 4) = "N. Guia "
   gridl.TextMatrix(0, 5) = "  Saldo"
   gridl.TextMatrix(0, 6) = "Fecha Vcto."
   gridl.TextMatrix(0, 7) = "Concepto    "
   gridl.TextMatrix(0, 8) = "F.V. Orig"
   gridl.TextMatrix(0, 9) = "Saldo Inicial"
   gridl.TextMatrix(0, 10) = "C/P"
   gridl.TextMatrix(0, 11) = "Serdoc"
   gridl.TextMatrix(0, 12) = "NumDoc"
   gridl.TextMatrix(0, 13) = "Tipdoc"
   gridl.TextMatrix(0, 14) = "Codcia"
   gridl.TextMatrix(0, 15) = "codclie"

End Sub

Private Sub BolFac_Click()

If PUB_TIPMOV = 0 Then
  LLENA_NUM_DIARIO
  LLENA_NUM_RECIBO
End If

If PUB_TIPMOV <> 0 Then
   If LK_FLAG_FACTURACION = "V" And (PUB_TIPMOV = 10 Or PUB_TIPMOV = 97) Then
      If ven_llave.EOF Then
          MsgBox "Falta Definir el Vendedor...", 48, Pub_Titulo
          Exit Sub
      Else
          If ven_llave!VEM_CODVEN <> Val(i_codven.Text) Then
             MsgBox "Falta Definir el Vendedor...", 48, Pub_Titulo
             If i_codven.Visible And i_codven.Enabled Then i_codven.SetFocus
          Exit Sub
          End If
      End If
   End If
   llena_numfac
   
End If




i_cambio.Visible = True
i_cambio.Value = 0
i_numfac.Locked = True
i_numser.Locked = True
i_numfac.BackColor = QBColor(7)
i_numser.BackColor = QBColor(7)
If LK_CODTRA = 2401 Then i_fecha_compra.Visible = False
End Sub


Private Sub CAMBIA_MONEDA()
Exit Sub
Dim WRES
Dim fila As Integer
Dim ws_curr As Currency
fila = 2
  If Val(WS_TIPO_CAMBIO) = 0 Then WS_TIPO_CAMBIO = LK_TIPO_CAMBIO

  WRES = InputBox("Ingrese Tipo de cambio :", , WS_TIPO_CAMBIO)
  If WRES = "" Then
  Exit Sub
  End If
  If Val(WRES) <= 0 Then
    Exit Sub
  End If
  WS_TIPO_CAMBIO = WRES
  If (Abs(WS_TIPO_CAMBIO - LK_TIPO_CAMBIO) * 100 / LK_TIPO_CAMBIO) > 5 Then
     WS_TIPO_CAMBIO = 0
     MsgBox "Cambiar Tipo De Cambio... Mucha diferencia"
  End If


If pub_signo_car = 0 Then
   If Mid(Tipo_Cambio.Caption, 2, 1) = "S" Then
      Tipo_Cambio.Caption = "&Dollares " & WS_TIPO_CAMBIO
   Else
      Tipo_Cambio.Caption = "&Soles " & WS_TIPO_CAMBIO
   End If
Exit Sub
End If

Do Until fila = 999
   If Mid(Tipo_Cambio.Caption, 2, 1) = "S" Then
      gridl.TextMatrix(fila, 6) = Val(grid_fac.TextMatrix(fila, 6)) / WS_TIPO_CAMBIO
      grid_fac.TextMatrix(fila, 6) = Format(Val(grid_fac.TextMatrix(fila, 6)), "#####.0000")
   Else
   ws_curr = Val(grid_fac.TextMatrix(fila, 6)) * WS_TIPO_CAMBIO
   grid_fac.TextMatrix(fila, 6) = ws_curr
   grid_fac.TextMatrix(fila, 6) = Format(Val(grid_fac.TextMatrix(fila, 6)), "#######.0000")
   End If
   fila = fila + 1
   If fila = grid_fac.Rows Then
      fila = 999
   End If
Loop
If Mid(Tipo_Cambio.Caption, 2, 1) = "S" Then
   i_descto.Text = redondea(Val(i_descto.Text) / WS_TIPO_CAMBIO)
   i_gastos.Text = redondea(Val(i_gastos.Text) / WS_TIPO_CAMBIO)
   Tipo_Cambio.Caption = "&Dollares " & WS_TIPO_CAMBIO
Else
   i_descto.Text = redondea(Val(i_descto.Text) * WS_TIPO_CAMBIO)
   i_gastos.Text = redondea(Val(i_gastos.Text) * WS_TIPO_CAMBIO)
   Tipo_Cambio.Caption = "&Soles " & WS_TIPO_CAMBIO
End If

calcula_totales
grid_fac.SetFocus

End Sub
Private Sub calcula_totales()
Dim wdesctopor As Currency
Dim RES_DEUDA As Currency
Dim WSIMBOL As String
Dim WTC As Currency
Dim wsumadol As Currency
Dim ws_bruto_bak
Dim ww_descto2 As Currency
Dim SUB_PESO As Currency
Dim ws_ex_igv As Currency
Dim WS_MONEDA As String
Dim WS_IEV As Currency
fila = 2
SUB_PESO = 0
SUB_FLETE = 0
ww_descto2 = 0
SUB_CANT = 0
ws_bruto_bak = WS_BRUTO
WS_BRUTO = 0
SUB_UNIDAD = 0
pub_ojo = ""
If SUT_LLAVE.EOF Then Exit Sub
PUB_DESCTO = 0
PUB_DESCTO = 0
PUB_IMPTO = 0
WS_IEV = 0
wdesctopor = 0
Dim subtotal As Currency

Do Until fila = 999
   If grid_fac.Rows >= 4 Then
   If LK_CODTRA = 2414 And fila = 3 And Val(grid_fac.TextMatrix(3, 4)) > 0 Then
      grid_fac.TextMatrix(fila, 7) = grid_fac.TextMatrix(2, 7)
      grid_fac.TextMatrix(fila, 6) = Format(Val(grid_fac.TextMatrix(fila, 7)) / Val(grid_fac.TextMatrix(fila, 4)), "0.0000")
      subtotal = Val(grid_fac.TextMatrix(fila, 7))
      GoTo i2414
   End If
   End If
   
'   If (PUB_TIPMOV = 20 And Val(grid_fac.TextMatrix(fila, 4)) > 0 And grid_fac.TextMatrix(fila, 25) = "1") Or (PUB_TIPMOV = 10 And SUT_LLAVE!SUT_precio = "2" And Val(grid_fac.TextMatrix(fila, 4)) > 0) Then
   If (PUB_TIPMOV = 20 And Val(grid_fac.TextMatrix(fila, 4)) > 0 And grid_fac.TextMatrix(fila, 25) = "1") Then
      grid_fac.TextMatrix(fila, 6) = Format(Val(grid_fac.TextMatrix(fila, 7)) / Val(grid_fac.TextMatrix(fila, 4)), "0.0000")
      subtotal = Val(grid_fac.TextMatrix(fila, 7))
   Else
      subtotal = Val(grid_fac.TextMatrix(fila, 4)) * Val(grid_fac.TextMatrix(fila, 6))
      subtotal = redondea(subtotal)
      grid_fac.TextMatrix(fila, 7) = subtotal
   End If
 If PUB_TIPMOV = 10 And Val(grid_fac.TextMatrix(fila, 38)) = 1 Then
   GoTo pasa_2401
End If
   If grid_fac.TextMatrix(fila, 23) = "A" And Val(grid_fac.TextMatrix(fila, 24)) <> 0 Then
      If Not cli_llave.EOF And Val(i_codcli.Text) > 0 Then
        If PUB_TIPMOV = 10 And Val(grid_fac.TextMatrix(fila, 24)) <> 0 Then
         'ws_ex_igv = redondea(subtotal * (LK_IGV - Val(grid_fac.TextMatrix(fila, 24)))) / 100 + ws_ex_igv
         ws_ex_igv = ws_ex_igv + redondea((redondea(subtotal / (LK_IGV / 100 + 1)) - subtotal) * (Val(grid_fac.TextMatrix(fila, 24)) / 100)) ' )) / 100 + ws_ex_igv)
        Else
          'ws_ex_igv = redondea(subtotal * (LK_IGV - Val(grid_fac.TextMatrix(fila, 24)))) / 100 + ws_ex_igv
          ws_ex_igv = ws_ex_igv + Abs(redondea(subtotal * ((LK_IGV * (Val(grid_fac.TextMatrix(fila, 24)) / 100))) / 100))
        End If
      End If
   End If
   If grid_fac.TextMatrix(fila, 23) <> "A" And Val(grid_fac.TextMatrix(fila, 24)) <> 0 Then
      If Not cli_llave.EOF And Val(i_codcli.Text) > 0 Then
        If PUB_TIPMOV = 10 And Val(grid_fac.TextMatrix(fila, 24)) <> 0 Then
          WS_IEV = Format(subtotal / (1 + (Val(grid_fac.TextMatrix(fila, 24)) / 100)), "0.00")
          'ws_iev = subtotal - ws_iev
         End If
      End If
   End If
   
   If Left(SUT_LLAVE!SUT_DESCTO, 1) = "M" Then
      If PUB_TIPMOV = 10 Or PUB_TIPMOV = 93 Then
      If Val(grid_fac.TextMatrix(fila, 10)) <> 0 Then
            ww_descto2 = subtotal * Val(grid_fac.TextMatrix(fila, 10)) / 100
            ww_descto2 = redondea(ww_descto2)
            PUB_DESCTO = PUB_DESCTO + ww_descto2
      End If
      End If
   End If
   
   If Mid(SUT_LLAVE!SUT_DESCTO, 2, 1) = "M" Then
      If PUB_TIPMOV = 10 Or PUB_TIPMOV = 93 Then
      If Val(grid_fac.TextMatrix(fila, 10)) = 100 Then
            ww_descto2 = subtotal * Val(grid_fac.TextMatrix(fila, 10)) / 100
            ww_descto2 = redondea(ww_descto2)
            PUB_DESCTO = PUB_DESCTO + ww_descto2
      End If
      End If
   End If
   
   If Val(Nulo_Valor0(SUT_LLAVE!SUT_DESCTO)) <> 0 Then
      If PUB_TIPMOV = 10 Then
            ww_descto2 = Val(grid_fac.TextMatrix(fila, 10))
            ww_descto2 = redondea(ww_descto2)
            PUB_DESCTO = PUB_DESCTO + ww_descto2
      End If
   End If
   
   
   grid_fac.TextMatrix(fila, 34) = redondea(Val(grid_fac.TextMatrix(fila, 37)) * Val(grid_fac.TextMatrix(fila, 4)))
   SUB_PESO = SUB_PESO + Val(grid_fac.TextMatrix(fila, 34))
   
   SUB_FLETE = Val(grid_fac.TextMatrix(fila, 8)) + SUB_FLETE
   SUB_CANT = Val(grid_fac.TextMatrix(fila, 4)) + SUB_CANT
   SUB_UNIDAD = Val(grid_fac.TextMatrix(fila, 3)) + SUB_UNIDAD
   If LK_CODTRA = 1401 Then wdesctopor = wdesctopor + Val(grid_fac.TextMatrix(fila, 10))
i2414:
   If LK_CODTRA = 2414 And fila = 3 Then subtotal = subtotal * -1
   WS_BRUTO = WS_BRUTO + subtotal
   
pasa_2401:

   fila = fila + 1
   If fila = grid_fac.Rows Then
      grid_fac.TextMatrix(1, 8) = SUB_FLETE
      grid_fac.TextMatrix(1, 3) = SUB_UNIDAD
      grid_fac.TextMatrix(1, 4) = SUB_CANT
      grid_fac.TextMatrix(1, 7) = WS_BRUTO
      grid_fac.TextMatrix(1, 10) = wdesctopor ' PUB_DESCTO
      grid_fac.TextMatrix(1, 34) = SUB_PESO
      fila = 999
   End If
   

Loop

   
   
ws_igv = 0
If i_fbg.Visible = True Or PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
   ws_igv = LK_IGV
   If i_fbg.Text = "G" Or i_fbg.Text = "P" Then ws_igv = 0
End If



FORMGEN.i_cantidad.Text = SUB_CANT
PUB_DESCTO = redondea(PUB_DESCTO)

WS_NETO = WS_BRUTO

If PUB_TIPMOV = 20 Then
   pub_flag_cambio = 1
   PUB_GASTOS = Val(i_gastos.Text)
   If Val(i_codcli.Text) > 0 Then
   
      WS_NETO = WS_BRUTO
      WS_BRUTO = redondea((WS_BRUTO / (1 + (ws_igv / 100)))) ' + PUB_GASTOS) * (ws_igv / 100)
      PUB_IMPTO = WS_NETO - WS_BRUTO ' (WS_BRUTO / (1 + (ws_igv / 100))) ' + PUB_GASTOS) * (ws_igv / 100)
      'PUB_IMPTO = PUB_IMPTO - ws_ex_igv
      'If ws_igv = 0 Then
      '   PUB_IMPTO = 0
      'Else
      '   PUB_IMPTO = redondea(PUB_IMPTO)
      'End If
      'WS_NETO = WS_BRUTO + PUB_GASTOS + PUB_IMPTO
   End If
' If Val(i_codcli.Text) > 0 Then
'      WS_NETO = WS_BRUTO
'      PUB_IMPTO = (WS_BRUTO + PUB_GASTOS) * (ws_igv / 100)
'      PUB_IMPTO = PUB_IMPTO - ws_ex_igv
'      If ws_igv = 0 Then
'         PUB_IMPTO = 0
'      Else
'         PUB_IMPTO = redondea(PUB_IMPTO)
'      End If
'      WS_NETO = WS_BRUTO + PUB_GASTOS + PUB_IMPTO
' End If
End If


PUB_FLETE = Val(FORMGEN.i_flete.Text)
If PUB_TIPMOV = 10 Or PUB_TIPMOV = 93 Or PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
   WS_NETO = WS_BRUTO - PUB_DESCTO
   If par_llave!PAR_VENTAS_IGV = "A" Then
      If PUB_DESCTO <> 0 Then
         PUB_DESCTO = redondea((PUB_DESCTO) / (1 + ws_igv / 100))
      End If
       ' ica
      If WS_IEV <> 0 Then
         WS_BRUTO = WS_IEV
      Else
         WS_BRUTO = (WS_NETO) / (1 + ws_igv / 100)
      End If
      WS_BRUTO = WS_BRUTO + PUB_DESCTO
      WS_BRUTO = redondea(WS_BRUTO)
      PUB_IMPTO = WS_NETO - WS_BRUTO + PUB_DESCTO
      If PUB_IMPTO = 0 And ws_igv > 0 Then
         PUB_IMPTO = redondea((WS_BRUTO + PUB_GASTOS) * (ws_igv / 100))
      End If
      PUB_IMPTO = PUB_IMPTO + ws_ex_igv
       ' ica
      If WS_IEV <> 0 Then PUB_IMPTO = 0
      WS_BRUTO = WS_BRUTO + Abs(ws_ex_igv)
      
   ElseIf par_llave!PAR_VENTAS_IGV = "E" Then
     ws_igv = 0
     If PUB_DESCTO <> 0 Then
         PUB_DESCTO = redondea((PUB_DESCTO) / (1 + ws_igv / 100))
      End If
      WS_NETO = WS_NETO + PUB_FLETE
      WS_BRUTO = (WS_NETO) / (1 + ws_igv / 100)
      WS_BRUTO = WS_BRUTO + PUB_DESCTO
      WS_BRUTO = redondea(WS_BRUTO)
      PUB_IMPTO = WS_NETO - WS_BRUTO + PUB_DESCTO
      If PUB_IMPTO = 0 And ws_igv > 0 Then
         PUB_IMPTO = redondea((WS_BRUTO + PUB_GASTOS) * (ws_igv / 100))
      End If
      PUB_IMPTO = PUB_IMPTO + ws_ex_igv
      WS_BRUTO = WS_BRUTO + Abs(ws_ex_igv)
   Else
      PUB_IMPTO = (WS_BRUTO - PUB_DESCTO + PUB_GASTOS) * (ws_igv / 100)
      If ws_igv = 0 Then
         PUB_IMPTO = 0
      Else
         PUB_IMPTO = redondea(PUB_IMPTO)
      End If
      WS_NETO = WS_BRUTO - PUB_DESCTO + PUB_GASTOS + PUB_IMPTO
   End If
End If
      

   FORMGEN.i_impto.Text = PUB_IMPTO
   'FORMGEN.i_descto.Text = PUB_DESCTO
   FORMGEN.i_subtotal.Text = WS_BRUTO
   
pub_cadena = TEXTOWO(1, grid_fac)

' QUITE ACV
 '  If LK_CODTRA = 2412 Or LK_CODTRA = 2410 Then  'REVISAR MONEDAS ....
 '     If WS_MONEDA = "S" Then
 '        lmoneda.Caption = "S/."
 '        WS_MONEDA_CLI = "S"
 '     ElseIf WS_MONEDA = "D" Then
 '         lmoneda.Caption = " $"
 '         WS_MONEDA_CLI = "D"
 '     Else
 '         lmoneda.Caption = " "
 '         WS_MONEDA_CLI = ""
 '     End If
 '  End If
   
   FORMGEN.i_neto.Text = WS_NETO
   If LK_CODTRA = 2401 And i_limcre_ant.Visible And pub_signo_ccm <> 0 Then
     i_limcre_ant.Text = WS_NETO
   End If
   If LK_CODTRA = 2412 Then
     gridl.TextMatrix(2, 5) = WS_NETO
     i_TEXTONCRE.Text = Trim(gridl.TextMatrix(2, 1)) & "/." & Format(gridl.TextMatrix(2, 2), "000") & "-" & gridl.TextMatrix(2, 3)
     calcula_totales2
   End If

   FORMGEN.i_impto2.Text = PUB_IMPTO
   'FORMGEN.i_descto.Text = PUB_DESCTO
   FORMGEN.i_bruto2.Text = WS_BRUTO - PUB_DESCTO
   

If LK_CODTRA = 1403 And Val(i_importe.Text) > 0 Then
      i_importe_amort.Text = WS_NETO - Val(i_importe.Text)
      If Val(i_importe_amort.Text) <> 0 Then
      If (pub_signo_car > 0 And Val(i_importe_amort.Text) > 0) Or (pub_signo_car < 0 And Val(i_importe_amort.Text) < 0) Then
      Else
         MsgBox "NO procede ...Cambiar opcion Aumento/disminucion..."
         cancela_todo
         Exit Sub
      End If
      End If
End If
If PUB_TIPMOV = 10 Then
TOTALES
Else
 i_camal.Visible = False
End If
If (PUB_TIPMOV = 97 Or PUB_TIPMOV = 98) And WS_NETO <> 0 Then FORMGEN.i_importe_amort.Text = WS_NETO

If PUB_TIPMOV <> 10 Then Exit Sub

If cli_llave.EOF Then Exit Sub

grid_fac.TextMatrix(1, 0) = ""

If SUT_LLAVE!SUT_SIGNO_CAR <> 1 Then Exit Sub
If Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 1 Then Exit Sub

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
If par_llave!PAR_MONEDA_FAC = "S" Then
     WTC = 1
Else
 If WTC = 0 Then
   MsgBox "Ingresar el Tipo de Cambio", 48, Pub_Titulo
   GoTo PA
 End If
End If
If Trim(i_ds.Text) = "S" Then
 WSIMBOL = "S/."
 wsumadol = Val(Nulo_Valor0(cli_llave!cli_limcre)) + Val(redondea((Nulo_Valor0(cli_llave!cli_limcre2) * WTC)))
 RES_DEUDA = pub_deuda
Else
 WSIMBOL = "US$."
 wsumadol = Val(redondea(Nulo_Valor0(cli_llave!cli_limcre) / WTC)) + Val(redondea(Val(Nulo_Valor0(cli_llave!cli_limcre2))))
 RES_DEUDA = redondea(Val(pub_deuda / WTC))
End If
grid_fac.TextMatrix(1, 0) = "Disp.: " & WSIMBOL & " " & Format(wsumadol - RES_DEUDA - WS_NETO, "##,###,###.00")
PA:

If LK_CODTRA <> 1111 Then
If LK_FLAG_EXED = "A" Then
   If Trim(i_ds.Text) = "S" Then
     If (RES_DEUDA + WS_NETO) > Nulo_Valor0(cli_llave!cli_limcre) Then
         If ws_bruto_bak <> WS_BRUTO Then MsgBox "Credito Excedido... "
      End If
   Else
     If (RES_DEUDA + WS_NETO) > Nulo_Valor0(cli_llave!cli_limcre2) Then
         If ws_bruto_bak <> WS_BRUTO Then MsgBox "Credito Excedido... "
      End If
   End If
End If
End If

End Sub
Private Sub graba_fffart()
      FFF_ART.Requery
      FFF_ART.AddNew
      FFF_ART!FFF_TIPMOV = PUB_TIPMOV
      FFF_ART!FFF_codcia = LK_CODCIA
      FFF_ART!FFF_numser = PUB_NUMSER
      FFF_ART!FFF_NUMFAC = PUB_NUMFAC
      FFF_ART!FFF_NUMSEC = WS_NUMSEC
      FFF_ART!FFF_fbg = PUB_FBG
      FFF_ART!FFF_TIPMOV_R = Val(grid_fac.TextMatrix(fila, 26))
      FFF_ART!FFF_codcia_R = grid_fac.TextMatrix(fila, 27)
      FFF_ART!FFF_numser_R = grid_fac.TextMatrix(fila, 28)
      FFF_ART!FFF_fbg_R = grid_fac.TextMatrix(fila, 29)
      FFF_ART!FFF_NUMFAC_R = Val(grid_fac.TextMatrix(fila, 30))
      FFF_ART!FFF_NUMSEC_R = grid_fac.TextMatrix(fila, 31)
      FFF_ART.Update
End Sub
Private Sub calcula_totales2()
Dim wvalor As Currency
Dim FILAX
Dim FF As Integer
FILAX = 2
Dim subtotal As Currency
Dim subtotal_che As Currency
Dim saldocar_s As Currency
Dim saldocar_d As Currency

Dim WS_MONEDA As String * 1
FF = 0
subtotal = 0
subtotal_che = 0
saldocar_s = 0
saldocar_d = 0

If i_ds.Visible = True Then
   If i_ds.Text = "S" Or i_ds.Text = "D" Then
   Else
      MsgBox "Seleccione Moneda..."
      Exit Sub
   End If
Else
  WS_MONEDA = i_ds.Text
End If
If i_ds.Visible = True And pub_signo_ccm <> 0 And LK_EMP <> "3AA" And LK_EMP <> "PIU" Then
   If i_ds.Text <> WS_MONEDA_CCM Then
      MsgBox "Moneda errada ..."
      Exit Sub
   End If
End If


If pub_signo_ccm <> 0 Then
   If WS_MONEDA_CLI <> WS_MONEDA_CCM Then
     ' acv  If WS_TIPO_CAMBIO = 0 Then Tipo_Cambio_Click
   End If
End If

Do Until FILAX = 999
   If Trim(gridl.TextMatrix(FILAX, 0)) <> "" Then
         If LK_CODTRA = 1455 And Val(gridl.TextMatrix(FILAX, 5)) <> 0 Then
            If Val(gridl.TextMatrix(FILAX, 5)) <> Val(gridl.TextMatrix(FILAX, 25)) Then
               MsgBox "Importe debe ser Cancelación"
               gridl.TextMatrix(FILAX, 5) = 0
            End If
         End If
         If gridl.TextMatrix(FILAX, 24) = "S" Then
            saldocar_s = saldocar_s + Val(gridl.TextMatrix(FILAX, 25))
         Else
            saldocar_d = saldocar_d + Val(gridl.TextMatrix(FILAX, 25))
         End If
       wvalor = Val(gridl.TextMatrix(FILAX, 5))
       If LK_CODTRA = 1455 Then
          If gridl.TextMatrix(FILAX, 24) = "S" And i_ds.Text = "D" Then
             wvalor = redondea(Val(gridl.TextMatrix(FILAX, 5)) / LK_TIPO_CAMBIO)
          ElseIf gridl.TextMatrix(FILAX, 24) = "D" And i_ds.Text = "S" Then
             wvalor = redondea(Val(gridl.TextMatrix(FILAX, 5)) * LK_TIPO_CAMBIO)
          End If
          If Val(gridl.TextMatrix(FILAX, 25)) < 0 Then
             subtotal = subtotal - wvalor
          Else
             subtotal = subtotal + wvalor
          End If
       Else
          If Val(gridl.TextMatrix(FILAX, 25)) < 0 Then
             subtotal = subtotal - Val(gridl.TextMatrix(FILAX, 5))
          Else
             subtotal = subtotal + Val(gridl.TextMatrix(FILAX, 5))
          End If
       End If
         If Val(gridl.TextMatrix(FILAX, 5)) <> 0 Then
            If FF = 0 Then
               WS_MONEDA = gridl.TextMatrix(FILAX, 24)
               LOC_MONEDA_DIG = gridl.TextMatrix(FILAX, 24)
               FF = 1
            Else
               If WS_MONEDA <> gridl.TextMatrix(FILAX, 24) And LK_CODTRA <> 1455 Then
                  MsgBox "HAY DIFERENTES MONEDAS...", 48, Pub_Titulo
                  cancelar_Click
                  Exit Sub
               End If
            End If
         End If
         
         If WS_MONEDA = "S" Then
            WS_MONEDA_CLI = "S"
         Else
            WS_MONEDA_CLI = "D"
         End If
         
         subtotal_che = subtotal_che + Val(gridl.TextMatrix(FILAX, 19))
         
    End If
   FILAX = FILAX + 1
   If FILAX = gridl.Rows Then
      FILAX = 999
   End If
Loop
   gridl.TextMatrix(1, 4) = "S/." & saldocar_s
   gridl.TextMatrix(1, 3) = "$" & saldocar_d
   gridl.TextMatrix(1, 5) = subtotal
   gridl.TextMatrix(1, 19) = subtotal_che
   ' CAMTEXT
   gridl.TextMatrix(0, 5) = "Efec S/."
   If LK_CODTRA = 1455 And i_ds.Text = "D" Then
      gridl.TextMatrix(0, 5) = "Efec US$."
      gridl.TextMatrix(1, 5) = Format(subtotal, "0.00")
   End If
   
   
   
   FORMGEN.i_importe.Text = subtotal + subtotal_che
   
'   If Tipo_Cambio.Visible = False Then Exit Sub
   
   If pub_signo_ccm <> 0 And LK_CODTRA <> 2725 Then
      If Trim(WS_MONEDA_CCM) = Trim(WS_MONEDA_CLI) Then Exit Sub
      If PUB_CP = "P" Then
      If WS_MONEDA_CCM = "S" Then
         If WS_MONEDA_CLI = "D" Then FORMGEN.i_importe.Text = redondea(subtotal * LK_TIPO_CAMBIO)
      Else
         If WS_MONEDA_CLI = "S" Then FORMGEN.i_importe.Text = redondea(subtotal / LK_TIPO_CAMBIO)
      End If
      End If
    Else
      If i_ds.Text = "S" Then
         If WS_MONEDA = "D" Then FORMGEN.i_importe.Text = redondea(subtotal * LK_TIPO_CAMBIO)
      Else
         If WS_MONEDA = "S" Then FORMGEN.i_importe.Text = redondea(subtotal / LK_TIPO_CAMBIO)
      End If
      
   End If
   
End Sub
Private Sub CALCULA_TOTALES_CANJE()
Dim FILAX
FILAX = 2
Dim subtotal As Currency

Do Until FILAX = 999
   subtotal = subtotal + Val(grid_canje.TextMatrix(FILAX, 3))
   grid_canje.TextMatrix(FILAX, 1) = FILAX - 1 & ".-"
   FILAX = FILAX + 1
   If FILAX = grid_canje.Rows Then
      FILAX = 999
   End If
Loop
grid_canje.TextMatrix(1, 3) = subtotal
FORMGEN.i_importe = subtotal + Val(gridl.TextMatrix(1, 5))
End Sub


Private Sub TOMA_DATOS()
'INICIALIZA_INPUT
Dim pos2 As Integer
Dim CARAC As String


pub_numplan = Val(FORMGEN.i_numplan.Text)
PUB_NUM_OPER = Val(FORMGEN.i_num_oper.Text)

PUB_CODCLIE = Val(FORMGEN.i_codcli.Text)
PUB_CANT_CHEQ = Val(FORMGEN.i_cant_cheq.Text)
PUB_NUM_INI = Val(FORMGEN.i_num_ini.Text)
PUB_CODVEN = Val(FORMGEN.i_codven.Text)
PUB_BRUTO2 = Val(FORMGEN.i_bruto2.Text)
PUB_IMPTO2 = Val(FORMGEN.i_impto2.Text)
PUB_IMPORTE = Val(FORMGEN.i_importe.Text)
' ICA
If LK_CODTRA = 2412 Then
 PUB_NETO = Val(FORMGEN.i_neto.Text)
 PUB_SUBTOTAL = PUB_BRUTO2
 PUB_IMPTO = PUB_IMPTO2
Else
 PUB_NETO = Val(FORMGEN.i_neto.Text)
 PUB_SUBTOTAL = Val(FORMGEN.i_subtotal.Text)
 PUB_IMPTO = Val(FORMGEN.i_impto.Text)
End If

If LK_CODTRA = 2748 Then
PUB_IMPTO = PUB_IMPTO2
PUB_SUBTOTAL = PUB_BRUTO2
End If

PUB_FLETE = Val(FORMGEN.i_flete.Text)

PUB_DESCTO = Val(FORMGEN.i_descto.Text)
PUB_GASTOS = Val(FORMGEN.i_gastos.Text)
If LK_CODTRA = 2748 Then
 PUB_GASTOS = Val(FORMGEN.i_importe.Text)
End If
PUB_CANTIDAD = Val(FORMGEN.i_cantidad.Text)

PUB_NUM_CHEQUE = FORMGEN.i_num_cheque.Text
PUB_NOMBRE_BANCO = FORMGEN.i_nombre_banco.Text

PUB_SERDOC = Val(FORMGEN.i_serdoc.Text)

PUB_NUMSER = Val(FORMGEN.i_numser.Text)


PUB_NUMFAC = Val(FORMGEN.i_numfac.Text)

PUB_NUMSER_C = Val(FORMGEN.i_numser_c.Text)
PUB_NUMFAC_C = Val(FORMGEN.i_numfac_c.Text)

PUB_USUARIO = RTrim(FORMGEN.i_CODUSU.Text)
PUB_TIPO_BLOQ_act1 = Left(FORMGEN.i_tipo_bloq_act1.Text, 1)
PUB_TIPO_BLOQ_act2 = Left(FORMGEN.i_tipo_bloq_act2.Text, 1)
PUB_TIPO_BLOQ_act3 = Left(FORMGEN.i_tipo_bloq_act3.Text, 1)
PUB_TIPO_BLOQ_act4 = Left(FORMGEN.i_tipo_bloq_act4.Text, 1)
PUB_TIPO_BLOQ_ant1 = Left(FORMGEN.i_tipo_bloq_ant1.Text, 1)
PUB_TIPO_BLOQ_ant2 = Left(FORMGEN.i_tipo_bloq_ant2.Text, 1)
PUB_TIPO_BLOQ_ant3 = Left(FORMGEN.i_tipo_bloq_ant3.Text, 1)
PUB_TIPO_BLOQ_ant4 = Left(FORMGEN.i_tipo_bloq_ant4.Text, 1)

PUB_LIMCRE_ACT = Val(FORMGEN.i_limcre.Text)
PUB_LIMCRE_ANT = Val(FORMGEN.i_limcre_ant.Text)
PUB_CODBAN = Val(FORMGEN.i_codban.Text)
pub_dias = Val(FORMGEN.i_dias.Text)
PUB_CONCEPTO = FORMGEN.i_concepto.Text
PUB_CHENUM = Val(FORMGEN.i_chenum.Text)
PUB_CHESER = FORMGEN.i_cheser.Text
If i_fbg.Visible = False Then
   PUB_FBG = ""
Else
   PUB_FBG = FORMGEN.i_fbg.Text
End If
If PUB_TIPMOV = 97 Then PUB_FBG = "N"
If PUB_TIPMOV = 98 Then PUB_FBG = "D"
If PUB_TIPMOV = 20 Then PUB_FBG = ""

PUB_CHESEC = 0
PUB_NUMGUIA = Val(FORMGEN.i_numguia.Text)
PUB_SERGUIA = Val(FORMGEN.i_serguia.Text)
PUB_TURNO = Val(Right(cmdtipo.Text, 6)) ' Val(FORMGEN.i_turno.Text)
If LK_CODTRA = 1401 Then
  PUB_TURNO = Val(Right(i_situacion.Text, 6))
  FORMGEN.t_doc.Text = Trim(i_num_ini.Text)
  PUB_PEDFAC = Val((i_num_lote.Text))
  
End If


PUB_IMPORTE_AMORT = Val(FORMGEN.i_importe_amort.Text)
If LK_CODTRA = 2748 Or (LK_CODTRA = 2741 And LK_EMP = "HER") Then
  PUB_IMPORTE = PUB_IMPORTE_AMORT
End If
PUB_DS = i_ds.Text ' verificar esto """""


PUB_NOMPRO = Trim(FORMGEN.i_nompro.Text)
PUB_NUMDOC = Val(FORMGEN.i_numdoc.Text)
PUB_CODART = Val(FORMGEN.i_codart.Text)
PUB_PRECIO = Val(FORMGEN.i_precio.Text)
If IsDate(FORMGEN.i_fecha_vcto.Text) Then
    PUB_FECHA_VCTO = CDate(FORMGEN.i_fecha_vcto.Text)
Else
    PUB_FECHA_VCTO = 0
End If
If FORMGEN.i_def.ListCount = 0 Then
   PUB_SECUENCIA = 0
Else
   pos2 = InStr(1, FORMGEN.i_def.Text, ".", 1)
   CARAC = Mid(FORMGEN.i_def.Text, 1, pos2)
   PUB_SECUENCIA = Val(CARAC)
End If

End Sub

Private Sub Boton_Letras_Click()
   If LK_CODTRA = 2728 Then
      pub_cadena = grid_liq.FormatString
      grid_liq.Clear
      grid_liq.FormatString = pub_cadena
      grid_liq.Visible = False
   End If
   If FORMGEN.i_codcli.Text <> "" Then
    If LK_CODTRA = 2412 Then
      If pub_signo_car = 0 Then Exit Sub
    End If
      PROCESA_GRIDL
      i_situacion.ListIndex = -1
      i_situacion_ant = ""
     If i_fbg.Text = "B" Then
     '  FORMGEN.i_codven.Text = "2"
     Else
     '  FORMGEN.i_codven.Text = "1"
     End If
      i_codven_KeyPress 13
   End If
   PUB_IMPORTE = 0
   If LK_CODTRA = 2412 Then
     If grid_fac.Visible = True Then
       grid_fac.COL = 1
       grid_fac.SetFocus
     End If
  End If
End Sub
Private Sub Boton_Recepcion_Click()

If LK_CODTRA = 1405 Then GoTo recep_guia
Dim ww_numfac
grid_fac.Clear
pasa_cabeza
grid_trans.Clear
pasa_cabeza_grid_trans
Screen.MousePointer = 11
pub_cadena = "SELECT *  FROM facart WHERE  FAR_TRANSITO = 'T' AND FAR_ESTADO <> 'E' AND FAR_OTRA_CIA = ? AND FAR_ESTADO2 <> 'L' ORDER BY FAR_OTRA_CIA, FAR_FBG, FAR_NUMSER, FAR_NUMFAC,FAR_NUMSEC"
Set PSFAR_TRANS = CN.CreateQuery("", pub_cadena)
PSFAR_TRANS.rdoParameters(0) = " "
Set FAR_TRANS = PSFAR_TRANS.OpenResultset(rdOpenKeyset, rdConcurValues)
PSFAR_TRANS.rdoParameters(0) = LK_CODCIA
FAR_TRANS.Requery
If FAR_TRANS.EOF Then
   Screen.MousePointer = 0
   MsgBox "No hay Registros para Recepcion"
   Exit Sub
End If
fila = 0
ww_numfac = 0
Do Until FAR_TRANS.EOF
   GoSub muestra
   FAR_TRANS.MoveNext
Loop
grid_trans.ColWidth(0) = 1300
grid_trans.ColWidth(7) = 1300
FORMGEN.grid_trans.ColWidth(6) = 0
FORMGEN.grid_trans.Cols = 8
FORMGEN.grid_trans.ColWidth(1) = 2200
FORMGEN.grid_trans.ColWidth(2) = 1500
FORMGEN.grid_trans.ColWidth(3) = 2000
FORMGEN.grid_trans.ColWidth(4) = 1000
FORMGEN.grid_trans.ColWidth(5) = 1000
FORMGEN.grid_trans.ColWidth(6) = 0
FORMGEN.grid_trans.ColWidth(7) = 0



grid_trans.Visible = True
grid_trans.Height = Frame4.Height - 1000
grid_trans.Width = Frame4.Width
label_grid.Top = Frame4.Top - 400
grid_trans.Left = 0
label_grid.Caption = "Documentos LLegados <Seleccionar con ENTER>"
label_grid.Left = 0
label_grid.Width = Frame4.Width
grid_trans.Top = Frame4.Top
label_grid.Visible = True
grid_trans.Left = 0
grid_trans.SetFocus
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE

Screen.MousePointer = 0
Exit Sub

muestra:
      If Val(i_numguia.Text) > 0 Then
      If FAR_TRANS!far_numguia = Val(i_numguia.Text) And FAR_TRANS!far_serguia = Val(i_serguia.Text) Then
      Else
      Return
      End If
      End If
      
      fila = fila + 1
      grid_trans.Rows = fila + 1
      grid_trans.TextMatrix(fila, 0) = FAR_TRANS!far_numfac
      SQ_OPER = 1
      PUB_CODCIA = FAR_TRANS!FAR_CODCIA
      LEER_PAR_LLAVE
      
      PUB_KEY = FAR_TRANS!far_codart
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ART_LLAVE
      grid_trans.TextMatrix(fila, 1) = art_LLAVE!art_nombre
      grid_trans.TextMatrix(fila, 2) = FAR_TRANS!far_cantidad
      grid_trans.TextMatrix(fila, 3) = par_llave!PAR_NOMBRE
      If FAR_TRANS!far_transito = "X" Then
         grid_trans.TextMatrix(fila, 5) = "Recibido"
         grid_trans.Row = fila
         grid_trans.COL = 0
         grid_trans.CellBackColor = vb3DLight
         grid_trans.COL = 1
         grid_trans.CellBackColor = vb3DLight
      Else
         grid_trans.TextMatrix(fila, 5) = "Por Recibir"
      End If
      grid_trans.TextMatrix(fila, 6) = par_llave!PAR_CODCIA
      grid_trans.TextMatrix(fila, 7) = FAR_TRANS!far_numser
      Return
Screen.MousePointer = 0

Exit Sub


'**Recepcion de Guias de Compras.**
'**Recepcion de Guias de Compras.**
'**Recepcion de Guias de Compras.**
recep_guia:

'grid_fac.Clear
pasa_cabeza
grid_trans.Clear
pasa_cabeza_grid_trans
Screen.MousePointer = 11

pub_cadena = "SELECT FAR_CANTIDAD, FAR_TRANSITO, FAR_NUMSER_C, FAR_NUMFAC_C, FAR_FECHA_COMPRA  FROM FACART WHERE  FAR_CODCIA = ? AND FAR_TIPMOV = 20 AND FAR_TRANSITO = 'P' AND FAR_ESTADO <> 'E' ORDER BY FAR_FECHA_COMPRA,FAR_NUMSER_C,FAR_NUMFAC_C"
pub_cadena = "SELECT * FROM FACART WHERE  FAR_CODCIA = ? and FAR_CODCLIE = ? AND FAR_CP = 'P' AND FAR_TIPMOV = 20 AND FAR_TRANSITO = 'P' AND FAR_ESTADO <> 'E' AND ((FAR_CANTIDAD - FAR_CANTIDAD_P) >0 ) ORDER BY FAR_FECHA_COMPRA,FAR_NUMSER_C, FAR_NUMFAC_C"
Set PSFAR_TRANS = CN.CreateQuery("", pub_cadena)
PSFAR_TRANS.rdoParameters(0) = " "
PSFAR_TRANS.rdoParameters(1) = 0
Set FAR_TRANS = PSFAR_TRANS.OpenResultset(rdOpenKeyset, rdConcurValues)
PSFAR_TRANS.rdoParameters(0) = LK_CODCIA
PSFAR_TRANS.rdoParameters(1) = Val(i_codcli.Text)
FAR_TRANS.Requery
If FAR_TRANS.EOF Then
   Screen.MousePointer = 0
   MsgBox "No hay Registros para Recepcion", 48, Pub_Titulo
   Exit Sub
End If
fila = 0
ww_numfac = 0
Do Until FAR_TRANS.EOF
   GoSub muestra2
   FAR_TRANS.MoveNext
Loop


grid_trans.ColWidth(7) = 1300
FORMGEN.grid_trans.ColWidth(6) = 0
FORMGEN.grid_trans.Cols = 8
FORMGEN.grid_trans.ColWidth(0) = 800
FORMGEN.grid_trans.ColWidth(1) = 1000
FORMGEN.grid_trans.ColWidth(2) = 3500
FORMGEN.grid_trans.ColWidth(3) = 1200
FORMGEN.grid_trans.ColWidth(4) = 1200
FORMGEN.grid_trans.ColWidth(5) = 0
FORMGEN.grid_trans.ColWidth(6) = 0
FORMGEN.grid_trans.ColWidth(7) = 0



grid_trans.Visible = True
grid_trans.Height = Frame4.Height - 1000
grid_trans.Width = Frame4.Width
label_grid.Top = Frame4.Top - 400
label_grid.Caption = "Documentos <Seleccionar con ENTER>"
''Print label_grid.Caption

grid_trans.Left = 0

label_grid.Left = 0
label_grid.Width = Frame4.Width
grid_trans.Top = Frame4.Top
label_grid.Visible = True
grid_trans.Left = 0
grid_trans.SetFocus
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE

Screen.MousePointer = 0
Exit Sub

muestra2:
      If Val(i_numfac_c.Text) > 0 Then
        If FAR_TRANS!FAR_NUMFAC_C = Val(i_numfac_c.Text) Then
        Else
          Return
        End If
      End If
      If Val(i_numser_c.Text) > 0 Then
        If FAR_TRANS!FAR_NUMSER_C = Val(i_numser_c.Text) Then
        Else
          Return
        End If
      End If
      fila = fila + 1
      grid_trans.Rows = fila + 1
      grid_trans.TextMatrix(fila, 0) = FAR_TRANS!FAR_NUMSER_C
      grid_trans.TextMatrix(fila, 1) = FAR_TRANS!FAR_NUMFAC_C
      SQ_OPER = 1
      PUB_CODCIA = FAR_TRANS!FAR_CODCIA
      LEER_PAR_LLAVE
      PUB_KEY = FAR_TRANS!far_codart
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ART_LLAVE
      grid_trans.TextMatrix(fila, 2) = art_LLAVE!art_nombre
      grid_trans.TextMatrix(fila, 3) = FAR_TRANS!far_descri
      grid_trans.TextMatrix(fila, 4) = Format((FAR_TRANS!far_cantidad - FAR_TRANS!far_cantidad_p) / FAR_TRANS!FAR_equiv, "0.000")

      grid_trans.TextMatrix(fila, 6) = FAR_TRANS!far_numser
      grid_trans.TextMatrix(fila, 7) = FAR_TRANS!far_numfac
    
      Return
      
Screen.MousePointer = 0

Screen.MousePointer = 0

End Sub

Private Sub cancelar_Click()
If fraPedidos.Visible = True Then
 fraPedidos.Visible = False
 cmdcanped_Click
End If
det_lot.Rows = 1
det_lot.Clear
i_def.Enabled = True
LOC_MONEDA_DIG = ""
FILAX = 0
cancela_todo
Grid_all.Visible = False
FORMGEN.gridl.Clear
FORMGEN.grid_fac.Clear
pasa_cabeza_L
pasa_cabeza
If grid_canje.Visible = True Then pasa_cabeza_canje
grid_fac.Rows = 3
grabar.Enabled = True
If LK_CODTRA = 2401 Or LK_CODTRA = 1401 Then
 'i_fecha_compra.Visible = False
 
 i_cambio.Value = 0
 i_cambio.Visible = False
End If

i_TEXTONCRE.Text = ""
i_numfac_c.Text = ""
i_numser_c.Text = ""

If i_def.Visible And i_def.Enabled Then
  i_def.SetFocus
End If
End Sub

Private Sub chkBoletas_Click()
Dim C As Integer
If chkBoletas.Value = 1 Then
  For C = 0 To i_fbg.ListCount - 1
   i_fbg.ListIndex = C
   If Trim(i_fbg.Text) = "B" Then Exit For
  Next C
  BolFac_Click
  FORMGEN.txtSerieDocIniBol.Text = Trim(i_numser.Text)
  FORMGEN.txtNumDocIniBol.Text = 0
  FORMGEN.txtNumDocIniBol.Text = Trim(i_numfac.Text)
  Azul txtSerieDocIniBol, txtSerieDocIniBol
Else
  FORMGEN.txtSerieDocIniBol.Text = ""
  FORMGEN.txtNumDocIniBol.Text = ""
  FORMGEN.txtSerieDocBolFin.Text = ""
  FORMGEN.txtNumDocFinBol.Text = ""
End If
If Grid_detalle.Visible And Grid_detalle.Rows > 2 Then Grid_detalle.Rows = 2
End Sub

Private Sub chkFacturas_Click()
Dim C As Integer
If chkFacturas.Value = 1 Then
  For C = 0 To i_fbg.ListCount - 1
   i_fbg.ListIndex = C
   If Trim(i_fbg.Text) = "F" Then Exit For
  Next C
  BolFac_Click
  FORMGEN.txtSerieDocIniFac.Text = Trim(i_numser.Text)
  FORMGEN.txtNumDocIniFac.Text = 0
  FORMGEN.txtNumDocIniFac.Text = Trim(i_numfac.Text)
  Azul txtSerieDocIniFac, txtSerieDocIniFac
Else
  FORMGEN.txtSerieDocIniFac.Text = ""
  FORMGEN.txtSerieDocFacFin.Text = ""
  FORMGEN.txtNumDocIniFac.Text = ""
  FORMGEN.txtNumDocFacFin.Text = ""
End If
If Grid_detalle.Visible And Grid_detalle.Rows > 2 Then Grid_detalle.Rows = 2
End Sub

Private Sub cmdcanped_Click()
Grid_detalle.Rows = 2
Azul2 txtCampo1, txtCampo1
End Sub

Private Sub cmdchequea_Click()
Dim CAP_CODVEN As Integer
pub_cadena = "SELECT * FROM VEMAEST WHERE  VEM_CODCIA = ? "
Set PSFAR_TRANS = CN.CreateQuery("", pub_cadena)
PSFAR_TRANS.rdoParameters(0) = LK_CODCIA
Set FAR_TRANS = PSFAR_TRANS.OpenResultset(rdOpenKeyset, rdConcurValues)
CAP_CODVEN = Val(txtCodVen.Text)
Do Until FAR_TRANS.EOF
 txtCodVen.Text = Val(FAR_TRANS!VEM_CODVEN)
 If Val(txtCodVen.Text) >= CAP_CODVEN Then
 Else
   GoTo SIGUE_VEN
 End If
 chkBoletas.Value = 0
 chkFacturas.Value = 1
 cmdmostrar_Click
 If Grid_detalle.Rows > 2 Then
     Exit Sub
 End If
 chkFacturas.Value = 0
 chkBoletas.Value = 1
 cmdmostrar_Click
 If Grid_detalle.Rows > 2 Then
     Exit Sub
 End If
SIGUE_VEN:
 FAR_TRANS.MoveNext
Loop
If Grid_detalle.Rows <= 2 Then
     MsgBox "Ok. !!! No Existe Pedidos Pendientes", 48, Pub_Titulo
End If
End Sub

Private Sub cmddd_Click()
Load frmdeudas
frmdeudas.Show 1
End Sub

Private Sub CmdEsp_Click()
Dim xcI As Integer
If LK_CODTRA = 2412 Then
    If Val(Trim(FORMGEN.i_codcli.Text)) = 0 Then
      MsgBox "Ingrese Cliente", 48, Pub_Titulo
      Exit Sub
    End If
Else
End If
Load FrmNC
If Val(FrmNC.tcodcli.Text) <> 0 Then
 If Val(FrmNC.tcodcli.Text) <> Val(Trim(FORMGEN.i_codcli.Text)) Then
   Unload FrmNC
   Load FrmNC
 End If
End If
FrmNC.tcodcli.Text = Trim(FORMGEN.i_codcli.Text)
FrmNC.lblnomcli.Caption = Trim(FORMGEN.i_nomCLI.Caption)
FrmNC.Show 1
i_importe_amort.Text = Format(FrmNC.lbltotnc.Caption, "0.00")
WS_NETO = Val(i_importe_amort.Text)
If LK_CODTRA = 2412 Then
     gridl.TextMatrix(2, 5) = WS_NETO
     i_TEXTONCRE.Text = Trim(gridl.TextMatrix(2, 1)) & "/." & Format(gridl.TextMatrix(2, 2), "000") & "-" & gridl.TextMatrix(2, 3)
     calcula_totales2
End If
If LK_CODTRA = 2774 Then
    For xcI = 1 To FrmNC.gridcabe.Rows - 1
      If Val(FrmNC.gridcabe.TextMatrix(xcI, 12)) <> 0 Then
         If i_fbg.List(0) = FrmNC.gridcabe.TextMatrix(xcI, 14) Then
         i_fbg.ListIndex = 0
         ElseIf i_fbg.List(1) = FrmNC.gridcabe.TextMatrix(xcI, 14) Then
         i_fbg.ListIndex = 1
         ElseIf i_fbg.List(2) = FrmNC.gridcabe.TextMatrix(xcI, 14) Then
         i_fbg.ListIndex = 2
         ElseIf i_fbg.List(3) = FrmNC.gridcabe.TextMatrix(xcI, 14) Then
         i_fbg.ListIndex = 3
         End If
         i_numser_c = FrmNC.gridcabe.TextMatrix(xcI, 15)
         i_numfac_c = FrmNC.gridcabe.TextMatrix(xcI, 16)
         loc_key = 0
         i_numfac_c_KeyPress 13
         If loc_key <> 777 Then
           textovarl.Text = FrmNC.gridcabe.TextMatrix(xcI, 12)
         End If
         loc_key = 0
      End If
    Next xcI

End If
End Sub

Private Sub cmdGenerar_Click()
Dim wp2 As String
Dim wp1 As String
Dim DEDOS As Currency
Dim Sec_SerieFactura As Currency
Dim Sec_SerieBoleta As Currency
Dim Sec_NumFactura As Currency
Dim Sec_NumBoleta As Currency
Dim C As Integer

PUB_VISITA = 0

pub_cadena = "SELECT  * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = ? " & _
 "AND PED_NUMSER = ? AND PED_NUMFAC = ? "
Set PSLOC_PED_LLAVE = CN.CreateQuery("", pub_cadena)
PSLOC_PED_LLAVE(0) = 0
PSLOC_PED_LLAVE(1) = 0
PSLOC_PED_LLAVE(2) = 0
PSLOC_PED_LLAVE(3) = 0
Set loc_ped_llave = PSLOC_PED_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

PUB_PROCESO = 1
Sec_SerieBoleta = Val(FORMGEN.txtSerieDocIniBol)
Sec_SerieFactura = Val(FORMGEN.txtSerieDocIniFac)
Sec_NumBoleta = Val(FORMGEN.txtNumDocIniBol) - 1
Sec_NumFactura = Val(FORMGEN.txtNumDocIniFac) - 1


With Grid_detalle
 If .Rows <= 2 Then
     MsgBox "No Existe Datos para Generar", 48, Pub_Titulo
     Exit Sub
 End If
DEDOS = 0
pub_mensaje = "Generar Todos los Documentos --> [Aceptar], o  uno a uno -->[Cancelar] "
Pub_Respuesta = MsgBox(pub_mensaje, vbOKCancel + vbInformation + vbDefaultButton1, Pub_Titulo)
If Pub_Respuesta = vbOK Then
   DEDOS = 1
End If
     
For C = 2 To .Rows - 1
    PSLOC_PED_LLAVE(0) = LK_CODCIA
    PSLOC_PED_LLAVE(1) = 201
    PSLOC_PED_LLAVE(2) = .TextMatrix(C, 2)
    PSLOC_PED_LLAVE(3) = Val(.TextMatrix(C, 3))
    loc_ped_llave.Requery
    If UCase(Left(.TextMatrix(C, 5), 1)) = "B" Then
        Sec_NumBoleta = Sec_NumBoleta + 1
    Else
        Sec_NumFactura = Sec_NumFactura + 1
    End If
    If loc_ped_llave.EOF Then
       MsgBox "No Existe "
    Else
      GoSub llena_2401
    End If
    If DEDOS = 0 Then
    '  DEDOS = 0
     pub_mensaje = "Generar la " & Trim(i_fbg.Text) & "/ " & Format(Trim(i_numser.Text), "000") & " " & Format(i_numfac.Text, "0000000")
     Pub_Respuesta = MsgBox(pub_mensaje, vbOKCancel + vbInformation + vbDefaultButton1, Pub_Titulo)
     If Pub_Respuesta = vbCancel Then
        Exit For
        'GoTo pasa_ped
     End If
    End If
    cmdGenerar.SetFocus
    grabar_Click
    cmdGenerar.SetFocus
pasa_ped:
Next C
'If DEDOS <> 0 Then grabar_Click 'SI FALTA PARA IMPRIMI ..TODO

End With

PUB_PROCESO = 0
cancelar_Click
fraPedidos.Visible = False

Exit Sub
'llena para grabar 2401
llena_2401:

loc_ped_llave.MoveFirst
For fila = 0 To i_def.ListCount - 1
  i_def.ListIndex = fila
  If Val(Left(i_def.Text, 2)) = Val(loc_ped_llave!PED_CONDI) Then
     Exit For
  End If
Next fila
i_def_LostFocus

i_codcli.Text = loc_ped_llave!PED_CODCLIE
i_codcli_KeyPress 13
i_codcli_LostFocus
i_codven.Text = loc_ped_llave!PED_codven
i_codven_KeyPress 13
i_dias.Text = loc_ped_llave!ped_DIAS

'forzar el documento a Registrar
If loc_ped_llave!ped_FBG = "F" Then
  i_fbg.ListIndex = 0
Else
  i_fbg.ListIndex = 1
End If
If UCase(Left(Grid_detalle.TextMatrix(C, 5), 1)) = "B" Then
    PUB_NUMSER = Sec_SerieBoleta
    PUB_NUMFAC = Sec_NumBoleta
Else
    PUB_NUMSER = Sec_SerieFactura
    PUB_NUMFAC = Sec_NumFactura
End If

i_cambio.Value = 1
i_cambio.Visible = True
i_numser.Text = PUB_NUMSER
i_numfac.Text = PUB_NUMFAC

    

i_dircli_GotFocus
For fila = 0 To i_dircli.ListCount - 1
   i_dircli.ListIndex = fila
  If Val(Trim(Right(i_dircli.Text, 8))) = Val(loc_ped_llave!ped_DIRCLI) Then
    Exit For
  End If
Next fila

For fila = 0 To cmdtipo.ListCount - 1
  cmdtipo.ListIndex = fila
  If Val(Trim(Right(cmdtipo.Text, 8))) = Val(loc_ped_llave!PED_TIPVTA) Then
    Exit For
  End If
Next fila


PUB_PEDSER = loc_ped_llave!PED_NUMSER
PUB_PEDFAC = loc_ped_llave!PED_NUMFAC

fila = 2
Do Until loc_ped_llave.EOF
   grid_fac.Rows = fila + 2
   
   PUB_KEY = loc_ped_llave!PED_CODART
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ART_LLAVE
   SQ_OPER = 1
   PUB_CODART = PUB_KEY
   pu_codcia = LK_CODCIA
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
      MsgBox "Datos de Articulos ...errados..Revisar"
      Exit Sub
   End If
   PUB_CODART = loc_ped_llave!PED_CODART
   pu_codcia = LK_CODCIA
   PUB_SECUEN = loc_ped_llave!PED_NUM_UNIDAD ' ver
   SQ_OPER = 1
   LEER_PRE_LLAVE
   grid_fac.TextMatrix(fila, 17) = loc_ped_llave!PED_NUMPRE
   grid_fac.TextMatrix(fila, 34) = redondea(pre_llave!pre_PESO * loc_ped_llave!PED_cantidad)
   grid_fac.TextMatrix(fila, 37) = pre_llave!pre_PESO
   grid_fac.TextMatrix(fila, 43) = Nulo_Valor0(pre_llave!PRE_LITRO)
   grid_fac.TextMatrix(fila, 0) = art_LLAVE!art_nombre
   grid_fac.TextMatrix(fila, 1) = loc_ped_llave!PED_CODART
   grid_fac.TextMatrix(fila, 16) = loc_ped_llave!PED_CODART
   grid_fac.TextMatrix(fila, 2) = 0
   grid_fac.TextMatrix(fila, 3) = loc_ped_llave!PED_UNIDAD
   grid_fac.TextMatrix(fila, 11) = arm_llave!arm_cospro
   If Nulo_Valor0(loc_ped_llave!PED_EQUIV) > 0 Then
   grid_fac.TextMatrix(fila, 4) = loc_ped_llave!PED_cantidad / Nulo_Valor0(loc_ped_llave!PED_EQUIV)
   Else
   grid_fac.TextMatrix(fila, 4) = loc_ped_llave!PED_cantidad
   End If
   
   'If SUT_LLAVE!SUT_PRECIO <> "4" Then grid_fac.TextMatrix(fila, 6) = loc_ped_llave!ped_COSPRO
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(loc_ped_llave!PED_EQUIV)
   grid_fac.TextMatrix(fila, 5) = Nulo_Valors(loc_ped_llave!PED_UNIDAD)
   grid_fac.TextMatrix(fila, 6) = Nulo_Valors(loc_ped_llave!PED_PRECIO)
   grid_fac.TextMatrix(fila, 8) = 0
   wp1 = ""
   If Val(loc_ped_llave!PED_descto_pre) <> 0 Then
      wp1 = Format(loc_ped_llave!PED_descto_pre, "0.0")
      If Val(loc_ped_llave!PED_descto) <> 0 Then
        wp1 = wp1 & "+"
      End If
   End If
   wp2 = ""
   If Val(loc_ped_llave!PED_descto) <> 0 Then
      wp2 = Format(loc_ped_llave!PED_descto, "0.0")
   End If
   
   grid_fac.TextMatrix(fila, 10) = wp1 & wp2 '  Format(loc_ped_llave!PED_descto, "#.0#") & "+" & Format(loc_ped_llave!PED_descto_pre, "#.0#")
   
   grid_fac.TextMatrix(fila, 12) = SUT_LLAVE!SUT_SIGNO_ARM
   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   grid_fac.TextMatrix(fila, 26) = 10 'loc_ped_llave!ped_TIPMOV
   grid_fac.TextMatrix(fila, 27) = loc_ped_llave!PED_CODCIA
   grid_fac.TextMatrix(fila, 28) = loc_ped_llave!PED_NUMSER
   grid_fac.TextMatrix(fila, 29) = loc_ped_llave!ped_FBG
   grid_fac.TextMatrix(fila, 30) = loc_ped_llave!PED_NUMFAC
   grid_fac.TextMatrix(fila, 31) = loc_ped_llave!PED_NUMSEC
   grid_fac.TextMatrix(fila, 33) = loc_ped_llave!PED_CODART
   grid_fac.TextMatrix(fila, 18) = loc_ped_llave!PED_DIA_VISITA
   grid_fac.TextMatrix(fila, 45) = Trim(loc_ped_llave!PED_OFERTA)
   
   PUB_CODUSU = loc_ped_llave!PED_CODUSU
   
    fila = fila + 1
OTRO:
   loc_ped_llave.MoveNext
Loop
label_grid.Visible = False
calcula_totales
If grid_fac.Visible Then grid_fac.SetFocus
Screen.MousePointer = 0

Return
End Sub

Private Sub cmdltaceptar_Click()
Dim xcta As Integer
Dim TOT_UM_MINIMAS As Currency
Dim canti As Currency

TOT_UM_MINIMAS = 0
flag_lotes_chequea = ""
 
For xcta = 2 To gridlt.Rows - 1
  If Not IsDate(gridlt.TextMatrix(xcta, 4)) Then
     MsgBox "Verificar Fecha", 48, Pub_Titulo
     Exit Sub
  End If
  If Trim(gridlt.TextMatrix(xcta, 0)) = "" Then
     MsgBox "No procede nro de Lote", 48, Pub_Titulo
     Exit Sub
  End If
  TOT_UM_MINIMAS = TOT_UM_MINIMAS + (Val(gridlt.TextMatrix(xcta, 3)) * Val(gridlt.TextMatrix(xcta, 11)))
Next xcta
canti = Val(lblltcantidad.Caption) * Val(Right$(lblltunidad.Caption, 10))
TOT_UM_MINIMAS = redondea(TOT_UM_MINIMAS)
canti = redondea(canti)

If Val(TOT_UM_MINIMAS) <> Val(canti) Then
  MsgBox "Cantidad Total No coresponden.  anotar--> Boton : <ACEPTAR>", 48, Pub_Titulo
  flag_lotes_chequea = "A"
  wlot_falso_flag = "1"
  If gridlt.Rows > 2 And gridlt.Visible = True Then
     gridlt.SetFocus
  End If
  Exit Sub
End If

PASA_BOT_ACEPTAR

fralotes.Visible = False
grid_fac.SetFocus
Exit Sub
End Sub

Private Sub cmdltcancelar_Click()
fralotes.Visible = False
End Sub

Private Sub cmdmostrar_Click()
Dim wanaliza As Integer
Dim wanaflag As String * 1
Dim wdescto As String
Dim COND As String
Dim WTIPO As String

Dim PSLOC_PEDIDOS As rdoQuery
Dim loc_pedidos   As rdoResultset
pub_cadena = "SELECT DISTINCT PED_TIPMOV, PED_NUMSER, PED_NUMFAC, PED_CODCLIE, PED_CODVEN, PED_FECHA , PED_CONDI, PED_CODUSU, PED_NOMCLIE, PED_IGV,PED_BRUTO, PED_FBG, ped_situacion , PED_TIPVTA " & _
 "FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = ? " & _
 "AND (PED_FECHA >= ? AND PED_FECHA <= ?) AND PED_CODVEN = ? AND PED_FBG IN (?,?) AND PED_SITUACION NOT IN ('P','E')  ORDER BY ped_fecha, ped_numfac "
Set PSLOC_PEDIDOS = CN.CreateQuery("", pub_cadena)
PSLOC_PEDIDOS(0) = LK_CODCIA
PSLOC_PEDIDOS(1) = 201
PSLOC_PEDIDOS(2) = txtCampo1.Text
PSLOC_PEDIDOS(3) = txtCampo2.Text
PSLOC_PEDIDOS(4) = Val(txtCodVen.Text)
PSLOC_PEDIDOS(5) = ""
PSLOC_PEDIDOS(6) = ""
If chkFacturas.Value = 1 Then PSLOC_PEDIDOS(5) = "F"
If chkBoletas.Value = 1 Then PSLOC_PEDIDOS(6) = "B"

Set loc_pedidos = PSLOC_PEDIDOS.OpenResultset(rdOpenKeyset, rdConcurValues)
'FechaIniFac.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
'FechaIniBol.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
With Grid_detalle
.Rows = 2
.Cols = 10
.TextMatrix(0, 0) = "Fecha"
.TextMatrix(0, 1) = "Usuario"
.TextMatrix(0, 2) = "Serie"
.TextMatrix(0, 3) = "Numero"
.TextMatrix(0, 4) = "CodVen"
.TextMatrix(0, 5) = "TipDoc"
.TextMatrix(0, 6) = "CodCli"
.TextMatrix(0, 7) = "Descripción"
.TextMatrix(0, 8) = "Importe"
.TextMatrix(0, 9) = "(%)Desct"
.ColWidth(0) = 1000
.ColWidth(1) = 800
.ColWidth(2) = 600
.ColWidth(3) = 900
.ColWidth(4) = 500
.ColWidth(5) = 1200
.ColWidth(6) = 700
.ColWidth(7) = 2100
.ColWidth(8) = 1000
.ColWidth(9) = 0
.Visible = False
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
PUB_TIPREG = 65
Do Until loc_pedidos.EOF
  wanaflag = ""
  For wanaliza = 0 To lstcanal.ListCount - 1
   If lstcanal.Selected(wanaliza) Then
    If Val(loc_pedidos!PED_TIPVTA) = Val(Right(lstcanal.List(wanaliza), 6)) Then
       wanaflag = "A"
       Exit For
    End If
   End If
  Next wanaliza
  If wanaflag <> "A" Then GoTo pasa_ped
  
 .Rows = .Rows + 1
 
  If loc_pedidos!PED_CONDI = 1 Then
   COND = "CAS"
  ElseIf loc_pedidos!PED_CONDI = 2 Then
   COND = "CNR"
  ElseIf loc_pedidos!PED_CONDI = 4 Then
   COND = "CRD"
  End If
  
  PUB_NUMTAB = loc_pedidos!PED_TIPVTA
  LEER_TAB_LLAVE
  WTIPO = ""
  If Not tab_llave.EOF Then WTIPO = Left(tab_llave!tab_nomlargo, 5)

 .TextMatrix(.Rows - 1, 0) = Format(loc_pedidos!PED_fecha, "dd/mm/yyyy")
 .TextMatrix(.Rows - 1, 1) = Trim(loc_pedidos!PED_CODUSU)
 .TextMatrix(.Rows - 1, 2) = Format(loc_pedidos!PED_NUMSER, "000")
 .TextMatrix(.Rows - 1, 3) = Format(loc_pedidos!PED_NUMFAC, "000000")
 .TextMatrix(.Rows - 1, 4) = Format(loc_pedidos!PED_codven, "000")
 
  If loc_pedidos!ped_FBG = "F" Then
       .TextMatrix(.Rows - 1, 5) = "Fact. " & COND & " - " & WTIPO
  End If
 If loc_pedidos!ped_FBG = "B" Then
       .TextMatrix(.Rows - 1, 5) = "Bol. " & COND & " - " & WTIPO
  End If

 .TextMatrix(.Rows - 1, 6) = loc_pedidos!PED_CODCLIE
 .TextMatrix(.Rows - 1, 7) = Trim(loc_pedidos!PED_NOMCLIE)
 .TextMatrix(.Rows - 1, 8) = Format(Val(loc_pedidos!PED_IGV) + Val(loc_pedidos!PED_BRUTO), "#,##0.00")
 wdescto = ""
' If Val(loc_pedidos!PED_descto) <> 0 Then
'    wdescto = Format(Val(loc_pedidos!PED_descto), "#0.0#")
'    If Val(loc_pedidos!PED_descto_pre) <> 0 Then wdescto = wdescto + "+"
' End If
' If Val(loc_pedidos!PED_descto_pre) <> 0 Then
'    wdescto = wdescto + Format(Val(loc_pedidos!PED_descto_pre), "#0.0#")
' End If
' If wdescto <> "" Then
'   .TextMatrix(.Rows - 1, 9) = wdescto
' End If
pasa_ped:
 loc_pedidos.MoveNext
 
Loop
 chkFacturas_Click
 chkBoletas_Click
 .Visible = True
End With


End Sub

Private Sub cmdpedidos_Click()
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
Else
Exit Sub
End If

' CLI_OTRO_CONTR

Load FrmControl
FrmControl.d_codigo.Caption = i_codcli.Text
FrmControl.d_nombre.Caption = ""
FrmControl.d_fecha.Caption = i_fecha_compra.Text
FrmControl.d_serie.Caption = i_numser_c.Text
FrmControl.d_numero.Caption = i_numfac_c.Text
FrmControl.d_numser.Caption = i_numser.Text
FrmControl.d_numfac.Caption = i_numfac.Text
FrmControl.d_nombre.Caption = FORMGEN.i_nomCLI.Caption

FrmControl.flag_signo.Caption = "1"

FrmControl.Show 1


Exit Sub
If LK_CODTRA <> 2401 Then
  MsgBox "No Procede para Esta Transaccion Solo Para Ventas", 48, Pub_Titulo
  Exit Sub
End If
txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
Grid_detalle.Rows = 2
PUB_TIPREG = 65
PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_TAB_LLAVE
lstcanal.Clear
Do Until tab_mayor.EOF
   lstcanal.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
   lstcanal.Selected(lstcanal.ListCount - 1) = True
   tab_mayor.MoveNext
Loop
fraPedidos.Visible = True
Azul2 txtCampo1, txtCampo1

End Sub


Private Sub cmdtemp_Click()
If LK_CODTRA = 2401 Or LK_CODTRA = 1401 Then
Else
Exit Sub
End If

Dim xc As Integer
Dim I As Integer
Dim var_c As String
Dim WINI As Integer
Dim WFIN As Integer
Dim RUTA As String
Dim xcol As Integer
Dim xcont As Integer
Dim WLINEA As String
RUTA = "C:\" & "MovLog.txt"
WLINEA = ""
msgrid.Cols = grid_fac.Cols
msgrid.Rows = 1
On Error GoTo SALE
Open RUTA For Input As #1
xc = 0
msgrid.Clear
Do While Not EOF(1)
  xc = xc + 1
  msgrid.Rows = msgrid.Rows + 1
  Line Input #1, WLINEA
  'msgrid.Visible = True
  var_c = ""
  WINI = 1
  WFIN = 1
  xcol = 0
  xcont = 0
  For I = 1 To Len(Trim(WLINEA))
    xcont = xcont + 1
    If Mid(WLINEA, I, 1) = "|" Then
      WFIN = xcont - 1
      var_c = Mid(Trim(WLINEA), WINI, WFIN)
      WINI = I + 1
      If var_c = "@" Then
      Else
        msgrid.TextMatrix(xc, xcol) = var_c
      End If
      xcont = 0
      xcol = xcol + 1
     End If
     
  Next I
Loop
Close #1
 msgrid.Visible = True
 msgrid.SetFocus
 Exit Sub
SALE:
 Close #1
 Resume Next
End Sub

Private Sub cmdtipo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
NUMERO = FORMGEN.cmdtipo.WhatsThisHelpID
avanza_campo

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
    SendKeys "%{UP}"
    DoEvents
End If

End Sub

Private Sub cta_gas1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If Trim(Left(cta_gas1.Text, 1)) = "*" Then
       BUSCAR_CTA 1, cta_gas1
       Exit Sub
     End If
     Azul imp_gas1, imp_gas1
End If


End Sub

Private Sub cta_gas2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul imp_gas2, imp_gas2
End If

End Sub


Private Sub Diario_Click()
Dim WS_SALDO As Currency
Dim Tit As String
Dim I As Integer
Dim success%
Dim flag_cuenta As Integer
If FLAG_DIARIO = "A" Then Exit Sub
If Left(Diario.Caption, 3) <> "Dia" Then
   Diario.Caption = "Diario"
   FLAG_DIARIO = "A"
   Exit Sub
Else
   FLAG_DIARIO = ""
   Diario.Caption = "Detener"
End If


i_num_oper.Text = ""
Screen.MousePointer = 11
'success% = SetWindowPos(FrmProcesa.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
'FrmProcesa.Show
'MUESTRA_PROCESO "Un Momento ..."
SQ_OPER = 1
If IsDate(i_fecha_oper.Text) Then
   PUB_FECHA = i_fecha_oper.Text
Else
   i_fecha_oper.Text = Format(LK_FECHA_DIA, "dd/mm/yy")
   PUB_FECHA = LK_FECHA_DIA
End If
pu_codcia = LK_CODCIA

LEER_ALL_LLAVE

fila = 0
If all_llave.EOF = True Then
   'FrmProcesa.Hide
   FLAG_DIARIO = ""
   MsgBox "No hay Registros....", 48, Pub_Titulo
   Screen.MousePointer = 0
   GoTo fin
End If

Barra.Min = 0
Barra.Value = 0
Barra.max = all_llave.RowCount
Barra.Visible = True

FORMGEN.Grid_all.Visible = False
Tit = FORMGEN.Grid_all.FormatString
Grid_all.Clear
Grid_all.Rows = 1
FORMGEN.Grid_all.FormatString = Tit
Grid_all.Cols = 23
'Grid_all.Cols = Grid_all.Cols + 1
FORMGEN.Grid_all.TextMatrix(0, 10) = "Hora"
FORMGEN.Grid_all.TextMatrix(0, 11) = "Fec.Proc."
FORMGEN.Grid_all.ColWidth(1) = 800
FORMGEN.Grid_all.ColWidth(8) = 1200
FORMGEN.Grid_all.ColWidth(9) = 500
FORMGEN.Grid_all.ColWidth(10) = 500
FORMGEN.Grid_all.ColWidth(11) = 800
FORMGEN.Grid_all.ColWidth(14) = 0
FORMGEN.Grid_all.ColWidth(15) = 0
FORMGEN.Grid_all.ColWidth(16) = 0
FORMGEN.Grid_all.ColWidth(17) = 0
FORMGEN.Grid_all.ColWidth(18) = 0
FORMGEN.Grid_all.ColWidth(19) = 0
FORMGEN.Grid_all.ColWidth(20) = 0
FORMGEN.Grid_all.ColWidth(21) = 0
FORMGEN.Grid_all.ColWidth(22) = 0

If LK_FLAG_SOS = "A" Then
  FORMGEN.Grid_all.ColWidth(0) = 0
  FORMGEN.Grid_all.ColWidth(1) = 0
End If
Do Until all_llave.EOF
    DoEvents
    If FLAG_DIARIO = "A" Then GoTo TERMINA
'    If LK_FLAG_SOS = "A" And all_llave!ALL_FLAG_SO <> "A" Then GoTo OTRO
    Barra.Value = Barra.Value + 1
    If Trim(FORMGEN.i_CODUSU.Text) = "" Then
    Else
       If Trim(FORMGEN.i_CODUSU.Text) <> Trim(all_llave!all_codusu) Then
          GoTo OTRO
       End If
    End If
    If Trim(FORMGEN.i_numguia.Text) = "" Then
    Else
       If Val(FORMGEN.i_numguia.Text) <> Val(all_llave!ALL_codtra) Then
          GoTo OTRO
       End If
    End If
    If LK_CODCIA <> all_llave!all_CODCIA Then
       GoTo OTRO
    End If
    If Val(i_importe.Text) <> 0 Then
    If (Val(FORMGEN.i_importe.Text) = Val(all_llave!ALL_IMPORTE_amort)) Or (Val(FORMGEN.i_importe.Text) = Val(all_llave!ALL_IMPORTE)) Then
    Else
       GoTo OTRO
    End If
    End If
   
    If Val(i_TEXTONCRE.Text) <> 0 Then
    If (Val(FORMGEN.i_TEXTONCRE.Text) = Val(all_llave!ALL_NUMOPER)) Then
    Else
       GoTo OTRO
    End If
    End If
    If Val(i_numfac_c.Text) <> 0 Then
    If (Val(FORMGEN.i_numfac_c.Text) = Val(all_llave!all_numfac)) Then
    Else
       GoTo OTRO
    End If
    End If

       fila = fila + 1
       FORMGEN.Grid_all.Rows = fila + 1
       FORMGEN.Grid_all.Row = fila
       FORMGEN.Grid_all.TextMatrix(fila, 0) = all_llave!ALL_NUMOPER
       If all_llave!all_flag_ext = "E" Or Val(all_llave!ALL_codtra) = 1111 Then
          FORMGEN.Grid_all.COL = 0
          FORMGEN.Grid_all.CellBackColor = vb3DLight
       End If
       
       FORMGEN.Grid_all.TextMatrix(fila, 1) = Nulo_Valor0(all_llave!ALL_numoper2)
       If all_llave!all_flag_ext = "E" Or all_llave!ALL_codtra = 1111 Then
          FORMGEN.Grid_all.COL = 1
          FORMGEN.Grid_all.CellBackColor = vb3DLight
       End If
       
       FORMGEN.Grid_all.TextMatrix(fila, 2) = all_llave!ALL_codtra
       If all_llave!all_flag_ext = "E" Or Val(all_llave!ALL_codtra) = 1111 Then
          FORMGEN.Grid_all.COL = 2
          FORMGEN.Grid_all.CellBackColor = vb3DLight
       End If
       
       FORMGEN.Grid_all.TextMatrix(fila, 3) = all_llave!all_codusu
       If all_llave!all_flag_ext = "E" Or Val(all_llave!ALL_codtra) = 1111 Then
          FORMGEN.Grid_all.COL = 3
          FORMGEN.Grid_all.CellBackColor = vb3DLight
       End If
       
       FORMGEN.Grid_all.COL = 4
       If all_llave!all_flag_ext = "E" Or Val(all_llave!ALL_codtra) = 1111 Then
          FORMGEN.Grid_all.CellBackColor = vb3DLight
       End If
       
       pu_cp = all_llave!ALL_CP
       pu_codclie = Val(all_llave!ALL_CODCLIE)
       If pu_codclie <> 0 Then
          SQ_OPER = 1
          pu_codcia = LK_CODCIA
          LEER_CLI_LLAVE
          If cli_llave.EOF Then
             FORMGEN.Grid_all.Text = ".........."
          Else
             FORMGEN.Grid_all.Text = cli_llave(3)
          End If
       End If
       
       
       FORMGEN.Grid_all.COL = 5
       If all_llave!ALL_codtra = 2580 Then
         FORMGEN.Grid_all.Text = all_llave!ALL_LIMCRE_ACT
       Else
         FORMGEN.Grid_all.Text = all_llave!ALL_IMPORTE_amort
       End If
       If all_llave!all_flag_ext = "E" Or all_llave!ALL_codtra = 1111 Then
          FORMGEN.Grid_all.CellBackColor = vb3DLight
       End If
       FORMGEN.Grid_all.COL = 6
       If all_llave!ALL_codtra = 2580 Then
         FORMGEN.Grid_all.Text = all_llave!ALL_LIMCRE_ANT
       Else
         FORMGEN.Grid_all.Text = all_llave!ALL_IMPORTE
       End If
       If all_llave!all_flag_ext = "E" Or all_llave!ALL_codtra = 1111 Then
          FORMGEN.Grid_all.CellBackColor = vb3DLight
       End If
       
       FORMGEN.Grid_all.COL = 7
       If Val(all_llave!all_numfac) <> 0 Then
          FORMGEN.Grid_all.Text = Nulo_Valors(all_llave!ALL_FBG) & "./ " & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
       Else
         If Val(all_llave!ALL_NUMDOC) <> 0 Then
           FORMGEN.Grid_all.Text = all_llave!ALL_NUMDOC
         End If
         If all_llave!ALL_SIGNO_CCM = 0 And (all_llave!ALL_codtra = 2735 Or all_llave!ALL_codtra = 2748 Or all_llave!ALL_codtra = 5714) Then
            FORMGEN.Grid_all.Text = all_llave!all_numfac
         End If
       End If
       
       
       If all_llave!all_flag_ext = "E" Or Val(all_llave!ALL_codtra) = 1111 Then
          FORMGEN.Grid_all.CellBackColor = vb3DLight
       End If
       
       FORMGEN.Grid_all.COL = 8
       If all_llave!ALL_SIGNO_CCM <> 0 Then
          SQ_OPER = 1
          PUB_CODBAN = all_llave!all_codban
          pu_codcia = all_llave!all_CODCIA
          LEER_CCM_LLAVE
          If Not ccm_llave.EOF Then
            FORMGEN.Grid_all.Text = Trim(Format(all_llave!all_chenum, "00000000") & " /" & Trim(ccm_llave!ccm_nomcorto))
          End If
       End If
       If all_llave!all_flag_ext = "E" Or Val(all_llave!ALL_codtra) = 1111 Then
          FORMGEN.Grid_all.CellBackColor = vb3DLight
       End If
       
       FORMGEN.Grid_all.COL = 9
       FORMGEN.Grid_all.Text = all_llave!ALL_CODVEN
       If all_llave!all_flag_ext = "E" Or all_llave!ALL_codtra = 1111 Then
          FORMGEN.Grid_all.CellBackColor = vb3DLight
       End If
       
       FORMGEN.Grid_all.COL = 10
       FORMGEN.Grid_all.Text = DatePart("h", Nulo_Valor0(all_llave!ALL_HORA)) & ":" & DatePart("n", Nulo_Valor0(all_llave!ALL_HORA))
       If all_llave!all_flag_ext = "E" Or Val(all_llave!ALL_codtra) = 1111 Then
          FORMGEN.Grid_all.CellBackColor = vb3DLight
       End If
       FORMGEN.Grid_all.COL = 11
       FORMGEN.Grid_all.Text = all_llave!ALL_FECHA_DIA
       
        FORMGEN.Grid_all.COL = 12
        FORMGEN.Grid_all.Text = all_llave!ALL_MONEDA_CAJA
        FORMGEN.Grid_all.COL = 13
        FORMGEN.Grid_all.Text = all_llave!ALL_TIPO_CAMBIO
       'If all_llave!ALL_tipmov = 10 And all_llave!ALL_SIGNO_CAR <> 0 Then
        FORMGEN.Grid_all.COL = 14
        FORMGEN.Grid_all.Text = all_llave!ALL_TIPDOC
        FORMGEN.Grid_all.COL = 15
        FORMGEN.Grid_all.Text = Nulo_Valor0(all_llave!ALL_serdoc)
        FORMGEN.Grid_all.COL = 16
        FORMGEN.Grid_all.Text = all_llave!ALL_NUMDOC
       'End If
       FORMGEN.Grid_all.COL = 17
       FORMGEN.Grid_all.Text = all_llave!ALL_CODCLIE
       FORMGEN.Grid_all.COL = 18
       FORMGEN.Grid_all.Text = all_llave!ALL_SIGNO_CAR
       FORMGEN.Grid_all.COL = 19
       If Not IsNull(all_llave!ALL_FECHA_SUNAT) Then FORMGEN.Grid_all.Text = all_llave!ALL_FECHA_SUNAT
       FORMGEN.Grid_all.COL = 20
       FORMGEN.Grid_all.Text = all_llave!all_numser_c
       FORMGEN.Grid_all.COL = 21
       FORMGEN.Grid_all.Text = all_llave!all_numfac_c
       FORMGEN.Grid_all.COL = 22
       FORMGEN.Grid_all.Text = all_llave!ALL_CP
OTRO:

       all_llave.MoveNext
Loop
TERMINA:
If fila = 0 Then
  Diario.Caption = "Diario"
  FLAG_DIARIO = ""
  Barra.Visible = False
  Screen.MousePointer = 0
  MsgBox "NO Existe Datos para esta consulta.", 48, Pub_Titulo
  Azul2 i_fecha_oper, i_fecha_oper
Else
  Diario.Caption = "Diario"
  FLAG_DIARIO = ""
  FORMGEN.Grid_all.Visible = True
  FORMGEN.Grid_all.COL = 1
  If FORMGEN.Grid_all.Rows > 1 Then
    FORMGEN.Grid_all.Row = Grid_all.Rows - 1
  End If
  FORMGEN.Grid_all.SetFocus
  Barra.Visible = False
  Screen.MousePointer = 0
End If
Diario.Caption = "Diario"
FLAG_DIARIO = ""
fin:
Diario.Caption = "Diario"
FLAG_DIARIO = ""
End Sub

Private Sub Diario_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.Diario.WhatsThisHelpID
avanza_campo
fin:

End Sub

Private Sub Form_Activate()
Unload RCRYSTAL
End Sub

Private Sub Form_Click()
FORMGEN.WhatsThisMode
End Sub

 Private Sub Form_DblClick()
''' CASO DE ESPECIAL ACV ICA
  REPROCESO
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim iFormCount As Integer
If KeyCode = 112 Then ' F1
    Load frmsucesos
    frmsucesos.Show 1
    Exit Sub
End If
If KeyCode = 120 Then ' F9
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      MDIForm1.Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda nombre de articulo."
      LK_FLAG_ALTERNO = " "
      MDIForm1.Toolbar1.Buttons.Item(12).Image = 18 '********** tenr en cuent
      'DoEvents
    Else
      MDIForm1.Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda sub-codigo de articulo."
      LK_FLAG_ALTERNO = "A"
      MDIForm1.Toolbar1.Buttons.Item(12).Image = 5
    End If
   If Forms.count - 1 > 0 Then
     For iFormCount = Forms.count - 1 To 1 Step -1
        If iFormCount <> 1 Then
           If UCase(Forms(iFormCount).Name) = "FRMARTI" Then
                'MsgBox UCase(Forms(iFormCount).Name)
                frmARTI.PROCESO_CANCELAR
                frmARTI.PROCESO_ARTI
           ElseIf UCase(Forms(iFormCount).Name) = "KARDEX" Then
                KARDEX.ListView1.Visible = False
                KARDEX.i_codart2.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FRMPREUNIDAD" Then
                'FrmPreUnidad.PRO_GOTFOCUS
                'FrmPreUnidad.ListView1.Visible = False
                'FrmPreUnidad.Txt_key.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FRMREPO" Then
                FrmRepo.ListView1.Visible = False
                FrmRepo.Txt_key.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FORMGEN" Then
                FORMGEN.LV_ART.Visible = False
                FORMGEN.fraprecios.Visible = False
                FORMGEN.TEXTOVAR.Text = ""
           End If
        End If
       Next iFormCount
   End If
   
 Exit Sub
End If
If KeyCode = 116 Then ' F5
 Azul TRANS, TRANS
 Exit Sub
End If
If KeyCode = 115 Then ' F4
 If cancelar.Enabled Then cancelar_Click
 Exit Sub
End If
If KeyCode = 114 Then ' F3
 If i_def.Visible Then
    'i_codcli.Text = ""
    i_def.Enabled = True
    i_def.SetFocus
    SendKeys "%{DOWN}"
 End If
 Exit Sub
End If
If KeyCode = 117 Then ' F6
 If grabar.Enabled Then grabar_Click
 Exit Sub
End If

End Sub

Private Sub Form_Load()
Dim ws_indice As Integer
Dim ws_ancho As Integer
cmddd.Visible = False
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
  cmddd.Visible = True
End If
flag_lotes_chequea = ""
frafondo.Height = 5895
loc_mofi_venta = ""
loc_acceso_des = ""
For fila = 1 To lk_OTROS_Count
   If Val(lk_OTROS(fila)) = 5 Then 'modificar Venta
     loc_mofi_venta = "A"
   End If
   If Val(lk_OTROS(fila)) = 12 Then 'Acceso a Descto.
     loc_acceso_des = "A"
   End If
Next fila
loc_doucumento = ""
PUB_ACU_TOTAL = 0
flag_del = 0
'DoEvents
LK_QUEDA = ""
If LK_FECHA_DIA = #1/1/1900# Then
  Screen.MousePointer = 0
  MsgBox "La Compañia No Tiene Definida su Fecha de Trabajo. Verificar!!!", 48, Pub_Titulo
  Unload FORMGEN
  Exit Sub
End If
FORMGEN.grid_fac.FontName = "Arial"
FORMGEN.grid_fac.FontSize = 8
PUB_PROCESO = 0

FORMGEN.grid_fac.Cols = 50
FORMGEN.grid_fac.ColWidth(0) = 3700 'nombre
If LK_EMP = "PIU" Then
  FORMGEN.grid_fac.ColWidth(1) = 1100 'codigo
Else
  FORMGEN.grid_fac.ColWidth(1) = 800 'codigo
End If
FORMGEN.grid_fac.ColWidth(2) = 600 'jabas
FORMGEN.grid_fac.ColWidth(3) = 400 'unidades
FORMGEN.grid_fac.ColWidth(4) = 900 'cantidad
FORMGEN.grid_fac.ColWidth(5) = 1000 'unidades
FORMGEN.grid_fac.ColWidth(6) = 1000 'precio
FORMGEN.grid_fac.ColWidth(7) = 1300 'subtotal
FORMGEN.grid_fac.ColWidth(8) = 600 'flete
FORMGEN.grid_fac.ColWidth(9) = 2000 'mortalidad
FORMGEN.grid_fac.ColWidth(10) = 700 'FLAG_DESCTO=D
FORMGEN.grid_fac.ColWidth(11) = 900 'COSTO PROMEDIO
FORMGEN.grid_fac.ColWidth(12) = 700  'signo del arm + -1
FORMGEN.grid_fac.ColWidth(13) = 0 'STOCK
FORMGEN.grid_fac.ColWidth(14) = 0  'EQUIV
FORMGEN.grid_fac.ColWidth(15) = 0 'DESCRIP. UNIDAD
FORMGEN.grid_fac.ColWidth(16) = 0 'codigo original
FORMGEN.grid_fac.ColWidth(17) = 0 'NUMERO DE PRECIO 1.2,3,4,5
FORMGEN.grid_fac.ColWidth(18) = 0 'relacion con un pedido
FORMGEN.grid_fac.ColWidth(19) = 0 'orden en i_precios
FORMGEN.grid_fac.ColWidth(20) = 0 'NUMSEC PARA UPDATE DE PEDIDOS
FORMGEN.grid_fac.ColWidth(21) = 0 'FLAG DE STOCK
FORMGEN.grid_fac.ColWidth(22) = 0 'numero de autorizacion
FORMGEN.grid_fac.ColWidth(23) = 0 'FLAG DE EXONERACION IGV
FORMGEN.grid_fac.ColWidth(24) = 0 '% IGV DE EXONERACION
FORMGEN.grid_fac.ColWidth(25) = 0 'FLAG PARA COMPRA SUBTOTAL DIGITADO
FORMGEN.grid_fac.ColWidth(26) = 0 'RECEP. TIPMOV
FORMGEN.grid_fac.ColWidth(27) = 0 'RECEP. CODCIA
FORMGEN.grid_fac.ColWidth(28) = 0 'RECEP. NUMSER
FORMGEN.grid_fac.ColWidth(29) = 0 'RECEP. FBG
FORMGEN.grid_fac.ColWidth(30) = 0 'RECEP. NUMFAC
FORMGEN.grid_fac.ColWidth(31) = 0 'RECEP. NUMSEC
FORMGEN.grid_fac.ColWidth(32) = 0 'precio de venta cash dpv
FORMGEN.grid_fac.ColWidth(33) = 0 'PRE_CODART
FORMGEN.grid_fac.ColWidth(34) = 300 'PESO TOTAL X ART. COL37 * COL4
FORMGEN.grid_fac.ColWidth(35) = 0 'precio BASE
FORMGEN.grid_fac.ColWidth(36) = 0 'DISPONIBLE
FORMGEN.grid_fac.ColWidth(37) = 0 'PESO DEL PROD.
FORMGEN.grid_fac.ColWidth(38) = 0 'Flag Max Item
FORMGEN.grid_fac.ColWidth(39) = 0 'Nro. Serie para relación 1405
FORMGEN.grid_fac.ColWidth(40) = 0 'Nro. Documento para relación 1405
FORMGEN.grid_fac.ColWidth(41) = 0 'Nro. secuancia  relación 1405
FORMGEN.grid_fac.ColWidth(42) = 0 'campo de los desctos.
FORMGEN.grid_fac.ColWidth(43) = 0 'campo para los litros
FORMGEN.grid_fac.ColWidth(44) = 0 'Solo campo para el Porc. del Numero de precio que se escoja
FORMGEN.grid_fac.ColWidth(45) = 0 'Solo campo para el Porc. del Numero de precio que se escoja
FORMGEN.grid_fac.ColWidth(46) = 0 'PARA EL EXTORNO ASIGNA EL FAR_ESTADO2
FORMGEN.grid_fac.ColWidth(47) = 0 'solo para los extornos asigna el codigo de lote
FORMGEN.grid_fac.ColWidth(48) = 0 'ALMACENA LA CANTIDAD POR CADA LOTE PARA EL EXTORNO
FORMGEN.grid_fac.ColWidth(49) = 0 'ALMACENA EL DESCTO EN EFECTIVO DEL PRODUCTO


FORMGEN.gridl.Cols = 28
FORMGEN.gridl.ColWidth(0) = 3000
FORMGEN.gridl.ColWidth(1) = 500
FORMGEN.gridl.ColWidth(2) = 500
FORMGEN.gridl.ColWidth(3) = 700
FORMGEN.gridl.ColWidth(4) = 1200
FORMGEN.gridl.ColWidth(5) = 1000
FORMGEN.gridl.ColWidth(6) = 1000
FORMGEN.gridl.ColWidth(7) = 500
FORMGEN.gridl.ColWidth(8) = 1000
FORMGEN.gridl.ColWidth(9) = 1000
FORMGEN.gridl.ColWidth(10) = 0
FORMGEN.gridl.ColWidth(11) = 0
FORMGEN.gridl.ColWidth(12) = 0
FORMGEN.gridl.ColWidth(13) = 0
FORMGEN.gridl.ColWidth(14) = 0
FORMGEN.gridl.ColWidth(15) = 0
FORMGEN.gridl.ColWidth(16) = 0
FORMGEN.gridl.ColWidth(17) = 0
FORMGEN.gridl.ColWidth(18) = 400
FORMGEN.gridl.ColWidth(19) = 1500
FORMGEN.gridl.ColWidth(20) = 1500
FORMGEN.gridl.ColWidth(21) = 1500
FORMGEN.gridl.ColWidth(22) = 1500
FORMGEN.gridl.ColWidth(23) = 2500
FORMGEN.gridl.ColWidth(24) = 0 ' CAR_MONEDA
FORMGEN.gridl.ColWidth(25) = 0 ' CAR_IMPORTE ESPEJO
FORMGEN.gridl.ColWidth(26) = 0 ' CAR_fecha vcto espejo

   If LK_EMP = "PAR" Then
    gridl.ColWidth(7) = 0
    gridl.ColWidth(8) = 0
    gridl.ColWidth(9) = 0
    gridl.ColWidth(10) = 0
    gridl.ColWidth(11) = 0
    gridl.ColWidth(12) = 0
    gridl.ColWidth(13) = 0
    gridl.ColWidth(14) = 0
    gridl.ColWidth(15) = 0
    gridl.ColWidth(16) = 0
    gridl.ColWidth(17) = 0
    gridl.ColWidth(18) = 0
    gridl.ColWidth(19) = 0
    gridl.ColWidth(20) = 0
    gridl.ColWidth(21) = 0
   End If



FORMGEN.gridl.RowHeight(0) = 300
FORMGEN.gridl.RowHeight(1) = 350

pasa_cabeza_L

pasa_cabeza_L2

FORMGEN.grid_canje.Cols = 8
FORMGEN.grid_canje.ColWidth(0) = 0 'codigo cliente
FORMGEN.grid_canje.ColWidth(1) = 450 'orden
FORMGEN.grid_canje.ColWidth(2) = 400 'tipo doc
FORMGEN.grid_canje.ColWidth(3) = 1000 'importe
FORMGEN.grid_canje.ColWidth(4) = 1100 'fecha
FORMGEN.grid_canje.ColWidth(5) = 1300 'banco
FORMGEN.grid_canje.ColWidth(6) = 800 'n.doc
FORMGEN.grid_canje.ColWidth(7) = 1200 'anotacion

FORMGEN.grid_canje.RowHeight(2) = 250


FORMGEN.grid_canje.RowHeight(0) = 250

grid_fac.RowHeight(1) = 300
grid_fac.RowHeight(2) = 300


FORMGEN.Frame2.Top = 3100
label_nomart.Top = 2800
label_nomart.Top = Frame2.Top - 220
label_precio.Top = Frame2.Top - 220
label_nomart.Left = 0
label_precio.Left = 6000
label_nomart.Visible = False
label_precio.Visible = False

'FORMGEN.i_image_llave.Left = 7000
'FORMGEN.i_image_llave.Top = 3000
'FORMGEN.i_image_llave.Visible = True

FORMGEN.grid_fac.ColAlignment(0) = 1
FORMGEN.grid_fac.ColAlignment(1) = 1
FORMGEN.grid_fac.ColAlignment(5) = 1

FORMGEN.grid_autorizacion.FontName = "Sans Serif"
FORMGEN.grid_autorizacion.FontSize = 10
FORMGEN.grid_canje.FontName = "Courier New"
FORMGEN.grid_canje.FontSize = 10

FORMGEN.Grid_all.FontName = "Sans Serif"
FORMGEN.Grid_all.FontSize = 7




FORMGEN.Grid_all.Top = 2280
FORMGEN.Grid_all.Left = 10
FORMGEN.Grid_all.Width = 10100
FORMGEN.Grid_all.Height = 4000

FORMGEN.i_concepto.Width = 3000
FORMGEN.i_nompro.Width = 3000
FORMGEN.i_nombre_banco.Width = 1500

'FORMGEN.nomtra.Width = 2500




'DoEvents




FORMGEN.gridl.Top = 1500
FORMGEN.gridl.Left = 0
FORMGEN.gridl.Width = 7500
FORMGEN.gridl.Height = 2000
grid_autorizacion.Left = 0
grid_autorizacion.Top = 2500
grid_autorizacion.Width = 7400
grid_autorizacion.Height = 1500




'DoEvents
numarchi = 2
PS_GEN(0) = 0
GEN.Requery

LK_COBRADOR = 1
PUB_FECHA = LK_FECHA_DIA

PUB_TIPREG = 999
'PUB_TIPREG = 777
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
         ws_indice = tab_mayor!tab_numtab
         TABLA_TAG(ws_indice) = Val(tab_mayor!tab_nomcorto)
         tab_mayor.MoveNext
Loop




Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = 10
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
        FORMGEN.i_cant_cheq.AddItem tab_mayor!tab_nomcorto
        tab_mayor.MoveNext
    Loop
    
    FORMGEN.i_codsunat.Clear
    PUB_TIPREG = 50
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
        FORMGEN.i_codsunat.AddItem tab_mayor!tab_nomlargo & String(40, " ") & tab_mayor!tab_numtab
        tab_mayor.MoveNext
    Loop
    If FORMGEN.i_codsunat.ListCount <> 0 Then FORMGEN.i_codsunat.ListIndex = 0
    
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
    
    
    
    CONTA = -1
    PUB_TIPREG = 52
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
        FORMGEN.i_mortal.AddItem Left(tab_mayor!tab_nomlargo, 20) & "  " & tab_mayor!tab_numtab
        tab_mayor.MoveNext
    Loop
    CONTA = 0
    PUB_TIPREG = 45
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until CONTA = 5
        CONTA = CONTA + 1
        If tab_mayor.EOF Then
           tab_precioss(CONTA) = "....."
        Else
           tab_precioss(CONTA) = tab_mayor!tab_nomlargo
           tab_mayor.MoveNext
        End If
    Loop
    
    PUB_CODBAN = 0
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    LEER_CCM_LLAVE
    FORMGEN.i_bancos.AddItem "Ninguno  ....                       "
    Do Until ccm_mayor.EOF
        If ccm_mayor!CCM_CODBAN <> 1 Then FORMGEN.i_bancos.AddItem Left(ccm_mayor!CCM_NOMBRE, 30) & "                              " & ccm_mayor!CCM_CODBAN
        ccm_mayor.MoveNext
    Loop
    
   
If LK_FLAG_FACTURACION = "V" Then i_fbg.AddItem "P"
If LK_EMP = "CAM" Then i_fbg.AddItem "P"
If LK_FLAG_SOS = "A" Or LK_EMP = "3AA" Then
  i_fbg.Clear
  i_fbg.AddItem "F"
  i_fbg.AddItem "B"
End If

   textovar_canje.Width = 30
   textovar_canje.Height = 30


i_def.Visible = False
grabar.Enabled = False
cancelar.Enabled = False
tab_derL(5) = 6
tab_derL(6) = 7
tab_derL(7) = 5
LOC_MONEDA_DIG = ""
'CABEZA DE GRID_PRECIOS
Dim fil As Integer
grid_unid.Clear
grid_unid.Cols = 16
grid_unid.Rows = 2
grid_unid.ColWidth(0) = 1300 ' unidad
grid_unid.ColWidth(1) = 700  ' equivalencia
grid_unid.ColWidth(2) = 1  'c. Repos.
grid_unid.ColWidth(3) = 900  ' cos. base
grid_unid.ColWidth(4) = 550   ' % 1
grid_unid.ColWidth(5) = 650   ' p 1
grid_unid.ColWidth(6) = 550   ' % 2
grid_unid.ColWidth(7) = 650   ' p 2
grid_unid.ColWidth(8) = 550   ' % 3
grid_unid.ColWidth(9) = 650   ' p 3
grid_unid.ColWidth(10) = 550   ' % 4
grid_unid.ColWidth(11) = 650   ' p 4
grid_unid.ColWidth(12) = 550   ' % 5
grid_unid.ColWidth(13) = 650   ' p 5
grid_unid.ColWidth(14) = 1   ' FLAG
grid_unid.ColWidth(15) = 1   ' SECUENCIA
grid_unid.TextMatrix(1, 0) = "Unidades"
grid_unid.TextMatrix(1, 1) = "Equiv."
grid_unid.TextMatrix(1, 2) = "C.Rep."
grid_unid.TextMatrix(1, 3) = "REP.C/IGV"
grid_unid.TextMatrix(1, 4) = "( % )  "
grid_unid.TextMatrix(1, 5) = "  Valor"
grid_unid.TextMatrix(1, 6) = "( % )  "
grid_unid.TextMatrix(1, 7) = "  Valor"
grid_unid.TextMatrix(1, 8) = "( % )  "
grid_unid.TextMatrix(1, 9) = "  Valor"
grid_unid.TextMatrix(1, 10) = "( % )  "
grid_unid.TextMatrix(1, 11) = "  Valor"
grid_unid.TextMatrix(1, 12) = "( % )  "
grid_unid.TextMatrix(1, 13) = "  Valor"
grid_unid.MergeRow(0) = True
    SQ_OPER = 2
    PUB_TIPREG = 45
    PUB_CODCIA = LK_CODCIA
    LEER_TAB_LLAVE
    fil = 4
    Do Until tab_mayor.EOF
      grid_unid.TextMatrix(0, fil + tab_mayor!tab_numtab - 1) = Trim(tab_mayor!tab_nomlargo)
      grid_unid.TextMatrix(0, fil + tab_mayor!tab_numtab) = Trim(tab_mayor!tab_nomlargo)
      tab_mayor.MoveNext
      fil = fil + 1
   Loop
   ws_ancho = grid_unid.ColWidth(0) + grid_unid.ColWidth(1) + grid_unid.ColWidth(2) + grid_unid.ColWidth(3)
   For fil = 4 To 13
       If Trim(grid_unid.TextMatrix(0, fil)) = "" Then grid_unid.ColWidth(fil) = 0
       ws_ancho = ws_ancho + grid_unid.ColWidth(fil)
   Next fil
  fcomun.Top = 400
  fcomun.Width = ws_ancho
  grid_unid.Width = ws_ancho
  grid_unid.Left = 0
  fcomun.Left = 9500 - ws_ancho
  fcomun.Height = 1400
  grid_unid.Height = 800
  If FORMGEN.i_ds.Visible And FORMGEN.i_ds.ListCount > 0 Then
    FORMGEN.i_ds.ListIndex = 0
  End If
  i_numfac.BackColor = QBColor(7)
  i_numser.BackColor = QBColor(7)
  
  pub_cadena = "SELECT FAR_SERGUIA, FAR_NUMGUIA, FAR_FBG, FAR_NUMFAC FROM FACART WHERE FAR_SERGUIA = ? AND FAR_NUMGUIA = ? AND FAR_CODCIA = ? AND FAR_ESTADO='N' "
  Set PSGUI_MENOR = CN.CreateQuery("", pub_cadena)
  PSGUI_MENOR.rdoParameters(0) = 0
  PSGUI_MENOR.rdoParameters(1) = 0
  PSGUI_MENOR.rdoParameters(2) = " "
  Set gui_menor = PSGUI_MENOR.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
  If LK_EMP = "HER" Then
   pub_cadena = "SELECT CAA_CODCIA, CAA_RECIBO  FROM CARACU WHERE CAA_CODCIA = ? AND CAA_SERIE = ? AND CAA_CP = 'C' AND CAA_ESTADO<>'E' ORDER BY CAA_RECIBO DESC"
   Set PSCAR_MENOR2 = CN.CreateQuery("", pub_cadena)
   PSCAR_MENOR2(0) = ""
   PSCAR_MENOR2(1) = 0
   PSCAR_MENOR2.MaxRows = 1
   Set car_menor2 = PSCAR_MENOR2.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
  End If
  
  pub_cadena = "SELECT ALL_CODCIA, ALL_CODTRA, ALL_SECUENCIA, ALL_SERIE_REC, ALL_NUM_RECIBO  FROM ALLOG WHERE ALL_CODCIA =? AND ALL_CODTRA = ? AND ALL_SECUENCIA = ? AND ALL_SERIE_REC = ? AND ALL_FLAG_EXT <>'E' ORDER BY ALL_NUM_RECIBO DESC"
  Set PSALL_REC = CN.CreateQuery("", pub_cadena)
  PSALL_REC(0) = ""
  PSALL_REC(1) = 0
  PSALL_REC(2) = 0
  PSALL_REC(3) = 0
  PSALL_REC.MaxRows = 1
  Set ALL_REC = PSALL_REC.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

  
  
  
  pub_cadena = "SELECT * FROM HISPRE WHERE HPR_CODCIA = ? AND HPR_FECHA = ? ORDER BY HPR_NUMOPER DESC"
  Set PSNEW_PRECIO = CN.CreateQuery("", pub_cadena)
  PSNEW_PRECIO(0) = 0
  PSNEW_PRECIO(1) = Date
  PSNEW_PRECIO.MaxRows = 1
  Set loc_new_llave = PSNEW_PRECIO.OpenResultset(rdOpenForwardOnly, rdConcurValues)
  
  
  Tipo_Cambio.Caption = "T/C = " & LK_TIPO_CAMBIO
  FORMGEN.i_cant_cheq.Text = ""
  
End Sub

Public Sub PROCESA_CELDAS2()
If grid_canje.COL < 3 Then Exit Sub
If grid_canje.COL = 5 Then Exit Sub
textovar_canje.Locked = False

If PUB_TIPMOV = 10 And LK_EMP = "DPV" And grid_canje.COL = 4 Then
    textovar_canje.Locked = True
End If

    textovar_canje.Left = grid_canje.Left + grid_canje.CellLeft
    textovar_canje.Width = grid_canje.CellWidth
    textovar_canje.Height = grid_canje.CellHeight
    textovar_canje.Top = grid_canje.Top + grid_canje.CellTop
    textovar_canje.Text = grid_canje.Text
    If textovar_canje.Visible = False Then textovar_canje.Visible = True
    textovar_bak = textovar_canje.Text
    'textovar_canje.SetFocus
     ' CAMTEXT
    On Error GoTo SALE
    Azul3 textovar_canje, textovar_canje
    i_bancos.Visible = False
    Exit Sub
SALE:
textovar_canje.SetFocus
End Sub
Public Sub PROCESA_CELDASL()
'If gridl.Col < 3 Then Exit Sub
'If gridl.Col = 5 Then Exit Sub
If gridl.CellWidth < 0 Then Exit Sub
textovarl.Locked = False


    textovarl.Left = gridl.Left + gridl.CellLeft
    textovarl.Width = gridl.CellWidth
    textovarl.Height = gridl.CellHeight
    textovarl.Top = gridl.Top + gridl.CellTop
    textovarl.Text = gridl.Text
    If textovarl.Visible = False Then textovarl.Visible = True
    textovar_bak = textovarl.Text
    textovarl.SetFocus
    i_bancos.Visible = False

End Sub


Private Sub grabar_Click()
Dim ctaprint As Integer
Dim P As Printer
Dim ws_fecha_vcto_orig As Date
Dim WS_NRO_ITEMS As Integer
Dim kcuentaf As Integer
Dim descrip As String
Dim nrolinelotes As Integer
Dim asigna_flag_lote As String * 1
Dim WS_TRANSAC  As Integer
Dim WK_MODI_1401 As String * 1
Dim wtventa As Currency
Dim loc_flag_1111  As String * 1
Dim LOC_FLAG_ALLOG As String * 1
'On Error GoTo Error_fatal
Dim WW_FECHA_COBRO As Date
Dim dispositivo As Integer
Dim WW_NUM_OPER As Currency
Dim ws_fecha_ant As Date
Dim ww_numdoc As Double
Dim ww_numser As Integer
Dim wformula, wformula1, wformula2
Dim ext_conta2 As Integer
Dim DIA, mes, ano
Dim er As rdoError
Dim pub_mensaje As String
Dim ws_cospro, ws_compra, ws_valor_ant As Currency
Dim flag_repite As Integer

Dim wS_saldo_car As Currency
Dim WS_NUM_OPER2 As Currency
Dim WW_NUM_OPER2 As Currency
Dim ws_codcia As String * 2

Dim WS_IMPORTE_AMORT_DOLL As Currency
Dim SS_IMPORTE_AMORT As Currency
Dim SS_IMPORTE_AMORT_DOLL As Currency
Dim WS_TOT As Currency
Dim FILAXX As Integer
Dim WZONA As String
Dim WSUBZONA As String
Dim WSLUGAR As String
Dim WS_AUTOCON As String
Dim WS_MONEDA As String * 1
Dim ws_mon_canje As String * 1
Const ingre = 2
Const MODIF = 1
Dim I As Integer
Dim N As Integer
Dim LOC_SALDO_CAR As Currency
Dim WS_DESCTO1, WS_DESCTO2 As Currency
Dim WS_CANTIDAD As Currency
Dim count As Integer
Dim ww_concepto As String * 150
Dim zona_nombre As String
Dim subzona_nombre  As String
Dim WS_IMPRESION_LET As Long
Dim WS_CHENUM As Currency
Dim WS_NUM_OPER  As Integer
Dim WS_DATOS As String * 2
Dim ws_flag As Integer
Dim WS_TRANSITO As String * 1
Dim WS_FLAG2 As Integer
Dim subtotal As Currency
Dim WS_BRUTO2 As Currency
Dim ws_diferencia As Currency
Dim WS_CORRELATIVO As Double
Dim ws_codart As Long
Dim NETO As Currency
Dim fx As Integer
Dim FLAG As Boolean
Dim ws_flag_arm As Integer
Dim ws_flag_che As Integer
Dim ws_flag_far As Integer
Dim ws_flag_car As Integer
Dim ws_flag_par As Integer
Dim ws_flag_ccm As Integer
Dim ws_flag_ven As Integer
Dim filafac As Integer

WS_FLAG_FACART = ""
Dim RUTA
Dim plaza_tasa1
Dim plaza_tasa2
WK_MODI_1401 = ""
LOC_FLAG_ALLOG = ""
ext_conta2 = 1
inicio_ext:   'inicio loop de extornos


If consis() = False Then Exit Sub

ww_concepto = ""

ww_codtra_ext = 0
filafac = 1
WS_NUMSEC = 0



'If pub_signo_ccm <> 0 And pub_signo_car <> 0 Then
'   If ws_moneda = "C" Then
      


'If ws_moneda = "A" Then
'   If pub_signo_car <> 0 Then
'      MsgBox "Operacion debe ser en efectivo ..."
'      exit sub
'   End If
'End If
'If Trim(ws_moneda) <> "" Then
'   If Mid(Tipo_Cambio.Caption, 2, 1) = "D" Then
'      MsgBox "Falta Cambiar a Soles ..."
'      exit sub
'   End If
'End If
'PASO:
WS_TRANSAC = 0
If LK_CODTRA = 2401 And PUB_PROCESO = 0 Then
    PUB_CODUSU = LK_CODUSU
    If LK_CODUSU = "OPER20" Or LK_CODUSU = "BOT-OPER20" Or LK_CODUSU = "SUPERVISOR" Or LK_CODUSU = "ADMIN" Then
        pub_mensaje = "- Grabar e Imprimir (Si)" & Chr(13) & "- Solo grabar (No).." & Chr(13) & "- No Grabar (Cancelar).."
        Pub_Respuesta = MsgBox(pub_mensaje, vbYesNoCancel + vbInformation + vbDefaultButton2, Pub_Titulo)
        WS_TRANSAC = 0
        If Pub_Respuesta = vbNo Then
            WS_TRANSAC = 0
        ElseIf Pub_Respuesta = vbYes Then
            WS_TRANSAC = 9
        Else
            grid_fac.COL = 1
            grid_fac.Row = grid_fac.Rows - 1
            grid_fac.SetFocus
            If TEXTOVAR.Visible Then TEXTOVAR.SetFocus
            Exit Sub
        End If
     Else ' otros usuarios
        pub_mensaje = "Desea Grabar e Imprimir...?"
        Pub_Respuesta = MsgBox(pub_mensaje, vbYesNo + vbInformation + vbDefaultButton2, Pub_Titulo)
        WS_TRANSAC = 0
        If Pub_Respuesta = vbYes Then
            WS_TRANSAC = 9
        ElseIf Pub_Respuesta = vbNo Then
            grid_fac.COL = 1
            grid_fac.Row = grid_fac.Rows - 1
            grid_fac.SetFocus
            If TEXTOVAR.Visible Then TEXTOVAR.SetFocus
            Exit Sub
        End If
     End If
     
 If WS_TRANSAC = 9 Then
    ctaprint = 0
    For Each P In Printers
        ctaprint = ctaprint + 1
    Next P
    If Val(par_llave!PAR_DEVICE_FBG) > ctaprint Then
       MsgBox "Problema Interno de Windows. (Tema: Impresion directa)" & Chr(13) & "Por favor salga del programa y cierre la sesion de su usuario/Reiniciar el Equipo e Ingrese Nuevamente al Sistema", 48, Pub_Titulo
       Exit Sub
    End If
 End If
End If

exito = True


grabar.Enabled = False
Screen.MousePointer = 11

pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)

If LK_CODTRA = 2401 Then ww_respuesta = vbNo
' Pendiente para Chequear el Stock por Pedido
loc_stock = ""
If consis2() = False Then
   exito = False
   GoTo SALIDA_TOTAL
End If


'repite: ' quite

Barra.Visible = True
'Barra
Barra.max = 90
Barra.Min = 62

Barra.Value = 62



N = 62
Do Until Val(tra_llave(N)) = 0 Or N = 76 Or exito = False


NUMERO = tra_llave(N)
On NUMERO GoSub CON1, CON2, CON3, CON4, CON5, CON6, CON7, CON8, CON9, CON10, CON11, CON12, CON13, CON14, CON15, CON16, CON17, CON18, CON19, CON20, CON21, CON22, CON23
Barra.Value = N

N = N + 1
Loop
SALIDA_TOTAL:
If exito = False Then
   con_llave.Close
   Barra.Visible = False
   Screen.MousePointer = 0
   If Left(pub_mensaje_err, 18) = "Ojo Cantidad Mayor" And LK_USU_STOCK = "X" Then
      pub_mensaje = " ¿Desea Grabar de todas maneras ... ?"
      ww_respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
      If ww_respuesta = vbYes Then GoTo inicio_ext
   End If
   MsgBox pub_mensaje_err, 48, Pub_Titulo
   PUB_TIPDOC = Nulo_Valors(SUT_LLAVE!sut_TIPDOC)
   grabar.Enabled = True
   grabar.SetFocus
   GoTo errorr
End If
' ALERTA DE MENSAJES POR ERROR DE SQL.
wtventa = Val(i_neto.Text)
CN.Execute "Begin Transaction", rdExecDirect
'On Error GoTo ALERTA
repite:
manda_numero
LOC_FLAG_ALLOG = ""
loc_flag_1111 = ""
WS_NUM_OPER2 = PUB_NUM_OPER_XXX


N = 77
Do Until Val(tra_llave(N)) = 0 Or N = 91
NUMERO = tra_llave(N)
 On NUMERO GoSub ACT1, ACT2, ACT3, ACT4, ACT5, ACT6, ACT7, ACT8, ACT9, ACT10, ACT11, ACT12, ACT13, ACT14, ACT15, ACT16, ACT17, ACT18, ACT19, ACT20, ACT21
Barra.Value = N
N = N + 1
Loop
' Actuliza las NC por docuemto
If LK_CODTRA = 2774 Then Aplica_NC
Barra.Value = 90
If Trim(LOC_FLAG_ALLOG) = "" Or (LK_CODTRA = 1111 And Trim(loc_flag_1111) = "") Then
 CN.Execute "Rollback Transaction", rdExecDirect
 con_llave.Close
 Barra.Visible = False
 MsgBox "Operación incompleta, NO SE REGISTRO!, Favor de Reiniciar el Sistema..." & Chr(13) & "El Sisema se Cerrará." & Chr(13) & "Si el problema continua llamar a su proveedor ", vbCritical, Pub_Titulo
 End
 GoTo fin
End If
If LK_CODTRA = 1401 And i_cambio.Value = 1 And Trim(WK_MODI_1401) = "" Then
 CN.Execute "Rollback Transaction", rdExecDirect
 con_llave.Close
 Barra.Visible = False
 MsgBox "Operación incompleta, NO SE REGISTRO!, Favor de Reiniciar el Sistema..." & Chr(13) & "El Sisema se Cerrará." & Chr(13) & "Si el problema continua llamar a su proveedor ", vbCritical, Pub_Titulo
 End
 GoTo fin
End If
WK_MODI_1401 = ""
LOC_FLAG_ALLOG = ""
loc_flag_1111 = ""

' PARA PIDA PROCESO DE COSTEO
If (LK_CODTRA = 2407 And LK_EMP_PTO = "A") Or PUB_TIPMOV = 20 Then
  'SQ_OPER = 1
  'PUB_CODCIA = LK_CODCIA
  'LEER_PAR_LLAVE
  'par_llave.Edit
  'par_llave!par_flag_costos = "8"
  'par_llave.Update
End If

CN.Execute "Commit Transaction", rdExecDirect
con_llave.Close
i_def.Enabled = True
On Error GoTo 0
LOC_FLAG_ALLOG = ""
det_lot.Rows = 1
det_lot.Clear

' <<<< PROCEDIMIENTO PARA LOS REPORTES >>>>
'If LK_CODTRA = 2401 And Val(i_numguia.Text) > 0 Then
'   MsgBox " Imprimir en Consulta de Documentos "
'Else


If PUB_PROCESO = 1 And chkPrint.Value = 1 Then
  dispositivo = REP_TRANSAC
Else
 If WS_TRANSAC = 9 Then
    dispositivo = REP_TRANSAC
  End If
End If
'End If
dispositivo = 1
'End If

If LK_CODTRA = 1401 Then
 llama_anexo
End If

final1

N = 103
Do Until Val(tra_llave(N)) = 0 Or N = 107
NUMERO = tra_llave(103)
 On NUMERO GoSub REP1, REP2, REP3

N = N + 1
Loop
i_cambio.Value = 0
i_importe.Text = ""
i_importe_amort.Text = ""
cancela_todo
Grid_all.Visible = False
pasa_cabeza

WS_BRUTO = 0
SUB_CANT = 0
SUB_FLETE = 0
SUB_JABAS = 0
SUB_UNIDAD = 0
'WS_MONEDA_CLI = ""
WS_MONEDA_CCM = ""

'If LK_EMP = "3AA" Then
 If LK_CODTRA <> 2401 Then
    grid_fac.Rows = 3
 Else
    VER_LINEAS
    calcula_totales
 End If
'Else
' grid_fac.Rows = 3
'End If


FILAX = 0
Barra.Visible = False
grabar.Enabled = True
label_nomart.Caption = ""
label_precio.Caption = ""
i_limcre_ant.Text = ""
i_dias.Text = ""
i_impto2.Text = ""
If LK_CODTRA = 2412 Then
  i_concepto.Text = ""
  i_TEXTONCRE.Text = ""
End If
If LK_CODTRA = 2401 Or LK_CODTRA = 1401 Or LK_CODTRA = 2402 Then
 i_cambio.Visible = False
 i_numfac.Locked = True
 i_numser.Locked = True
 i_numfac.BackColor = QBColor(7)
 i_numser.BackColor = QBColor(7)
 If LK_CODTRA = 2401 Or LK_CODTRA = 2402 Then
 i_fecha_compra.Visible = False
 End If
 i_numfac_c.Text = ""
 i_numser_c.Text = ""
 i_numfac_c.Text = ""
 If cmdtipo.ListCount > 0 Then cmdtipo.ListIndex = 0
End If


' AQUI INICIALIZO VARIABLES QUE NO SE CEREAN  AL ENTRAR AL COMMAND1.CLICK

fila = 0
grabar.Enabled = True
If gridl.Visible Then pasa_cabeza_L

If Nulo_Valors(par_llave!PAR_FLAG_RECIBOS) = "A" And pub_signo_car = -1 Then
   BolFac.Caption = "Nº " & PUB_CODART
End If



'If PUB_TIPMOV <> 0 Then
'   pasa_cabeza
'   If dispositivo = 1 Then
'     i_def.SetFocus
'   End If
'   exit sub
'End If
Screen.MousePointer = 0

If dispositivo = 1 Or dispositivo = -1 Then
 If LK_EMP = "HER" And LK_CODTRA = 2401 Then
   '  i_def.ListIndex = 1
   '  i_DEF_KeyPress 13
     GoTo CAJA
     Exit Sub
 Else
     FORMGEN.i_def.SetFocus
     SendKeys "%{DOWN}"
 End If
End If
If FORMGEN.i_importe.Visible = True And (PUB_IMPORTE = 0 Or PUB_CHESEC <> 0) And LK_CODTRA = 5714 Then
     FORMGEN.i_importe.SetFocus
End If



If LK_CODTRA = 1111 Then
   ext_conta2 = ext_conta2 + 1
   If ext_conta2 > ext_conta Then
   Else
     i_num_oper.Text = Val(i_num_oper.Text) + 1
     GoTo inicio_ext
   End If
End If

pub_signo_arm = SUT_LLAVE!SUT_SIGNO_ARM 'ASEGURAMOS POR LA TRANS. 1409
pub_signo_car = SUT_LLAVE!SUT_SIGNO_CAR 'ASEGURAMOS TODO
PUB_TIPMOV = SUT_LLAVE!SUT_TIPMOV 'ASEGURAMOS TODO POR  LA 1409
CAJA:

If LK_CODTRA = 2401 And LK_CAJERO = "A" And PUB_PROCESO <> 1 Then
 PUB_ACU_TOTAL = PUB_ACU_TOTAL + Val(wtventa)
 loc_doucumento = loc_doucumento + " / " + Trim(i_numfac.Text)
 If Trim(grid_fac.TextMatrix(2, 0)) = "" Then
 End If
 grid_fac.Row = 2
 grid_fac.COL = 1
 grid_fac.SetFocus
End If
If Trim(grid_fac.TextMatrix(2, 0)) = "" Then
 PUB_ACU_TOTAL = 0
 loc_doucumento = ""
End If
If LK_CODTRA = 2580 Then
  i_fbg.ListIndex = -1
  i_dias.Text = ""
End If
GoTo fin

CON1:
  ws_codart = 0
  exito = ANEXO_CON1()
Return

CON2:

'If PUB_CODART = 0 Then
'   pub_mensaje_err = "Articulo Invalido ..."
'   FORMGEN.i_codart.SetFocus
'   exito = False
'End If

'PSART_LLAVE.rdoParameters(0) = PUB_CODART
'art_LLAVE.Requery
'If art_LLAVE.EOF = True Then
'   pub_mensaje_err = "ARTICULO NO EXISTE...."
'   FORMGEN.i_codart.SetFocus
'   exito = False
'End If

Return


CON3:
exito = ANEXO_CON3()
Return

CON4:
Return

CON5:
Return

CON6:
exito = ANEXO_CON6()
Return

CON7:
exito = Anexo_con7()
Return
   
CON8:
Return

CON9:
Return

CON10:
Return

CON11:

Return

CON12:
exito = Anexo_con12()
Return

CON13:
exito = ANEXO_CON13()
Return

CON14:
exito = ANEXO_CON14()


Return

CON15:
 exito = ANEXO_CON15()
Return

CON16:
SQ_OPER = 1
LEER_TAB_LLAVE

Return

CON17:
exito = ANEXO_CON17()
Return

CON18:
 exito = ANEXO_CON18()
Return

CON19:
 exito = ANEXO_CON19()
Return

CON20:
exito = ANEXO_CON20()
Return


CON21:
exito = ANEXO_CON21()

If pub_signo_car <> 0 Then
SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_codclie = PUB_CODCLIE
pu_cp = PUB_CP
LEER_CAR_LLAVE
If car_llave.EOF = True Then
   pub_mensaje_err = "DOCUMENTO NO EXISTE..."
   exito = False
End If
SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_codclie = PUB_CODCLIE
pu_cp = PUB_CP
LEER_CAA_LLAVE
If caa_histo.EOF = True Then
   pub_mensaje_err = "DOCUMENTO EN CARACU NO EXISTE..."
   exito = False
End If
End If



If PUB_TIPMOV = 0 And Trim(WS_FLAG_FACART) = "" Then
   Return
End If
'If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then Return
SQ_OPER = 5
LEER_FAR_LLAVE
' en caso de notas de credito, puede ser sin facart
'If far_menor3.EOF = True Then
'   MsgBox "!!! NO HAY ESE DOCUMENTO ...", 48, Pub_Titulo
'   exito = False
'   Return
'End If
If far_menor3.EOF Then
   MsgBox "!!! NO HAY ESE DOCUMENTO ...", 48, Pub_Titulo
   exito = False
   Return
End If
If ww_codtra_ext = 2103 Then
   lee_ped_llave
   If Nulo_Valor0(ped_llave!PED_ingresos) <> 0 Or Nulo_Valor0(ped_llave!PED_salidas) <> 0 Then
      MsgBox "!!! NO PROCEDE ANULACION YA EXISTEN MOVIMIENTOS... ...", 48, Pub_Titulo
      exito = False
      Return
   End If
End If


far_menor3.MoveLast


If far_menor3!far_estado = "E" Then
   pub_mensaje_err = "Ya Extornado  ..."
   exito = False
   Return
End If

If far_menor3!far_codart = 0 Then
       Return
End If

far_menor3.MoveFirst
fila = 2
Do Until far_menor3.EOF

   grid_fac.Rows = fila + 2
   WS_NUMSEC = far_menor3!FAR_NUMSEC
   
   If LK_CODTRA = 1111 And far_menor3!FAR_TIPMOV = 180 Then WS_NUMSEC = WS_NUMSEC + 1
   
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_KEY = far_menor3!far_codart
   LEER_ART_LLAVE
   'SOLO PARA LA 2499
   If PUB_TIPMOV <> 80 Then grid_fac.TextMatrix(fila, 0) = art_LLAVE!art_nombre '& art_LLAVE!art_alterno
   
   grid_fac.TextMatrix(fila, 16) = far_menor3!far_codart
   grid_fac.TextMatrix(fila, 33) = far_menor3!far_codart
   grid_fac.TextMatrix(fila, 2) = Nulo_Valor0(far_menor3!far_JABAS)
   grid_fac.TextMatrix(fila, 3) = Nulo_Valor0(far_menor3!far_UNIDADES)
   grid_fac.TextMatrix(fila, 11) = far_menor3!FAR_COSPRO
   If far_menor3!FAR_equiv > 0 Then
      grid_fac.TextMatrix(fila, 4) = far_menor3!far_cantidad / far_menor3!FAR_equiv
   Else
      grid_fac.TextMatrix(fila, 4) = far_menor3!far_cantidad
   End If
   grid_fac.TextMatrix(fila, 6) = far_menor3!FAR_PRECIO
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(far_menor3!FAR_equiv)
   grid_fac.TextMatrix(fila, 5) = Nulo_Valors(far_menor3!far_descri)
   grid_fac.TextMatrix(fila, 8) = Nulo_Valor0(far_menor3!FAR_FLETE)
   
   If LK_EMP = "HER" Or LK_EMP = "3AA" Then
   ' ICA
     grid_fac.TextMatrix(fila, 10) = far_menor3!far_PORDESCTOS ' far_menor3!FAR_PORDESCTO1
     grid_fac.TextMatrix(fila, 32) = far_menor3!far_precio_neto ' far_menor3!FAR_PORDESCTO1
     grid_fac.TextMatrix(fila, 34) = Format(far_menor3!far_PESO * far_menor3!far_cantidad, "0.00")
     grid_fac.TextMatrix(fila, 37) = far_menor3!far_PESO
   Else
     grid_fac.TextMatrix(fila, 10) = far_menor3!FAR_DESCTO
   End If
   grid_fac.TextMatrix(fila, 17) = Nulo_Valor0(far_menor3!far_num_precio)
   grid_fac.TextMatrix(fila, 11) = far_menor3!FAR_COSPRO
   grid_fac.TextMatrix(fila, 14) = far_menor3!FAR_equiv
   ' grid_fac.TextMatrix(fila, 18) = Nulo_Valor0(far_menor3!far_pedsec)
   grid_fac.TextMatrix(fila, 19) = Nulo_Valor0(far_menor3!FAR_ORDEN_UNIDADES)
   
   grid_fac.TextMatrix(fila, 39) = far_menor3!FAR_NUMSER_C
   grid_fac.TextMatrix(fila, 40) = far_menor3!FAR_NUMFAC_C
   grid_fac.TextMatrix(fila, 46) = far_menor3!FAR_ESTADO2
   grid_fac.TextMatrix(fila, 47) = far_menor3!FAR_CODLOT ' codigo de lote
   grid_fac.TextMatrix(fila, 48) = far_menor3!far_cantidad_p ' codigo de lote
   
   If far_menor3!FAR_TIPMOV = 180 And far_menor3!far_transito = "T" Then
     grid_fac.TextMatrix(fila, 12) = far_menor3!far_signo_arm
   Else
     grid_fac.TextMatrix(fila, 12) = far_menor3!far_signo_arm * -1
   End If
   WS_MONEDA_CLI = Nulo_Valors(far_menor3!FAR_MONEDA)
   If WS_MONEDA_CLI = "S" Then
      i_ds.ListIndex = 0
   ElseIf WS_MONEDA_CLI = "D" Then
      i_ds.ListIndex = 1
   Else
      i_ds.ListIndex = -1
   End If
   
   i_moneda.Text = Nulo_Valors(far_menor3!FAR_MONEDA)
   lmoneda.Caption = Nulo_Valors(far_menor3!FAR_MONEDA)
   PUB_PEDSER = Nulo_Valor0(far_menor3!FAR_PEDSER)
   PUB_PEDFAC = Nulo_Valor0(far_menor3!FAR_PEDFAC)
   
   far_menor3.MoveNext
   fila = fila + 1
   
Loop
WS_NUMSEC = WS_NUMSEC + 200
calcula_totales

PUB_FECHA = LK_FECHA_DIA

Return

CON22:
exito = ANEXO_CON22()
Return

CON23:
exito = ANEXO_CON23()
Return


GoTo fin

ACT1:

GoSub ACT2

If pub_flag_cambio = 0 And PUB_TIPMOV = 20 And i_cambio.Value = 1 Then Return
If PUB_TIPMOV = 20 And Check1.Value = 1 And Val(i_neto.Text) = 0 Then Return
If LK_CODTRA = 1111 And ww_codtra_ext = 2107 Then Return
If LK_CODTRA = 1111 And ww_codtra_ext = 1402 Then Return

flag_salto = 1

If (PUB_TIPMOV = 0 Or pub_signo_arm = 0) And SUT_LLAVE!SUT_FLAG_TIPO <> "P" Then
   Return
End If
    
If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
   If Val(grid_fac.TextMatrix(1, 7)) = 0 Then Return
End If

If Val(i_neto.Text) = 0 And Trim(grid_fac.TextMatrix(2, 4)) = "ANULACION" Then
   GoTo ACT13
End If

ww_respuesta = vbYes  'PARTE DE 1409

SIGUE_1409:

If LK_CODTRA = 1122 Or LK_CODTRA = 1111 Then
Else
If i_cambio.Visible = False Or i_cambio.Value = 0 Then llena_numfac
End If
If LK_CODTRA <> 1111 Then
  TODO_LOTES  ' TOMA LOS LOTES AUTOMATICAMENTE DE ACUERDO A FECHA VCTO
  If flag_lotes_chequea = "A" Then
    CN.Execute "Rollback Transaction", rdExecDirect
    con_llave.Close
    MsgBox "Cantidad No Corresponde -  Existe Codigo con Lotes en Negativos , pro favor Revisar o Avisar al Administrador.!!!" & Chr(13) & "Codigo del Producto : " & det_lot.TextMatrix(kcuentaf, 1), 48, Pub_Titulo
    GoTo fin
  End If
  
  For kcuentaf = 0 To det_lot.Rows - 1
   If Val(det_lot.TextMatrix(kcuentaf, 3)) < 0 Then
        CN.Execute "Rollback Transaction", rdExecDirect
        con_llave.Close
        MsgBox "Existe Codigo con Lotes en Negativos , pro favor Revisar o Avisar al Administrador.!!!" & Chr(13) & "Codigo del Producto : " & det_lot.TextMatrix(kcuentaf, 1) & " Cantidad :" & det_lot.TextMatrix(kcuentaf, 3), 48, Pub_Titulo
        GoTo fin
   End If
  Next kcuentaf
  
  
   ' cheque las filas en la 2401 si pasa del parametro establesido
   If LK_CODTRA = 2401 Then
     nrolinelotes = det_lot.Rows - 1
     descrip = ARTICULOLOTE
     nrolinelotes = Val(Left(descrip, 3))
     descrip = Mid(descrip, 4, Len(descrip))
     If Trim(FORMGEN.i_fbg.Text) = "B" Then
      If PUB_CODVEN = 10 Then
          WS_NRO_ITEMS = 13 ' TEMPORAL
      Else
          WS_NRO_ITEMS = Nulo_Valor0(par_llave!par_BOL_lines)
      End If
      If nrolinelotes <= Val(WS_NRO_ITEMS) Then
      Else
         CN.Execute "Rollback Transaction", rdExecDirect
         con_llave.Close
         
         MsgBox "El Tamaño del Documento no alcanza a los Item de los Articulos, pasan de lo predeterminado." & Chr(13) & "Total de Filas para Imprimir : " & Format(nrolinelotes, "00") & Chr(13) & "Tope de filas del Documento : " & Format(WS_NRO_ITEMS, "00") & Chr(13) & descrip, 48, Pub_Titulo
         GoTo fin
      End If
     ElseIf Trim(FORMGEN.i_fbg.Text) = "F" Then
      If nrolinelotes <= Val(par_llave!par_fac_lines) Then
      Else
         CN.Execute "Rollback Transaction", rdExecDirect
         con_llave.Close
         MsgBox "El Tamaño del Documento no alcanza a los Item de los Articulos, pasan de lo predeterminado." & Chr(13) & "Total de Filas para Imprimir : " & Format(nrolinelotes, "00") & Chr(13) & "Tope de filas del Documento : " & Format(par_llave!par_fac_lines, "00") & Chr(13) & descrip, 48, Pub_Titulo
         GoTo fin
      End If
     ElseIf Trim(FORMGEN.i_fbg.Text) = "G" Then
      If nrolinelotes <= Val(par_llave!par_GUIA_lines) Then
      Else
         CN.Execute "Rollback Transaction", rdExecDirect
         con_llave.Close
         MsgBox "El Tamaño del Documento no alcanza a los Item de los Articulos, pasan de lo predeterminado." & Chr(13) & "Total de Filas para Imprimir : " & Format(nrolinelotes, "00") & Chr(13) & "Tope de filas del Documento : " & Format(par_llave!par_fac_lines, "00") & Chr(13) & descrip, 48, Pub_Titulo
         GoTo fin
      End If
     End If
   End If
End If
SUB_CANT = 0
fila = 2
WS_BRUTO2 = 0
FLAG = False
Do While FLAG = False
   If Val(grid_fac.TextMatrix(fila, 16)) = 0 Or (Val(grid_fac.TextMatrix(fila, 3)) = 0 And Val(grid_fac.TextMatrix(fila, 4)) = 0) Then
      GoTo OTROMAS
   End If
   If Val(grid_fac.TextMatrix(fila, 38)) = 1 Then GoTo OTROMAS
   PUB_CODART = Val(grid_fac.TextMatrix(fila, 16))
   
   PUB_JABAS = Val(FORMGEN.grid_fac.TextMatrix(fila, 2))
   PUB_UNIDAD = Val(FORMGEN.grid_fac.TextMatrix(fila, 3))
   
   PUB_PRECIO2 = Val(FORMGEN.grid_fac.TextMatrix(fila, 6))
   
   PUB_COSPRO = Val(FORMGEN.grid_fac.TextMatrix(fila, 11))
   If LK_CODTRA = 1405 Then
    PUB_NUMSER_C = Val(FORMGEN.grid_fac.TextMatrix(fila, 39))
    PUB_NUMFAC_C = Val(FORMGEN.grid_fac.TextMatrix(fila, 40))
    PUB_NUMDOC = Val(FORMGEN.grid_fac.TextMatrix(fila, 41))
   End If
   
   
   
sigue:
   
   
   If pub_signo_arm = 0 And SUT_LLAVE!SUT_FLAG_TIPO <> "P" Then
      CN.Execute "Rollback Transaction", rdExecDirect
      MsgBox "pub_signo_arm = 0 ... en ACT1, Verificar ..."
      GoTo fin
      
      GoTo arm_no
   End If
   
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
     CN.Execute "Rollback Transaction", rdExecDirect
     MsgBox "error grave en ARTICULO ..."
     GoTo fin
   End If
   WS_CANTIDAD = Val(FORMGEN.grid_fac.TextMatrix(fila, 4)) * Val(FORMGEN.grid_fac.TextMatrix(fila, 14))
   If SUT_LLAVE!SUT_FLAG_TIPO = "P" Then GoTo SALTA_T
   
   If PUB_PROCESO = 1 Then GoTo SALTA_T
   ' If LK_CODTRA = 1111 And ww_codtra_ext = 2401 Then GoTo SALTA_T solo para pedidos y despues canje por ventas
   If LK_CODTRA = 1111 And Trim(grid_fac.TextMatrix(fila, 46)) = "L" Then GoTo SALTA_T
   
   If PUB_TIPMOV = 179 And LK_CODTRA = 2411 Then
            If LK_ACTIVA = "A" Then
               If PUB_SO = "A" Then
                  WS_CANTIDAD = (Val(arm_llave!ARM_saldo_s) - (Val(FORMGEN.grid_fac.TextMatrix(fila, 4)) * Val(FORMGEN.grid_fac.TextMatrix(fila, 14)))) * -1
                 ' PUB_JABAS = (Val(arm_llave!ARM_SALDO_S2) - (Val(FORMGEN.grid_fac.TextMatrix(fila, 3)))) * -1
               Else
                  WS_CANTIDAD = (Val(arm_llave!ARM_Saldo_n) - (Val(FORMGEN.grid_fac.TextMatrix(fila, 4)) * Val(FORMGEN.grid_fac.TextMatrix(fila, 14)))) * -1
                  PUB_JABAS = (Val(arm_llave!ARM_SALDO_N2) - (Val(FORMGEN.grid_fac.TextMatrix(fila, 3)))) * -1
               End If
            Else
                WS_CANTIDAD = (Val(arm_llave!ARM_STOCK) - (Val(FORMGEN.grid_fac.TextMatrix(fila, 4)) * Val(FORMGEN.grid_fac.TextMatrix(fila, 14)))) * -1
                PUB_JABAS = (Val(arm_llave!arm_stock2) - (Val(FORMGEN.grid_fac.TextMatrix(fila, 3)))) * -1
            End If
            If WS_CANTIDAD < 0 Then
              pub_signo_arm = -1
            Else
              pub_signo_arm = 1
            End If
            WS_CANTIDAD = Abs(WS_CANTIDAD)
   End If
   
If LK_CODTRA <> 1122 And LK_CODTRA <> 1111 Then
    If Nulo_Valors(SUT_LLAVE!SUT_UNIDADES) <> "A" And Val(FORMGEN.grid_fac.TextMatrix(fila, 14)) <> 1 Then
        CN.Execute "Rollback Transaction", rdExecDirect
        MsgBox "error grave en ARTICULO ..."
        GoTo fin
    End If
End If
                                                
     ' CAMTEXT
   If PUB_TIPMOV = 70 Or PUB_TIPMOV = 75 Or PUB_TIPMOV = 180 Or Trim(Nulo_Valors(SUT_LLAVE!sut_is)) = "A" Or PUB_TIPMOV = 93 Or LK_CODTRA = 2425 Or LK_CODTRA = 2415 Or Check1.Value = 1 Then    'ok 2425 unico caso
      pub_signo_arm = Val(FORMGEN.grid_fac.TextMatrix(fila, 12))
   End If
   If LK_CODTRA = 1111 Then
     If all_llave!ALL_tipmov = 102 Then  '  PUB_TIPMOV = 75 Or PUB_TIPMOV = 180 Or Trim(Nulo_Valors(SUT_LLAVE!sut_is)) = "A" Or PUB_TIPMOV = 93 Or LK_CODTRA = 2425 Or LK_CODTRA = 2415 Or Check1.Value = 1 Then    'ok 2425 unico caso
       pub_signo_arm = Val(FORMGEN.grid_fac.TextMatrix(fila, 12))
     End If
   End If
   
   If pub_signo_arm = -1 Or pub_signo_arm = 1 Then
   Else
      CN.Execute "Rollback Transaction", rdExecDirect
      MsgBox "error grave en SIGNO ARM ..."
      GoTo fin
   End If

      CONTADOR = CONTADOR + 1
      
         arm_llave.Edit
         arm_llave!ARM_STOCK = Val(arm_llave!ARM_STOCK) + WS_CANTIDAD * pub_signo_arm
                                                            ' PUB_UNIDAD
         arm_llave!arm_stock2 = Val(arm_llave!arm_stock2) + PUB_JABAS * pub_signo_arm
         
         If PUB_SO = "A" Then
            arm_llave!ARM_saldo_s = Val(arm_llave!ARM_saldo_s) + WS_CANTIDAD * pub_signo_arm
            'arm_llave!ARM_SALDO_S2 = Val(arm_llave!ARM_SALDO_S2) + PUB_JABAS * pub_signo_arm
         Else
            arm_llave!ARM_Saldo_n = Val(arm_llave!ARM_Saldo_n) + WS_CANTIDAD * pub_signo_arm
            arm_llave!ARM_SALDO_N2 = Val(arm_llave!ARM_SALDO_N2) + PUB_JABAS * pub_signo_arm
         End If
         If pub_signo_arm = 1 And PUB_TIPMOV <> 0 Then Act_sin_stock


 
         If pub_signo_arm = -1 Then
            arm_llave!ARM_SALIDAS = Val(arm_llave!ARM_SALIDAS) + Val(WS_CANTIDAD)
         Else
         If pub_signo_arm = 1 Then
            arm_llave!ARM_INGRESOS = Val(arm_llave!ARM_INGRESOS) + Val(WS_CANTIDAD)
         End If
         End If
         If LK_CODTRA = 1401 Then
           If i_fecha_compra.Visible Then
              arm_llave!ARM_FECHA_ULT = i_fecha_compra.Text
           Else
             arm_llave!ARM_FECHA_ULT = LK_FECHA_DIA
           End If
         End If
         arm_llave.Update
SALTA_T:
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      PUB_KEY = arm_llave!ARM_codart
      LEER_ART_LLAVE
      If art_LLAVE.EOF Then
        CN.Execute "Rollback Transaction", rdExecDirect
        MsgBox "error grave en ARTICULO ..."
        GoTo fin
      End If
      
      If LK_CODTRA = 1111 Then
       If EXTORNA_LOTE = False Then GoTo fin
      End If
      

      If Trim(art_LLAVE!art_nombre) = Trim(FORMGEN.grid_fac.TextMatrix(fila, 0)) Then
      Else
         CN.Execute "Rollback Transaction", rdExecDirect
         MsgBox "Error Grave con Codigos Alternos, Verificar o digitar el Producto Nuevamente..." & Chr(13) & "FILA : " & fila & Chr(13) & "Productos : " & Trim(art_LLAVE!art_nombre), 48, Pub_Titulo
         GoTo fin
      End If
      
      If LK_CODTRA = 2403 And PUB_SECUENCIA = 11 Then
         PUB_PRECIO2 = Format(PUB_PRECIO2 / Val(FORMGEN.grid_fac.TextMatrix(fila, 14)), "0.000")
         arm_llave.Edit
         arm_llave!arm_cospro = PUB_PRECIO2
         PUB_COSPRO = PUB_PRECIO2
         arm_llave.Update
      End If
         
         
      If PUB_TIPMOV = 20 Then
         art_LLAVE.Edit
         art_LLAVE!art_codclie = PUB_CODCLIE
         art_LLAVE.Update
      End If
      
      ws_flag_arm = MODIF
      
arm_no:
   ' PARA LOS LOTES NEMOS EN LA 1111
      If LK_CODTRA = 1111 Then
        ParaLot_count = 0
        GoTo SINLOTE
      End If
      verif_lotes Val(arm_llave!ARM_codart), fila
      asigna_flag_lote = ""
      If ParaLot_count = 1 Then asigna_flag_lote = "A"
Otro_lote:
      PUB_LOT_CANT = ParaLot_lotcant(ParaLot_count)
      PUB_CODLOT = ParaLot_codlot(ParaLot_count)
      If Trim(PUB_CODLOT) = "" Then
         CN.Execute "Rollback Transaction", rdExecDirect
         MsgBox "Verifique los Codigo de Lotes ..." & Chr(13) & "FILA : " & fila & Chr(13) & "Productos : " & Trim(art_LLAVE!art_nombre), 48, Pub_Titulo
         GoTo fin
       End If

      If Not IsDate(ParaLot_fechalot(ParaLot_count)) Then
        PUB_FECHA_LOT = LK_FECHA_DIA
      Else
        PUB_FECHA_LOT = ParaLot_fechalot(ParaLot_count)
      End If
      ParaLot_count = ParaLot_count - 1
SINLOTE:
      far_llave.AddNew
      
       ' lotes
    If LK_CODTRA <> 1111 Then
      PSLOT_LLAVE(0) = LK_CODCIA
      PSLOT_LLAVE(1) = arm_llave!ARM_codart
      PSLOT_LLAVE(2) = PUB_CODLOT
      lot_llave.Requery
      If lot_llave.EOF Then
       lot_llave.AddNew
       lot_llave!LOT_CODCIA = LK_CODCIA
       lot_llave!lot_nrolote = PUB_CODLOT
       lot_llave!lot_codart = arm_llave!ARM_codart
       lot_llave!lot_fecha_vcto = PUB_FECHA_LOT
       lot_llave!lot_saldos = 0
      Else
       lot_llave.Edit
      End If
      lot_llave!lot_codclie = 0
      lot_llave!lot_saldos = lot_llave!lot_saldos + (PUB_LOT_CANT * pub_signo_arm)
      lot_llave.Update
     End If
      far_llave!far_cantidad_p = PUB_LOT_CANT
      far_llave!FAR_CODLOT = PUB_CODLOT
      If asigna_flag_lote = "A" Then
        far_llave!FAR_ESTADO2 = "N"
      Else
       If ParaLot_count = 1 Then
          far_llave!FAR_ESTADO2 = "N"
       Else
          far_llave!FAR_ESTADO2 = "L"
       End If
      End If
      far_llave!FAR_TIPMOV = PUB_TIPMOV
      far_llave!FAR_CODCIA = LK_CODCIA
      If PUB_TIPMOV = 20 Then
        far_llave!far_cod_sunat = "01" ' SIEMPRE FACTURA  PARA COMPRAS DE MERCADERIA
      Else
        far_llave!far_cod_sunat = Val(Right(i_codsunat.Text, 5))
      End If
      far_llave!far_numser = PUB_NUMSER
      far_llave!FAR_CODVEN = PUB_CODVEN
      far_llave!far_numfac = PUB_NUMFAC
      far_llave!FAR_CLIENTE = Trim(t_nombre.Text)
      far_llave!FAR_DOCCLI = Trim(FORMGEN.t_doc.Text)
      If PUB_CODCLIE = 1 Then
          far_llave!FAR_dircli = Trim(FORMGEN.t_direc.Text)
      Else
          far_llave!FAR_dircli = Trim(FORMGEN.i_dircli.Text)
      End If
      WS_NUMSEC = WS_NUMSEC + 1

      far_llave!FAR_NUMSEC = WS_NUMSEC
      If pub_signo_arm = 0 Then
         far_llave!far_codart = PUB_CODART
         far_llave!FAR_STOCK = 0
      Else
         far_llave!FAR_STOCK = Val(arm_llave!ARM_STOCK)
         far_llave!far_codart = Val(arm_llave!ARM_codart)
      End If
      far_llave!far_cantidad = Val(WS_CANTIDAD)
'      far_llave!FAR_PESO = Val(FORMGEN.grid_fac.TextMatrix(fila, 4))
      
      SUB_CANT = SUB_CANT + WS_CANTIDAD
      
      far_llave!far_signo_car = pub_signo_car
      far_llave!far_signo_arm = pub_signo_arm
      
      far_llave!far_codclie = PUB_CODCLIE
      far_llave!FAR_RUC = PUB_RUC
      
      WS_MONEDA_CLI = i_ds.Text
      
      far_llave!FAR_MONEDA = WS_MONEDA_CLI
      far_llave!FAR_EX_IGV = Nulo_Valors(grid_fac.TextMatrix(fila, 23))
      
      far_llave!FAR_cp = PUB_CP
      If i_fecha_compra.Visible = True Then
       far_llave!FAR_fecha_compra = i_fecha_compra.Text
      Else
       far_llave!FAR_fecha_compra = LK_FECHA_DIA
      End If
      
      
      
      'AGREGE ACV
      If LK_CODTRA = 1111 Then
        far_llave!FAR_fecha_compra = i_fecha_compra.Text
      End If
      
      far_llave!far_estado = "N"
      
      
      If LK_CODTRA = 2401 And Trim(SUT_LLAVE!SUT_abreviado) = "CONS" Then
         far_llave!FAR_ESTADO2 = "C"
      End If
         

' QUIEN A PUESTO ESTO!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      If PUB_TIPMOV = 20 Then
         'If i_fbg.Text = "F" Then far_llave!FAR_estado2 = "P"
      End If
'????????????????????????????????????????????????
      If LK_CODTRA = 1111 Or LK_CODTRA = 1133 Then
         far_llave!far_estado = "E"
         far_llave!FAR_ESTADO2 = "E"
      End If
 

      far_llave!FAR_COSPRO = PUB_COSPRO
      'YO AGREGO
      If LK_CODTRA = 2407 Then
      far_llave!FAR_COSPRO_ANT = Val(arm_llave!arm_cospro)
      Else
      far_llave!FAR_COSPRO_ANT = PUB_COSPRO
      End If
      far_llave!FAR_PRECIO = PUB_PRECIO2
      
      far_llave!FAR_equiv = Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 14)))
      far_llave!FAR_FBG = PUB_FBG
      
      far_llave!far_IMPTO = PUB_IMPTO
      far_llave!FAR_TOT_FLETE = PUB_FLETE
      far_llave!FAR_FLETE = Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 8)))
      If PUB_DESCTO <> 0 Then
         far_llave!FAR_DESCTO = Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 10)))
      Else
        far_llave!FAR_DESCTO = 0
      End If
      ' VERIFICAR EL CAMPO NO EXISTE EN BDATOS
      far_llave!far_PORDESCTOS = grid_fac.TextMatrix(fila, 10)
      
      far_llave!FAR_TOT_DESCTO = PUB_DESCTO
      far_llave!FAR_GASTOS = PUB_GASTOS
      far_llave!FAR_BRUTO = PUB_SUBTOTAL
      far_llave!FAR_NUMDOC = PUB_NUMDOC
      
      far_llave!far_numguia = PUB_NUMGUIA ' Val(i_numguia.Text)
      far_llave!far_serguia = PUB_SERGUIA 'Val(i_serguia.Text)
      
      far_llave!FAR_pordescto1 = Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 49)))   ' Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 10)))
      far_llave!FAR_SUBTOTAL = Val(Nulo_Valor0(grid_fac.TextMatrix(fila, 7)))
      
      far_llave!FAR_costeo = " "
      far_llave!FAR_COSTEO_REAL = " "
      If (PUB_TIPMOV = 20 Or LK_CODTRA = 2414 Or LK_CODTRA = 2403 Or LK_CODTRA = 2409) And pub_signo_arm = 1 Then
         If LK_EMP = "HER" Then
            If Val(grid_fac.TextMatrix(fila, 37)) <> 1 Then
              far_llave!FAR_costeo = "A"
              far_llave!FAR_COSTEO_REAL = "A"
            End If
         Else
           far_llave!FAR_costeo = "A"
           far_llave!FAR_COSTEO_REAL = "A"
         End If
         'PENDIENTE
         If WS_MONEDA_CLI = "D" Then
            If Tipo_Cambio.Caption = "&Dollares" Then
               far_llave!FAR_tipo_cambio = 1
            Else
                                                            'ACV
               far_llave!FAR_tipo_cambio = LK_TIPO_CAMBIO ' WS_TIPO_CAMBIO
            End If
         Else
            If Tipo_Cambio.Caption = "&Soles" Then
               far_llave!FAR_tipo_cambio = 1
            Else
               far_llave!FAR_tipo_cambio = LK_TIPO_CAMBIO ' ACV  WS_TIPO_CAMBIO
            End If
         End If
      Else
            far_llave!FAR_tipo_cambio = 0
      End If
      If PUB_TIPMOV = 101 And (LK_CODCIA = "10" Or LK_CODCIA = "04" Or LK_CODCIA = "30" Or LK_CODCIA = "02") Then      ' SOLO PARA ALMACENES DEFECTUOSOS
             far_llave!FAR_costeo = "A"
             far_llave!FAR_COSTEO_REAL = "A"
      End If
'      If PUB_TIPMOV = 94 Then far_llave!FAR_COSTEO = "A"
      
'      If PUB_TIPMOV = 75 And pub_signo_arm = 1 Then far_llave!FAR_COSTEO = "A"
      far_llave!FAR_DIAS = pub_dias
      
      far_llave!FAR_fecha = LK_FECHA_DIA
      
      far_llave!FAR_NUMSER_C = PUB_NUMSER_C
      far_llave!FAR_NUMFAC_C = PUB_NUMFAC_C
      far_llave!FAR_NUMOPER = PUB_NUM_OPER_XXX
      far_llave!far_NUMOPER2 = PUB_NUM_OPER_XXX
      ' quite en arequipa
      
      'If i_cambio.Value = 1 And i_cambio.Visible = True And LK_CODTRA <> 2401 Then
      '   far_llave!FAR_fecha_compra = WW_FECHA_ANT
      '   far_llave!far_NUMOPER2 = WW_NUM_OPER
      'End If
      If i_cambio.Value = 1 And i_cambio.Visible = True And i_fecha_compra.Visible = True And pub_signo_arm = 1 Then
         If far_llave!FAR_fecha_compra <> LK_FECHA_DIA Then
         far_llave!far_NUMOPER2 = 0
      End If
      End If
      
      
      far_llave!far_precio_neto = Val(grid_fac.TextMatrix(fila, 35)) 'GUARDA PRECIO DE LISTA
      
   '   If Check1.Value = 1 And PUB_TIPMOV = 20 Then
   '      far_llave!FAR_fecha = PUB_FECHA
   '      far_llave!FAR_NUMOPER = WW_NUM_OPER
   '   End If
      
      
      far_llave!far_otra_cia = " "
      far_llave!far_subtra = SUT_LLAVE!sut_descripcion
      If LK_CODTRA = 1409 And pub_signo_arm = -1 Then
         far_llave!far_subtra = "Transferencia"
      End If
      
      If i_cias.Visible = True Then
         far_llave!far_transito = "T"
         far_llave!far_subtra = Left(SUT_LLAVE!sut_descripcion, 7) & " A " & i_cias.Text
         far_llave!far_otra_cia = Right(i_cias.Text, 2)
      ElseIf Trim(SUT_LLAVE!SUT_FLAG_TIPO) = "P" Then
         far_llave!far_transito = Trim(SUT_LLAVE!SUT_FLAG_TIPO)
         far_llave!far_subtra = Trim(SUT_LLAVE!sut_descripcion)
      Else
         far_llave!far_transito = " "
      End If
         
      If Boton_Recepcion.Visible Then
         far_llave!far_subtra = Left(SUT_LLAVE!sut_descripcion, 10) & " De " & grid_trans.TextMatrix(1, 3)
         far_llave!far_otra_cia = grid_trans.TextMatrix(grid_trans.Row, 6)
         far_llave!far_transito = "Y"
      End If
      If LK_CODTRA = 2437 Then
         far_llave!far_otra_cia = grid_fac.TextMatrix(fila, 10)
      End If
      far_llave!far_descri = Nulo_Valors(grid_fac.TextMatrix(fila, 5))
      far_llave!far_JABAS = PUB_JABAS '
      far_llave!far_UNIDADES = PUB_UNIDAD
      far_llave!far_mortal = 0 ' Val(Nulo_Valor0(Mid(grid_fac.TextMatrix(fila, 9), 23, 5)))
      far_llave!far_num_precio = Val(grid_fac.TextMatrix(fila, 17))
      far_llave!far_PESO = Val(grid_fac.TextMatrix(fila, 37))
      far_llave!far_litro = Val(grid_fac.TextMatrix(fila, 43))
      
      far_llave!FAR_TIPDOC = PUB_TIPDOC
      
      
      far_llave!FAR_ORDEN_UNIDADES = Val(grid_fac.TextMatrix(fila, 19))
      far_llave!far_concepto = PUB_CONCEPTO
      If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
         far_llave!far_concepto = Nulo_Valors(i_TEXTONCRE.Text)
      End If
      far_llave!far_codusu = PUB_CODUSU
'      far_llave!far_CODTRA = SUT_LLAVE!SUT_CODTRA
'      far_llave!far_SUBCODTRA = SUT_LLAVE!SUT_SECUENCIA
      far_llave!FAR_HORA = Format(Now, "hh:mm:ss AMPM")
      far_llave!FAR_NUM_LOTE = SUT_LLAVE!sut_secuencia ' fue para piura sec=10 trans=2403
      'If i_codsunat.ListIndex >= 0 Then far_llave!far_COD_SUNAT = Val(Right(i_codsunat.Text, 5))
      far_llave!FAR_PEDSER = PUB_PEDSER
      
      ' AGREGE ACV
      If i_TEXTONCRE.Visible And PUB_TIPMOV <> 10 Then
         far_llave!FAR_PEDFAC = Val(i_TEXTONCRE.Text)
      Else
         far_llave!FAR_PEDFAC = PUB_PEDFAC
      End If
      far_llave!far_turno = PUB_TURNO ' PARA EL CANAL DE DISTRIBUCION
      far_llave!far_pedsec = Val(grid_fac.TextMatrix(fila, 18)) ' PARA EL DIA DE VISITA
      PUB_VISITA = Val(grid_fac.TextMatrix(fila, 18))
      far_llave!FAR_FLAG_SO = PUB_SO
      If LK_CODTRA = 2401 Then
         PUB_NUMFAC_C = PUB_PEDFAC ' el Graba Numero de Pedido
         PUB_NUMSER_C = PUB_VISITA ' el dia de Visita
         far_llave!FAR_NUMFAC_C = PUB_PEDFAC
         far_llave!FAR_NUMSER_C = PUB_NUMSER_C
      End If
      far_llave!far_key_dircli = Val(Right(i_dircli.Text, 8))
      If PUB_TIPMOV = 10 Then
          far_llave!FAR_OC = Trim(grid_fac.TextMatrix(fila, 45)) ' Left(i_TEXTONCRE.Text, 15)
      End If
        If i_fecha_oper.Visible And IsDate(i_fecha_oper.Text) Then
           far_llave!FAR_fecha_pro = i_fecha_oper.Text
        Else
           far_llave!FAR_fecha_pro = LK_FECHA_DIA
        End If
        If i_fecha_can.Visible And IsDate(i_fecha_can.Text) Then
         far_llave!FAR_fecha_can = i_fecha_can.Text
        Else
         far_llave!FAR_fecha_can = LK_FECHA_DIA
        End If
'        far_llave!far_ISLA = PUB_ISLA
        far_llave!far_turno = PUB_TURNO
        far_llave.Update
        If ParaLot_count <= 0 Then
        Else
         GoTo Otro_lote
        End If
      pub_autkey = Val(FORMGEN.grid_fac.TextMatrix(fila, 22))

      If Boton_Recepcion.Visible And LK_EMP_PTO = "A" Then graba_fffart

      ws_flag_far = ingre
OTROMAS:

   fila = fila + 1
   If fila > grid_fac.Rows - 1 Then
      FLAG = True
   End If
Loop
loc_flag_1111 = "A"
   

If LK_CODTRA = 1409 And ww_respuesta = vbYes Then
     ww_respuesta = vbNo
     i_fbg.ListIndex = 2
     pub_signo_arm = -1
     PUB_TIPMOV = 5
     PUB_SO = ""
     GoTo SIGUE_1409
End If


'If pub_autkey <> 0 Then
'   aut_llave.MoveFirst
'   Do Until aut_llave.EOF
'         aut_llave.Edit
'         aut_llave!aut_estado = "A"
'         aut_llave.Update
'         aut_llave.MoveNext
'   Loop
'   ' SENDMAIL "alancosme@almacen", "AUTORIZACIONES", "PASE AUTORIZADO.... EJECUTADA"'
'End If

flag_salto = 0
FORMGEN.i_subtotal.Text = Val(FORMGEN.i_subtotal.Text)
PUB_SUBTOTAL = FORMGEN.i_subtotal.Text
Return


ACT2:
' agrege acv 22-07
If PUB_TIPMOV = 20 And LK_CODTRA = 1401 Then
 If Val(par_llave!PAR_SER_KARDEX) <> Val(FORMGEN.i_numser.Text) Then
  par_llave.Edit
  par_llave!PAR_SER_KARDEX = Val(FORMGEN.i_numser.Text)
  par_llave.Update
 End If
End If
If PUB_TIPMOV = 93 And LK_CODTRA = 2414 Then
 If Val(par_llave!par_serie) <> Val(FORMGEN.i_numser.Text) Then
  par_llave.Edit
  par_llave!par_serie = Val(FORMGEN.i_numser.Text)
  par_llave.Update
 End If
End If

WW_FECHA_ANT = LK_FECHA_DIA
WW_NUM_OPER = PUB_NUM_OPER_XXX

If i_cambio.Value = 0 Or i_cambio.Visible = False Then Return
If PUB_PROCESO = 1 Then Return

far_llave.MoveFirst

If far_llave.EOF Then Return

If LK_CODTRA = 1401 Then
   far_llave.MoveLast
   If Val(i_codcli.Text) <> far_llave!far_codclie Then pub_flag_cambio = 1
   far_llave.MoveFirst
End If

If pub_flag_cambio = 0 And LK_CODTRA = 1401 Then
   Do Until far_llave.EOF
      far_llave.Edit
      far_llave!FAR_NUMFAC_C = PUB_NUMFAC_C
      far_llave!FAR_NUMSER_C = PUB_NUMSER_C
      far_llave!far_numguia = PUB_NUMGUIA
      If i_fecha_compra.Visible Then
        far_llave!FAR_fecha_compra = i_fecha_compra.Text
      Else
        far_llave!FAR_fecha_compra = LK_FECHA_DIA
      End If
      If LK_FLAG_GRIFO = "A" And LK_CODTRA = 1401 Then
        far_llave!far_turno = Val(i_turno.Text)
        far_llave!FAR_PEDFAC = Val(i_TEXTONCRE.Text)
      End If
      far_llave!FAR_DIAS = pub_dias
      far_llave!far_serguia = Val(i_serguia.Text)
      far_llave!far_numguia = PUB_NUMGUIA
      far_llave!FAR_MONEDA = WS_MONEDA_CLI
      
      pub_signo_car = far_llave!far_signo_car
      far_llave.Update
      PUB_NUMFAC = far_llave!far_numfac
      PUB_NUMSER = far_llave!far_numser
      If far_llave!far_estado <> "E" Then PUB_NUM_OPER_EXT = far_llave!FAR_NUMOPER

      
      far_llave.MoveNext
   Loop
Else
   far_llave.MoveLast      ' PARA VER CUAL ES EL VALIDO...CUANDO HAY VARIAS ANULACIONES DEL MISMO DOC.
   WS_NUMSEC = far_llave!FAR_NUMSEC
   If far_llave!far_estado = "E" Then Return

   SQ_OPER = 5
   PUB_FECHA = far_llave!FAR_fecha
   WW_FECHA_ANT = far_llave!FAR_fecha_compra
   WW_NUM_OPER = far_llave!far_NUMOPER2
   PUB_NUM_OPER_EXT = far_llave!FAR_NUMOPER
   
   far_llave.MoveFirst ' ES PARA SU EXTORNO EN ACT2
   
   LEER_FAR_LLAVE
   If far_menor3.EOF = True Then
      CN.Execute "Rollback Transaction", rdExecDirect
      MsgBox "error grave en ARTICULO ..."
      GoTo fin
   End If
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_FECHA = PUB_FECHA
    LEER_ALL_LLAVE
    Do Until all_llave.EOF
      ' quite en arequipa
      ' far_llave!far_NUMOPER2
      If PUB_NUM_OPER_EXT = all_llave!ALL_NUMOPER Then
        Exit Do
      Else
        all_llave.MoveNext
      End If
    Loop
    If all_llave.EOF Then
          CN.Execute "Rollback Transaction", rdExecDirect
          MsgBox "error grave en ALLOG ..."
          GoTo fin
    End If
    all_llave.Edit
    all_llave!all_flag_ext = "E"
    all_llave.Update
    WK_MODI_1401 = "A"
    
    
 
'End If
 

   
Do Until far_llave.EOF
   If far_llave!far_estado = "E" Then GoTo SALTA_OTRO
   I = 0
   far_menor3.AddNew
   Do Until I = 79
      If Not IsNull(far_llave(I)) Then far_menor3(I) = far_llave(I)
      I = I + 1
   Loop
   WS_NUMSEC = WS_NUMSEC + 200
   far_menor3!FAR_NUMSEC = WS_NUMSEC
   
   far_menor3!far_estado = "E"
   
   far_menor3!far_signo_arm = far_llave!far_signo_arm * -1
   pub_signo_arm = far_menor3!far_signo_arm
   
   'AGREGE ACV
   'far_menor3!far_NUMOPER = PUB_NUM_OPER_XXX
   'Print PUB_NUM_OPER_XXX
   far_menor3.Update
   
   PUB_CODART = far_llave!far_codart
   far_llave.Edit
   far_llave!far_estado = "E"
   far_llave.Update
   
   If PUB_CODART = 0 Then GoTo SALTA_OTRO
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
     CN.Execute "Rollback Transaction", rdExecDirect
     MsgBox "error grave en ARTICULO ..."
     GoTo fin
   End If
   
   WS_CANTIDAD = far_llave!far_cantidad
   PUB_JABAS = far_llave!far_JABAS
   
   'AGREGE ACV
   pub_dias = far_llave!FAR_DIAS
'   pub_signo_arm = far_llave!far_SIGNO_aRM * -1
   
   If pub_signo_arm = -1 Or pub_signo_arm = 1 Then
   Else
      CN.Execute "Rollback Transaction", rdExecDirect
      MsgBox "error grave en SIGNO ARM ..."
      GoTo fin
   End If

      arm_llave.Edit
      arm_llave!ARM_STOCK = Val(arm_llave!ARM_STOCK) + WS_CANTIDAD * pub_signo_arm
      arm_llave!arm_stock2 = Val(arm_llave!arm_stock2) + PUB_JABAS * pub_signo_arm
      
      If PUB_SO = "A" Then
         arm_llave!ARM_saldo_s = Val(arm_llave!ARM_saldo_s) + WS_CANTIDAD * pub_signo_arm
         'arm_llave!ARM_SALDO_S2 = Val(arm_llave!ARM_SALDO_S2) + PUB_JABAS * pub_signo_arm
      Else
         arm_llave!ARM_Saldo_n = Val(arm_llave!ARM_Saldo_n) + WS_CANTIDAD * pub_signo_arm
         arm_llave!ARM_SALDO_N2 = Val(arm_llave!ARM_SALDO_N2) + PUB_JABAS * pub_signo_arm
      End If
      

      
      If pub_signo_arm = -1 Then
         arm_llave!ARM_SALIDAS = Val(arm_llave!ARM_SALIDAS) + Val(WS_CANTIDAD)
      Else
      If pub_signo_arm = 1 Then
         arm_llave!ARM_INGRESOS = Val(arm_llave!ARM_INGRESOS) + Val(WS_CANTIDAD)
      End If
      End If
      If LK_CODTRA = 1401 Then
         If i_fecha_compra.Visible Then
             arm_llave!ARM_FECHA_ULT = i_fecha_compra.Text
         Else
             arm_llave!ARM_FECHA_ULT = LK_FECHA_DIA
         End If
      End If
      arm_llave.Update
       'AGREGE ACV
      If LK_CODTRA = 1401 And i_cambio.Value = 1 Then
         pub_signo_arm = 1
      End If
      
      pub_signo_car = far_llave!far_signo_car


   PUB_NUMFAC = far_llave!far_numfac
   PUB_NUMSER = far_llave!far_numser
SALTA_OTRO:
   far_llave.MoveNext
  Loop
  
End If

If pub_signo_car = 0 Then GoTo ffin


If LK_CODTRA = 2414 Then GoTo ffin


SQ_OPER = 5
pu_codcia = LK_CODCIA
LEER_CAR_LLAVE
If car_far2.EOF = True Then
      pub_mensaje = "!!! ERROR EN ACTUALIZACION de CANJE..."
      exito = False
      GoTo Error_fatal
End If


SQ_OPER = 1
pu_cp = cli_llave!CLI_CP
pu_codclie = car_far2!CAR_CODCLIE
pu_codcia = LK_CODCIA
PUB_SERDOC = car_far2!car_SERDOC
PUB_NUMDOC = car_far2!car_NUMDOC
PUB_TIPDOC = car_far2!car_TIPDOC
LEER_CAR_LLAVE
If Val(car_llave!CAR_IMP_INI) = 0 Then GoTo SSS
car_llave.Edit
car_llave!CAR_SITUACION = "E"
car_llave.Update


SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_cp = PUB_CP
PUB_NUM_OPER = PUB_NUM_OPER_EXT
LEER_CAA_LLAVE
If caa_histo.EOF = True Then
   MsgBox "DOCUMENTO EN CARACU NO EXISTE..."
   exito = False
End If
Do Until caa_histo.EOF
    caa_histo.Edit
    caa_histo!CAA_ESTADO = "E"
    caa_histo.Update
    caa_histo.MoveNext
Loop

SSS:
' AGREGE ACV PARA CARACU
If Val(i_codcli.Text) <> car_llave!CAR_CODCLIE Then pub_flag_cambio = 1

If pub_flag_cambio = 0 And LK_CODTRA = 1401 Then
car_llave.Edit
car_llave!car_numser_c = PUB_NUMSER_C
car_llave!car_numfac_c = PUB_NUMFAC_C
car_llave!car_NUMGUIA = PUB_NUMGUIA
car_llave!car_SERguia = i_serguia.Text
If i_fecha_compra.Visible Then
  car_llave!car_fecha_sunat = CDate(i_fecha_compra.Text)
Else
  car_llave!car_fecha_sunat = LK_FECHA_DIA
End If
car_llave!car_fecha_vcto = PUB_FECHA_VCTO
car_llave!car_fecha_VCTO_orig = PUB_FECHA_VCTO
car_llave.Update
pub_signo_car = 0

GoTo ffin
End If

SQ_OPER = 1
pu_cp = all_llave!ALL_CP
pu_codcia = all_llave!all_CODCIA
pu_codclie = all_llave!ALL_CODCLIE
PUB_CODCLIE = all_llave!ALL_CODCLIE
LEER_CLI_LLAVE



pub_signo_car = -1
PUB_IMPORTE_AMORT = Val(car_far2!CAR_IMP_INI)
'If LK_CODTRA = 2748 And i_cambio.Value = 1 Then PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
fx = 32000
GoSub ACT8
fx = 0
'GoSub ACT7
pub_signo_car = SUT_LLAVE!SUT_SIGNO_CAR
PUB_IMPORTE_AMORT = Val(i_neto.Text)
'If LK_CODTRA = 2748 And i_cambio.Value = 1 Then
'Else
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
'End If

SQ_OPER = 1
pu_cp = SUT_LLAVE!SUT_cp
pu_codcia = LK_CODCIA
pu_codclie = Val(i_codcli.Text)
PUB_CODCLIE = pu_codclie
LEER_CLI_LLAVE


ffin:
pub_signo_arm = SUT_LLAVE!SUT_SIGNO_ARM
pub_signo_car = SUT_LLAVE!SUT_SIGNO_CAR


Return






ACT3:
' EL CAMPO SIGNO_CAR SOLO VALE EN LA 2745 Y PARA CONSULTAS
If LK_CODTRA = 2412 Then Aplica_NC
If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
   If Val(gridl.TextMatrix(1, 5)) <> 0 Then Return
End If

If LK_CODTRA <> 5712 And PUB_IMPORTE_AMORT = 0 And PUB_TIPMOV <> 20 Then Return

If pub_flag_cambio = 0 And PUB_TIPMOV = 20 And i_cambio.Value = 1 Then Return

If pub_signo_car = 0 Then
   Return
End If

SQ_OPER = 3
pu_codcia = LK_CODCIA
pu_cp = PUB_CP
LEER_CAR_LLAVE
If car_menor.EOF Then
   PUB_NUMDOC = 1
Else
   PUB_NUMDOC = car_menor!car_NUMDOC + 1
End If

car_llave.AddNew
car_llave!CAR_CODCLIE = PUB_CODCLIE
car_llave!car_codcia = LK_CODCIA
car_llave!car_NUMGUIA = PUB_NUMGUIA
car_llave!car_SERguia = Val(i_serguia.Text)
car_llave!car_TIPDOC = PUB_TIPDOC
car_llave!CAR_cp = PUB_CP
car_llave!car_SERDOC = PUB_SERDOC
car_llave!car_NUMDOC = PUB_NUMDOC
car_llave!CAR_FECHA_INGR = LK_FECHA_DIA
car_llave!car_fecha_vcto = PUB_FECHA_VCTO
If LK_CODTRA = 1455 And PUB_TIPDOC = "LR" Then
car_llave!car_fecha_VCTO_orig = ws_fecha_vcto_orig
Else
car_llave!car_fecha_VCTO_orig = PUB_FECHA_VCTO
End If
car_llave!CAR_FLAG_SO = PUB_SO

If i_fecha_compra.Visible Then
   car_llave!car_fecha_sunat = CDate(i_fecha_compra.Text)
   ' agrege acv
   ws_fecha_emision = CDate(i_fecha_compra.Text)
Else
   car_llave!car_fecha_sunat = LK_FECHA_DIA
   ' agrege acv
   ws_fecha_emision = LK_FECHA_DIA
End If

car_llave!CAR_COMISION = 0

car_llave!CAR_SITUACION = " "
car_llave!CAR_NUM_REN = PUB_SECUENCIA
' agrege acv
If LK_CODTRA = 1401 Then PUB_CONCEPTO = PUB_CONCEPTO & "Desct. (%) :" & PUB_DESCTO
car_llave!car_concepto = PUB_CONCEPTO
If LK_EMP = "3AA" And PUB_TIPMOV = 10 Then
   car_llave!car_concepto = SUT_LLAVE!sut_descripcion
End If
car_llave!car_nombre_banco = PUB_NOMBRE_BANCO
car_llave!car_NUM_CHEQUE = PUB_NUM_CHEQUE
car_llave!car_SIGNO_CAJA = pub_signo_caja
car_llave!car_NUMGUIA = PUB_NUMGUIA
' CAMTEX
If LK_CODTRA = 5712 Then PUB_IMPORTE_AMORT = PUB_IMPORTE

WS_IMPORTE_AMORT = PUB_IMPORTE_AMORT

WS_TOT = WS_IMPORTE_AMORT
If LK_CODTRA = 2741 And Trim(SUT_LLAVE!SUT_abreviado) <> "PUB" Then
 ' PUB_FBG = PUB_FBG
  PUB_NUMSER = PUB_NUMSER_C
  PUB_NUMFAC = PUB_NUMFAC_C
  'car_llave!CAR_IMP_INI = Val(InputBox("Importe Original = ", "Ingrese Importe:", WS_IMPORTE_AMORT * pub_signo_car, 2500, 3300))
  car_llave!CAR_IMP_INI = WS_IMPORTE_AMORT * pub_signo_car
  car_llave!CAR_IMP_INI = LOC_IMP_INI
Else
  car_llave!CAR_IMP_INI = WS_IMPORTE_AMORT * pub_signo_car
End If
car_llave!car_importe = WS_IMPORTE_AMORT * pub_signo_car
LOC_SALDO_CAR = WS_IMPORTE_AMORT
wS_saldo_car = car_llave!car_importe
car_llave!car_codtra = LK_CODTRA
'If PUB_PRECIO2 <> 0 Then CREO QUE NO SE USA
'   car_llave!car_PRECIO = 1 / PUB_PRECIO2
'Else
car_llave!car_PRECIO = LOC_DEFINI ' 0
'End If

car_llave!CAR_MONEDA = WS_MONEDA_CLI

If LK_EMP = "PLA" And LK_CODTRA = 2745 Then
   car_llave!CAR_MONEDA = SUT_LLAVE!SUT_MONEDA_CAJA
End If

car_llave!CAR_signo_car = pub_signo_car
car_llave!car_NUMSER = PUB_NUMSER
car_llave!car_NUMFAC = PUB_NUMFAC
car_llave!CAR_TIPMOV = PUB_TIPMOV
car_llave!car_FBG = PUB_FBG
car_llave!CAR_codven = PUB_CODVEN
car_llave!CAR_COBRADOR = PUB_CODVEN
car_llave!car_numser_c = PUB_NUMSER_C
car_llave!car_numfac_c = PUB_NUMFAC_C
car_llave!CAR_NUMOPER = PUB_NUM_OPER_XXX

' CAMTEX
If PUB_TIPDOC = "LE" And LK_CODTRA <> 2741 Then car_llave!car_numfac_c = Val(PUB_NUM_CHEQUE)

' CAMTEX
If LK_CODTRA = 1455 Then car_llave!car_numfac_c = PUB_NUM_CHEQUE

car_llave!CAR_codban = PUB_CODBAN

car_llave.Update
ws_flag_car = ingre

GoSub registra_caa

Return

ACT4:
pub_signo_arm = 0  'ASEGURA NO ARTICULOS....CUANDO SE EXTORNA
If LK_CODTRA = 2401 And Val(i_neto.Text) = 0 Then Return
If LK_CODTRA = 1455 Then i_concepto.Text = ""


'revisar bien para las monedas.... 25/03/00
If LK_CODTRA = 2401 Then GoTo SALTA_CANJE
If LK_CODTRA = 2752 Then GoTo SALTA_CANJE
If LK_CODTRA = 2750 Then GoTo SALTA_CANJE
FILAX = 2
pub_signo_car = -1
Do Until FILAX > gridl.Rows - 1
   If Val(gridl.TextMatrix(FILAX, 5)) = 0 Then GoTo pasando
   
   PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
   PUB_NUMDOC = gridl.TextMatrix(FILAX, 12)
   PUB_SERDOC = gridl.TextMatrix(FILAX, 11)
   PUB_TIPDOC = gridl.TextMatrix(FILAX, 13)
   PUB_IMPORTE_AMORT = Val(gridl.TextMatrix(FILAX, 5))
   i_concepto.Text = Trim(i_concepto.Text) & gridl.TextMatrix(FILAX, 1) & " " & gridl.TextMatrix(FILAX, 2) & "-" & gridl.TextMatrix(FILAX, 3) & "..."
   PUB_CONCEPTO = "Cancel. X Canje"
   SQ_OPER = 1
   pu_cp = PUB_CP
   pu_codclie = PUB_CODCLIE
   pu_codcia = LK_CODCIA
   LEER_CAR_LLAVE
   If car_llave.EOF Then
      pub_mensaje = "!!! ERROR EN ACTUALIZACION de CANJE..."
      exito = False
      GoTo Error_fatal
   End If
   car_llave.Edit
ws_fecha_vcto_orig = car_llave!car_fecha_VCTO_orig
WS_TOT = PUB_IMPORTE_AMORT
WS_IMPORTE_AMORT = PUB_IMPORTE_AMORT
car_llave!car_importe = car_llave!car_importe + WS_IMPORTE_AMORT * pub_signo_car
wS_saldo_car = car_llave!car_importe
LOC_SALDO_CAR = car_llave!car_importe
If pub_signo_car = -1 Then
   car_llave!car_fecha_vcto = PUB_FECHA_VCTO
End If
' camtex
If LK_CODTRA <> 1455 Then
WS_MONEDA_CLI = car_llave!CAR_MONEDA
End If
ws_mon_canje = car_llave!CAR_MONEDA
ws_fecha_emision = car_llave!car_fecha_sunat
car_llave.Update
GoSub registra_caa
ws_flag_car = MODIF
GoSub ACT7
pasando:
FILAX = FILAX + 1
Loop

SALTA_CANJE:
If Val(grid_canje.TextMatrix(1, 3)) = Val(i_neto.Text) And LK_CODTRA = 2401 Then
   PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
   pub_signo_car = -1
   PUB_IMPORTE_AMORT = Val(grid_canje.TextMatrix(1, 3))
   PUB_CONCEPTO = "Cancel. X Canje"
   SQ_OPER = 1
   pu_cp = PUB_CP
   pu_codclie = PUB_CODCLIE
   pu_codcia = LK_CODCIA
   LEER_CAR_LLAVE
   If car_llave.EOF Then
      pub_mensaje = "!!! ERROR EN ACTUALIZACION de CANJE..."
      exito = False
      GoTo Error_fatal
   End If

   car_llave.Edit
   WS_TOT = PUB_IMPORTE_AMORT
   WS_IMPORTE_AMORT = PUB_IMPORTE_AMORT
   car_llave!car_importe = car_llave!car_importe + WS_IMPORTE_AMORT * pub_signo_car
   wS_saldo_car = car_llave!car_importe
   LOC_SALDO_CAR = car_llave!car_importe
   car_llave.Update
   GoSub registra_caa
   PUB_TIPMOV = 0
   GoSub ACT7
End If



FILAX = 2

Do Until FILAX > grid_canje.Rows - 1
If Val(grid_canje.TextMatrix(FILAX, 3)) <> 0 Then
PUB_TIPDOC = grid_canje.TextMatrix(FILAX, 2)
If PUB_TIPDOC = "CH" Then
   pub_signo_car = -1
   pub_signo_caja = 0
End If

If LK_CODTRA = 1455 Then
   pub_signo_car = 1
   WS_MONEDA_CLI = Trim(i_ds.Text) ' ws_mon_canje
End If

If PUB_TIPDOC = "LE" Then pub_signo_car = 1

PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
PUB_NUMGUIA = 0
PUB_FECHA_VCTO = grid_canje.TextMatrix(FILAX, 4)
PUB_CONCEPTO = grid_canje.TextMatrix(FILAX, 7)
If LK_CODTRA = 1455 Then PUB_CONCEPTO = i_concepto.Text

If SUT_LLAVE!SUT_TIPMOV = 10 Then PUB_CONCEPTO = BolFac.Caption & " " & PUB_CONCEPTO
PUB_IMPORTE_AMORT = grid_canje.TextMatrix(FILAX, 3)
PUB_CHENUM = Val(Left(grid_canje.TextMatrix(FILAX, 6), 9))
PUB_NUM_CHEQUE = Val(Left(grid_canje.TextMatrix(FILAX, 6), 9))
PUB_CODBAN = Val(Right(grid_canje.TextMatrix(FILAX, 5), 7))
PUB_NUMSER = 0
PUB_NUMFAC = 0
PUB_TIPMOV = 0

GoSub ACT3
GoSub ACT7

'AUTOCANCELACION DE  CHEQUE POR CONTADO
If SUT_LLAVE!SUT_signo_caja = 1 And PUB_TIPDOC = "CH" Then
   pub_signo_car = 1
   pub_signo_caja = 0
   PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
   SQ_OPER = 1
   pu_cp = PUB_CP
   pu_codclie = PUB_CODCLIE
   pu_codcia = LK_CODCIA
   LEER_CAR_LLAVE
   If car_llave.EOF Then
      pub_mensaje = "!!! ERROR EN ACTUALIZACION de CANJE..."
      exito = False
      GoTo Error_fatal
   End If
   GoSub ACT8
   WS_MONEDA_CLI = car_llave!CAR_MONEDA
   GoSub ACT7
End If
End If
FILAX = FILAX + 1
Loop

Return


ACT5:
'PARTE DEL CODIGO

If ANEXO_ACT5 = True Then Return

If (i_codban2.Visible = True Or Val(Nulo_Valor0(SUT_LLAVE!sut_codban)) = 1) And flag_repite = 0 Then
   GoSub ACT7
   PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
   If i_codban2.Visible = True Then
      PUB_CODBAN = i_codban2.Text
      If LK_CODTRA <> 5318 Then
       i_importe.Text = i_importe_amort.Text
      End If
   Else
      PUB_CODBAN = 1
   End If
   PUB_CONCEPTO = PUB_ABREVIADO & " :" & FORMGEN.i_nomban.Caption
   If LK_CODTRA = 5318 Then
   If pub_signo_ccm = -1 Then
     PUB_CONCEPTO = "TRANSF. CARGO"
   Else
     PUB_CONCEPTO = "TRANSF. ABONO"
   End If
   End If

   'PUB_CHENUM = 0
   'i_chenum.Text = 0
   SQ_OPER = 1
   LEER_CCM_LLAVE
   flag_repite = 1
   PUB_IMPORTE = Val(i_importe.Text)
   pub_signo_ccm = 1
   GoTo ACT5
End If


Return

ACT6:
ANEXO_ACT6
Return

ACT7:

If i_cambio.Value = 0 And WS_ESTADO = 3 Then LLENA_RECIBO_ALL
If WS_ESTADO = 99 Or WS_ESTADO = 98 Or WS_ESTADO = 97 Then LLENA_NUM_RECIBO
LLENA_NUM_DIARIO


If PUB_NUM_OPER_XXX = 0 Then Return
all_llave.AddNew
all_llave!ALL_NUMOPER = PUB_NUM_OPER_XXX
all_llave!all_CODCIA = LK_CODCIA
all_llave!ALL_codtra = LK_CODTRA
all_llave!all_flag_ext = Nulo_Valors(tra_llave!TRA_FLAG_EXT)
If PUB_TIPMOV = 20 And Check1.Value = 1 Then
   all_llave!all_flag_ext = "X"
End If
all_llave!ALL_CODCLIE = PUB_CODCLIE
all_llave!ALL_CODART = PUB_CODART
all_llave!ALL_IMPORTE_amort = PUB_IMPORTE_AMORT
all_llave!all_codusu = PUB_CODUSU
all_llave!ALL_FBG = PUB_FBG
all_llave!ALL_CODVEN = PUB_CODVEN

all_llave!ALL_IMPORTE = PUB_IMPORTE
all_llave!ALL_IMPORTE_DOLL = 0
If LK_CODTRA = 2410 Then all_llave!ALL_IMPORTE_DOLL = Val(i_importe_amort.Text)
If PUB_TIPMOV = 20 Then
 all_llave!ALL_CODSUNAT = "01"  ' SIEMPRE FACTURA  PARA COMPRAS DE MERCADERIA
ElseIf PUB_TIPMOV = 97 And LK_CODTRA = 2410 Then
 all_llave!ALL_CODSUNAT = "07"  ' SIEMPRE 2410  PARA COMPRAS DE MERCADERIA n/c
Else
 all_llave!ALL_CODSUNAT = Val(Right(i_codsunat.Text, 5))
End If
If PUB_TIPMOV = 20 Then 'PENDIENTE
   If Tipo_Cambio.Visible = True And cli_llave!CLI_MONEDA = "D" Then
      If Mid(Tipo_Cambio.Caption, 2, 1) = "S" Then all_llave!ALL_IMPORTE_DOLL = WS_IMPORTE_AMORT
   End If
End If
all_llave!ALL_NUMDOC = PUB_NUMDOC
all_llave!ALL_CP = PUB_CP
all_llave!ALL_TIPDOC = PUB_TIPDOC
all_llave!all_numfac_c = PUB_NUMFAC_C
all_llave!all_numser_c = PUB_NUMSER_C
all_llave!all_codban = PUB_CODBAN
all_llave!all_concepto = PUB_CONCEPTO
all_llave!all_chenum = PUB_CHENUM
all_llave!ALL_FECHA_DIA = LK_FECHA_DIA '**
all_llave!ALL_FECHA_VCTO = PUB_FECHA_VCTO '**

If IsDate(ws_fecha_emision) Then all_llave!all_fecha_ant = ws_fecha_emision
' quite acv
'If IsDate(ws_fecha_anterior) Then all_llave!all_fecha_ant = ws_fecha_anterior

all_llave!ALL_CANTIDAD = PUB_CANTIDAD
all_llave!ALL_NUMSER = PUB_NUMSER
all_llave!all_numfac = PUB_NUMFAC
all_llave!all_neto = PUB_NETO  '**IMPORTE
all_llave!ALL_BRUTO = PUB_SUBTOTAL
all_llave!ALL_tipmov = PUB_TIPMOV
all_llave!ALL_IMPTO = PUB_IMPTO
all_llave!ALL_flete = PUB_FLETE
all_llave!ALL_HORA = Now
all_llave!ALL_LIMCRE_ACT = PUB_LIMCRE_ACT
all_llave!ALL_LIMCRE_ANT = PUB_LIMCRE_ANT
all_llave!ALL_DESCTO = PUB_DESCTO
all_llave!ALL_GASTOS = PUB_GASTOS
all_llave!ALL_PRECIO = PUB_PRECIO2
'all_llave!ALL_FLAG_SO = PUB_SO
all_llave!ALL_RUC = PUB_RUC
all_llave!ALL_MONEDA_CLI = WS_MONEDA_CLI
If (LK_CODTRA = 2725 Or LK_CODTRA = 2774 Or LK_CODTRA = 2770) And LK_EMP = "HER" Then
 all_llave!ALL_MONEDA_CLI = PUB_DS
End If

all_llave!ALL_moneda_ccm = WS_MONEDA_CCM
all_llave!ALL_MONEDA_CAJA = PUB_DS

If LK_EMP = "PLA" And LK_CODTRA = 2725 Then
   all_llave!ALL_MONEDA_CAJA = SUT_LLAVE!SUT_MONEDA_CAJA
   all_llave!ALL_MONEDA_CLI = SUT_LLAVE!SUT_MONEDA_CAJA
End If

'all_llave!all_SECUENCIA = PUB_SECUENCIA
If Val(loc_secuencia) <> Val(SUT_LLAVE!sut_secuencia) Then
 MsgBox "Favor Avisar al Administrador este mensaje: < Secuencia a cambiado ...Verificar en Transaccion >", vbInformation, "Aviso..."
 all_llave!all_SECUENCIA = loc_secuencia 'SUT_LLAVE!SUT_SECUENCIA
Else
 all_llave!all_SECUENCIA = SUT_LLAVE!sut_secuencia
End If
all_llave!ALL_SIGNO_CAR = pub_signo_car
all_llave!ALL_signo_caja = pub_signo_caja

all_llave!ALL_SIGNO_CCM = pub_signo_ccm
all_llave!all_sIGNO_ARM = pub_signo_arm

If WS_ESTADO = 3 Or WS_ESTADO = 2 Or LK_CODTRA = 5350 Then
   all_llave!ALL_NUM_RECIBO = Val(i_numfac.Text)
   all_llave!ALL_SERIE_REC = Val(i_numser.Text)
Else
   all_llave!ALL_NUM_RECIBO = 0
   all_llave!ALL_SERIE_REC = 0
End If

all_llave!all_chenum = PUB_CHENUM
all_llave!ALL_CHESEC = PUB_CHESEC
all_llave!ALL_CHESER = PUB_CHESER
all_llave!ALL_SUBTRA = SUT_LLAVE!sut_descripcion

If car_llave.RowCount > 0 Then
   all_llave!ALL_SITUACION = Nulo_Valors(car_llave!CAR_SITUACION)
Else
   all_llave!ALL_SITUACION = ""
End If
all_llave!all_codtra_ext = LK_CODTRA 'Left(FORMGEN.TRANS.Text, 4)

all_llave!ALL_TIPO_CAMBIO = LK_TIPO_CAMBIO ' ACV WS_TIPO_CAMBIO
'acv
'If WS_TIPO_CAMBIO <= 1 Then
  all_llave!ALL_TIPO_CAMBIO = LK_TIPO_CAMBIO
'End If

If PUB_TIPMOV = 10 Or PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
   If all_llave!ALL_FBG = "F" Then
          WS_AUTOCON = "Fact. " & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
   ElseIf all_llave!ALL_FBG = "B" Then
          WS_AUTOCON = "Bolet." & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
   ElseIf all_llave!ALL_FBG = "G" Or all_llave!ALL_FBG = "P" Then
          WS_AUTOCON = "Guia. " & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
   ElseIf all_llave!ALL_FBG = "N" Then
          WS_AUTOCON = "N.CRED. " & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
   ElseIf all_llave!ALL_FBG = "D" Then
          WS_AUTOCON = "N.DEB.  " & all_llave!ALL_NUMSER & "-" & all_llave!all_numfac
   End If
   WS_AUTOCON = WS_AUTOCON & ":" & Trim(Left(cli_llave!cli_nombre, 25))
End If

If Val(all_llave!all_codban) <> 0 And Not ccm_llave.EOF Then
   If all_llave!all_chenum <> 0 And SUT_LLAVE!SUT_cp = "P" And Not cli_llave.EOF And PUB_CODCLIE <> 0 Then
      WS_AUTOCON = SUT_LLAVE!SUT_abreviado & ":" & all_llave!all_chenum & ", Fact." & all_llave!all_numfac_c & ":" & Trim(Left(cli_llave!cli_nombre, 18)) & ":" & Left(ccm_llave!CCM_NOMBRE, 15)
   ElseIf all_llave!all_chenum <> 0 Then
      WS_AUTOCON = SUT_LLAVE!SUT_abreviado & ":" & all_llave!all_chenum & ":" & Trim(Left(ccm_llave!CCM_NOMBRE, 15))
   Else
      WS_AUTOCON = Trim(ccm_llave!CCM_NOMBRE) & ":" & i_concepto.Text
   End If
End If
      
If pub_signo_car = -1 And pub_signo_caja = 1 And Val(gridl.TextMatrix(1, 5)) <> 0 Then
   WS_AUTOCON = ""
   If car_llave!car_FBG = "F" Then
          WS_AUTOCON = "Fact. " & car_llave!car_NUMSER & "-" & car_llave!car_NUMFAC
   ElseIf car_llave!car_FBG = "B" Then
          WS_AUTOCON = "Bolet." & car_llave!car_NUMSER & "-" & car_llave!car_NUMFAC
   ElseIf car_llave!car_FBG = "G" Or car_llave!car_FBG = "P" Then
          WS_AUTOCON = "Guia. " & car_llave!car_NUMSER & "-" & car_llave!car_NUMFAC
   End If
   WS_AUTOCON = WS_AUTOCON & " " & Trim(Left(cli_llave!cli_nombre, 18))
End If
 
If pub_signo_caja = 1 And Val(gridl.TextMatrix(1, 19)) <> 0 And PUB_TIPDOC = "CH" Then
   WS_AUTOCON = "Cheque N. " & PUB_NUM_CHEQUE
   WS_AUTOCON = WS_AUTOCON & ":" & Left(cli_llave!cli_nombre, 18)
End If

If Val(grid_canje.TextMatrix(1, 3)) <> 0 And letche.ToolTipText = "CH" And PUB_TIPMOV = 10 Then
   WS_AUTOCON = WS_AUTOCON & ":" & " Cheque N. " & Val(grid_canje.TextMatrix(2, 6))
End If
If Trim(PUB_CONCEPTO) <> "" Then
   WS_AUTOCON = WS_AUTOCON & ":" & PUB_CONCEPTO
End If
'   If ws_codclie <> 0 Then
'         WS_NOMCLI = Left(cli_llave!cli_nombre, 18) & ":"
'   End If
'   If all_llave!ALL_SIGNO_CAR <> 0 Then
'       WS_NOMCLI = Left(WS_NOMCLI, 30) & car_llave!CAR_fbg & "-." & car_llave!CAR_numSER & "-" & car_llave!CAR_numfac
'   End If
WS_AUTOCON = Left(SUT_LLAVE!sut_descripcion, 18) & ":" & PUB_CONCEPTO
If Not cli_llave.EOF And PUB_CODCLIE <> 0 Then
   If Trim(WS_AUTOCON) = "" Then WS_AUTOCON = Left(SUT_LLAVE!sut_descripcion, 25) & "  : " & Left(cli_llave!cli_nombre, 18) & ":" & all_llave!all_concepto
End If
If LK_CODTRA = 5350 Then
  WS_AUTOCON = Nulo_Valors(i_TEXTONCRE.Text)
End If
all_llave!ALL_numoper2 = WS_NUM_OPER2
all_llave!ALL_autocon = WS_AUTOCON
If i_fecha_compra.Visible = True Then
  all_llave!ALL_FECHA_SUNAT = i_fecha_compra.Text
Else
  all_llave!ALL_FECHA_SUNAT = LK_FECHA_DIA
End If
If fragas.Visible = True Then
  all_llave!ALL_IMPG1 = Val(imp_gas1.Text)
  all_llave!ALL_IMPG2 = Val(imp_gas2.Text)
  all_llave!ALL_CTAG1 = cta_gas1.Text
  all_llave!ALL_CTAG2 = cta_gas2.Text
Else
  all_llave!ALL_IMPG1 = 0
  all_llave!ALL_IMPG2 = 0
  all_llave!ALL_CTAG1 = " "
  all_llave!ALL_CTAG2 = " "
End If
If i_fecha_oper.Visible And IsDate(i_fecha_oper.Text) Then
     all_llave!ALL_FECHA_PRO = i_fecha_oper.Text
Else
    all_llave!ALL_FECHA_PRO = LK_FECHA_DIA
End If
If i_fecha_can.Visible And IsDate(i_fecha_can.Text) Then
    all_llave!ALL_FECHA_CAN = i_fecha_can.Text
Else
    all_llave!ALL_FECHA_CAN = LK_FECHA_DIA
End If
If LK_CODTRA = 5318 And i_fecha_oper.Visible Then ' forzar para registrar otra Fecha de giro.
  If pub_signo_ccm = -1 Then all_llave!ALL_FECHA_CAN = i_fecha_oper.Text
End If

If LK_CODTRA = 2401 And i_limcre_ant.Visible Then
  all_llave!ALL_flete = Val(i_limcre_ant.Text)
End If
all_llave!ALL_NUMGUIA = PUB_NUMGUIA
If (all_llave!ALL_NUMOPER <> all_llave!ALL_numoper2) And all_llave!ALL_codtra = 2410 Then
 all_llave!ALL_FACART = " "
 all_llave!ALL_tipmov = 0
Else
 all_llave!ALL_FACART = WS_FLAG_FACART
End If
 

all_llave.Update
WK_MODI_1401 = "A"
LOC_FLAG_ALLOG = "A"
loc_flag_1111 = "A"
' SACT. PARA PEDIDOS PROCESADO MARCA = 'P'
If LK_CODTRA = 2401 And fraPedidos.Visible = True Then
 pub_cadena = "UPDATE PEDIDOS SET PED_SITUACION = 'P', PED_FORMA = '" & PUB_FBG & "/" & Format(PUB_NUMSER, "000") & "-" & Format(PUB_NUMFAC, "00000000") & "' WHERE PED_CODCIA = '" & LK_CODCIA & "' AND PED_NUMSER = " & PUB_PEDSER & " AND PED_NUMFAC = " & PUB_PEDFAC & " AND PED_TIPMOV = 201"
 CN.Execute pub_cadena, rdExecDirect
End If
If LK_CODTRA = 1111 And ww_codtra_ext = 2401 Then
  If LOC_PEDIDO <> -1 Then
    'Anular o no Anular el pedido
    ANULAR_PED
  Else
    pub_cadena = "UPDATE PEDIDOS SET PED_SITUACION = ' ', PED_FORMA = ' ' WHERE PED_CODCIA = '" & LK_CODCIA & "' AND PED_NUMSER = " & PUB_PEDSER & " AND PED_NUMFAC = " & PUB_PEDFAC & " AND PED_TIPMOV = 201"
    CN.Execute pub_cadena, rdExecDirect
  End If
End If

Return


ACT8:
If pub_signo_car = 0 Then
   Return
End If

car_llave.Edit

WS_TOT = PUB_IMPORTE_AMORT
WS_IMPORTE_AMORT = PUB_IMPORTE_AMORT

If gridl.Visible = True Then
   WS_TOT = Val(gridl.TextMatrix(1, 9))
End If

car_llave!car_importe = car_llave!car_importe + WS_IMPORTE_AMORT * pub_signo_car
If LK_CODTRA = 1111 Then car_llave!CAR_SITUACION = "E"


wS_saldo_car = car_llave!car_importe

LOC_SALDO_CAR = car_llave!car_importe

If LK_CODTRA = 1111 Then
   car_llave!car_fecha_vcto = ws_fecha_anterior
Else
   ws_fecha_anterior = car_llave!car_fecha_vcto
End If


If (FORMGEN.i_fecha_vcto.Visible = True And LK_CODTRA <> 1111) Or gridl.Visible = True Or grid_liq.Visible = True Then
   car_llave!car_fecha_vcto = PUB_FECHA_VCTO
End If

If FORMGEN.i_concepto.Visible = True And LK_CODTRA <> 2412 And LK_CODTRA <> 2410 And LK_CODTRA <> 5360 Then
   car_llave!car_concepto = PUB_CONCEPTO
End If

If PUB_IMPORTE_AMORT <> 0 And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 And LK_CODTRA <> 1133 And car_llave!car_importe <> 0 And LK_EMP <> "HER" Then
   car_llave!CAR_NUM_REN = car_llave!CAR_NUM_REN + 1
End If

If LK_EMP = "HER" Then car_llave!CAR_COBRADOR = Val(gridl.TextMatrix(FILAX, 7))
If LK_CODTRA = 2725 And LK_EMP = "PAR" Then
  WW_FECHA_COBRO = gridl.TextMatrix(FILAX, 22)
Else
  WW_FECHA_COBRO = LK_FECHA_DIA
End If
If LK_CODTRA = 1111 Or LK_CODTRA = 1122 Then  ' para mantener enlace del doc.
 PUB_NUMSER_C = car_llave!car_numser_c
 PUB_NUMFAC_C = car_llave!car_numfac_c
 If ww_codtra_ext = 2105 Then
    car_llave!car_concepto = ""
    car_llave!car_numser_c = 0
    car_llave!car_numfac_c = 0
 End If
 If ww_codtra_ext = 2101 Then
    car_llave!car_numfac_c = -1  ' es para que no aparezca como canje
 End If
 End If
 
If LK_CODTRA = 2410 Or LK_CODTRA = 2412 Then
  PUB_NUMSER_C = Val(FORMGEN.i_numser_c.Text)
  PUB_NUMFAC_C = Val(FORMGEN.i_numfac_c.Text)
End If

car_llave.Update

GoSub registra_caa

ws_flag_car = MODIF
Return

ACT9:
ANEXO_ACT9
Return

ACT10:
 ANEXO_ACT10
Return

ACT11:

If i_numplan.Visible = True And i_codcli.Visible = False Then
   par_llave.Edit
   ' par_llave!par_planilla = par_llave!par_planilla + 1
   'i_numplan.Text = par_llave!par_planilla
   pub_numplan = Val(i_numplan.Text)
   par_llave.Update
End If

If Val(gridl.TextMatrix(1, 5)) = 0 And (PUB_TIPMOV = 97 Or PUB_TIPMOV = 98) Then
   GoSub ACT7
   Return
End If

If Val(gridl.TextMatrix(2, 12)) = 0 And Val(gridl.TextMatrix(1, 5)) = 0 Then GoTo continua

If i_cambio.Visible = False Or i_cambio.Value = 0 Then LLENA_NUM_RECIBO


FILAX = 2


If SUT_LLAVE!SUT_signo_ccm = -1 Then
   i_concepto.Text = "O/P : " & "N. " & i_chenum.Text & "  " & Trim(Left(ccm_llave!CCM_NOMBRE, 20))
End If

Do Until FILAX > gridl.Rows - 1
      If SUT_LLAVE!SUT_SIGNO_CAR = 2 Then
         If Val(gridl.TextMatrix(FILAX, 25)) < 0 Then
            pub_signo_car = 1
         Else
            pub_signo_car = -1
         End If
      Else
        pub_signo_car = SUT_LLAVE!SUT_SIGNO_CAR
      End If
      PUB_CODCLIE = Val(gridl.TextMatrix(FILAX, 15))
      PUB_CP = gridl.TextMatrix(FILAX, 10)
      PUB_NUMDOC = Val(gridl.TextMatrix(FILAX, 12))
      PUB_SERDOC = Val(gridl.TextMatrix(FILAX, 11))
      PUB_TIPDOC = gridl.TextMatrix(FILAX, 13)
      If LK_EMP = "PIU" Then
         PUB_IMPORTE_AMORT = Val(gridl.TextMatrix(FILAX, 5)) + Val(gridl.TextMatrix(FILAX, 19))
      Else
         PUB_IMPORTE_AMORT = Val(gridl.TextMatrix(FILAX, 5))
      End If
      
      If IsDate(gridl.TextMatrix(FILAX, 6)) = True Then PUB_FECHA_VCTO = CDate(gridl.TextMatrix(FILAX, 6))
      
      If PUB_IMPORTE_AMORT <> 0 Then
         ww_concepto = Trim(gridl.TextMatrix(FILAX, 1)) & "./ " & gridl.TextMatrix(FILAX, 2) & "-" & gridl.TextMatrix(FILAX, 3)
      End If
      If Trim(gridl.TextMatrix(FILAX, 23)) <> "" Then
        ww_concepto = Trim(gridl.TextMatrix(FILAX, 23))
        'gridl.TextMatrix(FILAX, 23) = ""
        i_concepto.Text = ""
      End If
      If Trim(i_concepto.Text) = "" Then i_concepto.Text = ww_concepto
      If LK_CODTRA <> 2412 Then
         i_concepto.Text = Trim(i_concepto.Text) & " " & ww_concepto
      End If
      
      SQ_OPER = 1
      pu_cp = PUB_CP
      pu_codclie = PUB_CODCLIE
      pu_codcia = LK_CODCIA
      LEER_CAR_LLAVE
      If car_llave.EOF Then
         pub_mensaje = "!!! ERROR EN ACTUALIZACION de CANJE..."
         exito = False
         GoTo Error_fatal
      End If
      WS_MONEDA_CLI = car_llave!CAR_MONEDA
      If LK_CODTRA = 2410 Then WS_MONEDA_CLI = i_ds.Text
      
      If LK_CODTRA = 2770 Or LK_CODTRA = 2774 Then
        PUB_FBG = car_llave!car_FBG
        PUB_NUMFAC = car_llave!car_NUMFAC
        PUB_NUMSER = car_llave!car_NUMSER
      End If
            
      ws_mon_canje = car_llave!CAR_MONEDA
      ' agrege acv
      ws_fecha_emision = car_llave!car_fecha_sunat
      pub_signo_ccm = 0
      If PUB_IMPORTE_AMORT = 0 Then
         If car_llave!car_fecha_vcto <> PUB_FECHA_VCTO Then
            car_llave.Edit
            car_llave!car_fecha_vcto = PUB_FECHA_VCTO
            car_llave.Update
         End If
         If car_llave!CAR_COBRADOR <> Val(gridl.TextMatrix(FILAX, 7)) And LK_CODTRA = 2770 Then
            car_llave.Edit
            car_llave!CAR_COBRADOR = Val(gridl.TextMatrix(FILAX, 7))
            car_llave.Update
         End If
          ' no se debe Editar o modificar
         'If CDate(car_llave!car_fecha_sunat) <> CDate(gridl.TextMatrix(FILAX, 8)) Then
         '   car_llave.Edit
         '   car_llave!car_fecha_sunat = CDate(gridl.TextMatrix(FILAX, 8))
         '   car_llave.Update
         'End If
         LOC_FLAG_ALLOG = "A"
      Else
         PUB_NUMSER_C = Val(gridl.TextMatrix(FILAX, 2))
         PUB_NUMFAC_C = Val(gridl.TextMatrix(FILAX, 3))
         PUB_SO = gridl.TextMatrix(FILAX, 27)
         ' CAMTEX
         If LK_CODTRA = 2735 Then
             PUB_CONCEPTO = PUB_CONCEPTO + " " + Trim(i_concepto.Text)
             i_concepto.Text = ""
         ElseIf LK_CODTRA = 2725 And PUB_SECUENCIA = 50 Then
             PUB_CONCEPTO = "Canc. con Adelanto-" & Trim(i_concepto.Text)
         Else
             PUB_CONCEPTO = Trim(i_concepto.Text)
         End If
'         PUB_CONCEPTO = Trim(gridl.TextMatrix(FILAX, 1)) & "./ " & gridl.TextMatrix(FILAX, 2) & "-" & gridl.TextMatrix(FILAX, 3)
        If LK_CODTRA = 2735 And pub_signo_car = -1 Then
          PUB_NUMGUIA = PUB_NUMDOC
        End If
         PUB_CODCLIE = car_llave!CAR_CODCLIE
         GoSub CON1
         GoSub ACT8
         GoSub ACT7
         PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
      End If
      
      FILAX = FILAX + 1
Loop
   pub_signo_ccm = SUT_LLAVE!SUT_signo_ccm
   If pub_signo_ccm = 0 Then GoTo continua
   PUB_IMPORTE_AMORT = 0
   If pub_signo_ccm = 0 Then
      PUB_NUM_OPER_XXX = 0
      Return
   End If
   
   pub_signo_caja = 0
   If pub_signo_ccm = 1 Then
     ' pub_signo_caja = -1
      PUB_CONCEPTO = "Dep.Bco. :" & ww_concepto
   ElseIf pub_signo_ccm = -1 Then
    ' CAMTEX
     If LK_CODTRA = 2735 Then
      PUB_CONCEPTO = PUB_CONCEPTO & "- " & Trim(Left(cli_llave!cli_nombre, 25))
     Else
      PUB_CONCEPTO = Trim(ww_concepto) & "- " & Trim(Left(cli_llave!cli_nombre, 25))
     End If
   End If
   pub_signo_car = 0
   If LK_CODTRA = 2725 And LK_EMP = "3AA" Or LK_EMP = "PIU" Then
   Else
    GoSub ACT5
    GoSub ACT7
   End If

continua:

If Val(gridl.TextMatrix(1, 19)) = 0 Then
   textovarl.Visible = False
   Return
End If

FILAX = 2
Do Until FILAX > gridl.Rows - 1
If Val(gridl.TextMatrix(FILAX, 19)) <> 0 Then
If LK_EMP = "PIU" Then
   PUB_TIPDOC = "CD"
   pub_signo_car = 1
Else
   PUB_TIPDOC = "CH"
   pub_signo_car = -1
   If LK_EMP = "HER" Then
     pub_signo_caja = 0
   End If
End If

PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
PUB_NUMGUIA = 0
PUB_FECHA_VCTO = gridl.TextMatrix(FILAX, 22)
PUB_CONCEPTO = gridl.TextMatrix(FILAX, 1) & "/." & gridl.TextMatrix(FILAX, 2) & "-" & gridl.TextMatrix(FILAX, 3)
If LK_CODTRA = 2770 Or LK_CODTRA = 2774 Then
  PUB_FBG = gridl.TextMatrix(FILAX, 1)
  PUB_NUMSER = gridl.TextMatrix(FILAX, 2)
  PUB_NUMFAC = gridl.TextMatrix(FILAX, 3)
End If


PUB_IMPORTE_AMORT = gridl.TextMatrix(FILAX, 19)
PUB_NUM_CHEQUE = gridl.TextMatrix(FILAX, 20)
PUB_CODCLIE = gridl.TextMatrix(FILAX, 15)
'PUB_CODVEN = 0 ' lo quito por que quiero que grave el codigo de vendedor ** alan

GoSub CON1
PUB_NUMSER = 0
PUB_NUMFAC = 0
PUB_TIPMOV = 0
WS_MONEDA_CLI = gridl.TextMatrix(FILAX, 24)
ws_mon_canje = gridl.TextMatrix(FILAX, 24)

GoSub ACT3
GoSub ACT7
End If

FILAX = FILAX + 1
Loop
FORMGEN.i_bancos.Visible = False
textovarl.Visible = False
PUB_NUM_OPER_XXX = 0
Return

ACT12:
ANEXO_ACT12
Return

ACT13:

    GoSub ACT2
      If LK_EMP = "HER" And i_cambio.Value = 1 And PUB_TIPMOV = 10 Then GoTo DETL
      If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
         If Val(grid_fac.TextMatrix(1, 7)) <> 0 Then Return
      End If

      If i_cambio.Visible = False Or i_cambio.Value = 0 Then llena_numfac
      
      far_llave.AddNew
      
     ''  far_llave!far_COD_SUNAT = Val(Right(i_codsunat.Text, 5))

      ANEXO_ACT13
      
      
      If LK_CODTRA = 2410 Then WS_MONEDA_CLI = i_ds.Text
      If LK_CODTRA = 2412 Then
        far_llave!far_IMPTO = Val(FORMGEN.i_impto2.Text)
        far_llave!FAR_BRUTO = Val(FORMGEN.i_bruto2.Text)
        If far_llave!FAR_BRUTO = 0 Then
            pub_mensaje = "!!! Operacion Invalida...NO PROCEDE reiniciar sistema "
            exito = False
            GoTo Error_fatal
        End If
      End If

      PUB_SUBTOTAL = PUB_BRUTO2
      PUB_IMPTO = PUB_IMPTO2
      far_llave!FAR_SUBTOTAL = PUB_IMPORTE_AMORT
      far_llave!FAR_DIAS = pub_dias
      far_llave!FAR_NUMDOC = 0
      far_llave!far_numguia = PUB_NUMGUIA
      far_llave!FAR_tipo_cambio = LK_TIPO_CAMBIO ' acv WS_TIPO_CAMBIO
      'If WS_TIPO_CAMBIO <= 1 Then
      '    far_llave!FAR_tipo_cambio = LK_TIPO_CAMBIO
      'End If
      far_llave!FAR_DESCTO = 0
      far_llave!FAR_TOT_DESCTO = 0
      far_llave!far_JABAS = 0
      far_llave!far_UNIDADES = 0
      
   
      If WS_MONEDA_CLI = "S" Or WS_MONEDA_CLI = "D" Then
      Else
         MsgBox "fallo la moneda"
         End
      End If
      
      far_llave!FAR_MONEDA = WS_MONEDA_CLI
      
      far_llave!far_PESO = 0
      far_llave!far_codusu = PUB_CODUSU
      If i_fecha_compra.Visible Then
        far_llave!FAR_fecha_compra = i_fecha_compra.Text
      Else
        far_llave!FAR_fecha_compra = LK_FECHA_DIA
      End If
'     far_llave!far_CODTRA = SUT_LLAVE!SUT_CODTRA
      If Val(i_neto.Text) = 0 And Trim(grid_fac.TextMatrix(2, 4)) = "ANULACION" And PUB_TIPMOV = 10 Then
         far_llave!far_estado = "E"
         far_llave!FAR_ESTADO2 = "A"
      End If
      
      far_llave!FAR_equiv = 1
      far_llave!far_descri = ""
      far_llave!FAR_equiv = 1
      far_llave!far_descri = ""
      far_llave!FAR_NUM_LOTE = SUT_LLAVE!sut_secuencia
      If PUB_TIPMOV = 20 Then
        far_llave!far_cod_sunat = "01" ' SIEMPRE FACTURA  PARA COMPRAS DE MERCADERIA
      ElseIf PUB_TIPMOV = 97 And LK_CODTRA = 2410 Then
        far_llave!far_cod_sunat = "07"  ' SIEMPRE 2410  PARA COMPRAS DE MERCADERIA n/c
      Else
        far_llave!far_cod_sunat = Val(Right(i_codsunat.Text, 5))
      End If
      far_llave!far_concepto = PUB_CONCEPTO
      far_llave!far_signo_arm = pub_signo_arm
      far_llave!FAR_CODVEN = PUB_CODVEN
      WS_FLAG_FACART = "A"
      If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
         far_llave!far_concepto = Nulo_Valors(i_TEXTONCRE.Text)
      End If
      far_llave!FAR_DIAS = pub_dias
      far_llave!far_numguia = Val(i_numguia.Text)
      far_llave!far_serguia = Val(i_serguia.Text)
      far_llave!far_subtra = Trim(i_def.Text)
      far_llave!far_signo_car = pub_signo_car
      far_llave!FAR_FLAG_SO = PUB_SO
      far_llave!FAR_HORA = Format(Now, "hh:mm:ss AMPM")
        If i_fecha_oper.Visible And IsDate(i_fecha_oper.Text) Then
         far_llave!FAR_fecha_pro = i_fecha_oper.Text
        Else
         far_llave!FAR_fecha_pro = LK_FECHA_DIA
        End If
        If i_fecha_can.Visible And IsDate(i_fecha_can.Text) Then
         far_llave!FAR_fecha_can = i_fecha_can.Text
        Else
         far_llave!FAR_fecha_can = LK_FECHA_DIA
        End If
      far_llave.Update
DETL:

Return

ACT14:
WS_IMPORTE_AMORT = 0
car_llave.Edit
car_llave!CAR_SITUACION = Left(i_situacion.Text, 1)
car_llave.Update
GoSub registra_caa
 
'Boton_Compras_Click 'no se usa
Return

ACT15:
If grid_canje.Visible = False Then Return
FILAX = 2
Do Until FILAX > grid_canje.Rows - 1
If Val(grid_canje.TextMatrix(FILAX, 3)) <> 0 Then
PUB_TIPDOC = grid_canje.TextMatrix(FILAX, 2)
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
PUB_NUMGUIA = 0
PUB_FECHA_VCTO = grid_canje.TextMatrix(FILAX, 6)
PUB_CONCEPTO = grid_canje.TextMatrix(FILAX, 4)
PUB_IMPORTE_AMORT = grid_canje.TextMatrix(FILAX, 3)
PUB_NUM_CHEQUE = grid_canje.TextMatrix(FILAX, 5)
PUB_CODCLIE = grid_canje.TextMatrix(FILAX, 8)
PUB_CODVEN = grid_canje.TextMatrix(FILAX, 9)

GoSub CON1
PUB_NUMSER = 0
PUB_NUMFAC = 0
PUB_TIPMOV = 0
i_fbg.ListIndex = -1
GoSub ACT3
GoSub ACT7
End If

FILAX = FILAX + 1
Loop

grid_canje.Top = 4300
grid_canje.Height = 700
grid_canje.Clear
'gridl.Clear
FORMGEN.i_bancos.Visible = False

Return

ACT16:
 
If pub_signo_car <> 0 Then
   caa_histo.Edit
   caa_histo!CAA_ESTADO = "E"
   ww_numdoc = caa_histo!caa_numfac_c
   ww_numser = caa_histo!caa_numser_c
   caa_histo.Update
End If

If PUB_TIPMOV = 0 And Trim(WS_FLAG_FACART) = "" Then
   Return
End If
ANEXO_ACT16
loc_flag_1111 = "A"
Return

ACT17:
ANEXO_ACT17

Return

ACT18:
ANEXO_ACT18
Return

ACT19:
   ANEXO_ACT19
Return

ACT20:
   ANEXO_ACT20
Return

ACT21:
   ANEXO_ACT21
Return

REP1:
Return
REP2:

Return

REP3:

'PUB_TIPREG = 20
'PUB_NUMTAB = cli_llave!CLI_CASA_ZONA
'SQ_OPER = 1
'LEER_TAB_LLAVE
'If tab_llave.EOF = True Then
'   pub_mensaje_err = "ERROR EN ZONA.."
'End If
'zona_nombre = tab_llave!TAB_NOMLARGO

'PUB_TIPREG = 30
'PUB_NUMTAB = cli_llave!CLI_CASA_SUBZONA
'SQ_OPER = 1
'LEER_TAB_LLAVE
'subzona_nombre = tab_llave!TAB_NOMLARGO


'*** REPORTE DE LETRA ***
'IMP_LETRA CStr(WS_IMPRESION_LET), CStr(PUB_FECHA_VCTO), CStr(PUB_NETO), CStr(PUB_FECHA), CStr(PUB_FECHA_VCTO), CStr(PUB_NETO), cli_llave!cli_nombre, Nulo_Valors(cli_llave!CLI_CASA_DIREC), cli_llave!CLI_CASA_NUM, zona_nombre, subzona_nombre

Return
   

'Screen.MousePointer = 1






GoTo fin


Error_fatal:
    CN.Execute "Rollback Transaction", rdExecDirect
    MsgBox pub_mensaje
    pub_mensaje = "Se ha producido un error " & "al abrir la conexión:" & Err & " - " & Error & vbCr
    For Each er In rdoErrors
        pub_mensaje = pub_mensaje & er.Description & ":" & er.Number & vbCr
        MsgBox pub_mensaje, 48, Pub_Titulo
    Next er

    

End
'    Resume AbandonCn
errorr:

fin:
Screen.MousePointer = 0
grabar.Enabled = True
PUB_CODBAN = 0
PUB_CODCLIE = 0
pasa_def
FORMGEN.Barra.Visible = False
'FORMGEN.i_def.SetFocus

Exit Sub

registra_caa:
If WS_IMPORTE_AMORT = 0 Then Return
If pub_signo_car <> 0 Then
cli_llave.Edit
WS_CANTIDAD = WS_IMPORTE_AMORT * pub_signo_car
cli_llave!cli_SALDO = Nulo_Valor0(cli_llave!cli_SALDO) + WS_CANTIDAD
If pub_signo_car = -1 Then cli_llave!cli_fecha_fac = PUB_FECHA_VCTO
' AGREGE IF PARA NO LIMPIAR EL CLI_MONEDA POR QUE WS_MONEDA_CLI = ""
If Trim(WS_MONEDA_CLI) = "S" Or Trim(WS_MONEDA_CLI) = "D" Then cli_llave!CLI_MONEDA = WS_MONEDA_CLI
cli_llave.Update
End If

caa_histo.AddNew
caa_histo!CAA_CODCLIE = PUB_CODCLIE
caa_histo!caa_codcia = LK_CODCIA
caa_histo!CAA_TIPDOC = PUB_TIPDOC
caa_histo!CAA_CP = PUB_CP
caa_histo!CAA_NUM_OPER = PUB_NUM_OPER_XXX
If grid_liq.Visible = True Then
   caa_histo!caa_INTVEN = grid_liq.TextMatrix(1, 2)
   caa_histo!caa_DIASV = grid_liq.TextMatrix(1, 6)
   caa_histo!caa_DIASA = grid_liq.TextMatrix(1, 7)
   caa_histo!caa_tasav = Val(i_tasav.Text)
Else
   caa_histo!caa_INTVEN = 0
   caa_histo!caa_DIASV = 0
   caa_histo!caa_DIASA = 0
   caa_histo!caa_tasav = 0
End If
caa_histo!caa_TIPO_CAMBIO = LK_TIPO_CAMBIO
caa_histo!caa_serdoc = PUB_SERDOC
caa_histo!CAA_NUMDOC = PUB_NUMDOC
caa_histo!CAA_FECHA = LK_FECHA_DIA
caa_histo!CAA_FECHA_VCTO = PUB_FECHA_VCTO
caa_histo!caa_situacion = PUB_SITUACION_ACT
caa_histo!caa_concepto = PUB_CONCEPTO
caa_histo!CAA_IMPORTE = WS_IMPORTE_AMORT * pub_signo_car

caa_histo!CAA_TOTAL = Abs(WS_TOT) * pub_signo_car
caa_histo!CAA_SALDO = Nulo_Valor0(cli_llave!cli_SALDO)
caa_histo!caa_SALDO_car = wS_saldo_car
If pub_signo_car <> 0 Then
   caa_histo!CAA_SIGNO_CAJA = PUB_SECUENCIA
Else
   caa_histo!CAA_SIGNO_CAJA = 0
End If
caa_histo!CAA_SIGNO_CAJA_REAL = pub_signo_caja
caa_histo!CAA_SIGNO_CAR = pub_signo_car
caa_histo!CAA_TIPMOV = PUB_TIPMOV
caa_histo!CAA_hora = Now
caa_histo!CAA_CODUSU = PUB_CODUSU
caa_histo!CAA_ESTADO = "N"
If Not cli_llave.EOF And cli_llave.RowCount > 0 Then
   If cli_llave!CLI_CODCLIE = Val(i_codcli.Text) Then
      caa_histo!CAa_NOMBRE = cli_llave!cli_nombre
   End If
End If

caa_histo!CAA_NUMPLAN = pub_numplan
If LK_CODTRA = 1111 Then caa_histo!CAA_ESTADO = "E"
If LK_CODTRA = 1122 Then caa_histo!CAA_ESTADO = "E"
If LK_CODTRA = 1133 Then caa_histo!CAA_ESTADO = "E"
If fx = 32000 Then caa_histo!CAA_ESTADO = "E"
If SUT_LLAVE!SUT_SIGNO_CAR = 2 And LK_CODTRA <> 2727 Then caa_histo!CAA_ESTADO = "E"

If i_fecha_compra.Visible Then
   caa_histo!CAA_FECHA_COBRO = CDate(i_fecha_compra.Text)
Else
   caa_histo!CAA_FECHA_COBRO = LK_FECHA_DIA
End If

If ws_flag_car = ingre Then
   caa_histo!CAa_NUM_CHEQUE = PUB_NUM_CHEQUE
   caa_histo!CAa_numser = PUB_NUMSER
   caa_histo!CAa_numfac = PUB_NUMFAC
   caa_histo!caa_numser_c = PUB_NUMSER_C
   caa_histo!caa_numfac_c = PUB_NUMFAC_C
   If LK_CODTRA = 2741 Then
    caa_histo!CAa_numser = PUB_NUMSER_C
    caa_histo!CAa_numfac = PUB_NUMFAC_C
   End If
   caa_histo!CAa_numGUIA = PUB_NUMGUIA
   caa_histo!CAa_SERGUIA = Val(i_serguia.Text)
   caa_histo!CAA_FBG = PUB_FBG
   caa_histo!CAA_CODVEN = PUB_CODVEN
   caa_histo!caa_situacion = " "
   caa_histo!caa_FLAG_SO = PUB_SO
Else
   caa_histo!CAa_NUM_CHEQUE = Nulo_Valors(car_llave!car_NUM_CHEQUE)
   caa_histo!CAa_numser = Nulo_Valor0(car_llave!car_NUMSER)
   caa_histo!CAa_numfac = Nulo_Valor0(car_llave!car_NUMFAC)
   If LK_CODTRA = 2412 Then
     caa_histo!caa_numser_c = PUB_NUMSER
     caa_histo!caa_numfac_c = PUB_NUMFAC
   Else
     caa_histo!caa_numser_c = PUB_NUMSER_C
     caa_histo!caa_numfac_c = PUB_NUMFAC_C
   End If
   If LK_CODTRA = 1122 Or LK_CODTRA = 1111 Then
      If ww_codtra_ext = 2412 Or ww_codtra_ext = 2410 Then
         caa_histo!caa_numfac_c = ww_numdoc
         caa_histo!caa_numser_c = ww_numser
      End If
   End If
   
   caa_histo!CAA_NOTA = PUB_FBG
   caa_histo!CAA_FBG = Nulo_Valors(car_llave!car_FBG)
   If LK_CODTRA = 2770 And i_codven.Visible = True Then
      caa_histo!CAA_CODVEN = PUB_CODVEN
   Else
      caa_histo!CAA_CODVEN = Nulo_Valor0(car_llave!CAR_codven)
   End If
   If PUB_SECUENCIA = 1 And LK_EMP = "HER" Then caa_histo!CAA_CODVEN = Nulo_Valor0(car_llave!CAR_codven)
      caa_histo!CAa_numGUIA = Nulo_Valor0(car_llave!car_NUMGUIA)
   caa_histo!CAa_SERGUIA = Nulo_Valor0(car_llave!car_SERguia)
   caa_histo!caa_situacion = Nulo_Valors(car_llave!CAR_SITUACION)
   caa_histo!caa_FLAG_SO = Nulo_Valors(car_llave!CAR_FLAG_SO)
   caa_histo!caa_signo_ccm = pub_signo_ccm
   caa_histo!caa_codban = PUB_CODBAN
End If
   

caa_histo!caa_codTRA = LK_CODTRA

If WS_ESTADO = 2 And (LK_CODTRA = 2725 Or LK_CODTRA = 2770 Or LK_CODTRA = 2774) Then
   caa_histo!CAA_RECIBO = Val(i_numfac.Text) 'ES EL MISMO CAMPO
   caa_histo!CAA_SERIE = Val(i_numser.Text)
Else
   caa_histo!CAA_RECIBO = 0
   caa_histo!CAA_SERIE = 0
End If
If LK_CODTRA = 2728 And LK_EMP = "PLA" Then
  caa_histo!caa_situacion = PLAZA_FLAG_MANUAL
End If
caa_histo.Update
loc_flag_1111 = "A"
Return


Exit Sub
ALERTA:
MsgBox Err.Description
 CN.Execute "Rollback Transaction", rdExecDirect
 con_llave.Close
 Barra.Visible = False
 MsgBox "Intente nuevamente ,si persiste favor de Reiniciar el Sistema..." & Chr(13) & "Si el problema continua llamar a su proveedor ", vbCritical, Pub_Titulo
 grabar.Enabled = True
' Resume Next

End Sub

Private Sub Grid_all_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
  Grid_all_MouseDown 2, 0, 0, 0
End If
End Sub

Private Sub Grid_all_KeyPress(KeyAscii As Integer)
Dim WTC As Currency
If KeyAscii = 27 Then
 Grid_all.Visible = False
 Diario.SetFocus
End If
If KeyAscii = 13 And Grid_all.COL = 13 Then
pu_codcia = LK_CODCIA
PUB_FECHA = i_fecha_oper.Text
LEER_ALL_LLAVE
Do Until all_llave.EOF
   If all_llave!ALL_NUMOPER = Grid_all.TextMatrix(Grid_all.Row, 0) Then Exit Do
   all_llave.MoveNext
Loop
WTC = 0
WTC = Val(InputBox("Ingrese Tipo de Cambio :"))
If WTC = 0 Then Exit Sub

all_llave.Edit
all_llave!ALL_TIPO_CAMBIO = WTC
all_llave.Update
End If
End Sub

Private Sub Grid_all_LostFocus()
FORMGEN.Grid_all.Visible = False

End Sub

Private Sub Grid_all_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim WNUMERO
Dim WS_CANTIDAD As Currency
If LK_CODTRA <> 1111 Then
   GoTo fin
End If
If Button <> 2 Then
   GoTo fin
End If
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
Else
 If Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 2770 Then
   MsgBox "Solo puede Anular una Cobranza el Usuario: SUPERVISOR !!!", 48, Pub_Titulo
   Grid_all.SetFocus
  Exit Sub
 End If
End If
If CDate(Grid_all.TextMatrix(Grid_all.Row, 11)) <> LK_FECHA_DIA Then
If loc_mofi_venta <> "A" And Val(Grid_all.TextMatrix(Grid_all.Row, 2)) Then
  'BolFac_Click
  MsgBox "No tiene Acceso a modificar documentos anteriores a la fecha del día", vbExclamation, Pub_Titulo
  Exit Sub
End If '
End If

If Trim(Grid_all.TextMatrix(Grid_all.Row, 12)) = "E" Then
      MsgBox "Ya Extornado...", 48, Pub_Titulo
      Grid_all.SetFocus
      GoTo fin
End If
If (Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 1401 Or Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 2401) And Val(Grid_all.TextMatrix(Grid_all.Row, 18)) <> 0 Then
    pu_cp = Grid_all.TextMatrix(Grid_all.Row, 22)
    pu_codclie = Val(Grid_all.TextMatrix(Grid_all.Row, 17))
    pu_codcia = LK_CODCIA
    PUB_SERDOC = Val(Grid_all.TextMatrix(Grid_all.Row, 15))
    PUB_NUMDOC = Val(Grid_all.TextMatrix(Grid_all.Row, 16))
    PUB_TIPDOC = Grid_all.TextMatrix(Grid_all.Row, 14)
    SQ_OPER = 1
    LEER_CAR_LLAVE
    If car_llave.EOF Then
        MsgBox "Verificar Documento....", 48, Pub_Titulo
    Else
      If Val(car_llave!car_importe) <> Val(Grid_all.TextMatrix(Grid_all.Row, 5)) Then
        MsgBox "Documento  tiene amortizaciones", 48, Pub_Titulo
        Grid_all.SetFocus
        Exit Sub
      End If
    End If
End If
If Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 2401 And Val(Grid_all.TextMatrix(Grid_all.Row, 18)) <> 0 Then
   If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
   Else
      MsgBox "No se puede anular Ventas al Credito o Reparto. solo el Supervisor", 48, Pub_Titulo
      Exit Sub
   End If
    pu_cp = "C"
    pu_codclie = Val(Grid_all.TextMatrix(Grid_all.Row, 17))
    pu_codcia = LK_CODCIA
    PUB_SERDOC = Val(Grid_all.TextMatrix(Grid_all.Row, 15))
    PUB_NUMDOC = Val(Grid_all.TextMatrix(Grid_all.Row, 16))
    PUB_TIPDOC = Grid_all.TextMatrix(Grid_all.Row, 14)
    SQ_OPER = 1
    LEER_CAR_LLAVE
    If car_llave.EOF Then
      MsgBox "Verificar Venta....", 48, Pub_Titulo
    Else
      If Val(car_llave!car_importe) <> Val(Grid_all.TextMatrix(Grid_all.Row, 5)) Then
       MsgBox "Venta Tiene Amortizaciones", 48, Pub_Titulo
       Grid_all.SetFocus
       Exit Sub
      End If
    End If
End If
LOC_PEDIDO = -1
If Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 2401 Then
    ' CONFIRMAR LA ANULACION DE PEDIDO
    '******************
    pub_mensaje = "Desea Anular Tambien el Pedido Relacionado... ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       LOC_PEDIDO = -1
    Else
       LOC_PEDIDO = Val(Grid_all.TextMatrix(Grid_all.Row, 21))
    End If
End If


If (Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 1401) And Val(Grid_all.TextMatrix(Grid_all.Row, 18)) <> 0 Then
  SQ_OPER = 6
  pu_codcia = LK_CODCIA
  If Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 1401 Then
   PU_TIPMOV = 20
  Else
   PU_TIPMOV = 0
  End If
  PUB_CODCLIE = Val(Grid_all.TextMatrix(Grid_all.Row, 17))
  PUB_NUMSER_C = Val(Grid_all.TextMatrix(Grid_all.Row, 20))
  PUB_NUMFAC_C = Val(Grid_all.TextMatrix(Grid_all.Row, 21))
  LEER_CAR_LLAVE
  If far_llave2.EOF Then
     MsgBox "No Procede ... Extornar la Cancelación...", 48, Pub_Titulo
     Grid_all.SetFocus
     Exit Sub
     
  End If
End If
  
If Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 1401 Or Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 2401 Or Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 2414 Or Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 2407 Or Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 2409 Or Val(Grid_all.TextMatrix(Grid_all.Row, 2)) = 2403 Then
 SQ_OPER = 1
 PUB_TIPREG = 60
 WNUMERO = Format(Grid_all.TextMatrix(Grid_all.Row, 19), "mm") & Format(Grid_all.TextMatrix(Grid_all.Row, 19), "yyyy")
 PUB_NUMTAB = Val(WNUMERO)
 PUB_CODCIA = LK_CODCIA
 LEER_TAB_LLAVE
 If Not tab_llave.EOF Then
   If tab_llave!tab_codart = 1 Then
     BolFac_Click
     MsgBox "El documento pertenece a un Periodo cerrado, no procede", vbInformation, Pub_Titulo
     Exit Sub
   End If
 End If
End If


FORMGEN.i_num_oper.Text = Grid_all.TextMatrix(Grid_all.Row, 1)
REGRESA:
ext_conta = 0
fila = 1
Grid_all.Visible = False
Screen.MousePointer = 11
Do Until fila = Grid_all.Rows
   Grid_all.COL = 1
   Grid_all.Row = fila
   If Val(FORMGEN.i_num_oper.Text) = Grid_all.Text And Trim(Grid_all.TextMatrix(Grid_all.Row, 12)) <> "E" Then
      FORMGEN.Grid_all.CellBackColor = vbRed
      ext_conta = ext_conta + 1
      If ext_conta = 1 Then PUB_NUM_OPER = Grid_all.TextMatrix(Grid_all.Row, 0)
   Else
      Grid_all.RowHeight(Grid_all.Row) = 0
   End If
   fila = fila + 1
Loop
Screen.MousePointer = 0
Grid_all.Visible = True
If ext_conta > 1 Then
   pub_mensaje = " estas " & ext_conta & " operaciones"
Else
   pub_mensaje = " esta Operacion"
End If
'agrege acv
PUB_CONCEPTO = ""
Grid_all.Visible = True
pub_mensaje = MsgBox("Esta seguro de Extornar " & pub_mensaje, 36, Pub_Titulo)
If pub_mensaje = vbNo Then
   FORMGEN.i_num_oper.Text = 0
   Grid_all.Visible = False
   Diario.SetFocus
   GoTo fin
Else
   FORMGEN.i_num_oper.Text = PUB_NUM_OPER
   Grid_all.Visible = True
   MsgBox "Dar click en Grabar... para registrar el extorno.", 48, Pub_Titulo
   grabar.SetFocus
End If

NUMERO = FORMGEN.Diario.WhatsThisHelpID
avanza_campo

fin:

End Sub

Private Sub grid_autorizacion_GotFocus()
Dim cambiar_color As Boolean
Dim Tit As String
grid_autorizacion.Clear
grid_autorizacion.Cols = 4
grid_autorizacion.ColWidth(0) = 600
grid_autorizacion.ColWidth(1) = 3000
grid_autorizacion.ColWidth(2) = 2700
grid_autorizacion.ColWidth(3) = 1500

grid_autorizacion.TextMatrix(0, 0) = "N.Autoriz."
grid_autorizacion.TextMatrix(0, 1) = "Nombre del Cliente"
grid_autorizacion.TextMatrix(0, 2) = "Articulos ..."
grid_autorizacion.TextMatrix(0, 3) = "Concepto"

grid_autorizacion.Rows = 2
pub_autkey = 30000
pu_codcia = LK_CODCIA
PUB_FECHA = LK_FECHA_DIA
SQ_OPER = 3
LEER_AUT_LLAVE
If aut_menor.EOF = True Then
   MsgBox "Lo siento ... No tiene autorizaciones ", 48, Pub_Titulo
   grid_autorizacion.Visible = False
   GoTo fin
End If
fila = 1
Do Until aut_menor.EOF

If Trim(LK_CODUSU) = Trim(aut_menor!aut_codusu_final) And aut_menor!AUT_SECUENCIA = 1 Then

cambiar_color = False
If Nulo_Valors(aut_menor!aut_estado) = "A" Or Nulo_Valors(aut_menor!aut_estado) = "X" Then
   cambiar_color = True
End If
grid_autorizacion.Rows = grid_autorizacion.Rows + 1

grid_autorizacion.Row = fila
grid_autorizacion.COL = 0
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If
grid_autorizacion.Text = aut_menor!aut_key


grid_autorizacion.COL = 1
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If
SQ_OPER = 1
pu_cp = "C"
pu_codcia = LK_CODCIA
pu_codclie = aut_menor!aut_codclie
LEER_CLI_LLAVE

grid_autorizacion.Text = cli_llave!cli_nombre



SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
PUB_KEY = aut_menor!aut_codart
LEER_ART_LLAVE
grid_autorizacion.COL = 2
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If
grid_autorizacion.Text = Left(art_LLAVE!art_nombre, 25) & "..."


grid_autorizacion.COL = 3
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If


grid_autorizacion.Text = Nulo_Valors(aut_menor!AUT_concepto)
fila = fila + 1

End If
aut_menor.MoveNext

Loop
fin:

End Sub

Private Sub grid_autorizacion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If grid_autorizacion.Cols <> 4 Then Exit Sub
pub_autkey = grid_autorizacion.TextMatrix(grid_autorizacion.Row, 0)

grid_autorizacion.Clear
grid_autorizacion.Cols = 7
grid_autorizacion.ColWidth(0) = 600
grid_autorizacion.ColWidth(1) = 2400
grid_autorizacion.ColWidth(2) = 0
grid_autorizacion.ColWidth(3) = 1900
grid_autorizacion.ColWidth(4) = 1200
grid_autorizacion.ColWidth(5) = 1200
grid_autorizacion.ColWidth(6) = 0

grid_autorizacion.TextMatrix(0, 0) = "N.Autoriz."
grid_autorizacion.TextMatrix(0, 1) = "Nombre del Cliente"
grid_autorizacion.TextMatrix(0, 3) = "Articulos "
grid_autorizacion.TextMatrix(0, 4) = "Cantidad"
grid_autorizacion.TextMatrix(0, 5) = "Precio"


pu_codcia = LK_CODCIA
SQ_OPER = 1
LEER_AUT_LLAVE
If aut_llave.EOF = True Then
   MsgBox "Lo siento ... no hay autorizaciones"
   grid_autorizacion.Visible = False
   GoTo fin
End If
fila = 1
Do Until aut_llave.EOF
If Nulo_Valors(aut_llave!aut_estado) = "A" Then
   MsgBox "Ya procesada... "
   grid_autorizacion.Visible = False
   Exit Do
End If
If Nulo_Valors(aut_llave!aut_estado) = "X" Then
   MsgBox "Anulada ... "
   grid_autorizacion.Visible = False
   Exit Do
End If

grid_autorizacion.Row = fila
If aut_llave!AUT_SECUENCIA = 1 Then
grid_autorizacion.COL = 0
grid_autorizacion.Text = aut_llave!aut_key
grid_autorizacion.COL = 1
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
PUB_CODCLIE = aut_llave!aut_codclie
LEER_CLI_LLAVE
grid_autorizacion.Text = cli_llave!cli_nombre
End If

grid_autorizacion.COL = 2
grid_autorizacion.Text = aut_llave!aut_codart



SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
PUB_KEY = aut_llave!aut_codart
LEER_ART_LLAVE
grid_autorizacion.COL = 3
grid_autorizacion.Text = Left(art_LLAVE!art_nombre, 35)


grid_autorizacion.COL = 4
grid_autorizacion.Text = aut_llave!aut_CANTIDAD

grid_autorizacion.COL = 5
grid_autorizacion.Text = aut_llave!aut_PRECIO

grid_autorizacion.COL = 6
grid_autorizacion.Text = aut_llave!aut_codclie


aut_llave.MoveNext

fila = fila + 1
Loop
fin:

End Sub

Private Sub grid_autorizacion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 45 And KeyCode <> 46 Then
   GoTo fin
End If
If LK_CODTRA = 2401 And KeyCode = 46 Then Exit Sub

If grid_autorizacion.Cols <> 7 Then Exit Sub

If KeyCode = 46 Then
   pub_mensaje = " ¿Desea Eliminar esta Autorizacion  ... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then Exit Sub
   SQ_OPER = 1
   pub_autkey = grid_autorizacion.TextMatrix(1, 0)
   PUB_FECHA = LK_FECHA_DIA
   pu_codcia = LK_CODCIA
   LEER_AUT_LLAVE
   Do Until aut_llave.EOF
   aut_llave.Edit
   aut_llave!aut_estado = "X"
   aut_llave.Update
   aut_llave.MoveNext
   Loop
   MsgBox "Autorizacion Eliminada... "
   grid_autorizacion.Visible = False
End If
   
   
   
   
If Val(grid_fac.TextMatrix(1, 4)) <> 0 Then
   MsgBox "No debe haber datos para la transferencia ..."
   GoTo fin
End If

SQ_OPER = 1
pub_autkey = grid_autorizacion.TextMatrix(1, 0)
PUB_FECHA = LK_FECHA_DIA
pu_codcia = LK_CODCIA
LEER_AUT_LLAVE
If aut_llave.EOF = True Then
   MsgBox "Numero de Autorizacion errado"
   GoTo fin
End If

fila = 2
Do Until aut_llave.EOF
   grid_fac.Rows = fila + 2
   
   PUB_KEY = aut_llave!aut_codart
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ART_LLAVE
   If art_LLAVE.EOF And PUB_KEY <> 0 Then
      MsgBox "Error Grave en arti..."
      Exit Sub
   End If
   PUB_CODART = aut_llave!aut_codart
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ARM_LLAVE
   
   
   i_codcli.Text = aut_llave!aut_codclie
   If PUB_KEY <> 0 Then grid_fac.TextMatrix(fila, 0) = art_LLAVE!art_nombre
   grid_fac.TextMatrix(fila, 1) = aut_llave!aut_codart
   grid_fac.TextMatrix(fila, 16) = aut_llave!aut_codart
   grid_fac.TextMatrix(fila, 33) = aut_llave!aut_codart
   grid_fac.TextMatrix(fila, 2) = 0
   grid_fac.TextMatrix(fila, 3) = 0
   grid_fac.TextMatrix(fila, 11) = arm_llave!arm_cospro
   grid_fac.TextMatrix(fila, 4) = aut_llave!aut_CANTIDAD
   grid_fac.TextMatrix(fila, 6) = aut_llave!aut_PRECIO
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(aut_llave!aut_EQUIV)
   grid_fac.TextMatrix(fila, 5) = Nulo_Valors(aut_llave!aut_descri)
   grid_fac.TextMatrix(fila, 8) = 0
   grid_fac.TextMatrix(fila, 10) = 0
   grid_fac.TextMatrix(fila, 12) = -1
   grid_fac.TextMatrix(fila, 11) = arm_llave!arm_cospro
   grid_fac.TextMatrix(fila, 22) = aut_llave!aut_key
   fila = fila + 1
   aut_llave.MoveNext
Loop
calcula_totales
grid_autorizacion.Visible = False






fin:

End Sub

Private Sub grid_autorizacion_LostFocus()
grid_autorizacion.Visible = False
End Sub

Private Sub grid_canje_EnterCell()
If Trim(letche.ToolTipText) = "" Then
   MsgBox "Dar Click Derecho y seleccionar Tipo de Doc "
   Exit Sub
 End If
PROCESA_CELDAS2
End Sub

Private Sub grid_canje_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

 If grid_canje.COL = 5 Then
    i_bancos.Left = grid_canje.Left + grid_canje.CellLeft
    i_bancos.Width = grid_canje.CellWidth
    i_bancos.Top = grid_canje.Top + grid_canje.CellTop
    textovar_canje.Width = 30
    textovar_canje.Height = 30
    i_bancos.Visible = True
    i_bancos_Click
Else
PROCESA_CELDAS2
End If
End Sub


Private Sub Grid_detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
 If Grid_detalle.Rows <= 3 Then
  Grid_detalle.Rows = 2
 Else
  Grid_detalle.RemoveItem Grid_detalle.Row
 ' chkFacturas_Click
 ' chkBoletas_Click

 End If
End If

End Sub

Private Sub grid_fac_DblClick()
Dim ww_desc As Currency
Dim FILAX As Integer
If SUT_LLAVE.EOF Then Exit Sub
If grid_fac.COL = 1 And PUB_TIPMOV = 20 And LK_EMP = "HER" Then
 grid_fac.COL = 1
 If grid_fac.CellBackColor = QBColor(7) Then
   grid_fac.CellBackColor = QBColor(15)
   grid_fac.TextMatrix(grid_fac.Row, 37) = 0
 Else
   grid_fac.TextMatrix(grid_fac.Row, 37) = 1
   grid_fac.CellBackColor = QBColor(7)
 End If
 Exit Sub
End If
If grid_fac.COL = 10 And Val(SUT_LLAVE!SUT_DESCTO) <> 0 Then
   ww_desc = 0
   pub_montodescto = Val(i_neto.Text) * Val(SUT_LLAVE!SUT_DESCTO) / 100
   fila = 2
   Do Until fila = grid_fac.Rows
   If PUB_TIPMOV = 10 And grid_fac.TextMatrix(1, 7) > 0 Then
      grid_fac.TextMatrix(fila, 10) = Format((pub_montodescto * grid_fac.TextMatrix(fila, 7) / Val(grid_fac.TextMatrix(1, 7))), "######.00")
      ww_desc = Val(grid_fac.TextMatrix(fila, 10)) + ww_desc
      If Val(grid_fac.TextMatrix(fila, 10)) <> 0 Then FILAX = fila
   End If
   fila = fila + 1
   Loop
   ww_desc = ww_desc - pub_montodescto
   grid_fac.TextMatrix(FILAX, 10) = Val(grid_fac.TextMatrix(FILAX, 10)) + ww_desc
   calcula_totales
   Exit Sub
End If
If grid_fac.COL = 10 And (PUB_TIPMOV = 10 Or PUB_TIPMOV = 93) And LK_EMP <> "3AA" Then
 If Trim(grid_fac.TextMatrix(grid_fac.Row, 0)) = "" Then
  Exit Sub
 End If
 If Val(grid_fac.TextMatrix(grid_fac.Row, 10)) <> 0 Then
    grid_fac.TextMatrix(grid_fac.Row, 10) = ""
 Else
    grid_fac.TextMatrix(grid_fac.Row, 10) = "D"
 End If
 calcula_totales
End If
Exit Sub
End Sub

Private Sub grid_fac_EnterCell()
If flag_del = 1 Then Exit Sub
If SUT_LLAVE.EOF Then Exit Sub
If LK_CODTRA = 1405 And grid_fac.COL <> 4 Then Exit Sub
If LK_CODTRA = 2414 And grid_fac.Row > 3 Then Exit Sub

If LK_CODTRA = 2414 And grid_fac.Row > 3 Then Exit Sub
If Val(grid_fac.TextMatrix(grid_fac.Row, 22)) <> 0 And LK_CODTRA = 2401 Then Exit Sub
If LK_CODTRA = 2407 Then Exit Sub
If grid_canje.Visible = True Then
   grid_canje.Visible = False
   textovar_canje.Visible = False
End If

If LK_CODTRA = 2414 Then
   If grid_fac.Row = 2 Then
      grid_fac.TextMatrix(2, 12) = "-1 Salida"
   Else
      grid_fac.TextMatrix(3, 12) = "+1 Ingreso"
   End If
End If

If grid_fac.COL = 12 Then
   If LK_CODTRA = 2414 Then Exit Sub
   'grid_fac.Text = i_is.Text
   procesa_celdas
   Exit Sub
End If

i_unidades.Visible = False
i_precios.Visible = False
i_mortal.Visible = False
i_is.Visible = False
If LK_CODTRA <> 2415 Then i_is.Visible = False

TEXTOVAR.Visible = False
If Trim(grid_fac.TextMatrix(grid_fac.Row, 0)) <> "" Then
   label_nomart.Caption = grid_fac.TextMatrix(grid_fac.Row, 0)
   If Val(grid_fac.TextMatrix(grid_fac.Row, 14)) <> 0 Then
    label_precio.Caption = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 13)) / Val(grid_fac.TextMatrix(grid_fac.Row, 14)), "0.000")
   End If
   label_nomart.Visible = True
   label_precio.Visible = True
Else
   label_nomart.Visible = False
   label_precio.Visible = False
End If
If FORMGEN.grid_fac.Text = "" Or (FORMGEN.grid_fac.COL = 6 And (PUB_TIPMOV = 20)) Or (SUT_LLAVE!SUT_PRECIO = "2" And PUB_TIPMOV = 10) Then
   procesa_celdas
End If
End Sub
Private Sub grid_fac_KeyPress(KeyAscii As Integer)
If LK_CODTRA = 2414 And grid_fac.Row > 3 Then Exit Sub
flag_salto = 1

'If pub_signo_ped = -1 And grid_fac.Col <> 4 Then Exit Sub


If KeyAscii <> 13 Then Exit Sub

If LK_CODTRA = 1405 And grid_fac.COL <> 4 Then Exit Sub
If grid_fac.COL = 6 Then
  If LK_CODTRA <> 2412 Then
      If LK_PRECIO = "C" Then Exit Sub
  End If
End If

If Val(grid_fac.TextMatrix(grid_fac.Row, 22)) <> 0 And LK_CODTRA = 2401 Then Exit Sub
flag_salto = 0
If grid_fac.COL = 5 Or grid_fac.COL = 6 Then
Else
procesa_celdas
Exit Sub
End If

SQ_OPER = 1
PUB_KEY = Nulo_Valor0(grid_fac.TextMatrix(grid_fac.Row, 16))
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
If art_LLAVE.EOF Then Exit Sub

If grid_fac.COL = 5 Then
  If LK_CODTRA <> 2414 Then
    grid_fac.TextMatrix(grid_fac.Row, 6) = ""
    verifica_unidades
    procesa_celdas
  End If
Else
   PUB_SECUEN = Val(Mid(grid_fac.TextMatrix(grid_fac.Row, 5), 23, 3))
   verifica_precios
   procesa_celdas
End If

End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo SALE
Dim det_suc As rdoResultset
Dim fx As Integer
Dim dflag As Integer

If KeyCode = 113 Then
 grid_fac_KeyPress 13
 Exit Sub
End If
If KeyCode = 119 Then
  If Val(grid_fac.TextMatrix(grid_fac.Row, 16)) = 0 Then
     MsgBox "Seleccione un codigo de Articulo", 48, Pub_Titulo
     Exit Sub
  End If
    PUB_KEY = Val(grid_fac.TextMatrix(grid_fac.Row, 16))
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = PUB_KEY
    LEER_PRE_LLAVE
    If pre_mayor.EOF Then GoTo PASA_PRE
  Load frmprecios
  frmprecios.unidad.Clear
  Do Until pre_mayor.EOF
     frmprecios.unidad.AddItem pre_mayor!pre_UNIDAD & String(80, " ") & pre_mayor!PRE_EQUIV
     pre_mayor.MoveNext
  Loop
  frmprecios.unidad.ListIndex = 0
  frmprecios.lblnombre.Caption = grid_fac.TextMatrix(grid_fac.Row, 0)
  frmprecios.Show 1
  loc_secuencia = 0
  If Val(frmprecios.txtprecio.Text) = 0 And Val(Left(frmprecios.cmdopcion.Text, 1)) = 2 Then GoTo ACT_FAL
  If Val(frmprecios.txtprecio.Text) <> 0 Then
ACT_FAL:
    PSNEW_PRECIO(0) = LK_CODCIA
    PSNEW_PRECIO(1) = LK_FECHA_DIA
    loc_new_llave.Requery
    If loc_new_llave.EOF Then
      loc_secuencia = 1
    Else
      loc_secuencia = Val(loc_new_llave!hpr_numoper) + 1
    End If
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = PUB_KEY
    PUB_SECUEN = frmprecios.unidad.ListIndex
    LEER_PRE_LLAVE
    If pre_llave.EOF Then GoTo PASA_PRE
    
    loc_new_llave.AddNew
    loc_new_llave!hpr_codcia = LK_CODCIA
    loc_new_llave!hpr_FECHA = LK_FECHA_DIA
    loc_new_llave!hpr_numoper = loc_secuencia
    loc_new_llave!hpr_codart = PUB_KEY
    loc_new_llave!hpr_TIPMOV = Val(Left(frmprecios.cmdopcion.Text, 1))
    loc_new_llave!hpr_precio = Val(frmprecios.txtprecio.Text)
    loc_new_llave!hpr_empresa = Trim(frmprecios.txtprove.Text)
    loc_new_llave!hpr_codusu = LK_CODUSU
    loc_new_llave!hpr_unidad = Left(frmprecios.unidad.Text, 20)
    loc_new_llave!hpr_EQUIV = Val(Right(frmprecios.unidad.Text, 12))
    loc_new_llave!hpr_pre1 = pre_llave!PRE_PRE1
    loc_new_llave!hpr_pre2 = pre_llave!PRE_PRE2
    loc_new_llave!hpr_pre3 = pre_llave!PRE_PRE3
    loc_new_llave!hpr_pre4 = pre_llave!PRE_PRE4
    loc_new_llave!hpr_pre5 = pre_llave!PRE_PRE5
    loc_new_llave.Update
    
   pub_cadena = "DELETE FROM CONSULART WHERE CON_CODART = " & PUB_KEY & " AND CON_CODCIA = '" & LK_CODCIA & "'"
   CN.Execute pub_cadena, rdExecDirect
    
  End If
  
PASA_PRE:
  Unload frmprecios
End If
If KeyCode = 45 And grid_fac.COL = 1 Then
   grid_fac.Rows = grid_fac.Rows + 1
   Exit Sub
End If
If KeyCode = 45 And grid_fac.COL = 4 Then
  CONFIGAR_LOTES LK_CODCIA, Val(grid_fac.TextMatrix(grid_fac.Row, 16)), Val(grid_fac.TextMatrix(grid_fac.Row, 4)), Left(grid_fac.TextMatrix(grid_fac.Row, 5), 10), Val(Right(grid_fac.TextMatrix(grid_fac.Row, 5), 8)), grid_fac.Row
  PASA_BOT_ACEPTAR
  Exit Sub
End If

If KeyCode = 32 And grid_fac.TextMatrix(grid_fac.Row, 1) <> "" Then
  If grid_fac.Rows = 3 Then Exit Sub
   grid_fac.COL = 1
   If grid_fac.CellBackColor = vbBlue Then
      grid_fac.CellBackColor = vbWhite
   Else
      grid_fac.CellBackColor = vbBlue
   End If
End If


If KeyCode <> 46 Then Exit Sub
   pub_mensaje = " ¿Desea Eliminar las filas seleccionadas ... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then Exit Sub
   
dflag = 0
OTRO:
fx = grid_fac.Row
FILAX = 2
'grid_fac.COL = 1
flag_del = 1
Do Until FILAX > grid_fac.Rows - 1
   grid_fac.Row = FILAX
   If grid_fac.CellBackColor = vbBlue And FILAX >= 2 And grid_fac.Rows <> 3 Then
      BORRA_ITEM_LOTE Val(grid_fac.TextMatrix(grid_fac.Row, 16)), grid_fac.Row
      grid_fac.RemoveItem (grid_fac.Row)
      dflag = 1
      fx = 0
      grid_fac.Refresh
      Exit Do
   End If
   FILAX = FILAX + 1
Loop

If fx = 0 Then GoTo OTRO

If dflag = 0 And grid_fac.Rows > 3 Then
   grid_fac.Row = fx
   BORRA_ITEM_LOTE Val(grid_fac.TextMatrix(grid_fac.Row, 16)), grid_fac.Row
   grid_fac.RemoveItem (grid_fac.Row)
End If

If grid_fac.Rows = 3 Then
 label_nomart.Caption = ""
 label_precio.Caption = ""
End If
flag_del = 0

grid_fac.Refresh
calcula_totales
grid_fac.SetFocus
Exit Sub
SALE:
  MsgBox Err.Description, 48, Pub_Titulo
End Sub

Private Sub grid_fac_LeaveCell()
label_nomart.Caption = ""
label_precio.Caption = ""
End Sub

Private Sub grid_fac_LostFocus()
'i_mortal.Visible = False
'textovar.Visible = False
End Sub

Private Sub grid_fac_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   Dim FILAS
   If Trim(SUT_LLAVE!SUT_abreviado) = "CONS" Then
      FILAS = 2
      Do Until FILAS = 999
      grid_fac.TextMatrix(FILAS, 6) = 0
      FILAS = FILAS + 1
      If FILAS = grid_fac.Rows Then
         FILAS = 999
      End If
   Loop
   calcula_totales
   End If
   pub_mensaje = "Colocar papel para la Impresora "
   Pub_Respuesta = MsgBox(pub_mensaje, vbInformation + vbOKCancel, Pub_Titulo)
   If Pub_Respuesta = 1 Then
      IMP_GRID
      Exit Sub
   End If

End If

End Sub

Private Sub grid_fac_Scroll()
TEXTOVAR.Visible = False
i_mortal.Visible = False
i_is.Visible = False
i_precios.Visible = False
i_unidades.Visible = False
End Sub
Private Sub grid_trans_KeyPress(KeyAscii As Integer)
Dim longit As String * 15
If KeyAscii <> 13 Then Exit Sub
If LK_CODTRA = 1405 Then GoTo JALA_REC
If Left(grid_trans.TextMatrix(grid_trans.Row, 5), 1) = "R" Then
   MsgBox "Ya Rececpcionado...", 48, Pub_Titulo
   grid_trans.SetFocus
   Exit Sub
End If

FAR_TRANS.MoveFirst
fila = 2
Do Until FAR_TRANS.EOF
   
   If Val(FAR_TRANS!far_numfac) = Val(grid_trans.TextMatrix(grid_trans.Row, 0)) And Val(FAR_TRANS!far_numser) = Val(grid_trans.TextMatrix(grid_trans.Row, 7)) Then
   Else
   GoTo OTRO
   End If
   
   If FAR_TRANS!far_transito <> "T" Then GoTo OTRO
   
   
   If Val(FAR_TRANS!FAR_CODCIA) <> Val(grid_trans.TextMatrix(grid_trans.Row, 6)) Then GoTo OTRO
        
        
   grid_fac.Rows = fila + 2
   PUB_KEY = FAR_TRANS!far_codart
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ART_LLAVE
   
   SQ_OPER = 1
   PUB_CODART = PUB_KEY
   pu_codcia = LK_CODCIA
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
      MsgBox "Datos de Articulos ...errados..Revisar"
      Exit Sub
   End If
   PUB_CODART = FAR_TRANS!far_codart
   pu_codcia = LK_CODCIA
   PUB_SECUEN = FAR_TRANS!far_num_precio
   SQ_OPER = 1
   LEER_PRE_LLAVE
   grid_fac.TextMatrix(fila, 34) = redondea(pre_llave!pre_PESO * FAR_TRANS!far_cantidad)
   grid_fac.TextMatrix(fila, 37) = pre_llave!pre_PESO
   grid_fac.TextMatrix(fila, 43) = Nulo_Valor0(pre_llave!PRE_LITRO)
   i_codcli.Text = FAR_TRANS!far_codclie
'   i_concepto.text = FAR_TRANS!far_concepto
   grid_fac.TextMatrix(fila, 0) = art_LLAVE!art_nombre
   grid_fac.TextMatrix(fila, 1) = FAR_TRANS!far_codart
   grid_fac.TextMatrix(fila, 16) = FAR_TRANS!far_codart
   grid_fac.TextMatrix(fila, 2) = FAR_TRANS!far_JABAS
   grid_fac.TextMatrix(fila, 3) = FAR_TRANS!far_UNIDADES
   grid_fac.TextMatrix(fila, 11) = FAR_TRANS!FAR_COSPRO
   If Nulo_Valor0(FAR_TRANS!FAR_equiv) > 0 Then
   grid_fac.TextMatrix(fila, 4) = FAR_TRANS!far_cantidad / Nulo_Valor0(FAR_TRANS!FAR_equiv)
   Else
   grid_fac.TextMatrix(fila, 4) = FAR_TRANS!far_cantidad
   End If
'   If FAR_TRANS!far_codart = 498689 Then Stop
   If SUT_LLAVE!SUT_PRECIO <> "4" Then grid_fac.TextMatrix(fila, 6) = FAR_TRANS!FAR_COSPRO * Nulo_Valor0(FAR_TRANS!FAR_equiv)
   longit = Nulo_Valors(FAR_TRANS!far_descri)
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(FAR_TRANS!FAR_equiv)
   grid_fac.TextMatrix(fila, 5) = longit & "       " & FAR_TRANS!FAR_ORDEN_UNIDADES & "       " & FAR_TRANS!FAR_equiv
   grid_fac.TextMatrix(fila, 6) = Nulo_Valors(FAR_TRANS!FAR_COSPRO) * Nulo_Valor0(FAR_TRANS!FAR_equiv)
  
   grid_fac.TextMatrix(fila, 8) = longit & "       " & FAR_TRANS!FAR_ORDEN_UNIDADES & "       " & FAR_TRANS!FAR_equiv
   grid_fac.TextMatrix(fila, 10) = FAR_TRANS!FAR_DESCTO
   grid_fac.TextMatrix(fila, 12) = SUT_LLAVE!SUT_SIGNO_ARM
   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   grid_fac.TextMatrix(fila, 26) = FAR_TRANS!FAR_TIPMOV
   grid_fac.TextMatrix(fila, 27) = FAR_TRANS!FAR_CODCIA
   grid_fac.TextMatrix(fila, 28) = FAR_TRANS!far_numser
   grid_fac.TextMatrix(fila, 29) = FAR_TRANS!FAR_FBG
   grid_fac.TextMatrix(fila, 30) = FAR_TRANS!far_numfac
   grid_fac.TextMatrix(fila, 31) = FAR_TRANS!FAR_NUMSEC
   grid_fac.TextMatrix(fila, 33) = FAR_TRANS!far_codart
   
    fila = fila + 1
OTRO:
   FAR_TRANS.MoveNext
Loop
grid_trans.Visible = False
label_grid.Visible = False
calcula_totales
If grid_fac.Visible Then grid_fac.SetFocus
Screen.MousePointer = 0

Exit Sub
' ***  JALA RECEPCION DE MERCADERIA ***
JALA_REC:

FAR_TRANS.MoveFirst
grid_trans.Visible = False

fila = grid_fac.Rows - 1
Do Until FAR_TRANS.EOF
   If Val(FAR_TRANS!FAR_NUMFAC_C) = Val(grid_trans.TextMatrix(grid_trans.Row, 1)) And Val(FAR_TRANS!FAR_NUMSER_C) = Val(grid_trans.TextMatrix(grid_trans.Row, 0)) Then
   Else
    GoTo otro2
   End If
   
   If FAR_TRANS!far_transito <> "P" Then GoTo otro2
        
   grid_fac.Rows = fila + 2
   PUB_KEY = FAR_TRANS!far_codart
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ART_LLAVE
   
   SQ_OPER = 1
   PUB_CODART = PUB_KEY
   pu_codcia = LK_CODCIA
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
      MsgBox "Datos de Articulos ...errados..Revisar"
      Exit Sub
   End If
   PUB_CODART = FAR_TRANS!far_codart
   pu_codcia = LK_CODCIA
   PUB_SECUEN = FAR_TRANS!far_num_precio
   SQ_OPER = 1
   LEER_PRE_LLAVE
   grid_fac.TextMatrix(fila, 34) = redondea(pre_llave!pre_PESO * (FAR_TRANS!far_cantidad - FAR_TRANS!far_cantidad_p))
   grid_fac.TextMatrix(fila, 37) = pre_llave!pre_PESO
   grid_fac.TextMatrix(fila, 43) = Nulo_Valor0(pre_llave!PRE_LITRO)
         
   i_codcli.Text = FAR_TRANS!far_codclie
'   i_concepto.text = FAR_TRANS!far_concepto
   grid_fac.TextMatrix(fila, 0) = art_LLAVE!art_nombre
   grid_fac.TextMatrix(fila, 1) = FAR_TRANS!far_codart
   grid_fac.TextMatrix(fila, 16) = FAR_TRANS!far_codart
   grid_fac.TextMatrix(fila, 2) = FAR_TRANS!far_JABAS
   grid_fac.TextMatrix(fila, 3) = FAR_TRANS!far_UNIDADES
   grid_fac.TextMatrix(fila, 11) = FAR_TRANS!FAR_COSPRO
   If Nulo_Valor0(FAR_TRANS!FAR_equiv) > 0 Then
   grid_fac.TextMatrix(fila, 4) = (FAR_TRANS!far_cantidad - FAR_TRANS!far_cantidad_p) / Nulo_Valor0(FAR_TRANS!FAR_equiv)
   Else
   grid_fac.TextMatrix(fila, 4) = (FAR_TRANS!far_cantidad - FAR_TRANS!far_cantidad_p)
   End If
   
   If SUT_LLAVE!SUT_PRECIO <> "4" Then grid_fac.TextMatrix(fila, 6) = FAR_TRANS!FAR_COSPRO
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(FAR_TRANS!FAR_equiv)
   grid_fac.TextMatrix(fila, 5) = Nulo_Valors(FAR_TRANS!far_descri)
   grid_fac.TextMatrix(fila, 6) = Nulo_Valors(FAR_TRANS!FAR_PRECIO)
  
   grid_fac.TextMatrix(fila, 8) = Nulo_Valor0(FAR_TRANS!FAR_FLETE)
   grid_fac.TextMatrix(fila, 10) = FAR_TRANS!FAR_DESCTO
   grid_fac.TextMatrix(fila, 12) = SUT_LLAVE!SUT_SIGNO_ARM
   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   grid_fac.TextMatrix(fila, 26) = FAR_TRANS!FAR_TIPMOV
   grid_fac.TextMatrix(fila, 27) = FAR_TRANS!FAR_CODCIA
   grid_fac.TextMatrix(fila, 28) = FAR_TRANS!far_numser
   grid_fac.TextMatrix(fila, 29) = FAR_TRANS!FAR_FBG
   grid_fac.TextMatrix(fila, 30) = FAR_TRANS!far_numfac
   grid_fac.TextMatrix(fila, 31) = FAR_TRANS!FAR_NUMSEC
   grid_fac.TextMatrix(fila, 33) = FAR_TRANS!far_codart
   
   grid_fac.TextMatrix(fila, 39) = Trim(grid_trans.TextMatrix(grid_trans.Row, 0))
   grid_fac.TextMatrix(fila, 40) = Trim(grid_trans.TextMatrix(grid_trans.Row, 1))
   grid_fac.TextMatrix(fila, 41) = FAR_TRANS!FAR_NUMSEC   ') Trim(grid_trans.TextMatrix(grid_trans.Row, 8))
   
   
   
    fila = fila + 1
otro2:
   FAR_TRANS.MoveNext
Loop
grid_trans.Visible = False
label_grid.Visible = False
calcula_totales
If grid_fac.Visible Then grid_fac.SetFocus
Screen.MousePointer = 0


End Sub

Private Sub gridl_DblClick()
Dim FILA_aCT As Integer


If LK_CODTRA <> 2702 Then Exit Sub


If pub_signo_car = 0 Then
   i_serdoc.Text = Val(gridl.TextMatrix(gridl.Row, 11))
   i_numdoc.Text = Val(gridl.TextMatrix(gridl.Row, 12))
   PUB_TIPDOC = gridl.TextMatrix(gridl.Row, 13)
   i_numguia.Text = gridl.TextMatrix(gridl.Row, 16)
   For FILAX = 0 To i_situacion.ListCount - 1
       i_situacion.ListIndex = FILAX
       If Trim(gridl.TextMatrix(gridl.Row, 17)) = Left(i_situacion.Text, 1) Then
           i_situacion_ant = Left(i_situacion.Text, 20)
          Exit For
       End If
   Next FILAX

   End If
   
   
   FILA_aCT = gridl.Row
   If pub_signo_car = 0 Then
   FILAX = 2
   Do Until FILAX > gridl.Rows - 1
   If FILA_aCT <> FILAX Then
      gridl.RowHeight(FILAX) = 0
   End If
   FILAX = FILAX + 1
   Loop
   End If

End Sub

Private Sub gridl_EnterCell()
'If LK_CODTRA = 2774 Then
'   If gridl.Col <> 5 Then Exit Sub
'End If
If LK_CODTRA = 1455 Then Exit Sub

'TEXTOVAR_CANJE.Visible = False
If gridl.COL = 21 Then Exit Sub
If gridl.COL = 5 Or gridl.COL = 6 Or gridl.COL = 7 Or gridl.COL >= 17 Then
   PROCESA_CELDASL
Else
   textovarl.Visible = False
End If

End Sub

Private Sub gridl_KeyPress(KeyAscii As Integer)


If KeyAscii <> 13 Then Exit Sub

 If gridl.COL = 21 Then
    i_bancos.Left = gridl.Left + gridl.CellLeft
    i_bancos.Width = gridl.CellWidth
    i_bancos.Top = gridl.Top + gridl.CellTop
    textovar_canje.Width = 30
    textovar_canje.Height = 30
    i_bancos.Visible = True
    i_bancos_Click
Else
If gridl.COL = 5 Or gridl.COL = 8 Or gridl.COL = 6 Or gridl.COL = 7 Or gridl.COL >= 17 Then PROCESA_CELDASL
End If



End Sub

Private Sub gridl_KeyUp(KeyCode As Integer, Shift As Integer)
Dim FILAX, fx, FILA_aCT As Integer





If KeyCode = 46 Then
If gridl.Rows <= 3 Then
Else
   gridl.RemoveItem (gridl.Row)
   gridl.Refresh
   calcula_totales2
   Exit Sub
 End If
End If


   


    
End Sub

Private Sub gridl_LeaveCell()
' If textovarL.Enabled = True And textovarL.Visible = True Then
'   textovarL.SetFocus
'   textovarL.SelStart = 0
'   textovarL.SelLength = Len(textovarL)
' End If
' gridl.CellBackColor = QBColor(15)
End Sub

Private Sub gridl_Scroll()

textovarl.Height = 30
textovarl.Width = 30
i_bancos.Visible = False
End Sub




Private Sub gridlt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then
  'ASIGNA_NEW_LOTE Val(gridlt.TextMatrix(gridlt.Row, 9)), "0.000", Val(gridlt.TextMatrix(gridlt.Row, 10))
  ASIGNA_NEW_LOTE Val(gridlt.TextMatrix(gridlt.Row, 9)), "0.00", Trim(gridlt.TextMatrix(gridlt.Row, 1)), Val(gridlt.TextMatrix(gridlt.Row, 11)), Val(gridlt.TextMatrix(gridlt.Row, 10))
  ' CONFIGAR_LOTES LK_CODCIA, Val(grid_fac.TextMatrix(grid_fac.Row, 16)), Val(grid_fac.TextMatrix(grid_fac.Row, 4)), grid_fac.TextMatrix(grid_fac.Row, 5), grid_fac.Row
  CONFIGAR_LOTES LK_CODCIA, Val(grid_fac.TextMatrix(grid_fac.Row, 16)), Val(grid_fac.TextMatrix(grid_fac.Row, 4)), Left(grid_fac.TextMatrix(grid_fac.Row, 5), 10), Val(Right(grid_fac.TextMatrix(grid_fac.Row, 5), 8)), grid_fac.Row
End If
If KeyCode = 46 Then
  If gridlt.Rows > 3 Then
   BORRA_ITEM_LOTE Val(grid_fac.TextMatrix(grid_fac.Row, 16)), grid_fac.Row, gridlt.TextMatrix(gridlt.Row, 0)
   gridlt.RemoveItem (gridlt.Row)
   gridlt.Row = gridlt.Row
  End If
End If
   
End Sub

Private Sub gridlt_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub

'If Trim(gridigv.TextMatrix(gridigv.Row, 9)) <> "8" Then
'  If Trim(gridigv.TextMatrix(gridigv.Row, 0)) = "" Then Exit Sub
'  If Trim(gridigv.TextMatrix(gridigv.Row, 1)) <> "" And gridigv.Col = 2 Or gridigv.Col = 3 Then GoTo leer
'  If Trim(gridigv.TextMatrix(gridigv.Row, 8)) <> "0" Then Exit Sub
'End If


'If gridlt.COL = 0 Then Exit Sub ' a execiopn de un ingreso nuevo
If gridlt.COL = 1 Then Exit Sub
If gridlt.COL = 2 Then Exit Sub
'   a = Val(gridigv.TextMatrix(gridigv.Row - 1, 0))
'   a = a + 1
'  gridigv.TextMatrix(gridigv.Row, 0) = a

'If WMODO = "I" Or WMODO = "C" Then
    textolote.Left = gridlt.Left + gridlt.CellLeft
    textolote.Width = gridlt.CellWidth
    textolote.Height = gridlt.CellHeight
    textolote.Top = gridlt.Top + gridlt.CellTop
    textolote.Text = gridlt.TextMatrix(gridlt.Row, gridlt.COL)
    textolote.Visible = True
    Azul textolote, textolote
    textolote.SetFocus
'End If

End Sub

Private Sub i_bancos_Click()
If LK_CODTRA = 2770 Or LK_CODTRA = 2725 Then
   gridl.Text = i_bancos.Text
End If
grid_canje.Text = i_bancos.Text
End Sub

Private Sub i_bancos_GotFocus()
i_bancos.ListIndex = -1
End Sub

Private Sub i_bancos_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If LK_CODTRA = 2770 Or LK_CODTRA = 2725 Then
   If gridl.Visible = True Then gridl.COL = 22
End If

If grid_canje.Visible = True Then grid_canje.COL = 6

End Sub

Private Sub i_bruto2_GotFocus()
Azul i_bruto2, i_bruto2
End Sub

Private Sub i_bruto2_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_bruto2, KeyAscii

If KeyAscii <> 13 Then Exit Sub

NUMERO = FORMGEN.i_bruto2.WhatsThisHelpID

avanza_campo

End Sub

Private Sub i_bruto2_LostFocus()
Dim wBruto As Currency
If LK_CODTRA = 2750 Then
  i_importe_amort.Text = Val(i_bruto2.Text) + Val(i_impto2.Text)
  Exit Sub
End If
If LK_CODTRA = 2748 Then
' i_importe_amort.Text = Val(Val(i_impto2.Text) + Val(i_bruto2.Text) + Val(i_importe.Text))
' wbruto = i_bruto2.Text
'i_importe_amort.Text = redondea(Val(i_bruto2.Text) * (1 + LK_IGV / 100))
' i_importe_amort.Text = i_importe_amort.Text + Val(i_importe.Text)
  i_impto2.Text = redondea(Val(i_bruto2.Text) * (LK_IGV / 100))
  i_importe.Text = redondea(Abs(Val(i_bruto2.Text) + Val(i_impto2.Text) - Val(i_importe_amort.Text)))
Else
 i_importe_amort.Text = Val(Val(i_impto2.Text) + Val(i_bruto2.Text))
End If

End Sub

Private Sub i_camal_GotFocus()
If cli_llave.EOF = False Then
   i_camal.Text = cli_llave!CLI_CASA2
End If

End Sub

Private Sub i_camal_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   Exit Sub
End If
NUMERO = FORMGEN.i_camal.WhatsThisHelpID
avanza_campo

End Sub

Private Sub i_cambio_Click()
If i_cambio.Value = 1 Then
   i_numfac.Locked = False
   i_numser.Locked = False
   i_numfac.BackColor = QBColor(15)
   i_numser.BackColor = QBColor(15)
   If PUB_PROCESO <> 1 Then
       Azul i_numfac, i_numfac
   End If
Else
   BolFac_Click
End If

End Sub

Private Sub i_cantidad_GotFocus()
Azul i_cantidad, i_cantidad

End Sub

Private Sub i_cantidad_LostFocus()
If Val(i_precio.Text) <> 0 Then
i_importe_amort.Text = Val(i_cantidad.Text) * Val(i_precio.Text)
End If

If LK_CODTRA = 2499 Then
WS_BRUTO = i_cantidad.Text * i_precio.Text
i_subtotal = WS_BRUTO
i_neto.Text = WS_BRUTO
PUB_SUBTOTAL = WS_BRUTO
End If

End Sub

Private Sub i_cias_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   Exit Sub
End If
NUMERO = FORMGEN.i_cias.WhatsThisHelpID
avanza_campo

End Sub

Private Sub i_cias_LostFocus()
'YO AGREGO
'If LK_CODTRA = 2409 And Right(i_cias.text, 2) = "00" Then
'  If par_llave!par_flag_costos <> 9 Then
'    MsgBox "!!! Compañia ... debe Costear Productos antes de Enviar ", 48, Pub_Titulo
'    Exit Sub
'  End If
'End If
End Sub

Private Sub i_codart_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then Exit Sub
NUMERO = FORMGEN.i_codart.WhatsThisHelpID

avanza_campo
fin:

End Sub

Private Sub i_codban_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not LV_CCM.Visible Then
 Exit Sub
End If
If KeyCode = 13 Then Exit Sub
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codcli.Text = "" Then
  loc_key = 1
  Set LV_CCM.SelectedItem = LV_CCM.ListItems(loc_key)
'  LV_ccm.Visible = False
  LV_CCM.ListItems.Item(loc_key).Selected = True
  LV_CCM.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_CCM.ListItems.count Then loc_key = LV_CCM.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_CCM.ListItems.count Then loc_key = LV_CCM.ListItems.count
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
  LV_CCM.ListItems.Item(loc_key).Selected = True
  LV_CCM.ListItems.Item(loc_key).EnsureVisible
  i_codban.Text = Trim(LV_CCM.ListItems.Item(loc_key).Text) & " "
  DoEvents
  i_codban.SelStart = Len(i_codban.Text)
  DoEvents
fin:

End Sub

Private Sub i_codban_KeyPress(KeyAscii As Integer)
On Error GoTo SALE
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.

If KeyAscii = 27 Then
   i_codban.Text = ""
   i_nomban.Caption = ""
   i_codban.SetFocus
   GoTo fin
End If

If KeyAscii <> 13 Then
   GoTo fin
End If
If IsNumeric(FORMGEN.i_codban.Text) Then
PUB_CODBAN = Val(FORMGEN.i_codban.Text)
Else
 PUB_CODBAN = Trim(LV_CCM.ListItems.Item(loc_key).SubItems(1)) 'Val(FORMGEN.i_codban.text)
 FORMGEN.i_codban.Text = PUB_CODBAN
End If
If Len(i_codban.Text) = 0 Then
   Exit Sub
End If
If PUB_CODBAN <> 0 Then
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CCM_LLAVE
   If ccm_llave.EOF Then
    Azul FORMGEN.i_codban, FORMGEN.i_codban
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    FORMGEN.i_codban.SetFocus
    FORMGEN.i_nomban.Caption = ""
    GoTo fin
   Else
    NUMERO = FORMGEN.i_codban.WhatsThisHelpID
    If LMONEDAB.Visible Then
      If ccm_llave!CCM_MONEDA = "S" Then
        LMONEDAB.Caption = "S/."
      Else
        LMONEDAB.Caption = "$."
      End If
    End If
    avanza_campo
    FORMGEN.i_nomban.Caption = ccm_llave(2)
   End If
Else
 On Error GoTo VERI
   If loc_key <> 0 Then VALOR = UCase(LV_CCM.ListItems.Item(loc_key).Text)
   If Trim(UCase(i_codban.Text)) = Left(VALOR, Len(Trim(i_codban.Text))) Then
   Else
      Exit Sub
   End If
   If loc_key = 0 Then Exit Sub
   i_codban.Text = Trim(LV_CCM.ListItems.Item(loc_key).SubItems(1))
   PUB_CODBAN = Val(FORMGEN.i_codban.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CCM_LLAVE
   If LMONEDAB.Visible Then
      If ccm_llave!CCM_MONEDA = "S" Then
        LMONEDAB.Caption = "S/."
      Else
        LMONEDAB.Caption = "$."
      End If
   End If
   FORMGEN.i_nomban.Caption = ccm_llave(2)
   NUMERO = FORMGEN.i_codban.WhatsThisHelpID
   avanza_campo

End If
LV_CCM.Visible = False


Exit Sub
VERI:
fin:

Exit Sub
SALE:
Resume Next

End Sub



Private Sub i_codban_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(i_codban.Text) = 0 Then
   LV_CCM.Visible = False
   Exit Sub
End If
If LV_CCM.Visible = False Then
    var = Asc(i_codban.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 9
    archi = "SELECT * FROM CCMAEST WHERE CCM_CODBAN > 1 AND CCM_CODCIA = '" & par_llave!PAR_CIACCM & "' AND CCM_NOMBRE BETWEEN '" & i_codban.Text & "' AND  '" & var & "' ORDER BY CCM_NOMBRE"
    PROC_LISVIEW LV_CCM
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If LV_CCM.Visible Then
  Set itmFound = LV_CCM.FindItem(LTrim(i_codban.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_CCM.ListItems.count Then
      LV_CCM.ListItems.Item(LV_CCM.ListItems.count).EnsureVisible
   Else
     LV_CCM.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  
  Exit Sub
End If

End Sub

Private Sub i_codban_GotFocus()
Azul i_codban, i_codban
loc_key = 0
If pub_signo_ccm = 0 And LK_CODTRA <> 5715 And LK_CODTRA <> 231 Then
   NUMERO = FORMGEN.i_codban.WhatsThisHelpID
   avanza_campo
End If

End Sub

Private Sub i_codban_LostFocus()
WS_MONEDA_CCM = ""
If LK_CODTRA = 5714 Then
If ccm_llave.EOF = False And Val(i_codban.Text) > 0 Then
   If ccm_llave!CCM_CODBAN = Val(i_codban.Text) Then
     WS_MONEDA_CCM = ccm_llave!CCM_MONEDA
      If ccm_llave!CCM_MONEDA = "D" Then
         lmoneda.Caption = "S/."
      Else
         lmoneda.Caption = " $"
      End If
      If LK_CODTRA = 2738 Then
        WS_MONEDA_CLI = ccm_llave!CCM_MONEDA
      End If
   End If
End If
Exit Sub
End If

If ccm_llave.EOF = False And Val(i_codban.Text) > 0 Then
   If ccm_llave!CCM_CODBAN = Val(i_codban.Text) Then
     WS_MONEDA_CCM = ccm_llave!CCM_MONEDA
      If ccm_llave!CCM_MONEDA = "D" Then
         lmoneda.Caption = " $"
      Else
         lmoneda.Caption = "S/."
      End If
      If LK_CODTRA = 2738 Then
        WS_MONEDA_CLI = ccm_llave!CCM_MONEDA
      End If
   End If
End If

End Sub
Private Sub i_codban2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not LV_CCM.Visible Then
 Exit Sub
End If
If KeyCode = 13 Then Exit Sub
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codcli.Text = "" Then
  loc_key = 1
  Set LV_CCM.SelectedItem = LV_CCM.ListItems(loc_key)
'  LV_ccm.Visible = False
  LV_CCM.ListItems.Item(loc_key).Selected = True
  LV_CCM.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_CCM.ListItems.count Then loc_key = LV_CCM.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_CCM.ListItems.count Then loc_key = LV_CCM.ListItems.count
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
  LV_CCM.ListItems.Item(loc_key).Selected = True
  LV_CCM.ListItems.Item(loc_key).EnsureVisible
  i_codban2.Text = Trim(LV_CCM.ListItems.Item(loc_key).Text) & " "
  DoEvents
  i_codban2.SelStart = Len(i_codban2.Text)
  DoEvents
fin:

End Sub

Private Sub i_codban2_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.

If KeyAscii = 27 Then
   i_codban2.Text = ""
   i_nomban2.Caption = ""
   i_codban2.SetFocus
   GoTo fin
End If

If KeyAscii <> 13 Then
   GoTo fin
End If
PUB_CODBAN = Val(FORMGEN.i_codban2.Text)
If Len(i_codban2.Text) = 0 Then
   Exit Sub
End If
If PUB_CODBAN <> 0 Then
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CCM_LLAVE
   If ccm_llave.EOF Then
    Azul FORMGEN.i_codban2, FORMGEN.i_codban2
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    FORMGEN.i_codban2.SetFocus
    FORMGEN.i_nomban2.Caption = ""
    GoTo fin
   Else
    NUMERO = FORMGEN.i_codban2.WhatsThisHelpID
    avanza_campo
    FORMGEN.i_nomban2.Caption = ccm_llave(2)
   End If
Else
   If loc_key <> 0 Then VALOR = UCase(LV_CCM.ListItems.Item(loc_key).Text)
   If Trim(UCase(i_codban2.Text)) = Left(VALOR, Len(Trim(i_codban2.Text))) Then
   Else
      Exit Sub
   End If
   If loc_key = 0 Then Exit Sub
   i_codban2.Text = Trim(LV_CCM.ListItems.Item(loc_key).SubItems(1))
   PUB_CODBAN = Val(FORMGEN.i_codban2.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CCM_LLAVE
   FORMGEN.i_nomban2.Caption = ccm_llave(2)
   NUMERO = FORMGEN.i_codban2.WhatsThisHelpID
   avanza_campo

End If
LV_CCM.Visible = False



fin:
End Sub
Private Sub i_codban2_KeyUp(KeyCode As Integer, Shift As Integer)

Dim var
If Len(i_codban2.Text) = 0 Then
   LV_CCM.Visible = False
   Exit Sub
End If
If LV_CCM.Visible = False Then
    var = Asc(i_codban2.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 9
    archi = "SELECT * FROM CCMAEST WHERE CCM_CODBAN > 1 AND CCM_CODCIA = '" & LK_CODCIA & "' AND CCM_NOMBRE BETWEEN '" & i_codban.Text & "' AND  '" & var & "' ORDER BY CCM_NOMBRE"
    PROC_LISVIEW LV_CCM
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If LV_CCM.Visible Then
  Set itmFound = LV_CCM.FindItem(LTrim(i_codban2.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_CCM.ListItems.count Then
      LV_CCM.ListItems.Item(LV_CCM.ListItems.count).EnsureVisible
   Else
     LV_CCM.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  
  Exit Sub
End If

End Sub


Private Sub i_codban2_LostFocus()
If ccm_llave.EOF = False And Val(i_codban2.Text) > 0 Then
   If ccm_llave!CCM_CODBAN = Val(i_codban2.Text) Then
      If ccm_llave!CCM_MONEDA = "D" Then
         'Lmoneda(1).Caption = " $"
      Else
         'Lmoneda(1).Caption = "S/."
      End If
   End If
End If

End Sub

Private Sub i_codcli_Change()
If i_codcli.Text = "" Then
  If LK_CODTRA = 2401 Then
      i_condi.Clear
  End If
  FORMGEN.i_nomCLI.Caption = ""
  FORMGEN.LV_CLI.Visible = False
  BLOQ(0).Value = 0
  BLOQ(1).Value = 0
  BLOQ(2).Value = 0
  BLOQ(3).Value = 0
End If

End Sub

Private Sub i_codcli_GotFocus()
loc_key = 0
End Sub


Private Sub i_codcli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not LV_CLI.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codcli.Text = "" Then
  loc_key = 1
  Set LV_CLI.SelectedItem = LV_CLI.ListItems(loc_key)
'  LV_CLI.Visible = False
  LV_CLI.ListItems.Item(loc_key).Selected = True
  LV_CLI.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_CLI.ListItems.count Then loc_key = LV_CLI.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_CLI.ListItems.count Then loc_key = LV_CLI.ListItems.count
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
  LV_CLI.ListItems.Item(loc_key).Selected = True
  LV_CLI.ListItems.Item(loc_key).EnsureVisible
  i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
'  DoEvents
  i_codcli.SelStart = Len(i_codcli.Text)
'  DoEvents
fin:

End Sub

Private Sub i_codcli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(i_codcli.Text) = 0 Or IsNumeric(i_codcli.Text) Then
   LV_CLI.Visible = False
   Exit Sub
End If
If LV_CLI.Visible = False Or Len(Trim(i_codcli.Text)) = 1 Then
   loc_key = 0
    var = Asc(i_codcli.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 1
    archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = '" & PUB_CP & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & i_codcli.Text & "' AND  '" & var & "' ORDER BY CLI_NOMBRE"
    PROC_LISVIEW LV_CLI, 3000
    loc_key = 0
    If LV_CLI.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If LV_CLI.Visible Then
  Set itmFound = LV_CLI.FindItem(LTrim(i_codcli.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_CLI.ListItems.count Then
      LV_CLI.ListItems.Item(LV_CLI.ListItems.count).EnsureVisible
   Else
     LV_CLI.ListItems.Item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If


End Sub

Private Sub i_codcli_LostFocus()
Dim xcuenta As Integer
Dim defecto As Integer
Dim key_descto As String * 2
Dim tf As Integer
If Not IsNumeric(i_codcli.Text) Then
 Exit Sub
End If
Dim ww_num As Integer
If cli_llave.EOF = False And Val(i_codcli.Text) > 0 Then
   If cli_llave!CLI_CODCLIE <> Val(i_codcli.Text) Then
      i_codcli.Text = ""
      Exit Sub
   End If
End If

If cli_llave.EOF = False And Val(i_codcli.Text) > 0 Then
   ww_num = Val(Nulo_Valors(cli_llave!CLI_nucleo))
   If PUB_TIPMOV = 10 And ww_num <> 0 Then
      ' LIMPIA PRECIOS.
    For fila = 2 To FORMGEN.grid_fac.Rows - 1
       FORMGEN.grid_fac.TextMatrix(fila, 6) = "0.00"
    Next fila
    End If
End If






'SE USA CUANDO ENTRE CIAS UNA VENTA ES COMPRA EN OTRA CIA
'If cli_llave.RowCount > 0 And Val(i_codcli.Text) > 0 Then
'   If cli_llave!CLI_CODCLIE = Val(i_codcli.Text) Then
'      vuelca_datos
'      If PUB_TIPMOV = 20 Then calcula_totales
'      If SUT_LLAVE!SUT_tipmov_ref <> 0 Then
'         If cli_llave!CLI_GRUPO <> SUT_LLAVE!SUT_tipmov_ref Then
'            MsgBox "Cliente no pertenece al grupo... "
'            i_codcli.Text = ""
'            Exit Sub
'         End If
'       End If
'   End If
'End If


If cli_llave.RowCount > 0 And Val(i_codcli.Text) > 0 And PUB_TIPMOV <> 10 And PUB_TIPMOV <> 20 Then
   If cli_llave!CLI_CODCLIE = Val(i_codcli.Text) Then
      If Nulo_Valors(cli_llave!CLI_MONEDA) = "D" Then
   '      WS_MONEDA_CLI = "D"
   '      i_moneda.Text = " $"
   '      Lmoneda.Caption = "  $"
      Else
  '       Lmoneda.Caption = "S/."
  '       i_moneda.Text = "S/."
  '       WS_MONEDA_CLI = "S"
      End If
End If
End If



If cli_llave.EOF = False And Val(i_codcli.Text) > 0 Then
   If cli_llave!CLI_CODCLIE = Val(i_codcli.Text) Then
      If Left(SUT_LLAVE!SUT_DESCTO, 1) = "P" And Nulo_Valor0(cli_llave!CLI_PORDESCTO) <> 0 Then MsgBox "Cliente tiene Descto de : " & Format(cli_llave!CLI_PORDESCTO, "0.00") & "%", 48, Pub_Titulo
      If LK_CODTRA <> 2401 Then
        i_limcre_ant.Text = Nulo_Valor0(cli_llave!cli_limcre)
      End If
      If Trim(i_ds.Text) = "S" Then
         i_limcre.Text = Nulo_Valor0(cli_llave!cli_limcre)
      Else
         i_limcre.Text = Nulo_Valor0(cli_llave!cli_limcre2)
      End If
      If LK_CODTRA = 2401 And i_cambio.Value <> 1 And PUB_PROCESO = 1 Then
        For fila = 0 To i_fbg.ListCount - 1
          i_fbg.ListIndex = fila
          If Trim(i_fbg.Text) = Trim(cli_llave!CLI_TIPO) Then Exit For
        Next fila
      End If
      If LK_CODTRA = 2580 Then
        For fila = 0 To i_fbg.ListCount - 1
          i_fbg.ListIndex = fila
          If Trim(i_fbg.Text) = Trim(cli_llave!CLI_TIPO) Then Exit For
        Next fila
        i_impto2.Text = Nulo_Valor0(cli_llave!cli_limcre2)
        i_dias.Text = Val(Nulo_Valor0(cli_llave!CLI_AUTO1))
       
       For tf = 0 To i_condi.ListCount - 1
           If Val(cli_llave!cli_DIAS_FAC) = Val(Right(i_condi.List(tf), 6)) Then
              i_condi.ListIndex = tf
           End If
        Next tf
      End If
      If LK_CODTRA = 2401 And LK_EMP <> "HER" Then
         i_condi.Clear
         xcuenta = -1
         For fila = 1 To 30 Step 2
           key_descto = Mid(Trim(Nulo_Valors(cli_llave!CLI_CASA1)), fila, 2)
           If Trim(key_descto) = "" Then Exit For
           SQ_OPER = 1
           PUB_CODCIA = LK_CODCIA
           PUB_TIPREG = 333
           PUB_NUMTAB = key_descto
           LEER_TAB_LLAVE
           If Not tab_llave.EOF Then
              i_condi.AddItem Trim(tab_llave!tab_nomlargo) & String(80, " ") & tab_llave!tab_numtab
              xcuenta = xcuenta + 1
              'CAMBIAR
               ' Val(cli_llave!CLI_SUBGRUPO)
              If loc_var_descto = Val(tab_llave!tab_numtab) Then
                 defecto = xcuenta
              End If
          End If
         Next fila
         If i_condi.ListCount = 1 Then
            i_condi.ListIndex = 0
            i_condi.Enabled = False
         ElseIf i_condi.ListCount = 0 Then
            MsgBox "No Existe definición de Descuentos.", 48, Pub_Titulo
            i_condi.Enabled = False
            i_codcli.Text = ""
         Else
           i_condi.ListIndex = defecto
           i_condi.Enabled = True
      '     i_condi.SetFocus
         End If
         
      End If
      
      If Frame1.Visible = True Then
         BLOQ(0).Value = 0
         BLOQ(1).Value = 0
         BLOQ(2).Value = 0
         BLOQ(3).Value = 0
         If cli_llave!CLI_TIPO_BLOQ1 = "1" Then BLOQ(0).Value = 1
         If cli_llave!CLI_TIPO_BLOQ2 = "1" Then BLOQ(1).Value = 1
         If cli_llave!CLI_TIPO_BLOQ3 = "1" Then BLOQ(2).Value = 1
         If cli_llave!CLI_TIPO_BLOQ4 = "1" Then BLOQ(3).Value = 1
      End If
   End If
End If
If PUB_TIPMOV = 10 And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 1 And PUB_PROCESO = 0 Then
   pub_deuda = CAR_TOT_CC("C", pu_codcia, cli_llave!CLI_CODCLIE)
   If pub_deuda <> 0 Then
   If PUB_FLAG_VENCIDO_CC = 1 And Nulo_Valor0(par_llave!PAR_DIAS_VENC_CC) <> 0 Then
      MsgBox "Venta al Contado NO PROCEDE - Pendientes: " & Chr(13) & pub_mensaje, 48, Pub_Titulo
      fracli.Visible = False
      i_codcli.SetFocus
      i_codcli.Text = ""
      Exit Sub
   Else
      MsgBox "Contado Pendientes: " & Chr(13) & pub_mensaje, 48, Pub_Titulo
   End If
   End If
End If
If PUB_TIPMOV = 10 And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 And PUB_PROCESO = 0 Then
If Not cli_llave.EOF Then

    If LK_FLAG_EXED = "A" Then
    If Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) <> 1 And PUB_TIPMOV = 10 And pub_signo_car <> 0 And Nulo_Valor0(cli_llave!cli_limcre) = 0 And Nulo_Valor0(cli_llave!cli_limcre2) = 0 Then
       MsgBox "Cliente No tiene Limite de Credito.", 48, Pub_Titulo
    End If
    End If
   If SUT_LLAVE!SUT_SIGNO_CAR = 1 Then
      pu_codcia = LK_CODCIA
      pub_deuda = CAR_TOT_CPX2("C", pu_codcia, cli_llave!CLI_CODCLIE)
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
         i_codcli.Text = ""
         Exit Sub
      End If
      i_dias.Text = Val(Nulo_Valor0(cli_llave!CLI_AUTO1))
'      If PUB_FLAG_DOC > Nulo_Valor0(cli_llave!CLI_AUTO1) Then
'         MsgBox "Cliente alcanzo el tope de Documentos " & Chr(13) & "Emiitidos : " & PUB_FLAG_DOC & Chr(13) & "Autorizados : " & Trim(Nulo_Valor0(cli_llave!CLI_AUTO1)) & Chr(13) & " No procede la Venta", 48, Pub_Titulo
'         i_codcli.Text = ""
'         Exit Sub
'      End If

   End If
End If
End If
LV_CLI.Visible = False

End Sub

Private Sub i_codsunat_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

NUMERO = FORMGEN.i_codsunat.WhatsThisHelpID

avanza_campo

End Sub

Private Sub i_codsunat_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then
    PUB_TIPREG = 50
    PUB_CODCIA = "00"
    Load FrmDatArti
    FrmDatArti.Caption = Trim(Left(i_codsunat.Text, 30)) & " TAB_TIPREG = " & PUB_TIPREG
    FrmDatArti.Show 1
    PUB_TIPREG = 50
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
    i_codsunat.Clear
    Do Until tab_mayor.EOF
       i_codsunat.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
       tab_mayor.MoveNext
    Loop
    If i_codsunat.ListCount > 0 Then i_codsunat.ListIndex = 0
    i_codsunat.SetFocus
    SendKeys "%{UP}"
    DoEvents
End If

End Sub

Private Sub i_codven_Change()
If i_codven.Text = "" Then i_nomven.Caption = ""

End Sub

Private Sub i_codven_GotFocus()
Azul i_codven, i_codven
End Sub

Private Sub i_codven_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.

If Not LV_VEN.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codcli.Text = "" Then
  loc_key = 1
  Set LV_VEN.SelectedItem = LV_VEN.ListItems(loc_key)
  LV_VEN.ListItems.Item(loc_key).Selected = True
  LV_VEN.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_VEN.ListItems.count Then loc_key = LV_VEN.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_VEN.ListItems.count Then loc_key = LV_VEN.ListItems.count
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
  LV_VEN.ListItems.Item(loc_key).Selected = True
  LV_VEN.ListItems.Item(loc_key).EnsureVisible
  i_codven.Text = Trim(LV_VEN.ListItems.Item(loc_key).Text) & " "
  DoEvents
  i_codven.SelStart = Len(i_codven.Text)
  DoEvents
fin:

End Sub

Private Sub i_codven_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
Dim pos1, CARAC, I, izq, der

If KeyCode = 45 Then
   var = InputBox("Digite clave de acceso : ")
   If var = "" Then Exit Sub
   If var <> "3A" Then Exit Sub
   GoTo transfer
End If

If Len(i_codven.Text) = 0 Or IsNumeric(i_codven.Text) Then
   LV_VEN.Visible = False
   Exit Sub
End If
If LV_VEN.Visible = False Or Len(i_codven.Text) = 1 Then
    loc_key = 0
    var = Asc(i_codven.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 2
    archi = "SELECT VEM_CODVEN , VEM_CODCIA, VEM_NOMBRE  FROM VEMAEST WHERE  VEM_CODCIA = '" & LK_CODCIA & "' AND VEM_NOMBRE BETWEEN '" & i_codven.Text & "' AND  '" & var & "' ORDER BY VEM_NOMBRE"
    PROC_LISVIEW LV_VEN
    loc_key = 0
    If LV_VEN.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If LV_VEN.Visible Then
  Set itmFound = LV_VEN.FindItem(LTrim(i_codcli.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_VEN.ListItems.count Then
      LV_VEN.ListItems.Item(LV_VEN.ListItems.count).EnsureVisible
   Else
     LV_VEN.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If

Exit Sub

transfer:

   Dim PSFAR_REG As rdoQuery
   Dim FAR_REG As rdoResultset
   LKCHEK = False
   pub_total_2455 = 0
   pub_cadena = "SELECT * FROM facart WHERE FAR_CODCIA = ? AND FAR_TIPMOV = 27 AND FAR_FECHA = ?   ORDER BY FAR_FBG, FAR_NUMSER, FAR_NUMFAC "
   Set PSFAR_REG = CN.CreateQuery("", pub_cadena)
   PSFAR_REG(0) = LK_CODCIA
   PSFAR_REG(1) = LK_FECHA_DIA
   Set FAR_REG = PSFAR_REG.OpenResultset(rdOpenKeyset, rdConcurValues)
   FAR_REG.Requery
   If FAR_REG.EOF Then
      MsgBox "No hay datos para Procesar ..Revisar Procedimiento"
      Exit Sub
   End If
   If FAR_REG!far_transito = "X" Then
      FAR_REG.MoveLast
      If FAR_REG!far_transito = "T" Then
         MsgBox "El sistema detectó que el proceso fue interrumpido ...Se iniciará el proceso complementario ... "
      End If
      FAR_REG.MoveLast
      PUB_NUMFAC = 0
      Do Until FAR_REG.BOF
         If FAR_REG!far_transito = "X" Then
            If PUB_NUMFAC <> 0 And PUB_NUMFAC <> Val(FAR_REG!far_numfac) Then Exit Do
            PU_NUMFAC = FAR_REG!far_numfac
            PU_NUMSER = FAR_REG!far_numser
            pu_codcia = LK_CODCIA
            PU_TIPMOV = 10
            PU_FBG = FAR_REG!FAR_FBG
            SQ_OPER = 1
            LEER_FAR_LLAVE
            If far_llave.EOF = True Then
               FAR_REG.Edit
               FAR_REG!far_transito = "T"
               FAR_REG.Update
               PUB_NUMFAC = FAR_REG!far_numfac
            End If
         End If
         
         FAR_REG.MovePrevious
      Loop
      
      FAR_REG.Requery
      End If
      

   pub_cadena = "SELECT * FROM facart WHERE FAR_CODCIA = ? AND FAR_TIPMOV = 27 AND FAR_FECHA = ?  AND FAR_TRANSITO= 'T'  ORDER BY FAR_FBG, FAR_NUMSER, FAR_NUMFAC "
   Set PSFAR_REG = CN.CreateQuery("", pub_cadena)
   PSFAR_REG(0) = LK_CODCIA
   PSFAR_REG(1) = LK_FECHA_DIA
   Set FAR_REG = PSFAR_REG.OpenResultset(rdOpenKeyset, rdConcurValues)
   FAR_REG.Requery
   If FAR_REG.EOF Then
      MsgBox "No hay datos para Procesar ..Revisar Procedimiento"
      Exit Sub
   End If
   
   PUB_PROCESO = 1
   Do Until FAR_REG.EOF
      GoSub llena_datos
      grabar_Click
   Loop
   PUB_PROCESO = 0
   MsgBox "PROCESO TERMINADO  ...", 48, Pub_Titulo
   
Exit Sub

llena_datos:
Dim WS_PRECIO As Currency
Dim WS_DATOS As String
Dim subtotal As Currency
Dim WW_CODART As Long
Dim ww_cantidad As Currency
Dim ww_codclie As Long

PUB_SUBTOTAL = 0
SUB_JABAS = 0
SUB_CANT = 0
SUB_FLETE = 0
SUB_UNIDAD = 0
FILAX = 0
grid_fac.Clear
pasa_cabeza
SUB_JABAS = 0
SUB_UNIDAD = 0
SUB_CANT = 0
PUB_SUBTOTAL = 0

   
FORMGEN.i_codcli.Text = FAR_REG!FAR_NUMDOC
pu_codclie = FAR_REG!FAR_NUMDOC

pu_cp = "C"
pu_codcia = LK_CODCIA
SQ_OPER = 1
LEER_CLI_LLAVE
If cli_llave.EOF = False Then
   DoEvents
   i_nomCLI.Visible = True
   FORMGEN.i_nomCLI.Caption = Trim(cli_llave(3))
End If

    
'WS_FBG = far_llave!far_fbg
'xxxxxxxx

WS_DATOS = "SI"
WW_CODART = FAR_REG!far_numguia
FILAX = 1
PUB_NUMSER = FAR_REG!far_numser
PUB_NUMFAC = FAR_REG!far_numfac
PUB_FBG = FAR_REG!FAR_FBG

i_serguia.Text = FAR_REG!FAR_NUMSER_C
i_numguia.Text = FAR_REG!FAR_NUMFAC_C
dias_bloqueo = 0
If FAR_REG!far_signo_car = 1 Then
    dias_bloqueo = FAR_REG!FAR_DIAS
    i_dias.Text = FAR_REG!FAR_DIAS
    If FAR_REG!FAR_NUM_LOTE = 50 Then
     i_def.ListIndex = 4
    Else
     i_def.ListIndex = 0
    End If
Else
  If FAR_REG!FAR_NUM_LOTE = 25 Then
    i_def.ListIndex = 1
  ElseIf FAR_REG!FAR_NUM_LOTE = 30 Then
    i_def.ListIndex = 2
  ElseIf FAR_REG!FAR_NUM_LOTE = 40 Then
    i_def.ListIndex = 3
  ElseIf FAR_REG!FAR_NUM_LOTE = 50 Then
    dias_bloqueo = FAR_REG!FAR_DIAS
    i_dias.Text = FAR_REG!FAR_DIAS
    i_def.ListIndex = 4
  End If
End If

pos1 = InStr(1, FORMGEN.i_def.List(i_def.ListIndex), ".", 1)
pos1 = pos1 - 1
CARAC = Mid(FORMGEN.i_def.List(i_def.ListIndex), 1, pos1)
PUB_SECUENCIA = Val(CARAC)
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_SUT_LLAVE
If SUT_LLAVE.EOF Then
   MsgBox "No existe Definicion en Sub_Transacciones... "
   End
End If
pasa_def




If PUB_FBG = "F" Then
   i_fbg.ListIndex = 0
ElseIf PUB_FBG = "B" Then
   i_fbg.ListIndex = 1
End If



i_cambio.Value = 1
i_cambio.Visible = True
i_numser.Text = PUB_NUMSER
i_numfac.Text = PUB_NUMFAC
DoEvents

If FAR_REG!FAR_MONEDA = "D" Then
   i_ds.ListIndex = 1
ElseIf FAR_REG!FAR_MONEDA = "S" Then
   i_ds.ListIndex = 0
End If


Do Until WS_DATOS = "NO"
   FILAX = FILAX + 1
   grid_fac.Rows = FILAX + 1
   WW_CODART = FAR_REG!far_numguia
   grid_fac.TextMatrix(FILAX, 1) = WW_CODART
   grid_fac.TextMatrix(FILAX, 16) = WW_CODART
   PUB_KEY = WW_CODART
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
      MsgBox "No Existe articulo..=" & WW_CODART
      End
   Else
      grid_fac.TextMatrix(FILAX, 0) = art_LLAVE!art_nombre
   End If
 
   PUB_CODART = WW_CODART
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ARM_LLAVE
  If arm_llave.EOF Then
     MsgBox "Error Grave en arti..."
     End
  End If

   
   
   PUB_PRECIO2 = FAR_REG!FAR_PRECIO
   
   grid_fac.TextMatrix(FILAX, 1) = art_LLAVE!art_alterno
   grid_fac.TextMatrix(FILAX, 4) = FAR_REG!far_cantidad
   grid_fac.TextMatrix(FILAX, 5) = FAR_REG!far_descri
   
   grid_fac.TextMatrix(FILAX, 6) = PUB_PRECIO2
   grid_fac.TextMatrix(FILAX, 7) = FAR_REG!FAR_SUBTOTAL
   grid_fac.TextMatrix(FILAX, 8) = 0
   grid_fac.TextMatrix(FILAX, 9) = 0
   grid_fac.TextMatrix(FILAX, 10) = FAR_REG!FAR_pordescto1
   grid_fac.TextMatrix(FILAX, 11) = ""
   grid_fac.TextMatrix(FILAX, 12) = FAR_REG!far_signo_arm
   grid_fac.TextMatrix(FILAX, 13) = " "
   grid_fac.TextMatrix(FILAX, 19) = 0
   grid_fac.TextMatrix(FILAX, 14) = 1
   grid_fac.TextMatrix(FILAX, 34) = Format(FAR_REG!far_PESO * FAR_REG!far_cantidad, "0.00")
   grid_fac.TextMatrix(FILAX, 37) = Format(FAR_REG!far_PESO, "0.00")
   grid_fac.TextMatrix(FILAX, 16) = WW_CODART
   grid_fac.TextMatrix(FILAX, 33) = WW_CODART
   grid_fac.TextMatrix(FILAX, 11) = arm_llave!arm_cospro
   grid_fac.TextMatrix(FILAX, 11) = arm_llave!arm_cospro
   grid_fac.TextMatrix(FILAX, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(FILAX, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(FILAX, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)

   i_subtotal.Text = FAR_REG!FAR_BRUTO
   i_impto.Text = FAR_REG!far_IMPTO
   i_codven.Text = FAR_REG!FAR_CODVEN
   WS_BRUTO = i_subtotal.Text
   FAR_REG.Edit
   FAR_REG!far_transito = "X"
   FAR_REG.Update
   FAR_REG.MoveNext
   If FAR_REG.EOF = False Then
   If Val(PUB_NUMSER) = Val(FAR_REG!far_numser) And Val(PUB_NUMFAC) = Val(FAR_REG!far_numfac) And PUB_FBG = FAR_REG!FAR_FBG Then
   Else
      WS_DATOS = "NO"
   End If
   End If
   
   If FAR_REG.EOF = True Then
      WS_DATOS = "NO"
   End If
   
Loop
'pub_flag_cambio = 0
Return


End Sub

Private Sub i_codven_LostFocus()
If Trim(i_codven) = "" Then Exit Sub
If ven_llave.RowCount = 0 Then Exit Sub
If ven_llave.EOF And Len(i_codven.Text) > 0 Then
   MsgBox "Definir Codigo de Vendedor..."
'   i_codven.SetFocus
End If
If Not ven_llave.EOF Then
   If Val(i_codven.Text) <> ven_llave!VEM_CODVEN Then
      MsgBox "Revisar Codigo de Vendedor..."
      i_codven.Text = ""
   End If
End If


End Sub

Private Sub i_cheser_LostFocus()
If LK_CODTRA = 2720 Then
  BolFac_Click
  i_chenum.Text = i_numfac.Text
End If
If LK_CODTRA = 5715 Then Exit Sub
'If LK_CODTRA = 5318 Or LK_CODTRA = 2720 Or LK_CODTRA = 2738 Or LK_CODTRA = 5714 Or LK_CODTRA = 2748 Or LK_CODTRA = 2735 Or LK_CODTRA = 2748 Then Exit Sub
   PUB_CHESEC = 0
   PUB_CODBAN = Val(FORMGEN.i_codban.Text)
   PUB_CHESER = FORMGEN.i_cheser.Text
   SQ_OPER = 3
   PUB_FECHA = 0
   LEER_CHE_LLAVE
   If che_menor.EOF Then
    MsgBox "Talonario no Exite..", 48, Pub_Titulo
   Else
    If che_menor.EOF = False Then
       FORMGEN.i_chenum.Text = che_menor!CHE_CHENUM
    End If
   End If

End Sub

Private Sub i_condi_Click()
loc_var_descto = 0
If LK_CODTRA = 2401 And i_cambio.Value = 0 Then
   loc_var_descto = Val(Right(i_condi.Text, 6))
   If loc_var_descto = 0 Then
      MsgBox "Cliente no Tiene definición de Descuento.", 48, Pub_Titulo
   End If
   FORMGEN.grid_fac.Clear
   pasa_cabeza
   grid_fac.Rows = 3
   TEXTOVAR.Visible = False
End If

End Sub

Private Sub i_condi_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
NUMERO = FORMGEN.i_condi.WhatsThisHelpID
avanza_campo

End Sub

Private Sub i_descto_Click()
If PUB_TIPMOV = 20 And LK_EMP <> "HER" Then
  i_descto.Locked = True
Else
  i_descto.Locked = False
End If
End Sub

Private Sub i_descto_DblClick()
Dim ww_desc As Currency
If PUB_TIPMOV = 20 Then GoTo TIP_20

If PUB_TIPMOV = 10 Then GoTo TIP_10
Exit Sub
TIP_10:

Exit Sub
TIP_20:
'pub_mensaje = "Flete x Cantidad (Si), x Peso (No)...."
Pub_Respuesta = vbYes ' MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   ww_desc = 0
   fila = 2
   Do Until fila = grid_fac.Rows
    If Pub_Respuesta = vbYes Then
       If Val(grid_fac.TextMatrix(fila, 7)) <> 0 Then
          grid_fac.TextMatrix(fila, 10) = Format(Val(grid_fac.TextMatrix(fila, 7)) * (Val(i_descto.Text) / 100), "0.00")
       End If
    End If
    ww_desc = ww_desc + Val(grid_fac.TextMatrix(fila, 10))
    fila = fila + 1
   Loop
   calcula_totales
   MsgBox "Total Descuento Acumulado : " & ww_desc
   i_flete.SetFocus

End Sub

Private Sub i_descto_GotFocus()
Azul i_descto, i_descto
i_descto.MaxLength = 13
End Sub

Private Sub i_descto_LostFocus()
'calcula_totales
End Sub

Private Sub i_dias_Change()
If LK_CODTRA = 2401 And LK_EMP = "3AA" Then
    If dias_bloqueo = 0 Then
    '    i_dias.Text = 0
    Else
        If Val(i_dias.Text) > dias_bloqueo Then
          i_dias.Text = dias_bloqueo
        '  Azul i_dias, i_dias
        End If
    End If
End If
If grabar.Enabled = False Then Exit Sub
If IsDate(i_fecha_compra) = False Then Exit Sub
FORMGEN.i_fecha_vcto.Text = Str(DateAdd("d", Val(FORMGEN.i_dias.Text), i_fecha_compra.Text))
On Error GoTo SALE

Exit Sub
SALE:
End Sub

Private Sub i_dias_GotFocus()
If LK_CODTRA = 2580 Then
    i_dias.Locked = False
    Exit Sub
End If
If LK_CODTRA = 2402 Then
 If PUB_TIPMOV = 10 And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 1 Then
 '   i_dias.Locked = True
    Exit Sub
  End If
Else
 If PUB_TIPMOV = 10 And pub_signo_car = 1 And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 Then
    i_dias.Locked = False
    Exit Sub
 ElseIf PUB_TIPMOV = 10 And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) <> 1 And Nulo_Valor0(SUT_LLAVE!SUT_SIGNO_CAR) = 0 And grid_fac.Visible Then
     i_dias.Locked = True
     Exit Sub
 End If
 End If
 
' If LK_CODTRA = 2401 And LK_EMP = "3AA" Then
'   i_dias.Locked = True
' Else
'   i_dias.Locked = False
' End If
 i_dias.MaxLength = 3
 Azul i_dias, i_dias

End Sub

Private Sub i_dircli_GotFocus()
pub_cadena = "SELECT * FROM DIRCLI WHERE CODCIA=? AND CODCLI=? AND CP=?"
Set PSFAR_TRANS = CN.CreateQuery("", pub_cadena)
PSFAR_TRANS.rdoParameters(0) = " "
PSFAR_TRANS.rdoParameters(1) = 0
PSFAR_TRANS.rdoParameters(2) = " "
Set FAR_TRANS = PSFAR_TRANS.OpenResultset(rdOpenKeyset, rdConcurValues)
PSFAR_TRANS.rdoParameters(0) = LK_CODCIA
PSFAR_TRANS.rdoParameters(1) = Val(i_codcli.Text)
PSFAR_TRANS.rdoParameters(2) = PUB_CP
FAR_TRANS.Requery
i_dircli.Clear
Do Until FAR_TRANS.EOF
  i_dircli.AddItem Trim(FAR_TRANS!DIRCOMP) & String(60, " ") & Trim(FAR_TRANS!DIRCLI)
  FAR_TRANS.MoveNext
Loop
If i_dircli.ListCount = 1 Then i_dircli.ListIndex = 0

End Sub

Private Sub i_dircli_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
NUMERO = FORMGEN.i_dircli.WhatsThisHelpID
avanza_campo

End Sub

Private Sub i_ds_Click()

If i_ds.Text = "S" Then
   i_moneda.Text = "S/."
   WS_MONEDA_CLI = "S"
   ' AGREGE ACV
   If LK_CODTRA = 5345 Then
    LMONEDAB.Caption = "S/."
    lmoneda.Caption = " $."
   Else
    LMONEDAB.Caption = "S/."
    lmoneda.Caption = "S/."
   End If
   i_moneda.BackColor = QBColor(15)
   i_ds.BackColor = QBColor(15)
Else
   WS_MONEDA_CLI = "D"
   i_moneda.Text = " $"
   i_moneda.BackColor = QBColor(10)
   i_ds.BackColor = QBColor(10)
   If LK_CODTRA = 5345 Then
    LMONEDAB.Caption = " $"
    lmoneda.Caption = "S/."
   Else
    LMONEDAB.Caption = " $"
    lmoneda.Caption = " $."
   End If
End If
' camtext
If LK_CODTRA = 1455 Then calcula_totales2

If PUB_TIPMOV = 10 And LK_EMP <> "CAM" Then
   For FILAX = 2 To FORMGEN.grid_fac.Rows - 1
       PUB_CODART = Val(grid_fac.TextMatrix(FILAX, 16))
       If PUB_CODART > 0 Then
       pu_codcia = LK_CODCIA
       PUB_SECUENCIA = grid_fac.TextMatrix(FILAX, 12)
       SQ_OPER = 1
       LEER_PRE_LLAVE
       If WS_MONEDA_CLI = "D" Then
          grid_fac.TextMatrix(FILAX, 35) = Val(pre_llave!pre_pre11)
          grid_fac.TextMatrix(FILAX, 6) = redondea(Val(grid_fac.TextMatrix(FILAX, 35)) * (100 - Val(grid_fac.TextMatrix(FILAX, 10))) / 100)
          grid_fac.TextMatrix(FILAX, 7) = redondea(Val(grid_fac.TextMatrix(FILAX, 6)) * Val(grid_fac.TextMatrix(FILAX, 4)))
          grid_fac.TextMatrix(FILAX, 32) = Val(pre_llave!PRE_PRE55)
       Else
          grid_fac.TextMatrix(FILAX, 35) = Val(pre_llave!PRE_PRE1)
          grid_fac.TextMatrix(FILAX, 6) = redondea(Val(grid_fac.TextMatrix(FILAX, 35)) * (100 - Val(grid_fac.TextMatrix(FILAX, 10))) / 100)
          grid_fac.TextMatrix(FILAX, 7) = redondea(Val(grid_fac.TextMatrix(FILAX, 6)) * Val(grid_fac.TextMatrix(FILAX, 4)))
          grid_fac.TextMatrix(FILAX, 32) = Val(pre_llave!PRE_PRE5)
       End If
       End If
   Next FILAX
calcula_totales
End If

   
End Sub

Private Sub i_ds_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   Exit Sub
End If

NUMERO = FORMGEN.i_ds.WhatsThisHelpID
avanza_campo

End Sub
Private Sub i_fbg_Click()
i_numfac.Text = ""
i_numser.Text = ""
i_cambio.Value = 0
BolFac.Caption = ""
VER_LINEAS
calcula_totales
End Sub

Private Sub i_fbg_GotFocus()
textovarl.Height = 30
textovarl.Width = 30
If gridl.Top = grid_canje.Top Then grid_canje.Visible = False
End Sub

Private Sub i_fbg_LostFocus()
'i_numser_c.Text = ""
If Trim(i_fbg.Text) = "G" And LK_CODTRA = 1401 And LK_EMP = "CAM" And LK_FLAG_SOS <> "A" Then
  i_numser.Text = "9"
End If
If LK_CODTRA = 2401 Then ' SOLO EN VETAS FORZAR EL CODIGO DE VENDEDOR
  PSUSU_LLAVE(0) = LK_CODUSU
  usu_llave.Requery
  If LK_CODCIA = "01" Or LK_CODCIA = "30" Then
   If Trim(i_fbg.Text) = "B" And Val(i_codven.Text) = 10 Then GoTo bien_ven
  End If
  If Trim(i_fbg.Text) = "B" Then
    If Val(Nulo_Valor0(usu_llave!USU_CODTRA1)) <> 0 Then  ' BOLETAS
      i_codven.Text = usu_llave!USU_CODTRA1
      PUB_CODVEN = i_codven.Text
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      LEER_VEN_LLAVE
      If ven_llave.EOF Then
         Azul FORMGEN.i_codven, FORMGEN.i_codven
         MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
         FORMGEN.i_codven.SetFocus
      Else
        FORMGEN.i_nomven.Caption = ven_llave(2)
      End If
    End If
  ElseIf Trim(i_fbg.Text) = "F" Then
    If Val(Nulo_Valor0(usu_llave!USU_CODTRA2)) <> 0 Then  ' FACTURAS
      i_codven.Text = usu_llave!USU_CODTRA2
      PUB_CODVEN = i_codven.Text
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      LEER_VEN_LLAVE
      If ven_llave.EOF Then
         Azul FORMGEN.i_codven, FORMGEN.i_codven
         MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
         FORMGEN.i_codven.SetFocus
      Else
          FORMGEN.i_nomven.Caption = ven_llave(2)
      End If
    End If
  End If
End If
bien_ven:
'i_cambio.Value = 0
If pub_signo_car <> 0 Then calcula_totales
If LK_CODTRA = 1122 And PUB_TIPMOV = 97 Then PUB_FBG = "N"
If Trim(SUT_LLAVE!SUT_abreviado) = "CONS" And (LK_EMP = "PAR" Or PUB_SECUENCIA = 10) Then
   i_fbg.ListIndex = 2
End If

'BolFac_Click
End Sub

Private Sub i_fecha_can_Change()
If LK_CODTRA = 1111 Then Exit Sub
If LK_CODTRA = 2748 Then Exit Sub
If LK_CODTRA = 5318 Then Exit Sub
If IsDate(i_fecha_can.Text) Then
   If i_cambio.Value <> 1 Then
     i_fecha_oper.Text = Format(i_fecha_can.Text, "dd/mm/yy")
   End If
End If
End Sub

Private Sub i_fecha_can_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
If Not IsDate(i_fecha_can.Text) Then
  MsgBox "Fecha NO procede ...", 48, Pub_Titulo
  Azul2 i_fecha_can, i_fecha_can
  Exit Sub
End If
If LK_CODTRA = 2748 Then
  i_concepto.SetFocus
  Exit Sub
End If
NUMERO = FORMGEN.i_fecha_can.WhatsThisHelpID
avanza_campo
fin:

End Sub

Private Sub i_fecha_compra_Change()
If LK_CODTRA = 1111 Then Exit Sub
If IsDate(i_fecha_compra.Text) Then
  If i_cambio.Value <> 1 Then
   i_fecha_oper.Text = Format(i_fecha_compra.Text, "dd/mm/yy")
  End If
  If i_dias.Visible Then
    i_dias_Change
   End If
End If

End Sub
Private Sub i_fecha_compra_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
If Not IsDate(i_fecha_compra.Text) Then
  MsgBox "Fecha NO procede ...", 48, Pub_Titulo
  i_fecha_compra.SetFocus
  Exit Sub
End If
NUMERO = FORMGEN.i_fecha_compra.WhatsThisHelpID
avanza_campo
fin:

End Sub

Private Sub i_fecha_oper_Change()
If LK_CODTRA = 1111 Then Exit Sub
If LK_CODTRA = 5318 Then
If IsDate(i_fecha_oper.Text) Then
   If i_cambio.Value <> 1 Then
     i_fecha_can.Text = Format(i_fecha_oper.Text, "dd/mm/yy")
   End If
End If
End If
End Sub

Private Sub i_fecha_oper_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

If Not IsDate(i_fecha_oper.Text) Then
   MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
   Azul2 i_fecha_oper, i_fecha_oper
   Exit Sub
End If
If LK_CODTRA = 2748 And pub_signo_ccm <> 0 Then
    Azul i_codban, i_codban
    Exit Sub
ElseIf LK_CODTRA = 2748 And pub_signo_ccm = 0 Then
    i_ds.SetFocus
    Exit Sub
ElseIf LK_CODTRA = 5714 Then
    Azul cta_gas1, cta_gas1
    Exit Sub

End If


NUMERO = FORMGEN.i_fecha_oper.WhatsThisHelpID
avanza_campo

End Sub

Private Sub i_fecha_vcto_GotFocus()
Azul i_fecha_vcto, i_fecha_vcto
End Sub

Private Sub i_flete_DblClick()
Dim qsw As Currency
Dim ww_desc

If PUB_TIPMOV = 20 Or PUB_TIPMOV = 101 Then
Else
Exit Sub
End If
pub_mensaje = "Flete x Cantidad (Si), x Valor (No)...."
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   ww_desc = 0
   fila = 2
   Do Until fila = grid_fac.Rows
    If Pub_Respuesta = vbYes Then
       If Val(grid_fac.TextMatrix(1, 4)) <> 0 Then
          grid_fac.TextMatrix(fila, 8) = Format((Val(i_flete.Text) / Val(grid_fac.TextMatrix(1, 4))) * Val(grid_fac.TextMatrix(fila, 4)), "0.00")
       End If
    Else
       If Val(grid_fac.TextMatrix(1, 7)) <> 0 Then
          grid_fac.TextMatrix(fila, 8) = Format((Val(i_flete.Text) / Val(grid_fac.TextMatrix(1, 7))) * Val(grid_fac.TextMatrix(fila, 7)), "0.00")
       End If
    End If
     ww_desc = ww_desc + Val(grid_fac.TextMatrix(fila, 8))
   fila = fila + 1
   Loop
   If Val(ww_desc) <> Val(i_flete.Text) Then
      qsw = Val(i_flete.Text) - Val(ww_desc)

      grid_fac.TextMatrix(2, 8) = Val(grid_fac.TextMatrix(2, 8)) + qsw
      ww_desc = ww_desc + qsw
   End If
   calcula_totales
   MsgBox "Total Flete Acumulado : " & ww_desc
   i_flete.SetFocus

End Sub

Private Sub i_flete_GotFocus()
i_flete.MaxLength = 13
End Sub

Private Sub i_flete_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_flete, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If
PUB_FLETE = Val(FORMGEN.i_flete.Text)
i_flete_DblClick

'If LK_EMP = "HER" Then ' QUITAR DESPUES
   calcula_totales
'End If

If LK_EMP = "HER" Then Exit Sub 'QUITAR
fin:

End Sub

Private Sub i_flete_LostFocus()
'If LK_EMP = "HER" Then ' QUITAR DESPUES
   calcula_totales
'End If

End Sub

Private Sub i_gastos_Change()
PUB_GASTOS = Val(FORMGEN.i_gastos.Text)
calcula_totales
End Sub

Private Sub i_gastos_GotFocus()
i_gastos.MaxLength = 13
Azul i_gastos, i_gastos
End Sub
Private Sub i_importe_amort_Change()
If LK_CODTRA = 2738 Or LK_CODTRA = 2720 Then
 i_importe.Text = i_importe_amort.Text
Exit Sub
End If
If LK_CODTRA = 2750 Then Exit Sub
If i_bruto2.Visible = True Then
   i_impto2.Text = Val(i_importe_amort.Text) - Val(i_importe_amort.Text) / (1 + LK_IGV / 100)
   i_impto2.Text = Format(i_impto2.Text, "0.00")
   i_bruto2.Text = Val(i_importe_amort.Text) - Val(i_impto2.Text)
   i_neto.Text = i_importe_amort.Text
   WS_BRUTO = i_bruto2.Text
   i_impto.Text = i_impto2.Text
   i_subtotal.Text = i_bruto2.Text
   If LK_CODTRA = 2748 Then
    imp_gas1.Text = ""
    imp_gas2.Text = ""
    If Val(i_subtotal.Text) <> 0 Then
      If Trim(cta_gas1.Text) <> "" Then imp_gas1.Text = redondea(Val(i_subtotal.Text))
      If Trim(cta_gas2.Text) <> "" Then imp_gas2.Text = "0.00"
    End If
  End If
End If
End Sub

Private Sub i_importe_amort_GotFocus()
Azul i_importe_amort, i_importe_amort
End Sub

Private Sub i_importe_Change()
If i_codban2.Visible = True Then i_importe_amort.Text = ""
If LK_CODTRA = 5714 Then imp_gas1.Text = i_importe.Text
End Sub

Private Sub i_importe_LostFocus()
Dim ww_tc As Double
Dim wa_moneda As String * 1
If LK_CODTRA = 1111 Then Exit Sub
If LK_CODTRA = 2735 Then
    If ccm_llave.EOF Then Exit Sub
    If ccm_llave!CCM_MONEDA = LOC_MONEDA_DIG Then
     i_importe.Text = Val(gridl.TextMatrix(1, 5))
     Exit Sub
    End If
    wa_moneda = ccm_llave!CCM_MONEDA
ElseIf LK_CODTRA = 2725 Then
   If i_ds.Visible Then
     If Trim(i_ds.Text) = "" Then Exit Sub
     If Trim(i_ds.Text) = LOC_MONEDA_DIG Then
        i_importe.Text = Val(gridl.TextMatrix(1, 5))
        Exit Sub
     End If
     wa_moneda = Trim(i_ds.Text)
   Else
    i_importe.Text = Val(gridl.TextMatrix(1, 5))
   End If
End If



If gridl.Visible And Val(gridl.TextMatrix(1, 5)) <> 0 Then
   If wa_moneda = "S" Then
     ww_tc = Format(i_importe.Text / Val(gridl.TextMatrix(1, 5)), "0.0000000")
   Else
     ww_tc = Format(Val(gridl.TextMatrix(1, 5) / i_importe.Text), "0.0000000")
   End If
   If Abs(Val(ww_tc) - Val(LK_TIPO_CAMBIO)) > 0.4 Then
     MsgBox "Existe mucha Diferencia en el T.C. Anterior, si todavia desea aplicar su T.C. debe dar click en la opción de T.C. para cambiarlo ", 48, Pub_Titulo
     i_importe.Text = Format(Val(gridl.TextMatrix(1, 5)) * LK_TIPO_CAMBIO, "0.00")
     Exit Sub
   End If
   LK_TIPO_CAMBIO = ww_tc
   'MDIForm1.StatusBar1.Panels(3).Text = "T.C.= S/. " + Format(LK_TIPO_CAMBIO, "0.0000")
   MsgBox "Cambio el T.C. = " & LK_TIPO_CAMBIO, 48, Pub_Titulo
End If


End Sub

Private Sub i_impto_Change()
   If LK_CODTRA = 1401 Then
    i_neto.Text = Val(i_subtotal.Text) + Val(i_gastos.Text) + Val(i_impto.Text)
   Else
    i_neto.Text = Val(i_subtotal.Text) + Val(i_gastos.Text) - Val(i_descto.Text) + Val(i_impto.Text)
   End If

End Sub

Private Sub i_impto_GotFocus()
PUB_IMPTO = Val(i_impto.Text)
End Sub

Private Sub i_impto_LostFocus()
If PUB_IMPTO = 0 Then Exit Sub
If (Abs(Val(i_impto.Text) - PUB_IMPTO) * 100) / (PUB_IMPTO) > 2 Then
   MsgBox "No Procede...demasiada diferencia"
   i_impto.Text = PUB_IMPTO
End If
End Sub

Private Sub i_impto2_Change()
If LK_CODTRA = 2750 Then Exit Sub
If LK_CODTRA = 2748 Then
Else
 i_bruto2.Text = Val(i_importe_amort.Text) - Val(i_impto2.Text)
End If
End Sub

Private Sub i_impto2_GotFocus()
Azul i_impto2, i_impto2
End Sub

Private Sub i_impto2_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_impto2, KeyAscii

If KeyAscii <> 13 Then Exit Sub
If LK_CODTRA = 2748 And (LK_EMP = "3AA" Or LK_EMP = "PIU") Then
  Azul cta_gas1, cta_gas1
  Exit Sub
End If
NUMERO = FORMGEN.i_impto2.WhatsThisHelpID
avanza_campo


End Sub

Private Sub i_impto2_LostFocus()
If LK_CODTRA = 2748 Then Exit Sub
If LK_CODTRA = 2750 Then
  i_importe_amort.Text = Val(i_bruto2.Text) + Val(i_impto2.Text)
  Exit Sub
End If
If Val(i_importe_amort.Text) - Val(i_impto2.Text) = 0 Then Exit Sub
WS_IMPORTE = Val(i_impto2.Text) * 100 / (Val(i_importe_amort.Text) - Val(i_impto2.Text))
If WS_IMPORTE > LK_IGV - 0.5 And WS_IMPORTE < LK_IGV + 0.5 Then
Else
   MsgBox "Revisar el Igv ..."
   i_bruto2.Text = Val(i_importe_amort.Text) - Val(i_impto2.Text)
End If

i_neto.Text = i_importe_amort.Text

End Sub

Private Sub i_is_Click()
If grid_fac.COL = 12 Then grid_fac.Text = i_is.Text


If Val(i_is.Text) = -1 Then
   If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) > Val(grid_fac.TextMatrix(grid_fac.Row, 13)) Then
      MsgBox "Ojo stock insuficiente ..."
      If LK_USU_STOCK <> "A" Then
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         i_is.ListIndex = -1
         Exit Sub
      End If
   End If
End If

End Sub

Private Sub i_is_GotFocus()
'Exit Sub
TEXTOVAR.Visible = False
'i_is.Visible = True
If Val(grid_fac.Text) <> 0 Then
   If Val(grid_fac.Text) = 1 Then
      i_is.ListIndex = 0
   Else
      i_is.ListIndex = 1
   End If
Else
   grid_fac.Text = 1
   i_is.ListIndex = 0
End If

End Sub

Private Sub i_is_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
   
If grid_fac.TextMatrix(grid_fac.Row, 16) = "" Then
   grid_fac.COL = tab_derecha(grid_fac.COL)
   Exit Sub
End If
   
flag_salto = 1
If grid_fac.Row = grid_fac.Rows - 1 Then
   grid_fac.Rows = grid_fac.Rows + 1
End If
'grid_fac.SetFocus
'grid_fac.Row = grid_fac.Row + 1
'grid_fac.RowHeight(grid_fac.Row) = 315
flag_salto = 0

If grid_fac.COL = 12 Then i_is.Visible = False
'grid_fac.COL = tab_derecha(grid_fac.COL)
'grid_fac.Row = grid_fac.Row + 1
grid_fac.COL = 1
grid_fac.Row = grid_fac.Rows - 1
grid_fac.RowHeight(grid_fac.Row) = 315
grid_fac.SetFocus
grid_fac.Row = grid_fac.Rows - 1
TEXTOVAR.SetFocus

'grid_fac.RowHeight(grid_fac.Row) = 315


fin:


End Sub

Private Sub i_is_LostFocus()
i_is.Visible = False

End Sub

Private Sub i_limcre_GotFocus()
Azul i_limcre, i_limcre
End Sub

Private Sub i_mortal_Click()
grid_fac.Text = i_mortal.Text & "     " & i_mortal.ListIndex
End Sub

Private Sub i_mortal_GotFocus()
TEXTOVAR.Visible = False
i_mortal.Visible = True
If grid_fac.COL <> 9 Then Exit Sub
If Len(grid_fac.Text) <> 0 Then
   i_mortal.ListIndex = Right(grid_fac.Text, 5)
Else
   grid_fac.Text = i_mortal.Text & "     " & i_mortal.ListIndex
   i_mortal.ListIndex = 0
End If

End Sub
Private Sub i_mortal_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
   
If grid_fac.TextMatrix(grid_fac.Row, 16) = "" Then
   grid_fac.COL = tab_derecha(grid_fac.COL)
   Exit Sub
End If
   
flag_salto = 1
If grid_fac.Row = grid_fac.Rows - 1 Then
   grid_fac.Rows = grid_fac.Rows + 1
End If
grid_fac.SetFocus
grid_fac.Row = grid_fac.Row + 1
grid_fac.RowHeight(grid_fac.Row) = 315
flag_salto = 0

grid_fac.COL = tab_derecha(grid_fac.COL)

fin:

End Sub

Private Sub i_mortal_LostFocus()
i_mortal.Visible = False
End Sub

Private Sub i_neto_Change()
If i_cambio.Value = 1 And Val(i_neto.Text) <> 0 And pub_flag_cambio <> 99 Then
   pub_flag_cambio = 1
End If

End Sub

Private Sub i_neto_GotFocus()
i_neto.MaxLength = 17
Azul i_neto, i_neto
End Sub

Private Sub i_nombre_banco_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_nombre_banco.WhatsThisHelpID

avanza_campo
fin:

End Sub


Private Sub i_nompro_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_nompro.WhatsThisHelpID

avanza_campo
fin:

End Sub

Private Sub i_num_cheque_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_num_cheque.WhatsThisHelpID

avanza_campo
fin:

End Sub

Private Sub i_numfac_c_GotFocus()
textovarl.Height = 30
textovarl.Width = 30
Azul i_numfac_c, i_numfac_c
End Sub

Private Sub i_numfac_c_LostFocus()
If LK_CODTRA = 2748 Then
    SQ_OPER = 8
    PU_TIPMOV = PUB_TIPMOV
    pu_codcia = LK_CODCIA
    PU_NUMSER = Val(i_numser_c.Text)
    PU_FBG = "K"
    PU_FBG2 = " "
    PU_NUMFAC = Val(i_numfac_c.Text)
    LEER_FAR_LLAVE
    If Not far_menor4.EOF Then
      Do Until far_menor4.EOF
      If Val(far_menor4!far_codclie) = Val(i_codcli.Text) Then
         MsgBox " Documento ya Emitido, Secuencia Nro :" & Format(far_menor4!far_numfac, "000"), 48, Pub_Titulo
         i_numfac_c.Text = ""
         i_numser_c.Text = ""
         Azul i_numser_c, i_numser_c
         Exit Sub
      End If
      far_menor4.MoveNext
      Loop
    End If
End If
If LK_CODTRA = 2720 Then
    SQ_OPER = 1
    PU_TIPMOV = 10
    pu_codcia = LK_CODCIA
    PU_NUMSER = Val(i_numser_c.Text)
    PU_FBG = i_fbg.Text
    PU_NUMFAC = Val(i_numfac_c.Text)
    LEER_FAR_LLAVE
    If far_llave.EOF Then
        MsgBox "Documento no Existe.", 48, Pub_Titulo
        Azul i_numser_c, i_numser_c
        Exit Sub
    End If
    far_llave.MoveLast
    If Trim(far_llave!far_estado) = "E" Then
        MsgBox "Documento esta Anulado.", 48, Pub_Titulo
        Azul i_numser_c, i_numser_c
        Exit Sub
    End If
    If Val(far_llave!far_codclie) <> Val(i_codcli.Text) Then
        MsgBox "El Documento no pertenece a este Cliente.", 48, Pub_Titulo
        Azul i_numser_c, i_numser_c
        Exit Sub
    End If
    If WS_MONEDA_CCM = "D" Then
     If far_llave!FAR_MONEDA = "S" Then
      If JALAR_TC(far_llave!FAR_fecha_compra, 3) <> 0 Then i_importe_amort.Text = redondea((Val(far_llave!FAR_BRUTO) + Val(far_llave!far_IMPTO)) / JALAR_TC(far_llave!FAR_fecha_compra, 3))
     Else
       i_importe_amort.Text = redondea((Val(far_llave!FAR_BRUTO) + Val(far_llave!far_IMPTO)))
     End If
    Else
     If far_llave!FAR_MONEDA = "D" Then
       i_importe_amort.Text = redondea((Val(far_llave!FAR_BRUTO) + Val(far_llave!far_IMPTO)) * JALAR_TC(far_llave!FAR_fecha_compra, 3))
     Else
       i_importe_amort.Text = redondea((Val(far_llave!FAR_BRUTO) + Val(far_llave!far_IMPTO)))
     End If
    End If
End If

End Sub

Private Sub i_numfac_Change()
If LK_CODTRA = 2401 Or LK_CODTRA = 2402 Then
Else
Exit Sub
End If

'If Val(i_numfac.text) < pub_numfac_ult And i_cambio.Value = 1 And i_cambio.Visible Then
If i_cambio.Value = 1 And i_cambio.Visible Then
   i_fecha_compra.Left = 8200
   i_fecha_compra.Top = 1700
   LABELGEN(16).Caption = "Fec. Emisión:"
   LABELGEN(16).Top = i_fecha_compra.Top
   LABELGEN(16).Left = i_fecha_compra.Left - 1200
   LABELGEN(16).Visible = True
   i_fecha_compra.Visible = True
   
Else
   i_fecha_compra.Visible = False
End If

End Sub

Private Sub i_numfac_GotFocus()
'Stop
End Sub

Private Sub i_numfac_LostFocus()
Dim longit As String * 15
Dim WIGV_G As Currency
Dim WNUMERO As String
Dim WF As String * 1
Dim ww_flag As String * 1
Dim WS_FECHITA As Date
Dim wscodusu As String
If i_cambio.Value = 0 Then Exit Sub
WF = ""
ww_flag = "N"
'If Not cli_llave.EOF Then If Trim(Nulo_Valors(cli_llave!cli_cia_ref)) <> "" Then WF = "A"
pub_flag_cambio = 99
If Val(i_numfac.Text) = 0 Then Exit Sub
If LK_CODTRA = 1122 Then Exit Sub
Dim fin As Integer
Dim WS_NUMFAC, WS_NUMOPER
wscodusu = LK_CODUSU
If PUB_TIPMOV = 0 And LK_CODTRA <> 1403 Then Exit Sub

If PUB_TIPMOV = 20 And i_cambio.Value = 0 Then Exit Sub

If Not i_cambio.Visible Then Exit Sub

PU_NUMFAC = Val(i_numfac.Text)
pu_codcia = LK_CODCIA
PU_NUMSER = Val(i_numser.Text)
If i_fbg.Visible = False Or PUB_TIPMOV = 20 Then
  If LK_CODTRA <> 2748 Then PUB_FBG = ""
Else
   PUB_FBG = i_fbg.Text
End If
If PUB_TIPMOV = 97 Then PUB_FBG = "N"
If PUB_TIPMOV = 98 Then PUB_FBG = "D"
PU_TIPMOV = PUB_TIPMOV
PU_FBG = PUB_FBG
If LK_EMP = "PIU" And i_cambio.Value = 1 Then
  If PUB_TIPMOV = 5 Or PUB_TIPMOV = 100 Or PUB_TIPMOV = 101 Then
    PU_TIPMOV = 20
  End If
End If


'If WF = "A" Then
'   pu_codcia = cli_llave!cli_cia_ref
'   PU_TIPMOV = 10
'   PU_FBG = i_fbg.text
'End If

SQ_OPER = 1
LEER_FAR_LLAVE
If far_llave.EOF = True Then
 If PUB_PROCESO = 1 Then Exit Sub
pub_mensaje = "NO Existe datos en este Documento ...Desea saltar a este numero? "
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   BolFac_Click
End If
 If grid_fac.Visible Then grid_fac.SetFocus
 Exit Sub
End If

If LK_EMP = "PIU" And i_cambio.Value = 1 Then
  If PUB_TIPMOV = 5 Or PUB_TIPMOV = 100 Or PUB_TIPMOV = 101 Then
    PU_TIPMOV = PUB_TIPMOV
    GoTo LLENA
  End If
End If

far_llave.MoveLast
'If far_llave!far_estado <> "E" And LK_CODTRA <> 1401 And LK_CODTRA <> 2414 Then
'  MsgBox "Nª. Documento: " + i_numser.Text + " - " + i_numfac.Text + "  - está  Emitido...no procede", 48, Pub_Titulo
'  i_numfac.Text = ""
'  i_numfac.SetFocus
'  i_cambio.Value = 0
'  BolFac_Click
'  Exit Sub
'End If
If far_llave!FAR_fecha <> LK_FECHA_DIA Then
If loc_mofi_venta <> "A" And i_cambio.Value = 1 And LK_CODTRA = 2401 Then
  BolFac_Click
  MsgBox "No tiene Acceso a modificar documentos anteriores a la fecha del día", vbExclamation, Pub_Titulo
  Exit Sub
End If
End If
If LK_CODTRA = 1401 Or LK_CODTRA = 2401 Or LK_CODTRA = 2414 Or LK_CODTRA = 2407 Or LK_CODTRA = 2409 Or LK_CODTRA = 2403 Then
 SQ_OPER = 1
 PUB_TIPREG = 60
 WNUMERO = Format(far_llave!FAR_fecha_compra, "mm") & Format(far_llave!FAR_fecha_compra, "yyyy")
 PUB_NUMTAB = Val(WNUMERO)
 PUB_CODCIA = LK_CODCIA
 LEER_TAB_LLAVE
 If Not tab_llave.EOF Then
   If tab_llave!tab_codart = 1 Then
     BolFac_Click
     MsgBox "El documento pertenece a un Periodo cerrado, no procede", vbInformation, Pub_Titulo
     Exit Sub
   End If
 End If
End If

pub_mensaje = "Existe datos en este Documento : " & Trim(i_numser.Text) & " - " & i_numfac.Text & " ¿Desea Mostrar los Datos en pantalla?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
    far_llave.MoveLast
    wscodusu = far_llave!far_codusu
    If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
    Else
     If Trim(wscodusu) <> LK_CODUSU And PUB_TIPMOV = 10 Then
        BolFac_Click
        cancelar_Click
        MsgBox "No le corresponde a su Usuario !!!", 48, Pub_Titulo
        Exit Sub
     End If
    End If

'   BolFac_Click
    If (LK_CODTRA = 2401 And far_llave!far_signo_car <> 0) Then
        SQ_OPER = 4
        pu_codcia = LK_CODCIA
        PU_FBG = far_llave!FAR_FBG
        PUB_TIPDOC = far_llave!FAR_TIPDOC
        PUB_NUMSER = far_llave!far_numser
        PUB_NUMFAC = far_llave!far_numfac
        LEER_CAR_LLAVE
        If car_far.EOF Then
            MsgBox "Verificar Documento....", 48, Pub_Titulo
        Else
          If Val(car_far!car_importe) <> Val(car_far!CAR_IMP_INI) Then
            BolFac_Click
            MsgBox "Documento  tiene amortizaciones", 48, Pub_Titulo
            Exit Sub
          End If
        End If
    End If

    If i_fecha_compra.Visible Then Azul2 i_fecha_compra, i_fecha_compra
   Exit Sub
End If

grid_fac.Clear
grid_fac.Rows = 3
pasa_cabeza
fin = 0
far_llave.MoveLast
WS_NUMOPER = far_llave!FAR_NUMOPER
far_llave.MoveFirst

If LK_CODTRA = 2748 Then GoTo GO_2748

Do Until fin = 1
' CAMTEX
   If LK_CODTRA = 1401 And far_llave!far_estado = "E" And LK_FECHA_DIA <> far_llave!FAR_fecha Then GoTo pasa
   'If far_llave!far_estado = "E" Then GoTo pasa
   
   If far_llave!FAR_CODCIA <> LK_CODCIA And WF <> "A" Then
      MsgBox "!!! NO TE CORRESPONDE..."
      GoTo fin
   End If
   WS_NUMOPER = far_llave!FAR_NUMOPER
   ' AGREGE UNA OPCION MAS ACV
   WS_FECHITA = far_llave!FAR_fecha
   If LK_CODTRA = 1405 Then GoTo SALTAX
   If far_llave!FAR_fecha <> LK_FECHA_DIA And WF = "C" Then
      MsgBox "!!! NO ES DEL DIA...", 48, Pub_Titulo
      GoTo fin
   End If
   
   'WS_FECHITA = far_llave!FAR_fecha
   'If LK_CODTRA = 1405 Then GoTo SALTAX
   
pasa:
   far_llave.MoveNext
   If far_llave.EOF Then fin = 1
Loop

SALTAX:


far_llave.MoveFirst
If LK_CODTRA = 2401 And i_cambio.Value = 1 Then
 i_codcli.Text = far_llave!far_codclie
 i_dircli_GotFocus
 For fila = 0 To i_dircli.ListCount - 1
   i_dircli.ListIndex = fila
  If Val(Trim(Right(i_dircli.Text, 8))) = Val(far_llave!far_key_dircli) Then
    Exit For
  End If
 Next fila
 For fila = 0 To cmdtipo.ListCount - 1
  cmdtipo.ListIndex = fila
  If Val(Trim(Right(cmdtipo.Text, 8))) = Val(far_llave!far_turno) Then
    Exit For
  End If
 Next fila
End If

fila = 2
If LK_CODTRA = 1401 Or (LK_CODTRA = 2401 And far_llave!far_signo_car <> 0) Then
    SQ_OPER = 4
    pu_codcia = LK_CODCIA
    PU_FBG = far_llave!FAR_FBG
    PUB_TIPDOC = far_llave!FAR_TIPDOC
    PUB_NUMSER = far_llave!far_numser
    PUB_NUMFAC = far_llave!far_numfac
    LEER_CAR_LLAVE
    If car_far.EOF Then
        MsgBox "Observación: < Documento ha sido Anulado. >", 48, Pub_Titulo
    Else
      car_far.MoveLast
      If Val(car_far!car_importe) <> Val(car_far!CAR_IMP_INI) And car_far!CAR_SITUACION <> "E" Then
        BolFac_Click
        MsgBox "Documento  tiene amortizaciones", 48, Pub_Titulo
        Exit Sub
      End If
    End If
End If

If far_llave!FAR_TIPMOV = 98 Or (far_llave!FAR_TIPMOV = 97 And far_llave!far_codart = 0) Then
   If i_codcli.Visible Then i_codcli.Text = far_llave!far_codclie
   i_codven.Text = Nulo_Valor0(far_llave!FAR_CODVEN)
   i_concepto.Text = Trim(far_llave!far_concepto)
   i_bruto2.Text = far_llave!FAR_BRUTO
   i_impto2.Text = far_llave!far_IMPTO
   i_neto.Text = far_llave!FAR_BRUTO + far_llave!far_IMPTO
   i_importe_amort.Text = Val(far_llave!FAR_BRUTO) + Val(far_llave!far_IMPTO)
   'GoTo MUE
   Exit Sub
End If

Do Until far_llave.EOF
   grid_fac.Rows = fila + 2
   If far_llave!FAR_ESTADO2 = "L" Then GoTo PASA_LOTE
   If WS_FECHITA = far_llave!FAR_fecha And WS_NUMOPER = far_llave!FAR_NUMOPER Then
      PUB_KEY = far_llave!far_codart
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ART_LLAVE
      If art_LLAVE.EOF And PUB_KEY <> 0 Then
         MsgBox "Error Grave en arti..."
         Exit Sub
      End If
      PUB_CODART = far_llave!far_codart
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ARM_LLAVE
      If arm_llave.EOF Then
         MsgBox "Error Grave en arti..."
         Exit Sub
      End If
      PUB_CODART = far_llave!far_codart
      pu_codcia = LK_CODCIA
      PUB_SECUEN = far_llave!FAR_ORDEN_UNIDADES ' far_num_precio
      SQ_OPER = 1
      LEER_PRE_LLAVE
      If pre_llave.EOF Then
         MsgBox "Error Grave en arti..."
         Exit Sub
      End If
   If WF <> "A" Then i_codcli.Text = far_llave!far_codclie
   
   'i_dias.Text = far_llave!FAR_DIAS
   i_codven.Text = Nulo_Valor0(far_llave!FAR_CODVEN)
   If far_llave!far_estado = "E" Then
     i_concepto.Text = ""
   Else
     i_concepto.Text = far_llave!far_concepto
   End If
   If Nulo_Valor0(far_llave!far_numguia) <= 0 Then
      i_numguia.Text = ""
   Else
    i_numguia.Text = far_llave!far_numguia
   End If
   
   i_numfac_c.Text = far_llave!FAR_NUMFAC_C
   i_numser_c.Text = far_llave!FAR_NUMSER_C
   i_fecha_compra.Text = Format(far_llave!FAR_fecha_compra, "dd/mm/yy")
   'If LK_EMP <> "HER" Then
   If i_fecha_oper.Visible Then
   i_fecha_oper.Text = Format(far_llave!FAR_fecha_pro, "dd/mm/yy")
   End If
   'End If
   FORMGEN.i_fecha_vcto.Text = Str(DateAdd("d", Val(FORMGEN.i_dias.Text), i_fecha_compra.Text))
   If LK_FLAG_GRIFO = "A" Then
    i_turno.Text = far_llave!far_turno
    i_TEXTONCRE.Text = far_llave!FAR_PEDFAC
   End If
   If PUB_TIPMOV = 10 And LK_EMP <> "HER" Then
      i_TEXTONCRE.Text = far_llave!FAR_OC
   End If
   i_serguia.Text = far_llave!far_serguia
   
   If PUB_KEY <> 0 Then grid_fac.TextMatrix(fila, 0) = art_LLAVE!art_nombre
   If LK_FLAG_ORIGINAL = "A" Then
     grid_fac.TextMatrix(fila, 1) = art_LLAVE!ART_KEY
   Else
     grid_fac.TextMatrix(fila, 1) = art_LLAVE!art_alterno
   End If
   grid_fac.TextMatrix(fila, 16) = far_llave!far_codart
   grid_fac.TextMatrix(fila, 33) = far_llave!far_codart
   grid_fac.TextMatrix(fila, 2) = far_llave!far_JABAS
   grid_fac.TextMatrix(fila, 3) = far_llave!far_UNIDADES
   grid_fac.TextMatrix(fila, 11) = far_llave!FAR_COSPRO
   If Nulo_Valor0(far_llave!FAR_equiv) > 0 Then
   grid_fac.TextMatrix(fila, 4) = far_llave!far_cantidad / Nulo_Valor0(far_llave!FAR_equiv)
   Else
   grid_fac.TextMatrix(fila, 4) = far_llave!far_cantidad
   End If
   grid_fac.TextMatrix(fila, 6) = far_llave!FAR_PRECIO
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(far_llave!FAR_equiv)
   longit = Nulo_Valors(far_llave!far_descri)
   grid_fac.TextMatrix(fila, 5) = longit & "       " & far_llave!FAR_ORDEN_UNIDADES & "       " & far_llave!FAR_equiv
   'i_unidades.AddItem longit & "       " & pre_mayor!pre_secuencia & "       " & pre_mayor!pre_equiv
   grid_fac.TextMatrix(fila, 8) = Nulo_Valor0(far_llave!FAR_FLETE)
   If LK_EMP = "3AA" Or LK_EMP = "HER" Then
    grid_fac.TextMatrix(fila, 10) = Trim(far_llave!far_PORDESCTOS)
   Else
     grid_fac.TextMatrix(fila, 10) = far_llave!FAR_DESCTO
   End If
   ' ICA
   grid_fac.TextMatrix(fila, 32) = pre_llave!PRE_PRE5
   grid_fac.TextMatrix(fila, 35) = far_llave!far_precio_neto
   
   grid_fac.TextMatrix(fila, 11) = far_llave!FAR_COSPRO
   grid_fac.TextMatrix(fila, 12) = far_llave!far_signo_arm
   grid_fac.TextMatrix(fila, 42) = far_llave!far_PORDESCTOS
   grid_fac.TextMatrix(fila, 10) = far_llave!far_PORDESCTOS
   grid_fac.TextMatrix(fila, 49) = Nulo_Valor0(far_llave!FAR_pordescto1)
   
   WIGV_G = far_llave!far_IMPTO
 '  If LK_EMP = "3AA" Then
     If PUB_TIPMOV = 101 Then
       grid_fac.TextMatrix(fila, 34) = Format(pre_llave!pre_PESO * far_llave!far_cantidad, "0.00")
       grid_fac.TextMatrix(fila, 37) = Format(pre_llave!pre_PESO, "0.00")
       grid_fac.TextMatrix(fila, 43) = Format(Nulo_Valor0(pre_llave!PRE_LITRO), "0.000")
     Else
       grid_fac.TextMatrix(fila, 34) = Format(far_llave!far_PESO * far_llave!far_cantidad, "0.00")
       grid_fac.TextMatrix(fila, 37) = Format(far_llave!far_PESO, "0.00")
       grid_fac.TextMatrix(fila, 43) = Format(Nulo_Valor0(far_llave!far_litro), "0.00")
     End If
  ' End If
   If i_fbg.Visible And LK_ACTIVA = "A" Then
    If PUB_TIPMOV <> 0 And (i_fbg.Text = "F" Or i_fbg.Text = "B") Then
        grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_saldo_s
    Else
        grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_Saldo_n
    End If
   Else
      grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_STOCK
   End If
   If NJ.Visible And LK_ACTIVA = "A" Then
     If NJ.Text = "S" Then
       grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_saldo_s
     Else
       grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_Saldo_n
     End If
  End If
'   If WF = "A" Then grid_fac.TextMatrix(fila, 12) = pub_signo_arm
   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   If PUB_TIPMOV = 20 Then
    grid_fac.TextMatrix(fila, 7) = Format(far_llave!FAR_SUBTOTAL, "0.00")
    grid_fac.TextMatrix(fila, 25) = 1
   End If
   
   If far_llave!far_estado = "E" Then ww_flag = "S"
   fila = fila + 1
   End If
   WS_MONEDA_CLI = far_llave!FAR_MONEDA
   i_fecha_compra.Text = Format(far_llave!FAR_fecha_compra, "dd/mm/yy")
   i_fecha_can.Text = Format(far_llave!FAR_fecha_can, "dd/mm/yy")
   i_flete.Text = Nulo_Valor0(far_llave!FAR_TOT_FLETE)
   i_descto.Text = Nulo_Valor0(far_llave!FAR_TOT_DESCTO)
   t_nombre.Text = Trim(far_llave!FAR_CLIENTE)
   t_doc.Text = Trim(far_llave!FAR_DOCCLI)
   t_direc.Text = Trim(far_llave!FAR_dircli)
   wscodusu = far_llave!far_codusu
PASA_LOTE:
   far_llave.MoveNext
Loop

If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
Else
 If Trim(wscodusu) <> LK_CODUSU And PUB_TIPMOV = 10 Then
    BolFac_Click
    cancelar_Click
    MsgBox "No le corresponde a su Usuario !!!", 48, Pub_Titulo
    Exit Sub
 End If
End If

MUE:
calcula_totales
If PUB_TIPMOV = 20 Then
  i_impto.Text = Format(WIGV_G, "0.00")
End If
If i_codcli.Visible Then
  i_codcli.SetFocus
  i_codcli_KeyPress 13
  If PUB_TIPMOV = 10 And Val(i_codcli.Text) = 1 Then
    t_direc_KeyPress 13
  End If

End If

   If WS_MONEDA_CLI = "S" Then
     If Val(i_ds.ListCount) = 1 Then
      i_ds.ListIndex = 0
     Else
    '  i_ds.ListIndex = 1
      i_ds.ListIndex = 0
     End If
   ElseIf WS_MONEDA_CLI = "D" Then
      If Val(i_ds.ListCount) = 1 Then
        'i_ds.ListIndex = 1
        i_ds.ListIndex = 0
      Else
        'i_ds.ListIndex = 0
        i_ds.ListIndex = 1
      End If
   Else
      i_ds.ListIndex = -1
   End If
TEXTOVAR.Visible = False
If grid_fac.Visible Then grid_fac.SetFocus
If i_cambio.Value = 1 And i_cambio.Visible And i_fecha_compra.Visible = True Then i_fecha_compra.SetFocus

'pub_flag_cambio = 0
i_fecha_compra.Visible = True
If LK_CODTRA = 2401 Then
  LABELGEN(16).Caption = "Fec. Emis."
  i_fecha_compra.Top = i_codcli.Top
  LABELGEN(16).Top = i_fecha_compra.Top
  LABELGEN(16).Left = i_fecha_compra.Left - 1000
End If
Exit Sub
' RUTINA PARA JALAR UNA COMPRA EN LA 2403  VENUS

LLENA:
fila = 2
Do Until far_llave.EOF
      grid_fac.Rows = grid_fac.Rows + 1 ' fila + 2
      grid_fac.RowHeight(grid_fac.Rows - 1) = 315
      PUB_KEY = far_llave!far_codart
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ART_LLAVE
      If art_LLAVE.EOF And PUB_KEY <> 0 Then
         MsgBox "Error Grave en arti..."
         Exit Sub
      End If
      PUB_CODART = far_llave!far_codart
      pu_codcia = LK_CODCIA
      SQ_OPER = 1
      LEER_ARM_LLAVE
      If arm_llave.EOF Then
         MsgBox "Error Grave en arti..."
         Exit Sub
      End If
      
   i_concepto.Text = far_llave!far_concepto
   
   If PUB_KEY <> 0 Then grid_fac.TextMatrix(fila, 0) = art_LLAVE!art_nombre
   
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!art_alterno
   grid_fac.TextMatrix(fila, 16) = far_llave!far_codart
   grid_fac.TextMatrix(fila, 33) = far_llave!far_codart
   grid_fac.TextMatrix(fila, 2) = far_llave!far_JABAS
   grid_fac.TextMatrix(fila, 3) = far_llave!far_UNIDADES
   grid_fac.TextMatrix(fila, 11) = far_llave!FAR_COSPRO
   If Nulo_Valor0(far_llave!FAR_equiv) > 0 Then
   grid_fac.TextMatrix(fila, 4) = far_llave!far_cantidad / Nulo_Valor0(far_llave!FAR_equiv)
   Else
   grid_fac.TextMatrix(fila, 4) = far_llave!far_cantidad
   End If
   grid_fac.TextMatrix(fila, 6) = arm_llave!arm_cospro ' far_llave!far_PRECIO
   grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(far_llave!FAR_equiv)
   grid_fac.TextMatrix(fila, 5) = Nulo_Valors(far_llave!far_descri)
   grid_fac.TextMatrix(fila, 8) = Nulo_Valor0(far_llave!FAR_FLETE)
   grid_fac.TextMatrix(fila, 10) = far_llave!FAR_DESCTO
   grid_fac.TextMatrix(fila, 11) = far_llave!FAR_COSPRO
   grid_fac.TextMatrix(fila, 12) = far_llave!far_signo_arm
   grid_fac.TextMatrix(fila, 37) = far_llave!far_PESO
   'modi
   If i_fbg.Visible And LK_ACTIVA = "A" Then
    If PUB_TIPMOV <> 0 And (i_fbg.Text = "F" Or i_fbg.Text = "B") Then
       grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_saldo_s
    Else
      grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_Saldo_n
    End If
   Else
      grid_fac.TextMatrix(fila, 13) = arm_llave!ARM_STOCK
   End If
   If NJ.Visible And LK_ACTIVA = "A" Then
     If NJ.Text = "S" Then
       grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_saldo_s
     Else
       grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_Saldo_n
     End If
   End If

   grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
   
   If far_llave!far_estado = "E" Then ww_flag = "S"
   fila = fila + 1
   far_llave.MoveNext
Loop
If fila <> 2 Then
 calcula_totales
End If
i_cambio.Value = 0

fin:
Exit Sub
GO_2748:
On Error GoTo MIRA
   far_llave.MoveLast
   If i_codcli.Visible Then i_codcli.Text = far_llave!far_codclie
   i_importe_amort.Text = Val(far_llave!FAR_BRUTO) + Val(far_llave!far_IMPTO) + Val(far_llave!FAR_GASTOS)
   i_bruto2.Text = Val(far_llave!FAR_BRUTO)
   i_impto2.Text = Val(far_llave!far_IMPTO)
   i_importe.Text = Val(far_llave!FAR_GASTOS)
   'i_importe_amort.Text = Val(far_llave!FAR_BRUTO) + Val(far_llave!far_impto)
   i_dias.Text = far_llave!FAR_DIAS
   If Nulo_Valor0(far_llave!far_numguia) <= 0 Then
      i_numguia.Text = ""
   Else
    i_numguia.Text = far_llave!far_numguia
   End If
   
   i_numfac_c.Text = far_llave!FAR_NUMFAC_C
   i_numser_c.Text = far_llave!FAR_NUMSER_C
   i_fecha_compra.Text = Format(far_llave!FAR_fecha_compra, "dd/mm/yy")
   i_fecha_vcto.Text = Str(DateAdd("d", Val(FORMGEN.i_dias.Text), i_fecha_compra.Text))
   i_serguia.Text = far_llave!far_serguia
   For fila = 0 To i_def.ListCount - 1
      i_def.ListIndex = fila
      If Val(Right(i_def.Text, 6)) = far_llave!FAR_NUM_LOTE Then
          i_def_LostFocus
          Exit For
      End If
   Next fila
   For fila = 0 To i_codsunat.ListCount - 1
      i_codsunat.ListIndex = fila
      If Val(Right(i_codsunat.Text, 6)) = far_llave!far_cod_sunat Then
          Exit For
      End If
  Next fila
   
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_FECHA = far_llave!FAR_fecha
   LEER_ALL_LLAVE
   Do Until all_llave.EOF
    If all_llave!ALL_NUMOPER = far_llave!FAR_NUMOPER Then
      Exit Do
    End If
   all_llave.MoveNext
   Loop
   If all_llave.EOF Then
     MsgBox "Verificar...", 48, Pub_Titulo
     Exit Sub
   End If
   If far_llave!FAR_MONEDA = "S" Then
     i_ds.ListIndex = 0
   Else
     i_ds.ListIndex = 1
   End If
   
   i_dias.Text = far_llave!FAR_DIAS
   i_fecha_oper.Text = Format(all_llave!ALL_FECHA_PRO, "dd/mm/yy")
   i_fecha_can.Text = Format(all_llave!ALL_FECHA_CAN, "dd/mm/yy")
   i_concepto.Text = all_llave!all_concepto
   cta_gas1.Text = all_llave!ALL_CTAG1
   cta_gas2.Text = all_llave!ALL_CTAG2
   imp_gas1.Text = all_llave!ALL_IMPG1
   imp_gas2.Text = all_llave!ALL_IMPG2
   If all_llave!ALL_SIGNO_CCM <> 0 Then
     i_codban.Text = all_llave!all_codban
     i_cheser.Text = all_llave!ALL_CHESER
     i_chenum.Text = all_llave!all_chenum
   End If
   If i_codcli.Visible Then
     i_codcli_KeyPress 13
   End If
'   If far_llave!far_signo_ccm <> 0 Then
'    i_codban.Text = far_llave!far_codban
'   End If


Exit Sub
MIRA:
 MsgBox "Reintente Nuevamente", 48, Pub_Titulo
 Resume Next
End Sub

Private Sub i_numguia_GotFocus()
If Nulo_Valors(SUT_LLAVE!SUT_abreviado) = "GUIAS" Then

'If PUB_TIPMOV = 10 Or LK_CODTRA = 2409 Then
   PUB_SERGUIA = Val(i_serguia.Text)
   llena_numGUIA
End If
PUB_NUMGUIA = PU_NUMFAC
End Sub

Private Sub i_numguia_LostFocus()
If PUB_TIPMOV = 10 And i_numguia.Text <> "" And PUB_NUMGUIA <> Val(i_numguia.Text) Then
   pub_mensaje = "Confirma el numero de Guia de remision ? " & i_numguia.Text
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo)
   If Pub_Respuesta = vbNo Then llena_numGUIA
End If
If LK_CODTRA = 1401 Then
  Exit Sub
 SQ_OPER = 6
 PUB_SERGUIA = Val(i_serguia.Text)
 PUB_NUMGUIA = Val(i_numguia.Text)
 pu_codcia = LK_CODCIA
 LEER_FAR_LLAVE
 If far_guia.EOF = False Then
   MsgBox "Nro de Guia ya Existe "
   Exit Sub
 End If

End If

End Sub

Private Sub i_numplan_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   Exit Sub
End If

NUMERO = FORMGEN.i_numplan.WhatsThisHelpID
avanza_campo

End Sub

Private Sub i_numser_c_LostFocus()
Dim wdia1 As String
If LK_CODTRA = 2770 Then
   i_numser_c.MaxLength = 0
End If

If ((fragas.Visible And LK_CODTRA = 2748) Or LK_CODTRA = 1401) And LK_EMP = "3AA" And cli_llave.EOF = False And pub_signo_car <> 0 Then
    If InStr(1, cli_llave!CLI_AUTO1, "-") <> 0 Then
       If Val(Left(cli_llave!CLI_AUTO1, 3)) = Val(i_numser_c.Text) Then i_dias.Text = Val(Mid(cli_llave!CLI_AUTO1, 5, Len(Trim(cli_llave!CLI_AUTO1))))
    End If
    If InStr(1, cli_llave!cli_auto2, "-") <> 0 Then
       If Val(Left(cli_llave!cli_auto2, 3)) = Val(i_numser_c.Text) Then i_dias.Text = Val(Mid(cli_llave!cli_auto2, 5, Len(Trim(cli_llave!cli_auto2))))
    End If
End If
End Sub

Private Sub i_precio_LostFocus()
i_importe.Text = Val(i_cantidad.Text) * Val(i_precio.Text)

End Sub
Private Sub i_precios_GotFocus()
If i_precios.ListCount <> 0 Then
 i_precios.ListIndex = 0
 SendKeys "%{DOWN}"
End If
'textovar.Visible = False
'i_unidades.Visible = False
'i_mortal.Visible = False
'i_precios.Visible = True
'If Len(grid_fac.text) <> 0 Then
'   i_precios.ListIndex = Right(grid_fac.text, 5)
'Else
'   grid_fac.text = i_precios.text & "     " & i_precios.ListIndex
'   i_precios.ListIndex = 0
'End If

End Sub
Private Sub i_precios_KeyDown(KeyCode As Integer, Shift As Integer)
Dim WTE
If KeyCode <> 45 Then Exit Sub
If LK_CODTRA = 2412 Then GoTo PASAE
If LK_PRECIO <> "B" Then Exit Sub
PASAE:
WTE = InputBox("Digite Precio:")
If WTE = "" Then
  Exit Sub
End If
If Not IsNumeric(WTE) Then
  MsgBox "Valor Incorrecto ..., Intente nuevamente ", 48, Pub_Titulo
  Exit Sub
End If
grid_fac.Text = WTE
grid_fac.TextMatrix(grid_fac.Row, 35) = grid_fac.Text
' ULTIMO
grid_fac.TextMatrix(grid_fac.Row, 32) = grid_fac.Text


flag_salto = 1
If grid_fac.Row = grid_fac.Rows - 1 Then
   grid_fac.Rows = grid_fac.Rows + 1
End If
grid_fac.SetFocus
grid_fac.Row = grid_fac.Row + 1
grid_fac.RowHeight(grid_fac.Row) = 315
flag_salto = 0

grid_fac.COL = tab_derecha(grid_fac.COL)

End Sub
Private Sub i_precios_KeyPress(KeyAscii As Integer)
Dim WS_DIF_PRE As Currency
Dim wpor As Currency
Dim WS_COSTO As Currency
If KeyAscii = 27 Then
 i_precios.Visible = False
 grid_fac.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then Exit Sub
   
grid_fac.Text = Mid(i_precios.Text, 10, 10)
grid_fac.TextMatrix(grid_fac.Row, 35) = grid_fac.Text
grid_fac.TextMatrix(grid_fac.Row, 32) = grid_fac.Text

grid_fac.TextMatrix(grid_fac.Row, 17) = Right(i_precios.Text, 1)
If Right(i_precios.Text, 1) = 1 Then
    wpor = pre_llave!PRE_POR1
ElseIf Right(i_precios.Text, 1) = 2 Then
    wpor = pre_llave!PRE_POR2
ElseIf Right(i_precios.Text, 1) = 3 Then
    wpor = pre_llave!PRE_POR3
ElseIf Right(i_precios.Text, 1) = 4 Then
    wpor = pre_llave!PRE_POR4
ElseIf Right(i_precios.Text, 1) = 5 Then
    wpor = pre_llave!PRE_POR5
End If
grid_fac.TextMatrix(grid_fac.Row, 44) = wpor
grid_fac.TextMatrix(grid_fac.Row, 10) = "" 'wpor
grid_fac.TextMatrix(grid_fac.Row, 42) = wpor
SQ_OPER = 1
PUB_SECUEN = Val(grid_fac.TextMatrix(grid_fac.Row, 19))
pu_codcia = LK_CODCIA
PUB_CODART = art_LLAVE!ART_KEY
LEER_PRE_LLAVE
SQ_OPER = 1
PUB_CODART = art_LLAVE!ART_KEY
pu_codcia = LK_CODCIA
LEER_ARM_LLAVE
If arm_llave.EOF Then
 WS_COSTO = 0
Else
 WS_COSTO = Format(Val(arm_llave!arm_cospro) * Val(grid_fac.TextMatrix(grid_fac.Row, 14)), "0.0000")
End If


'WS_PRECIO = Val(pre_llave!PRE_PRE1)
If LK_EMP = "HER" And PUB_TIPMOV = 10 And LK_CODCIA = "01" Then
   If Val(grid_fac.TextMatrix(grid_fac.Row, 35)) < WS_COSTO Then
     par_llave.Requery
     If par_llave!PAR_FLAG_P <> "P" Then
       MsgBox "El Precio Elegido Es Menor que el Costo y NO PROCEDE !!!", 48, Pub_Titulo
       grid_fac.TextMatrix(grid_fac.Row, 35) = ""
       grid_fac.TextMatrix(grid_fac.Row, 32) = ""
       grid_fac.COL = grid_fac.COL - 1
       grid_fac.SetFocus
       grid_fac_KeyPress 13
       Exit Sub
     Else
       MsgBox "El Precio Elegido Es Menor que el Costo! ANOTAR !!!", 48, Pub_Titulo
     End If
   End If

End If
WS_DIF_PRE = 0
If pre_llave.EOF Then
    MsgBox "Avisar sobre Mensaje: Columna 19 no tiene el codigo de unidad", 48, Pub_Titulo
    GoTo pasa_desc
End If
If pub_signo_car = 0 Then
   If Nulo_Valor0(pre_llave!PRE_PREC1) <> 0 Then
     WS_DIF_PRE = Val(grid_fac.TextMatrix(grid_fac.Row, 35)) - Val(Nulo_Valor0(pre_llave!PRE_PREC1))
     If WS_DIF_PRE = 0 Then MsgBox "Avisar sobre posible error de descuento.", 48, Pub_Titulo
   End If
   grid_fac.TextMatrix(grid_fac.Row, 49) = WS_DIF_PRE
Else
   If Nulo_Valor0(pre_llave!PRE_PREC2) <> 0 Then
     WS_DIF_PRE = Val(grid_fac.TextMatrix(grid_fac.Row, 35)) - Val(Nulo_Valor0(pre_llave!PRE_PREC2))
     If WS_DIF_PRE = 0 Then MsgBox "Avisar sobre posible error de descuento.", 48, Pub_Titulo
   End If
   grid_fac.TextMatrix(grid_fac.Row, 49) = WS_DIF_PRE
End If
pasa_desc:

  
   
flag_salto = 1
If grid_fac.Row = grid_fac.Rows - 1 Then
   grid_fac.Rows = grid_fac.Rows + 1
End If
grid_fac.SetFocus
If tab_derecha(grid_fac.COL) = 1 Then grid_fac.Row = grid_fac.Row + 1
grid_fac.RowHeight(grid_fac.Row) = 315
flag_salto = 0

grid_fac.COL = tab_derecha(grid_fac.COL)

fin:



End Sub

Private Sub i_precios_LostFocus()
'grid_fac.text = i_precios.text & "     " & i_precios.ListIndex
calcula_totales
End Sub

Private Sub i_sd_Change()

End Sub

Private Sub i_serguia_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then Exit Sub

NUMERO = FORMGEN.i_serguia.WhatsThisHelpID

avanza_campo

End Sub

Private Sub i_situacion_KeyUp(KeyCode As Integer, Shift As Integer)
If LK_CODTRA <> 1401 Then Exit Sub
If KeyCode = 45 Then
    PUB_TIPREG = 70
    PUB_CODCIA = LK_CODCIA
    Load FrmDatArti
    FrmDatArti.Caption = Trim(Left(i_situacion.Text, 30)) & " TAB_TIPREG = " & PUB_TIPREG
    FrmDatArti.Show 1
    PUB_TIPREG = 70
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    i_situacion.Clear
    Do Until tab_mayor.EOF
       i_situacion.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
       tab_mayor.MoveNext
    Loop
    If i_situacion.ListCount > 0 Then i_situacion.ListIndex = 0
    i_situacion.SetFocus
    SendKeys "%{UP}"
    DoEvents
End If

End Sub

Private Sub i_subtotal_Change()
   i_neto.Text = Val(i_subtotal.Text) + Val(i_gastos.Text) - Val(i_descto.Text) + Val(i_impto.Text)
   
End Sub

Private Sub i_TEXTONCRE_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

If LK_CODTRA = 2748 Then
  i_concepto.SetFocus
  Exit Sub
End If

NUMERO = FORMGEN.i_TEXTONCRE.WhatsThisHelpID

avanza_campo

End Sub

Private Sub i_tipdoc_Click()
letche.ToolTipText = i_tipdoc.Text
End Sub


Private Sub i_turno_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then Exit Sub

NUMERO = FORMGEN.i_turno.WhatsThisHelpID

avanza_campo

End Sub

Private Sub i_turno_LostFocus()
If LK_FLAG_GRIFO = "A" Then
If i_turno.Text = "" Then Exit Sub
If Val(i_turno.Text) > 3 Or Val(i_turno.Text) = 0 Then
   MsgBox "Turno NO Existe...", 48, Pub_Titulo
   i_turno.Text = ""
   i_turno.SetFocus
End If
End If
End Sub

Private Sub i_unidades_GotFocus()

If LK_EMP = "3AA" And LK_CODTRA <> 2414 Then
 '  i_unidades_KeyPress 13
   If (LK_EMP = "3AA" Or LK_EMP = "HER") Then
     grid_fac.TextMatrix(grid_fac.Row, 6) = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 35)) * (100 - Val(grid_fac.TextMatrix(grid_fac.Row, 10))) / 100)
     grid_fac.TextMatrix(grid_fac.Row, 7) = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 6)) * Val(grid_fac.TextMatrix(grid_fac.Row, 4)))
'     If Not textovar.Visible Then grid_fac.SetFocus
   End If
   Exit Sub
End If
'i_unidades_KeyPress 13
 SendKeys "%{DOWN}"


'i_unidades.ListIndex = 0

'textovar.Visible = False
'i_precios.Visible = False
'i_mortal.Visible = False
'i_unidades.Visible = True
'If Len(grid_fac.text) <> 0 Then
'   i_unidades.ListIndex = Right(grid_fac.text, 5)
'Else
 '  grid_fac.text = i_unidades.text & "     " & i_unidades.ListIndex
 '  i_unidades.ListIndex = 0
'End If

End Sub
Private Sub i_unidades_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 i_unidades.Visible = False
 grid_fac.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then Exit Sub
If i_unidades.Text = "" Then Exit Sub
If Trim(Mid(i_unidades.Text, 23, 3)) = "" Then
   MsgBox "Reiniciar Equipo...- Unidades"
   End
End If

PUB_SECUEN = Val(Mid(i_unidades.Text, 23, 3))

verifica_precios
grid_fac.Text = i_unidades.Text
grid_fac.TextMatrix(grid_fac.Row, 14) = Right(i_unidades.Text, 7)
grid_fac.TextMatrix(grid_fac.Row, 19) = i_unidades.ListIndex

If PUB_TIPMOV = 20 Then
   grid_fac.TextMatrix(grid_fac.Row, 6) = Nulo_Valor0(pre_llave!PRE_COSTO) 'Nulo_Valor0(arm_llave!ARM_COSTO_ULT) * grid_fac.TextMatrix(grid_fac.Row, 14)
End If
If LK_CODTRA = 2403 Then  ' And (PUB_SECUENCIA = 10 Or PUB_SECUENCIA = 0) Then
   grid_fac.TextMatrix(grid_fac.Row, 6) = Nulo_Valor0(arm_llave!arm_cospro) * Val(grid_fac.TextMatrix(grid_fac.Row, 14))
End If
If LK_CODTRA = 2409 And LK_EMP = "3AA" Then
      grid_fac.TextMatrix(grid_fac.Row, 32) = pre_llave!PRE_PRE55
      grid_fac.TextMatrix(grid_fac.Row, 35) = Nulo_Valor0(arm_llave!arm_cospro)
      grid_fac.TextMatrix(grid_fac.Row, 6) = Nulo_Valor0(arm_llave!arm_cospro)
End If


If Not arm_llave.EOF And (pub_signo_arm = -1 Or (LK_CODTRA = 2414 And grid_fac.Row = 2)) Then
   PUB_CANTIDAD = Val(grid_fac.TextMatrix(grid_fac.Row, 4)) * Val(grid_fac.TextMatrix(grid_fac.Row, 14))
   If PUB_CANTIDAD > Val(grid_fac.TextMatrix(grid_fac.Row, 13)) And Nulo_Valors(art_LLAVE!art_flag_stock) <> "A" Then
       MsgBox "Ojo Cantidad Mayor que Stock ..." & "Existencia : " & Val(grid_fac.TextMatrix(grid_fac.Row, 13)), 48, Pub_Titulo
   End If
End If

' QUITE LK_CODTRA = 2403 Or
If (LK_CODTRA = 2425 Or tab_derecha(grid_fac.COL) = 1) Then
   flag_salto = 1
   If grid_fac.Row = grid_fac.Rows - 1 Then grid_fac.Rows = grid_fac.Rows + 1
      If LK_CODTRA = 2414 Then
      i_unidades.Visible = False
      End If
      grid_fac.SetFocus
      grid_fac.Row = grid_fac.Row + 1
      grid_fac.RowHeight(grid_fac.Row) = 315
End If
   
flag_salto = 0



If i_precios.ListCount = 1 And PUB_TIPMOV = 10 Then
   flag_salto = 1
   i_precios.ListIndex = 0
   grid_fac.TextMatrix(grid_fac.Row, 6) = Mid(i_precios.Text, 10, 10)
   If LK_EMP <> "HER" Then
   grid_fac.TextMatrix(grid_fac.Row, 35) = Mid(i_precios.Text, 10, 10)
   End If
   If LK_EMP = "HER" Then
 '     grid_fac.TextMatrix(grid_fac.Row, 32) = grid_fac.TextMatrix(grid_fac.Row, 35)
   End If
   grid_fac.TextMatrix(grid_fac.Row, 17) = Right(i_precios.Text, 1)

   If Trim(SUT_LLAVE!SUT_DESCTO) <> "" And grid_fac.COL = 5 And i_precios.ListCount = 1 Then
      grid_fac.COL = 10
      If Not TEXTOVAR.Visible Then
        'ICA
         grid_fac.SetFocus
      End If
   Else
      If grid_fac.Row = grid_fac.Rows - 1 Then grid_fac.Rows = grid_fac.Rows + 1
      grid_fac.Row = grid_fac.Row + 1
      grid_fac.RowHeight(grid_fac.Row) = 315
      flag_salto = 0
      grid_fac.COL = tab_derecha(6)
      If TEXTOVAR.Visible = True Then
         TEXTOVAR.SetFocus
      Else
         grid_fac.SetFocus
      End If
  End If
Else
   grid_fac.COL = tab_derecha(grid_fac.COL)
   If LK_CODTRA = 2412 Or LK_CODTRA = 2403 Then  ' And (PUB_SECUENCIA = 10 Or PUB_SECUENCIA = 0)) Then
      grid_fac.SetFocus
   End If
End If
End Sub

Private Sub i_unidades_LostFocus()
calcula_totales
If LK_CODTRA = 2414 Then grid_fac.SetFocus
End Sub

Private Sub imp_gas1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If cta_gas2.Visible Then
     Azul cta_gas2, cta_gas2
  Else
   grabar.SetFocus
  End If
End If

End Sub

Private Sub imp_gas2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  grabar.SetFocus
End If

End Sub


Private Sub Label9_Click()

End Sub


Private Sub letche_Click()
Dim ww
Dim I
If gridl.Visible = True Then
   ww = gridl.TextMatrix(2, 24)
   I = 2
   Do Until I = gridl.Rows
   If ww <> gridl.TextMatrix(I, 24) And Val(gridl.TextMatrix(I, 25)) <> 0 Then
'      MsgBox "No Procede en monedas diferentes...           "
'      Exit Sub
   End If
   I = I + 1
Loop

End If
pasa_cabeza_canje
If i_tipdoc.Visible = True Then letche.ToolTipText = i_tipdoc.Text

If Frame4.Visible = True Then
   grid_canje.Top = Frame4.Top + 1200
   grid_canje.Height = Frame4.Height - 1900
   grid_canje.Left = Frame4.Left
   grid_canje.Width = Frame4.Width
Else
   grid_canje.Top = gridl.Top + 1300
   grid_canje.Height = gridl.Height + 500
   grid_canje.Left = gridl.Left
   grid_canje.Width = gridl.Width
End If
If grid_canje.Visible = True Then
   grid_canje.Visible = False
Else
   grid_canje.Visible = True
SS:
   ww = InputBox("Ingrese Cantidad de Documentos")
   If ww = "" Or Val(ww) > 100 Or Val(ww) = 0 Then GoTo SS
   grid_canje.Rows = ww + 2
   I = 0
   Do Until I > ww + 1
   FORMGEN.grid_canje.RowHeight(I) = 350
   FORMGEN.grid_canje.TextMatrix(I, 3) = 0
   I = I + 1
   Loop
   
End If


textovar_canje.Visible = False
If grid_canje.Visible = True Then grid_canje.SetFocus
End Sub



Private Sub letche_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Then Exit Sub
If LK_CODTRA <> 2401 Then Exit Sub

If pub_signo_car = 0 And letche.ToolTipText = "CH" Then Exit Sub

If letche.ToolTipText = "CH" Then
   letche.Caption = "LETRAS"
   letche.ToolTipText = "LE"
   If LK_EMP = "DPV" Then
      grid_fac.Clear
      pasa_cabeza
   End If
Else
   letche.Caption = "CHEQUES"
   letche.ToolTipText = "CH"
End If
   


End Sub

Private Sub LisTransa_DblClick()
LisTransa_KeyPress 13
End Sub

Private Sub i_cant_cheq_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_cant_cheq.WhatsThisHelpID

avanza_campo
fin:

End Sub

Private Sub i_cantidad_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_cantidad, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_cantidad.WhatsThisHelpID

avanza_campo
fin:

End Sub

Private Sub i_chenum_GotFocus()

Azul i_chenum, i_chenum
If pub_signo_ccm = 0 And LK_CODTRA <> 5715 And LK_CODTRA <> 231 Then
   NUMERO = FORMGEN.i_chenum.WhatsThisHelpID
   avanza_campo
End If

End Sub

Private Sub i_chenum_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If
If LK_CODTRA = 2748 Then
  i_fecha_can.SetFocus
  i_fecha_can.SelStart = 0
  i_fecha_can.SelLength = 2
  Exit Sub
End If


NUMERO = FORMGEN.i_chenum.WhatsThisHelpID

avanza_campo
fin:


End Sub

Private Sub i_cheser_GotFocus()
Azul i_cheser, i_cheser
If pub_signo_ccm = 0 And LK_CODTRA <> 5715 And LK_CODTRA <> 231 Then
   NUMERO = FORMGEN.i_cheser.WhatsThisHelpID
   avanza_campo
End If

End Sub

Private Sub i_cheser_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_cheser.WhatsThisHelpID

avanza_campo
fin:

End Sub

Private Sub i_codcli_KeyPress(KeyAscii As Integer)
Dim var  As String
Dim VALOR As String
Dim tf As Integer
Dim wcta1 As String * 12
Dim wcta2 As String * 12
Dim WDIAS As Integer
Dim wCONDI As Integer
Dim wdia1 As String
Dim wDIA2 As String
Dim I
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
wCONDI = -1
If KeyAscii = 27 Then
 i_codcli.Text = ""
 Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
If KeyAscii = 13 And Left(i_codcli.Text, 1) = "+" Then GoTo buscar
On Error GoTo OJO
pu_cp = PUB_CP
pu_codclie = Val(FORMGEN.i_codcli.Text)
On Error GoTo 0
If Len(i_codcli.Text) = 0 Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(i_codcli.Text) = True Then
   If Len(Trim(i_codcli.Text)) = LK_DIG_RUC Then ' LONG DEL RUC
        PUB_RUC = Trim(i_codcli.Text)
        SQ_OPER = 4
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_ruc.EOF Then
           MsgBox "R.U.C. No Existe ", 48, Pub_Titulo
           Exit Sub
        End If
        i_codcli.Text = cli_ruc!CLI_CODCLIE
   End If
   On Error GoTo OJO
   SQ_OPER = 1
   pu_codclie = Val(i_codcli.Text)
   pu_codcia = LK_CODCIA
   pu_cp = PUB_CP
   LEER_CLI_LLAVE
   On Error GoTo 0
   wcta1 = ""
   wcta2 = ""
   If cli_llave.EOF Then
    Azul FORMGEN.i_codcli, FORMGEN.i_codcli
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    FORMGEN.i_codcli.SetFocus
    GoTo fin
   Else
      FORMGEN.i_nomCLI.Caption = Trim(cli_llave!cli_nombre) & " - " & Trim(cli_llave!cli_ruc_esposo)
      NUMERO = FORMGEN.i_codcli.WhatsThisHelpID
    '  If LK_EMP = "3AA" Or LK_EMP = "PIU" Then
      wcta1 = Trim(cli_llave!CLI_CUENTA_CONTAB2)
      wcta2 = Trim(Nulo_Valors(cli_llave!cli_CUENTA_CONTAB22))
      WDIAS = Nulo_Valor0(cli_llave!cli_dias_cred)
      wCONDI = Nulo_Valor0(cli_llave!cli_DIAS_FAC)
      wdia1 = Nulo_Valors(cli_llave!CLI_AUTO1)
      wDIA2 = Nulo_Valors(cli_llave!cli_auto2)
      PUB_RUC = Nulo_Valors(cli_llave!cli_ruc_esposo)
      GoSub PONER
     ''   End If
      NUMERO = FORMGEN.i_codcli.WhatsThisHelpID
      avanza_campo
   End If
   
Else
On Error GoTo sigue
   If loc_key <> 0 Then VALOR = UCase(LV_CLI.ListItems.Item(loc_key).Text)
   If Trim(UCase(i_codcli.Text)) = Left(VALOR, Len(Trim(i_codcli.Text))) Then
   Else
      Exit Sub
   End If
   If loc_key = 0 Then Exit Sub
   i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).SubItems(1))
   pu_codclie = Val(FORMGEN.i_codcli.Text)
   pu_cp = PUB_CP
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   FORMGEN.i_nomCLI.Caption = Trim(cli_llave!cli_nombre) & " - " & Trim(cli_llave!cli_ruc_esposo)
   PUB_RUC = Nulo_Valors(cli_llave!cli_ruc_esposo)
   NUMERO = FORMGEN.i_codcli.WhatsThisHelpID
   wcta1 = Trim(cli_llave!CLI_CUENTA_CONTAB2)
   wcta2 = Trim(Nulo_Valors(cli_llave!cli_CUENTA_CONTAB22))
   WDIAS = Nulo_Valor0(cli_llave!cli_dias_cred)
   wCONDI = Nulo_Valors(cli_llave!cli_DIAS_FAC)
   wdia1 = Nulo_Valors(cli_llave!CLI_AUTO1)
   wDIA2 = Nulo_Valors(cli_llave!cli_auto2)
   GoSub PONER
   NUMERO = FORMGEN.i_codcli.WhatsThisHelpID
   avanza_campo
End If
LV_CLI.Visible = False
If Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) <> 1 And PUB_TIPMOV = 10 And pub_signo_car <> 0 And Nulo_Valors(cli_llave!CLI_estado) <> "A" Then
   MsgBox "              !!! O J O !!! " + Chr(13) + "Cliente No está ACTIVO en el Sistema.", 48, Pub_Titulo
End If

Exit Sub
sigue:
If Err.Number = 35600 Then
  Exit Sub
End If
fin:
OJO:
Exit Sub
PONER:
If LK_CODTRA = 2750 Then Return
If Trim(PUB_RUC) <> "" Then
  If LK_DIG_RUC <> Len(Trim(PUB_RUC)) Then
     MsgBox "Ruc Antiguo.. Verificar en Maestro", 48, Pub_Titulo
     i_codcli.Text = ""
    Exit Sub
  End If
End If

'i_dias.Locked = True
dias_bloqueo = 0
If LK_EMP = "3AA" And (PUB_CP = "P" Or PUB_CP = "C") And (LK_CODTRA = 2748 Or LK_CODTRA = 2401) Then
 If wCONDI >= 0 Then
    tf = 0
    If Val(Left(i_def.Text, 2)) > wCONDI Then
        For tf = 0 To i_def.ListCount - 1
           If wCONDI = Val(Left(i_def.List(tf), 2)) Then
              MsgBox "Condición No Autorizada....Verificar", 48, Pub_Titulo
              MsgBox "La Maxima Condición es : " & Trim(i_def.List(tf)), 48, Pub_Titulo
              i_def.ListIndex = tf
              i_def_LostFocus
           End If
        Next tf
    End If
'    If tf = 0 Then i_def.ListIndex = I
 End If
'i_def.Enabled = False
End If
If PUB_TIPMOV = 10 And PUB_CP = "C" And Val(i_codcli.Text) = 1 Then
  fracli.Top = i_codcli.Top
  fracli.Left = i_codcli.Left + 300
  fracli.Visible = True
  If Trim(t_nombre.Text) = "" Then t_nombre.Text = "Clientes Varios"
  Azul t_nombre, t_nombre
  Exit Sub
End If

If PUB_CP = "P" Then
    If pub_signo_car <> 0 Then
      If InStr(1, wdia1, "-") = 0 Then i_dias.Text = Trim(wdia1)
      If Val(wdia1) <> 0 Then dias_bloqueo = Val(wdia1)     '      i_dias.Locked = True
     End If
     
     If Not cli_llave.EOF Then
       If cli_llave!CLI_programado = "S" Then
         i_ds.ListIndex = 0
       ElseIf cli_llave!CLI_programado = "D" Then
         i_ds.ListIndex = 1
       End If
     End If

End If

If fragas.Visible And LK_CODTRA = 2748 And i_cambio.Value = 0 Then
 cta_gas1.Text = Trim(wcta1)
 cta_gas2.Text = Trim(wcta2)
 imp_gas1.Text = ""
 imp_gas2.Text = ""
 If Val(i_subtotal.Text) <> 0 Then
   If Trim(cta_gas1.Text) <> "" Then imp_gas1.Text = redondea(Val(i_subtotal.Text))
   If Trim(cta_gas2.Text) <> "" Then imp_gas2.Text = "0.00"
 End If
 If Trim(wcta1) = "" Then imp_gas1.Text = ""
 If Trim(wcta2) = "" Then imp_gas2.Text = ""
 If pub_signo_ccm = 0 Then
  i_codban.Text = ""
  i_chenum.Text = ""
  i_cheser.Text = ""
 End If
 
End If
Return
Exit Sub
buscar:
If Left(i_codcli.Text, 2) = "++" Then
 var = Mid(i_codcli.Text, 3, Len(i_codcli.Text))
 numarchi = alta_vista_nombre(LV_CLI, var, PUB_CP, "D")
Else
 var = Mid(i_codcli.Text, 2, Len(i_codcli.Text))
 numarchi = alta_vista_nombre(LV_CLI, var, PUB_CP)
End If
If numarchi = 0 Then
  LV_CLI.Visible = False
  MsgBox "Alta Vista: No Existe .. Esta descripcion..", 48, Pub_Titulo
Else
  LV_CLI.Visible = True
  i_codcli.SetFocus
End If
loc_key = 1
Exit Sub

End Sub



Private Sub i_codven_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.

If KeyAscii = 27 Then
  i_codven.Text = ""
  LV_VEN.Visible = False
  Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo OJO
PUB_CODVEN = Val(FORMGEN.i_codven.Text)
On Error GoTo 0
If Len(i_codven.Text) = 0 Then
   Exit Sub
End If
If PUB_CODVEN <> 0 And IsNumeric(i_codven.Text) = True Then
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_VEN_LLAVE
   If ven_llave.EOF Then
    Azul FORMGEN.i_codven, FORMGEN.i_codven
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    FORMGEN.i_codven.SetFocus
    GoTo fin
   Else
      FORMGEN.i_nomven.Caption = ven_llave(2)
      NUMERO = FORMGEN.i_codven.WhatsThisHelpID
      avanza_campo

   End If
Else
On Error GoTo sigue
   If loc_key > 0 Then VALOR = UCase(LV_VEN.ListItems.Item(loc_key).Text)
   If Trim(UCase(i_codven.Text)) = Left(VALOR, Len(Trim(i_codven.Text))) Then
   Else
      Exit Sub
   End If
   
   If loc_key = 0 Then Exit Sub
   
   i_codven.Text = Trim(LV_VEN.ListItems.Item(loc_key).SubItems(1))
   PUB_CODVEN = Val(FORMGEN.i_codven.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_VEN_LLAVE
   FORMGEN.i_nomven.Caption = ven_llave(2)
   NUMERO = FORMGEN.i_codven.WhatsThisHelpID
   
   
   avanza_campo
End If
LV_VEN.Visible = False

fin:
sigue:
OJO:
End Sub



Private Sub i_concepto_GotFocus()
Azul i_concepto, i_concepto
End Sub

Private Sub i_concepto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
If LK_CODTRA = 2748 And fragas.Visible Then
  i_ds.SetFocus
 ' Azul cta_gas1, cta_gas1
  Exit Sub
End If

NUMERO = FORMGEN.i_concepto.WhatsThisHelpID

avanza_campo
fin:

End Sub
Private Sub i_def_GotFocus()
Dim pos1 As Integer
Dim CARAC As String
If FORMGEN.TRANS.Text = "" Then
   GoTo fin
End If
'SendKeys "%{DOWN}"
fin:
End Sub

Private Sub i_DEF_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = TABLA_TAG(tra_llave(2))
WS_INDICE_RETORNO = NUMERO
FORMGEN.Controls(NUMERO).SetFocus

fin:
End Sub


Private Sub i_def_LostFocus()
Dim pos1, CARAC, I, izq, der
If i_def.ListIndex = -1 Then Exit Sub
'Aqui se lee la contabilidad
pos1 = InStr(1, FORMGEN.i_def.List(i_def.ListIndex), ".", 1)
pos1 = pos1 - 1
''If LK_CODTRA <> 2748 Then
 CARAC = Mid(FORMGEN.i_def.List(i_def.ListIndex), 1, pos1)
''Else
'' CARAC = Right(FORMGEN.i_def.Text, 6)
''End If
WS_ESTADO = 0
PUB_SECUENCIA = Val(CARAC)
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_SUT_LLAVE
If SUT_LLAVE.EOF Then
   MsgBox "No existe Definicion en Sub_Transacciones... "
   GoTo fin
End If


pasa_def

   
If FORMGEN.Frame4.Visible = True Then
 '  FORMGEN.Frame2.Enabled = False
   grid_fac.ColWidth(2) = 0
   grid_fac.ColWidth(3) = 0
   grid_fac.ColWidth(5) = 0
   grid_fac.ColWidth(6) = 0
   grid_fac.ColWidth(7) = 0
   grid_fac.ColWidth(8) = 0
   grid_fac.ColWidth(9) = 0
   grid_fac.ColWidth(10) = 0
   grid_fac.ColWidth(11) = 0
   grid_fac.ColWidth(12) = 0
   
   If PUB_TIPMOV = 20 Then Frame2.Enabled = True
   
   If Len(Trim(SUT_LLAVE!SUT_JABAS)) <> 0 Then
      grid_fac.TextMatrix(0, 2) = SUT_LLAVE!SUT_JABAS
      grid_fac.ColWidth(2) = 600
   End If
   If Trim(SUT_LLAVE!SUT_DESFAC) = "VA" And PUB_TIPMOV = 10 Then
       lcodart(5).Caption = "Dsct.(valor)"
   End If
   
   
   If Nulo_Valors(SUT_LLAVE!sut_is) = "A" And LK_CODTRA <> 2414 Then
      grid_fac.ColWidth(12) = 1200
   End If
   If Nulo_Valor0(SUT_LLAVE!SUT_TIPMOV_REF) <> 0 And LK_CODTRA = 2401 Then
      grid_fac.ColWidth(42) = 1200
   End If
 
   
   If Nulo_Valors(SUT_LLAVE!SUT_UNIDADES) = "A" Then
      If LK_EMP = "3AA" Then
        grid_fac.ColWidth(5) = 700
        grid_fac.ColWidth(6) = 700
      Else
        grid_fac.ColWidth(5) = 1000
        grid_fac.ColWidth(6) = 1000
      End If
   End If
   
   If Nulo_Valors(SUT_LLAVE!SUT_PRECIO) = "1" Then
      grid_fac.ColWidth(6) = 0
   End If


   
   If PUB_TIPMOV = 20 Then
      grid_fac.ColWidth(6) = 1000
      grid_fac.ColWidth(8) = 600
'      grid_fac.ColWidth(7) = 900
      grid_fac.ColWidth(10) = 600
   End If
   
   If Trim(Nulo_Valors(SUT_LLAVE!SUT_DESCTO)) <> "" Then
      grid_fac.ColWidth(10) = 700
   End If
   
   I = 1
   GoSub destinos
   If grid_fac.ColWidth(2) <> 0 Then
      I = 2
      GoSub destinos
   End If
   
   If grid_fac.ColWidth(3) <> 0 Then
      I = 3
      GoSub destinos
   End If
   
   I = 4
   GoSub destinos
   
   If grid_fac.ColWidth(5) <> 0 Then
      I = 5
      GoSub destinos
   End If
   
   If grid_fac.ColWidth(6) <> 0 Then
      I = 6
      GoSub destinos
   End If
   If grid_fac.ColWidth(7) <> 0 Then
      I = 7
      GoSub destinos
   End If
   
   ' MNEW LK_CODTRA = 2403
   If LK_CODTRA = 2425 Then
      tab_derecha(5) = 1
   End If
   
   If grid_fac.ColWidth(8) <> 0 Then
      I = 8
      GoSub destinos
   End If
   If grid_fac.ColWidth(9) <> 0 Then
      I = 9
      GoSub destinos
   End If
   If grid_fac.ColWidth(10) <> 0 Then
      I = 10
      GoSub destinos
   End If
   If grid_fac.ColWidth(12) <> 0 Then
      I = 12
      GoSub destinos
   End If

   If SUT_LLAVE!SUT_PRECIO = "2" And PUB_TIPMOV = 10 Then
     tab_derecha(7) = 1
   End If
   
   grid_fac.ColWidth(6) = 1000
   grid_fac.ColWidth(7) = 900
   'If Nulo_Valors(SUT_LLAVE!SUT_mortal) = "1" Then
   '   grid_fac.ColWidth(9) = 2000
   'End If
   

   If LK_CODTRA = 2414 Then
      grid_fac.ColWidth(12) = 1200
   End If
   If LK_CODTRA = 2437 Then
       grid_fac.TextMatrix(0, 10) = "Ref."
   End If
   
   If Trim(SUT_LLAVE!SUT_DESCTO) <> "" Then
      grid_fac.ColWidth(10) = 700
   End If
   
   

End If
'BolFac_Click
   If PUB_TIPMOV = 93 Then
      grid_fac.ColWidth(10) = 400
   End If
   If PUB_TIPMOV = 101 Then
     grid_fac.ColWidth(34) = 800
     grid_fac.ColWidth(8) = 600
   End If
   

   If PUB_TIPMOV = 10 And SUT_LLAVE!SUT_SIGNO_CAR = 0 Then
      letche.Caption = "CHEQUES"
      letche.ToolTipText = "CH"
   End If
   If PUB_TIPMOV <> 0 And LK_CODTRA <> 2412 And LK_CODTRA <> 2410 And LK_CODTRA <> 2580 And LK_CODTRA <> 1122 And LK_CODTRA <> 2748 And LK_CODTRA <> 2402 And LK_CODTRA <> 2735 Then
        i_moneda.Visible = True
   Else
        i_moneda.Visible = False
   End If


If LK_CODTRA = 1401 And PUB_SECUENCIA = 10 And LK_EMP = "3AA" Then
  If i_fbg.ListCount >= 3 Then i_fbg.ListIndex = 2
End If
If LK_CODTRA = 2401 Then
      i_fbg.Clear
      i_fbg.AddItem "F"
      i_fbg.AddItem "B"
      If LK_CODUSU = "SUPERVISOR" Or LK_CODUSU = "ADMIN" Or LK_CODUSU = "OPER20" Or Left(LK_CODUSU, 3) = "BOT" Then
         i_fbg.AddItem "G"
         i_fbg.AddItem "P"
      End If
      If i_fbg.ListIndex = -1 Then i_fbg.ListIndex = 0
End If
If LK_CODTRA = 2741 Then
    If Trim(SUT_LLAVE!SUT_abreviado) = "PUB" Then
      i_fbg.Clear
      i_fbg.AddItem " "
    Else
      i_fbg.Clear
      i_fbg.AddItem "F"
      i_fbg.AddItem "B"
      i_fbg.AddItem " "
    End If
    i_fbg.ListIndex = 0
End If
If Trim(SUT_LLAVE!SUT_abreviado) = "CONS" And (PUB_SECUENCIA = 10 Or LK_EMP = "PAR") Then
  If i_fbg.ListCount > 0 Then i_fbg.ListIndex = 2
End If

If LK_CODTRA = 2748 And LK_EMP = "3AA" And i_cambio.Value = 0 Then
  i_dias.Text = Val(Mid(i_def.Text, 15, Len(i_def.Text)))
  fragas.Visible = True
  For fila = 0 To i_codsunat.ListCount - 1
      i_codsunat.ListIndex = fila
      If Val(Right(i_codsunat.Text, 6)) = 1 Then
          Exit For
      End If
  Next fila
End If
'i_dias.Locked = False
If LK_CODTRA = 2401 Then
  If SUT_LLAVE!SUT_SIGNO_CAR <> 0 And SUT_LLAVE!sut_TIPDOC <> "CC" Then
'     i_dias.Text = Trim(Left(SUT_LLAVE!sut_descripcion, 2))
  Else
     i_dias.Text = ""
  End If
'
End If


If LK_CODTRA = 2738 Or LK_CODTRA = 2735 Or LK_CODTRA = 5714 Then
  WS_ESTADO = 99
  FORMGEN.i_numser.Text = ""
End If
If LK_CODTRA = 5714 Then
 If pub_signo_ccm <> 1 Then
    i_importe_amort.Visible = False
    Tipo_Cambio.Visible = False
 Else
    i_importe_amort.Visible = True
    Tipo_Cambio.Visible = True
 End If
 FORMGEN.i_numser.Text = ""
 If SUT_LLAVE!SUT_signo_ccm = 1 Then
  WS_ESTADO = 98
 Else
  WS_ESTADO = 99
 End If
End If
If LK_CODTRA = 2720 Then
  WS_ESTADO = 98
  FORMGEN.i_numser.Text = ""
End If
If LK_CODTRA = 5318 Then
  WS_ESTADO = 97
  FORMGEN.i_numser.Text = ""
End If

If LK_CODTRA = 2412 Or LK_CODTRA = 2774 Then
  If LK_CODTRA = 2774 Then
    FORMGEN.CmdEsp.Left = 3000
  Else
    FORMGEN.CmdEsp.Left = 700
  End If
  FORMGEN.CmdEsp.Top = 1700
  FORMGEN.CmdEsp.Visible = True
End If


If LK_CODTRA = 2401 Or LK_CODTRA = 2412 Then GoTo SALT
If i_codban.Visible = True And LK_CODTRA <> 2725 Then GoTo SALT

i_numser.Text = ""
PSCNT_LLAVE.rdoParameters(0) = LK_CODCIA
PSCNT_LLAVE.rdoParameters(1) = LK_CODTRA
PSCNT_LLAVE.rdoParameters(2) = SUT_LLAVE!sut_secuencia
cnt_llave.Requery
If cnt_llave.EOF Then
   'CAMTEX
   If LK_EMP <> "CAM" And LK_EMP <> "HER" And LK_EMP <> "PIU" Then
     MsgBox "No ha definido Serie y Correlativo ..."
   End If
   GoTo SALT
End If


   If PUB_TIPMOV <> 0 Then
      FORMGEN.i_numser.Text = Nulo_Valor0(cnt_llave!cnt_serie)
      If Nulo_Valor0(cnt_llave!cnt_serie) > 0 Then WS_ESTADO = 1
   Else
      If SUT_LLAVE!SUT_SIGNO_CAR = -1 Then
         If Nulo_Valor0(cnt_llave!cnt_serie_RECIBO) > 0 Then WS_ESTADO = 2
         FORMGEN.i_numser.Text = Nulo_Valor0(cnt_llave!cnt_serie_RECIBO)
         PUB_NUMSER = Nulo_Valor0(cnt_llave!cnt_serie_RECIBO)
      ElseIf SUT_LLAVE!SUT_signo_caja <> 0 Then
         If Nulo_Valor0(cnt_llave!cnt_serie_CAJA) > 0 Then WS_ESTADO = 3
         FORMGEN.i_numser.Text = Nulo_Valor0(cnt_llave!cnt_serie_CAJA)
         PUB_NUMSER = Nulo_Valor0(cnt_llave!cnt_serie_CAJA)
      End If
   End If

SALT:

'If LK_CODTRA = 1401 Then
'   PUB_NUMSER = Nulo_Valor0(par_llave!PAR_SER_KARDEX)
'   FORMGEN.i_numser.Text = PUB_NUMSER
'End If
'If LK_CODTRA = 2414 Then
'   PUB_NUMSER = Nulo_Valor0(par_llave!par_serie)
'   FORMGEN.i_numser.Text = PUB_NUMSER
'End If
If LK_CODTRA = 1409 Then i_fbg.ListIndex = 0
If NJ.Visible Then
  If Trim(NJ.Text) = "N" Or Trim(NJ.Text) = "S" Then
  Else
    NJ.Text = "N"
  End If
End If
If LK_FLAG_SOS = "A" And NJ.Visible And LK_ACTIVA = "A" Then
 NJ.Text = "N"
 NJ.Locked = True
Else
 NJ.Locked = False
End If

If LK_EMP = "CAM" And LK_CODTRA = 1401 And PUB_SECUENCIA = 20 Then
   grid_fac.TextMatrix(0, 10) = "O/P"
End If
Exit Sub

destinos:
   izq = I - 1
   der = I + 1
   tab_derecha(I) = 1
   tab_izquierda(I) = 1
   Do Until der > 12 And izq < 1
      If der < 13 Then
      If grid_fac.ColWidth(der) <> 0 Then
         tab_derecha(I) = der
         der = 13
      End If
      End If
      
      If izq > 0 Then
         If grid_fac.ColWidth(izq) <> 0 Then
            tab_izquierda(I) = izq
            izq = 0
         End If
      End If
       
      der = der + 1
      izq = izq - 1
   Loop

  Return
fin:




End Sub






Private Sub i_descto_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_descto, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If
i_descto_DblClick
PUB_DESCTO = Val(FORMGEN.i_descto.Text)
FORMGEN.grid_fac.Row = 2
FORMGEN.grid_fac.COL = 1
FORMGEN.grid_fac.SetFocus


fin:

End Sub
Private Sub i_dias_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 i_dias.Text = ""
 Exit Sub
End If
If KeyAscii = 13 Then GoTo OTRO
SOLO_ENTERO KeyAscii

Dim FECHA_DIA As Date
Dim fecha_vcto As Variant
Dim pub_mensaje As String
If KeyAscii <> 13 Then
   GoTo fin
End If
'If ENTERO(Val(FORMGEN.i_dias.text)) = False Then
'   Azul FORMGEN.i_dias, FORMGEN.i_dias
'   GoTo fin
'End If

If LK_CODTRA = 2401 Then
If Val(i_dias.Text) > 30 Then
   pub_mensaje = " ¿Esta seguro del numero de Dias... ?"
     Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then   ' El usuario eligió
     i_dias.SetFocus
     Exit Sub
  End If
End If
End If
OTRO:
NUMERO = FORMGEN.i_dias.WhatsThisHelpID
avanza_campo
fin:

End Sub
Private Sub i_dias_LostFocus()
wtemporal = ""
FORMGEN.i_fecha_vcto.Text = Str(DateAdd("d", Val(FORMGEN.i_dias.Text), i_fecha_compra.Text))
If LK_CODTRA = 1401 Then
  If Weekday(FORMGEN.i_fecha_vcto.Text) = 1 Or Weekday(FORMGEN.i_fecha_vcto.Text) = 7 Then
     MsgBox "El Día no es laborable : " & UCase(Format(FORMGEN.i_fecha_vcto.Text, "dddd dd mmmm , yyyy")), 48, Pub_Titulo
  End If
End If

End Sub

Private Sub i_fbg_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = FORMGEN.i_fbg.WhatsThisHelpID

avanza_campo
fin:

End Sub


Private Sub i_fecha_vcto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
If IsDate(i_fecha_vcto) = False Then
   MsgBox "FECHA NO VALIDA ...", 48, Pub_Titulo
   GoTo fin
End If
If LK_CODTRA = 1111 And Diario.Visible Then
  Diario.SetFocus
  Exit Sub
End If
NUMERO = FORMGEN.i_fecha_vcto.WhatsThisHelpID
avanza_campo
fin:

End Sub
Private Sub i_gastos_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_gastos, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo SALE
PUB_GASTOS = Val(FORMGEN.i_gastos.Text)
grid_fac.SetFocus
On Error GoTo 0


fin:
Exit Sub
SALE:
Azul i_gastos, i_gastos
End Sub

Private Sub i_importe_amort_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_importe_amort, KeyAscii

If KeyAscii <> 13 Then Exit Sub
If LK_CODTRA <> 2770 Then
   NUMERO = FORMGEN.i_importe_amort.WhatsThisHelpID
   avanza_campo
End If




fin:

End Sub

Private Sub i_importe_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_importe, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If
If LK_CODTRA = 5350 Then
  i_importe_amort.Text = i_importe.Text
End If
If LK_CODTRA = 5714 And Val(i_importe.Text) = 0 Then
  Exit Sub
End If
If LK_CODTRA = 5714 Then
If Not Tipo_Cambio.Visible Then
   FORMGEN.i_fecha_can.SetFocus
   FORMGEN.i_fecha_can.SelStart = 0
   FORMGEN.i_fecha_can.SelLength = 2
   'Azul2 i_fecha_can, i_fecha_can
End If
End If



NUMERO = FORMGEN.i_importe.WhatsThisHelpID
avanza_campo
 
 
 
fin:

End Sub


Private Sub i_gastos_fijos_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_gastos_fijos, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_gastos_fijos.WhatsThisHelpID
avanza_campo
fin:

End Sub




Private Sub i_limcre_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_limcre, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If

If i_limcre.Text = "" Then Exit Sub
NUMERO = FORMGEN.i_limcre.WhatsThisHelpID
avanza_campo
fin:
'End Sub



'salta:

'FORMGEN.i_subtotal.Enabled = False
'FORMGEN.i_gastos.Enabled = False
'FORMGEN.i_descto.Enabled = False
'FORMGEN.i_impto.Enabled = False
'FORMGEN.i_neto.Enabled = False
'FORMGEN.i_dias.Enabled = False
'FORMGEN.i_fecha_vcto.Enabled = False
'FORMGEN.grid_fac.SetFocus
'FORMGEN.grid_fac.Row = 2
'FORMGEN.grid_fac.Col = 1

'fin:

End Sub



Private Sub i_num_ini_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_num_ini.WhatsThisHelpID

avanza_campo
fin:

End Sub

Private Sub i_num_oper_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_num_oper.WhatsThisHelpID
avanza_campo
fin:

End Sub

Private Sub i_numdoc_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_numdoc.WhatsThisHelpID
avanza_campo
fin:

End Sub

Private Sub I_NUMDOC_R_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_numdoc_r.WhatsThisHelpID
avanza_campo


fin:

End Sub
Private Sub i_numfac_c_KeyPress(KeyAscii As Integer)
Dim rec_fila As Integer
Dim wven  As String

SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If LK_EMP = "HER" And LK_CODTRA = 2401 Then GoTo ver_cambio
If LK_CODTRA = 2748 Or LK_CODTRA = 1401 Then GoTo chequeadocu
If LK_CODTRA = 1111 Then GoTo SALTA
If LK_CODTRA = 2412 Then GoTo SALTA
If LK_CODTRA = 2741 Then GoTo SALTA
If LK_CODTRA = 2410 Then GoTo SALTA
If LK_CODTRA = 1405 Then GoTo SALTA
If LK_CODTRA = 2437 Then GoTo SALTA

If LK_CODTRA = 2720 Then GoTo PRO_2720
If LK_CODTRA = 1122 Then If gridl.Visible = False Then GoTo SALTA

If LK_CODTRA = 2750 Then GoTo llamar_compra
If PUB_TIPMOV = 20 Then GoTo SALTA
   SQ_OPER = 4
   pu_codcia = LK_CODCIA
   PU_FBG = i_fbg.Text
   PUB_TIPDOC = SUT_LLAVE!sut_TIPDOC
   PUB_NUMSER = Val(i_numser_c.Text)
   PUB_NUMFAC = Val(i_numfac_c.Text)
   LEER_CAR_LLAVE
   If car_far.EOF Then
      MsgBox "Documento no Existe. o es Otro Tipo de Documento.", 48, Pub_Titulo
      i_numser_c.SetFocus
      Exit Sub
   End If

   Do Until car_far.EOF
      If car_far!car_importe <> 0 And car_far!CAR_cp = "C" Then Exit Do
      car_far.MoveNext
   Loop
   If car_far.EOF Then
      loc_key = 777
      MsgBox "Documento Cancelado..." & PU_FBG & "/" & PUB_NUMSER & "-" & PUB_NUMFAC, 48, Pub_Titulo
      i_numser_c.SetFocus
      Exit Sub
   End If
   If Trim(car_far!CAR_MONEDA) <> Trim(i_ds.Text) Then
      MsgBox "El Documento No Pertenece al Tipo de Moneda.", 48, Pub_Titulo
      Exit Sub
   End If
   If Val(i_codven.Text) = 0 And LK_CODTRA = 2770 And LK_EMP = "HER" Then
      MsgBox "Falta Cod. Vendedor / Cobrador ", 48, Pub_Titulo
      If i_codven.Visible Then i_codven.SetFocus
      Exit Sub
   End If
   
   If gridl.Visible = False Then gridl.Visible = True
   If textovarl.Visible = False Then textovarl.Visible = True
   If textovar_canje.Visible = False Then textovar_canje.Visible = True
   
   If LK_CODTRA = 2770 Or LK_CODTRA = 2774 Then
     For rec_fila = 2 To gridl.Rows - 1
        'gridl.TextMatrix(rec_fila, 15) = cli_llave!cli_codclie
        If car_far!CAR_cp = "P" Then
          If (Val(gridl.TextMatrix(rec_fila, 2)) = Nulo_Valor0(car_far!car_numser_c)) And (Val(gridl.TextMatrix(rec_fila, 3)) = Nulo_Valor0(car_far!car_numfac_c)) Then
             MsgBox "Documento esta en la Lista", 48, Pub_Titulo
             Exit Sub
          End If
        Else
          If (Val(gridl.TextMatrix(rec_fila, 2)) = Nulo_Valor0(car_far!car_NUMSER)) And (Val(gridl.TextMatrix(rec_fila, 3)) = Nulo_Valor0(car_far!car_NUMFAC)) Then
            ' MsgBox "Documento esta en la Lista", 48, Pub_Titulo
            ' Exit Sub
          End If
        End If
      Next rec_fila
   End If

   textovar_canje.Width = 30
   textovar_canje.Height = 30
   
   pasa_cabeza_canje
   pasa_cabeza_L

   pu_codclie = car_far!CAR_CODCLIE
   pu_codcia = car_far!car_codcia
   pu_cp = car_far!CAR_cp
   SQ_OPER = 1
   LEER_CLI_LLAVE
      
      FILAX = gridl.Rows - 1
      
      If gridl.TextMatrix(FILAX, 0) = "" Then
      Else
         gridl.Rows = gridl.Rows + 1
         FILAX = gridl.Rows - 1
      End If
      
      gridl.RowHeight(FILAX) = 350
      gridl.TextMatrix(FILAX, 15) = cli_llave!CLI_CODCLIE
      gridl.TextMatrix(FILAX, 0) = cli_llave!cli_nombre
      gridl.TextMatrix(FILAX, 13) = car_far!car_TIPDOC
      gridl.TextMatrix(FILAX, 1) = Nulo_Valors(car_far!car_FBG)
      If car_far!CAR_cp = "P" Then
          gridl.TextMatrix(FILAX, 2) = Nulo_Valor0(car_far!car_numser_c)
          gridl.TextMatrix(FILAX, 3) = Nulo_Valor0(car_far!car_numfac_c)
       Else
          gridl.TextMatrix(FILAX, 2) = Nulo_Valor0(car_far!car_NUMSER)
          gridl.TextMatrix(FILAX, 3) = Nulo_Valor0(car_far!car_NUMFAC)
       End If
       
       If car_far!CAR_MONEDA = "S" Then
          gridl.TextMatrix(FILAX, 4) = "S/." & car_far!car_importe
       Else
          gridl.TextMatrix(FILAX, 4) = "  $" & car_far!car_importe
       End If

       
       gridl.TextMatrix(FILAX, 5) = ""
       gridl.TextMatrix(FILAX, 6) = Format(car_far!car_fecha_vcto, "dd/mm/yy")
       gridl.TextMatrix(FILAX, 7) = Val(i_codven.Text)
       gridl.TextMatrix(FILAX, 9) = car_far!CAR_IMP_INI
       gridl.TextMatrix(FILAX, 8) = car_far!CAR_FECHA_INGR
       gridl.TextMatrix(FILAX, 10) = car_far!CAR_cp
       gridl.TextMatrix(FILAX, 11) = car_far!car_SERDOC
       gridl.TextMatrix(FILAX, 12) = car_far!car_NUMDOC
       gridl.TextMatrix(FILAX, 14) = car_far!car_codcia
       gridl.TextMatrix(FILAX, 18) = car_far!CAR_codven
       gridl.TextMatrix(FILAX, 22) = Format(LK_FECHA_DIA, "dd/mm/yyyy")
       gridl.TextMatrix(FILAX, 24) = car_far!CAR_MONEDA
       gridl.TextMatrix(FILAX, 25) = car_far!car_importe
       gridl.TextMatrix(FILAX, 26) = car_far!car_fecha_vcto
       
       If gridl.Visible = False Then gridl.Visible = True
       gridl.SetFocus
       flag_salto = 1
       gridl.Row = FILAX
       flag_salto = 0
       gridl.COL = 5


fin:
Exit Sub
SALTA:
   NUMERO = FORMGEN.i_numfac_c.WhatsThisHelpID
   avanza_campo

Exit Sub
chequeadocu:
If Val(i_numser_c.Text) = 0 And Val(i_numfac_c.Text) = 0 Then GoTo salta1
SQ_OPER = 8
PU_TIPMOV = PUB_TIPMOV
pu_codcia = LK_CODCIA
PU_NUMSER = Val(i_numser_c.Text)
If LK_CODTRA = 2748 Then
 PU_FBG = "K"
 PU_FBG2 = " "
Else
 PU_FBG = " "
 PU_FBG2 = " "
End If
PU_NUMFAC = Val(i_numfac_c.Text)
LEER_FAR_LLAVE
If Not far_menor4.EOF Then
  Do Until far_menor4.EOF
  If Val(far_menor4!far_codclie) = Val(i_codcli.Text) Then
     MsgBox " Documento ya Emitido, Secuencia de kardex Nro :" & Format(far_menor4!far_numfac, "000"), 48, Pub_Titulo
     i_numfac_c.Text = ""
     i_numser_c.Text = ""
     Azul i_numser_c, i_numser_c
     Exit Sub
  End If
  far_menor4.MoveNext
  Loop
End If
salta1:
NUMERO = FORMGEN.i_numfac_c.WhatsThisHelpID
avanza_campo
Exit Sub
llamar_compra:

SQ_OPER = 8
PU_TIPMOV = 20
pu_codcia = LK_CODCIA
PU_NUMSER = Val(i_numser_c.Text)
PU_FBG = " "
PU_FBG2 = " "
PU_NUMFAC = Val(i_numfac_c.Text)
LEER_FAR_LLAVE
If far_menor4.EOF Then
  MsgBox "Documento no Existe...", 48, Pub_Titulo
  Exit Sub
End If
far_menor4.MoveLast
i_dias.Text = Format(far_menor4!FAR_DIAS, "0")
i_impto2.Text = Format(far_menor4!far_IMPTO, "0.00")
i_bruto2.Text = Format(far_menor4!FAR_BRUTO, "0.00")
i_importe_amort.Text = Format(Val(far_menor4!FAR_BRUTO) + Val(far_menor4!far_IMPTO), "0.00")
i_fecha_compra.Text = Format(far_menor4!FAR_fecha_compra, "dd/mm/yy")
If far_menor4!FAR_MONEDA = "S" Then
 i_ds.ListIndex = 0
Else
 i_ds.ListIndex = 1
End If
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_FECHA = far_menor4!FAR_fecha
LEER_ALL_LLAVE
Do Until all_llave.EOF
If Val(all_llave!ALL_NUMOPER) = Val(far_menor4!FAR_NUMOPER) Then GoTo sigue0
all_llave.MoveNext
Loop
MsgBox "Documento no Existe.", 48, Pub_Titulo
i_numser_c.SetFocus
Exit Sub
sigue0:
SQ_OPER = 4
pu_codcia = LK_CODCIA
PU_FBG = far_menor4!FAR_FBG
PUB_TIPDOC = far_menor4!FAR_TIPDOC
PUB_NUMSER = far_menor4!far_numser
PUB_NUMFAC = far_menor4!far_numfac
'Print far_menor4!far_FECHA
LEER_CAR_LLAVE
If car_far.EOF Then
   MsgBox "Documento no Existe.", 48, Pub_Titulo
   i_numser_c.SetFocus
   Exit Sub
End If
Do Until car_far.EOF
 If car_far!car_importe <> 0 Then GoTo sigue
 
 car_far.MoveNext
Loop
 MsgBox "Documento Tiene Amortizaciones .. no procede ", 48, Pub_Titulo
Exit Sub
sigue:
If car_far!car_importe <> car_far!CAR_IMP_INI Then
  MsgBox "Documento Tiene Amortizaciones .. no procede ", 48, Pub_Titulo
  Exit Sub
End If
SQ_OPER = 1
pu_cp = car_far!CAR_cp
pu_codclie = car_far!CAR_CODCLIE
pu_codcia = car_far!car_codcia
PUB_TIPDOC = car_far!car_TIPDOC
PUB_FECHA = car_far!CAR_FECHA_INGR
PUB_NUM_OPER = car_far!CAR_NUMOPER
LEER_CAA_LLAVE
If caa_histo.EOF Then
   MsgBox "Documento no Existe.", 48, Pub_Titulo
   i_numser_c.SetFocus
   Exit Sub
End If
Do Until caa_histo.EOF
If caa_histo!CAA_ESTADO <> "E" Then GoTo sigue2
caa_histo.MoveNext
Loop
 MsgBox "Documento Tiene Amortizaciones .. no procede ", 48, Pub_Titulo
 Exit Sub
sigue2:


NUMERO = FORMGEN.i_numfac_c.WhatsThisHelpID
avanza_campo
Exit Sub
ver_cambio:
If Val(i_numfac_c.Text) = 0 Then GoTo PS
Dim walterno As String * 10
Dim wnomarti As String * 20
Dim wdescri  As String * 10
Dim wcantida  As String * 13
Dim WTIPO  As String * 10
PS_CAMBIO(0) = LK_CODCIA
PS_CAMBIO(1) = 102
PS_CAMBIO(2) = Val(i_numfac_c.Text)
PS_CAMBIO(3) = Val(i_numser_c.Text)
far_cambio.Requery
If far_cambio.EOF Then
 MsgBox "No Existe Doc. Interno", 48, Pub_Titulo
 Exit Sub
End If

pub_cadena = ""
pu_codclie = far_cambio!far_codclie
pu_cp = far_cambio!FAR_cp
pu_codcia = LK_CODCIA
SQ_OPER = 1
LEER_CLI_LLAVE

Do Until far_cambio.EOF
PUB_KEY = far_cambio!far_codart
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
walterno = art_LLAVE!art_alterno
wnomarti = art_LLAVE!art_nombre
wdescri = far_cambio!far_descri
wcantida = Format(far_cambio!far_cantidad, "0.00")
WTIPO = far_cambio!far_signo_arm
pub_cadena = pub_cadena + walterno & " " & wnomarti & " " & wdescri & " " & wcantida & " " & WTIPO & Chr(13)
wven = far_cambio!FAR_CODVEN
far_cambio.MoveNext
Loop
pub_mensaje = "Cliente: " & Trim(cli_llave!cli_nombre) & Chr(13) & "Detalle.-" & Chr(13) & pub_cadena & "Vendedor:" & wven & Chr(13) & Chr(13) & "Confirmar ...?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   i_numser_c.SetFocus
   Exit Sub
End If


PS:
NUMERO = FORMGEN.i_numfac_c.WhatsThisHelpID
avanza_campo
Exit Sub

PRO_2720:
NUMERO = FORMGEN.i_numfac_c.WhatsThisHelpID
avanza_campo

End Sub

Private Sub i_numfac_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If LK_CODTRA = 2401 Or LK_CODTRA = 2402 Then
  If i_fecha_compra.Visible Then Azul2 i_fecha_compra, i_fecha_compra
Else
NUMERO = FORMGEN.i_numfac.WhatsThisHelpID
avanza_campo
End If
fin:
End Sub



Private Sub i_numguia_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If LK_CODTRA = 2748 And pub_signo_ccm <> 0 Then
'  FORMGEN.i_fecha_oper.SetFocus
'  FORMGEN.i_fecha_oper.SelStart = 0
'  FORMGEN.i_fecha_oper.SelLength = 2
' Exit Sub
End If
If LK_CODTRA = 2748 And pub_signo_ccm = 0 And FORMGEN.i_TEXTONCRE.Visible Then
  FORMGEN.i_TEXTONCRE.SetFocus
  FORMGEN.i_TEXTONCRE.SelStart = 0
  FORMGEN.i_TEXTONCRE.SelLength = 2
  Exit Sub
End If
NUMERO = FORMGEN.i_numguia.WhatsThisHelpID
avanza_campo
fin:

End Sub

Private Sub i_numser_c_GotFocus()
'   Azul i_numser_c, i_numser_c
If LK_CODTRA = 2770 Then
  i_numser_c.MaxLength = 3
Else
  i_numser_c.MaxLength = 0
End If

End Sub

Private Sub i_numser_c_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = FORMGEN.i_numser_c.WhatsThisHelpID
avanza_campo
fin:

End Sub

Private Sub i_numser_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_numser.WhatsThisHelpID

avanza_campo
fin:

End Sub


Private Sub i_numser_r_KeyPress(KeyAscii As Integer)
Dim respuesta
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If

If LK_CODTRA = 2455 Then
   LKCHEK = False
   pub_total_2455 = 0
   Do Until LKCHEK = True
      llena_numfac
      DoEvents
      i_numdoc_r.Text = Val(i_numdoc_r.Text) + 1
      If LKCHEK = False Then
         pub_total_2455 = pub_total_2455 + PUB_SUBTOTAL
         grabar_Click
      End If
   Loop
   pub_mensaje = " ¿Es la venta total OK?" & pub_total_2455
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then
      MsgBox "LLAMAR A COMPUTO IMMEDIATAMENTE.... ERROR GRAVE..."
      MsgBox "LLAMAR A COMPUTO IMMEDIATAMENTE.... ERROR GRAVE..."
   End If
End If
   
NUMERO = FORMGEN.i_numser_r.WhatsThisHelpID
avanza_campo
   

fin:

End Sub

Private Sub i_precio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = FORMGEN.i_precio.WhatsThisHelpID
avanza_campo
fin:

End Sub



Private Sub i_subtotal_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL i_subtotal, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If

fin:

End Sub
Private Sub i_subtotal_LostFocus()
PUB_SUBTOTAL_BAK = Val(FORMGEN.i_subtotal.Text)
End Sub
Private Sub i_tipdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_tipdoc.WhatsThisHelpID

avanza_campo
fin:

End Sub





Private Sub i_codusu_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_CODUSU.WhatsThisHelpID
avanza_campo
fin:

End Sub

Private Sub i_tipo_bloq_act1_KeyPress(KeyAscii As Integer)

NUMERO = FORMGEN.i_tipo_bloq_act1.WhatsThisHelpID
avanza_campo
End Sub

Private Sub i_tipo_bloq_act2_KeyPress(KeyAscii As Integer)

NUMERO = FORMGEN.i_tipo_bloq_act2.WhatsThisHelpID
avanza_campo
End Sub

Private Sub i_tipo_bloq_act3_KeyPress(KeyAscii As Integer)

NUMERO = FORMGEN.i_tipo_bloq_act3.WhatsThisHelpID
avanza_campo

End Sub

Private Sub i_tipo_bloq_act4_KeyPress(KeyAscii As Integer)

NUMERO = FORMGEN.i_tipo_bloq_act4.WhatsThisHelpID
avanza_campo
End Sub
Private Sub LisTransa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  FORMGEN.TRANS.Text = Trim(Left(LisTransa.Text, 7))
  FORMGEN.LisTransa.Visible = False
  FORMGEN.TRANS.SetFocus
ElseIf KeyAscii = 27 Then
  KeyAscii = 0
  FORMGEN.LisTransa.Visible = False
  FORMGEN.TRANS.SetFocus
End If

End Sub

Private Sub LisTransa_LostFocus()
  FORMGEN.LisTransa.Visible = False
  FORMGEN.TRANS.SetFocus
End Sub

Private Sub LV_ART_DblClick()
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 loc_key = LV_ART.SelectedItem.Index
 TEXTOVAR.Text = Trim(LV_ART.ListItems.Item(loc_key).Text) & " "
 textovar_KeyPress 13
Else
 loc_key = LV_ART.SelectedItem.Index
 TEXTOVAR.Text = Trim(LV_ART.ListItems.Item(loc_key).Text) & " "
 textovar_KeyPress 13
End If

End Sub

Private Sub LV_ART_GotFocus()
If loc_key <> 0 Then
 Set LV_ART.SelectedItem = LV_ART.ListItems(loc_key)
 LV_ART.ListItems.Item(loc_key).Selected = True
 LV_ART.ListItems.Item(loc_key).EnsureVisible
End If

End Sub


Private Sub LV_ART_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = LV_ART.SelectedItem.Index
 If LK_CODTRA <> 2401 Then GoTo SALTA2401
  If fraprecios.Visible = False Then
   fraprecios.Top = LV_ART.Top + 3230
   fraprecios.Visible = True
  End If
  
  lbltitulostock.Caption = Left(lk_descrip_cias, 10)
  lbltitulostock2.Caption = Right(lk_descrip_cias, 10)
  lpstock2.Caption = Int(LV_ART.ListItems.Item(loc_key).SubItems(12))
  lpstock2.Caption = Format(lpstock2.Caption, "##,##0.00")
  lpstocku2.Caption = Format(Val((LV_ART.ListItems.Item(loc_key).SubItems(13))) - (Val(lpstock2.Caption) * Val(LV_ART.ListItems.Item(loc_key).SubItems(9))), "##,##0.000")
  
  lpstock3.Caption = Int(LV_ART.ListItems.Item(loc_key).SubItems(14))
  lpstock3.Caption = Format(lpstock3.Caption, "##,##0.00")
  lpstocku3.Caption = Format(Val((LV_ART.ListItems.Item(loc_key).SubItems(15))) - (Val(lpstock3.Caption) * Val(LV_ART.ListItems.Item(loc_key).SubItems(9))), "##,##0.000")
  
  
  lpnombre.Caption = Trim(LV_ART.ListItems.Item(loc_key))
  lpunidad.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(2))
  lpstock.Caption = Int(LV_ART.ListItems.Item(loc_key).SubItems(3))
  lpstocku.Caption = Format(Val((LV_ART.ListItems.Item(loc_key).SubItems(10))) - (Val(lpstock.Caption) * Val(LV_ART.ListItems.Item(loc_key).SubItems(9))), "##,##0.000")
  lpstock.Caption = Format(lpstock.Caption, "##,##0.00")
  lpp1.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(4))
  lpp2.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(5))
  lpp3.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(6))
  lpp4.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(7))
  lpp5.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(8))
  lpp6.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(11))
  
  lbllote.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(17))
  
  
  If Val(LV_ART.ListItems.Item(loc_key).SubItems(18)) = 1 Then
    fraprecios.BackColor = QBColor(3)
  ElseIf Val(LV_ART.ListItems.Item(loc_key).SubItems(16)) = 1 Then
    fraprecios.BackColor = QBColor(6)
  Else
    fraprecios.BackColor = QBColor(8)
  End If
  
SALTA2401:

 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  TEXTOVAR.Text = Trim(LV_ART.ListItems.Item(loc_key).Text) & " "
 Else
  TEXTOVAR.Text = Trim(LV_ART.ListItems.Item(loc_key).Text) & " "
 End If
End If

End Sub

Private Sub LV_ART_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 If TEXTOVAR.Visible Then
   TEXTOVAR.Text = ""
   TEXTOVAR.SetFocus
 End If
End If

If KeyAscii = 13 Then
 If TEXTOVAR.Visible Then
   textovar_KeyPress 13
 End If

End If
End Sub

Private Sub LV_ART_LostFocus()
LV_ART.Visible = False
fraprecios.Visible = False
End Sub

Private Sub LV_CCM_DblClick()
loc_key = LV_CCM.SelectedItem.Index
i_codban.Text = Trim(LV_CCM.ListItems.Item(loc_key).Text) & " "
i_codban_KeyPress 13

End Sub

Private Sub LV_CCM_GotFocus()
If loc_key <> 0 Then
 Set LV_CCM.SelectedItem = LV_CCM.ListItems(loc_key)
 LV_CCM.ListItems.Item(loc_key).Selected = True
 LV_CCM.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub LV_CCM_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
  loc_key = LV_CCM.SelectedItem.Index
  i_codban.Text = Trim(LV_CCM.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub LV_CCM_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 If i_codban.Visible Then
    FORMGEN.LV_CCM.Visible = False
    FORMGEN.i_codban.Text = ""
    FORMGEN.i_codban.SetFocus
 End If
End If
If KeyAscii = 13 Then
 If i_codban.Visible Then
    i_codban_KeyPress 13
 End If
End If

End Sub

Private Sub LV_CCM_LostFocus()
FORMGEN.LV_CCM.Visible = False
End Sub

Private Sub LV_CLI_DblClick()
loc_key = LV_CLI.SelectedItem.Index
i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
i_codcli_KeyPress 13

End Sub

Private Sub LV_CLI_GotFocus()
If loc_key <> 0 Then
 Set LV_CLI.SelectedItem = LV_CLI.ListItems(loc_key)
 LV_CLI.ListItems.Item(loc_key).Selected = True
 LV_CLI.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub LV_CLI_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
  loc_key = LV_CLI.SelectedItem.Index
  i_codcli.Text = Trim(LV_CLI.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub LV_CLI_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 If i_codcli.Visible Then
    FORMGEN.LV_CLI.Visible = False
    FORMGEN.i_codcli.Text = ""
    FORMGEN.i_codcli.SetFocus
 End If
End If
If KeyAscii = 13 Then
 If i_codcli.Visible Then
    i_codcli_KeyPress 13
 End If
End If
End Sub

Private Sub LV_CLI_LostFocus()
FORMGEN.LV_CLI.Visible = False
End Sub

Private Sub LV_VEN_DblClick()
loc_key = LV_VEN.SelectedItem.Index
i_codven.Text = Trim(LV_VEN.ListItems.Item(loc_key).Text) & " "
i_codven_KeyPress 13
End Sub

Private Sub LV_VEN_GotFocus()
If loc_key <> 0 Then
 Set LV_VEN.SelectedItem = LV_VEN.ListItems(loc_key)
 LV_VEN.ListItems.Item(loc_key).Selected = True
 LV_VEN.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub LV_VEN_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
  loc_key = LV_VEN.SelectedItem.Index
  i_codven.Text = Trim(LV_VEN.ListItems.Item(loc_key).Text) & " "
End If
End Sub

Private Sub LV_VEN_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 If i_codven.Visible Then
    i_codven.Text = ""
    i_codven.SetFocus
 End If
End If
If KeyAscii = 13 Then
 If i_codven.Visible Then
    i_codven_KeyPress 13
 End If
End If

End Sub

Private Sub LV_VEN_LostFocus()
LV_VEN.Visible = False
End Sub

Private Sub MOSTRAR_Click()

End Sub

Private Sub Option1_Click()

End Sub

Private Sub msgrid_DblClick()
Dim I As Integer
Dim J As Integer
If grid_fac.Rows <> 3 Then
  MsgBox "Documento de estar sin productos para acceder a los temporales", 48, Pub_Titulo
Else
  For I = 1 To msgrid.Rows - 1
   If I > 1 Then grid_fac.AddItem ""
   For J = 0 To msgrid.Cols - 1
      grid_fac.TextMatrix(I + 1, J) = msgrid.TextMatrix(I, J)
   Next J
  Next I
End If
msgrid.Visible = False
End Sub

Private Sub msgrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then msgrid.Visible = False
End Sub

Private Sub msgrid_LostFocus()
msgrid.Visible = False
End Sub

Private Sub NJ_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 83 And KeyAscii <> 8 And KeyAscii <> 78 And KeyAscii <> 13 Then
 KeyAscii = 0
 Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.NJ.WhatsThisHelpID
avanza_campo
fin:


End Sub

Private Sub NJ_LostFocus()

If NJ.Visible And LK_ACTIVA = "A" Then
  If Trim(NJ.Text) = "N" Or Trim(NJ.Text) = "S" Then
  Else
    NJ.Text = "N"
  End If
  If LK_CODTRA = 2403 And LK_EMP = "CAM" And LK_FLAG_SOS <> "A" Then
    i_numser.Text = "9"
  Else
    i_numser.Text = ""
  End If
End If

End Sub

Private Sub salir_Click()
Unload FORMGEN
LK_CODTRA = 0
End Sub

Private Sub t_direc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  fracli.Visible = False
  i_nomCLI.Caption = Trim(t_nombre.Text) & " - " & Trim(t_doc.Text) & " - " & Trim(t_direc.Text)
  NUMERO = FORMGEN.i_codcli.WhatsThisHelpID
  avanza_campo
End If
If KeyAscii = 27 Then
  i_codcli.Text = 0
  fracli.Visible = False
  i_codcli.SetFocus
End If

End Sub

Private Sub t_doc_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then Azul t_direc, t_direc
If KeyAscii = 27 Then
  i_codcli.Text = 0
  fracli.Visible = False
  i_codcli.SetFocus
End If

End Sub

Private Sub t_nombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul t_doc, t_doc
If KeyAscii = 27 Then
  i_codcli.Text = 0
  fracli.Visible = False
  i_codcli.SetFocus
End If
End Sub

Private Sub textolote_Change()
If gridlt.COL = 3 Then
gridlt.Text = Format(textolote.Text, "0.0000")
Else
gridlt.Text = Trim(textolote.Text)
End If
End Sub

Private Sub textolote_GotFocus()
temporal = gridlt.TextMatrix(gridlt.Row, gridlt.COL)
End Sub

Private Sub textolote_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
  textolote.Text = temporal
  textolote.Visible = False
  gridlt.SetFocus
  Exit Sub
End If
'If gridlt.COL = 1 Then Consistencias gridlt, textolote, KeyAscii
'If gridlt.COL = 4 Then Consistencias gridlt, textolote, KeyAscii
'If gridlt.COL = 5 Or gridlt.COL = 6 Then Consistencias gridlt, textolote, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If gridlt.COL = 1 Or gridlt.COL = 4 Then
  If Val(textolote.Text) > 99 Then
     Azul textolote, textolote
     Exit Sub
  End If
End If
If gridlt.Row >= gridlt.Rows - 1 Then
Else
  gridlt.Row = gridlt.Row + 1
End If
gridlt.SetFocus
textolote.Visible = False

fin:


End Sub

Private Sub TEXTOVAR_CANJE_Change()
grid_canje.Text = textovar_canje.Text
End Sub

Private Sub textovar_canje_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   textovar_canje.Text = textovar_bak
End If


If (KeyCode > 36 And KeyCode < 41) Or KeyCode = 13 Then
Else
   Exit Sub
End If


   If KeyCode = 38 Then
      If grid_canje.Row = 2 Then Exit Sub
   End If
   
   If KeyCode = 37 Then
      If grid_canje.COL = 0 Then Exit Sub
   End If
   If KeyCode = 40 Then
      If grid_canje.Row = grid_canje.Rows - 1 Then Exit Sub
   End If
   
If KeyCode = 13 Then
   If grid_canje.COL = 1 Or grid_canje.COL = 2 Then
   Else
   Exit Sub
   End If
End If

SALTO:

grid_canje.TextMatrix(grid_canje.Row, grid_canje.COL) = textovar_canje.Text
If LK_CODTRA <> 2401 Then
textovar_canje.Visible = False
If grid_canje.Visible Then grid_canje.SetFocus
End If


If KeyCode = 38 Then
   grid_canje.Row = grid_canje.Row - 1
ElseIf KeyCode = 40 Then
   grid_canje.Row = grid_canje.Row + 1
End If


If KeyCode = 37 Then
   grid_canje.COL = grid_canje.COL - 1
ElseIf KeyCode = 39 Or KeyCode = 13 Then
      If grid_canje.COL = 7 Then grid_canje.COL = 2
      grid_canje.COL = grid_canje.COL + 1
End If



End Sub

Private Sub TEXTOVAR_CANJE_KeyPress(KeyAscii As Integer)

'If Len(Trim(letche.ToolTipText)) <> 2 Then
'   MsgBox "Selecccionar Tipo Doc."
'   KeyAscii = 0
'   Exit Sub
'End If


If KeyAscii <> 13 Then Exit Sub


'If Len(Trim(letche.ToolTipText)) <> 2 Then
'   MsgBox "Selecccionar Tipo Doc."
'   textovar_canje.text = ""
'   Exit Sub
'End If
'If grid_canje.Col = 3 Then CALCULA_TOTALES_CANJE

'If grid_canje.Col = 3 Then
'   grid_canje.Col = 4
'   Exit Sub
'End If
'If grid_canje.Col = 4 Then
'   If IsDate(textovar_canje.text) = False Then
'      MsgBox "Fecha Incorrecta..."
'       Exit Sub
'   End If
'   grid_canje.Col = 5
'   Exit Sub
'End If
'If grid_canje.Col = 5 Then
'   grid_canje.Col = 6
'   Exit Sub
'End If
'If grid_canje.Col = 6 Then
'   grid_canje.Col = 7
'   Exit Sub
'End If


'If grid_canje.Col = 7 Then
'If grid_canje.Row = grid_canje.Rows - 1 And LK_CODTRA <> 2770 Then
'      grid_canje.Rows = grid_canje.Rows + 1
'      grid_canje.Row = grid_canje.Rows - 1
'      grid_canje.RowHeight(grid_canje.Row) = 350
'      grid_canje.Col = 3
'Else
'    grid_canje.Col = 3
'    grid_canje.Row = grid_canje.Row + 1
'End If
'End If

End Sub
Private Sub textovar_canje_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
Dim WS_TIPO As Integer
Dim WS_CALIDAD As Integer
Dim ws_codcia As String

If grid_canje.TextMatrix(grid_canje.Row, 3) <> "" Then
   grid_canje.TextMatrix(grid_canje.Row, 2) = letche.ToolTipText
End If
If PUB_TIPMOV = 10 Then
   grid_canje.TextMatrix(grid_canje.Row, 4) = i_fecha_vcto.Text
Else
'If grid_canje.Col = 3 Then grid_canje.TextMatrix(grid_canje.Row, 4) = LK_FECHA_DIA
End If


If KeyCode > 35 And KeyCode < 41 Then
   Exit Sub
End If

End Sub

Private Sub TEXTOVAR_CANJE_LostFocus()
CALCULA_TOTALES_CANJE
'i_tipdoc.Visible = False
'TEXTOVAR_CANJE.Visible = False

End Sub

Private Sub textovar_Change()
Dim ws_val_pre As Currency
Dim WW_PARCIAL As Currency
Dim WW_DESC1 As Currency
Dim WW_DESC2 As Currency
Dim pos1 As Integer
If (LK_EMP = "HER" Or LK_EMP = "3AA" Or LK_EMP = "PIU") And PUB_TIPMOV = 10 And grid_fac.COL = 10 Then
    pos1 = InStr(1, TEXTOVAR.Text, "+", 1)
    pos1 = pos1 - 1
    If pos1 < 0 Then
       pos1 = 5
       WW_DESC1 = Val(Mid(TEXTOVAR.Text, 1, pos1))
       WW_DESC2 = 0
    Else
       WW_DESC1 = Val(Mid(TEXTOVAR.Text, 1, pos1))
       pos1 = pos1 + 2
       WW_DESC2 = Val(Mid(TEXTOVAR.Text, pos1, 5))
    End If
End If



grid_fac.Text = TEXTOVAR.Text
If (LK_EMP = "HER" Or LK_EMP = "3AA" Or LK_EMP = "PIU") And PUB_TIPMOV = 10 Then
   If grid_fac.COL = 10 Then
     If loc_acceso_des <> "A" Then
         TEXTOVAR.Text = ""
         GoTo fin
     End If
      WW_PARCIAL = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 35)) * (100 - WW_DESC1) / 100)
      grid_fac.TextMatrix(grid_fac.Row, 6) = redondea(WW_PARCIAL * (100 - WW_DESC2) / 100)
      grid_fac.TextMatrix(grid_fac.Row, 7) = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 6)) * Val(grid_fac.TextMatrix(grid_fac.Row, 4)))
      If Val(grid_fac.TextMatrix(grid_fac.Row, 16)) = 0 Then GoTo fin
      SQ_OPER = 1
      PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 16))
      pu_codcia = LK_CODCIA
      LEER_ARM_LLAVE
      If arm_llave.EOF Then
         MsgBox "Digitar el Producto Nuevamente.", 48, Pub_Titulo
         Exit Sub
      Else
       ws_val_pre = Val(grid_fac.TextMatrix(grid_fac.Row, 6))
       If (arm_llave!arm_cospro * grid_fac.TextMatrix(grid_fac.Row, 14)) < ws_val_pre Then
       Else
         MsgBox "El Precio de Venta es Menor que el Costo. " & Chr(13) & "Costo Promedio es : " & Format((arm_llave!arm_cospro * grid_fac.TextMatrix(grid_fac.Row, 14)), "0.00"), 48, Pub_Titulo
         TEXTOVAR.Text = ""
       End If
      End If
      
   ElseIf grid_fac.COL = 6 Then
      ' quite precio digitado no altera el porcentaje de descto.. Alan
      'If Val(grid_fac.TextMatrix(grid_fac.Row, 35)) <> 0 Then grid_fac.TextMatrix(grid_fac.Row, 10) = redondea((Val(grid_fac.TextMatrix(grid_fac.Row, 35)) - Val(grid_fac.TextMatrix(grid_fac.Row, 6))) * 100 / Val(Val(grid_fac.TextMatrix(grid_fac.Row, 35))))
      grid_fac.TextMatrix(grid_fac.Row, 7) = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 6)) * Val(grid_fac.TextMatrix(grid_fac.Row, 4)))
   End If
End If
fin:
If grid_fac.COL = 7 And Trim(TEXTOVAR.Text) <> "" Then
   grid_fac.TextMatrix(grid_fac.Row, 25) = "1"
End If
If grid_fac.COL = 4 Or grid_fac.COL = 1 Then
   BORRA_ITEM_LOTE Val(grid_fac.TextMatrix(grid_fac.Row, 16)), grid_fac.Row
End If

If grid_fac.COL = 6 And Trim(TEXTOVAR.Text) <> "" Then grid_fac.TextMatrix(grid_fac.Row, 25) = ""

End Sub


Private Sub TEXTOVAR_GotFocus()
If LK_CODTRA = 2401 Then
 TEXTOVAR.Locked = False
 If grid_fac.COL <> 10 Then Exit Sub
 If Trim(SUT_LLAVE!SUT_DESCTO) = "F" Then
   TEXTOVAR.Locked = True
   Exit Sub
 End If
End If

End Sub

Private Sub textovar_KeyDown(KeyCode As Integer, Shift As Integer)
Dim WS_NRO_ITEMS As Integer
Dim tf
Dim WSPOR As Currency
Dim fil As Integer
Dim celda_anterior
flag_salto = 1

If KeyCode = 13 Then
  If grid_fac.COL = 4 And Trim(TEXTOVAR.Text) = "" Then Exit Sub
End If

If grid_fac.COL = 1 Then
  If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
    If TEXTOVAR.Text = "" Then
    Else
     GoTo BUSCANDO
    End If
  Else
   If IsNumeric(TEXTOVAR.Text) = True Or TEXTOVAR.Text = "" Then
   Else
     GoTo BUSCANDO
  End If
End If
End If

If KeyCode = 27 Then
   TEXTOVAR.Text = textovar_bak
End If


If (KeyCode > 36 And KeyCode < 41 And TEXTOVAR.Text = "") Or KeyCode = 13 Then
Else
   Exit Sub
End If


   If KeyCode = 38 Then
      If grid_fac.Row = 2 Then Exit Sub
   End If
   
   If KeyCode = 37 Then
      If grid_fac.COL = 0 Then Exit Sub
   End If
   If KeyCode = 40 Then
      If grid_fac.Row = grid_fac.Rows - 1 Then Exit Sub
   End If
   
If KeyCode = 13 Then
   If grid_fac.COL = 1 Or grid_fac.COL = 2 Or grid_fac.COL = 3 Or grid_fac.COL = 4 Or grid_fac.COL = 5 Or grid_fac.COL = 6 Or grid_fac.COL = 7 Or grid_fac.COL = 8 Or grid_fac.COL = 10 Then
   Else
   Exit Sub
   End If
End If

If Nulo_Valors(SUT_LLAVE!SUT_UNIDADES) <> "A" Then
If grid_fac.COL = 4 And Not arm_llave.EOF And pub_signo_arm = -1 Then
   If Val(TEXTOVAR.Text) > Val(grid_fac.TextMatrix(grid_fac.Row, 13)) Then
      MsgBox "Ojo Stock Negativo..." & "Existencia : " & Val(grid_fac.TextMatrix(grid_fac.Row, 13))
      If LK_USU_STOCK <> "A" Then
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         grid_fac.TextMatrix(grid_fac.Row, 13) = ""
         grid_fac.TextMatrix(grid_fac.Row, 16) = ""
         Exit Sub
      End If
   End If
End If
End If


If PUB_TIPMOV = 10 Then
    If Trim(i_fbg.Text) = "F" Then
      If Nulo_Valor0(par_llave!par_fac_lines) < grid_fac.Row - 1 Then
        If LK_EMP = "HER" Then
          MsgBox "Tope de lineas alcanzado..."
          grid_fac.TextMatrix(grid_fac.Row, 38) = ""
          grid_fac.TextMatrix(grid_fac.Row, 4) = ""
          grid_fac.TextMatrix(grid_fac.Row, 5) = ""
          grid_fac.TextMatrix(grid_fac.Row, 6) = ""
          grid_fac.TextMatrix(grid_fac.Row, 7) = ""
          grid_fac.TextMatrix(grid_fac.Row, 13) = ""
          grid_fac.TextMatrix(grid_fac.Row, 16) = ""
          Exit Sub
        End If
      End If
    ElseIf Trim(i_fbg.Text) = "B" Then
        If PUB_CODVEN = 10 Then
          WS_NRO_ITEMS = 13 ' TEMPORAL
        Else
          WS_NRO_ITEMS = Nulo_Valor0(par_llave!par_BOL_lines)
        End If
        
      If WS_NRO_ITEMS < grid_fac.Row - 1 Then
        If LK_EMP = "HER" Then
          MsgBox "Tope de lineas alcanzado..."
          grid_fac.TextMatrix(grid_fac.Row, 4) = ""
          grid_fac.TextMatrix(grid_fac.Row, 5) = ""
          grid_fac.TextMatrix(grid_fac.Row, 6) = ""
          grid_fac.TextMatrix(grid_fac.Row, 7) = ""
          grid_fac.TextMatrix(grid_fac.Row, 13) = ""
          grid_fac.TextMatrix(grid_fac.Row, 16) = ""
          Exit Sub
        End If
      End If
    ElseIf Trim(i_fbg.Text) = "G" Then
      If Nulo_Valor0(par_llave!par_GUIA_lines) < grid_fac.Row - 1 Then
        If LK_EMP = "HER" Then
          MsgBox "Tope de lineas alcanzado..."
          grid_fac.TextMatrix(grid_fac.Row, 4) = ""
          grid_fac.TextMatrix(grid_fac.Row, 5) = ""
          grid_fac.TextMatrix(grid_fac.Row, 6) = ""
          grid_fac.TextMatrix(grid_fac.Row, 7) = ""
          grid_fac.TextMatrix(grid_fac.Row, 13) = ""
          grid_fac.TextMatrix(grid_fac.Row, 16) = ""
          Exit Sub
        End If
      End If
    End If
End If

If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
      If Nulo_Valor0(par_llave!par_NOTAS_lines) < grid_fac.Row - 1 Then
         MsgBox "Tope de lineas alcanzado..."
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         grid_fac.TextMatrix(grid_fac.Row, 13) = ""
         grid_fac.TextMatrix(grid_fac.Row, 16) = ""
         Exit Sub
      End If
End If


If grid_fac.COL = 1 Then
   If TEXTOVAR.Text = "" And KeyCode <> 13 Then
         grid_fac.TextMatrix(grid_fac.Row, 1) = ""
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         grid_fac.TextMatrix(grid_fac.Row, 13) = ""
         grid_fac.TextMatrix(grid_fac.Row, 16) = ""
         label_nomart.Caption = ""
         calcula_totales
      GoTo SALTO
   End If
Else
   GoTo SALTO
End If
    SQ_OPER = 1
    On Error GoTo OJO
    PUB_KEY = Val(FORMGEN.TEXTOVAR.Text)
    On Error GoTo 0
    If PUB_KEY = 0 Then GoTo fin
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       grid_fac.TextMatrix(grid_fac.Row, 0) = ""
       label_nomart.Caption = ""
       GoTo fin
    Else
       If PUB_TIPMOV = 20 Then
          If SUT_LLAVE!sut_secuencia = 1 Then
          '   If art_LLAVE!art_flag_stock <> "M" Then
          '      MsgBox "No es Mercaderias ... ", 48, Pub_Titulo
          '      textovar.Text = ""
          '      GoTo fin
          '   End If
          'ElseIf art_LLAVE!art_flag_stock <> "A" Then
          '      MsgBox "No es Activo Fijo ... ", 48, Pub_Titulo
          '      textovar.Text = ""
          '      GoTo fin
          End If
       End If
       If PUB_TIPMOV = 10 Then
          If SUT_LLAVE!sut_secuencia <> 2 Then
             If art_LLAVE!art_flag_stock <> "M" And art_LLAVE!art_flag_stock <> "S" Then
                'MsgBox "No es Mercaderias ... ", 48, Pub_Titulo
                'textovar.Text = ""
                'Exit Sub
             End If
          End If
          If Val(Nulo_Valor0(art_LLAVE!ART_CALIDAD)) <> 1 Then
            MsgBox "No procede para la Venta... Articulo defectuoso", 48, Pub_Titulo
            Azul TEXTOVAR, TEXTOVAR
            Exit Sub
         End If
          
        End If
       grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE(2)
    End If
    grid_fac.TextMatrix(grid_fac.Row, 16) = art_LLAVE(0)
    
    SQ_OPER = 1
    PUB_CODART = Val(FORMGEN.TEXTOVAR.Text)
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    'modi
    If i_fbg.Visible And LK_ACTIVA = "A" Then
       If PUB_TIPMOV <> 0 And (i_fbg.Text = "F" Or i_fbg.Text = "B") Then
          grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_saldo_s
       Else
          grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_Saldo_n
       End If
    Else
      grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_STOCK
    End If
    If NJ.Visible And LK_ACTIVA = "A" Then
     If NJ.Text = "S" Then
       grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_saldo_s
     Else
       grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_Saldo_n
     End If
    End If
     
If Nulo_Valors(SUT_LLAVE!SUT_UNIDADES) = "A" And pub_signo_arm = -1 Then
  ' modi
   If Val(grid_fac.TextMatrix(grid_fac.Row, 13)) <= 0 Then
      MsgBox "Ojo Stock no disponible " & "Existencia : " & Val(grid_fac.TextMatrix(grid_fac.Row, 13))
      If LK_USU_STOCK <> "A" Then
         grid_fac.TextMatrix(grid_fac.Row, 1) = ""
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         grid_fac.TextMatrix(grid_fac.Row, 13) = ""
         grid_fac.TextMatrix(grid_fac.Row, 16) = ""
         TEXTOVAR.Text = ""
         Exit Sub
      End If
   End If
End If

    
    
    
grid_fac.TextMatrix(grid_fac.Row, 16) = art_LLAVE!ART_KEY
grid_fac.TextMatrix(grid_fac.Row, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
grid_fac.TextMatrix(grid_fac.Row, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
grid_fac.TextMatrix(grid_fac.Row, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
If LK_CODTRA = 2401 Then VER_LINEAS

If LK_EMP = "3AA" Then
  If VERIFICA_REPET = 1 Then
    grid_fac.TextMatrix(grid_fac.Row, 1) = ""
    grid_fac.TextMatrix(grid_fac.Row, 4) = ""
    grid_fac.TextMatrix(grid_fac.Row, 5) = ""
    grid_fac.TextMatrix(grid_fac.Row, 6) = ""
    grid_fac.TextMatrix(grid_fac.Row, 7) = ""
    grid_fac.TextMatrix(grid_fac.Row, 13) = ""
    grid_fac.TextMatrix(grid_fac.Row, 16) = ""
    TEXTOVAR.Text = ""
    Exit Sub
  End If
Else
Call VERIFICA_REPET
End If
grid_fac.TextMatrix(grid_fac.Row, 11) = arm_llave!arm_cospro
'grid_fac.TextMatrix(grid_fac.Row, 14) = 1

If Val(grid_fac.TextMatrix(grid_fac.Row, 12)) = 0 Then grid_fac.TextMatrix(grid_fac.Row, 12) = pub_signo_arm


    
    
grid_fac.CellForeColor = vbBlack
If Nulo_Valors(SUT_LLAVE!SUT_PRECIO) = "1" Then
   grid_fac.TextMatrix(grid_fac.Row, 6) = arm_llave!arm_cospro
End If
 

verifica_unidades

grid_fac.TextMatrix(grid_fac.Row, 11) = arm_llave!arm_cospro


SALTO:

grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL) = TEXTOVAR.Text

If grid_fac.COL <> 1 Then
If grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL) <> textovar_bak Then
   calcula_totales
End If
End If
PEPE:
TEXTOVAR.Visible = False
grid_fac.SetFocus



If KeyCode = 38 Then
   grid_fac.Row = grid_fac.Row - 1
ElseIf KeyCode = 40 Then
   grid_fac.Row = grid_fac.Row + 1
End If

If KeyCode = 37 Then
   grid_fac.COL = tab_izquierda(grid_fac.COL)
ElseIf KeyCode = 39 Or KeyCode = 13 Then
    If tab_derecha(grid_fac.COL) = 1 And grid_fac.TextMatrix(grid_fac.Row, 16) <> "" Then
      flag_salto = 1
      If grid_fac.Row = grid_fac.Rows - 1 Then
         grid_fac.Rows = grid_fac.Rows + 1
      End If
      grid_fac.Row = grid_fac.Row + 1
      grid_fac.RowHeight(grid_fac.Row) = 315
      flag_salto = 0
      grid_fac.COL = 1
    Else
      grid_fac.COL = tab_derecha(grid_fac.COL)
    End If
    
End If

GoTo fin

BUSCANDO:
If grid_fac.COL <> 1 Then Exit Sub

Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not LV_ART.Visible Then
 Exit Sub
End If

If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And TEXTOVAR.Text = "" Then
  loc_key = 1
  Set LV_ART.SelectedItem = LV_ART.ListItems(loc_key)
'  LV_CLI.Visible = False
  LV_ART.ListItems.Item(loc_key).Selected = True
  LV_ART.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_ART.ListItems.count Then loc_key = LV_ART.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_ART.ListItems.count Then loc_key = LV_ART.ListItems.count
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
  LV_ART.ListItems.Item(loc_key).Selected = True
  If LK_CODTRA <> 2401 Then GoTo SALTA2401
  If fraprecios.Visible = False Then
   fraprecios.Top = LV_ART.Top + 3230
   fraprecios.Visible = True
  End If
  lbltitulostock.Caption = Left(lk_descrip_cias, 10)
  lbltitulostock2.Caption = Right(lk_descrip_cias, 10)
  lpstock2.Caption = Int(LV_ART.ListItems.Item(loc_key).SubItems(12))
  lpstock2.Caption = Format(lpstock2.Caption, "##,##0.00")
  lpstocku2.Caption = Format(Val((LV_ART.ListItems.Item(loc_key).SubItems(13))) - (Val(lpstock2.Caption) * Val(LV_ART.ListItems.Item(loc_key).SubItems(9))), "##,##0.000")
  
  lpstock3.Caption = Int(LV_ART.ListItems.Item(loc_key).SubItems(14))
  lpstock3.Caption = Format(lpstock3.Caption, "##,##0.00")
  lpstocku3.Caption = Format(Val((LV_ART.ListItems.Item(loc_key).SubItems(15))) - (Val(lpstock3.Caption) * Val(LV_ART.ListItems.Item(loc_key).SubItems(9))), "##,##0.000")

  lpnombre.Caption = Trim(LV_ART.ListItems.Item(loc_key))
  lpunidad.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(2))
  lpstock.Caption = Int(LV_ART.ListItems.Item(loc_key).SubItems(3))
  lpstocku.Caption = Format(Val((LV_ART.ListItems.Item(loc_key).SubItems(10))) - (Val(lpstock.Caption) * Val(LV_ART.ListItems.Item(loc_key).SubItems(9))), "##,##0.000")
  lpstock.Caption = Format(lpstock.Caption, "##,##0.00")
  lpp1.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(4))
  lpp2.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(5))
  lpp3.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(6))
  lpp4.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(7))
  lpp5.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(8))
  lpp6.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(11))
  lbllote.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(17))
  If Val(LV_ART.ListItems.Item(loc_key).SubItems(18)) = 1 Then
    fraprecios.BackColor = QBColor(3)
  ElseIf Val(LV_ART.ListItems.Item(loc_key).SubItems(16)) = 1 Then
    fraprecios.BackColor = QBColor(6)
  Else
    fraprecios.BackColor = QBColor(8)
  End If
  
SALTA2401:
  LV_ART.ListItems.Item(loc_key).EnsureVisible
  TEXTOVAR.Text = Trim(LV_ART.ListItems.Item(loc_key).Text) & " "
  DoEvents
  
  GoTo XXX
  
  grid_unid.Visible = True
  SQ_OPER = 2
  pu_codcia = LK_CODCIA
  PUB_CODART = Trim(LV_ART.ListItems.Item(loc_key).SubItems(1))
  LEER_PRE_LLAVE
  
  grid_unid.Height = 1200
  LIMPIA_PREC
  fil = 1
Do Until pre_mayor.EOF
   fil = fil + 1
   grid_unid.Rows = fil + 1
   grid_unid.Row = fil
   grid_unid.RowHeight(fil) = 285
   grid_unid.TextMatrix(fil, 0) = Trim(pre_mayor!pre_UNIDAD)
   grid_unid.TextMatrix(fil, 1) = pre_mayor!PRE_EQUIV
   grid_unid.TextMatrix(fil, 2) = ""
   grid_unid.TextMatrix(fil, 3) = Format((Nulo_Valor0(pre_mayor!PRE_COSTO) * ((100 + LK_IGV) / 100)), "0.00")
   grid_unid.COL = 4
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fil, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE1) * 100) / Val(grid_unid.TextMatrix(fil, 3)) - 100
   grid_unid.TextMatrix(fil, 4) = Format(WSPOR, "0.0")
   grid_unid.TextMatrix(fil, 5) = Nulo_Valor0(pre_mayor!PRE_PRE1)
   grid_unid.COL = 6
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fil, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE2) * 100) / Val(grid_unid.TextMatrix(fil, 3)) - 100
   grid_unid.TextMatrix(fil, 6) = Format(WSPOR, "0.0")
   grid_unid.TextMatrix(fil, 7) = Nulo_Valor0(pre_mayor!PRE_PRE2)
   grid_unid.COL = 8
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fil, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE3) * 100) / Val(grid_unid.TextMatrix(fil, 3)) - 100
   grid_unid.TextMatrix(fil, 8) = Format(WSPOR, "0.0")
   grid_unid.TextMatrix(fil, 9) = Nulo_Valor0(pre_mayor!PRE_PRE3)
   grid_unid.COL = 10
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fil, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE4) * 100) / Val(grid_unid.TextMatrix(fil, 3)) - 100
   grid_unid.TextMatrix(fil, 10) = Format(WSPOR, "0.0")
   grid_unid.TextMatrix(fil, 11) = Nulo_Valor0(pre_mayor!PRE_PRE4)
   grid_unid.COL = 12
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fil, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE5) * 100) / Val(grid_unid.TextMatrix(fil, 3)) - 100
   grid_unid.TextMatrix(fil, 12) = Format(WSPOR, "0.0")
   grid_unid.TextMatrix(fil, 13) = Nulo_Valor0(pre_mayor!PRE_PRE5)
   grid_unid.TextMatrix(fil, 14) = Nulo_Valor0(pre_mayor!PRE_FLAG_UNIDAD)
   grid_unid.TextMatrix(fil, 15) = Nulo_Valor0(pre_mayor!pre_secuencia)
'   If Nulo_Valor0(pre_mayor!PRE_FLAG_UNIDAD) = "A" Then
'     LBLUNIDAD.Caption = Trim(pre_mayor!pre_unidad)
'   End If
   pre_mayor.MoveNext
Loop

grid_unid.Row = 1
grid_unid.COL = 0

XXX:
  TEXTOVAR.SelStart = Len(TEXTOVAR.Text)
  DoEvents
fin:
OJO:
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I, car
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
car = Chr(KeyAscii)
car = UCase(car)
KeyAscii = Asc(car)
If KeyAscii = 27 Then
  TEXTOVAR.Text = ""
  LV_ART.Visible = False
  fraprecios.Visible = False
  Exit Sub
End If
If KeyAscii <> 13 Then pub_ojo = "A"


'label_nomart.Caption = ""
'label_precio.Caption = ""
If grid_fac.COL = 3 Then
 Consistencias grid_fac, TEXTOVAR, KeyAscii, 1, 3
 If Len(TEXTOVAR.Text) > 13 Then
    KeyAscii = 0
    Exit Sub
 End If
End If
If grid_fac.COL = 2 Then
 If Len(TEXTOVAR.Text) > 13 Then
    KeyAscii = 0
    Exit Sub
 End If
 Consistencias grid_fac, TEXTOVAR, KeyAscii, 2, 2
End If
If grid_fac.COL = 4 Then
 If Len(TEXTOVAR.Text) > 13 Then
    KeyAscii = 0
    Exit Sub
 End If
  Consistencias grid_fac, TEXTOVAR, KeyAscii, 2, 4
End If
 
'Else
'textovar.MaxLength = 0
'End If
If grid_fac.COL = 5 Then Consistencias grid_fac, TEXTOVAR, KeyAscii, 2, 5
If grid_fac.COL = 6 Then Consistencias grid_fac, TEXTOVAR, KeyAscii, 2, 6
If grid_fac.COL = 7 Then Consistencias grid_fac, TEXTOVAR, KeyAscii, 2, 6
If grid_fac.COL = 8 Then Consistencias grid_fac, TEXTOVAR, KeyAscii, 2, 8


If grid_fac.COL = 1 And KeyAscii <> 13 Then
   label_nomart.Caption = ""
   label_precio.Caption = ""
   grid_fac.TextMatrix(grid_fac.Row, 0) = ""
   grid_fac.TextMatrix(grid_fac.Row, 4) = ""
   grid_fac.TextMatrix(grid_fac.Row, 5) = ""
   grid_fac.TextMatrix(grid_fac.Row, 6) = ""
   grid_fac.TextMatrix(grid_fac.Row, 7) = ""
   grid_fac.TextMatrix(grid_fac.Row, 16) = ""
   calcula_totales
End If

If KeyAscii <> 13 Then
   TEXTOVAR.ForeColor = vbRed
End If

If KeyAscii <> 13 Then
   GoTo fin
End If

flag_salto = 0
If grid_fac.COL <> 1 Then Exit Sub
'solo para codart....

If grid_fac.TextMatrix(grid_fac.Row, 1) <> TEXTOVAR.Text Then Exit Sub
If Trim(TEXTOVAR.Text) = "" Then Exit Sub
 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
     SQ_OPER = 3
     pu_alterno = TEXTOVAR.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul TEXTOVAR, TEXTOVAR
       Exit Sub
     End If
     PUB_KEY = art_llave_alt!ART_KEY
     PUB_CODART = art_llave_alt!ART_KEY
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     GoTo pasa
  
 Else
 On Error GoTo OJO
 PUB_KEY = Val(FORMGEN.TEXTOVAR.Text)
 On Error GoTo 0
 If Len(TEXTOVAR.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(TEXTOVAR.Text) = False Then
   PUB_KEY = 0
 End If
End If
If IsNumeric(FORMGEN.TEXTOVAR) Then
''If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = FORMGEN.TEXTOVAR.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul TEXTOVAR, TEXTOVAR
       GoTo fin
    Else
       If PUB_TIPMOV = 20 Then
          If SUT_LLAVE!sut_secuencia = 1 Then
             If art_LLAVE!art_flag_stock <> "M" Then
                'MsgBox "No es Mercaderias ... ", 48, Pub_Titulo
                'textovar.Text = ""
                'GoTo fin
             End If
          End If
       End If
       If PUB_TIPMOV = 10 Then
          If SUT_LLAVE!sut_secuencia <> 2 Then
             If art_LLAVE!art_flag_stock <> "M" Then
                'MsgBox "No es Mercaderias ... ", 48, Pub_Titulo
                'textovar.Text = ""
                'Exit Sub
             End If
          
          End If
        End If
    End If
    If PUB_TIPMOV = 75 Then
         If Val(Nulo_Valor0(art_LLAVE!ART_CODART2)) = 0 Then
            MsgBox "Articulo no esta relacionado..."
            Azul TEXTOVAR, TEXTOVAR
            Exit Sub
         End If
    End If
Else
  If VAR_ACTIVAR <> 99 And TEXTOVAR <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
     SQ_OPER = 3
     pu_alterno = TEXTOVAR.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul TEXTOVAR, TEXTOVAR
       Exit Sub
     Else
       If PUB_TIPMOV = 20 Then
          If SUT_LLAVE!sut_secuencia = 1 Then
             If art_llave_alt!art_flag_stock <> "M" Then
                'MsgBox "No es Mercaderias ... ", 48, Pub_Titulo
                'textovar.Text = ""
                'Exit Sub
             End If
       '   ElseIf art_llave_alt!art_flag_stock <> "A" Then
       '         MsgBox "No es Activo Fijo ... ", 48, Pub_Titulo
       '         textovar.Text = ""
       '         Exit Sub
        End If
        If PUB_TIPMOV = 10 Then
          If SUT_LLAVE!sut_secuencia <> 2 Then
             If art_llave_alt!art_flag_stock <> "M" Then
            '    MsgBox "No es Mercaderias ... ", 48, Pub_Titulo
            '    textovar.Text = ""
            '    Exit Sub
             End If
          
          End If
        End If
        
     End If
     End If
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
      PUB_KEY = art_llave_alt!ART_KEY
      LEER_ART_LLAVE
      If PUB_TIPMOV = 75 Then
         If Val(Nulo_Valor0(art_LLAVE!ART_CODART2)) = 0 Then
            MsgBox "Articulo no esta relacionado..."
            Azul TEXTOVAR, TEXTOVAR
            Exit Sub
         End If
      End If
  Else
    If loc_key > LV_ART.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    VALOR = UCase(LV_ART.ListItems.Item(loc_key).Text)
    If Trim(UCase(TEXTOVAR.Text)) = Left(VALOR, Len(Trim(TEXTOVAR.Text))) And Len(Trim(TEXTOVAR.Text)) <> 0 Then
      If VAR_ACTIVAR <> 99 Then
      TEXTOVAR.Text = Trim(LV_ART.ListItems.Item(loc_key).SubItems(1))
      Else
      TEXTOVAR.Text = Trim(LV_ART.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(LV_ART.ListItems.Item(loc_key).SubItems(2))
      Else
       PUB_KEY = Val(FORMGEN.TEXTOVAR.Text)
      End If
      LEER_ART_LLAVE
      If PUB_TIPMOV = 20 Then
         If SUT_LLAVE!sut_secuencia = 1 Then
             If art_LLAVE!art_flag_stock <> "M" Then
               ' MsgBox "No es Mercaderias ... ", 48, Pub_Titulo
               ' textovar.Text = ""
               ' Exit Sub
             End If
          End If
          'If art_LLAVE!art_flag_stock <> "A" Then
          '      MsgBox "No es Activo Fijo ... ", 48, Pub_Titulo
          '      TEXTOVAR.Text = ""
          '      Exit Sub
          'End If
       End If
      
      VAR_ACTIVAR = 0
    Else
      Exit Sub
    End If
  End If
End If
    SQ_OPER = 1
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       If VAR_ACTIVAR <> 99 Then
          PUB_CODART = PUB_KEY
       Else
          PUB_CODART = Val(LV_ART.ListItems.Item(loc_key).SubItems(2))
       End If
    Else
       PUB_CODART = Val(FORMGEN.TEXTOVAR.Text)
    End If
pasa:
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    'modi
    If i_fbg.Visible And LK_ACTIVA = "A" Then
       If PUB_TIPMOV <> 0 And (i_fbg.Text = "F" Or i_fbg.Text = "B") Then
          grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_saldo_s
       Else
          grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_Saldo_n
       End If
    Else
      grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_STOCK
    End If
    If art_LLAVE!art_situacion <> 0 Then
       MsgBox "Producto Esta Desactivado!!!! ", 48, Pub_Titulo
       MsgBox "Revisar Producto - Esta Desactivado!!!! ", 48, Pub_Titulo
    End If
    If PUB_TIPMOV = 10 Then
        If SUT_LLAVE!sut_secuencia <> 2 Then
           If art_LLAVE!art_flag_stock = "M" Or art_LLAVE!art_flag_stock = "P" Then
           Else
              'MsgBox "No es Mercaderias ... ", 48, Pub_Titulo
              'textovar.Text = ""
              'Exit Sub
           End If
        End If
        
        If NJ.Visible And LK_ACTIVA = "A" Then
          If NJ.Text = "S" Then
            grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_saldo_s
          Else
            grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_Saldo_n
          End If
        End If
        If Val(Nulo_Valor0(art_LLAVE!ART_CALIDAD)) <> 1 Then
            MsgBox "No procede para la Venta... Articulo defectuoso", 48, Pub_Titulo
            Azul TEXTOVAR, TEXTOVAR
            Exit Sub
         End If
    End If
    If PUB_TIPMOV = 75 Then
         If Val(Nulo_Valor0(art_LLAVE!ART_CODART2)) = 0 Then
            MsgBox "Articulo no esta relacionado...", 48, Pub_Titulo
            Azul TEXTOVAR, TEXTOVAR
            Exit Sub
         End If
   End If
   If PUB_TIPMOV = 102 Then
     'If Nulo_Valors(art_LLAVE!art_flag_cambio) <> "A" Then
     '    MsgBox "Producto NO Procede para Cambio", 48, Pub_Titulo
     '    Azul TEXTOVAR, TEXTOVAR
     '    Exit Sub
     'End If
   End If
         
   If Val(grid_fac.TextMatrix(grid_fac.Row, 13)) <= 0 And (pub_signo_arm = -1 Or (LK_CODTRA = 2414 And grid_fac.Row = 2)) Then
      MsgBox "Ojo Stock no disponible " & "Existencia : " & grid_fac.TextMatrix(grid_fac.Row, 13)
      GoTo todo
      If LK_USU_STOCK <> "A" Then
         grid_fac.TextMatrix(grid_fac.Row, 3) = ""
         grid_fac.TextMatrix(grid_fac.Row, 4) = ""
         grid_fac.TextMatrix(grid_fac.Row, 5) = ""
         grid_fac.TextMatrix(grid_fac.Row, 6) = ""
         grid_fac.TextMatrix(grid_fac.Row, 7) = ""
         grid_fac.TextMatrix(grid_fac.Row, 13) = ""
         grid_fac.TextMatrix(grid_fac.Row, 16) = ""
         TEXTOVAR.Text = ""
         Exit Sub
      End If
todo:
  End If
    
    

verifica_unidades

If Nulo_Valors(SUT_LLAVE!SUT_PRECIO) = "1" Then
   grid_fac.TextMatrix(grid_fac.Row, 6) = Nulo_Valor0(arm_llave!arm_cospro)
End If
grid_fac.CellForeColor = vbBlack
grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE!art_nombre
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
   If VAR_ACTIVAR <> 99 Then
   Else
    grid_fac.TextMatrix(grid_fac.Row, 1) = LV_ART.ListItems.Item(loc_key)
   End If
Else
  grid_fac.TextMatrix(grid_fac.Row, 1) = art_LLAVE!ART_KEY
End If
grid_fac.TextMatrix(grid_fac.Row, 11) = Nulo_Valor0(arm_llave!arm_cospro)
'modi
If i_fbg.Visible And LK_ACTIVA = "A" Then
   If PUB_TIPMOV <> 0 And (i_fbg.Text = "F" Or i_fbg.Text = "B") Then
      grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_saldo_s
   Else
      grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_Saldo_n
   End If
Else
  grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_STOCK
End If
If NJ.Visible And LK_ACTIVA = "A" Then
     If NJ.Text = "S" Then
       grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_saldo_s
     Else
       grid_fac.TextMatrix(grid_fac.Row, 13) = arm_llave!ARM_Saldo_n
     End If
End If
grid_fac.TextMatrix(grid_fac.Row, 16) = art_LLAVE!ART_KEY
grid_fac.TextMatrix(grid_fac.Row, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
grid_fac.TextMatrix(grid_fac.Row, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
grid_fac.TextMatrix(grid_fac.Row, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
' arreglar Despues con un campo del arti
If art_LLAVE!ART_CALIDAD <> 1 And PUB_TIPMOV = 10 Then
        MsgBox "No procede para la Venta ..Articulo es defectuoso"
        Azul TEXTOVAR, TEXTOVAR
        Exit Sub
End If
If PUB_TIPMOV = 102 Then
  'If Nulo_Valors(art_LLAVE!art_flag_cambio) <> "A" Then
  '    MsgBox "Producto NO Procede para Cambio", 48, Pub_Titulo
  '    Azul TEXTOVAR, TEXTOVAR
  '    Exit Sub
  'End If
End If

          
If LK_CODTRA = 2401 Then VER_LINEAS
If LK_EMP = "3AA" Then
  If VERIFICA_REPET = 1 Then
    grid_fac.TextMatrix(grid_fac.Row, 0) = ""
    grid_fac.TextMatrix(grid_fac.Row, 1) = ""
    grid_fac.TextMatrix(grid_fac.Row, 4) = ""
    grid_fac.TextMatrix(grid_fac.Row, 5) = ""
    grid_fac.TextMatrix(grid_fac.Row, 6) = ""
    grid_fac.TextMatrix(grid_fac.Row, 7) = ""
    grid_fac.TextMatrix(grid_fac.Row, 13) = ""
    grid_fac.TextMatrix(grid_fac.Row, 16) = ""
    TEXTOVAR.Text = ""
    Azul TEXTOVAR, TEXTOVAR
    Exit Sub
  End If
Else
Call VERIFICA_REPET
End If

If Val(grid_fac.TextMatrix(grid_fac.Row, 12)) = 0 Then grid_fac.TextMatrix(grid_fac.Row, 12) = pub_signo_arm

grid_fac.COL = tab_derecha(1)
If Not TEXTOVAR.Visible Then grid_fac.SetFocus

LV_ART.Visible = False
fraprecios.Visible = False
VAR_ACTIVAR = 0
   
fin:
OJO:
End Sub
Private Sub textovar_KeyUp(KeyCode As Integer, Shift As Integer)
'' CHEQUEAR AQUI...
Dim aux_codcia As String * 2
Dim var
Dim WS_TIPO As Integer
Dim WS_CALIDAD As Integer
Dim ws_codcia As String
' CAMTEX
If grid_fac.COL = 4 And KeyCode = 116 And LK_EMP = "CAM" Then
  LOC_BRUTO = 0
  'Load Frmcalculo
  'Frmcalculo.Show 1
  TEXTOVAR.Text = LOC_BRUTO
  Exit Sub
End If
If grid_fac.COL = 1 And KeyCode = 46 Then
   label_nomart.Caption = ""
   grid_fac.TextMatrix(grid_fac.Row, 0) = ""
   grid_fac.TextMatrix(grid_fac.Row, 5) = ""
   grid_fac.TextMatrix(grid_fac.Row, 6) = ""
End If

If KeyCode > 35 And KeyCode < 41 Then
   Exit Sub
End If
If grid_fac.COL <> 1 Then Exit Sub
' puede estar pruebas
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 If Len(TEXTOVAR.Text) = 0 Or TEXTOVAR.Text = "" Then
   VAR_ACTIVAR = 0
   LV_ART.Visible = False
   fraprecios.Visible = False
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
 If Len(TEXTOVAR.Text) = 0 Or IsNumeric(TEXTOVAR.Text) = True Then
   LV_ART.Visible = False
   fraprecios.Visible = False
   Exit Sub
 End If
End If
If LV_ART.Visible = False And KeyCode <> 13 And Len(TEXTOVAR.Text) = 2 Then
'If LV_ART.Visible = False And IsNumeric(textovar.text) = False Then
    If TEXTOVAR.Text = "" Then Exit Sub
    var = Asc(TEXTOVAR.Text)
    var = var + 1
    If var = 33 Or var = 91 Or var = 58 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    WS_TIPO = Val(Nulo_Valor0(SUT_LLAVE!SUT_art_gru))
    WS_CALIDAD = Val(Nulo_Valor0(SUT_LLAVE!SUT_CALIDAD))
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       ws_codcia = "00"
    End If
    If LK_CODTRA = 2409 Or LK_CODTRA = 2401 Then
          ' tentativo cambio
      '//////////////////
       numarchi = 7
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
         If WS_TIPO = 0 And WS_CALIDAD = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK, PRE_EQUIV , PRE_UNIDAD, PRE_PRE1,PRE_PRE2,PRE_PRE3,PRE_PRE4,PRE_PRE5,PRE_PRE6 FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND  ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
         If WS_CALIDAD = 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, PRE_UNIDAD, PRE_PRE1,PRE_PRE2,PRE_PRE3,PRE_PRE4,PRE_PRE5,PRE_PRE6 FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
         If WS_CALIDAD <> 0 And WS_TIPO = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV,  PRE_UNIDAD,PRE_PRE1,PRE_PRE2,PRE_PRE3,PRE_PRE4,PRE_PRE5,PRE_PRE6 FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
         If WS_CALIDAD <> 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO , PRE_EQUIV, PRE_PRE1, PRE_UNIDAD,PRE_PRE2,PRE_PRE3,PRE_PRE4,PRE_PRE5,PRE_PRE6 FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & " AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
      Else
         If WS_TIPO = 0 And WS_CALIDAD = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV,  PRE_UNIDAD,PRE_PRE1,PRE_PRE2,PRE_PRE3,PRE_PRE4,PRE_PRE5,PRE_PRE6 FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
         If WS_CALIDAD = 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV,  PRE_UNIDAD,PRE_PRE1,PRE_PRE2,PRE_PRE3,PRE_PRE4,PRE_PRE5,PRE_PRE6 FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND  ART_FAMILIA = " & WS_TIPO & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
         If WS_CALIDAD <> 0 And WS_TIPO = 0 Then
             If LK_CODCIA = "01" Then
                PUB_CODCIA = "03"
                If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
                  aux_codcia = "05" ' TEMPORAL EN 05  DEBE ESTAR EN 04
                Else
                  aux_codcia = "04" ' TEMPORAL EN 05  DEBE ESTAR EN 04
                End If
             ElseIf LK_CODCIA = "03" Then
                PUB_CODCIA = "01"
                aux_codcia = "04"
             ElseIf LK_CODCIA = "02" Then
                PUB_CODCIA = "01"
                aux_codcia = "03"
             Else
                PUB_CODCIA = "01"
                aux_codcia = "03"
             End If
             archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, cia_actual.ARM_STOCK as ARM_STOCK , PRE_EQUIV,  PRE_UNIDAD,PRE_PRE1,PRE_PRE2,PRE_PRE3,PRE_PRE4,PRE_PRE5,PRE_PRE6,cia_otra.ARM_STOCK as ARM_STOCK2, cia_SUBALM.ARM_STOCK as ARM_STOCK3, art_marca, ART_NUMERO FROM ARTI, ARTICULO as cia_actual,ARTICULO as cia_otra, ARTICULO as cia_SUBALM, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD = 'A') AND (ART_KEY = cia_actual.ARM_CODART) AND (ART_KEY = cia_otra.ARM_CODART) AND (ART_KEY = cia_SUBALM.ARM_CODART) AND ART_SITUACION <> 1 and ART_KEY <> 0 AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND (cia_actual.arm_codcia = '" & ws_codcia & "') AND (cia_otra.arm_codcia = '" & PUB_CODCIA & "') AND (cia_SUBALM.arm_codcia = '" & aux_codcia & "') AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
         End If
         If WS_CALIDAD <> 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV,  PRE_UNIDAD,PRE_PRE1,PRE_PRE2,PRE_PRE3,PRE_PRE4,PRE_PRE5,PRE_PRE6 FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND  (ART_CODCIA = ARM_CODCIA) AND ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND  ART_FAMILIA = " & WS_TIPO & " AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      End If
      '//////////////////

    Else
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        numarchi = 3
        If WS_TIPO = 0 And WS_CALIDAD = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK, PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND  ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
        If WS_CALIDAD = 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
        If WS_CALIDAD <> 0 And WS_TIPO = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
        If WS_CALIDAD <> 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_FAMILIA = " & WS_TIPO & " AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
      Else
        numarchi = 0
        If WS_TIPO = 0 And WS_CALIDAD = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION  FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_CODCIA = '" & ws_codcia & "' AND ART_SITUACION <> 1 AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
        If WS_CALIDAD = 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION  FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND  ART_FAMILIA = " & WS_TIPO & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
        If WS_CALIDAD <> 0 And WS_TIPO = 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION  FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
        If WS_CALIDAD <> 0 And WS_TIPO <> 0 Then archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION  FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND  (ART_CODCIA = ARM_CODCIA) AND ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND  ART_FAMILIA = " & WS_TIPO & " AND ART_CALIDAD = " & WS_CALIDAD & "  AND ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      End If
     End If
     PROC_LISVIEW LV_ART
     loc_key = 0
     If LV_ART.Visible Then
      loc_key = 1
     End If
    Exit Sub
End If

'If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
' Exit Sub
'End If

Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If LV_ART.Visible Then
  Set itmFound = LV_ART.FindItem(LTrim(TEXTOVAR.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_ART.ListItems.count Then
      LV_ART.ListItems.Item(LV_ART.ListItems.count).EnsureVisible
   Else
     LV_ART.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
    If LK_CODTRA <> 2401 Then GoTo SALTA2401
  If fraprecios.Visible = False Then
   fraprecios.Top = LV_ART.Top + 3230
   fraprecios.Visible = True
  End If
  lbltitulostock.Caption = Left(lk_descrip_cias, 10)
  lbltitulostock2.Caption = Right(lk_descrip_cias, 10)
  lpstock2.Caption = Int(LV_ART.ListItems.Item(loc_key).SubItems(12))
  lpstock2.Caption = Format(lpstock2.Caption, "##,##0.00")
  lpstocku2.Caption = Format(Val((LV_ART.ListItems.Item(loc_key).SubItems(13))) - (Val(lpstock2.Caption) * Val(LV_ART.ListItems.Item(loc_key).SubItems(9))), "##,##0.000")
  
  lpstock3.Caption = Int(LV_ART.ListItems.Item(loc_key).SubItems(14))
  lpstock3.Caption = Format(lpstock3.Caption, "##,##0.00")
  lpstocku3.Caption = Format(Val((LV_ART.ListItems.Item(loc_key).SubItems(15))) - (Val(lpstock3.Caption) * Val(LV_ART.ListItems.Item(loc_key).SubItems(9))), "##,##0.000")


  lpnombre.Caption = Trim(LV_ART.ListItems.Item(loc_key))
  lpunidad.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(2))
  lpstock.Caption = Int(LV_ART.ListItems.Item(loc_key).SubItems(3))
  lpstocku.Caption = Format(Val((LV_ART.ListItems.Item(loc_key).SubItems(10))) - (Val(lpstock.Caption) * Val(LV_ART.ListItems.Item(loc_key).SubItems(9))), "##,##0.000")
  lpstock.Caption = Format(lpstock.Caption, "##,##0.00")
  lpp1.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(4))
  lpp2.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(5))
  lpp3.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(6))
  lpp4.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(7))
  lpp5.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(8))
  lpp6.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(11))
  lbllote.Caption = Trim(LV_ART.ListItems.Item(loc_key).SubItems(17))
  
  If Val(LV_ART.ListItems.Item(loc_key).SubItems(18)) = 1 Then
    fraprecios.BackColor = QBColor(3)
  ElseIf Val(LV_ART.ListItems.Item(loc_key).SubItems(16)) = 1 Then
    fraprecios.BackColor = QBColor(6)
  Else
    fraprecios.BackColor = QBColor(8)
  End If
    
SALTA2401:


  End If
  Exit Sub
End If

End Sub
Private Sub textovar_LostFocus()
On Error GoTo SALE
calcula_totales
Exit Sub
SALE:
End Sub


Private Sub textovarL_Change()
gridl.Text = textovarl.Text
calcula_totales2
End Sub

Private Sub textovarl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   textovarl.Text = textovar_bak
End If


If (KeyCode > 36 And KeyCode < 41) Or KeyCode = 13 Then
Else
   Exit Sub
End If


   If KeyCode = 38 Then
      If gridl.Row = 2 Then Exit Sub
   End If
   
   If KeyCode = 37 Then
      If gridl.COL = 0 Then Exit Sub
   End If
   If KeyCode = 40 Then
      If gridl.Row = gridl.Rows - 1 Then Exit Sub
   End If
   
If KeyCode = 13 Then
   If gridl.COL = 1 Or gridl.COL = 2 Then
   Else
   Exit Sub
   End If
End If

If gridl.COL = 5 Then
If Val(gridl.TextMatrix(gridl.Row, 5)) > Val(gridl.TextMatrix(gridl.Row, 25)) And Val(gridl.TextMatrix(gridl.Row, 5)) <> 0 Then
   MsgBox "El importe ingresado es ayor que el Saldo...", 48, Pub_Titulo
   If LK_CODTRA = 2770 Then textovarl.Text = ""
End If
End If



SALTO:

gridl.TextMatrix(gridl.Row, gridl.COL) = textovarl.Text

textovarl.Visible = False
gridl.SetFocus



If KeyCode = 38 Then
   gridl.Row = gridl.Row - 1
ElseIf KeyCode = 40 Then
   gridl.Row = gridl.Row + 1
End If


If KeyCode = 37 Then
   gridl.COL = gridl.COL - 1
ElseIf KeyCode = 39 Or KeyCode = 13 Then
      If gridl.COL = 23 Then gridl.COL = 1
      gridl.COL = gridl.COL + 1
End If

End Sub

Private Sub textovarL_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I, car
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
car = Chr(KeyAscii)
car = UCase(car)
KeyAscii = Asc(car)



If gridl.COL = 5 Then
 Consistencias_rich gridl, textovarl, KeyAscii, 2, 4
 textovarl.MaxLength = 17
Else
 textovarl.MaxLength = 0
End If





If KeyAscii <> 13 Then Exit Sub
PROCESA_CELDASL

'If i_codcli.Visible = True Then
'If Trim(gridl.TextMatrix(gridl.Row, 0)) = "" Then Exit Sub
'End If


'If gridl.Col = 18 Then
'   gridl.Col = 5
'   Exit Sub
'End If
'If gridl.Col = 23 Then
'   gridl.Col = 5
'   Exit Sub
'End If

'If gridl.Col = 19 Then
'   calcula_totales2
 '  gridl.Col = 20
 '  Exit Sub
'End If
'If gridl.Col = 20 Then
'   gridl.Col = 21
'   Exit Sub
'End If
'If gridl.Col = 21 Then
'   gridl.Col = 22
'   Exit Sub
'End If
'If gridl.Col = 22 Then
'   gridl.Col = 19
'   Exit Sub
'End If
'If LK_EMP = "PLA" Then
'If gridl.Col = 5 Then
'   calcula_totales2
'   gridl.Col = 23
'   Exit Sub
'End If
'End If
If gridl.COL = 5 Then
       If Val(gridl.TextMatrix(gridl.Row, 5)) > Val(gridl.TextMatrix(gridl.Row, 25)) And Val(gridl.TextMatrix(gridl.Row, 5)) <> 0 Then
          MsgBox "El importe ingresado es ayor que el Saldo...", 48, Pub_Titulo
          If LK_CODTRA = 2770 Then
             Azul3 textovarl, textovarl
             Exit Sub
          End If
       End If
       gridl.COL = 6
       If textovarl.Visible Then
         textovarl.SelStart = 0
         textovarl.SelLength = 2
       End If
       If Val(gridl.TextMatrix(gridl.Row, 5)) <> 0 Then
         i_TEXTONCRE.Text = Trim(gridl.TextMatrix(gridl.Row, 1)) & "/." & Format(gridl.TextMatrix(gridl.Row, 2), "000") & "-" & gridl.TextMatrix(gridl.Row, 3)
         i_codven.Text = Trim(gridl.TextMatrix(gridl.Row, 7))
       End If
ElseIf gridl.COL = 6 Then
       If i_fbg.Visible = True Then
          i_fbg.SetFocus
       Else
       gridl.COL = 5
       End If
End If
        
If gridl.COL = 6 Or gridl.COL = 20 Then calcula_totales2

End Sub

Private Sub textovarL_KeyUp(KeyCode As Integer, Shift As Integer)

'Dim i_tasaa
'Dim plaza_tasa1
'Dim plaza_tasa2
'i_tasaa = 0
If KeyCode = 113 And gridl.COL = 5 And LK_CODTRA <> 2728 Then
   textovarl.Text = Abs(gridl.TextMatrix(gridl.Row, 25))
   calcula_totales2
   If i_fbg.Visible = True Then i_fbg.SetFocus
   Exit Sub
End If

'Dim WS_DIASV, ws_diasa, WS_INTVEN, WS_INTADE
'Dim WK_DIASV As Integer
'Dim WK_DIASA As Integer


'If (KeyCode = 115 Or KeyCode = 119) And (gridl.Col = 5) Then
'Else
'Exit Sub
'End If'


'fila = gridl.Row

'If LK_EMP = "PLA" Then
'   If gridl.TextMatrix(gridl.Row, 26) < LK_FECHA_DIA Then
'      gridl.TextMatrix(gridl.Row, 6) = LK_FECHA_DIA
'   Else
'      gridl.TextMatrix(gridl.Row, 6) = gridl.TextMatrix(gridl.Row, 26)
'   End If
'End If

'If IsDate(gridl.TextMatrix(gridl.Row, 6)) = False Then
'   MsgBox "Fecha esta errada.. Revisar "
'   Exit Sub
'End If

'   gridl.Visible = False
'   textovarl.Visible = False
'   grid_liq.Visible = True
'   grid_liq.Width = 7500
'   grid_liq.Height = 2000
'   grid_liq.Top = gridl.Top
'   grid_liq.Left = 1080
'   grid_liq.Row = 1
'   grid_liq.RowHeight(0) = 500
'   grid_liq.RowHeight(1) = 500
'   grid_liq.Rows = 2
'   grid_liq.Cols = 13
'   grid_liq.ColWidth(0) = 0
'   grid_liq.ColWidth(1) = 1000
'   grid_liq.ColWidth(2) = 1000
'   grid_liq.ColWidth(3) = 0
'   grid_liq.ColWidth(4) = 0
'   grid_liq.ColWidth(5) = 1000
'   grid_liq.ColWidth(6) = 1000
'   grid_liq.ColWidth(7) = 1000
'   grid_liq.TextMatrix(1, 1) = 0
'   grid_liq.TextMatrix(1, 2) = 0
   
'Dim J
'   WS_DIASV = DateDiff("d", gridl.TextMatrix(gridl.Row, 26), LK_FECHA_DIA)
'   If WS_DIASV < 0 Then
'      ws_diasa = DateDiff("d", gridl.TextMatrix(gridl.Row, 26), gridl.TextMatrix(gridl.Row, 6))
'   Else
'      ws_diasa = DateDiff("d", LK_FECHA_DIA, gridl.TextMatrix(gridl.Row, 6))
'   End If
   
'   WK_DIASA = ws_diasa
'   WK_DIASV = WS_DIASV
   
'If WK_DIASV > 0 Then
'   i_tasav.Text = Val(XLL.Cells(WK_DIASV, 1)) + Val(XLL.Cells(WK_DIASV, 2))
   ' PARA PLAZA GRAU
'   If LK_CODTRA = 2728 And LK_EMP = "PLA" Then
'      PLAZA_FLAG_MANUAL = " "
'      plaza_tasa1 = InputBox("Ingrese (%)Tasa  Compensatorio : Días: " & WK_DIASV, "Tasa de Interes", "0")
'      If plaza_tasa1 = "" Then GoTo pasa_tasa
'      plaza_tasa2 = InputBox("Ingrese (%)Tasa  Moratoria     : Días: " & WK_DIASV, "Tasa de Interes", "0")
'      If plaza_tasa2 = "" Then GoTo pasa_tasa
'      If Val(i_tasav.Text) <> (Val(plaza_tasa1) + Val(plaza_tasa2)) Then
'         pub_mensaje = "Importe de Tasas No son Correctas  ...   ¿Desea Continuar de todas maneras ... ?"
'         Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
'         If Pub_Respuesta = vbNo Then
'              GoTo pasa_tasa
'              Exit Sub
'         End If
'         PLAZA_FLAG_MANUAL = "M"
'      End If
'pasa_tasa:
'   End If
'End If

'   WS_INTADE = 0
'   WS_INTVEN = 0
'
'  If KeyCode = 119 Then  ' AMORTIZACION SOLO
'         FORMGEN.i_importe.Text = gridl.TextMatrix(fila, 25)
'         FORMGEN.i_importe_amort.Text = gridl.TextMatrix(fila, 5)
'         WS_INTVEN = Format((i_importe.Text * Val(i_tasav.Text) / 100), "0.00")
'         WS_INTADE = (i_importe.Text - i_importe_amort.Text) * Val(i_tasaa) / 100
'         WS_INTADE = Format(WS_INTADE, "0.00")
         
'         If Val(WS_INTVEN) < 0 Then WS_INTVEN = 0
'         If Val(WS_INTADE) < 0 Then WS_INTADE = 0
       
'         WS_IMPORTE = Val(i_importe_amort.Text) + Val(WS_INTVEN) + Val(WS_INTADE)
'         GoTo saltito
'  End If
      
      
'      WS_IMPORTE = 0
      
'      WK_DIASV = WS_DIASV
'      WK_DIASA = ws_diasa
'      If WK_DIASV < 0 Then WK_DIASV = 0
'      If WK_DIASA < 0 Then WK_DIASA = 0
         
            
'      i_importe_amort.Text = (Val(gridl.TextMatrix(fila, 5)) - Val(gridl.TextMatrix(fila, 4)) * (WK_DIASV * Val(i_tasav.Text) / 100 + WK_DIASA * i_tasaa / 100)) / (1 - Val(WK_DIASA) * Val(i_tasaa) / 100)
      
'      J = 0
'      Do Until WS_IMPORTE = Val(gridl.TextMatrix(fila, 5)) Or J > 5000
'      i_importe_amort = Format(i_importe_amort.Text, "0.00")
'      WS_INTVEN = Format((Val(gridl.TextMatrix(fila, 25)) * Val(i_tasav.Text) / 100), "0.00")
'      WS_INTADE = 0 ' (gridl.TextMatrix(fila, 4) - i_importe_amort.Text) * i_tasaa * WK_DIASA / 100
'      WS_INTADE = Format(WS_INTADE, "0.00")
      
'      If Val(WS_INTVEN) < 0 Then
'         WS_INTVEN = 0
'      End If
'      If Val(WS_INTADE) < 0 Then
'         WS_INTADE = 0
'      End If
      
      
'      J = J + 1
'      WS_IMPORTE = Val(i_importe_amort.Text) + Val(WS_INTADE) + Val(WS_INTVEN)
'
'      If WS_IMPORTE < Val(gridl.TextMatrix(fila, 5)) Then
'         i_importe_amort.Text = Val(i_importe_amort.Text) + 0.01
'      ElseIf WS_IMPORTE > Val(gridl.TextMatrix(fila, 5)) Then
'         i_importe_amort.Text = Val(i_importe_amort.Text) - 0.01
'      End If
'
'      Loop
      
'      If Val(i_importe_amort.Text) < 0 Or J >= 5000 Then
'         MsgBox "Monto Insuficiente...aumentar..."
'         i_importe_amort.Text = 0
'         WS_INTADE = 0
'         WS_INTVEN = 0
'         gridl.Visible = True
'         Exit Sub
'      End If
'saltito:
'   grid_liq.TextMatrix(1, 1) = i_importe_amort.Text
'   grid_liq.TextMatrix(1, 2) = WS_INTVEN
'   grid_liq.TextMatrix(1, 3) = WS_INTADE
   
'   grid_liq.TextMatrix(1, 5) = Val(grid_liq.TextMatrix(1, 1)) + Val(grid_liq.TextMatrix(1, 2)) + Val(grid_liq.TextMatrix(1, 3))
'   grid_liq.TextMatrix(1, 6) = WS_DIASV
'   grid_liq.TextMatrix(1, 7) = ws_diasa


End Sub

Private Sub textovarl_LostFocus()
calcula_totales2

End Sub

Private Sub Tipo_Cambio_Click()
Dim WRES
  WS_TIPO_CAMBIO = LK_TIPO_CAMBIO
sigue:
  WRES = InputBox("Ingrese Tipo de cambio :", , LK_TIPO_CAMBIO)
  If WRES = "" Then Exit Sub
  If Val(WRES) <= 0 Then GoTo sigue
  
  WS_TIPO_CAMBIO = WRES
  LK_TIPO_CAMBIO = WRES
  'MDIForm1.StatusBar1.Panels(3).Text = "T.C.= S/. " + Format(LK_TIPO_CAMBIO, "0.0000")
  If (Abs(WS_TIPO_CAMBIO - LK_TIPO_CAMBIO) * 100 / LK_TIPO_CAMBIO) > 10 Then
     WS_TIPO_CAMBIO = 0
     MsgBox "Cambiar Tipo De Cambio... Mucha diferencia"
     Exit Sub
  End If
Tipo_Cambio.Caption = "T/C = " & WS_TIPO_CAMBIO
If LK_CODTRA = 5345 Then
   If Trim(LMONEDAB.Caption) = "$" Then
      i_importe_amort.Text = Format(Val(i_importe.Text) * WS_TIPO_CAMBIO, "0.00")
   Else
     i_importe_amort.Text = Format(Val(i_importe.Text) / WS_TIPO_CAMBIO, "0.00")
   End If
   Exit Sub
End If
If LK_CODTRA = 5714 And pub_signo_ccm = 1 And Trim(LMONEDAB.Caption) = "$." Then
   i_importe_amort.Text = Format(Val(i_importe.Text) * WS_TIPO_CAMBIO, "0.00")
   Exit Sub
End If

If pub_signo_car <> 0 And pub_signo_ccm <> 0 Then calcula_totales2
If i_codban2.Visible = True Then
   PUB_CODBAN = Val(i_codban.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CCM_LLAVE
   If ccm_llave.EOF Then Exit Sub
   WS_MONEDA_CCM = ccm_llave!CCM_MONEDA
   PUB_CODBAN = Val(i_codban2.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CCM_LLAVE
   If ccm_llave.EOF Then Exit Sub

   If ccm_llave!CCM_MONEDA = WS_MONEDA_CCM Then
      i_importe_amort.Text = i_importe.Text
      'Exit Sub
   Else
   If ccm_llave!CCM_MONEDA = "D" Then
      lmoneda.Caption = " $."
      'LMONEDAB.Caption = " $"
      i_importe_amort.Text = Format(Val(i_importe.Text) / WS_TIPO_CAMBIO, "0.00")
   Else
      lmoneda.Caption = "S/."
  '    LMONEDAB.Caption = "S/."
      i_importe_amort.Text = Format(Val(i_importe.Text) * WS_TIPO_CAMBIO, "0.00")
   End If
   End If
   i_importe_amort.SetFocus
Exit Sub
End If
   
calcula_totales2
End Sub

Private Sub TRANS_Click()
TRANS.Locked = False
End Sub

Private Sub TRANS_GotFocus()
TRANS.Locked = False
'Azul TRANS, TRANS
End Sub

Private Sub TRANS_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not LV_TRA.Visible Then
 Exit Sub
End If
If KeyCode = 13 Then Exit Sub
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And TRANS.Text = "" Then
  loc_key = 1
  Set LV_TRA.SelectedItem = LV_TRA.ListItems(loc_key)
  LV_TRA.ListItems.Item(loc_key).Selected = True
  LV_TRA.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_TRA.ListItems.count Then loc_key = LV_TRA.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_TRA.ListItems.count Then loc_key = LV_TRA.ListItems.count
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
  LV_TRA.ListItems.Item(loc_key).Selected = True
  LV_TRA.ListItems.Item(loc_key).EnsureVisible
  TRANS.Text = Trim(LV_TRA.ListItems.Item(loc_key).Text) & " "
  DoEvents
  TRANS.SelStart = Len(TRANS.Text)
  DoEvents
fin:
End Sub

Private Sub Trans_KeyPress(KeyAscii As Integer)
Dim loc_codoper As String
Dim car As String
Dim SN As String
Dim Control As Object
Dim PRIMER As Integer
Dim pos1, pos2, DIF As Integer
Dim CARAC, CARAC2 As String
Dim I, J, WS_SECUENCIA As Integer
inicio:
'SOLO_DECIMAL TRANS, KeyAscii
pub_flag_cambio = 0
If KeyAscii = 27 Then
  TRANS.Text = ""
  Exit Sub
End If
If KeyAscii = 13 Then
  If TRANS.Text = "" Then
    TRANS.SetFocus
    Exit Sub
  End If
End If
If KeyAscii <> 13 Then
   Exit Sub
End If
If Trim(TRANS.Text) = "5555" And (LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR") Then
  ACT_FORMGEN
'  Load FrmActualiza
'  FrmActualiza.Show 1
  GoTo SALIR
End If


If IsNumeric(FORMGEN.TRANS.Text) Then
  pu_subtra = Trim(FORMGEN.TRANS.Text)
Else
  pu_subtra = LV_TRA.ListItems.Item(loc_key).SubItems(1)
End If
LV_TRA.Visible = False
FORMGEN.TRANS.Text = pu_subtra
SQ_OPER = 4
LEER_TRA_LLAVE
If subtra_llave.EOF Then
  MsgBox "No Existe Codigo de Operación.", 48, Pub_Titulo
  Azul TRANS, TRANS
  Exit Sub
End If
loc_codoper = Val(subtra_llave!TRA_KEY)



If LK_CODTRA = Val(loc_codoper) And grabar.Enabled = True Then
   If grid_fac.Visible Then grid_fac.SetFocus
   Exit Sub
End If

If Trim(loc_codoper) = "2101" Or Trim(loc_codoper) = "2105" Then
   MsgBox "Seleccione opcion de Grifo"
   GoTo SALIR
End If

If Trim(loc_codoper) = "2107" Then
   MsgBox "No Pertenece a este Módulo.", 48, Pub_Titulo
   GoTo SALIR
End If


If tra_llave.EOF Then
   GoTo SALTARIN
End If

frafondo.Visible = True
DoEvents
cancela_todo
FORMGEN.fragas.Visible = False
i_importe.Visible = False
grid_fac.Rows = 3
pasa_cabeza
' desaparece a todos los campos actuales
nn = 2
m_ind = 0
Do Until Val(tra_llave(nn)) = 0 Or nn = 62
         m_ind = m_ind + 1
         FORMGEN.LABELGEN(m_ind).Visible = False
         NUMERO = TABLA_TAG(tra_llave(nn))
         If TypeOf FORMGEN.Controls(NUMERO) Is TextBox Then
            FORMGEN.Controls(NUMERO).Text = ""
         End If
         FORMGEN.Controls(NUMERO).Visible = False
nn = nn + 4
Loop
SALTARIN:

FORMGEN.i_def.Visible = True
i_def.Clear

PUB_CODTRA = Val(loc_codoper)
LK_CODTRA = PUB_CODTRA
If Trim(loc_codoper) <> "" Then
   SQ_OPER = 1
   LEER_TRA_LLAVE
   If tra_llave.EOF Then
      MsgBox "TRANSACCION NO EXISTE...", 48, Pub_Titulo
      Azul TRANS, TRANS
      GoTo SALIR
   Else
   If tra_llave!TRA_KEY = 1409 And LK_FLAG_SOS = "A" Then
      MsgBox "TRANSACCION NO EXISTE...", 48, Pub_Titulo
      Azul TRANS, TRANS
      GoTo SALIR
   End If
     If Nulo_Valors(tra_llave!TRA_FLAG_ACTIVO) <> "A" Then
        Azul TRANS, TRANS
        GoTo SALIR
      End If
     LK_NOMTRA = tra_llave(1)
      LK_CODTRA = tra_llave(0)
      FORMGEN.nomtra.Caption = Trim(LK_NOMTRA)
   End If
End If
SN = "N"
I = 0
Do Until SN = "S" Or I = 10
   I = I + 1
   J = 0
   Do Until SN = "S" Or J = 10
     J = J + 1
     If lk_GRUPOS(J) = tra_llave(92 + I) And lk_GRUPOS(J) <> 0 Then
         SN = "S"
     End If
   Loop
Loop

'If SN = "S" Then
'   GoTo CACHETE
'End If


   J = 1
   Do Until lk_CODTRAS(J) = "" Or SN = "Y" Or J = 10
      If Left(lk_CODTRAS(J), 4) = tra_llave(0) Then
         SN = "Y"
         Exit Do
      End If
   J = J + 1
   Loop

If SN = "N" And LK_CODUSU <> "ADMIN" Then
   MsgBox "No Tiene los Derechos Asignados"
   GoTo SALIR
End If
i_fbg.ListIndex = 0
PS_GEN(0) = 0
GEN.Requery
PASO:

FORMGEN.nomtra.Visible = True
CARAC = loc_codoper
LLENA_CAMPOS2
pasa_cabeza
If gridl.Visible = True Then
   gridl.Clear
   gridl.Rows = 3
   pasa_cabeza_L
End If
FORMGEN.i_camal.Visible = False
If LK_CODTRA = 2410 Then
'  If pub_signo_arm = 0 Then
  Frame4.Visible = False
'  Else
'  Frame4.Visible = True
'  End If
End If
If LK_CODTRA = 1401 Then
   cheigv.Visible = True
   cheigv.Enabled = False
End If

If LK_CODTRA = 2410 Or LK_CODTRA = 2412 Or LK_CODTRA = 2418 Then
   FORMGEN.Frame4.Width = 10100
   FORMGEN.grid_fac.Width = FORMGEN.Frame4.Width - 100
   FORMGEN.Frame4.Height = 2100
   FORMGEN.grid_fac.Height = 1700
Else
   
   FORMGEN.Frame4.Width = LBLDET.Width - 100
   FORMGEN.grid_fac.Width = FORMGEN.Frame4.Width - 200
   FORMGEN.Frame4.Height = 3970
   FORMGEN.grid_fac.Height = 2700
End If
If (LK_CODTRA = 2412 Or LK_CODTRA = 2410) Then gridl.Visible = False
If LK_CODTRA = 2407 Then
   pub_cadena = "SELECT * FROM FFFART WHERE  FFF_CODCIA= ? "
   Set PSFFF_ART = CN.CreateQuery("", pub_cadena)
   PSFFF_ART.rdoParameters(0) = " "
   Set FFF_ART = PSFFF_ART.OpenResultset(rdOpenKeyset, rdConcurValues)
End If
If LK_CODTRA = 1401 Then
    pub_flag_cambio = 1
    PUB_NUMSER = Nulo_Valor0(par_llave!PAR_SER_KARDEX)
    FORMGEN.i_numser.Text = PUB_NUMSER
    FORMGEN.i_situacion.Clear
    PUB_TIPREG = 70
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
        FORMGEN.i_situacion.AddItem tab_mayor!tab_nomlargo & String(40, " ") & tab_mayor!tab_numtab
        tab_mayor.MoveNext
    Loop
    If FORMGEN.i_situacion.ListCount <> 0 Then FORMGEN.i_situacion.ListIndex = 0
    i_num_ini.Text = " "
    i_num_ini.Visible = True
    i_num_ini.Top = 1750
    i_num_ini.Left = 3100
    LABELGEN(16).Top = i_num_ini.Top
    LABELGEN(16).Left = i_num_ini.Left - 1000
    LABELGEN(16).Caption = "Cod.Distrib."
    LABELGEN(16).Visible = True
    
    i_num_lote.Text = " "
    i_num_lote.Top = 1750
    i_num_lote.Left = 5700
    i_num_lote.Visible = True
    LABELGEN(17).Top = i_num_lote.Top
    LABELGEN(17).Left = i_num_lote.Left - 1000
    LABELGEN(17).Caption = "Nº.Pedido."
    LABELGEN(17).Visible = True
    
    i_situacion.Top = 500
    i_situacion.Left = 7300
    LABELGEN(15).Top = i_situacion.Top
    LABELGEN(15).Left = i_situacion.Left - 1200
    LABELGEN(15).Caption = "ALQ.Almacen"
    LABELGEN(15).Visible = True
    
    i_situacion.Visible = True
End If
If LK_CODTRA = 2414 Then
   PUB_NUMSER = Nulo_Valor0(par_llave!par_serie)
   FORMGEN.i_numser.Text = PUB_NUMSER
End If


'FORMGEN.TRANS.Text = CARAC

If (LK_CODTRA = 5714 Or LK_CODTRA = 2748) And (LK_EMP = "3AA" Or LK_EMP = "PIU") Then
 fragas.Visible = True
 If LK_CODTRA = 2735 Then
 fragas.Left = 10
 Else
 fragas.Left = 10
 End If
 fragas.Top = 4280
 FORMGEN.imp_gas1.Visible = True
 FORMGEN.cta_gas1.Visible = True
 If LK_CODTRA = 5714 Then
  FORMGEN.imp_gas2.Visible = False
  FORMGEN.cta_gas2.Visible = False
 Else
  FORMGEN.imp_gas2.Visible = True
  FORMGEN.cta_gas2.Visible = True
 End If
 
End If
NUMERO = TABLA_TAG(tra_llave(2))
WS_INDICE_RETORNO = NUMERO
SN = "S"
If SN = "S" Then
SQ_OPER = 2
PUB_CODTRA = LK_CODTRA
LEER_SUT_LLAVE
'FORMGEN.i_def2.Clear
Do Until SUT_MAYOR.EOF
  If LK_CODTRA = 2725 And SUT_MAYOR!sut_secuencia = 100 And LK_CODUSU <> "ADMIN" Then GoTo SALE
  If LK_FLAG_SOS = "A" And LK_CODTRA = 2401 And SUT_MAYOR!sut_secuencia = 0 Then GoTo SALE
   If LK_EMP = "HER" Then
      If SUT_MAYOR!sut_secuencia = 2 And LK_CODTRA = 2401 Then
         If (LK_CODUSU = "OPER20" Or LK_CODUSU = "SUPERVISOR" Or LK_CODUSU = "ADMIN") Then
            FORMGEN.i_def.AddItem SUT_MAYOR!sut_secuencia & ".-" & SUT_MAYOR!sut_descripcion
         End If
      Else
          FORMGEN.i_def.AddItem SUT_MAYOR!sut_secuencia & ".-" & SUT_MAYOR!sut_descripcion
      End If
   Else
     'If LK_CODTRA = 2748 Then FORMGEN.i_def2.AddItem SUT_MAYOR!SUT_DESCRIPCION & String(80, " ") & SUT_MAYOR!SUT_SECUENCIA
     FORMGEN.i_def.AddItem SUT_MAYOR!sut_secuencia & ".-" & SUT_MAYOR!sut_descripcion
   End If
SALE:
SUT_MAYOR.MoveNext
Loop
GoTo todo
End If


CARAC2 = "N"
pos1 = 5
pos2 = 99
Do Until CARAC2 = "S" Or pos2 = 0
   pos1 = pos1 + 1
   pos2 = InStr(pos1, lk_CODTRAS(J), ".", 1)
   If pos2 = 0 Then
      pos2 = pos1 + 3
      CARAC2 = "S"
   End If
   DIF = pos2 - pos1
   CARAC = Mid(lk_CODTRAS(J), pos1, DIF)
   PUB_SECUENCIA = Val(CARAC)
   SQ_OPER = 1
   PUB_CODTRA = LK_CODTRA
   LEER_SUT_LLAVE
   FORMGEN.i_def.AddItem SUT_LLAVE!sut_secuencia & ".-" & SUT_LLAVE!sut_descripcion
pos1 = pos2
Loop
   
todo:
pos1 = InStr(1, loc_codoper, ".", 1)
pos1 = pos1 + 1
If pos1 = 1 Then
   PUB_SECUENCIA = 0
Else
   CARAC = Mid(loc_codoper, pos1)
   PUB_SECUENCIA = Val(CARAC)
End If
If PUB_SECUENCIA = 0 Then
   I = 0
   GoTo listo
End If
SN = "N"
I = 0

Do Until SN = "S" Or I > FORMGEN.i_def.ListCount - 1
   pos1 = InStr(1, FORMGEN.i_def.List(I), ".", 1)
   pos1 = pos1 - 1
   CARAC = Mid(FORMGEN.i_def.List(I), 1, pos1)

   If PUB_SECUENCIA = Val(CARAC) Then
      SN = "S"
      Exit Do
   End If
I = I + 1
Loop

If SN = "N" And LK_CODUSU <> "ADMIN" Then
   MsgBox "No Tiene los Derechos Asignados "
   GoTo SALIR
End If



listo:

If i_def.ListCount = 0 Then
   MsgBox "Falta Definir Sub TRansacciones"
   Exit Sub
End If

i_numfac.Locked = True
i_numser.Locked = True
i_cambio.Value = 0
i_cambio.Visible = False
i_numfac.BackColor = QBColor(7)
i_numser.BackColor = QBColor(7)
FORMGEN.i_def.Enabled = True
BolFac.Caption = ""
grabar.Enabled = True
cancelar.Enabled = True
FORMGEN.i_def.ListIndex = 0
FORMGEN.i_def.SetFocus
' SendKeys "%{UP}"
If LK_CODTRA = 1401 And i_fecha_oper.Visible Then i_fecha_oper.Text = Format(LK_FECHA_DIA, "dd/mm/yy")
 If LK_EMP = "3AA" And LK_CODTRA = 2401 Then
 'i_def.ListIndex = 1
 'i_DEF_KeyPress 13
   frafondo.Visible = False
' Exit Sub
End If
i_fecha_oper.Text = Format(LK_FECHA_DIA, "dd/mm/yy")
i_fecha_can.Text = Format(LK_FECHA_DIA, "dd/mm/yy")

If LK_CODTRA = 2748 And (LK_EMP = "3AA" Or LK_EMP = "PIU") Then
 i_fecha_can.Top = 3200
 i_fecha_can.Left = 1080
 i_fecha_can.Visible = True
 i_fecha_can.Text = Format(LK_FECHA_DIA, "dd/mm/yy")
 LABELGEN(15).Caption = "Fec. Pago:"
 LABELGEN(15).Visible = True
 LABELGEN(15).Left = i_fecha_can.Left - 1000
 LABELGEN(15).Top = i_fecha_can.Top
 'i_TEXTONCRE
 i_importe.Visible = True
 i_importe.Left = 8000
 i_importe.Top = 2400
 i_importe.Width = 1095
 i_importe.Height = 360
 i_importe.Text = ""
 
 LABELGEN(17).Caption = "Exone.:"
 LABELGEN(17).Visible = True
 LABELGEN(17).Left = i_importe.Left - 800
 LABELGEN(17).Top = i_importe.Top
 
 i_fecha_oper.Top = 1500
 i_fecha_oper.Left = 1080
 LABELGEN(16).Caption = "Fec. Contable:"
 LABELGEN(16).Visible = True
 LABELGEN(16).Left = i_fecha_oper.Left - 1000
 LABELGEN(16).Top = i_fecha_oper.Top
 
 i_fecha_oper.Visible = True
 
 i_concepto.Visible = True
 i_concepto.Text = ""
 i_concepto.Top = 3600
 i_concepto.Left = 1080
 'LABELGEN(18).Caption = "Concepto:"
 'LABELGEN(18).Visible = True
 'LABELGEN(18).Left = i_concepto.Left - 1000
 'LABELGEN(18).Top = i_concepto.Top
  
End If
If LK_CODTRA = 2580 Then
    pub_cadena = "SELECT SUT_SECUENCIA, SUT_DESCRIPCION FROM SUB_TRANSA WHERE SUT_CODTRA = ? ORDER BY SUT_SECUENCIA"
    Set PS_CAMBIO = CN.CreateQuery("", pub_cadena)
    PS_CAMBIO(0) = 0
    Set far_cambio = PS_CAMBIO.OpenResultset(rdOpenKeyset, rdConcurValues)
    PS_CAMBIO(0) = 2401
    far_cambio.Requery
    i_condi.Clear
    i_condi.AddItem "Opcional" & String(60, " ") & "-1"
    Do Until far_cambio.EOF
      i_condi.AddItem far_cambio!sut_descripcion & String(60, " ") & far_cambio!sut_secuencia
      far_cambio.MoveNext
    Loop
     i_condi.Width = 2950
End If
If LK_CODTRA = 2401 Then
   i_condi.Width = 1800
   i_condi.Clear
Else
   PUB_CODUSU = LK_CODUSU
End If

If LK_EMP = "HER" And LK_CODTRA = 2401 Then
   ' pub_cadena = "SELECT far_codven, far_signo_arm, FAR_CANTIDAD, FAR_DESCRI, FAR_CODART, far_CODCIA,FAR_CP,FAR_CODCLIE FROM facart WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND FAR_NUMFAC = ? AND FAR_NUMSER = ? AND FAR_ESTADO <> 'E' " ' ORDER BY  FAR_NUMSER, FAR_NUMFAC, FAR_NUMSEC"
   ' Set PS_CAMBIO = CN.CreateQuery("", pub_cadena)
   ' PS_CAMBIO (0)
   ' PS_CAMBIO (0)
   ' Set far_cambio = PS_CAMBIO.OpenResultset(rdOpenKeyset, rdConcurValues)
End If



frafondo.Visible = False
Exit Sub
SALIR:
grabar.Enabled = False
cancelar.Enabled = False
FORMGEN.i_def.Visible = False
nomtra.Caption = ""
FORMGEN.i_numplan.Text = ""
FORMGEN.BolFac.Caption = ""
i_def.Clear
frafondo.Visible = False

End Sub

Private Sub trans_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(TRANS.Text) = 0 Then
   LV_TRA.Visible = False
   Exit Sub
End If
If LV_TRA.Visible = False Then
    var = Asc(TRANS.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 9
    archi = "SELECT TRA_KEY, TRA_DESCRIPCION, TRA_SUBKEY FROM TRANSACCION WHERE TRA_FLAG_ACTIVO = 'A' AND (TRA_KEY >= 0  AND TRA_KEY < 8000) AND TRA_DESCRIPCION BETWEEN '" & TRANS.Text & "' AND  '" & var & "' ORDER BY TRA_DESCRIPCION"
    PROC_TRANSA LV_TRA
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If LV_TRA.Visible Then
  Set itmFound = LV_TRA.FindItem(LTrim(TRANS.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_TRA.ListItems.count Then
      LV_TRA.ListItems.Item(LV_TRA.ListItems.count).EnsureVisible
   Else
     LV_TRA.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If


End Sub

Private Sub TRANS_LostFocus()
TRANS.Locked = True

End Sub
Public Function consis() As Boolean
Dim wtotal As Currency
Dim wigv  As Currency
Dim WSUMA As Currency
Dim subtotal As Currency
Dim ww_so As String * 1
Dim CONTADOR As Integer
Dim filafac As Integer
Dim FILAXX As Integer
loc_secuencia = SUT_LLAVE!sut_secuencia
If LK_CODTRA = 2401 And PUB_PROCESO = 1 Then GoTo SALTARIN
If i_cambio.Visible = False Then i_cambio.Value = 0
If (LK_CODTRA = 2410 Or LK_CODTRA = 2412) And LK_EMP = "PIU" And Val(SUT_LLAVE!sut_secuencia) = 0 Then
  If Val(gridl.TextMatrix(1, 5)) = 0 Then
    pub_signo_car = 0
  End If
End If
ww_so = "X"
PUB_SO = "X"

If PUB_TIPMOV = 10 Or PUB_TIPMOV = 20 Then
   PUB_SO = ""
   If i_fbg.Text = "F" Or i_fbg.Text = "B" Then PUB_SO = "A"
End If
If LK_CODTRA = 2750 Then
 If far_menor4.EOF Then
   MsgBox "NO PROCEDE ... ", 48, Pub_Titulo
   GoTo salirf
 End If
 If car_far.EOF Then
   MsgBox "NO PROCEDE ... ", 48, Pub_Titulo
   GoTo salirf
 End If
 If car_far!CAR_CODCLIE <> Val(i_codcli.Text) Then
   MsgBox "Documento Tiene Amortizaciones .. no procede ", 48, Pub_Titulo
   GoTo salirf
 End If
 If car_far!car_importe <> car_far!CAR_IMP_INI Then
  MsgBox "Documento Tiene Amortizaciones .. no procede ", 48, Pub_Titulo
  GoTo salirf
 End If
End If
If LK_CODTRA = 1409 Then
   PUB_SO = ""
   If i_fbg.Text = "F" Or i_fbg.Text = "B" Then PUB_SO = "A"
End If

If LK_CODTRA = 2745 And PUB_SECUENCIA = 10 Then
   PUB_SO = "A"
End If


If NJ.Visible = True And LK_ACTIVA = "A" Then
   If NJ.Text = "S" Or NJ.Text = "N" Then
   Else
      MsgBox "Confirmar la Operación...", 48, Pub_Titulo
      GoTo salirf
   End If
   If NJ.Text = "S" Then PUB_SO = "A"
   If NJ.Text = "N" Then PUB_SO = " "
End If
If LK_CODTRA = 2725 Or LK_CODTRA = 2735 Then PUB_SO = ""
If LK_CODTRA = 2720 Then
  
  If Val(i_numfac_c.Text) = 0 Then
    MsgBox "No Procede no tiene documento...", 48, Pub_Titulo
    GoTo salirf
  Else
    SQ_OPER = 1
    PU_TIPMOV = 10
    pu_codcia = LK_CODCIA
    PU_NUMSER = Val(i_numser_c.Text)
    PU_FBG = i_fbg.Text
    PU_NUMFAC = Val(i_numfac_c.Text)
    LEER_FAR_LLAVE
    If far_llave.EOF Then
        MsgBox "Documento no Existe.", 48, Pub_Titulo
        Azul i_numser_c, i_numser_c
        GoTo salirf
    End If
    far_llave.MoveLast
    If Trim(far_llave!far_estado) = "E" Then
        MsgBox "Documento esta Anulado.", 48, Pub_Titulo
        Azul i_numser_c, i_numser_c
        GoTo salirf
    End If
    If Val(far_llave!far_codclie) <> Val(i_codcli.Text) Then
        MsgBox "El Documento no pertenece a este Cliente.", 48, Pub_Titulo
        Azul i_numser_c, i_numser_c
        GoTo salirf
    End If
    
End If
 

End If

' agrege acv
If LK_CODTRA = 2401 Or LK_CODTRA = 2725 Or LK_CODTRA = 2735 Then
 If Not i_concepto.Visible Then
   i_concepto.Text = ""
 End If
 If pub_signo_ccm <> 0 Then
   If Val(i_codban.Text) = 0 Then
     MsgBox "Banco no Corresponde...", 48, Pub_Titulo
     FORMGEN.i_codban.SetFocus
     GoTo salirf
   End If
   If Val(i_chenum.Text) = 0 Then
       MsgBox "Falta Nro de Dep/Ch/..", 48, Pub_Titulo
       i_chenum.SetFocus
       GoTo salirf
   End If
 End If
   
End If
' AGREGE ACV
If LK_CODTRA = 2401 Then
  If dir(GEN!gen_ruta_otros, vbDirectory) = "" Then
    MsgBox "Verifique las conexiones de Red, faltan las unidades de Reportes ", 48, Pub_Titulo
    GoTo salirf
  End If
 If Val(i_codcli.Text) <> 1 Then
   If Trim(i_dircli.Text) = "" Then  ' direccion de almacen
      MsgBox "Verificar el cliente no tiene direccion de almacen de destino", 48, Pub_Titulo
      i_moneda.SetFocus
      GoTo salirf
   End If
 End If
 If Val(Right(cmdtipo.Text, 8)) = 0 Then
    MsgBox "Verificar la Condición de la Venta", 48, Pub_Titulo
    cmdtipo.SetFocus
    GoTo salirf
 End If
End If
If LK_CODTRA = 1401 Then
 If i_ds.Text = "S" Then
    If Trim(UCase(i_moneda.Text)) <> "S/." Then
      MsgBox "Verificar la Moneda para registrar...", 48, Pub_Titulo
      i_moneda.SetFocus
      GoTo salirf
    End If
 Else
    If Trim(UCase(i_moneda.Text)) <> "$" Then
      MsgBox "Verificar la Moneda para registrar...", 48, Pub_Titulo
      i_chenum.SetFocus
      GoTo salirf
    End If
 End If
 If PUB_TIPDOC <> "FA" Then
     MsgBox "El Tipo de Documento es Incorrecto, Veficar", 48, Pub_Titulo
     i_def.SetFocus
     GoTo salirf
 End If
End If
If LK_FLAG_GRIFO = "A" And LK_CODTRA <> 1111 Then
If LK_CODTRA = 1401 And (Val(FORMGEN.i_turno.Text) > 3 Or Val(FORMGEN.i_turno.Text) <= 0) Then
   MsgBox "Turno NO Existe...", 48, Pub_Titulo
   FORMGEN.i_turno.SetFocus
   GoTo salirf
End If
End If
If i_fecha_oper.Visible Then
  If Not IsDate(i_fecha_oper.Text) Then
    MsgBox "Fecha no procede ...", 48, Pub_Titulo
    Azul2 i_fecha_oper, i_fecha_oper
    GoTo salirf
  End If
End If
If LK_CODTRA = 2748 And i_cambio.Value = 0 Then
    SQ_OPER = 8
    PU_TIPMOV = PUB_TIPMOV
    pu_codcia = LK_CODCIA
    PU_NUMSER = Val(i_numser_c.Text)
    PU_FBG = "K"
    PU_FBG2 = " "
    PU_NUMFAC = Val(i_numfac_c.Text)
    LEER_FAR_LLAVE
    If Not far_menor4.EOF Then
      Do Until far_menor4.EOF
      If Val(far_menor4!far_codclie) = Val(i_codcli.Text) Then
         MsgBox " Documento ya Emitido, Secuencia Nro :" & Format(far_menor4!far_numfac, "000"), 48, Pub_Titulo
         i_numfac_c.Text = ""
         i_numser_c.Text = ""
         Azul i_numser_c, i_numser_c
         GoTo salirf
      End If
      far_menor4.MoveNext
      Loop
    End If
End If



If (LK_EMP = "3AA" Or LK_EMP = "PIU") And LK_CODTRA = 2748 Then
  wigv = redondea(Val(i_bruto2.Text) * (LK_IGV / 100))
  If Abs(wigv - Val(i_impto2.Text)) > 0.01 Or Abs(wigv - Val(i_impto2.Text)) < 0.01 Then
'    MsgBox "Igv ..Revisar...", 48, Pub_Titulo
'    FORMGEN.i_impto2.SetFocus
'   GoTo salirf
  End If
  wtotal = Val(i_bruto2.Text) + Val(i_importe.Text) + Val(i_impto2.Text)
  If wtotal <> Val(i_neto.Text) Then
   MsgBox "Total  ..Revisar...", 48, Pub_Titulo
   FORMGEN.i_impto2.SetFocus
   GoTo salirf
  End If

 If pub_signo_car <> 0 And Val(i_dias.Text) = 0 Then
    i_dias.Text = Trim(i_dias.Text)
    MsgBox "Falta Nro de Dias ...", 48, Pub_Titulo
    i_dias.SetFocus
    GoTo salirf
 End If
 If Val(i_numfac_c.Text) = 0 Then
    MsgBox "Falta Nro Documento ...", 48, Pub_Titulo
    i_numfac_c.SetFocus
    GoTo salirf
 End If
 If Val(i_importe_amort.Text) = 0 Then
    MsgBox "Falta Importe de Documento ...", 48, Pub_Titulo
    i_importe_amort.SetFocus
    GoTo salirf
 End If
 
 If fragas.Visible Then
    WSUMA = Val(imp_gas1.Text) + Val(imp_gas2.Text)
    If WSUMA <> (Val(i_bruto2.Text) + Val(i_importe.Text)) Then
      MsgBox "No procede los importes contables con el Bruto...", 48, Pub_Titulo
      imp_gas1.SetFocus
      GoTo salirf
    End If
 End If
End If

If LK_EMP = "3AA" And (LK_CODTRA = 2748) Then
 If Not IsDate(i_fecha_can.Text) Then
    MsgBox "Fecha no Procede...", 48, Pub_Titulo
    Azul2 i_fecha_can, i_fecha_can
    GoTo salirf
 End If
 If pub_signo_ccm <> 0 Then
   If Anexo_con7() = False Then
      MsgBox "Falta datos del Banco ...", 48, Pub_Titulo
      GoTo salirf
   End If
 End If

End If

If pub_signo_ccm <> 0 Then
  If i_chenum.Visible And Val(i_chenum.Text) = 0 Then
    MsgBox "No procede... Cheque tiene que tener valor ...", 48, Pub_Titulo
    i_chenum.SetFocus
    GoTo salirf
  End If
End If

If LK_CODTRA = 2409 Then
  If Trim(FORMGEN.i_cias.Text) = "" Then
    MsgBox "No procede Configurar la Compañia de Destino ", 48, Pub_Titulo
    GoTo salirf
  End If
End If
If Val(Nulo_Valor0(SUT_LLAVE!sut_codban)) = 1 And i_codban.Visible = False Then
   i_codban.Text = 1
End If

If i_fecha_compra.Visible = True Then
   If IsDate(i_fecha_compra.Text) = False Then
        MsgBox "Fecha Errada ...", 48, Pub_Titulo
        Azul2 i_fecha_compra, i_fecha_compra
        GoTo salirf
   End If
End If

'If PUB_TIPMOV = 10 And i_numguia.Text <> "" And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 Then
'   PSGUI_MENOR.rdoParameters(0) = Nulo_Valor0(par_llave!PAR_R_SERIE)
'   PSGUI_MENOR.rdoParameters(1) = i_numguia.Text
'   PSGUI_MENOR.rdoParameters(2) = LK_CODCIA
'   gui_menor.Requery
'   If gui_menor.EOF = False Then
'      MsgBox "Ya existe la guia de remision del Doc: " & gui_menor!far_NUMGUIA, 48, Pub_Titulo
'      If i_numguia.Visible Then Azul i_numguia, i_numguia
'      GoTo salirf
'   End If
'End If
   
If LK_CODTRA = 2412 Then
''  If Val(gridl.TextMatrix(1, 5)) <> Val(i_importe_amort.Text) And Val(grid_fac.TextMatrix(1, 7)) = 0 Then
 '   MsgBox "Verificar Importes .. son diferentes", 48, Pub_Titulo
 '   GoTo salirf
 '' End If
End If
   
   


If Boton_Letras.Visible = True And LK_CODTRA <> 2412 And LK_CODTRA <> 2410 Then
      CONTADOR = 0
      For filafac = 2 To gridl.Rows - 1
           If LK_CODTRA = 2735 And (Val(gridl.TextMatrix(filafac, 5)) > Val(gridl.TextMatrix(filafac, 25))) And Val(gridl.TextMatrix(filafac, 5)) <> 0 Then
            MsgBox "Importe es mayor que el saldo del documento...", 48, Pub_Titulo
            GoTo salirf
          End If
          If Val(gridl.TextMatrix(filafac, 5)) <> 0 Then CONTADOR = 1
          If Val(gridl.TextMatrix(filafac, 19)) <> 0 Then CONTADOR = 1
      Next filafac
     If CONTADOR = 0 Then
        MsgBox "Importe errado...", 48, Pub_Titulo
        pub_mensaje = "No hay amortizaciones ...Desea Continuar.. ?"
        Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
        If Pub_Respuesta = vbNo Then GoTo salirf
     End If
End If

If letche.Visible And LK_CODTRA <> 2401 Then
   If grid_canje.Visible = False Then
        MsgBox "Faltan Datos ...", 48, Pub_Titulo
        GoTo salirf
   End If
   If Val(grid_canje.TextMatrix(1, 3)) = 0 And gridl.Visible = False And grid_fac.Visible = False Then
         MsgBox "Falta Ingresar Documentos ...", 48, Pub_Titulo
        GoTo salirf
   End If
   
End If

If LK_CODTRA = 2728 Then
   If Val(grid_liq.TextMatrix(1, 5)) = 0 Or grid_liq.Visible = False Then
        MsgBox "Faltan Datos ...", 48, Pub_Titulo
        GoTo salirf
   End If
End If
' PROCESO DE TRANF. NO PREGUNTA
If PUB_PROCESO <> 1 And LK_CODTRA = 2403 Then
 If pub_flag_cambio = 1 And i_cambio.Value = 1 Then
   pub_mensaje = "El Documento Anterior sera Reemplazado ... Confirma la operacion ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo)
   If Pub_Respuesta = vbNo Then GoTo salirf
 End If
End If
If Boton_Letras.Visible = True Then
If Val(gridl.TextMatrix(1, 5)) = 0 And (LK_CODTRA = 2412 Or LK_CODTRA = 2410) Then
         pub_mensaje = "No aplicable a ninguna factura...Desea Continuar? "
         'GoTo salirf
        Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo)
       '  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
         If Pub_Respuesta = vbNo Then GoTo salirf
Else
     If Trim(gridl.TextMatrix(2, 0)) = "" Then GoTo salirf
End If
End If
  
If Boton_Letras.Visible = True Or LK_CODTRA = 2770 Or LK_CODTRA = 2774 Then
          CONTADOR = 0
          For filafac = 2 To gridl.Rows - 1
              If Trim(gridl.TextMatrix(filafac, 6)) <> "" Then
                 If IsDate(gridl.TextMatrix(filafac, 6)) = False Then CONTADOR = 2
              End If
              If Trim(gridl.TextMatrix(filafac, 8)) <> "" Then
                 If IsDate(gridl.TextMatrix(filafac, 8)) = False Then CONTADOR = 7
              End If
              
              If Val(gridl.TextMatrix(filafac, 19)) <> 0 Then
                 If IsDate(gridl.TextMatrix(filafac, 22)) = False Then CONTADOR = 2
                 If Trim(gridl.TextMatrix(filafac, 20)) = "" Then
                        MsgBox "Colocar Identificador o numero de Cheque  al Nº.Doc: " & Trim(gridl.TextMatrix(filafac, 1)) & "/" & Trim(gridl.TextMatrix(filafac, 2)) & "-" & Trim(gridl.TextMatrix(filafac, 3)), 48, Pub_Titulo
                        CONTADOR = 6
                 End If
              End If
             
              If Val(gridl.TextMatrix(filafac, 5)) <> 0 And LK_ACTIVA = "A" Then
                 If ww_so <> "X" Then
                    If ww_so <> gridl.TextMatrix(filafac, 27) Then
                       MsgBox "Solo operaciones de un solo tipo ...", 48, Pub_Titulo
                       GoTo salirf
                    End If
                 End If
                 ww_so = gridl.TextMatrix(filafac, 27)
              End If
              
                    
              
              If LK_CODTRA = 2725 And LK_EMP = "PAR" And Not IsDate(gridl.TextMatrix(filafac, 22)) Then CONTADOR = 2

              If LK_CODTRA = 2770 Then
                If LK_EMP = "PIU" And Val(gridl.TextMatrix(filafac, 19)) <> 0 And Val(gridl.TextMatrix(filafac, 5)) <> 0 Then CONTADOR = 5
                If LK_EMP = "PIU" And Val(gridl.TextMatrix(filafac, 19)) <> 0 Then gridl.TextMatrix(filafac, 5) = gridl.TextMatrix(filafac, 19)
                If Trim(gridl.TextMatrix(filafac, 6)) <> "" Then
                 SQ_OPER = 1
                 pu_codcia = LK_CODCIA
                 PUB_CODVEN = gridl.TextMatrix(filafac, 7)
                 LEER_VEN_LLAVE
                 If ven_llave.EOF Then CONTADOR = 9
                End If
              End If
              
          Next filafac
            If CONTADOR = 2 Then
               MsgBox "Verificar Fecha ...", 48, Pub_Titulo
               If gridl.Visible And gridl.Enabled And textovarl.Visible Then textovarl.SetFocus
               GoTo salirf
            End If
            If CONTADOR = 9 Then
               MsgBox "Vendedor errado en fila ...Revisar "
               GoTo salirf
            End If
            If CONTADOR = 7 Then
               MsgBox "Fecha emision invalida ...Revisar "
               GoTo salirf
            End If
            
            If CONTADOR = 5 Then
               MsgBox "Si hay Efectivo y cheques realizar la operacion en dos lineas ... Llamando nuevamente al documento "
               GoTo salirf
            End If
            If CONTADOR = 6 Then
               GoTo salirf
            End If
         If Val(SUT_LLAVE!SUT_SIGNO_CAR) = 2 Then
            If Val(gridl.TextMatrix(1, 5)) <> 0 Then
                  MsgBox "Los importes de los documentos deben ser iguiales para la Liquidación...", 48, Pub_Titulo
                  GoTo salirf
            End If
         End If
End If

If ww_so <> "X" Then
If NJ.Visible = True Then
   If Trim(PUB_SO) <> Trim(ww_so) Then
      MsgBox "No concuerdan los tipos ...", 48, Pub_Titulo
      GoTo salirf
   End If
End If
End If


If LK_CODTRA = 2774 And LK_EMP <> "HER" Then
   If Val(gridl.TextMatrix(1, 5)) = 0 Then
      MsgBox "Importe errado...", 48, Pub_Titulo
      GoTo salirf
   End If
End If


If LK_CODTRA = 1455 Then
      If Val(gridl.TextMatrix(1, 5)) <> Val(grid_canje.TextMatrix(1, 3)) Or Val(gridl.TextMatrix(1, 5)) = 0 Then
         MsgBox "Importe no corresponde ...", 48, Pub_Titulo
         GoTo salirf
      End If
End If

If LK_CODTRA = 2401 Then
   If Val(i_neto.Text) <> Val(grid_canje.TextMatrix(1, 3)) And Val(grid_canje.TextMatrix(1, 3)) > 0 Then
       MsgBox "Importe no corresponde ...", 48, Pub_Titulo
       GoTo salirf
   End If
End If
If fragas.Visible Then
   If Trim(cta_gas1.Text) = "" Then
      If Val(imp_gas1.Text) <> 0 Then
        MsgBox "Definir la Cuenta de Gasto...", 48, Pub_Titulo
        cta_gas1.SetFocus
        GoTo salirf
      End If
   Else
        SQ_OPER = 1
        PUB_CUENTA = Trim(cta_gas1.Text)
        LEER_COM_LLAVE
        If com_llave.EOF Then
          MsgBox "Cuanta Contable No Existe. Verificar ", 48, Pub_Titulo
          Azul cta_gas1, cta_gas1
          GoTo salirf
        Else
          If LK_NIVEL_MAX <> Val(com_llave!COM_NIVEL) Then
            MsgBox "No Procede.. Cuanta no es Analitica...", 48, Pub_Titulo
            Azul cta_gas1, cta_gas1
            GoTo salirf
          End If
        End If
        
   End If
   If Trim(cta_gas2.Text) = "" Then
      If Val(imp_gas2.Text) <> 0 Then
        MsgBox "Definir la Cuenta de Gasto...", 48, Pub_Titulo
        cta_gas2.SetFocus
        GoTo salirf
      End If
   Else
        SQ_OPER = 1
        PUB_CUENTA = Trim(cta_gas2.Text)
        LEER_COM_LLAVE
        If com_llave.EOF Then
          MsgBox "Cuanta Contable No Existe. Verificar ", 48, Pub_Titulo
          Azul cta_gas2, cta_gas2
          GoTo salirf
        Else
          If LK_NIVEL_MAX <> Val(com_llave!COM_NIVEL) Then
            MsgBox "No Procede.. Cuanta no es Analitica...", 48, Pub_Titulo
            Azul cta_gas2, cta_gas2
            GoTo salirf
          End If
        End If
        
   End If
End If

If i_impto2.Visible = True And LK_CODTRA <> 2748 And LK_CODTRA <> 2580 Then
   If Val(i_importe_amort.Text) <> Val(Val(i_bruto2.Text) + Val(i_impto2.Text)) Then
       MsgBox "Importe Errado ...", 48, Pub_Titulo
       GoTo salirf
   End If
End If

If i_ds.Visible = True Then
   If i_ds.Text = "S" Or i_ds.Text = "D" Then
   Else
       MsgBox "Seleccionar Moneda...", 48, Pub_Titulo
       GoTo salirf
   End If
End If
' CHEQUEO DE TRANSACCIONES DE PROVEEDORE
'***************************************
If Val(par_llave!PAR_TASA_VENTA) <> 0 Then
 If LK_CODTRA = 1401 Or (LK_CODTRA = 2741 And PUB_CP = "P") Then
  pub_deuda = Val(i_neto.Text) + DeudaDia("P", i_fecha_vcto.Text)
  If pub_deuda >= Val(par_llave!PAR_TASA_VENTA) Then
    MsgBox "El Monto de Pago para el Credito esta al Tope al : " & Format(i_fecha_vcto, "dd/mm/yyyy") & Chr(13) & "Programarlo para otra fecha.", 48, Pub_Titulo
    'pub_mensaje = "El Monto del Tope de Cronograma de Pagos esta en su Limite. Revisar sus cronogramas...  ¿Desea Continuar... ?"
    'Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    'If Pub_Respuesta = vbNo Then
    Azul i_dias, i_dias
           GoTo salirf
    'End If
  End If
 End If
End If



'***************************************



' CHEQUEO CUALQUIER MOVIMIENTO DE INGRESO DE MERCADERIA
If PUB_TIPMOV <> 0 And pub_signo_arm = 1 Then
   pub_mensaje = ""
   pub_mensaje = CHEQUEO_DETALLE
  If pub_mensaje <> "" Then
     MsgBox "En Ingreso de Mercaderia necesita espesificar sus Lotes de Cada Producto." & Chr(13) & pub_mensaje, 48, Pub_Titulo
     GoTo salirf
  End If
End If


If grid_canje.Visible = True Then
CONTADOR = 0
For filafac = 2 To grid_canje.Rows - 1
    If Val(grid_canje.TextMatrix(filafac, 3)) <> 0 Then
       If IsDate(grid_canje.TextMatrix(filafac, 4)) = False Then CONTADOR = 2
    End If
    If Trim(grid_canje.TextMatrix(filafac, 2)) = "" Then CONTADOR = 3
    
    Next filafac
    If CONTADOR = 2 Then
       MsgBox "Verificar Fecha ...", 48, Pub_Titulo
       GoTo salirf
    End If
    If CONTADOR = 3 Then
       MsgBox "Falta llenar Tipo  de Documento...", 48, Pub_Titulo
       GoTo salirf
    End If

End If


If PUB_TIPMOV = 10 And pub_signo_car = 1 And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 Then
   If Val(i_dias.Text) = 0 And LK_CODTRA <> 2402 Then
      MsgBox "Seleccionar  Dias...", 48, Pub_Titulo
      Azul i_dias, i_dias
      GoTo salirf
   End If
End If


If pub_ojo = "A" Then
   MsgBox "Antes de Grabar ... Favor de Revisar Totales ...", 48, Pub_Titulo
   calcula_totales
   GoTo salirf
End If

If grid_fac.Visible And LK_CODTRA <> 2412 And LK_CODTRA <> 2410 And LK_CODTRA <> 2401 Then
 If PUB_TIPMOV <> 0 Then
    If LK_CODTRA = 1401 And FORMGEN.i_cambio.Value = 1 Then
    Else
    If Val(grid_fac.TextMatrix(1, 4)) = 0 Then
      MsgBox " Falta Seleccionar datos...", 48, Pub_Titulo
      GoTo salirf
    End If
    End If
 End If
End If
                                               
If grid_fac.Visible And LK_CODTRA = 2401 And Nulo_Valors(SUT_LLAVE!SUT_CERO) = "A" And Trim(SUT_LLAVE!SUT_abreviado) <> "CONS" Then
 If Val(grid_fac.TextMatrix(1, 7)) = 0 Then
    'llena_numfac
    If PUB_NUMFAC = 0 Then
       MsgBox "Faltan Datos.."
       GoTo salirf
    End If
    If Val(grid_fac.TextMatrix(2, 4)) <> 0 Then
       pub_mensaje = "Transferencia Gratuita :" & BolFac.Caption & "-" & i_numser.Text & " " & i_numfac.Text & " Continuar ..?"
    Else
       pub_mensaje = "Es  Anulación o en Blanco  del Documento :" & BolFac.Caption & "-" & i_numser.Text & " " & i_numfac.Text & " Continuar ..?"
    End If
    
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo)
    If Pub_Respuesta = vbNo Then
       MsgBox " Falta Seleccionar datos...", 48, Pub_Titulo
       GoTo salirf
    Else
        If Val(grid_fac.TextMatrix(2, 4)) = 0 Then grid_fac.TextMatrix(2, 4) = "ANULACION"
    End If
 End If
End If


If grid_fac.Visible = True And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 Then
   If LK_CODTRA = 2425 Then
   fila = 2
   Do Until fila > 2900
      If (Trim(grid_fac.TextMatrix(fila, 0)) <> "") Then
         SQ_OPER = 1
         PUB_KEY = Trim(grid_fac.TextMatrix(fila, 16))
         LEER_ART_LLAVE
         PUB_KEY = art_LLAVE!ART_CODART2
         LEER_ART_LLAVE
         If art_LLAVE.EOF Then
            MsgBox "Articulo no esta bien relacionado... " & Trim(grid_fac.TextMatrix(fila, 0))
            GoTo salirf
         End If
      End If
      fila = fila + 1
      If fila = grid_fac.Rows Then fila = 2998
   Loop
   FILAXX = grid_fac.Rows - 1
   grid_fac.Rows = grid_fac.Rows + 2
   fila = 2
   Do Until fila = FILAXX
      If (Trim(grid_fac.TextMatrix(fila, 0)) <> "") Then
         SQ_OPER = 1
         PUB_KEY = Trim(grid_fac.TextMatrix(fila, 16))
         LEER_ART_LLAVE
         PUB_KEY = art_LLAVE!ART_CODART2
         LEER_ART_LLAVE
         grid_fac.Rows = grid_fac.Rows + 1
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 16) = PUB_KEY
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 33) = PUB_KEY
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 0) = art_LLAVE!art_nombre
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 3) = FORMGEN.grid_fac.TextMatrix(fila, 3)
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 1) = PUB_KEY
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 4) = FORMGEN.grid_fac.TextMatrix(fila, 4)
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 5) = FORMGEN.grid_fac.TextMatrix(fila, 5)
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 6) = FORMGEN.grid_fac.TextMatrix(fila, 6)
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 7) = FORMGEN.grid_fac.TextMatrix(fila, 7)
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 11) = FORMGEN.grid_fac.TextMatrix(fila, 11)
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 14) = FORMGEN.grid_fac.TextMatrix(fila, 14)
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 15) = FORMGEN.grid_fac.TextMatrix(fila, 15)
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 12) = -1 * Val(FORMGEN.grid_fac.TextMatrix(fila, 12))
         
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 21) = FORMGEN.grid_fac.TextMatrix(fila, 21)
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 23) = FORMGEN.grid_fac.TextMatrix(fila, 23)
         FORMGEN.grid_fac.TextMatrix(FILAXX + fila, 24) = FORMGEN.grid_fac.TextMatrix(fila, 24)
      End If
      fila = fila + 1
   Loop
   
   End If
End If


CONTADOR = 0
If grid_fac.Visible = True And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 Then
   subtotal = 0
   fila = 2
   Do Until fila > 900
     
     If (grid_fac.TextMatrix(fila, 21)) = "A" Then GoTo mas
     
     If Trim(SUT_LLAVE!SUT_abreviado) = "CONS" Then GoTo mas

     If (Val(grid_fac.TextMatrix(fila, 4)) = 0 Or Trim(grid_fac.TextMatrix(fila, 6)) = "") And Val(grid_fac.TextMatrix(fila, 16)) <> 0 Then
            fila = 999
            Exit Do
     End If
     If Val(grid_fac.TextMatrix(fila, 16)) <> Val(grid_fac.TextMatrix(fila, 33)) And Val(grid_fac.TextMatrix(fila, 16)) > 0 Then
            fila = 888
            Exit Do
     End If
     
     If Val(grid_fac.TextMatrix(fila, 6)) < Val(grid_fac.TextMatrix(fila, 32)) And LK_EMP = "3AA" And PUB_TIPMOV = 10 Then
            MsgBox "El precio es Menor que el Minimoo ..." & grid_fac.TextMatrix(fila, 0) & Chr(13) & "Precio Minimo : " & grid_fac.TextMatrix(fila, 32), 48, Pub_Titulo
            fila = 887
            Exit Do
     End If
     If Val(grid_fac.TextMatrix(fila, 4)) = 0 And Val(grid_fac.TextMatrix(fila, 16)) <> 0 Then ' And Val(grid_fac.TextMatrix(fila, 35)) = 0 And LK_EMP = "HER" And PUB_TIPMOV = 10 Then
             MsgBox "Revisar  Item Nro. " & fila - 1 & " . posible Importe Unitario es 0.00 / digitar el producto nuevamente", 48, Pub_Titulo
            fila = 887
            Exit Do
     End If
     
     If LK_CODTRA = 2414 Then
        If Val(grid_fac.TextMatrix(fila, 4)) <> 0 Then CONTADOR = CONTADOR + 1
     End If
     
         

     
     
     If Nulo_Valors(SUT_LLAVE!SUT_CERO) = "A" Then GoTo mas
     
         If (Val(grid_fac.TextMatrix(fila, 4)) = 0 Or Val(grid_fac.TextMatrix(fila, 6)) = 0 Or Val(grid_fac.TextMatrix(fila, 7)) = 0) And Val(grid_fac.TextMatrix(fila, 16)) <> 0 And LK_CODTRA <> 2403 Then
            fila = 999
            Exit Do
         End If
         If (Trim(grid_fac.TextMatrix(fila, 4)) <> "") And Trim(grid_fac.TextMatrix(fila, 0)) = "" Then
            fila = 999
            Exit Do
         End If
         
         
mas:
      fila = fila + 1
      If fila = grid_fac.Rows Then fila = 998
   Loop
   If fila = 999 Then
     If LK_CODUSU = "ADMIN" Or LK_CODUSU = "ADMIN" Then
       pub_mensaje = "Esta seguro de Registrar precio cero  ?"
       Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo)
        If Pub_Respuesta = vbNo Then GoTo salirf
     Else
        MsgBox "Existe algun Producto con Precio 0 , no procede la Venta", 48, Pub_Titulo
        GoTo salirf
     End If
   End If
      
   If fila = 888 Then
      MsgBox " revisar unidades y precios  de algun articulo .."
      GoTo salirf
   End If
   If fila = 887 Then
      GoTo salirf
   End If
   If LK_CODTRA = 2414 And CONTADOR <> 2 Then
      MsgBox "Solo se permite 2 items..."
      GoTo salirf
   End If
   
   
   If fila = 997 Then
      GoTo salirf
   End If
   
If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then
   If Nulo_Valors(par_llave!PAR_VENTAS_IGV) = "A" Then
     ' If Val(i_importe_amort.Text) <> Val(grid_fac.TextMatrix(1, 6)) Or Val(grid_fac.TextMatrix(1, 7)) <> 0 Then
     '   MsgBox "Importe debe ser igual ... ", 48, Pub_Titulo
     '   GoTo salirf
     ' End If
   Else
'      If Val(i_bruto2.Text) <> Val(grid_fac.TextMatrix(1, 7)) Or Val(grid_fac.TextMatrix(1, 7)) <> 0 Then
'        MsgBox "Importe debe ser igual ... ", 48, Pub_Titulo
'        GoTo salirf
'      End If
    End If
    If Val(gridl.TextMatrix(1, 5)) <> 0 Then
     If Val(gridl.TextMatrix(1, 5)) <> Val(i_importe_amort.Text) Then
        MsgBox "Importe debe ser igual ... ", 48, Pub_Titulo
        GoTo salirf
     End If
    End If
    
    If Val(i_importe_amort.Text) = 0 Then
         MsgBox "Falta Importe ...", 48, Pub_Titulo
        GoTo salirf
    End If
    If Val(grid_fac.TextMatrix(1, 7)) = 0 Then
      If Val(i_importe_amort.Text) <> Val(Val(i_impto2.Text) + Val(i_bruto2.Text)) Then
      'If Val(i_importe_amort.text) <> Val(i_impto2.text) + Val(i_bruto2.text) Then
         MsgBox "Revisar Importes  ...", 48, Pub_Titulo
         GoTo salirf
      End If
   End If

Else
End If
End If

SALTARIN:
If PUB_TIPMOV <> 0 Then
    If LK_CODTRA <> 2748 Then
      FORMGEN.i_importe_amort.Text = Val(i_neto.Text)
    End If
End If

If PUB_TIPMOV = 20 And Not cli_llave.EOF Then
   If LK_MONEDA = "D" And cli_llave!CLI_MONEDA = "S" Then
       Tipo_Cambio_Click
   ElseIf LK_MONEDA = "S" And cli_llave!CLI_MONEDA = "D" Then
       Tipo_Cambio_Click
   End If
End If

TOMA_DATOS
'CAPTURA_DATOS

SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE
If par_llave!par_flag_cierre = 9 Then
   MsgBox "!!! Compañia ... Cerró Operaciones ... Llamar al Administrador ", 48, Pub_Titulo
   GoTo salirf
End If
If LK_FLAG_GRIFO <> "A" Then
If par_llave!PAR_FECHA_DIA <> LK_FECHA_DIA Then
   MsgBox "!!!FECHA YA NO COINCIDE CON LA ACTUAL , OTRO USUARIO A CERRADO EL DIA!!! SALGA Y REINICIE SU SISTEMA...", 48, Pub_Titulo
   End
   GoTo salirf
End If
End If



If i_chenum.Visible = False Or pub_signo_ccm = 0 Or LK_CODTRA = 2401 Or LK_CODTRA = 2725 Or LK_CODTRA = 2748 Or LK_CODTRA = 2735 Or LK_CODTRA = 5714 Or LK_CODTRA = 2720 Or LK_CODTRA = 5318 Or LK_CODTRA = 2738 Then
GoTo pasa_C
End If

PUB_CHESEC = 99
SQ_OPER = 1
LEER_CHE_LLAVE          ' CAMBIE
If che_llave.EOF Then
   MsgBox "Numero de Cheque errado ....", 48, Pub_Titulo
   GoTo salirf
End If



che_llave.MoveLast
If che_llave!CHE_FECHA > 0 / 1 / 1900 Then
   pub_mensaje = "Cheque ya procesado...desea operar con el mismo numero de cheque ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo)
   If Pub_Respuesta = vbNo Then
   GoTo salirf
   End If
End If

pasa_C:


'If Tipo_Cambio.Visible = True Then
'   If WS_TIPO_CAMBIO <> LK_TIPO_CAMBIO And WS_TIPO_CAMBIO = 0 Then Tipo_Cambio_Click
'End If

If PUB_TIPMOV <> 0 And PUB_TIPMOV = 20 Then   'QUITAR EL TIPMOV..PARA HER
   If Val(i_flete.Text) <> SUB_FLETE Then
      SUB_FLETE = 0
      MsgBox " Los Fletes no coinciden ....Revisar " & SUB_FLETE & " - " & i_flete.Text
      GoTo salirf
    End If
    If Val(i_descto.Text) <> PUB_DESCTO Then
       MsgBox " Los Desctos no coinciden ....Revisar ", 48, Pub_Titulo
       GoTo salirf
    End If
End If
' ICA

If LK_CODTRA = 1111 And LK_EMP <> "3AA" Then
  If ((CDate(FORMGEN.Grid_all.TextMatrix(Grid_all.Row, 11)) >= cop_llave!cop_fecha_proceso) Or (CDate(FORMGEN.Grid_all.TextMatrix(Grid_all.Row, 11)) <= cop_llave!cop_fecha_proceso2)) Then
  Else
    MsgBox "El Documento es del : " & CDate(FORMGEN.Grid_all.TextMatrix(Grid_all.Row, 11)) & " periodo cerrado... No Procede la anulación ", 48, Pub_Titulo
    GoTo salirf
  End If
End If

If LK_ACTIVA = "A" And LK_CODTRA <> 1111 And Left(Trim(Str(LK_CODTRA)), 1) <> "5" And LK_CODTRA <> 2745 And LK_CODTRA <> 2725 And LK_CODTRA <> 2580 And LK_CODTRA <> 2007 And LK_CODTRA <> 2741 Then
If PUB_SO = "X" And ww_so = "X" Then  ' ojo agregar un flag para activar ese control
    MsgBox "Falta Tipo de Operacion ..."
    GoTo salirf
End If
End If
If Val(i_codcli.Text) = 0 Then
 PUB_RUC = ""
End If


consis = True

Exit Function
salirf:
consis = False


End Function

Public Function consis2() As Boolean
Dim wau As Currency
wau = 0
If LK_EMP = "CAM" Then
    wau = 1
End If

If grid_fac.Visible = True And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 And ww_respuesta = vbNo And pub_signo_arm = -1 Then
   fila = 2
   Do Until fila > 900
   
     PUB_CANTIDAD = Val(grid_fac.TextMatrix(fila, 4))
     PUB_CODART = Val(grid_fac.TextMatrix(fila, 16))
     'If Trim(LK_USU_STOCK) = "" And PUB_CODART <> 0 And PUB_CANTIDAD <> 0 Then
     If Trim(LK_USU_STOCK) = "" And PUB_CODART <> 0 And PUB_CANTIDAD <> 0 And PUB_PROCESO = 0 Then
        pu_codcia = LK_CODCIA
        SQ_OPER = 1
        LEER_ARM_LLAVE
        PUB_CANTIDAD = Val(grid_fac.TextMatrix(fila, 4)) * Val(grid_fac.TextMatrix(fila, 14))
        If PUB_CANTIDAD > (Val(arm_llave!ARM_STOCK) + wau) And LK_ACTIVA <> "A" Then
           pub_mensaje_err = "Ojo Cantidad Mayor que Stock ..." & "Existencia : " & grid_fac.TextMatrix(fila, 0) & ":" & arm_llave!ARM_STOCK & " ......Salida Requerida :" & PUB_CANTIDAD
           loc_stock = Trim(grid_fac.TextMatrix(fila, 0)) & " Stock : " & Format((Val(arm_llave!ARM_STOCK) + wau), "0.00") & " - Por Pedido : " & Format(PUB_CANTIDAD, "0.00")
           GoTo salirf
        End If
        If LK_ACTIVA = "A" Then
           If PUB_SO = "A" Then
              If PUB_CANTIDAD > (Val(arm_llave!ARM_saldo_s) + wau) <> "A" Then
                 pub_mensaje_err = "Ojo Cantidad Mayor que Stock ...Tipo de Operacion no procede ..." & "Existencia : " & grid_fac.TextMatrix(fila, 0) & ":" & arm_llave!ARM_saldo_s & " ......Salida Requerida :" & PUB_CANTIDAD
                 GoTo salirf
              End If
           Else
              If PUB_CANTIDAD > (Val(arm_llave!ARM_Saldo_n) + wau) Then
                 pub_mensaje_err = "Ojo Cantidad Mayor que Stock ...Tipo de Operacion no procede ..." & "Existencia : " & grid_fac.TextMatrix(fila, 0) & ":" & arm_llave!ARM_Saldo_n & " ......Salida Requerida :" & PUB_CANTIDAD
                 GoTo salirf
              End If
           End If
        End If
          
      End If
      
      fila = fila + 1
      If fila = grid_fac.Rows Then fila = 998
   Loop
End If

consis2 = True

Exit Function
salirf:

consis2 = False

End Function

Public Sub REPROCESO()
'*************************
' CARAGA CUENTAS CORRIENTES
'*************************
Dim WS_FILA As Integer
Dim xl  As Object
Stop
Stop

If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Workbooks.Open "C:\CARGA\CTACTE2.xls"  ', 0, True, 4,  WPAS, WPAS
xl.Application.Visible = True
WS_FILA = 2
Do Until WS_FILA = 965000
 WS_FILA = WS_FILA + 1
 If Val((xl.Cells(WS_FILA, 6))) = 0 Then Exit Do
 SQ_OPER = 1
 pu_codclie = Val((xl.Cells(WS_FILA, 6)))
 pu_codcia = LK_CODCIA
 pu_cp = "C"
 LEER_CLI_LLAVE
 If cli_llave.EOF Then
   xl.Cells(WS_FILA, 12) = "NO EXISTE"
'   Stop
 End If
 SQ_OPER = 1
 pu_codcia = LK_CODCIA
 PUB_CODVEN = Val(xl.Cells(WS_FILA, 7))
 LEER_VEN_LLAVE
 If ven_llave.EOF Then Stop
 If Trim(xl.Cells(WS_FILA, 1)) = "" Then Exit Do
 
 If Left(Trim(xl.Cells(WS_FILA, 1)), 1) = "F" Then
   i_fbg.ListIndex = 0
 Else
   i_fbg.ListIndex = 1
 End If
 i_codcli.Text = Val((xl.Cells(WS_FILA, 6)))
 i_fecha_compra.Text = Format(Trim(xl.Cells(WS_FILA, 4)), "dd/mm/yy")
 i_fecha_vcto.Text = Format(Trim(xl.Cells(WS_FILA, 5)), "dd/mm/yyyy")
 If i_fecha_compra.Text > i_fecha_vcto.Text Then
   i_fecha_vcto.Text = Format(i_fecha_compra.Text, "dd/mm/yyyy")
 End If
 
  
 i_numser_c.Text = Left(Trim(xl.Cells(WS_FILA, 2)), 3)
 i_numfac_c.Text = Right(Trim(xl.Cells(WS_FILA, 2)), 7)
 i_codven.Text = Val(xl.Cells(WS_FILA, 7))
 i_concepto.Text = "Carga Incial de Saldos"
 LOC_IMP_INI = Val(xl.Cells(WS_FILA, 8))
 LOC_DEFINI = Val(xl.Cells(WS_FILA, 10))
 'i_importe.Text = Val(xl.Cells(WS_FILA, 9))
 i_importe_amort.Text = Val(xl.Cells(WS_FILA, 9))
 grabar_Click
 i_DEF_KeyPress 13
 Stop
Loop

MsgBox "TERMINO"



Exit Sub

'*************************
' PROCESO NO LOSE
Exit Sub
Dim WNUM As Integer
Dim FAR_TRANSX As rdoResultset
Dim PSFAR_TRANSX  As rdoQuery
Dim W_CUENTA As Integer
Dim ww_numser
Dim ww_numdoc
Dim ingre
Dim ws_flag_car
Dim WS_CANTIDAD As Currency
Dim WW_FECHA_COBRO As Date
Dim LOC_SALDO_CAR As Currency
Dim wS_saldo_car As Currency
Dim WS_TOT As Currency
Dim WS_FLAG_FACART
Dim ws_fecha_anterior As Date
Dim WS_IMPORTE_AMORT As Currency
Dim WWFECHA As Date

Dim LIST_ART As rdoQuery
Dim LISTA_ARTI As rdoResultset
Dim WEQUIV_REAL As Currency
Dim CANTI_RES As Currency

Exit Sub

pub_cadena = "SELECT FAR_SUBTRA, FAR_FBG, FAR_NUMFAC, FAR_NUMSER , FAR_FECHA, FAR_NUMDOC  FROM FACART  WHERE  FAR_CODCIA = ? AND FAR_FECHA >= ?  AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E' "
Set LIST_ART = CN.CreateQuery("", pub_cadena)
Set LISTA_ARTI = LIST_ART.OpenResultset(rdOpenKeyset, rdConcurValues)

MsgBox "LISTO"
WWFECHA = CDate("01/01/01")
LIST_ART(0) = LK_CODCIA
LIST_ART(1) = WWFECHA
LISTA_ARTI.Requery
If LISTA_ARTI.EOF Then
  MsgBox "NO HAY "
  Stop
  Exit Sub
End If
Print LISTA_ARTI.RowCount

pub_cadena = "SELECT CAR_NUM_REN FROM CARTERA  WHERE  CAR_CODCIA = ? AND CAR_FECHA_INGR = ? AND CAR_NUMSER = ? AND CAR_NUMFAC = ?  AND CAR_TIPMOV = 10 "
Set PSFAR_TRANSX = CN.CreateQuery("", pub_cadena)
Set FAR_TRANSX = PSFAR_TRANSX.OpenResultset(rdOpenKeyset, rdConcurValues)


Do Until LISTA_ARTI.EOF
'FAR_FECHA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? AND FAR_NUMDOC = ? "
    PSFAR_TRANSX(0) = LK_CODCIA
    PSFAR_TRANSX(1) = LISTA_ARTI!FAR_fecha
    PSFAR_TRANSX(2) = LISTA_ARTI!far_numser
    PSFAR_TRANSX(3) = LISTA_ARTI!far_numfac
'    PSFAR_TRANSX(3) = LISTA_ARTI!far_NUMDOC
    FAR_TRANSX.Requery
    Print FAR_TRANSX.RowCount
'    If FAR_TRANSX.EOF Then Stop
     WNUM = 0
    If Trim(LISTA_ARTI!far_subtra) = "Ventas al Credito-Mi" Then
     WNUM = 24
    ElseIf Trim(LISTA_ARTI!far_subtra) = "Ventas al Credito-Ma" Then
     WNUM = 26
    ElseIf Trim(LISTA_ARTI!far_subtra) = "Ventas al Contado-Mi" Then
      WNUM = 27
    ElseIf Trim(LISTA_ARTI!far_subtra) = "Ventas al Contado-Ma" Then
      WNUM = 28
    Else
   ''    Beep
    '  Stop
    End If
    DoEvents
     Print FAR_TRANSX.RowCount
     Print LISTA_ARTI.AbsolutePosition
     Print LISTA_ARTI.RowCount
'     nomtra.Visible = True
     nomtra.Caption = LISTA_ARTI.AbsolutePosition & " / " & LISTA_ARTI.RowCount
     If WNUM = 0 Then GoTo SALT
     Do Until FAR_TRANSX.EOF
        If FAR_TRANSX!CAR_NUM_REN <> WNUM Then
         FAR_TRANSX.Edit
         FAR_TRANSX!CAR_NUM_REN = WNUM
         FAR_TRANSX.Update
        End If
       FAR_TRANSX.MoveNext
     Loop
SALT:
  LISTA_ARTI.MoveNext
'  LISTA_ARTI.Requery
Loop

MsgBox "PROCESO TERMINADO"

Exit Sub



pub_cadena = "SELECT FAR_CODART FROM FACART  WHERE  FAR_CODCIA = ? AND FAR_FECHA >= ?  AND FAR_TIPMOV = 10 GROUP BY FAR_CODART "
Set LIST_ART = CN.CreateQuery("", pub_cadena)
Set LISTA_ARTI = LIST_ART.OpenResultset(rdOpenKeyset, rdConcurValues)

Stop
WWFECHA = CDate("06/08/00")
LIST_ART(0) = LK_CODCIA
LIST_ART(1) = WWFECHA
LISTA_ARTI.Requery
If LISTA_ARTI.EOF Then
  MsgBox "NO HAY "
  Stop
  Exit Sub
End If
Print LISTA_ARTI.RowCount

pub_cadena = "SELECT FAR_CODART,FAR_CANTIDAD, FAR_EQUIV, FAR_DESCRI, FAR_SIGNO_ARM FROM FACART  WHERE  FAR_CODCIA = ? AND FAR_FECHA >= ?  AND FAR_CODART = ? AND  FAR_TIPMOV = 10 ORDER BY FAR_FECHA, FAR_NUMOPER "
Set PSFAR_TRANSX = CN.CreateQuery("", pub_cadena)
Set FAR_TRANSX = PSFAR_TRANSX.OpenResultset(rdOpenKeyset, rdConcurValues)


Do Until LISTA_ARTI.EOF
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = LISTA_ARTI!far_codart
    LEER_PRE_LLAVE
    If pre_mayor.RowCount <> 2 Then
'      GoTo dale
    End If
    PSFAR_TRANSX(0) = LK_CODCIA
    PSFAR_TRANSX(1) = WWFECHA
    PSFAR_TRANSX(2) = LISTA_ARTI!far_codart
    FAR_TRANSX.Requery
    WEQUIV_REAL = 0
    Print FAR_TRANSX.RowCount
    Do Until FAR_TRANSX.EOF
        pre_mayor.MoveFirst
        Do Until pre_mayor.EOF
            If Trim(UCase(pre_mayor!pre_UNIDAD)) = Trim(UCase(FAR_TRANSX!far_descri)) Then
                WEQUIV_REAL = pre_mayor!PRE_EQUIV
'                GoTo SIG01
            End If
            pre_mayor.MoveNext
        Loop
        MsgBox "NO ENCONTRO "
        Stop
SIG013:
     If Val(WEQUIV_REAL) = Val(FAR_TRANSX!FAR_equiv) Then
          GoTo OTRO_1
     End If
     CANTI_RES = CANTI_RES + (redondea(FAR_TRANSX!far_cantidad * WEQUIV_REAL) - Val(FAR_TRANSX!far_cantidad)) * FAR_TRANSX!far_signo_arm
     FAR_TRANSX.Edit
     FAR_TRANSX!far_cantidad = redondea(FAR_TRANSX!far_cantidad * WEQUIV_REAL)
     FAR_TRANSX!FAR_equiv = WEQUIV_REAL
     FAR_TRANSX.Update

OTRO_1:
     FAR_TRANSX.MoveNext
    Loop
    
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = LISTA_ARTI!far_codart
    LEER_ARM_LLAVE
    arm_llave.Edit
    arm_llave!ARM_STOCK = arm_llave!ARM_STOCK + CANTI_RES
    arm_llave.Update
    
dale8:
  LISTA_ARTI.MoveNext
Loop

MsgBox "PROCESO TERMINADO"

Exit Sub

MsgBox "YA LO USE EN ICA"
pub_cadena = "SELECT * FROM CARTERA  WHERE  CAR_TIPDOC = 'CC' AND CAR_FECHA_INGR >= '2000/09/01'  AND CAR_CODCIA = '02' ORDER BY CAR_FECHA_INGR"
Set PSALL_LLAVE = CN.CreateQuery("", pub_cadena)
Set all_llave = PSALL_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
all_llave.Requery
If all_llave.EOF Then
  MsgBox "NO HAY"
  Exit Sub
End If
Print all_llave.RowCount
pub_cadena = "SELECT CAA_CODVEN , CAA_ESTADO  FROM CARACU WHERE CAA_CODCIA = '02' AND CAA_TIPDOC = ? AND CAA_SERDOC = ? AND  CAA_NUMDOC = ? AND CAA_CODCLIE = ? "
Set PSFAR_TRANSX = CN.CreateQuery("", pub_cadena)
Set FAR_TRANSX = PSFAR_TRANSX.OpenResultset(rdOpenKeyset, rdConcurValues)


Do Until all_llave.EOF
Print all_llave!CAR_FECHA_INGR
Print all_llave.RowCount
Print all_llave.AbsolutePosition
PSFAR_TRANSX(0) = all_llave!car_TIPDOC
PSFAR_TRANSX(1) = all_llave!car_SERDOC
PSFAR_TRANSX(2) = all_llave!car_NUMDOC
PSFAR_TRANSX(3) = all_llave!CAR_CODCLIE
FAR_TRANSX.Requery
If FAR_TRANSX.EOF Then Stop
Do Until FAR_TRANSX.EOF
 If FAR_TRANSX!CAA_ESTADO = "E" Then GoTo PA
 If Val(FAR_TRANSX!CAA_CODVEN) <> Val(all_llave!CAR_codven) Then
'  If FAR_TRANSX!CAA_CODVEN <> 1 Then Stop
  FAR_TRANSX.Edit
  FAR_TRANSX!CAA_CODVEN = all_llave!CAR_codven
  FAR_TRANSX.Update
 End If
PA:
 FAR_TRANSX.MoveNext
Loop
all_llave.MoveNext
Loop
MsgBox "TERMINO"

Exit Sub
pub_cadena = "SELECT * FROM ALLOG WHERE  ALL_CODCIA = ? AND ALL_FLAG_EXT = 'E' AND (ALL_CODTRA = 2401 OR ALL_CODTRA = 1401)  ORDER BY ALL_FECHA_DIA, ALL_NUMOPER "
pub_cadena = "SELECT * FROM ALLOG WHERE  ALL_CODCIA = ? AND ALL_FLAG_EXT = 'E' AND ALL_CODTRA = 2412 ORDER BY ALL_FECHA_DIA, ALL_NUMOPER "
Set PSALL_LLAVE = CN.CreateQuery("", pub_cadena)
PSALL_LLAVE(0) = ""
Set all_llave = PSALL_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
PSALL_LLAVE(0) = LK_CODCIA
all_llave.Requery
If all_llave.EOF Then
  MsgBox "NO HAY"
  Exit Sub
End If
PUB_NUM_OPER_XXX = 2000
MsgBox all_llave.RowCount
W_CUENTA = 0
Do Until all_llave.EOF
   Print all_llave.AbsolutePosition
   GoSub REPRO
OTRITO:
   all_llave.MoveNext
Loop
MsgBox "TERMINO .. " & W_CUENTA
' FIN
Exit Sub

REPRO:
'EMPIEZA
PUB_NUM_OPER = all_llave!ALL_NUMOPER
PUB_NUM_OPER_EXT = all_llave!ALL_NUMOPER
PUB_CODCIA = all_llave!all_CODCIA
PUB_CODCIAL = all_llave!all_CODCIA

pu_codclie = all_llave!ALL_CODCLIE
pu_cp = all_llave!ALL_CP
pu_codcia = LK_CODCIA
SQ_OPER = 1
LEER_CLI_LLAVE

ww_codtra_ext = all_llave!ALL_codtra

PUB_NUMDOC = Nulo_Valor0(all_llave!ALL_NUMDOC)
PUB_CODART = all_llave!ALL_CODART
PUB_IMPORTE_AMORT = all_llave!ALL_IMPORTE_amort
WS_IMPORTE_AMORT = all_llave!ALL_IMPORTE_amort

If all_llave!ALL_tipmov = 20 Then
   If Val(all_llave!ALL_IMPORTE_DOLL) <> 0 Then WS_IMPORTE_AMORT = all_llave!ALL_IMPORTE_DOLL
End If

PUB_IMPORTE = all_llave!ALL_IMPORTE
PUB_CP = all_llave!ALL_CP
PUB_CODCIA = all_llave!ALL_CP
PUB_CODCLIE = all_llave!ALL_CODCLIE
PUB_TIPDOC = all_llave!ALL_TIPDOC
PUB_TIPMOV = all_llave!ALL_tipmov
PUB_CODBAN = all_llave!all_codban
PUB_CHENUM = all_llave!all_chenum
PUB_CHESER = Nulo_Valor0(all_llave!ALL_CHESER)
PUB_CHESEC = Nulo_Valor0(all_llave!ALL_CHESEC)
PUB_SERDOC = Nulo_Valor0(all_llave!ALL_serdoc)
PUB_NUMDOC = Nulo_Valor0(all_llave!ALL_NUMDOC)
ws_fecha_anterior = Nulo_Valor0(all_llave!all_fecha_ant)
PUB_CANTIDAD = all_llave!ALL_CANTIDAD
PUB_NUMSER = all_llave!ALL_NUMSER
PUB_FBG = all_llave!ALL_FBG
PUB_CODVEN = all_llave!ALL_CODVEN
PUB_NUMFAC = all_llave!all_numfac
PUB_SECUENCIA = all_llave!all_SECUENCIA
pub_signo_ccm = Nulo_Valor0(all_llave!ALL_SIGNO_CCM) * -1
pub_signo_car = Nulo_Valor0(all_llave!ALL_SIGNO_CAR) * -1
pub_signo_arm = Nulo_Valor0(all_llave!all_sIGNO_ARM) * -1
pub_signo_caja = Nulo_Valor0(all_llave!ALL_signo_caja) * -1
WS_FLAG_FACART = Nulo_Valors(all_llave!ALL_FACART)
PUB_CONCEPTO = "Extorno - " & all_llave!all_concepto
PUB_FECHA = all_llave!ALL_FECHA_DIA

If LK_CODTRA = 1122 Then PUB_CONCEPTO = "Anulacion - " & all_llave!all_concepto
'If (Val(all_llave!ALL_IMPORTE_AMORT) + Val(all_llave!ALL_IMPORTE)) = 0 And LK_CODTRA = 1111 Then
' pub_signo_car = 0
'End If

If pub_signo_car = 0 Then
     GoTo OTRITO
End If
If pub_signo_car <> 0 Then
SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_codclie = PUB_CODCLIE
pu_cp = PUB_CP
LEER_CAR_LLAVE
If car_llave.EOF = True Then
   'pub_mensaje_err = "DOCUMENTO NO EXISTE..."
   GoTo OTRITO
End If
Print car_llave!car_FBG & car_llave!car_NUMFAC
If car_llave!car_importe <= 0 Then
  GoTo OTRITO
End If
SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_codclie = PUB_CODCLIE
pu_cp = PUB_CP
LEER_CAA_LLAVE
If caa_histo.EOF = True Then
   'pub_mensaje_err = "DOCUMENTO EN CARACU NO EXISTE..."
   Stop
   GoTo OTRITO
End If
End If
W_CUENTA = W_CUENTA + 1
If pub_signo_car <> 0 Then
   caa_histo.Edit
   caa_histo!CAA_ESTADO = "E"
   ww_numdoc = caa_histo!caa_numfac_c
   ww_numser = caa_histo!caa_numser_c
   caa_histo.Update
End If


car_llave.Edit

WS_TOT = PUB_IMPORTE_AMORT
WS_IMPORTE_AMORT = PUB_IMPORTE_AMORT

If gridl.Visible = True Then
   WS_TOT = Val(gridl.TextMatrix(1, 9))
End If

car_llave!car_importe = car_llave!car_importe + WS_IMPORTE_AMORT * pub_signo_car
wS_saldo_car = car_llave!car_importe

LOC_SALDO_CAR = car_llave!car_importe

If LK_CODTRA = 1111 Then
   car_llave!car_fecha_vcto = ws_fecha_anterior
Else
   ws_fecha_anterior = car_llave!car_fecha_vcto
End If


If (FORMGEN.i_fecha_vcto.Visible = True And LK_CODTRA <> 1111) Or gridl.Visible = True Or grid_liq.Visible = True Then
   car_llave!car_fecha_vcto = PUB_FECHA_VCTO
End If



If FORMGEN.i_concepto.Visible = True And LK_CODTRA <> 2412 And LK_CODTRA <> 2410 Then
   car_llave!car_concepto = PUB_CONCEPTO
End If

If PUB_IMPORTE_AMORT <> 0 And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 And LK_CODTRA <> 1133 And car_llave!car_importe <> 0 Then
   car_llave!CAR_NUM_REN = car_llave!CAR_NUM_REN + 1
End If

If LK_EMP = "HER" Then car_llave!CAR_COBRADOR = Val(gridl.TextMatrix(FILAX, 7))
If LK_CODTRA = 2725 And LK_EMP = "PAR" Then
  WW_FECHA_COBRO = gridl.TextMatrix(FILAX, 22)
Else
  WW_FECHA_COBRO = LK_FECHA_DIA
End If
If LK_CODTRA = 1111 Or LK_CODTRA = 1122 Then  ' para mantener enlace del doc.
 PUB_NUMSER_C = car_llave!car_numser_c
 PUB_NUMFAC_C = car_llave!car_numfac_c
 If ww_codtra_ext = 2105 Then
    car_llave!car_concepto = ""
    car_llave!car_numser_c = 0
    car_llave!car_numfac_c = 0
 End If
 If ww_codtra_ext = 2101 Then
    car_llave!car_numfac_c = -1  ' es para que no aparezca como canje
 End If
 End If
 
If LK_CODTRA = 2410 Or LK_CODTRA = 2412 Then
  PUB_NUMSER_C = Val(FORMGEN.i_numser_c.Text)
  PUB_NUMFAC_C = Val(FORMGEN.i_numfac_c.Text)
End If

car_llave.Update
'GoSub manda_numero
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
GoSub registra_caa
'ws_flag_car = modif
Return

Stop
Exit Sub


registra_caa:
If pub_signo_car <> 0 Then
cli_llave.Edit
WS_CANTIDAD = WS_IMPORTE_AMORT * pub_signo_car
cli_llave!cli_SALDO = Nulo_Valor0(cli_llave!cli_SALDO) + WS_CANTIDAD
If pub_signo_car = -1 Then cli_llave!cli_fecha_fac = PUB_FECHA_VCTO
' AGREGE IF PARA NO LIMPIAR EL CLI_MONEDA POR QUE WS_MONEDA_CLI = ""
If Trim(WS_MONEDA_CLI) = "S" Or Trim(WS_MONEDA_CLI) = "D" Then cli_llave!CLI_MONEDA = WS_MONEDA_CLI
cli_llave.Update
End If

caa_histo.AddNew
caa_histo!CAA_CODCLIE = PUB_CODCLIE
caa_histo!caa_codcia = LK_CODCIA
caa_histo!CAA_TIPDOC = PUB_TIPDOC
caa_histo!CAA_CP = PUB_CP
caa_histo!CAA_NUM_OPER = PUB_NUM_OPER_XXX
If grid_liq.Visible = True Then
   caa_histo!caa_INTVEN = grid_liq.TextMatrix(1, 2)
   caa_histo!caa_DIASV = grid_liq.TextMatrix(1, 6)
   caa_histo!caa_DIASA = grid_liq.TextMatrix(1, 7)
   caa_histo!caa_tasav = Val(i_tasav.Text)
Else
   caa_histo!caa_INTVEN = 0
   caa_histo!caa_DIASV = 0
   caa_histo!caa_DIASA = 0
   caa_histo!caa_tasav = 0
End If

caa_histo!caa_serdoc = PUB_SERDOC
caa_histo!CAA_NUMDOC = PUB_NUMDOC
caa_histo!CAA_FECHA = LK_FECHA_DIA
caa_histo!CAA_FECHA_VCTO = PUB_FECHA_VCTO
caa_histo!caa_situacion = PUB_SITUACION_ACT
caa_histo!caa_concepto = PUB_CONCEPTO
caa_histo!CAA_IMPORTE = WS_IMPORTE_AMORT * pub_signo_car
caa_histo!CAA_TOTAL = Abs(WS_TOT) * pub_signo_car
caa_histo!CAA_SALDO = Nulo_Valor0(cli_llave!cli_SALDO)
caa_histo!caa_SALDO_car = wS_saldo_car
If pub_signo_car <> 0 Then
   caa_histo!CAA_SIGNO_CAJA = PUB_SECUENCIA
Else
   caa_histo!CAA_SIGNO_CAJA = 0
End If
caa_histo!CAA_SIGNO_CAJA_REAL = pub_signo_caja
caa_histo!CAA_SIGNO_CAR = pub_signo_car
caa_histo!CAA_TIPMOV = PUB_TIPMOV
caa_histo!CAA_hora = Now
caa_histo!CAA_CODUSU = LK_CODUSU
caa_histo!CAA_ESTADO = "N"
If Not cli_llave.EOF And cli_llave.RowCount > 0 Then
   If cli_llave!CLI_CODCLIE = Val(i_codcli.Text) Then
      caa_histo!CAa_NOMBRE = cli_llave!cli_nombre
   End If
End If

caa_histo!CAA_NUMPLAN = pub_numplan
If LK_CODTRA = 1111 Then caa_histo!CAA_ESTADO = "E"
If LK_CODTRA = 1122 Then caa_histo!CAA_ESTADO = "E"
If LK_CODTRA = 1133 Then caa_histo!CAA_ESTADO = "E"

If i_fecha_compra.Visible Then
  caa_histo!CAA_FECHA_COBRO = CDate(i_fecha_compra.Text) 'WW_FECHA_COBRO
Else
 caa_histo!CAA_FECHA_COBRO = LK_FECHA_DIA
End If
If ws_flag_car = ingre Then
   caa_histo!CAa_NUM_CHEQUE = PUB_NUM_CHEQUE
   caa_histo!CAa_numser = PUB_NUMSER
   caa_histo!CAa_numfac = PUB_NUMFAC
   caa_histo!caa_numser_c = PUB_NUMSER_C
   caa_histo!caa_numfac_c = PUB_NUMFAC_C
   If LK_CODTRA = 2741 Then
    caa_histo!CAa_numser = PUB_NUMSER_C
    caa_histo!CAa_numfac = PUB_NUMFAC_C
   End If
   caa_histo!CAa_numGUIA = PUB_NUMGUIA
   caa_histo!CAa_SERGUIA = Val(i_serguia.Text)
   caa_histo!CAA_FBG = PUB_FBG
   caa_histo!CAA_CODVEN = PUB_CODVEN
   caa_histo!caa_situacion = " "
Else
   caa_histo!CAa_NUM_CHEQUE = Nulo_Valors(car_llave!car_NUM_CHEQUE)
   caa_histo!CAa_numser = Nulo_Valor0(car_llave!car_NUMSER)
   caa_histo!CAa_numfac = Nulo_Valor0(car_llave!car_NUMFAC)
   If LK_CODTRA = 2412 Then
     caa_histo!caa_numser_c = PUB_NUMSER
     caa_histo!caa_numfac_c = PUB_NUMFAC
   Else
     caa_histo!caa_numser_c = PUB_NUMSER_C
     caa_histo!caa_numfac_c = PUB_NUMFAC_C
   End If

   caa_histo!CAA_NOTA = PUB_FBG
   caa_histo!CAA_FBG = Nulo_Valors(car_llave!car_FBG)
   If LK_CODTRA = 2770 And i_codven.Visible = True Then
      caa_histo!CAA_CODVEN = PUB_CODVEN
   Else
      caa_histo!CAA_CODVEN = Nulo_Valor0(car_llave!CAR_codven)
   End If
   caa_histo!CAa_numGUIA = Nulo_Valor0(car_llave!car_NUMGUIA)
   caa_histo!CAa_SERGUIA = Nulo_Valor0(car_llave!car_SERguia)
   caa_histo!caa_situacion = Nulo_Valors(car_llave!CAR_SITUACION)
   caa_histo!caa_signo_ccm = pub_signo_ccm
   caa_histo!caa_codban = PUB_CODBAN
End If
caa_histo!CAA_RECIBO = PUB_CODART
If LK_CODTRA = 2728 And LK_EMP = "PLA" Then
  caa_histo!caa_situacion = PLAZA_FLAG_MANUAL
End If

caa_histo.Update
Return


manda_numero:
SQ_OPER = 2
PUB_FECHA = LK_FECHA_DIA
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
If all_menor.EOF = False Then
   PUB_NUM_OPER_XXX = all_menor!ALL_NUMOPER
Else
   PUB_NUM_OPER_XXX = 0
End If
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
'WS_NUM_OPER2 = PUB_NUM_OPER_XXX
Return


End Sub
Public Sub lee_ped_llave()
  PSped_llave.rdoParameters(0) = 177
  PSped_llave.rdoParameters(1) = LK_CODCIA
  PSped_llave.rdoParameters(2) = far_menor3!FAR_PEDSER
  PSped_llave.rdoParameters(3) = far_menor3!FAR_PEDFAC
  ped_llave.Requery
End Sub
Public Sub elim_ped()
  Do Until ped_llave.EOF
     ped_llave.Edit
     ped_llave!PED_ROLLOS = 0
     ped_llave!PED_cantidad = 0
     ped_llave!PED_CONTACTO = "Extornado..."
     ped_llave.Update
     ped_llave.MoveNext
  Loop
End Sub
Public Sub ped_2210()
  'ped_llave.MoveFirst
  'Do Until ped_llave.EOF
  '   If ped_llave!PED_NUMSEC = far_menor3!far_pedsec Then
  '      ped_llave.Edit
  '      ped_llave!PED_ingresos = Val(ped_llave!PED_ingresos) - Val(far_menor3!far_cantidad)
  '      ped_llave!PED_UNID_I = Val(ped_llave!PED_UNID_I) - Val(far_menor3!far_UNIDADES)
  '      ped_llave.Update
  '   End If
  '   ped_llave.MoveNext
  'Loop

End Sub



Public Sub ped_2212()
  'ped_llave.MoveFirst
  'Do Until ped_llave.EOF
  '   If ped_llave!PED_NUMSEC = far_menor3!far_pedsec Then
  '      ped_llave.Edit
  '      ped_llave!PED_salidas = Val(ped_llave!PED_salidas) - Val(far_menor3!far_cantidad)
  '      ped_llave!PED_UNID_S = Val(ped_llave!PED_UNID_S) - Val(far_menor3!far_UNIDADES)'''
'
'        ped_llave.Update
'     End If
'     ped_llave.MoveNext
'  Loop

End Sub
Public Function ANEXO_CON20() As Boolean
Dim CONTADOR As Integer
ANEXO_CON20 = True
CONTADOR = 0
Do Until CONTADOR = PUB_CANT_CHEQ
PUB_CHESEC = 0
PUB_CHENUM = CONTADOR + PUB_NUM_INI
CONTADOR = CONTADOR + 1
SQ_OPER = 1
LEER_CHE_LLAVE
If Not che_llave.EOF Then
       pub_mensaje_err = "Numero de Serie y Cheque ya existe..." & PUB_CHESER & "-" & PUB_CHENUM
       ANEXO_CON20 = False
End If
Loop

End Function
Public Function ANEXO_CON21() As Boolean
ANEXO_CON21 = True

If IsDate(i_fecha_oper.Text) = False Then
       ANEXO_CON21 = False
       Exit Function
End If
PUB_FECHA = i_fecha_oper.Text

SQ_OPER = 1
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
Do Until all_llave.EOF
  If PUB_NUM_OPER = all_llave!ALL_NUMOPER Then
    Exit Do
  Else
    all_llave.MoveNext
  End If
Loop
     

If all_llave.EOF Then
   pub_mensaje_err = "Numero de Operacion Incorrecto...."
   ANEXO_CON21 = False
   Exit Function
End If

If all_llave!all_flag_ext = "E" Then
   pub_mensaje_err = "Operacion ya Extornada........"
   ANEXO_CON21 = False
   Exit Function
End If
If all_llave!all_flag_ext = "X" Then
   pub_mensaje_err = "Operacion NO Extornable........"
   ANEXO_CON21 = False
   Exit Function
End If


If all_llave!all_CODCIA <> LK_CODCIA Then
   pub_mensaje_err = "Cia. no coincide..."
   ANEXO_CON21 = False
   Exit Function
End If




PUB_NUM_OPER = all_llave!ALL_NUMOPER
PUB_NUM_OPER_EXT = all_llave!ALL_NUMOPER
PUB_CODCIA = all_llave!all_CODCIA
PUB_CODCIAL = all_llave!all_CODCIA
PUB_CODCLIE = all_llave!ALL_CODCLIE
ww_codtra_ext = all_llave!ALL_codtra

PUB_NUMDOC = Nulo_Valor0(all_llave!ALL_NUMDOC)
PUB_CODART = all_llave!ALL_CODART
PUB_IMPORTE_AMORT = all_llave!ALL_IMPORTE_amort
WS_IMPORTE_AMORT = all_llave!ALL_IMPORTE_amort

If all_llave!ALL_tipmov = 20 Then
   If Val(all_llave!ALL_IMPORTE_DOLL) <> 0 Then WS_IMPORTE_AMORT = all_llave!ALL_IMPORTE_DOLL
End If

PUB_IMPORTE = all_llave!ALL_IMPORTE
PUB_CP = all_llave!ALL_CP
PUB_TIPDOC = all_llave!ALL_TIPDOC
PUB_TIPMOV = all_llave!ALL_tipmov
PUB_CODBAN = all_llave!all_codban
PUB_CHENUM = all_llave!all_chenum
PUB_CHESER = Nulo_Valor0(all_llave!ALL_CHESER)
PUB_CHESEC = Nulo_Valor0(all_llave!ALL_CHESEC)
PUB_SERDOC = Nulo_Valor0(all_llave!ALL_serdoc)
PUB_NUMDOC = Nulo_Valor0(all_llave!ALL_NUMDOC)
ws_fecha_anterior = Nulo_Valor0(all_llave!all_fecha_ant)
PUB_CANTIDAD = all_llave!ALL_CANTIDAD
PUB_NUMSER = all_llave!ALL_NUMSER
PUB_FBG = all_llave!ALL_FBG
PUB_CODVEN = all_llave!ALL_CODVEN
PUB_NUMFAC = all_llave!all_numfac
PUB_SECUENCIA = all_llave!all_SECUENCIA
pub_signo_ccm = Nulo_Valor0(all_llave!ALL_SIGNO_CCM) * -1
pub_signo_car = Nulo_Valor0(all_llave!ALL_SIGNO_CAR) * -1
pub_signo_arm = Nulo_Valor0(all_llave!all_sIGNO_ARM) * -1
pub_signo_caja = Nulo_Valor0(all_llave!ALL_signo_caja) * -1
WS_FLAG_FACART = Nulo_Valors(all_llave!ALL_FACART)
PUB_CONCEPTO = "Extorno - " & all_llave!all_concepto
'PUB_SO = all_llave!ALL_FLAG_SO

If LK_CODTRA = 1122 Then PUB_CONCEPTO = "Anulacion - " & all_llave!all_concepto
If (Val(all_llave!ALL_IMPORTE_amort) + Val(all_llave!ALL_IMPORTE)) = 0 And LK_CODTRA = 1111 Then
 pub_signo_car = 0
End If

If pub_signo_ccm <> 0 Then
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_CCM_LLAVE
   If ccm_llave.EOF Then
      ANEXO_CON21 = False
      Exit Function
   End If
End If


End Function


Public Function ANEXO_CON1() As Boolean
Dim valor_venta As Currency
ANEXO_CON1 = True
If PUB_CP = "C" Or PUB_CP = "P" Then
Else
   GoTo fin
End If
SQ_OPER = 1
If PUB_CODCLIE = 0 Then
   pub_mensaje_err = "Cliente Invalido ..."
   ANEXO_CON1 = False
End If
pu_codcia = LK_CODCIA
pu_cp = PUB_CP
pu_codclie = PUB_CODCLIE
LEER_CLI_LLAVE
If cli_llave.EOF = True Then
   pub_mensaje_err = "Cliente/Proveedor no existe...."
   ANEXO_CON1 = False
   GoTo fin
End If
If PUB_CP = "C" And cli_llave!CLI_CODCLIE = 1 Then
Else
 i_nomCLI.Caption = Trim(cli_llave!cli_nombre) & " - " & Trim(cli_llave!cli_ruc_esposo)
 PUB_RUC = Nulo_Valors(cli_llave!cli_ruc_esposo)
End If
If LK_CODTRA = 1111 Or LK_CODTRA = 1122 Then GoTo fin
If LK_CODTRA = 2401 And PUB_PROCESO = 1 Then GoTo fin


If PUB_TIPMOV = 10 And PUB_FBG = "B" Then
   If Trim(i_ds.Text) = "D" Then
      valor_venta = redondea(Val(i_neto.Text) * JALAR_TC(LK_FECHA_DIA, 3))
   Else
      valor_venta = Val(i_neto.Text)
   End If
   If valor_venta >= Val(GEN!GEN_UIT) And Trim(cli_llave!cli_RUC_ESPOSA) = "" Then
      pub_mensaje_err = "Falta DNI para la Boleta ...por pasar de la 1/2 UIT que es = S/. " & Format(GEN!GEN_UIT, "#,##0.00")
      ANEXO_CON1 = False
      GoTo fin
   End If
End If

   If PUB_TIPMOV <> 0 And LK_CODTRA <> 2748 And LK_CODTRA <> 2412 And LK_CODTRA <> 2410 And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 Then
      If Trim(i_moneda.Text) = "" Then
         pub_mensaje_err = "Falta La Moneda ..."
         ANEXO_CON1 = False
         GoTo fin
      End If
   Else
   ' QUITE ACV
    'If PUB_TIPMOV <> 0 And LK_CODTRA <> 2748 And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 Then
    '  If Trim(lmoneda.Caption) = "" Then
    '     pub_mensaje_err = "Falta LA Moneda ..."
    '     exito = False
    '     Return
    '  End If
    'End If
   End If
     

If PUB_TIPMOV <> 0 And LK_CODTRA <> 1111 And LK_CODTRA <> 1122 Then
If WS_MONEDA_CLI = "S" Or WS_MONEDA_CLI = "D" Then
Else
   pub_mensaje_err = "Falta LA Moneda ..."
   ANEXO_CON1 = False
   GoTo fin
End If
End If

If (PUB_TIPMOV = 10 And PUB_FBG = "F") And LK_CODTRA <> 1111 Then
   If Trim(cli_llave!cli_ruc_esposo) = "" And Nulo_Valors(par_llave!PAR_VENTAS_IGV) <> "E" Then
      pub_mensaje_err = "Falta el N. Ruc ..."
      ANEXO_CON1 = False
      GoTo fin
   End If
End If

If LK_CODTRA <> 1122 And LK_CODTRA <> 2412 And LK_CODTRA <> 2410 Then
If (Nulo_Valors(cli_llave!CLI_TIPO_BLOQ1) = "1" Or Nulo_Valors(cli_llave!CLI_TIPO_BLOQ2) = "1" Or Nulo_Valors(cli_llave!CLI_TIPO_BLOQ3) = "1" Or Nulo_Valors(cli_llave!CLI_TIPO_BLOQ4) = "1") And pub_signo_car = 1 And LK_CODTRA <> 1111 And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 Then
   pub_mensaje_err = "Credito Bloqueado ..."
   ANEXO_CON1 = False
End If
End If

fin:


End Function
Public Function ANEXO_CON14() As Boolean

SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_codclie = PUB_CODCLIE
pu_cp = PUB_CP
LEER_CAR_LLAVE
If car_llave.EOF = False Then
   pub_mensaje_err = "Ya existe Cartera ... "
   ANEXO_CON14 = False
End If
End Function

Public Sub ANEXO_ACT17()
Dim CONTADO As Integer
CONTADOR = 0
Do Until CONTADOR = PUB_CANT_CHEQ
che_llave.AddNew
che_llave!CHE_CODBAN = PUB_CODBAN
If LK_EMP = "3AA" Then
    che_llave!CHE_CODCIA = par_llave!PAR_CIACCM
Else
    che_llave!CHE_CODCIA = LK_CODCIA
End If

che_llave!che_cheser = PUB_CHESER
che_llave!che_chesec = 0
che_llave!CHE_CHENUM = CONTADOR + PUB_NUM_INI
CONTADOR = CONTADOR + 1
che_llave!CHE_IMPORTE = 0
che_llave!CHE_FECHA = 0
che_llave!CHE_CODUSU = LK_CODUSU
che_llave!che_concepto = " "
che_llave!CHE_NUMOPER = CONTADOR + 30000
che_llave!che_saldo = 0
che_llave!CHE_SIGNO_CCM = 0
che_llave!che_estado = " "
che_llave!che_abreviado = ""
che_llave!CHE_codtra = LK_CODTRA
che_llave!che_chenum_ext = 0
che_llave.Update
Loop
PUB_CHENUM = 0

End Sub
Public Sub ANEXO_ACT19()

   If PUB_TIPMOV = 97 Then
   If Nulo_Valors(par_llave!PAR_FLAG_NCRE) = "A" Then
      par_llave.Edit
      par_llave!PAR_FLAG_NCRE = ""
      par_llave.Update
   End If
   End If
   
   If PUB_TIPMOV = 98 Then
   If Nulo_Valors(par_llave!PAR_FLAG_NDEB) = "A" Then
      par_llave.Edit
      par_llave!PAR_FLAG_NDEB = ""
      par_llave.Update
   End If
   End If

   
   If FORMGEN.i_fbg.Visible = False Then GoTo CHAU
   
   
   If LK_FLAG_FACTURACION = "A" And PUB_TIPMOV = 10 Then
   If par_llave!PAR_FLAG_G = "A" Or par_llave!PAR_FLAG_B = "A" Or par_llave!PAR_FLAG_F = "A" Then
      par_llave.Edit
      If PUB_FBG = "G" And par_llave!PAR_FLAG_G = "A" Then
         par_llave!PAR_FLAG_G = ""
      ElseIf PUB_FBG = "B" And par_llave!PAR_FLAG_B = "A" Then
         par_llave!PAR_FLAG_B = ""
      ElseIf PUB_FBG = "F" And par_llave!PAR_FLAG_F = "A" Then
         par_llave!PAR_FLAG_F = ""
      End If
      par_llave.Update
   End If
   End If
   
   
   
If LK_FLAG_FACTURACION = "V" And PUB_TIPMOV = 10 Then
   If (ven_llave!VEM_FLAG_G = "A" Or ven_llave!VEM_FLAG_B = "A" Or ven_llave!VEM_FLAG_F = "A" Or ven_llave!VEM_FLAG_P = "A") Then
        ven_llave.Edit
        If PUB_FBG = "G" And ven_llave!VEM_FLAG_G = "A" Then
           ven_llave!VEM_FLAG_G = ""
        ElseIf PUB_FBG = "B" And ven_llave!VEM_FLAG_B = "A" Then
           ven_llave!VEM_FLAG_B = ""
        ElseIf PUB_FBG = "F" And ven_llave!VEM_FLAG_F = "A" Then
           ven_llave!VEM_FLAG_F = ""
        ElseIf PUB_FBG = "P" And ven_llave!VEM_FLAG_P = "A" Then
           ven_llave!VEM_FLAG_P = ""
        End If
        ven_llave.Update
    End If
End If

Exit Sub
CHAU:
End Sub



Public Function ANEXO_CON3() As Boolean
Dim WTC  As Currency
Dim WMO As String
Dim wsumadol As Currency
Dim RES_DEUDA As Currency
ANEXO_CON3 = True
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_VEN_LLAVE
If ven_llave.EOF Then
   pub_mensaje_err = "Vendedor no existe... "
   FORMGEN.i_codven.SetFocus
   ANEXO_CON3 = False
End If
If LK_CODTRA = 2401 Then
  If ven_llave!VEM_CODVEN = 0 Then
    pub_mensaje_err = "Vendedor no procede... "
    FORMGEN.i_codven.SetFocus
    ANEXO_CON3 = False
  End If
End If
If LK_CODTRA = 2770 Then Exit Function
If PUB_PROCESO = 1 Then Exit Function


If PUB_TIPMOV = 97 Or PUB_TIPMOV = 98 Then Exit Function

If Not cli_llave.EOF And Nulo_Valors(par_llave!par_flag_cred) <> "A" And Nulo_Valor0(SUT_LLAVE!SUT_FLAG_CC) = 0 Then
   If SUT_LLAVE!SUT_SIGNO_CAR = 1 Then
      pu_codcia = LK_CODCIA
      pub_deuda = CAR_TOT_CPX2("C", pu_codcia, cli_llave!CLI_CODCLIE)
      If PUB_FLAG_VENCIDO = 1 And LK_FLAG_LIMITE <> "A" And LK_FLAG_LIMITE <> "C" Then
         pub_mensaje_err = "CLIENTE TIENE OBLIGACIONES VENCIDAS ... "
         ANEXO_CON3 = False
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
If WTC = 0 Then
  pub_mensaje_err = "Venta falta parametros ...INGRESE TIPO DE CAMBIO DEL DIA"
  ANEXO_CON3 = False
  GoTo PA
End If
If Trim(i_ds.Text) = "S" Then
 wsumadol = Val(Nulo_Valor0(cli_llave!cli_limcre)) + Val(redondea((Nulo_Valor0(cli_llave!cli_limcre2) * WTC)))
 RES_DEUDA = pub_deuda
 WMO = "S/."
Else
 wsumadol = Val(redondea(Nulo_Valor0(cli_llave!cli_limcre) / WTC)) + Val(redondea(Val(Nulo_Valor0(cli_llave!cli_limcre2))))
 RES_DEUDA = redondea(Val(pub_deuda / WTC))
 WMO = "US$."
End If
'grid_fac.TextMatrix(1, 0) = "Disp.: " & WSIMBOL & " " & Format(wsumadol - RES_DEUDA - WS_NETO, "##,###,###.00")
'wsumadol = Val(Nulo_Valor0(cli_llave!cli_limcre)) + redondea((Nulo_Valor0(cli_llave!cli_limcre2) * WTC))
If (RES_DEUDA + Val(i_neto.Text)) > wsumadol And LK_FLAG_LIMITE <> "B" And LK_FLAG_LIMITE <> "C" Then
   pub_mensaje_err = "LIMITE DE CREDITO EXCEDIDO ...SALDO POR ATENDER : " & WMO & " " & Format(wsumadol - RES_DEUDA, "0.00")
   ANEXO_CON3 = False
End If
PA:
      
   End If
End If

If Nulo_Valors(SUT_LLAVE!SUT_CERO) <> "A" And LK_CODTRA <> 2745 Then
If WS_BRUTO = 0 Then
   pub_mensaje_err = "Importe errado... "
   ANEXO_CON3 = False
End If
End If



End Function

Public Sub manda_numero()
SQ_OPER = 2
PUB_FECHA = LK_FECHA_DIA
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
If all_menor.EOF = False Then
   PUB_NUM_OPER_XXX = all_menor!ALL_NUMOPER
Else
   PUB_NUM_OPER_XXX = 0
End If
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1

End Sub



Public Function Anexo_con7() As Boolean

Anexo_con7 = True
If SUT_LLAVE!SUT_signo_ccm = 0 Then
  Exit Function
End If
flag_repite = 0
PUB_CODBAN = Val(i_codban.Text)
SQ_OPER = 1
pu_codcia = LK_CODCIA
LEER_CCM_LLAVE
If ccm_llave.EOF Then
   pub_mensaje_err = "banco no existe..."
   FORMGEN.i_codban.SetFocus
   Anexo_con7 = False
   Exit Function
End If

If PUB_IMPORTE = 0 And LK_CODTRA <> 2748 Then
    pub_mensaje = "Es Anulacion de Cheque ..Desea Continuar ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
        Anexo_con7 = False
        Exit Function
    End If
End If


WS_MONEDA_CCM = ccm_llave!CCM_MONEDA
   
   
If gridl.Visible = True And pub_signo_car <> 0 Then
   If WS_MONEDA_CCM <> WS_MONEDA_CLI Then
         If Val(i_importe.Text) = Val(gridl.TextMatrix(1, 9)) Then
            pub_mensaje_err = "Falta considerar el tipo de cambio..."
            Anexo_con7 = False
         End If
      End If
End If

End Function

Public Sub final1()
 If grid_liq.Visible = True Then
    pub_cadena = grid_liq.FormatString
    grid_liq.Clear
    grid_liq.FormatString = pub_cadena
    grid_liq.Visible = False
 End If
 
End Sub
Public Function Anexo_con12() As Boolean
Anexo_con12 = True
If LK_CODTRA = 2401 And PUB_PROCESO = 1 Then GoTo SALTA


If i_cambio.Value = 1 And i_cambio.Visible = True Then Exit Function


If Val(i_numguia.Text) <> 0 And PUB_TIPMOV = 10 Then
SQ_OPER = 6
PUB_SERGUIA = Val(i_serguia.Text)
PUB_NUMGUIA = Val(i_numguia.Text)
pu_codcia = LK_CODCIA
LEER_FAR_LLAVE
If far_guia.EOF = False Then
   pub_mensaje_err = "Guia ya existe en Doc: " & far_guia!FAR_FBG & " - " & far_guia!far_numfac
   Anexo_con12 = False
   Exit Function
End If
End If

If i_cambio.Visible = False Or i_cambio.Value = 0 Then Exit Function

SALTA:
SQ_OPER = 1
PU_TIPMOV = PUB_TIPMOV
PU_NUMSER = Val(i_numser.Text)
PU_NUMFAC = Val(i_numfac.Text)
pu_codcia = LK_CODCIA
If LK_CODTRA = 1401 Or LK_CODTRA = 2414 Then
   PU_FBG = ""
Else
   PU_FBG = i_fbg.Text
End If
LEER_FAR_LLAVE
If far_llave.EOF = False Then
   Do Until far_llave.EOF
      WS_NUMSEC = far_llave!FAR_NUMSEC
      If far_llave!far_estado <> "E" Then
         If PUB_PROCESO = 1 Then
          loc_ped_llave.MoveFirst
          If Not loc_ped_llave.EOF Then
             pub_mensaje_err = "El Documento " & Trim(far_llave!FAR_FBG) & "/" & far_llave!far_numser & "-" & far_llave!far_numfac & " No esta extornado " & Chr(13) & "Para el Pedido : " & loc_ped_llave!PED_NUMFAC
          Else
            pub_mensaje_err = "El Documento " & Trim(far_llave!FAR_FBG) & "/" & far_llave!far_numser & "-" & far_llave!far_numfac & " No esta extornado "
          End If
         Else
           pub_mensaje_err = "Documento no esta extornado..."
         End If
          Anexo_con12 = False
          Exit Function
      End If
      far_llave.MoveNext
   Loop
   WS_NUMSEC = WS_NUMSEC + 200
End If

End Function

Public Function ANEXO_CON15() As Boolean
ANEXO_CON15 = True
If PUB_CHENUM = 0 Then Exit Function
If (LK_EMP = "3AA" Or LK_EMP = "PIU") And (LK_CODTRA = 2748 Or LK_CODTRA = 2735 Or LK_CODTRA = 5714 Or LK_CODTRA = 2738 Or LK_CODTRA = 2720) Then Exit Function

If SUT_LLAVE!SUT_signo_ccm = 0 Then Exit Function
'Verifica si es un cheque ya trabajado.. que existe en archivo...
  PUB_CHESEC = 99
  SQ_OPER = 1
  LEER_CHE_LLAVE
  If che_llave.EOF Then
     pub_mensaje_err = "Numero de Cheque no Existe ...."
     ANEXO_CON15 = False
     Exit Function
  End If
  che_llave.MoveLast
' ?????? REVISAR ESTO EL IGUAL ???
If che_llave!CHE_FECHA > #1/1/1900# And che_llave!CHE_FECHA <> LK_FECHA_DIA And che_llave!che_estado <> "E" Then
   pub_mensaje_err = "Cheque ya procesado...otro dia..."
   ANEXO_CON15 = False
   Exit Function
End If
If i_fecha_can.Visible And IsDate(i_fecha_can.Text) = False Then
   pub_mensaje_err = "Fecha Invalidad o no procede..."
   ANEXO_CON15 = False
   Exit Function
End If
If che_llave!CHE_FECHA > #1/1/1900# Then
   PUB_CHESEC = che_llave!che_chesec + 1
   Exit Function
End If
' ese cheque es el que sigue en la secuencia ????
' lectura de todos los cheques de fecha 0 y serie xxx

   PUB_CHESEC = 0
   SQ_OPER = 3
   PUB_FECHA = 0
   LEER_CHE_LLAVE
   If che_menor.EOF Then
      pub_mensaje_err = "Numero de Serie errado ...."
      ANEXO_CON15 = False
      Exit Function
   End If
che_menor.MoveFirst
   
If PUB_CHENUM = che_menor!CHE_CHENUM Then
Else
'   muestra_cheques... CAMBIAR A CHE_MENOR
' QUITE CAMTEX
   'pub_mensaje_err = "No corresponde numero de Cheque ...?"
   'exito = False
   'Return
End If

   

End Function

Public Function ANEXO_CON6() As Boolean
ANEXO_CON6 = True
If i_codban2.Visible = True Then
   PUB_CODBAN = Val(i_codban2.Text)
   pu_codcia = LK_CODCIA
   LEER_CCM_LLAVE
   If ccm_llave.EOF Then
      pub_mensaje_err = "banco2 no existe..."
      FORMGEN.i_codban2.SetFocus
      ANEXO_CON6 = False
      Exit Function
   End If
End If
End Function
Public Function ANEXO_CON18() As Boolean
ANEXO_CON18 = True
SQ_OPER = 1
pu_cp = PUB_CP
pu_codclie = PUB_CODCLIE
pu_codcia = LK_CODCIA
LEER_CAR_LLAVE
If car_llave.EOF = True Then
   pub_mensaje_err = "DOCUMENTO EN CARTERA NO EXISTE..."
   ANEXO_CON18 = False
   Exit Function
End If
End Function
Public Function ANEXO_CON22() As Boolean
ANEXO_CON22 = True

SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_cp = PUB_CP
pu_codclie = PUB_CODCLIE
LEER_CAR_LLAVE
If car_llave.EOF = True Then
    pub_mensaje_err = "DOCUMENTO NO EXISTE..."
    ANEXO_CON22 = False
    Exit Function
End If

If car_llave!CAR_FECHA_INGR <> LK_FECHA_DIA Then
    pub_mensaje_err = "NO PUEDE MODIFICAR DOCS... DE OTRO DIA..."
    ANEXO_CON22 = False
    Exit Function
End If

End Function
Public Sub ANEXO_ACT13()

      far_llave!FAR_TIPMOV = PUB_TIPMOV
      far_llave!FAR_CODCIA = LK_CODCIA
      far_llave!far_numser = PUB_NUMSER
      far_llave!far_numfac = PUB_NUMFAC
      far_llave!FAR_NUMSEC = WS_NUMSEC + 1
      far_llave!far_codart = 0
      far_llave!far_precio_neto = 0
      far_llave!far_cantidad = PUB_CANTIDAD
      far_llave!far_UNIDADES = 0
      far_llave!far_JABAS = 0
      far_llave!far_mortal = 0
      far_llave!far_codclie = PUB_CODCLIE
      far_llave!FAR_RUC = PUB_RUC
      far_llave!FAR_cp = PUB_CP
      far_llave!far_transito = " "
      far_llave!FAR_STOCK = 0
      far_llave!FAR_COSPRO = PUB_COSPRO
      far_llave!FAR_PRECIO = PUB_PRECIO
      far_llave!FAR_FBG = PUB_FBG
      If LK_CODTRA = 2748 Then
        far_llave!FAR_FBG = "K"
      End If
      far_llave!far_IMPTO = PUB_IMPTO
      far_llave!FAR_TOT_DESCTO = 0
      far_llave!FAR_DESCTO = 0
      far_llave!FAR_GASTOS = PUB_GASTOS
      far_llave!FAR_BRUTO = PUB_SUBTOTAL
      
      far_llave!far_num_precio = "0"
      far_llave!far_otra_cia = ""
      far_llave!FAR_fecha = LK_FECHA_DIA
      far_llave!FAR_NUMSER_C = PUB_NUMSER_C
      far_llave!FAR_NUMFAC_C = PUB_NUMFAC_C
      far_llave!FAR_NUMOPER = PUB_NUM_OPER_XXX
      far_llave!far_tipo_bloq_act1 = PUB_TIPO_BLOQ_act1
      far_llave!FAR_CLIENTE = Trim(FORMGEN.t_nombre.Text)
      far_llave!FAR_DOCCLI = Trim(FORMGEN.t_doc.Text)
      far_llave!FAR_dircli = Trim(FORMGEN.t_direc.Text)
      'far_llave!far_tipo_bloq_act2 = PUB_TIPO_BLOQ_act2
      'far_llave!far_tipo_bloq_act3 = PUB_TIPO_BLOQ_act3
      far_llave!FAR_TIPO_BLOQ_ANT1 = PUB_TIPO_BLOQ_ant1
      far_llave!far_tipo_bloq_ant2 = PUB_TIPO_BLOQ_ant2
      far_llave!far_key_dircli = Val(Right(i_dircli.Text, 8))
      far_llave!far_estado = "N"
      far_llave!FAR_ESTADO2 = "N"

End Sub
Public Sub ANEXO_ACT10()
If LK_CODTRA <> 2407 Then Exit Sub
If i_cambio.Value = 1 And i_cambio.Visible Then Exit Sub
FAR_TRANS.MoveFirst
Do Until FAR_TRANS.EOF
   If Val(FAR_TRANS!far_numfac) = Val(grid_trans.TextMatrix(grid_trans.Row, 0)) Then
      FAR_TRANS.Edit
      FAR_TRANS!far_transito = "X"
      FAR_TRANS.Update
   End If
   FAR_TRANS.MoveNext
Loop


If Frame1.Visible = True Then
   cli_llave.Edit
   cli_llave!CLI_TIPO_BLOQ1 = BLOQ(0).Value
   cli_llave!CLI_TIPO_BLOQ2 = BLOQ(1).Value
   cli_llave!CLI_TIPO_BLOQ3 = BLOQ(2).Value
   cli_llave!CLI_TIPO_BLOQ4 = BLOQ(3).Value
   cli_llave.Update
End If

End Sub
Public Sub ANEXO_ACT12()
cli_llave.Edit
If Frame1.Visible = True Then
   cli_llave!CLI_TIPO_BLOQ1 = BLOQ(0).Value
   cli_llave!CLI_TIPO_BLOQ2 = BLOQ(1).Value
   cli_llave!CLI_TIPO_BLOQ3 = BLOQ(2).Value
   cli_llave!CLI_TIPO_BLOQ4 = BLOQ(3).Value
End If
If LK_CODTRA = 2580 Then
  If Trim(i_ds.Text) = "S" Then
     If Nulo_Valor0(cli_llave!cli_limcre) <> PUB_LIMCRE_ACT Then cli_llave!cli_limcre = PUB_LIMCRE_ACT
  Else
     If Nulo_Valor0(cli_llave!cli_limcre2) <> PUB_LIMCRE_ACT Then cli_llave!cli_limcre2 = PUB_LIMCRE_ACT
  End If
  cli_llave!CLI_AUTO1 = i_dias.Text
  cli_llave!cli_DIAS_FAC = Val(Right(i_condi.Text, 6))
  cli_llave!CLI_TIPO = Trim(i_fbg.Text)
  
End If

cli_llave.Update

End Sub

Public Function ANEXO_CON17() As Boolean
ANEXO_CON17 = True
If PUB_FBG = "B" Or PUB_FBG = "F" Then
   If PUB_IMPTO = 0 And Nulo_Valors(SUT_LLAVE!SUT_CERO) <> "A" Then
       pub_mensaje_err = "Falta el Igv.."
       ANEXO_CON17 = False
       Exit Function
   End If
End If



End Function
Public Function ANEXO_CON19() As Boolean
ANEXO_CON19 = True
If LK_CODTRA <> 2407 Then Exit Function
If i_cias.ListIndex = -1 Then
   pub_mensaje_err = "Seleccion Destino ..."
   ANEXO_CON19 = False
   Exit Function
End If


End Function

Public Function ANEXO_CON23() As Boolean
ANEXO_CON23 = True
If IsDate(i_fecha_compra.Text) = False Then
   pub_mensaje_err = "Falta Fecha de Emision o es Invalidad..."
   ANEXO_CON23 = False
   Exit Function
End If

End Function

Public Sub VER_LINEAS()
Dim WS_NRO_ITEMS As Integer
Dim h As Integer
LK_QUEDA = ""
For h = 2 To grid_fac.Rows - 1
 If Trim(grid_fac.TextMatrix(h, 1)) = "" Then GoTo dale
  grid_fac.TextMatrix(h, 38) = ""
  If Left(grid_fac.TextMatrix(h, 0), 1) = "*" Then
     grid_fac.TextMatrix(h, 0) = Mid(grid_fac.TextMatrix(h, 0), 2, Len(grid_fac.TextMatrix(h, 0)))
  End If
  grid_fac.CellBackColor = QBColor(15)
  If Trim(i_fbg.Text) = "F" Then
     If Nulo_Valor0(par_llave!par_fac_lines) < h - 1 Then
       grid_fac.TextMatrix(h, 38) = "1"
       If Left(grid_fac.TextMatrix(h, 0), 1) <> "*" Then
         grid_fac.TextMatrix(h, 0) = "*" & grid_fac.TextMatrix(h, 0)
       End If
       LK_QUEDA = "A"
     End If
  End If
  If Trim(i_fbg.Text) = "B" Then
      If PUB_CODVEN = 10 Then
          WS_NRO_ITEMS = 13 ' TEMPORAL
      Else
          WS_NRO_ITEMS = Nulo_Valor0(par_llave!par_BOL_lines)
      End If
        
      If WS_NRO_ITEMS < h - 1 Then
         grid_fac.TextMatrix(h, 38) = "1"
         If Left(grid_fac.TextMatrix(h, 0), 1) <> "*" Then
            grid_fac.TextMatrix(h, 0) = "*" & grid_fac.TextMatrix(h, 0)
         End If
         LK_QUEDA = "A"
      End If
  End If
dale:
Next h
End Sub
Private Sub TOTALES()
Dim ww_TOTAL As Currency
fila = 2
ww_TOTAL = 0
For fila = 2 To grid_fac.Rows - 1
  ww_TOTAL = ww_TOTAL + Val(grid_fac.TextMatrix(fila, 7))
Next fila
 i_camal.Enabled = False
 i_camal.Top = 5760
 i_camal.Left = 8880
 i_camal.BackColor = QBColor(14)
 i_camal.Font.Size = 9
 i_camal.Font.Size = 10
 i_camal.Width = 1300
 i_camal.Height = 300
 i_camal.Visible = True
 i_camal.Text = ww_TOTAL
End Sub


Public Sub ANEXO_ACT18()
SQ_OPER = 1
' quite este ACV
'PUB_FECHA = Grid_all.TextMatrix(2, 11)
PUB_FECHA = FORMGEN.Grid_all.TextMatrix(1, 11)
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
Do Until all_llave.EOF
  If PUB_NUM_OPER_EXT = all_llave!ALL_NUMOPER Then
    all_llave.Edit
    all_llave!all_flag_ext = "E"
    all_llave.Update
    Exit Do
  Else
    all_llave.MoveNext
  End If
Loop
If all_llave.EOF Then
   MsgBox "Error en Extorno de Allog ..."
   Exit Sub
End If

If all_llave!ALL_codtra = 1455 Then pub_signo_ccm = 0
If all_llave!ALL_codtra = 2725 And LK_EMP = "3AA" Then pub_signo_ccm = 0
If all_llave!ALL_codtra = 2401 And LK_EMP = "3AA" Then pub_signo_ccm = 0
If all_llave!ALL_codtra = 2738 Then GoTo SSA
If all_llave!ALL_codtra = 2720 Then GoTo SSA
If pub_signo_ccm <> 0 Then
   SQ_OPER = 2
   PUB_FECHA = all_llave!ALL_FECHA_DIA
   LEER_CHE_LLAVE
   If che_oper.EOF = True Then
      MsgBox "Error en Bancos..no existe tal cuenta..."
      Exit Sub
   End If
   che_oper.Edit
   che_oper!che_estado = "E"
   che_oper.Update
   PUB_CHESEC = che_oper!che_chesec + 50
   PUB_CHENUM_EXT = PUB_CHENUM
   PUB_CHENUM = 0
   PUB_ABREVIADO = "EXT"
End If
SSA:
End Sub
Public Function ANEXO_CON13() As Boolean
Dim fx As Integer
ANEXO_CON13 = True

If LK_CODTRA = 1133 Then PUB_FBG = "N"

SQ_OPER = 1
PU_TIPMOV = PUB_TIPMOV
PU_NUMSER = PUB_NUMSER_C
PU_NUMFAC = PUB_NUMFAC_C
pu_codcia = LK_CODCIA
PU_FBG = PUB_FBG
LEER_FAR_LLAVE
If far_llave.EOF = True Then
   pub_mensaje_err = "No existe Documento ..."
   LKCHEK = True
   ANEXO_CON13 = False
   Exit Function
End If

If LK_CODTRA <> 1122 And LK_CODTRA <> 1133 Then Exit Function
fx = 0
Do Until fx = 99
   If far_llave!far_estado <> "E" Then fx = 99
   If fx <> 99 Then
      far_llave.MoveNext
   End If
   If far_llave.EOF Then
      pub_mensaje_err = "DOCUMENTO YA  ANULADO ..."
      ANEXO_CON13 = False
      Exit Function
   End If
Loop

PUB_CODCIA = far_llave!FAR_CODCIA
PUB_CODCIAL = far_llave!FAR_CODCIA
PUB_CODCLIE = far_llave!far_codclie
WS_IMPORTE_AMORT = Val(far_llave!FAR_BRUTO) + Val(far_llave!far_IMPTO) - Val(far_llave!FAR_TOT_DESCTO)
PUB_IMPORTE_AMORT = WS_IMPORTE_AMORT
PUB_CP = far_llave!FAR_cp
PUB_TIPMOV = far_llave!FAR_TIPMOV
PUB_CODBAN = 0
PUB_CHENUM = 0
PUB_CHESER = 0
PUB_CHESEC = 0
PUB_SERDOC = 0
PUB_NUMDOC = 0
PUB_CANTIDAD = 0
PUB_NUMSER = far_llave!FAR_NUMSER_C
PUB_FBG = far_llave!FAR_FBG
PUB_CODVEN = far_llave!FAR_CODVEN
PUB_NUMFAC = far_llave!FAR_NUMFAC_C
i_numfac.Text = far_llave!far_numfac

PUB_SECUENCIA = far_llave!FAR_NUMSEC
PUB_NUM_INI = 0
pub_signo_car = Nulo_Valor0(far_llave!far_signo_car) * -1
If pub_signo_car <> 0 Then PUB_TIPDOC = "FA"

pub_signo_arm = 1
pub_signo_caja = 0
PUB_CONCEPTO = "Anulado - " & far_llave!far_concepto


If pub_signo_car <> 0 Then
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    pu_codclie = PUB_CODCLIE
    pu_cp = PUB_CP
    fx = 0
    LEER_CAR_LLAVE
    Do Until car_mayor.EOF Or fx = 99
    'revisar cambie pub_numfac por pu_numfac para anul de ventas
        If PU_NUMFAC = Val(car_mayor!car_NUMFAC) And PU_NUMSER = Val(car_mayor!car_NUMSER) Then
           fx = 99
        Else
           car_mayor.MoveNext
        End If
    Loop
    If Not car_mayor.EOF Then
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_codclie = PUB_CODCLIE
    PUB_SERDOC = car_mayor!car_SERDOC
    PUB_NUMDOC = car_mayor!car_NUMDOC
    PUB_TIPDOC = car_mayor!car_TIPDOC
    pu_cp = PUB_CP
    LEER_CAR_LLAVE
'no importa si cancelo...
'    If car_llave!CAR_IMPORTE <> car_llave!CAR_IMP_INI Then
'       pub_mensaje_err = "DOCUMENTO NO PUEDE SER ANULADO...SALDO DIFERENTE "
'       exito = False
'       Return
'    End If
    
    End If
    
     ANEXO_CON13 = LEE_CARACU()
End If

SQ_OPER = 5
PUB_FECHA = far_llave!FAR_fecha
PUB_NUM_OPER_EXT = far_llave!FAR_NUMOPER
LEER_FAR_LLAVE
'Print far_menor3!far_codclie
If far_menor3.EOF = True Then
   pub_mensaje_err = "!!! NO HAY ESE DOCUMENTO o YA ANULADO..."
   exito = False
   Return
End If

If LK_CODTRA = 1122 Then
If far_menor3!FAR_fecha = LK_FECHA_DIA Then
   pub_mensaje_err = "!!! Use el extorno de Operaciones ... "
   ANEXO_CON13 = False
   Exit Function
End If
End If
fila = 2
Do Until far_menor3.EOF

   grid_fac.Rows = fila + 2
   WS_NUMSEC = far_menor3!FAR_NUMSEC
   
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_KEY = far_menor3!far_codart
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then Exit Do
   grid_fac.TextMatrix(fila, 0) = art_LLAVE!art_nombre
   
   grid_fac.TextMatrix(fila, 16) = far_menor3!far_codart
   grid_fac.TextMatrix(fila, 33) = far_menor3!far_codart
   grid_fac.TextMatrix(fila, 2) = far_menor3!far_JABAS
   grid_fac.TextMatrix(fila, 3) = far_menor3!far_UNIDADES
   grid_fac.TextMatrix(fila, 11) = far_menor3!FAR_COSPRO
   grid_fac.TextMatrix(fila, 4) = far_menor3!far_cantidad / far_menor3!FAR_equiv
   grid_fac.TextMatrix(fila, 6) = far_menor3!FAR_PRECIO
   grid_fac.TextMatrix(fila, 14) = far_menor3!FAR_equiv
   grid_fac.TextMatrix(fila, 5) = far_menor3!far_descri
   grid_fac.TextMatrix(fila, 8) = far_menor3!FAR_FLETE
   grid_fac.TextMatrix(fila, 10) = far_menor3!FAR_DESCTO
   grid_fac.TextMatrix(fila, 11) = far_menor3!FAR_COSPRO
   grid_fac.TextMatrix(fila, 12) = far_menor3!far_signo_arm * -1
   
   
   far_menor3.MoveNext
   fila = fila + 1
   
Loop


WS_NUMSEC = WS_NUMSEC + 200
calcula_totales


End Function

Public Sub ANEXO_ACT6()

End Sub
Public Sub ANEXO_ACT21()
Dim WF As Integer

FAR_TRANS.MoveFirst
Do Until FAR_TRANS.EOF

   For WF = 2 To grid_fac.Rows - 1
    If Val(FAR_TRANS!FAR_NUMSER_C) = Val(grid_fac.TextMatrix(WF, 39)) And Val(FAR_TRANS!FAR_NUMFAC_C) = Val(grid_fac.TextMatrix(WF, 40)) Then
      If Val(FAR_TRANS!far_codart) = Val(grid_fac.TextMatrix(WF, 16)) And Val(FAR_TRANS!FAR_NUMSEC) = Val(grid_fac.TextMatrix(WF, 41)) Then
        FAR_TRANS.Edit
        FAR_TRANS!far_cantidad_p = FAR_TRANS!far_cantidad_p + (Val(grid_fac.TextMatrix(WF, 4)) * Val(grid_fac.TextMatrix(WF, 14)))
        FAR_TRANS.Update
        'para no volver a actualizarlo
        grid_fac.TextMatrix(WF, 16) = -99999
        Exit For
      End If
    End If
   Next WF
FAR_TRANS.MoveNext
Loop


End Sub

Public Sub ANEXO_ACT20()
far_menor4.MoveFirst
Do Until far_menor4.EOF
far_menor4.Edit
far_menor4!FAR_BRUTO = Val(i_bruto2.Text)
far_menor4!far_IMPTO = Val(i_impto2.Text)
far_menor4!FAR_MONEDA = i_ds.Text
far_menor4!FAR_fecha_compra = i_fecha_compra.Text
far_menor4!FAR_NUMFAC_C = Val(i_numfac_c.Text)
far_menor4!FAR_NUMSER_C = Val(i_numser_c.Text)
far_menor4!FAR_DIAS = Val(i_dias.Text)
far_menor4.Update
far_menor4.MoveNext
Loop
all_llave.Edit
all_llave!ALL_BRUTO = Val(i_bruto2.Text)
all_llave!ALL_IMPTO = Val(i_impto2.Text)
all_llave!ALL_IMPORTE_amort = Val(i_bruto2.Text) + Val(i_impto2.Text)
all_llave!all_numfac_c = Val(i_numfac_c.Text)
all_llave!all_numser_c = Val(i_numser_c.Text)
all_llave!ALL_FECHA_SUNAT = i_fecha_compra.Text
all_llave!ALL_FECHA_VCTO = CDate(i_fecha_vcto.Text)
all_llave.Update

car_far.Edit
car_far!car_fecha_sunat = CDate(i_fecha_compra.Text)
car_far!car_importe = Val(i_bruto2.Text) + Val(i_impto2.Text)
car_far!CAR_IMP_INI = Val(i_bruto2.Text) + Val(i_impto2.Text)
car_far!car_numfac_c = Val(i_numfac_c.Text)
car_far!car_numser_c = Val(i_numser_c.Text)
car_far!CAR_MONEDA = i_ds.Text
car_far!car_fecha_vcto = CDate(i_fecha_vcto.Text)
car_far!car_fecha_VCTO_orig = i_fecha_vcto.Text

car_far.Update

caa_histo.Edit
caa_histo!CAA_IMPORTE = Val(i_bruto2.Text) + Val(i_impto2.Text)
caa_histo!CAA_SALDO = Val(i_bruto2.Text) + Val(i_impto2.Text)
caa_histo!caa_SALDO_car = Val(i_bruto2.Text) + Val(i_impto2.Text)
caa_histo!CAA_TOTAL = Val(i_bruto2.Text) + Val(i_impto2.Text)
caa_histo!caa_numfac_c = Val(i_numfac_c.Text)
caa_histo!caa_numser_c = Val(i_numser_c.Text)
caa_histo.Update

End Sub

Public Sub ANEXO_ACT9()
Dim WS_CANTIDAD As Currency
Dim FLAG
PUB_NUM = 0
pu_codcia = LK_CODCIA
PUB_FECHA = #1/1/1900#
pub_autkey = 30000
SQ_OPER = 3
LEER_AUT_LLAVE
If aut_menor.EOF = False Then
   aut_menor.MoveLast
   PUB_NUM = aut_menor!aut_key
End If

WS_NUMSEC = 0
fila = 2
FLAG = False
Do While FLAG = False
   If Val(grid_fac.TextMatrix(fila, 16)) = 0 Or (Val(grid_fac.TextMatrix(fila, 3)) = 0 And Val(grid_fac.TextMatrix(fila, 4)) = 0) Then
     GoTo OTROMAS2
   End If
PUB_NUM = PUB_NUM + 1
   PUB_CODART = grid_fac.TextMatrix(fila, 16)
   PUB_PRECIO2 = Val(FORMGEN.grid_fac.TextMatrix(fila, 6))
aut_menor.AddNew
aut_menor!aut_key = PUB_NUM
WS_NUMSEC = WS_NUMSEC + 1
aut_menor!AUT_SECUENCIA = WS_NUMSEC
aut_menor!AUT_CODUSU = LK_CODUSU
aut_menor!AUT_CODCIA = LK_CODCIA
aut_menor!aut_codusu_final = PUB_USUARIO
aut_menor!aut_PRECIO = PUB_PRECIO2
aut_menor!aut_estado = ""
aut_menor!AUT_FECHA = LK_FECHA_DIA
aut_menor!aut_descri = Nulo_Valors(grid_fac.TextMatrix(fila, 5))
aut_menor!aut_codart = PUB_CODART
WS_CANTIDAD = Val(FORMGEN.grid_fac.TextMatrix(fila, 4))
aut_menor!aut_CANTIDAD = WS_CANTIDAD
aut_menor!aut_PRECIO = PUB_PRECIO2
If LK_EMP = "HER" Then
 aut_menor!aut_codclie = PUB_CODVEN
Else
 aut_menor!aut_codclie = PUB_CODCLIE
End If
aut_menor!aut_EQUIV = Val(FORMGEN.grid_fac.TextMatrix(fila, 14)) '
aut_menor!aut_NUMOPER = PUB_NUM_OPER_XXX
aut_menor.Update
OTROMAS2:
   fila = fila + 1
   If fila > grid_fac.Rows - 1 Then
      FLAG = True
   End If

Loop

End Sub

Public Sub IMP_GRID()
Dim wnombre As String * 30
Dim WPRECIO As String * 8
Dim wcantidad As String * 8
Dim wsuttotal As String * 10
Dim wcod As String * 10
Dim P As Integer

Printer.Print "FECHA: " & LK_FECHA_DIA
Printer.Print "USUARIO: " & LK_CODUSU
Printer.Print ""
wnombre = "PRODUCTO"
Printer.Print "CODIGO    "; " "; wnombre; "CANT."; "  "; "  PRECIO TOTAL "
For P = 2 To FORMGEN.grid_fac.Rows - 1
  If Trim(FORMGEN.grid_fac.TextMatrix(P, 1)) <> "" Then
    wcod = Trim(FORMGEN.grid_fac.TextMatrix(P, 1))
    wnombre = Trim(FORMGEN.grid_fac.TextMatrix(P, 0))
    WPRECIO = Format(Trim(FORMGEN.grid_fac.TextMatrix(P, 6)), "##0.00")
    wcantidad = Format(Trim(FORMGEN.grid_fac.TextMatrix(P, 4)), "##0.00")
    wsuttotal = Format(Trim(FORMGEN.grid_fac.TextMatrix(P, 7)), "##0.00")
    Printer.Print wcod; " "; wnombre; " "; wcantidad; " "; Trim(Left(FORMGEN.grid_fac.TextMatrix(P, 5), 8)); " "; WPRECIO; " "; wsuttotal
  End If
Next P
Printer.EndDoc

End Sub

Private Sub transferImport_Click()
On Error GoTo Xverficar
Dim longit As String * 15
Dim IFIL As Integer
Dim WS_PRECIO As Currency
Dim WS_DATOS As String
Dim subtotal As Currency
Dim WW_CODART As Long
Dim ww_cantidad As Currency
Dim ww_codclie As Long

Dim var As String
Dim pos1, CARAC, I, izq, der
Dim TF_TIPMOV



If grid_fac.Rows <> 3 Then Exit Sub
pub_cadena = "SELECT *  FROM facart WHERE  FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? AND FAR_NUMFAC  = ? AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E'  ORDER BY  FAR_FBG, FAR_NUMSER, FAR_NUMFAC,FAR_NUMSEC"
Set PSFAR_TRANS = CN.CreateQuery("", pub_cadena)
If LK_CODCIA = "05" Or LK_CODCIA = "03" Or LK_CODCIA = "30" Then
  PSFAR_TRANS.rdoParameters(0) = "01"
Else
 PSFAR_TRANS.rdoParameters(0) = "03"
End If


PSFAR_TRANS.rdoParameters(1) = i_fbg.Text
PSFAR_TRANS.rdoParameters(2) = i_numser_c.Text
PSFAR_TRANS.rdoParameters(3) = i_numfac_c.Text
Set FAR_TRANS = PSFAR_TRANS.OpenResultset(rdOpenKeyset, rdConcurValues)
FAR_TRANS.Requery
If FAR_TRANS.EOF Then
 MsgBox "Documento no Existe no Esta Anulado ", 48
 Exit Sub
End If
IFIL = 2

Dim rs_ruccli As rdoResultset
If LK_CODCIA = "01" Then
pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_RUC_ESPOSO = '20104174852' AND CLI_CP =  'P' "
Else
pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_RUC_ESPOSO = '20183040406 ' AND CLI_CP =  'P' "
End If
Set rs_ruccli = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not rs_ruccli.EOF Then
  If Val(rs_ruccli!CLI_CODCLIE) <> Val(i_codcli.Text) Then
     MsgBox "No corresponde el Proveedor", 48, Pub_Titulo
     Exit Sub
  End If
End If

det_lot.Rows = 1
det_lot.Clear

Do Until FAR_TRANS.EOF
   If Trim(FAR_TRANS!FAR_ESTADO2) = "L" Then GoTo lotes
   grid_fac.Rows = IFIL + 1
   grid_fac.RowHeight(grid_fac.Rows - 1) = 285
   SQ_OPER = 1
   PUB_KEY = FAR_TRANS!far_codart
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
      MsgBox "Verificar no Existe Codigo , " & FAR_TRANS!far_codart, 48
      Exit Sub
   End If
   grid_fac.TextMatrix(IFIL, 1) = art_LLAVE!art_alterno
   grid_fac.TextMatrix(IFIL, 0) = art_LLAVE!art_nombre
   grid_fac.TextMatrix(IFIL, 16) = FAR_TRANS!far_codart
   grid_fac.TextMatrix(IFIL, 33) = FAR_TRANS!far_codart
   grid_fac.TextMatrix(IFIL, 2) = FAR_TRANS!far_JABAS
   grid_fac.TextMatrix(IFIL, 3) = FAR_TRANS!far_UNIDADES
   longit = Nulo_Valors(FAR_TRANS!far_descri)
   grid_fac.TextMatrix(IFIL, 11) = FAR_TRANS!FAR_COSPRO
   If Nulo_Valor0(FAR_TRANS!FAR_equiv) > 0 Then
   grid_fac.TextMatrix(IFIL, 4) = FAR_TRANS!far_cantidad / Nulo_Valor0(FAR_TRANS!FAR_equiv)
   Else
   grid_fac.TextMatrix(IFIL, 4) = FAR_TRANS!far_cantidad
   End If
   grid_fac.TextMatrix(IFIL, 6) = FAR_TRANS!FAR_PRECIO
   grid_fac.TextMatrix(IFIL, 14) = Nulo_Valor0(FAR_TRANS!FAR_equiv)
   grid_fac.TextMatrix(IFIL, 5) = longit & "       " & FAR_TRANS!FAR_ORDEN_UNIDADES & "       " & FAR_TRANS!FAR_equiv  ' Nulo_Valor0(FAR_TRANS!FAR_FLETE)
   grid_fac.TextMatrix(IFIL, 8) = longit & "       " & FAR_TRANS!FAR_ORDEN_UNIDADES & "       " & FAR_TRANS!FAR_equiv  ' Nulo_Valor0(FAR_TRANS!FAR_FLETE)
   
   grid_fac.TextMatrix(IFIL, 10) = FAR_TRANS!FAR_DESCTO
   
   grid_fac.TextMatrix(IFIL, 11) = FAR_TRANS!FAR_COSPRO
   grid_fac.TextMatrix(IFIL, 12) = FAR_TRANS!far_signo_arm
   grid_fac.TextMatrix(IFIL, 37) = FAR_TRANS!far_PESO
   grid_fac.TextMatrix(IFIL, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(IFIL, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(IFIL, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
lotes:
   det_lot.Rows = det_lot.Rows + 1
   det_lot.TextMatrix(det_lot.Rows - 1, 1) = FAR_TRANS!far_codart ' codigo
   det_lot.TextMatrix(det_lot.Rows - 1, 2) = FAR_TRANS!FAR_CODLOT ' NRO LOTE
   det_lot.TextMatrix(det_lot.Rows - 1, 7) = FAR_TRANS!far_descri   ' DESCRIP UNIDAD
   det_lot.TextMatrix(det_lot.Rows - 1, 5) = 0
   det_lot.TextMatrix(det_lot.Rows - 1, 8) = FAR_TRANS!FAR_equiv  '), "0.0000")   ' SALDO INCIAL
   det_lot.TextMatrix(det_lot.Rows - 1, 3) = FAR_TRANS!far_cantidad_p
   det_lot.TextMatrix(det_lot.Rows - 1, 6) = "01/01/2003"
   det_lot.TextMatrix(det_lot.Rows - 1, 9) = FAR_TRANS!FAR_equiv
   det_lot.TextMatrix(det_lot.Rows - 1, 10) = IFIL
   
   
      
 '  ASIGNA_NEW_LOTE Val(gridlt.TextMatrix(IFIL, 9)), "0.00", Trim(gridlt.TextMatrix(IFIL, 1)), Val(gridlt.TextMatrix(IFIL, 11)), Val(gridlt.TextMatrix(IFIL, 10))
'   CONFIGAR_LOTES LK_CODCIA, Val(grid_fac.TextMatrix(IFIL, 16)), Val(grid_fac.TextMatrix(IFIL, 4)), Left(grid_fac.TextMatrix(IFIL, 5), 10), Val(Right(grid_fac.TextMatrix(IFIL, 5), 8)), IFIL
 

IFIL = IFIL + 1
FAR_TRANS.MoveNext
Loop


calcula_totales

MsgBox "Pase OK."
Exit Sub
'VAR = InputBox("Digite : ACEPTO", "")
'If VAR = "" Then Exit Sub
'If VAR <> "3A" Then Exit Sub
''   GoTo transfer
If LK_CODTRA = 2409 Then
  TF_TIPMOV = 26
ElseIf LK_CODTRA = 2401 Then
  TF_TIPMOV = 27
ElseIf LK_CODTRA = 2403 Then
  TF_TIPMOV = 28
Else
  MsgBox "Proceso Automatica no procede en esta Transación", 48, Pub_Titulo
  Exit Sub
End If
pub_mensaje = "Continuar con el proceso Automatico... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If


   Dim PSFAR_REG As rdoQuery
   Dim FAR_REG As rdoResultset
   LKCHEK = False
   pub_total_2455 = 0

   pub_cadena = "SELECT * FROM facart WHERE FAR_CODCIA = ? AND FAR_TIPMOV = " & TF_TIPMOV & " AND FAR_FECHA = ?   ORDER BY FAR_FBG, FAR_NUMSER, FAR_NUMFAC "
   Set PSFAR_REG = CN.CreateQuery("", pub_cadena)
   PSFAR_REG(0) = 0
   PSFAR_REG(1) = LK_FECHA_DIA
   Set FAR_REG = PSFAR_REG.OpenResultset(rdOpenKeyset, rdConcurValues)
   PSFAR_REG(0) = LK_CODCIA
   PSFAR_REG(1) = LK_FECHA_DIA

   FAR_REG.Requery
   If FAR_REG.EOF Then
      MsgBox "No hay datos para Procesar ..Revisar Procedimientos anteriores", 48, Pub_Titulo
      Exit Sub
   End If
   If FAR_REG!far_transito = "X" Then
      FAR_REG.MoveLast
      If FAR_REG!far_transito = "T" Then
         MsgBox "El sistema detectó que el proceso fue interrumpido ...Se iniciará el proceso complementario ... "
      End If
      FAR_REG.MoveLast
      PUB_NUMFAC = 0
      Do Until FAR_REG.BOF
         If FAR_REG!far_transito = "X" Then
            If PUB_NUMFAC <> 0 And PUB_NUMFAC <> Val(FAR_REG!far_numfac) Then Exit Do
            PU_NUMFAC = FAR_REG!far_numfac
            PU_NUMSER = FAR_REG!far_numser
            pu_codcia = LK_CODCIA
            PU_TIPMOV = 10
            PU_FBG = FAR_REG!FAR_FBG
            SQ_OPER = 1
            LEER_FAR_LLAVE
            If far_llave.EOF = True Then
               FAR_REG.Edit
               FAR_REG!far_transito = "T"
               FAR_REG.Update
               PUB_NUMFAC = FAR_REG!far_numfac
            End If
         End If
         
         FAR_REG.MovePrevious
      Loop
      
      FAR_REG.Requery
      End If
      

   pub_cadena = "SELECT * FROM facart WHERE FAR_CODCIA = ? AND FAR_TIPMOV = " & TF_TIPMOV & " AND FAR_FECHA = ?  AND FAR_TRANSITO= 'T'  ORDER BY FAR_FBG, FAR_NUMSER, FAR_NUMFAC "
   Set PSFAR_REG = CN.CreateQuery("", pub_cadena)
   PSFAR_REG(0) = 0
   PSFAR_REG(1) = LK_FECHA_DIA
   Set FAR_REG = PSFAR_REG.OpenResultset(rdOpenKeyset, rdConcurValues)
   PSFAR_REG(0) = LK_CODCIA
   PSFAR_REG(1) = LK_FECHA_DIA
   FAR_REG.Requery
   If FAR_REG.EOF Then
       MsgBox "No hay datos para Procesar ..Revisar Procedimientos anteriores", 48, Pub_Titulo
      Exit Sub
   End If
   
   PUB_PROCESO = 1
   Do Until FAR_REG.EOF
      GoSub llena_datos
      calcula_totales

      grabar_Click
   Loop
   PUB_PROCESO = 0
   MsgBox "El Ciclo de Transferencia ha sido TERMINADO SATISFACTORIAMENTE.", 48, Pub_Titulo

Exit Sub

llena_datos:

PUB_SUBTOTAL = 0
SUB_JABAS = 0
SUB_CANT = 0
SUB_FLETE = 0
SUB_UNIDAD = 0
FILAX = 0
grid_fac.Clear
pasa_cabeza
SUB_JABAS = 0
SUB_UNIDAD = 0
SUB_CANT = 0
PUB_SUBTOTAL = 0

   
FORMGEN.i_codcli.Text = FAR_REG!FAR_NUMDOC
pu_codclie = FAR_REG!FAR_NUMDOC

pu_cp = "C"
pu_codcia = LK_CODCIA
SQ_OPER = 1
LEER_CLI_LLAVE
If cli_llave.EOF = False Then
   DoEvents
   i_nomCLI.Visible = True
   FORMGEN.i_nomCLI.Caption = Trim(cli_llave(3))
End If

    
'WS_FBG = far_llave!far_fbg
'xxxxxxxx

WS_DATOS = "SI"
WW_CODART = FAR_REG!far_numguia
FILAX = 1
PUB_NUMSER = FAR_REG!far_numser
PUB_NUMFAC = FAR_REG!far_numfac
PUB_FBG = FAR_REG!FAR_FBG

i_serguia.Text = FAR_REG!FAR_NUMSER_C
i_numguia.Text = FAR_REG!FAR_NUMFAC_C
dias_bloqueo = 0
If FAR_REG!FAR_TIPMOV = 27 Then
 If FAR_REG!far_signo_car = 1 Then
    dias_bloqueo = FAR_REG!FAR_DIAS
    i_dias.Text = FAR_REG!FAR_DIAS
    If FAR_REG!FAR_NUM_LOTE = 50 Then
     i_def.ListIndex = 4
    Else
     i_def.ListIndex = 0
    End If
 Else
   If FAR_REG!FAR_NUM_LOTE = 25 Then
     i_def.ListIndex = 1
   ElseIf FAR_REG!FAR_NUM_LOTE = 30 Then
     i_def.ListIndex = 2
   ElseIf FAR_REG!FAR_NUM_LOTE = 40 Then
     i_def.ListIndex = 3
   ElseIf FAR_REG!FAR_NUM_LOTE = 50 Then
     dias_bloqueo = FAR_REG!FAR_DIAS
     i_dias.Text = FAR_REG!FAR_DIAS
     i_def.ListIndex = 4
   End If
 End If
End If
On Error GoTo VERI100
If FAR_REG!FAR_TIPMOV = 26 Then
End If
If FAR_REG!FAR_TIPMOV = 28 Then
   
   i_def.ListIndex = FAR_REG!FAR_NUM_LOTE
End If
On Error GoTo 0
pos1 = InStr(1, FORMGEN.i_def.List(i_def.ListIndex), ".", 1)
pos1 = pos1 - 1
CARAC = Mid(FORMGEN.i_def.List(i_def.ListIndex), 1, pos1)
PUB_SECUENCIA = Val(CARAC)
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_SUT_LLAVE
If SUT_LLAVE.EOF Then
   MsgBox "No existe Definicion en Sub_Transacciones... "
   End
End If
pasa_def




If PUB_FBG = "F" Then
   i_fbg.ListIndex = 0
ElseIf PUB_FBG = "B" Then
   i_fbg.ListIndex = 1
End If



i_cambio.Value = 1
i_cambio.Visible = True
i_numser.Text = PUB_NUMSER
i_numfac.Text = PUB_NUMFAC
DoEvents

If FAR_REG!FAR_MONEDA = "D" Then
   i_ds.ListIndex = 1
ElseIf FAR_REG!FAR_MONEDA = "S" Then
   i_ds.ListIndex = 0
End If


Do Until WS_DATOS = "NO"
   FILAX = FILAX + 1
   grid_fac.Rows = FILAX + 1
   WW_CODART = FAR_REG!far_numguia
   grid_fac.TextMatrix(FILAX, 1) = WW_CODART
   grid_fac.TextMatrix(FILAX, 16) = WW_CODART
   PUB_KEY = WW_CODART
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
      MsgBox "No Existe articulo..=" & WW_CODART
      End
   Else
      grid_fac.TextMatrix(FILAX, 0) = art_LLAVE!art_nombre
   End If
 
   PUB_CODART = WW_CODART
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ARM_LLAVE
  If arm_llave.EOF Then
     MsgBox "Error Grave en arti..."
     End
  End If

   
   
   PUB_PRECIO2 = FAR_REG!FAR_PRECIO
   
   grid_fac.TextMatrix(FILAX, 1) = art_LLAVE!art_alterno
   grid_fac.TextMatrix(FILAX, 4) = FAR_REG!far_cantidad
   grid_fac.TextMatrix(FILAX, 5) = FAR_REG!far_descri
   
   grid_fac.TextMatrix(FILAX, 6) = PUB_PRECIO2
   grid_fac.TextMatrix(FILAX, 7) = FAR_REG!FAR_SUBTOTAL
   grid_fac.TextMatrix(FILAX, 8) = 0
   grid_fac.TextMatrix(FILAX, 9) = 0
   grid_fac.TextMatrix(FILAX, 10) = FAR_REG!FAR_pordescto1
   grid_fac.TextMatrix(FILAX, 11) = ""
   grid_fac.TextMatrix(FILAX, 12) = FAR_REG!far_signo_arm
   grid_fac.TextMatrix(FILAX, 13) = " "
   grid_fac.TextMatrix(FILAX, 19) = 0
   grid_fac.TextMatrix(FILAX, 14) = 1
   grid_fac.TextMatrix(FILAX, 34) = Format(FAR_REG!far_PESO * FAR_REG!far_cantidad, "0.00")
   grid_fac.TextMatrix(FILAX, 37) = Format(FAR_REG!far_PESO, "0.00")
   grid_fac.TextMatrix(FILAX, 16) = WW_CODART
   grid_fac.TextMatrix(FILAX, 33) = WW_CODART
   grid_fac.TextMatrix(FILAX, 11) = arm_llave!arm_cospro
   grid_fac.TextMatrix(FILAX, 11) = arm_llave!arm_cospro
   grid_fac.TextMatrix(FILAX, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
   grid_fac.TextMatrix(FILAX, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
   grid_fac.TextMatrix(FILAX, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)

   i_subtotal.Text = FAR_REG!FAR_BRUTO
   i_impto.Text = FAR_REG!far_IMPTO
   i_codven.Text = FAR_REG!FAR_CODVEN
   WS_BRUTO = i_subtotal.Text
   FAR_REG.Edit
   FAR_REG!far_transito = "X"
   FAR_REG.Update
   FAR_REG.MoveNext
   If FAR_REG.EOF = False Then
   If Val(PUB_NUMSER) = Val(FAR_REG!far_numser) And Val(PUB_NUMFAC) = Val(FAR_REG!far_numfac) And PUB_FBG = FAR_REG!FAR_FBG Then
   Else
      WS_DATOS = "NO"
   End If
   End If
   
   If FAR_REG.EOF = True Then
      WS_DATOS = "NO"
   End If
   
Loop
'pub_flag_cambio = 0
Return

Exit Sub
VERI100:
  MsgBox "EL PARAMETRO DE LA SECUANCIA NO ES VALIDAD . INFORMAR A PRINCIPAL.. Proceso Cancelado", 48, Pub_Titulo
Exit Sub
Xverficar:
Resume Next

Exit Sub



'******************
mover:  ' SOLO PARA SALCAR TODA LA MERCADERIA DE ALMACION CASA - DILVISA
'******************

IFIL = 2


pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODCIA = '" & LK_CODCIA & "' AND ARM_STOCK <> 0 "
Set rs_ruccli = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If rs_ruccli.EOF Then
   MsgBox "NO EXISTE DATOS", 48, Pub_Titulo
   Exit Sub
End If


Do Until rs_ruccli.EOF
   grid_fac.Rows = IFIL + 1
   grid_fac.RowHeight(grid_fac.Rows - 1) = 285
   SQ_OPER = 1
   PUB_KEY = rs_ruccli!ARM_codart
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
      MsgBox "Verificar no Existe Codigo , " & rs_ruccli!ARM_codart, 48
      Exit Sub
   End If
   grid_fac.TextMatrix(IFIL, 1) = art_LLAVE!art_alterno
   grid_fac.TextMatrix(IFIL, 0) = art_LLAVE!art_nombre
   grid_fac.TextMatrix(IFIL, 16) = rs_ruccli!ARM_codart
   grid_fac.TextMatrix(IFIL, 33) = rs_ruccli!ARM_codart
   grid_fac.TextMatrix(IFIL, 2) = 0
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_CODART = rs_ruccli!ARM_codart
   PUB_SECUEN = 0
   LEER_PRE_LLAVE
   longit = Nulo_Valors(pre_llave!pre_UNIDAD)
   grid_fac.TextMatrix(IFIL, 3) = longit
   grid_fac.TextMatrix(IFIL, 11) = rs_ruccli!arm_cospro
   grid_fac.TextMatrix(IFIL, 4) = rs_ruccli!ARM_STOCK

   grid_fac.TextMatrix(IFIL, 6) = rs_ruccli!arm_cospro
   grid_fac.TextMatrix(IFIL, 14) = 1
   grid_fac.TextMatrix(IFIL, 5) = longit & "       " & 0 & "       " & 1 ' Nulo_Valor0(FAR_TRANS!FAR_FLETE)
   grid_fac.TextMatrix(IFIL, 8) = longit & "       " & 0 & "       " & 1 ' Nulo_Valor0(FAR_TRANS!FAR_FLETE)
   
   grid_fac.TextMatrix(IFIL, 10) = 0
   
   grid_fac.TextMatrix(IFIL, 11) = rs_ruccli!arm_cospro
   grid_fac.TextMatrix(IFIL, 12) = -1
   grid_fac.TextMatrix(IFIL, 37) = 0
   grid_fac.TextMatrix(IFIL, 21) = "M"
   grid_fac.TextMatrix(IFIL, 23) = 0
   grid_fac.TextMatrix(IFIL, 24) = 0
IFIL = IFIL + 1
rs_ruccli.MoveNext
Loop
MsgBox "TERMINO"


End Sub

Public Function ANEXO_ACT5() As Boolean
ANEXO_ACT5 = False
If pub_signo_ccm = 0 Then
   ANEXO_ACT5 = True
   Exit Function
End If
If LK_CODTRA = 2748 Then PUB_IMPORTE = PUB_IMPORTE_AMORT
ccm_llave.Edit
ccm_llave!CCM_SALDO = ccm_llave!CCM_SALDO + PUB_IMPORTE * pub_signo_ccm
ccm_llave.Update
' ESTPY QUITANDO NO  SE USA  ---  ws_flag_ccm = 1 'MODIF ' ES LO MISMO
If LK_CODTRA = 231 Then
   che_llave.Edit
   che_llave!che_concepto = PUB_CONCEPTO
   che_llave.Update
   ANEXO_ACT5 = True
End If

If i_codban2.Visible = True Then
    PUB_CONCEPTO = PUB_ABREVIADO & " :" & FORMGEN.i_nomban2.Caption
End If
If LK_CODTRA = 5318 Then
  If pub_signo_ccm = 1 Then
    PUB_CONCEPTO = "ABONO " + Trim(i_concepto.Text)
  Else
    PUB_CONCEPTO = "CARGO " + Trim(i_concepto.Text)
  End If
End If


If (LK_EMP = "3AA" Or LK_EMP = "PIU") And (LK_CODTRA = 5318 Or LK_CODTRA = 2748 Or LK_CODTRA = 2735 Or LK_CODTRA = 5714) Then GoTo DDD
If (LK_CODTRA = 2725) Then GoTo DDD
If PUB_CHENUM = 0 Or LK_CODTRA = 1111 Or PUB_CHESEC <> 0 And LK_CODTRA <> 2741 Then
DDD:
   che_llave.AddNew
   che_llave!CHE_CODBAN = PUB_CODBAN
   'If LK_EMP = "3AA" Then
    che_llave!CHE_CODCIA = par_llave!PAR_CIACCM
   'Else
  '  che_llave!CHE_CODCIA = LK_CODCIA
  ' End If
   che_llave!che_cheser = Val(PUB_CHESER)
   che_llave!che_chesec = PUB_CHESEC
   che_llave!CHE_CHENUM = PUB_CHENUM
   che_llave!CHE_IMPORTE = PUB_IMPORTE
   che_llave!che_abreviado = PUB_ABREVIADO
   che_llave!CHE_FECHA = LK_FECHA_DIA
   If i_fecha_compra.Visible Then
     che_llave!CHE_FECHA_EMISION = i_fecha_compra.Text
     che_llave!CHE_FECHA_COBRO = i_fecha_compra.Text
   Else
     che_llave!CHE_FECHA_EMISION = LK_FECHA_DIA
   End If
   If i_fecha_can.Visible And PUB_TIPMOV <> 10 Then
     che_llave!CHE_FECHA_COBRO = i_fecha_can.Text
   Else
     che_llave!CHE_FECHA_COBRO = LK_FECHA_DIA
   End If
   If IsDate(i_fecha_can.Text) And LK_CODTRA = 2735 Then
     che_llave!CHE_FECHA_COBRO = i_fecha_can.Text
     che_llave!CHE_FECHA_EMISION = i_fecha_can.Text
   'Else
    ' che_llave!CHE_FECHA_COBRO = LK_FECHA_DIA
   End If
   If IsDate(i_fecha_oper.Text) And LK_CODTRA = 5310 Then
     che_llave!CHE_FECHA_COBRO = i_fecha_oper.Text
     che_llave!CHE_FECHA_EMISION = i_fecha_oper.Text
   'Else
    ' che_llave!CHE_FECHA_COBRO = LK_FECHA_DIA
   End If
   che_llave!CHE_CODUSU = LK_CODUSU
   che_llave!che_concepto = PUB_CONCEPTO
   che_llave!CHE_NUMOPER = PUB_NUM_OPER_XXX
   ' CAMTEX
   che_llave!CHE_NUMOPER2 = PUB_NUM_OPER_XXX
   che_llave!che_saldo = ccm_llave!CCM_SALDO
   che_llave!CHE_SIGNO_CCM = pub_signo_ccm
   che_llave!CHE_codtra = LK_CODTRA
   che_llave!che_chenum_ext = PUB_CHENUM_EXT
   che_llave!che_estado = " "
   che_llave.Update
   ' NO DE USA   --ws_flag_che = ingre
 Else
   che_llave.Edit
   che_llave!CHE_IMPORTE = PUB_IMPORTE
   che_llave!CHE_FECHA = LK_FECHA_DIA
   If i_fecha_can.Visible Then
     che_llave!CHE_FECHA_EMISION = i_fecha_can.Text
   Else
     che_llave!CHE_FECHA_EMISION = LK_FECHA_DIA
   End If
   If LK_CODTRA = 2741 And i_fecha_oper.Visible And PUB_TIPMOV <> 10 Then
     che_llave!CHE_FECHA_COBRO = i_fecha_oper.Text
     che_llave!CHE_FECHA_EMISION = i_fecha_oper.Text
     GoTo BIEN_FEC
   End If
   
   If i_fecha_can.Visible And PUB_TIPMOV <> 10 Then
     che_llave!CHE_FECHA_COBRO = i_fecha_can.Text
   Else
     che_llave!CHE_FECHA_COBRO = LK_FECHA_DIA
   End If
   If IsDate(i_fecha_can.Text) And LK_CODTRA = 2735 Then
     che_llave!CHE_FECHA_COBRO = i_fecha_can.Text
     che_llave!CHE_FECHA_EMISION = i_fecha_can.Text
   End If
BIEN_FEC:
   che_llave!CHE_CODUSU = LK_CODUSU
   If che_llave!CHE_FECHA_COBRO <> che_llave!CHE_FECHA_EMISION Then
       che_llave!che_estado = "T"
   End If
   ' CAMTEX
   If PUB_IMPORTE = 0 Then
     che_llave!che_concepto = "ANULACION DE CHEQUE"
     che_llave!che_estado = " "
   Else
     che_llave!che_concepto = PUB_CONCEPTO
   End If
   che_llave!che_abreviado = PUB_ABREVIADO
   che_llave!che_abreviado = PUB_ABREVIADO
   che_llave!CHE_NUMOPER = PUB_NUM_OPER_XXX
   ' CAMTEX
   che_llave!CHE_NUMOPER2 = PUB_NUM_OPER_XXX
   che_llave!che_saldo = ccm_llave!CCM_SALDO
   che_llave!CHE_SIGNO_CCM = pub_signo_ccm
   che_llave.Update
   ' NO SE USA   -    ws_flag_che = MODIF
End If

End Function


Public Sub LLENA_NUM_DIARIO()
If LK_CODTRA = 5350 Then
   PSALL_DIARIO(0) = LK_CODCIA
   PSALL_DIARIO(1) = LK_FECHA_DIA
   PSALL_DIARIO(2) = LK_CODTRA
   all_diario.Requery
   If all_diario.EOF Then
      i_numfac.Text = 1
   Else
      i_numfac.Text = Val(Nulo_Valor0(all_diario!ALL_NUM_RECIBO)) + 1
   End If
End If
End Sub

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not IsDate(txtCampo1.Text) Then
      MsgBox "Fecha No Procede", 48, Pub_Titulo
      Exit Sub
    End If
    Azul2 txtCampo2, txtCampo2
End If
End Sub

Private Sub txtCampo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not IsDate(txtCampo1.Text) Then
      MsgBox "Fecha No Procede", 48, Pub_Titulo
      Exit Sub
    End If
    Azul txtCodVen, txtCodVen
End If

End Sub

Private Sub txtCodVen_Change()
If txtCodVen.Text = "" Then
  i_nombreven.Caption = ""
  If Grid_detalle.Visible And Grid_detalle.Rows > 2 Then Grid_detalle.Rows = 2
End If
End Sub

Private Sub txtCodVen_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.

If Not LV_VEN.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txtCodVen.Text = "" Then
  loc_key = 1
  Set LV_VEN.SelectedItem = LV_VEN.ListItems(loc_key)
  LV_VEN.ListItems.Item(loc_key).Selected = True
  LV_VEN.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > LV_VEN.ListItems.count Then loc_key = LV_VEN.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > LV_VEN.ListItems.count Then loc_key = LV_VEN.ListItems.count
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
  LV_VEN.ListItems.Item(loc_key).Selected = True
  LV_VEN.ListItems.Item(loc_key).EnsureVisible
  txtCodVen.Text = Trim(LV_VEN.ListItems.Item(loc_key).Text) & " "
  DoEvents
  txtCodVen.SelStart = Len(i_codven.Text)
  DoEvents
fin:


End Sub

Private Sub txtCodVen_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.

If KeyAscii = 27 Then
  txtCodVen.Text = ""
  LV_VEN.Visible = False
  Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo OJO
PUB_CODVEN = Val(FORMGEN.txtCodVen.Text)
On Error GoTo 0
If Len(txtCodVen.Text) = 0 Then
   Exit Sub
End If
If PUB_CODVEN <> 0 And IsNumeric(txtCodVen.Text) = True Then
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_VEN_LLAVE
   If ven_llave.EOF Then
    Azul FORMGEN.txtCodVen, FORMGEN.txtCodVen
    MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
    FORMGEN.i_codven.SetFocus
    GoTo fin
   Else
    FORMGEN.i_nombreven.Caption = ven_llave(2)
   End If
Else
On Error GoTo sigue
   If loc_key > 0 Then VALOR = UCase(LV_VEN.ListItems.Item(loc_key).Text)
   If Trim(UCase(txtCodVen.Text)) = Left(VALOR, Len(Trim(txtCodVen.Text))) Then
   Else
      Exit Sub
   End If
   
   If loc_key = 0 Then Exit Sub
   
   txtCodVen.Text = Trim(LV_VEN.ListItems.Item(loc_key).SubItems(1))
   PUB_CODVEN = Val(FORMGEN.txtCodVen.Text)
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   LEER_VEN_LLAVE
   FORMGEN.i_nombreven.Caption = ven_llave(2)
End If
LV_VEN.Visible = False

fin:
sigue:
OJO:
End Sub

Private Sub txtCodVen_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
Dim pos1, CARAC, I, izq, der

If KeyCode = 45 Then
 
End If

If Len(txtCodVen.Text) = 0 Or IsNumeric(txtCodVen.Text) Then
   LV_VEN.Visible = False
   Exit Sub
End If
If LV_VEN.Visible = False Or Len(txtCodVen.Text) = 1 Then
    loc_key = 0
    var = Asc(txtCodVen.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 2
    archi = "SELECT VEM_CODVEN , VEM_CODCIA, VEM_NOMBRE  FROM VEMAEST WHERE  VEM_CODCIA = '" & LK_CODCIA & "' AND VEM_NOMBRE BETWEEN '" & i_codven.Text & "' AND  '" & var & "' ORDER BY VEM_NOMBRE"
    PROC_LISVIEW LV_VEN
    loc_key = 0
    If LV_VEN.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If LV_VEN.Visible Then
  Set itmFound = LV_VEN.FindItem(LTrim(i_codcli.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > LV_VEN.ListItems.count Then
      LV_VEN.ListItems.Item(LV_VEN.ListItems.count).EnsureVisible
   Else
     LV_VEN.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If

Exit Sub

Return



End Sub

Private Sub txtNumDocIniBol_Change()
Dim WsNUM As String
If Grid_detalle.Rows - 2 <= 0 Then Exit Sub
WsNUM = Format(Val(FORMGEN.txtNumDocIniBol.Text) + ver_nro_pedidos("B"), "######")
FORMGEN.txtNumDocFinBol.Text = WsNUM
FORMGEN.txtSerieDocBolFin.Text = FORMGEN.txtSerieDocIniBol.Text

End Sub

Private Sub txtNumDocIniBol_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
   ' Azul2 FechaIniBol, FechaIniBol
End If

End Sub

Private Sub txtNumDocIniFac_Change()
Dim WsNUM As String
If Grid_detalle.Rows - 2 <= 0 Then Exit Sub
WsNUM = Format(Val(FORMGEN.txtNumDocIniFac.Text) + ver_nro_pedidos("F"), "######")
FORMGEN.txtNumDocFacFin.Text = WsNUM
FORMGEN.txtSerieDocFacFin.Text = FORMGEN.txtSerieDocIniFac.Text

End Sub

Public Function ver_nro_pedidos(WFBG As String) As Integer
Dim C As Integer
Dim CONTA As Integer
ver_nro_pedidos = 0
CONTA = 0
If Grid_detalle.Rows - 1 <= 2 Then Exit Function
For C = 2 To Grid_detalle.Rows - 1
  If WFBG = UCase(Left(Grid_detalle.TextMatrix(C, 5), 1)) Then
  CONTA = CONTA + 1
  End If
Next
If CONTA <> 0 Then CONTA = CONTA - 1
ver_nro_pedidos = CONTA
End Function

Private Sub txtNumDocIniFac_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  ' Azul2 FechaIniFac, FechaIniFac
End If

End Sub

Private Sub txtSerieDocIniBol_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  Azul txtNumDocIniBol, txtNumDocIniBol
End If

End Sub

Private Sub txtSerieDocIniFac_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  Azul txtNumDocIniFac, txtNumDocIniFac
End If
End Sub

Public Sub ANEXO_ACT16()
'VIENE DE CONSIS_21 EXTORNOS
far_menor3.MoveFirst
If ww_codtra_ext = 2210 Or ww_codtra_ext = 2212 Or ww_codtra_ext = 2103 Then
   lee_ped_llave
End If
If ww_codtra_ext = 1402 Then
   Do Until far_menor3.EOF
      far_menor3.Edit
      far_menor3.Delete
      far_menor3.MoveNext
   Loop
   Exit Sub
End If

Do Until far_menor3.EOF
   PUB_NUMSER_C = far_menor3!FAR_NUMSER_C
   PUB_NUMFAC_C = far_menor3!FAR_NUMFAC_C
   PUB_CODVEN = far_menor3!FAR_CODVEN
   PUB_NUMGUIA = far_menor3!far_numguia
   PUB_SERGUIA = far_menor3!far_serguia
   If LK_EMP <> "HER" Then
     PUB_ISLA = far_menor3!far_ISLA
     PUB_TURNO = far_menor3!far_turno
   End If
   
   i_fecha_compra.Text = Format(far_menor3!FAR_fecha_compra, "dd/mm/yy")
   ' AGREGE ACV
   pub_dias = far_menor3!FAR_DIAS
   If LK_CODTRA = 1111 And (ww_codtra_ext = 2105 Or ww_codtra_ext = 2107) Then
         SQ_OPER = 1
         PU_TIPMOV = 10
         'AGREGE ACV ESTE IF
'         If ww_codtra_ext = 2105 Then
'          PU_NUMSER = far_menor3!far_NUMSER
'          PU_NUMFAC = far_menor3!far_NUMFAC
'         Else
          PU_NUMSER = far_menor3!FAR_PEDSER
          PU_NUMFAC = far_menor3!FAR_PEDFAC
'         End If
         pu_codcia = LK_CODCIA
         PU_FBG = far_menor3!far_fbg2
         LEER_FAR_LLAVE
         If far_llave.EOF Then
            pub_mensaje = "!!! Operacion Invalida..."
            exito = False
            Exit Sub
            'GoTo Error_fatal
         End If
            
         'Do Until far_llave.EOF
         '   If far_llave!FAR_NUMSEC = far_menor3!far_pedsec Then
         '      far_llave.Edit
         '      far_llave!FAR_CONSIG = far_llave!FAR_CONSIG - far_menor3!far_cantidad
         '      far_llave!far_concepto = " "
         '      far_llave.Update
         '      Exit Do
         '   End If
         '   far_llave.MoveNext
         'Loop
      far_menor3.Delete
   Else
    far_menor3.Edit
    far_menor3!far_estado = "E"
    far_menor3!FAR_ESTADO2 = "E"
    If LK_CODTRA = 1111 And LK_FECHA_DIA <> i_fecha_oper.Text Then
       far_menor3!FAR_ESTADO2 = "A"
    End If
    far_menor3.Update
  End If
  If ww_codtra_ext = 2212 Then ped_2212
  If ww_codtra_ext = 2210 Then ped_2210

   far_menor3.MoveNext
Loop
If ww_codtra_ext = 2103 Then elim_ped
far_llave.Requery

End Sub

Public Sub ANULAR_PED()
Dim cat As Currency

   If LOC_PEDIDO = -1 Then Exit Sub
   SQ_OPER = 1
   PUB_TIPMOV = 201
   pu_codcia = LK_CODCIA
   PUB_PEDSER = 100
   PUB_PEDFAC = LOC_PEDIDO
   LEER_PED_LLAVE
   Do Until ped_llave.EOF
       SQ_OPER = 1
       PUB_CODART = ped_llave!PED_CODART
       pu_codcia = LK_CODCIA
       LEER_ARM_LLAVE
       cat = Val(ped_llave!PED_cantidad) * Val(ped_llave!PED_EQUIV)
       arm_llave.Edit
       arm_llave!ARM_STOCK = arm_llave!ARM_STOCK + cat
       arm_llave!ARM_INGRESOS = arm_llave!ARM_INGRESOS + cat
       arm_llave.Update
       ped_llave.MoveNext
   Loop
   If PUB_PEDFAC <> 0 Then
      pub_signo_arm = -1
      pub_cadena = "UPDATE PEDIDOS SET PED_SITUACION = 'E' WHERE PED_TIPMOV = 201 AND PED_CODCIA = '" & LK_CODCIA & "' AND PED_NUMSER = " & PUB_PEDSER & " AND PED_NUMFAC = " & PUB_PEDFAC
      CN.Execute pub_cadena, rdExecDirect
   End If
End Sub

Public Sub CONFIGAR_LOTES(LT_CODCIA As String, LT_CODART As Currency, LT_CANTIDAD As Currency, LT_DESCRIP As String, LT_EQUIV As Integer, FILA_GRID As Integer, Optional NOMUESTRA)

Dim xcta As Integer
Dim XCTA2 As Integer
Dim JALA_LOTE As rdoResultset
Dim LT_SALDO As Currency
Dim flag_lotes As String * 1
Dim WCANTIDAD_INI  As Currency
WCANTIDAD_INI = LT_CANTIDAD

pub_cadena = "SELECT * FROM LOTE WHERE LOT_CODCIA = '" & LT_CODCIA & "' AND LOT_CODART = " & LT_CODART & " AND LOT_SALDOS <> 0  ORDER BY LOT_FECHA_VCTO"
Set JALA_LOTE = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
LT_SALDO = LT_CANTIDAD
gridlt.Clear
gridlt.Rows = 2
gridlt.Cols = 13
gridlt.ColWidth(0) = 1200  'NRO DE LOTE
gridlt.ColWidth(1) = 800  '  DESCRIP DE EQUIV
gridlt.ColWidth(2) = 1100
gridlt.ColWidth(3) = 1100
gridlt.ColWidth(4) = 1200
gridlt.ColWidth(5) = 0
gridlt.ColWidth(6) = 0
gridlt.ColWidth(7) = 0
gridlt.ColWidth(8) = 0
gridlt.ColWidth(9) = 0 ' codigo de arti
gridlt.ColWidth(10) = 0 ' FILA DE GRID_FAC
gridlt.ColWidth(11) = 0 ' EQUIV
gridlt.ColWidth(12) = 0 ' Nro de Lote original
gridlt.TextMatrix(0, 0) = "Nº LOTE"
gridlt.TextMatrix(0, 1) = "UNIDAD"
gridlt.TextMatrix(0, 2) = "SALDO"
gridlt.TextMatrix(0, 3) = "SOLICITA"
gridlt.TextMatrix(0, 4) = "FECHAVCTO"
gridlt.TextMatrix(0, 10) = "fila del grid_fac"
' chequeo existe datos en el detalle para mostrar
'------------------------------------------------
otra_vez:
gridlt.Rows = 2
flag_lotes = ""
For xcta = 1 To det_lot.Rows - 1
 If Val(det_lot.TextMatrix(xcta, 1)) = LT_CODART And FILA_GRID = Val(det_lot.TextMatrix(xcta, 4)) Then
   flag_lotes = "A"
   gridlt.Rows = gridlt.Rows + 1
   gridlt.TextMatrix(gridlt.Rows - 1, 0) = det_lot.TextMatrix(xcta, 2) ' NRO LOTE
   gridlt.TextMatrix(gridlt.Rows - 1, 1) = det_lot.TextMatrix(xcta, 7)  ' DESCRIP UNIDAD
   gridlt.TextMatrix(gridlt.Rows - 1, 2) = Format(Val(det_lot.TextMatrix(xcta, 5)) / Val(det_lot.TextMatrix(xcta, 8)), "0.0000")   ' SALDO INCIAL
   gridlt.TextMatrix(gridlt.Rows - 1, 3) = Format(Val(det_lot.TextMatrix(xcta, 3)) / Val(det_lot.TextMatrix(xcta, 8)), "0.0000")  'CANTIDAD PARA MOVER
   gridlt.TextMatrix(gridlt.Rows - 1, 4) = det_lot.TextMatrix(xcta, 6) ' FECHA VCTO
   gridlt.TextMatrix(gridlt.Rows - 1, 9) = LT_CODART
   gridlt.TextMatrix(gridlt.Rows - 1, 10) = FILA_GRID
   gridlt.TextMatrix(gridlt.Rows - 1, 11) = det_lot.TextMatrix(xcta, 8)  ' equiv de unidad
   gridlt.TextMatrix(gridlt.Rows - 1, 12) = det_lot.TextMatrix(xcta, 2)  ' equiv de unidad
 End If

Next xcta
If flag_lotes = "A" Then
 GoTo mues
Else
 If JALA_LOTE.EOF Then
   ASIGNA_NEW_LOTE LT_CODART, LT_CANTIDAD, LT_DESCRIP, LT_EQUIV, FILA_GRID
   GoTo otra_vez
 End If
End If
xcta = 1
If Not JALA_LOTE.EOF Then
  LT_CANTIDAD = LT_CANTIDAD * LT_EQUIV
End If
Do Until JALA_LOTE.EOF
   If LT_CANTIDAD <= 0 Then
       LT_SALDO = 0
       LT_CANTIDAD = 0
       GoSub muestra
       GoTo OTRO
   End If
   If LT_CANTIDAD <= JALA_LOTE!lot_saldos Then
       LT_SALDO = LT_CANTIDAD
       LT_CANTIDAD = 0
       GoSub muestra
   Else
      If JALA_LOTE.RowCount = JALA_LOTE.AbsolutePosition Then
          LT_SALDO = LT_CANTIDAD  'JALA_LOTE!LOT_SALDOS
          LT_CANTIDAD = 0
      Else
          LT_SALDO = JALA_LOTE!lot_saldos
        ' LT_CANTIDAD = 0
       End If
       GoSub muestra
       LT_CANTIDAD = LT_CANTIDAD - JALA_LOTE!lot_saldos
   End If
OTRO:
  JALA_LOTE.MoveNext
Loop

mues:
PUB_KEY = LT_CODART
pu_codcia = LT_CODCIA
SQ_OPER = 1
LEER_ART_LLAVE
If Not art_LLAVE.EOF Then
  lblltcod.Caption = art_LLAVE!art_alterno
  lblltnom.Caption = Trim(art_LLAVE!art_nombre)
End If
lblltcantidad.Caption = Format(WCANTIDAD_INI, "0.000")
lblltunidad.Caption = Trim(LT_DESCRIP) & String(80, "  ") & LT_EQUIV
If Not IsMissing(NOMUESTRA) Then
 If NOMUESTRA = 1 Then Exit Sub
End If

fralotes.Visible = True
If gridlt.Rows > 2 Then
 gridlt.Row = 2
 gridlt.COL = 1
End If
gridlt.SetFocus
Exit Sub
muestra:
 xcta = xcta + 1
 gridlt.Rows = gridlt.Rows + 1
 If Val(LT_SALDO) <> 0 Then
 gridlt.TextMatrix(xcta, 0) = "X"
 Else
 gridlt.TextMatrix(xcta, 0) = ""
 End If
 gridlt.TextMatrix(xcta, 0) = Trim(JALA_LOTE!lot_nrolote)
 gridlt.TextMatrix(xcta, 1) = Trim(LT_DESCRIP)
 gridlt.TextMatrix(xcta, 2) = Format(Val(JALA_LOTE!lot_saldos) / Val(Right(LT_EQUIV, 8)), "0.0000")
 gridlt.TextMatrix(xcta, 3) = Format(Val(LT_SALDO) / Val(Right(LT_EQUIV, 8)), "0.0000")
 gridlt.TextMatrix(xcta, 4) = Format(JALA_LOTE!lot_fecha_vcto, "dd/mm/yy")
 gridlt.TextMatrix(xcta, 9) = JALA_LOTE!lot_codart
 gridlt.TextMatrix(xcta, 10) = FILA_GRID
 gridlt.TextMatrix(xcta, 11) = LT_EQUIV
 gridlt.TextMatrix(xcta, 12) = Trim(JALA_LOTE!lot_nrolote)

Return
End Sub

Public Sub TODO_LOTES()
Dim JALA_LOTE As rdoResultset
Dim LT_SALDO As Currency
Dim xcuenta As Integer
Dim xcuenta2 As Integer
Dim LT_CANTIDAD As Currency
Dim pasa_act As String * 1

det_lot.ColWidth(0) = 0 ' codigo de cia
det_lot.ColWidth(1) = 400 ' codigo Interno  de Articulo
det_lot.ColWidth(2) = 400 ' codigo nro de lote
det_lot.ColWidth(3) = 400 ' cantidad
det_lot.ColWidth(4) = 200 ' fila del grid_fac
det_lot.ColWidth(5) = 200 ' saldo de lote
For xcuenta = 2 To grid_fac.Rows - 1
  PUB_CODART = Val(grid_fac.TextMatrix(xcuenta, 16))
  If PUB_CODART = 0 Then Exit For
  LT_CANTIDAD = Val(grid_fac.TextMatrix(xcuenta, 4))
  CONFIGAR_LOTES LK_CODCIA, Val(grid_fac.TextMatrix(xcuenta, 16)), LT_CANTIDAD, grid_fac.TextMatrix(xcuenta, 5), Val(grid_fac.TextMatrix(xcuenta, 14)), xcuenta, 1
  wlot_falso_flag = ""
  cmdltaceptar_Click
  If wlot_falso_flag = "1" Then Exit For
'  pasa_act = ""
'  GoSub chequeo_lista ' chequeo que si esta en la lista
'  If Trim(pasa_act) <> "A" Then GoTo sigue_for
'  pub_cadena = "SELECT * FROM LOTE WHERE LOT_CODCIA = '" & LK_CODCIA & "' AND LOT_CODART = " & PUB_CODART & " AND LOT_SALDOS <> 0  ORDER BY LOT_FECHA_VCTO"
'  Set JALA_LOTE = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
'  LT_CANTIDAD = Val(grid_fac.TextMatrix(xcuenta, 4))
'  LT_SALDO = 0
'  Do Until JALA_LOTE.EOF
'    If LT_CANTIDAD <= 0 Then Exit Do
'    If LT_CANTIDAD <= JALA_LOTE!LOt_SALDOS Then
'       LT_SALDO = LT_CANTIDAD
'       LT_CANTIDAD = 0
'       det_lot.Rows = det_lot.Rows + 1
'       det_lot.TextMatrix(det_lot.Rows - 1, 0) = LK_CODCIA
'       det_lot.TextMatrix(det_lot.Rows - 1, 1) = PUB_CODART
'       det_lot.TextMatrix(det_lot.Rows - 1, 2) = Trim(JALA_LOTE!LOT_NROLOTE)
'       det_lot.TextMatrix(det_lot.Rows - 1, 3) = LT_SALDO
'       det_lot.TextMatrix(det_lot.Rows - 1, 4) = xcuenta
'       det_lot.TextMatrix(det_lot.Rows - 1, 5) = JALA_LOTE!LOt_SALDOS
'       det_lot.TextMatrix(det_lot.Rows - 1, 6) = JALA_LOTE!lot_fecha_vcto
'    Else
'       LT_SALDO = JALA_LOTE!LOt_SALDOS
'       det_lot.Rows = det_lot.Rows + 1
'       det_lot.TextMatrix(det_lot.Rows - 1, 0) = LK_CODCIA
'       det_lot.TextMatrix(det_lot.Rows - 1, 1) = PUB_CODART
'       det_lot.TextMatrix(det_lot.Rows - 1, 2) = Trim(JALA_LOTE!LOT_NROLOTE)
'       det_lot.TextMatrix(det_lot.Rows - 1, 3) = LT_SALDO
'       det_lot.TextMatrix(det_lot.Rows - 1, 4) = xcuenta
'       det_lot.TextMatrix(det_lot.Rows - 1, 5) = JALA_LOTE!LOt_SALDOS
'       det_lot.TextMatrix(det_lot.Rows - 1, 6) = JALA_LOTE!lot_fecha_vcto
'       LT_CANTIDAD = LT_CANTIDAD - JALA_LOTE!LOt_SALDOS
'    End If
'    JALA_LOTE.MoveNext
'  Loop
sigue_for:

Next xcuenta

Exit Sub
' chequeo Lista
chequeo_lista:
pasa_act = "A"
For xcuenta2 = 1 To det_lot.Rows - 1
   If Val(det_lot.TextMatrix(xcuenta2, 1)) = PUB_CODART And Val(det_lot.TextMatrix(xcuenta2, 4)) = xcuenta Then
      pasa_act = ""
      Exit For
   End If
Next xcuenta2

Return
End Sub


Public Sub verif_lotes(lot_codart As Currency, lot_fila As Integer)
Dim xcuenta As Integer

For xcuenta = 1 To 100
 ParaLot_codlot(xcuenta) = ""
 ParaLot_lotcant(xcuenta) = 0
 ParaLot_fechalot(xcuenta) = 0
Next xcuenta
ParaLot_count = 0
For xcuenta = 1 To det_lot.Rows - 1
 If lot_codart = Val(det_lot.TextMatrix(xcuenta, 1)) And lot_fila = Val(det_lot.TextMatrix(xcuenta, 4)) Then
    If Val(det_lot.TextMatrix(xcuenta, 3)) <> 0 Then
     ParaLot_count = ParaLot_count + 1
     ParaLot_codlot(ParaLot_count) = Trim(det_lot.TextMatrix(xcuenta, 2))
     ParaLot_lotcant(ParaLot_count) = Format(det_lot.TextMatrix(xcuenta, 3), "0.00")
     ParaLot_fechalot(ParaLot_count) = det_lot.TextMatrix(xcuenta, 6)
    End If
 End If
Next xcuenta

End Sub

Public Sub ASIGNA_NEW_LOTE(LT_CODART As Currency, LT_CANTIDAD As Currency, LT_DESCRIP As String, LT_EQUIV As Integer, FILA_GRID As Integer)
det_lot.Rows = det_lot.Rows + 1
det_lot.TextMatrix(det_lot.Rows - 1, 0) = LK_CODCIA
det_lot.TextMatrix(det_lot.Rows - 1, 1) = LT_CODART
det_lot.TextMatrix(det_lot.Rows - 1, 2) = "(*)" 'Format(LT_CODART, "0") & Format(LK_FECHA_DIA, "ddmm") & Format(Now, "HHMMSS")
det_lot.TextMatrix(det_lot.Rows - 1, 3) = LT_CANTIDAD * LT_EQUIV
det_lot.TextMatrix(det_lot.Rows - 1, 4) = FILA_GRID
det_lot.TextMatrix(det_lot.Rows - 1, 5) = "0.00"
det_lot.TextMatrix(det_lot.Rows - 1, 6) = LK_FECHA_DIA
det_lot.TextMatrix(det_lot.Rows - 1, 7) = LT_DESCRIP
det_lot.TextMatrix(det_lot.Rows - 1, 8) = LT_EQUIV
det_lot.TextMatrix(det_lot.Rows - 1, 9) = det_lot.TextMatrix(det_lot.Rows - 1, 2)
End Sub

Public Sub PASA_BOT_ACEPTAR()
' PASA DE LA MUESTRA AL DETALLATE AL PULSAR EL BOTON ACEPTAR
Dim JALA_LOTE As rdoResultset
Dim LT_SALDO As Currency
Dim xcuenta As Integer
Dim xcuenta2 As Integer
Dim LT_CANTIDAD As Currency
Dim pasa_act As String * 1
Dim FILA_GRID As Integer
Dim NRO_LOTE  As String
Dim wencuentra As String * 1
For xcuenta = 2 To gridlt.Rows - 1
    PUB_CODART = Val(gridlt.TextMatrix(xcuenta, 9))
    FILA_GRID = Val(gridlt.TextMatrix(xcuenta, 10))
    NRO_LOTE = Trim(gridlt.TextMatrix(xcuenta, 12))
    If PUB_CODART = 0 Then Exit For
    'If gridlt.TextMatrix(xcuenta, 3) = 0 Then GoTo sigue_for
    wencuentra = ""
    For xcuenta2 = 1 To det_lot.Rows - 1
      If PUB_CODART = Val(det_lot.TextMatrix(xcuenta2, 1)) And FILA_GRID = Val(det_lot.TextMatrix(xcuenta2, 4)) And NRO_LOTE = Trim(det_lot.TextMatrix(xcuenta2, 9)) Then
        NRO_LOTE = Trim(gridlt.TextMatrix(xcuenta, 0))
        gridlt.TextMatrix(xcuenta, 12) = NRO_LOTE
        det_lot.TextMatrix(xcuenta2, 2) = NRO_LOTE
        det_lot.TextMatrix(xcuenta2, 9) = NRO_LOTE
        det_lot.TextMatrix(xcuenta2, 3) = Val(gridlt.TextMatrix(xcuenta, 3)) * Val(gridlt.TextMatrix(xcuenta, 11))
        det_lot.TextMatrix(xcuenta2, 6) = Trim(gridlt.TextMatrix(xcuenta, 4))
        wencuentra = "A"
        Exit For
      End If
    Next xcuenta2
    If wencuentra <> "A" Then
       det_lot.Rows = det_lot.Rows + 1
       det_lot.TextMatrix(det_lot.Rows - 1, 0) = LK_CODCIA
       det_lot.TextMatrix(det_lot.Rows - 1, 1) = PUB_CODART
       det_lot.TextMatrix(det_lot.Rows - 1, 2) = NRO_LOTE
       det_lot.TextMatrix(det_lot.Rows - 1, 3) = Val(gridlt.TextMatrix(xcuenta, 3)) * Val(gridlt.TextMatrix(xcuenta, 11))
       det_lot.TextMatrix(det_lot.Rows - 1, 4) = FILA_GRID
       det_lot.TextMatrix(det_lot.Rows - 1, 5) = Val(gridlt.TextMatrix(xcuenta, 2)) * Val(gridlt.TextMatrix(xcuenta, 11))
       det_lot.TextMatrix(det_lot.Rows - 1, 6) = Trim(gridlt.TextMatrix(xcuenta, 4))
       det_lot.TextMatrix(det_lot.Rows - 1, 7) = Trim(gridlt.TextMatrix(xcuenta, 1))
       det_lot.TextMatrix(det_lot.Rows - 1, 8) = Trim(gridlt.TextMatrix(xcuenta, 11))
       det_lot.TextMatrix(det_lot.Rows - 1, 9) = NRO_LOTE
    End If
sigue_for:
Next xcuenta

End Sub
Public Sub BORRA_ITEM_LOTE(LT_CODART As Currency, FILA_GRID As Integer, Optional WNROLOTE)
Dim ix As Integer
If Not IsMissing(WNROLOTE) Then
 For ix = 1 To det_lot.Rows - 1
   If WNROLOTE = det_lot.TextMatrix(ix, 9) And LT_CODART = Val(det_lot.TextMatrix(ix, 1)) And FILA_GRID = Val(det_lot.TextMatrix(ix, 4)) Then
     If det_lot.Rows <= 2 Then
       det_lot.Rows = 1
       Exit For
     Else
       det_lot.RemoveItem (ix)
       ix = ix - 1
       Exit For
     End If
   End If
 Next ix
 Exit Sub
End If

On Error GoTo SALE_CODART
For ix = 1 To det_lot.Rows - 1
  If LT_CODART = Val(det_lot.TextMatrix(ix, 1)) And FILA_GRID = Val(det_lot.TextMatrix(ix, 4)) Then
    If det_lot.Rows <= 2 Then
      det_lot.Rows = 1
      Exit For
    Else
      det_lot.RemoveItem (ix)
      ix = ix - 1
    End If
  End If
Next ix
SALE_CODART:
End Sub
Public Function CHEQUEO_DETALLE() As String
' CHEQUEO PARA VER SI ESTAN TODOS LO ARTICULOS EDITAR SOLO EN INGRESOS

Dim xcuenta As Integer
Dim xcuenta2 As Integer
Dim pasa_act As String
Dim FILA_GRID As Integer
Dim NRO_LOTE  As String
Dim wencuentra As String * 1
CHEQUEO_DETALLE = ""
wencuentra = "A"
pasa_act = ""
For xcuenta = 2 To grid_fac.Rows - 1
    PUB_CODART = Val(grid_fac.TextMatrix(xcuenta, 16))
    FILA_GRID = xcuenta
    If PUB_CODART = 0 Then
     Exit For
    End If
    'If gridlt.TextMatrix(xcuenta, 3) = 0 Then GoTo sigue_for
    For xcuenta2 = 1 To det_lot.Rows - 1
      If PUB_CODART = Val(det_lot.TextMatrix(xcuenta2, 1)) And FILA_GRID = Val(det_lot.TextMatrix(xcuenta2, 4)) Then
        wencuentra = "A"
        pasa_act = ""
        Exit For
      Else
        wencuentra = ""
        pasa_act = "Fila:" & Format(xcuenta - 1, "00") & " " & Trim(grid_fac.TextMatrix(xcuenta, 0))
      End If
    Next xcuenta2
    If det_lot.Rows <= 1 Then
       wencuentra = ""
       pasa_act = "Todos."
    End If
Next xcuenta
If wencuentra <> "A" Then
  CHEQUEO_DETALLE = pasa_act
End If

End Function



Private Function ARTICULOLOTE() As String
Dim I As Integer
Dim TCOD As Currency
Dim C As Integer
Dim tnombre As String * 40
Dim cda_lotes As Integer
ARTICULOLOTE = "Productos con mas de un codigo de lote: " & Chr(13) & "    Producto                                  Cantidad   " & Chr(13)

TCOD = Val(det_lot.TextMatrix(1, 1))
C = 0
cda_lotes = 0
For I = 1 To FORMGEN.det_lot.Rows - 1
 If Val(det_lot.TextMatrix(I, 3)) <> 0 Then
   cda_lotes = cda_lotes + 1
 End If
 
 If TCOD <> Val(det_lot.TextMatrix(I, 1)) Then
   ' c = c + 1
   If C > 1 Then
    PUB_KEY = TCOD
    pu_codcia = LK_CODCIA
    SQ_OPER = 1
    LEER_ART_LLAVE
    If Not art_LLAVE.EOF Then
      tnombre = Trim(art_LLAVE!art_nombre)
      ARTICULOLOTE = ARTICULOLOTE & "- " & tnombre & " : " & C & Chr(13)
    Else
      ARTICULOLOTE = ARTICULOLOTE & TCOD & Chr(13)
    End If
   End If
   TCOD = Val(det_lot.TextMatrix(I, 1))
   C = 1
 Else
  If Val(det_lot.TextMatrix(I, 3)) = 0 Then
  Else
    C = C + 1
  End If
 End If
siguecontando:
Next I
If C > 1 Then
 PUB_KEY = TCOD
 pu_codcia = LK_CODCIA
 SQ_OPER = 1
 LEER_ART_LLAVE
 If Not art_LLAVE.EOF Then
 '   ARTICULOLOTE = ARTICULOLOTE & "- " & Trim(art_LLAVE!ART_NOmBRE) & "Cantidad : " & c & Chr(13)
    tnombre = Trim(art_LLAVE!art_nombre)
    ARTICULOLOTE = ARTICULOLOTE & "- " & tnombre & " : " & C & Chr(13)
 Else
   ARTICULOLOTE = ARTICULOLOTE & TCOD & Chr(13)
 End If
End If
ARTICULOLOTE = Format(cda_lotes, "000") & ARTICULOLOTE

End Function
Public Function TEXTOWO(WMODO As Integer, msgrid As MSFlexGrid, Optional wWRITE) As String
Dim I As Integer
Dim J As Integer
Dim RUTA As String
Dim WLINEA As String
On Error GoTo SALE
If Trim(msgrid.TextMatrix(2, 1)) = "" And Trim(msgrid.TextMatrix(2, 0)) = "" Then Exit Function

RUTA = "C:\" & "MovLog.txt"
WLINEA = ""
    Open RUTA For Output As #1
    For I = 2 To msgrid.Rows - 1
       WLINEA = ""
       For J = 0 To msgrid.Cols - 1
        If Trim(msgrid.TextMatrix(I, J)) = "" Then
         WLINEA = WLINEA + "@|"
        Else
         WLINEA = WLINEA + Trim(msgrid.TextMatrix(I, J)) & "|"
        End If
       Next J
       WLINEA = Mid(WLINEA, 1, Len(WLINEA) - 1)
       Print #1, WLINEA
    Next I
    Close #1

Exit Function

SALE:
 If Err.Number = 53 Then
    TEXTOWO = ""
 Else
    MsgBox Err.Number & Err.Description, 48, Pub_Titulo
 End If
 
 Close #1
 
End Function

Public Sub llama_anexo()
Load FrmControl
FrmControl.d_codigo.Caption = i_codcli.Text
FrmControl.d_nombre.Caption = ""
FrmControl.d_fecha.Caption = i_fecha_compra.Text
FrmControl.d_serie.Caption = i_numser_c.Text
FrmControl.d_numero.Caption = i_numfac_c.Text
FrmControl.d_numser.Caption = i_numser.Text
FrmControl.d_numfac.Caption = i_numfac.Text
FrmControl.d_nombre.Caption = FORMGEN.i_nomCLI.Caption


FrmControl.Show 1

End Sub


Public Sub Act_sin_stock()

pub_cadena = "DELETE FROM HISPRE WHERE HPR_CODCIA = '" & arm_llave!ARM_CODCIA & "' AND HPR_CODART = " & arm_llave!ARM_codart & " AND HPR_TIPMOV = 2 "
CN.Execute pub_cadena, rdExecDirect


End Sub

Public Function EXTORNA_LOTE() As Boolean
EXTORNA_LOTE = True
PSLOT_LLAVE(0) = LK_CODCIA
PSLOT_LLAVE(1) = arm_llave!ARM_codart
PSLOT_LLAVE(2) = Trim(FORMGEN.grid_fac.TextMatrix(fila, 47))
lot_llave.Requery
If lot_llave.EOF Then
    PSLOT_LLAVE2(0) = LK_CODCIA
    PSLOT_LLAVE2(1) = arm_llave!ARM_codart
    lot_llave2.Requery
    If Not lot_llave2.EOF Then
        lot_llave2.Edit
        lot_llave2!lot_saldos = lot_llave2!lot_saldos + (Val(FORMGEN.grid_fac.TextMatrix(fila, 48)) * pub_signo_arm)
        lot_llave2.Update
        Exit Function
    Else
        CN.Execute "Rollback Transaction", rdExecDirect
        con_llave.Close
        MsgBox "Nro de Lote no Encontrado  : " & Trim(FORMGEN.grid_fac.TextMatrix(fila, 47)), 48, Pub_Titulo
        EXTORNA_LOTE = False
        Exit Function
    End If
End If
lot_llave.Edit
lot_llave!lot_saldos = lot_llave!lot_saldos + (Val(FORMGEN.grid_fac.TextMatrix(fila, 48)) * pub_signo_arm)
lot_llave.Update
fin:
End Function

Public Sub Aplica_NC()
Dim xcI As Integer
Dim cadena_update As String
FrmNC.Calcula_NC
For xcI = 1 To FrmNC.gridcabe.Rows - 1
  If Val(FrmNC.gridcabe.TextMatrix(xcI, 12)) <> 0 Then
    cadena_update = FrmNC.gridcabe.TextMatrix(xcI, 18)
    CN.Execute cadena_update, rdExecDirect
  End If
Next xcI

End Sub

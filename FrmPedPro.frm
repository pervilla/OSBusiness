VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmPedPro 
   Caption         =   "Pedidos a Proveedores"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   9450
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView2 
      Height          =   735
      Left            =   9120
      TabIndex        =   27
      Top             =   7320
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   735
      Left            =   6000
      TabIndex        =   23
      Top             =   7560
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
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
   Begin VB.TextBox txtcondi 
      Height          =   885
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   84
      Top             =   1320
      Width           =   5745
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00808000&
      Height          =   4875
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   4425
      Begin VB.TextBox txtmax 
         Height          =   285
         Left            =   1080
         TabIndex        =   82
         Text            =   "15"
         Top             =   4200
         Width           =   735
      End
      Begin VB.ComboBox cmbemp 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtprecioneto 
         Height          =   285
         Left            =   3600
         TabIndex        =   65
         Top             =   1560
         Width           =   765
      End
      Begin VB.TextBox txtdn4_3 
         Height          =   285
         Left            =   3480
         TabIndex        =   58
         Top             =   3960
         Width           =   795
      End
      Begin VB.TextBox txtdn4_2 
         Height          =   285
         Left            =   3480
         TabIndex        =   57
         Top             =   3600
         Width           =   795
      End
      Begin VB.TextBox txtbonif 
         Height          =   285
         Left            =   2040
         TabIndex        =   51
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtdn4 
         Height          =   285
         Left            =   3480
         TabIndex        =   49
         Top             =   3240
         Width           =   795
      End
      Begin VB.TextBox txtdn2 
         Height          =   285
         Left            =   3480
         TabIndex        =   48
         Top             =   2520
         Width           =   795
      End
      Begin VB.TextBox txtdn3 
         Height          =   285
         Left            =   3480
         TabIndex        =   47
         Top             =   2880
         Width           =   795
      End
      Begin VB.TextBox txtdn1 
         Height          =   285
         Left            =   3480
         TabIndex        =   38
         Top             =   2160
         Width           =   795
      End
      Begin VB.TextBox txtd3 
         Height          =   285
         Left            =   1080
         TabIndex        =   37
         Top             =   2880
         Width           =   675
      End
      Begin VB.TextBox txtd2 
         Height          =   285
         Left            =   1080
         TabIndex        =   36
         Top             =   2520
         Width           =   675
      End
      Begin VB.TextBox txtd1 
         Height          =   285
         Left            =   1080
         TabIndex        =   35
         Top             =   2160
         Width           =   675
      End
      Begin VB.ComboBox cmbtipo 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4440
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtcantidad 
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtprecio 
         Height          =   285
         Left            =   2880
         TabIndex        =   13
         Top             =   1560
         Width           =   645
      End
      Begin VB.TextBox i_codart2 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox unidad 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1560
         Width           =   945
      End
      Begin VB.TextBox txttotal 
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Max.Item:"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Empresa :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "NETO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   28
         Left            =   3720
         TabIndex        =   67
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "PRECIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   27
         Left            =   3600
         TabIndex        =   66
         Top             =   1125
         Width           =   705
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "BiMestral"
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   26
         Left            =   2640
         TabIndex        =   61
         Top             =   3600
         Width           =   645
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "TriMestral"
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   25
         Left            =   2640
         TabIndex        =   60
         Top             =   3960
         Width           =   810
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "Mensual"
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   24
         Left            =   2640
         TabIndex        =   59
         Top             =   3240
         Width           =   720
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Bonific."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   23
         Left            =   2040
         TabIndex        =   53
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Afecta"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   22
         Left            =   1320
         TabIndex        =   50
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "Vales"
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   20
         Left            =   2640
         TabIndex        =   46
         Top             =   2880
         Width           =   390
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "MM"
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   19
         Left            =   2640
         TabIndex        =   45
         Top             =   2520
         Width           =   270
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "Efect."
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   18
         Left            =   2640
         TabIndex        =   44
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "Desctos. NO Afectos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   17
         Left            =   2760
         TabIndex        =   43
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "Acumulado"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   42
         Top             =   2880
         Width           =   795
      End
      Begin VB.Label lblcontrol 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Financiero"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   41
         Top             =   2520
         Width           =   750
      End
      Begin VB.Label lblcontrol 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Item"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   40
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "Desctos. Afectos :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   360
         TabIndex        =   39
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "C/I.G.V."
         Height          =   195
         Index           =   9
         Left            =   3000
         TabIndex        =   34
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label i_nomarti 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   24
         Top             =   240
         Width           =   2385
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Total S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   21
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Tipo de Descuento :"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   20
         Top             =   4560
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Producto:"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Alignment       =   2  'Center
         Caption         =   "C a n t i d a d "
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
         Index           =   2
         Left            =   1440
         TabIndex        =   18
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label lblcontrol 
         AutoSize        =   -1  'True
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   3000
         TabIndex        =   17
         Top             =   1125
         Width           =   450
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Unidad"
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame fradetall 
      Caption         =   "Detalle del Movimiento :"
      Height          =   2655
      Left            =   4440
      TabIndex        =   70
      Top             =   3360
      Visible         =   0   'False
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid grdidetall 
         Height          =   1965
         Left            =   120
         TabIndex        =   71
         Top             =   600
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3466
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label lunidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   79
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblgridall 
         Caption         =   "Unidad:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lsaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   77
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lentregado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         TabIndex        =   76
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lcantidad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   75
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblgridall 
         Caption         =   "Entregado:"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   74
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblgridall 
         Caption         =   "Saldo :"
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   73
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblgridall 
         Caption         =   "Cantidad:"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   72
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmddd 
      BackColor       =   &H80000004&
      Caption         =   "&Cronogramas de  Pagos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   64
      Tag             =   "0"
      Top             =   6360
      Width           =   1560
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Height          =   625
      Left            =   4200
      Picture         =   "FrmPedPro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6360
      Width           =   1545
   End
   Begin VB.CommandButton CmdImprimir 
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
      Height          =   625
      Left            =   2400
      Picture         =   "FrmPedPro.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6360
      Width           =   1545
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   10680
      Picture         =   "FrmPedPro.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6360
      Width           =   825
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   600
      Picture         =   "FrmPedPro.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6360
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Proveedor :"
      ForeColor       =   &H00808000&
      Height          =   1320
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11745
      Begin VB.CommandButton CmdAnterior 
         Caption         =   "&Ant."
         Height          =   495
         Left            =   9600
         Picture         =   "FrmPedPro.frx":0CD8
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "&Sig"
         Height          =   495
         Left            =   10440
         Picture         =   "FrmPedPro.frx":111A
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtglosa 
         Height          =   285
         Left            =   720
         TabIndex        =   55
         Top             =   960
         Width           =   6225
      End
      Begin VB.CheckBox cheliq 
         Caption         =   "Pedidos Confirmado"
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
         Height          =   495
         Left            =   9360
         TabIndex        =   52
         Top             =   120
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.TextBox txtrepre 
         Height          =   285
         Left            =   1320
         TabIndex        =   33
         Top             =   600
         Width           =   2745
      End
      Begin VB.CheckBox chenumfac 
         Caption         =   "Digitar Numero"
         Height          =   195
         Left            =   6120
         TabIndex        =   29
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtnumfac 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8400
         TabIndex        =   28
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txt_cli 
         BackColor       =   &H00E0E0E0&
         DataField       =   "ART_KEY"
         DataSource      =   "Data1"
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
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2790
         TabIndex        =   6
         Top             =   270
         Width           =   6315
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Glosa : "
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   56
         Top             =   975
         Width           =   585
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Representante:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Proveedor :"
         Height          =   255
         Index           =   10
         Left            =   90
         TabIndex        =   8
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Interno : "
         Height          =   255
         Index           =   13
         Left            =   7680
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.Label d_fecha 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Fecha : "
         Height          =   255
         Index           =   11
         Left            =   4200
         TabIndex        =   7
         Top             =   600
         Width           =   765
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdidet 
      Height          =   3645
      Left            =   4440
      TabIndex        =   1
      Top             =   2520
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6429
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblcontrol 
      Caption         =   "Condición:"
      Height          =   255
      Index           =   21
      Left            =   4800
      TabIndex        =   85
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10320
      TabIndex        =   63
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Total Pedido="
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   9120
      TabIndex        =   62
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   54
      Top             =   6240
      Width           =   11775
   End
   Begin VB.Label flag_signo 
      Caption         =   "1"
      Height          =   225
      Left            =   3180
      TabIndex        =   26
      Top             =   1110
      Width           =   435
   End
   Begin VB.Label lblcontrol 
      Caption         =   "Maestros del Pedido :"
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   2
      Top             =   2280
      Width           =   3015
   End
End
Attribute VB_Name = "FrmPedPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_cp As String
Dim PRE_RELCOMP As rdoQuery
Dim relcomp_llave  As rdoResultset
Dim loc_key As Integer
Dim PRE_UNIDADES As rdoQuery
Dim preuni_llave  As rdoResultset

Private Sub cheliq_Click()
txtglosa.SetFocus
End Sub

Private Sub chenumfac_Click()
If chenumfac.Value = 1 Then
  llenanumfac
  txtnumfac.Enabled = True
  Azul txtnumfac, txtnumfac
Else
  txtnumfac.Enabled = False
End If
End Sub

Private Sub cmbemp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then unidad.SetFocus
End Sub

Private Sub cmbtipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  i_codart2.SetFocus
End If
End Sub

Private Sub cmbtipo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = cmbTipo.ListIndex
PUB_TIPREG = Mid(cmbTipo.ToolTipText, 13, Len(cmbTipo.ToolTipText))
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
LLENADOS cmbTipo, 43
cmbTipo.SetFocus
SendKeys "%{up}"


End Sub

Private Sub cmdAgregar_Click()
Dim ps_ultimo As rdoResultset
If Val(txt_cli.Text) = 0 Then
    MsgBox "Debe Ingresar Datos del Proveedor", 48, Pub_Titulo
    Azul txt_cli, txt_cli
    Exit Sub
End If
If grdidet.Rows - 1 = 0 Then
    MsgBox "Debe Ingresar Datos para registrar el Pedido", 48, Pub_Titulo
    Azul txt_cli, txt_cli
    Exit Sub
End If


If chenumfac.Value = 1 Then

Else
    pub_cadena = "SELECT top 1 rel_numfac FROM PEDPRO WHERE REL_CODCIA = '" & LK_CODCIA & "' AND REL_CP = 'P' order by rel_numfac desc "
    Set ps_ultimo = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    If ps_ultimo.EOF Then
      txtnumfac.Text = "1"
    Else
      txtnumfac.Text = Val(ps_ultimo!rel_numfac) + 1
    End If
End If
  
PRE_RELCOMP(0) = LK_CODCIA
PRE_RELCOMP(1) = "P"
PRE_RELCOMP(2) = Val(txt_cli.Text)
PRE_RELCOMP(3) = 0
PRE_RELCOMP(4) = Val(txtnumfac.Text)
relcomp_llave.Requery
If Not relcomp_llave.EOF Then
 pub_mensaje = "Reemplazar la infomación...   ¿Desea Continuar... ?"
 Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
 If Pub_Respuesta = vbNo Then
    Exit Sub
 End If
 pub_cadena = "DELETE FROM PEDPRO WHERE REL_CODCIA = '" & LK_CODCIA & "' AND REL_CP = 'P' AND REL_CODPRO = " & Val(txt_cli.Text) & " AND REL_NUMSER = 0 AND REL_NUMFAC =  " & txtnumfac.Text & " "
 CN.Execute pub_cadena, rdExecDirect
End If
Screen.MousePointer = 13
For fila = 1 To grdidet.Rows - 1
    relcomp_llave.AddNew
    relcomp_llave!rel_codcia = LK_CODCIA
    relcomp_llave!rel_cp = "P"
    relcomp_llave!rel_codpro = Val(txt_cli.Text)
    relcomp_llave!rel_numser = 0
    relcomp_llave!rel_numfac = Val(txtnumfac.Text)
    relcomp_llave!REL_NUMSEC = fila
    relcomp_llave!rel_FECHA_COMPRA = Trim(d_fecha.Caption)
    relcomp_llave!rel_numser_c = 0
    relcomp_llave!rel_numfac_c = 0
    relcomp_llave!rel_codtip = Val(Right(grdidet.TextMatrix(fila, 1), 8))
    relcomp_llave!rel_codart = Val(grdidet.TextMatrix(fila, 2))
    relcomp_llave!rel_descri = Trim(Left(grdidet.TextMatrix(fila, 4), 20))
    relcomp_llave!rel_equiv = Val(Trim(Right(grdidet.TextMatrix(fila, 4), 8)))
    relcomp_llave!rel_cantidad = Val(grdidet.TextMatrix(fila, 5)) * Val(relcomp_llave!rel_equiv)
    relcomp_llave!rel_total = Val(grdidet.TextMatrix(fila, 8))
    relcomp_llave!rel_precio = Val(grdidet.TextMatrix(fila, 6))
    relcomp_llave!rel_GLOSA = Trim(txtglosa.Text) ' Trim((grdidet.TextMatrix(fila, 7)))
    relcomp_llave!rel_CONDI = Trim(txtcondi.Text) ' Trim((grdidet.TextMatrix(fila, 7)))
    
    relcomp_llave!rel_signo_carga = Val((grdidet.TextMatrix(fila, 9)))
    relcomp_llave!rel_FECHA = grdidet.TextMatrix(fila, 11)
    relcomp_llave!rel_LIQUIDO = cheliq.Value
    
    relcomp_llave!rel_bonif = Val(grdidet.TextMatrix(fila, 13))
    relcomp_llave!rel_d1 = Val(grdidet.TextMatrix(fila, 14))
    relcomp_llave!rel_d2 = Val(grdidet.TextMatrix(fila, 15))
    relcomp_llave!rel_d3 = Val(grdidet.TextMatrix(fila, 16))
    relcomp_llave!rel_dn1 = Val(grdidet.TextMatrix(fila, 17))
    relcomp_llave!rel_dn2 = Val(grdidet.TextMatrix(fila, 18))
    relcomp_llave!rel_dn3 = Val(grdidet.TextMatrix(fila, 19))
    relcomp_llave!rel_dn4 = Val(grdidet.TextMatrix(fila, 20))
    relcomp_llave!rel_dn4_2 = Val(grdidet.TextMatrix(fila, 21))
    relcomp_llave!rel_dn4_3 = Val(grdidet.TextMatrix(fila, 22))
    relcomp_llave!rel_repre = Trim(txtrepre.Text)
    relcomp_llave!rel_cantidad_p = Val(grdidet.TextMatrix(fila, 24)) * Val(relcomp_llave!rel_equiv)
    relcomp_llave!rel_precio_pro = Val(grdidet.TextMatrix(fila, 25))
    relcomp_llave!rel_precio_neto = Val(grdidet.TextMatrix(fila, 23))
    relcomp_llave!rel_saldo = relcomp_llave!rel_cantidad
    relcomp_llave!rel_codusu = LK_CODUSU
    relcomp_llave!rel_codcia_ref = Left(grdidet.TextMatrix(fila, 27), 2)
    relcomp_llave.Update
Next fila
Screen.MousePointer = 0
pub_mensaje = "Datos Guardados correctamente!!! ...   ¿Desea Imprimir el Pedido... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
  cmdcancelar_Click
   Exit Sub
End If
CmdImprimir_Click
cmdcancelar_Click
'Unload FrmPedPro
End Sub

Private Sub CmdAnterior_Click()
If chenumfac.Value = 0 Then Exit Sub
If txtnumfac.Text >= 1 Then
  txtnumfac.Text = Val(txtnumfac.Text) - 1
  txtnumfac_KeyPress 13
End If
End Sub

Private Sub cmdcancelar_Click()
txtnumfac.Text = ""
d_fecha.Caption = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txt_cli.Text = ""
lblCliente.Caption = ""

i_codart2.Text = ""
unidad.ListIndex = -1
txtcantidad.Text = ""
cmbTipo.ListIndex = -1
txttotal.Text = ""
txtprecio.Tag = ""
txtglosa.Text = ""
txtcondi.Text = ""
chenumfac.Value = 0
cheliq.Value = 0
txtcantidad.Text = ""
txtprecio.Text = ""
txtprecio.Tag = ""
txttotal.Text = ""
lbltotal.Caption = "0.00"
txtbonif.Text = ""
txtd1.Text = ""
txtd2.Text = ""
txtd3.Text = ""
txtdn1.Text = ""
txtdn2.Text = ""
txtdn3.Text = ""
txtdn4.Text = ""
txtdn4_2.Text = ""
txtdn4_3.Text = ""
cabe
Azul txt_cli, txt_cli

End Sub

Private Sub cmdCerrar_Click()
Unload FrmPedPro
End Sub

Private Sub cmddd_Click()
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
   Exit Sub
End If

Load frmdeudas
frmdeudas.txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
frmdeudas.txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
frmdeudas.Show 1
End Sub

Private Sub CmdImprimir_Click()
Screen.MousePointer = 11
Reportes.Connect = PUB_ODBC
Reportes.Destination = crptToWindow  '= crptToPrinter
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
Reportes.Formulas(1) = ""
Reportes.WindowTitle = "Orden de Pedido " & Reportes.ReportFileName
Reportes.ReportFileName = PUB_RUTA_OTRO + "PEDPRO.RPT"
'{RELCOMPRA.REL_CODCIA} = "01"
pub_cadena = "{RELCOMPRA.REL_CODCIA} = '" & LK_CODCIA & "' AND {RELCOMPRA.REL_NUMFAC} =  " & Val(txtnumfac.Text) & " "
Reportes.SelectionFormula = pub_cadena
Reportes.Action = 1
Screen.MousePointer = 0
End Sub

Private Sub cmdSiguiente_Click()
If chenumfac.Value = 0 Then Exit Sub
  txtnumfac.Text = Val(txtnumfac.Text) + 1
  txtnumfac_KeyPress 13
End Sub

Private Sub Form_Activate()
If LK_CODUSU = "ADMIN" Then
 txtmax.Locked = False
Else
 txtmax.Locked = True
End If
Screen.MousePointer = 0
loc_cp = "P"
'If fraanexos.Visible = True Then
' GoTo anexo_act
'End If
Exit Sub
PRE_RELCOMP(0) = LK_CODCIA
PRE_RELCOMP(1) = "P"
PRE_RELCOMP(2) = Val(d_codigo.Caption)
PRE_RELCOMP(3) = Val(d_numser.Caption)
PRE_RELCOMP(4) = Val(d_numfac.Caption)
relcomp_llave.Requery
If relcomp_llave.EOF Then
If Val(flag_signo.Caption) = 1 Then
   cmdsigno.ListIndex = 0
ElseIf Val(flag_signo.Caption) = -1 Then
   cmdsigno.ListIndex = 1
End If

   Exit Sub
End If

cabe
cheliq.Value = relcomp_llave!rel_LIQUIDO
Do Until relcomp_llave.EOF
    grdidet.Rows = grdidet.Rows + 1
    PUB_TIPREG = 43
    PUB_NUMTAB = relcomp_llave!rel_codtip
    SQ_OPER = 1
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
       grdidet.TextMatrix(grdidet.Rows - 1, 1) = tab_llave!TAB_NOMLARGO & String(80, " ") & tab_llave!TAB_NUMTAB
    End If
    grdidet.TextMatrix(grdidet.Rows - 1, 2) = relcomp_llave!rel_codart
    grdidet.TextMatrix(grdidet.Rows - 1, 0) = Format(relcomp_llave!rel_FECHA, "dd/mm/yy")
    If Val(relcomp_llave!rel_codart) <> 0 Then
     SQ_OPER = 1
     PUB_KEY = relcomp_llave!rel_codart
     pu_codcia = relcomp_llave!rel_codcia
     LEER_ART_LLAVE
     
     grdidet.TextMatrix(grdidet.Rows - 1, 3) = Trim(art_LLAVE!art_nombre)
     grdidet.TextMatrix(grdidet.Rows - 1, 4) = relcomp_llave!rel_descri & String(80, " ") & relcomp_llave!rel_equiv
     grdidet.TextMatrix(grdidet.Rows - 1, 5) = Format(relcomp_llave!rel_cantidad / relcomp_llave!rel_equiv, "0.00")
     grdidet.TextMatrix(grdidet.Rows - 1, 6) = relcomp_llave!rel_precio
    End If
    grdidet.TextMatrix(grdidet.Rows - 1, 7) = Trim(relcomp_llave!rel_GLOSA)
    grdidet.TextMatrix(grdidet.Rows - 1, 8) = relcomp_llave!rel_total
    grdidet.TextMatrix(grdidet.Rows - 1, 9) = relcomp_llave!rel_signo_carga
    grdidet.TextMatrix(grdidet.Rows - 1, 10) = "A"
    grdidet.TextMatrix(grdidet.Rows - 1, 11) = relcomp_llave!rel_FECHA
    grdidet.TextMatrix(grdidet.Rows - 1, 12) = relcomp_llave!rel_LIQUIDO
    If Val(relcomp_llave!rel_signo_carga) = -1 Then
      grdidet.Row = grdidet.Rows - 1
      grdidet.COL = 1
      grdidet.CellBackColor = QBColor(14)
    End If
 relcomp_llave.MoveNext
Loop
If Val(flag_signo.Caption) = 1 Then
   cmdsigno.ListIndex = 0
ElseIf Val(flag_signo.Caption) = -1 Then
   cmdsigno.ListIndex = 1
End If

Exit Sub
anexo_act:

pub_cadena = "SELECT * FROM PEDPRO WHERE REL_CODCIA = ? AND REL_CP = ? AND REL_CODPRO = ? AND " & archi & "  ORDER BY REL_NUMSER, REL_NUMFAC ,REL_NUMSEC"
Set PRE_RELCOMP = CN.CreateQuery("", pub_cadena)
PRE_RELCOMP(0) = 0
PRE_RELCOMP(1) = 0
PRE_RELCOMP(2) = 0
Set relcomp_llave = PRE_RELCOMP.OpenResultset(rdOpenKeyset, rdConcurValues)
PRE_RELCOMP(0) = LK_CODCIA
PRE_RELCOMP(1) = "P"
PRE_RELCOMP(2) = Val(d_codigo.Caption)
relcomp_llave.Requery

CABE_MUESTRA
Do Until relcomp_llave.EOF
    GRID_MST.Rows = GRID_MST.Rows + 1
    PUB_TIPREG = 43
    PUB_NUMTAB = relcomp_llave!rel_codtip
    SQ_OPER = 1
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
       GRID_MST.TextMatrix(GRID_MST.Rows - 1, 0) = tab_llave!TAB_NOMLARGO & String(80, " ") & tab_llave!TAB_NUMTAB
    End If
    GRID_MST.TextMatrix(GRID_MST.Rows - 1, 1) = relcomp_llave!rel_codart
    If Val(relcomp_llave!rel_codart) <> 0 Then
     SQ_OPER = 1
     PUB_KEY = relcomp_llave!rel_codart
     pu_codcia = relcomp_llave!rel_codcia
     LEER_ART_LLAVE
     GRID_MST.TextMatrix(GRID_MST.Rows - 1, 2) = Trim(art_LLAVE!art_nombre)
     GRID_MST.TextMatrix(GRID_MST.Rows - 1, 3) = relcomp_llave!rel_descri & String(80, " ") & relcomp_llave!rel_equiv
     GRID_MST.TextMatrix(GRID_MST.Rows - 1, 4) = Format(relcomp_llave!rel_cantidad / relcomp_llave!rel_equiv, "0.00")
     GRID_MST.TextMatrix(GRID_MST.Rows - 1, 5) = relcomp_llave!rel_precio
    End If
    GRID_MST.TextMatrix(GRID_MST.Rows - 1, 6) = Trim(relcomp_llave!rel_GLOSA)
    GRID_MST.TextMatrix(GRID_MST.Rows - 1, 7) = relcomp_llave!rel_total

 relcomp_llave.MoveNext
Loop

Exit Sub



End Sub

Private Sub Form_Load()
CenterMe FrmPedPro
loc_key = 0
cabe
PUB_CODCIA = LK_CODCIA
LLENADOS cmbTipo, 43
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ?  ORDER BY PRE_CODART"
Set PRE_UNIDADES = CN.CreateQuery("", pub_cadena)
PRE_UNIDADES(0) = 0
PRE_UNIDADES(1) = 0
Set preuni_llave = PRE_UNIDADES.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM PEDPRO WHERE REL_CODCIA = ? AND REL_CP = ? AND REL_CODPRO = ? AND REL_NUMSER = ? AND REL_NUMFAC = ?  ORDER BY REL_NUMSER, REL_NUMFAC "
Set PRE_RELCOMP = CN.CreateQuery("", pub_cadena)
PRE_RELCOMP(0) = 0
PRE_RELCOMP(1) = 0
PRE_RELCOMP(2) = 0
PRE_RELCOMP(3) = 0
PRE_RELCOMP(4) = 0
Set relcomp_llave = PRE_RELCOMP.OpenResultset(rdOpenKeyset, rdConcurValues)
d_fecha.Caption = Format(LK_FECHA_DIA, "dd/mm/yyyy")
cmdcancelar.Enabled = True
'If LK_CODCIA = "05" Then
PSPAR_MULTI(0) = LK_CODCIA
par_multi.Requery
cmbemp.AddItem par_multi!par_codcia & " - " & Trim(par_multi!par_nombre)

'Else
'    PSPAR_MULTI(0) = "01"
'    par_multi.Requery
'    cmbemp.AddItem par_multi!par_codcia & " - " & Trim(par_multi!par_nombre)
'    PSPAR_MULTI(0) = "03"
'    par_multi.Requery
'    cmbemp.AddItem par_multi!par_codcia & " - " & Trim(par_multi!par_nombre)
'    PSPAR_MULTI(0) = "07"
'    par_multi.Requery
'    cmbemp.AddItem par_multi!par_codcia & " - " & Trim(par_multi!par_nombre)
'End If
End Sub


Private Sub grdidet_DblClick()
Dim ps_pedpro As rdoResultset

fradetall.Visible = True
cabe_detalle
If grdidet.Rows <= 1 Then Exit Sub
PUB_CODART = Val(grdidet.TextMatrix(grdidet.Row, 2))
If PUB_CODART = 0 Then Exit Sub
pub_cadena = "SELECT * FROM PEDPROALL WHERE REL_CODCIA = '" & LK_CODCIA & "' AND REL_CODART = " & PUB_CODART & " AND REL_CODART = " & PUB_CODART & " AND  REL_SER_INT = 0 AND   REL_FAC_INT = " & txtnumfac.Text & " "
Set ps_pedpro = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
fradetall.Caption = "Detalle del Producto: " & Trim(grdidet.TextMatrix(grdidet.Row, 3))
lunidad.Caption = Trim(grdidet.TextMatrix(grdidet.Row, 4))
lcantidad.Caption = Format(Val(grdidet.TextMatrix(grdidet.Row, 5)), "0.00")
lsaldo.Caption = Format(Val(grdidet.TextMatrix(grdidet.Row, 24)), "0.00")
lentregado.Caption = Format(Val(grdidet.TextMatrix(grdidet.Row, 5)) - Val(grdidet.TextMatrix(grdidet.Row, 24)), "0.00")
Do Until ps_pedpro.EOF
    grdidetall.Rows = grdidetall.Rows + 1
    grdidetall.TextMatrix(grdidetall.Rows - 1, 0) = ps_pedpro!rel_FECHA
    grdidetall.TextMatrix(grdidetall.Rows - 1, 1) = ps_pedpro!rel_numser
    grdidetall.TextMatrix(grdidetall.Rows - 1, 2) = ps_pedpro!rel_numfac
    SQ_OPER = 1
    pu_cp = "P"
    pu_codclie = ps_pedpro!rel_codpro
    PUB_CODCIA = LK_CODCIA
    LEER_CLI_LLAVE
    If Not cli_llave.EOF Then
       grdidetall.TextMatrix(grdidetall.Rows - 1, 3) = Trim(cli_llave!cli_nombre)
    Else
      grdidetall.TextMatrix(grdidetall.Rows - 1, 3) = "No Existe este Codigo.." & pu_codclie
    End If
    grdidetall.TextMatrix(grdidetall.Rows - 1, 4) = ps_pedpro!rel_descri
    grdidetall.TextMatrix(grdidetall.Rows - 1, 5) = Format(ps_pedpro!rel_cantidad / ps_pedpro!rel_equiv, "0.00")
    grdidetall.TextMatrix(grdidetall.Rows - 1, 6) = ps_pedpro!rel_precio
ps_pedpro.MoveNext
Loop
grdidetall.Visible = True
grdidetall.SetFocus

End Sub

Private Sub grdidet_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo sale
If KeyCode = 46 Then
  If Val(grdidet.TextMatrix(grdidet.Row, 5)) <> Val(grdidet.TextMatrix(grdidet.Row, 24)) Then
    MsgBox "No puede eliminar este Registro, por tener movimientos.", 48, Pub_Titulo
    Exit Sub
  End If
  If grdidet.Rows = 2 Then
    grdidet.Rows = 1
  ElseIf grdidet.Rows > 2 Then
    grdidet.RemoveItem grdidet.Row
  End If
End If
sum_total
sale:
End Sub

Private Sub grdidetall_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 grdidet.COL = 2
 grdidet.SetFocus
End If
End Sub

Private Sub grdidetall_LostFocus()
fradetall.Visible = False
End Sub

Private Sub i_codart2_Change()
If i_codart2.Text = "" Then
  i_nomarti.Caption = ""
  VAR_ACTIVAR = 0
End If

End Sub

Private Sub i_codart2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codart2.Text = "" Then
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
  i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  i_codart2.SelStart = Len(i_codart2.Text)

     If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        SQ_OPER = 3
        pu_alterno = i_codart2.Text
        pu_codcia = LK_CODCIA
        LEER_ART_LLAVE
        If art_llave_alt.EOF Then
           MsgBox "Codigo No Existe ...", 48, Pub_Titulo
           Exit Sub
        End If
        PUB_CODART = art_llave_alt!art_key
     Else
        SQ_OPER = 1

        PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
        pu_codcia = LK_CODCIA
        LEER_ART_LLAVE
        If art_LLAVE.EOF Then
           MsgBox "Codigo No Existe ...", 48, Pub_Titulo
           Exit Sub
        End If
        PUB_CODART = art_LLAVE!art_key
        i_nomarti.Caption = art_LLAVE!art_nombre
     End If

fin:

End Sub
Private Sub i_codart2_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I, car
Dim itmFound As MSComctlLib.ListItem
Dim unid_max As Integer
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView1.Visible = False
 i_codart2.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
If i_codart2.Text = "" Then
  txtglosa.SetFocus
  Exit Sub
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
Else
 On Error GoTo mucho
 PUB_KEY = Val(i_codart2.Text)
 On Error GoTo 0
 If Len(i_codart2.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(i_codart2.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = i_codart2.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       GoTo fin
    End If
    If Val(art_LLAVE!art_numero) = 3 Then
           MsgBox "Producto en Proceso de Liquidación ", 48, Pub_Titulo
     Exit Sub
    End If
    WCOD_ORIGINAL = art_LLAVE!art_key
    i_nomarti.Caption = Trim(art_LLAVE!art_nombre)
    'i_codart2.text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
    PRE_UNIDADES(0) = LK_CODCIA
    PRE_UNIDADES(1) = art_LLAVE!art_key
    preuni_llave.Requery
    unidad.Clear
    unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!PRE_UNIDAD & String(80, " ") & preuni_llave!PRE_EQUIV
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
         unid_max = preuni_llave.AbsolutePosition - 1
'         LBLUNIDAD.Caption = Trim(preuni_llave!pre_UNIDAD)
      '   Exit Do
       End If
      preuni_llave.MoveNext
    Loop
    ListView1.Visible = False
    If unidad.ListCount > 0 Then unidad.ListIndex = unid_max
    cmbemp.SetFocus
    Exit Sub
Else
  If ListView1.Visible = False And VAR_ACTIVAR <> 99 And i_codart2.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = i_codart2.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       Exit Sub
     End If
     WCOD_ORIGINAL = art_llave_alt!art_key
     'i_codart2.text = Trim(art_llave_alt!ART_NOMBRE)
     If art_llave_alt!ART_MONEDA = "D" Then
      lblMoneda.Caption = "Moneda : US$."
     Else
      lblMoneda.Caption = "Moneda : S/ ."
     End If
     i_nomarti.Caption = Trim(art_llave_alt!art_nombre)
     ListView1.Visible = False
     
     PRE_UNIDADES(0) = LK_CODCIA
     PRE_UNIDADES(1) = art_llave_alt!art_key
     preuni_llave.Requery
     unidad.Clear
     unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!PRE_UNIDAD & String(80, " ") & preuni_llave!PRE_EQUIV
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
          unid_max = preuni_llave.AbsolutePosition - 1
          LBLUNIDAD.Caption = Trim(preuni_llave!PRE_UNIDAD)
          Exit Do
        End If
        preuni_llave.MoveNext
     Loop
     If unidad.ListCount > 0 Then unidad.ListIndex = unid_max
      cmbemp.SetFocus
     'If Frame2.Visible And i_maximo.Visible Then i_maximo.SetFocus
     'If Frame1.Visible And txtCampo1.Visible Then Azul2 txtCampo1, txtCampo1
     
     'CmdProcesa.SetFocus
     Exit Sub
  Else
    If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
    If Trim(UCase(i_codart2.Text)) = Left(VALOR, Len(Trim(i_codart2.Text))) And Len(Trim(i_codart2.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       PUB_KEY = i_codart2.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If Val(art_LLAVE!art_numero) = 3 Then
           MsgBox "Producto en Proceso de Liquidación ", 48, Pub_Titulo
     Exit Sub
       End If
    
    
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul i_codart2, i_codart2
        Exit Sub
      End If
      WCOD_ORIGINAL = art_LLAVE!art_key
      i_nomarti.Caption = Trim(art_LLAVE!art_nombre)
    PRE_UNIDADES(0) = LK_CODCIA
    PRE_UNIDADES(1) = art_LLAVE!art_key
    preuni_llave.Requery
    unidad.Clear
    unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!PRE_UNIDAD & String(80, " ") & preuni_llave!PRE_EQUIV
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
         unid_max = preuni_llave.AbsolutePosition - 1
'         LBLUNIDAD.Caption = Trim(preuni_llave!pre_UNIDAD)
      '   Exit Do
       End If
      preuni_llave.MoveNext
    Loop
    ListView1.Visible = False
    If unidad.ListCount > 0 Then unidad.ListIndex = unid_max
      cmbemp.SetFocus
      ListView1.Visible = False
      If unidad.ListCount > 0 Then unidad.ListIndex = unid_max
      i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      ListView1.Visible = False
      'If Frame2.Visible And i_maximo.Visible Then i_maximo.SetFocus
      'If Frame1.Visible And txtCampo1.Visible Then Azul2 txtCampo1, txtCampo1
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
Azul i_codart2, i_codart2
  

End Sub

Private Sub i_codart2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
Dim ws_codcia As String * 2
If KeyCode = 13 Then Exit Sub
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  If Len(i_codart2.Text) = 0 Or i_codart2.Text = "" Then
    ListView1.Visible = False
    Exit Sub
  End If
  If i_codart2.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf i_codart2.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(i_codart2.Text, 1) = "*" Then
   i_codart2.Text = Mid(i_codart2.Text, 2, Len(i_codart2.Text))
   i_codart2.SelStart = Len(i_codart2.Text)
  End If
Else
 If Len(i_codart2.Text) = 0 Or IsNumeric(i_codart2.Text) = True Then
   ListView1.Visible = False
   Exit Sub
 End If
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(i_codart2.Text) = 1 Then
    If i_codart2.Text = "" Then Exit Sub
    var = Asc(i_codart2.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      ws_codcia = "00"
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  ART_SITUACION <> 1 AND  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  ART_SITUACION <> 1  AND (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
    End If
   ' If Len(I_CODART2.text) > 1 And ListView1.ListItems.count = 0 Then
   ' Else
     PROC_LISVIEW ListView1
   ' End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(i_codart2.Text), lvwText, , lvwPartial)
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

Public Sub LLENADOS(cont As ComboBox, tip As Integer)
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
       If PUB_TIPREG = 35 And LK_EMP = "PAR" Then
          If Val(tab_mayor!tab_codart) = WW_CODVEN Then
            cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
          End If
       Else
           cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
       End If
       tab_mayor.MoveNext
    Loop
End Sub


Private Sub txt_cli_GotFocus()
Azul txt_cli, txt_cli
lblCliente.Caption = ""
End Sub
Private Sub txt_cli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not ListView2.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_cli.Text = "" Then
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
'  KeyCode = 0
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).Text) & " "
  DoEvents
  txt_cli.SelStart = Len(txt_cli.Text)
  DoEvents
fin:

End Sub
Private Sub txt_cli_KeyPress(KeyAscii As Integer)
Dim var As String
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If KeyAscii = 27 Then
 ListView2.Visible = False
 txt_cli.Text = ""
 Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
If KeyAscii = 13 And Left(txt_cli.Text, 1) = "+" Then GoTo buscar
On Error GoTo ERROR_CODIGO
pu_codclie = Val(txt_cli.Text)
On Error GoTo 0
If Len(txt_cli.Text) = 0 Then
   Exit Sub
End If

If pu_codclie <> 0 And IsNumeric(txt_cli.Text) = True Then
   SQ_OPER = 1
   pu_cp = loc_cp
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
     lblCliente.Caption = ""
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Azul txt_cli, txt_cli
     GoTo fin
   Else
     lblCliente.Caption = Trim(cli_llave!cli_nombre)
     LOC_RUC = Trim(cli_llave!cli_ruc_esposo)
   End If
'   If Pantalla.Visible And Pantalla.Enabled Then
'     Pantalla.SetFocus
'   End If
    txtrepre.SetFocus
Else
   If loc_key > ListView2.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   VALOR = UCase(ListView2.ListItems.Item(loc_key).Text)
   If Trim(UCase(txt_cli.Text)) = Left(VALOR, Len(Trim(txt_cli.Text))) Then
   Else
      Exit Sub
   End If
   txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).SubItems(1))
   pu_codclie = Val(txt_cli.Text)
   SQ_OPER = 1
   pu_cp = loc_cp
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   If Not cli_llave.EOF Then
    lblCliente.Caption = Trim(ListView2.ListItems.Item(loc_key).Text)
    LOC_RUC = Trim(cli_llave!cli_ruc_esposo)
   End If
   txtrepre.SetFocus
End If

dale:
ListView2.Visible = False
fin:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul txt_cli, txt_cli
Exit Sub

buscar:
var = Mid(txt_cli.Text, 2, Len(txt_cli.Text))
numarchi = alta_vista_nombre(ListView2, var, loc_cp)
If numarchi = 0 Then
  ListView2.Visible = False
  MsgBox "Alta Vista: No Existe .. Esta descripcion..", 48, Pub_Titulo
Else
  ListView2.Visible = True
  txt_cli.SetFocus
End If
loc_key = 1


End Sub

Private Sub txt_cli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(txt_cli.Text) = 0 Or IsNumeric(txt_cli.Text) = True Then
   ListView2.Visible = False
   Exit Sub
End If
If ListView2.Visible = False And KeyCode <> 13 Then
    var = Asc(txt_cli.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    ElseIf var = 58 Then
       var = "A"
    Else
       var = Chr(var)
    End If
    numarchi = 1
    'archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM  FROM CLIENTES WHERE  CLI_CP = '" & loc_cp & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_cli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    archi = "SELECT TOP 2000  CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = '" & loc_cp & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_cli.Text & "' AND  '" & var & "' ORDER BY CLI_NOMBRE"
'    If Trim(txt_cli.text) <> "" And ListView1.ListItems.count = 0 Then
'    Else
     PROC_LISVIEW ListView2, 3000
     loc_key = 0
     If ListView2.Visible Then
      loc_key = 1
     End If
 '   End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If ListView2.Visible Then
  Set itmFound = ListView2.FindItem(LTrim(txt_cli.Text), lvwText, , lvwPartial)
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


Private Sub txtbonif_Change()
cal
End Sub

Private Sub txtbonif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtprecio, txtprecio
End If

End Sub

Private Sub txtcantidad_Change()
cal
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'   grid_fac.TextMatrix(grid_fac.Row, 6) = Nulo_Valor0(arm_llave!ARM_COSTO_ULT) '* grid_fac.TextMatrix(grid_fac.Row, 14)
 If Val(i_codart2.Text) = 0 Then GoTo pasabof
   If Test_Compra(LK_CODCIA, Val(i_codart2.Text), (Val(txtcantidad.Text) * Val(Right(unidad.Text, 20))), False) = False Then
        MsgBox "Cantidad Ingresada , es Mayor que el Promedio de Ventas , Cordinar con el Administrador, No Procede!!!" & Chr(13) & pub_cadena, 48, Pub_Titulo
'        textovar.SetFocus
'        Azul txtcantidad, txtcantidad
'        Exit Sub
   End If
pasabof:
  Azul txtbonif, txtbonif
End If
End Sub

Private Sub txtcondi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   i_codart2.SetFocus
End If

End Sub

Private Sub txtd1_Change()
cal
End Sub

Private Sub txtd1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul txtd2, txtd2
End If

End Sub

Private Sub txtd2_Change()
cal
End Sub

Private Sub txtd2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtd3, txtd3
End If

End Sub

Private Sub txtd3_Change()
cal
End Sub

Private Sub txtd3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtdn1, txtdn1
End If

End Sub

Private Sub txtdn1_Change()
cal
End Sub

Private Sub txtdn1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtdn2, txtdn2
End If

End Sub

Private Sub txtdn2_Change()
cal
End Sub

Private Sub txtdn2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtdn3, txtdn3
End If

End Sub

Private Sub txtdn3_Change()
cal
End Sub

Private Sub txtdn3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtdn4, txtdn4
End If

End Sub

Private Sub txtdn4_2_Change()
cal
End Sub

Private Sub txtdn4_2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Azul txtdn4_3, txtdn4_3
End If

End Sub

Private Sub txtdn4_3_Change()
cal
End Sub

Private Sub txtdn4_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Azul txttotal, txttotal
End If

End Sub

Private Sub txtdn4_Change()
cal
End Sub

Private Sub txtdn4_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
Azul txtdn4_2, txtdn4_2
End If

End Sub

Private Sub txtglosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtcondi.SetFocus
End If
End Sub


Private Sub txtnumfac_KeyPress(KeyAscii As Integer)
Dim consul_relcomp_llave As rdoResultset
Dim CONSUL_PRE_RELCOMP As rdoQuery
If KeyAscii <> 13 Then Exit Sub

pub_cadena = "SELECT * FROM PEDPRO WHERE REL_CODCIA = ? AND REL_CP = ?  AND REL_NUMSER = ? AND REL_NUMFAC = ?   ORDER BY REL_NUMSER, REL_NUMFAC ,REL_NUMSEC"
Set CONSUL_PRE_RELCOMP = CN.CreateQuery("", pub_cadena)
CONSUL_PRE_RELCOMP(0) = LK_CODCIA
CONSUL_PRE_RELCOMP(1) = "P"
CONSUL_PRE_RELCOMP(2) = 0
CONSUL_PRE_RELCOMP(3) = Val(txtnumfac.Text)
Set consul_relcomp_llave = CONSUL_PRE_RELCOMP.OpenResultset(rdOpenKeyset, rdConcurValues)
If consul_relcomp_llave.EOF Then
   cabe
  MsgBox "No Existe el Numero Registrado", 48, Pub_Titulo
  Exit Sub
End If

cabe

txtglosa.Text = Trim(consul_relcomp_llave!rel_GLOSA)
txtcondi.Text = Trim(consul_relcomp_llave!rel_CONDI)
cheliq.Value = consul_relcomp_llave!rel_LIQUIDO
txtrepre.Text = Trim(consul_relcomp_llave!rel_repre)
txt_cli.Text = consul_relcomp_llave!rel_codpro
txt_cli_KeyPress 13

Do Until consul_relcomp_llave.EOF
    grdidet.Rows = grdidet.Rows + 1
    PUB_TIPREG = 43
    PUB_NUMTAB = consul_relcomp_llave!rel_codtip
    SQ_OPER = 1
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
       grdidet.TextMatrix(grdidet.Rows - 1, 1) = tab_llave!TAB_NOMLARGO & String(80, " ") & tab_llave!TAB_NUMTAB
    End If
    grdidet.TextMatrix(grdidet.Rows - 1, 2) = consul_relcomp_llave!rel_codart
    grdidet.TextMatrix(grdidet.Rows - 1, 0) = Format(consul_relcomp_llave!rel_FECHA, "dd/mm/yy")
    If Val(consul_relcomp_llave!rel_codart) <> 0 Then
     SQ_OPER = 1
     PUB_KEY = consul_relcomp_llave!rel_codart
     pu_codcia = consul_relcomp_llave!rel_codcia
     LEER_ART_LLAVE
     
     grdidet.TextMatrix(grdidet.Rows - 1, 3) = Trim(art_LLAVE!art_nombre)
     grdidet.TextMatrix(grdidet.Rows - 1, 4) = consul_relcomp_llave!rel_descri & String(80, " ") & consul_relcomp_llave!rel_equiv
     grdidet.TextMatrix(grdidet.Rows - 1, 5) = Format(consul_relcomp_llave!rel_cantidad / consul_relcomp_llave!rel_equiv, "0.00")
     grdidet.TextMatrix(grdidet.Rows - 1, 6) = consul_relcomp_llave!rel_precio
    End If
    grdidet.TextMatrix(grdidet.Rows - 1, 7) = Trim(consul_relcomp_llave!rel_GLOSA)
    grdidet.TextMatrix(grdidet.Rows - 1, 8) = consul_relcomp_llave!rel_total
    grdidet.TextMatrix(grdidet.Rows - 1, 9) = consul_relcomp_llave!rel_signo_carga
    grdidet.TextMatrix(grdidet.Rows - 1, 10) = "A"
    grdidet.TextMatrix(grdidet.Rows - 1, 11) = consul_relcomp_llave!rel_FECHA
    grdidet.TextMatrix(grdidet.Rows - 1, 12) = consul_relcomp_llave!rel_LIQUIDO
    
    grdidet.TextMatrix(grdidet.Rows - 1, 13) = consul_relcomp_llave!rel_bonif
    grdidet.TextMatrix(grdidet.Rows - 1, 14) = consul_relcomp_llave!rel_d1
    grdidet.TextMatrix(grdidet.Rows - 1, 15) = consul_relcomp_llave!rel_d2
    grdidet.TextMatrix(grdidet.Rows - 1, 16) = consul_relcomp_llave!rel_d3
    grdidet.TextMatrix(grdidet.Rows - 1, 17) = consul_relcomp_llave!rel_dn1
    grdidet.TextMatrix(grdidet.Rows - 1, 18) = consul_relcomp_llave!rel_dn2
    grdidet.TextMatrix(grdidet.Rows - 1, 19) = consul_relcomp_llave!rel_dn3
    grdidet.TextMatrix(grdidet.Rows - 1, 20) = consul_relcomp_llave!rel_dn4
    grdidet.TextMatrix(grdidet.Rows - 1, 21) = consul_relcomp_llave!rel_dn4_2
    grdidet.TextMatrix(grdidet.Rows - 1, 22) = consul_relcomp_llave!rel_dn4_3
    grdidet.TextMatrix(grdidet.Rows - 1, 23) = consul_relcomp_llave!rel_precio_neto
    grdidet.TextMatrix(grdidet.Rows - 1, 24) = Format(consul_relcomp_llave!rel_saldo / consul_relcomp_llave!rel_equiv, "0.00")
    grdidet.TextMatrix(grdidet.Rows - 1, 25) = consul_relcomp_llave!rel_precio_pro
    PSPAR_MULTI(0) = Trim(consul_relcomp_llave!rel_codcia_ref)
    par_multi.Requery
    If par_multi.EOF Then
      grdidet.TextMatrix(grdidet.Rows - 1, 27) = LK_CODCIA
    Else
      grdidet.TextMatrix(grdidet.Rows - 1, 27) = par_multi!par_codcia & " - " & Trim(par_multi!par_nombre)
    End If

 consul_relcomp_llave.MoveNext
Loop
'If Val(flag_signo.Caption) = 1 Then
'   cmdsigno.ListIndex = 0
'ElseIf Val(flag_signo.Caption) = -1 Then
'   cmdsigno.ListIndex = 1
'End If
sum_total


End Sub

Private Sub txtprecio_Change()
If Val(txtprecio.Text) <> 0 Then
 txtprecioneto.Text = 0
 txtprecioneto.Locked = True
 txtd1.Text = 0
 txtd2.Text = 0
 txtd3.Text = 0
 txtd1.Locked = False
 txtd2.Locked = False
 txtd3.Locked = False
Else
 txtprecioneto.Locked = False
 txtd1.Text = 0
 txtd2.Text = 0
 txtd3.Text = 0
 txtd1.Locked = True
 txtd2.Locked = True
 txtd3.Locked = True

End If
cal
End Sub

Private Sub txtprecio_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  Azul txtprecioneto, txtprecioneto
End If
End Sub

Private Sub txtprecioneto_Change()
cal
End Sub

Private Sub txtprecioneto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtd1, txtd1
End If
End Sub

Private Sub txtrepre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtglosa.SetFocus
End If
End Sub

Private Sub txttotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Val(txtmax.Text) <= grdidet.Rows - 1 Then
   MsgBox "llego al maximo tope de Item.", 48, Pub_Titulo
   Azul i_codart2, i_codart2
   Exit Sub
 End If

add_registro
End If
End Sub

Private Sub unidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtcantidad, txtcantidad
End If
End Sub
Public Sub cabe()
grdidet.Clear
grdidet.Cols = 28
grdidet.Rows = 1
grdidet.ColWidth(0) = 800 ' Fecha
grdidet.ColWidth(1) = 0   ' tipo
grdidet.ColWidth(2) = 600 ' codigo internot
grdidet.ColWidth(3) = 1300  ' descrip prod
grdidet.ColWidth(4) = 500  ' unidad
grdidet.ColWidth(5) = 800 ' cantidad
grdidet.ColWidth(6) = 700 ' precio
grdidet.ColWidth(7) = 0 ' glosa
grdidet.ColWidth(8) = 900 ' total
grdidet.ColWidth(9) = 0 ' Signo de cargo
grdidet.ColWidth(10) = 0 ' No Editar
grdidet.ColWidth(11) = 0 ' FECHA DE PROCESO
grdidet.ColWidth(12) = 0 ' FLAG DE CIERRE

grdidet.ColWidth(13) = 0  ' cantidad bonifi
grdidet.ColWidth(14) = 0  ' des3
grdidet.ColWidth(15) = 0  ' des3
grdidet.ColWidth(16) = 0  ' des3
grdidet.ColWidth(17) = 0  ' des no afect 1
grdidet.ColWidth(18) = 0  ' des no afect 2
grdidet.ColWidth(19) = 0  ' des no afect 3
grdidet.ColWidth(20) = 0  ' des no afect 4
grdidet.ColWidth(21) = 0  ' des no afect 4 - 2
grdidet.ColWidth(22) = 0  ' des no afect 4 - 3
grdidet.ColWidth(23) = 0  ' des no afect 4 - 3
grdidet.ColWidth(24) = 900  ' saldo de pendiente
grdidet.ColWidth(25) = 0  ' PRECIO NETO CON DESCTO.
grdidet.ColWidth(26) = 0  ' condicion de pago.
grdidet.ColWidth(27) = 800  ' Codigo de empresa

grdidet.TextMatrix(0, 0) = "Fecha"
grdidet.TextMatrix(0, 1) = "Tipo"
grdidet.TextMatrix(0, 2) = "Codigo"
grdidet.TextMatrix(0, 3) = "Descripción"
grdidet.TextMatrix(0, 4) = "Unid."
grdidet.TextMatrix(0, 5) = "Cantid."
grdidet.TextMatrix(0, 6) = "Precio"
grdidet.TextMatrix(0, 7) = "Glosa"
grdidet.TextMatrix(0, 8) = "Total"

grdidet.TextMatrix(0, 13) = "CantBoni" ' cantidad bonifi
grdidet.TextMatrix(0, 14) = "Desc1" ' des3
grdidet.TextMatrix(0, 15) = "Desc2" ' des3
grdidet.TextMatrix(0, 16) = "Desc3" 'des3
grdidet.TextMatrix(0, 17) = "DnAf1" ' des no afect 1
grdidet.TextMatrix(0, 18) = "DnAf2" ' des no afect 2
grdidet.TextMatrix(0, 19) = "DnAf3" ' des no afect 3
grdidet.TextMatrix(0, 20) = "DnAf4" ' des no afect 4
grdidet.TextMatrix(0, 23) = "Precio Neto"
grdidet.TextMatrix(0, 24) = "Saldo"
grdidet.TextMatrix(0, 27) = "Empresa"




End Sub

Public Sub CABE_MUESTRA()

GRID_MST.Clear
GRID_MST.Cols = 8
GRID_MST.Rows = 1
GRID_MST.ColWidth(0) = 1000 ' tipo
GRID_MST.ColWidth(1) = 900 ' codigo internot
GRID_MST.ColWidth(2) = 1300  ' descrip prod
GRID_MST.ColWidth(3) = 900  ' unidad
GRID_MST.ColWidth(4) = 900 ' cantidad
GRID_MST.ColWidth(5) = 900 ' precio
GRID_MST.ColWidth(6) = 2500 ' glosa
GRID_MST.ColWidth(7) = 900 ' total


7 GRID_MST.TextMatrix(0, 0) = "Tipo"
GRID_MST.TextMatrix(0, 1) = "Codigo"
GRID_MST.TextMatrix(0, 2) = "Descripción"
GRID_MST.TextMatrix(0, 3) = "Unidad"
GRID_MST.TextMatrix(0, 4) = "Cantidad"
GRID_MST.TextMatrix(0, 5) = "Precio"
GRID_MST.TextMatrix(0, 6) = "Glosa"
GRID_MST.TextMatrix(0, 7) = "Total"


End Sub

Public Sub add_registro()
If Trim(cmbemp.Text) = "" Then
  MsgBox "Seleccionar Empresa Destino ", 48, Pub_Titulo
  cmbemp.SetFocus
  Exit Sub
End If
grdidet.Rows = grdidet.Rows + 1
grdidet.TextMatrix(grdidet.Rows - 1, 0) = Format(LK_FECHA_DIA, "dd/mm/yy")
grdidet.TextMatrix(grdidet.Rows - 1, 1) = Trim(cmbTipo.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 2) = Trim(i_codart2.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 3) = Trim(i_nomarti.Caption)
grdidet.TextMatrix(grdidet.Rows - 1, 4) = Trim(unidad.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 5) = Val(txtcantidad.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 6) = Trim(txtprecio.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 7) = Trim(txtglosa.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 26) = Trim(txtcondi.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 8) = Val(txttotal.Text)
'grdidet.TextMatrix(grdidet.Rows - 1, 9) = Trim(Left(cmdsigno.Text, 2))
grdidet.TextMatrix(grdidet.Rows - 1, 11) = LK_FECHA_DIA
grdidet.TextMatrix(grdidet.Rows - 1, 12) = cheliq.Value

grdidet.TextMatrix(grdidet.Rows - 1, 13) = txtbonif.Text
grdidet.TextMatrix(grdidet.Rows - 1, 14) = txtd1.Text
grdidet.TextMatrix(grdidet.Rows - 1, 15) = txtd2.Text
grdidet.TextMatrix(grdidet.Rows - 1, 16) = txtd3.Text
grdidet.TextMatrix(grdidet.Rows - 1, 17) = txtdn1.Text
grdidet.TextMatrix(grdidet.Rows - 1, 18) = txtdn2.Text
grdidet.TextMatrix(grdidet.Rows - 1, 19) = txtdn3.Text
grdidet.TextMatrix(grdidet.Rows - 1, 20) = txtdn4.Text
grdidet.TextMatrix(grdidet.Rows - 1, 21) = txtdn4_2.Text
grdidet.TextMatrix(grdidet.Rows - 1, 22) = txtdn4_3.Text
grdidet.TextMatrix(grdidet.Rows - 1, 23) = txtprecioneto.Text
grdidet.TextMatrix(grdidet.Rows - 1, 24) = Val(txtcantidad.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 25) = txtprecio.Tag
grdidet.TextMatrix(grdidet.Rows - 1, 27) = Trim(cmbemp.Text)
grdidet.ColAlignment(27) = 1



    


'i_codart2.Text = ""
'i_nomarti.Caption = ""
'unidad.Clear
cmbemp.ListIndex = -1
txtcantidad.Text = ""
txtprecio.Text = ""
txtprecio.Tag = ""
txttotal.Text = ""
i_codart2.SetFocus
txtbonif.Text = ""
txtd1.Text = ""
txtd2.Text = ""
txtd3.Text = ""
txtdn1.Text = ""
txtdn2.Text = ""
txtdn3.Text = ""
txtdn4.Text = ""
txtdn4_2.Text = ""
txtdn4_3.Text = ""
sum_total
'cmbtipo.SetFocus
End Sub

Public Sub cal()
Dim es_prec As Currency
If Val(txtprecioneto.Text) <> 0 Then
 es_prec = Val(txtprecioneto.Text)
Else
 es_prec = (Val(txtprecio.Text) - (Val(txtprecio.Text) * (Val(txtd1.Text) / 100)))
 es_prec = (es_prec - (Val(es_prec) * (Val(txtd2.Text) / 100)))
 es_prec = (es_prec - (Val(es_prec) * (Val(txtd3.Text) / 100)))
End If
txtprecio.Tag = Val(es_prec)
txttotal.Text = Format(Val(es_prec) * Val(txtcantidad.Text), "0.00")

End Sub
Public Sub sum_total()
Dim es_prec As Currency
For fila = 1 To grdidet.Rows - 1
es_prec = es_prec + Val(grdidet.TextMatrix(fila, 8))
Next fila
lbltotal.Caption = Format(Val(es_prec), "##,##0.00")

End Sub

Public Sub cabe_detalle()
grdidetall.Clear
grdidetall.Cols = 7
grdidetall.Rows = 1
grdidetall.ColWidth(0) = 900 ' Fecha
grdidetall.ColWidth(1) = 700 ' SERIE INT
grdidetall.ColWidth(2) = 900  ' NUMERO INT
grdidetall.ColWidth(3) = 2100  ' PROVEEDOR
grdidetall.ColWidth(4) = 800 ' unidad
grdidetall.ColWidth(5) = 800 ' cantidad
grdidetall.ColWidth(6) = 800 ' precio

grdidetall.TextMatrix(0, 0) = "Fecha"
grdidetall.TextMatrix(0, 1) = "Ser/Int"
grdidetall.TextMatrix(0, 2) = "Fac/Int"
grdidetall.TextMatrix(0, 3) = "Proveedor"
grdidetall.TextMatrix(0, 4) = "Unidad"
grdidetall.TextMatrix(0, 5) = "Cantidad"
grdidetall.TextMatrix(0, 6) = "Precio"



End Sub


Public Sub llenanumfac()
Dim ps_ultimo As rdoResultset
pub_cadena = "SELECT top 1 rel_numfac FROM PEDPRO WHERE REL_CODCIA = '" & LK_CODCIA & "' AND REL_CP = 'P' order by rel_numfac desc "
Set ps_ultimo = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not ps_ultimo.EOF Then
txtnumfac.Text = Val(ps_ultimo!rel_numfac)
End If

End Sub
Public Function Test_Compra(wr_codcia As String, wr_codart As Currency, wr_cantidad_min As Currency, ForzarTrue As Boolean) As Boolean
Dim wr_descrip As String
Dim wr_cant_compra As Currency
Dim wr_vstock As Currency
Dim wr_vmaxi As Currency
Dim wr_equiv As Integer
If ForzarTrue = True Then
  Test_Compra = True
  Exit Function
End If

If wr_codcia = "01" Or wr_codcia = "03" Or wr_codcia = "07" Or wr_codcia = "10" Then
  Test_Compra = False
Else
  Test_Compra = True
  Exit Function
End If
wr_cant_compra = wr_cantidad_min
wr_equiv = 0
wr_descrip = ""

pub_cadena = "select pre_unidad, pre_equiv from precios where pre_codcia = '" & wr_codcia & "' and pre_codart = " & wr_codart & " and pre_flag_unidad = 'A'"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If X.EOF Then
Exit Function
Else
wr_descrip = Trim(X!PRE_UNIDAD)
wr_equiv = Val(X!PRE_EQUIV)
End If

pub_cadena = "select * from v_stockgen where arm_codart = " & wr_codart & " "
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If X.EOF Then
  wr_vstock = 0
Else
  wr_vstock = Format(Val(Nulo_Valor0(X!stockGen)) / wr_equiv, "0.00")
End If
pub_cadena = "select * from v_stockmaxgen where art_key = " & wr_codart & " "
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If X.EOF Then
  wr_vmaxi = 0
Else
  wr_vmaxi = Format(Val(Nulo_Valor0(X!StockMaxGen) / wr_equiv), "0.00")
End If
wr_cant_compra = Format(Val(Val(wr_cant_compra) / wr_equiv), "0.00")
pub_cadena = "Unidad : " & Trim(wr_descrip) & " " & Chr(13) & "Promedio( 1 Meses): " & Val(wr_vmaxi * 1) & Chr(13) & "StockGeneral : " & Val(wr_vstock) & Chr(13) & "Lo que estas comprando : " & Val(wr_cant_compra) & Chr(13) & "Solo Permite : " & Format(Val(wr_vmaxi * 1) - Val(wr_vstock), "0.00") & " " & Trim(wr_descrip)
If ((Val(wr_cant_compra)) + Val(wr_vstock)) > (Val(wr_vmaxi) * 1) Then
  Test_Compra = False
Else
  Test_Compra = True
End If

End Function


VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form frmCLI 
   Caption         =   "Clientes / Proveedores"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   1080
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "client.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7410
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstVen 
      Height          =   1185
      Left            =   6480
      Style           =   1  'Checkbox
      TabIndex        =   137
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdprocli 
      Caption         =   "Tipo Cliente"
      Height          =   375
      Left            =   10560
      TabIndex        =   136
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame fraplaca 
      Caption         =   "Opcion para Grifos - Descto. Especial"
      Height          =   855
      Left            =   2760
      TabIndex        =   125
      Top             =   7200
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdmante 
         Caption         =   "Editar &Placas"
         Height          =   420
         Left            =   3840
         TabIndex        =   128
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmddescto 
         Caption         =   "&Editar Desc."
         Height          =   420
         Left            =   2280
         TabIndex        =   126
         Top             =   240
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid grid_des 
         Height          =   1095
         Left            =   120
         TabIndex        =   127
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1931
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   8421376
      End
   End
   Begin VB.Frame F14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   35
      Top             =   7560
      Visible         =   0   'False
      Width           =   4125
      Begin MSFlexGridLib.MSFlexGrid ListExiste 
         Height          =   1455
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   4
         BackColorBkg    =   8421376
      End
      Begin VB.CommandButton CmdEscapa 
         Caption         =   "E&scapar"
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
         Left            =   6720
         TabIndex        =   40
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdconfirma 
         Caption         =   "Con&firmar"
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
         Left            =   5280
         TabIndex        =   39
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Op 
         Caption         =   "Ignorar la Lista "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   38
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton Op 
         Caption         =   "Seleccionar uno de la Lista "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   1800
         Width           =   2535
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   495
      Left            =   9240
      TabIndex        =   28
      Top             =   7680
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      Appearance      =   1
      NumItems        =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   0
      TabIndex        =   41
      Tag             =   "32"
      Top             =   4425
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      TabHeight       =   450
      ForeColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Direccion Fiscal"
      TabPicture(0)   =   "client.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblnom(7)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblnom(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblnom(9)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblnom(10)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblnom(11)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblnom(31)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblnom(22)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblnom(28)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblnom(27)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbldocr"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtSubZona"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtZona"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Txtnumdir"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Txtdireccion"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtZonaNew"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtLugarCasa"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtregpublico2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "letraotorgado"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "otrocontrato"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtciarela"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Direccion Almacen"
      TabPicture(1)   =   "client.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "cboProvincia"
      Tab(1).Control(2)=   "cboDireccion"
      Tab(1).Control(3)=   "txtNumDirTrabajo"
      Tab(1).Control(4)=   "TxtLugarTrab"
      Tab(1).Control(5)=   "txtpropiedad2"
      Tab(1).Control(6)=   "TxtSubZonaTrabajo"
      Tab(1).Control(7)=   "TxtZonaTrabajo"
      Tab(1).Control(8)=   "txtDirTrabajo"
      Tab(1).Control(9)=   "label(36)"
      Tab(1).Control(10)=   "lblnom(15)"
      Tab(1).Control(11)=   "lblnom(32)"
      Tab(1).Control(12)=   "lblnom(21)"
      Tab(1).Control(13)=   "lblnom(17)"
      Tab(1).Control(14)=   "lblnom(18)"
      Tab(1).Control(15)=   "lblnom(14)"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Otras Opciones"
      TabPicture(2)   =   "client.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblnom(29)"
      Tab(2).Control(1)=   "lblnom(33)"
      Tab(2).Control(2)=   "lblnom(35)"
      Tab(2).Control(3)=   "g_fechafac"
      Tab(2).Control(4)=   "lblnom(23)"
      Tab(2).Control(5)=   "lblnom(24)"
      Tab(2).Control(6)=   "lblnom(34)"
      Tab(2).Control(7)=   "lblnom(25)"
      Tab(2).Control(8)=   "lblnom(20)"
      Tab(2).Control(9)=   "g_diasfac"
      Tab(2).Control(10)=   "LblDatos(20)"
      Tab(2).Control(11)=   "Label9"
      Tab(2).Control(12)=   "lblnom(19)"
      Tab(2).Control(13)=   "txtNucleo"
      Tab(2).Control(14)=   "txtDTX"
      Tab(2).Control(15)=   "txtpordes"
      Tab(2).Control(16)=   "t_diascred"
      Tab(2).Control(17)=   "t_diasfac"
      Tab(2).Control(18)=   "t_fechafac"
      Tab(2).Control(19)=   "txtprog"
      Tab(2).Control(20)=   "txtautovaluo"
      Tab(2).Control(21)=   "txtregpublico1"
      Tab(2).Control(22)=   "copia"
      Tab(2).Control(23)=   "ListBloqueos"
      Tab(2).Control(24)=   "lisdescto"
      Tab(2).Control(25)=   "txtpropiedad1"
      Tab(2).ControlCount=   26
      Begin VB.TextBox txtciarela 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9360
         TabIndex        =   146
         Text            =   "00"
         Top             =   2280
         Width           =   495
      End
      Begin VB.CheckBox otrocontrato 
         Caption         =   "-"
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
         Left            =   7800
         TabIndex        =   143
         Top             =   1800
         Visible         =   0   'False
         WhatsThisHelpID =   7
         Width           =   375
      End
      Begin VB.CheckBox letraotorgado 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7800
         TabIndex        =   139
         Top             =   1560
         Visible         =   0   'False
         WhatsThisHelpID =   10
         Width           =   435
      End
      Begin VB.TextBox txtpropiedad1 
         DataField       =   "ART_MARGEN"
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
         Height          =   285
         Left            =   -74880
         MaxLength       =   30
         TabIndex        =   133
         Top             =   1770
         WhatsThisHelpID =   8
         Width           =   540
      End
      Begin VB.TextBox txtregpublico2 
         DataField       =   "ART_MARGEN"
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
         Height          =   525
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   131
         Top             =   1875
         Visible         =   0   'False
         WhatsThisHelpID =   12
         Width           =   7140
      End
      Begin VB.ListBox lisdescto 
         Height          =   1635
         Left            =   -68160
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   129
         Top             =   600
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         Height          =   2055
         Left            =   -66360
         TabIndex        =   119
         Top             =   360
         Width           =   1455
         Begin VB.CommandButton cmdDireccion 
            Caption         =   "Agregar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   240
            TabIndex        =   122
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Modificar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   240
            TabIndex        =   121
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Eliminar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   240
            TabIndex        =   120
            Top             =   1440
            Width           =   975
         End
      End
      Begin VB.ComboBox cboProvincia 
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
         Height          =   315
         Left            =   -69120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1680
         WhatsThisHelpID =   6
         Width           =   2655
      End
      Begin VB.ComboBox cboDireccion 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   480
         Width           =   8295
      End
      Begin VB.ListBox ListBloqueos 
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
         ItemData        =   "client.frx":0496
         Left            =   -73560
         List            =   "client.frx":0498
         TabIndex        =   116
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton copia 
         Caption         =   "Copia a Otra Cia."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71520
         TabIndex        =   114
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtregpublico1 
         DataField       =   "ART_MARGEN"
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
         Height          =   285
         Left            =   -74880
         TabIndex        =   111
         Top             =   1200
         Visible         =   0   'False
         WhatsThisHelpID =   9
         Width           =   2220
      End
      Begin VB.TextBox txtNumDirTrabajo 
         DataSource      =   "3"
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
         Left            =   -67320
         MaxLength       =   4
         TabIndex        =   23
         Top             =   1080
         WhatsThisHelpID =   3
         Width           =   855
      End
      Begin VB.ComboBox TxtLugarTrab 
         DataSource      =   "2"
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
         Left            =   -74760
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1080
         WhatsThisHelpID =   2
         Width           =   1695
      End
      Begin VB.TextBox txtpropiedad2 
         DataField       =   "ART_MARGEN"
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
         Height          =   285
         Left            =   -74760
         MaxLength       =   30
         TabIndex        =   27
         Top             =   2230
         WhatsThisHelpID =   11
         Width           =   8100
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
         Left            =   -74745
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1680
         WhatsThisHelpID =   6
         Width           =   2775
      End
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
         Left            =   -71880
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1680
         WhatsThisHelpID =   5
         Width           =   2655
      End
      Begin VB.TextBox txtDirTrabajo 
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
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   22
         Top             =   1080
         WhatsThisHelpID =   1
         Width           =   5535
      End
      Begin VB.TextBox txtautovaluo 
         DataField       =   "ART_MARGEN"
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
         Height          =   285
         Left            =   -74880
         TabIndex        =   100
         Top             =   600
         Visible         =   0   'False
         WhatsThisHelpID =   16
         Width           =   2220
      End
      Begin VB.TextBox txtprog 
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
         Left            =   -72480
         MaxLength       =   1
         TabIndex        =   98
         Top             =   1200
         Visible         =   0   'False
         WhatsThisHelpID =   19
         Width           =   495
      End
      Begin VB.TextBox t_fechafac 
         DataSource      =   "4"
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
         Left            =   -71160
         MaxLength       =   12
         TabIndex        =   93
         Top             =   1800
         Visible         =   0   'False
         WhatsThisHelpID =   4
         Width           =   1215
      End
      Begin VB.TextBox t_diasfac 
         DataSource      =   "4"
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
         Left            =   -71160
         MaxLength       =   12
         TabIndex        =   92
         Top             =   1200
         Visible         =   0   'False
         WhatsThisHelpID =   4
         Width           =   1215
      End
      Begin VB.TextBox t_diascred 
         DataSource      =   "4"
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
         Left            =   -69720
         MaxLength       =   12
         TabIndex        =   91
         Top             =   1200
         Visible         =   0   'False
         WhatsThisHelpID =   4
         Width           =   975
      End
      Begin VB.TextBox txtpordes 
         DataField       =   "ART_COSPRO"
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
         Height          =   300
         Left            =   -72480
         MaxLength       =   12
         TabIndex        =   89
         Text            =   " "
         Top             =   1760
         Width           =   1125
      End
      Begin VB.TextBox txtDTX 
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
         Left            =   -71880
         MaxLength       =   1
         TabIndex        =   86
         Top             =   600
         Visible         =   0   'False
         WhatsThisHelpID =   19
         Width           =   495
      End
      Begin VB.TextBox txtNucleo 
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
         Left            =   -72480
         MaxLength       =   2
         TabIndex        =   85
         Top             =   600
         Visible         =   0   'False
         WhatsThisHelpID =   13
         Width           =   495
      End
      Begin VB.ComboBox TxtLugarCasa 
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
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox txtZonaNew 
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
         Left            =   255
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox Txtdireccion 
         DataField       =   "ART_PLANCHA"
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
         Height          =   285
         Left            =   2160
         TabIndex        =   15
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox Txtnumdir 
         DataField       =   "ART_COSPRO"
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
         Height          =   285
         Left            =   8640
         MaxLength       =   4
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox TxtZona 
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
         Left            =   3720
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1200
         Width           =   3015
      End
      Begin VB.ComboBox TxtSubZona 
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
         Left            =   6960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label11 
         Caption         =   "Compañia de Relación:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7560
         TabIndex        =   147
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lbldocr 
         Caption         =   "Documentación  en Regla"
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
         Height          =   495
         Left            =   8280
         TabIndex        =   145
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblnom 
         Caption         =   "Contrato a Plazo"
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
         Height          =   315
         Index           =   27
         Left            =   8160
         TabIndex        =   144
         Tag             =   "28"
         Top             =   1800
         Visible         =   0   'False
         WhatsThisHelpID =   7
         Width           =   195
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Letra Otorgado"
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
         Height          =   195
         Index           =   28
         Left            =   8280
         TabIndex        =   140
         Tag             =   "29"
         Top             =   1560
         Visible         =   0   'False
         WhatsThisHelpID =   10
         Width           =   1530
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
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
         Index           =   19
         Left            =   -74880
         TabIndex        =   134
         Tag             =   "20"
         Top             =   1560
         Visible         =   0   'False
         WhatsThisHelpID =   8
         Width           =   1005
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
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
         Index           =   22
         Left            =   240
         TabIndex        =   132
         Tag             =   "23"
         Top             =   1650
         Visible         =   0   'False
         WhatsThisHelpID =   12
         Width           =   945
      End
      Begin VB.Label Label9 
         Caption         =   "Otras Lista de Descuento."
         Height          =   255
         Left            =   -68040
         TabIndex        =   130
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
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
         Index           =   36
         Left            =   -74760
         TabIndex        =   118
         Tag             =   "19"
         Top             =   1440
         WhatsThisHelpID =   6
         Width           =   450
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Bloqueos :"
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
         Index           =   20
         Left            =   -74760
         TabIndex        =   117
         Top             =   2160
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label g_diasfac 
         AutoSize        =   -1  'True
         Caption         =   "Dias p' Factr."
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
         Height          =   195
         Left            =   -71160
         TabIndex        =   94
         Tag             =   "25"
         Top             =   960
         Visible         =   0   'False
         WhatsThisHelpID =   15
         Width           =   1155
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Propiedades 2"
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
         Height          =   195
         Index           =   20
         Left            =   -74880
         TabIndex        =   112
         Tag             =   "21"
         Top             =   960
         Visible         =   0   'False
         WhatsThisHelpID =   9
         Width           =   1230
      End
      Begin VB.Label lblnom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Nº. Dir."
         DataSource      =   "3"
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
         Index           =   15
         Left            =   -67200
         TabIndex        =   110
         Tag             =   "16"
         Top             =   840
         WhatsThisHelpID =   3
         Width           =   675
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Lugar"
         DataSource      =   "2"
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
         Index           =   32
         Left            =   -74760
         TabIndex        =   109
         Tag             =   "33"
         Top             =   840
         WhatsThisHelpID =   2
         Width           =   495
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Referencias:"
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
         Index           =   21
         Left            =   -74760
         TabIndex        =   108
         Tag             =   "22"
         Top             =   2040
         WhatsThisHelpID =   11
         Width           =   1095
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
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
         Index           =   17
         Left            =   -71880
         TabIndex        =   107
         Tag             =   "18"
         Top             =   1440
         WhatsThisHelpID =   5
         Width           =   615
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
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
         Index           =   18
         Left            =   -69000
         TabIndex        =   106
         Tag             =   "19"
         Top             =   1440
         WhatsThisHelpID =   6
         Width           =   810
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Dirección Almacen :"
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
         Index           =   14
         Left            =   -72960
         TabIndex        =   105
         Tag             =   "15"
         Top             =   840
         WhatsThisHelpID =   1
         Width           =   1725
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Propiedades 1"
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
         Height          =   195
         Index           =   25
         Left            =   -74880
         TabIndex        =   101
         Tag             =   "26"
         Top             =   360
         Visible         =   0   'False
         WhatsThisHelpID =   16
         Width           =   1230
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Programado :"
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
         Height          =   195
         Index           =   34
         Left            =   -72480
         TabIndex        =   99
         Tag             =   "35"
         Top             =   960
         Visible         =   0   'False
         WhatsThisHelpID =   13
         Width           =   1140
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "."
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
         Height          =   195
         Index           =   24
         Left            =   -69480
         TabIndex        =   97
         Tag             =   "25"
         Top             =   2040
         Visible         =   0   'False
         WhatsThisHelpID =   15
         Width           =   75
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Placa Auto:"
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
         Height          =   195
         Index           =   23
         Left            =   -69720
         TabIndex        =   96
         Tag             =   "24"
         Top             =   960
         Visible         =   0   'False
         WhatsThisHelpID =   14
         Width           =   1005
      End
      Begin VB.Label g_fechafac 
         AutoSize        =   -1  'True
         Caption         =   "Fecha p' Factr."
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
         Height          =   195
         Left            =   -71160
         TabIndex        =   95
         Tag             =   "25"
         Top             =   1560
         Visible         =   0   'False
         WhatsThisHelpID =   15
         Width           =   1305
      End
      Begin VB.Label lblnom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Descto. Fact."
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
         Height          =   195
         Index           =   35
         Left            =   -72405
         TabIndex        =   90
         Tag             =   "9"
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Op 2"
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
         Height          =   195
         Index           =   33
         Left            =   -71760
         TabIndex        =   88
         Tag             =   "34"
         Top             =   360
         Visible         =   0   'False
         WhatsThisHelpID =   19
         Width           =   420
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Op 1"
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
         Height          =   195
         Index           =   29
         Left            =   -72480
         TabIndex        =   87
         Tag             =   "30"
         Top             =   360
         Visible         =   0   'False
         WhatsThisHelpID =   13
         Width           =   420
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Lugar"
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
         Index           =   31
         Left            =   240
         TabIndex        =   84
         Tag             =   "32"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
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
         Index           =   11
         Left            =   240
         TabIndex        =   83
         Tag             =   "12"
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
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
         Index           =   10
         Left            =   6960
         TabIndex        =   82
         Tag             =   "11"
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
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
         Index           =   9
         Left            =   3720
         TabIndex        =   81
         Tag             =   "10"
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Dirección  :"
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
         Index           =   6
         Left            =   2160
         TabIndex        =   80
         Tag             =   "7"
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label lblnom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "N° Dir."
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
         Index           =   7
         Left            =   8640
         TabIndex        =   79
         Tag             =   "8"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "CLIENTES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4455
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   10215
      Begin VB.ComboBox cmbtipocli 
         Height          =   315
         Left            =   7290
         Style           =   2  'Dropdown List
         TabIndex        =   148
         Top             =   4035
         Width           =   2445
      End
      Begin VB.TextBox t_ciarela 
         Height          =   285
         Left            =   3960
         TabIndex        =   141
         Text            =   "00"
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cboDiaVisita 
         Height          =   315
         Left            =   3855
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   3465
         Width           =   2985
      End
      Begin VB.ComboBox Cmbcate 
         Height          =   315
         Left            =   5400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   123
         Top             =   4035
         Width           =   1575
      End
      Begin VB.TextBox txttelefono2 
         DataSource      =   "4"
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
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   6
         Top             =   2295
         WhatsThisHelpID =   4
         Width           =   1575
      End
      Begin VB.TextBox txttelefono1 
         DataField       =   "ART_COSPRO"
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
         Height          =   300
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   4
         Text            =   " "
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox txtsubgrupo 
         Height          =   315
         Left            =   2805
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4035
         Width           =   2415
      End
      Begin VB.ComboBox cmbgrupo 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4030
         Width           =   2415
      End
      Begin VB.TextBox txtprendas 
         DataField       =   "ART_COLOR"
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
         Height          =   285
         Left            =   5400
         MaxLength       =   30
         TabIndex        =   11
         Top             =   3480
         WhatsThisHelpID =   17
         Width           =   1380
      End
      Begin VB.TextBox txtlimite 
         DataField       =   "ART_COSPRO"
         DataSource      =   "Data1"
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
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   9
         Top             =   2880
         Width           =   1365
      End
      Begin VB.ComboBox txtestado 
         Height          =   315
         ItemData        =   "client.frx":049A
         Left            =   7320
         List            =   "client.frx":04A4
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   3460
         Width           =   2415
      End
      Begin VB.TextBox txtRUCempresa 
         DataField       =   "ART_UNIDAD"
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
         Height          =   300
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   8
         Top             =   2880
         Width           =   1245
      End
      Begin VB.TextBox tcuenta22 
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
         Left            =   5400
         TabIndex        =   70
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TxtEmpresa 
         DataField       =   "ART_MARGEN"
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
         Height          =   285
         Left            =   240
         MaxLength       =   40
         TabIndex        =   10
         Top             =   3480
         Width           =   3375
      End
      Begin VB.OptionButton OptNombre 
         Caption         =   "Por la Razon Razón Social."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Index           =   0
         Left            =   7440
         TabIndex        =   68
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton OptNombre 
         Caption         =   "Por el Gerente o Rep. Legal."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   1
         Left            =   7440
         TabIndex        =   67
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton OptNombre 
         Caption         =   "Por el Contacto."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Index           =   2
         Left            =   7440
         TabIndex        =   66
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Txtesposa 
         DataField       =   "ART_IGV"
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
         Height          =   285
         Left            =   240
         MaxLength       =   40
         TabIndex        =   7
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox txtRUCesposa 
         DataField       =   "ART_UNIDAD"
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
         Height          =   300
         Left            =   255
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2295
         Width           =   1695
      End
      Begin VB.ComboBox CmbCGP 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         ItemData        =   "client.frx":04BB
         Left            =   7440
         List            =   "client.frx":04C5
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   450
         Width           =   2415
      End
      Begin VB.TextBox tcuenta 
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
         Left            =   5400
         TabIndex        =   56
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox tcuenta2 
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
         Left            =   5400
         TabIndex        =   55
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox condi 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtauto1 
         DataField       =   "ART_PLANCHA"
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
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   53
         Top             =   2880
         Visible         =   0   'False
         WhatsThisHelpID =   14
         Width           =   900
      End
      Begin VB.TextBox txtauto2 
         DataField       =   "ART_COSPRO"
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
         Height          =   285
         Left            =   8640
         MaxLength       =   30
         TabIndex        =   52
         Top             =   2880
         Visible         =   0   'False
         WhatsThisHelpID =   15
         Width           =   1020
      End
      Begin VB.TextBox txtRUCesposo 
         DataField       =   "ART_UNIDAD"
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
         Height          =   300
         Left            =   240
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtesposo 
         DataField       =   "ART_COSTO"
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
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox txt_key 
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
         Height          =   285
         Left            =   240
         MaxLength       =   8
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtnombre 
         DataField       =   "ART_NOMBRE"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo Cliente"
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
         Height          =   210
         Left            =   7305
         TabIndex        =   149
         Top             =   3825
         Width           =   1845
      End
      Begin VB.Label Label10 
         Caption         =   "Cia Relac."
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
         Left            =   3960
         TabIndex        =   142
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "División"
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
         Index           =   36
         Left            =   5400
         TabIndex        =   124
         Tag             =   "37"
         Top             =   3840
         Width           =   690
      End
      Begin VB.Label lblnom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Fax."
         DataSource      =   "4"
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
         Index           =   16
         Left            =   2160
         TabIndex        =   113
         Tag             =   "17"
         Top             =   2040
         Visible         =   0   'False
         WhatsThisHelpID =   4
         Width           =   405
      End
      Begin VB.Label lblnom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Telefono :"
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
         Index           =   8
         Left            =   2160
         TabIndex        =   104
         Tag             =   "9"
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Clase de Negocio"
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
         Index           =   13
         Left            =   2790
         TabIndex        =   103
         Tag             =   "14"
         Top             =   3840
         Width           =   1515
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Negocio"
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
         Index           =   12
         Left            =   240
         TabIndex        =   102
         Tag             =   "13"
         Top             =   3840
         Width           =   1425
      End
      Begin VB.Label Label8 
         Caption         =   "Opciones para busqueda"
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
         Height          =   375
         Left            =   7440
         TabIndex        =   78
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Opciones para sus Creditos"
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
         Left            =   7320
         TabIndex        =   77
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor ."
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
         Index           =   26
         Left            =   4080
         TabIndex        =   76
         Tag             =   "27"
         Top             =   3240
         Visible         =   0   'False
         WhatsThisHelpID =   17
         Width           =   945
      End
      Begin VB.Label lbllimite 
         Caption         =   "Limite de Cred."
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
         Left            =   5400
         TabIndex        =   75
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Estado :"
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
         Left            =   7320
         TabIndex        =   74
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label lblnom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "D.N.I."
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
         Index           =   5
         Left            =   3840
         TabIndex        =   72
         Tag             =   "6"
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Cta. Nat 2 :"
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
         Left            =   5400
         TabIndex        =   71
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Contacto:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   69
         Tag             =   "5"
         Top             =   3240
         Width           =   840
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Gerente / Representate Legal"
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
         Index           =   2
         Left            =   240
         TabIndex        =   65
         Tag             =   "3"
         Top             =   2640
         Width           =   2565
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "D.N.I.:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   64
         Tag             =   "4"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Determinar:"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   7440
         TabIndex        =   63
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lcuenta 
         Caption         =   "Cta. Activo:"
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
         Left            =   5400
         TabIndex        =   61
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cta. Nat 1 :"
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
         Left            =   5400
         TabIndex        =   60
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label g_fechacred 
         AutoSize        =   -1  'True
         Caption         =   "Dias Cred.:"
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
         Left            =   7320
         TabIndex        =   59
         Tag             =   "25"
         Top             =   2640
         Visible         =   0   'False
         WhatsThisHelpID =   15
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Condición"
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
         Left            =   7320
         TabIndex        =   58
         Tag             =   "25"
         Top             =   2040
         Visible         =   0   'False
         WhatsThisHelpID =   15
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Otr.Dias :"
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
         Left            =   8760
         TabIndex        =   57
         Tag             =   "25"
         Top             =   2640
         Visible         =   0   'False
         WhatsThisHelpID =   15
         Width           =   825
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "RUC "
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
         Index           =   1
         Left            =   240
         TabIndex        =   51
         Tag             =   "2"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Nombre / Razon Social"
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
         Index           =   0
         Left            =   240
         TabIndex        =   50
         Tag             =   "1"
         Top             =   840
         Width           =   1995
      End
      Begin VB.Label lblvar 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
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
         Index           =   0
         Left            =   240
         TabIndex        =   49
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion de Busqueda :"
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
         Index           =   30
         Left            =   2160
         TabIndex        =   48
         Tag             =   "31"
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdcontab 
      Caption         =   "Relacionar a Contabilidad"
      Height          =   375
      Left            =   4440
      TabIndex        =   45
      Top             =   7560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdcontab2 
      Caption         =   "Relacionar a Contabilidad"
      Height          =   375
      Left            =   960
      TabIndex        =   43
      Top             =   7560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   10440
      Picture         =   "client.frx":04EE
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4080
      Width           =   1300
   End
   Begin VB.Timer Parpadea 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   360
      Top             =   7080
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
      Left            =   10440
      Picture         =   "client.frx":0930
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5160
      Width           =   1300
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
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
      Left            =   10440
      Picture         =   "client.frx":0A7A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   840
      Width           =   1300
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
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
      Left            =   10440
      Picture         =   "client.frx":0BC4
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3000
      Width           =   1300
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
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
      Left            =   10440
      Picture         =   "client.frx":1006
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1920
      Width           =   1300
   End
   Begin ComctlLib.ProgressBar PB2 
      Height          =   135
      Left            =   5040
      TabIndex        =   42
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   0
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   195
      Left            =   10440
      TabIndex        =   44
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label fechahora 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   10440
      TabIndex        =   138
      Top             =   120
      Width           =   1335
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
      Left            =   10320
      TabIndex        =   46
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Index           =   5
      Left            =   10320
      TabIndex        =   115
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label LblMensaje 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   31
      Top             =   5160
      Width           =   75
   End
End
Attribute VB_Name = "frmCLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ws_Prov_bloq  As String * 1
Dim tempo_ruc As String
Dim wCOM_NIVEL(6) As Integer
Dim NIVEL_MAX As Integer
Dim llave1 As String
Dim wGARANTES As Integer
Dim CU As Integer
Dim pasa As Integer
Dim UNICO As String
Dim CIA_REF As String * 2
Dim loc_key  As Integer
Dim LOC_TIPREG  As Integer
Dim loc_ultcod As Currency
Dim PSCLILOC_LLAVE As rdoQuery
Dim PSCLILOC_MAYOR As rdoQuery
Dim cliloc_llave As rdoResultset
Dim cliloc_mayor As rdoResultset
Dim PSPAR_CLI As rdoQuery
Dim par_llave_cli As rdoResultset
'agregado
'06/12/2001
Dim PS_dir As rdoQuery
Dim llave_dir As rdoResultset
'*****************************
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset

Dim LOC_CANCELA As Integer
Dim LOC_CTA_CLI As String * 12
Dim LOC_DES_CLI As String * 50
Dim LOC_NIVEL As Integer
Dim LOC_CTA_SUP As String
Dim LOC_FLAG_AFEC As String * 1
Dim LOC_ESTADO As String * 1
Dim LOC_TIPO_CTA As Integer
Dim LOC_SIGNO_D As Integer
Dim LOC_SIGNO_H As Integer
Dim LOC_ACT_PAS As Integer
Dim COD_ORIGINAL As Currency
Dim LOC_CTA_CLI2 As String * 12
Dim LOC_DES_CLI2 As String * 50
Dim LOC_NIVEL2 As Integer
Dim LOC_CTA_SUP2 As String
Dim LOC_FLAG_AFEC2 As String * 1
Dim LOC_ESTADO2 As String * 1
Dim LOC_TIPO_CTA2 As Integer
Dim LOC_SIGNO_D2 As Integer
Dim LOC_SIGNO_H2 As Integer
Dim LOC_ACT_PAS2 As Integer
Dim PSPLAC_LLAVE As rdoQuery
Dim cliplac_llave As rdoResultset

Dim SUTRA As rdoQuery
Dim sutra_llave As rdoResultset
Dim LOC_TIPOCLI As String * 1
Dim ntipocli As Integer
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
Public Sub LLENA_LISTAS(cont As ListBox, tip As Integer, wcodclie As Currency)
    PUB_TIPREG = tip
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        wcodclie = tab_mayor!TAB_NOMLARGO
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
        tab_mayor.MoveNext
    Loop
    
End Sub

Public Sub LLENA_GRUPOS(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    PUB_CODCIA = LK_CODCIA
    If tip = 66 Or tip = 144 Then
    PUB_CODCIA = "00"
    End If
    
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    If PUB_TIPREG = 333 Then lisdescto.Clear
    Do Until tab_mayor.EOF
         cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        If PUB_TIPREG = 333 Then lisdescto.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
    
End Sub
Public Sub LLENA_FILTROS(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    cont.AddItem "TODOS " + String(60, "  ") + "T"
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
    cont.ListIndex = 0
End Sub


Public Sub LLENA_BLOQ()
   ListBloqueos.Clear
   PUB_TIPREG = 7
   PUB_CODCIA = "00"
   SQ_OPER = 2
   LEER_TAB_LLAVE
   Do Until tab_mayor.EOF
        If cliloc_llave!CLI_TIPO_BLOQ1 = Left(tab_mayor!tab_nomcorto, 1) Then
            ListBloqueos.AddItem tab_mayor!tab_nomcorto
        End If
        If cliloc_llave!CLI_TIPO_BLOQ2 = Left(tab_mayor!tab_nomcorto, 1) Then
            ListBloqueos.AddItem tab_mayor!tab_nomcorto
        End If
        If cliloc_llave!CLI_TIPO_BLOQ3 = Left(tab_mayor!tab_nomcorto, 1) Then
            ListBloqueos.AddItem tab_mayor!tab_nomcorto
        End If
        If cliloc_llave!CLI_TIPO_BLOQ4 = Left(tab_mayor!tab_nomcorto, 1) Then
            ListBloqueos.AddItem tab_mayor!tab_nomcorto
        End If
        tab_mayor.MoveNext
   Loop
End Sub

Public Sub Mayuscula(Optional tecla)
'CONVIERTE TODA A MAYUSCULAS LETRAS
Dim car As String, Longt As Integer
car = Chr$(tecla)
car = UCase$(Chr$(tecla))
tecla = Asc(car)
If Not car < "a" Or car > "z" Then
  If tecla <> 209 Then
        tecla = 0
        Beep
  End If
End If
End Sub


Public Sub BLOQUEA_TEXT()
    t_ciarela.Enabled = False
    cboDiaVisita.Enabled = False
    txtnombre.Enabled = False
    txtesposo.Enabled = False
    Txtesposa.Enabled = False
    TxtEmpresa.Enabled = False
    txtdireccion.Enabled = False
    Txtnumdir.Enabled = False
    TxtZona.Enabled = False
    TxtSubZona.Enabled = False
    txtZonaNew.Enabled = False
    txtDirTrabajo.Enabled = False
    txtNumDirTrabajo.Enabled = False
    frmCLI.TxtZonaTrabajo.Enabled = False
    TxtSubZonaTrabajo.Enabled = False
    txtRUCesposo.Enabled = False
    txtRUCesposa.Enabled = False
    txtRUCempresa.Enabled = False
    frmCLI.txtpropiedad2.Enabled = False
    frmCLI.txtpropiedad1.Enabled = False
    frmCLI.txtregpublico1.Enabled = False
    frmCLI.txtregpublico2.Enabled = False
    frmCLI.txtautovaluo.Enabled = False
    frmCLI.txtauto1.Enabled = False
    frmCLI.txtauto2.Enabled = False
    frmCLI.txtprendas.Enabled = False
    frmCLI.txttelefono1.Enabled = False
    frmCLI.txttelefono2.Enabled = False
    frmCLI.otrocontrato.Enabled = False
    frmCLI.letraotorgado.Enabled = False
    frmCLI.cmbgrupo.Enabled = False
    frmCLI.txtsubgrupo.Enabled = False
    frmCLI.txtNucleo.Enabled = False
    frmCLI.tcuenta.Enabled = False
    frmCLI.txtciarela.Enabled = False
    frmCLI.OptNombre(0).Enabled = False
    frmCLI.OptNombre(1).Enabled = False
    frmCLI.OptNombre(2).Enabled = False
    frmCLI.txtestado.Enabled = False
    frmCLI.TxtLugarCasa.Enabled = False
    frmCLI.TxtLugarTrab.Enabled = False
    frmCLI.txtlimite.Enabled = False

    txtDTX.Enabled = False
    txtprog.Enabled = False
    cmdcontab.Enabled = False
    cmdcontab2.Enabled = False
    tcuenta2.Enabled = False
    tcuenta22.Enabled = False
    frmCLI.txtpordes.Enabled = False
    g_fechafac.Enabled = False
    g_diasfac.Enabled = False
    t_fechafac.Enabled = False
    t_diasfac.Enabled = False
    t_diascred.Enabled = False
    Cmbcate.Enabled = False
    lisdescto.Enabled = False
        
    cmbtipocli.Enabled = False
        
End Sub
Public Sub DESBLOQUEA_TEXT()
    t_ciarela.Enabled = True
    cboDiaVisita.Enabled = True
    txtesposo.Enabled = True
    Txtesposa.Enabled = True
    TxtEmpresa.Enabled = True
    txtdireccion.Enabled = True
    Txtnumdir.Enabled = True
    TxtZona.Enabled = True
    TxtSubZona.Enabled = True
    txtZonaNew.Enabled = True
    txtDirTrabajo.Enabled = True
    txtNumDirTrabajo.Enabled = True
    frmCLI.TxtZonaTrabajo.Enabled = True
    TxtSubZonaTrabajo.Enabled = True
    txtRUCesposo.Enabled = True
    txtRUCesposa.Enabled = True
    txtRUCempresa.Enabled = True
    frmCLI.txtpropiedad2.Enabled = True
    frmCLI.txtpropiedad1.Enabled = True
    frmCLI.txtregpublico1.Enabled = True
    frmCLI.txtregpublico2.Enabled = True
    frmCLI.txtautovaluo.Enabled = True
    frmCLI.txtauto1.Enabled = True
    frmCLI.txtauto2.Enabled = True
    frmCLI.txtprendas.Enabled = True
    frmCLI.txttelefono1.Enabled = True
    frmCLI.txttelefono2.Enabled = True
    frmCLI.otrocontrato.Enabled = True
    frmCLI.letraotorgado.Enabled = True
    frmCLI.cmbgrupo.Enabled = True
    frmCLI.txtsubgrupo.Enabled = True
    frmCLI.txtNucleo.Enabled = True
    frmCLI.OptNombre(0).Enabled = True
    frmCLI.OptNombre(1).Enabled = True
    frmCLI.OptNombre(2).Enabled = True
    frmCLI.tcuenta.Enabled = True
    frmCLI.txtciarela.Enabled = True
    frmCLI.txtestado.Enabled = True
    frmCLI.TxtLugarCasa.Enabled = True
    frmCLI.TxtLugarTrab.Enabled = True
    frmCLI.txtlimite.Enabled = True

    txtDTX.Enabled = True
    txtprog.Enabled = True
    cmdcontab.Enabled = True
    cmdcontab2.Enabled = True
    tcuenta2.Enabled = True
    tcuenta22.Enabled = True
    frmCLI.txtpordes.Enabled = True
    
    g_fechafac.Enabled = True
    g_diasfac.Enabled = True
    t_fechafac.Enabled = True
    t_diasfac.Enabled = True
    t_diascred.Enabled = True
    Cmbcate.Enabled = True
    lisdescto.Enabled = True
    
    cmbtipocli.Enabled = True
    
End Sub



Private Sub cboDiaVisita_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
Exit Sub
wpos = cboDiaVisita.ListIndex
PUB_TIPREG = 66
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
'DoEvents

LLENA_GRUPOS frmCLI.cboDiaVisita, 66

cboDiaVisita.SetFocus
SendKeys "%{up}"

End Sub

'*************************
'PROCEDIMIENTO PARA BUSCAR
'************************
Private Sub cboDireccion_Click()
Dim SQL As String
Dim PS_DIREC As rdoQuery
Dim llave_Direc As rdoResultset
SQL = "select * FROM DIRCLI where codcia=? and DIRCLI=? AND codcli=? and cp=?"
'sql = "SELECT * FROM DIRCLI WHERE CODCIA=? AND CODCLI=? AND CP=?"
 Set PS_DIREC = CN.CreateQuery("", SQL)
  PS_DIREC.rdoParameters(0) = " "
  PS_DIREC.rdoParameters(1) = 0
  PS_DIREC.rdoParameters(2) = 0
  PS_DIREC.rdoParameters(3) = " "
  Set llave_Direc = PS_DIREC.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  PS_DIREC(0) = LK_CODCIA
  PS_DIREC(1) = Val(frmCLI.cboDireccion.ItemData(cboDireccion.ListIndex))
  PS_DIREC(2) = Val(frmCLI.Txt_key)
  'OJO
  If Trim(Left$(CmbCGP.Text, 1)) = "C" Then
  PS_DIREC(3) = "C"
  Else
  PS_DIREC(3) = "P"
  End If
  llave_Direc.Requery
  If llave_Direc.EOF Then Exit Sub
  
  Do Until llave_Direc.EOF
   ASIGNA_INT TxtLugarTrab, Val(llave_Direc!CLI_LUGAR_TRAB)
   ASIGNA_INT TxtZonaTrabajo, Val(llave_Direc!cli_TRAB_ZONA)
   ASIGNA_INT cboProvincia, Val(llave_Direc!CLI_CASA_SUBZONA)
   ASIGNA_INT TxtSubZonaTrabajo, Val(llave_Direc!cli_TRAB_SUBZONA)
   txtNumDirTrabajo = llave_Direc!NUMERO
   txtDirTrabajo = llave_Direc!direc
   txtpropiedad2 = llave_Direc!ref
   llave_Direc.MoveNext
  Loop
End Sub

Private Sub cboProvincia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtpropiedad2.SetFocus
End If
End Sub

Private Sub cboProvincia_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos
If KeyCode <> 45 Then
  Exit Sub
End If

wpos = cboProvincia.ListIndex
PUB_TIPREG = Mid(cboProvincia.ToolTipText, 13, Len(cboProvincia.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA cboProvincia, 30
' LLENA_ZONA TxtSubZonaTrabajo, 35
cboProvincia.SetFocus
SendKeys "%{up}"


End Sub

Private Sub Cmbcate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 frmCLI.TxtLugarCasa.SetFocus
 SendKeys "%{up}"
End If

End Sub

Private Sub Cmbcate_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = Cmbcate.ListIndex
PUB_TIPREG = Mid(Cmbcate.ToolTipText, 13, Len(Cmbcate.ToolTipText))
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "SUB - GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_GRUPOS Cmbcate, 230
Cmbcate.SetFocus
SendKeys "%{up}"

End Sub

Private Sub CmbCGP_Click()
If llave1 <> "X" Then
  If ws_Prov_bloq = "A" Then
     If Left(CmbCGP.Text, 1) = "P" Then
       CmbCGP.ListIndex = 0
       MsgBox "No tiene Acceso a Proveedores...", 48, Pub_Titulo
       Exit Sub
     End If
  End If
  Txt_key.Enabled = False
  If Trim(txtnombre.Text) <> "" Then
    LIMPIA_CLI
  End If
  CmbCGP_KeyPress 13
End If
End Sub

Private Sub CmbCGP_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.Text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub CmbCGP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmbCGP.Text = "" Then
       CmbCGP.SetFocus
       SendKeys "%{UP}"
       Exit Sub
    End If
'    ALLVISIBLE
    If Left(CmbCGP.Text, 1) = "P" Then
      'frmCLI.SSTab1.TabCaption(0) = "&Datos Proveedor - Principales"
      'frmCLI.SSTab1.TabCaption(1) = "&Datos Proveedor - Opcionales"
       Frame2.Caption = "PROVEEDORES"
       Label3(0).Caption = "Rel./Vendedor"
       LOC_TIPREG = 310 ' PROVEEDORES
       Screen.MousePointer = 11
       ETIQUETA_CLI
       Screen.MousePointer = 0
       lbllimite.Visible = False
       txtlimite.Visible = False
       lcuenta.Caption = "Cta. Pasivo:"
       txtauto1.Locked = False
       txtauto2.Locked = False
        SUTRA.rdoParameters(0) = 2748
        sutra_llave.Requery
        condi.Clear
        condi.AddItem "Opcional" & String(60, " ") & "-1"
        Do Until sutra_llave.EOF
        condi.AddItem sutra_llave!sut_descripcion & String(70, " ") & Str(sutra_llave!sut_secuencia)
        sutra_llave.MoveNext
        Loop
    Else
      Label3(0).Caption = "Cuenta2"
      SUTRA.rdoParameters(0) = 2401
      sutra_llave.Requery
      condi.Clear
      condi.AddItem "Opcional" & String(60, " ") & "-1"
      Do Until sutra_llave.EOF
        condi.AddItem sutra_llave!sut_descripcion & String(70, " ") & Str(sutra_llave!sut_secuencia)
        sutra_llave.MoveNext
      Loop
      ' txtauto1.Locked = True
      txtauto2.Locked = True
      lcuenta.Caption = "Cta. Activo:"
      'frmCLI.SSTab1.TabCaption(0) = "&Datos Clientes - Principales"
      'frmCLI.SSTab1.TabCaption(1) = "&Datos Clientes - Opcionales"
      Frame2.Caption = "CLIENTES"
      LOC_TIPREG = 300 ' CLIENTES
      Screen.MousePointer = 11
      ETIQUETA_CLI
      lbllimite.Visible = True
      txtlimite.Visible = True
      Screen.MousePointer = 0
    End If
      If Left(CmbCGP.Text, 1) = "C" Then
         LLENA_GRUPOS frmCLI.cmbgrupo, 222
      Else
         LLENA_GRUPOS frmCLI.cmbgrupo, 223
      End If
      If Left(CmbCGP.Text, 1) = "C" Then
          LLENA_GRUPOS txtsubgrupo, 333
      Else
          LLENA_GRUPOS txtsubgrupo, 334
      End If

    frmCLI.Txt_key.Locked = False
    frmCLI.Txt_key.Enabled = True
    frmCLI.Txt_key.SetFocus
    
End If
End Sub

Private Sub cmbgrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtsubgrupo.SetFocus
  SendKeys "%{up}"
End If
End Sub

Private Sub cmbgrupo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = cmbgrupo.ListIndex
PUB_TIPREG = Mid(cmbgrupo.ToolTipText, 13, Len(cmbgrupo.ToolTipText))
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
'DoEvents
If Left(CmbCGP.Text, 1) = "C" Then
  LLENA_GRUPOS frmCLI.cmbgrupo, 222
Else
  LLENA_GRUPOS frmCLI.cmbgrupo, 223
End If
cmbgrupo.SetFocus
SendKeys "%{up}"


End Sub

Private Sub cmbtipocli_Click()
If LOC_TIPOCLI = "A" Then
    If Val(Right(Trim(cmbtipocli.Text), 4)) <> ntipocli Then
        If LK_CODUSU <> "ADMIN" Then
            MsgBox "No Procede , Solo el Administrador", 48, Pub_Titulo
            ASIGNA_INT cmbtipocli, ntipocli
            cmbtipocli.SetFocus
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub cmbtipocli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 frmCLI.TxtLugarCasa.SetFocus
 SendKeys "%{up}"
End If
End Sub

Private Sub cmbtipocli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = cmbgrupo.ListIndex
PUB_TIPREG = Mid(cmbtipocli.ToolTipText, 13, Len(cmbtipocli.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
'DoEvents
If Left(CmbCGP.Text, 1) = "C" Then
  LLENA_GRUPOS frmCLI.cmbtipocli, 144
Else
  LLENA_GRUPOS frmCLI.cmbtipocli, 144
End If
cmbtipocli.SetFocus
SendKeys "%{up}"
End Sub

Private Sub cmdAgregar_Click()
'On Error GoTo ESCAPA
If Trim(CmbCGP.Text) = "" Then
   MENSAJE_CLI "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(cmdAgregar.Caption, 2) = "&A" And cmdAgregar.Enabled = True Then
    LOC_TIPOCLI = ""
    cmdAgregar.Caption = "&Grabar"
    Frame2.Enabled = True
    SSTab1.Enabled = True
    cmdcancelar.Enabled = True
    CmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    'AGREGADO
    'If Trim(Left$(CmbCGP, 1)) = "C" Then
    cmdcancel.Enabled = True
    cmdDelete.Enabled = True
    cmdDireccion.Enabled = True
    cboDireccion.Enabled = True
    cboProvincia.Enabled = True
    
    DESBLOQUEA_TEXT
    If LK_EMP <> "PAR" Then
     Txt_key.Locked = True
    End If
    LIMPIA_CLI
    If Left(CmbCGP.Text, 1) = "C" Then
        If cboDiaVisita.ListCount > 0 Then cboDiaVisita.ListIndex = 0
        frmCLI.OptNombre(0).Value = True
        frmCLI.Txt_key = GENERA_CODI
    ElseIf Left(CmbCGP.Text, 1) = "P" Then
        frmCLI.OptNombre(0).Value = True
        frmCLI.Txt_key = GENERA_PRO
    End If
    frmCLI.txtesposo.SetFocus
    Txt_key.ToolTipText = ""
    CmbCGP.Enabled = False
    If frmCLI.cmbgrupo.ListCount <> 0 Then frmCLI.cmbgrupo.ListIndex = 0
    frmCLI.txtestado.ListIndex = 0
    frmCLI.SSTab1.tab = 0
    frmCLI.t_fechafac.Text = LK_FECHA_DIA
    pasa = 1
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
    TxtZona.ListIndex = 0
    TxtSubZona.ListIndex = 0
    txtZonaNew.ListIndex = 0
    TxtLugarCasa.ListIndex = 0
    txtauto1.Text = "07"
    'If LK_EMP <> "HER" Then
     TxtLugarTrab.ListIndex = 0
     TxtZonaTrabajo.ListIndex = 0
     TxtSubZonaTrabajo.ListIndex = 0
    'End If
    If condi.ListCount > 2 Then condi.ListIndex = 2
    condi.Enabled = False
    lisdescto.Enabled = False
    If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
       lisdescto.Enabled = True
    End If
    ASIGNA_INT cmbtipocli, 7
    cmbtipocli.Enabled = False
    If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Or LK_CODUSU = "SUPERVISOR" Then
        cmbtipocli.Enabled = True
    End If
    'AGREGAMOS EN BLANCO
Else
 If Trim(frmCLI.txtesposo.Text) = "" Then
   MsgBox "Ingrese Nombre...", 48, Pub_Titulo
   frmCLI.txtesposo.SetFocus
  Exit Sub
 End If
  If Left(CmbCGP.Text, 1) = "C" Then
      If pasa = 1 Then
     '    If EXISTE_CLI("C", Left(frmCLI.txtesposo.Text, 15), Trim(Txt_key.Text)) Then
      '      MENSAJE_CLI " Existen algunos clientes con estos NOMBRES .."
      '      frmCLI.ListExiste.SetFocus
      '      Exit Sub
      '   End If
      End If
      pasa = 0
      If par_llave!PAR_CONTABILIDAD = "" Then
        GoTo PASACONTAB
      End If
      If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(CmbCGP.Text, 1) = "C" Then
            GoTo PASACONTAB
      ElseIf Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(CmbCGP.Text, 1) = "P" Then
            GoTo PASACONTAB
      End If
      If Trim(LOC_CTA_CLI) = "" And Trim(LOC_DES_CLI) = "" Then
            If Trim(tcuenta.Text) = "" Then
            pub_mensaje = "Desea Relacionarlo a Contabilidad ?"
            Pub_Respuesta = MsgBox(pub_mensaje, vbYesNoCancel + vbQuestion + vbDefaultButton3, Pub_Titulo)
            If Pub_Respuesta = vbYes Then
                cmdcontab_Click
                If LOC_CANCELA = 1 Then
                    MsgBox "Intente Nuevamente colocar Cta. Contable.", 48, Pub_Titulo
                    Exit Sub
                 Else
                    LOC_CANCELA = 0
                End If
            ElseIf Pub_Respuesta = vbCancel Then
               Exit Sub
            End If
         End If
      End If
PASACONTAB:
     If Not CONSIS_CLI Then
          Exit Sub
     End If
     If LK_EMP = "PAR" And COD_ORIGINAL <> Val(Txt_key.Text) Then
      SQ_OPER = 1
      pu_codclie = Val(Txt_key.Text)
      pu_cp = "C"
      pu_codcia = LK_CODCIA
      LEER_CLILOC_LLAVE
      If Not cliloc_llave.EOF Then
         MsgBox "Cliente Existe en Compañia ..", 48, Pub_Titulo
         Azul Txt_key, Txt_key
         Exit Sub
      End If
     End If
     On Error GoTo VERLO_GRABAR
     pub_cadena = "SELECT * FROM CONTROLL"
     CN.Execute "Begin Transaction", rdExecDirect
     Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
     frmCLI.Txt_key = GENERA_CODI
     If wGARANTES = 1 Then
           GRABAR_CLI "G"
     ElseIf wGARANTES = 2 Then
           GRABAR_CLI "H"
     Else
           GRABAR_CLI "C"
     End If
     CN.Execute "Commit Transaction", rdExecDirect
     con_llave.Close
     On Error GoTo 0
     MENSAJE_CLI "Registro,   AGREGADO ... "
     wGARANTES = 0
  ElseIf Left(CmbCGP.Text, 1) = "P" Then

      If pasa = 1 Then
         If EXISTE_CLI("P", Left(frmCLI.txtesposo.Text, 15), Trim(Txt_key.Text)) Then
            MENSAJE_CLI " Existen algunos Proveedor con estos NOMBRES .."
            frmCLI.ListExiste.SetFocus
            Exit Sub
         End If
      End If
       pasa = 0
       If Not CONSIS_CLI Then
          Exit Sub
       End If
       'SQ_OPER = 1
       'pu_codclie = Val(txt_key.text)
       'pu_cp = "P"
       'pu_codcia = LK_CODCIA
       'LEER_CLILOC_LLAVE
       'If Not cliloc_llave.EOF Then
       '   MsgBox "Proveedor Existe en Compañia ..", 48, Pub_Titulo
       '   Exit Sub
       'End If
       If Trim(LOC_CTA_CLI) = "" And Trim(LOC_DES_CLI) = "" And LK_CODCIA = "03" Then
         If Trim(tcuenta.Text) = "" Then
            pub_mensaje = "Antes de Grabar.  Desea Relacionarlo a Contabilidad ?"
            Pub_Respuesta = MsgBox(pub_mensaje, vbYesNoCancel + vbQuestion + vbDefaultButton3, Pub_Titulo)
            If Pub_Respuesta = vbYes Then
              ''   cmdcontab_Click
                If LOC_CANCELA = 1 Then
                    MsgBox "Intente Nuevamente colocar Cta. Contable.", 48, Pub_Titulo
                    Exit Sub
                 Else
                    LOC_CANCELA = 0
                End If
             ElseIf Pub_Respuesta = vbCancel Then
               Exit Sub
            End If
         End If
       End If
       On Error GoTo VERLO_GRABAR
       CN.Execute "Begin Transaction", rdExecDirect
       pub_cadena = "SELECT * FROM CONTROLL"
       Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
       frmCLI.Txt_key = GENERA_PRO
       GRABAR_CLI "P"
       con_llave.Close
       CN.Execute "Commit Transaction", rdExecDirect
       On Error GoTo 0
       MENSAJE_CLI "Proveedor , AGREGADO... "
    End If
    cmdAgregar.Caption = "&Agregar"
    cmdEliminar.Enabled = True
    CmdModificar.Enabled = True

    BLOQUEA_TEXT
    Txt_key.Locked = False
    CmbCGP.Enabled = True
    Screen.MousePointer = 0
    frmCLI.SSTab1.tab = 0
    Txt_key.ToolTipText = ""
    LIMPIA_CLI
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
End If
Exit Sub
    
ESCAPA:
   If Err.Number = 40002 Then
      Screen.MousePointer = 0
      MsgBox "El Codigo generado ya existe " & Chr(13) & "Se procede a generar el siguiente codigo y a continuación " & Chr(13) & "Intente Grabar Nuevamente...", 48, Pub_Titulo
      frmCLI.Txt_key = GENERA_CODI
      Resume Next
      Exit Sub
   Else
      Screen.MousePointer = 0
      MsgBox Err.Number & "  " & Err.Description & "   ... "
      cmdcancelar_Click
'       Resume Next
   End If
   Exit Sub
VERLO_GRABAR:
'    If con_llave Is Nothing Then
     con_llave.Close
     MsgBox Err.Description
     CN.Execute "Rollback Transaction", rdExecDirect
'    End If
    cmdcancelar_Click
fin:
End Sub

Private Sub cmdagregar_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.Text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   If frmCLI.Txt_key.Visible Then
      frmCLI.Txt_key.SetFocus
   End If
End If

End Sub


Private Sub cmdcancel_Click()
Dim SQL As String
Dim dir As String
Dim strRef As String
Dim strDir As String
On Error GoTo ErrorHandle
If cmdcancel.Caption = "Cancelar" Then
  cmdcancel.Caption = "Modificar"
  cmdDireccion.Caption = "Agregar"
  cmdDelete.Enabled = True
  cboDireccion.ListIndex = 0
  Exit Sub
ElseIf cmdcancel.Caption = "Grabar" Then
 cmdcancel.Caption = "Modificar"
 cmdDireccion.Caption = "Agregar"
 cmdDelete.Enabled = True
 strDir = Trim(txtDirTrabajo)
 strRef = Trim(txtpropiedad2)
 dir = Trim(Left$(TxtLugarTrab, 10))
    dir = dir + " " + Trim(Left$(strDir, 30))
    dir = dir + " " + Trim(Left$(txtNumDirTrabajo, 30))
    dir = dir + " Zn. " + Trim(Left$(TxtSubZonaTrabajo.Text, 30))
    dir = dir + ", Dt. " + Trim(Left$(TxtZonaTrabajo, 30))
    dir = dir + ", Pr. " + Trim(Left$(cboProvincia, 30))
 SQL = "UPDATE DIRCLI " & _
 "SET DIREC='" & Trim(txtDirTrabajo) & "',ref= '" & Trim(txtpropiedad2) & "',cli_lugar_trab='" & _
 Val(Right(frmCLI.TxtLugarTrab, 4)) & "',cli_trab_zona='" & _
 Val(Right(frmCLI.TxtZonaTrabajo.Text, 4)) & "',cli_casa_subzona='" & _
 Val(Right(frmCLI.cboProvincia, 4)) & "',cli_trab_subzona='" & _
 Val(Right(frmCLI.TxtSubZonaTrabajo, 4)) & "',Numero='" & _
 Val(txtNumDirTrabajo) & "',dircomp='" & dir & "' " & _
 "WHERE CODCIA='" & LK_CODCIA & "' AND CODCLI='" & Val(frmCLI.Txt_key) & "' AND DIRCLI='" & Val(cboDireccion.ItemData(cboDireccion.ListIndex)) & "'"
 CN.Execute SQL
 LLENA_DIRECCIONES
 cboDireccion.ListIndex = 0
 
ElseIf cmdcancel.Caption = "Modificar" Then
 cmdDelete.Enabled = False
 cmdDireccion.Caption = "Cancelar"
 cmdcancel.Caption = "Grabar"
End If
Exit Sub
ErrorHandle:
 Select Case Err.Number
  Case Is = 381
   MsgBox "Posiblemente no ha seleccionado ó no existe ninguna direción,", vbInformation, "Direcciones"
  Exit Sub
 End Select

End Sub

Private Sub cmdcancelar_Click()
'agregado
Frame2.Enabled = True
SSTab1.Enabled = True
cmdDelete.Enabled = False
cmdcancel.Enabled = False
cmdDireccion.Enabled = False
cboProvincia.Enabled = False
cboDireccion.Enabled = False
LOC_TIPOCLI = ""
If Txt_key.Visible = False Then
  Exit Sub
End If
If Left(cmdAgregar.Caption, 2) = "&A" And Left(CmdModificar.Caption, 2) = "&M" Then
    LIMPIA_CLI
    cmdcancelar.Enabled = True
    Txt_key.Locked = False
    MENSAJE_CLI "Proceso Cancelado... !!!    "
    Txt_key.Enabled = True
    Txt_key.SetFocus
    frmCLI.SSTab1.tab = 0
    Screen.MousePointer = 0
    pasa = 0
    cmdcontab.Enabled = False
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
    Exit Sub
End If
     Screen.MousePointer = 11
     If Left(CmdModificar.Caption, 2) = "&G" Then
        CmdModificar.Caption = "&Modificar"
        If Left(CmbCGP.Text, 1) = "C" Then
           LLENA_CLI 1, "C"
        Else
           LLENA_CLI 1, "P"
        End If
        Txt_key.Locked = True
     Else
        GoSub ELI_TABLAS
        cmdAgregar.Caption = "&Agregar"
        cmdcontab.Enabled = False
        LIMPIA_CLI
        Txt_key.Locked = False
        Txt_key.SetFocus
     End If
     cmdAgregar.Enabled = True
     cmdEliminar.Enabled = True
     CmdModificar.Enabled = True

     Txt_key.ToolTipText = ""
     wGARANTES = 0
     BLOQUEA_TEXT
     MENSAJE_CLI "Proceso Cancelado... !!!    "
     CmbCGP.Enabled = True
     frmCLI.SSTab1.tab = 0
     Screen.MousePointer = 0
     LOC_CTA_CLI = ""
     LOC_CTA_CLI2 = ""
     pasa = 0


Exit Sub
ELI_TABLAS:
If LK_FLAG_GRIFO <> "A" Then Return
pu_codclie = Val(Txt_key.Text)
If pu_codclie = 0 Then Return
PSPLAC_LLAVE(0) = LK_CODCIA
PSPLAC_LLAVE(1) = 2101
PSPLAC_LLAVE(2) = pu_codclie
cliplac_llave.Requery
Do Until cliplac_llave.EOF
  cliplac_llave.Delete
  cliplac_llave.MoveNext
Loop

PSPLAC_LLAVE(0) = LK_CODCIA
PSPLAC_LLAVE(1) = 2301
PSPLAC_LLAVE(2) = pu_codclie
cliplac_llave.Requery
Do Until cliplac_llave.EOF
  cliplac_llave.Delete
  cliplac_llave.MoveNext
Loop


Return
End Sub

Private Sub cmdCancelar_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.Text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub


Private Sub cmdCerrar_Click()
Dim iFormCount As Integer
Dim WCODI As String
cmdcancelar_Click
frmCLI.Hide
If LK_EMP = "3AA" Then
 If Forms.count - 1 > 0 Then
  For iFormCount = Forms.count - 1 To 1 Step -1
     If iFormCount <> 1 Then
        If UCase(Forms(iFormCount).Name) = "FORMGEN" And Left(CmbCGP.Text, 1) = "C" Then
              If FORMGEN.i_codcli.Visible Then
                FORMGEN.i_codcli.SetFocus
              End If
        End If
     End If
  Next iFormCount
 End If
End If
If LK_FLAG_GRIFO = "A" Then
 'If Forms.count - 1 > 0 Then
 ' For iFormCount = Forms.count - 1 To 1 Step -1
 '    If iFormCount <> 1 Then
 '       If UCase(Forms(iFormCount).Name) = "FORM_GRIFO" And Left(CmbCGP.Text, 1) = "C" Then
 '             If FORM_GRIFO.i_codcli.Visible Then
 '               FORM_GRIFO.i_codcli.SetFocus
 '             End If
 '       End If
 '    End If
 ' Next iFormCount
 'End If
End If


End Sub

Private Sub cmdCerrar_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.Text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdCerrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmCLI.Txt_key.SetFocus
End If
End Sub

Private Sub cmdconfirma_Click()
  If Op(0).Value And Left(frmCLI.CmbCGP, 1) = "C" Then
     frmCLI.Txt_key.Text = ListExiste.TextMatrix(ListExiste.Row, 1)
     pasa = 1
     frmCLI.F14.Visible = False
     cmdAgregar_Click
     Exit Sub
  End If
  If Op(0).Value And Left(frmCLI.CmbCGP, 1) = "P" Then
    frmCLI.txtnombre.Text = ListExiste.TextMatrix(ListExiste.Row, 2)
    frmCLI.Txt_key.Text = ListExiste.TextMatrix(ListExiste.Row, 1)
     pasa = 1
     frmCLI.F14.Visible = False
     If Left(cmdAgregar.Caption, 2) = "&G" And cmdAgregar.Enabled = True Then cmdAgregar_Click
     If Left(CmdModificar.Caption, 2) = "&G" And CmdModificar.Enabled = True Then CmdModificar_Click
     Exit Sub
  End If
  If Op(1).Value Then
     pasa = 0
     frmCLI.F14.Visible = False
     If Left(cmdAgregar.Caption, 2) = "&G" And cmdAgregar.Enabled = True Then cmdAgregar_Click
     If Left(CmdModificar.Caption, 2) = "&G" And CmdModificar.Enabled = True Then CmdModificar_Click
     Exit Sub
  End If
  MsgBox "Seleccione una de las dos Opciones ..", 48, Pub_Titulo
End Sub

Private Sub cmdcontab_Click()
If par_llave!PAR_CONTABILIDAD <> "A" Then
  Exit Sub
End If
Exit Sub
If Left(CmbCGP.Text, 1) = "C" Then
  If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(cmdcontab.Caption, 2) = "&Q" Then
      tcuenta.Text = ""
      cmdcontab.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
ElseIf Left(CmbCGP.Text, 1) = "P" Then
  If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(cmdcontab.Caption, 2) = "&Q" Then
      tcuenta.Text = ""
      cmdcontab.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
End If
If Left(cmdcontab.Caption, 2) = "&Q" Then
    pub_mensaje = "Confirmar la eliminación de la Cuenta : " & tcuenta.Text & " , Continuar ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
    SQ_OPER = 1
    PUB_CUENTA = Trim(tcuenta.Text)
    LEER_COM_LLAVE
    If com_llave.EOF Then
      tcuenta.Text = ""
    Else
      com_llave.Delete
      tcuenta.Text = ""
      CmdModificar_Click
    End If
    cmdcontab.Caption = "Relacionar a Con&tabilidad"
    Exit Sub
End If
LOC_CANCELA = 0
If txtesposo.Text = "" Then
 MsgBox "Ingrese Descripción del cliente..", 48, Pub_Titulo
 Azul txtesposo, txtesposo
 Exit Sub
End If
If Left(CmbCGP.Text, 1) = "C" Then
    LK_TABLA = "CLIENTE"
    archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND ( COM_CUENTA >= '" & 12 & "' AND COM_CUENTA < '" & 13 & "' OR COM_CUENTA >= '" & 16 & "' AND COM_CUENTA < '" & 17 & "' OR COM_CUENTA >= '" & 14 & "' AND COM_CUENTA < '" & 15 & "' ) ORDER BY COM_CUENTA"
Else
   LK_TABLA = "PROVEEDOR"
   archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND (COM_CUENTA >= '" & 42 & "' AND COM_CUENTA < '" & 43 & "' OR COM_CUENTA >= '" & 46 & "' AND COM_CUENTA < '" & 47 & "' )  ORDER BY COM_CUENTA"
End If
LOC_CTA_CLI = ""
LOC_DES_CLI = ""
pb.Visible = True
DoEvents
Load frmBuscacta
frmBuscacta.lbltabla.Caption = LK_TABLA
pb.Visible = False
frmBuscacta.Show 1
LOC_CTA_CLI = Trim(frmBuscacta.tcuenta)
LOC_DES_CLI = Trim(frmBuscacta.tnombre.Text)
LOC_NIVEL = Val(frmBuscacta.txtdatos(0).Text)
LOC_CTA_SUP = Trim(frmBuscacta.txtdatos(1).Text)
LOC_FLAG_AFEC = Trim(frmBuscacta.txtdatos(2).Text)
LOC_ESTADO = Trim(frmBuscacta.txtdatos(3).Text)
LOC_TIPO_CTA = Val(frmBuscacta.txtdatos(4).Text)
LOC_SIGNO_D = Val(frmBuscacta.txtdatos(5).Text)
LOC_SIGNO_H = Val(frmBuscacta.txtdatos(6).Text)
LOC_ACT_PAS = Val(frmBuscacta.txtdatos(7).Text)
tcuenta = Trim(LOC_CTA_CLI)
If Trim(LOC_DES_CLI) = "" And Trim(LOC_DES_CLI) = "" Then
 LOC_CANCELA = 1
Else
 LOC_CANCELA = 0
End If
Unload frmBuscacta
If LOC_CANCELA = 1 Then
 Exit Sub
End If
If Left(CmbCGP.Text, 1) = "C" Then
   If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" Then
     Exit Sub
   End If
ElseIf Left(CmbCGP.Text, 1) = "P" Then
   If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" Then
     Exit Sub
   End If
End If

If Left(CmdModificar.Caption, 2) = "&G" Then
   CmdModificar_Click
End If
End Sub

Private Sub cmdcontab2_Click()
If par_llave!PAR_CONTABILIDAD <> "A" Then
  Exit Sub
End If
If Left(CmbCGP.Text, 1) = "P" Then
  If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(cmdcontab2.Caption, 2) = "&Q" Then
      tcuenta2.Text = ""
      cmdcontab2.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
End If
If Left(CmbCGP.Text, 1) = "C" Then
  If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(cmdcontab2.Caption, 2) = "&Q" Then
      tcuenta2.Text = ""
      cmdcontab2.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
End If

If Left(cmdcontab2.Caption, 2) = "&Q" Then
    pub_mensaje = "Confirmar la eliminación de la Cuenta : " & tcuenta2.Text & " , Continuar ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
     tcuenta2.Text = ""
     CmdModificar_Click
     cmdcontab2.Caption = "Relacionar a Con&tabilidad"
    Exit Sub
End If
LOC_CANCELA = 0
If txtesposo.Text = "" Then
 MsgBox "Ingrese Descripción del cliente..", 48, Pub_Titulo
 Azul txtesposo, txtesposo
 Exit Sub
End If
If Left(CmbCGP.Text, 1) = "C" Then
   LK_TABLA = "CLIENTES2"
   archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND (COM_CUENTA >= '" & 70 & "' AND COM_CUENTA < '" & 71 & "' OR COM_CUENTA >= '" & 75 & "' AND COM_CUENTA < '" & 78 & "') ORDER BY COM_CUENTA"
End If
If Left(CmbCGP.Text, 1) = "P" Then
   LK_TABLA = "PROVEEDOR2"
   archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND (COM_CUENTA >= '" & 60 & "' AND COM_CUENTA < '" & 61 & "' OR COM_CUENTA >= '" & 62 & "' AND COM_CUENTA < '" & 68 & "' OR COM_CUENTA >= '" & 33 & "' AND COM_CUENTA < '" & 39 & "' ) ORDER BY COM_CUENTA"
End If
LOC_CTA_CLI2 = ""
LOC_DES_CLI2 = ""
PB2.Visible = True
DoEvents
Load frmBuscacta
frmBuscacta.lbltabla.Caption = LK_TABLA
pb.Visible = False
frmBuscacta.Show 1
LOC_CTA_CLI2 = Trim(frmBuscacta.tcuenta)
LOC_DES_CLI2 = Trim(frmBuscacta.tnombre.Text)
LOC_NIVEL2 = Val(frmBuscacta.txtdatos(0).Text)
LOC_CTA_SUP2 = Trim(frmBuscacta.txtdatos(1).Text)
LOC_FLAG_AFEC2 = Trim(frmBuscacta.txtdatos(2).Text)
LOC_ESTADO2 = Trim(frmBuscacta.txtdatos(3).Text)
LOC_TIPO_CTA2 = Val(frmBuscacta.txtdatos(4).Text)
LOC_SIGNO_D2 = Val(frmBuscacta.txtdatos(5).Text)
LOC_SIGNO_H2 = Val(frmBuscacta.txtdatos(6).Text)
LOC_ACT_PAS2 = Val(frmBuscacta.txtdatos(7).Text)
tcuenta2 = Trim(LOC_CTA_CLI2)
If Trim(LOC_DES_CLI2) = "" And Trim(LOC_DES_CLI2) = "" Then
 LOC_CANCELA = 1
Else
 LOC_CANCELA = 0
End If
Unload frmBuscacta
If LOC_CANCELA = 1 Then
 Exit Sub
End If
If Left(CmbCGP.Text, 1) = "P" Then
   If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" Then
     Exit Sub
   End If
End If

If Left(CmdModificar.Caption, 2) = "&G" Then
   CmdModificar_Click
End If

End Sub

Private Sub cmdDelete_Click()
  Dim SQL As String
  With cboDireccion
  On Error GoTo ErrorDelete
  SQL = "DELETE FROM DIRCLI WHERE CODCIA='" & LK_CODCIA & "' " & _
        "AND DIRCLI='" & Val(.ItemData(.ListIndex)) & "' AND " & _
        "CODCLI='" & Val(frmCLI.Txt_key) & "' AND CP= '" & Left(CmbCGP.Text, 1) & "'"
  End With
  
  If MsgBox("Esta seguro de Eliminar esta dirección", vbYesNo, "Eliminar Dirección") = vbYes Then
  
  CN.Execute SQL
  LLENA_DIRECCIONES
  cboDireccion.SetFocus
  cboDireccion.ListIndex = 0
  Else
  cboDireccion.ListIndex = 0
  cboDireccion.SetFocus
  End If
Exit Sub
ErrorDelete:
 Select Case Err.Number
  Case Is = 381
  MsgBox "Posiblemente no ha seleccionado ó no existe ninguna direción,", vbInformation, "Direcciones"
  Exit Sub
 End Select
End Sub

Private Sub cmddescto_Click()
pu_codclie = Val(Txt_key.Text)
If pu_codclie = 0 Then Exit Sub
'PUB_TIPREG = 2301
'PUB_CODCIA = LK_CODCIA
'Load FrmDatplac
'FrmDatplac.Caption = "Tabla de Descutntos " & PUB_TIPREG
'FrmDatplac.Show 1
'DoEvents
'LLENA_DESCTO

End Sub

Private Sub cmdDireccion_Click()
'If Not frmCLI.txt_key = "" And Left$(cmdModificar.Caption, 2) = "&G" Then
'PUB_TIPREG = -5
'PUB_CODCIA = LK_CODCIA
'Load FrmDatArti
'FrmDatArti.Caption = "DIRECCIONES - CLIENTES"
'FrmDatArti.Show 1
'Else
' MsgBox "Primero seleccione un cliente", vbInformation, "Mensageria"
'End If
Dim strDir As String
Dim strRef As String
Dim SQL As String
Dim dir As String
On Error GoTo ErrorHandle

'llave_rep01.Requery
If cmdDireccion.Caption = "Grabar" Then
    strDir = Trim(txtDirTrabajo)
    strRef = Trim(txtpropiedad2)
    dir = Trim(Left$(TxtLugarTrab, 10))
    dir = dir + " " + Trim(Left$(strDir, 30))
    dir = dir + " " + Trim(Left$(txtNumDirTrabajo, 30))
    dir = dir + " Zn. " + Trim(Left$(TxtSubZonaTrabajo.Text, 30))
    dir = dir + ", Dt. " + Trim(Left$(TxtZonaTrabajo, 30))
    dir = dir + ", Pr. " + Trim(Left$(cboProvincia, 30))
    If strDir = "" Then
     MsgBox "Dato ingresado no valido, Intentelo nuevamente", vbInformation, "Dirección"
     Exit Sub
    End If
     SQL = "insert into dircli " & _
     "(codcia,codcli,cp,direc,ref,CLI_LUGAR_TRAB, " & _
     "CLI_TRAB_ZONA,CLI_CASA_SUBZONA,CLI_TRAB_SUBZONA,NUMERO,DIRCOMP) " & _
     "values('" & LK_CODCIA & "','" & Val(frmCLI.Txt_key) & "', '" & Trim(Left$(CmbCGP.Text, 1)) & "','" & strDir & "','" & strRef & "','" & _
     Val(Right(frmCLI.TxtLugarTrab, 4)) & "','" & _
     Val(Right(frmCLI.TxtZonaTrabajo.Text, 4)) & "','" & _
     Val(Right(frmCLI.cboProvincia, 4)) & "','" & _
     Val(Right(frmCLI.TxtSubZonaTrabajo, 4)) & "','" & _
     Val(txtNumDirTrabajo) & "','" & dir & "')"
     CN.Execute SQL
     LLENA_DIRECCIONES
     cboDireccion.SetFocus
     cmdDireccion.Caption = "Agregar"
     cmdcancel.Caption = "Modificar"
    cmdDelete.Enabled = True
 ElseIf cmdDireccion.Caption = "Agregar" Then
  cmdDelete.Enabled = False
  cmdDireccion.Caption = "Grabar"
  cmdcancel.Caption = "Cancelar"
  TxtLugarTrab.ListIndex = -1
  TxtZonaTrabajo.ListIndex = -1
  TxtSubZonaTrabajo.ListIndex = -1
  cboProvincia.ListIndex = -1
  txtNumDirTrabajo = ""
  txtDirTrabajo = ""
  txtpropiedad2 = ""
  TxtLugarTrab.SetFocus
 ElseIf cmdDireccion.Caption = "Cancelar" Then
 cmdDireccion.Caption = "Agregar"
 cmdDelete.Enabled = True
 cmdcancel.Caption = "Modificar"
 cboDireccion.ListIndex = 0
 End If
 'llave_rep01.Requery
ErrorHandle:
Exit Sub
End Sub

Private Sub cmdEliminar_Click()
Dim wcias As String
On Error GoTo sale
If Len(Txt_key) = 0 Or Len(txtnombre) = 0 Then
   MENSAJE_CLI "NO a seleccionado NADA ... !"
   Exit Sub
End If
  Dim PS_REP01 As rdoQuery
  Dim llave_rep01 As rdoResultset
  Screen.MousePointer = 11
  LblMensaje.Visible = True
  LblMensaje.Caption = "Verificando Data.  un Momento..."
  pub_cadena = "SELECT FAR_CODCIA FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODCLIE = ? "
  Set PS_REP01 = CN.CreateQuery("", pub_cadena)
  PS_REP01.rdoParameters(0) = " "
  PS_REP01.rdoParameters(1) = 0
  PS_REP01.MaxRows = 1
  Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = cliloc_llave!cli_codclie
  llave_rep01.Requery
  If Not llave_rep01.EOF Then
     LblMensaje.Visible = False
     Screen.MousePointer = 0
     MsgBox "NO se Puede Eliminar ...  CLIENTE  TIENE H I S T O R I A.. (en Movimientos)", 48, Pub_Titulo
     Exit Sub
  End If
  pub_cadena = "SELECT CAR_CODCLIE FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_CODCLIE = ? "
  Set PS_REP01 = CN.CreateQuery("", pub_cadena)
  PS_REP01.rdoParameters(0) = " "
  PS_REP01.rdoParameters(1) = 0
  PS_REP01.MaxRows = 1
  Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = cliloc_llave!cli_codclie
  llave_rep01.Requery
  If Not llave_rep01.EOF Then
     LblMensaje.Visible = False
     Screen.MousePointer = 0
     MsgBox "NO se Puede Eliminar ...  CLIENTE  TIENE H I S T O R I A.. (en saldos Iniciales)", 48, Pub_Titulo
     Exit Sub
  End If
  
  Screen.MousePointer = 0
  LblMensaje.Caption = ""
  If Trim(Nulo_Valors(GEN!gen_cli_cias)) <> "" Then
    wcias = Trim(GEN!gen_cli_cias)
    MsgBox "O J O ...  Al Eliminar este Cliente tambien debe hacerlo con las demas Compañias relacionadas : " & wcias, 48, Pub_Titulo
  End If
'  If Trim(tcuenta.Text) <> "" And LK_EMP <> "CAM" Then
  '  pub_mensaje = " ¿Desea Eliminar el Registro, y su Relacion a Contabilidad .. ?"
'  Else
    pub_mensaje = " ¿Desea Eliminar el Registro... ?"
'  End If
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
    Screen.MousePointer = 11
    cliloc_llave.Delete
    pub_cadena = "Delete dircli where CODCIA = '" & LK_CODCIA & "' and  CODCLI = " & Trim(Txt_key.Text) & "  and   CP  = '" & Trim(Left(CmbCGP.Text, 1)) & "' "
    CN.Execute pub_cadena
    frmCLI.Txt_key.Text = ""
    frmCLI.Txt_key.Locked = False
    'If Trim(tcuenta.Text) <> "" And LK_EMP <> "CAM" Then
    '        SQ_OPER = 1
    '        PUB_CUENTA = Trim(tcuenta.Text)
    '        LEER_COM_LLAVE
    '        If com_llave.EOF Then
    '            tcuenta.Text = ""
    '        Else
    '            com_llave.Delete
    '            tcuenta.Text = ""
    '        End If
    'End If
    'If Trim(tcuenta2.Text) <> "" And LK_EMP <> "CAM" Then
    ' SQ_OPER = 1
    ' PUB_CUENTA = Trim(tcuenta2.Text)
    ' LEER_COM_LLAVE
    ' If com_llave.EOF Then
    '     tcuenta2.Text = ""
    ' Else
    '     com_llave.Delete
    '     tcuenta2.Text = ""
    ' End If
    'End If
   ' cmdcontab.Caption = "Relacionar a Con&tabilidad"
   ' cmdcontab2.Caption = "Relacionar a Con&tabilidad"
    LIMPIA_CLI
    MENSAJE_CLI "Registro   ELIMINADO ... "
    Screen.MousePointer = 0
  End If
  Screen.MousePointer = 0
Exit Sub
sale:
    MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    cmdcancelar_Click
    Screen.MousePointer = 0

End Sub

Private Sub cmdEliminar_GotFocus()
If ListView1.Visible Then
frmCLI.Txt_key.Text = ""
frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdEliminar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmCLI.Txt_key.SetFocus
End If

End Sub

Private Sub CmdEscapa_Click()
  frmCLI.F14.Visible = False
'  PASA = 0
  If frmCLI.txtesposo.Enabled Then
      frmCLI.txtesposo.SetFocus
  End If
End Sub

Private Sub cmdmante_Click()
pu_codclie = Val(Txt_key.Text)
If pu_codclie = 0 Then Exit Sub
'PUB_TIPREG = 2101
'PUB_CODCIA = LK_CODCIA
'Load FrmDatplac
'FrmDatplac.Caption = "Placas de Clientes : " & PUB_TIPREG
'FrmDatplac.Show 1
'cmdmante.SetFocus
'DoEvents

End Sub

Private Sub CmdModificar_Click()
If Len(Txt_key) = 0 Or Len(txtnombre) = 0 Then
   MENSAJE_CLI "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(CmdModificar.Caption, 2) = "&M" Then
    LOC_TIPOCLI = "A"
    CmdModificar.Caption = "&Grabar"
    cmdEliminar.Enabled = False
    cmdAgregar.Enabled = False
    cmdcancelar.Enabled = True
    CmbCGP.Enabled = False
    condi.Enabled = False
    'agregado
    'If Trim(Left$(CmbCGP, 1)) = "C" Then
    cmdcancel.Enabled = True
    cmdDelete.Enabled = True
    cmdDireccion.Enabled = True
    cboDireccion.Enabled = True
    'End If
    cboProvincia.Enabled = True
    
    DESBLOQUEA_TEXT
    lisdescto.Enabled = False
    If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
       lisdescto.Enabled = True
    End If
    Txt_key.Locked = True
    If frmCLI.txtesposo.Enabled Then frmCLI.txtesposo.SetFocus
    pasa = 1
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
    If Left(cmdcancel.Caption, 1) = "M" Then
       cmdcancel_Click
    End If

 Else
   If Left(CmbCGP.Text, 1) = "C" Then
      If pasa = 1 Then
         If EXISTE_CLI("C", Left(frmCLI.txtesposo.Text, 15), Trim(Txt_key.Text)) Then
            MENSAJE_CLI " Existen algunos clientes con estos NOMBRES .."
            frmCLI.ListExiste.SetFocus
            Exit Sub
         End If
      End If
      pasa = 0
   ElseIf Left(CmbCGP.Text, 1) = "P" Then
     If pasa = 1 Then
      If EXISTE_CLI("P", Left(frmCLI.txtesposo.Text, 15), Trim(Txt_key.Text)) Then
         MENSAJE_CLI " Existen algunos Proveedor con estos NOMBRES .."
         frmCLI.ListExiste.SetFocus
         Exit Sub
      End If
    End If
    pasa = 0
   End If
   If Not CONSIS_CLI Then
         '  "NO SE PUEDE.."
      Exit Sub
   End If
   If par_llave!PAR_CONTABILIDAD = "" Then
      GoTo PASACONTAB
   End If
   If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(CmbCGP.Text, 1) = "C" Then
      GoTo PASACONTAB
   ElseIf Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(CmbCGP.Text, 1) = "P" Then
      GoTo PASACONTAB
   End If
   If Left(cmdcontab.Caption, 2) <> "&Q" Or Left(cmdcontab2.Caption, 2) <> "&Q" Then
      If Trim(tcuenta.Text) <> "" Or Trim(tcuenta2.Text) <> "" Then
         GRABA_CONTAB LK_CODCIA
      End If
      
   End If
PASACONTAB:
   
    If Trim(tempo_ruc) <> Trim(txtRUCesposo.Text) Then
        pub_mensaje = "El Nro. R.U.C. ha cambiado, el sistema actualizará la información.  ¿Desea Continuar... ?"
        Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
        If Pub_Respuesta = vbNo Then
           Exit Sub
        End If
    End If
    Screen.MousePointer = 11
    GRABAR_CLI "C"
    MENSAJE_CLI "Registro , MODIFICADO... "
    CmdModificar.Caption = "&Modificar"
    frmCLI.SSTab1.tab = 0
    Screen.MousePointer = 0
    cmdcancelar.Enabled = True
    cmdEliminar.Enabled = True
    cmdAgregar.Enabled = True
    If Left(cmdcancel.Caption, 1) = "G" Then
       cmdcancel_Click
    End If
    BLOQUEA_TEXT
    
    Txt_key.Locked = True
    CmbCGP.Enabled = True
    cmdcancelar.SetFocus
    Screen.MousePointer = 0
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
  
End If
End Sub

Private Sub cmdModificar_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.Text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdprocli_Click()
'Load frmtipocal
'frmtipocal.Show 1
End Sub

Private Sub condi_Click()
 t_diasfac.Text = Val(Right(condi.Text, 6))
End Sub

Private Sub condi_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   Azul txtauto1, txtauto1
 End If
End Sub

Private Sub copia_Click()
Dim VALOR
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVIDOR" Then
Else
MsgBox "No procede , Solo el Administrador", 48, Pub_Titulo
Exit Sub
End If
pub_mensaje = "si .. GENERA CODIGO 0 "
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbYes Then
    SQ_OPER = 1
    pu_codclie = 0
    pu_cp = " "
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If Not cli_llave.EOF Then Exit Sub
    
    cliloc_llave.AddNew
    cliloc_llave!CLI_CP = " "
    cliloc_llave!cli_codclie = 0
    cliloc_llave!cli_SALDO = 0
    cliloc_llave!CLI_DET_TOT = "D"
    cliloc_llave!CLI_MONEDA = " "
    cliloc_llave!CLI_CODCIA = LK_CODCIA
    cliloc_llave!CLI_NOMBRE_ESPOSO = " "
    cliloc_llave!CLI_NOMBRE_ESPOSA = " "
    cliloc_llave!CLI_NOMBRE_EMPRESA = " "
    cliloc_llave!cli_nombre = " "
    cliloc_llave!CLI_CASA_DIREC = " "
    cliloc_llave!CLI_CASA_NUM = 0
    cliloc_llave!CLI_CASA_ZONA = 0
    cliloc_llave!CLI_LUGAR_CASA = 0
    cliloc_llave!CLI_LUGAR_TRAB = 0
    cliloc_llave!CLI_CASA_SUBZONA = 0
    cliloc_llave!cli_zona_new = 0
    cliloc_llave!CLI_TRAB_DIREC = " "
    cliloc_llave!CLI_TRAB_NUM = 0
    cliloc_llave!cli_TRAB_ZONA = 0
    cliloc_llave!cli_TRAB_SUBZONA = 0
    cliloc_llave!cli_ruc_esposo = " "
    cliloc_llave!cli_RUC_ESPOSA = " "
    cliloc_llave!CLI_RUC_EMPRESA = " "
    cliloc_llave!CLI_CASA1 = " "
    cliloc_llave!CLI_CASA2 = " "
    cliloc_llave!CLI_regpub1 = " "
    cliloc_llave!CLI_REGPUB2 = " "
    cliloc_llave!CLI_AUTOAVALUO = " "
    cliloc_llave!CLI_AUTO1 = " "
    cliloc_llave!cli_auto2 = " "
    cliloc_llave!CLI_PRENDA = " "
    cliloc_llave!cli_telef1 = " "
    cliloc_llave!CLI_TELEF2 = " "
    cliloc_llave!CLI_OTRO_CONTR = 0
    cliloc_llave!CLI_LETRA = 0
    cliloc_llave!CLI_GRUPO = 0
    cliloc_llave!CLI_SUBGRUPO = 0
    cliloc_llave!CLI_nucleo = 0
    cliloc_llave!cli_estado = " "
    cliloc_llave!CLI_programado = " "
    cliloc_llave!CLI_CUENTA_CONTAB = " "
    cliloc_llave!CLI_CUENTA_CONTAB2 = " "
    cliloc_llave!CLI_DET_TOT = " "
    cliloc_llave!CLI_LIMCRE = Val(txtlimite.Text)
    cliloc_llave.Update
    MsgBox "Cliente Creado ", 48, Pub_Titulo

   Exit Sub
End If

   If Val(frmCLI.Txt_key.Text) <= 0 Then
      MsgBox " Consulte  y  despues Copiar.."
      Exit Sub
   End If
   
    VALOR = InputBox("La Compañia a donde copiar los datos : ", "COMPAÑIA", "03")
    If VALOR = "" Then Exit Sub
    If Trim(VALOR) = LK_CODCIA Then
       MsgBox "No Procede .. "
       Exit Sub
    End If
    If Len(Trim(VALOR)) <> 2 Then
       MsgBox "de dos digitos  "
       Exit Sub
    End If
    
    cliloc_llave.AddNew
    cliloc_llave!CLI_CP = Left(CmbCGP.Text, 1)
    PUB_CODCIA = VALOR
    cliloc_llave!cli_codclie = GENERA_CODI
    PUB_CODCIA = LK_CODCIA
    cliloc_llave!cli_SALDO = 0
    cliloc_llave!CLI_DET_TOT = "D"
    cliloc_llave!CLI_MONEDA = " "
    cliloc_llave!CLI_CODCIA = VALOR
    cliloc_llave!CLI_NOMBRE_ESPOSO = txtesposo.Text
    cliloc_llave!CLI_NOMBRE_ESPOSA = Txtesposa.Text
    cliloc_llave!CLI_NOMBRE_EMPRESA = TxtEmpresa.Text
    ASIGNA_123
    cliloc_llave!cli_nombre = frmCLI.txtnombre.Text
    cliloc_llave!CLI_CASA_DIREC = txtdireccion.Text
    cliloc_llave!CLI_CASA_NUM = Val(Txtnumdir.Text)
    cliloc_llave!CLI_CASA_ZONA = Val(Right(TxtZona.Text, 4))
    cliloc_llave!CLI_LUGAR_CASA = Val(Right(TxtLugarCasa.Text, 4))
    cliloc_llave!CLI_LUGAR_TRAB = Val(Right(TxtLugarTrab.Text, 4))
    cliloc_llave!CLI_CASA_SUBZONA = Val(Right(TxtSubZona.Text, 4))
    cliloc_llave!cli_zona_new = Val(Right(txtZonaNew.Text, 4))
    cliloc_llave!CLI_TRAB_DIREC = txtDirTrabajo.Text
    cliloc_llave!CLI_TRAB_NUM = Nulo_Valor0(txtNumDirTrabajo.Text)
    cliloc_llave!cli_TRAB_ZONA = Val(Right(frmCLI.TxtZonaTrabajo.Text, 4))
    cliloc_llave!cli_TRAB_SUBZONA = Val(Right(TxtSubZonaTrabajo.Text, 4))
    cliloc_llave!cli_ruc_esposo = txtRUCesposo.Text
    cliloc_llave!cli_RUC_ESPOSA = txtRUCesposa.Text
    cliloc_llave!CLI_RUC_EMPRESA = txtRUCempresa.Text
    cliloc_llave!CLI_CASA1 = frmCLI.txtpropiedad1.Text
    cliloc_llave!CLI_CASA2 = frmCLI.txtpropiedad2.Text
    cliloc_llave!CLI_regpub1 = frmCLI.txtregpublico1.Text
    cliloc_llave!CLI_REGPUB2 = frmCLI.txtregpublico2.Text
    cliloc_llave!CLI_AUTOAVALUO = frmCLI.txtautovaluo.Text
    cliloc_llave!CLI_AUTO1 = frmCLI.txtauto1.Text
    cliloc_llave!cli_auto2 = frmCLI.txtauto2.Text
    cliloc_llave!CLI_PRENDA = frmCLI.txtprendas.Text
    cliloc_llave!cli_telef1 = frmCLI.txttelefono1.Text
    cliloc_llave!CLI_TELEF2 = frmCLI.txttelefono2.Text
    cliloc_llave!CLI_OTRO_CONTR = frmCLI.otrocontrato.Value
    cliloc_llave!CLI_LETRA = frmCLI.letraotorgado.Value
    cliloc_llave!CLI_GRUPO = Val(Right(frmCLI.cmbgrupo.Text, 4))
    cliloc_llave!CLI_SUBGRUPO = Val(Right(frmCLI.txtsubgrupo.Text, 4))
    cliloc_llave!CLI_nucleo = frmCLI.txtNucleo.Text
    cliloc_llave!cli_estado = Left(frmCLI.txtestado.Text, 1)
    cliloc_llave!CLI_programado = Nulo_Valors(txtprog.Text)
    '  <<< Actualiza La Cta. solo de la Cia Actual >>>
    cliloc_llave!CLI_CUENTA_CONTAB = Trim(frmCLI.tcuenta.Text)
    cliloc_llave!CLI_CUENTA_CONTAB2 = Trim(frmCLI.tcuenta2.Text)
    If txtDTX.Text = "" Then
      txtDTX.Text = " "
    End If
    cliloc_llave!CLI_DET_TOT = txtDTX.Text
    cliloc_llave!CLI_LIMCRE = Val(txtlimite.Text)
cliloc_llave.Update
MsgBox "Proceso Copiado .... ", 48, Pub_Titulo
Unload frmCLI
End Sub


Private Sub Form_Activate()
'frmCLI.Refresh
End Sub

Private Sub Form_DblClick()
If LK_CODUSU <> "ADMIN" Then Exit Sub



MsgBox "Solo Admin Chequea Relacion con Cli_Zona_new"

OPCINAL

 Exit Sub


Screen.MousePointer = 11
SQ_OPER = 2
pu_cp = Left(CmbCGP.Text, 1)
pu_codclie = 0
pu_codcia = LK_CODCIA

LEER_CLI_LLAVE
Do Until cli_mayor.EOF

SQ_OPER = 1
PUB_TIPREG = 222
PUB_NUMTAB = Nulo_Valor0(cli_mayor!CLI_GRUPO)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_llave.EOF Then
 'MsgBox "NO HAY " & cli_mayor!CLI_NOMBRE
 cli_mayor.Edit
 cli_mayor!CLI_GRUPO = 1
 cli_mayor.Update
End If
cli_mayor.MoveNext
Loop
Screen.MousePointer = 0
MsgBox "TERMINO"
Exit Sub


MsgBox "SOLO ADMINISTRADOR"

pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = ?   AND CLI_CP =  'C' ORDER BY  CLI_CODCLIE"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
llave_rep01.Requery
Do Until llave_rep01.EOF
Print llave_rep01!cli_codclie
 If IsNull(llave_rep01!cli_RUC_ESPOSA) Then
   llave_rep01.Edit
   llave_rep01!cli_RUC_ESPOSA = " "
   llave_rep01.Update
 End If
llave_rep01.MoveNext
Loop


MsgBox "TERMINO"



Exit Sub
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ?   AND ALL_FLAG_EXT <> 'E' AND ALL_CODCLIE <> 0   ORDER BY  ALL_CODCLIE"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
llave_rep01.Requery
Do Until llave_rep01.EOF
   cmdAgregar.Caption = llave_rep01.AbsolutePosition & "/ " & llave_rep01.RowCount
    DoEvents
    SQ_OPER = 1
    pu_codclie = llave_rep01!ALL_CODCLIE
    pu_cp = llave_rep01!ALL_CP
    pu_codcia = llave_rep01!ALL_CODCIA
    LEER_CLI_LLAVE
    If Not cli_llave.EOF Then
      llave_rep01.Edit
      llave_rep01!ALL_RUC = Trim(cli_llave!cli_ruc_esposo)
      llave_rep01.Update
    Else
      llave_rep01.Edit
      llave_rep01!ALL_RUC = " "
      llave_rep01.Update
    End If
'End If
llave_rep01.MoveNext
Loop
MsgBox "TERMINO.."

Exit Sub

pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_PRO >=?   AND ALL_FLAG_EXT <> 'E' AND ALL_CODCLIE <> 0   AND ALL_TIPMOV = 97 AND ALL_CP = 'P' AND ALL_CODTRA = 2410 ORDER BY  ALL_CODCLIE"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
'PS_REP01(1) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
  fila = 0
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = "01/06/01"
'PS_REP01(1) = LK_FECHA_DIA
llave_rep01.Requery
Do Until llave_rep01.EOF
   cmdAgregar.Caption = llave_rep01.AbsolutePosition & "/ " & llave_rep01.RowCount
    DoEvents
'If Trim(Nulo_Valors(llave_rep01!ALL_RUC)) = "" Then
pu_cp = llave_rep01!ALL_CP
pu_codclie = llave_rep01!ALL_CODCLIE
pu_codcia = llave_rep01!ALL_CODCIA
PUB_TIPDOC = llave_rep01!ALL_TIPDOC
PUB_FECHA = llave_rep01!ALL_FECHA_DIA
PUB_NUM_OPER = llave_rep01!ALL_NUMOPER
Print llave_rep01!ALL_CODTRA
SQ_OPER = 1
LEER_CAA_LLAVE
caa_histo.Requery
'fila = 2
If Not caa_histo.EOF Then
 If caa_histo!CAA_FECHA_COBRO <> llave_rep01!ALL_FECHA_SUNAT Then
  caa_histo.Edit
  caa_histo!CAA_FECHA_COBRO = llave_rep01!ALL_FECHA_SUNAT
  caa_histo.Update
  fila = fila + 1
 
'  MsgBox "" & llave_rep01!ALL_NUMFAC
 End If
End If

llave_rep01.MoveNext
Loop
MsgBox "TERMINO.." & fila
Exit Sub

Exit Sub
PASE

End Sub

Private Sub Form_Load()
Dim mst_vend As rdoResultset
Dim I As Integer
COD_ORIGINAL = 0
LOC_CTA_CLI = ""
LOC_DES_CLI = ""
LOC_CTA_CLI2 = ""
LOC_DES_CLI2 = ""
ntipocli = 0
If Not cop_llave.EOF Then
For I = 1 To 6
  If cop_llave.rdoColumns(I) <> 0 Then
     wCOM_NIVEL(I) = cop_llave.rdoColumns(I)
     NIVEL_MAX = I
  End If
Next I
End If

pub_cadena = "SELECT SUT_DESCRIPCION, SUT_SECUENCIA FROM SUB_TRANSA WHERE SUT_CODTRA = ?  ORDER BY SUT_SECUENCIA "
Set SUTRA = CN.CreateQuery("", pub_cadena)
SUTRA.rdoParameters(0) = 0
Set sutra_llave = SUTRA.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM PARGEN WHERE PAR_CODCIA = ?  ORDER BY PAR_CODCIA "
Set PSPAR_CLI = CN.CreateQuery("", pub_cadena)
PSPAR_CLI.rdoParameters(0) = " "
Set par_llave_cli = PSPAR_CLI.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  
pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP=? AND CLI_CODCLIE  = ? AND CLI_CODCIA = ? ORDER BY CLI_CP ,CLI_CODCLIE"
Set PSCLILOC_LLAVE = CN.CreateQuery("", pub_cadena)
PSCLILOC_LLAVE.rdoParameters(0) = " "
PSCLILOC_LLAVE.rdoParameters(1) = 0
PSCLILOC_LLAVE.rdoParameters(2) = " "
Set cliloc_llave = PSCLILOC_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP = ? AND CLI_CODCLIE  >= ? AND CLI_CODCIA = ? ORDER BY CLI_CP ,CLI_CODCLIE"
Set PSCLILOC_MAYOR = CN.CreateQuery("", pub_cadena)
PSCLILOC_MAYOR.rdoParameters(0) = " "
PSCLILOC_MAYOR.rdoParameters(1) = 0
PSCLILOC_MAYOR.rdoParameters(2) = " "
Set cliloc_mayor = PSCLILOC_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_CODCLIE  >= ? ORDER BY CLI_CP ,CLI_CODCLIE"
Set PSCLI_MAYOR2 = CN.CreateQuery("", pub_cadena)
PSCLI_MAYOR2.rdoParameters(0) = " "
PSCLI_MAYOR2.rdoParameters(1) = " "
PSCLI_MAYOR2.rdoParameters(2) = 0
Set cli_mayor2 = PSCLI_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ? AND TAB_CODCLIE = ? ORDER BY TAB_NOMLARGO"
Set PSPLAC_LLAVE = CN.CreateQuery("", pub_cadena)
PSPLAC_LLAVE.rdoParameters(0) = 0
PSPLAC_LLAVE.rdoParameters(1) = 0
PSPLAC_LLAVE.rdoParameters(2) = 0
Set cliplac_llave = PSPLAC_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)


'If LK_EMP = "HER" Then
 pub_cadena = "SELECT CLI_NOMBRE,CLI_RUC_ESPOSO FROM CLIENTES WHERE CLI_CODCIA = ? AND CLI_CP = ? AND CLI_RUC_ESPOSO = ? and CLI_CODCLIE <> ?  ORDER BY CLI_CODCLIE"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 PS_REP01.rdoParameters(0) = " "
 PS_REP01.rdoParameters(1) = " "
 PS_REP01.rdoParameters(2) = " "
 PS_REP01.rdoParameters(3) = 0
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
'End If

ACCESO_CLI
loc_ultcod = 0
' para el reporte de bancos...reinicio
frmCLI.F14.Left = 90
frmCLI.F14.Top = 3360
frmCLI.F14.Height = 2415
frmCLI.F14.Width = 9380
llave1 = ""
UNICO = ""
pasa = 0
For pasa = 0 To frmCLI.ListExiste.COL - 1
  frmCLI.ListExiste.COL = pasa
  frmCLI.ListExiste.FixedAlignment(pasa) = 2
Next pasa
pasa = 0
frmCLI.ListExiste.Cols = 5
frmCLI.ListExiste.ColWidth(0) = 350
frmCLI.ListExiste.ColWidth(1) = 800
frmCLI.ListExiste.ColWidth(2) = 3300
frmCLI.ListExiste.ColWidth(3) = 2000
frmCLI.ListExiste.ColWidth(4) = 3000
wGARANTES = 0
'ALLINVISIBLE
BLOQUEA_TEXT
LLENA_ZONA TxtZona, 20
LLENA_ZONA TxtSubZona, 30
LLENA_ZONA txtZonaNew, 35
LLENA_ZONA TxtZonaTrabajo, 20
LLENA_ZONA TxtSubZonaTrabajo, 35
LLENA_GRUPOS cmbgrupo, 222
LLENA_GRUPOS txtsubgrupo, 333
LLENA_GRUPOS Cmbcate, 230

'milton
LLENA_GRUPOS cmbtipocli, 144
'***********
LLENA_ZONA TxtLugarCasa, 25
'AGREGADO 29/11/2001
LLENA_ZONA cboProvincia, 30
LLENA_DIRECCIONES

'**************************
LLENA_ZONA TxtLugarTrab, 25
'frmCLI.SSTab1.TabCaption(0) = "&Datos Clientes - Principales"
'frmCLI.SSTab1.TabCaption(1) = "&Datos Clientes - Opcionales"
'llena dias de visita
pub_cadena = "SELECT VEM_CODCIA , VEM_CODVEN, VEM_NOMBRE FROM VEMAEST WHERE VEM_CODCIA = '" & LK_CODCIA & "' AND  VEM_DESACTIVO <> 'A' ORDER BY VEM_NOMBRE"
Set mst_vend = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
mst_vend.Requery
Do Until mst_vend.EOF
    cboDiaVisita.AddItem mst_vend!VEM_NOMBRE & String(80, " ") & mst_vend!VEM_CODVEN
  mst_vend.MoveNext
Loop
LOC_TIPREG = 300 ' CLIENTES
ETIQUETA_CLI
llave1 = "X"
CmbCGP.ListIndex = 0
llave1 = ""
Screen.MousePointer = 0
Txt_key.MaxLength = 15
cmdcontab.Enabled = False
t_diasfac.Visible = True



If LK_FLAG_GRIFO = "A" Then
    fraplaca.Visible = True
    cmdmante.Visible = True
    g_fechafac.Visible = True
    g_diasfac.Visible = True
    t_fechafac.Visible = True
Else
    cmdmante.Visible = False
    fraplaca.Visible = False
    g_fechafac.Visible = False
    g_diasfac.Visible = False
    t_fechafac.Visible = False
End If
SUTRA.rdoParameters(0) = 2401
sutra_llave.Requery
condi.Clear
condi.AddItem "Opcional" & String(60, " ") & "-1"
Do Until sutra_llave.EOF
condi.AddItem sutra_llave!sut_descripcion & String(70, " ") & Str(sutra_llave!sut_secuencia)
sutra_llave.MoveNext
Loop
' txtauto1.Locked = True
txtauto2.Locked = True
txtautovaluo.Visible = False
txtregpublico1.Visible = False
frmCLI.Txt_key.TabIndex = 0
copia.Visible = True
ws_Prov_bloq = ""
For fila = 1 To lk_OTROS_Count
   If Val(lk_OTROS(fila)) = 17 Then 'modificar NOMBRE DEL CLIENTE O PROVEEDORE
     ws_Prov_bloq = "A"
   End If
Next fila


Exit Sub
Resume Next
End Sub
'*******************************************
'AGREGADO:
'PROCEDIMIENTO PARA LLENAR LA S DIRECCIONES
'*******************************************
Sub LLENA_DIRECCIONES()
Dim SQL As String
'SQL = "select d.DIRCLI,d.DirComp,t.tab_nomlargo+' '+d.direc+' '+str(d.numero)as Direc " & _
"from dircli d inner join tablas t on d.cli_lugar_trab=t.tab_numtab " & _
"and tab_tipreg='25' " & _
"where d.codcia=? and d.codcli=? and d.cp=?"
SQL = "select d.DIRCLI,d.DirComp " & _
"from dircli  d " & _
"where d.codcia=? and d.codcli=? and d.cp=?"

'sql = "SELECT * FROM DIRCLI WHERE CODCIA=? AND CODCLI=? AND CP=?"
 Set PS_dir = CN.CreateQuery("", SQL)
  PS_dir.rdoParameters(0) = " "
  PS_dir.rdoParameters(1) = 0
  PS_dir.rdoParameters(2) = " "
  Set llave_dir = PS_dir.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  PS_dir(0) = LK_CODCIA
  PS_dir(1) = Val(frmCLI.Txt_key)
  If Trim(Left$(CmbCGP.Text, 1)) = "C" Then
  PS_dir(2) = "C"
  Else
  PS_dir(2) = "P"
  End If
  llave_dir.Requery
  If llave_dir.EOF Then
   cboDireccion.Clear
   Exit Sub
  End If
  cboDireccion.Clear
  Do Until llave_dir.EOF
   cboDireccion.AddItem llave_dir!DIRCOMP
   cboDireccion.ItemData(cboDireccion.NewIndex) = Val(llave_dir!DIRCLI)
   llave_dir.MoveNext
  Loop
  'SELECCIONAR EL PRIMER ITEM
  cboDireccion.ListIndex = 0
End Sub
Public Sub ALLINVISIBLE()
    frmCLI.lcuenta.Visible = False
    Txt_key.Visible = False
    txtnombre.Visible = False
    txtesposo.Visible = False
    Txtesposa.Visible = False
    TxtEmpresa.Visible = False
    txtdireccion.Visible = False
    Txtnumdir.Visible = False
    TxtZona.Visible = False
    TxtSubZona.Visible = False
    txtZonaNew.Visible = False
    txtDirTrabajo.Visible = False
    txtNumDirTrabajo.Visible = False
    frmCLI.TxtZonaTrabajo.Visible = False
    TxtSubZonaTrabajo.Visible = False
    txtRUCesposo.Visible = False
    txtRUCesposa.Visible = False
    txtRUCempresa.Visible = False
    frmCLI.txtpropiedad2.Visible = False
    frmCLI.txtpropiedad1.Visible = False
    frmCLI.txtregpublico1.Visible = False
    frmCLI.txtregpublico2.Visible = False
    frmCLI.txtautovaluo.Visible = False
    frmCLI.txtauto1.Visible = False
    frmCLI.txtauto2.Visible = False
    frmCLI.txtprendas.Visible = False
    frmCLI.txttelefono1.Visible = False
    frmCLI.txttelefono2.Visible = False
    frmCLI.otrocontrato.Visible = False
    frmCLI.letraotorgado.Visible = False
    frmCLI.ListBloqueos.Visible = False
    frmCLI.OptNombre(0).Visible = False
    frmCLI.OptNombre(1).Visible = False
    frmCLI.OptNombre(2).Visible = False
    frmCLI.cmbgrupo.Visible = False
    frmCLI.txtsubgrupo.Visible = False
    frmCLI.txtNucleo.Visible = False
    frmCLI.tcuenta.Visible = False
    frmCLI.txtestado.Visible = False
    frmCLI.TxtLugarCasa.Visible = False
    frmCLI.TxtLugarTrab.Visible = False
    frmCLI.txtpordes.Visible = False
    t_diascred.Visible = False
End Sub
Public Sub ALLVISIBLE()
    frmCLI.lcuenta.Visible = True
    Txt_key.Visible = True
    txtnombre.Visible = True
    txtesposo.Visible = True
    Txtesposa.Visible = True
    TxtEmpresa.Visible = True
    txtdireccion.Visible = True
    Txtnumdir.Visible = True
    TxtZona.Visible = True
    TxtSubZona.Visible = True
    txtZonaNew.Visible = True
    txtDirTrabajo.Visible = True
    txtNumDirTrabajo.Visible = True
    frmCLI.TxtZonaTrabajo.Visible = True
    TxtSubZonaTrabajo.Visible = True
    txtRUCesposo.Visible = True
    txtRUCesposa.Visible = True
    txtRUCempresa.Visible = True
    frmCLI.txtpropiedad2.Visible = True
    frmCLI.txtpropiedad1.Visible = True
    frmCLI.txtregpublico1.Visible = True
    frmCLI.txtregpublico2.Visible = True
    frmCLI.txtautovaluo.Visible = True
    frmCLI.txtauto1.Visible = True
    frmCLI.txtauto2.Visible = True
    frmCLI.txtprendas.Visible = True
    frmCLI.txttelefono1.Visible = True
    frmCLI.txttelefono2.Visible = True
    frmCLI.otrocontrato.Visible = True
    frmCLI.letraotorgado.Visible = True
    frmCLI.ListBloqueos.Visible = True
    frmCLI.OptNombre(0).Visible = True
    frmCLI.OptNombre(1).Visible = True
    frmCLI.OptNombre(2).Visible = True
    frmCLI.cmbgrupo.Visible = True
    frmCLI.txtsubgrupo.Visible = True
    frmCLI.txtNucleo.Visible = True
    frmCLI.tcuenta.Visible = True
    frmCLI.txtestado.Visible = True
    frmCLI.TxtLugarCasa.Visible = True
    frmCLI.TxtLugarTrab.Visible = True
    frmCLI.txtpordes.Visible = True
    t_diascred.Visible = True
End Sub

Public Sub LLENA_123()
  If cliloc_llave!CLI_123 = 1 Then
       frmCLI.OptNombre(0).Value = True
  ElseIf cliloc_llave!CLI_123 = 2 Then
       frmCLI.OptNombre(1).Value = True
  ElseIf cliloc_llave!CLI_123 = 3 Then
       frmCLI.OptNombre(2).Value = True
       Exit Sub
  End If

End Sub
Public Sub ASIGNA_123()
  If frmCLI.OptNombre(0).Value Then
     frmCLI.txtnombre.Text = Nulo_Valors(frmCLI.txtesposo.Text)
     cliloc_llave!CLI_123 = 1
  ElseIf frmCLI.OptNombre(1).Value Then
     frmCLI.txtnombre.Text = Nulo_Valors(frmCLI.Txtesposa.Text)
     cliloc_llave!CLI_123 = 2
  ElseIf frmCLI.OptNombre(2).Value Then
     frmCLI.txtnombre.Text = Nulo_Valors(frmCLI.TxtEmpresa.Text)
     cliloc_llave!CLI_123 = 3
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'cli_mayor2.Close
 fila = 0
 pub_cadena = ""
End Sub


Private Sub Label3_DblClick(Index As Integer)
Dim WGUARDA_IMP As Currency
If LK_CODUSU <> "ADMIN" Then Exit Sub
Stop
pub_cadena = "SELECT * FROM ALLOG WHERE  ALL_CODTRA = 2735 AND (ALL_CODCIA = '01' OR ALL_CODCIA = '02') AND ALL_FLAG_EXT <> 'E' AND ALL_MONEDA_CLI = 'D' ORDER BY ALL_FECHA_DIA, ALL_NUMOPER "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
llave_rep01.Requery
Print llave_rep01.RowCount
Do Until llave_rep01.EOF
If llave_rep01!ALL_SIGNO_CAR = -1 Then
  WGUARDA_IMP = Val(llave_rep01!ALL_NUMDOC)
End If
  llave_rep01.Edit
  llave_rep01!ALL_NUMGUIA = WGUARDA_IMP
  llave_rep01.Update
'End If

llave_rep01.MoveNext
Loop

MsgBox "TERMINO "
Exit Sub

pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FLAG_EXT <> 'E' AND ALL_CODTRA = 5318 ORDER BY ALL_FECHA_DIA , ALL_NUMOPER"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = LK_CODCIA
llave_rep01.Requery

Do Until llave_rep01.EOF
If llave_rep01!ALL_SIGNO_CCM = -1 Then
  WGUARDA_IMP = Val(llave_rep01!ALL_IMPORTE)
Else
  llave_rep01.Edit
  llave_rep01!ALL_IMPORTE = WGUARDA_IMP
  llave_rep01.Update
End If

llave_rep01.MoveNext
Loop

MsgBox "TERMINO "
Exit Sub



MsgBox "Solo Admin Chequea Relacion con Cli_Zona_new"
Screen.MousePointer = 11
SQ_OPER = 2
pu_cp = Left(CmbCGP.Text, 1)
pu_codclie = 0
pu_codcia = LK_CODCIA

LEER_CLI_LLAVE
Do Until cli_mayor.EOF

SQ_OPER = 1
PUB_TIPREG = 35
PUB_NUMTAB = Nulo_Valor0(cli_mayor!cli_zona_new)
PUB_CODCIA = "00"
LEER_TAB_LLAVE
If tab_llave.EOF Then
cli_mayor.Edit
cli_mayor!cli_zona_new = 4
cli_mayor.Update
End If
cli_mayor.MoveNext
Loop
Screen.MousePointer = 0
End Sub

Private Sub lblnom_DblClick(Index As Integer)
If Trim(LK_CODUSU) <> "ADMIN" And Trim(LK_CODUSU) <> "SUPERVISOR" Then
 Exit Sub
End If
If Trim(lblnom(Index).Tag) = "" Then
 Exit Sub
End If
Dim wnombre
wnombre = InputBox("Ingrese la Nueva Descripción para este Campo :", Pub_Titulo, Trim(lblnom(Index).Caption))
If wnombre = "" Then
  Screen.MousePointer = 0
  Exit Sub
End If
Screen.MousePointer = 11
SQ_OPER = 1
PUB_TIPREG = LOC_TIPREG
PUB_NUMTAB = Val(lblnom(Index).Tag)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_llave.EOF Then
  tab_llave.AddNew
Else
  tab_llave.Edit
End If
  tab_llave!tab_codcia = LK_CODCIA
  tab_llave!tab_tipreg = LOC_TIPREG
  tab_llave!TAB_NUMTAB = Val(lblnom(Index).Tag)
  tab_llave!TAB_NOMLARGO = Left(wnombre, 40)
  tab_llave!tab_nomcorto = Left(wnombre, 10)
  tab_llave.Update
  lblnom(Index).Caption = Left(wnombre, 40)
Screen.MousePointer = 0
End Sub

Private Sub letraotorgado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.letraotorgado.TabIndex
End If

End Sub

Private Sub lisdescto_ItemCheck(Item As Integer)
  If Screen.MousePointer = 11 Then Exit Sub
  If Val(Trim(Right(lisdescto.List(Item), 6))) = Val(Trim(Right(txtsubgrupo.Text, 6))) Then
      lisdescto.Selected(Item) = True
      MsgBox "Esta Cambiando la opcion por Defecto. ", 48, Pub_Titulo
      Exit Sub
  End If

End Sub

Private Sub ListExiste_Click()
Dim d, C, a As Integer
End Sub
Private Sub ListExiste_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmCLI.F14.Visible = False
     frmCLI.txtesposo.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub ListExiste_LostFocus()
If frmCLI.ListExiste.Visible = False Then
    Exit Sub
End If
End Sub

Private Sub ListView1_DblClick()
 loc_key = ListView1.SelectedItem.Index
 Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 txt_key_KeyPress 13
End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.Item(loc_key).Selected = True
 ListView1.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView1.Visible = False
 Txt_key.Text = ""
 Txt_key.SetFocus
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

Private Sub lstVen_KeyPress(KeyAscii As Integer)
Dim I
If KeyAscii = 27 Then
  lstVen.Visible = False
  tcuenta22.SetFocus
  Exit Sub
End If

If KeyAscii = 13 Then
   tcuenta22.Text = ""
   tcuenta22.Text = Left(tcuenta22.Text, 4)
   For I = 0 To lstVen.ListCount - 1
     lstVen.ListIndex = I
     If lstVen.Selected(I) Then
         tcuenta22.Text = tcuenta22.Text & "." & Trim(Left(lstVen.Text, 3))
     End If
   Next I
   lstVen.Visible = False
End If

End Sub

Private Sub otrocontrato_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.otrocontrato.TabIndex
End If
End Sub

Private Sub PARPADEA_Timer()
 CU = CU + 1
 LblMensaje.Visible = True
 If CU > 4 Then
   CU = 0
   PARPADEA.Enabled = False
   LblMensaje.Visible = False
 End If
End Sub

Public Sub ASIGNA_INT(WCONTROL As ComboBox, txt As Integer)
For fila = 0 To WCONTROL.ListCount - 1
    If Val(Trim(Right(WCONTROL.List(fila), 8))) = txt Then
        WCONTROL.ListIndex = fila
        Exit Sub
    End If
Next fila
End Sub
Public Sub ASIGNA_subgrupo(WCONTROL As ComboBox, txt As String)
For fila = 0 To WCONTROL.ListCount - 1
    If Val(Trim(Right(WCONTROL.List(fila), 3))) = Val(txt) Then
        WCONTROL.ListIndex = fila
        Exit Sub
    End If
Next fila
End Sub

Public Sub LLENA_CLI(ban As Integer, CG As String)
Dim ws_flag As String
Dim key_descto As String * 2
Dim I As Integer
Screen.MousePointer = 11
    If ban = 0 Then
        '**  BAN = 0 BUSCA DATOS NUEVAMENTE
        If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
         Else
          Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
        End If
        pu_cp = Left(CmbCGP.Text, 2)
        pu_codclie = Val(Txt_key.Text)
        SQ_OPER = 1
        pu_codcia = LK_CODCIA
        LEER_CLILOC_LLAVE
    End If
    loc_ultcod = Val(cliloc_llave!cli_codclie)
    frmCLI.Txt_key.Text = cliloc_llave!cli_codclie
    LLENA_123
    txtnombre.Text = Nulo_Valors(cliloc_llave!cli_nombre)
    txtnombre.MaxLength = cliloc_llave(3).Size
    txtesposo.Text = Trim(Nulo_Valors(cliloc_llave!CLI_NOMBRE_ESPOSO))
    txtesposo.MaxLength = cliloc_llave(4).Size
    Txtesposa.Text = Trim(Nulo_Valors(cliloc_llave!CLI_NOMBRE_ESPOSA))
    TxtEmpresa.Text = Trim(Nulo_Valors(cliloc_llave!CLI_NOMBRE_EMPRESA))
    txtdireccion.Text = Trim(Nulo_Valors(cliloc_llave!CLI_CASA_DIREC))
    txtdireccion.MaxLength = cliloc_llave(10).Size
    Txtnumdir.Text = Trim(Nulo_Valor0(cliloc_llave!CLI_CASA_NUM))
    ASIGNA_INT TxtZona, Nulo_Valor0(cliloc_llave!CLI_CASA_ZONA)
    ASIGNA_INT TxtSubZona, Nulo_Valor0(cliloc_llave!CLI_CASA_SUBZONA)
    ASIGNA_INT txtZonaNew, Nulo_Valor0(cliloc_llave!cli_zona_new)
    'QUITADO
    txtDirTrabajo.Text = Trim(Nulo_Valors(cliloc_llave!CLI_TRAB_DIREC))
    txtDirTrabajo.MaxLength = cliloc_llave(14).Size
    txtNumDirTrabajo.Text = Trim(Nulo_Valor0(cliloc_llave!CLI_TRAB_NUM))
    'OJOQUITADO
    ASIGNA_INT TxtZonaTrabajo, Nulo_Valor0(cliloc_llave!cli_TRAB_ZONA)
    'AGREGADO PARA SIGNAR LA PROVINCIA 06/12/2001
    ASIGNA_INT cboProvincia, Nulo_Valor0(cliloc_llave!cli_TRAB_PROV)
    
    ASIGNA_INT TxtSubZonaTrabajo, Nulo_Valor0(cliloc_llave!cli_TRAB_SUBZONA)
    ASIGNA_INT TxtLugarCasa, Nulo_Valor0(cliloc_llave!CLI_LUGAR_CASA)
    'OJO QUITADO
    
    ASIGNA_INT TxtLugarTrab, Nulo_Valor0(cliloc_llave!CLI_LUGAR_TRAB)
    ASIGNA_INT cboDiaVisita, Nulo_Valor0(cliloc_llave!cli_dia_visita)
    
    'AGREGADO 29/11/2001
    LLENA_DIRECCIONES
    
    txtRUCesposo.Text = Trim(Nulo_Valors(cliloc_llave!cli_ruc_esposo))
    tempo_ruc = Trim(Nulo_Valors(cliloc_llave!cli_ruc_esposo))
    If LK_DIG_RUC <> 0 Then txtRUCesposo.MaxLength = LK_DIG_RUC
    txtRUCesposa.Text = Trim(Nulo_Valors(cliloc_llave!cli_RUC_ESPOSA))
    txtRUCempresa.Text = Trim(Nulo_Valors(cliloc_llave!CLI_RUC_EMPRESA))
    frmCLI.txtpropiedad1.Text = Trim(Nulo_Valors(cliloc_llave!CLI_CASA1))
    frmCLI.txtpropiedad2.Text = Trim(Nulo_Valors(cliloc_llave!CLI_CASA2))
    frmCLI.txtregpublico1.Text = Trim(Nulo_Valors(cliloc_llave!CLI_regpub1))
    frmCLI.txtregpublico2.Text = Trim(Nulo_Valors(cliloc_llave!CLI_REGPUB2))
    frmCLI.txtautovaluo.Text = Trim(Nulo_Valors(cliloc_llave!CLI_AUTOAVALUO))
    frmCLI.txtauto1.Text = Trim(Nulo_Valors(cliloc_llave!CLI_AUTO1))
    frmCLI.txtauto2.Text = Trim(Nulo_Valors(cliloc_llave!cli_auto2))
    frmCLI.txtprendas.Text = Trim(Nulo_Valors(cliloc_llave!CLI_PRENDA))
    frmCLI.txttelefono1.Text = Trim(Nulo_Valors(cliloc_llave!cli_telef1))
    frmCLI.txttelefono2.Text = Trim(Nulo_Valors(cliloc_llave!CLI_TELEF2))
    frmCLI.otrocontrato.Value = Nulo_Valor0(cliloc_llave!CLI_OTRO_CONTR)
    frmCLI.letraotorgado.Value = Nulo_Valor0(cliloc_llave!CLI_LETRA)
    LLENA_BLOQ
    ASIGNA_INT cmbgrupo, Nulo_Valors(cliloc_llave!CLI_GRUPO)
    ASIGNA_INT Cmbcate, Nulo_Valor0(cliloc_llave!CLI_division)
    ASIGNA_subgrupo txtsubgrupo, Nulo_Valors(cliloc_llave!CLI_SUBGRUPO)
    frmCLI.txtNucleo.Text = Nulo_Valor0(cliloc_llave!CLI_nucleo)
    ASIGNA_INT condi, Nulo_Valor0(cliloc_llave!cli_DIAS_FAC)
    ASIGNA_INT cmbtipocli, Nulo_Valor0(cliloc_llave!CLI_TIPOCLI) ' MILTON
    If Nulo_Valors(cliloc_llave!cli_estado) = "A" Then
      frmCLI.txtestado.ListIndex = 0
    Else
    frmCLI.txtestado.ListIndex = 1
    End If
    frmCLI.t_ciarela.Text = Nulo_Valors(cliloc_llave!CLI_CIA_REF)
    frmCLI.txtprog.Text = Nulo_Valors(cliloc_llave!CLI_programado)
    frmCLI.tcuenta.Text = Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB)
    frmCLI.txtciarela.Text = Trim(cliloc_llave!CLI_CIARELA)
    If Trim(Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB)) <> "" Then
        cmdcontab.Caption = "&Quitar Relacion Contable"
    Else
        cmdcontab.Caption = "Relacionar a Con&tabilidad"
    End If
    frmCLI.tcuenta2.Text = Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB2)
    frmCLI.tcuenta22.Text = Nulo_Valors(cliloc_llave!cli_cuenta_contab22)
    If Trim(Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB2)) <> "" Then
        cmdcontab2.Caption = "&Quitar Relacion Contable"
    Else
        cmdcontab2.Caption = "Relacionar a Con&tabilidad"
    End If
    frmCLI.txtlimite.Text = Nulo_Valor0(cliloc_llave!CLI_LIMCRE)
    txtDTX.Text = Nulo_Valors(cliloc_llave!CLI_DET_TOT)
    frmCLI.txtpordes.Text = Nulo_Valor0(cliloc_llave!CLI_PORDESCTO)
    t_fechafac.Text = Format(cliloc_llave!cli_fecha_fac, "dd/mm/yyyy")
    t_diasfac.Text = Nulo_Valor0(cliloc_llave!cli_DIAS_FAC)
    frmCLI.t_diascred.Text = Nulo_Valor0(cliloc_llave!cli_dias_cred)
    frmCLI.fechahora.Caption = Trim(cliloc_llave!cli_fechahora)
    pu_codclie = Val(Txt_key.Text)
    If LK_FLAG_GRIFO = "A" Then
      LLENA_DESCTO
    End If
   '******* numtab tipo cliente
   ntipocli = Val(Right(Trim(cmbtipocli.Text), 4))
   '***************************
   For fila = 1 To 30 Step 2
     key_descto = Mid(Trim(Nulo_Valors(cliloc_llave!CLI_CASA1)), fila, 2)
     If Trim(key_descto) = "" Then Exit For
     If lisdescto.ListCount > 0 Then
      For I = 0 To lisdescto.ListCount - 1
        If Val(Right(lisdescto.List(I), 6)) = Val(key_descto) Then
          lisdescto.Selected(I) = True
        End If
      Next I
     End If
   Next fila
   
ws_flag = ""
For fila = 1 To lk_OTROS_Count
   If Val(lk_OTROS(fila)) = 7 Then 'modificar NOMBRE DEL CLIENTE O PROVEEDORE
     ws_flag = "A"
   End If
Next fila
If ws_flag = "A" Then
  txtesposo.Locked = False
  Frame2.Enabled = True
  SSTab1.Enabled = True
Else
  txtesposo.Locked = True
  Frame2.Enabled = False
  SSTab1.Enabled = False
End If

   Screen.MousePointer = 0
End Sub

Public Sub LIMPIA_CLI()
    t_ciarela.Text = ""
    fechahora.Caption = ""
    tempo_ruc = ""
    Txt_key.Text = ""
    txtnombre.Text = ""
    txtesposo.Text = ""
    Txtesposa.Text = ""
    TxtEmpresa.Text = ""
    txtdireccion.Text = ""
    Txtnumdir.Text = ""
    TxtZona.ListIndex = -1
    TxtSubZona.ListIndex = -1
    txtZonaNew.ListIndex = -1
    'agregado
    cboProvincia.ListIndex = -1
    cboDireccion.Clear
    
    TxtLugarCasa.ListIndex = -1
    TxtLugarTrab.ListIndex = -1
    cboDiaVisita.ListIndex = -1
    txtDirTrabajo.Text = ""
    txtNumDirTrabajo.Text = ""
    frmCLI.TxtZonaTrabajo.ListIndex = -1
    TxtSubZonaTrabajo.ListIndex = -1
    txtRUCesposo.Text = ""
    txtRUCesposa.Text = ""
    txtRUCempresa.Text = ""
    frmCLI.txtpropiedad2.Text = ""
    frmCLI.txtpropiedad1.Text = ""
    frmCLI.txtregpublico1.Text = ""
    frmCLI.txtregpublico2.Text = ""
    frmCLI.txtautovaluo.Text = ""
    frmCLI.txtauto1.Text = ""
    frmCLI.txtauto2.Text = ""
    frmCLI.txtprendas.Text = ""
    frmCLI.txttelefono1.Text = ""
    frmCLI.txttelefono2.Text = ""
    frmCLI.otrocontrato.Value = 0
    frmCLI.letraotorgado.Value = 0
    frmCLI.ListBloqueos.Clear
    frmCLI.cmbgrupo.ListIndex = -1
    frmCLI.txtsubgrupo.ListIndex = -1
    frmCLI.txtNucleo.Text = ""
    frmCLI.txtestado.ListIndex = -1
    frmCLI.tcuenta.Text = ""
    frmCLI.txtciarela.Text = ""
    frmCLI.OptNombre(0).Value = False
    frmCLI.OptNombre(1).Value = False
    frmCLI.OptNombre(2).Value = False
    frmCLI.txtlimite.Text = ""
    frmCLI.txtDTX = ""
    txtprog.Text = ""
    LOC_CTA_CLI = ""
    LOC_DES_CLI = ""
    tcuenta2.Text = ""
    LOC_CTA_CLI2 = ""
    LOC_DES_CLI2 = ""
    cmdcontab.Caption = "Relacionar a Con&tabilidad"
    cmdcontab2.Caption = "Relacionar a Con&tabilidad"
    frmCLI.txtpordes.Text = ""
    COD_ORIGINAL = 0
    t_fechafac.Text = ""
    t_diasfac.Text = ""
    t_diascred.Text = ""
    tcuenta22.Text = ""
    frmCLI.grid_des.Clear
    frmCLI.condi.ListIndex = -1
    frmCLI.grid_des.Rows = 1
    Cmbcate.ListIndex = -1
    For fila = 0 To lisdescto.ListCount - 1
      lisdescto.Selected(fila) = False
    Next fila
    txtesposo.Locked = False
    
    cmbtipocli.ListIndex = -1
    ntipocli = 0
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.tab = 0 Then
 If txtesposo.Enabled And txtesposo.Visible Then
   txtesposo.SetFocus
 End If
Else
 If txtDirTrabajo.Enabled And txtDirTrabajo.Visible Then
   txtDirTrabajo.SetFocus
 End If
End If
End Sub

Private Sub SSTab1_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.Text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub t_diascred_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
    frmCLI.txtdireccion.SetFocus
    Exit Sub
    If frmCLI.CmdModificar.Enabled Then
          frmCLI.CmdModificar.SetFocus
    Else
           frmCLI.cmdAgregar.SetFocus
    End If
 End If
End Sub

Private Sub t_diasfac_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  frmCLI.t_fechafac.SetFocus
End If

End Sub

Private Sub t_fechacred_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii

End Sub

Private Sub t_fechafac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  frmCLI.t_diascred.SetFocus
End If

End Sub

Private Sub tcuenta_GotFocus()
Azul tcuenta, tcuenta
End Sub

Private Sub tcuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If Trim(Left(tcuenta.Text, 1)) = "*" Then
       BUSCAR_CTA 1, tcuenta
       Exit Sub
     End If
     Azul tcuenta2, tcuenta2
End If
End Sub

Private Sub tcuenta2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If Trim(Left(tcuenta2.Text, 1)) = "*" Then
       BUSCAR_CTA 1, tcuenta2
       Exit Sub
     End If
     Azul tcuenta22, tcuenta22
End If
End Sub

Private Sub tcuenta22_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    LLENA_VEN 1
    lstVen.Visible = True
    lstVen.SetFocus
End If

End Sub

Private Sub txt_key_GotFocus()
 Azul Txt_key, Txt_key
End Sub

Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo sale
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If KeyCode = 113 Then
 If CmbCGP.ListIndex = 1 Then
  CmbCGP.ListIndex = 0
 Else
  CmbCGP.ListIndex = 1
 End If
End If
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And Txt_key.Text = "" Then
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
  Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  Txt_key.SelStart = Len(Txt_key.Text)
fin:
Exit Sub
sale:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim var As String
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As ListItem
On Error GoTo SALCODI
If KeyAscii = 13 And Val(Txt_key.Text) = 1 Then
    Txt_key.Text = ""
    Exit Sub
End If
If KeyAscii = 13 Then
If LK_EMP = "PAR" And Left(cmdAgregar.Caption, 2) = "&G" Then
  Azul txtesposo, txtesposo
  Exit Sub
End If
End If
If KeyAscii = 13 And Left(Txt_key.Text, 1) = "+" Then GoTo buscar
If KeyAscii = 27 And Trim(txtnombre.Text) = "" Then
 Txt_key.Text = ""
End If
If KeyAscii <> 13 Or Left(cmdAgregar.Caption, 2) = "&G" Or Left(CmdModificar.Caption, 2) = "&G" Then
   GoTo fin
End If
   
On Error GoTo CODI_ERR
pu_codclie = Val(Txt_key.Text)
On Error GoTo 0
If Len(Txt_key.Text) = 0 Then
   Exit Sub
End If
'fra2.Refresh
If pu_codclie <> 0 And IsNumeric(Txt_key.Text) = True Then
   If Len(Trim(Txt_key.Text)) = LK_DIG_RUC Then ' LONG DEL RUC
        pu_cp = Left(CmbCGP.Text, 1)
        PUB_RUC = Trim(Txt_key.Text)
        SQ_OPER = 4
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_ruc.EOF Then
           MsgBox "R.U.C. No Existe ", 48, Pub_Titulo
           Exit Sub
        End If
        Txt_key.Text = cli_ruc!cli_codclie
   End If
    SQ_OPER = 1
   On Error GoTo mucho
   pu_codcia = LK_CODCIA
   pu_cp = Left(CmbCGP.Text, 1)
   pu_codclie = Val(Txt_key.Text)
   LEER_CLILOC_LLAVE
   On Error GoTo 0
   If cliloc_llave.EOF Then
     Azul Txt_key, Txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Txt_key.SetFocus
     GoTo fin
   End If
   Screen.MousePointer = 11
   ListView1.Visible = False
   cmdcancelar.Enabled = True
   If Left(CmbCGP.Text, 1) = "C" Then
         LLENA_CLI 1, "C"
   End If
   If Left(CmbCGP.Text, 1) = "P" Then
         LLENA_CLI 1, "P"
   End If
   frmCLI.Txt_key.Locked = True
   frmCLI.CmdModificar.SetFocus
   Screen.MousePointer = 0
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(Txt_key.Text)) = Left(VALOR, Len(Trim(Txt_key.Text))) Then
   Else
      Exit Sub
   End If
   ListView1.Visible = False
   cmdcancelar.Enabled = True
   If Left(CmbCGP.Text, 1) = "C" Then
         LLENA_CLI 0, "C"
   End If
   If Left(CmbCGP.Text, 1) = "P" Then
         LLENA_CLI 0, "P"
   End If
   frmCLI.Txt_key.Locked = True
   cmdcancelar.Enabled = True
   frmCLI.CmdModificar.SetFocus
End If
dale:
ListView1.Visible = False
fin:
mucho:
CODI_ERR:
Exit Sub

buscar:
var = Mid(Txt_key.Text, 2, Len(Txt_key.Text))
numarchi = alta_vista_nombre(ListView1, var, Left(CmbCGP.Text, 1))
If numarchi = 0 Then
  ListView1.Visible = False
  MsgBox "Alta Vista: No Existe .. Esta descripcion..", 48, Pub_Titulo
Else
  ListView1.Visible = True
  Txt_key.SetFocus
End If
loc_key = 1
Exit Sub
SALCODI:
MsgBox Err.Description & " Intente Nuevamente ", 48, Pub_Titulo
Unload frmCLI
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim NADA
Dim var
If Len(Txt_key.Text) = 0 Or IsNumeric(Txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(Txt_key.Text) = 1 Then
    If Txt_key.Text = "" Then Txt_key.Text = " "
    var = Asc(Txt_key.Text)
    var = var + 1
    NADA = var
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 1
    'archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM FROM CLIENTES WHERE  CLI_CP = '" & Left(CmbCGP.Text, 1) & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_key.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    archi = "SELECT TOP 2000 CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = '" & Left(CmbCGP.Text, 1) & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY CLI_NOMBRE"
    PROC_LISVIEW ListView1
    loc_key = 1
    If NADA = 33 Or NADA = 91 Then
      If ListView1.Visible = False Then
        loc_key = 0
        MsgBox "No existe Datos ...", 48, Pub_Titulo
        Txt_key.Text = ""
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
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(Txt_key.Text), lvwText, , lvwPartial)
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

Private Sub txtauto1_GotFocus()
Azul txtauto1, txtauto1
End Sub

Private Sub txtauto1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TxtLugarCasa.SetFocus
   ''SIGUE_CAMPO frmCLI.txtauto1.TabIndex
End If
End Sub

Private Sub txtauto2_GotFocus()
Azul txtauto2, txtauto2
End Sub

Private Sub txtauto2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SIGUE_CAMPO frmCLI.txtauto2.TabIndex
End If
End Sub

Private Sub txtautovaluo_GotFocus()
Azul txtautovaluo, txtautovaluo
End Sub

Private Sub txtautovaluo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SIGUE_CAMPO frmCLI.txtautovaluo.TabIndex
End If
End Sub

Private Sub txtciarela_KeyPress(KeyAscii As Integer)
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
Else
KeyAscii = 0
End If

End Sub

Private Sub Txtdireccion_GotFocus()
Azul txtdireccion, txtdireccion

End Sub

Private Sub txtdireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   frmCLI.Txtnumdir.SetFocus
End If
End Sub

Private Sub Txtdireccion_LostFocus()
'If Left(CmbCGP.Text, 1) = "C" Then
'  If Trim(txtDirTrabajo.Text) = "" Then
    txtDirTrabajo.Text = Trim(txtdireccion.Text)
'  End If
'End If
End Sub

Private Sub txtDirTrabajo_GotFocus()
Azul txtDirTrabajo, txtDirTrabajo

End Sub

Private Sub txtDirTrabajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNumDirTrabajo.SetFocus
    Exit Sub
   SIGUE_CAMPO frmCLI.txtDirTrabajo.TabIndex
End If
End Sub

Private Sub txtDTX_GotFocus()
Azul txtDTX, txtDTX
End Sub

Private Sub txtDTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtDTX.TabIndex
End If

End Sub

Private Sub TxtEmpresa_GotFocus()
 Azul TxtEmpresa, TxtEmpresa
End Sub

Private Sub TxtEmpresa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtprendas.SetFocus
End If
End Sub

Private Sub Txtesposa_GotFocus()
 Azul Txtesposa, Txtesposa
End Sub

Private Sub Txtesposa_KeyPress(KeyAscii As Integer)
  
If KeyAscii = 13 Then
    If frmCLI.txtRUCempresa.Visible Then
       frmCLI.txtRUCempresa.SetFocus
    Else
        frmCLI.txttelefono2.SetFocus
    End If
End If
End Sub

Private Sub txtesposo_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
  KeyAscii = 0
  Exit Sub
End If
If KeyAscii = 13 Then
  frmCLI.txtRUCesposo.SetFocus
End If
End Sub

Public Sub GRABAR_CLI(wCGH As String)
Dim PS_REP09 As rdoQuery
Dim llave_rep09 As rdoResultset
  
Dim VAR_CIAS As String * 2
Dim TOTCIAS As String
Dim xcuenta As Integer
Dim Modo As String * 1
PS_GEN(0) = 0
GEN.Requery
TOTCIAS = Nulo_Valors(GEN!gen_cli_cias)
xcuenta = 1
For fila = 1 To 30
    If Trim(Mid(TOTCIAS, xcuenta, 2)) = LK_CODCIA Then
         GoTo SIGUE_PASA
    End If
    xcuenta = xcuenta + 2
Next fila
GoTo CIA_ACTUAL

SIGUE_PASA:
If Trim(TOTCIAS) <> "" And Left(CmbCGP.Text, 1) = "C" Then
  xcuenta = 1
  For fila = 1 To 30
    If Trim(Mid(TOTCIAS, xcuenta, 2)) = "" Then
      Exit For
    Else
       PSPAR_CLI(0) = Mid(TOTCIAS, xcuenta, 2)
       par_llave_cli.Requery
       If par_llave_cli.EOF Then
       '     MsgBox "No Grabo en la Compañia : " + Mid(TOTCIAS, xcuenta, 2) + " No Existe", 48, Pub_Titulo
       Else
           VAR_CIAS = Mid(TOTCIAS, xcuenta, 2)
           If Left(CmdModificar.Caption, 2) = "&G" Then
             If VAR_CIAS = LK_CODCIA Then GoTo pasa
             SQ_OPER = 1
             pu_cp = wCGH
             pu_codclie = Val(frmCLI.Txt_key.Text)
             pu_codcia = VAR_CIAS
             LEER_CLILOC_LLAVE
             If cliloc_llave.EOF Then
'                MsgBox "No Grabo en la Compañia : " + VAR_CIAS + " No Existe cliente ", 48, Pub_Titulo
             Else
               cliloc_llave.Edit
               Modo = "E"
               GoSub grabar
               cliloc_llave.Update
             End If
pasa:
           Else
             cliloc_llave.AddNew
             Modo = "A"
             GoSub grabar
             cliloc_llave.Update
           End If
      End If
    End If
    xcuenta = xcuenta + 2
  Next fila
  'ACTUALIZA POR ULTIMO LA CIA ACTUAL PARA MANTENER LA LLAVE ACTIVA
  If Left(CmdModificar.Caption, 2) = "&G" Then
    VAR_CIAS = LK_CODCIA
    SQ_OPER = 1
    pu_cp = wCGH
    pu_codclie = Val(frmCLI.Txt_key.Text)
    pu_codcia = VAR_CIAS
    LEER_CLILOC_LLAVE
    If cliloc_llave.EOF Then
      MsgBox "No Grabo en la Compañia : " + VAR_CIAS + " No Existe cliente ", 48, Pub_Titulo
    Else
      cliloc_llave.Edit
      cliloc_llave!CLI_LIMCRE = Val(frmCLI.txtlimite.Text)
      Modo = "E"
      GoSub grabar
      cliloc_llave.Update
    End If
  End If
  
Else
CIA_ACTUAL:
  VAR_CIAS = LK_CODCIA
  If Left(CmdModificar.Caption, 2) = "&G" Then
    cliloc_llave.Edit
    Modo = "E"
  Else
    GRABA_CONTAB VAR_CIAS
    cliloc_llave.AddNew
    Modo = "A"
  End If
  GoSub grabar
  cliloc_llave.Update
  
  Exit Sub
End If
 
Exit Sub
   
grabar:
    If Modo = "A" Then
       cliloc_llave!CLI_CP = wCGH
       cliloc_llave!cli_codclie = Val(frmCLI.Txt_key.Text)
       cliloc_llave!cli_SALDO = 0
       cliloc_llave!CLI_DET_TOT = "D"
       cliloc_llave!CLI_MONEDA = "S"
       cliloc_llave!CLI_limcre2 = 0
       cliloc_llave!CLI_CUENTA_CONTAB2 = "1"
       If Left(CmbCGP.Text, 1) = "C" Then
        loc_ultcod = Val(frmCLI.Txt_key.Text)
       End If
    Else
'      If Trim(tempo_ruc) <> Trim(txtRUCesposo.Text) Then
'            pub_cadena = "SELECT ALL_RUC FROM ALLOG WHERE ALL_CODCIA = ?   AND ALL_CODCLIE = ? AND ALL_CP = ? AND ALL_FLAG_EXT <> 'E'"
'            Set PS_REP09 = CN.CreateQuery("", pub_cadena)
'            PS_REP09(0) = 0
'            PS_REP09(1) = 0
'            PS_REP09(2) = 0
'            Set llave_rep09 = PS_REP09.OpenResultset(rdOpenKeyset, rdConcurValues)
'            PS_REP09(0) = LK_CODCIA
'            PS_REP09(1) = Val(frmCLI.txt_key.Text)
'            PS_REP09(2) = Left(CmbCGP.Text, 1)
'            llave_rep09.Requery
'            PB.Visible = True
'            pbValue = 0
'            pbMin = 0
'            If llave_rep09.RowCount <> 0 Then pbMax = llave_rep09.RowCount
'
'            Do Until llave_rep09.EOF
'              pbValue = pbValue + 1
'                llave_rep09.Edit
'                llave_rep09!ALL_RUC = Trim(txtRUCesposo.Text)
'                llave_rep09.Update
'              llave_rep09.MoveNext
'            Loop
'            PB.Visible = False
'       End If
       
    End If
    cliloc_llave!CLI_CODCIA = VAR_CIAS
    cliloc_llave!CLI_NOMBRE_ESPOSO = txtesposo.Text
    cliloc_llave!CLI_NOMBRE_ESPOSA = Txtesposa.Text
    cliloc_llave!CLI_NOMBRE_EMPRESA = TxtEmpresa.Text
    ASIGNA_123
    cliloc_llave!cli_nombre = frmCLI.txtnombre.Text
    cliloc_llave!CLI_CASA_DIREC = txtdireccion.Text
    cliloc_llave!CLI_CASA_NUM = Val(Txtnumdir.Text)
    cliloc_llave!CLI_CASA_ZONA = Val(Right(TxtZona.Text, 4))
    cliloc_llave!CLI_LUGAR_CASA = Val(Right(TxtLugarCasa.Text, 4))
    cliloc_llave!CLI_LUGAR_TRAB = Val(Right(TxtLugarTrab.Text, 4))
    
    cliloc_llave!cli_dia_visita = Val(Right(cboDiaVisita.Text, 4))
    
    cliloc_llave!CLI_CASA_SUBZONA = Val(Right(TxtSubZona.Text, 4))
    cliloc_llave!cli_zona_new = Val(Right(txtZonaNew.Text, 4))
    cliloc_llave!CLI_TRAB_DIREC = txtDirTrabajo.Text
    cliloc_llave!CLI_TRAB_NUM = Nulo_Valor0(txtNumDirTrabajo.Text)
    cliloc_llave!cli_TRAB_ZONA = Val(Right(frmCLI.TxtZonaTrabajo.Text, 4))
    'AGREGADO PARA GRABAR LA PROVINCIA DE TRABAJO 06/12/2001
    cliloc_llave!cli_TRAB_PROV = Val(Right(frmCLI.cboProvincia.Text, 4))
    cliloc_llave!CLI_CIA_REF = Trim(frmCLI.t_ciarela.Text)
    
    'AGREGADO PARA GRABAR LA PRIMERA DIRECCION
    If cboDireccion.Text = "" Then 'And Trim(Left$(CmbCGP.Text, 1)) = "C" Then
     Dim strDir As String
     Dim strRef As String
     Dim SQL As String
     Dim dir As String
      strDir = Trim(txtDirTrabajo)
      strRef = Trim(txtpropiedad2)
      dir = Trim(Left$(TxtLugarTrab, 10))
      dir = dir + " " + Trim(Left$(strDir, 30))
      dir = dir + " " + Trim(Left$(txtNumDirTrabajo, 30))
      dir = dir + " Zn. " + Trim(Left$(TxtSubZonaTrabajo.Text, 30))
      dir = dir + ", Dt. " + Trim(Left$(TxtZonaTrabajo, 30))
      dir = dir + ", Pr. " + Trim(Left$(cboProvincia, 30))
      If strDir = "" Then
'       MsgBox "Dato ingresado no valido, Intentelo nuevamente", vbInformation, "Dirección"
       'GoTo SALTAdire
       strDir = "-"
      End If
       SQL = "insert into dircli " & _
       "(codcia,codcli,cp,direc,ref,CLI_LUGAR_TRAB, " & _
       "CLI_TRAB_ZONA,CLI_CASA_SUBZONA,CLI_TRAB_SUBZONA,NUMERO,DIRCOMP) " & _
       "values('" & LK_CODCIA & "','" & Val(frmCLI.Txt_key) & "','" & Trim(Left$(CmbCGP.Text, 1)) & "','" & strDir & "','" & strRef & "','" & _
       Val(Right(frmCLI.TxtLugarTrab, 4)) & "','" & _
       Val(Right(frmCLI.TxtZonaTrabajo.Text, 4)) & "','" & _
       Val(Right(frmCLI.cboProvincia, 4)) & "','" & _
       Val(Right(frmCLI.TxtSubZonaTrabajo, 4)) & "','" & _
       Val(txtNumDirTrabajo) & "','" & dir & "')"
       CN.Execute SQL
SALTAdire:
    End If
    cliloc_llave!cli_TRAB_SUBZONA = Val(Right(TxtSubZonaTrabajo.Text, 4))
    cliloc_llave!cli_ruc_esposo = txtRUCesposo.Text
    cliloc_llave!cli_RUC_ESPOSA = txtRUCesposa.Text
    cliloc_llave!CLI_RUC_EMPRESA = txtRUCempresa.Text
    cliloc_llave!CLI_CASA1 = frmCLI.txtpropiedad1.Text
    cliloc_llave!CLI_CASA2 = frmCLI.txtpropiedad2.Text
    cliloc_llave!CLI_regpub1 = frmCLI.txtregpublico1.Text
    cliloc_llave!CLI_REGPUB2 = frmCLI.txtregpublico2.Text
    cliloc_llave!CLI_AUTOAVALUO = frmCLI.txtautovaluo.Text
    cliloc_llave!CLI_AUTO1 = frmCLI.txtauto1.Text
    cliloc_llave!cli_auto2 = frmCLI.txtauto2.Text
    cliloc_llave!CLI_PRENDA = frmCLI.txtprendas.Text
    cliloc_llave!cli_telef1 = frmCLI.txttelefono1.Text
    cliloc_llave!CLI_TELEF2 = frmCLI.txttelefono2.Text
    cliloc_llave!CLI_OTRO_CONTR = frmCLI.otrocontrato.Value
    cliloc_llave!CLI_LETRA = frmCLI.letraotorgado.Value
    cliloc_llave!CLI_GRUPO = Val(Right(frmCLI.cmbgrupo.Text, 4))
    cliloc_llave!CLI_SUBGRUPO = Val(Right(frmCLI.txtsubgrupo.Text, 4))
    cliloc_llave!CLI_TIPOCLI = Val(Right(frmCLI.cmbtipocli.Text, 4)) ' MILTON
    cliloc_llave!CLI_division = Val(Right(frmCLI.Cmbcate.Text, 4))
    cliloc_llave!CLI_nucleo = frmCLI.txtNucleo.Text
    cliloc_llave!cli_estado = Left(frmCLI.txtestado.Text, 1)
    cliloc_llave!CLI_programado = Nulo_Valors(txtprog.Text)
    cliloc_llave!CLI_PORDESCTO = Nulo_Valor0(txtpordes.Text)
    cliloc_llave!cli_fechahora = Format(Now, "dd/mm/yy") & Chr(13) & Format(Now, "hh:mm AMPM") & Chr(13) & LK_CODUSU
    If LK_FLAG_GRIFO = "A" Then
     cliloc_llave!cli_fecha_fac = Format(t_fechafac.Text, "dd/mm/yyyy")
    Else
     cliloc_llave!cli_fecha_fac = Format(LK_FECHA_DIA, "dd/mm/yyyy")
    End If
     cliloc_llave!cli_DIAS_FAC = Val(t_diasfac.Text)
     cliloc_llave!cli_dias_cred = Val(frmCLI.t_diascred.Text)
   '  <<< Actualiza La Cta. solo de la Cia Actual >>>
   '' If VAR_CIAS = LK_CODCIA Then
      cliloc_llave!CLI_CUENTA_CONTAB = Trim(frmCLI.tcuenta.Text)
      cliloc_llave!CLI_CIARELA = Trim(frmCLI.txtciarela.Text)
    
      cliloc_llave!cli_cuenta_contab22 = Trim(frmCLI.tcuenta22.Text)
   '' End If
    If txtDTX.Text = "" Then
     txtDTX.Text = " "
    End If
    cliloc_llave!CLI_DET_TOT = txtDTX.Text
    If Trim(TOTCIAS) = "" Then
      cliloc_llave!CLI_LIMCRE = Val(txtlimite.Text)
    End If
    pub_cadena = ""
    For fila = 0 To lisdescto.ListCount - 1
      If lisdescto.Selected(fila) Then
           pub_cadena = pub_cadena + Format(Trim(Right(lisdescto.List(fila), 6)), "00")
      End If
    Next fila
    cliloc_llave!CLI_CASA1 = Trim(pub_cadena)
     
   
   
Return
End Sub

Public Sub MENSAJE_CLI(TEXTO As String)
  LblMensaje.Caption = TEXTO
  PARPADEA.Enabled = True
End Sub
Public Function GENERA_PRO() As Double
Dim NUMCAD, FIJO As String
Dim DIGI As String * 2
Dim I, VINT1, VINT2, VINT3, VINT4 As Double
Dim VSTR1, VSTR2, VSTR3, VSTR4 As String
Dim VFIJO As Double
Dim VVARI As Integer
Dim STRpub_cadena As String
Dim INTpub_cadena As Double
pu_cp = "P"
pu_codclie = 0
SQ_OPER = 2
pu_codcia = LK_CODCIA
LEER_CLILOC_LLAVE
If cliloc_mayor.EOF Then
    NUMCAD = "1"
Else
    cliloc_mayor.MoveLast
    NUMCAD = cliloc_mayor!cli_codclie
End If
VINT2 = 0
NUMCAD = Trim(NUMCAD)
VINT1 = Len(NUMCAD)
If NUMCAD = "1" Or NUMCAD = "2" Or NUMCAD = "0" Then
  VINT2 = 1
  VINT1 = 2
End If
If VINT1 > 1 Then
    VSTR4 = Val(Mid(NUMCAD, 1, VINT1 - 2)) + 1
End If

For I = 1 To VINT1 - 2
   VSTR1 = Mid(VSTR4, I, 1)
   VINT2 = VINT2 + Val(VSTR1)
Next I
VINT3 = VINT2 * 13 - 5

VSTR3 = Right(CStr(VINT3), 2)
If Len(VSTR3) = 1 Then
  VSTR3 = "0" & VSTR3
End If
FIJO = VSTR4
STRpub_cadena = FIJO & VSTR3
INTpub_cadena = Val(STRpub_cadena)
GENERA_PRO = INTpub_cadena
End Function

Public Function GENERA_CODI() As Double
Dim NUMCAD, FIJO As String
Dim DIGI As String * 2
Dim I, VINT1, VINT2, VINT3, VINT4 As Double
Dim VSTR1, VSTR2, VSTR3, VSTR4 As String
Dim VFIJO As Double
Dim VVARI As Integer
Dim STRpub_cadena As String
Dim INTpub_cadena As Double
pu_cp = "C"
pu_codclie = 0
SQ_OPER = 2
PUB_CODCIA = LK_CODCIA
pu_codcia = PUB_CODCIA
LEER_CLILOC_LLAVE

If cliloc_mayor.EOF Then
    NUMCAD = "1"
    If LK_EMP = "HER" Then
      INTpub_cadena = Val(NUMCAD)
      If COD_ORIGINAL <> 0 And INTpub_cadena <> Val(Txt_key.Text) Then
        INTpub_cadena = Val(Txt_key.Text)
        GoTo GEN
      End If
      COD_ORIGINAL = INTpub_cadena
      GoTo GEN
    End If
Else
    cliloc_mayor.MoveLast
    NUMCAD = cliloc_mayor!cli_codclie
    If LK_EMP = "HER" Then
      INTpub_cadena = Val(NUMCAD) + 1
      If COD_ORIGINAL <> 0 And INTpub_cadena <> Val(Txt_key.Text) Then
      '  INTpub_cadena = Val(Txt_key.Text)
      '  GoTo GEN
      End If
      COD_ORIGINAL = INTpub_cadena
      GoTo GEN
    End If
End If

VINT2 = 0
NUMCAD = Trim(NUMCAD)
VINT1 = Len(NUMCAD)
If NUMCAD = "1" Or NUMCAD = "2" Or NUMCAD = "0" Then
  VINT2 = 1
  VINT1 = 2
End If
If VINT1 > 1 Then
    VSTR4 = Val(Mid(NUMCAD, 1, VINT1 - 2)) + 1
End If
For I = 1 To VINT1 - 2
   VSTR1 = Mid(VSTR4, I, 1)
   VINT2 = VINT2 + Val(VSTR1)
Next I
VINT3 = VINT2 * 9

VSTR3 = Right(CStr(VINT3), 2)
If Len(VSTR3) = 1 Then
  VSTR3 = "0" & VSTR3
End If
FIJO = VSTR4
STRpub_cadena = FIJO & VSTR3
INTpub_cadena = Val(STRpub_cadena)
GEN:
GENERA_CODI = INTpub_cadena

End Function
Public Function CONSIS_CLI() As Boolean
Dim wruc As Integer
If Left(CmdModificar.Caption, 2) = "&G" Then
   If Left(txtestado.Text, 1) = "D" Then
    pub_cadena = "select  sum(car_importe) as valor from cartera where car_codcia = '" & LK_CODCIA & "' and car_cp = '" & Left(CmbCGP.Text, 1) & "' and car_codclie = " & Val(frmCLI.Txt_key.Text) & " and car_importe <> 0"
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    If Not X.EOF Then
        If Nulo_Valor0(X!VALOR) <> 0 Then
            MsgBox "Al DesActivar - NO Procede, el codigo tiene Saldo en Cuentas.", 48, Pub_Titulo
            CONSIS_CLI = False
            GoTo ESCAPA
        End If
    End If
    End If
 End If
If frmCLI.OptNombre(0).Value Then
    If Trim(frmCLI.txtesposo.Text) = "" Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Datos Principal ..."
        txtesposo.SetFocus
        GoTo ESCAPA
    End If
ElseIf frmCLI.OptNombre(1).Value Then
    If Trim(frmCLI.Txtesposa.Text) = "" Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Datos Principal ..."
        Txtesposa.SetFocus
        GoTo ESCAPA
    End If
ElseIf frmCLI.OptNombre(2).Value Then
    If Trim(frmCLI.TxtEmpresa.Text) = "" Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Datos Principal ..."
        TxtEmpresa.SetFocus
        GoTo ESCAPA
    End If
End If

If Len(frmCLI.txtesposo.Text) = 0 And Len(frmCLI.Txtesposa.Text) = 0 And Len(frmCLI.TxtEmpresa.Text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Algun Nombre  ..."
        txtesposo.SetFocus
        GoTo ESCAPA
ElseIf frmCLI.OptNombre(0).Value And Len(frmCLI.txtesposo.Text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Nombre  NO Puede estar en Blanco ..."
        txtesposo.SetFocus
        GoTo ESCAPA
ElseIf frmCLI.OptNombre(1).Value And Len(frmCLI.Txtesposa.Text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Nombre  NO Puede estar en Blanco ..."
        Txtesposa.SetFocus
        GoTo ESCAPA
ElseIf frmCLI.OptNombre(2).Value And Len(frmCLI.TxtEmpresa.Text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Nombre  NO Puede estar en Blanco ..."
        TxtEmpresa.SetFocus
        GoTo ESCAPA
ElseIf Len(frmCLI.txtesposo.Text) = 0 And Len(frmCLI.txtRUCesposo.Text) > 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "RUC  debe estar en Blanco ..."
        txtRUCesposo.SetFocus
        GoTo ESCAPA
'ElseIf Len(frmCLI.Txtesposa.Text) = 0 And Len(frmCLI.txtRUCesposa.Text) > 0 Then
'        CONSIS_CLI = False
'        MENSAJE_CLI "L.E.  debe estar en Blanco ..."
'        txtRUCesposa.SetFocus
'        GoTo ESCAPA
ElseIf Len(frmCLI.TxtEmpresa.Text) = 0 And Len(frmCLI.txtRUCempresa.Text) > 0 Then
     If LK_EMP <> "PLA" Then
        CONSIS_CLI = False
        MENSAJE_CLI "RUC  debe estar en Blanco ..."
        txtRUCempresa.SetFocus
        GoTo ESCAPA
     End If
End If
wruc = 8
If LK_DIG_RUC <> 0 Then wruc = LK_DIG_RUC
If frmCLI.txtRUCesposo.Text <> "" Then
  If Len(Trim(frmCLI.txtRUCesposo.Text)) <> wruc Then
       CONSIS_CLI = False
       MENSAJE_CLI "R.U.C. de No es Validad ..."
       frmCLI.txtRUCesposo.SetFocus
       GoTo ESCAPA
   End If
Else
  If Left(CmbCGP.Text, 1) = "P" Then
    MsgBox "Necesita Nro de R.U.C... dato no opcional", 48, Pub_Titulo
    CONSIS_CLI = False
    frmCLI.txtRUCesposo.SetFocus
    GoTo ESCAPA
  End If
End If

If frmCLI.txtRUCesposa.Text <> "" Then
    If Len(Trim(frmCLI.txtRUCesposa.Text)) = 8 Or Len(Trim(frmCLI.txtRUCesposa.Text)) = 12 Then
    Else
       CONSIS_CLI = False
       MENSAJE_CLI "D.N.I. de No es Validad ..."
       frmCLI.txtRUCesposa.SetFocus
       GoTo ESCAPA
    End If
End If
If LK_EMP <> "PLA" Then
 If Left(CmbCGP.Text, 1) = "C" Then
  If frmCLI.txtRUCempresa.Text <> "" Then
    If Len(Trim(frmCLI.txtRUCempresa.Text)) <> 8 Then
       CONSIS_CLI = False
       MENSAJE_CLI "D.N.I. de No es Validad ..."
       txtRUCempresa.SetFocus
       GoTo ESCAPA
    End If
  End If
 End If
End If
If frmCLI.txtRUCesposo.Text <> "" Then
'CLI_CODCIA = ? AND CLI_CP = 'C' AND CLI_RUC_ESPOSO = ? and CLI_CODCLIE <> ?
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = Left(frmCLI.CmbCGP, 1)
 PS_REP01(2) = frmCLI.txtRUCesposo.Text
 PS_REP01(3) = frmCLI.Txt_key.Text
 llave_rep01.Requery
 If Not llave_rep01.EOF Then
    MsgBox "RUC Existe en otro Cliente : " + Trim(llave_rep01!cli_nombre), 48, Pub_Titulo
    LK_ACCESO_REPORT = ""
    Load frmclave2
    Screen.MousePointer = 0
    frmclave2.Show 1
    If LK_ACCESO_REPORT <> "A" Then
      CONSIS_CLI = False
      Azul frmCLI.txtRUCesposo, frmCLI.txtRUCesposo
      GoTo ESCAPA
    End If
 End If
End If
If Trim(par_llave!PAR_CONTABILIDAD) = "A" Then
  SQ_OPER = 1
  PUB_CUENTA = Trim(tcuenta.Text)
  LEER_COM_LLAVE
  If com_llave.EOF Then
    MsgBox "Cuanta Contable No Existe. Verificar ", 48, Pub_Titulo
    CONSIS_CLI = False
    Azul frmCLI.tcuenta, frmCLI.tcuenta
    GoTo ESCAPA
  End If
End If
If frmCLI.txtNucleo = "" Then
   frmCLI.txtNucleo.Text = " "
End If

If Trim(frmCLI.t_ciarela.Text) <> "" Then
  pub_cadena = "SELECT * FROM PARGEN WHERE PAR_CODCIA = '" & Trim(frmCLI.t_ciarela.Text) & "'"
  Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
  If X.EOF Then
    MsgBox "Codigo de Cia no existe", 48, Pub_Titulo
    CONSIS_CLI = False
    Azul t_ciarela, t_ciarela
    GoTo ESCAPA
   End If
End If

If Left(CmbCGP.Text, 1) = "C" Then
 If Val(Right(cboDiaVisita.Text, 6)) = 0 Then
    MsgBox "Asignar a un Vemdedor.!!!", 48, Pub_Titulo
    CONSIS_CLI = False
    cboDiaVisita.SetFocus
    GoTo ESCAPA
 End If
 If Trim(TxtLugarCasa.Text) = "" Then
    MsgBox "Dato no es opcional ,Lugar.", 48, Pub_Titulo
    CONSIS_CLI = False
    TxtLugarCasa.SetFocus
    GoTo ESCAPA
 End If
 If Trim(TxtZona.Text) = "" Then
    MsgBox "Dato no es opcional ,Definir.", 48, Pub_Titulo
    CONSIS_CLI = False
    TxtZona.SetFocus
    GoTo ESCAPA
 End If
 If Trim(TxtSubZona.Text) = "" Then
    MsgBox "Dato no es opcional ,Definir.", 48, Pub_Titulo
    CONSIS_CLI = False
    TxtSubZona.SetFocus
    GoTo ESCAPA
 End If
 If Trim(txtZonaNew.Text) = "" Then
    MsgBox "Dato no es opcional ,Definir.", 48, Pub_Titulo
    CONSIS_CLI = False
    txtZonaNew.SetFocus
    GoTo ESCAPA
 End If
End If
If LK_FLAG_GRIFO = "A" And Left(CmbCGP.Text, 1) = "C" Then
 If Not IsDate(frmCLI.t_fechafac.Text) Then
    MsgBox "Fecha para la Facturacion no procede.", 48, Pub_Titulo
    CONSIS_CLI = False
    Azul frmCLI.t_fechafac, frmCLI.t_fechafac
    GoTo ESCAPA
 End If
 If Trim(frmCLI.t_fechafac.Text) = "" Then
    frmCLI.t_fechafac.Text = LK_FECHA_DIA
 End If
Else
   frmCLI.t_fechafac.Text = LK_FECHA_DIA
End If
 


CONSIS_CLI = True
ESCAPA:
End Function

Private Sub txtestado_GotFocus()
'Azul txtestado, txtestado
End Sub

Private Sub txtestado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   'SIGUE_CAMPO frmCLI.txtestado.TabIndex
   
End If
End Sub

Private Sub txtlimite_GotFocus()
Azul txtlimite, txtlimite
End Sub

Private Sub txtlimite_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Left(cmdAgregar.Caption, 2) = "&G" Then
   cmdAgregar.SetFocus
 Else
   CmdModificar.SetFocus
 End If
End If

End Sub

Private Sub TxtLugarCasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frmCLI.txtdireccion.SetFocus
End If

End Sub

Private Sub TxtLugarCasa_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtLugarCasa.ListIndex
PUB_TIPREG = Mid(TxtLugarCasa.ToolTipText, 13, Len(TxtLugarCasa.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtLugarCasa, 25
LLENA_ZONA TxtLugarTrab, 25
TxtLugarCasa.SetFocus
SendKeys "%{up}"

End Sub

Private Sub TxtLugarCasa_LostFocus()
On Error GoTo sigue
'If Left(CmbCGP.Text, 1) = "C" Then
'  If Trim(TxtLugarTrab.Text) = "" Then
    TxtLugarTrab.ListIndex = TxtLugarCasa.ListIndex
'  End If
'End If

Exit Sub
sigue:
End Sub

Private Sub TxtLugarTrab_GotFocus()
SendKeys "%{Down}"
End Sub

Private Sub TxtLugarTrab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtDirTrabajo.SetFocus
Exit Sub
   SIGUE_CAMPO frmCLI.TxtLugarTrab.TabIndex

End If

End Sub


Private Sub TxtLugarTrab_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtLugarTrab.ListIndex
PUB_TIPREG = Mid(TxtLugarTrab.ToolTipText, 13, Len(TxtLugarTrab.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtLugarCasa, 25
LLENA_ZONA TxtLugarTrab, 25
TxtLugarTrab.SetFocus
SendKeys "%{up}"

End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtesposo.SetFocus

End If
End Sub

Private Sub txtNucleo_GotFocus()
Azul txtNucleo, txtNucleo
End Sub

Private Sub txtnucleo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtNucleo.TabIndex
End If

End Sub

Private Sub Txtnumdir_GotFocus()
Azul Txtnumdir, Txtnumdir

End Sub

Private Sub Txtnumdir_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
   txtZonaNew.SetFocus
   SendKeys "%{up}"
End If
End Sub

Private Sub Txtnumdir_LostFocus()
On Error GoTo sigue
'If Left(CmbCGP.Text, 1) = "C" Then
'  If Val(txtNumDirTrabajo.Text) = 0 Then
    txtNumDirTrabajo.Text = Txtnumdir.Text
'  End If
'End If
Exit Sub
sigue:
End Sub

Private Sub txtnumdirtrabajo_GotFocus()
Azul txtNumDirTrabajo, txtNumDirTrabajo

End Sub

Private Sub txtnumdirtrabajo_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  TxtSubZonaTrabajo.SetFocus
  SendKeys "%{UP}"
'  SIGUE_CAMPO frmCLI.txtnumdirtrabajo.TabIndex
End If
End Sub

Private Sub txtpordes_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL txtpordes, KeyAscii
If KeyAscii = 13 Then
  frmCLI.TxtZona.SetFocus
  SendKeys "%{up}"
End If

End Sub

Private Sub txtprendas_GotFocus()
Azul txtprendas, txtprendas
End Sub

Private Sub txtprendas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   frmCLI.cmbgrupo.SetFocus
    SendKeys "%{up}"
End If
End Sub

Private Sub txtprog_GotFocus()
 Azul txtprog, txtprog
End Sub

Private Sub txtprog_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtprog.TabIndex

End If

End Sub

Private Sub txtpropiedad1_GotFocus()
Azul txtpropiedad1, txtpropiedad1
End Sub

Private Sub txtpropiedad1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SIGUE_CAMPO frmCLI.txtpropiedad1.TabIndex
End If
End Sub

Private Sub txtpropiedad2_GotFocus()
Azul txtpropiedad2, txtpropiedad2
End Sub

Private Sub txtpropiedad2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtpropiedad2.TabIndex
End If
End Sub

Private Sub txtregpublico1_GotFocus()
Azul txtregpublico1, txtregpublico1
End Sub

Private Sub txtregpublico1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtregpublico1.TabIndex
End If
End Sub

Private Sub txtregpublico2_GotFocus()
Azul txtregpublico2, txtregpublico2
End Sub

Private Sub txtregpublico2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtregpublico2.TabIndex
End If
End Sub

Private Sub txtRUCempresa_GotFocus()
Azul txtRUCempresa, txtRUCempresa
End Sub

Private Sub txtRUCempresa_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
    frmCLI.TxtEmpresa.SetFocus
End If
End Sub

Private Sub txtRUCesposa_GotFocus()
Azul txtRUCesposa, txtRUCesposa
End Sub

Private Sub txtRUCesposa_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
    frmCLI.txttelefono2.SetFocus
End If
End Sub

Private Sub txtRUCesposo_GotFocus()
 Azul txtRUCesposo, txtRUCesposo
End Sub

Private Sub txtRUCesposo_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 If LK_EMP = "3AA" Then
  Azul tcuenta, tcuenta
  Else
  frmCLI.txttelefono1.SetFocus
 End If
End If
End Sub

Private Sub txtsubgrupo_Click()
  Dim I As Integer
  Exit Sub
      For I = 0 To txtsubgrupo.ListCount - 1
        If Val(Right(lisdescto.List(I), 6)) = Val(Right(txtsubgrupo.Text, 6)) Then
          lisdescto.Selected(I) = True
         Else
          lisdescto.Selected(I) = False
        End If
      Next I
End Sub

Private Sub txtsubgrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 frmCLI.Cmbcate.SetFocus
 SendKeys "%{up}"
End If
End Sub


Private Sub txtsubgrupo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = txtsubgrupo.ListIndex
PUB_TIPREG = Mid(txtsubgrupo.ToolTipText, 13, Len(txtsubgrupo.ToolTipText))
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "SUB - GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
If Left(CmbCGP.Text, 1) = "C" Then
  LLENA_GRUPOS txtsubgrupo, 333
Else
  LLENA_GRUPOS txtsubgrupo, 334
End If

txtsubgrupo.SetFocus
SendKeys "%{up}"
End Sub

Private Sub TxtSubZona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 txtregpublico2.SetFocus
End If
End Sub

Private Sub TxtSubZona_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtSubZona.ListIndex
PUB_TIPREG = Mid(TxtSubZona.ToolTipText, 13, Len(TxtSubZona.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtSubZona, 30
LLENA_ZONA TxtSubZonaTrabajo, 35
TxtSubZona.SetFocus
SendKeys "%{up}"


End Sub

Private Sub TxtSubZona_LostFocus()
On Error GoTo sigue
'If Left(CmbCGP.Text, 1) = "C" Then
  'If Trim(cboProvincia.Text) = "" Then
      cboProvincia.ListIndex = TxtSubZona.ListIndex
  'End If
'End If
Exit Sub
sigue:

End Sub

Private Sub TxtSubZonaTrabajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TxtZonaTrabajo.SetFocus
   SendKeys "%{UP}"
  'SIGUE_CAMPO TxtSubZonaTrabajo.TabIndex
End If
End Sub


Private Sub TxtSubZonaTrabajo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtSubZonaTrabajo.ListIndex
PUB_TIPREG = Mid(TxtSubZonaTrabajo.ToolTipText, 13, Len(TxtSubZonaTrabajo.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtSubZonaTrabajo, 35
LLENA_ZONA TxtSubZona, 30

TxtSubZonaTrabajo.SetFocus
SendKeys "%{up}"

End Sub

Private Sub txttelefono1_GotFocus()
Azul txttelefono1, txttelefono1
End Sub


Private Sub txttelefono1_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  txtRUCesposa.SetFocus
  Azul txtpordes, txtpordes
End If
End Sub

Private Sub txttelefono1_LostFocus()
On Error GoTo sigue
If Left(CmbCGP.Text, 1) = "C" Then
  If Trim(txttelefono2.Text) = "" Then
    txttelefono2.Text = txttelefono1.Text
  End If
End If

Exit Sub
sigue:
End Sub

Private Sub txttelefono2_GotFocus()
Azul txttelefono2, txttelefono2
End Sub

Private Sub txttelefono2_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  'SIGUE_CAMPO frmCLI.txttelefono2.TabIndex
  Txtesposa.SetFocus
End If
End Sub

Private Sub TxtZona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frmCLI.TxtSubZona.SetFocus
    SendKeys "%{up}"
End If
End Sub


Private Sub TxtZona_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtZona.ListIndex
PUB_TIPREG = Mid(TxtZona.ToolTipText, 13, Len(TxtZona.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtZona, 20
LLENA_ZONA TxtZonaTrabajo, 20
TxtZona.SetFocus
SendKeys "%{up}"

End Sub

Private Sub TxtZona_LostFocus()
On Error GoTo sigue
'If Left(CmbCGP.Text, 1) = "C" Then
  'If Trim(TxtZonaTrabajo.Text) = "" Then
      TxtZonaTrabajo.ListIndex = TxtZona.ListIndex
  'End If
'End If

Exit Sub
sigue:

End Sub

Private Sub txtZonaNew_LostFocus()
On Error GoTo sigue
'If Left(CmbCGP.Text, 1) = "C" Then
  'If Trim(TxtSubZonaTrabajo.Text) = "" Then
      TxtSubZonaTrabajo.ListIndex = txtZonaNew.ListIndex
 ' End If
'End If
Exit Sub
sigue:
End Sub

Private Sub TxtZonaTrabajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cboProvincia.SetFocus
  SendKeys "%{UP}"

'   SIGUE_CAMPO frmCLI.TxtZonaTrabajo.TabIndex
End If
End Sub

Public Function EXISTE_CLI(wcp As String, VALOR1 As String, WCODI As String) As Boolean
Dim var
Dim tempo
tempo = Left(Trim(VALOR1), Len(VALOR1) - 1)
var = Asc(Right(Trim(VALOR1), 1))
var = var + 1
If var = 91 Then
  var = "ZZZZZZZZ"
Else
  var = Chr(var)
End If
tempo = tempo + var
archi = "SELECT * FROM CLIENTES WHERE  CLI_CODCLIE <> " & WCODI & " AND CLI_CP = '" & wcp & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & VALOR1 & "' AND  '" & tempo & "' ORDER BY CLI_NOMBRE"
ListExiste.Clear
Set PSX = CN.CreateQuery("", archi)
Set X = PSX.OpenResultset(rdOpenKeyset)
X.Requery
EXISTE_CLI = False
If X.EOF Then
 frmCLI.ListExiste.Clear
 GoTo fin
End If

If wcp = "P" Then
    F14.Caption = "Lista de Proveedores parecidos ... "
End If
frmCLI.ListExiste.TextMatrix(0, 0) = "Cia"
frmCLI.ListExiste.TextMatrix(0, 1) = "Codigo "
If wcp = "C" Then
    F14.Caption = "Lista de Clientes parecidos ... "
End If
frmCLI.ListExiste.TextMatrix(0, 2) = lblnom(0).Caption
frmCLI.ListExiste.TextMatrix(0, 3) = lblnom(2).Caption
frmCLI.ListExiste.TextMatrix(0, 4) = lblnom(6).Caption & " " & lblnom(7).Caption

fila = 0
frmCLI.ListExiste.Rows = 2
Do Until X.EOF
    fila = fila + 1
    frmCLI.ListExiste.TextMatrix(fila, 0) = Nulo_Valors(X!CLI_CODCIA)
    frmCLI.ListExiste.TextMatrix(fila, 1) = Nulo_Valors(X!cli_codclie)
    frmCLI.ListExiste.TextMatrix(fila, 2) = Nulo_Valors(X!CLI_NOMBRE_ESPOSO)
    frmCLI.ListExiste.TextMatrix(fila, 3) = Nulo_Valors(X!CLI_NOMBRE_ESPOSA)
    frmCLI.ListExiste.TextMatrix(fila, 4) = Nulo_Valors(X!CLI_CASA_DIREC) & "  # " & Nulo_Valors(X!CLI_CASA_NUM)
    EXISTE_CLI = True
    frmCLI.ListExiste.Rows = frmCLI.ListExiste.Rows + 1
    X.MoveNext
Loop

If EXISTE_CLI Then
    frmCLI.ListExiste.Rows = frmCLI.ListExiste.Rows - 1
    Op(0).Value = False
    Op(0).Enabled = False
    Op(1).Value = True
    frmCLI.F14.Visible = True
    frmCLI.ListExiste.Row = 1
    frmCLI.ListExiste.COL = 1
    frmCLI.ListExiste.SetFocus
End If
GoTo fin
Exit Function

CHECKERROR:
MsgBox Err.Description
fin:

End Function


Private Sub TxtZonaTrabajo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtZonaTrabajo.ListIndex
PUB_TIPREG = Mid(TxtZonaTrabajo.ToolTipText, 13, Len(TxtZonaTrabajo.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtZonaTrabajo, 20
LLENA_ZONA TxtZona, 20
TxtZonaTrabajo.SetFocus
SendKeys "%{up}"
End Sub
Private Sub TxtZonanew_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frmCLI.TxtZona.SetFocus
    SendKeys "%{up}"
  
End If
End Sub


Private Sub TxtZonanew_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
PUB_TIPREG = Mid(txtZonaNew.ToolTipText, 13, Len(txtZonaNew.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA txtZonaNew, 35
txtZonaNew.SetFocus
SendKeys "%{up}"

End Sub


Public Sub ETIQUETA_CLI()
SQ_OPER = 1
PUB_TIPREG = LOC_TIPREG
PUB_CODCIA = LK_CODCIA
For fila = 0 To lblnom.count - 1
 PUB_NUMTAB = Val(lblnom(fila).Tag)
 LEER_TAB_LLAVE
 If tab_llave.EOF Then
 Else
 If fila = 30 Then
' Stop
 End If
  lblnom(fila).Caption = Trim(tab_llave!TAB_NOMLARGO)
 End If
Next fila
End Sub

Public Sub LEER_CLILOC_LLAVE()
Select Case SQ_OPER
Case 1
PSCLILOC_LLAVE.rdoParameters(0) = pu_cp
PSCLILOC_LLAVE.rdoParameters(1) = pu_codclie
PSCLILOC_LLAVE.rdoParameters(2) = pu_codcia
cliloc_llave.Requery
GoTo salida

Case 2
PSCLILOC_MAYOR.rdoParameters(0) = pu_cp
PSCLILOC_MAYOR.rdoParameters(1) = pu_codclie
PSCLILOC_MAYOR.rdoParameters(2) = pu_codcia
cliloc_mayor.Requery
GoTo salida

Case 3
GoTo salida

End Select


salida:

End Sub

Public Sub ACCESO_CLI()
On Error GoTo Ver
Dim tAcceso As String
Dim W1 As String * 2
Dim I, wPosF, WPosV, cuenta As Integer
Dim sal As Boolean
Dim cade As String
Dim WNUM As Integer
Dim F As Integer
Dim a As Integer
tAcceso = Trim(Nulo_Valors(par_llave!par_acceso_cli))

WNUM = 0
wPosF = 0
WPosV = 0
cuenta = 0
WPosV = Len(tAcceso)
cade = Trim(tAcceso)
cuenta = 0
wPosF = 1
a = 0
For I = 1 To Len(cade)
If Mid(cade, I, 1) = "." Then
  a = a + 1
End If
Next I
Do Until cuenta = a
   cuenta = cuenta + 1
   wPosF = InStr(wPosF, cade, ".", 1) + 1
   WNUM = Mid(cade, wPosF, 2)
   If Right(WNUM, 1) = "." Then
     WNUM = Left(WNUM, 2)
     wPosF = wPosF - 1
   End If
   ' WNUM tiene el codigo a mostrar
   GoSub muestracli
Loop

Exit Sub

muestracli:
fila = 0
Do Until fila >= frmCLI.Controls.count
 If TypeOf frmCLI.Controls(fila) Is Timer Then
   GoTo OTRITO
 End If
 If TypeOf frmCLI.Controls(fila) Is MSFlexGrid Then
   GoTo OTRITO
 End If
 If TypeOf frmCLI.Controls(fila) Is Line Then
   GoTo OTRITO
 End If
 If frmCLI.Controls(fila).WhatsThisHelpID = WNUM Then
   frmCLI.Controls(fila).Visible = True
 End If
 
OTRITO:
 fila = fila + 1
Loop
Return
Exit Sub
Ver:
MsgBox Err.Description, 48, Pub_Titulo
Resume Next

End Sub

Public Sub SIGUE_CAMPO(WTAG As Integer)
' 40 ES EL MAXIMO DE CAMPOS DISPONIBLE
On Error GoTo sigue
Dim wmax As Integer
Dim cuenta As Integer
wmax = 42
fila = WTAG
Do Until fila >= wmax
 fila = fila + 1
 cuenta = 0
 Do Until cuenta >= frmCLI.Controls.count - 1
  If TypeOf frmCLI.Controls(cuenta) Is Timer Then
    GoTo OTRITO
  End If
  If TypeOf frmCLI.Controls(cuenta) Is MSFlexGrid Then
    GoTo OTRITO
  End If
'  MsgBox frmCLI.Controls(fila).Name
  If TypeOf frmCLI.Controls(cuenta) Is OptionButton Then
    GoTo OTRITO
  End If
  If frmCLI.Controls(cuenta).TabIndex = fila Then
    If frmCLI.Controls(cuenta).Visible Then
         frmCLI.Controls(cuenta).SetFocus
         Exit Sub
    End If
  End If
OTRITO:
  cuenta = cuenta + 1
 Loop
Loop
If frmCLI.CmdModificar.Enabled Then
   frmCLI.CmdModificar.SetFocus
Else
   frmCLI.cmdAgregar.SetFocus
End If
Exit Sub
sigue:
Resume Next
End Sub

Public Sub GRABA_CONTAB(wcia As String)
Dim flagpase As String * 1
Exit Sub
 If Left(CmbCGP.Text, 1) = "C" Then
   If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" Then
     Exit Sub
   End If
 ElseIf Left(CmbCGP.Text, 1) = "P" Then
   If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" Then
     Exit Sub
   End If
 End If
If Trim(LOC_CTA_CLI) <> "" Then
 If Trim(tcuenta.Text) <> "" Then
   flagpase = ""
LETRAS:
   On Error GoTo cuenta1
  
    com_llave.AddNew
    com_llave!COM_CODCIA = wcia
    com_llave!com_cuenta = LOC_CTA_CLI
    com_llave!com_descripcion = LOC_DES_CLI
    com_llave!COM_NIVEL = LOC_NIVEL
    com_llave!com_cuenta_sup = LOC_CTA_SUP
    com_llave!COM_FLAG_AFECTACION = LOC_FLAG_AFEC
    com_llave!com_ESTADO = LOC_ESTADO
    com_llave!COM_TIPO_CTA = LOC_TIPO_CTA
    com_llave!com_signo_d = LOC_SIGNO_D
    com_llave!com_signo_h = LOC_SIGNO_H
    com_llave!COM_act_pas = LOC_ACT_PAS
    com_llave!com_signo_h = LOC_SIGNO_H
    com_llave!COM_act_pas = LOC_ACT_PAS
    com_llave!COM_DEB_MES = 0
    com_llave!COM_HAB_MES = 0
    com_llave!COM_DEB_ANO = 0
    com_llave!COM_HAB_ANO = 0
    com_llave!com_cuenta_AUTOM_D = " "
    com_llave!com_cuenta_AUTO_H = " "
    com_llave!COM_CUENTA_AUTOM_D2 = " "
    com_llave!COM_CUENTA_AUTOM_D3 = " "
    com_llave!COM_CUENTA_AUTOM_D4 = " "
    com_llave!COM_CUENTA_AUTOM_D5 = " "
    com_llave!COM_POR_AUTOM_D = 0
    com_llave!COM_POR_AUTOM_D2 = 0
    com_llave!COM_POR_AUTOM_D3 = 0
    com_llave!COM_POR_AUTOM_D4 = 0
    com_llave!COM_POR_AUTOM_D5 = 0
    com_llave!COM_CENTRO_COSTOS = " "
  com_llave.Update
  If Left(CmbCGP.Text, 1) = "P" And Left(Trim(LOC_CTA_CLI), 2) = "42" And flagpase <> "A" Then
    flagpase = "A"
    'LOC_CTA_CLI = "42101"
    LOC_CTA_CLI = Mid(Trim(LOC_CTA_CLI), 1, 2) + "3" + Mid(Trim(LOC_CTA_CLI), 4, Len(Trim(LOC_CTA_CLI)))
    GoTo LETRAS
  End If
  On Error GoTo 0
  cmdcontab.Caption = "&Quitar Relación Contable"
 End If
End If

Exit Sub
If Trim(LOC_CTA_CLI2) <> "" Then
 If Trim(tcuenta2.Text) <> "" Then
    On Error GoTo CUENTA2
    com_llave.AddNew
    com_llave!COM_CODCIA = wcia
    com_llave!com_cuenta = LOC_CTA_CLI2
    com_llave!com_descripcion = LOC_DES_CLI2
    com_llave!COM_NIVEL = LOC_NIVEL2
    com_llave!com_cuenta_sup = LOC_CTA_SUP2
    com_llave!COM_FLAG_AFECTACION = LOC_FLAG_AFEC2
    com_llave!com_ESTADO = LOC_ESTADO2
    com_llave!COM_TIPO_CTA = LOC_TIPO_CTA2
    com_llave!com_signo_d = LOC_SIGNO_D2
    com_llave!com_signo_h = LOC_SIGNO_H2
    com_llave!COM_act_pas = LOC_ACT_PAS2
    com_llave!com_signo_h = LOC_SIGNO_H2
    com_llave!COM_act_pas = LOC_ACT_PAS2
    com_llave!COM_DEB_MES = 0
    com_llave!COM_HAB_MES = 0
    com_llave!COM_HAB_ANO = 0
    com_llave!COM_DEB_ANO = 0
    com_llave!com_cuenta_AUTOM_D = ""
    com_llave!com_cuenta_AUTO_H = ""
    com_llave!COM_CUENTA_AUTOM_D2 = " "
    com_llave!COM_CUENTA_AUTOM_D3 = " "
    com_llave!COM_CUENTA_AUTOM_D4 = " "
    com_llave!COM_CUENTA_AUTOM_D5 = " "
    com_llave!COM_POR_AUTOM_D = 0
    com_llave!COM_POR_AUTOM_D2 = 0
    com_llave!COM_POR_AUTOM_D3 = 0
    com_llave!COM_POR_AUTOM_D4 = 0
    com_llave!COM_POR_AUTOM_D5 = 0
    com_llave!COM_CENTRO_COSTOS = " "
    com_llave.Update
    On Error GoTo 0
    cmdcontab2.Caption = "&Quitar Relación Contable"
 End If
End If
Exit Sub

cuenta1:
If Err.Number = 40002 Then
  MsgBox "Cuenta Existe, NO Procede. Cta.: " & LOC_CTA_CLI, 48, Pub_Titulo
   tcuenta.Text = ""
  com_llave.CancelUpdate
End If
Exit Sub

CUENTA2:
If Err.Number = 40002 Then
  MsgBox "Cuenta Existe, NO Procede. Cta.: " & LOC_CTA_CLI2, 48, Pub_Titulo
  tcuenta2.Text = ""
  com_llave.CancelUpdate
End If
Exit Sub

End Sub

Public Sub LLENA_DESCTO()

PSPLAC_LLAVE(0) = LK_CODCIA
PSPLAC_LLAVE(1) = 2301
PSPLAC_LLAVE(2) = pu_codclie
cliplac_llave.Requery
frmCLI.grid_des.Cols = 4
frmCLI.grid_des.Clear
frmCLI.grid_des.ColWidth(0) = 0
frmCLI.grid_des.ColWidth(1) = 1600
frmCLI.grid_des.ColWidth(2) = 600
frmCLI.grid_des.ColWidth(3) = 600
frmCLI.grid_des.Rows = 1
frmCLI.grid_des.TextMatrix(0, 0) = "Cod."
frmCLI.grid_des.TextMatrix(0, 1) = "Descrip."
frmCLI.grid_des.TextMatrix(0, 2) = "P.Contado"
frmCLI.grid_des.TextMatrix(0, 3) = "P.Credito"
SQ_OPER = 1
pu_codcia = LK_CODCIA
Do Until cliplac_llave.EOF
  frmCLI.grid_des.Rows = frmCLI.grid_des.Rows + 1
  frmCLI.grid_des.TextMatrix(frmCLI.grid_des.Rows - 1, 0) = cliplac_llave!tab_codart
  PUB_KEY = cliplac_llave!tab_codart
  LEER_ART_LLAVE
  If Not art_LLAVE.EOF Then frmCLI.grid_des.TextMatrix(frmCLI.grid_des.Rows - 1, 1) = art_LLAVE!art_nombre
  frmCLI.grid_des.TextMatrix(frmCLI.grid_des.Rows - 1, 2) = Format(cliplac_llave!TAB_NOMLARGO, "0.00")
  frmCLI.grid_des.TextMatrix(frmCLI.grid_des.Rows - 1, 3) = Format(cliplac_llave!tab_nomcorto, "0.00")
  cliplac_llave.MoveNext
Loop
grid_des.SetFocus
End Sub

Public Sub PASE()
Exit Sub
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_TIPMOV = ? AND ALL_CODCIA = ? AND ALL_NUMFAC = ? AND ALL_FECHA_DIA = ? AND ALL_FLAG_EXT <> 'E' ORDER BY ALL_NUMFAC"
Set PSPAR_CLI = CN.CreateQuery("", pub_cadena)
PSPAR_CLI.rdoParameters(0) = 0
PSPAR_CLI.rdoParameters(1) = 0
PSPAR_CLI.rdoParameters(2) = 0
PSPAR_CLI.rdoParameters(3) = LK_FECHA_DIA


Set par_llave_cli = PSPAR_CLI.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM FACART WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_ESTADO <> 'E' ORDER BY FAR_NUMFAC"
Set PSCLILOC_LLAVE = CN.CreateQuery("", pub_cadena)
PSCLILOC_LLAVE.rdoParameters(0) = 0
PSCLILOC_LLAVE.rdoParameters(1) = 0
PSCLILOC_LLAVE.rdoParameters(2) = LK_FECHA_DIA
Set cliloc_llave = PSCLILOC_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
PSCLILOC_LLAVE.rdoParameters(0) = 99
PSCLILOC_LLAVE.rdoParameters(1) = LK_CODCIA
PSCLILOC_LLAVE.rdoParameters(2) = "23/04/2001"
cliloc_llave.Requery
Print cliloc_llave.RowCount
Do Until cliloc_llave.EOF

PSPAR_CLI.rdoParameters(0) = cliloc_llave!FAR_tipmov
PSPAR_CLI.rdoParameters(1) = LK_CODCIA
PSPAR_CLI.rdoParameters(2) = cliloc_llave!far_numfac
PSPAR_CLI.rdoParameters(3) = cliloc_llave!far_fecha
par_llave_cli.Requery
If par_llave_cli.RowCount > 1 Then Stop
cliloc_llave.Edit
If cliloc_llave!far_numfac <> par_llave_cli!all_numfac Then Stop
cliloc_llave!FAR_fecha_can = par_llave_cli!ALL_FECHA_CAN
cliloc_llave!FAR_fecha_pro = par_llave_cli!ALL_FECHA_PRO
cliloc_llave.Update

cliloc_llave.MoveNext
Loop
MsgBox "TERMINO"
End Sub

Public Sub OPCINAL()
Dim WS_FILA  As Integer
Dim xl As Object
Stop

Stop
If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
xl.Workbooks.Open "C:\CARGA\CLIENTES.xls"  ', 0, True, 4,  WPAS, WPAS
' xl.Application.Visible = True

DoEvents
'lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
DoEvents
Dim art_rela As rdoResultset
Dim PSART_RELA As rdoQuery
pub_cadena = "select * from DIRCLI where CODCIA = ? and CODCLI = ? AND CP = ?"
Set PSART_RELA = CN.CreateQuery("", pub_cadena)
PSART_RELA(0) = 0
PSART_RELA(1) = 0
PSART_RELA(2) = 0
Set art_rela = PSART_RELA.OpenResultset(rdOpenKeyset, rdConcurValues)
PSART_RELA(0) = LK_CODCIA
art_rela.Requery
WS_FILA = 2
Do Until WS_FILA = 80000
 WS_FILA = WS_FILA + 1
 DoEvents
 cmdAgregar.Caption = WS_FILA & " / 19315"
 If Val(Trim(xl.Cells(WS_FILA, 1))) = 0 Then
   Exit Do
 End If
 GoSub GRABA_CLIENTES
Loop
MsgBox "tERMINO"
Exit Sub
' ************************************
GRABA_CLIENTES:
    If Trim(xl.Cells(WS_FILA, 16)) <> "" Then
        xl.Cells(WS_FILA, 20) = "******CESE*********"
        GoTo sale_cli
    End If
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = Val(xl.Cells(WS_FILA, 1))
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If Not cli_llave.EOF Then
'       xl.Application.Visible = True
       xl.Cells(WS_FILA, 20) = "******DUPLICADO *****"
'      MsgBox cli_llave!cli_CODCLIE
    GoTo sale_cli
    End If
    cliloc_llave.AddNew
    cliloc_llave!CLI_CP = "C"
    cliloc_llave!cli_codclie = Val(xl.Cells(WS_FILA, 1))
    cliloc_llave!cli_SALDO = 0
    cliloc_llave!CLI_DET_TOT = "D"
    cliloc_llave!CLI_MONEDA = "S"
    cliloc_llave!CLI_limcre2 = 0
    cliloc_llave!CLI_CODCIA = LK_CODCIA
    cliloc_llave!CLI_NOMBRE_ESPOSO = Trim(xl.Cells(WS_FILA, 2))
    cliloc_llave!CLI_NOMBRE_ESPOSA = ""
    cliloc_llave!CLI_NOMBRE_EMPRESA = ""
    cliloc_llave!CLI_123 = 1
    cliloc_llave!cli_nombre = Trim(xl.Cells(WS_FILA, 2))
    cliloc_llave!CLI_CASA_DIREC = Trim(xl.Cells(WS_FILA, 6))
    cliloc_llave!CLI_CASA_NUM = 0
    cliloc_llave!CLI_CASA_ZONA = 1
    cliloc_llave!CLI_LUGAR_CASA = 1
    cliloc_llave!CLI_LUGAR_TRAB = 1
    cliloc_llave!cli_dia_visita = Val((xl.Cells(WS_FILA, 9)))
    cliloc_llave!CLI_CASA_SUBZONA = 1
    cliloc_llave!cli_zona_new = Val(xl.Cells(WS_FILA, 7))
    cliloc_llave!CLI_TRAB_DIREC = ""
    cliloc_llave!CLI_TRAB_NUM = 1
    cliloc_llave!cli_TRAB_ZONA = 1
    'AGREGADO PARA GRABAR LA PROVINCIA DE TRABAJO 06/12/2001
    cliloc_llave!cli_TRAB_PROV = 1
    'AGREGADO PARA GRABAR LA PRIMERA DIRECCION
    cliloc_llave!cli_TRAB_SUBZONA = 1
    cliloc_llave!cli_ruc_esposo = Trim(xl.Cells(WS_FILA, 3))
    cliloc_llave!cli_RUC_ESPOSA = Trim(xl.Cells(WS_FILA, 4))
    cliloc_llave!CLI_RUC_EMPRESA = ""
    cliloc_llave!CLI_CASA1 = ""
    cliloc_llave!CLI_CASA2 = ""
    cliloc_llave!CLI_regpub1 = ""
    cliloc_llave!CLI_REGPUB2 = Trim(xl.Cells(WS_FILA, 19))
    cliloc_llave!CLI_AUTOAVALUO = 0
    cliloc_llave!CLI_AUTO1 = ""
    cliloc_llave!cli_auto2 = ""
    cliloc_llave!CLI_PRENDA = ""
    cliloc_llave!cli_telef1 = Trim(xl.Cells(WS_FILA, 12))
    cliloc_llave!CLI_TELEF2 = Trim(xl.Cells(WS_FILA, 13))
    cliloc_llave!CLI_OTRO_CONTR = 0
    cliloc_llave!CLI_LETRA = 0
    cliloc_llave!CLI_GRUPO = Trim(xl.Cells(WS_FILA, 18)) ' TIPO DE NEGOCIO
    cliloc_llave!CLI_SUBGRUPO = 0
    cliloc_llave!CLI_division = 0
    cliloc_llave!CLI_nucleo = 0
    If Trim(xl.Cells(WS_FILA, 16)) <> "" Then
      cliloc_llave!cli_estado = "D" ' ACT / DESCT
      cliloc_llave!cli_fecha_fac = Format(xl.Cells(WS_FILA, 16), "dd/mm/yyyy")
    Else
      cliloc_llave!cli_estado = "A" ' ACT / DESCT
    End If
    cliloc_llave!CLI_programado = 0
    cliloc_llave!CLI_PORDESCTO = 0
    
    cliloc_llave!cli_DIAS_FAC = 0
    cliloc_llave!cli_dias_cred = 0
    cliloc_llave!CLI_CUENTA_CONTAB = ""
    cliloc_llave!CLI_CUENTA_CONTAB2 = ""
    cliloc_llave!cli_cuenta_contab22 = ""
    cliloc_llave!CLI_DET_TOT = ""
    cliloc_llave!CLI_LIMCRE = 0
    cliloc_llave!CLI_CASA1 = 0
'     xl.Application.Visible = True
    'On Error GoTo siiii
'     cliloc_llave.CancelUpdate


    cliloc_llave.Update
    art_rela.AddNew
    art_rela!CODCIA = LK_CODCIA
    art_rela!DIRCLI = 1
    art_rela!CODCLI = Val(xl.Cells(WS_FILA, 1))
    art_rela!CP = "C"
    art_rela!direc = Trim(xl.Cells(WS_FILA, 6))
    art_rela!ref = Trim(xl.Cells(WS_FILA, 19))
    art_rela!CLI_LUGAR_TRAB = 1
    art_rela!cli_TRAB_ZONA = 1
    art_rela!CLI_CASA_SUBZONA = Val(xl.Cells(WS_FILA, 7))
    SQ_OPER = 1
    PUB_TIPREG = 35
    PUB_NUMTAB = Val(xl.Cells(WS_FILA, 7))
    PUB_CODCIA = "00"
    LEER_TAB_LLAVE
    art_rela!cli_TRAB_SUBZONA = 1
    art_rela!NUMERO = 0
    art_rela!DIRCOMP = Trim(xl.Cells(WS_FILA, 6)) & " - " & Trim(tab_llave!TAB_NOMLARGO)
    art_rela.Update
sale_cli:
Return



SQ_OPER = 2
pu_codclie = 0
pu_codcia = LK_CODCIA
LEER_CLI_LLAVE

xl.Workbooks.Open "C:\LLEVO\CLIENTES.xls"  ', 0, True, 4,  WPAS, WPAS
xl.Application.Visible = True
WS_FILA = 3

Do Until Trim(xl.Cells(WS_FILA, 1)) = ""
    If Val(Trim(xl.Cells(WS_FILA, 1))) = 0 Then
      MsgBox "no agregado"
      GoTo SALTA_ARTI
    End If
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = Val(xl.Cells(WS_FILA, 1))
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If cli_llave.EOF Then
'     Stop
     GoTo SALTA_ARTI
    End If
'     xl.Application.Visible = False
  art_rela.AddNew
  art_rela!CODCIA = LK_CODCIA
  art_rela!DIRCLI = 1
  art_rela!CODCLI = cli_llave!cli_codclie
  art_rela!CP = "C"
  art_rela!direc = Trim(cli_llave!CLI_CASA_DIREC)
  art_rela!ref = Trim(xl.Cells(WS_FILA, 18))
  art_rela!CLI_LUGAR_TRAB = 1
  art_rela!cli_TRAB_ZONA = 1
  art_rela!CLI_CASA_SUBZONA = 1
  art_rela!cli_TRAB_SUBZONA = 1
  art_rela!NUMERO = 0
  art_rela!DIRCOMP = Trim(cli_llave!CLI_CASA_DIREC)
  art_rela.Update
cli_llave.Edit
cli_llave!CLI_REGPUB2 = Trim(xl.Cells(WS_FILA, 18))
cli_llave!CLI_CASA2 = Trim(xl.Cells(WS_FILA, 18))
cli_llave!cli_telef1 = Trim(xl.Cells(WS_FILA, 12))
cli_llave.Update

  
'     xl.Application.Visible = False
SALTA_ARTI:
    WS_FILA = WS_FILA + 1
Loop
MsgBox " TEWRMINO "
Exit Sub
siiii:
cliloc_llave.CancelUpdate
GoTo sale_cli

Exit Sub
 ' *******- agregar zonas
 If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
xl.Workbooks.Open "C:\CARGAXXX\TABGRUP.xls"  ', 0, True, 4,  WPAS, WPAS
xl.Application.Visible = True
WS_FILA = 1
Do Until WS_FILA = 989898
WS_FILA = WS_FILA + 1
If Val(xl.Cells(WS_FILA, 3)) = 0 Then Exit Do
tab_llave.Requery
tab_llave.AddNew
tab_llave!tab_codcia = LK_CODCIA
tab_llave!tab_tipreg = 123  ' MODIFICABLE PARA CUALQUIER TIPREG
tab_llave!TAB_NUMTAB = Val(Trim(xl.Cells(WS_FILA, 3)))
tab_llave!TAB_NOMLARGO = xl.Cells(WS_FILA, 4)
tab_llave!tab_nomcorto = xl.Cells(WS_FILA, 4)
tab_llave!tab_codart = Val(Left(Trim(xl.Cells(WS_FILA, 3)), 2))
tab_llave!tab_codclie = 0
tab_llave!TAB_contable2 = 0

tab_llave.Update
Loop


Exit Sub
 '******VENDEDORES ***************
If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
xl.Workbooks.Open "C:\CARGA\TABVEN.xls"  ', 0, True, 4,  WPAS, WPAS
xl.Application.Visible = True
WS_FILA = 2
ven_llave.Requery


Do Until WS_FILA = 989898
WS_FILA = WS_FILA + 1
If Val(xl.Cells(WS_FILA, 1)) = 0 Then Exit Do
ven_llave.AddNew
ven_llave!VEM_CODVEN = Val(xl.Cells(WS_FILA, 1))
ven_llave!VEM_NOMBRE = Trim(xl.Cells(WS_FILA, 5))
ven_llave!vem_codcia = LK_CODCIA
ven_llave!VEM_SERIE_G = 0
ven_llave!VEM_NUMFAC_G_INI = 0
ven_llave!vem_serie_b = 0
ven_llave!VEM_NUMFAC_B_INI = 0
ven_llave!vem_serie_f = 0
ven_llave!VEM_NUMFAC_F_INI = 0
ven_llave!VEM_NUMFAC_G_FIN = 0
ven_llave!VEM_NUMFAC_B_FIN = 0
ven_llave!VEM_NUMFAC_F_FIN = 0
ven_llave!VEM_SERIE_P = 0
ven_llave!VEM_NUMFAC_P_INI = 0
ven_llave!VEM_NUMFAC_P_FIN = 0
ven_llave!VEM_FLAG_P = " "
ven_llave!VEM_SERIE_N = 0
ven_llave!VEM_SERIE_D = 0
ven_llave!VEM_NUMFAC_N_INI = 0
ven_llave!VEM_NUMFAC_D_INI = 0
ven_llave!VEM_NUMFAC_N_FIN = 0
ven_llave!VEM_NUMFAC_D_FIN = 0
ven_llave!VEM_FLAG_N = " "
ven_llave!VEM_FLAG_D = " "
ven_llave!VEM_FECHA_ING = Format(Trim(xl.Cells(WS_FILA, 8)), "dd/mm/yyyy")
ven_llave!VEM_DIRECCION = ""
ven_llave!VEM_TELE_CASA = ""
ven_llave!VEM_TELE_CELU = ""
ven_llave!VEM_SERIE_R = 0
ven_llave!VEM_FLAG_G = " "
ven_llave!VEM_FLAG_B = " "
ven_llave!VEM_FLAG_F = " "
ven_llave.Update


Loop


Exit Sub
SQ_OPER = 2
PUB_KEY = 0
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
If art_mayor.EOF Then
  Exit Sub
End If
Do Until art_mayor.EOF
 SQ_OPER = 1
 PUB_NUMTAB = art_mayor!art_familia
 PUB_TIPREG = 122
 LEER_TAB_LLAVE
 If tab_llave.EOF Then
   Stop
 End If
 PUB_TIPREG = 123
 PUB_NUMTAB = art_mayor!art_subfam
 LEER_TAB_LLAVE
 If tab_llave.EOF Then
 Print art_mayor!art_key
  MsgBox art_mayor!ART_alterno & "  " & art_mayor!art_subfam
'   Stop
 Else
    SQ_OPER = 1
    PUB_NUMTAB = tab_llave!tab_codart
    PUB_TIPREG = 122
    LEER_TAB_LLAVE
    If tab_llave.EOF Then
      Stop
    End If
 End If
 art_mayor.MoveNext
Loop

MsgBox "TERMINO"
End Sub
Public Sub LLENA_VEN(WSTit As Integer)
Dim ps_vendedor As rdoResultset
Dim W1 As String * 2
Dim I, wPosF, WPosV, cuenta As Integer
Dim sal As Boolean
Dim cade As String
Dim WNUM As Integer
Dim F As Integer
Dim a As Integer
WNUM = 0
wPosF = 0
WPosV = 0
lstVen.Clear

pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = '" & LK_CODCIA & "' AND  VEM_DESACTIVO <> 'A' "
Set ps_vendedor = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
  
cuenta = 0
PS_PAR(0) = " "
lstVen.Clear
Do Until ps_vendedor.EOF
     lstVen.AddItem Format(ps_vendedor!VEM_CODVEN, "000") & " " & ps_vendedor!VEM_NOMBRE
     cuenta = cuenta + 1
     ps_vendedor.MoveNext
Loop

cuenta = 0
WPosV = Len(tcuenta22.Text)
cade = Trim(tcuenta22.Text)
cuenta = 0
wPosF = 1
a = 0
For I = 1 To Len(cade)
If Mid(cade, I, 1) = "." Then
  a = a + 1
End If
Next I
Do Until cuenta = a
   cuenta = cuenta + 1
   wPosF = InStr(wPosF, cade, ".", 1) + 1
   WNUM = Mid(cade, wPosF, 3)
   If Right(WNUM, 1) = "." Then
     WNUM = Left(WNUM, 3)
     wPosF = wPosF - 1
   End If
   For I = 0 To lstVen.ListCount - 1
    lstVen.ListIndex = I
    If Trim(Left(lstVen.Text, 3)) = Format(CStr(WNUM), "00") Then
       lstVen.Selected(I) = True
       Exit For
    End If
   Next I
Loop

End Sub


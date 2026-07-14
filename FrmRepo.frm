VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025601-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmRepo 
   Caption         =   "Reportes de Inventarios"
   ClientHeight    =   5895
   ClientLeft      =   195
   ClientTop       =   1485
   ClientWidth     =   11220
   ControlBox      =   0   'False
   ForeColor       =   &H00808000&
   Icon            =   "FrmRepo.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5895
   ScaleWidth      =   11220
   Begin VB.Frame franego 
      Height          =   1530
      Left            =   3345
      TabIndex        =   67
      Top             =   3600
      Visible         =   0   'False
      Width           =   3000
      Begin VB.ListBox tipon 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   105
         Style           =   1  'Checkbox
         TabIndex        =   68
         Top             =   315
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.Label LBL2 
         Caption         =   "Tipo de Negocio :"
         Height          =   255
         Index           =   15
         Left            =   75
         TabIndex        =   69
         Top             =   105
         Width           =   1815
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   495
      Left            =   8040
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.Frame lisdocu 
      Height          =   3735
      Left            =   6885
      TabIndex        =   63
      Top             =   150
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton cmdrela 
         Caption         =   "Lista Relación."
         Height          =   375
         Left            =   600
         TabIndex        =   66
         Top             =   3240
         Width           =   2175
      End
      Begin VB.ListBox docu 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   64
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label LBL2 
         Caption         =   "Lista de Documentos :"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   65
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.CommandButton SALIR 
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
      Height          =   375
      Left            =   8400
      TabIndex        =   18
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame fraconsi 
      Caption         =   "Selección de datos "
      Height          =   5175
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   11055
      Begin VB.CheckBox chesub 
         Caption         =   "Mostar solo Productos de Sub Almacen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   6600
         TabIndex        =   72
         Top             =   3840
         Width           =   4215
      End
      Begin VB.Frame frazonas 
         Height          =   3495
         Left            =   3360
         TabIndex        =   54
         Top             =   120
         Width           =   2895
         Begin VB.ListBox listdistrito 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   4
            Top             =   300
            Width           =   2655
         End
         Begin VB.ListBox listprovincia 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Left            =   120
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Label LBL2 
            Caption         =   "Distritos"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label LBL2 
            Caption         =   "Zonas"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   55
            Top             =   1680
            Width           =   1095
         End
      End
      Begin VB.ListBox lstcanal 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   165
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   70
         Top             =   3495
         Width           =   2655
      End
      Begin VB.Frame fraven 
         Height          =   2415
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   3135
         Begin VB.ListBox listven 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1980
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   3
            Top             =   360
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label LBL2 
            Caption         =   "Vendedores para Filtrar"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   58
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.ComboBox moneda 
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
         ItemData        =   "FrmRepo.frx":0442
         Left            =   7320
         List            =   "FrmRepo.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame fradocu 
         Height          =   2535
         Left            =   6410
         TabIndex        =   33
         Top             =   120
         Width           =   2775
         Begin VB.CheckBox cheguia 
            Caption         =   "Guias"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox fac 
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   7
            TabIndex        =   7
            Top             =   525
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CheckBox chefac 
            Caption         =   "Facturas"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   975
         End
         Begin VB.CheckBox chebol 
            Caption         =   "Boletas"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox fac 
            Height          =   285
            Index           =   1
            Left            =   840
            MaxLength       =   10
            TabIndex        =   8
            Top             =   525
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox bol 
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   7
            TabIndex        =   11
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox guia 
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   7
            TabIndex        =   15
            Top             =   2085
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox bol 
            Height          =   285
            Index           =   1
            Left            =   840
            MaxLength       =   10
            TabIndex        =   12
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox guia 
            Height          =   285
            Index           =   1
            Left            =   960
            MaxLength       =   7
            TabIndex        =   16
            Top             =   2085
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox fac 
            Height          =   285
            Index           =   2
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   9
            Top             =   525
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox bol 
            Height          =   285
            Index           =   2
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   13
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label LBL2 
            AutoSize        =   -1  'True
            Caption         =   "Serie"
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
            Index           =   5
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label LBL2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Inicial"
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
            Index           =   6
            Left            =   840
            TabIndex        =   46
            Top             =   360
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label LBL2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Final"
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
            Index           =   7
            Left            =   1800
            TabIndex        =   44
            Top             =   360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label LBL2 
            AutoSize        =   -1  'True
            Caption         =   "Serie"
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
            Index           =   8
            Left            =   120
            TabIndex        =   38
            Top             =   1155
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label LBL2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Inicial"
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
            Index           =   9
            Left            =   840
            TabIndex        =   37
            Top             =   1155
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label LBL2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Final"
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
            Index           =   10
            Left            =   1800
            TabIndex        =   36
            Top             =   1155
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label LBL2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Final"
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
            Index           =   12
            Left            =   1080
            TabIndex        =   35
            Top             =   1920
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label LBL2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Inicial"
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
            Index           =   11
            Left            =   240
            TabIndex        =   34
            Top             =   1905
            Visible         =   0   'False
            Width           =   570
         End
      End
      Begin VB.CommandButton cmdmostrar 
         Caption         =   "&Mostrar / Imprimir"
         Height          =   375
         Left            =   6600
         TabIndex        =   17
         Top             =   3240
         Width           =   2175
      End
      Begin MSMask.MaskEdBox fechaini 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox fechafin 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_cli 
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
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LBL2 
         Caption         =   "Canal de Distribución"
         Height          =   180
         Index           =   16
         Left            =   195
         TabIndex        =   71
         Top             =   3285
         Width           =   1950
      End
      Begin VB.Label lmoneda 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   6600
         TabIndex        =   62
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblcliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   60
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label LBL2 
         Caption         =   "Cliente :"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   59
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label LBL2 
         Caption         =   "Fecha Final"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LBL2 
         Caption         =   "Fecha Inicial"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fra2 
      Height          =   2895
      Left            =   720
      TabIndex        =   19
      Top             =   0
      Width           =   9135
      Begin VB.Frame opciones 
         Height          =   2895
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton Option1 
            Caption         =   "Solo Ventas del Dia"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   43
            Top             =   1800
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Proveedor"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   41
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Familias"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Articulo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   42
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label etiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Opciones"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.TextBox txtMaxCol 
         Height          =   285
         Left            =   5160
         MaxLength       =   2
         TabIndex        =   52
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox cmbtipo 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton CmdProcesa 
         Caption         =   "&Listar Articulo (s)"
         Height          =   495
         Left            =   7440
         TabIndex        =   53
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ListBox PROV 
         Height          =   1635
         Left            =   1680
         Style           =   1  'Checkbox
         TabIndex        =   50
         Top             =   360
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.ListBox lineas 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1680
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ListBox listapro 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1680
         TabIndex        =   49
         Top             =   1440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txt_key 
         Height          =   315
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbltipo 
         Caption         =   "Acumulado por :"
         Height          =   255
         Left            =   5160
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblmax 
         Caption         =   "Maximo de Columnas"
         Height          =   255
         Left            =   5160
         TabIndex        =   24
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sub Familia"
         Height          =   195
         Left            =   1680
         TabIndex        =   23
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Familia :"
         Height          =   195
         Left            =   1680
         TabIndex        =   22
         Top             =   120
         Width           =   570
      End
      Begin VB.Label lblopcion 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2880
         TabIndex        =   20
         Top             =   240
         Width           =   2085
      End
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   5400
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   840
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblmensa 
      Caption         =   "Procensando  . . . "
      Height          =   255
      Left            =   2400
      TabIndex        =   28
      Top             =   5160
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "FrmRepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_cp  As String * 1
Dim xl As Object
Dim llave1  As String
Dim loc_key  As Integer
Dim PSART_LOC As rdoQuery
Dim PSART_LOC2 As rdoQuery
Dim PSPRO_V As rdoQuery
Dim artloc_llave As rdoResultset
Dim artloc_mayor As rdoResultset
Dim PRO_V As rdoResultset
Dim PSFAR_VENTAS As rdoQuery
Dim far_ventas As rdoResultset
Dim loc_row As Integer
Dim temporal As String
Dim LOC_DETENER  As Integer
Dim PRE_UNIDADES  As rdoQuery
Dim preuni_llave  As rdoResultset
Dim WCOD_ORIGINAL As Currency
Dim Wfile As String
Dim REP_FECHA1 As String
Dim REP_FECHA2 As String
Dim DIA As String
Dim mes As String
Dim ano As String
Dim DIA1 As String
Dim MES1 As String
Dim ANO1 As String



Private Sub bol_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
 If Index = 0 Then
  bol(1).SetFocus
 End If
 If Index = 1 Then
  bol(2).SetFocus
 End If
 If Index = 2 Then
  cheguia.SetFocus
 End If
 Exit Sub
End If
SOLO_ENTERO KeyAscii

End Sub

Private Sub chebol_Click()
If Wfile = "IMP_FACTURACION" Then
 chefac.Value = 0
 cheguia.Value = 0
End If

If chebol.Value = 1 Then
 LBL2(8).Visible = True
 LBL2(9).Visible = True
 LBL2(10).Visible = True
 bol(0).Visible = True
 bol(1).Visible = True
 bol(2).Visible = True
 Azul bol(0), bol(0)
Else
 LBL2(8).Visible = False
 LBL2(9).Visible = False
 LBL2(10).Visible = False
 bol(0).Visible = False
 bol(1).Visible = False
 bol(2).Visible = False
End If

End Sub

Private Sub chebol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 chebol.SetFocus
End If

End Sub

Private Sub chefac_Click()
If Wfile = "IMP_FACTURACION" Then
 chebol.Value = 0
 cheguia.Value = 0
End If
If chefac.Value = 1 Then
 LBL2(5).Visible = True
 LBL2(6).Visible = True
 LBL2(7).Visible = True
 fac(0).Visible = True
 fac(1).Visible = True
 fac(2).Visible = True
 Azul fac(0), fac(0)
Else
 LBL2(5).Visible = False
 LBL2(6).Visible = False
 LBL2(7).Visible = False
 fac(0).Visible = False
 fac(1).Visible = False
 fac(2).Visible = False
End If

End Sub

Private Sub chefac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 chebol.SetFocus
End If

End Sub

Private Sub cheguia_Click()
If Wfile = "IMP_FACTURACION" Then
 chebol.Value = 0
 chefac.Value = 0
End If

If cheguia.Value = 1 Then
 LBL2(11).Visible = True
 LBL2(12).Visible = True
 guia(0).Visible = True
 guia(1).Visible = True
 Azul guia(0), guia(0)
 Else
 LBL2(11).Visible = False
 LBL2(12).Visible = False
 guia(0).Visible = False
 guia(1).Visible = False
End If

End Sub

Private Sub cheguia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 listven.SetFocus
End If

End Sub

Private Sub cmbTipo_Click()
'If txtMaxCol.text <> "" Then
' txtMaxCol.SetFocus
'End If
End Sub

Private Sub cmbtipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtMaxCol.SetFocus
End If
End Sub

Private Sub cmdmostrar_Click()

Dim wformula, wformula1, wformula2, wformula3, wformula4
Dim Modo, Modo1
Dim wsFECHA1, wsFECHA2
Dim Wche, wkSELECT
Dim wfiltra1, wfiltra2, wfiltra3, wfecha, WDOCU
Dim wscliente As String
Dim wcodcia As String
If Wfile = "IMP_FACTURACION" Then
 If FrmRepo.chebol.Value = 0 And FrmRepo.chefac.Value = 0 And FrmRepo.cheguia.Value = 0 Then
   MsgBox "Seleccione un tipo de documento.", 48, Pub_Titulo
   Exit Sub
 End If
 IMP_FACTURACION
 Exit Sub
End If
Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(retra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390

If Wfile = "PLANILLA" Or Wfile = "CONSOLIDADO" Or Wfile = "CONTADO" Or Wfile = "PEDTN.RPT" Or Wfile = "HRESUVE" Then
  GoTo CONSO
End If
If Wfile = "VENTASXCLI" Then
  GoTo VENTASXCLI
End If
Exit Sub

CONSO:
On Error GoTo procancela
cmdmostrar.Enabled = False
If Right(fechaini.Text, 2) = "__" Then
  wsFECHA1 = Left(fechaini.Text, 8)
Else
 wsFECHA1 = Trim(fechaini.Text)
End If
If Not IsDate(wsFECHA1) Then
    cmdmostrar.Enabled = True
    MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
    Azul2 fechaini, fechaini
    Exit Sub
End If
If Right(fechafin.Text, 2) = "__" Then
  wsFECHA2 = Left(fechafin.Text, 8)
Else
  wsFECHA2 = Trim(fechafin.Text)
End If
If Not IsDate(wsFECHA2) Then
 cmdmostrar.Enabled = True
 MsgBox "Fecha Invalidad ..", 48, pu_titulo
 Azul2 fechafin, fechafin
 Exit Sub
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 cmdmostrar.Enabled = True
 MsgBox "Fecha Invalidad ..", 48, pu_titulo
 Azul2 fechaini, fechaini
 Exit Sub
End If
REP_FECHA1 = wsFECHA1
REP_FECHA2 = wsFECHA2

If Wfile = "PLANILLA" Or Wfile = "CONSOLIDADO" And LK_EMP = "PAR" Then
  pub_mensaje = "Desea Mostrar los Documentos del rango de Fechas ? "
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta <> vbNo Then
     PROCESA_DOCU
     Exit Sub
  End If
End If

lblmensa.Visible = True
DoEvents
pbMin = 0
pbMax = 10
pbValue = 0
pb.Visible = True
If wsFECHA1 = wsFECHA2 Then
 wfecha = wsFECHA1
Else
 wfecha = wsFECHA1 & " AL " & wsFECHA2
End If
If Wfile = "CONTADO" Then
 If wsFECHA1 = wsFECHA2 Then
  wfecha = "VENTAS AL CONTADO DEL DIA " & wsFECHA1
 Else
  wfecha = "VENTAS AL CONTADO DEL DIA " & wsFECHA1 & " AL " & wsFECHA2
 End If
End If

pbValue = pbValue + 1
CADENITA = ""
CADEFECHA = ""
wfiltra1 = ""
Wche = 0
Modo1 = "{FACART.FAR_CODVEN} in [" ''F']"
For fila = 0 To listven.ListCount - 1
  listven.ListIndex = fila
  If listven.Selected(fila) And Trim(listven.Text) <> "" Then
    Wche = 1
    wkSELECT = Str(Val(Left(listven.Text, 3)))
    wfiltra1 = wfiltra1 + wkSELECT + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
 CADENITA = Left(Modo1, Len(Modo1) - 1) & "] AND "
 wfiltra1 = Left(wfiltra1, Len(wfiltra1) - 1)
End If
pbValue = pbValue + 1
wfiltra2 = ""
Wche = 0
Modo1 = "{CLIENTES.CLI_TRAB_ZONA} in ["
For fila = 0 To listdistrito.ListCount - 1
  listdistrito.ListIndex = fila
  If listdistrito.Selected(fila) And Trim(listdistrito.Text) <> "" Then
    Wche = 1
    wkSELECT = Trim(Right(listdistrito.Text, 10))
    wfiltra2 = wfiltra2 + Trim(Left(listdistrito.Text, 30)) + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
 CADENITA = CADENITA + Left(Modo1, Len(Modo1) - 1) & "] AND "
 wfiltra2 = Left(wfiltra2, Len(wfiltra2) - 1)
End If

wfiltra2 = ""
Wche = 0
Modo1 = "{CLIENTES.CLI_GRUPO} in ["
For fila = 0 To tipon.ListCount - 1
  tipon.ListIndex = fila
  If tipon.Selected(fila) And Trim(tipon.Text) <> "" Then
    Wche = 1
    wkSELECT = Trim(Right(tipon.Text, 10))
    wfiltra2 = wfiltra2 + Trim(Left(tipon.Text, 30)) + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
 CADENITA = CADENITA + Left(Modo1, Len(Modo1) - 1) & "] AND "
 wfiltra2 = Left(wfiltra2, Len(wfiltra2) - 1)
End If

wfiltra2 = ""
Wche = 0
Modo1 = "{FACART.FAR_TURNO} in ["
For fila = 0 To lstcanal.ListCount - 1
  lstcanal.ListIndex = fila
  If lstcanal.Selected(fila) And Trim(lstcanal.Text) <> "" Then
    Wche = 1
    wkSELECT = Trim(Right(lstcanal.Text, 10))
    wfiltra2 = wfiltra2 + Trim(Left(lstcanal.Text, 30)) + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
 CADENITA = CADENITA + Left(Modo1, Len(Modo1) - 1) & "] AND "
 wfiltra2 = Left(wfiltra2, Len(wfiltra2) - 1)
End If



pbValue = pbValue + 1
wfiltra3 = ""
Wche = 0
Modo1 = "{CLIENTES.CLI_ZONA_NEW} in ["
For fila = 0 To listprovincia.ListCount - 1
  listprovincia.ListIndex = fila
  If listprovincia.Selected(fila) And Trim(listprovincia.Text) <> "" Then
    Wche = 1
    wkSELECT = Trim(Right(listprovincia.Text, 10))
    wfiltra3 = wfiltra3 + Trim(Left(listprovincia.Text, 30)) + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
 CADENITA = CADENITA + Left(Modo1, Len(Modo1) - 1) & "] AND "
 wfiltra3 = Left(wfiltra3, Len(wfiltra3) - 1)
End If
pbValue = pbValue + 1

wfiltra = ""
If wfiltra2 <> "" Then
 wfiltra = wfiltra + "DISTRITO : " & wfiltra2
End If
If wfiltra3 <> "" Then
 wfiltra = wfiltra + " ZONA : " & wfiltra3
End If
If wfiltra1 <> "" Then
 wfiltra = wfiltra + " VENDEDOR : " & wfiltra1
End If
pbValue = pbValue + 1
DIA = Day(wsFECHA1)
mes = Month(wsFECHA1)
ano = Year(wsFECHA1)

DIA1 = Day(wsFECHA2)
MES1 = Month(wsFECHA2)
ANO1 = Year(wsFECHA2)

If Wfile = "CONTADO" Then
 CADEFECHA = CADEFECHA + " {FACART.FAR_SIGNO_CAR} > -1 AND {FACART.FAR_SIGNO_CAR} < 1   AND {FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FECHA} >= Date ( " & ano & "," & mes & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
ElseIf Wfile = "PEDTN.RPT" Then
 CADEFECHA = CADEFECHA + "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 10  AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FECHA} >= Date ( " & ano & "," & mes & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
Else
 CADEFECHA = CADEFECHA + "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 10 AND {PRECIOS.PRE_FLAG_UNIDAD} = 'A' AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FECHA} >= Date ( " & ano & "," & mes & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
End If
CADENITA = CADENITA + CADEFECHA + " AND "
WDOCUF = ""
wdocuB = ""
Modo1 = "X"
cade = ""
If chefac.Value = 1 And chebol.Value = 0 And cheguia.Value = 0 Then
  Modo1 = "{FACART.FAR_FBG} in ['F'] "
  wdocuB = "FAC."
  If Trim(fac(0).Text) <> "" Then
    cade = cade + "{FACART.FAR_NUMSER} = '" & Trim(fac(0).Text) & "' "
    WDOCUF = WDOCUF & " - SERIE : " & Trim(fac(0).Text)
    If Trim(fac(1).Text) <> "" And Trim(fac(2).Text) <> "" Then
      cade = cade + "AND {FACART.FAR_NUMFAC} >= " & Trim(fac(1).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(fac(2).Text) & " "
      WDOCUF = WDOCUF & " Nº. " & Trim(fac(1).Text) & " AL " & Trim(fac(2).Text)
    End If
  End If
End If

If chefac.Value = 0 And chebol.Value = 1 And cheguia.Value = 0 Then
  Modo1 = "{FACART.FAR_FBG} in ['B'] "
  wdocuB = " BOL."
  If Trim(bol(0).Text) <> "" Then
    cade = cade + "{FACART.FAR_NUMSER} = '" & Trim(bol(0).Text) & "' "
    wdocuB = wdocuB & " - SERIE : " & Trim(bol(0).Text)
    If Trim(bol(1).Text) <> "" And Trim(bol(2).Text) <> "" Then
      cade = cade + "AND {FACART.FAR_NUMFAC} >= " & Trim(bol(1).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(bol(2).Text) & " "
      wdocuB = wdocuB & " Nº. " & Trim(bol(1).Text) & " AL " & Trim(bol(2).Text)
    End If
  End If
End If

If chefac.Value = 0 And chebol.Value = 0 And cheguia.Value = 1 Then
  Modo1 = "{FACART.FAR_FBG} in ['G'] "
  wdocuB = " GUIA"
  If Trim(guia(0).Text) <> "" And Trim(guia(1).Text) <> "" Then
    cade = cade + "{FACART.FAR_NUMFAC} >= " & Trim(guia(0).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(guia(1).Text) & " "
    wdocug = wdocug & " Nº. " & Trim(guia(0).Text) & " AL " & Trim(guia(1).Text)
  End If
End If

If chefac.Value = 1 And chebol.Value = 1 And cheguia.Value = 0 Then
  Modo1 = "{FACART.FAR_FBG} in ['F'] "
  wdocuB = " FAC.,BOL."
  If Trim(fac(0).Text) <> "" Then
    cade = cade + "{FACART.FAR_NUMSER} = '" & Trim(fac(0).Text) & "' "
    WDOCUF = WDOCUF & " - SERIE : " & Trim(fac(0).Text)
    If Trim(fac(1).Text) <> "" And Trim(fac(2).Text) <> "" Then
      cade = cade + "AND {FACART.FAR_NUMFAC} >= " & Trim(fac(1).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(fac(2).Text) & " "
      WDOCUF = WDOCUF & " Nº. " & Trim(fac(1).Text) & " AL " & Trim(fac(2).Text)
    End If
  End If
  wdocug = ""
  If Trim(bol(0).Text) <> "" Then
    If cade <> "" Then
       cade = cade + "OR {FACART.FAR_FBG} in ['B'] AND {FACART.FAR_NUMSER} = '" & Trim(bol(0).Text) & "' "
       wdocug = "1"
    Else
       Modo1 = "{FACART.FAR_FBG} in ['F','B'] "
       cade = cade + " {FACART.FAR_NUMSER} = '" & Trim(bol(0).Text) & "' "
    End If
    WDOCUF = WDOCUF & ", - SERIE : " & Trim(bol(0).Text)
    If Trim(bol(1).Text) <> "" And Trim(bol(2).Text) <> "" Then
      If wdocug = "1" Then
        cade = cade + "AND {FACART.FAR_NUMFAC} >= " & Trim(bol(1).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(bol(2).Text) & " "
      Else
        cade = cade + "AND {FACART.FAR_NUMFAC} >= " & Trim(bol(1).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(bol(2).Text) & " "
      End If
      WDOCUF = WDOCUF & ", Nº. " & Trim(bol(1).Text) & " AL " & Trim(bol(2).Text)
    Else
      cade = cade + " "
    End If
  Else
   If cade <> "" Then
     cade = cade + "OR ({FACART.FAR_FBG} in ['B']"
    Else
     Modo1 = "{FACART.FAR_FBG} in ['F','B'] "
   End If
  End If
End If

If chefac.Value = 1 And chebol.Value = 0 And cheguia.Value = 1 Then
  Modo1 = "{FACART.FAR_FBG} in ['F'] "
  wdocuB = " FAC.,GUIA"
  If Trim(fac(0).Text) <> "" Then
    cade = cade + "{FACART.FAR_NUMSER} = '" & Trim(fac(0).Text) & "' "
    WDOCUF = WDOCUF & " - SERIE : " & Trim(fac(0).Text)
    If Trim(fac(1).Text) <> "" And Trim(fac(2).Text) <> "" Then
      cade = cade + "AND {FACART.FAR_NUMFAC} >= " & Trim(fac(1).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(fac(2).Text) & " "
      WDOCUF = WDOCUF & " Nº. " & Trim(fac(1).Text) & " AL " & Trim(fac(2).Text)
    End If
  End If
  If Trim(guia(0).Text) <> "" And Trim(guia(1).Text) <> "" Then
    If cade <> "" Then
      cade = cade + " OR ({FACART.FAR_FBG} in ['G'] AND {FACART.FAR_NUMFAC} >= " & Trim(guia(0).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(guia(1).Text) & ") "
    Else
     cade = cade + " ({FACART.FAR_FBG} in ['G'] AND {FACART.FAR_NUMFAC} >= " & Trim(guia(0).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(guia(1).Text) & ") "
    End If
    WDOCUF = WDOCUF & ", Nº. " & Trim(guia(0).Text) & " AL " & Trim(guia(1).Text)
  End If
End If
If chefac.Value = 0 And chebol.Value = 1 And cheguia.Value = 1 Then
  Modo1 = "{FACART.FAR_FBG} in ['B','G'] "
  wdocuB = " BOL.,GUIA"
  If Trim(bol(0).Text) <> "" Then
    cade = cade + "{FACART.FAR_NUMSER} = '" & Trim(bol(0).Text) & "' "
    WDOCUF = WDOCUF & ", - SERIE : " & Trim(bol(0).Text)
    If Trim(bol(1).Text) <> "" And Trim(bol(2).Text) <> "" Then
      cade = cade + "AND {FACART.FAR_NUMFAC} >= " & Trim(bol(1).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(bol(2).Text) & " "
      WDOCUF = WDOCUF & ", Nº. " & Trim(bol(1).Text) & " AL " & Trim(bol(2).Text)
    End If
  End If
  If Trim(guia(0).Text) <> "" And Trim(guia(1).Text) <> "" Then
    cade = cade + "OR ({FACART.FAR_NUMFAC} >= " & Trim(guia(0).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(guia(1).Text) & ") "
    WDOCUF = WDOCUF & ", Nº. " & Trim(guia(0).Text) & " AL " & Trim(guia(1).Text)
  End If
End If
pbValue = pbValue + 1
If chefac.Value = 1 And chebol.Value = 1 And cheguia.Value = 1 Then
  Modo1 = "{FACART.FAR_FBG} in ['F','B','G'] "
  wdocuB = " FAC.,BOL.,GUIA"
  If Trim(fac(0).Text) <> "" Then
    cade = cade + "{FACART.FAR_NUMSER} = '" & Trim(fac(0).Text) & "' "
    WDOCUF = WDOCUF & " - SERIE : " & Trim(fac(0).Text)
    If Trim(fac(1).Text) <> "" And Trim(fac(2).Text) <> "" Then
      cade = cade + "AND {FACART.FAR_NUMFAC} >= " & Trim(fac(1).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(fac(2).Text) & " "
      WDOCUF = WDOCUF & " Nº. " & Trim(fac(1).Text) & " AL " & Trim(fac(2).Text)
    End If
  End If
  If Trim(bol(0).Text) <> "" Then
    cade = cade + "OR ({FACART.FAR_NUMSER} = '" & Trim(bol(0).Text) & "' "
    WDOCUF = WDOCUF & ", - SERIE : " & Trim(bol(0).Text)
    If Trim(bol(1).Text) <> "" And Trim(bol(2).Text) <> "" Then
      cade = cade + "AND {FACART.FAR_NUMFAC} >= " & Trim(bol(1).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(bol(2).Text) & ") "
      WDOCUF = WDOCUF & ", Nº. " & Trim(bol(1).Text) & " AL " & Trim(bol(2).Text)
    Else
      cade = cade + ")"
    End If
  End If
  If Trim(guia(0).Text) <> "" And Trim(guia(1).Text) <> "" Then
    cade = cade + "OR ({FACART.FAR_NUMFAC} >= " & Trim(guia(0).Text) & " AND {FACART.FAR_NUMFAC} <= " & Trim(guia(1).Text) & ") "
    WDOCUF = WDOCUF & ", Nº. " & Trim(guia(0).Text) & " AL " & Trim(guia(1).Text)
  End If
End If
If Modo1 <> "X" Then
 CADENITA_TEMPO = CADENITA + "(" + Modo1 + " AND "
 CADENITA = CADENITA + Modo1 + " AND "
 If cade <> "" Then
  CADENITA = CADENITA_TEMPO
  CADENITA = CADENITA + cade + ") AND "
 Else
   CADENITA = CADENITA + Modo1 + " AND "
 End If
 
End If
pbValue = pbValue + 1
If Trim(Right(CADENITA, 2)) = "D" Then
 CADENITA = Left(CADENITA, Len(CADENITA) - 4)
End If
If Trim(Right(CADENITA, 2)) = "R" Then
 CADENITA = Left(CADENITA, Len(CADENITA) - 4)
End If
pbValue = pbValue + 1
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.Formulas(3) = ""
If retra_llave!tra_rep1 = "1" Then
  If LK_EMP_PTO = "A" Then
    Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\PTOVTA\" & "hresu.rpt"
    wcodcia = "00"
  Else
    Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hresu.rpt"
    wcodcia = LK_CODCIA
  End If
Else
  Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hresu.rpt"
  wcodcia = LK_CODCIA
End If
If Wfile = "HRESUVE" Then
  Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hresuve.rpt"
End If
If Wfile = "CONTADO" Then
  Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "CONTADO.rpt"
End If
If Wfile = "PEDTN.RPT" Then
  Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "PEDTN.RPT"
  wcodcia = LK_CODCIA
End If
If Wfile = "PEDTN.RPT" Then
 If Trim(CADENITA) <> "" Then
   CADENITA = CADENITA + " AND {FACART.FAR_MONEDA} = '" & Left(moneda.Text, 1) & "'"
 Else
   CADENITA = "{FACART.FAR_MONEDA} = '" & Left(moneda.Text, 1) & "'"
 End If
End If
If Wfile = "CONSOLIDADO" Then
 If Trim(CADENITA) <> "" Then
   CADENITA = CADENITA + " AND {ARTI.ART_MONEDA} = '" & Left(moneda.Text, 1) & "'"
 Else
   CADENITA = "{ARTI.ART_MONEDA} = '" & Left(moneda.Text, 1) & "'"
 End If
 If chesub.Value = 1 Then
  If Trim(CADENITA) <> "" Then
    CADENITA = CADENITA + " AND {ARTI.ART_ESTADO} = '" & chesub.Value & "'"
  Else
    CADENITA = "{ARTI.ART_ESTADO} = '" & chesub.Value & "'" ' PARA FILTAR SOLO SUB ALMACEN
  End If
 End If
 
End If

DoEvents
' agrege verificar muy importante ica
If LK_EMP = "HER" And LK_ICA = "A" Then
    Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hguia.rpt"
    wformula1 = "FECHA_EMI=  '" & wfecha & "'"
Else

wformula1 = "FECHA=  '" & wfecha & "'"
wformula2 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
wformula3 = "ZONAVEN=  '" & wfiltra & "'"
wformula4 = "DOCU=  '" & wdocuB & WDOCUF & "'"
End If

pbValue = pbValue + 1
Reportes.Formulas(0) = wformula1
Reportes.Formulas(1) = wformula2
Reportes.Formulas(2) = wformula3
Reportes.Formulas(3) = wformula4

'CADENITA = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_CODCIA} = '01' AND {FACART.FAR_FECHA} >= Date ( 2000,1,13) AND {FACART.FAR_FECHA} <= Date ( 2000,1,13) AND ({FACART.FAR_FBG} in ['F']  AND {FACART.FAR_NUMSER} = '1' AND {FACART.FAR_NUMFAC} >= 100 AND {FACART.FAR_NUMFAC} <= 500 OR {FACART.FAR_FBG} in ['B'] AND {FACART.FAR_NUMSER} = '1' AND {FACART.FAR_NUMFAC} >= 100 AND {FACART.FAR_NUMFAC} <= 100 )"

Reportes.SelectionFormula = CADENITA
' Debug.Print CADENITA
Reportes.Action = 1
If LK_EMP <> "PIU" Then
 Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "HRESUVD.rpt"
 Reportes.Action = 1
 Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hresu4.rpt"
 Reportes.Action = 1
 
End If
If Wfile <> "PEDTN.RPT" Then
  'CONSO_CAMBIO
End If

If LK_EMP = "HER" And LK_ICA = "A" Then
   Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hguia1.rpt"
   wformula1 = "FECHA_EMI=  '" & LK_FECHA_DIA & "'"
   Reportes.Action = 1
   Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hguia2.rpt"
   wformula1 = "FECHA_EMI=  '" & LK_FECHA_DIA & "'"
   Reportes.Action = 1
End If
pbValue = pbValue + 1
pb.Visible = False
lblmensa.Visible = False
cmdmostrar.Enabled = True
Exit Sub
' PROCEDIMIENTO
VENTASXCLI:
SQ_OPER = 1
pu_cp = loc_cp
pu_codcia = LK_CODCIA
pu_codclie = Val(txt_cli.Text)
LEER_CLI_LLAVE
If cli_llave.EOF Then
  lblCliente.Caption = ""
  MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
  Azul txt_cli, txt_cli
  Exit Sub
Else
  lblCliente.Caption = Trim(cli_llave!cli_nombre)
  wscliente = Trim(cli_llave!cli_nombre)
End If
cmdmostrar.Enabled = False
If Right(fechaini.Text, 2) = "__" Then
  wsFECHA1 = Left(fechaini.Text, 8)
Else
 wsFECHA1 = Trim(fechaini.Text)
End If
If Not IsDate(wsFECHA1) Then
    cmdmostrar.Enabled = True
    MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
    Exit Sub
End If
If Right(fechafin.Text, 2) = "__" Then
  wsFECHA2 = Left(fechafin.Text, 8)
Else
  wsFECHA2 = Trim(fechafin.Text)
End If
If Not IsDate(wsFECHA2) Then
 cmdmostrar.Enabled = True
 MsgBox "Fecha Invalidad ..", 48, pu_titulo
 Exit Sub
End If
lblmensa.Visible = True
DoEvents
pbMin = 0
pbMax = 6
pbValue = 0
pb.Visible = True
If wsFECHA1 = wsFECHA2 Then
 wfecha = "VENTAS X PRODUCTO - DEL DIA " & wsFECHA1
Else
 wfecha = "VENTAS X PRODUCTO - DEL DIA  " & wsFECHA1 & " AL " & wsFECHA2
End If
pbValue = pbValue + 1
DIA = Day(wsFECHA1)
mes = Month(wsFECHA1)
ano = Year(wsFECHA1)
pbValue = pbValue + 1
DIA1 = Day(wsFECHA2)
MES1 = Month(wsFECHA2)
ANO1 = Year(wsFECHA2)
pbValue = pbValue + 1
CADENITA = ""
CADEFECHA = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 10 AND {PRECIOS.PRE_FLAG_UNIDAD} = 'A' AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FECHA} >= Date ( " & ano & "," & mes & "," & DIA & ") AND {FACART.FAR_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
CADENITA = CADENITA + CADEFECHA + " AND {FACART.FAR_CP} = 'C' AND {FACART.FAR_CODCLIE} = " & pu_codclie
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.Formulas(3) = ""
pbValue = pbValue + 1
If retra_llave!tra_rep1 = "1" Then
  If LK_EMP_PTO = "A" Then
    Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\PTOVTA\" & "hclicom.rpt"
    wcodcia = "00"
  Else
    Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hclicom.rpt"
    wcodcia = LK_CODCIA
  End If
Else
   Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hclicom.rpt"
   wcodcia = LK_CODCIA
End If

DoEvents
wformula1 = "FECHA=  '" & wfecha & "'"
wformula2 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
wformula3 = "CLIENTE=  'CLIENTE : " & wscliente & "'"
pbValue = pbValue + 1
Reportes.Formulas(0) = wformula1
Reportes.Formulas(1) = wformula2
Reportes.Formulas(2) = wformula3

Reportes.SelectionFormula = CADENITA

Reportes.Action = 1
lblCliente.Caption = wscliente
pbValue = pbValue + 1
pb.Visible = False
lblmensa.Visible = False
cmdmostrar.Enabled = True



Exit Sub
procancela:
MsgBox Err.Description, 48, Pub_Titulo
cmdmostrar.Enabled = True
Resume Next
'Unload FrmRepo
End Sub

Private Sub CmdProcesa_Click()
MsgBox "PROCESO ESTA EN CONSTRUCCION ....", 48, Pub_Titulo
Exit Sub
Dim LETRAS(52) As String * 2
Dim Modo As String * 1
Dim wclave As Integer
Dim wclave2 As Integer
Dim f1 As Integer
Dim c1 As Integer
Dim clave_far As String
Dim suma_cantid As Currency
Dim wequiv As Currency
Dim wunidad As String
CmdProcesa.Enabled = False
opciones.Enabled = False
DoEvents
If Left(CmdProcesa.Caption, 2) = "De" Then
  LOC_DETENER = 1
  CmdProcesa.Enabled = False
  CmdProcesa.Caption = "un momento..."
  DoEvents
  Exit Sub
End If
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
     ENTRADA = WCOD_ORIGINAL
Else
     ENTRADA = Val(Txt_key.Text)
End If
lblmensa.Visible = True
DoEvents
lblmensa.Caption = "Activando Reporte . . . "
DoEvents
Modo = ""
ws_clave = "0"
usu.Requery
Do Until usu.EOF
  If Trim(usu!usu_key) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop

If option1(0).Value Then
   If Trim(Txt_key.Text) = "" Then
     MsgBox "Seleccione Codigo, para Procesar ", 48, Pub_Titulo
     GoTo fin
   End If
   pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ?  AND ART_KEY = ?  ORDER BY ART_NOMBRE"
   Modo = "1"
ElseIf option1(1).Value Then
  If lineas.ListCount = 0 And listapro.ListCount = 0 Then
    MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
    GoTo fin
  End If
  If Trim(lineas.Text) = "" And Trim(listapro.Text) = "" Then
    MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
    GoTo fin
  End If
  If Trim(listapro.Text) = "" And Trim(lineas.Text) <> "" Then
     pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ?  AND ART_FAMILIA = ?  ORDER BY ART_FAMILIA, ART_SUBFAM, ART_NOMBRE"
  ElseIf Trim(listapro.Text) <> "" And Trim(lineas.Text) = "" Then
     pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ?  AND ART_SUBFAM = ?  ORDER BY ART_FAMILIA, ART_SUBFAM, ART_NOMBRE"
  Else
    pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ?  AND ART_FAMILIA = ? AND ART_SUBFAM = ? ORDER BY ART_FAMILIA, ART_SUBFAM, ART_NOMBRE"
  End If
  Modo = "2"
ElseIf option1(2).Value Then
   If Trim(PROV.Text) = "" Then
     MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
     GoTo fin
   End If
   Wche = 0
   wkSELECT = ""
   For I = 0 To PROV.ListCount - 1
     PROV.ListIndex = I
     If PROV.Selected(I) Then
       Wche = 1
       wkSELECT = wkSELECT + " ART_CODCLIE = " & Trim(Right(PROV.Text, 10)) & " OR"
      End If
   Next I
   If Wche = 0 Then
     MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
     GoTo fin
   End If
   wkSELECT = Left(wkSELECT, Len(wkSELECT) - 3)
   pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND " & wkSELECT & " ORDER BY ART_CODCLIE, ART_NOMBRE"
   'PSART_LOC(1) = Trim(Right(PROV.text, 10))
   Modo = "3"
ElseIf option1(3).Value Then
   pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ?  ORDER BY  ART_NOMBRE"
   Modo = "4"
ElseIf option1(4).Value Then
   clave_far = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA = ? AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E'  ORDER BY FAR_CODART"
   CLAVE_ART = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY = ? ORDER BY  ART_NOMBRE"
   Modo = "5"
   GoSub VENTAS_DIA
  GoTo escapar
End If

Set PSART_LOC = CN.CreateQuery("", pub_cadena)
Set artloc_llave = PSART_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)

If Modo = "1" Then
   PSART_LOC(0) = LK_CODCIA
   PSART_LOC(1) = ENTRADA
ElseIf Modo = "2" Then
    If Trim(listapro.Text) = "" And Trim(lineas.Text) <> "" Then
       PSART_LOC(0) = LK_CODCIA
       PSART_LOC(1) = Right(lineas.Text, 10)
    ElseIf Trim(listapro.Text) <> "" And Trim(lineas.Text) = "" Then
       PSART_LOC(0) = LK_CODCIA
       PSART_LOC(1) = Right(listapro.Text, 10)
    Else
       PSART_LOC(0) = LK_CODCIA
       PSART_LOC(1) = Right(lineas.Text, 10)
       PSART_LOC(2) = Right(listapro.Text, 10)
    End If
ElseIf Modo = "3" Then
   PSART_LOC(0) = LK_CODCIA
   'PSART_LOC(1) = Right(PROV.text, 10)
ElseIf Modo = "4" Then
   PSART_LOC(0) = LK_CODCIA
End If
  pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E' ORDER BY FAR_FECHA DESC"
  Set PSFAR_VENTAS = CN.CreateQuery("", pub_cadena)
  Set far_ventas = PSFAR_VENTAS.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  ws_intervalo = Trim(Right(cmbTipo.Text, 4))
  artloc_llave.Requery
  If artloc_llave.EOF Then
    MsgBox "No existen Articulos en la Lista", 48, Pub_Titulo
    GoTo fin
  End If
  GoSub WEXCEL
  xl.Worksheets(1).Activate
  DoEvents
  lblmensa.Caption = "Procesando . . . "
  DoEvents
  tt_fecha_d = LK_FECHA_DIA
  tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_d) + 1
  If ws_intervalo = "d" Then
    For I = 1 To Val(txtMaxCol.Text)
      tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_d) + 1
      xl.Cells(4, I + 6) = "'" & Format(tt_fecha_d, "dd - mm")
      tt_fecha_d = tt_fecha_h - 1
    Next I
  ElseIf ws_intervalo = "ww" Then
      'MAS ABAJO ESTA
  ElseIf ws_intervalo = "m" Then
    For I = 1 To Val(txtMaxCol.Text)
      tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_d) + 1
      xl.Cells(4, I + 6) = "'" & UCase(Format(tt_fecha_d, "mmm - yy"))
      tt_fecha_d = tt_fecha_h - 1
    Next I
  ElseIf ws_intervalo = "yyyy" Then
   For I = 1 To Val(txtMaxCol.Text)
      tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_d) + 1
      xl.Cells(4, I + 6) = "' Año  " & UCase(Format(tt_fecha_d, "yyyy"))
      tt_fecha_d = tt_fecha_h - 1
    Next I
  End If
 ' AQUI PARA EL SELECT
 If ws_intervalo = "m" Then
       w_mes = "01/" & Format(LK_FECHA_DIA, "mm") & "/" & Format(LK_FECHA_DIA, "yyyy")
       tt_fecha_h = w_mes
       fecha_final = tt_fecha_h
       For I = 1 To Val(txtMaxCol.Text) - 1
        tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_h)
        fecha_final = tt_fecha_h
       Next I
        PSFAR_VENTAS(2) = fecha_final
  ElseIf ws_intervalo = "ww" Then
       WSFECHA = LK_FECHA_DIA
       For I = 1 To 7
         wdias = Format(WSFECHA, "w")
         If wdias = 1 Then
           Exit For
         End If
         WSFECHA = DateAdd("d", -1, WSFECHA)
       Next I
       tt_fecha_h = WSFECHA
       fecha_final = tt_fecha_h
       For I = 1 To Val(txtMaxCol.Text) - 1
        tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_h)
        fecha_final = tt_fecha_h
       Next I
       PSFAR_VENTAS(2) = fecha_final
       ' CABECARA
       tt_fecha_h = WSFECHA
       xl.Cells(4, 7) = "'" & Format(tt_fecha_h, "dd.mm") & " al " & Format(LK_FECHA_DIA, "dd.mm")
       For I = 1 To Val(txtMaxCol.Text) - 1
         tt_fecha_d = tt_fecha_h - 1
         tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_h)
         xl.Cells(4, I + 7) = "'" & Format(tt_fecha_h, "dd.mm") & " al " & Format(tt_fecha_d, "dd.mm")
       Next I
  ElseIf ws_intervalo = "d" Then
       tt_fecha_h = DateAdd(ws_intervalo, (Val(txtMaxCol.Text) * -1) + 1, LK_FECHA_DIA)
       fecha_final = tt_fecha_h
       PSFAR_VENTAS(2) = fecha_final
  ElseIf ws_intervalo = "yyyy" Then
       w_mes = "01/01" & "/" & Format(LK_FECHA_DIA, "yyyy")
       tt_fecha_h = w_mes
       fecha_final = tt_fecha_h
       For I = 1 To Val(txtMaxCol.Text) - 1
        tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_h)
        fecha_final = tt_fecha_h
       Next I
       PSFAR_VENTAS(2) = fecha_final
  End If
  If Left(CmdProcesa.Caption, 2) = "&L" Then
    CmdProcesa.Caption = "Detener..."
    CmdProcesa.Enabled = True
    DoEvents
    LOC_DETENER = 0
  End If
  GoSub LETRAS
  PSFAR_VENTAS(0) = LK_CODCIA
  PSFAR_VENTAS(3) = LK_FECHA_DIA
  FILA_INICIAL = 5
  fila = 0
  pbMax = artloc_llave.RowCount
  pbMin = 0
  pbValue = 0
  pb.Visible = True
  DoEvents
  f1 = FILA_INICIAL - 1
  wcodclie = Nulo_Valor0(artloc_llave!art_codclie)
  If Modo = "2" Then
    If Trim(Left(lineas.Text, 40)) <> "" And Trim(Left(listapro.Text, 40)) = "" Then
        xl.Cells(3, 1) = Trim(Left(lineas.Text, 40))
    ElseIf Trim(Left(lineas.Text, 40)) = "" And Trim(Left(listapro.Text, 40)) <> "" Then
        xl.Cells(3, 1) = Trim(Left(listapro.Text, 40))
    Else
       xl.Cells(3, 1) = Trim(Left(lineas.Text, 40)) & " & " & Trim(Left(listapro.Text, 40))
    End If
  End If
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  pu_cp = "P"
  pu_codclie = wcodclie
  LEER_CLI_LLAVE
  If cli_llave.EOF Then
    ' MsgBox "Codigo de Proveedor NO EXISTE.., Consulte al Administrador.", 48, Pub_Titulo
  Else
     f1 = f1 + 1
     xl.Cells(f1, 2) = Trim(cli_llave!cli_nombre)
     xl.Cells(f1, 2).Font.Bold = True
  End If
  Do Until artloc_llave.EOF
    pbValue = pbValue + 1
    If LOC_DETENER = 1 Then
       GoTo escapar
    End If
    If Val(txtMaxCol.Text) <> 0 Then
     PSFAR_VENTAS(1) = artloc_llave!art_key
     far_ventas.Requery
     DoEvents
     If far_ventas.EOF Then
       If Modo = "1" Then
         MsgBox "No Exite Ventas . . ", 48, Pub_Titulo
         GoTo fin
       End If
       GoTo SALTO
     End If
    End If
    fila = fila + 1
    f1 = f1 + 1
    If wcodclie <> artloc_llave!art_codclie Then
       SQ_OPER = 1
       pu_codcia = LK_CODCIA
       pu_cp = "P"
       pu_codclie = artloc_llave!art_codclie
       LEER_CLI_LLAVE
       If cli_llave.EOF Then
          MsgBox "Codigo de Proveedor NO EXISTE.., Consulte al Administrador.", 48, Pub_Titulo
       Else
          xl.Cells(f1, 2) = Trim(cli_llave!cli_nombre)
          xl.Cells(f1, 2).Font.Bold = True
       End If
       wcodclie = artloc_llave!art_codclie
       f1 = f1 + 1
    End If
    xl.Cells(f1, 1) = fila
    xl.Cells(f1, 2) = Trim(artloc_llave!art_nombre)
    lblmensa.Caption = "Procesando . . .  " & Trim(artloc_llave!art_nombre)
    DoEvents
    SQ_OPER = 1
    PUB_CODART = artloc_llave!art_key
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    wequiv = 0
    wunidad = ""
    If arm_llave.EOF Then
       MsgBox "El Producto: " & Trim(artloc_llave!art_nombre) & " NO titne Registro en Tabla Articulo .. consultar con el Administrador", 48, Pub_Titulo
    Else
       PRE_UNIDADES(0) = LK_CODCIA
       PRE_UNIDADES(1) = artloc_llave!art_key
       PRE_UNIDADES(2) = "A"
       preuni_llave.Requery
       If preuni_llave.EOF Then
         MsgBox "El Producto: " & Trim(artloc_llave!art_nombre) & " NO titne Registro en Unidades .. consultar con el Administrador", 48, Pub_Titulo
         wequiv = 1
       Else
         wequiv = preuni_llave!PRE_EQUIV
         wunidad = Left(preuni_llave!PRE_UNIDAD, 12)
       End If
       xl.Cells(f1, 3) = Trim(wunidad)
       xl.Cells(f1, 4) = Format(arm_llave!arm_stock / wequiv, "0.00")
       xl.Cells(f1, 5) = Format(arm_llave!ARM_COSPRO * wequiv, "0.00")
       xl.Cells(f1, 6) = Format(arm_llave!arm_stock * arm_llave!ARM_COSPRO, "0.00")
    End If
    If Val(txtMaxCol.Text) <= 0 Then
       DoEvents
      GoTo SALTO
    End If
   
    tt_fecha_d = LK_FECHA_DIA
    If ws_intervalo = "ww" Then
       tt_fecha_h = WSFECHA
    ElseIf ws_intervalo = "m" Then
       tt_fecha_h = "01/" & Format(LK_FECHA_DIA, "mm") & "/" & Format(LK_FECHA_DIA, "yyyy")
    ElseIf ws_intervalo = "d" Then
       tt_fecha_h = LK_FECHA_DIA
    Else
       tt_fecha_h = "01/01" & "/" & Format(LK_FECHA_DIA, "yyyy")
    End If
    c1 = 6
    tot_col = 0
    WS_CANTIDAD = 0
    ws_flag = 0
    suma_cantid = 0
    far_ventas.MoveFirst
    WWFECHA = far_ventas!far_fecha
    Do Until far_ventas.EOF
      If far_ventas!far_fecha < CDate(tt_fecha_h) Then
       If ws_intervalo = "d" Or ws_intervalo = "ww" Then
        If WWFECHA <> tt_fecha_h Then
          If ws_intervalo = "m" Or ws_intervalo = "ww" Or ws_intervalo = "d" Then
             tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_h)
          Else
             tt_fecha_h = DateAdd(ws_intervalo, -1, far_ventas!far_fecha) + 1
          End If
          c1 = c1 + 1
          GoTo SIGUE_LOOP
        End If
       End If
       If ws_intervalo = "m" Then
         If WS_CANTIDAD = 0 Then
             If ws_intervalo = "m" Or ws_intervalo = "ww" Or ws_intervalo = "d" Then
               tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_h)
             Else
               tt_fecha_h = DateAdd(ws_intervalo, -1, far_ventas!far_fecha) + 1
             End If
             c1 = c1 + 1
             GoTo SIGUE_LOOP
         End If
       End If
         ws_flag = 1
         GoSub IMPRIME_CANT
         tot_col = c1 - 5
         If ws_intervalo = "m" Then
          If tot_col = Val(txtMaxCol.Text) + 1 Then
            GoTo SALTO
          End If
         Else
          If tot_col = Val(txtMaxCol.Text) Then
            GoTo SALTO
          End If
         End If
         If ws_intervalo = "m" Or ws_intervalo = "ww" Or ws_intervalo = "d" Then
           tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_h)
         Else
           tt_fecha_h = DateAdd(ws_intervalo, -1, far_ventas!far_fecha) + 1
         End If
         WS_CANTIDAD = 0
         GoSub ACU_CANT
         WWFECHA = far_ventas!far_fecha
      Else
          GoSub ACU_CANT
          WWFECHA = far_ventas!far_fecha
      End If
      If far_ventas!far_fecha < CDate(fecha_final) Then
        GoTo sal
      End If
      wwfecha_FIN = far_ventas!far_fecha
      far_ventas.MoveNext
SIGUE_LOOP:
    Loop
sal:
   If ws_flag = 0 Then
   If ws_intervalo = "d" Or ws_intervalo = "ww" Or ws_intervalo = "m" Then
     Do Until wwfecha_FIN >= tt_fecha_h
      If WWFECHA <> tt_fecha_h Then 'Or WS_CANTIDAD = 0 Then
       'If WS_CANTIDAD = 0 Then
         If ws_intervalo = "m" Or ws_intervalo = "ww" Or ws_intervalo = "d" Then
            tt_fecha_h = DateAdd(ws_intervalo, -1, tt_fecha_h)
         Else
            tt_fecha_h = DateAdd(ws_intervalo, -1, far_ventas!far_fecha) + 1
         End If
         c1 = c1 + 1
       End If
     Loop
   End If
      GoSub IMPRIME_CANT
   End If
SALTO:
   artloc_llave.MoveNext
  Loop
escapar:
  lblmensa.Caption = "Mostrando Reportes . . . "
  DoEvents
  xl.Cells(f1 + 1, 2) = "TOTAL .-"
  
  wmax = 5
  WINI = 5
  If Modo = "5" Then
    wmax = 8
    WINI = 4
  End If
  wranF = "A" & f1 + 1 & ":" & Trim(LETRAS(wmax + 5)) & f1 + 1
  xl.Range(wranF).Font.Bold = True
  If f1 <> 4 Then
  For I = WINI To wmax
    wranF = Trim(LETRAS(I)) & f1 + 1
    wran1 = Trim(LETRAS(I)) & FILA_INICIAL
    wran2 = Trim(LETRAS(I)) & f1
    xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  Next I
  End If
  If Modo = "5" Then
     xl.Worksheets("Hoja2").Protect ws_clave
  Else
     xl.Worksheets("Hoja1").Protect ws_clave
  End If
  LOC_DETENER = 0
  CmdProcesa.Caption = "&Listar Articulo (s)"
  DoEvents
  CmdProcesa.Enabled = True
  DoEvents
  opciones.Enabled = True
  xl.DisplayAlerts = False
  xl.Application.Visible = True

Set xl = Nothing
pb.Visible = False
lblmensa.Visible = False
 
Exit Sub

IMPRIME_CANT:
   c1 = c1 + 1
   If WS_CANTIDAD <> 0 Then xl.Cells(f1, c1) = Format(WS_CANTIDAD, "0.00")
Return

ACU_CANT:
  WS_CANTIDAD = WS_CANTIDAD + (far_ventas!FAR_cantidad / wequiv)
  ws_flag = 0
Return


WEXCEL:
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  lblmensa.Caption = "Abriendo , Archivo VENTAS_VARIOS.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\VENTAS_VARIOS.xls", 0, True, 4, WPAS, WPAS
Return

LETRAS:
 For I = 1 To 26
  LETRAS(I) = Chr(64 + I)
 Next I
 For I = 27 To 52
  LETRAS(I) = Trim(LETRAS(1)) & Chr(38 + I)
 Next I

Return
'  SOLO VENTAS DEL DIA
VENTAS_DIA:
  pub_cadena = CLAVE_ART
  Set PSART_LOC = CN.CreateQuery("", pub_cadena)
  Set artloc_llave = PSART_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  pub_cadena = clave_far
  Set PSFAR_VENTAS = CN.CreateQuery("", pub_cadena)
  Set far_ventas = PSFAR_VENTAS.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  ws_intervalo = "d"
  PSFAR_VENTAS(0) = LK_CODCIA
  PSFAR_VENTAS(1) = LK_FECHA_DIA
  far_ventas.Requery
  If far_ventas.EOF Then
    MsgBox "No existen Articulos en la Lista", 48, Pub_Titulo
    GoTo fin
  End If
  GoSub WEXCEL
  xl.Worksheets(2).Activate
  DoEvents
  lblmensa.Caption = "Procesando . . . "
  DoEvents
  GoSub LETRAS
  FILA_INICIAL = 5
  fila = 0
  pbMax = far_ventas.RowCount
  pbMin = 0
  pbValue = 0
  pb.Visible = True
  DoEvents
  f1 = FILA_INICIAL - 1
  WS_CANTIDAD = 0
  WS_BRUTO = 0
  WS_CREDITO = 0
  WS_Contado = 0
  WS_COSTO = 0
  WSCODIGO = far_ventas!far_codart
  PRE_UNIDADES(0) = LK_CODCIA
  PRE_UNIDADES(1) = WSCODIGO
  PRE_UNIDADES(2) = "A"
  preuni_llave.Requery
  If preuni_llave.EOF Then
     MsgBox "El Producto: " & Trim(artloc_llave!art_nombre) & " NO titne Registro en Unidades .. consultar con el Administrador", 48, Pub_Titulo
  Else
     wequiv = preuni_llave!PRE_EQUIV
     wunidad = Left(preuni_llave!PRE_UNIDAD, 12)
  End If
  xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
  Do Until far_ventas.EOF
    pbValue = pbValue + 1
    If WSCODIGO <> far_ventas!far_codart Then
        fila = fila + 1
        f1 = f1 + 1
        PSART_LOC(0) = LK_CODCIA
        PSART_LOC(1) = WSCODIGO
        artloc_llave.Requery
        If artloc_llave.EOF Then
          MsgBox "Codigo de articulo de FACART no existe en Articulo, consulte al Adminisrador", 48, pub_tiulo
          GoTo fin
        End If
        xl.Cells(f1, 1) = fila
        xl.Cells(f1, 2) = Trim(artloc_llave!art_nombre)
        lblmensa.Caption = "Procesando . . .  " & Trim(artloc_llave!art_nombre)
        DoEvents
        SQ_OPER = 1
        PUB_CODART = artloc_llave!art_key
        pu_codcia = LK_CODCIA
        LEER_ARM_LLAVE
        If arm_llave.EOF Then
           MsgBox "El Producto: " & Trim(artloc_llave!art_nombre) & " NO titne Registro en Tabla Articulo .. consultar con el Administrador", 48, Pub_Titulo
        End If
        xl.Cells(f1, 3) = wunidad
        xl.Cells(f1, 4) = Format(WS_CANTIDAD, "0.00")
        xl.Cells(f1, 5) = Format(WS_BRUTO, "0.00")
        xl.Cells(f1, 6) = Format(WS_CREDITO, "0.00")
        xl.Cells(f1, 7) = Format(WS_Contado, "0.00")
        xl.Cells(f1, 8) = WS_COSTO
        xl.Cells(f1, 9) = WS_BRUTO - WS_COSTO
        If WS_BRUTO <> 0 Then
          xl.Cells(f1, 10) = Format(((WS_BRUTO - WS_COSTO) * 100) / WS_BRUTO, "0.00")
        End If
        WSCODIGO = far_ventas!far_codart
        PRE_UNIDADES(0) = LK_CODCIA
        PRE_UNIDADES(1) = WSCODIGO
        PRE_UNIDADES(2) = "A"
        preuni_llave.Requery
        If preuni_llave.EOF Then
           MsgBox "El Producto: " & Trim(artloc_llave!art_nombre) & " NO titne Registro en Unidades .. consultar con el Administrador", 48, Pub_Titulo
        Else
           wequiv = preuni_llave!PRE_EQUIV
           wunidad = Left(preuni_llave!PRE_UNIDAD, 12)
        End If
        WS_CANTIDAD = 0
        WS_BRUTO = 0
        WS_CREDITO = 0
        WS_Contado = 0
        WS_COSTO = 0
    End If
    WS_CANTIDAD = WS_CANTIDAD + (far_ventas!FAR_cantidad / wequiv)
    WS_valor = (far_ventas!FAR_cantidad / far_ventas!FAR_equiv) * far_ventas!FAR_PRECIO
    If far_ventas!FAR_DESCTO <> 0 Then
       WS_valor = 0
    End If
    If Val(Nulo_Valor0(far_ventas!far_signo_car)) = 0 Then
          WS_Contado = WS_Contado + WS_valor
    Else
          WS_CREDITO = WS_CREDITO + WS_valor
    End If
    WS_BRUTO = WS_BRUTO + WS_valor
    WS_COSTO = WS_COSTO + (far_ventas!FAR_cantidad * far_ventas!FAR_COSPRO)
    far_ventas.MoveNext
  Loop
    fila = fila + 1
    f1 = f1 + 1
    PSART_LOC(0) = LK_CODCIA
    PSART_LOC(1) = WSCODIGO
    artloc_llave.Requery
    If artloc_llave.EOF Then
      MsgBox "Codigo de articulo de FACART no existe en Articulo, consulte al Adminisrador", 48, pub_tiulo
      GoTo fin
    End If
    PRE_UNIDADES(0) = LK_CODCIA
    PRE_UNIDADES(1) = WSCODIGO
    PRE_UNIDADES(2) = "A"
    preuni_llave.Requery
    If preuni_llave.EOF Then
        MsgBox "El Producto: " & Trim(artloc_llave!art_nombre) & " NO titne Registro en Unidades .. consultar con el Administrador", 48, Pub_Titulo
    Else
        wequiv = preuni_llave!PRE_EQUIV
        wunidad = Left(preuni_llave!PRE_UNIDAD, 12)
    End If
    xl.Cells(f1, 1) = fila
    xl.Cells(f1, 2) = Trim(artloc_llave!art_nombre)
    lblmensa.Caption = "Procesando . . .  " & Trim(artloc_llave!art_nombre)
    DoEvents
    SQ_OPER = 1
    PUB_CODART = artloc_llave!art_key
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    If arm_llave.EOF Then
       MsgBox "El Producto: " & Trim(artloc_llave!art_nombre) & " NO titne Registro en Tabla Articulo .. consultar con el Administrador", 48, Pub_Titulo
    End If
    xl.Cells(f1, 3) = wunidad
    xl.Cells(f1, 4) = Format(WS_CANTIDAD, "0.00")
    xl.Cells(f1, 5) = Format(WS_BRUTO, "0.00")
    xl.Cells(f1, 6) = Format(WS_CREDITO, "0.00")
    xl.Cells(f1, 7) = Format(WS_Contado, "0.00")
    xl.Cells(f1, 8) = WS_COSTO
    xl.Cells(f1, 9) = WS_BRUTO - WS_COSTO
    If WS_BRUTO <> 0 Then
          xl.Cells(f1, 10) = Format(((WS_BRUTO - WS_COSTO) * 100) / WS_BRUTO, "0.00")
    End If
Return
' HASTA AQUI

GoTo fin

Exit Sub
fin:
'Dim dd As Excel.Application
If Not xl Is Nothing Then
  xl.Application.Visible = True
  Set xl = Nothing
End If

pb.Visible = False
lblmensa.Visible = False
LOC_DETENER = 0
CmdProcesa.Caption = "&Listar Articulo (s)"
DoEvents
CmdProcesa.Enabled = True
DoEvents
opciones.Enabled = True

End Sub


Private Sub cmdrela_Click()
Dim WPLA As String * 1
WPLA = ""
If Wfile = "PLANILLA" Then
 WPLA = "A"
End If

Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(retra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390

CADENITA = ""
wkSELECT = ""
Wche = 0
For fila = 0 To docu.ListCount - 1
  docu.ListIndex = fila
  WNUMSER = Str(Val(Mid(docu.Text, 4, 3)))
  wnumfac = Str(Val(Mid(docu.Text, 10, 16)))
  WFBG = Left(docu.Text, 1)
  If docu.Selected(fila) And Trim(listven.Text) <> "" Then
    If WPLA = "A" Then
      If Wche = 0 Then
       wkSELECT = "({CARTERA.CAR_FBG} = '" & Trim(WFBG) & "' AND {CARTERA.CAR_NUMSER} = " & Trim(WNUMSER) & " AND {CARTERA.CAR_NUMFAC} = " & wnumfac & ") "
      Else
       wkSELECT = wkSELECT + " OR ({CARTERA.CAR_FBG} = '" & Trim(WFBG) & "' AND {CARTERA.CAR_NUMSER} = " & Trim(WNUMSER) & " AND {CARTERA.CAR_NUMFAC} = " & wnumfac & ") "
      End If
     Else
      If Wche = 0 Then
       wkSELECT = "({FACART.FAR_FBG} = '" & Trim(WFBG) & "' AND {FACART.FAR_NUMSER} = '" & Trim(WNUMSER) & "' AND {FACART.FAR_NUMFAC} = " & wnumfac & ") "
      Else
       wkSELECT = wkSELECT + " OR ({FACART.FAR_FBG} = '" & Trim(WFBG) & "' AND {FACART.FAR_NUMSER} = '" & Trim(WNUMSER) & "' AND {FACART.FAR_NUMFAC} = " & wnumfac & ") "
      End If
     End If
    Wche = 1
  End If
  docu.ListIndex = fila
Next fila
If WPLA = "A" Then
 If wkSELECT = "" Then
   wkSELECT = "({CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' )"
 Else
   wkSELECT = "(" + wkSELECT + ") AND ( {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' )"
 End If
Else
 If wkSELECT = "" Then
  wkSELECT = "( {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_ESTADO} <> 'E')"
 Else
  wkSELECT = "(" + wkSELECT + ") AND ( {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_ESTADO} <> 'E')"
 End If
End If
CADENITA = wkSELECT
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.Formulas(3) = ""
If retra_llave!tra_rep1 = "1" Then
  If LK_EMP_PTO = "A" Then
    Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\PTOVTA\" & "hresu.rpt"
    wcodcia = "00"
  Else
    Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hresu.rpt"
    wcodcia = LK_CODCIA
  End If
Else
   Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hresu.rpt"
   wcodcia = LK_CODCIA
End If
If WPLA = "A" Then
  Reportes.ReportFileName = PUB_RUTA_OTRO & "PLANILLA.rpt"
  wcodcia = LK_CODCIA
End If
If Wfile = "CONSOLIDADO" Then
 If Trim(CADENITA) <> "" Then
   CADENITA = CADENITA + " AND {ARTI.ART_MONEDA} = '" & Left(moneda.Text, 1) & "'"
 Else
   CADENITA = "{ARTI.ART_MONEDA} = '" & Left(moneda.Text, 1) & "'"
 End If
End If

DoEvents
If WPLA = "A" Then
 wfecha = LK_FECHA_DIA
 wformula1 = "DIA=  '" & wfecha & "'"
 wformula2 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
 Reportes.Formulas(0) = wformula1
 Reportes.Formulas(1) = wformula2
Else
 wformula1 = "FECHA=  '" & wfecha & "'"
 wformula2 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
 wformula3 = "ZONAVEN=  '" & wfiltra & "'"
 Reportes.Formulas(0) = wformula1
 Reportes.Formulas(1) = wformula2
 Reportes.Formulas(2) = wformula3

End If

Reportes.SelectionFormula = CADENITA
Reportes.Action = 1
If LK_EMP = "PAR" Then 'And Wfile = "CONSOLIDADO" Then
 Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hresu2.rpt"
 Reportes.Action = 1
End If

pb.Visible = False
lblmensa.Visible = False
cmdmostrar.Enabled = True
Exit Sub

End Sub

Private Sub docu_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 lisdocu.Visible = False
 fechaini.SetFocus
End If
End Sub

Private Sub docu_LostFocus()
' lisdocu.Visible = False
End Sub

Private Sub fac_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
 If Index = 0 Then
  fac(1).SetFocus
 End If
 If Index = 1 Then
  fac(2).SetFocus
 End If
 If Index = 2 Then
  chebol.SetFocus
 End If
 Exit Sub
End If
SOLO_ENTERO KeyAscii

End Sub

Private Sub fechafin_GotFocus()
fechafin.SetFocus
fechafin.SelStart = 0
fechafin.SelLength = Len(fechafin.Text)
FrmRepo.ListView1.Visible = False
End Sub

Private Sub fechafin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If listven.Visible Then
 listven.SetFocus
End If
If txt_cli.Visible Then
 txt_cli.SetFocus
End If
End If

End Sub

Private Sub fechaini_GotFocus()
fechaini.SetFocus
fechaini.SelStart = 0
fechaini.SelLength = Len(fechaini.Text)
FrmRepo.ListView1.Visible = False
End Sub

Private Sub fechaini_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 fechafin.SetFocus
End If
End Sub

Private Sub Form_Activate()
If fechaini.Visible And fechaini.Enabled Then
 fechaini.SetFocus
ElseIf Txt_key.Enabled And Txt_key.Visible Then
 Txt_key.SetFocus
End If
End Sub

Private Sub Form_Load()
Screen.MousePointer = 0
If retra_llave.EOF Then
 Exit Sub
End If
Wfile = Trim(retra_llave(3))
FrmRepo.Caption = Trim(retra_llave(1))
FrmRepo.txt_cli.Visible = False
FrmRepo.lblCliente.Visible = False
If Wfile = "ANALISIS" Then
  FrmRepo.fraconsi.Visible = False
  FrmRepo.fra2.Visible = True
  GoTo ANA
End If
If Wfile = "PLANILLA" Or Wfile = "CONSOLIDADO" Or Wfile = "CONTADO" Or Wfile = "PEDTN.RPT" Or Wfile = "HRESUVE" Then
  FrmRepo.franego.Visible = True
  FrmRepo.tipon.Visible = True
  PUB_CODCIA = LK_CODCIA
  LLENADOS tipon, 222
  FrmRepo.fraconsi.Visible = True
  FrmRepo.fra2.Visible = False
  If Wfile = "PEDTN.RPT" Then FrmRepo.frazonas.Visible = False
  FrmRepo.Lmoneda.Visible = True
  FrmRepo.moneda.Visible = True
  PUB_CODCIA = LK_CODCIA
  LLENADOS lstcanal, 65
  GoTo CONSO
End If
If Wfile = "IMP_FACTURACION" Then
  FrmRepo.fraconsi.Visible = True
  FrmRepo.fra2.Visible = False
  FrmRepo.frazonas.Visible = False
  FrmRepo.chefac.Value = 1
  GoTo FACTU
End If

If Wfile = "VENTASXCLI" Then
  FrmRepo.txt_cli.Visible = True
  FrmRepo.lblCliente.Visible = True
  FrmRepo.fraconsi.Visible = True
  FrmRepo.fra2.Visible = False
  FrmRepo.fraven.Visible = False
  FrmRepo.frazonas.Visible = False
  FrmRepo.fradocu.Visible = False
  GoTo CONSO
End If
If Wfile = "ORDEN_COMPRA" Then
  FrmRepo.txt_cli.Visible = True
  FrmRepo.lblCliente.Visible = True
  FrmRepo.fraconsi.Visible = True
  FrmRepo.fra2.Visible = False
  FrmRepo.fraven.Visible = False
  FrmRepo.frazonas.Visible = False
  FrmRepo.fradocu.Visible = False
  GoTo CONSO
End If


Exit Sub
ANA:
Dim fech As String
Dim wmes As String * 10
Dim codi
Dim VARX As String
Dim r As Integer
Dim WT As Integer
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ? AND PRE_FLAG_UNIDAD = ? ORDER BY PRE_CODART"
Set PRE_UNIDADES = CN.CreateQuery("", pub_cadena)
Set preuni_llave = PRE_UNIDADES.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP = 'P'  AND CLI_CODCIA = ? ORDER BY CLI_NOMBRE"
Set PSPRO_V = CN.CreateQuery("", pub_cadena)
Set PRO_V = PSPRO_V.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PSPRO_V(0) = LK_CODCIA
PRO_V.Requery
Do Until PRO_V.EOF
    PROV.AddItem PRO_V!cli_nombre & String(15, " ") & PRO_V!cli_codclie
    PRO_V.MoveNext
Loop
PROV.Visible = False
PUB_CODCIA = LK_CODCIA
LLENADOS lineas, 122
LLENADOS listapro, 123
ListView1.Visible = False
cmbTipo.AddItem "Dias                           d   "
cmbTipo.AddItem "Semanas                        ww  "
cmbTipo.AddItem "Meses                          m   "
cmbTipo.AddItem "Años                           yyyy"
txtMaxCol.Text = "'"
cmbTipo.ListIndex = 0
txtMaxCol.Text = "0"
Txt_key.Visible = True
lineas.Visible = False
listapro.Visible = False
PROV.Visible = False
Label1.Visible = True
Label1.Caption = "Codigo :"
Label2.Visible = False
Txt_key.MaxLength = 15
Exit Sub

CONSO:
If Wfile <> "VENTASXCLI" Then
 If par_llave!PAR_DEFAULT_FAC = "D" Then
  FrmRepo.moneda.ListIndex = 1
 Else
  FrmRepo.moneda.ListIndex = 0
 End If
  
 
 LLENA_VEND
 PUB_CODCIA = "00"
 LLENADOS listdistrito, 20
 LLENADOS listprovincia, 35
End If
LLENA_FECHAS
loc_cp = "C"
cheguia.Enabled = False
Exit Sub

FACTU:
LLENA_VEND
LLENA_FECHAS
Exit Sub

End Sub

Private Sub guia_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
 If Index = 0 Then
  guia(1).SetFocus
 End If
 If Index = 1 Then
  cmdmostrar.SetFocus
 End If
 Exit Sub
End If
SOLO_ENTERO KeyAscii

End Sub

Private Sub lineas_DblClick()
lineas_KeyPress 13
End Sub

Private Sub lineas_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
  Exit Sub
End If
CmdProcesa_Click
End Sub


Private Sub listapro_DblClick()
listapro_KeyPress 13
End Sub

Private Sub listapro_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Exit Sub
End If
CmdProcesa_Click
End Sub

Private Sub listdistrito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 listprovincia.SetFocus
End If

End Sub

Private Sub listprovincia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 cmdmostrar.SetFocus
End If

End Sub

Private Sub listven_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If frazonas.Visible Then
  listdistrito.SetFocus
 ElseIf fradocu.Visible Then
  chefac.SetFocus
 End If
End If

End Sub

Private Sub ListView1_DblClick()
 loc_key = ListView1.SelectedItem.Index
 txt_cli.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 txt_cli_KeyPress 13
End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.Item(loc_key).Selected = True
 ListView1.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Txt_key.Text = ""
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
txt_key_KeyPress 13
End Sub

Private Sub Option1_Click(Index As Integer)
txtMaxCol.Text = ""
If option1(0).Value = True Then
   lblopcion.Caption = ""
   Txt_key.Text = ""
   Txt_key.Visible = True
   lineas.Visible = False
   listapro.Visible = False
   PROV.Visible = False
   Label1.Visible = True
   Label1.Caption = "Codigo :"
   Label2.Visible = False
   lbltipo.Visible = True
   lblmax.Visible = True
   txtMaxCol.Visible = True
   cmbTipo.Visible = True
   Txt_key.SetFocus
ElseIf option1(1).Value = True Then
   lblopcion.Caption = ""
   Txt_key.Text = ""
   lineas.Visible = True
   listapro.Visible = True
   If listapro.ListCount > 0 Then
     listapro.ListIndex = -1
   End If
   If lineas.ListCount > 0 Then
      lineas.ListIndex = -1
   End If
   PROV.Visible = False
   Label1.Visible = True
   Label1.Caption = "Familia :"
   Label2.Visible = True
   lbltipo.Visible = True
   lblmax.Visible = True
   txtMaxCol.Visible = True
   cmbTipo.Visible = True
   lineas.SetFocus
ElseIf option1(2).Value = True Then
   For I = 0 To PROV.ListCount - 1
      PROV.Selected(I) = False
   Next I
   lblopcion.Caption = ""
   Txt_key.Text = ""
   lineas.Visible = False
   listapro.Visible = False
   PROV.Visible = True
   Label1.Caption = "Proveedor"
   Label1.Visible = True
   Label2.Visible = False
   Txt_key.Visible = False
   lbltipo.Visible = True
   lblmax.Visible = True
   txtMaxCol.Visible = True
   cmbTipo.Visible = True
   PROV.SetFocus
ElseIf option1(3).Value = True Then
   lineas.Visible = False
   listapro.Visible = False
   PROV.Visible = False
   Label1.Visible = False
   Label2.Visible = False
   Txt_key.Visible = False
   lbltipo.Visible = True
   lblmax.Visible = True
   txtMaxCol.Visible = True
   cmbTipo.Visible = True
ElseIf option1(4).Value = True Then
   lbltipo.Visible = False
   lblmax.Visible = False
   txtMaxCol.Visible = False
   cmbTipo.Visible = False
   lineas.Visible = False
   listapro.Visible = False
   PROV.Visible = False
   Label1.Visible = False
   Label2.Visible = False
   Txt_key.Visible = False
End If

End Sub
Private Sub PROV_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Exit Sub
End If
CmdProcesa_Click
End Sub

Private Sub salir_Click()
 Unload FrmRepo
End Sub

Public Sub LLENADOS(cont As ListBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    If tip <> 65 Then
      cont.AddItem " "
    End If
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Private Sub Txt_key_Change()
If Txt_key.Text = "" Then
 lblopcion.Caption = ""
End If
End Sub

Private Sub txt_key_GotFocus()
 Txt_key.Text = ""
End Sub

Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
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
  Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text)
  DoEvents
  Txt_key.SelStart = Len(Txt_key.Text)
  DoEvents
fin:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
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
If LK_FLAG_ALTERNO <> "A" And LK_FLAG_ORIGINAL = "A" Then
 On Error GoTo ERROR_CODIGO
 pu_codclie = Val(Txt_key.Text)
 On Error GoTo 0
Else
 pu_alterno = Trim(Txt_key.Text)
End If
If Len(Txt_key.Text) = 0 Then
   Exit Sub
End If

If LK_FLAG_ALTERNO <> "A" And LK_FLAG_ORIGINAL = "A" And pu_codclie <> 0 And IsNumeric(Txt_key.Text) = True Then
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   On Error GoTo ERROR_CODIGO
   PUB_KEY = pu_codclie
   On Error GoTo 0
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
     lblopcion.Caption = ""
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Azul Txt_key, Txt_key
     GoTo fin
   Else
   lblopcion.Caption = Trim(art_LLAVE!art_nombre)
   Txt_key.Text = Trim(art_LLAVE!art_key)
   cmbTipo.SetFocus
   End If
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(Txt_key.Text)) = Left(VALOR, Len(Trim(Txt_key.Text))) Then
   Else
      Exit Sub
   End If
   If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
     lblopcion.Caption = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
     SQ_OPER = 3
     pu_alterno = Trim(ListView1.ListItems.Item(loc_key).Text)
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     WCOD_ORIGINAL = art_llave_alt!art_key
     Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text)
   Else
    lblopcion.Caption = Trim(ListView1.ListItems.Item(loc_key).Text)
    Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
   End If
   If cmbTipo.Visible And cmbTipo.Enabled Then cmbTipo.SetFocus
End If
dale:
ListView1.Visible = False
fin:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul Txt_key, Txt_key
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 If Len(Txt_key.Text) = 0 Or Txt_key.Text = "" Then
   ListView1.Visible = False
   Exit Sub
 End If
Else
 If Len(Txt_key.Text) = 0 Or IsNumeric(Txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
 End If
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(Txt_key) = 1 Then
    var = Asc(Txt_key.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND  (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_ALTERNO BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND  (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_NOMBRE BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
    End If
  '  If Len(txt_key.text) > 1 And ListView1.ListItems.count = 0 Then
  '  Else
     PROC_LISVIEW ListView1
  '  End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
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
   DoEvents
  End If
  Exit Sub
End If


End Sub


Private Sub txtMaxCol_GotFocus()
 Azul txtMaxCol, txtMaxCol
 temporal = txtMaxCol.Text
End Sub

Private Sub txtMaxCol_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 CmdProcesa.SetFocus
End If
End Sub

Private Sub txtMaxCol_LostFocus()
If IsNumeric(txtMaxCol.Text) Then
 If Val(txtMaxCol.Text) <= 100 Then
   Exit Sub
 End If
End If
txtMaxCol.Text = temporal

End Sub
Public Sub SOLO_ENTERO(Optional tecla)
'CONVIERTE TODA A MAYUSCULAS LETRAS
Dim car As String, Longt As Integer
car = Chr$(tecla)
car = UCase$(Chr$(tecla))
tecla = Asc(car)
If car < "0" Or car > "9" Then
    If tecla <> 8 And tecla <> 13 Then
        tecla = 0
        Beep
    End If
End If
End Sub

Private Sub txt_cli_GotFocus()
Azul txt_cli, txt_cli
lblCliente.Caption = ""
End Sub
Private Sub txt_cli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_cli.Text = "" Then
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
'  KeyCode = 0
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  txt_cli.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  txt_cli.SelStart = Len(txt_cli.Text)
  DoEvents
fin:

End Sub
Private Sub txt_cli_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If KeyAscii = 27 Then
 ListView1.Visible = False
 txt_cli.Text = ""
 Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
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
   End If
   If cmdmostrar.Visible And cmdmostrar.Enabled Then
     cmdmostrar.SetFocus
   End If
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(txt_cli.Text)) = Left(VALOR, Len(Trim(txt_cli.Text))) Then
   Else
      Exit Sub
   End If
   lblCliente.Caption = Trim(ListView1.ListItems.Item(loc_key).Text)
   txt_cli.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
   If cmdmostrar.Visible And cmdmostrar.Enabled Then
     cmdmostrar.SetFocus
   End If
End If

dale:
ListView1.Visible = False
fin:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul txt_cli, txt_cli

End Sub

Private Sub txt_cli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(txt_cli.Text) = 0 Or IsNumeric(txt_cli.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Then
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
     PROC_LISVIEW ListView1
     loc_key = 0
     If ListView1.Visible Then
      loc_key = 1
     End If
 '   End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(txt_cli.Text), lvwText, , lvwPartial)
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


Public Sub IMP_FACTURACION()
On ERRO GoTo FINTODO
Dim wranF, wran1, wran2, WPAS
Dim LETRAS(24) As String * 1
Dim FECHA_INICIO
Dim WSFECHA As Date
Dim TABLAGEN2() As String * 30
Dim I, WS_SUBGRUPO
Dim c1 As Integer
Dim f1 As Integer
Dim nuc_repo  As rdoResultset
Dim PSNUC_REPO As rdoQuery
Dim far_r  As rdoResultset
Dim PSFAR As rdoQuery
Dim Mensaje, titulo, valorpred As String
Dim wvalor
Dim VALOR As Double
Dim xcuenta As Integer
Dim PAG_HOJAS As Integer
Dim WS_BRUTO As Currency
Dim w_fbg As String * 1
Dim w_serie As Integer
Dim w_numfac_ini As Currency
Dim w_numfac_fin As Currency
Dim llave_rep03 As rdoResultset
Dim PS_REP03 As rdoQuery
Dim WS_FLETE As Currency

Dim wformula, wformula1, wformula2, wformula3, wformula4
Dim Modo, Modo1
Dim wsFECHA1, wsFECHA2
Dim Wche, wkSELECT
Dim wfiltra1, wfiltra2, wfiltra3, wfecha, WDOCU
Dim wscliente As String
If Not SON_FECHAS() Then
 Exit Sub
End If
pub_cadena = ""
GoSub ARMA_SELE
pub_cadena = pub_cadena + " AND FAR_CODCIA = '" + LK_CODCIA + "' AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E' "
'pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC >= ? AND FAR_NUMFAC <= ? AND FAR_ESTADO <> 'E' ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
Set PSFAR = CN.CreateQuery("", pub_cadena)
Set far_r = PSFAR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


far_r.Requery

If far_r.EOF = True Then
   MsgBox "!!! NO EXISTEN Documentos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
FrmRepo.pbMax = far_r.RowCount
FrmRepo.pbMin = 0
FrmRepo.pbValue = 0
FrmRepo.pb.Visible = True
FrmRepo.lblmensa.Visible = True
DoEvents
If far_r!FAR_CODCIA <> LK_CODCIA Then
   MsgBox "!!! NO TE CORRESPONDE...", 48, Pub_Titulo
   GoTo CANCELA
End If
If far_r!far_fecha <> LK_FECHA_DIA Then
'   MsgBox "!!! NO ES DEL DIA...", 48, Pub_Titulo
'   GoTo CANCELA
End If
ws_clave = "0"
usu.Requery
Do Until usu.EOF
  If Trim(usu!usu_key) = "ADMIN" Then
    ws_clave = Trim(usu!USU_CLAVE)
    Exit Do
  End If
  usu.MoveNext
Loop

GoSub WEXCEL
If FrmRepo.chebol.Value = 1 Then
   xl.Worksheets(1).Activate
ElseIf FrmRepo.chefac.Value = 1 Then
   xl.Worksheets(2).Activate
ElseIf FrmRepo.cheguia.Value = 1 Then
   xl.Worksheets(3).Activate
End If
GoSub LETRAS
wmaxitem = 15
waumenta = 0
WS_SALDO = 0
xcuenta = 0
f1 = 0
FILA_INICIAL = 0
wnumfac = far_r!far_numfac
f1 = f1 + 1
xl.Cells(f1, 9) = "'" & far_r!far_numser & " - " & far_r!far_numfac
f1 = f1 + 1
SQ_OPER = 1
pu_codcia = LK_CODCIA
pu_cp = "C"
pu_codclie = far_r!far_codclie
LEER_CLI_LLAVE
If cli_llave.EOF Then
  MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
  GoTo CANCELA
End If
xl.Cells(f1, 2) = Trim(cli_llave!cli_nombre)
xl.Cells(f1, 7) = "'" & Format(far_r!far_fecha, "dd/mm/yyyy")
f1 = f1 + 1
xl.Cells(f1, 2) = Trim(cli_llave!CLI_CASA_DIREC) & " # " & cli_llave!CLI_CASA_NUM
If FrmRepo.chefac.Value = 1 Then
  xl.Cells(f1, 7) = "'" & Trim(cli_llave!cli_ruc_esposo)
End If
f1 = f1 + 2
fila = 0
WS_BRUTO = 0
SUB_CANT = 0
subtotal = 0
PUB_DESCTO = 0
If LK_EMP = "HER" Then
    WS_FLETE = 0
End If
Do Until far_r.EOF
   FrmRepo.pbValue = FrmRepo.pbValue + 1
   If wnumfac <> far_r!far_numfac Then
    waumenta = wmaxitem - fila
    f1 = f1 + waumenta + 1
    xl.Cells(f1 - 1, 2) = "SON: " & CONVER_LETRAS(WS_BRUTO, far_r!FAR_MONEDA)
    xl.Cells(f1, 2) = WS_BRUTO - WS_IMPTO + WS_DESCTO
    xl.Cells(f1, 3) = WS_DESCTO
    xl.Cells(f1, 4) = WS_IMPTO
    xl.Cells(f1, 8) = WS_BRUTO
    f1 = f1 + 1
    xl.Cells(f1, 9) = "'" & far_r!far_numser & " - " & far_r!far_numfac
    f1 = f1 + 1
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    pu_cp = "C"
    pu_codclie = far_r!far_codclie
    LEER_CLI_LLAVE
    If cli_llave.EOF Then
      MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
      GoTo CANCELA
    End If
    xl.Cells(f1, 2) = Trim(cli_llave!cli_nombre)
    xl.Cells(f1, 7) = "'" & Format(far_r!far_fecha, "dd/mm/yyyy")
    f1 = f1 + 1
    xl.Cells(f1, 2) = Trim(cli_llave!CLI_CASA_DIREC) & " # " & cli_llave!CLI_CASA_NUM
    If FrmRepo.chefac.Value = 1 Then
     xl.Cells(f1, 7) = "'" & Trim(cli_llave!cli_ruc_esposo)
    End If
    wnumfac = far_r!far_numfac
    f1 = f1 + 2
    fila = 0
    WS_BRUTO = 0
    SUB_CANT = 0
    subtotal = 0
    PUB_DESCTO = 0
   End If
   fila = fila + 1
   PUB_KEY = far_r!far_codart
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
      MsgBox "Error Grave en arti..."
      GoTo CANCELA
   Else
      xl.Cells(f1, 2) = Trim(art_LLAVE!art_nombre)
   End If
   xl.Cells(f1, 5) = far_r!far_descri
   xl.Cells(f1, 6) = Format(far_r!FAR_cantidad / far_r!FAR_equiv, "0.00")
   'If FrmRepo.cheguia.Value = 1 Then
    xl.Cells(f1, 7) = Format(far_r!FAR_PRECIO, "0.00")
   'End If
   subtotal = Format(far_r!FAR_PRECIO * (far_r!FAR_cantidad / far_r!FAR_equiv), "0.00")
   If far_r!FAR_DESCTO <> 0 Then
     subtotal = 0
   End If
   subtotal = redondea(subtotal)
   'If FrmRepo.cheguia.Value = 1 Then
    xl.Cells(f1, 8) = subtotal
   'End If
   WS_BRUTO = WS_BRUTO + subtotal
   SUB_CANT = SUB_CANT + (far_r!FAR_cantidad / far_r!FAR_equiv)
   WS_DESCTO = far_r!FAR_TOT_DESCTO
   WS_IMPTO = far_r!far_impto
   WS_GASTOS = far_r!FAR_GASTOS
   If LK_EMP = "HER" And far_r!FAR_TOT_FLETE <> 0 Then
    WS_FLETE = far_r!FAR_TOT_FLETE
   End If
   f1 = f1 + 1
   far_r.MoveNext
Loop
   waumenta = wmaxitem - fila
   f1 = f1 + waumenta + 1
   xl.Cells(f1 - 1, 2) = "SON: " & CONVER_LETRAS(WS_BRUTO, far_r!FAR_MONEDA)
   'If FrmRepo.chefac.Value = 1 Then
    If LK_EMP = "HER" Then
      WS_BRUTO = WS_BRUTO + WS_FLETE
      WS_IMPTO = WS_IMPTO + WS_FLETE
    End If
   xl.Cells(f1, 2) = WS_BRUTO - WS_IMPTO + WS_DESCTO
   xl.Cells(f1, 3) = WS_DESCTO
   xl.Cells(f1, 4) = WS_IMPTO
  'End If
  'If Left(CMBFGB.text, 1) = "F" Or Left(CMBFGB.text, 1) = "B" Then
   xl.Cells(f1, 8) = WS_BRUTO
  'End If
  DoEvents
  FrmRepo.lblmensa.Caption = "Mostrando Hoja de Calculo  . . . "
 'DoEvents
  
  'wranF = "B" & F1 & ":B" & F1
  'xl.Range(wranF).Font.Bold = True
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Worksheets(2).Protect ws_clave
  xl.Worksheets(3).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
'  FRMIMP.lblProceso.Visible = False
'  FRMIMP.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
   FrmRepo.pb.Visible = False
   FrmRepo.lblmensa.Visible = False
'  FRMIMP.pantalla.Enabled = True
'  FRMIMP.pantalla.Caption = "Por &Pantalla"
'  FRMIMP.lblProceso.Visible = False

Exit Sub


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

WEXCEL:

  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
 ' FRMIMP.lblProceso.Caption = "Abriendo , Archivo Facturación.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\Facturacion.xls", 0, True, 4, WPAS, WPAS
Return

Exit Sub
CANCELA:
  'FRMIMP.pantalla.Enabled = True
  'FRMIMP.pantalla.Caption = "Por &Pantalla"
  'FRMIMP.lblProceso.Visible = False
  LOC_CANCELA = 0
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmRepo


Exit Sub

ARMA_SELE:

pub_cadena = "SELECT * FROM FACART WHERE "
pub_cadena = pub_cadena + " FAR_FECHA >= '" + Format(REP_FECHA1, "yyyy/mm/dd") + "' AND  FAR_FECHA <= '" + Format(REP_FECHA1, "yyyy/mm/dd") + "' "
wfiltra = ""
Wche = 0
For fila = 0 To FrmRepo.listven.ListCount - 1
  FrmRepo.listven.ListIndex = fila
  If FrmRepo.listven.Selected(fila) Then
    Wche = 1
    wfiltra = wfiltra + " FAR_CODVEN = " + Str(Val(Left(FrmRepo.listven.Text, 4))) + " OR "
  End If
Next fila
If Wche <> 0 Then
 wfiltra = Left(wfiltra, Len(wfiltra) - 3)
 pub_cadena = pub_cadena + " AND (" + wfiltra + ") "
End If
wfiltra = ""
If chefac.Value = 1 Then
 wfiltra = wfiltra + " FAR_FBG = 'F' OR"
End If
If chebol.Value = 1 Then
wfiltra = wfiltra + " FAR_FBG = 'B' OR"
End If
If cheguia.Value = 1 Then
 wfiltra = wfiltra + " FAR_FBG = 'G' OR"
End If
If Trim(wfiltra) <> "" Then
 If Right(wfiltra, 1) = "R" Then
  wfiltra = Mid(wfiltra, 1, Len(wfiltra) - Len(Right(wfiltra, 2)))
 End If
 pub_cadena = pub_cadena + " AND (" + wfiltra + ") "
End If
wfiltra = ""
If chefac.Value = 1 Then
  If Trim(fac(0).Text) <> "" Then
    wfiltra = wfiltra + "FAR_NUMSER = " & Trim(fac(0).Text)
    If Trim(fac(1).Text) <> "" And Trim(fac(2).Text) <> "" Then
      wfiltra = wfiltra + " AND FAR_NUMFAC >= " & Trim(fac(1).Text) & " AND FAR_NUMFAC <= " & Trim(fac(2).Text)
    End If
  End If
End If
If Trim(wfiltra) <> "" Then
 pub_cadena = pub_cadena + " AND " + wfiltra + " "
End If
wfiltra = ""
If chebol.Value = 1 Then
  If Trim(bol(0).Text) <> "" Then
    wfiltra = wfiltra + "FAR_NUMSER = " & Trim(bol(0).Text)
    If Trim(bol(1).Text) <> "" And Trim(bol(2).Text) <> "" Then
      wfiltra = wfiltra + " AND FAR_NUMFAC >= " & Trim(bol(1).Text) & " AND FAR_NUMFAC <= " & Trim(bol(2).Text)
    End If
  End If
End If
If Trim(wfiltra) <> "" Then
 pub_cadena = pub_cadena + " AND " + wfiltra + " "
End If
wfiltra = ""
If cheguia.Value = 1 Then
  If Trim(guia(0).Text) <> "" And Trim(guia(1).Text) <> "" Then
    wfiltra = wfiltra + "FAR_NUMFAC >= " & Trim(guia(0).Text) & " AND FAR_NUMFAC <= " & Trim(guia(1).Text)
  End If
End If
If Trim(wfiltra) <> "" Then
 pub_cadena = pub_cadena + " AND " + wfiltra + " "
End If

Return
 
End Sub


Public Function SON_FECHAS() As Boolean
SON_FECHAS = True
If Right(FrmRepo.fechaini.Text, 2) = "__" Then
  REP_FECHA1 = Left(FrmRepo.fechaini.Text, 8)
Else
 REP_FECHA1 = Trim(FrmRepo.fechaini.Text)
End If
If Not IsDate(REP_FECHA1) Then
    MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
    Azul2 FrmRepo.fechaini, FrmRepo.fechaini
    GoTo fin
End If
If Right(FrmRepo.fechafin.Text, 2) = "__" Then
  REP_FECHA2 = Left(FrmRepo.fechafin.Text, 8)
Else
  REP_FECHA2 = Trim(FrmRepo.fechafin.Text)
End If
If Not IsDate(REP_FECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, pu_titulo
 Azul2 FrmRepo.fechafin, FrmRepo.fechafin
 GoTo fin
End If
If CDate(REP_FECHA1) > CDate(REP_FECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, pu_titulo
 Azul2 FrmRepo.fechaini, FrmRepo.fechaini
  GoTo fin
End If

Exit Function
fin:
SON_FECHAS = False

End Function

Public Sub LLENA_VEND()
pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
Set PSPRO_V = CN.CreateQuery("", pub_cadena)
PSPRO_V(0) = 0
Set PRO_V = PSPRO_V.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PSPRO_V(0) = LK_CODCIA
PRO_V.Requery
listven.Clear
Do Until PRO_V.EOF
    codi = PRO_V!VEM_CODVEN
    listven.AddItem Format(codi, "000") & " - " & Trim(PRO_V!VEM_NOMBRE)
    PRO_V.MoveNext
Loop
listven.Visible = True

End Sub

Public Sub LLENA_FECHAS()
fechaini.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
fechaini.Mask = "##/##/####"
fechaini.Enabled = True
fechafin.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
fechafin.Mask = "##/##/####"
fechafin.Enabled = True
End Sub

Public Sub PROCESA_DOCU()
Dim PSPRO1 As rdoQuery
Dim pro1_llave As rdoResultset
pub_cadena = "SELECT far_numfac, far_numser, far_fbg FROM facart WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10 ORDER BY  FAR_FECHA, FAR_FBG, FAR_NUMSER, FAR_NUMFAC"
Set PSPRO1 = CN.CreateQuery("", pub_cadena)
Set pro1_llave = PSPRO1.OpenResultset(rdOpenKeyset, rdConcurValues)
PSPRO1(0) = LK_CODCIA
PSPRO1(1) = REP_FECHA1
PSPRO1(2) = REP_FECHA2
pro1_llave.Requery
If pro1_llave.EOF Then
  Exit Sub
End If
WFBG = pro1_llave!far_fbg
WNUMSER = pro1_llave!far_numser
wnumfac = pro1_llave!far_numfac
  docu.AddItem Trim(pro1_llave!far_fbg) & "/ " & Format(pro1_llave!far_numser, "000") & " - " & Format(pro1_llave!far_numfac, "0000000")
Do Until pro1_llave.EOF
  If Trim(WFBG) = Trim(pro1_llave!far_fbg) And Trim(WNUMSER) = Trim(pro1_llave!far_numser) And Val(wnumfac) = Val(pro1_llave!far_numfac) Then
  Else
    docu.AddItem Trim(pro1_llave!far_fbg) & "/ " & Format(pro1_llave!far_numser, "000") & " - " & Format(pro1_llave!far_numfac, "0000000")
  End If
  WFBG = pro1_llave!far_fbg
  WNUMSER = pro1_llave!far_numser
  wnumfac = pro1_llave!far_numfac
  pro1_llave.MoveNext
Loop
lisdocu.Visible = True
docu.SetFocus
End Sub



Public Sub CONSO_CAMBIO()
Dim PSPRO1 As rdoQuery
Dim pro1_llave As rdoResultset

Dim Modo, Modo1
Dim wsFECHA1, wsFECHA2
Dim Wche, wkSELECT
Dim wfiltra1, wfiltra2, wfiltra3, wfecha, WDOCU
Dim wscliente As String
Dim wcodcia As String

CONSO:
CADENITA = ""
On Error GoTo procancela
cmdmostrar.Enabled = False
If Right(fechaini.Text, 2) = "__" Then
  wsFECHA1 = Left(fechaini.Text, 8)
Else
 wsFECHA1 = Trim(fechaini.Text)
End If
If Not IsDate(wsFECHA1) Then
    cmdmostrar.Enabled = True
    MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
    Azul2 fechaini, fechaini
    Exit Sub
End If
If Right(fechafin.Text, 2) = "__" Then
  wsFECHA2 = Left(fechafin.Text, 8)
Else
  wsFECHA2 = Trim(fechafin.Text)
End If
If Not IsDate(wsFECHA2) Then
 cmdmostrar.Enabled = True
 MsgBox "Fecha Invalidad ..", 48, pu_titulo
 Azul2 fechafin, fechafin
 Exit Sub
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 cmdmostrar.Enabled = True
 MsgBox "Fecha Invalidad ..", 48, pu_titulo
 Azul2 fechaini, fechaini
 Exit Sub
End If
REP_FECHA1 = wsFECHA1
REP_FECHA2 = wsFECHA2

If Wfile = "PLANILLA" Or Wfile = "CONSOLIDADO" And LK_EMP = "PAR" Then
  pub_mensaje = "Desea Mostrar los Documentos del rango de Fechas ? "
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta <> vbNo Then
     PROCESA_DOCU
     Exit Sub
  End If
End If


lblmensa.Visible = True
DoEvents
pbMin = 0
pbMax = 10
pbValue = 0
pb.Visible = True
If wsFECHA1 = wsFECHA2 Then
 wfecha = "CONSOLIDADO DE CAMBIOS DE PRODUCTOS DEL DIA " & wsFECHA1
Else
 wfecha = "CONSOLIDADO DE CAMBIOS DE PRODUCTOS DEL DIA " & wsFECHA1 & " AL " & wsFECHA2
End If
If Wfile = "CONTADO" Then
 If wsFECHA1 = wsFECHA2 Then
  wfecha = "VENTAS AL CONTADO DEL DIA " & wsFECHA1
 Else
  wfecha = "VENTAS AL CONTADO DEL DIA " & wsFECHA1 & " AL " & wsFECHA2
 End If
End If

pbValue = pbValue + 1
CADENITA = ""
CADEFECHA = ""
wfiltra1 = ""
Wche = 0
Modo1 = "FAR_CODVEN IN ("
For fila = 0 To listven.ListCount - 1
  listven.ListIndex = fila
  If listven.Selected(fila) And Trim(listven.Text) <> "" Then
    Wche = 1
    wkSELECT = Str(Val(Left(listven.Text, 3)))
    wfiltra1 = wfiltra1 + wkSELECT + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
 CADENITA = Left(Modo1, Len(Modo1) - 1) & ") AND "
 wfiltra1 = Left(wfiltra1, Len(wfiltra1) - 1)
End If
pbValue = pbValue + 1
wfiltra2 = ""
Wche = 0
Modo1 = "CLI_TRAB_ZONA in ("
For fila = 0 To listdistrito.ListCount - 1
  listdistrito.ListIndex = fila
  If listdistrito.Selected(fila) And Trim(listdistrito.Text) <> "" Then
    Wche = 1
    wkSELECT = Trim(Right(listdistrito.Text, 10))
    wfiltra2 = wfiltra2 + Trim(Left(listdistrito.Text, 30)) + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
 CADENITA = CADENITA + Left(Modo1, Len(Modo1) - 1) & ") AND "
 wfiltra2 = Left(wfiltra2, Len(wfiltra2) - 1)
End If

wfiltra2 = ""
Wche = 0
Modo1 = "CLI_GRUPO in ("
For fila = 0 To tipon.ListCount - 1
  tipon.ListIndex = fila
  If tipon.Selected(fila) And Trim(tipon.Text) <> "" Then
    Wche = 1
    wkSELECT = Trim(Right(tipon.Text, 10))
    wfiltra2 = wfiltra2 + Trim(Left(tipon.Text, 30)) + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
 CADENITA = CADENITA + Left(Modo1, Len(Modo1) - 1) & ") AND "
 wfiltra2 = Left(wfiltra2, Len(wfiltra2) - 1)
End If



pbValue = pbValue + 1
wfiltra3 = ""
Wche = 0
Modo1 = "CLI_ZONA_NEW in ("
For fila = 0 To listprovincia.ListCount - 1
  listprovincia.ListIndex = fila
  If listprovincia.Selected(fila) And Trim(listprovincia.Text) <> "" Then
    Wche = 1
    wkSELECT = Trim(Right(listprovincia.Text, 10))
    wfiltra3 = wfiltra3 + Trim(Left(listprovincia.Text, 30)) + ","
    Modo1 = Modo1 + wkSELECT + ","
  End If
Next fila
If Wche <> 0 Then
 CADENITA = CADENITA + Left(Modo1, Len(Modo1) - 1) & ") AND "
 wfiltra3 = Left(wfiltra3, Len(wfiltra3) - 1)
End If
pbValue = pbValue + 1

wfiltra = ""
If wfiltra2 <> "" Then
 wfiltra = wfiltra + "DISTRITO : " & wfiltra2
End If
If wfiltra3 <> "" Then
 wfiltra = wfiltra + " ZONA : " & wfiltra3
End If
If wfiltra1 <> "" Then
 wfiltra = wfiltra + " VENDEDOR : " & wfiltra1
End If
pbValue = pbValue + 1
DIA = Day(wsFECHA1)
mes = Month(wsFECHA1)
ano = Year(wsFECHA1)

DIA1 = Day(wsFECHA2)
MES1 = Month(wsFECHA2)
ANO1 = Year(wsFECHA2)

CADEFECHA = CADEFECHA + "FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10 AND PRE_FLAG_UNIDAD = 'A' AND FAR_CODCIA = '" & LK_CODCIA & "' AND FAR_FECHA >= '" & Format(wsFECHA1, "yyyy/mm/dd") & "' AND FAR_FECHA <= '" & Format(wsFECHA1, "yyyy/mm/dd") & "' "

CADENITA = CADENITA + CADEFECHA + " AND "
WDOCUF = ""
wdocuB = ""
Modo1 = "X"
cade = ""
If chefac.Value = 1 And chebol.Value = 0 And cheguia.Value = 0 Then
  Modo1 = "FAR_FBG in ('F') "
  wdocuB = "FAC."
  If Trim(fac(0).Text) <> "" Then
    cade = cade + "FAR_NUMSER = '" & Trim(fac(0).Text) & "' "
    WDOCUF = WDOCUF & " - SERIE : " & Trim(fac(0).Text)
    If Trim(fac(1).Text) <> "" And Trim(fac(2).Text) <> "" Then
      cade = cade + "AND FAR_NUMFAC >= " & Trim(fac(1).Text) & " AND FAR_NUMFAC <= " & Trim(fac(2).Text) & " "
      WDOCUF = WDOCUF & " Nº. " & Trim(fac(1).Text) & " AL " & Trim(fac(2).Text)
    End If
  End If
End If

If chefac.Value = 0 And chebol.Value = 1 And cheguia.Value = 0 Then
  Modo1 = "FAR_FBG in ('B') "
  wdocuB = " BOL."
  If Trim(bol(0).Text) <> "" Then
    cade = cade + "FAR_NUMSER = '" & Trim(bol(0).Text) & "' "
    wdocuB = wdocuB & " - SERIE : " & Trim(bol(0).Text)
    If Trim(bol(1).Text) <> "" And Trim(bol(2).Text) <> "" Then
      cade = cade + "AND FAR_NUMFAC >= " & Trim(bol(1).Text) & " AND FAR_NUMFAC <= " & Trim(bol(2).Text) & " "
      wdocuB = wdocuB & " Nº. " & Trim(bol(1).Text) & " AL " & Trim(bol(2).Text)
    End If
  End If
End If

If chefac.Value = 0 And chebol.Value = 0 And cheguia.Value = 1 Then
  Modo1 = "FAR_FBG in ('G') "
  wdocuB = " GUIA"
  If Trim(guia(0).Text) <> "" And Trim(guia(1).Text) <> "" Then
    cade = cade + "FAR_NUMFAC >= " & Trim(guia(0).Text) & " AND FAR_NUMFAC <= " & Trim(guia(1).Text) & " "
    wdocug = wdocug & " Nº. " & Trim(guia(0).Text) & " AL " & Trim(guia(1).Text)
  End If
End If

If chefac.Value = 1 And chebol.Value = 1 And cheguia.Value = 0 Then
  Modo1 = "FAR_FBG in ('F') "
  wdocuB = " FAC.,BOL."
  If Trim(fac(0).Text) <> "" Then
    cade = cade + "FAR_NUMSER = '" & Trim(fac(0).Text) & "' "
    WDOCUF = WDOCUF & " - SERIE : " & Trim(fac(0).Text)
    If Trim(fac(1).Text) <> "" And Trim(fac(2).Text) <> "" Then
      cade = cade + "AND FAR_NUMFAC >= " & Trim(fac(1).Text) & " AND FAR_NUMFAC <= " & Trim(fac(2).Text) & " "
      WDOCUF = WDOCUF & " Nº. " & Trim(fac(1).Text) & " AL " & Trim(fac(2).Text)
    End If
  End If
  wdocug = ""
  If Trim(bol(0).Text) <> "" Then
    If cade <> "" Then
       cade = cade + "OR FAR_FBG in ('B') AND FAR_NUMSER = '" & Trim(bol(0).Text) & "' "
       wdocug = "1"
    Else
       Modo1 = "FAR_FBG in ('F','B') "
       cade = cade + " FAR_NUMSER = '" & Trim(bol(0).Text) & "' "
    End If
    WDOCUF = WDOCUF & ", - SERIE : " & Trim(bol(0).Text)
    If Trim(bol(1).Text) <> "" And Trim(bol(2).Text) <> "" Then
      If wdocug = "1" Then
        cade = cade + "AND FAR_NUMFAC >= " & Trim(bol(1).Text) & " AND FAR_NUMFAC <= " & Trim(bol(2).Text) & " "
      Else
        cade = cade + "AND FAR_NUMFAC >= " & Trim(bol(1).Text) & " AND FAR_NUMFAC <= " & Trim(bol(2).Text) & " "
      End If
      WDOCUF = WDOCUF & ", Nº. " & Trim(bol(1).Text) & " AL " & Trim(bol(2).Text)
    Else
      cade = cade + " "
    End If
  Else
   If cade <> "" Then
     cade = cade + "OR (FAR_FBG in ('B')"
    Else
     Modo1 = "FAR_FBG in ('F','B') "
   End If
  End If
End If

If chefac.Value = 1 And chebol.Value = 0 And cheguia.Value = 1 Then
  Modo1 = "FAR_FBG in ('F') "
  wdocuB = " FAC.,GUIA"
  If Trim(fac(0).Text) <> "" Then
    cade = cade + "FAR_NUMSER = '" & Trim(fac(0).Text) & "' "
    WDOCUF = WDOCUF & " - SERIE : " & Trim(fac(0).Text)
    If Trim(fac(1).Text) <> "" And Trim(fac(2).Text) <> "" Then
      cade = cade + "AND FAR_NUMFAC >= " & Trim(fac(1).Text) & " AND FAR_NUMFAC <= " & Trim(fac(2).Text) & " "
      WDOCUF = WDOCUF & " Nº. " & Trim(fac(1).Text) & " AL " & Trim(fac(2).Text)
    End If
  End If
  If Trim(guia(0).Text) <> "" And Trim(guia(1).Text) <> "" Then
    If cade <> "" Then
      cade = cade + " OR (FAR_FBG in ('G') AND FAR_NUMFAC >= " & Trim(guia(0).Text) & " AND FAR_NUMFAC <= " & Trim(guia(1).Text) & ") "
    Else
     cade = cade + " (FAR_FBG in ('G') AND FAR_NUMFAC >= " & Trim(guia(0).Text) & " AND FAR_NUMFAC <= " & Trim(guia(1).Text) & ") "
    End If
    WDOCUF = WDOCUF & ", Nº. " & Trim(guia(0).Text) & " AL " & Trim(guia(1).Text)
  End If
End If
If chefac.Value = 0 And chebol.Value = 1 And cheguia.Value = 1 Then
  Modo1 = "FAR_FBG in ('B','G') "
  wdocuB = " BOL.,GUIA"
  If Trim(bol(0).Text) <> "" Then
    cade = cade + "FAR_NUMSER = '" & Trim(bol(0).Text) & "' "
    WDOCUF = WDOCUF & ", - SERIE : " & Trim(bol(0).Text)
    If Trim(bol(1).Text) <> "" And Trim(bol(2).Text) <> "" Then
      cade = cade + "AND FAR_NUMFAC >= " & Trim(bol(1).Text) & " AND FAR_NUMFAC <= " & Trim(bol(2).Text) & " "
      WDOCUF = WDOCUF & ", Nº. " & Trim(bol(1).Text) & " AL " & Trim(bol(2).Text)
    End If
  End If
  If Trim(guia(0).Text) <> "" And Trim(guia(1).Text) <> "" Then
    cade = cade + "OR (FAR_NUMFAC >= " & Trim(guia(0).Text) & " AND FAR_NUMFAC <= " & Trim(guia(1).Text) & ") "
    WDOCUF = WDOCUF & ", Nº. " & Trim(guia(0).Text) & " AL " & Trim(guia(1).Text)
  End If
End If
pbValue = pbValue + 1
If chefac.Value = 1 And chebol.Value = 1 And cheguia.Value = 1 Then
  Modo1 = "FAR_FBG in ('F','B','G') "
  wdocuB = " FAC.,BOL.,GUIA"
  If Trim(fac(0).Text) <> "" Then
    cade = cade + "FAR_NUMSER = '" & Trim(fac(0).Text) & "' "
    WDOCUF = WDOCUF & " - SERIE : " & Trim(fac(0).Text)
    If Trim(fac(1).Text) <> "" And Trim(fac(2).Text) <> "" Then
      cade = cade + "AND FAR_NUMFAC >= " & Trim(fac(1).Text) & " AND FAR_NUMFAC <= " & Trim(fac(2).Text) & " "
      WDOCUF = WDOCUF & " Nº. " & Trim(fac(1).Text) & " AL " & Trim(fac(2).Text)
    End If
  End If
  If Trim(bol(0).Text) <> "" Then
    cade = cade + "OR (FAR_NUMSER = '" & Trim(bol(0).Text) & "' "
    WDOCUF = WDOCUF & ", - SERIE : " & Trim(bol(0).Text)
    If Trim(bol(1).Text) <> "" And Trim(bol(2).Text) <> "" Then
      cade = cade + "AND FAR_NUMFAC >= " & Trim(bol(1).Text) & " AND FAR_NUMFAC <= " & Trim(bol(2).Text) & ") "
      WDOCUF = WDOCUF & ", Nº. " & Trim(bol(1).Text) & " AL " & Trim(bol(2).Text)
    Else
      cade = cade + ")"
    End If
  End If
  If Trim(guia(0).Text) <> "" And Trim(guia(1).Text) <> "" Then
    cade = cade + "OR (FAR_NUMFAC >= " & Trim(guia(0).Text) & " AND FAR_NUMFAC <= " & Trim(guia(1).Text) & ") "
    WDOCUF = WDOCUF & ", Nº. " & Trim(guia(0).Text) & " AL " & Trim(guia(1).Text)
  End If
End If
If Modo1 <> "X" Then
 CADENITA_TEMPO = CADENITA + "(" + Modo1 + " AND "
 CADENITA = CADENITA + Modo1 + " AND "
 If cade <> "" Then
  CADENITA = CADENITA_TEMPO
  CADENITA = CADENITA + cade + ") AND "
 Else
   CADENITA = CADENITA + Modo1 + " AND "
 End If
 
End If
pbValue = pbValue + 1
If Trim(Right(CADENITA, 2)) = "D" Then
 CADENITA = Left(CADENITA, Len(CADENITA) - 4)
End If
If Trim(Right(CADENITA, 2)) = "R" Then
 CADENITA = Left(CADENITA, Len(CADENITA) - 4)
End If
pbValue = pbValue + 1
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.Formulas(3) = ""
Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hresu2.rpt"
wcodcia = LK_CODCIA
DoEvents
wformula1 = "FECHA=  '" & wfecha & "'"
wformula2 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
wformula3 = "ZONAVEN=  '" & wfiltra & "'"
wformula4 = "DOCU=  '" & wdocuB & WDOCUF & "'"
pbValue = pbValue + 1
Reportes.Formulas(0) = wformula1
Reportes.Formulas(1) = wformula2
Reportes.Formulas(2) = wformula3
Reportes.Formulas(3) = wformula4
Modo1 = "(FAR_CODART = ART_KEY) AND (FAR_CODCIA = ART_CODCIA) AND (FAR_CODCLIE = CLI_CODCLIE) AND (FAR_CODCIA = CLI_CODCIA) AND (PRE_CODART = ART_KEY) AND (PRE_CODCIA = ART_CODCIA) AND FAR_NUMFAC_C <> 0 "
pub_cadena = "SELECT DISTINCT FAR_NUMSER_C, FAR_NUMFAC_C FROM FACART, CLIENTES,ARTI,PRECIOS WHERE " & Modo1 & " AND " & CADENITA
Set PSPRO1 = CN.CreateQuery("", pub_cadena)
Set pro1_llave = PSPRO1.OpenResultset(rdOpenKeyset, rdConcurValues)
pro1_llave.Requery
If pro1_llave.EOF Then

  If LK_EMP <> "PIU" Then
'     MsgBox "No hay Cambio de Productos", 48, Pub_Titulo
  End If
  GoTo NOCAM2
End If
CADENITA = "{FACART.FAR_NUMFAC} IN ["
Do Until pro1_llave.EOF
If pro1_llave.AbsolutePosition = 20 Then
  CADENITA = Left(CADENITA, Len(CADENITA) - 1) + "]"
  CADENITA = CADENITA + " OR {FACART.FAR_NUMFAC} IN ["
End If
If pro1_llave.AbsolutePosition = 40 Then
  CADENITA = Left(CADENITA, Len(CADENITA) - 1) + "]"
  CADENITA = CADENITA + " OR {FACART.FAR_NUMFAC} IN ["
End If
If pro1_llave.AbsolutePosition = 60 Then
  CADENITA = Left(CADENITA, Len(CADENITA) - 1) + "]"
  CADENITA = CADENITA + " OR {FACART.FAR_NUMFAC} IN ["
End If
If pro1_llave.AbsolutePosition = 80 Then
  CADENITA = Left(CADENITA, Len(CADENITA) - 1) + "]"
  CADENITA = CADENITA + " OR {FACART.FAR_NUMFAC} IN ["
End If
If pro1_llave.AbsolutePosition = 100 Then
  CADENITA = Left(CADENITA, Len(CADENITA) - 1) + "]"
  CADENITA = CADENITA + " OR {FACART.FAR_NUMFAC} IN ["
End If
If pro1_llave.AbsolutePosition = 120 Then
  CADENITA = Left(CADENITA, Len(CADENITA) - 1) + "]"
  CADENITA = CADENITA + " OR {FACART.FAR_NUMFAC} IN ["
End If
If pro1_llave.AbsolutePosition = 140 Then
  CADENITA = Left(CADENITA, Len(CADENITA) - 1) + "]"
  CADENITA = CADENITA + " OR {FACART.FAR_NUMFAC} IN ["
End If
If pro1_llave.AbsolutePosition = 160 Then
  CADENITA = Left(CADENITA, Len(CADENITA) - 1) + "]"
  CADENITA = CADENITA + " OR {FACART.FAR_NUMFAC} IN ["
End If

CADENITA = CADENITA + Str(pro1_llave!FAR_NUMFAC_C) + ","
pro1_llave.MoveNext
Loop
CADENITA = Left(CADENITA, Len(CADENITA) - 1) + "]"
NOCAM:
pub_cadena = "{FACART.FAR_SIGNO_ARM} = -1 AND {FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 102 AND {PRECIOS.PRE_FLAG_UNIDAD} = 'A' AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND  (" & CADENITA & ")"
Reportes.SelectionFormula = pub_cadena
Reportes.WindowTitle = "Reporte :  " & wfecha
Reportes.Action = 1
Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hresu3.rpt"
wfecha = "CONSOLIDADO DE DOCUMENTOS POR CAMBIOS DE PRODUCTOS "
Reportes.WindowTitle = "Reporte :  " & wfecha
Reportes.Action = 1
NOCAM2:
Exit Sub

Stop
Reportes.Action = 1
If LK_EMP = "HER" And LK_ICA = "A" Then
   Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hguia1.rpt"
   wformula1 = "FECHA_EMI=  '" & LK_FECHA_DIA & "'"
   Reportes.Action = 1
   Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hguia2.rpt"
   wformula1 = "FECHA_EMI=  '" & LK_FECHA_DIA & "'"
   Reportes.Action = 1
End If
pbValue = pbValue + 1
pb.Visible = False
lblmensa.Visible = False
cmdmostrar.Enabled = True
Exit Sub
' PROCEDIMIENTO

procancela:
MsgBox Err.Description, 48, Pub_Titulo
cmdmostrar.Enabled = True
Resume Next
'Unload FrmRepo

End Sub

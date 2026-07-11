VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FORM_PED 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ordenes de Compra"
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
   Icon            =   "FORM_PED.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView1 
      Height          =   495
      Left            =   8040
      TabIndex        =   35
      Top             =   7680
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.Frame fraListar 
      Caption         =   "Seleccionar"
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
      Height          =   6615
      Left            =   2190
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   9675
      Begin VB.CheckBox cherot 
         Caption         =   "Mostrar solo Rotativos"
         Height          =   405
         Left            =   7350
         TabIndex        =   75
         Top             =   2190
         Width           =   1875
      End
      Begin MSComctlLib.ProgressBar probar 
         Height          =   225
         Left            =   750
         TabIndex        =   74
         Top             =   4410
         Visible         =   0   'False
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin RichTextLib.RichTextBox TEXTOVARPRO 
         Height          =   375
         Left            =   1050
         TabIndex        =   72
         Top             =   5190
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
         MultiLine       =   0   'False
         TextRTF         =   $"FORM_PED.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox fami 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   120
         TabIndex        =   8
         Top             =   630
         Width           =   3375
      End
      Begin VB.TextBox i_codart2 
         Height          =   315
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   34
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton opcomo 
         Caption         =   "x División"
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   0
         Left            =   4200
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opcomo 
         Caption         =   "x Sub División"
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   1
         Left            =   5880
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opcomo 
         Caption         =   "x Proveedor"
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opcomo 
         Caption         =   "x Articulo"
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.ListBox lisarti 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1440
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   2910
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Retornar"
         Height          =   465
         Left            =   5310
         Picture         =   "FORM_PED.frx":04C6
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6060
         Width           =   1035
      End
      Begin VB.CommandButton cmdpasar 
         Caption         =   "&Aceptar"
         Height          =   465
         Left            =   3120
         Picture         =   "FORM_PED.frx":0610
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6060
         Width           =   1215
      End
      Begin VB.ListBox subfami 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   3690
         TabIndex        =   9
         Top             =   660
         Width           =   3525
      End
      Begin MSFlexGridLib.MSFlexGrid gridarti 
         Height          =   3165
         Left            =   120
         TabIndex        =   73
         Tag             =   "9999"
         Top             =   2820
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   5583
         _Version        =   393216
         Rows            =   3
         Cols            =   3
         FixedRows       =   2
         FixedCols       =   2
         BackColorBkg    =   8421376
         FocusRect       =   2
         HighLight       =   2
         AllowUserResizing=   3
         Appearance      =   0
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
      Begin VB.Label Label8 
         Caption         =   "Un Momento... procesando informaciï¿½n."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1650
         TabIndex        =   69
         Top             =   3960
         Width           =   6135
      End
      Begin VB.Label lblarti 
         Caption         =   "Codigo de Articulo :"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lfami 
         Caption         =   "División:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Tag             =   "9999"
         Top             =   465
         Width           =   885
      End
      Begin VB.Label lsubfami 
         Caption         =   "Sub División:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   135
         Left            =   4200
         TabIndex        =   13
         Tag             =   "9999"
         Top             =   465
         Width           =   1365
      End
      Begin VB.Label larti 
         Caption         =   "Articulos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   135
         Left            =   210
         TabIndex        =   10
         Tag             =   "9999"
         Top             =   2640
         Width           =   885
      End
   End
   Begin VB.CommandButton MODIF 
      Caption         =   "&Modificar"
      Height          =   525
      Left            =   10440
      Picture         =   "FORM_PED.frx":0A52
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox MSChart1 
      Height          =   975
      Left            =   480
      ScaleHeight     =   915
      ScaleWidth      =   6075
      TabIndex        =   43
      Top             =   7440
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.CommandButton cmdIngreso 
      Caption         =   "&Ingreso"
      Height          =   525
      Left            =   10440
      Picture         =   "FORM_PED.frx":0B9C
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "Ce&rrar"
      Height          =   525
      Left            =   10440
      Picture         =   "FORM_PED.frx":0FDE
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame f1 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   11775
      Begin VB.CheckBox cheprecios 
         Caption         =   "Formato Especial de Precios."
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
         Height          =   225
         Left            =   7890
         TabIndex        =   76
         Top             =   120
         Width           =   3735
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Lista de Articulos"
         Height          =   615
         Left            =   9720
         Picture         =   "FORM_PED.frx":1128
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   960
         Width           =   1815
      End
      Begin VB.ListBox liscia 
         Height          =   60
         Left            =   11160
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   67
         Top             =   360
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.ComboBox moneda 
         Height          =   315
         ItemData        =   "FORM_PED.frx":156A
         Left            =   6000
         List            =   "FORM_PED.frx":1574
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chesin 
         Caption         =   "Orden Sin Valorizar"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   7560
         TabIndex        =   58
         Top             =   1020
         Width           =   1980
      End
      Begin VB.CheckBox checia 
         Caption         =   "02 - Calcular Sucursal"
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   5160
         TabIndex        =   57
         Top             =   120
         Width           =   2220
      End
      Begin VB.CheckBox cheneg 
         Caption         =   "Quitar los Negativos"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   7560
         TabIndex        =   56
         Top             =   1320
         Width           =   2340
      End
      Begin VB.ComboBox fpago 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   1635
         Width           =   2052
      End
      Begin VB.ComboBox txtagencia 
         Height          =   315
         Left            =   1920
         TabIndex        =   51
         Text            =   "txtagencia"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FORM_PED.frx":158C
         Left            =   120
         List            =   "FORM_PED.frx":158E
         Style           =   1  'Simple Combo
         TabIndex        =   48
         Top             =   1635
         Width           =   1200
      End
      Begin VB.TextBox txtcontacto 
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
         Left            =   2880
         TabIndex        =   46
         Top             =   1635
         Width           =   2175
      End
      Begin VB.ComboBox codpro 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   360
         Width           =   4935
      End
      Begin VB.TextBox txtdoc 
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
         Height          =   300
         Left            =   720
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox tserie 
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
         Height          =   300
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtdias 
         Height          =   300
         Left            =   8880
         MaxLength       =   3
         TabIndex        =   22
         Text            =   "7"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox tpromedio 
         Height          =   300
         Left            =   11040
         MaxLength       =   3
         TabIndex        =   21
         Text            =   "0"
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSMask.MaskEdBox fecha1 
         Height          =   285
         Left            =   8880
         TabIndex        =   70
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox fecha2 
         Height          =   285
         Left            =   10320
         TabIndex        =   71
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
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
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Compañia:"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   11160
         TabIndex        =   68
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Moneda:"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   5160
         TabIndex        =   64
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha:"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   5160
         TabIndex        =   62
         Top             =   480
         Width           =   735
      End
      Begin VB.Label fecha 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6000
         TabIndex        =   61
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lcodart 
         Caption         =   "F/Pago"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   50
         Tag             =   "9999"
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label lcodart 
         Caption         =   "Enviar a:"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   49
         Tag             =   "9999"
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lcodart 
         Caption         =   "Att."
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   47
         Tag             =   "9999"
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Proveedor :"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lcodart 
         Caption         =   "Nï¿½. Doc"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Tag             =   "9999"
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lcodart 
         Caption         =   "Cantidad Dias a Solicitar :"
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   1
         Left            =   7440
         TabIndex        =   28
         Tag             =   "9999"
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label lcodart 
         Caption         =   "Fecha para Promediar de :"
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   2
         Left            =   7440
         TabIndex        =   27
         Tag             =   "9999"
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lcodart 
         Caption         =   "Agencia:"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   26
         Tag             =   "9999"
         Top             =   840
         Width           =   765
      End
   End
   Begin VB.CommandButton cancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancelar"
      Height          =   525
      Left            =   10440
      Picture         =   "FORM_PED.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5160
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Tag             =   "0"
      Top             =   3480
      Visible         =   0   'False
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   0
      Min             =   77
      Max             =   91
   End
   Begin VB.Frame ESTADO 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Articulos :"
      Enabled         =   0   'False
      ForeColor       =   &H00808000&
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Tag             =   "100"
      Top             =   2040
      Width           =   11655
      Begin VB.CommandButton Imprimir 
         Caption         =   "Im&primir"
         Height          =   525
         Left            =   10320
         Picture         =   "FORM_PED.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox imax 
         Height          =   288
         Left            =   10200
         MaxLength       =   2
         TabIndex        =   53
         Text            =   "18"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox UNIDAD 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txttotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8640
         TabIndex        =   39
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtigv 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         TabIndex        =   38
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txtvalorv 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   37
         Top             =   4320
         Width           =   975
      End
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
         MultiLine       =   0   'False
         TextRTF         =   $"FORM_PED.frx":1CDC
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   4095
         Left            =   120
         TabIndex        =   0
         Tag             =   "9999"
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   7223
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         BackColorBkg    =   8421376
         FocusRect       =   2
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
      End
      Begin VB.Label lblpeso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   60
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Peso(Kg.):"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   59
         Top             =   4320
         Width           =   900
      End
      Begin VB.Label lblmoneda 
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8160
         TabIndex        =   55
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Maximo de Item :"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   10200
         TabIndex        =   54
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Total :"
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   3
         Left            =   7320
         TabIndex        =   42
         Tag             =   "9999"
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "I.G.V. :"
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   2
         Left            =   5280
         TabIndex        =   41
         Tag             =   "9999"
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Valor:"
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   40
         Tag             =   "9999"
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label momen 
         Caption         =   "Un Momento ..."
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   120
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
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
      Left            =   -330
      TabIndex        =   65
      Top             =   6870
      Width           =   11895
   End
End
Attribute VB_Name = "FORM_PED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As Object
Dim LOC_FECHA_INI As Date
Dim LOC_FECHA_FIN As Date
Dim loc_cancel As String * 1
Dim PSPRO_V As rdoQuery
Dim PRO_V As rdoResultset
Dim LOC_CODCIA As String * 2
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_REP03 As rdoQuery
Dim llave_rep03 As rdoResultset
Dim PS_REP04 As rdoQuery
Dim llave_rep04 As rdoResultset

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
Dim FLAG_MODI As Integer
Option Explicit

Private Sub cancelar_Click()
WMODO = ""
cmdIngreso.Caption = "&Ingreso"
f1.Enabled = False
ESTADO.Enabled = False
pb.Visible = False
Barra.Visible = False
fila = 0
SUM_D = 0
SUM_H = 0
LIMPIA_DATOS
CABE_MAN
FORM_PED.MODIF.Enabled = True
fecha.Caption = ""


'grid_fac.SetFocus
FORM_PED.MSChart1.Visible = False
End Sub

Private Sub cancelar_GotFocus()
fraListar.Visible = False
End Sub

Private Sub cmdcancel_Click()
On Error GoTo sale

fraListar.Visible = False
If grid_fac.Enabled Then grid_fac.SetFocus
Exit Sub
sale:
End Sub

Private Sub cmdIngreso_Click()
Dim flag_precio As String * 1
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

FLAG = False
ws_tot_debe = Val(grid_fac.TextMatrix(1, 4))
ws_tot_haber = Val(grid_fac.TextMatrix(1, 5))

'If ws_tot_debe = 0 And ws_tot_haber = 0 Then
'  MsgBox "Ingrese los Vouchers ..", 48, Pub_Titulo
'  If grid_fac.Rows > 2 Then
'     grid_fac.Col = 1
'     grid_fac.Row = 2
'     grid_fac.SetFocus
'   End If
  Barra.Visible = False
'  GoTo fin
'End If
If grid_fac.Rows > Val(imax.Text) + 2 Then
    MsgBox "LLego al Maximo de Item... Generar otra Orden de Compra", 48, Pub_Titulo
    cmdmostrar.SetFocus
    GoTo fin
End If
If grid_fac.Rows <= 2 Then
  MsgBox "Ingresar datos ...", 48, Pub_Titulo
  cmdmostrar.SetFocus
  GoTo fin
End If
If Trim(codpro.Text) = "" Then
  MsgBox "Ingresar Proveedor ...", 48, Pub_Titulo
  codpro.SetFocus
  GoTo fin
End If
If Trim(txtagencia.Text) = "" Then
  MsgBox "Ingresar Agencia de Transporte ...", 48, Pub_Titulo
  codpro.SetFocus
  GoTo fin
End If
suma_grid
flag_precio = ""
For fila = 2 To grid_fac.Rows - 1
 If grid_fac.TextMatrix(fila, 1) <> "" Then
  If Val(grid_fac.TextMatrix(fila, 2)) <= 0 Then
    MsgBox "Verificar, cantidad en cero o menor. - " & grid_fac.TextMatrix(fila, 1) & " : " & grid_fac.TextMatrix(fila, 0), 48, Pub_Titulo
    grid_fac.SetFocus
    GoTo fin
  End If
  If Val(grid_fac.TextMatrix(fila, 4)) = 0 Then
    flag_precio = "A"
    ''MsgBox "Verificar hay algun precio en 0 .", 48, Pub_Titulo
    ''grid_fac.SetFocus
 '   GoTo fin
  End If
End If
Next fila
If flag_precio = "A" Then
  pub_mensaje = "Existen precios com valor Cero.   ¿Desea Continuar... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
     GoTo fin
  End If
End If
If FLAG_MODI <> 1 Then
pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)

PSTEMP_LLAVE.rdoParameters(0) = LK_CODCIA
PSTEMP_LLAVE.rdoParameters(1) = Val(tserie.Text)
PSTEMP_LLAVE.rdoParameters(2) = Val(txtdoc.Text)
temp_llave.Requery
If Not temp_llave.EOF Then
   txtdoc.Text = Format(Val(txtdoc.Text) + 1, "0")
   con_llave.Close
   MsgBox "Existe Nro. de Docuemnto, el sistema generï¿½ el correlativo,  intente grabar nuevamente...", 48, Pub_Titulo
   Exit Sub
End If
End If


Screen.MousePointer = 11
DoEvents
Barra.Visible = True
DoEvents
Barra.Min = 0
Barra.max = fila
Barra.Value = 0
exito = True
Barra.Value = 1
If FLAG_MODI = 1 Then
   temp_llave.MoveFirst
   Do Until temp_llave.EOF
      temp_llave.Edit
      temp_llave.Delete
      temp_llave.MoveNext
   Loop
End If

GoSub ACT1

pub_mensaje = "Usted. desea imprimir la Orden de Compra... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbYes Then
  Imprimir_Click
End If
fila = 1
SUM_D = 0
SUM_H = 0
CABE_MAN
LIMPIA_DATOS
fila = 0
If Pub_Respuesta <> vbYes Then
cancelar.SetFocus
End If
Barra.Visible = False
cmdIngreso.Caption = "&Ingreso"

GoTo fin

ACT1:

fila = 1
ws_nro_voucher = ws_nro_voucher + 1
FLAG = False
WS_NRO_MOV = 0
fila = 2
Do While FLAG = False
   If Trim(grid_fac.TextMatrix(fila, 1)) = "" Then GoTo pasa
    ' grabo todo
   temp_llave.AddNew
   temp_llave!PED_CODCIA = LK_CODCIA
   temp_llave!PED_FECHA = LK_FECHA_DIA
   temp_llave!PED_NUMSER = Trim(tserie.Text)
   temp_llave!PED_NUMFAC = Val(txtdoc.Text)
   temp_llave!PED_NUMSEC = WS_NRO_MOV
   temp_llave!PED_CANTIDAD = Val(grid_fac.TextMatrix(fila, 2))
   temp_llave!PED_PRECIO = Val(grid_fac.TextMatrix(fila, 4))
   temp_llave!PED_CODUSU = LK_CODUSU
   temp_llave!PED_IGV = Val(txtigv.Text)
   temp_llave!PED_BRUTO = Val(txtvalorv.Text)
   temp_llave!PED_CONTACTO = Nulo_Valors(txtcontacto.Text)
   temp_llave!PED_ESTADO = "N"
   temp_llave!PED_CODUSU = LK_CODUSU
   temp_llave!PED_codart = Val(grid_fac.TextMatrix(fila, 9))
   temp_llave!PED_UNIDAD = Trim(grid_fac.TextMatrix(fila, 3))
   If cheprecios.Value = 1 Then
     temp_llave!PED_OFERTA = Trim(grid_fac.TextMatrix(fila, 8))
   Else
    temp_llave!PED_OFERTA = txtagencia.Text
   End If
   temp_llave!PED_NOMCLIE = Combo1.Text
   temp_llave!PED_FORMA = fpago.Text
   temp_llave!PED_RUCCLIE = txtagencia.ListIndex
   temp_llave!PED_MONEDA = Left(moneda.Text, 1)
   temp_llave!PED_SUBTOTAL = Trim(grid_fac.TextMatrix(fila, 5))
   temp_llave!PED_EQUIV = 1
   temp_llave!PED_CODCLIE = Val(Right(codpro.Text, 9))
   temp_llave!PED_TIPMOV = 20
   temp_llave!PED_TRANSP = Val(Right(txtagencia.Text, 6))
   temp_llave!PED_DIA_VISITA = cheprecios.Value
   temp_llave.Update
pasa:
   fila = fila + 1
   WS_NRO_MOV = WS_NRO_MOV + 1
   If fila >= FORM_PED.grid_fac.Rows Then
      FLAG = True
   End If
  
Loop
If FLAG_MODI <> 1 Then
 con_llave.Close
End If

FLAG_MODI = 0
Return

Screen.MousePointer = 1


Exit Sub
End If
Dim wser As String
Dim wnumfac As String

cmdIngreso.Caption = "&Grabar"
f1.Enabled = True
ESTADO.Enabled = True
LIMPIA_DATOS
FORM_PED.MODIF.Enabled = False
CABE_MAN
WMODO = "I"
PSTEMP_MAYOR(0) = LK_CODCIA
PSTEMP_MAYOR(1) = Val(tserie.Text)
temp_mayor.Requery
If temp_mayor.EOF Then
 wnumfac = 1
Else
 wnumfac = Val(Nulo_Valor0(temp_mayor!PED_NUMFAC)) + 1
End If
txtdoc.Text = wnumfac
'txtagencia.text = Nulo_Valors(par_llave!PAR_AGE_EMP)
If codpro.ListCount > 0 Then codpro.ListIndex = 0
fecha.Caption = Format(LK_FECHA_DIA, "dd/mm/yyyy")
moneda.ListIndex = 0
Exit Sub

error_fatal:
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
Screen.MousePointer = 0
Exit Sub
sale:
If Err.Number = 6 Then
  MsgBox "Verficar Importe.", 48, Pub_Titulo
  If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR
  FORM_PED.Barra.Visible = False
  Screen.MousePointer = 0
  grid_fac.SetFocus
Else
  MsgBox Err.Description, 48, Pub_Titulo
End If

End Sub



Private Sub cmdmostrar_Click()
If Val(txtdias.Text) = 0 And Val(tpromedio.Text) = 0 Then
  GoTo pasa
End If


If Val(txtdias.Text) <= 0 Then
  MsgBox "Definir cantidad de dias para la orden de compra.", 48, Pub_Titulo
  txtdias.SetFocus
  Exit Sub
End If
'If Val(tpromedio.Text) <= 0 Then
'  MsgBox "Definir dias para el Promedio.", 48, Pub_Titulo
'  tpromedio.SetFocus
'  Exit Sub
'End If
pasa:
fraListar.Visible = True
lisarti.Clear
opcomo(0).SetFocus
End Sub

Private Sub cmdpasar_Click()
Dim FLAG_MAX As String * 1
Dim wsFECHA1
Dim wsFECHA2
Dim WPUEDE
Dim fx As Integer
Dim WCODIGO As Currency
Dim wcanti As Currency
Dim costo As Currency
Dim wdia, wpro As Integer
Dim wcalculo As Currency
Dim stockact As Currency
Dim stockmin As Currency
Dim tdias As Integer
If Left(cmdpasar.Caption, 3) = "&De" Then
  cmdpasar.Enabled = False
  loc_cancel = "A"
  Exit Sub
End If
loc_cancel = ""
cmdpasar.Caption = "&Detener"
   
If Right(fecha1.Text, 2) = "__" Then
     wsFECHA1 = Left(fecha1.Text, 8)
Else
     wsFECHA1 = Trim(fecha1.Text)
End If
If Right(fecha2.Text, 2) = "__" Then
     wsFECHA2 = Left(fecha2.Text, 8)
Else
     wsFECHA2 = Trim(fecha2.Text)
End If
If Not IsDate(wsFECHA1) Then
   MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
   Exit Sub
End If
If Not IsDate(wsFECHA2) Then
   MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
   Exit Sub
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
   MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
   Exit Sub
End If


LOC_FECHA_INI = wsFECHA1
LOC_FECHA_FIN = wsFECHA2

If checia.Value = 1 Then
 LOC_CODCIA = "02"
Else
 LOC_CODCIA = LK_CODCIA
End If


tdias = Val(txtdias.Text)
wdia = Val(txtdias.Text)
wpro = Val(tpromedio.Text)
If wpro = 0 And wdia = 0 Then GoTo pasa
If wdia <= 0 Then
 MsgBox "cantidad de Dias. No Procede..", 48, Pub_Titulo
 GoTo TERMINAR
End If
'If wpro <= 0 Then
'  MsgBox "Dias para el Promedio. No Procede..", 48, Pub_Titulo
'  GoTo TERMINAR
'End If
'If wdia > wpro Then
'  MsgBox "Cantidad de Dias no procede en el rango de promedio..", 48, Pub_Titulo
'  GoTo TERMINAR
'End If
pasa:
If lisarti.ListCount = 0 Then
 lisarti.Visible = False
 MsgBox "No Procede...", 48, Pub_Titulo
 GoTo TERMINAR
End If
DoEvents
lisarti.Visible = False
DoEvents
pb.Visible = True
pb.Min = 0
pb.max = lisarti.ListCount
pb.Value = 0
SQ_OPER = 1
pu_codcia = LOC_CODCIA
wcalculo = 0
PUB_SECUEN = 0
GoTo PASA_DEL_GRID

''
Exit Sub
For fila = 0 To lisarti.ListCount - 1
 If grid_fac.Rows >= Val(imax.Text) + 2 Then
     MsgBox "Llego al Maximo de Item.. Generar otra Orden de Compra", 48, Pub_Titulo
     GoTo TERMINAR
 End If
 DoEvents
 If loc_cancel = "A" Then
   cmdpasar.Caption = "&Aceptar"
   cmdpasar.Enabled = True
   loc_cancel = ""
   Exit For
 End If
 pb.Value = pb.Value + 1
 lisarti.ListIndex = fila
 If lisarti.Selected(fila) = True Then
    PUB_KEY = Val(Right(lisarti, 8))
    PUB_CODART = Val(Right(lisarti, 8))
    SQ_OPER = 1
    LEER_ART_LLAVE
    stockact = 0
    If LISCIA.ListCount > 0 Then
     For fx = 0 To LISCIA.ListCount - 1
       LISCIA.ListIndex = fx
       If LISCIA.Selected(fx) Then
         pu_codcia = Left(LISCIA.Text, 2)
         LEER_ARM_LLAVE
         stockact = stockact + Nulo_Valor0(arm_llave!arm_stock)
       End If
      Next fx
    Else
      pu_codcia = LK_CODCIA
      LEER_ARM_LLAVE
      stockact = Nulo_Valor0(arm_llave!arm_stock)
    End If
    pu_codcia = LOC_CODCIA
    SQ_OPER = 2
    LEER_PRE_LLAVE
    pre_mayor.MoveLast
    costo = 0
    WCODIGO = Val(Right(lisarti, 8))
    If wpro = 0 And wdia = 0 Then
      wcanti = 0
    Else
      wcanti = PROCESA_CANTIDAD(WCODIGO, tdias)
    End If
    If Not art_LLAVE.EOF Then
      costo = Nulo_Valor0(arm_llave!ARM_COSTO_ULT)
      'stockact = Nulo_Valor0(arm_llave!ARM_STOCK)
otro_minimo:
      If Nulo_Valor0(art_LLAVE!ART_STOCK_MIN) = 0 Then
         WPUEDE = InputBox("El Producto no tiene Stock Minimo " & Chr(13) & "Articulo :" & Chr(13) & Trim(art_LLAVE!art_nombre) & Chr(13) & "Colocar Stock en Unidades Minimas :", "Stock Minimo.", 0#)
         If WPUEDE = "" Then
             stockmin = 0 ' pasa con Stock minimo 0
            GoTo otro_producto
         End If
         If WPUEDE = "" Then GoTo pasa_minimo
         If Val(WPUEDE) = 0 Then GoTo otro_minimo
          art_LLAVE.Edit
          art_LLAVE!ART_STOCK_MIN = Val(WPUEDE)
          art_LLAVE.Update
          art_LLAVE.Requery
pasa_minimo:
          stockmin = Val(WPUEDE)
      Else
         stockmin = Nulo_Valor0(art_LLAVE!ART_STOCK_MIN)
      End If
otro_producto:
      wcalculo = wcanti + stockmin - stockact
      If cheneg.Value = 1 And wcalculo < 0 Then GoTo sigueart
      If wpro = 0 And wdia = 0 Then
          wcalculo = 0
      End If
    End If
    grid_fac.Rows = grid_fac.Rows + 1
    grid_fac.RowHeight(grid_fac.Rows - 1) = 285
    grid_fac.TextMatrix(grid_fac.Rows - 1, 0) = "'" & Trim(art_LLAVE!art_nombre)  ' Trim(Mid(lisarti.Text, 11, 30))
    grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = Trim(art_LLAVE!ART_alterno)
    If cheprecios.Value = 0 Then
     grid_fac.TextMatrix(grid_fac.Rows - 1, 2) = Format(wcalculo / pre_mayor!PRE_EQUIV, "#######")
    Else
     grid_fac.TextMatrix(grid_fac.Rows - 1, 2) = 0
    End If
    grid_fac.TextMatrix(grid_fac.Rows - 1, 3) = pre_mayor!PRE_UNIDAD
    If chesin.Value = 0 Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = Format(costo, "0.00")
      grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = Format((Val(grid_fac.TextMatrix(grid_fac.Rows - 1, 4)) * Val(grid_fac.TextMatrix(grid_fac.Rows - 1, 2))), "####0.00")
    End If
    If cheprecios.Value = 1 Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = Format(wcalculo, "0.00")
    End If
    grid_fac.TextMatrix(grid_fac.Rows - 1, 6) = Format(wcanti / pre_mayor!PRE_EQUIV, "#######")
    grid_fac.TextMatrix(grid_fac.Rows - 1, 7) = Format(stockact / pre_mayor!PRE_EQUIV, "#######")
    grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = Format(stockmin / pre_mayor!PRE_EQUIV, "#######")
    grid_fac.TextMatrix(grid_fac.Rows - 1, 9) = PUB_KEY
    grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = wcalculo
    grid_fac.TextMatrix(grid_fac.Rows - 1, 11) = costo
    grid_fac.TextMatrix(grid_fac.Rows - 1, 12) = pre_mayor!PRE_SECUENCIA

    grid_fac.TextMatrix(grid_fac.Rows - 1, 14) = wcanti
    grid_fac.TextMatrix(grid_fac.Rows - 1, 15) = stockact
    grid_fac.TextMatrix(grid_fac.Rows - 1, 16) = stockmin
    grid_fac.TextMatrix(grid_fac.Rows - 1, 17) = pre_mayor!PRE_EQUIV
    grid_fac.TextMatrix(grid_fac.Rows - 1, 18) = pre_mayor!pre_PESO
    grid_fac.ColWidth(12) = pre_mayor!PRE_SECUENCIA
    
 End If
sigueart:
Next fila
DoEvents
lisarti.Visible = True
TERMINAR:
pb.Visible = False
cmdpasar.Caption = "&Aceptar"
cmdpasar.Enabled = True
loc_cancel = ""
suma_grid
lisarti.Clear


Exit Sub
PASA_DEL_GRID:
FLAG_MAX = ""

For fila = 2 To gridarti.Rows - 1
   If FLAG_MAX <> "A" Then
      If grid_fac.Rows >= Val(imax.Text) + 2 Then
        pub_mensaje = "LLego al Maximo de Item ¿Desea Continuar... ?"
        Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
        If Pub_Respuesta = vbNo Then
           Exit For
        End If
        FLAG_MAX = "A"
      End If
    End If
    pb.Value = pb.Value + 1
    If Val(gridarti.TextMatrix(fila, 9)) = 0 Then GoTo PASA_POR_ALTO

     grid_fac.Rows = grid_fac.Rows + 1
    grid_fac.RowHeight(grid_fac.Rows - 1) = 285
    grid_fac.TextMatrix(grid_fac.Rows - 1, 0) = "'" & gridarti.TextMatrix(fila, 1)
    
    grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = gridarti.TextMatrix(fila, 10)
    If chesin.Value = 0 Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = 0 ' Format(costo, "0.00")
      grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = 0 'Format((Val(grid_fac.TextMatrix(grid_fac.Rows - 1, 4)) * Val(grid_fac.TextMatrix(grid_fac.Rows - 1, 2))), "####0.00")
    End If
    If cheprecios.Value = 1 Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 2) = 1
      grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = Trim(gridarti.TextMatrix(fila, 9))
    Else
      grid_fac.TextMatrix(grid_fac.Rows - 1, 2) = Trim(gridarti.TextMatrix(fila, 9))
    End If
    
    
    grid_fac.TextMatrix(grid_fac.Rows - 1, 3) = gridarti.TextMatrix(fila, 2)
    
    
    
    
    grid_fac.TextMatrix(grid_fac.Rows - 1, 6) = 0 ' Format(wcanti / pre_mayor!PRE_EQUIV, "#######")
    grid_fac.TextMatrix(grid_fac.Rows - 1, 7) = 0 ' Format(stockact / pre_mayor!PRE_EQUIV, "#######")
    If cheprecios.Value = 1 Then
      grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = gridarti.TextMatrix(fila, 12) ' Format(stockmin / pre_mayor!PRE_EQUIV, "#######")
    End If
    grid_fac.TextMatrix(grid_fac.Rows - 1, 9) = gridarti.TextMatrix(fila, 11)
    grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = 0
    grid_fac.TextMatrix(grid_fac.Rows - 1, 11) = 0
    grid_fac.TextMatrix(grid_fac.Rows - 1, 12) = 0

    grid_fac.TextMatrix(grid_fac.Rows - 1, 14) = 0
    grid_fac.TextMatrix(grid_fac.Rows - 1, 15) = 0
    grid_fac.TextMatrix(grid_fac.Rows - 1, 16) = 0
    grid_fac.TextMatrix(grid_fac.Rows - 1, 17) = 0
    grid_fac.TextMatrix(grid_fac.Rows - 1, 18) = 0
PASA_POR_ALTO:
Next fila
lisarti.Visible = True
pb.Visible = False
lisarti.Visible = False
cmdpasar.Caption = "&Aceptar"
cmdpasar.Enabled = True
loc_cancel = ""
suma_grid
gridarti.Rows = 2
lisarti.Clear


End Sub

Private Sub cmdproorden_Click()

Dim ws_conta As Integer
Dim WS_NOMCLI As String * 20
Dim WS_NOMBAN As String * 20
Dim WS_SALDO As Currency
Dim WS_MONEDA As String * 1
Dim ww_moneda
Dim ws_mensaje
Dim WS_FBG, WS_LARGO, ws_nomche, ws_codclie
Dim WS_SALDO_ING As Currency
Dim WS_SALDO_SAL As Currency
Dim todos_cli

Dim wsFECHA1
Dim ws_codcia As String
Dim xcuenta As Integer
Dim docus As String
Dim docus2 As String
Dim wmonto As Currency
Dim WINGRESOS As Currency
Dim TOT_INGRESOS As Currency
Dim WSALDO_CAJA As Currency



wsFECHA1 = LK_FECHA_DIA

DoEvents
DoEvents
GoSub WEXCEL

xl.Worksheets(1).Activate
WS_SALDO = 0
xcuenta = 0
f1 = 5


WSALDO_CAJA = 0
f1 = f1 + 1
ws_conta = 0

VAMOS:
WS_SALDO = 0
WSALDO_CAJA = 0
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
Else
  PUB_CODCIA = LK_CODCIA
End If


SQ_OPER = 1
LEER_PAR_LLAVE
WS_SALDO = par_llave!PAR_SALDO_CAJA_ayer
WSALDO_CAJA = WSALDO_CAJA + WS_SALDO
xl.Cells(f1, 2) = "Saldo Anterior:"
xl.Cells(f1, 5) = WS_SALDO
'all_llave.MoveNext
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ? AND ALL_FBG = ?  AND ALL_TIPMOV = 10 AND ALL_SIGNO_CAJA = 1  AND all_flag_ext <> 'E' AND (ALL_CODTRA <> 1111 OR ALL_CODTRA <> 1122)  ORDER BY  ALL_FECHA_DIA, ALL_FBG, ALL_NUMSER, ALL_NUMFAC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT PAR_NOMBRE FROM PARGEN WHERE PAR_CODCIA = ? "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)

xcuenta = 1
f1 = 6
f1 = f1 + 1
xl.Cells(f1, 1) = "MAS  : INGRESOS"
f1 = f1 + 1
xl.Cells(f1, 1) = 1
xl.Cells(f1, 2) = "VENTAS AL CONTADO:"
WINGRESOS = 0
TOT_INGRESOS = 0
For fila = 1 To 30
   ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
   If Trim(ws_codcia) = "00" Then GoTo OTRA_CIA
   If Trim(ws_codcia) = "" Then Exit For
    WINGRESOS = 0
    PS_REP01(0) = ws_codcia
    PS_REP01(1) = wsFECHA1
    PS_REP02(0) = ws_codcia
    llave_rep02.Requery
    f1 = f1 + 1
    xl.Cells(f1, 2) = Trim(llave_rep02!par_nombre)
    ' SOLO PARA FACTURAS
    PS_REP01(2) = "F"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      ' xl.Cells(f1, 3) = "F/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "FACTURAS"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
    'SOLO PARA BOLETAS
    PS_REP01(2) = "B"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      ' xl.Cells(f1, 3) = "B/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "BOLETAS"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
    ' SOLO PARA GUIAS
    PS_REP01(2) = "G"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      'xl.Cells(f1, 3) = "G/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "GUIAS"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
    ' SOLO PARA VENTAS ADMINISTRADORES P
    PS_REP01(2) = "P"
    llave_rep01.Requery
    wmonto = 0
    If Not llave_rep01.EOF Then
       docus = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
      Do Until llave_rep01.EOF
        wmonto = wmonto + llave_rep01!ALL_IMPORTE_AMORT
        docus2 = Trim(llave_rep01!ALL_NUMSER) & "-" & llave_rep01!all_numfac
        llave_rep01.MoveNext
      Loop
      f1 = f1 + 1
      xl.Cells(f1, 3) = "P/. " & docus & " AL " & docus2
      xl.Cells(f1, 3) = "ADMINISTRACION"
      xl.Cells(f1, 4) = wmonto
      WINGRESOS = WINGRESOS + wmonto
    End If
      f1 = f1 + 1
      xl.Cells(f1, 4) = WINGRESOS
      TOT_INGRESOS = TOT_INGRESOS + WINGRESOS
OTRA_CIA:
   xcuenta = xcuenta + 2
 Next fila
 f1 = f1 + 1
 xl.Cells(f1, 5) = TOT_INGRESOS
 WSALDO_CAJA = WSALDO_CAJA + TOT_INGRESOS

DoEvents

DoEvents
xcuenta = 1
xl.Cells(2, 2) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
xl.Cells(2, 5) = "'" & Format(wsFECHA1, "dd mmm yyyy")
xl.DisplayAlerts = False
xl.Worksheets("Hoja1").Range("A1:X51").Locked = True
xl.Worksheets("Hoja1").Protect PUB_CLAVE
xl.Application.Visible = True
DoEvents
Set xl = Nothing
Screen.MousePointer = 0


Exit Sub


WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  
  DoEvents
  
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\CAJA.xls", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

Exit Sub

'LLENA_VALOR:
'For I = QJ To WDIF
'  If I >= 3 Then
'    xl.Cells(F1 + 7, I) = Format(LOC_VALOR, "0.00")
'  End If
'Next I
'Return

Exit Sub
CANCELA:
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FORM_PED
Exit Sub



End Sub

Private Sub codpro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul txtdias, txtdias
End If
End Sub

Private Sub codpro_LostFocus()

SQ_OPER = 1
pu_codclie = Val(Right(codpro.Text, 8))
pu_cp = "P"
pu_codcia = LK_CODCIA
LEER_CLI_LLAVE

If cli_llave.EOF Then Exit Sub
txtcontacto.Text = cli_llave!CLI_NOMBRE_ESPOSA


End Sub


Private Sub fami_Click()
Dim wpos As Integer
Dim WFAMI2 As Integer
If opcomo(1) = True Then
    lisarti.Clear
    If Trim(fami.Text) = "" Then
     subfami.Clear
     Exit Sub
    End If
    Screen.MousePointer = 11
    wpos = fami.ListIndex
    WFAMI2 = Val(Trim(Right(fami.Text, 6)))
    LLENADO_SUBFAM WFAMI2
    On Error GoTo sigue
    subfami.ListIndex = wpos
    Screen.MousePointer = 0
    Exit Sub
sigue:
    Resume Next
End If
If opcomo(0) = True Then
 lisarti.Clear
 LLENA_ARTI Val(Right(fami.Text, 6)), 0
End If

End Sub

Private Sub fecha1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   fecha2.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim xcuenta As Integer
'On Error GoTo SALE
checia.Visible = False
'If LK_EMP = "3AA" Then
' checia.Visible = True
'End If

Wsec = 0
LOC_CANCELA = 0
fila = 0
wfila_act = 0
WSELE = ""
Dim ws_indice As Integer
Dim cade
WMODO = ""

pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP = 'P'  AND CLI_CODCIA = ? ORDER BY CLI_NOMBRE"
Set PSPRO_V = CN.CreateQuery("", pub_cadena)
PSPRO_V.rdoParameters(0) = " "
Set PRO_V = PSPRO_V.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PSPRO_V(0) = LK_CODCIA
PRO_V.Requery
codpro.Clear
Do Until PRO_V.EOF
    codpro.AddItem PRO_V!cli_nombre & String(60, " ") & PRO_V!cli_codclie
    PRO_V.MoveNext
Loop

pub_cadena = "SELECT * FROM TRANSPORTE ORDER  BY TRN_NOMBRE"
Set PSPRO_V = CN.CreateQuery("", pub_cadena)
Set PRO_V = PSPRO_V.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PRO_V.Requery
txtagencia.Clear
Do Until PRO_V.EOF
    txtagencia.AddItem PRO_V!TRN_NOMBRE & String(80, " ") & PRO_V!TRN_KEY
    PRO_V.MoveNext
Loop


pub_cadena = "SELECT SUM(FAR_CANTIDAD)AS CANTIDAD FROM FACART WHERE (FAR_CODCIA = ? or FAR_CODCIA = ? or FAR_CODCIA = ?) AND FAR_CODART = ? and (FAR_FECHA >= ? AND FAR_FECHA <= ?) AND FAR_TIPMOV = 10  AND FAR_ESTADO  <> 'E' GROUP BY FAR_CODART"
Set PSLOC_WARTI = CN.CreateQuery("", pub_cadena)
PSLOC_WARTI.rdoParameters(0) = 0
PSLOC_WARTI.rdoParameters(1) = 0
PSLOC_WARTI.rdoParameters(2) = 0
PSLOC_WARTI.rdoParameters(3) = 0
PSLOC_WARTI.rdoParameters(4) = LK_FECHA_DIA
PSLOC_WARTI.rdoParameters(5) = LK_FECHA_DIA
Set llave_sum_arti = PSLOC_WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)



pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_NUMSER = ? AND PED_TIPMOV= 20  ORDER BY  PED_NUMFAC DESC "
Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
PSTEMP_MAYOR.rdoParameters(0) = " "
PSTEMP_MAYOR.rdoParameters(1) = 0

PSTEMP_MAYOR.MaxRows = 1
Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA= ? AND PED_NUMSER = ? AND PED_NUMFAC= ? AND PED_TIPMOV = 20 ORDER BY PED_CODCIA"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP_LLAVE.rdoParameters(0) = " "
PSTEMP_LLAVE.rdoParameters(1) = 0
PSTEMP_LLAVE.rdoParameters(2) = 0
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
temp_llave.Requery


pub_cadena = "SELECT ART_KEY  FROM ARTI WHERE ART_CODCIA = ? AND ART_ALTERNO =?"
Set PS_REP04 = CN.CreateQuery("", pub_cadena)
PS_REP04.rdoParameters(0) = ""
PS_REP04.rdoParameters(1) = ""
Set llave_rep04 = PS_REP04.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >=? AND COV_FECHA_VOUCHER <=?  ORDER BY COV_CODCIA, COV_NRO_VOUCHER, COV_DH" ', COV_NRO_MOV"
Set PSCOV_MAYOR = CN.CreateQuery("", pub_cadena)
PSCOV_MAYOR.rdoParameters(0) = " "
PSCOV_MAYOR.rdoParameters(1) = LK_FECHA_DIA
PSCOV_MAYOR.rdoParameters(2) = LK_FECHA_DIA
Set cov_mayor = PSCOV_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ?  AND COV_FECHA_VOUCHER=? AND COV_NRO_VOUCHER = ? AND COV_NRO_MOV = ? ORDER BY COV_CODCIA, COV_FECHA_VOUCHER, COV_NRO_VOUCHER, COV_NRO_MOV"
Set PSCOV_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOV_LLAVE(0) = " "
PSCOV_LLAVE(1) = LK_FECHA_DIA
PSCOV_LLAVE(2) = 0
PSCOV_LLAVE(3) = 0
Set cov_llave = PSCOV_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COPARAM WHERE COP_CODCIA = ?"
Set PSCOP_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOP_LLAVE.rdoParameters(0) = " "
Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
'If LK_EMP = "3AA" Then
'   Combo1.AddItem "AV. PROLONGACION UNION N. 2218 - TRUJILLO"
'   Combo1.AddItem "AV. SANCHEZ CERRO N. 1675 - PIURA"
'   Combo1.ListIndex = 0
'End If
''If LK_EMP_PTO = "A" Then
'' PSCOP_LLAVE.rdoParameters(0) = "00"
''Else
 PSCOP_LLAVE.rdoParameters(0) = par_llave!PAR_CIACON
'End If

cop_llave.Requery
PUB_CODCIA = LK_CODCIA
LLENA_COMBO fpago, 11
fila = 0
DoEvents
LIMPIA_DATOS
CABE_MAN
tserie.Text = 0


 LISCIA.Visible = True
 LISCIA.Clear
 xcuenta = 0
 For fila = 1 To 30 Step 2
   PUB_CODCIA = Mid(Trim(par_llave!par_art_cias), fila, 2)
   If Trim(PUB_CODCIA) = "" Then Exit For
   xcuenta = xcuenta + 1
   PSPAR_MULTI(0) = PUB_CODCIA
   par_multi.Requery
   LISCIA.AddItem PUB_CODCIA & " - " & Trim(par_multi!par_nombre)
   Combo1.AddItem Trim(par_multi!PAR_DIR_EMP) ' "AV. SANCHEZ CERRO N. 1675 - PIURA"
 Next fila
 
 For fila = 0 To LISCIA.ListCount - 1
  LISCIA.ListIndex = fila
  If Left(LISCIA.Text, 2) = LK_CODCIA Then LISCIA.Selected(fila) = True
 Next fila

imax.Text = Format(GEN!GEN_TASA_LEG_VENC, "##")
fecha1.Mask = "##/##/####"
fecha1.Text = "__/__/____"
fecha1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
fecha2.Mask = "##/##/####"
fecha2.Text = "__/__/____"
fecha2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")


Exit Sub
sale:
MsgBox "Depurar: " & Err.Description, 48, Pub_Titulo
Resume Next
End Sub

Private Sub fpago_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 45 Then
 Exit Sub
End If
Dim wpos
wpos = fpago.ListIndex
PUB_TIPREG = Mid(fpago.ToolTipText, 13, Len(fpago.ToolTipText))
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_COMBO fpago, 11
On Error GoTo sigue
fpago.ListIndex = wpos
On Error GoTo 0
fpago.SetFocus
SendKeysSeguro VK_UP, True
Exit Sub
sigue:
Resume Next


End Sub

Private Sub grid_fac_DblClick()
Exit Sub
End Sub

Private Sub grid_fac_GotFocus()
fraListar.Visible = False
End Sub

Private Sub grid_fac_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub
If grid_fac.Rows <= 2 Then Exit Sub
If grid_fac.COL = 1 Then Exit Sub
If grid_fac.COL >= 6 Then Exit Sub


'GoTo leer
'Exit Sub
'End If
If WMODO = "I" Then
    'If Trim(grid_fac.TextMatrix(grid_fac.Row, 0)) = "" Then Exit Sub
    'If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) <> "" And grid_fac.Col = 2 Or grid_fac.Col = 3 Then GoTo leer
    'If Trim(grid_fac.TextMatrix(grid_fac.Row, 8)) <> "0" Then Exit Sub
End If


'If grid_fac.Col = 1 And WMODO = "I" Then
'   a = Val(grid_fac.TextMatrix(grid_fac.Row - 1, 0))
'   a = a + 1
'  grid_fac.TextMatrix(grid_fac.Row, 0) = a
'End If

If WMODO = "I" Or FLAG_MODI = 1 Then
  If grid_fac.COL = 3 Then
    unidad.Left = grid_fac.Left + grid_fac.CellLeft
    unidad.Width = grid_fac.CellWidth
    'UNIDAD.Height = grid_fac.CellHeight
    unidad.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 1200 '480
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 9))
    LEER_PRE_LLAVE
    unidad.Clear
    unidad.Visible = True
    Do Until pre_mayor.EOF
     unidad.AddItem Trim(pre_mayor!PRE_UNIDAD) & String(30, " ") & pre_mayor!PRE_SECUENCIA
     pre_mayor.MoveNext
    Loop
    'wfila_act = grid_fac.Row
    On Error GoTo pasa
    unidad.ListIndex = Val(grid_fac.TextMatrix(grid_fac.Row, 12))
    On Error GoTo 0
    unidad.Visible = True
    
    unidad.SetFocus
    SendKeysSeguro VK_UP, True
     Exit Sub
  End If
    TEXTOVAR.Left = grid_fac.Left + grid_fac.CellLeft
    TEXTOVAR.Width = grid_fac.CellWidth
    TEXTOVAR.Height = grid_fac.CellHeight
    TEXTOVAR.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 2025 '480
    TEXTOVAR.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
    wfila_act = grid_fac.Row
    TEXTOVAR.Visible = True
    Azul3 TEXTOVAR, TEXTOVAR
    TEXTOVAR.SetFocus
End If
Exit Sub
pasa:
Resume Next
End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim sale As Boolean

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


Private Sub gridarti_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub

'If Trim(gridarti.TextMatrix(gridarti.Row, 9)) <> "8" Then
'  If Trim(gridarti.TextMatrix(gridarti.Row, 0)) = "" Then Exit Sub
'  If Trim(gridarti.TextMatrix(gridarti.Row, 1)) <> "" And gridarti.Col = 2 Or gridarti.Col = 3 Then GoTo leer
'  If Trim(gridarti.TextMatrix(gridarti.Row, 8)) <> "0" Then Exit Sub
'End If


If gridarti.COL <> 9 And gridarti.COL <> 12 Then Exit Sub
veri_suma
'   a = Val(gridarti.TextMatrix(gridarti.Row - 1, 0))
'   a = a + 1
'  gridarti.TextMatrix(gridarti.Row, 0) = a
'End If
'If WMODO = "I" Or WMODO = "C" Then
    TEXTOVARPRO.Left = gridarti.Left + gridarti.CellLeft
    TEXTOVARPRO.Width = gridarti.CellWidth
    TEXTOVARPRO.Height = gridarti.CellHeight
    TEXTOVARPRO.Top = gridarti.Top + gridarti.CellTop
    TEXTOVARPRO.Text = gridarti.TextMatrix(gridarti.Row, gridarti.COL)
    TEXTOVARPRO.Visible = True
    Azul3 TEXTOVARPRO, TEXTOVARPRO
    TEXTOVARPRO.SetFocus
'End If
End Sub

Private Sub gridarti_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim sale As Boolean
Dim Wsec

'If WMODO = "C" Then Exit Sub
If KeyCode = 46 Then
  If gridarti.COL = 9 Then
    gridarti.Text = ""
  End If
End If


'If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
'End If


If Left(gridarti.TextMatrix(gridarti.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
  'If WMODO <> "C" Then Exit Sub
  tC = gridarti.COL
  For fila = 1 To gridarti.Cols - 1
      gridarti.COL = fila
      If gridarti.CellBackColor = QBColor(12) Then
         gridarti.CellBackColor = QBColor(15)
         gridarti.TextMatrix(gridarti.Row, 9) = "9"
      Else
         gridarti.CellBackColor = QBColor(12)
         gridarti.TextMatrix(gridarti.Row, 9) = "-1"
      End If
  Next fila
  gridarti.COL = tC
  gridarti.SetFocus
  Exit Sub
End If
If KeyCode = 45 Then
    Wsec = Wsec + 1
    If Trim(gridarti.TextMatrix(gridarti.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(gridarti.TextMatrix(gridarti.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(gridarti.TextMatrix(gridarti.Row, 4)) = 0 And Val(gridarti.TextMatrix(gridarti.Row, 5)) = 0 Then Exit Sub
    gridarti.AddItem "", gridarti.Row + 1
    gridarti.TextMatrix(gridarti.Row + 1, 0) = "MAN. " & Format(gridarti.TextMatrix(gridarti.Row, 10), "dd/mm/yyyy")
    gridarti.TextMatrix(gridarti.Row + 1, 6) = Wsec
    gridarti.TextMatrix(gridarti.Row + 1, 8) = gridarti.TextMatrix(gridarti.Row, 8)
    gridarti.TextMatrix(gridarti.Row + 1, 3) = gridarti.TextMatrix(gridarti.Row, 3)
    gridarti.TextMatrix(gridarti.Row + 1, 7) = gridarti.TextMatrix(gridarti.Row, 7)
    gridarti.TextMatrix(gridarti.Row + 1, 10) = gridarti.TextMatrix(gridarti.Row, 10)
    gridarti.TextMatrix(gridarti.Row + 1, 11) = "8"
    gridarti.Row = gridarti.Row + 1
    gridarti.COL = 1
    gridarti.SetFocus
End If
Exit Sub
If KeyCode = 46 Then
If gridarti.Rows <= 3 Then
Else
   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(gridarti.TextMatrix(gridarti.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     gridarti.SetFocus
     Exit Sub
   Else
     gridarti.RowHeight(gridarti.Row) = 1
     gridarti.Row = gridarti.Row + 1
    
   'gridarti.RemoveItem (gridarti.Row)
   'gridarti.Refresh
   gridarti.SetFocus
   End If
End If
End If
'gridarti.SetFocus
Exit Sub



End Sub

Private Sub gridarti_Scroll()
TEXTOVARPRO.Visible = False
End Sub


Private Sub TEXTOVARPRO_Change()
gridarti.Text = Format(TEXTOVARPRO.Text, "0.0000")
End Sub

Private Sub TEXTOVARPRO_GotFocus()
 temporal = gridarti.TextMatrix(gridarti.Row, gridarti.COL)
End Sub

Private Sub TEXTOVARPRO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  TEXTOVARPRO.Text = temporal
  TEXTOVARPRO.Visible = False
  gridarti.SetFocus
  Exit Sub
End If
'If gridarti.COL = 1 Then Consistencias gridarti, TEXTOVARPRO, KeyAscii
If gridarti.COL = 9 Then Consistencias gridarti, TEXTOVARPRO, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If
If gridarti.COL = 1 Or gridarti.COL = 4 Then
  If Val(TEXTOVARPRO.Text) > 99 Then
    Azul3 TEXTOVARPRO, TEXTOVARPRO
    Exit Sub
  End If
End If



' COLOCAR EDICION


If gridarti.Row >= gridarti.Rows - 1 Then
Else
  gridarti.Row = gridarti.Row + 1
End If
gridarti.SetFocus
TEXTOVARPRO.Visible = False

fin:

End Sub



Private Sub i_codart2_Change()
If i_codart2.Text = "" Then
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
  DoEvents
fin:

End Sub
Private Sub i_codart2_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I, car
Dim itmFound As MSComctlLib.ListItem
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView1.Visible = False
 i_codart2.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
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
    If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       GoTo fin
    End If
    WCOD_ORIGINAL = art_LLAVE!art_key
    'i_codart2.text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
    lisarti.AddItem art_LLAVE!ART_alterno & " " & Trim(art_LLAVE!art_nombre) & String(60, " ") & art_LLAVE!art_key
    lisarti.Selected(lisarti.ListCount - 1) = True
    i_codart2.Text = ""
    ListView1.Visible = False
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
     If art_llave_alt!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       GoTo fin
     End If
     WCOD_ORIGINAL = art_llave_alt!art_key
     'i_codart2.text = Trim(art_llave_alt!ART_NOMBRE)
     lisarti.AddItem art_llave_alt!ART_alterno & " " & Trim(art_llave_alt!art_nombre) & String(60, " ") & art_llave_alt!art_key
     lisarti.Selected(lisarti.ListCount - 1) = True
     ListView1.Visible = False
     i_codart2.Text = ""
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
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul i_codart2, i_codart2
        Exit Sub
      End If
      If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       GoTo fin
      End If
      WCOD_ORIGINAL = art_LLAVE!art_key
      lisarti.AddItem art_LLAVE!ART_alterno & " " & Trim(art_LLAVE!art_nombre) & String(60, " ") & art_LLAVE!art_key
      lisarti.Selected(lisarti.ListCount - 1) = True
      i_codart2.Text = ""
      ListView1.Visible = False
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
    var = Asc(i_codart2.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND (PRE_FLAG_UNIDAD ='A') AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV , ART_SITUACION FROM ARTI, ARTICULO, PRECIOS WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND (PRE_FLAG_UNIDAD ='A') AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      
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

Private Sub i_codart2_LostFocus()
' ListView1.Visible = False
End Sub



Private Sub imax_LostFocus()
GEN.Requery
 GEN.Edit
 GEN!GEN_TASA_LEG_VENC = Val(imax.Text)
 GEN.Update
GEN.Requery
End Sub

Private Sub Imprimir_Click()
On Error GoTo FINTODO
Dim wser As String * 3
Dim WSRUTA As String
Dim wRuta As String
Dim rmoneda As String * 1
wRuta = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\"
Reportes.Connect = PUB_ODBC
Reportes.Destination = crptToWindow  '= crptToPrinter
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 800
Reportes.WindowHeight = 490
Reportes.Formulas(1) = ""
PUB_NETO = Val(txttotal.Text)
PUB_FECHA = fecha.Caption
PU_NUMSER = Val((tserie.Text))

If Left(moneda.Text, 1) = "D" Then
   rmoneda = "D"
Else
   rmoneda = "S"
End If

PU_NUMFAC = Val((txtdoc.Text))
Reportes.Formulas(1) = "SON_EFECTIVO=  'SON: " & CONVER_LETRAS(PUB_NETO, rmoneda) & "'"
If cheprecios.Value = 0 Then
  Reportes.WindowTitle = "ORDEN DE COMPRA  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
 Reportes.ReportFileName = wRuta + "ORDEN.RPT"
 wser = PU_NUMSER
 pub_cadena = "{PEDIDOS.PED_ESTADO} = 'N' AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND {PEDIDOS.PED_NUMSER}= '" & wser & "' AND {PEDIDOS.PED_NUMFAC} = " & PU_NUMFAC
 Reportes.SelectionFormula = pub_cadena
Else
 Reportes.WindowTitle = "COTIZACION: " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
 Reportes.ReportFileName = wRuta + "ORDEN2.RPT"
 wser = PU_NUMSER
 pub_cadena = "{PEDIDOS.PED_ESTADO} = 'N' AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND {PEDIDOS.PED_NUMSER}= '" & wser & "' AND {PEDIDOS.PED_NUMFAC} = " & PU_NUMFAC
 Reportes.SelectionFormula = pub_cadena
End If

Reportes.WindowTitle = Reportes.WindowTitle & " Archivo: " & Trim(Reportes.ReportFileName)
On Error GoTo accion
Reportes.Action = 1
On Error GoTo 0
Exit Sub

FINTODO:
accion:
 MsgBox Err.Description, 48, Pub_Titulo
 MsgBox "Reintente Nuevamente ..", 48, Pub_Titulo

End Sub

Private Sub lisarti_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If KeyCode = 113 Then
  For I = 0 To lisarti.ListCount - 1
   lisarti.ListIndex = I
   lisarti.Selected(I) = True
  Next I
End If
If KeyCode = 114 Then
  For I = 0 To lisarti.ListCount - 1
   lisarti.ListIndex = I
   lisarti.Selected(I) = False
  Next I
End If
End Sub

Private Sub ListView1_DblClick()
 loc_key = ListView1.SelectedItem.Index
 i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 i_codart2_KeyPress 13
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
 i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView1.Visible = False
 i_codart2.Text = ""
 i_codart2.SetFocus
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

Private Sub modif_Click()
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
FLAG_MODI = 1
tserie.Enabled = True
txtdoc.Enabled = True
FORM_PED.MODIF.Enabled = False

f1.Enabled = True
cmdIngreso.Caption = "&Grabar"
ESTADO.Enabled = True
tserie.SetFocus


End Sub

Private Sub moneda_Click()
If Left(moneda.Text, 1) = "S" Then
 lblMoneda.Caption = "S/."
Else
 lblMoneda.Caption = "US$."
End If
End Sub

Private Sub MSChart1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 FORM_PED.MSChart1.Visible = False
End If
End Sub

Private Sub MSChart1_LostFocus()
 FORM_PED.MSChart1.Visible = False
End Sub

Private Sub opcomo_Click(Index As Integer)
Select Case Index
Case 0
  lblarti.Visible = False
  i_codart2.Visible = False
  fami.Visible = True
  lfami.Visible = True
  subfami.Clear
  subfami.Enabled = False
  fami.Clear
  fami.Enabled = True
  PUB_CODCIA = LK_CODCIA
  LLENADOS fami, 122
Case 1
  lblarti.Visible = False
  i_codart2.Visible = False
  fami.Visible = True
  lfami.Visible = True
  subfami.Clear
  subfami.Enabled = True
  fami.Clear
  fami.Enabled = True
  PUB_CODCIA = LK_CODCIA
  LLENADOS fami, 122
Case 2
  lblarti.Visible = False
  i_codart2.Visible = False
  fami.Visible = True
  lfami.Visible = True
  subfami.Clear
  subfami.Enabled = False
  fami.Clear
  fami.Enabled = False
  LLENA_ARTI_PRO Val(Right(codpro.Text, 10))
Case 3
  fami.Visible = False
  lfami.Visible = False
  lblarti.Visible = True
  lisarti.Clear
  i_codart2.Visible = True
  i_codart2.Text = ""
  i_codart2.SetFocus
End Select


End Sub

Private Sub opcomo_DblClick(Index As Integer)
opcomo_Click Index
End Sub

Private Sub salir_Click()
Unload FORM_PED
End Sub


Public Sub LIMPIA_DATOS()
cheprecios.Value = 0
codpro.ListIndex = -1
'tserie.text = ""
txtdoc.Text = ""
grid_fac.Clear
'txtagencia.text = ""
txtigv.Text = ""
txtvalorv.Text = ""
txttotal.Text = ""
lblpeso.Caption = ""

End Sub

Public Sub CABE_MAN()
grid_fac.Cols = 20
grid_fac.Rows = 2
grid_fac.Clear
fila = 0
grid_fac.ColWidth(0) = 3500
grid_fac.ColWidth(1) = 900
grid_fac.ColWidth(2) = 900
grid_fac.ColWidth(3) = 550
grid_fac.ColWidth(4) = 900
grid_fac.ColWidth(5) = 1000
grid_fac.ColWidth(6) = 800
grid_fac.ColWidth(7) = 800
grid_fac.ColWidth(8) = 900
grid_fac.ColWidth(9) = 0
grid_fac.ColWidth(10) = 0
grid_fac.ColWidth(11) = 0
grid_fac.ColWidth(12) = 0
grid_fac.ColWidth(13) = 0
grid_fac.ColWidth(14) = 0
grid_fac.ColWidth(15) = 0
grid_fac.ColWidth(16) = 0
grid_fac.ColWidth(17) = 0
grid_fac.ColWidth(18) = 0

grid_fac.TextMatrix(0, 0) = "Articulo"
grid_fac.TextMatrix(0, 1) = "Codigo"
grid_fac.TextMatrix(0, 2) = "Cantidad"
grid_fac.TextMatrix(0, 3) = "Unidad"
grid_fac.TextMatrix(0, 4) = "Precios"
grid_fac.TextMatrix(0, 5) = "Sub Total"
grid_fac.TextMatrix(0, 6) = "Prom. Ventas"
grid_fac.TextMatrix(0, 7) = "Stock Actual"
grid_fac.TextMatrix(0, 8) = "Stock Min"

End Sub
Public Sub suma_grid()
'On Error GoTo SALE
Dim WF As Integer
Dim ww_subt As Currency
Dim SUM_PESO As Currency
WF = 2
Dim fx As Integer
fx = 1
SUM_H = 0
SUM_D = 0
SUM_PESO = 0
If grid_fac.Rows <= WF Then Exit Sub
Do While fx = 1
    If grid_fac.TextMatrix(WF, 0) <> "" Then
      If Val(grid_fac.TextMatrix(WF, 19)) <> 5 Then
         ww_subt = Format(Val(Val(grid_fac.TextMatrix(WF, 2)) * Val(grid_fac.TextMatrix(WF, 4))), "0.00")
         grid_fac.TextMatrix(WF, 5) = ww_subt
      Else
         ww_subt = Val(grid_fac.TextMatrix(WF, 5))
         If Val(grid_fac.TextMatrix(WF, 2)) <> 0 Then
         grid_fac.TextMatrix(WF, 4) = Format(Val(Val(grid_fac.TextMatrix(WF, 5)) / Val(grid_fac.TextMatrix(WF, 2))), "0.0000")
         End If
      End If
      SUM_H = SUM_H + ww_subt
      SUM_D = SUM_D + Val(grid_fac.TextMatrix(WF, 4))
      SUM_PESO = SUM_PESO + (Val(grid_fac.TextMatrix(WF, 18)) * Val(grid_fac.TextMatrix(WF, 2)))
    End If
    WF = WF + 1
    If WF = grid_fac.Rows Then
        fx = 0
    Else
        If Trim(grid_fac.TextMatrix(WF, 0)) = "" Then fx = 0
    End If
Loop
   fila = WF - 1
   grid_fac.TextMatrix(1, 5) = Format(SUM_H, "####0.00")
   lblpeso.Caption = Format(SUM_PESO, "#,##0.00")
  ' txttotal.text = Format(SUM_H, "###,##0.00")
  ' txtigv.text = Format((SUM_H / LK_IGV), "###,##0.00")
   'txtvalorv.text = Format(Val(txttotal.text) - Val(txtigv.text), "###,##0.00")
   If LK_EMP = "HER" Then
    txtvalorv.Text = SUM_H
    txtigv.Text = Format(SUM_H * ((LK_IGV / 100)), "0.00")
    txttotal.Text = Format(SUM_H * (1 + (LK_IGV / 100)), "0.00")  ' Val(Format(SUM_H, "#####0.00")) + Val(Format(txtigv.Text, "#####0.00"))
   Else
    txtigv.Text = SUM_H - Format((SUM_H / (1 + (LK_IGV / 100))), "####0.00")
    txttotal.Text = Format(SUM_H, "#####0.00")
    txtvalorv.Text = Val(txttotal.Text) - Val(txtigv.Text)
   End If
   
   
Exit Sub
sale:
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
grid_fac.Cols = 6
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

Private Sub SALIR_GotFocus()
fraListar.Visible = False
End Sub

Private Sub subfami_Click()
If Trim(subfami.Text) = "" Then
Exit Sub
End If
LLENA_ARTI 0, Val(Right(subfami.Text, 6))
End Sub
Private Sub textovar_Change()
If grid_fac.COL = 1 Then
Else
 If grid_fac.COL = 2 Then
  grid_fac.Text = TEXTOVAR.Text
 Else
  grid_fac.Text = Format(TEXTOVAR.Text, "0.00")
 End If
 If grid_fac.COL = 5 Then grid_fac.TextMatrix(grid_fac.Row, 19) = grid_fac.COL
End If
suma_grid
suma_subtotal

End Sub

Private Sub TEXTOVAR_GotFocus()
temporal = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
'SOLO_DECIMAL TEXTOVAR, KeyAscii
If KeyAscii = 27 Then
  TEXTOVAR.Text = temporal
  TEXTOVAR.Visible = False
  grid_fac.SetFocus
  Exit Sub
End If
If grid_fac.COL = 2 Or grid_fac.COL = 4 Then Consistencias grid_fac, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then Exit Sub

TEXTOVAR.Visible = False
grid_fac.SetFocus

fin:

End Sub

Private Sub textovar_LostFocus()
'TEXTOVAR.Visible = False
If TEXTOVAR.Visible Then
   TEXTOVAR.Visible = False
   grid_fac.Row = wfila_act
'   grid_fac.SetFocus
   Exit Sub
End If

End Sub

Public Sub LLENADO_SUBFAM(wfami As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = 123
    PUB_CODCIA = LK_CODCIA
    PUB_CODART = wfami
    SQ_OPER = 3
    LEER_TAB_LLAVE
    subfami.ToolTipText = "TAB_TIPREG = 123"
    subfami.Clear
    Do Until tab_menor.EOF
        DoEvents
        subfami.AddItem tab_menor!TAB_NOMLARGO & String(50, " ") & Trim(CStr(tab_menor!TAB_NUMTAB))
        DoEvents
        CONTA = CONTA + 1
        tab_menor.MoveNext
    Loop
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
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENA_COMBO(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    cont.AddItem " "
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub



Public Sub LLENA_ARTI_PRO(wcodclie As Currency)
Dim WARTI As rdoQuery
Dim wllave_arti As rdoResultset

pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_CODCLIE = ? AND ART_CALIDAD = 1 ORDER BY ART_NOMBRE "
Set WARTI = CN.CreateQuery("", pub_cadena)
WARTI.rdoParameters(0) = " "
WARTI.rdoParameters(1) = 0
Set wllave_arti = WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)
WARTI(0) = LK_CODCIA
WARTI(1) = wcodclie
wllave_arti.Requery
lisarti.Clear
If wllave_arti.EOF Then
  lisarti.Clear
  Exit Sub
End If
Do Until wllave_arti.EOF
  lisarti.AddItem wllave_arti!ART_alterno & " " & wllave_arti!art_nombre & String(120, " ") & wllave_arti!art_key
wllave_arti.MoveNext
Loop


End Sub
Public Sub LLENA_ARTI(wfami As Integer, WSUBFAMI As Integer)
Dim xpinta As Integer
Dim xfil  As Integer
Dim WARTI As rdoQuery
Dim wllave_arti As rdoResultset

Dim WBOT As rdoQuery
Dim wllave_bot As rdoResultset

Dim wvalor As Integer
Dim CADE_ROT As String
  
pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODCIA = ? AND ARM_CODART = ? "
Set WBOT = CN.CreateQuery("", pub_cadena)
WBOT(0) = 0
WBOT(1) = 0
Set wllave_bot = WBOT.OpenResultset(rdOpenKeyset, rdConcurValues)
If LK_CODCIA = "01" Then
 WBOT(0) = "04"
Else
 WBOT(0) = LK_CODCIA
End If
CADE_ROT = ""
If cherot.Value = 1 Then
  CADE_ROT = " AND art_plancha = 1 "
End If

If wvalor = 0 Then
  pub_cadena = "SELECT ARM_COSPRO, ART_STOCK_MAX , PRE_COSTO_ANT , ART_KEY , art_alterno, art_nombre, PRE_EQUIV ,PRE_PRE6 PRE_UNIDAD, ARM_STOCK  , ((ARM_STOCK / PRE_EQUIV) - (ART_STOCK_MAX / PRE_EQUIV)) AS DIF , ART_PLANCHA  FROM ARTI, articulo, precios WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART ) AND PRE_FLAG_UNIDAD = 'A' AND ART_CODCIA = ? AND ART_CALIDAD = 1 AND ART_SITUACION <> '1' AND ART_FLAG_STOCK = 'M' " & CADE_ROT & " ORDER BY DIF ASC   "
End If
If wfami <> 0 Then
  pub_cadena = "SELECT ARM_COSPRO, ART_STOCK_MAX , PRE_COSTO_ANT , ART_KEY , art_alterno, art_nombre, PRE_EQUIV ,PRE_PRE6, PRE_UNIDAD, ARM_STOCK  , ((ARM_STOCK / PRE_EQUIV) - (ART_STOCK_MAX / PRE_EQUIV)) AS DIF , ART_PLANCHA FROM ARTI, articulo, precios WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART ) AND PRE_FLAG_UNIDAD = 'A' AND ART_CODCIA = ? AND ART_FAMILIA = ? AND ART_CALIDAD = 1 AND ART_SITUACION <> '1' and ART_FLAG_STOCK = 'M' " & CADE_ROT & " ORDER BY DIF ASC   "
  wvalor = wfami
ElseIf WSUBFAMI <> 0 Then
  pub_cadena = "SELECT ARM_COSPRO, ART_STOCK_MAX ,   PRE_COSTO_ANT, ART_KEY , art_alterno, art_nombre, PRE_EQUIV ,PRE_PRE6 , PRE_UNIDAD, ARM_STOCK , ((ARM_STOCK/PRE_EQUIV) - (ART_STOCK_MAX/PRE_EQUIV)) AS DIF , ART_PLANCHA FROM ARTI, articulo, precios WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART ) AND PRE_FLAG_UNIDAD = 'A' AND ART_CODCIA = ? AND ART_SUBFAM = ? AND ART_CALIDAD = 1 AND ART_SITUACION <> '1' AND ART_FLAG_STOCK = 'M'  " & CADE_ROT & " ORDER BY DIF ASC "
  wvalor = WSUBFAMI
End If
Set WARTI = CN.CreateQuery("", pub_cadena)
WARTI(0) = " "
If wvalor <> 0 Then WARTI(1) = 0
Set wllave_arti = WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)
WARTI(0) = LK_CODCIA
If wvalor <> 0 Then WARTI(1) = wvalor
wllave_arti.Requery
lisarti.Clear
If wllave_arti.EOF Then
  lisarti.Clear
  Exit Sub
End If
probar.Value = 0
probar.Min = 0
probar.max = wllave_arti.RowCount
probar.Visible = True
gridarti.Visible = False
DoEvents
gridarti.Clear
If cheprecios.Value = 0 Then
 gridarti.Cols = 12
Else
 gridarti.Cols = 13
End If
gridarti.TextMatrix(0, 0) = "Item"
gridarti.TextMatrix(0, 1) = "Descripcion del Producto"
gridarti.TextMatrix(0, 2) = "Unid."
gridarti.TextMatrix(0, 2) = "Activa"
gridarti.TextMatrix(0, 3) = "Costo"
gridarti.TextMatrix(1, 3) = "Repos."
gridarti.TextMatrix(0, 4) = "Costo"
gridarti.TextMatrix(1, 4) = "Ptromd"
gridarti.TextMatrix(0, 5) = "Stock "
gridarti.TextMatrix(1, 5) = "Actual "

gridarti.TextMatrix(0, 6) = "Stock "
gridarti.TextMatrix(1, 6) = "Maximo"

gridarti.TextMatrix(0, 7) = "Stock "
gridarti.TextMatrix(1, 7) = "Difer."

gridarti.TextMatrix(0, 8) = "Stock "
gridarti.TextMatrix(1, 8) = "Botica"
If cheprecios.Value = 0 Then
  gridarti.TextMatrix(0, 9) = "Cantidad"
  gridarti.TextMatrix(1, 9) = "Solicitar"
Else
  gridarti.TextMatrix(0, 9) = "Precio a "
  gridarti.TextMatrix(1, 9) = "Cotizar"
  gridarti.TextMatrix(1, 12) = "Oferta"
  gridarti.ColWidth(12) = 2000
End If

gridarti.ColWidth(0) = 300
gridarti.ColWidth(1) = 2500
gridarti.ColWidth(2) = 500
gridarti.ColWidth(3) = 700
gridarti.ColWidth(4) = 700
gridarti.ColWidth(5) = 700
gridarti.ColWidth(6) = 700
gridarti.ColWidth(7) = 700
gridarti.ColWidth(8) = 700
gridarti.ColWidth(9) = 1000


gridarti.ColWidth(10) = 0
gridarti.ColWidth(11) = 0

gridarti.Rows = 2



xfil = 1
Do Until wllave_arti.EOF
  probar.Value = probar.Value + 1
  DoEvents
  xfil = xfil + 1
  gridarti.Rows = gridarti.Rows + 1
  gridarti.RowHeight(gridarti.Rows - 1) = 285
  lisarti.AddItem wllave_arti!ART_alterno & " " & wllave_arti!art_nombre & String(120, " ") & wllave_arti!art_key
  If LK_CODCIA = "01" Then
    WBOT(0) = "04"
  Else
   WBOT(0) = LK_CODCIA
  End If
  WBOT(1) = wllave_arti!art_key
  wllave_bot.Requery
  If wllave_bot.EOF Then
     MsgBox "Avisar el Codigo no Existe (SUBALMACEN): " & wllave_arti!art_key, 48, Pub_Titulo
     GoTo sale_otro
  End If
'  If Val(wllave_arti!PRE_COSTO_ANT) = 1610.7 Then Stop
  gridarti.TextMatrix(xfil, 0) = Format(xfil - 1, "00")
  gridarti.TextMatrix(xfil, 1) = Trim(wllave_arti!art_nombre)
  gridarti.TextMatrix(xfil, 2) = Trim(wllave_arti!PRE_UNIDAD)
  gridarti.TextMatrix(xfil, 3) = Format(wllave_arti!PRE_cosTO_ant, "0.00")
  gridarti.TextMatrix(xfil, 4) = Format(wllave_arti!ARM_COSPRO * wllave_arti!PRE_EQUIV, "0.00")
  gridarti.TextMatrix(xfil, 5) = Format((Val(wllave_arti!arm_stock) + Val(wllave_bot!arm_stock)) / Val(wllave_arti!PRE_EQUIV), "0.00")
  gridarti.TextMatrix(xfil, 6) = Format(wllave_arti!ART_STOCK_MAX / wllave_arti!PRE_EQUIV, "0.00")
  gridarti.TextMatrix(xfil, 7) = Format(Val(gridarti.TextMatrix(xfil, 5)) - Val(gridarti.TextMatrix(xfil, 6)), "0.00")
  If LK_CODCIA = "01" Then
    WBOT(0) = "03"
  Else
    WBOT(0) = LK_CODCIA
  End If
  WBOT(1) = wllave_arti!art_key
  wllave_bot.Requery
  If wllave_bot.EOF Then
     MsgBox "Avisar el Codigo no Existe (BOTICA): " & wllave_arti!art_key, 48, Pub_Titulo
     GoTo sale_otro
  End If
  gridarti.TextMatrix(xfil, 8) = Format((wllave_bot!arm_stock) / wllave_arti!PRE_EQUIV, "0.00")
  gridarti.TextMatrix(xfil, 9) = Format(wllave_arti!PRE_PRE6, "0.00")

  gridarti.TextMatrix(xfil, 10) = wllave_arti!ART_alterno
  gridarti.TextMatrix(xfil, 11) = wllave_arti!art_key
  
  If wllave_arti!art_plancha = 1 Then
    gridarti.Row = xfil
    For xpinta = 0 To gridarti.Cols - 1
       gridarti.COL = xpinta
       gridarti.CellForeColor = vbRed
    Next xpinta
  End If
sale_otro:
wllave_arti.MoveNext
Loop
gridarti.Visible = True
probar.Visible = False
DoEvents
End Sub

Private Sub tpromedio_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
'If KeyAscii = 13 Then
' Azul txtagencia, txtagencia
'End If
End Sub

Private Sub tserie_GotFocus()
fraListar.Visible = False
End Sub

Private Sub tserie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
txtdoc.SetFocus

End Sub

Private Sub txtagencia_GotFocus()
temporal = Trim(txtagencia.Text)
End Sub

Private Sub txtagencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtcontacto.SetFocus
End If
End Sub

Private Sub txtagencia_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 45 Then
PUB_TIPREG = -10
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "Mantenimiento de Transportistas"
FrmDatArti.Show 1
PRO_V.Requery
txtagencia.Clear
Do Until PRO_V.EOF
    txtagencia.AddItem Trim(PRO_V!TRN_NOMBRE) & String(80, " ") & PRO_V!TRN_KEY
    PRO_V.MoveNext
Loop
txtagencia.SetFocus
DoEvents

End If

End Sub

Private Sub txtagencia_LostFocus()
On Error GoTo sale
If Trim(temporal) = Trim(txtagencia.Text) Then Exit Sub
par_llave.Edit
par_llave!PAR_AGE_EMP = Trim(txtagencia.Text)
par_llave.Update
par_llave.Requery
Exit Sub
sale:
End Sub

Private Sub txtcontacto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   fpago.SetFocus
End If

End Sub

Private Sub txtdias_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  Azul tpromedio, tpromedio
End If
End Sub

Private Sub txtdoc_GotFocus()
fraListar.Visible = False
End Sub

Public Function PROCESA_CANTIDAD(WCODART As Currency, wdias As Integer) As Currency
Dim fx As Integer
Dim fecha1
Dim fecha2
Dim wdividir As Integer
Dim wflag  As String * 1
Dim WARTI As rdoQuery
Dim wllave_arti As rdoResultset
Dim wrango As Integer
Dim sum_cantidad As Currency
Dim WQ_cANTIDAD As Currency
PSLOC_WARTI(0) = ""
PSLOC_WARTI(1) = ""
PSLOC_WARTI(2) = ""

For fx = 0 To LISCIA.ListCount - 1
 LISCIA.ListIndex = fx
 If LISCIA.Selected(fx) Then
  If Val(PSLOC_WARTI(0)) = 0 Then
    PSLOC_WARTI(0) = Left(LISCIA.Text, 2)
    GoTo dale
  End If
  If Val(PSLOC_WARTI(1)) = 0 Then
    PSLOC_WARTI(1) = Left(LISCIA.Text, 2)
    GoTo dale
  End If
  If Val(PSLOC_WARTI(2)) = 0 Then
    PSLOC_WARTI(2) = Left(LISCIA.Text, 2)
    GoTo dale
  End If
 End If
dale:
Next fx
If Val(PSLOC_WARTI(0)) = 0 And Val(PSLOC_WARTI(1)) = 0 And Val(PSLOC_WARTI(2)) = 0 Then
  PSLOC_WARTI(0) = LK_CODCIA
End If

PSLOC_WARTI(3) = WCODART
wdividir = 0
wrango = Val(tpromedio.Text) * -1
PROCESA_CANTIDAD = 0
fecha1 = LOC_FECHA_INI ' DateAdd("d", wrango, LK_FECHA_DIA)
fecha2 = DateAdd("d", wdias, fecha1)
wflag = "S"
Do Until wflag = "N"
    If fecha1 > LOC_FECHA_FIN Then 'LK_FECHA_DIA Then
       Exit Do
    End If
    PSLOC_WARTI(4) = fecha1
    PSLOC_WARTI(5) = fecha2
    llave_sum_arti.Requery
    If Not llave_sum_arti.EOF Then
      If Nulo_Valor0(llave_sum_arti!cantidad) <> 0 Then
        wdividir = wdividir + 1
        sum_cantidad = sum_cantidad + Nulo_Valor0(llave_sum_arti!cantidad)
      End If
    End If
    
    fecha1 = DateAdd("d", 1, fecha2)
    fecha2 = DateAdd("d", wdias, fecha1)
Loop
If sum_cantidad <> 0 Then
WQ_cANTIDAD = Format((sum_cantidad / wdividir), "000000")
Else
  sum_cantidad = 0
End If
PROCESA_CANTIDAD = WQ_cANTIDAD

End Function

Private Sub txtdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
txtagencia.SetFocus
End Sub

Private Sub txtdoc_LostFocus()
Screen.MousePointer = 11
DoEvents
Dim indice As Integer
Barra.Visible = True
DoEvents
exito = True
fila = 1
SUM_D = 0
SUM_H = 0
CABE_MAN
'LIMPIA_DATOS
GoSub ACT1

fila = 0
Barra.Visible = False
grid_fac.SetFocus
GoTo fin

ACT1:
PSTEMP_LLAVE.rdoParameters(0) = LK_CODCIA
PSTEMP_LLAVE.rdoParameters(1) = Val(tserie.Text)
PSTEMP_LLAVE.rdoParameters(2) = Val(txtdoc.Text)
temp_llave.Requery
If temp_llave.EOF Then
 MsgBox "No Existe Docuemnto...", 48, Pub_Titulo
 GoTo borra
End If

Barra.Min = 0
Barra.max = temp_llave.RowCount
Barra.Value = 0

fila = 2
indice = 0
For fila = 0 To codpro.ListCount - 1
''Do Until indice = codpro.ListCount - 1
    codpro.ListIndex = fila
    If Val(Right(codpro.Text, 9)) = temp_llave!PED_CODCLIE Then Exit For
    ''indice = indice + 1
Next fila
''Loop

'indice = 0
'Do Until indice = txtagencia.ListCount - 1
'    txtagencia.ListIndex = indice
If cheprecios.Value = 0 Then
    txtagencia.Text = UCase(Nulo_Valors(temp_llave!PED_OFERTA))
Else
    txtagencia.Text = ""
End If
'    indice = indice + 1
'Loop
'txtagencia.Text = Nulo_Valors(temp_llave!PED_OFERTA)
txtcontacto.Text = Trim(Nulo_Valors(temp_llave!PED_CONTACTO))
Combo1.Text = Trim(Nulo_Valors(temp_llave!PED_NOMCLIE))
On Error GoTo PASES
For fila = 0 To fpago.ListCount - 1
 fpago.ListIndex = fila
 If Trim(UCase(Trim(Left(fpago.Text, 30)))) = Trim(UCase(Left(temp_llave!PED_FORMA, 30))) Then
   Exit For
 End If
Next fila
If Nulo_Valors(temp_llave!PED_MONEDA) = "S" Then
 moneda.ListIndex = 0
Else
 moneda.ListIndex = 1
End If
On Error GoTo 0
fecha.Caption = Format(temp_llave!PED_FECHA, "dd/mm/yyyy")
fila = 0
fila = 2
cheprecios.Value = Nulo_Valor0(temp_llave!PED_DIA_VISITA)

Do Until temp_llave.EOF
   grid_fac.Rows = fila + 2
   grid_fac.TextMatrix(fila, 2) = Format(temp_llave!PED_CANTIDAD, "0.00")
   grid_fac.TextMatrix(fila, 4) = Format(temp_llave!PED_PRECIO, "0.0000")
   grid_fac.TextMatrix(fila, 3) = temp_llave!PED_UNIDAD
   grid_fac.TextMatrix(fila, 9) = temp_llave!PED_codart
   grid_fac.TextMatrix(fila, 5) = temp_llave!PED_SUBTOTAL
   If cheprecios.Value = 1 Then
     grid_fac.TextMatrix(fila, 8) = temp_llave!PED_OFERTA
   End If
   grid_fac.TextMatrix(fila, 19) = 5
   SQ_OPER = 2
   PUB_CODART = temp_llave!PED_codart
   pu_codcia = LK_CODCIA
   LEER_PRE_LLAVE
   grid_fac.TextMatrix(fila, 18) = 0
   Do Until pre_mayor.EOF
    If Val(pre_mayor!PRE_EQUIV) = Val(temp_llave!PED_EQUIV) Then
        grid_fac.TextMatrix(fila, 18) = Val(pre_mayor!pre_PESO)
        Exit Do
    End If
   pre_mayor.MoveNext
   Loop
   
   SQ_OPER = 1
   PUB_KEY = temp_llave!PED_codart
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   grid_fac.TextMatrix(fila, 0) = "'" & art_LLAVE!art_nombre
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!ART_alterno

   If fila = 2 Then
   End If
   fila = fila + 1
   Barra.Value = fila - 2
   temp_llave.MoveNext
Loop
suma_grid
f1.Enabled = True
ESTADO.Enabled = True

Return

fin:
Screen.MousePointer = 0
Exit Sub
borra:
cancelar_Click
Screen.MousePointer = 0
Exit Sub
PASES:
Resume Next


End Sub

Private Sub UNIDAD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 45 Then Exit Sub


SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 9))
PUB_SECUEN = Val(Right(unidad.Text, 4))
LEER_PRE_LLAVE
grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!PRE_SECUENCIA

grid_fac.TextMatrix(grid_fac.Row, 3) = InputBox("Unidad: ")

'grid_fac.TextMatrix(grid_fac.Row, 4) = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 11)) / Val(grid_fac.TextMatrix(grid_fac.Row, 17)), "0.00")
'grid_fac.TextMatrix(grid_fac.Row, 17) = pre_llave!pre_equiv
'grid_fac.TextMatrix(grid_fac.Row, 6) = Format(grid_fac.TextMatrix(grid_fac.Row, 14) / pre_llave!pre_equiv, "#######")
'grid_fac.TextMatrix(grid_fac.Row, 7) = Format(grid_fac.TextMatrix(grid_fac.Row, 15) / pre_llave!pre_equiv, "#######")
'grid_fac.TextMatrix(grid_fac.Row, 8) = Format(grid_fac.TextMatrix(grid_fac.Row, 16) / pre_llave!pre_equiv, "#######")
    


unidad.Visible = False
'suma_grid
grid_fac.COL = 4
grid_fac.SetFocus


End Sub

Private Sub unidad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 9))
PUB_SECUEN = Val(Right(unidad.Text, 4))
LEER_PRE_LLAVE
grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!PRE_SECUENCIA
grid_fac.TextMatrix(grid_fac.Row, 3) = Trim(Left(unidad.Text, 12))
grid_fac.TextMatrix(grid_fac.Row, 4) = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 11)) / Val(grid_fac.TextMatrix(grid_fac.Row, 17)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 17) = pre_llave!PRE_EQUIV
grid_fac.TextMatrix(grid_fac.Row, 6) = Format(grid_fac.TextMatrix(grid_fac.Row, 14) / pre_llave!PRE_EQUIV, "#######")
grid_fac.TextMatrix(grid_fac.Row, 7) = Format(grid_fac.TextMatrix(grid_fac.Row, 15) / pre_llave!PRE_EQUIV, "#######")
grid_fac.TextMatrix(grid_fac.Row, 8) = Format(grid_fac.TextMatrix(grid_fac.Row, 16) / pre_llave!PRE_EQUIV, "#######")
    


unidad.Visible = False
suma_grid
grid_fac.COL = 4
grid_fac.SetFocus

End Sub

Public Sub veri_suma()
Dim I As Integer
Dim item_suma As Integer
item_suma = 0
For I = 1 To gridarti.Rows - 1
   If Val(gridarti.TextMatrix(I, 9)) <> 0 Then
     item_suma = item_suma + 1
   End If
Next I
 If (grid_fac.Rows + item_suma) >= Val(imax.Text) + 2 Then
     MsgBox "Llego al Maximo de Item.. Generar otra Orden de Compra. Maximo = " & imax.Text, 48, Pub_Titulo
 End If

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

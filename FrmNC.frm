VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmNC 
   Caption         =   "Establecer Descuentos por Nota de Creditos"
   ClientHeight    =   6735
   ClientLeft      =   2565
   ClientTop       =   1920
   ClientWidth     =   11400
   Icon            =   "FrmNC.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11400
   Begin VB.Frame frafechas 
      Height          =   1695
      Left            =   6720
      TabIndex        =   68
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
      Begin MSMask.MaskEdBox tf2 
         Height          =   285
         Left            =   120
         TabIndex        =   69
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tf1 
         Height          =   285
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lblnc 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ini "
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   72
         Top             =   120
         Width           =   705
      End
      Begin VB.Label lblnc 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin "
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   70
         Top             =   840
         Width           =   750
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   615
      Left            =   8640
      TabIndex        =   55
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
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
   Begin VB.Frame fraop 
      Caption         =   "Opciones de Calculo "
      Height          =   1935
      Left            =   0
      TabIndex        =   42
      Top             =   240
      Width           =   8415
      Begin VB.OptionButton opnc 
         Caption         =   "Prod. C/ Descto. C.Reparto"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   75
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Frame fradoc 
         Height          =   735
         Left            =   2280
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   4215
         Begin VB.ComboBox tfbg 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox tnumero 
            Height          =   285
            Left            =   2160
            TabIndex        =   54
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox tserie 
            Height          =   285
            Left            =   1440
            TabIndex        =   53
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblnc 
            AutoSize        =   -1  'True
            Caption         =   "Tipo :"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   405
         End
      End
      Begin VB.OptionButton opnc 
         Caption         =   "Productos de Selección"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   73
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Frame fralin 
         Height          =   1455
         Left            =   2520
         TabIndex        =   65
         Top             =   360
         Visible         =   0   'False
         Width           =   4215
         Begin VB.ListBox lineas 
            Height          =   960
            Left            =   720
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   67
            Top             =   160
            Width           =   3375
         End
         Begin VB.Label lblnc 
            AutoSize        =   -1  'True
            Caption         =   "Linea :"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.OptionButton opnc 
         Caption         =   "Desc. por Linea"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   64
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox tcodcli 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame frapro 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   2160
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   5775
         Begin VB.TextBox i_codart2 
            Height          =   285
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label LBLNOMARTI 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "#"
            Height          =   255
            Left            =   1320
            TabIndex        =   50
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label lblnc 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Prod."
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   49
            Top             =   0
            Width           =   1035
         End
      End
      Begin VB.OptionButton opnc 
         Caption         =   "Desc. por Documento"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton opnc 
         Caption         =   "Desc. por Producto."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Label lblnc 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblnomcli 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "#"
         Height          =   255
         Left            =   2520
         TabIndex        =   46
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label lblnc 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   45
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdmst 
      Caption         =   "Mostrar"
      Height          =   495
      Left            =   8640
      TabIndex        =   41
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdactpre 
      Caption         =   "&Aplicar N/C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   3840
      Picture         =   "FrmNC.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5640
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   345
      Left            =   1590
      TabIndex        =   39
      Top             =   4110
      Visible         =   0   'False
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame fralista 
      Caption         =   "Lista Detalle :"
      Height          =   255
      Left            =   8760
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   975
      Begin VB.CommandButton cmdesc 
         Caption         =   "Retornar"
         Height          =   555
         Left            =   4320
         TabIndex        =   24
         Top             =   2730
         Width           =   915
      End
      Begin VB.ListBox lstdoc 
         Height          =   645
         Left            =   270
         TabIndex        =   21
         Top             =   2715
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ListBox lstcosto 
         Height          =   2790
         Left            =   240
         TabIndex        =   20
         Top             =   450
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Facturas NO Confirmadas"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Producto c/ Cambio de Costo"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   10920
      Picture         =   "FrmNC.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   435
   End
   Begin VB.Frame fracambio 
      Caption         =   "Productos con mucha variación :"
      Height          =   3195
      Left            =   -30
      TabIndex        =   26
      Top             =   2250
      Visible         =   0   'False
      Width           =   11415
      Begin RichTextLib.RichTextBox TEXTOVARpre 
         Height          =   375
         Left            =   3600
         TabIndex        =   38
         Top             =   510
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"FrmNC.frx":09CE
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
      Begin VB.CommandButton Command4 
         Caption         =   "Actualizar Lista."
         Height          =   315
         Left            =   270
         TabIndex        =   29
         Top             =   6030
         Width           =   2265
      End
      Begin VB.CommandButton cmdRET 
         Caption         =   "Retornar"
         Height          =   315
         Left            =   5220
         TabIndex        =   28
         Top             =   6000
         Width           =   1785
      End
      Begin MSFlexGridLib.MSFlexGrid gridcabe 
         Height          =   2925
         Left            =   0
         TabIndex        =   27
         ToolTipText     =   "[Enter] = para Editar"
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5159
         _Version        =   393216
         Rows            =   1
         Cols            =   10
         FixedCols       =   5
         BackColorBkg    =   8421376
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Cambios:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8640
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   1425
      Begin VB.TextBox fechafinal 
         Height          =   285
         Left            =   3120
         TabIndex        =   14
         Top             =   270
         Visible         =   0   'False
         Width           =   1335
      End
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"FrmNC.frx":0A5C
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
      Begin MSFlexGridLib.MSFlexGrid gridigv 
         Height          =   2565
         Left            =   90
         TabIndex        =   1
         ToolTipText     =   "[Enter] = para Editar"
         Top             =   840
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   4524
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         BackColorBkg    =   8421376
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "hasta"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Digitar [enter] para modificar valores:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Mostar del :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   6450
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proceso de Costeo de Articulo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1425
      Begin VB.Frame Frame3 
         Caption         =   "Costeo por Compra"
         Height          =   825
         Left            =   210
         TabIndex        =   32
         Top             =   810
         Width           =   5025
         Begin VB.TextBox txtnumero 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3270
            TabIndex        =   35
            Text            =   "0"
            Top             =   390
            Width           =   945
         End
         Begin VB.TextBox txtserie 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2340
            TabIndex        =   34
            Text            =   "0"
            Top             =   390
            Width           =   765
         End
         Begin VB.CheckBox chefac 
            Caption         =   "Costear por Nro. Interno"
            Height          =   375
            Left            =   60
            TabIndex        =   33
            Top             =   240
            Width           =   2145
         End
         Begin VB.Label Label4 
            Caption         =   "Numero:"
            Height          =   195
            Index           =   1
            Left            =   3270
            TabIndex        =   37
            Top             =   180
            Width           =   645
         End
         Begin VB.Label Label4 
            Caption         =   "Serie :"
            Height          =   195
            Index           =   0
            Left            =   2310
            TabIndex        =   36
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.CheckBox cheactpre 
         Caption         =   "Pre Visualizar Precios."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   1560
         TabIndex        =   30
         Top             =   3270
         Value           =   1  'Checked
         Width           =   2505
      End
      Begin VB.CheckBox chelives 
         Caption         =   "Actualizar a Botica Lives"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   300
         TabIndex        =   25
         Top             =   1980
         Width           =   1455
      End
      Begin VB.TextBox txtava 
         Height          =   285
         Left            =   4560
         TabIndex        =   18
         Text            =   "0"
         Top             =   2220
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TLINEA 
         Height          =   285
         Left            =   4620
         TabIndex        =   16
         Top             =   330
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TART 
         Height          =   288
         Left            =   3180
         TabIndex        =   12
         Top             =   300
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.CheckBox CHEART 
         Caption         =   "Solo un Articulo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   210
         TabIndex        =   11
         Top             =   300
         Width           =   1695
      End
      Begin MSComctlLib.ProgressBar ProgBar 
         Height          =   195
         Left            =   420
         TabIndex        =   5
         Top             =   1710
         Visible         =   0   'False
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Costear Articulos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   2040
         Picture         =   "FrmNC.frx":0AEA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Opciones para Cambio de Lista de Precios:"
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
         Height          =   195
         Left            =   270
         TabIndex        =   31
         Top             =   3060
         Width           =   3795
      End
      Begin VB.Label LART 
         Caption         =   "Codigo Alterno:"
         Height          =   255
         Left            =   1980
         TabIndex        =   13
         Top             =   330
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label POR 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Procesando.."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   1950
         Visible         =   0   'False
         Width           =   4815
      End
   End
   Begin VB.Label lbltotdes 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Height          =   255
      Left            =   7560
      TabIndex        =   60
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lbltotdoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Height          =   255
      Left            =   5880
      TabIndex        =   62
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Total Documento(s)"
      Height          =   255
      Left            =   5880
      TabIndex        =   63
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Total Ajustado"
      Height          =   255
      Left            =   7560
      TabIndex        =   61
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lbltotnc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Height          =   255
      Left            =   9360
      TabIndex        =   59
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Total N/Credito"
      Height          =   255
      Left            =   9360
      TabIndex        =   58
      Top             =   5640
      Width           =   1335
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
      Left            =   0
      TabIndex        =   17
      Top             =   6480
      Width           =   11415
   End
End
Attribute VB_Name = "FrmNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WCOD_ORIGINAL As Currency
Dim loc_key As Integer
Dim actllave_precio As rdoResultset
Dim PSACTLLAVE As rdoQuery

Dim new_precio As rdoResultset
Dim PSNEWLLAVE As rdoQuery

Dim temporal
Dim temfecha

Private Sub chefac_Click()
If chefac.Value = 1 Then
txtserie.Enabled = True
txtnumero.Enabled = True
Azul txtserie, txtserie
Else
txtserie.Enabled = False
txtnumero.Enabled = False
End If
End Sub


Private Sub cmdactpre_Click()
Dim xcI As Integer
    For xcI = 1 To FrmNC.gridcabe.Rows - 1
      If Val(FrmNC.gridcabe.TextMatrix(xcI, 12)) < 0 Then
         MsgBox "Verificar un Valor de Descuento no puede estar em Negativo.", 48, Pub_Titulo
         Exit Sub
      End If
    Next xcI

 FrmNC.Hide
End Sub

Private Sub cmdesc_Click()
fralista.Visible = False
End Sub

Private Sub cmdmst_Click()
Screen.MousePointer = 11
Dim WS_TP As String
Dim WLINEAS As String
Dim FILTRO_ADD As String
Dim I As Integer
Dim VALOR As Currency
Dim rslistcambio As rdoResultset
Dim PSlistcambio As rdoQuery

Dim pre_caja As rdoResultset
Dim PSPRE_CAJA As rdoQuery

gridcabe.Cols = 19
gridcabe.Rows = 1
gridcabe.TextMatrix(0, 0) = "Ord"
gridcabe.TextMatrix(0, 1) = "Fecha"
gridcabe.TextMatrix(0, 2) = "Vend."
gridcabe.TextMatrix(0, 3) = "Documento"
gridcabe.TextMatrix(0, 4) = "Sec."
gridcabe.TextMatrix(0, 5) = "Producto"
gridcabe.TextMatrix(0, 6) = "Cantidad"
gridcabe.TextMatrix(0, 7) = "Unidad"
gridcabe.TextMatrix(0, 8) = "Precio"
gridcabe.TextMatrix(0, 9) = "Total"
gridcabe.TextMatrix(0, 10) = "P.x NC"
gridcabe.TextMatrix(0, 11) = "T.x NC"
gridcabe.TextMatrix(0, 12) = "DescEfc."

gridcabe.ColWidth(0) = 400
gridcabe.ColWidth(1) = 800
gridcabe.ColWidth(2) = 400
gridcabe.ColWidth(3) = 1400
gridcabe.ColWidth(4) = 500
gridcabe.ColWidth(5) = 2000
gridcabe.ColWidth(6) = 800
gridcabe.ColWidth(7) = 700
gridcabe.ColWidth(8) = 800
gridcabe.ColWidth(9) = 800
gridcabe.ColWidth(10) = 800
gridcabe.ColWidth(11) = 800
gridcabe.ColWidth(12) = 800
gridcabe.ColWidth(13) = 0
gridcabe.ColWidth(14) = 0
gridcabe.ColWidth(15) = 0
gridcabe.ColWidth(16) = 0
gridcabe.ColWidth(17) = 0
gridcabe.ColWidth(18) = 0
'MsgBox PUB_TIPDOC

FILTRO_ADD = " " ' AND FAR_SIGNO_CAR = " & pub_signo_car * -1

If opnc(0).Value = True Then ' POR PRODUCTO
    If Val(i_codart2.Text) = 0 Then
      MsgBox "Ingrese Codigo del Producto", 48, Pub_Titulo
      Azul i_codart2, i_codart2
      Exit Sub
    End If
    pub_cadena = "SELECT FAR_FBG, FAR_NUMSER, FAR_NUMFAC , FAR_NUMSEC, FAR_CODCIA, FAR_DESCRI , FAR_FECHA, FAR_CODART, FAR_DESCRI, FAR_EQUIV , FAR_PRECIO , FAR_CANTIDAD , FAR_CODVEN , FAR_SUBTOTAL , FAR_PORDESCTO1 , FAR_TIPDOC , FAR_SIGNO_CAR, ART_NOMBRE , FAR_ESTADO2 FROM FACART, ARTI " & _
    " WHERE (FAR_CODCIA = ART_CODCIA) AND (FAR_CODART = ART_KEY) AND FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND " & _
    " FAR_CODART = ?  AND FAR_CODCLIE = ? AND FAR_CP = 'C' AND FAR_ESTADO <> 'E' " & FILTRO_ADD & " ORDER BY FAR_FECHA, FAR_FBG, FAR_NUMSER, FAR_NUMFAC "
    Set PSlistcambio = CN.CreateQuery("", pub_cadena)
    PSlistcambio(0) = 10
    PSlistcambio(1) = LK_CODCIA
    PSlistcambio(2) = tf1.Text
    PSlistcambio(3) = tf2.Text
    PSlistcambio(4) = Val(i_codart2.Text)
    PSlistcambio(5) = Val(tcodcli.Text)
    
ElseIf opnc(1).Value = True Then ' POR PRODUCTO ' POR DOCUMENTO
    pub_cadena = "SELECT FAR_FBG, FAR_NUMSER, FAR_NUMFAC , FAR_NUMSEC , FAR_DESCRI , FAR_CODCIA, FAR_FECHA, FAR_CODART, FAR_DESCRI, FAR_EQUIV , FAR_PRECIO , FAR_CANTIDAD , FAR_CODVEN , FAR_SUBTOTAL ,FAR_PORDESCTO1 , FAR_TIPDOC , FAR_SIGNO_CAR, ART_NOMBRE , FAR_ESTADO2 FROM FACART, ARTI " & _
    "WHERE (FAR_CODCIA = ART_CODCIA) AND (FAR_CODART = ART_KEY) AND FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? AND " & _
    " FAR_NUMFAC = ?  AND FAR_CODCLIE = ?  AND FAR_CP = 'C' AND FAR_ESTADO <> 'E'  " & FILTRO_ADD & " ORDER BY FAR_NUMSEC"
    Set PSlistcambio = CN.CreateQuery("", pub_cadena)
    PSlistcambio(0) = 10
    PSlistcambio(1) = LK_CODCIA
    PSlistcambio(2) = Trim(tfbg.Text)
    PSlistcambio(3) = Val(tserie.Text)
    PSlistcambio(4) = Val(tnumero.Text)
    PSlistcambio(5) = Val(tcodcli.Text)
ElseIf opnc(2).Value = True Then ' POR LINEAS
    WLINEAS = ""
    For I = 0 To lineas.ListCount - 1
      If lineas.Selected(I) Then
        WLINEAS = WLINEAS + Trim(Right(lineas.List(I), 8)) & ","
      End If
    Next I
    If WLINEAS <> "" Then
      WLINEAS = " AND ART_FAMILIA IN (" & Mid(WLINEAS, 1, Len(WLINEAS) - 1) & ")"
    Else
       WLINEAS = ""
    End If
    pub_cadena = "SELECT FAR_FBG, FAR_NUMSER, FAR_NUMFAC , FAR_NUMSEC, FAR_CODCIA, FAR_DESCRI , FAR_FECHA, FAR_CODART, FAR_DESCRI, FAR_EQUIV , FAR_PRECIO , FAR_CANTIDAD , FAR_CODVEN , FAR_SUBTOTAL ,  FAR_TIPDOC , FAR_SIGNO_CAR, FAR_PORDESCTO1 , ART_NOMBRE , FAR_ESTADO2 FROM FACART, ARTI " & _
    " WHERE (FAR_CODCIA = ART_CODCIA) AND (FAR_CODART = ART_KEY) AND FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ?  " & _
    " " & WLINEAS & " AND FAR_CODCLIE = ? AND FAR_CP = 'C' AND FAR_ESTADO <> 'E'  " & FILTRO_ADD & " ORDER BY FAR_FECHA, FAR_FBG, FAR_NUMSER, FAR_NUMFAC "
    Set PSlistcambio = CN.CreateQuery("", pub_cadena)
    PSlistcambio(0) = 10
    PSlistcambio(1) = LK_CODCIA
    PSlistcambio(2) = tf1.Text
    PSlistcambio(3) = tf2.Text
    PSlistcambio(4) = Val(tcodcli.Text)
ElseIf opnc(3).Value = True Then ' POR producto de Seleccion
    pub_cadena = "SELECT FAR_FBG, FAR_NUMSER, FAR_NUMFAC , FAR_NUMSEC, FAR_CODCIA, FAR_DESCRI , FAR_FECHA, FAR_CODART, FAR_DESCRI, FAR_EQUIV , FAR_PRECIO , FAR_CANTIDAD , FAR_CODVEN , FAR_SUBTOTAL ,FAR_PORDESCTO1 , FAR_TIPDOC , FAR_SIGNO_CAR, ART_NOMBRE , FAR_ESTADO2 FROM FACART, ARTI " & _
    " WHERE (FAR_CODCIA = ART_CODCIA) AND (FAR_CODART = ART_KEY) AND FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND " & _
    " ART_LINEA = 1 AND FAR_CP = 'C' AND FAR_ESTADO <> 'E'  " & FILTRO_ADD & " ORDER BY FAR_CODART "
    Set PSlistcambio = CN.CreateQuery("", pub_cadena)
    PSlistcambio(0) = 10
    PSlistcambio(1) = LK_CODCIA
    PSlistcambio(2) = tf1.Text
    PSlistcambio(3) = tf2.Text
ElseIf opnc(4).Value = True Then ' POR producto de Seleccion C/ Contado Reparto.
    pub_cadena = "SELECT FAR_FBG, FAR_NUMSER, FAR_NUMFAC , FAR_NUMSEC, FAR_CODCIA, FAR_DESCRI , FAR_FECHA, FAR_CODART, FAR_DESCRI, FAR_EQUIV , FAR_PRECIO , FAR_CANTIDAD , FAR_CODVEN , FAR_SUBTOTAL ,FAR_PORDESCTO1 , FAR_TIPDOC , FAR_SIGNO_CAR, ART_NOMBRE , FAR_ESTADO2 FROM FACART, ARTI " & _
    " WHERE (FAR_CODCIA = ART_CODCIA) AND (FAR_CODART = ART_KEY) AND FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND " & _
    " ART_LINEA = 1  AND FAR_CP = 'C' AND FAR_ESTADO <> 'E'  " & FILTRO_ADD & " ORDER BY FAR_CODART "
    Set PSlistcambio = CN.CreateQuery("", pub_cadena)
    PSlistcambio(0) = 10
    PSlistcambio(1) = LK_CODCIA
    PSlistcambio(2) = tf1.Text
    PSlistcambio(3) = tf2.Text
End If
fracambio.Visible = False
DoEvents
Set rslistcambio = PSlistcambio.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
If rslistcambio.EOF Then
  Screen.MousePointer = 0
  MsgBox "Nos Existe Datos "
Else
 pb.Visible = True
 pbMin = 0
 pbMax = 1 'rslistcambio.RowCount
 pbValue = 0
End If

Do Until rslistcambio.EOF
' pbValue = pbValue + 1
 If rslistcambio!FAR_ESTADO2 = "L" Then GoTo dale_arti
 DoEvents
 gridcabe.Rows = gridcabe.Rows + 1
 gridcabe.TextMatrix(gridcabe.Rows - 1, 0) = rslistcambio.AbsolutePosition
 gridcabe.TextMatrix(gridcabe.Rows - 1, 1) = Format(rslistcambio!far_fecha, "dd/mm/yy")
 gridcabe.TextMatrix(gridcabe.Rows - 1, 2) = Format(rslistcambio!FAR_CODVEN, "00")
 WS_TP = ""
 If rslistcambio!FAR_TIPDOC = "FA" And rslistcambio!far_signo_car = 0 Then
   WS_TP = "EFE"
 ElseIf rslistcambio!FAR_TIPDOC = "FA" And rslistcambio!far_signo_car <> 0 Then
   WS_TP = "CRE"
 ElseIf rslistcambio!FAR_TIPDOC = "CC" And rslistcambio!far_signo_car <> 0 Then
   WS_TP = "CC"
 End If
 gridcabe.TextMatrix(gridcabe.Rows - 1, 3) = rslistcambio!far_fbg & "/" & rslistcambio!far_numser & " - " & rslistcambio!far_numfac & " " & WS_TP
 gridcabe.TextMatrix(gridcabe.Rows - 1, 4) = rslistcambio!far_numsec
 gridcabe.TextMatrix(gridcabe.Rows - 1, 5) = rslistcambio!art_nombre
 gridcabe.TextMatrix(gridcabe.Rows - 1, 6) = Format(rslistcambio!FAR_cantidad / rslistcambio!FAR_equiv, "0.00")
 gridcabe.TextMatrix(gridcabe.Rows - 1, 7) = rslistcambio!far_descri
 gridcabe.TextMatrix(gridcabe.Rows - 1, 8) = Format(rslistcambio!FAR_PRECIO, "0.0000")
 gridcabe.TextMatrix(gridcabe.Rows - 1, 9) = Format(rslistcambio!far_subtotal, "0.00")
 If opnc(3).Value Or opnc(4).Value Then
    gridcabe.TextMatrix(gridcabe.Rows - 1, 10) = Format(Val(rslistcambio!FAR_PRECIO) - Val(rslistcambio!FAR_pordescto1), "0.0000")
    VALOR = Val(gridcabe.TextMatrix(gridcabe.Rows - 1, 10)) * Val(gridcabe.TextMatrix(gridcabe.Rows - 1, 6))
    gridcabe.TextMatrix(gridcabe.Rows - 1, 11) = Format(VALOR, "0.00")
    VALOR = Val(gridcabe.TextMatrix(gridcabe.Rows - 1, 9)) - Val(gridcabe.TextMatrix(gridcabe.Rows - 1, 11))
    gridcabe.TextMatrix(gridcabe.Rows - 1, 12) = Format(VALOR, "0.00")
    If VALOR = 0 Then
      gridcabe.TextMatrix(gridcabe.Rows - 1, 12) = ""
    End If
 Else
    gridcabe.TextMatrix(gridcabe.Rows - 1, 10) = Format(rslistcambio!FAR_PRECIO, "0.0000")
    gridcabe.TextMatrix(gridcabe.Rows - 1, 11) = Format(rslistcambio!far_subtotal, "0.00")
    gridcabe.TextMatrix(gridcabe.Rows - 1, 12) = ""
 End If
 ' acceso al registro
 
 gridcabe.TextMatrix(gridcabe.Rows - 1, 13) = rslistcambio!far_fecha
 gridcabe.TextMatrix(gridcabe.Rows - 1, 14) = rslistcambio!far_fbg
 gridcabe.TextMatrix(gridcabe.Rows - 1, 15) = rslistcambio!far_numser
 gridcabe.TextMatrix(gridcabe.Rows - 1, 16) = rslistcambio!far_numfac
 gridcabe.TextMatrix(gridcabe.Rows - 1, 17) = rslistcambio!far_numsec
 gridcabe.TextMatrix(gridcabe.Rows - 1, 18) = ""
dale_arti:
rslistcambio.MoveNext
Loop
Calcula_NC
Screen.MousePointer = 0
pb.Visible = False
fracambio.Visible = True
DoEvents

End Sub

Private Sub cmdRET_Click()
fracambio.Visible = False
End Sub

Private Sub Command1_Click()
Dim ACTLIVES As String
Dim WLINEA As Integer
Dim WCOSPRO_SUP As Currency
Dim WCODART As Currency
Dim FLAG_93 As String * 1
Dim wfec As String
Dim WS_SALDO As Currency
Dim WTEXTO As String
Dim CONTADOR As Integer
Dim wk_equiv As Currency
Dim indice As Integer
Dim WS_TIPO_CAMBIO As Double

Dim farartcosto  As rdoResultset
Dim PSFARCOSTO As rdoQuery

Dim far_cambiox  As rdoResultset
Dim PSFAR_CAMBIO As rdoQuery
Dim PSFAR_COSTO As rdoQuery
Dim Far_Cost As rdoResultset
Dim ultimo_costo, ultimo_cambio As Currency
Dim PSFAR_COSTO2 As rdoQuery
Dim Far_Cost2 As rdoResultset
Dim Far_envio As rdoResultset
Dim PSFAR_ENVIO As rdoQuery
Dim PSFAR_COMPR As rdoQuery
Dim Far_Compr As rdoResultset
Dim PSFAR_COMPR2 As rdoQuery
Dim Far_Compr2 As rdoResultset
Dim WPASA_PRECIO As Currency
Dim WPASA_COSPRO_ANT As Currency
Dim ws_cospro As Currency
Dim ws_valor_ant As Currency
Dim ws_compra As Currency
Dim ws_ajuste, ws_perdida, WS_AJUSTE2 As Currency
Dim WS_NOMART As String * 22
Dim WR_IMPORTE1 As String * 12
Dim WR_IMPORTE2 As String * 12
Dim WR_IMPORTE3 As String * 12
Dim WR_IMPORTE4 As String * 12
Dim WR_IMPORTE5 As String * 12
Dim WR_IMPORTE6 As String * 12
Dim WR_IMPORTE7 As String * 12
Dim WR_IMPORTE8 As String * 12
Dim WS_VALOR1, ws_valor2, ws_Valor3, ws_valor4, ws_valor5, ws_valor6, ws_valor7 As Currency
Dim wCODARTI As Currency
Dim flagcito As String * 1
Dim WMONEDA As String

Dim PSFAR_OTRO As rdoQuery
Dim Far_otro As rdoResultset
If fechafinal.Visible Then
  If Not IsDate(fechafinal.Text) Then
    MsgBox "Fecha no procede.", 48, Pub_Titulo
    Exit Sub
  End If
  If CDate(txtfecha.Text) > CDate(fechafinal.Text) Then
    MsgBox "No procede Fecha Final no puede ser mayor que la Inicial.", 48, Pub_Titulo
    Exit Sub
  End If
End If
WCODART = 0
If CHEART.Value = 1 Then
    SQ_OPER = 3
    pu_alterno = Trim(TART.Text)
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    art_llave_alt.Requery
    If art_llave_alt.EOF Then
       MsgBox "Codigo no Existe / No Procede ", 48, Pub_Titulo
       Exit Sub
    End If
    WCODART = art_llave_alt!art_key
End If

If Not IsDate(txtfecha.Text) Then
  MsgBox "Fecha no Prcede. ", 48, Pub_Titulo
  Exit Sub
End If

If gridigv.Visible = False Then
  MsgBox "Determine fecha inicial y final . Ademas pulse <Enter> en la fecha final", 48, Pub_Titulo
  Exit Sub
End If
If Not CONSIS_PERIODOS(CDate(txtfecha.Text)) Then
  Exit Sub
End If


pub_mensaje = "Proceso de Costeo de Articulos de la Fecha : " & Chr(13) & " " & Format(txtfecha.Text, "dd/mm/yyyy") & " al " & Format(LK_FECHA_DIA, "dd/mm/yyyy") & Chr(13) & "¿ Continuar? ..."
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
If Not CONSIS_TC(CDate(txtfecha.Text)) Then
  Exit Sub
End If
 
lstcosto.Clear
lstcosto.AddItem "Productos Costeados :"
lstdoc.Clear
lstdoc.AddItem "Nro Interno de Documentos No Confirmados:"
ProgBar.Visible = True
DoEvents
'POR(0).Visible = True
DoEvents
'POR(1).Visible = True
DoEvents
POR(2).Visible = True
DoEvents
ProgBar.Value = 0
ProgBar.Min = 0
POR(2).Caption = "Iniciando Proceso. . ."
ultimo_costo = 0
ultimo_cambio = 0
' proceso de consistencia de fechas con t.c.
FLAG_93 = ""
DoEvents
Command1.Enabled = False

pub_cadena = "DELETE FROM NEWPRECIOS WHERE PRE_CODCIA = '" & LK_CODCIA & "'"
CN.Execute pub_cadena, rdExecDirect

pub_cadena = "SELECT FAR_COSTEO, FAR_COSTEO_REAL, FAR_PRECIO FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_NUMGUIA = ? AND FAR_TIPMOV = 101  AND FAR_ESTADO <> 'E' and far_transito <> 'P' "
Set PSFAR_ENVIO = CN.CreateQuery("", pub_cadena)
PSFAR_ENVIO(0) = 0
PSFAR_ENVIO(1) = 0
PSFAR_ENVIO(2) = 0
Set Far_envio = PSFAR_ENVIO.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT FAR_CANTIDAD,FAR_COSTEO, FAR_COSTEO_REAL, FAR_PRECIO FROM FACART WHERE FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? AND FAR_TIPMOV = 93 AND FAR_SIGNO_ARM = 1 AND FAR_ESTADO <> 'E' and far_transito <> 'P' "
Set PSFAR_CAMBIO = CN.CreateQuery("", pub_cadena)
PSFAR_CAMBIO(0) = 0
PSFAR_CAMBIO(1) = 0
PSFAR_CAMBIO(2) = 0
Set far_cambiox = PSFAR_CAMBIO.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT FAR_CANTIDAD,FAR_COSTEO, FAR_COSTEO_REAL, FAR_PRECIO FROM FACART WHERE FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? AND FAR_TIPMOV = 93 AND FAR_SIGNO_ARM = 1 and far_transito <> 'P' "
Set PSFAR_OTRO = CN.CreateQuery("", pub_cadena)
PSFAR_OTRO(0) = 0
PSFAR_OTRO(1) = 0
PSFAR_OTRO(2) = 0
Set Far_otro = PSFAR_OTRO.OpenResultset(rdOpenKeyset, rdConcurValues)




'pub_cadena = "SELECT FAR_FECHA_COMPRA, FAR_COSPRO_SUP, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_STOCK, FAR_COSPRO FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA < ? AND FAR_CODART = ? and far_estado <>'E' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_NUMOPER2"
'Set PS_REP03 = CN.CreateQuery("", pub_cadena)
'PS_REP03(0) = 0
'PS_REP03(1) = LK_FECHA_DIA
'PS_REP03(2) = 0
'Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

' pub_cadena = "SELECT FAR_FECHA_COMPRA, FAR_COSPRO_SUP, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_STOCK, FAR_COSPRO                                                                                                                                                                                                                                                                                    FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA < ? AND FAR_CODART = ? and far_estado <>'E' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_NUMOPER2"
pub_cadena = "SELECT FAR_FECHA_COMPRA, FAR_FECHA_CAN,FAR_COSPRO_SUP, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_STOCK, FAR_COSPRO FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA < ? AND FAR_CODART = ? and far_estado <>'E' and far_estado2 <>'L' and far_transito <> 'P' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_TIPMOV, FAR_NUMOPER2"

                                                                                                                                                                                                                                                        
Set PSFAR_COSTO2 = CN.CreateQuery("", pub_cadena)
PSFAR_COSTO2(0) = 0
PSFAR_COSTO2(1) = LK_FECHA_DIA
PSFAR_COSTO2(2) = 0
Set Far_Cost2 = PSFAR_COSTO2.OpenResultset(rdOpenKeyset, rdConcurReadOnly)



'pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_FECHA_COMPRA >= ? AND  FAR_FECHA_COMPRA <= ?   and far_estado<>'E' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_NUMOPER2"


pub_cadena = "SELECT FAR_NUM_LOTE,FAR_PRECIO_NETO,FAR_COSPRO_ANT,FAR_COSPRO_SUP, FAR_STOCK , FAR_TIPO_CAMBIO, FAR_COSTEO_REAL, FAR_NUMGUIA, FAR_NUMSER, FAR_NUMFAC, FAR_OTRA_CIA, FAR_CODART, FAR_COSPRO, FAR_FLETE, FAR_BRUTO, FAR_DESCTO, FAR_GASTOS, FAR_TOT_DESCTO, FAR_SIGNO_ARM, FAR_CANTIDAD, FAR_TIPMOV, FAR_MONEDA, FAR_FECHA_COMPRA, FAR_FECHA_CAN, FAR_COSTEO, FAR_ESTADO, FAR_PRECIO, FAR_EQUIV  FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_FECHA_COMPRA >= ? AND  FAR_FECHA_COMPRA <= ?   and far_estado <>'E' and far_estado2 <>'L' and far_transito <> 'P' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC ,FAR_TIPMOV, FAR_NUMOPER2"
Set PSFAR_COSTO = CN.CreateQuery("", pub_cadena)
PSFAR_COSTO(0) = LK_CODCIA
PSFAR_COSTO(1) = 0
PSFAR_COSTO(2) = LK_FECHA_DIA
PSFAR_COSTO(3) = LK_FECHA_DIA
Set Far_Cost = PSFAR_COSTO.OpenResultset(rdOpenKeyset, rdConcurValues)

'¡DistinctRow
'DISTINCTROW
If chefac.Value = 1 Then
 pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ?  AND  FAR_ESTADO <> 'E' AND  FAR_ESTADO <> 'L' AND ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' )  and far_tipmov = 20 and far_numser = " & Val(txtserie.Text) & " and far_numfac = " & txtnumero.Text & " ORDER BY FAR_CODART"
Else
pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ?  AND  FAR_ESTADO <> 'E' AND  FAR_ESTADO <> 'L' AND ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' )  ORDER BY FAR_CODART"
End If
If CHEART.Value = 1 Then
  pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND FAR_CODART= " & WCODART & "  AND  FAR_ESTADO <> 'E' AND FAR_ESTADO2 <> 'L' AND ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' ) ORDER BY FAR_CODART"
End If
If Val(TLINEA.Text) <> 0 Then
  WLINEA = Val(TLINEA.Text)
  ' PRUEBA
  pub_cadena = "SELECT Distinct FAR_CODART FROM FACART, ARTI WHERE (ART_CODCIA = FAR_CODCIA) AND (FAR_CODART = ART_KEY) AND ART_LINEA = " & WLINEA & " AND FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND FAR_ESTADO <> 'E' AND FAR_ESTADO2 <> 'L' AND ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' ) ORDER BY FAR_CODART"
End If
' TEMPORAL
''''pub_cadena = "SELECT Distinct FAR_CODART FROM FACART, ARTI WHERE (ART_CODCIA = FAR_CODCIA) AND (FAR_CODART = ART_KEY) AND  FAR_TIPMOV = 6 AND FAR_NUM_LOTE <> 10 AND FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND ( FAR_COSTEO_REAL = 'A'  OR FAR_FECHA_COMPRA <> FAR_FECHA  OR FAR_MONEDA= 'D' )  AND FAR_ESTADO<>'E'  ORDER BY FAR_CODART"
Set PSFAR_COMPR = CN.CreateQuery("", pub_cadena)
PSFAR_COMPR(0) = LK_CODCIA
PSFAR_COMPR(1) = LK_FECHA_DIA
PSFAR_COMPR(2) = LK_FECHA_DIA
Set Far_Compr = PSFAR_COMPR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PSFAR_COMPR.rdoParameters(0) = LK_CODCIA
PSFAR_COMPR.rdoParameters(1) = CDate(txtfecha.Text)
If fechafinal.Visible Then
 PSFAR_COMPR.rdoParameters(2) = fechafinal.Text
Else
 PSFAR_COMPR.rdoParameters(2) = LK_FECHA_DIA
End If


WTEXTO = ""
CALCULA_93:

Far_Compr.Requery
ProgBar.Min = 0
ProgBar.Value = 0
If Not Far_Compr.EOF Then ProgBar.max = Far_Compr.RowCount ''+ 2

'Do Until Far_Compr.EOF
'F1 = F1 + 1
'xl.Cells(F1, 1) = Far_Compr!far_codart

'Loop
'
If Far_Compr.EOF Then
  GoTo pasa
   MsgBox "No Movimientos para Costear....", 48, Pub_Titulo
    GoTo flag_avanza
   Unload FrmNC
   Unload FORMGEN
   Exit Sub
End If
'OPEN_LOG "Proceso de Costeo"

POR(2).Caption = "Actualizando Costos. . ."
DoEvents
wCODARTI = Far_Compr!far_codart
flagcito = "A"
WPASA_COSPRO_ANT = 0
WCOSPRO_SUP = 0
Do Until Far_Compr.EOF
   If Val(txtava.Text) <> 0 Then
      If Far_Compr.AbsolutePosition < Val(txtava.Text) Then GoTo OTRO_ART
   End If
   ProgBar.Value = Far_Compr.AbsolutePosition
   POR(2).Caption = "Costeando... " & ProgBar.Value & " de " & ProgBar.max
   DoEvents
   'PSFAR_COSTO2.rdoParameters(0) = LK_CODCIA
   'PSFAR_COSTO2.rdoParameters(1) = Far_Compr!far_codart
   'PSFAR_COSTO2.rdoParameters(2) = CDate(txtfecha.Text)
   'Far_Cost2.Requery
   PSFAR_COSTO.rdoParameters(0) = LK_CODCIA
   PSFAR_COSTO.rdoParameters(1) = Far_Compr!far_codart
   PSFAR_COSTO.rdoParameters(2) = CDate(txtfecha.Text)
   If fechafinal.Visible Then
     PSFAR_COSTO.rdoParameters(3) = fechafinal.Text
   Else
     PSFAR_COSTO.rdoParameters(3) = LK_FECHA_DIA
   End If

   Far_Cost.Requery
   If Far_Cost.EOF Then
      MsgBox "Ver Mensaje : entre tablas facart : " & Far_Compr!far_codart
      End
   End If
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_KEY = Far_Cost!far_codart
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
      MsgBox "ERROR GRAVE ARTI... " & Chr(13) & "Codcia :" & pu_codcia & Chr(13) & "Codigo de interno de Arti: " & PUB_KEY
      End
   End If
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_CODART = Far_Cost!far_codart
  ' If PUB_CODART = 642 Then Stop
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
      MsgBox "ERROR GRAVE en ARTICULO... " & Chr(13) & "Codcia :" & pu_codcia & Chr(13) & "Codigo de interno de Arti: " & PUB_KEY
      End
   End If

   WS_NOMART = art_LLAVE!art_nombre
   ultimo_costo = arm_llave!ARM_COSTO_ULT
   WPASA_COSPRO_ANT = Nulo_Valor0(Far_Cost!FAR_COSPRO_ANT)
   
   If LK_EMP = "HER" Or LK_EMP = "PIU" Then
       ws_cospro = Far_Cost!FAR_COSPRO
   End If
   
   If CDate(txtfecha.Text) = GEN!GEN_FECHA_REF Then
      If Far_Cost!far_signo_arm = -1 Then
        WS_SALDO = (0 + (Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad) * -1)
      Else
        WS_SALDO = (0 + ((Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad) * -1)) '' + arm_llave!ARM_STOCK - Far_Cost2(1)
      End If
      If Far_Cost!FAR_COSTEO = "A" And Far_Cost!far_estado <> "E" And Far_Cost!far_signo_arm = 1 Then
      Else
      WS_SALDO = 0
      End If
   Else
      If Far_Cost!far_signo_arm = -1 Then
        WS_SALDO = (Nulo_Valor0(Far_Cost!FAR_STOCK) + ((Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad) * -1))
      Else
        WS_SALDO = (Nulo_Valor0(Far_Cost!FAR_STOCK) + ((Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad) * -1))
      End If
   End If
   If WS_SALDO <= 0 Then WS_SALDO = 0
 '  Far_Cost.Requery
   ProgBar.Value = 0
   ProgBar.max = Far_Cost.RowCount
   If Far_Cost!FAR_COSTEO = "A" And Far_Cost!far_estado <> "E" And Far_Cost!far_signo_arm = 1 And LK_EMP <> "PIU" Then
      WCOSPRO_SUP = Nulo_Valor0(Far_Cost!far_cospro_sup)
      ws_cospro = Nulo_Valor0(Far_Cost!far_cospro_sup)
   Else
      WCOSPRO_SUP = Nulo_Valor0(Far_Cost!FAR_COSPRO)
      ws_cospro = Nulo_Valor0(Far_Cost!FAR_COSPRO)
   End If
   PSFAR_COSTO2(0) = LK_CODCIA
   PSFAR_COSTO2(1) = Format(txtfecha.Text, "dd/mm/yyyy")
   PSFAR_COSTO2(2) = Far_Cost!far_codart
   Far_Cost2.Requery
   Far_Cost2.MoveLast
   If Far_Cost2.EOF Then
     WS_SALDO = 0
   Else
     WS_SALDO = Far_Cost2!FAR_STOCK
     WCOSPRO_SUP = Far_Cost2!FAR_COSPRO
     ws_cospro = Far_Cost2!FAR_COSPRO
   End If
   
   Do Until Far_Cost.EOF
      ProgBar.Value = Far_Cost.AbsolutePosition
      
      DoEvents
      WS_TIPO_CAMBIO = 1
      WS_SALDO = WS_SALDO + (Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad)
      If Far_Cost!FAR_MONEDA = "D" Then WS_TIPO_CAMBIO = JALAR(Far_Cost!FAR_fecha_compra, Far_Cost!FAR_fecha_compra)
      Far_Cost.Edit
      
      
      If Far_Cost!FAR_COSTEO = "A" And Far_Cost!far_estado <> "E" And Far_Cost!far_signo_arm = 1 And Far_Cost!FAR_tipmov = 101 And (LK_CODCIA = "04" Or LK_CODCIA = "10") Then
         ws_cospro = Format(Far_Cost!FAR_PRECIO / Far_Cost!FAR_equiv, "0.0000")
         GoTo NADA
      End If
      
      If Far_Cost!FAR_COSTEO = "A" And Far_Cost!far_estado <> "E" And Far_Cost!far_signo_arm = 1 Then
       ' ver si esta Fact esta confirmada
       'If Far_Cost!far_tipo_bloq_act3 <> "1" Then Call buscadoc(Far_Cost!far_numfac)
            
      ' Es temporal para regulizar costos
        ''If Far_Cost!FAR_TIPMOV = 6 And Far_Cost!FAR_PRECIO <> 0 And Far_Cost!FAR_NUM_LOTE <> 10 Then
        '' Far_Cost!FAR_PRECIO = ws_cospro
        '' GoTo NADA
        ''End If
         WPASA_COSPRO_ANT = ws_cospro
         wk_equiv = Far_Cost!FAR_equiv
         'If Far_Cost!FAR_TIPMOV = 101 And Far_Cost!far_cantidad = 100 And Far_Cost!far_codart = 79119 Then Stop
         'WS_VALOR1 = (WS_SALDO - Far_Cost!FAR_CANTIDAD)
         'ws_valor2 = ws_cospro
         ws_valor_ant = (WS_SALDO - Far_Cost!FAR_cantidad) * ws_cospro
         'ws_Valor3 = ws_valor_ant
         ws_igv = 0
         ws_compra = WS_TIPO_CAMBIO * Far_Cost!FAR_cantidad * Far_Cost!FAR_PRECIO / wk_equiv
         ws_ajuste = 0
         If LK_EMP = "3AA" Then
           WS_AJUSTE2 = Nulo_Valor0(Far_Cost!FAR_FLETE)
         Else
            WS_AJUSTE2 = 0
           If Far_Cost!FAR_tipmov <> 101 Then
             WS_AJUSTE2 = Nulo_Valor0(Far_Cost!FAR_FLETE) * WS_TIPO_CAMBIO
           End If
         End If
         WS_AJUSTE2 = WS_AJUSTE2 - (Nulo_Valor0(Far_Cost!FAR_DESCTO) * WS_TIPO_CAMBIO)
         ws_valor4 = Far_Cost!FAR_cantidad
         ws_compra = (ws_compra + ws_ajuste) + WS_AJUSTE2
         'ws_valor5 = ws_compra
         Far_Cost!far_precio_neto = ws_compra
         ultimo_costo = Far_Cost!FAR_PRECIO
         ultimo_cambio = Far_Cost!FAR_tipo_cambio
         If ws_valor_ant < 0 And LK_EMP = "3AA" Then
            ws_valor_ant = 0
         End If
         'ws_valor6 = (ws_valor_ant + ws_compra)
         If WS_SALDO <> 0 Then
            ws_cospro = (ws_valor_ant + ws_compra) / WS_SALDO
         Else
            ws_cospro = ((ws_valor_ant + ws_compra) / Far_Cost!FAR_cantidad)
         End If
         'ws_valor7 = ws_cospro
         If Far_Cost!FAR_COSPRO <> ws_cospro Then
'           Print Far_Cost!FAR_fecha_compra
'           Print Far_Cost!far_codart
          If Far_Cost!FAR_fecha_compra <> "20/06/01" Then
              If Nulo_Valors(art_LLAVE!ART_MONEDA) = "S" Then
                WMONEDA = "S/. "
              Else
                WMONEDA = "US$. "
             End If
           End If
            lstcosto.AddItem Trim(art_LLAVE!art_nombre)
            lstcosto.AddItem "Costo Anterior: " & WMONEDA & Format(Far_Cost!FAR_COSPRO, "0.0000")
            lstcosto.AddItem "Nuevo Costo: " & WMONEDA & Format(ws_cospro, "0.0000")
            'WTEXTO = WTEXTO + Trim(art_LLAVE!art_nombre) & Chr(13) & "Costo Anterior: " & WMONEDA & Format(Far_Cost!FAR_COSPRO, "0.0000") & Chr(13) & "Nuevo Costo: " & WMONEDA & Format(ws_cospro, "0.0000") & Chr(13)
         End If
         If LK_EMP = "HER" Then
          '  ws_cospro = Val(Format((Far_Cost!FAR_PRECIO / Far_Cost!FAR_equiv) * WS_TIPO_CAMBIO, "0.0000"))
         End If
      End If
NADA:
      Far_Cost!far_cospro_sup = WCOSPRO_SUP
      Far_Cost!FAR_COSPRO = ws_cospro
      Far_Cost!FAR_COSPRO_ANT = WPASA_COSPRO_ANT
      WCOSPRO_SUP = Far_Cost!FAR_COSPRO
      'If Far_Cost!FAR_TIPMOV = 100 Then
        'Far_Cost!FAR_PRECIO = ws_cospro
      'End If
      If Far_Cost!FAR_tipmov = 93 Then
        If Far_Cost!far_signo_arm = -1 Then
            Far_Cost!FAR_PRECIO = ws_cospro
            PSFAR_CAMBIO(0) = LK_CODCIA
            PSFAR_CAMBIO(1) = Far_Cost!far_numser
            PSFAR_CAMBIO(2) = Far_Cost!far_numfac
            far_cambiox.Requery
            If Not far_cambiox.EOF Then
                far_cambiox.Edit
                far_cambiox!FAR_PRECIO = Format(redondea(Far_Cost!FAR_cantidad * Far_Cost!FAR_COSPRO) / far_cambiox!FAR_cantidad, "0.0000")
                far_cambiox!FAR_COSTEO = "A"
                far_cambiox!FAR_COSTEO_REAL = "A"
                far_cambiox.Update
            End If
          End If
      End If
      If Far_Cost!FAR_tipmov = 100 Then
        'Far_Cost!FAR_PRECIO = ws_cospro
        'PSFAR_ENVIO(0) = Far_Cost!far_otra_cia
        'PSFAR_ENVIO(1) = Far_Cost!far_codart
        'PSFAR_ENVIO(2) = Far_Cost!far_numguia
        'Far_envio.Requery
        'Do Until Far_envio.EOF
        '    Far_envio.Edit
        '    Far_envio!FAR_PRECIO = Far_Cost!FAR_COSPRO
        '    Far_envio!FAR_costeo = "A"
        '    Far_envio!FAR_COSTEO_REAL = "A"
        '    Far_envio.Update
        'Far_envio.MoveNext
        'Loop
      End If
      If WS_TIPO_CAMBIO <> 1 Then Far_Cost!FAR_tipo_cambio = WS_TIPO_CAMBIO
      If Val(Far_Cost!FAR_STOCK) <> Val(WS_SALDO) Then
          Print "ACTUALIZA FAR_STOCK"
      End If
      Far_Cost!FAR_STOCK = WS_SALDO
   
      Far_Cost.Update
      Far_Cost.MoveNext
Loop
        POR(2).Caption = "Actualizando Costos de Articulo. . ."
        actualizar_precio art_LLAVE!art_key, ws_cospro, arm_llave!ARM_COSPRO
        
        'arm_llave.Edit
        'arm_llave!ARM_COSPRO = ws_cospro
        'If ultimo_costo <> 0 Then
        '  arm_llave!ARM_COSTO_ULT = redondea(ultimo_costo)
        'End If
        'arm_llave.Update
        
        ' TEMPORAL
        '----------
        If chelives.Value = 1 Then
          ACTLIVES = "PRE_COSTO = (" & ws_cospro & "* PRE_EQUIV)"
          pub_cadena = "UPDATE PRECIOS SET " & ACTLIVES & " WHERE  PRE_CODCIA = '03' AND PRE_CODART = " & art_LLAVE!art_key
          CN.Execute pub_cadena, rdExecDirect
        End If
        '--------------------

      
        
        indice = 0
        SQ_OPER = 2
        PUB_CODART = art_LLAVE!art_key
        pu_codcia = LK_CODCIA
        LEER_PRE_LLAVE
        If pre_mayor.EOF Then MsgBox "Error en Precios..."
            Do Until pre_mayor.EOF
              pre_mayor.Edit
              If ws_valor4 > 0 Then
                  pre_mayor!PRE_COSTO = Format(ws_compra * pre_mayor!PRE_EQUIV / ws_valor4, "0.0000")
              Else
                  pre_mayor!PRE_COSTO = Nulo_Valor0(pre_mayor!PRE_cosTO_ant)
              End If

              pre_mayor.Update
              pre_mayor.MoveNext
         Loop

OTRO_ART:
    ProgBar.max = Far_Compr.RowCount
    Far_Compr.MoveNext
Loop
pasa:

'If FLAG_93 <> "A" Then
If FLAG_93 = "z" Then
 FLAG_93 = "A"
 pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND ((FAR_TIPMOV = 93 AND FAR_SIGNO_ARM = 1) OR FAR_TIPMOV = 100)  AND (FAR_COSTEO_REAL = 'A'  OR FAR_FECHA_COMPRA <> FAR_FECHA  OR FAR_MONEDA= 'D')  AND FAR_ESTADO <>'E'  ORDER BY FAR_CODART" ' , FAR_CODCIA,FAR_FECHA_CAN"
 If CHEART.Value = 1 Then
   pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODART =  " & WCODART & " AND FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ?  AND ((FAR_TIPMOV = 93 AND FAR_SIGNO_ARM = 1) OR FAR_TIPMOV = 100)  AND (FAR_COSTEO_REAL = 'A' )  AND FAR_ESTADO<>'E'  ORDER BY FAR_CODART" ' , FAR_CODCIA,FAR_FECHA_CAN"
 End If
 If Val(TLINEA.Text) <> 0 Then
  WLINEA = Val(TLINEA.Text)
  ' PRUEBA
  pub_cadena = "SELECT Distinct FAR_CODART FROM FACART, ARTI WHERE (ART_CODCIA = FAR_CODCIA) AND (FAR_CODART = ART_KEY) AND ART_LINEA = " & WLINEA & "  AND FAR_TIPMOV = 93 AND FAR_SIGNO_ARM = 1 AND  FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND (FAR_COSTEO_REAL = 'A'  OR FAR_FECHA_COMPRA <> FAR_FECHA  OR FAR_MONEDA= 'D' )  AND FAR_ESTADO<>'E'  ORDER BY FAR_CODART"
 End If


 Set PSFAR_COMPR = CN.CreateQuery("", pub_cadena)
 PSFAR_COMPR(0) = LK_CODCIA
 PSFAR_COMPR(1) = LK_FECHA_DIA
 PSFAR_COMPR(2) = LK_FECHA_DIA
 Set Far_Compr = PSFAR_COMPR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PSFAR_COMPR.rdoParameters(0) = LK_CODCIA
 PSFAR_COMPR.rdoParameters(1) = CDate(txtfecha.Text)
 If fechafinal.Visible Then
   PSFAR_COMPR.rdoParameters(2) = fechafinal.Text
 Else
   PSFAR_COMPR.rdoParameters(2) = LK_FECHA_DIA
 End If
 GoTo CALCULA_93
End If

cabeprecios
Muestra_Preicios 0

        


'ProgBar.Value = ProgBar.Value + 1
'CLOSE_LOG

POR(2).Caption = "Terminando Proceso. . ."
DoEvents
flag_avanza:
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_PAR_LLAVE
par_llave.Edit
par_llave!par_flag_cierre = 8
par_llave!par_flag_costos = 9
par_llave.Update
''ProgBar.Value = ProgBar.Value + 1
salida:
CHEART.Value = 0
'If Trim(WTEXTO) <> "" Then
 fralista.Visible = True
 'MsgBox " P r o c e s o   d e   C o s t e o ", 48, Pub_Titulo ' & Chr(13) & WTEXTO, vbInformation, Pub_Titulo
 
'End If
POR(2).Caption = ""
MsgBox "Proceso Terminado.", vbInformation, Pub_Titulo
Command1.Enabled = True
ProgBar.Visible = False
'Unload Costos
'Unload FORMGEN
sale:
Exit Sub

PROCESO_RECECPION:
If LK_EMP_PTO = "A" Then
  If LK_CODCIA <> "00" Then
    SQ_OPER = 1
    PUB_CODCIA = "00"
    LEER_PAR_LLAVE
    If par_llave!par_flag_costos <> "9" Then
       MsgBox "Costear Primero en Cia Principal ..."
       Unload FrmNC
       GoTo fin
    End If
End If
End If

pub_cadena = "SELECT * FROM FFFART WHERE FFF_CODCIA=?  ORDER BY FFF_CODCIA"
Set PSFAR_COSTO = CN.CreateQuery("", pub_cadena)
PSFAR_COSTO(0) = LK_CODCIA
Set Far_Cost = PSFAR_COSTO.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT FAR_COSPRO, FAR_PRECIO, FAR_COSTEO, FAR_COSTEO_REAL FROM FACART WHERE FAR_TIPMOV =? AND FAR_CODCIA= ? AND FAR_NUMSER = ? AND FAR_FBG = ? AND FAR_NUMFAC = ? AND FAR_NUMSEC = ?   ORDER BY FAR_CODCIA"
Set PSFAR_COMPR = CN.CreateQuery("", pub_cadena)
PSFAR_COMPR(0) = 0
PSFAR_COMPR(1) = LK_CODCIA
PSFAR_COMPR(2) = 0
PSFAR_COMPR(3) = " "
PSFAR_COMPR(4) = 0
PSFAR_COMPR(5) = 0
Set Far_Compr = PSFAR_COMPR.OpenResultset(rdOpenKeyset, rdConcurValues)

PSFAR_COSTO(0) = LK_CODCIA
Far_Cost.Requery
    
Do Until Far_Cost.EOF
    PSFAR_COMPR(0) = Far_Cost!FFF_TIPMOV_R
    PSFAR_COMPR(1) = Far_Cost!FFF_codcia_R
    PSFAR_COMPR(2) = Far_Cost!FFF_numser_R
    PSFAR_COMPR(3) = Far_Cost!FFF_fbg_R
    PSFAR_COMPR(4) = Far_Cost!FFF_NUMFAC_R
    PSFAR_COMPR(5) = Far_Cost!FFF_NUMSEC_R
    Far_Compr.Requery
    If Far_Compr.EOF Then
      MsgBox "Registro No puede Costear: TIPMOV:" & Far_Cost!FFF_TIPMOV_R & " CODCIA:" & Far_Cost!FFF_codcia_R & " NUMSER:" & Far_Cost!FFF_numser_R & " FBG:" & Far_Cost!FFF_fbg_R & " NUMFAC:" & Far_Cost!FFF_NUMFAC_R & " NUMSEC:" & Far_Cost!FFF_NUMSEC_R, 48, Pub_Titulo
      GoTo SIGUE_R
    End If
    WPASA_PRECIO = Far_Compr!FAR_COSPRO ' JALO SU COSTO PROMEDIO
    
    PSFAR_COMPR(0) = Far_Cost!FFF_TIPMOV
    PSFAR_COMPR(1) = Far_Cost!FFF_codcia
    PSFAR_COMPR(2) = Val(Far_Cost!FFF_numser)
    PSFAR_COMPR(3) = Far_Cost!FFF_fbg
    PSFAR_COMPR(4) = Far_Cost!FFF_NUMFAC
    PSFAR_COMPR(5) = Far_Cost!FFF_NUMSEC
    Far_Compr.Requery
    If Far_Compr.EOF Then
      MsgBox "Registro No puede Costear: TIPMOV:" & Far_Cost!FFF_TIPMOV & " CODCIA:" & Far_Cost!FFF_codcia & " NUMSER:" & Far_Cost!FFF_numser & " FBG:" & Far_Cost!FFF_fbg & " NUMFAC:" & Far_Cost!FFF_NUMFAC & " NUMSEC:" & Far_Cost!FFF_NUMSEC, 48, Pub_Titulo
      GoTo SIGUE_R
    End If
    Far_Compr.Edit
    Far_Compr!FAR_PRECIO = WPASA_PRECIO ' ASIGNO EL COSTO AL PRECIO
    Far_Compr!FAR_COSTEO = "A"
    Far_Compr!FAR_COSTEO_REAL = "A"
    Far_Compr.Update
    
SIGUE_R:
    Far_Cost.MoveNext
Loop

Return
fin:
ProgBar.Visible = False
POR(0).Visible = False
POR(1).Visible = False
POR(2).Visible = False
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Unload FrmNC
End Sub


Private Sub CHEART_Click()
If CHEART.Value = 1 Then
  TART.Visible = True
  LART.Visible = True
  TART.SetFocus
Else
  TART.Visible = False
  LART.Visible = False
End If
End Sub

Private Sub Form_Activate()
CenterMe FrmNC
If txtfecha.Visible Then
 txtfecha_KeyPress 13
 Azul2 txtfecha, txtfecha
End If
If Trim(FORMGEN.i_fbg) <> "" And Trim(FORMGEN.i_numfac_c) <> "" Then
  opnc(1).Value = True
  opnc_Click 1
Else
  opnc_Click 0
End If
If LK_CODTRA = 2774 Then
    opnc(0).Enabled = False
    opnc(1).Enabled = False
    opnc(2).Enabled = False
    opnc(3).Enabled = False
    opnc(4).Value = True
    opnc_Click 4
End If

End Sub

Private Sub Form_Load()

tf1.Mask = "##/##/####"
tf1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
tf2.Mask = "##/##/####"
tf2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
tf1.Visible = True
tf2.Visible = True
End Sub

Private Sub gridcabe_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub

'If Trim(gridcabe.TextMatrix(gridcabe.Row, 9)) <> "8" Then
'  If Trim(gridcabe.TextMatrix(gridcabe.Row, 0)) = "" Then Exit Sub
'  If Trim(gridcabe.TextMatrix(gridcabe.Row, 1)) <> "" And gridcabe.Col = 2 Or gridcabe.Col = 3 Then GoTo leer
'  If Trim(gridcabe.TextMatrix(gridcabe.Row, 8)) <> "0" Then Exit Sub
'End If


If gridcabe.COL = 10 Or gridcabe.COL = 11 Then
Else
 Exit Sub
End If
'   a = Val(gridcabe.TextMatrix(gridcabe.Row - 1, 0))
'   a = a + 1
'  gridcabe.TextMatrix(gridcabe.Row, 0) = a
'End If
'If WMODO = "I" Or WMODO = "C" Then
    TEXTOVARpre.Left = gridcabe.Left + gridcabe.CellLeft
    TEXTOVARpre.Width = gridcabe.CellWidth
    TEXTOVARpre.Height = gridcabe.CellHeight
    TEXTOVARpre.Top = gridcabe.Top + gridcabe.CellTop
    TEXTOVARpre.Text = gridcabe.TextMatrix(gridcabe.Row, gridcabe.COL)
    TEXTOVARpre.Visible = True
    Azul3 TEXTOVARpre, TEXTOVARpre
    TEXTOVARpre.SetFocus
'End If
End Sub

Private Sub gridcabe_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim sale As Boolean
Dim Wsec

'If WMODO = "C" Then Exit Sub

'If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
'End If


If Left(gridcabe.TextMatrix(gridcabe.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
  'If WMODO <> "C" Then Exit Sub
  tC = gridcabe.COL
  For fila = 1 To gridcabe.Cols - 1
      gridcabe.COL = fila
      If gridcabe.CellBackColor = QBColor(12) Then
         gridcabe.CellBackColor = QBColor(15)
         gridcabe.TextMatrix(gridcabe.Row, 9) = "9"
      Else
         gridcabe.CellBackColor = QBColor(12)
         gridcabe.TextMatrix(gridcabe.Row, 9) = "-1"
      End If
  Next fila
  gridcabe.COL = tC
  gridcabe.SetFocus
  Exit Sub
End If
If KeyCode = 45 Then
    Wsec = Wsec + 1
    If Trim(gridcabe.TextMatrix(gridcabe.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(gridcabe.TextMatrix(gridcabe.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(gridcabe.TextMatrix(gridcabe.Row, 4)) = 0 And Val(gridcabe.TextMatrix(gridcabe.Row, 5)) = 0 Then Exit Sub
    gridcabe.AddItem "", gridcabe.Row + 1
    gridcabe.TextMatrix(gridcabe.Row + 1, 0) = "MAN. " & Format(gridcabe.TextMatrix(gridcabe.Row, 10), "dd/mm/yyyy")
    gridcabe.TextMatrix(gridcabe.Row + 1, 6) = Wsec
    gridcabe.TextMatrix(gridcabe.Row + 1, 8) = gridcabe.TextMatrix(gridcabe.Row, 8)
    gridcabe.TextMatrix(gridcabe.Row + 1, 3) = gridcabe.TextMatrix(gridcabe.Row, 3)
    gridcabe.TextMatrix(gridcabe.Row + 1, 7) = gridcabe.TextMatrix(gridcabe.Row, 7)
    gridcabe.TextMatrix(gridcabe.Row + 1, 10) = gridcabe.TextMatrix(gridcabe.Row, 10)
    gridcabe.TextMatrix(gridcabe.Row + 1, 11) = "8"
    gridcabe.Row = gridcabe.Row + 1
    gridcabe.COL = 1
    gridcabe.SetFocus
End If
Exit Sub
If KeyCode = 46 Then
If gridcabe.Rows <= 3 Then
Else
   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(gridcabe.TextMatrix(gridcabe.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     gridcabe.SetFocus
     Exit Sub
   Else
     gridcabe.RowHeight(gridcabe.Row) = 1
     gridcabe.Row = gridcabe.Row + 1
    
   'gridcabe.RemoveItem (gridcabe.Row)
   'gridcabe.Refresh
   gridcabe.SetFocus
   End If
End If
End If
'gridcabe.SetFocus
Exit Sub



End Sub

Private Sub gridcabe_Scroll()
TEXTOVARpre.Visible = False
End Sub

Private Sub i_codart2_Change()
If i_codart2.Text = "" Then
  LBLNOMARTI.Caption = ""
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
        LBLNOMARTI.Caption = art_LLAVE!art_nombre
     End If

     
  
  DoEvents
fin:

End Sub
Private Sub i_codart2_KeyPress(KeyAscii As Integer)
Dim VAR_ACTIVAR As Integer
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
    WCOD_ORIGINAL = art_LLAVE!art_key
    LBLNOMARTI.Caption = Trim(art_LLAVE!art_nombre)
    i_codart2.Text = WCOD_ORIGINAL
    'i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
    ListView1.Visible = False
    GoTo SIGUE_DATO
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
     LBLNOMARTI.Caption = Trim(art_llave_alt!art_nombre)
     ListView1.Visible = False
     GoTo SIGUE_DATO
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
      WCOD_ORIGINAL = art_LLAVE!art_key
      LBLNOMARTI.Caption = Trim(art_LLAVE!art_nombre)
      ListView1.Visible = False
      i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      ListView1.Visible = False
      GoTo SIGUE_DATO
      Exit Sub
    Else
      Exit Sub
    End If
    
  End If
End If

SIGUE_DATO:
Azul2 tf1, tf1

Exit Sub
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
Dim VAR_ACTIVAR As Integer
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
      'If chedes.Value = 1 Then
        archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
      'Else
      '  archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & ws_codcia & "' AND  ART_SITUACION <> 1 AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
      'End If
    Else
      numarchi = 0
      'If chedes.Value = 1 Then
       archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      'Else
      ' archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_SITUACION <> 1 AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      'End If
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



Private Sub opnc_Click(Index As Integer)
If Index = 0 Then
    frapro.Visible = True
    frafechas.Visible = True
    fradoc.Visible = False
    fralin.Visible = False
    i_codart2.Text = ""
    LBLNOMARTI.Caption = ""
    Azul i_codart2, i_codart2
ElseIf Index = 1 Then
    fradoc.Visible = True
    frafechas.Visible = False
    frapro.Visible = False
    fralin.Visible = False
    tfbg.Clear
    tfbg.AddItem "F"
    tfbg.AddItem "B"
    tfbg.AddItem "G"
    tfbg.AddItem "P"
    If Trim(FORMGEN.i_fbg) <> "" And Trim(FORMGEN.i_numfac_c) <> "" Then
        If FORMGEN.i_fbg = "F" Then
        FrmNC.tfbg.ListIndex = 0
        ElseIf FORMGEN.i_fbg = "B" Then
        FrmNC.tfbg.ListIndex = 1
        ElseIf FORMGEN.i_fbg = "G" Then
        FrmNC.tfbg.ListIndex = 2
        ElseIf FORMGEN.i_fbg = "P" Then
        FrmNC.tfbg.ListIndex = 3
        End If
        FrmNC.tserie.Text = Trim(FORMGEN.i_numser_c)
        FrmNC.tnumero.Text = Trim(FORMGEN.i_numfac_c)
    End If
    tfbg.SetFocus
ElseIf Index = 2 Then
    frafechas.Visible = True
    fradoc.Visible = False
    frapro.Visible = False
    fralin.Visible = True
    lineas.Clear
    PUB_CODCIA = LK_CODCIA
    LLENADOS lineas, 122
ElseIf Index = 3 Then
    frafechas.Visible = True
    fradoc.Visible = False
    frapro.Visible = False
    fralin.Visible = False
ElseIf Index = 4 Then
    frafechas.Visible = True
    fradoc.Visible = False
    frapro.Visible = False
    fralin.Visible = False
End If
End Sub

Private Sub textovarpre_Change()
gridcabe.Text = Format(TEXTOVARpre.Text, "0.0000")
End Sub

Private Sub textovarpre_GotFocus()
 temporal = gridcabe.TextMatrix(gridcabe.Row, gridcabe.COL)
End Sub

Private Sub textovarpre_KeyPress(KeyAscii As Integer)
Dim VALOR As Currency
If KeyAscii = 27 Then
  TEXTOVARpre.Text = temporal
  TEXTOVARpre.Visible = False
  gridcabe.SetFocus
  Exit Sub
End If
'If gridcabe.COL = 1 Then Consistencias gridcabe, TEXTOVARpre, KeyAscii
'If gridcabe.COL = 4 Then Consistencias gridcabe, TEXTOVARpre, KeyAscii
Consistencias gridcabe, TEXTOVARpre, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If gridcabe.COL = 1 Or gridcabe.COL = 4 Then
  If Val(TEXTOVARpre.Text) > 99 Then
    Azul3 TEXTOVARpre, TEXTOVARpre
    Exit Sub
  End If
End If

'PUB_CAL_INI = gridcabe.TextMatrix(gridcabe.Row, 2)
'PUB_CAL_FIN = gridcabe.TextMatrix(gridcabe.Row, 2)
'pu_codcia = LK_CODCIA
'SQ_OPER = 1
'PUB_CODCIA = LK_CODCIA
'LEER_CAL_LLAVE
'cal_llave.Edit
'If gridcabe.COL = 4 Then
'   cal_llave!cal_tipo_cambio = Val(TEXTOVARpre.Text)
'End If
'If gridcabe.COL = 1 Then
'   cal_llave!CAL_TC_MERCA = Val(TEXTOVARpre.Text)
'   If Format(LK_FECHA_DIA, "dd/mm/yyyy") = Format(gridcabe.TextMatrix(gridcabe.Row, 0), "dd/mm/yyyy") Then
'      LK_TIPO_CAMBIO = Val(TEXTOVARpre.Text)
'      'MDIForm1.StatusBar1.Panels(3).Text = "T.C.= S/. " + Format(LK_TIPO_CAMBIO, "0.0000")
'   End If
'End If
'If gridcabe.COL = 5 Then
'   cal_llave!cal_tc_ingre = Val(TEXTOVARpre.Text)
'End If
'If gridcabe.COL = 6 Then
'   cal_llave!cal_tc_salid = Val(TEXTOVARpre.Text)
'End If

'cal_llave.Update


'VALOR = (Val(TEXTOVARpre.Text) * 100) / Val(gridcabe.TextMatrix(gridcabe.Row, 4)) - 100
'If VALOR = -100 Then VALOR = 0

If gridcabe.COL = 10 Then
  gridcabe.CellForeColor = QBColor(9)
  If (Val(gridcabe.TextMatrix(gridcabe.Row, 8)) - Val(TEXTOVARpre.Text)) < 0 Then
     MsgBox "No procede, el precio es mayor que el actual", 48, Pub_Titulo
     TEXTOVARpre.Visible = True
     Azul3 TEXTOVARpre, TEXTOVARpre
     Exit Sub
  End If
  VALOR = Val(gridcabe.TextMatrix(gridcabe.Row, 10)) * Val(gridcabe.TextMatrix(gridcabe.Row, 6))
  gridcabe.TextMatrix(gridcabe.Row, 11) = Format(VALOR, "0.00")
  VALOR = Val(gridcabe.TextMatrix(gridcabe.Row, 9)) - Val(gridcabe.TextMatrix(gridcabe.Row, 11))
  gridcabe.TextMatrix(gridcabe.Row, 12) = Format(VALOR, "0.00")
  If VALOR = 0 Then
    gridcabe.TextMatrix(gridcabe.Row, 12) = ""
    gridcabe.CellForeColor = QBColor(0)
  End If
  Calcula_NC
End If



If gridcabe.Row >= gridcabe.Rows - 1 Then
Else
  gridcabe.Row = gridcabe.Row + 1
End If
gridcabe.SetFocus
TEXTOVARpre.Visible = False

fin:

End Sub


Private Sub gridiGV_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub

'If Trim(gridigv.TextMatrix(gridigv.Row, 9)) <> "8" Then
'  If Trim(gridigv.TextMatrix(gridigv.Row, 0)) = "" Then Exit Sub
'  If Trim(gridigv.TextMatrix(gridigv.Row, 1)) <> "" And gridigv.Col = 2 Or gridigv.Col = 3 Then GoTo leer
'  If Trim(gridigv.TextMatrix(gridigv.Row, 8)) <> "0" Then Exit Sub
'End If


'If gridigv.Col = 1 And WMODO = "I" Then
'   a = Val(gridigv.TextMatrix(gridigv.Row - 1, 0))
'   a = a + 1
'  gridigv.TextMatrix(gridigv.Row, 0) = a
'End If
'If WMODO = "I" Or WMODO = "C" Then
    TEXTOVAR.Left = gridigv.Left + gridigv.CellLeft
    TEXTOVAR.Width = gridigv.CellWidth
    TEXTOVAR.Height = gridigv.CellHeight
    TEXTOVAR.Top = gridigv.Top + gridigv.CellTop
    TEXTOVAR.Text = gridigv.TextMatrix(gridigv.Row, gridigv.COL)
    TEXTOVAR.Visible = True
    Azul3 TEXTOVAR, TEXTOVAR
    TEXTOVAR.SetFocus
'End If
End Sub

Private Sub gridiGV_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim sale As Boolean
Dim Wsec

'If WMODO = "C" Then Exit Sub

'If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
'End If


If Left(gridigv.TextMatrix(gridigv.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
  'If WMODO <> "C" Then Exit Sub
  tC = gridigv.COL
  For fila = 1 To gridigv.Cols - 1
      gridigv.COL = fila
      If gridigv.CellBackColor = QBColor(12) Then
         gridigv.CellBackColor = QBColor(15)
         gridigv.TextMatrix(gridigv.Row, 9) = "9"
      Else
         gridigv.CellBackColor = QBColor(12)
         gridigv.TextMatrix(gridigv.Row, 9) = "-1"
      End If
  Next fila
  gridigv.COL = tC
  gridigv.SetFocus
  Exit Sub
End If
If KeyCode = 45 Then
    Wsec = Wsec + 1
    If Trim(gridigv.TextMatrix(gridigv.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(gridigv.TextMatrix(gridigv.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(gridigv.TextMatrix(gridigv.Row, 4)) = 0 And Val(gridigv.TextMatrix(gridigv.Row, 5)) = 0 Then Exit Sub
    gridigv.AddItem "", gridigv.Row + 1
    gridigv.TextMatrix(gridigv.Row + 1, 0) = "MAN. " & Format(gridigv.TextMatrix(gridigv.Row, 10), "dd/mm/yyyy")
    gridigv.TextMatrix(gridigv.Row + 1, 6) = Wsec
    gridigv.TextMatrix(gridigv.Row + 1, 8) = gridigv.TextMatrix(gridigv.Row, 8)
    gridigv.TextMatrix(gridigv.Row + 1, 3) = gridigv.TextMatrix(gridigv.Row, 3)
    gridigv.TextMatrix(gridigv.Row + 1, 7) = gridigv.TextMatrix(gridigv.Row, 7)
    gridigv.TextMatrix(gridigv.Row + 1, 10) = gridigv.TextMatrix(gridigv.Row, 10)
    gridigv.TextMatrix(gridigv.Row + 1, 11) = "8"
    gridigv.Row = gridigv.Row + 1
    gridigv.COL = 1
    gridigv.SetFocus
End If
Exit Sub
If KeyCode = 46 Then
If gridigv.Rows <= 3 Then
Else
   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(gridigv.TextMatrix(gridigv.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     gridigv.SetFocus
     Exit Sub
   Else
     gridigv.RowHeight(gridigv.Row) = 1
     gridigv.Row = gridigv.Row + 1
    
   'gridiGV.RemoveItem (gridiGV.Row)
   'gridiGV.Refresh
   gridigv.SetFocus
   End If
End If
End If
'gridiGV.SetFocus
Exit Sub



End Sub

Private Sub gridigv_Scroll()
TEXTOVAR.Visible = False
End Sub


Private Sub textovar_Change()
gridigv.Text = Format(TEXTOVAR.Text, "0.0000")
End Sub

Private Sub TEXTOVAR_GotFocus()
 temporal = gridigv.TextMatrix(gridigv.Row, gridigv.COL)
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  TEXTOVAR.Text = temporal
  TEXTOVAR.Visible = False
  gridigv.SetFocus
  Exit Sub
End If
If gridigv.COL = 1 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If gridigv.COL = 4 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If gridigv.COL = 5 Or gridigv.COL = 6 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If gridigv.COL = 1 Or gridigv.COL = 4 Then
  If Val(TEXTOVAR.Text) > 99 Then
    Azul3 TEXTOVAR, TEXTOVAR
    Exit Sub
  End If
End If

PUB_CAL_INI = gridigv.TextMatrix(gridigv.Row, 2)
PUB_CAL_FIN = gridigv.TextMatrix(gridigv.Row, 2)
pu_codcia = LK_CODCIA
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_CAL_LLAVE
cal_llave.Edit
If gridigv.COL = 4 Then
   cal_llave!cal_tipo_cambio = Val(TEXTOVAR.Text)
End If
If gridigv.COL = 1 Then
   cal_llave!CAL_TC_MERCA = Val(TEXTOVAR.Text)
   If Format(LK_FECHA_DIA, "dd/mm/yyyy") = Format(gridigv.TextMatrix(gridigv.Row, 0), "dd/mm/yyyy") Then
      LK_TIPO_CAMBIO = Val(TEXTOVAR.Text)
      'MDIForm1.StatusBar1.Panels(3).Text = "T.C.= S/. " + Format(LK_TIPO_CAMBIO, "0.0000")
   End If
End If
If gridigv.COL = 5 Then
   cal_llave!cal_tc_ingre = Val(TEXTOVAR.Text)
End If
If gridigv.COL = 6 Then
   cal_llave!cal_tc_salid = Val(TEXTOVAR.Text)
End If

cal_llave.Update
If gridigv.Row >= gridigv.Rows - 1 Then
Else
  gridigv.Row = gridigv.Row + 1
End If
gridigv.SetFocus
TEXTOVAR.Visible = False

fin:

End Sub


Private Sub tf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Azul2 tf2, tf2
End If
End Sub

Private Sub tf2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cmdmst.SetFocus
End If
End Sub

Private Sub tfbg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul tserie, tserie

End Sub

Private Sub Timer1_Timer()
'lblcierre.Visible = Not lblcierre.Visible
End Sub

Public Sub Muestra_tc(wfecha_ini As Date)
Dim wdiaI, wdiaF As String
Dim wmesM As String
gridigv.Clear
gridigv.Cols = 7
gridigv.Rows = 2
gridigv.ColWidth(0) = 1000
gridigv.ColWidth(1) = 900
gridigv.ColWidth(2) = 0
gridigv.ColWidth(3) = 0
gridigv.ColWidth(4) = 900
gridigv.ColWidth(5) = 900
gridigv.ColWidth(6) = 900

gridigv.TextMatrix(0, 0) = "Fecha"
gridigv.TextMatrix(0, 1) = "Publicación"
gridigv.TextMatrix(1, 0) = "-"
gridigv.Row = 1
gridigv.COL = 1

gridigv.TextMatrix(1, 1) = "Compra"
gridigv.CellForeColor = QBColor(9)
gridigv.CellFontBold = True
gridigv.CellAlignment = 4
gridigv.TextMatrix(0, 4) = "---------"
gridigv.COL = 4
gridigv.TextMatrix(1, 4) = "Venta"
gridigv.CellForeColor = QBColor(4)
gridigv.CellFontBold = True
gridigv.CellAlignment = 4

gridigv.COL = 5
gridigv.TextMatrix(0, 5) = "Vigente"
gridigv.TextMatrix(1, 5) = "Ingresos"
gridigv.CellForeColor = QBColor(6)
gridigv.CellFontBold = True
gridigv.CellAlignment = 4
gridigv.COL = 6
gridigv.TextMatrix(0, 6) = "--------"
gridigv.TextMatrix(1, 6) = "Egresos"
gridigv.CellForeColor = QBColor(6)
gridigv.CellFontBold = True
gridigv.CellAlignment = 4



PUB_CAL_INI = wfecha_ini ' wdiaI & "/" & wmesM & "/" & PUB_CAL_ANO
PUB_CAL_FIN = LK_FECHA_DIA 'wdiaF & "/" & wmesM & "/" & PUB_CAL_ANO
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
cal_llave.MoveFirst
fila = 1
gridigv.Rows = 2
Do Until cal_llave.EOF
  fila = fila + 1
  gridigv.Rows = gridigv.Rows + 1
  gridigv.RowHeight(gridigv.Rows - 1) = 285
  gridigv.TextMatrix(fila, 0) = Format(cal_llave!cal_fecha, "dd/mm/yyyy")
  gridigv.TextMatrix(fila, 1) = Format(cal_llave!CAL_TC_MERCA, "0.0000")
  gridigv.TextMatrix(fila, 2) = Format(cal_llave!cal_fecha, "dd/mm/yyyy")
  gridigv.TextMatrix(fila, 4) = Format(cal_llave!cal_tipo_cambio, "0.0000")
  gridigv.TextMatrix(fila, 5) = Format(cal_llave!cal_tc_ingre, "0.0000")
  gridigv.TextMatrix(fila, 6) = Format(cal_llave!cal_tc_salid, "0.0000")
  cal_llave.MoveNext
Loop
gridigv.Visible = True
gridigv.COL = 1
If gridigv.Rows > 2 Then gridigv.Row = 2
If gridigv.Visible Then gridigv.SetFocus

End Sub

Public Function CONSIS_TC(wfecha_ini As Date) As Boolean
Dim WRES As String
PUB_CAL_INI = wfecha_ini ' wdiaI & "/" & wmesM & "/" & PUB_CAL_ANO
PUB_CAL_FIN = LK_FECHA_DIA
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
WRES = "Las Siguientes Fechas no tienen tipo de cambio:" & Chr(13)
fila = 0
Do Until cal_llave.EOF
If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 And cal_llave!CAL_LABORABLE = "S" Then
  WRES = WRES + Format(cal_llave!cal_fecha, "dd/mm/yyyy") + " "
  fila = 1
End If
cal_llave.MoveNext
Loop
WRES = WRES + "." + Chr(13) + "Consulte su tabla de tipo de cambios"
If fila = 1 Then
  MsgBox WRES, 48, Pub_Titulo
  Muestra_tc wfecha_ini
  CONSIS_TC = False
Else
   CONSIS_TC = True
End If

End Function
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


Private Sub tnumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cmdmst.SetFocus
End If

End Sub

Private Sub tserie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul tnumero, tnumero
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And IsDate(txtfecha.Text) = True Then
   Muestra_tc txtfecha.Text
Else
   gridigv.Clear
End If
End Sub

Public Function JALAR(wfecha_ini As Date, wfecha_fin As Date) As Currency
PUB_CAL_INI = wfecha_ini
PUB_CAL_FIN = wfecha_fin
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
JALAR = cal_llave!cal_tipo_cambio
End Function

Private Sub txtfecha_LostFocus()
If txtfecha.Text = "" Then Exit Sub
If Not IsDate(txtfecha.Text) Then
  MsgBox "Fecha no procede. ", 48, Pub_Titulo
  Azul2 txtfecha, txtfecha
  Exit Sub
End If
If txtfecha.Text > LK_FECHA_DIA Then
  MsgBox "Fecha no Puede ser mayor a la del dia.", 48, Pub_Titulo
  Azul2 txtfecha, txtfecha
  Exit Sub
End If


End Sub


Public Function CONSIS_PERIODOS(wfecha_ini As Date) As Boolean
Dim WRES As Integer
Dim wfecha As Date
Dim WNUMERO As Currency
CONSIS_PERIODOS = True
SQ_OPER = 1
PUB_TIPREG = 60

WRES = DateDiff("M", wfecha_ini, FrmNC.fechafinal.Text) + 1
wfecha = wfecha_ini
For fila = 1 To WRES
 If fila <> 1 Then
   wfecha = DateAdd("M", 1, wfecha_ini)
 End If

WNUMERO = Format(wfecha, "mm") & Format(wfecha, "yyyy")
PUB_NUMTAB = Val(WNUMERO)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If Not tab_llave.EOF Then
  If tab_llave!tab_codart = 1 Then
    MsgBox "E l  p e r i o d o: " & Chr(13) & UCase(Format(wfecha, "mmmm - yyyy")) & Chr(13) & "e s t a   C e r r a d o.", vbInformation, Pub_Titulo
    CONSIS_PERIODOS = False
    Exit Function
  End If
End If
 
 
Next fila
End Function


Private Function buscadoc(nro_interno As Currency)
Dim I As Integer
For I = 0 To lstdoc.ListCount - 1
 If Val(Left(lstdoc.List(I), 5)) = Val(nro_interno) Then
  Exit Function
 End If
Next I
lstdoc.AddItem Format(nro_interno, "00000")
End Function

Public Sub cabeprecios()

End Sub

Public Sub actualizar_precio(WCODART As Currency, w_newcosto As Currency, w_antwcosto As Currency)
Dim cant As Integer
Dim WS_COSTO As Currency
Dim WS_PRECIO As Currency
Dim res_costo_UNI As Currency
Dim res_costo_CAJA As Currency

Dim res_new_costo_UNI As Currency
Dim res_new_costo_CAJA As Currency

Dim cto_llave As rdoResultset
Dim PSCTOLLAVE  As rdoQuery

pub_cadena = "SELECT COUNT(PRE_CODART) AS CANTIDAD FROM PRECIOS WHERE PRE_CODART = ? AND PRE_CODCIA = ?"
Set PSCTOLLAVE = CN.CreateQuery("", pub_cadena)
PSCTOLLAVE(0) = 0
PSCTOLLAVE(1) = 0
Set cto_llave = PSCTOLLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)


  

pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ? "
Set PSACTLLAVE = CN.CreateQuery("", pub_cadena)
PSACTLLAVE(0) = 0
PSACTLLAVE(1) = 0
Set actllave_precio = PSACTLLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

PSACTLLAVE(0) = LK_CODCIA
PSACTLLAVE(1) = WCODART
actllave_precio.Requery
Do Until actllave_precio.EOF
    PSCTOLLAVE(0) = actllave_precio!PRE_CODART
    PSCTOLLAVE(1) = actllave_precio!PRE_CODCIA
    cto_llave.Requery
    cant = 0
    If Not cto_llave.EOF Then
      cant = cto_llave!cantidad
    End If
    
  'If actllave_precio!PRE_FLAG_UNIDAD = "A" Then
    res_costo_CAJA = Format(w_newcosto * actllave_precio!PRE_EQUIV, "0.00")
    WS_COSTO = 0
    'If Val(Format(actllave_precio!PRE_COSTO, "0.00")) = Val(res_costo_CAJA) Then
     ' pasa normal
    ' GoTo OTROPRECIO
   ' Else
      WS_COSTO = Format(w_newcosto * actllave_precio!PRE_EQUIV, "0.0000")
    'End If
    
    gridcabe.Rows = gridcabe.Rows + 1
    new_precio.AddNew
    
    new_precio!PRE_CODCIA = actllave_precio!PRE_CODCIA
    new_precio!PRE_CODART = actllave_precio!PRE_CODART
    new_precio!PRE_SECUENCIA = actllave_precio!PRE_SECUENCIA
    new_precio!pre_FLAG_UNIDAD = actllave_precio!pre_FLAG_UNIDAD
    new_precio!PRE_cosTO_ant = w_antwcosto * actllave_precio!PRE_EQUIV
    new_precio!PRE_COSTO_NEW = WS_COSTO
    If actllave_precio!PRE_POR1 <> 0 Then
      WS_PRECIO = Val(WS_COSTO * (1 + (actllave_precio!PRE_POR1 / 100)))
      WS_PRECIO = REDONDEA_PRECIO(WS_PRECIO, actllave_precio!PRE_EQUIV, cant)
      new_precio!PRE_PRE1 = WS_PRECIO
    Else
      new_precio!PRE_PRE1 = 0
    End If
    
    '------------------------------
    If actllave_precio!PRE_POR2 <> 0 Then
     WS_PRECIO = Val(WS_COSTO * (1 + (actllave_precio!PRE_POR2 / 100)))
     WS_PRECIO = REDONDEA_PRECIO(WS_PRECIO, actllave_precio!PRE_EQUIV, cant)
     new_precio!PRE_PRE2 = WS_PRECIO
    Else
     new_precio!PRE_PRE2 = 0
    End If
    '------------------------------
    If actllave_precio!PRE_POR3 <> 0 Then
     WS_PRECIO = Val(WS_COSTO * (1 + (actllave_precio!PRE_POR3 / 100)))
     WS_PRECIO = REDONDEA_PRECIO(WS_PRECIO, actllave_precio!PRE_EQUIV, cant)
     new_precio!PRE_PRE3 = WS_PRECIO
    Else
     new_precio!PRE_PRE3 = 0
    
    End If
    '------------------------------
    If actllave_precio!PRE_POR4 <> 0 Then
     WS_PRECIO = Val(WS_COSTO * (1 + (actllave_precio!PRE_POR4 / 100)))
     WS_PRECIO = REDONDEA_PRECIO(WS_PRECIO, actllave_precio!PRE_EQUIV, cant)
     new_precio!PRE_PRE4 = WS_PRECIO
    Else
     new_precio!PRE_PRE4 = 0
    End If
    '------------------------------
    If actllave_precio!PRE_POR5 <> 0 Then
     WS_PRECIO = Val(WS_COSTO * (1 + (actllave_precio!PRE_POR5 / 100)))
     WS_PRECIO = REDONDEA_PRECIO(WS_PRECIO, actllave_precio!PRE_EQUIV, cant)
     new_precio!PRE_PRE5 = WS_PRECIO
    Else
     new_precio!PRE_PRE5 = 0
    End If
    '------------------------------
    If actllave_precio!PRE_POR6 <> 0 Then
     WS_PRECIO = Val(WS_COSTO * (1 + (actllave_precio!PRE_POR6 / 100)))
     WS_PRECIO = REDONDEA_PRECIO(WS_PRECIO, actllave_precio!PRE_EQUIV, cant)
     new_precio!PRE_PRE6 = WS_PRECIO
    Else
     new_precio!PRE_PRE6 = 0
    End If
    '------------------------------
    new_precio!PRE_DIF = Format(Val(new_precio!PRE_cosTO_ant) - Val(new_precio!PRE_COSTO_NEW), "0.0000")
    new_precio.Update
  'Else
  '  res_costo_UNI = Format(w_newcosto * actllave_precio!PRE_EQUIV, "0.00")
  'End If
OTROPRECIO:
  actllave_precio.MoveNext
Loop

End Sub
Public Function REDONDEA_PRECIO(PRE_ORIG As Currency, wflag As Currency, TOTREG As Integer) As Currency
Dim newvalor As Currency
Dim wdec As String
Dim wdec2 As String
If wflag = 1 And TOTREG = 1 Then
  newvalor = Format(PRE_ORIG, "0.00")
  REDONDEA_PRECIO = Format(Val(newvalor), "0.00")
ElseIf wflag <> 1 And TOTREG <> 1 Then
  newvalor = Format(PRE_ORIG, "0.00")
  REDONDEA_PRECIO = Format(Val(newvalor), "0.00")
Else
 newvalor = Format(PRE_ORIG, "0.000")
 REDONDEA_PRECIO = Format(Val(newvalor), "0.000")
End If

End Function


Public Sub Muestra_Preicios(wporc As Currency)
Dim I As Integer
Dim VALOR As Currency
Dim rslistcambio As rdoResultset
Dim PSlistcambio As rdoQuery

Dim pre_caja As rdoResultset
Dim PSPRE_CAJA As rdoQuery

gridcabe.Cols = 24
gridcabe.Rows = 1
gridcabe.TextMatrix(0, 0) = "Cod.Int"
gridcabe.TextMatrix(0, 1) = "Descripción"
gridcabe.TextMatrix(0, 2) = "Unid."

gridcabe.TextMatrix(0, 3) = "Cos.Anter."

gridcabe.TextMatrix(0, 4) = "Cos.Calc."
gridcabe.TextMatrix(0, 5) = "Porc.1(%)"
gridcabe.TextMatrix(0, 6) = "P1.Actual"
gridcabe.TextMatrix(0, 7) = "P1.Calc."
gridcabe.TextMatrix(0, 8) = "Porc.2(%)"
gridcabe.TextMatrix(0, 9) = "P2.Actual"
gridcabe.TextMatrix(0, 10) = "P2.Calc."
gridcabe.TextMatrix(0, 11) = "Porc.3(%)"
gridcabe.TextMatrix(0, 12) = "P3.Actual"
gridcabe.TextMatrix(0, 13) = "P3.Calc."
gridcabe.TextMatrix(0, 14) = "Porc.4(%)"
gridcabe.TextMatrix(0, 15) = "P4.Actual"
gridcabe.TextMatrix(0, 16) = "P4.Calc."

gridcabe.TextMatrix(0, 17) = "Porc.5(%)"
gridcabe.TextMatrix(0, 18) = "P5.Actual"
gridcabe.TextMatrix(0, 19) = "P5.Calc."

gridcabe.TextMatrix(0, 20) = "Porc.6(%)"
gridcabe.TextMatrix(0, 21) = "P6.Actual"
gridcabe.TextMatrix(0, 22) = "P6.Calc."


gridcabe.ColWidth(23) = 0
gridcabe.ColWidth(0) = 0
gridcabe.ColWidth(1) = 1900
gridcabe.ColWidth(2) = 700
gridcabe.ColWidth(3) = 800
gridcabe.ColWidth(4) = 800
gridcabe.ColWidth(5) = 800
gridcabe.ColWidth(6) = 800
gridcabe.Row = gridcabe.Rows - 1
gridcabe.COL = 6
gridcabe.CellForeColor = vbRed

gridcabe.ColWidth(7) = 950
gridcabe.ColWidth(8) = 800
gridcabe.COL = 9
gridcabe.CellForeColor = vbRed
gridcabe.ColWidth(9) = 800
gridcabe.ColWidth(10) = 950
gridcabe.ColWidth(11) = 800
gridcabe.COL = 12
gridcabe.CellForeColor = vbRed
gridcabe.ColWidth(12) = 800
gridcabe.COL = 15
gridcabe.CellForeColor = vbRed
gridcabe.ColWidth(13) = 950
gridcabe.ColWidth(14) = 800
gridcabe.COL = 18
gridcabe.CellForeColor = vbRed
gridcabe.ColWidth(15) = 800
gridcabe.ColWidth(16) = 950
gridcabe.COL = 21
gridcabe.CellForeColor = vbRed
gridcabe.ColWidth(17) = 800
gridcabe.ColWidth(18) = 800
gridcabe.ColWidth(19) = 950
gridcabe.ColWidth(20) = 800
gridcabe.ColWidth(21) = 800
gridcabe.ColWidth(22) = 950
gridcabe.Row = 0
For I = 0 To 23
  gridcabe.COL = I
  gridcabe.CellFontBold = True
Next I


pub_cadena = "SELECT * FROM PRECIOS, ARTI WHERE (PRE_CODART = ART_KEY ) AND (PRE_CODCIA = ART_CODCIA) AND PRE_CODCIA = ? AND PRE_CODART = ? AND PRE_SECUENCIA = ? "
Set PSPRE_CAJA = CN.CreateQuery("", pub_cadena)
PSPRE_CAJA(0) = 0
PSPRE_CAJA(1) = 0
PSPRE_CAJA(2) = 0
Set pre_caja = PSPRE_CAJA.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM NEWPRECIOS WHERE PRE_CODCIA = ? ORDER BY PRE_CODART,PRE_DIF "
Set PSACTLLAVE = CN.CreateQuery("", pub_cadena)
PSACTLLAVE(0) = LK_CODCIA
Set actllave_precio = PSACTLLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
actllave_precio.Requery
gridcabe.Visible = False
DoEvents
Do Until actllave_precio.EOF

PSPRE_CAJA(0) = LK_CODCIA
PSPRE_CAJA(1) = actllave_precio!PRE_CODART
PSPRE_CAJA(2) = actllave_precio!PRE_SECUENCIA
pre_caja.Requery
If pre_caja.EOF Then
  MsgBox "AVISAR CODIGO NO EXISTE : " & pre_caja!PRE_CODART
  GoTo pasa
End If
gridcabe.Rows = gridcabe.Rows + 1
gridcabe.RowHeight(gridcabe.Rows - 1) = 300
If actllave_precio!pre_FLAG_UNIDAD = "A" Then
gridcabe.Row = gridcabe.Rows - 1
  gridcabe.RowHeight(gridcabe.Rows - 1) = 300
  For I = 0 To 23
  gridcabe.COL = I
  gridcabe.CellBackColor = QBColor(7)
  Next I
Else
gridcabe.Row = gridcabe.Rows - 1
  gridcabe.RowHeight(gridcabe.Rows - 1) = 300
  For I = 5 To 23
  gridcabe.COL = I
  gridcabe.CellBackColor = QBColor(15)
  Next I

End If

gridcabe.TextMatrix(gridcabe.Rows - 1, 0) = pre_caja!PRE_CODART
gridcabe.TextMatrix(gridcabe.Rows - 1, 1) = pre_caja!art_nombre
gridcabe.TextMatrix(gridcabe.Rows - 1, 2) = Trim(pre_caja!PRE_UNIDAD)

gridcabe.TextMatrix(gridcabe.Rows - 1, 3) = actllave_precio!PRE_cosTO_ant  ' ANTERIOR
gridcabe.TextMatrix(gridcabe.Rows - 1, 4) = actllave_precio!PRE_COSTO_NEW ' ACTUL
gridcabe.COL = 5
gridcabe.Row = gridcabe.Rows - 1
gridcabe.CellForeColor = QBColor(3)
gridcabe.CellFontBold = True

gridcabe.TextMatrix(gridcabe.Rows - 1, 5) = pre_caja!PRE_POR1
gridcabe.TextMatrix(gridcabe.Rows - 1, 6) = pre_caja!PRE_PRE1
gridcabe.TextMatrix(gridcabe.Rows - 1, 7) = actllave_precio!PRE_PRE1

gridcabe.COL = 8
gridcabe.Row = gridcabe.Rows - 1
gridcabe.CellForeColor = QBColor(3)
gridcabe.CellFontBold = True

gridcabe.TextMatrix(gridcabe.Rows - 1, 8) = pre_caja!PRE_POR2
gridcabe.TextMatrix(gridcabe.Rows - 1, 9) = pre_caja!PRE_PRE2
gridcabe.TextMatrix(gridcabe.Rows - 1, 10) = actllave_precio!PRE_PRE2

gridcabe.COL = 11
gridcabe.Row = gridcabe.Rows - 1
gridcabe.CellForeColor = QBColor(3)
gridcabe.CellFontBold = True

gridcabe.TextMatrix(gridcabe.Rows - 1, 11) = pre_caja!PRE_POR3
gridcabe.TextMatrix(gridcabe.Rows - 1, 12) = pre_caja!PRE_PRE3
gridcabe.TextMatrix(gridcabe.Rows - 1, 13) = actllave_precio!PRE_PRE3


gridcabe.COL = 14
gridcabe.Row = gridcabe.Rows - 1
gridcabe.CellForeColor = QBColor(3)
gridcabe.CellFontBold = True
gridcabe.TextMatrix(gridcabe.Rows - 1, 14) = pre_caja!PRE_POR4
gridcabe.TextMatrix(gridcabe.Rows - 1, 15) = pre_caja!PRE_PRE4
gridcabe.TextMatrix(gridcabe.Rows - 1, 16) = actllave_precio!PRE_PRE4

gridcabe.COL = 17
gridcabe.Row = gridcabe.Rows - 1
gridcabe.CellForeColor = QBColor(3)
gridcabe.CellFontBold = True

gridcabe.TextMatrix(gridcabe.Rows - 1, 17) = pre_caja!PRE_POR5
gridcabe.TextMatrix(gridcabe.Rows - 1, 18) = pre_caja!PRE_PRE5
gridcabe.TextMatrix(gridcabe.Rows - 1, 19) = actllave_precio!PRE_PRE5

gridcabe.COL = 20
gridcabe.Row = gridcabe.Rows - 1
gridcabe.CellForeColor = QBColor(3)
gridcabe.CellFontBold = True

gridcabe.TextMatrix(gridcabe.Rows - 1, 20) = pre_caja!PRE_POR6
gridcabe.TextMatrix(gridcabe.Rows - 1, 21) = pre_caja!PRE_PRE6
gridcabe.TextMatrix(gridcabe.Rows - 1, 22) = actllave_precio!PRE_PRE6
' secuencia
gridcabe.TextMatrix(gridcabe.Rows - 1, 23) = actllave_precio!PRE_SECUENCIA


pasa:
actllave_precio.MoveNext
Loop
gridcabe.Visible = True
DoEvents
fracambio.Visible = True


End Sub

Private Sub txtnumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtnumero, txtnumero

End Sub

Public Sub ACT_PRECIO_POR(WCODIGO As Currency, Wsec As Currency, OTRA_CIA As String)


Exit Sub

End Sub

Public Sub Calcula_NC()
Dim I As Integer
Dim ws_dif As Currency
Dim ws_des As Currency
Dim ws_doc As Currency
ws_dif = 0
ws_des = 0
ws_doc = 0
For I = 1 To gridcabe.Rows - 1
  ws_dif = ws_dif + Val(gridcabe.TextMatrix(I, 12))
  ws_des = ws_des + Val(gridcabe.TextMatrix(I, 11))
  ws_doc = ws_doc + Val(gridcabe.TextMatrix(I, 9))
  If Val(gridcabe.TextMatrix(I, 12)) <> 0 Then
    If opnc(0).Value Or opnc(3).Value Or opnc(4).Value Then
      gridcabe.TextMatrix(I, 18) = "UPDATE FACART SET FAR_NUMFAC_C = " & PUB_NUMFAC & " , FAR_NUMSER_C = " & PUB_NUMSER & " , FAR_GASTOS = " & Val(gridcabe.TextMatrix(I, 12)) & " WHERE FAR_CODCIA = '" & LK_CODCIA & "' AND FAR_FECHA = '" & Format(gridcabe.TextMatrix(I, 13), "dd/mm/yyyy") & "' AND FAR_FBG = '" & gridcabe.TextMatrix(I, 14) & "' AND FAR_NUMSER = '" & gridcabe.TextMatrix(I, 15) & "' AND FAR_NUMFAC = " & gridcabe.TextMatrix(I, 16) & " AND FAR_NUMSEC = " & gridcabe.TextMatrix(I, 17) & " AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E'"
    ElseIf opnc(1).Value Then
      gridcabe.TextMatrix(I, 18) = "UPDATE FACART SET FAR_NUMFAC_C = " & PUB_NUMFAC & " , FAR_NUMSER_C = " & PUB_NUMSER & " , FAR_FLETE = " & Val(gridcabe.TextMatrix(I, 12)) & " WHERE FAR_CODCIA = '" & LK_CODCIA & "' AND FAR_FECHA = '" & Format(gridcabe.TextMatrix(I, 13), "dd/mm/yyyy") & "' AND FAR_FBG = '" & gridcabe.TextMatrix(I, 14) & "' AND FAR_NUMSER = '" & gridcabe.TextMatrix(I, 15) & "' AND FAR_NUMFAC = " & gridcabe.TextMatrix(I, 16) & " AND FAR_NUMSEC = " & gridcabe.TextMatrix(I, 17) & " AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E'"
    ElseIf opnc(2).Value Then
      gridcabe.TextMatrix(I, 18) = "UPDATE FACART SET FAR_NUMFAC_C = " & PUB_NUMFAC & " , FAR_NUMSER_C = " & PUB_NUMSER & " , FAR_TOT_FLETE = " & Val(gridcabe.TextMatrix(I, 12)) & " WHERE FAR_CODCIA = '" & LK_CODCIA & "' AND FAR_FECHA = '" & Format(gridcabe.TextMatrix(I, 13), "dd/mm/yyyy") & "' AND FAR_FBG = '" & gridcabe.TextMatrix(I, 14) & "' AND FAR_NUMSER = '" & gridcabe.TextMatrix(I, 15) & "' AND FAR_NUMFAC = " & gridcabe.TextMatrix(I, 16) & " AND FAR_NUMSEC = " & gridcabe.TextMatrix(I, 17) & " AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E'"
    End If
  Else
    gridcabe.TextMatrix(I, 18) = " "
  End If
Next I
lbltotnc.Caption = Format(ws_dif, "0.00")
lbltotdes.Caption = Format(ws_des, "0.00")
lbltotdoc.Caption = Format(ws_doc, "0.00")


End Sub

Public Sub LLENADOS(cont As ListBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
'    cont.AddItem " "
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub


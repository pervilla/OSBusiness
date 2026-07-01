VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FORM_COTV 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Cotizacoines"
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
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4890
   ScaleWidth      =   6600
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin VB.Frame frault 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      TabIndex        =   64
      Top             =   5880
      Visible         =   0   'False
      Width           =   9375
      Begin VB.Label lcodart 
         BackStyle       =   0  'Transparent
         Caption         =   "Margenes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   71
         Tag             =   "9999"
         Top             =   0
         Width           =   975
      End
      Begin VB.Label lvcosto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2040
         TabIndex        =   70
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "V. Costo:"
         Height          =   195
         Index           =   4
         Left            =   1080
         TabIndex        =   69
         Tag             =   "9999"
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Marg. %"
         Height          =   195
         Index           =   5
         Left            =   5280
         TabIndex        =   68
         Tag             =   "9999"
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lmargen 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6120
         TabIndex        =   67
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lutilidad 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4200
         TabIndex        =   66
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Utilidad :"
         Height          =   195
         Index           =   6
         Left            =   3240
         TabIndex        =   65
         Tag             =   "9999"
         Top             =   120
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   285
      Left            =   10440
      TabIndex        =   63
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   128
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
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
      Left            =   4200
      TabIndex        =   38
      Top             =   2040
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   42
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
         TabIndex        =   40
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
         Left            =   120
         TabIndex        =   43
         Top             =   840
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
         TabIndex        =   41
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
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   1335
      End
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
      Left            =   8400
      TabIndex        =   34
      Top             =   240
      Width           =   495
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
      Left            =   8880
      TabIndex        =   33
      Top             =   240
      Width           =   975
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
      Height          =   375
      Left            =   8640
      TabIndex        =   32
      Top             =   600
      Width           =   735
   End
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
      Height          =   375
      Left            =   9480
      TabIndex        =   31
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdimp 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   30
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdconsulta 
      Caption         =   "Co&nsultas"
      Height          =   375
      Left            =   2400
      TabIndex        =   29
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdIngreso 
      Caption         =   "&Ingreso"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame f1 
      Caption         =   "Datos del Solicitante :"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton cmdact 
         BackColor       =   &H00FFFF80&
         Caption         =   "Update"
         Height          =   315
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton c_condi 
         Caption         =   "&Condiciones"
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
         Left            =   6000
         TabIndex        =   46
         Top             =   120
         Width           =   1455
      End
      Begin VB.Frame fraacceso 
         Height          =   975
         Left            =   6000
         TabIndex        =   76
         Top             =   480
         Width           =   2415
         Begin VB.CheckBox chelima 
            Caption         =   "Cotizacion Especial."
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
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chemarg 
            Caption         =   "Cotización con Margenes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   435
            Left            =   120
            TabIndex        =   77
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.TextBox txtvend 
         Height          =   285
         Left            =   4200
         TabIndex        =   74
         Top             =   840
         Width           =   495
      End
      Begin VB.CheckBox Che_Pedidos 
         Caption         =   "Opciones Pre-Pedido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   8640
         TabIndex        =   72
         Top             =   1080
         Width           =   1455
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
         Height          =   285
         Left            =   720
         TabIndex        =   27
         Top             =   1200
         Width           =   5175
      End
      Begin VB.ComboBox moneda 
         Height          =   315
         ItemData        =   "FORM_COTV.frx":0000
         Left            =   3000
         List            =   "FORM_COTV.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtruc 
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
         Left            =   1560
         TabIndex        =   20
         Top             =   840
         Width           =   1335
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
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblvend 
         Caption         =   "."
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
         Left            =   4800
         TabIndex        =   75
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label LBLHORA 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   73
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lcodart 
         Caption         =   "Numero:"
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
         Index           =   7
         Left            =   7560
         TabIndex        =   37
         Tag             =   "9999"
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lcodart 
         Caption         =   "Atte."
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
         Index           =   5
         Left            =   240
         TabIndex        =   28
         Tag             =   "9999"
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label lblcli 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lcodart 
         Caption         =   "Moneda : "
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   24
         Tag             =   "9999"
         Top             =   720
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lcodart 
         Caption         =   "R.U.C. /D.N.I.:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Tag             =   "9999"
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label lcodart 
         Caption         =   "Nombre / Razon Social:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Tag             =   "9999"
         Top             =   240
         Width           =   2925
      End
   End
   Begin VB.CommandButton cancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   6720
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Tag             =   "0"
      Top             =   3120
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
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Tag             =   "100"
      Top             =   1560
      Width           =   10575
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   52
         Top             =   3360
         Width           =   9255
         Begin VB.Label lblvcto 
            BorderStyle     =   1  'Fixed Single
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
            Height          =   255
            Left            =   1320
            TabIndex        =   85
            Top             =   675
            Width           =   7335
         End
         Begin VB.Label Label8 
            Caption         =   "Lote Vcto:"
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
            TabIndex        =   84
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "PRE6 :"
            Height          =   255
            Left            =   6120
            TabIndex        =   62
            Top             =   390
            Width           =   615
         End
         Begin VB.Label TXTPRE6 
            BackColor       =   &H00E0E0E0&
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
            Height          =   285
            Left            =   6840
            TabIndex        =   61
            Top             =   390
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Costo Promedio:"
            Height          =   255
            Left            =   3360
            TabIndex        =   60
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Stock de Almacenes :"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   390
            Width           =   1935
         End
         Begin VB.Label lblstockg 
            BackColor       =   &H00E0E0E0&
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
            Height          =   285
            Left            =   2160
            TabIndex        =   58
            Top             =   390
            Width           =   975
         End
         Begin VB.Label lblcosprog 
            BackColor       =   &H00E0E0E0&
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
            Height          =   285
            Left            =   4920
            TabIndex        =   57
            Top             =   390
            Width           =   975
         End
         Begin VB.Label lblpro 
            Caption         =   "Producto :"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   975
         End
         Begin VB.Label stock 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF00&
            Height          =   255
            Left            =   7560
            TabIndex        =   55
            Top             =   120
            Width           =   1095
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
            Left            =   6480
            TabIndex        =   54
            Top             =   120
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
            Left            =   1200
            TabIndex        =   53
            Top             =   120
            Width           =   5175
         End
      End
      Begin VB.TextBox txtmayor 
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
         Left            =   9480
         TabIndex        =   50
         Text            =   "500"
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Otras Opciones."
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
         Left            =   8520
         TabIndex        =   49
         Top             =   3720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtmax 
         Height          =   285
         Left            =   9600
         TabIndex        =   47
         Text            =   "30"
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox PRECIOS 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox UNIDAD 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txttotal 
         Enabled         =   0   'False
         Height          =   255
         Left            =   6960
         TabIndex        =   14
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtigv 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4080
         TabIndex        =   13
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtvalorv 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1680
         TabIndex        =   12
         Top             =   4800
         Width           =   975
      End
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   2760
         TabIndex        =   6
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
         TextRTF         =   $"FORM_COTV.frx":0020
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   3015
         Left            =   120
         TabIndex        =   0
         Tag             =   "9999"
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         FocusRect       =   2
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
      End
      Begin VB.Label Label9 
         Caption         =   "Deuda :"
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
         Left            =   3960
         TabIndex        =   82
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lbldeuda 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4680
         TabIndex        =   81
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Credito Disp.:"
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
         Left            =   6360
         TabIndex        =   80
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblcredito 
         Alignment       =   1  'Right Justify
         Caption         =   "."
         Height          =   255
         Left            =   8040
         TabIndex        =   79
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblmayor 
         Caption         =   "Mayores que:"
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
         Left            =   8280
         TabIndex        =   51
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Max.Item"
         Height          =   255
         Left            =   9600
         TabIndex        =   48
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label i_moneda 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         Height          =   195
         Left            =   6600
         TabIndex        =   25
         Top             =   4800
         Width           =   285
      End
      Begin VB.Label Label3 
         Caption         =   "Total :"
         Height          =   195
         Index           =   3
         Left            =   5880
         TabIndex        =   17
         Tag             =   "9999"
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "I.G.V. :"
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   16
         Tag             =   "9999"
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Valor Venta:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   15
         Tag             =   "9999"
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Lista de Articulos :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Tag             =   "9999"
         Top             =   120
         Width           =   3735
      End
      Begin VB.Label momen 
         Caption         =   "Un Momento ..."
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   360
      Top             =   7320
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
      TabIndex        =   36
      Tag             =   "9999"
      Top             =   240
      Width           =   525
   End
   Begin VB.Label lcodart 
      Caption         =   "Nº. Doc"
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   35
      Tag             =   "9999"
      Top             =   240
      Width           =   1125
   End
End
Attribute VB_Name = "FORM_COTV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_flag_bloq As String
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
Dim st_codcia1 As String
Dim st_codcia2 As String
Dim st_codcia3 As String
Dim st_codcia4 As String
Dim st_codcia5 As String
Dim st_codcia6 As String
Dim st_codcia7 As String

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
cmdImp.Visible = False
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
cmdIngreso.Enabled = True
'grid_fac.SetFocus

End Sub

Private Sub Che_Pedidos_Click()
If Che_Pedidos.Value = 1 Then
txtmax.Text = 11
Else
txtmax.Text = 30
End If
End Sub

Private Sub chemarg_Click()
If chemarg.Value = 1 Then
frault.Visible = True
Else
frault.Visible = False
End If
End Sub

Private Sub cmdact_Click()
Dim wnumfac  As Currency
Dim wser As Integer

If WMODO <> "I" Then
  MsgBox "Solo para Nuevos Documentos.", 48, Pub_Titulo
  Exit Sub
End If
Screen.MousePointer = 11
cmdact.Enabled = False
WMODO = "I"
PSTEMP_MAYOR(0) = LK_CODCIA
temp_mayor.Requery
If temp_mayor.EOF Then
 wser = 301
 wnumfac = 1
Else
 wser = Nulo_Valors(temp_mayor!PED_NUMSER)
 wnumfac = Val(Nulo_Valor0(temp_mayor!PED_NUMFAC)) + 1
End If
tserie.Text = wser
txtdoc.Text = wnumfac
Screen.MousePointer = 0
cmdact.Enabled = True
End Sub

Private Sub cmdconsulta_Click()
cmdImp.Visible = True
cmdIngreso.Enabled = False
tserie.Locked = False
txtdoc.Locked = False
tserie.Enabled = True
txtdoc.Enabled = True
'siguiente.Enabled = True
'anterior.Enabled = True
f1.Enabled = True
tserie.Text = "301"
tserie.Locked = True
Azul txtdoc, txtdoc

End Sub

Private Sub cmdImp_Click()
Call REP_CONSUL
End Sub

Private Sub cmdIngreso_Click()
Dim vtacospro As Currency
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
If Trim(txtcli.Text) = "" Then
  MsgBox "Nombre del Cliente ", 48, Pub_Titulo
  txtcli.SetFocus
  Exit Sub
End If
If Val(txttotal.Text) <= 0 Then
  MsgBox "Ingrese Datos ", 48, Pub_Titulo
  grid_fac.SetFocus
  Exit Sub
End If

If grid_fac.Rows = 3 Then
 If grid_fac.TextMatrix(2, 0) = "" Then
   MsgBox "Ingrese Datos de Productos ", 48, Pub_Titulo
   grid_fac.SetFocus
   Exit Sub
 End If
End If

Barra.Visible = False
suma_grid
For fila = 2 To grid_fac.Rows - 1
If grid_fac.TextMatrix(fila, 1) <> "" Then
  If Val(grid_fac.TextMatrix(fila, 2)) <= 0 Then
    MsgBox "Verificar, cantidad en cero o menor. - " & grid_fac.TextMatrix(fila, 1) & " : " & grid_fac.TextMatrix(fila, 0), 48, Pub_Titulo
    grid_fac.SetFocus
    GoTo fin
  End If
  If Val(grid_fac.TextMatrix(fila, 4)) = 0 Then
    MsgBox "Verificar hay algun precio en 0 .", 48, Pub_Titulo
    grid_fac.SetFocus
    GoTo fin
  End If
  If Left(moneda.Text, 1) = "S" Then
    vtacospro = Val(grid_fac.TextMatrix(fila, 18))
    vtacospro = Format(vtacospro, "0.00")
    pub_cadena = "S/. " & Format(vtacospro, "0.00") & " (Inc. IGV)"
  Else
    SQ_OPER = 1
    PUB_CAL_INI = LK_FECHA_DIA
    PUB_CAL_FIN = LK_FECHA_DIA
    LEER_CAL_LLAVE
    If cal_llave.EOF Then
      MsgBox "No Existe Calendario Verificar", 48, Pub_Titulo
      vtacospro = 0
    Else
      If Nulo_Valor0(cal_llave!cal_tipo_cambio) Then
        vtacospro = Format(Val(grid_fac.TextMatrix(fila, 18)) / Val(cal_llave!cal_tipo_cambio), "0.00")
      Else
        vtacospro = 0
      End If
    End If
    vtacospro = Format(vtacospro, "0.00")
    pub_cadena = "US$. " & Format(vtacospro, "0.00") & " (Inc. IGV)"
  End If
  If Val(grid_fac.TextMatrix(fila, 4)) < vtacospro Then
      MsgBox Trim(grid_fac.TextMatrix(fila, 0)) & Chr(13) & "Costo Promedio es Mayor que el Precio de Venta.  " & Chr(13) & "Cos.Prom.: " & pub_cadena, 48, Pub_Titulo
      If LK_CODUSU = "SUPERVISOR" Or LK_CODUSU = "ADMIN" Then
         MsgBox "Usuario Autorizado a registrar la Venta", 48, Pub_Titulo
      Else
         grid_fac.SetFocus
         GoTo fin
      End If
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
GoSub ACT1
Call REP_CONSUL
fila = 1
SUM_D = 0
SUM_H = 0
CABE_MAN
LIMPIA_DATOS
fila = 0
'cancelar.SetFocus
CABE_MAN
Barra.Visible = False
cmdIngreso.Caption = "&Ingreso"

GoTo fin

ACT1:
If Trim(txtdoc.Text) <> 0 Then
 pub_cadena = "DELETE PEDIDOS WHERE PED_CODCIA = '" & LK_CODCIA & "' AND PED_NUMSER = " & Trim(tserie.Text) & " AND PED_NUMFAC = " & Trim(txtdoc.Text)
 CN.Execute pub_cadena, rdExecDirect
End If
fila = 1
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
   temp_llave!PED_ESTADO = "N"
   temp_llave!PED_CODUSU = LK_CODUSU
   temp_llave!PED_codart = Val(grid_fac.TextMatrix(fila, 10))
   temp_llave!PED_UNIDAD = Trim(grid_fac.TextMatrix(fila, 3))
   temp_llave!PED_EQUIV = Val(grid_fac.TextMatrix(fila, 12))
   temp_llave!PED_NOMCLIE = Trim(FORM_COTV.lblcli.Caption)
   temp_llave!PED_RUCCLIE = Trim(txtruc.Text) ' Trim(fbg.Text)
   temp_llave!PED_CODCLIE = Val(txtcli.Text)
   temp_llave!PED_TIPMOV = 301
   temp_llave!PED_HORA = Format(Now, "hh:mm:ss AMPM")
   temp_llave!PED_descto = Val(grid_fac.TextMatrix(fila, 5))
   temp_llave!PED_MONEDA = Left(Trim(moneda.Text), 1)
   temp_llave!PED_CONTACTO = txtatte.Text
   temp_llave!PED_FORMA = Trim(forma.Text)
   temp_llave!PED_TIEMPO = Trim(tiempo.Text)
   temp_llave!PED_OFERTA = Trim(oferta.Text)
   temp_llave!PED_SUBTOTAL = Val(grid_fac.TextMatrix(fila, 6))
   temp_llave!PED_FLAG_PRE = Val(Che_Pedidos.Value)
   temp_llave!ped_codven = Val(txtvend.Text)
   temp_llave.Update
pasa:
   fila = fila + 1
   WS_NRO_MOV = WS_NRO_MOV + 1
   If fila >= FORM_COTV.grid_fac.Rows Then
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
CABE_MAN
WMODO = "I"
PSTEMP_MAYOR(0) = LK_CODCIA
temp_mayor.Requery
If temp_mayor.EOF Then
 wser = 301
 wnumfac = 1
Else
 wser = Nulo_Valors(temp_mayor!PED_NUMSER)
 wnumfac = Val(Nulo_Valor0(temp_mayor!PED_NUMFAC)) + 1
End If
tserie.Text = wser
txtdoc.Text = wnumfac

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
txtcli.SetFocus
moneda.ListIndex = 0
Exit Sub

error_fatal:
    pub_mensaje = "Se ha producido un error " & "al abrir la conexión:" & Err & " - " & Error & vbCr
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
  If textovar.Visible Then Azul3 textovar, textovar
  FORM_COTV.Barra.Visible = False
  Screen.MousePointer = 0
  grid_fac.SetFocus
Else
  MsgBox Err.Description, 48, Pub_Titulo
End If

End Sub





Private Sub Command1_Click()
condi.Visible = False
c_condi.SetFocus
End Sub

Private Sub Form_Load()
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
loc_flag_bloq = ""
For fila = 1 To lk_OTROS_Count
       If Val(lk_OTROS(fila)) = 50 Then loc_flag_bloq = "A"
Next fila

fraacceso.Enabled = True
Che_Pedidos.Value = 0
If loc_flag_bloq = "A" Then
  fraacceso.Enabled = False
  Che_Pedidos.Value = 1
End If

pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_NUMSER = ? and PED_NUMFAC = ?  ORDER BY PED_NUMSEC"
Set PSLOC_WARTI = CN.CreateQuery("", pub_cadena)
PSLOC_WARTI(0) = 0
PSLOC_WARTI(1) = 0
PSLOC_WARTI(2) = 0
PSLOC_WARTI(3) = 0
Set llave_sum_arti = PSLOC_WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = 301  ORDER BY  PED_NUMFAC DESC "
Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
PSTEMP_MAYOR(0) = LK_CODCIA
PSTEMP_MAYOR.MaxRows = 1
Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PEDIDOS WHERE  PED_CODCIA = '00' AND PED_TIPMOV = 301 ORDER BY PED_CODCIA"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
temp_llave.Requery

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
cmdImp.Visible = False
txtruc.MaxLength = LK_DIG_RUC


st_codcia1 = LK_CODCIA
StockC1 = "Actual"
st_codcia2 = LK_CODCIA
StockC2 = "Actual"
st_codcia3 = LK_CODCIA
StockC3 = "Actual"
st_codcia4 = LK_CODCIA
StockC4 = "Actual"
st_codcia5 = LK_CODCIA
st_codcia6 = LK_CODCIA
st_codcia7 = LK_CODCIA
StockC5 = "Actual"
SQ_OPER = 1
PUB_TIPREG = 58
PUB_NUMTAB = 1
LEER_TAB_LLAVE
If Not tab_llave.EOF Then
   st_codcia1 = Mid(tab_llave!tab_nomlargo, 1, 2)
   pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = '" & st_codcia1 & "' AND TAB_TIPREG = 58 AND TAB_NUMTAB = 1"
   Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   PSPAR_MULTI(0) = st_codcia1
   par_multi.Requery
   If Not par_multi.EOF Then
   'lbltitulostock.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
'   lpstock.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
'   lpstocku.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
   End If
   StockC1 = Trim(X!tab_nomcorto)
   
   st_codcia2 = Mid(tab_llave!tab_nomlargo, 3, 2)
   pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = '" & st_codcia2 & "' AND TAB_TIPREG = 58 AND TAB_NUMTAB = 1"
   Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   PSPAR_MULTI(0) = st_codcia2
   par_multi.Requery
   If Not par_multi.EOF Then
 '  lbltitulostock2.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
 '  lpstock2.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
 '  lpstocku2.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
   End If
   StockC2 = Trim(X!tab_nomcorto)
   
   st_codcia3 = Mid(tab_llave!tab_nomlargo, 5, 2)
   pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = '" & st_codcia3 & "' AND TAB_TIPREG = 58 AND TAB_NUMTAB = 1"
   Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   PSPAR_MULTI(0) = st_codcia3
   par_multi.Requery
   If Not par_multi.EOF Then
  ' Label10.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
  ' lpstock3.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
  ' lpstocku3.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
   End If
   StockC3 = Trim(X!tab_nomcorto)
   
   st_codcia4 = Mid(tab_llave!tab_nomlargo, 7, 2)
   pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = '" & st_codcia4 & "' AND TAB_TIPREG = 58 AND TAB_NUMTAB = 1"
   Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   PSPAR_MULTI(0) = st_codcia4
   par_multi.Requery
   If Not par_multi.EOF Then
   'lbltitulostock3.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
   'lpstock4.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
   'lpstocku4.ForeColor = QBColor(Left(par_multi!PAR_COLOR, 2))
   End If
   StockC4 = Trim(X!tab_nomcorto)
   
   st_codcia5 = Mid(tab_llave!tab_nomlargo, 9, 2)
   pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = '" & st_codcia5 & "' AND TAB_TIPREG = 58 AND TAB_NUMTAB = 1"
   Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   PSPAR_MULTI(0) = st_codcia5
   par_multi.Requery
   If Not par_multi.EOF Then
   End If
   StockC5 = Trim(X!tab_nomcorto)
   
   st_codcia6 = Mid(tab_llave!tab_nomlargo, 11, 2)
   pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = '" & st_codcia6 & "' AND TAB_TIPREG = 58 AND TAB_NUMTAB = 1"
   Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   PSPAR_MULTI(0) = st_codcia6
   par_multi.Requery
   If Not par_multi.EOF Then
   End If
   StockC6 = Trim(X!tab_nomcorto)
   
   
   st_codcia7 = Mid(tab_llave!tab_nomlargo, 13, 2)
   pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = '" & st_codcia7 & "' AND TAB_TIPREG = 58 AND TAB_NUMTAB = 1"
   Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   PSPAR_MULTI(0) = st_codcia7
   par_multi.Requery
   If Not par_multi.EOF Then
   End If
   StockC7 = Trim(X!tab_nomcorto)
   
End If

'lbltitulostock.Caption = StockC1
'lbltitulostock2.Caption = StockC2
'Label10.Caption = StockC3
'lbltitulostock3.Caption = StockC4
'lbltitulostock4.Caption = StockC5




Exit Sub
sale:
MsgBox "Depurar: " & Err.Description, 48, Pub_Titulo
Resume Next
End Sub



Private Sub forma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tiempo.SetFocus
End If
End Sub

Private Sub grid_fac_EnterCell()
textovar.Visible = False
textovar.Text = Trim(grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL))
textovar.Left = grid_fac.Left + grid_fac.CellLeft
textovar.Width = grid_fac.CellWidth
textovar.Height = grid_fac.CellHeight
textovar.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 1560 '480
If grid_fac.COL = 1 Then
 If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
  textovar.Visible = True
  textovar.SetFocus
 End If
End If
If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) <> "" Then
If Val(grid_fac.TextMatrix(grid_fac.Row, 12)) = 0 Then Exit Sub
 stock.Caption = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 15)) / Val(grid_fac.TextMatrix(grid_fac.Row, 12)), "0.00")
 unid.Caption = grid_fac.TextMatrix(grid_fac.Row, 16)
 nomarti.Caption = grid_fac.TextMatrix(grid_fac.Row, 0)
 lblstockg.Caption = grid_fac.TextMatrix(grid_fac.Row, 19)
 lblcosprog.Caption = grid_fac.TextMatrix(grid_fac.Row, 20)
Else
 stock.Caption = ""
 unid.Caption = ""
 nomarti.Caption = ""
 lblstockg.Caption = ""
 lblcosprog.Caption = ""
End If

End Sub

Private Sub grid_fac_KeyPress(KeyAscii As Integer)
Dim wflag_Codven As String
Dim a As Integer
Dim t, WC
Dim wprecios As String * 12
Static CONS
Dim wactivo As Integer
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
    unidad.Left = grid_fac.Left + grid_fac.CellLeft
    unidad.Width = grid_fac.CellWidth
    unidad.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 1550 '480
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
    LEER_PRE_LLAVE
    unidad.Clear
    unidad.Visible = True
    wactivo = 0
    Do Until pre_mayor.EOF
     unidad.AddItem Trim(pre_mayor!pre_unidad) & String(30, " ") & pre_mayor!PRE_SECUENCIA
     If pre_mayor!pre_FLAG_UNIDAD = "A" Then
       wactivo = pre_mayor.AbsolutePosition - 1
     End If
     pre_mayor.MoveNext
    Loop
    On Error GoTo pasa
    unidad.ListIndex = wactivo
    grid_fac.TextMatrix(grid_fac.Row, 13) = wactivo
    On Error GoTo 0
    unidad.Visible = True
    unidad.SetFocus
    SendKeys "%{up}"
     Exit Sub
End If

PSUSU_LLAVE(0) = LK_CODUSU
usu_llave.Requery
wflag_Codven = ""
If Val(usu_llave!usu_codven) <> 0 And Val(txtvend.Text) = 0 Then
    MsgBox "Necesita Codigo de Vendedor", 48, Pub_Titulo
    Azul txtvend, txtvend
    Exit Sub
    wflag_Codven = "A"
End If
    
If grid_fac.COL = 4 Then
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
      grid_fac.SetFocus
      Exit Sub
    End If
    
    PRECIOS.Left = grid_fac.Left + grid_fac.CellLeft
    PRECIOS.Width = grid_fac.CellWidth + 600
    PRECIOS.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 1550 '480
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
    LEER_PRE_LLAVE
    PRECIOS.Clear
    PRECIOS.Width = 1980
    PRECIOS.Visible = True
    Do Until pre_mayor.EOF
      If Val(grid_fac.TextMatrix(grid_fac.Row, 11)) <> Val(pre_mayor!PRE_SECUENCIA) Then GoTo salta_precios
      If Left(moneda.Text, 1) = "S" Then
      If Val(txtvend.Text) <> 0 Then
      If Val(ven_llave!VEM_AM_LISTA) <> 0 Then
            If Val(ven_llave!VEM_AM_LISTA) = 1 Then
             PRECIOS.AddItem Left(tab_precioss(1), 7) & "  " & Format(pre_llave!PRE_PRE1, "0.000") & String(20, " ") & "1"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 2 Then
             PRECIOS.AddItem Left(tab_precioss(2), 7) & "  " & Format(pre_llave!pre_pre2, "0.000") & String(20, " ") & "2"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 3 Then
             PRECIOS.AddItem Left(tab_precioss(3), 7) & "  " & Format(pre_llave!PRE_PRE3, "0.000") & String(20, " ") & "3"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 4 Then
             PRECIOS.AddItem Left(tab_precioss(4), 7) & "  " & Format(pre_llave!PRE_PRE4, "0.000") & String(20, " ") & "4"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 5 Then
             PRECIOS.AddItem Left(tab_precioss(5), 7) & "  " & Format(pre_llave!PRE_PRE5, "0.000") & String(20, " ") & "5"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 6 Then
             PRECIOS.AddItem "Lista.6" & "  " & Format(pre_llave!PRE_PRE6, "0.000") & String(20, " ") & "6"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 7 Then
             PRECIOS.AddItem "Lista.7" & "  " & Format(pre_llave!PRE_PRE7, "0.000") & String(20, " ") & "7"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 8 Then
             PRECIOS.AddItem "Lista.8" & "  " & Format(pre_llave!PRE_PRE8, "0.000") & String(20, " ") & "8"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 9 Then
             PRECIOS.AddItem "Lista.9" & "  " & Format(pre_llave!PRE_PRE9, "0.000") & String(20, " ") & "9"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 10 Then
             PRECIOS.AddItem "Lista.10" & "  " & Format(pre_llave!PRE_PRECHI1, "0.000") & String(20, " ") & "10"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 11 Then
             PRECIOS.AddItem "Lista.11" & "  " & Format(pre_llave!PRE_PREC1, "0.000") & String(20, " ") & "11"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 12 Then
             PRECIOS.AddItem "Lista.12" & "  " & Format(pre_llave!PRE_PREC2, "0.000") & String(20, " ") & "12"
            ElseIf Val(ven_llave!VEM_AM_LISTA) = 13 And Val(Nulo_Valor0(pre_llave!PRE_OP1)) <> 0 Then
                If art_LLAVE!ART_ORDEN = 0 Then
                 MsgBox "Producto No permitido.", 48, Pub_Titulo
                 GoTo salta_precios
                Else
                 PRECIOS.AddItem "Lista.13" & "  " & Format(pre_llave!PRE_OP1, "0.000") & String(20, " ") & "13"
                 End If
            End If
            'PRECIOS.AddItem ""
            GoTo salta_precios
     End If
     End If
          wprecios = pre_mayor!PRE_PRE1
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!pre_pre2
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(1), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PRE3
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(2), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PRE4
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(3), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PRE5
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(4), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PRE6
          If Val(wprecios) <> 0 Then PRECIOS.AddItem "LISTA6--" & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PRE7
          If Val(wprecios) <> 0 Then PRECIOS.AddItem "LISTA7--" & "= " & wprecios & String(60, " ") & "1"
           wprecios = pre_mayor!PRE_PRE8
          If Val(wprecios) <> 0 Then PRECIOS.AddItem "LISTA8--" & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PRECHI1
          If Val(wprecios) <> 0 Then PRECIOS.AddItem "LISTA10-" & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PREC1
          If Val(wprecios) <> 0 Then PRECIOS.AddItem "LISTA11--" & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PREC2
          If Val(wprecios) <> 0 Then PRECIOS.AddItem "LISTA12-" & "= " & wprecios & String(60, " ") & "1"
          'End If
       Else
          wprecios = pre_mayor!PRE_PRE11
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PRE22
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PRE33
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PRE44
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PRE55
          If Val(wprecios) <> 0 Then PRECIOS.AddItem Left(PRE_ETIQUETA(0), 8) & "= " & wprecios & String(60, " ") & "1"
          wprecios = pre_mayor!PRE_PRE66
          If Val(wprecios) <> 0 Then PRECIOS.AddItem "LISTA6--" & "= " & wprecios & String(60, " ") & "1"

       End If
salta_precios:
     pre_mayor.MoveNext
    Loop
    On Error GoTo pasa
    If PRECIOS.ListCount <= 0 Then
      PRECIOS.Visible = False
      MsgBox "Definir precios....", 48, Pub_Titulo
      grid_fac.COL = 1
      grid_fac.SetFocus
      Exit Sub
    End If
    PRECIOS.ListIndex = 0
    On Error GoTo 0
    
    PRECIOS.Visible = True
    PRECIOS.SetFocus
    SendKeys "%{up}"
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
    textovar.Visible = True
    Azul3 textovar, textovar
    textovar.SetFocus
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



Private Sub ListView1_DblClick()
' loc_key = ListView1.SelectedItem.Index
' TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
' TEXTOVAR_KeyPress 13
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
'If KeyAscii = 27 Then
' ListView1.Visible = False
' TEXTOVAR.Text = ""
' TEXTOVAR.SetFocus
' Exit Sub
'End If
'If KeyAscii <> 13 Then
' Exit Sub
'End If
'ListView1_DblClick
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
          If Val(grid_fac.TextMatrix(fila, 14)) = 2 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!pre_pre2)
          If Val(grid_fac.TextMatrix(fila, 14)) = 3 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE3)
          If Val(grid_fac.TextMatrix(fila, 14)) = 4 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE4)
          If Val(grid_fac.TextMatrix(fila, 14)) = 5 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE5)
       Else
          If Val(grid_fac.TextMatrix(fila, 14)) = 1 Then grid_fac.TextMatrix(fila, 13) = Val(pre_llave!PRE_PRE11)
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

Private Sub moneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtvend.SetFocus
End If
End Sub


Private Sub PRECIOS_GotFocus()
On Error GoTo sale
grid_fac.TextMatrix(grid_fac.Row, 13) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 14) = Val(Right(PRECIOS.Text, 3))
sale:
Exit Sub
End Sub

Private Sub PRECIOS_KeyPress(KeyAscii As Integer)
On Error Resume Next

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

grid_fac.TextMatrix(grid_fac.Row, 4) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.000")
grid_fac.TextMatrix(grid_fac.Row, 13) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 14) = Val(Right(PRECIOS.Text, 3))

PRECIOS.Visible = False
suma_grid
grid_fac.COL = 5
grid_fac_KeyPress 13


End Sub

Private Sub PRECIOS_KeyUp(KeyCode As Integer, Shift As Integer)
Dim ww As String
Dim wpre As Currency
If KeyCode = 45 Then
    If Trim(LK_CODUSU) <> "ADMIN" Then
       Exit Sub
    End If
'seIf grid_fac.COL = 6 Then
ww = InputBox("Digite Precios :", "Ingreso de Precio", "0")
wpre = Val(ww)
grid_fac.TextMatrix(grid_fac.Row, 4) = wpre
If Val(grid_fac.TextMatrix(grid_fac.Row, 13)) <> 0 Then
  grid_fac.TextMatrix(grid_fac.Row, 5) = redondea((Val(grid_fac.TextMatrix(grid_fac.Row, 13)) - wpre) * 100 / Val(Val(grid_fac.TextMatrix(grid_fac.Row, 13))))
End If
PRECIOS.Visible = False
suma_grid
'grid_fac.Visible = True
grid_fac.SetFocus


End If
End Sub

Private Sub salir_Click()
Unload FORM_COTV
End Sub


Public Sub LIMPIA_DATOS()
txtvend.Text = ""
LBLHORA.Caption = ""
grid_fac.Enabled = True
lblcli.Caption = ""
txtatte.Text = ""
f1.Enabled = True
txtcli.Text = ""
txtruc.Text = ""
tserie.Text = ""
txtdoc.Text = ""
grid_fac.Clear

txtigv.Text = ""
txtvalorv.Text = ""
txttotal.Text = ""
textovar.Visible = False
stock.Caption = ""
unid.Caption = ""
nomarti.Caption = ""
oferta.Text = ""
forma.Text = ""
tiempo.Text = ""
moneda.ListIndex = -1

End Sub

Public Sub CABE_MAN()
grid_fac.Cols = 22
grid_fac.Rows = 2
grid_fac.Clear
fila = 0
grid_fac.ColWidth(0) = 2400 ' nombre arti
grid_fac.ColWidth(1) = 900 ' codigo arti
grid_fac.ColWidth(2) = 800 ' cantidad
grid_fac.ColWidth(3) = 900 ' unidad
grid_fac.ColWidth(4) = 800 ' precio
grid_fac.ColWidth(5) = 700 ' decto.
grid_fac.ColWidth(6) = 1000 ' sub total
grid_fac.ColWidth(7) = 800 ' peso
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
grid_fac.ColWidth(18) = 0
grid_fac.ColWidth(19) = 0 ' Stock de todos los almacnes
grid_fac.ColWidth(20) = 0  ' Costo Promedio
grid_fac.ColWidth(21) = 0  ' FLAG IGV

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
grid_fac.TextMatrix(0, 18) = "Costo Promedio"

grid_fac.RowHeight(1) = 320

End Sub
Public Sub suma_grid()
Dim exvalor As Currency
Dim sumexo As Currency
'On Error GoTo SALE
Dim WF As Integer
WF = 2
Dim fx As Integer
Dim wcantid As Currency
Dim wpeso As Currency
Dim SUM_VCOSTO As Currency
fx = 1
SUM_H = 0
SUM_D = 0
SUM_VCOSTO = 0
wcantid = 0
sumexo = 0
Do While fx = 1
    'If Left(grid_fac.TextMatrix(WF, 0), 1) <> "T" Then
      SUM_D = SUM_D + Val(grid_fac.TextMatrix(WF, 4))
      SUM_H = SUM_H + Val(Val(grid_fac.TextMatrix(WF, 2)) * Val(grid_fac.TextMatrix(WF, 4)))
      If Trim(grid_fac.TextMatrix(WF, 21)) = "A" Then
        
        SUM_VCOSTO = SUM_VCOSTO + ((Val(grid_fac.TextMatrix(WF, 2))) * Val(grid_fac.TextMatrix(WF, 20)))
        exvalor = ((Val(grid_fac.TextMatrix(WF, 2))) * Val(grid_fac.TextMatrix(WF, 4)))
        sumexo = sumexo + (exvalor - (Val(exvalor) / (1 + (LK_IGV / 100))))
      Else
        SUM_VCOSTO = SUM_VCOSTO + Val(Format((Val(grid_fac.TextMatrix(WF, 2)) * Val(grid_fac.TextMatrix(WF, 20)) / (1 + (LK_IGV / 100))), "0.00"))
      End If
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
   grid_fac.TextMatrix(1, 0) = "Totales = "
   grid_fac.TextMatrix(1, 6) = Format(SUM_H, "####0.00")
   grid_fac.TextMatrix(1, 2) = Format(wcantid, "####0.00")
   grid_fac.TextMatrix(1, 7) = Format(wpeso, "####0.00")
   txttotal.Text = Format(SUM_H, "#####0.00")
   txtvalorv.Text = Format(((SUM_H / ((100 + LK_IGV) / 100))) + sumexo, "####0.00")
   txtigv.Text = Format(SUM_H - Val(txtvalorv.Text), "#####0.00")
   
   lutilidad.Caption = Format(Val(txtvalorv.Text) - SUM_VCOSTO, "#####0.00")
   lvcosto.Caption = Format(SUM_VCOSTO, "#####0.00")
   If SUM_VCOSTO <> 0 Then
   lmargen.Caption = Format(((Val(Format(lutilidad.Caption, "0.00"))) * 100) / SUM_VCOSTO, "#####0.00")
   Else
   lmargen.Caption = ""
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
If Not textovar.Visible Then Exit Sub
If grid_fac.COL = 5 Then
Exit Sub
 grid_fac.TextMatrix(grid_fac.Row, 4) = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 13)) * (100 - Val(textovar.Text)) / 100)
End If
If grid_fac.COL = 1 Then
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
    grid_fac.Text = textovar.Text
    stock.Caption = ""
    unid.Caption = ""
    nomarti.Caption = ""
    suma_grid
Else
 If grid_fac.COL = 2 Then
  grid_fac.Text = textovar.Text
 Else
  grid_fac.Text = Format(textovar.Text, "0.00")
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
End If
If grid_fac.COL <> 1 Then Exit Sub
Dim strFindMe As String
Dim itmFound As Object     ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And textovar.Text = "" Then
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
  textovar.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  textovar.SelStart = Len(textovar.Text)
  PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
  If chelima.Value = 1 Then
  TXTPRE6.Caption = Format(Val(ListView1.ListItems.Item(loc_key).SubItems(11)), "0.00")
  Else
  TXTPRE6.Caption = ""
  
  End If
  lblstockg.Caption = Jala_StockGen(Val(PUB_KEY), Val(ListView1.ListItems.Item(loc_key).SubItems(9)))
  lblvcto.Caption = ListView1.ListItems.Item(loc_key).SubItems(17)
  DoEvents
  DoEvents
fin:

End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
'SOLO_DECIMAL TEXTOVAR, KeyAscii
If KeyAscii = 27 Then
  If textovar.Text = "" Then
    textovar.Visible = False
    grid_fac.SetFocus
    Exit Sub
  End If
  textovar.Text = "" ' temporal
  'TEXTOVAR.Visible = False
  'grid_fac.SetFocus
  ListView1.Visible = False
  Exit Sub
End If
If grid_fac.COL = 5 And KeyAscii <> 13 Then
    KeyAscii = 0
    Exit Sub
End If

If grid_fac.COL = 2 Or grid_fac.COL = 4 Or grid_fac.COL = 5 Then Consistencias grid_fac, textovar, KeyAscii
If KeyAscii <> 13 Then Exit Sub
If grid_fac.COL = 1 Then
  If (grid_fac.Rows - 1) = Val(txtmax.Text) + 2 Then
    MsgBox "LLego al tipo de Filas .", 48
    textovar.SetFocus
    Exit Sub
  End If
End If

If grid_fac.COL = 2 Then
 
 textovar.Visible = False
 If Val(arm_llave!arm_stock) - Val(grid_fac.TextMatrix(grid_fac.Row, 2)) <= 0 Then
      MsgBox "Stock es :" & Format(arm_llave!arm_stock, "0.00") & "  /  Aplicando la cantidad : " & Format(Val(arm_llave!arm_stock) - Val(grid_fac.TextMatrix(grid_fac.Row, 2)), "0.00"), 48, Pub_Titulo
 End If
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
 textovar.Visible = False
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
 textovar.Visible = True
 textovar.SetFocus
 Exit Sub
End If




If grid_fac.COL <> 1 Then Exit Sub

Dim VALOR As String
Dim tf As Integer
Dim I, car
Dim itmFound As Object
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView1.Visible = False
 textovar.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
Else
 On Error GoTo mucho
 PUB_KEY = Val(textovar.Text)
 On Error GoTo 0
 If Len(textovar.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(textovar.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = textovar.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul3 textovar, textovar
       GoTo fin
    End If
    If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul3 textovar, textovar
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
    grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_unidad
    grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!arm_stock
    grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!pre_equiv
    grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!PRE_SECUENCIA
    grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE!ART_NOMBRE
    grid_fac.TextMatrix(grid_fac.Row, 10) = art_LLAVE!ART_KEY
    grid_fac.TextMatrix(grid_fac.Row, 18) = Nulo_Valor0(arm_llave!ARM_COSPRO)
    grid_fac.TextMatrix(grid_fac.Row, 21) = art_LLAVE!ART_EX_IGV
    
    grid_fac.TextMatrix(grid_fac.Row, 19) = Jala_StockGen(art_LLAVE!ART_KEY, pre_llave!pre_equiv)
    grid_fac.TextMatrix(grid_fac.Row, 20) = Format(Nulo_Valor0(arm_llave!ARM_COSPRO) * pre_llave!pre_equiv, "0.000")
    
    
    ListView1.Visible = False
    textovar.Visible = False
    grid_fac.COL = 2
    If Trim(grid_fac.Text) <> "" Then
      grid_fac.SetFocus
      Exit Sub
    End If
    textovar.Visible = True
    textovar.SetFocus
    Exit Sub
Else
  If ListView1.Visible = False And VAR_ACTIVAR <> 99 And textovar.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = textovar.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul3 textovar, textovar
       Exit Sub
     End If
     If art_llave_alt!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul3 textovar, textovar
       GoTo fin
     End If
     ListView1.Visible = False
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_CODART = WCOD_ORIGINAL
     LEER_ARM_LLAVE
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_SECUEN = 0
     LEER_PRE_LLAVE
     grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_unidad
     grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!arm_stock
     grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!pre_equiv
     grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!PRE_SECUENCIA
    
     grid_fac.TextMatrix(grid_fac.Row, 0) = art_llave_alt!ART_NOMBRE
     grid_fac.TextMatrix(grid_fac.Row, 10) = art_llave_alt!ART_KEY
     grid_fac.TextMatrix(grid_fac.Row, 18) = Nulo_Valor0(arm_llave!ARM_COSPRO)
     grid_fac.TextMatrix(grid_fac.Row, 19) = Jala_StockGen(art_LLAVE!ART_KEY, pre_llave!pre_equiv)
     grid_fac.TextMatrix(grid_fac.Row, 20) = Format(Nulo_Valor0(arm_llave!ARM_COSPRO) * pre_llave!pre_equiv, "0.000")
     grid_fac.TextMatrix(grid_fac.Row, 21) = art_LLAVE!ART_EX_IGV
     textovar.Visible = False
     ListView1.Visible = False
     grid_fac.COL = 2
     If Trim(grid_fac.Text) <> "" Then
       grid_fac.SetFocus
       Exit Sub
     End If
     textovar.Visible = True
     Azul3 textovar, textovar
     Exit Sub
  Else
    If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
    If Trim(UCase(textovar.Text)) = Left(VALOR, Len(Trim(textovar.Text))) And Len(Trim(textovar.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        textovar.Text = Trim(ListView1.ListItems.Item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       textovar.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       textovar.Text = Trim(ListView1.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       PUB_KEY = textovar.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul3 textovar, textovar
        Exit Sub
      End If
      If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul3 textovar, textovar
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
      grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_unidad
      grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!arm_stock
      grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!pre_equiv
      grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!PRE_SECUENCIA
      ListView1.Visible = False
      grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE!ART_NOMBRE
      grid_fac.TextMatrix(grid_fac.Row, 10) = art_LLAVE!ART_KEY
      grid_fac.TextMatrix(grid_fac.Row, 18) = Nulo_Valor0(arm_llave!ARM_COSPRO)
      grid_fac.TextMatrix(grid_fac.Row, 19) = Jala_StockGen(art_LLAVE!ART_KEY, pre_llave!pre_equiv)
      grid_fac.TextMatrix(grid_fac.Row, 20) = Format(Nulo_Valor0(arm_llave!ARM_COSPRO) * pre_llave!pre_equiv, "0.000")
      grid_fac.TextMatrix(grid_fac.Row, 21) = art_LLAVE!ART_EX_IGV
      grid_fac.COL = 2
      If Trim(grid_fac.Text) <> "" Then
        grid_fac.SetFocus
        Exit Sub
      End If
      textovar.Visible = True
      textovar.SetFocus
     
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
Azul3 textovar, textovar
  

Exit Sub

End Sub

Private Sub textovar_KeyUp(KeyCode As Integer, Shift As Integer)
Dim SoloconStock  As String
Dim aux_codcia As String
Dim aux_codcia2 As String
If grid_fac.COL <> 1 Then Exit Sub
' busca arti
Dim var
SoloconStock = ""
If KeyCode = 13 Then Exit Sub
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  If Len(textovar.Text) = 0 Or Trim(textovar.Text) = "" Then
    ListView1.Visible = False
    Exit Sub
  End If
  If textovar.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf textovar.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(textovar.Text, 1) = "*" Then
   textovar.Text = Mid(textovar.Text, 2, Len(textovar.Text))
   textovar.SelStart = Len(textovar.Text)
  End If
Else
 If Len(textovar.Text) = 0 Or IsNumeric(textovar.Text) = True Then
   ListView1.Visible = False
   Exit Sub
 End If
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(textovar.Text) = 1 Then
    var = Asc(textovar.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      'archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK FROM ARTI, ARTICULO WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND  ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_ALTERNO BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK ,PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_KEY  <> 1 and ART_CODCIA = '" & LK_CODCIA & "' AND ART_ALTERNO BETWEEN '" & textovar.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 7
      ' ANTERIOR      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK ,PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND ART_SITUACION <> 1 AND ARM_STOCK <> 0 AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_KEY  <> 1 and ART_CODCIA = '" & LK_CODCIA & "' AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
        PUB_CODCIA = "02"
        aux_codcia = "04" ' TEMPORAL EN 05  DEBE ESTAR EN 04
        aux_codcia2 = "08"
        
       If chelima.Value = 1 Then
        '   archi = "SELECT TOP 2000  ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, cia_actual.ARM_STOCK as ARM_STOCK , PRE_EQUIV,  PRE_UNIDAD,PRE_PRE1,PRE_PRE2,PRE_PRE3,PRE_PRE4,PRE_PRE5,PRE_PRE6,cia_otra.ARM_STOCK as ARM_STOCK2, cia_SUBALM.ARM_STOCK as ARM_STOCK3, cia_JOLMEDO.ARM_STOCK as ARM_STOCK4 ,ALM_CIA5.ARM_STOCK as ARM_STOCK5, art_marca, ART_NUMERO, ART_ESTADO FROM ARTI, ARTICULO as cia_actual,ARTICULO as cia_otra, ARTICULO as cia_SUBALM, ARTICULO as cia_JOLMEDO, ARTICULO as ALM_CIA5, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD = 'A') AND (ART_KEY = cia_actual.ARM_CODART) AND (ART_KEY = cia_otra.ARM_CODART) AND (ART_KEY = cia_SUBALM.ARM_CODART)  AND ART_ORDEN = 1 AND " & _
        '"  (ART_KEY = cia_JOLMEDO.ARM_CODART) AND (ART_KEY = ALM_CIA5.ARM_CODART)  AND ART_SITUACION <> 1 and ART_KEY <> 0 AND ART_CALIDAD = 1  AND ART_CODCIA = '" & LK_CODCIA & "' AND (cia_actual.arm_codcia = '" & st_codcia1 & "') AND (cia_otra.arm_codcia = '" & st_codcia2 & "') AND (cia_SUBALM.arm_codcia = '" & st_codcia3 & "') AND (cia_JOLMEDO.arm_codcia = '" & st_codcia4 & "')  AND (ALM_CIA5.arm_codcia = '" & st_codcia5 & "')  AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & var & "' " & SoloconStock & " ORDER BY ART_NOMBRE"
         archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, cia_actual.ARM_STOCK as ARM_STOCK , PRE_EQUIV,  PRE_UNIDAD,PRE_PRE1,PRE_PRE2,PRE_PRE3,PRE_PRE4,PRE_PRE5,PRE_PRE6,cia_otra.ARM_STOCK as ARM_STOCK2, cia_SUBALM.ARM_STOCK as ARM_STOCK3, cia_JOLMEDO.ARM_STOCK as ARM_STOCK4 ,ALM_CIA5.ARM_STOCK as ARM_STOCK5,ALM_CIA6.ARM_STOCK as ARM_STOCK6 ,ALM_CIA7.ARM_STOCK as ARM_STOCK7, art_marca, ART_NUMERO, ART_ESTADO , PRE_CANT , PRE_OP1 FROM ARTI, ARTICULO as cia_actual,ARTICULO as cia_otra, ARTICULO as cia_SUBALM, ARTICULO as cia_JOLMEDO, ARTICULO as ALM_CIA5, ARTICULO as ALM_CIA6, ARTICULO as ALM_CIA7, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD = 'A') AND (ART_KEY = cia_actual.ARM_CODART) AND (ART_KEY = cia_otra.ARM_CODART) AND (ART_KEY = cia_SUBALM.ARM_CODART)  AND ART_ORDEN = 1  AND " & _
        "  (ART_KEY = cia_JOLMEDO.ARM_CODART) AND (ART_KEY = ALM_CIA5.ARM_CODART)  AND (ART_KEY = ALM_CIA6.ARM_CODART) AND (ART_KEY = ALM_CIA7.ARM_CODART)  AND ART_SITUACION <> 1 and ART_KEY <> 0 AND ART_CALIDAD = 1  AND ART_CODCIA = '" & LK_CODCIA & "' AND (cia_actual.arm_codcia = '" & st_codcia1 & "') AND (cia_otra.arm_codcia = '" & st_codcia2 & "') AND (cia_SUBALM.arm_codcia = '" & st_codcia3 & "') AND (cia_JOLMEDO.arm_codcia = '" & st_codcia4 & "')  AND (ALM_CIA5.arm_codcia = '" & st_codcia5 & "') AND (ALM_CIA6.arm_codcia = '" & st_codcia6 & "')  AND (ALM_CIA7.arm_codcia = '" & st_codcia7 & "')  AND ART_NOMBRE BETWEEN '" & textovar.Text & "' AND  '" & var & "' " & SoloconStock & " ORDER BY ART_NOMBRE"
   Else
        SoloconStock = ""
        archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, cia_actual.ARM_STOCK as ARM_STOCK , PRE_EQUIV,  PRE_UNIDAD,PRE_PRE1,PRE_PRE2,PRE_PRE3,PRE_PRE4,PRE_PRE5,PRE_PRE6,cia_otra.ARM_STOCK as ARM_STOCK2, cia_SUBALM.ARM_STOCK as ARM_STOCK3, cia_JOLMEDO.ARM_STOCK as ARM_STOCK4 ,ALM_CIA5.ARM_STOCK as ARM_STOCK5,ALM_CIA6.ARM_STOCK as ARM_STOCK6 ,ALM_CIA7.ARM_STOCK as ARM_STOCK7, art_marca, ART_NUMERO, ART_ESTADO,  PRE_CANT , PRE_OP1  FROM ARTI, ARTICULO as cia_actual,ARTICULO as cia_otra, ARTICULO as cia_SUBALM, ARTICULO as cia_JOLMEDO, ARTICULO as ALM_CIA5, ARTICULO as ALM_CIA6, ARTICULO as ALM_CIA7, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD = 'A') AND (ART_KEY = cia_actual.ARM_CODART) AND (ART_KEY = cia_otra.ARM_CODART) AND (ART_KEY = cia_SUBALM.ARM_CODART)  AND " & _
        "  (ART_KEY = cia_JOLMEDO.ARM_CODART) AND (ART_KEY = ALM_CIA5.ARM_CODART)  AND (ART_KEY = ALM_CIA6.ARM_CODART) AND (ART_KEY = ALM_CIA7.ARM_CODART)  AND ART_SITUACION <> 1 and ART_KEY <> 0 AND ART_CALIDAD = 1  AND ART_CODCIA = '" & LK_CODCIA & "' AND (cia_actual.arm_codcia = '" & st_codcia1 & "') AND (cia_otra.arm_codcia = '" & st_codcia2 & "') AND (cia_SUBALM.arm_codcia = '" & st_codcia3 & "') AND (cia_JOLMEDO.arm_codcia = '" & st_codcia4 & "')  AND (ALM_CIA5.arm_codcia = '" & st_codcia5 & "') AND (ALM_CIA6.arm_codcia = '" & st_codcia6 & "')  AND (ALM_CIA7.arm_codcia = '" & st_codcia7 & "')  AND ART_NOMBRE BETWEEN '" & textovar.Text & "' AND  '" & var & "' " & SoloconStock & " ORDER BY ART_NOMBRE"
      End If
    End If
    PROC_LISVIEW ListView1, , "2"
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As Object     ' Variable FoundItem.
If ListView1.Visible Then
   
  Set itmFound = ListView1.FindItem(LTrim(textovar.Text), lvwText, , lvwPartial)
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

Private Sub txtatte_KeyPress(KeyAscii As Integer)
On Error GoTo sale
If KeyAscii = 13 Then
 grid_fac.Row = 2
 grid_fac.COL = 2
 grid_fac.COL = 1
 textovar.Visible = True
 textovar.SetFocus
End If
Exit Sub
sale:
End Sub

Private Sub txtdoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 PUB_NUMSER = Val(tserie.Text)
 PUB_NUMFAC = Val(txtdoc.Text)
 LLENA_DOCU
End If
End Sub

Private Sub txtruc_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  If Len(Trim(txtruc.Text)) <> LK_DIG_RUC Then
    MsgBox "R.U.C. No procede ", 48, Pub_Titulo
    Azul txtruc, txtruc
    Exit Sub
  Else
    moneda.SetFocus
  End If
End If

End Sub

Private Sub txtvend_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
If Val(txtvend.Text) = 0 Then
End If
End If

If KeyAscii = 13 Then
   txtatte.SetFocus
End If
End Sub

Private Sub txtvend_LostFocus()
If Val(txtvend.Text) <> 0 Then
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODVEN = Val(txtvend.Text)
    LEER_VEN_LLAVE
    If ven_llave.EOF Then
        MsgBox "Codigo No Existe!!!", 48, Pub_Titulo
        txtvend.Text = ""
        Azul txtvend, txtvend
        Exit Sub
    End If
    lblvend.Caption = Trim(ven_llave!VEM_NOMBRE)
    PSUSU_LLAVE(0) = LK_CODUSU
    usu_llave.Requery
    If Val(usu_llave!usu_codven) <> 0 Then
        If Val(usu_llave!usu_codven) <> PUB_CODVEN Then
            MsgBox "Usted no Tiene acceso a cambiar el vendedor", 48, Pub_Titulo
            txtvend.Text = Val(usu_llave!usu_codven)
            Azul txtvend, txtvend
            'moneda.SetFocus
        End If
    End If
    
End If
End Sub

Private Sub unidad_KeyPress(KeyAscii As Integer)
Dim wt_res  As Currency
Dim wt_mult  As Currency
Dim wt_equiv  As Currency

If KeyAscii = 27 Then
 unidad.Visible = False
 grid_fac.SetFocus
End If


If KeyAscii <> 13 Then Exit Sub
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
PUB_SECUEN = Val(Right(unidad.Text, 4))
LEER_PRE_LLAVE


If LK_FLAG_PARTES = "A" Then
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
    LEER_PRE_LLAVE
    wt_equiv = 1
    Do Until pre_mayor.EOF
     If pre_mayor!pre_FLAG_UNIDAD = "A" Then
        wt_equiv = pre_mayor!pre_equiv
     End If
    pre_mayor.MoveNext
    Loop
PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
SQ_OPER = 1
PUB_KEY = PUB_CODART
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
If Not art_LLAVE.EOF Then
   If Val(art_LLAVE!ART_MARGEN) <> 0 And (wt_equiv <> Val(pre_llave!pre_equiv)) Then
      wt_mult = Val(art_LLAVE!ART_MARGEN)
      If wt_mult = 0 Then
          MsgBox "Sin definición de partes minimas", 48, Pub_Titulo
          GoTo SINEQUIV
      End If
      wt_res = (Val(grid_fac.TextMatrix(grid_fac.Row, 2)) / wt_mult)
      If (Int(wt_res) - wt_res) <> 0 Then
        MsgBox "No Procede esta Cantidad, Verificar en las Unidades de Venta " & Chr(13) & "Unidades Minimas : " & wt_mult, 48, Pub_Titulo
SINEQUIV:
        unidad.Visible = False
        grid_fac.TextMatrix(grid_fac.Row, 3) = ""
        grid_fac.TextMatrix(grid_fac.Row, 4) = ""
        grid_fac.TextMatrix(grid_fac.Row, 5) = ""
        suma_grid
        grid_fac.COL = 2
        grid_fac.SetFocus
        Exit Sub
      End If
   End If
End If
End If
grid_fac.TextMatrix(grid_fac.Row, 3) = Trim(Left(unidad.Text, 12))
grid_fac.TextMatrix(grid_fac.Row, 4) = "" 'Format(Val(grid_fac.TextMatrix(grid_fac.Row, 11)) / Val(grid_fac.TextMatrix(grid_fac.Row, 17)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 7) = redondea(Nulo_Valor0(pre_llave!pre_PESO) * Val(grid_fac.TextMatrix(grid_fac.Row, 2)))
grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!PRE_SECUENCIA
grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!pre_equiv
grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_unidad
grid_fac.TextMatrix(grid_fac.Row, 19) = Format(Jala_StockGen(pre_llave!PRE_codart, pre_llave!pre_equiv), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 20) = Format(Nulo_Valor0(grid_fac.TextMatrix(grid_fac.Row, 18)) * pre_llave!pre_equiv, "0.000")
stock.Caption = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 15)) / Val(grid_fac.TextMatrix(grid_fac.Row, 12)), "0.00")
unid.Caption = grid_fac.TextMatrix(grid_fac.Row, 16)
nomarti.Caption = grid_fac.TextMatrix(grid_fac.Row, 0)

unidad.Visible = False
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
LOC_TIPMOV = 301
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
  

    FORM_COTV.Reportes.Connect = PUB_ODBC
    FORM_COTV.Reportes.Destination = crptToWindow  '= crptToPrinter
    FORM_COTV.Reportes.WindowLeft = 2
    FORM_COTV.Reportes.WindowTop = 70
    FORM_COTV.Reportes.WindowWidth = 635
    FORM_COTV.Reportes.WindowHeight = 390
    FORM_COTV.Reportes.Formulas(1) = ""
    PUB_NETO = Val(txttotal.Text)
    PU_NUMSER = Val((tserie.Text))
    PU_NUMFAC = Val((txtdoc.Text))
    FORM_COTV.Reportes.Formulas(1) = ""
    FORM_COTV.Reportes.Formulas(1) = "SON_EFECTIVO=  ' " & CONVER_LETRAS(PUB_NETO, WMONEDA) & "'"
    FORM_COTV.Reportes.WindowTitle = "GUIA DE COTIZACION  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "00000000")
    If chemarg.Value = 1 Then
    FORM_COTV.Reportes.ReportFileName = wRuta + "COTIOF.RPT"
    Else
    FORM_COTV.Reportes.ReportFileName = wRuta + "COTI.RPT"
    End If
    pub_cadena = "{PEDIDOS.PED_TIPMOV} = " & LOC_TIPMOV & " AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND  {PEDIDOS.PED_NUMSER} = '" & PU_NUMSER & "' AND {PEDIDOS.PED_NUMFAC} = " & PU_NUMFAC
    FORM_COTV.Reportes.SelectionFormula = pub_cadena
    On Error GoTo accion
    FORM_COTV.Reportes.Action = 1
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
Dim itmFound As Object     ' Variable FoundItem.
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
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As Object
On Error GoTo SALCODI

If KeyAscii = 27 Then
 txtcli.Text = ""
 lblcli.Caption = ""
 lbldeuda.Caption = ""
 lblcredito.Caption = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
pu_codclie = Val(txtcli.Text)
On Error GoTo CODI_ERR

If pu_codclie <> 0 And IsNumeric(txtcli.Text) = True Then
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
   FORM_COTV.lblcli.Caption = cli_llave!cli_nombre
   If Trim(cli_llave!cli_ruc_esposo) <> "" Then
     txtruc.Text = cli_llave!cli_ruc_esposo
   Else
     txtruc.Text = cli_llave!cli_RUC_ESPOSA
   End If
   moneda.SetFocus
   GoTo ver_credito
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
   FORM_COTV.lblcli.Caption = cli_llave!cli_nombre
   If Trim(cli_llave!cli_ruc_esposo) <> "" Then
     txtruc.Text = cli_llave!cli_ruc_esposo
   Else
     txtruc.Text = cli_llave!cli_RUC_ESPOSA
   End If
   moneda.SetFocus
   GoTo ver_credito

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
ver_credito:
pub_deuda = CAR_TOT_CPX("C", LK_CODCIA, cli_llave!cli_codclie)
lbldeuda.Caption = Format(pub_deuda, "0.00")
lblcredito.Caption = Format(Val(cli_llave!CLI_LIMCRE) - pub_deuda, "0.00")


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
    archi = "SELECT TOP 2000  CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = 'C' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txtcli.Text & "' AND  '" & var & "' ORDER BY CLI_NOMBRE"
    PROC_LISVIEW ListView1
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
Dim itmFound As Object     ' Variable FoundItem.
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
PSLOC_WARTI(1) = 301
PSLOC_WARTI(2) = PUB_NUMSER
PSLOC_WARTI(3) = PUB_NUMFAC
llave_sum_arti.Requery
If llave_sum_arti.EOF Then
  tserie.Text = PUB_NUMSER
  txtdoc.Text = PUB_NUMFAC
  MsgBox "No Existe Cotización.", 48, Pub_Titulo
  Exit Sub
End If
PSUSU_LLAVE(0) = LK_CODUSU
usu_llave.Requery
If Val(usu_llave!usu_codven) <> 0 Then
If Val(Nulo_Valor0(llave_sum_arti!ped_codven)) <> Val(usu_llave!usu_codven) Then
    MsgBox "Usted no Tiene Acceso a ver Esta Infomración.", 48, Pub_Titulo
    Exit Sub
End If
End If



txtigv.Text = llave_sum_arti!PED_IGV
txtvalorv.Text = llave_sum_arti!PED_BRUTO
txttotal.Text = Format(llave_sum_arti!PED_IGV + llave_sum_arti!PED_BRUTO, "0.00")
FORM_COTV.lblcli.Caption = llave_sum_arti!PED_NOMCLIE
txtruc.Text = llave_sum_arti!PED_RUCCLIE
txtcli.Text = llave_sum_arti!PED_CODCLIE

If llave_sum_arti!PED_MONEDA = "S" Then
 moneda.ListIndex = 0
 i_moneda.Caption = "S/."
 grid_fac.TextMatrix(1, 4) = "S/."
Else
 moneda.ListIndex = 1
 i_moneda.Caption = "US$."
 grid_fac.TextMatrix(1, 4) = "US$."
End If
txtvend.Text = Nulo_Valor0(llave_sum_arti!ped_codven)
tserie.Text = PUB_NUMSER
txtdoc.Text = PUB_NUMFAC
txtatte.Text = Nulo_Valors(llave_sum_arti!PED_CONTACTO)
Che_Pedidos.Value = Val(llave_sum_arti!PED_FLAG_PRE)
LBLHORA.Caption = Format(llave_sum_arti!PED_FECHA, "dd/mm/yyyy") & " " & llave_sum_arti!PED_HORA

fila = 2
Do Until llave_sum_arti.EOF
   SQ_OPER = 1
   PUB_KEY = llave_sum_arti!PED_codart
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   grid_fac.Rows = grid_fac.Rows + 1
   grid_fac.RowHeight(grid_fac.Rows - 1) = 285
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!art_alterno
   grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   If LK_EMP = "3AA" Then
     grid_fac.TextMatrix(fila, 11) = 0
     grid_fac.TextMatrix(fila, 14) = 1
   End If
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_CODART = PUB_KEY
   LEER_ARM_LLAVE
   grid_fac.TextMatrix(fila, 2) = llave_sum_arti!PED_CANTIDAD
   grid_fac.TextMatrix(fila, 4) = llave_sum_arti!PED_PRECIO
   grid_fac.TextMatrix(fila, 10) = llave_sum_arti!PED_codart
   grid_fac.TextMatrix(fila, 3) = llave_sum_arti!PED_UNIDAD
   grid_fac.TextMatrix(fila, 12) = llave_sum_arti!PED_EQUIV
   grid_fac.TextMatrix(fila, 5) = llave_sum_arti!PED_descto
   grid_fac.TextMatrix(fila, 18) = Nulo_Valor0(arm_llave!ARM_COSPRO)
   grid_fac.TextMatrix(fila, 19) = Format(Jala_StockGen(llave_sum_arti!PED_codart, llave_sum_arti!PED_EQUIV), "0.00")
   grid_fac.TextMatrix(fila, 20) = Format(Nulo_Valor0(arm_llave!ARM_COSPRO) * llave_sum_arti!PED_EQUIV, "0.000")
   grid_fac.TextMatrix(fila, 21) = art_LLAVE!ART_EX_IGV
   fila = fila + 1
   llave_sum_arti.MoveNext
Loop
suma_grid
grid_fac.Enabled = True
'f1.Enabled = False
ESTADO.Enabled = True
Azul txtdoc, txtdoc
cmdIngreso.Caption = "&Grabar"
cmdIngreso.Enabled = True
tserie.Enabled = False
txtdoc.Enabled = False
End Sub

Public Function Jala_StockGen(WS_CODART As Currency, ws_equiv As Currency) As Currency
Dim ps_stockg As rdoResultset
pub_cadena = "SELECT SUM(ARM_STOCK) AS STOCKG FROM ARTICULO WHERE  ARM_CODART = " & WS_CODART & " AND ARM_CODCIA IN ('01','02','04','07','08')"
Set ps_stockg = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If ps_stockg.EOF Then
  Jala_StockGen = 0
Else
  If ws_equiv <> 0 Then
    Jala_StockGen = Format(ps_stockg!STOCKG / ws_equiv, "0.00")
  Else
    Jala_StockGen = ps_stockg!STOCKG
  End If
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_CODART = WS_CODART
  LEER_ARM_LLAVE
  If arm_llave.EOF Then
    lblcosprog.Caption = 0
  Else
    lblcosprog.Caption = Format(arm_llave!ARM_COSPRO * ws_equiv, "0.0000")
  End If
End If
End Function


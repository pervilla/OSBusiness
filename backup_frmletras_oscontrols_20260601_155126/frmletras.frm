VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F6E4F630-E903-11D5-8BB9-0080AD40A177}#1.18#0"; "OSControlsUser.ocx"
Begin VB.Form frmletras 
   Caption         =   "Consulta de Letras"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Reportes 
      Left            =   690
      Top             =   7170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraLetra 
      Height          =   7155
      Left            =   10920
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   11775
      Begin TabDlg.SSTab SSTab1 
         Height          =   4425
         Left            =   75
         TabIndex        =   39
         Top             =   2655
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   7805
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         ForeColor       =   10235904
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Relacion de Letras"
         TabPicture(0)   =   "frmletras.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraDatos(1)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Edicion de Letras"
         TabPicture(1)   =   "frmletras.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fraDatos(0)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame fraDatos 
            Height          =   4020
            Index           =   1
            Left            =   -74910
            TabIndex        =   73
            Top             =   315
            Width           =   11415
            Begin VB.OptionButton OptDocId 
               Caption         =   "Print Ruc"
               Height          =   195
               Index           =   1
               Left            =   3945
               TabIndex        =   97
               Top             =   3750
               Value           =   -1  'True
               Width           =   1740
            End
            Begin VB.OptionButton OptDocId 
               Caption         =   "Print Dni"
               Height          =   195
               Index           =   0
               Left            =   3960
               TabIndex        =   96
               Top             =   3510
               Width           =   1740
            End
            Begin VB.CheckBox chkPrint 
               Caption         =   "Directo a Impresora"
               Height          =   195
               Left            =   6060
               TabIndex        =   91
               Top             =   3765
               Width           =   1680
            End
            Begin VB.CheckBox chkConyugue 
               Caption         =   "Imprimir Conyugue"
               Height          =   195
               Left            =   6060
               TabIndex        =   87
               Top             =   3525
               Width           =   1680
            End
            Begin VB.CommandButton cmdImprimir 
               BackColor       =   &H00F5F1EC&
               Caption         =   "Imprimir Letras"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   195
               Style           =   1  'Graphical
               TabIndex        =   86
               Top             =   3525
               Width           =   1530
            End
            Begin VB.CommandButton cmdImprimir1 
               BackColor       =   &H00F5F1EC&
               Caption         =   "Imprimir en Fact."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1755
               Style           =   1  'Graphical
               TabIndex        =   85
               Top             =   3525
               Width           =   1755
            End
            Begin VB.CommandButton cmdCancelar1 
               BackColor       =   &H00F5F1EC&
               Caption         =   "Cancelar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   10050
               Style           =   1  'Graphical
               TabIndex        =   82
               Top             =   3525
               Width           =   1290
            End
            Begin VB.CommandButton cmdGrabar1 
               BackColor       =   &H00F5F1EC&
               Caption         =   "Grabar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   8730
               Style           =   1  'Graphical
               TabIndex        =   81
               Top             =   3525
               Width           =   1290
            End
            Begin MSComctlLib.ListView lvwLetras 
               Height          =   3195
               Index           =   1
               Left            =   105
               TabIndex        =   75
               Top             =   255
               Width           =   11070
               _ExtentX        =   19526
               _ExtentY        =   5636
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
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
               MouseIcon       =   "frmletras.frx":0038
               NumItems        =   0
            End
         End
         Begin VB.Frame fraDatos 
            Height          =   4020
            Index           =   0
            Left            =   90
            TabIndex        =   40
            Top             =   315
            Width           =   11415
            Begin VB.CommandButton cmdDistImpLet 
               BackColor       =   &H00F5F1EC&
               Caption         =   "Dist. Importe"
               Height          =   360
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   90
               Tag             =   "104"
               Top             =   3615
               Width           =   1170
            End
            Begin VB.TextBox txtSubTotal 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2265
               TabIndex        =   89
               Top             =   3645
               Width           =   1125
            End
            Begin MSComCtl2.MonthView txtFechaEntrega1 
               Height          =   2370
               Left            =   3255
               TabIndex        =   77
               Top             =   3870
               Visible         =   0   'False
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   4180
               _Version        =   393216
               ForeColor       =   -2147483630
               BackColor       =   -2147483633
               Appearance      =   1
               StartOfWeek     =   49348609
               CurrentDate     =   37978
            End
            Begin MSComCtl2.MonthView txtFechaDevolucion1 
               Height          =   2370
               Left            =   4185
               TabIndex        =   80
               Top             =   3765
               Visible         =   0   'False
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   4180
               _Version        =   393216
               ForeColor       =   -2147483630
               BackColor       =   -2147483633
               Appearance      =   1
               StartOfWeek     =   49348609
               CurrentDate     =   37978
            End
            Begin VB.TextBox txtCodUniBko 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8130
               TabIndex        =   84
               Top             =   1020
               Width           =   3105
            End
            Begin VB.CommandButton cmdF2 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   10005
               TabIndex        =   79
               Top             =   2925
               Width           =   345
            End
            Begin VB.CommandButton cmdF1 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   7110
               TabIndex        =   78
               Top             =   2925
               Width           =   345
            End
            Begin VB.ComboBox cboBanco 
               Height          =   315
               Left            =   3540
               Style           =   2  'Dropdown List
               TabIndex        =   50
               Top             =   1590
               Width           =   4365
            End
            Begin VB.TextBox txtObservaciones 
               Height          =   285
               Left            =   3525
               TabIndex        =   49
               Top             =   2205
               Width           =   4365
            End
            Begin VB.ComboBox cboSituacion 
               Height          =   315
               Left            =   3525
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   1005
               Width           =   4365
            End
            Begin VB.CheckBox chkProtestada 
               Height          =   195
               Left            =   4800
               TabIndex        =   47
               Top             =   2790
               Width           =   225
            End
            Begin VB.TextBox txtImporte 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   10095
               TabIndex        =   46
               Top             =   1455
               Width           =   1080
            End
            Begin VB.TextBox txtAmort 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   10110
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   1815
               Width           =   1080
            End
            Begin VB.TextBox txtSaldos 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   10125
               Locked          =   -1  'True
               TabIndex        =   44
               Top             =   2160
               Width           =   1080
            End
            Begin VB.CommandButton cmdGrabar 
               BackColor       =   &H00F5F1EC&
               Caption         =   "Grabar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   8745
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   3525
               Width           =   1290
            End
            Begin VB.CommandButton cmdCancelar 
               BackColor       =   &H00F5F1EC&
               Caption         =   "Cerrar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   10050
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   3525
               Width           =   1290
            End
            Begin VB.CheckBox chkAceptada 
               Height          =   195
               Left            =   4785
               TabIndex        =   41
               Top             =   3180
               Width           =   225
            End
            Begin MSComCtl2.DTPicker txtFechaEmi 
               Height          =   315
               Left            =   3525
               TabIndex        =   52
               Top             =   420
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   556
               _Version        =   393216
               Format          =   49348609
               CurrentDate     =   37977
            End
            Begin OSControlsUser.ctlMaskEdBox txtFechaVcto 
               Height          =   270
               Left            =   8115
               TabIndex        =   53
               Top             =   420
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   476
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               Format          =   "ddddd"
            End
            Begin OSControlsUser.ctlMaskEdBox txtFechaEntrega 
               Height          =   270
               Left            =   6030
               TabIndex        =   54
               Top             =   2925
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   476
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               Format          =   "ddddd"
            End
            Begin OSControlsUser.ctlMaskEdBox txtFechaDevolucion 
               Height          =   270
               Left            =   8895
               TabIndex        =   55
               Top             =   2925
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   476
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               Format          =   "ddddd"
            End
            Begin MSComCtl2.UpDown CountDias 
               Height          =   270
               Left            =   7590
               TabIndex        =   56
               Top             =   450
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   476
               _Version        =   393216
               Value           =   1
               Max             =   1000
               Min             =   1
               Enabled         =   -1  'True
            End
            Begin MSComctlLib.ListView lvwLetras 
               Height          =   3345
               Index           =   2
               Left            =   105
               TabIndex        =   76
               Top             =   255
               Width           =   3300
               _ExtentX        =   5821
               _ExtentY        =   5900
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
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
               MouseIcon       =   "frmletras.frx":0352
               NumItems        =   0
            End
            Begin VB.TextBox txtdias 
               Height          =   330
               Left            =   7200
               TabIndex        =   51
               Text            =   "1"
               Top             =   420
               Width           =   660
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubTotal"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   35
               Left            =   1455
               TabIndex        =   88
               Top             =   3675
               Width           =   750
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Codigo Unico"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   34
               Left            =   8130
               TabIndex        =   83
               Top             =   810
               Width           =   1080
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Emisión"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   4
               Left            =   3525
               TabIndex        =   71
               Top             =   180
               Width           =   1185
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Vencimiento"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   5
               Left            =   8115
               TabIndex        =   70
               Top             =   180
               Width           =   1590
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dias"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   6
               Left            =   7215
               TabIndex        =   69
               Top             =   180
               Width           =   360
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Situación"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   7
               Left            =   3525
               TabIndex        =   68
               Top             =   750
               Width           =   780
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Banco"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   8
               Left            =   3525
               TabIndex        =   67
               Top             =   1365
               Width           =   510
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Observaciones"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   9
               Left            =   3525
               TabIndex        =   66
               Top             =   1965
               Width           =   1245
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Protestada"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   10
               Left            =   3645
               TabIndex        =   65
               Top             =   2790
               Width           =   945
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Entrega"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   11
               Left            =   6015
               TabIndex        =   64
               Top             =   2700
               Width           =   1200
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Devolución"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   12
               Left            =   8895
               TabIndex        =   63
               Top             =   2700
               Width           =   1470
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Importe"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   13
               Left            =   8175
               TabIndex        =   62
               Top             =   1515
               Width           =   705
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Amortizaciones"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   14
               Left            =   8175
               TabIndex        =   61
               Top             =   1815
               Width           =   1320
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saldos"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   15
               Left            =   8175
               TabIndex        =   60
               Top             =   2130
               Width           =   555
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Aceptada"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   195
               Index           =   23
               Left            =   3675
               TabIndex        =   59
               Top             =   3180
               Width           =   810
            End
            Begin VB.Label lblMoneda 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   240
               Left            =   9585
               TabIndex        =   58
               Top             =   1410
               Width           =   60
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   1110
               Index           =   16
               Left            =   8100
               TabIndex        =   57
               Top             =   1395
               Width           =   3180
            End
            Begin VB.Label lblCaption 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H009C3000&
               Height          =   780
               Index           =   17
               Left            =   5475
               TabIndex        =   72
               Top             =   2625
               Width           =   5850
            End
         End
      End
      Begin VB.ComboBox cbovendedor 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1665
         Width           =   3480
      End
      Begin VB.ComboBox cbomoneda 
         Height          =   315
         ItemData        =   "frmletras.frx":066C
         Left            =   8085
         List            =   "frmletras.frx":0676
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1665
         Width           =   900
      End
      Begin VB.TextBox txtCambio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8100
         TabIndex        =   30
         Top             =   2160
         Width           =   885
      End
      Begin MSComCtl2.DTPicker txtFecEmiGrl 
         Height          =   315
         Left            =   1590
         TabIndex        =   33
         Top             =   2160
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49348609
         CurrentDate     =   37977
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   29
         Left            =   300
         TabIndex        =   37
         Top             =   1665
         Width           =   810
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   30
         Left            =   7230
         TabIndex        =   36
         Top             =   1665
         Width           =   675
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T/Cambio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   31
         Left            =   7200
         TabIndex        =   35
         Top             =   2160
         Width           =   825
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fec. Emisión"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   32
         Left            =   300
         TabIndex        =   34
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblTotal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   8940
         TabIndex        =   24
         Top             =   975
         Width           =   1110
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblSaldo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   8895
         TabIndex        =   23
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Documento :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   25
         Left            =   7230
         TabIndex        =   22
         Top             =   975
         Width           =   1545
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo de Documento :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   24
         Left            =   7230
         TabIndex        =   21
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label lblDocumento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   8280
         TabIndex        =   9
         Top             =   210
         Width           =   3240
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRuc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1230
         TabIndex        =   8
         Top             =   975
         Width           =   630
      End
      Begin VB.Label lblDireccion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1245
         TabIndex        =   7
         Top             =   660
         Width           =   630
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1230
         TabIndex        =   6
         Top             =   345
         Width           =   630
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   22
         Left            =   7215
         TabIndex        =   5
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUC :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   21
         Left            =   315
         TabIndex        =   4
         Top             =   975
         Width           =   420
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   20
         Left            =   315
         TabIndex        =   3
         Top             =   660
         Width           =   750
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   18
         Left            =   315
         TabIndex        =   1
         Top             =   345
         Width           =   600
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00404040&
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
         ForeColor       =   &H009C3000&
         Height          =   1260
         Index           =   19
         Left            =   60
         TabIndex        =   2
         Top             =   150
         Width           =   11640
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H009C3000&
         Height          =   1260
         Index           =   33
         Left            =   60
         TabIndex        =   38
         Top             =   1380
         Width           =   11640
      End
   End
   Begin VB.Frame fraDocumentos 
      Height          =   7155
      Left            =   0
      TabIndex        =   10
      Top             =   -15
      Width           =   11775
      Begin VB.TextBox b_numfac 
         Height          =   285
         Left            =   4440
         TabIndex        =   103
         Top             =   600
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwCanjes 
         Height          =   5820
         Left            =   120
         TabIndex        =   94
         Top             =   1230
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   10266
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox cmbCP 
         Height          =   315
         ItemData        =   "frmletras.frx":068C
         Left            =   1215
         List            =   "frmletras.frx":0696
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   225
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker txtfecha1 
         Height          =   315
         Left            =   8625
         TabIndex        =   28
         Top             =   285
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49348609
         CurrentDate     =   37977
      End
      Begin VB.CheckBox chkClientes 
         BackColor       =   &H00404040&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   13
         Top             =   285
         Value           =   1  'Checked
         Width           =   885
      End
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00F5F1EC&
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10335
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   1260
      End
      Begin VB.CommandButton cmdCerrar 
         BackColor       =   &H00F5F1EC&
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6660
         Width           =   1260
      End
      Begin OSControlsUser.OSFindItem txtCP 
         Height          =   300
         Left            =   1230
         TabIndex        =   14
         Top             =   615
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         Locked          =   0   'False
      End
      Begin MSComctlLib.ListView lvwDocumentos 
         Height          =   2310
         Left            =   5295
         TabIndex        =   15
         Top             =   1440
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   4075
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         MouseIcon       =   "frmletras.frx":06EC
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker txtfecha2 
         Height          =   315
         Left            =   8625
         TabIndex        =   29
         Top             =   675
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49348609
         CurrentDate     =   37977
      End
      Begin MSComctlLib.ListView lvwLetras 
         Height          =   2445
         Index           =   0
         Left            =   5310
         TabIndex        =   74
         Top             =   4095
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   4313
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         MouseIcon       =   "frmletras.frx":0A06
         NumItems        =   0
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Docum.:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Index           =   40
         Left            =   3600
         TabIndex        =   104
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Index           =   39
         Left            =   3600
         TabIndex        =   102
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblsalcanje 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   8520
         TabIndex        =   101
         Top             =   3840
         Width           =   1050
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo: Canje:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   38
         Left            =   7440
         TabIndex        =   100
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label lbltotcanje 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10560
         TabIndex        =   99
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tot. Canje:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   37
         Left            =   9600
         TabIndex        =   98
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Determinar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Index           =   36
         Left            =   165
         TabIndex        =   93
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Doble Click para editar>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   28
         Left            =   9705
         TabIndex        =   27
         Top             =   3825
         Width           =   1845
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   27
         Left            =   5400
         TabIndex        =   26
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Letras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   26
         Left            =   5550
         TabIndex        =   25
         Top             =   3855
         Width           =   540
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razon Soc"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   19
         Top             =   645
         Width           =   975
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Index           =   1
         Left            =   7740
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Index           =   2
         Left            =   7755
         TabIndex        =   17
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblCP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2595
         TabIndex        =   16
         Top             =   675
         Width           =   60
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00404040&
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
         ForeColor       =   &H009C3000&
         Height          =   930
         Index           =   3
         Left            =   45
         TabIndex        =   20
         Top             =   150
         Width           =   11655
      End
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
      TabIndex        =   95
      Top             =   7185
      Width           =   11895
   End
End
Attribute VB_Name = "frmletras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim loListItem As MSComctlLib.ListItem
Dim loListItem1 As MSComctlLib.ListItem
Dim loListItem2 As MSComctlLib.ListItem
Dim sCP As String
Dim WS_NUM_OPER2 As Integer
Dim sCodClie As Long
Dim sNumDocLetra As Integer
Dim sSerDocLetra As Integer
Dim VARTEMP As String
Dim sSaldoDocumento As Double
Dim sTotalDocumento As Double
Dim MonedaOriginal As String
Dim FlagMoney As Integer
Dim sDocumentos As String
Dim sConyugue As String
Dim sFechaDia As String
Dim sNumOper1 As Long
Dim RowActual As Integer
Dim RowActual1 As Integer

Private Sub cbomoneda_Click()
Dim I As Integer
Dim sImporte As Double
Dim Total_Cambio As Double
Dim TotalTmp As Double
Dim TotalxLetra As Double
Dim sTotalDocumento1 As Double

    If FlagMoney = 0 Then Exit Sub
    If MonedaOriginal = "S" And Left(cbomoneda.Text, 1) = "D" Then
        sTotalDocumento1 = Format(sTotalDocumento / Val(txtCambio.Text), "#0.00")
    ElseIf MonedaOriginal = "D" And Left(cbomoneda.Text, 1) = "S" Then
        sTotalDocumento1 = Format(sTotalDocumento * Val(txtCambio.Text), "#0.00")
    ElseIf MonedaOriginal = Left(cbomoneda.Text, 1) Then
        sTotalDocumento1 = sTotalDocumento
    End If
    
'    If FlagMoney = 1 And Left(cbomoneda.Text, 1) = "D" Then
'        Total_Cambio = Format(sTotalDocumento1 / Val(txtCambio.Text), "#0.00")
'    ElseIf FlagMoney = 1 And Left(cbomoneda.Text, 1) = "S" Then
'        Total_Cambio = Format(sTotalDocumento1 * Val(txtCambio.Text), "#0.00")
'    End If
    TotalxLetra = Format(sTotalDocumento1 / lvwLetras(1).ListItems.count, "0.00")
    
    For I = 1 To lvwLetras(1).ListItems.count
        lvwLetras(1).ListItems(I).ListSubItems(4).Text = Left(cbomoneda.Text, 1)
        'sImporte = lvwLetras(1).ListItems(i).ListSubItems(7).Text
        'sImporte = Format(sImporte / Val(txtCambio.Text), "#0.00")
        'lvwLetras(1).ListItems(i).ListSubItems(7).Text = sImporte
        '*****************
        'If i = lvwLetras(1).ListItems.count Then
        '    lvwLetras(1).ListItems(i).ListSubItems(7).Text = sTotalDocumento1 - TotalTmp
        'Else
        '    lvwLetras(1).ListItems(i).ListSubItems(7).Text = TotalxLetra
        '    TotalTmp = TotalTmp + TotalxLetra
        'End If
        '*****************
        TotalxLetra = Val(lvwLetras(1).ListItems(I).ListSubItems(7).Text)
        If MonedaOriginal = "S" And Left(cbomoneda.Text, 1) = "D" Then
            TotalxLetra = Format(TotalxLetra / Val(txtCambio.Text), "#0.00")
        ElseIf MonedaOriginal = "D" And Left(cbomoneda.Text, 1) = "S" Then
            TotalxLetra = Format(TotalxLetra * Val(txtCambio.Text), "#0.00")
        ElseIf MonedaOriginal = Left(cbomoneda.Text, 1) Then
            TotalxLetra = Val(lvwLetras(1).ListItems(I).ListSubItems(15).Text)
        End If
        
        If I = lvwLetras(1).ListItems.count Then
            lvwLetras(1).ListItems(I).ListSubItems(7).Text = Format(sTotalDocumento1 - TotalTmp, "#0.00")
            lvwLetras(1).ListItems(I).ListSubItems(8).Text = Format(sTotalDocumento1 - TotalTmp, "#0.00")
            lvwLetras(1).ListItems(I).ListSubItems(12).Text = Format(sTotalDocumento1 - TotalTmp, "#0.00")
            lvwLetras(1).ListItems(I).ListSubItems(13).Text = Format(sTotalDocumento1 - TotalTmp, "#0.00")
        Else
            lvwLetras(1).ListItems(I).ListSubItems(7).Text = Format(TotalxLetra, "#0.00")
            lvwLetras(1).ListItems(I).ListSubItems(8).Text = Format(TotalxLetra, "#0.00")
            lvwLetras(1).ListItems(I).ListSubItems(12).Text = Format(TotalxLetra, "#0.00")
            lvwLetras(1).ListItems(I).ListSubItems(13).Text = Format(TotalxLetra, "#0.00")
            TotalTmp = TotalTmp + TotalxLetra
        End If
        

        
        
        
'            sImporte = lvwLetras(1).ListItems(i).ListSubItems(12).Text
'            sImporte = Format(sImporte / Val(txtCambio.Text), "#0.00")
'            lvwLetras(1).ListItems(i).ListSubItems(12).Text = sImporte
'
'            sImporte = lvwLetras(1).ListItems(i).ListSubItems(13).Text
'            sImporte = Format(sImporte / Val(txtCambio.Text), "#0.00")
'            lvwLetras(1).ListItems(i).ListSubItems(13).Text = sImporte
    Next I
End Sub

Private Sub cboSituacion_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 45 Then
      Exit Sub
    End If
    PUB_TIPREG = 160
    PUB_CODCIA = "00"
    Load FrmDatArti
    FrmDatArti.Caption = "Situacion  -  TAB_TIPREG = " & PUB_TIPREG
    FrmDatArti.Show 1
    DoEvents
    VARTEMP = LK_EMP_PTO
    LK_EMP_PTO = "A"
    LlenadoCbo cboSituacion, 160
    LK_EMP_PTO = VARTEMP
    cboSituacion.SetFocus
    SendKeys "%{up}"
End Sub

Private Sub chkClientes_Click()
    lblCP.Caption = ""
    txtCP.TEXTO = ""
End Sub

Private Sub cmbCP_Click()
    If Trim(Right(cmbCP.Text, 3)) <> sCP Then
        lvwCanjes.ListItems.Clear
        lvwDocumentos.ListItems.Clear
        lvwLetras(0).ListItems.Clear
    End If
    sCP = Trim(Right(cmbCP.Text, 3))
End Sub

Private Sub CmdCancelar_Click()
    If Val(txtSubTotal.Text) <> Val(sTotalDocumento) Then
 '       MsgBox "Error en el Importe no Coinciden", vbCritical, Pub_Titulo
'        Exit Sub
    End If
    fraLetra.Visible = False
    fraDocumentos.Visible = True
    txtFechaVcto.Text = "__/__/____"
    'txtFechaEmi.Text = "__/__/____"
    txtdias.Text = 0
    txtImporte.Text = "0.00"
    txtSaldos.Text = "0.00"
    txtAmort.Text = "0.00"
    cboSituacion.ListIndex = -1
    cboBanco.ListIndex = -1
    txtObservaciones.Text = ""
    chkProtestada.Value = 0
    chkAceptada.Value = 0
    txtFechaEntrega.Text = "__/__/____"
    txtFechaDevolucion.Text = "__/__/____"
    lblSaldo.Caption = "0.00"
    lblTotal.Caption = "0.00"
    txtFechaEntrega1.Visible = False
    txtFechaDevolucion1.Visible = False
    FlagMoney = 0
End Sub

Private Sub cmdCancelar1_Click()
    fraLetra.Visible = False
    fraDocumentos.Visible = True
    FlagMoney = 0
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdF1_Click()
    txtFechaEntrega1.Left = 6045
    txtFechaEntrega1.Top = 550
    txtFechaEntrega1.Visible = True
    txtFechaEntrega1.SetFocus
End Sub

Private Sub cmdF2_Click()
    txtFechaDevolucion1.Top = 550
    txtFechaDevolucion1.Left = 8760
    txtFechaDevolucion1.Visible = True
    txtFechaDevolucion1.SetFocus
End Sub

Private Sub cmdGrabar_Click()
Dim sx_tipdoc As String
Dim sx_numdoc As Long
Dim sx_codtra As Integer
Dim sx_subtra As String

sx_tipdoc = Left(lvwLetras(0).ListItems(RowActual).Text, 2)
sx_numdoc = lvwLetras(2).ListItems(RowActual).ListSubItems(6).Text
On Error GoTo Handler

    If Consistencias = False Then
        Exit Sub
    End If
    pub_cadena = "SELECT * FROM CONTROLL"
    CN.Execute "Begin Transaction", rdExecDirect
    Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
    
    'ALL_MONEDA_CAJA = '" & Left(cbomoneda.Text, 1) & "', ALL_TIPO_CAMBIO=" & Val(txtCambio.Text) & ",
    archi = "UPDATE ALLOG SET  ALL_FECHA_SUNAT = '" & txtFechaEmi.Value & "', ALL_CODVEN=" & Val(Right(cbovendedor.Text, 6)) & " WHERE ALL_TIPDOC='" & sx_tipdoc & "' AND ALL_CODTRA=1455 AND ALL_CODCIA='" & LK_CODCIA & "' AND ALL_FECHA_DIA='" & sFechaDia & "' AND ALL_NUMOPER = " & sNumOper1
    CN.Execute archi
    
    car_llave.Edit
    car_llave("car_FECHA_SUNAT") = txtFechaEmi
    car_llave("car_voucher") = Val(Right(cboSituacion.Text, 6)) 'situacion
    car_llave("car_codban") = Val(Right(cboBanco.Text, 6)) 'banco
    car_llave("car_nombre_banco") = Left(cboBanco.Text, 30) 'banco
    car_llave("car_concepto") = txtObservaciones.Text 'observacion
    car_llave("car_placa") = chkProtestada.Value 'protestada
    car_llave("car_precio") = chkAceptada.Value 'aceptada
    If IsDate(txtFechaEntrega.Text) Then
        car_llave("car_fecha_entrega") = txtFechaEntrega.Text 'fechaentrega
    End If
    If IsDate(txtFechaDevolucion.Text) Then
        car_llave("car_fecha_devo") = txtFechaDevolucion.Text 'fechadevolucion
    End If
    car_llave("CAR_FECHA_VCTO") = txtFechaVcto.Text
    car_llave("CAR_FECHA_VCTO_ORIG") = txtFechaVcto.Text
    car_llave("CAR_CODVEN") = Val(Right(cbovendedor.Text, 6))
    car_llave("CAR_COBRADOR") = Val(Right(cbovendedor.Text, 6))
                
    car_llave("CAR_IMPORTE") = Val(txtSaldos.Text)
    car_llave("CAR_IMP_INI") = Val(txtImporte.Text)
    car_llave("CAR_CODUNIBKO") = Left(txtCodUniBko.Text, 30)
    car_llave.Update
    
    SQ_OPER = 1
    pu_cp = sCP
    pu_codclie = sCodClie
    pu_codcia = LK_CODCIA
    PUB_TIPDOC = sx_tipdoc
    LEER_CAA_LLAVE
    caa_histo.Requery
    If Not caa_histo.EOF Then
        caa_histo.Edit
        caa_histo("CAA_FECHA_COBRO") = txtFechaEmi.Value
        caa_histo("CAA_FECHA_VCTO") = txtFechaVcto.Text
        
        caa_histo("CAA_IMPORTE") = Val(txtImporte.Text)
        'caa_histo ("CAA_SALDO") 'NO TOMAR EN CUENTA
        caa_histo("CAA_TOTAL") = Val(txtImporte.Text)
        caa_histo("CAA_SALDO_CAR") = Val(txtImporte.Text)
            
        caa_histo.Update
    End If
    'GRABAR COMO TRANSACCION PARA MODULO DE PASE A CONTABILIDAD
    If txtCodUniBko.Text <> "" Then
        archi = "SELECT * FROM allog WHERE all_flag_ext <> 'E' and all_cp = '" & sCP & "' and ALL_TIPDOC='" & sx_tipdoc & "' AND ALL_CODTRA in (1466,1499) AND ALL_CODCIA='" & LK_CODCIA & "' AND ALL_tipdoc = '" & sx_tipdoc & "' and all_numdoc=" & sx_numdoc
        Set PSX = CN.CreateQuery("", archi)
        Set X = PSX.OpenResultset(rdOpenKeyset, rdConcurValues)
        X.Requery
        If X.EOF Then
            sx_codtra = 1466
            sx_subtra = "Abono de Bancos"
        Else
            If chkProtestada.Value = 1 Then
                sx_codtra = 1499
                sx_subtra = "Protesto de Letras"
            Else
                sx_codtra = 1466
                sx_subtra = "Abono de Bancos"
            End If
            Do While Not X.EOF
                If X!ALL_CODTRA = sx_codtra Then
                    X.Edit
                    X!all_flag_ext = "E"
                    X.Update
                End If
                X.MoveNext
            Loop
        End If
        
savetra:
        FORMGEN.manda_numero
        all_llave.AddNew
        all_llave!ALL_NUMOPER = PUB_NUM_OPER_XXX
        all_llave!ALL_CODCIA = LK_CODCIA
        all_llave!ALL_CODTRA = sx_codtra
        all_llave!all_flag_ext = ""
        all_llave!ALL_CODCLIE = sCodClie
        all_llave!ALL_CODART = 0
        all_llave!ALL_IMPORTE_AMORT = Val(txtSaldos.Text)
        all_llave!all_codusu = LK_CODUSU
        all_llave!ALL_FBG = ""
        all_llave!ALL_CODVEN = Val(Right(cbovendedor.Text, 6))
        all_llave!ALL_IMPORTE = Val(txtImporte.Text)
        all_llave!ALL_IMPORTE_DOLL = 0
        all_llave!ALL_NUMDOC = sx_numdoc
        all_llave!ALL_CP = sCP
        all_llave!ALL_TIPDOC = sx_tipdoc
        all_llave!all_numfac_c = 0
        all_llave!all_numser_c = 0
        all_llave!all_codban = Val(Right(cboBanco.Text, 6)) 'banco
        all_llave!all_concepto = Left(txtCodUniBko.Text, 30)
        all_llave!all_chenum = 0
        all_llave!ALL_FECHA_DIA = LK_FECHA_DIA
        all_llave!ALL_FECHA_SUNAT = txtFechaEmi.Value
        all_llave!ALL_FECHA_VCTO = txtFechaVcto.Text
        all_llave!ALL_CANTIDAD = 0
        all_llave!ALL_NUMSER = 0
        all_llave!all_numfac = 0
        all_llave!all_neto = 0
        all_llave!ALL_BRUTO = 0
        all_llave!ALL_tipmov = 0
        all_llave!ALL_IMPTO = 0
        all_llave!ALL_flete = 0
        all_llave!ALL_HORA = Now
        all_llave!ALL_DESCTO = 0
        all_llave!ALL_GASTOS = 0
        all_llave!ALL_PRECIO = 0
        all_llave!ALL_MONEDA_CLI = cbomoneda.Text
        all_llave!ALL_moneda_ccm = cbomoneda.Text
        all_llave!ALL_MONEDA_CAJA = cbomoneda.Text
        all_llave!all_SECUENCIA = 0
        all_llave!ALL_SIGNO_CAR = 0
        all_llave!ALL_signo_caja = 0
        all_llave!ALL_SIGNO_CCM = 0
        all_llave!all_sIGNO_ARM = 0
        all_llave!all_chenum = 0
        all_llave!ALL_CHESEC = 0
        all_llave!ALL_CHESER = 0
        all_llave!ALL_SUBTRA = ""
        all_llave!ALL_TIPO_BLOQ_ACT = 0
        all_llave!ALL_TIPO_BLOQ_ANT = 0
        all_llave!all_codtra_ext = 0
        all_llave!ALL_TIPO_CAMBIO = Val(txtCambio.Text)
        all_llave!ALL_RUC = 0
        all_llave!ALL_NUMOPER2 = PUB_NUM_OPER_XXX
        all_llave.Update

    End If
    CN.Execute "Commit Transaction", rdExecDirect
    con_llave.Close
    MsgBox "Los datos se grabaron correctamente", vbInformation, Pub_Titulo
    LoadLetras
    ClearLetras
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical, Pub_Titulo
    con_llave.Close
    CN.Execute "Rollback Transaction", rdExecDirect
End Sub

Private Sub cmdGrabar1_Click()
Dim I As Integer
Dim sDias As Integer
Dim sNumOper As Long

    pub_cadena = "SELECT * FROM CONTROLL"
    CN.Execute "Begin Transaction", rdExecDirect
    Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
    
    For I = 1 To lvwDocumentos.ListItems.count
        sNumOper = Val(lvwDocumentos.ListItems(I).ListSubItems(8).Text)
        archi = "UPDATE ALLOG SET ALL_MONEDA_CAJA = '" & Left(cbomoneda.Text, 1) & "', ALL_TIPO_CAMBIO=" & Val(txtCambio.Text) & ", ALL_FECHA_SUNAT = '" & txtFecEmiGrl.Value & "', ALL_CODVEN=" & Val(Right(cbovendedor.Text, 6)) & " WHERE ALL_CODCIA = '" & LK_CODCIA & "' AND ALL_FECHA_DIA='" & sFechaDia & "' AND ALL_NumOper = " & sNumOper
        CN.Execute archi
    Next I
    
    For I = 1 To lvwLetras(1).ListItems.count
        SQ_OPER = 1
        pu_cp = sCP
        pu_codclie = sCodClie
        pu_codcia = LK_CODCIA
        PUB_TIPDOC = Left(lvwLetras(1).ListItems(I).Text, 2)
        PUB_FECHA = lvwLetras(0).ListItems(I).ListSubItems(8).Text  ' X("ALL_FECHA")
        PUB_NUM_OPER = lvwLetras(1).ListItems(I).ListSubItems(10).Text ' X("ALL_NUMOPER")
        PUB_SERDOC = lvwLetras(1).ListItems(I).ListSubItems(16).Text
        PUB_NUMDOC = lvwLetras(1).ListItems(I).ListSubItems(9).Text
        LEER_CAA_LLAVE
        caa_histo.Requery
        If Not caa_histo.EOF Then
            caa_histo.Edit
            caa_histo("CAA_FECHA_COBRO") = lvwLetras(1).ListItems(I).ListSubItems(1).Text
            sDias = Val(lvwLetras(1).ListItems(I).ListSubItems(2).Text)
            caa_histo("CAA_FECHA_VCTO") = lvwLetras(1).ListItems(I).ListSubItems(3).Text
            caa_histo("CAA_CODVEN") = Val(Right(cbovendedor.Text, 6))
            caa_histo("CAA_TIPO_CAMBIO") = Val(txtCambio.Text)
            caa_histo("CAA_IMPORTE") = Val(lvwLetras(1).ListItems(I).ListSubItems(8).Text)
            'caa_histo ("CAA_SALDO") 'NO TOMAR EN CUENTA
            caa_histo("CAA_TOTAL") = Val(lvwLetras(1).ListItems(I).ListSubItems(7).Text)
            caa_histo("CAA_SALDO_CAR") = Val(lvwLetras(1).ListItems(I).ListSubItems(8).Text)
            caa_histo.Update
        End If
        LEER_CAR_LLAVE
        If Not car_llave.EOF Then
            car_llave.Edit
            car_llave("CAR_FECHA_SUNAT") = lvwLetras(1).ListItems(I).ListSubItems(1).Text
            car_llave("CAR_FECHA_VCTO") = lvwLetras(1).ListItems(I).ListSubItems(3).Text
            car_llave("CAR_FECHA_VCTO_ORIG") = lvwLetras(1).ListItems(I).ListSubItems(3).Text
            car_llave("CAR_CODVEN") = Val(Right(cbovendedor.Text, 6))
            car_llave("CAR_COBRADOR") = Val(Right(cbovendedor.Text, 6))
            car_llave("CAR_IMPORTE") = Val(lvwLetras(1).ListItems(I).ListSubItems(8).Text)
            car_llave("CAR_IMP_INI") = Val(lvwLetras(1).ListItems(I).ListSubItems(7).Text)
            car_llave("CAR_MONEDA") = Left(cbomoneda.Text, 1)
            car_llave.Update
        End If
        
        archi = "UPDATE ALLOG SET ALL_MONEDA_CAJA = '" & Left(cbomoneda.Text, 1) & "', ALL_TIPO_CAMBIO=" & Val(txtCambio.Text) & ", ALL_FECHA_SUNAT = '" & txtFecEmiGrl.Value & "', ALL_CODVEN=" & Val(Right(cbovendedor.Text, 6)) & " WHERE ALL_CODCIA='" & LK_CODCIA & "' AND ALL_FECHA_DIA='" & sFechaDia & "' AND ALL_NUMOPER=" & PUB_NUM_OPER
        CN.Execute archi
        
    Next I
    
    CN.Execute "Commit Transaction", rdExecDirect
    con_llave.Close
    MsgBox "Los datos se grabaron correctamente", vbInformation, Pub_Titulo
    
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical, Pub_Titulo
    con_llave.Close
    CN.Execute "Rollback Transaction", rdExecDirect
End Sub

Private Sub CmdImprimir_Click()
Dim NumDocs As String
Dim vbresp As Integer
Dim I As Integer
Dim rmoneda As String
Dim Importe As Currency
Dim STIPDOC As String
Dim sRuta As String

On Error GoTo Handler
    
    If lvwLetras(0).ListItems.count = 0 Then Exit Sub
    
    vbresp = MsgBox("Si desea Imprimir todas las Lstras de click en <Si>" & vbCrLf & "Si solo desea imprimir la Letra seleccionada de click en <NO>" & vbCrLf & "De lo contrario click en <CANCELAR>", vbQuestion + vbYesNoCancel, Pub_Titulo)
    rmoneda = Left(cbomoneda.Text, 1)
    'vbresp = vbYes
    If vbresp = vbYes Then
        For I = 1 To lvwLetras(0).ListItems.count
            STIPDOC = Left(lvwLetras(0).ListItems(I).Text, 2)
            Importe = lvwLetras(0).ListItems(I).ListSubItems(5).Text
            NumDocs = lvwLetras(0).ListItems(I).ListSubItems(6).Text
            Reportes.Formulas(2) = "SON='SON: " & CONVER_LETRAS(Importe, rmoneda) & "'"
            GoSub Imprimir
        Next I
    ElseIf vbresp = vbNo Then
        'Importe = lvwLetras(1).ListItems(RowActual1).ListSubItems(7).Text
        'NumDocs = lvwLetras(1).ListItems(RowActual1).ListSubItems(9).Text
        STIPDOC = Left(lvwLetras(1).ListItems(RowActual1).Text, 2)
        Importe = lvwLetras(1).ListItems(RowActual1).ListSubItems(7).Text
        NumDocs = lvwLetras(1).ListItems(RowActual1).ListSubItems(9).Text
        Reportes.Formulas(2) = "SON='SON: " & CONVER_LETRAS(Importe, rmoneda) & "'"
        GoSub Imprimir
    Else
        Exit Sub
    End If
    Exit Sub
Imprimir:
    If chkConyugue.Value = 1 Then
        Reportes.Formulas(1) = "CONYUGUE='" & sConyugue & "'"
    Else
        Reportes.Formulas(1) = ""
    End If
    If OptDocId(0).Value = True Then
        Reportes.Formulas(3) = "DOCUMENTO='" & cli_llave("cli_ruc_esposa") & "'"
    ElseIf OptDocId(1).Value = True Then
        Reportes.Formulas(3) = "DOCUMENTO='" & cli_llave("cli_ruc_esposo") & "'"
    Else
        Reportes.Formulas(3) = ""
    End If
   ' sRuta = Mid(PUB_RUTA_OTRO, 2, Len(PUB_RUTA_OTRO) - 1)
    'sRuta = "C" & sRuta
    sRuta = PUB_RUTA_OTRO
    
    Reportes.Connect = PUB_ODBC
    Reportes.ReportFileName = sRuta & "LETRAS.RPT"
    Reportes.WindowTitle = "Letras de Clientes"
    If chkPrint.Value = 1 Then
        Reportes.Destination = crptToPrinter
    Else
        Reportes.Destination = crptToWindow
    End If
    Reportes.SelectionFormula = " {CARTERA.CAR_TIPDOC} = '" & STIPDOC & "'  AND {CARTERA.CAR_NUMDOC} = " & NumDocs & " AND {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' AND {CARTERA.CAR_CODCLIE} = " & sCodClie & " AND {CARTERA.CAR_CP} = '" & sCP & "'"
    Reportes.Action = 1
    Return
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub

Private Sub cmdImprimir1_Click()
Dim STIPDOC As String
Dim NumDocs As String
Dim I As Integer
Dim sRuta As String


On Error GoTo Handler
    Reportes.Formulas(1) = ""
    Reportes.Formulas(2) = ""
    Reportes.Formulas(3) = ""
    If lvwLetras(0).ListItems.count = 0 Then Exit Sub
    Reportes.Formulas(2) = ""
    NumDocs = ""
    STIPDOC = ""
    For I = 1 To lvwLetras(0).ListItems.count
        NumDocs = NumDocs & " ( {CARTERA.CAR_TIPDOC} = '" & Left(lvwLetras(0).ListItems(I).Text, 2) & "' AND {CARTERA.CAR_NUMDOC} = " & lvwLetras(0).ListItems(I).ListSubItems(6).Text & ") OR "
        STIPDOC = STIPDOC & "'" & Left(lvwLetras(0).ListItems(I).Text, 2) & "',"
    Next I
    NumDocs = Mid(NumDocs, 1, Len(NumDocs) - 3)
    
    sRuta = Mid(PUB_RUTA_OTRO, 2, Len(PUB_RUTA_OTRO) - 1)
    sRuta = "C" & sRuta
    
    Reportes.Connect = PUB_ODBC
    Reportes.ReportFileName = sRuta & "LETRASFAC.RPT"
    If chkPrint.Value = 1 Then
        Reportes.Destination = crptToPrinter
    Else
        Reportes.Destination = crptToWindow
    End If
    'Reportes.SelectionFormula = " {CARTERA.CAR_NUMDOC} IN " & NumDocs & " AND {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "'"
    Reportes.SelectionFormula = "(" & NumDocs & ") AND {CARTERA.CAR_CODCIA} = '" & LK_CODCIA & "' AND {CARTERA.CAR_CODCLIE} = " & sCodClie & " AND {CARTERA.CAR_CP} = '" & sCP & "'"
    Reportes.Action = 1
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub

Private Sub cmdmostrar_Click()
  Screen.MousePointer = 13
    LoadCanjes
    Screen.MousePointer = 0
End Sub

Private Sub CountDias_Change()
    txtdias.Text = CountDias.Value
    txtFechaVcto.Text = Format(DateAdd("d", Val(txtdias.Text), txtFechaEmi.Value), "dd/mm/yyyy")
End Sub

Private Sub Form_DblClick()
Exit Sub
        
        'archi = "SELECT * FROM CARTERA WHERE CAR_cp='C' and CAR_CODCIA='01' AND CAR_TIPDOC IN ('LE','RE','RF','FR') AND CAR_CODUNIBKO<>'' AND CAR_CODTRA=1455 AND CAR_SITUACION <> 'E' AND CAR_FECHA_INGR >'01/01/2005' AND CAR_FECHA_INGR < '31/03/2005' order by CAR_FECHA_INGR"1466
        archi = "SELECT * FROM CARTERA WHERE CAR_cp='C' and CAR_CODCIA='01' AND CAR_TIPDOC IN ('LE','RE','RF','FR') AND CAR_CODUNIBKO<>'' AND CAR_CODTRA=1455 AND CAR_SITUACION <> 'E' AND CAR_FECHA_INGR >'01/01/2005' AND CAR_FECHA_INGR < '31/03/2005' and car_placa='1' order by CAR_FECHA_INGR" '1499
        Set PSX = CN.CreateQuery("", archi)
        Set X = PSX.OpenResultset(rdOpenForwardOnly)
        X.Requery
        
    Do While Not X.EOF
        SQ_OPER = 2
        PUB_FECHA = X!CAR_FECHA_INGR
        pu_codcia = LK_CODCIA
        LEER_ALL_LLAVE
        If all_menor.EOF = False Then
           PUB_NUM_OPER_XXX = all_menor!ALL_NUMOPER
        Else
           PUB_NUM_OPER_XXX = 0
        End If
        PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
        all_llave.AddNew
        If Val(par_llave!PAR_FLAG_EXP) = 9 Then
          all_llave!ALL_NUMOPER = PUB_NUM_OPER_XXX + 5000
        Else
          all_llave!ALL_NUMOPER = PUB_NUM_OPER_XXX
        End If
        all_llave!ALL_CODCIA = LK_CODCIA
        all_llave!ALL_CODTRA = 1499
        all_llave!all_flag_ext = ""
        all_llave!ALL_CODCLIE = X!CAR_codclie
        all_llave!ALL_CODART = 0
        all_llave!ALL_IMPORTE_AMORT = Val(X!CAR_IMP_INI)
        all_llave!all_codusu = LK_CODUSU
        all_llave!ALL_FBG = ""
        all_llave!ALL_CODVEN = X!CAR_codven
        all_llave!ALL_IMPORTE = X!CAR_IMP_INI
        all_llave!ALL_IMPORTE_DOLL = 0
        all_llave!ALL_NUMDOC = X!car_NUMDOC
        all_llave!ALL_CP = X!CAR_CP
        all_llave!ALL_TIPDOC = X!CAR_TIPDOC
        all_llave!all_numfac_c = 0
        all_llave!all_numser_c = 0
        all_llave!all_codban = X!CAR_codban
        all_llave!all_concepto = X!CAR_CODUNIBKO
        all_llave!all_chenum = 0
        all_llave!ALL_FECHA_DIA = X!CAR_FECHA_INGR
        all_llave!ALL_FECHA_SUNAT = X!car_fecha_sunat
        all_llave!ALL_FECHA_VCTO = X!car_fecha_vcto
        all_llave!ALL_CANTIDAD = 0
        all_llave!ALL_NUMSER = 0
        all_llave!all_numfac = 0
        all_llave!all_neto = 0
        all_llave!ALL_BRUTO = 0
        all_llave!ALL_tipmov = 0
        all_llave!ALL_IMPTO = 0
        all_llave!ALL_flete = 0
        all_llave!ALL_HORA = Now
        all_llave!ALL_DESCTO = 0
        all_llave!ALL_GASTOS = 0
        all_llave!ALL_PRECIO = 0
        all_llave!ALL_MONEDA_CLI = X!CAR_MONEDA
        all_llave!ALL_moneda_ccm = X!CAR_MONEDA
        all_llave!ALL_MONEDA_CAJA = X!CAR_MONEDA
        all_llave!all_SECUENCIA = 0
        all_llave!ALL_SIGNO_CAR = 0
        all_llave!ALL_signo_caja = 0
        all_llave!ALL_SIGNO_CCM = 0
        all_llave!all_sIGNO_ARM = 0
        all_llave!all_chenum = 0
        all_llave!ALL_CHESEC = 0
        all_llave!ALL_CHESER = 0
        all_llave!ALL_SUBTRA = ""
        all_llave!ALL_TIPO_BLOQ_ACT = 0
        all_llave!ALL_TIPO_BLOQ_ANT = 0
        all_llave!all_codtra_ext = 0
        all_llave!ALL_TIPO_CAMBIO = 0
        all_llave!ALL_RUC = 0
        all_llave!ALL_NUMOPER2 = PUB_NUM_OPER_XXX
        all_llave.Update
        X.MoveNext
    Loop
End Sub

Private Sub Form_Load()

    FormatLvw
    txtfecha1.Value = LK_FECHA_DIA
    txtfecha2.Value = LK_FECHA_DIA
    fraLetra.Top = 20
    fraLetra.Left = 20
    
    VARTEMP = LK_EMP_PTO
    LK_EMP_PTO = "A"
    LlenadoCbo cboSituacion, 160
    LK_EMP_PTO = VARTEMP
    
    LoadVendedores
    
    archi = "SELECT * FROM CCMAEST WHERE CCM_CODCIA='" & LK_CODCIA & "' "
    Set PSX = CN.CreateQuery("", archi)
    Set X = PSX.OpenResultset(rdOpenForwardOnly)
    X.Requery
    Do While Not X.EOF
        cboBanco.AddItem X("CCM_NOMBRE") & Space(30) & X("CCM_CODBAN")
        X.MoveNext
    Loop
    cmbCP.ListIndex = 0
End Sub

Private Sub LoadCanjes()
Dim tNumOper2 As Integer
Dim iCount As Integer
Dim Icount2 As Integer
Dim dImporte As Double
    archi = "SELECT DISTINCT ALL_FECHA_DIA,ALL_FECHA_SUNAT,ALL_MONEDA_CLI,ALL_IMPORTE,ALL_NUMOPER2,ALL_CODCLIE,ALL_tipo_cambio, ALL_codven, ALL_NUM_RECIBO FROM ALLOG WHERE ALL_CP = '" & sCP & "' AND ALL_TIPDOC in ('LE','RE','RF','FR') AND ALL_CODTRA=1455 AND ALL_FLAG_EXT<>'E' AND ALL_CODCIA='" & LK_CODCIA & "' AND ALL_FECHA_DIA >= '" & txtfecha1.Value & "' AND ALL_FECHA_DIA <= '" & txtfecha2.Value & "' and all_signo_car=1"
    If chkClientes.Value = 0 Then
        archi = archi & " AND ALL_CODCLIE = " & Val(txtCP.TEXTO)
    End If
    If Len(Trim(b_numfac.Text)) <> 0 Then
        archi = archi & " AND (ALL_CONCEPTO LIKE  '%" & Trim(b_numfac.Text) & "%')"
    End If
    archi = archi & " ORDER BY all_fecha_dia, ALL_NUMOPER2"
    Set PSX = CN.CreateQuery("", archi)
    Set X = PSX.OpenResultset(rdOpenForwardOnly)
    X.Requery
    lvwCanjes.ListItems.Clear
    lvwDocumentos.ListItems.Clear
    lvwLetras(0).ListItems.Clear
    Do While Not X.EOF
        iCount = iCount + 1
'        If tNumOper2 = X("ALL_NUMOPER2") And iCount <> 1 Then
'            If X("ALL_MONEDA_CLI") = "D" And X("ALL_MONEDA_CAJA") = "S" Then
'                dImporte = dImporte + X("ALL_IMPORTE_AMORT") * X("ALL_TIPO_CAMBIO")
'            Else
'                dImporte = dImporte + X("ALL_IMPORTE_AMORT")
'            End If
'            GoTo OTRO
'        End If
'        Icount2 = Icount2 + 1
'        If Icount2 > 1 Then
'            lvwCanjes.ListItems(Icount2 - 1).ListSubItems(3).Text = Format(dImporte, "0.00")
'            dImporte = 0
'        End If
'        If X("ALL_MONEDA_CLI") = "D" And X("ALL_MONEDA_CAJA") = "S" Then
'            dImporte = dImporte + X("ALL_IMPORTE_AMORT") * X("ALL_TIPO_CAMBIO")
'        Else
'            dImporte = dImporte + X("ALL_IMPORTE_AMORT")
'        End If
        
        Set loListItem = lvwCanjes.ListItems.Add(, "n" & X("ALL_NUMOPER2") & X("ALL_FECHA_DIA"), X("ALL_FECHA_SUNAT"))
        SQ_OPER = 1
        pu_cp = sCP
        pu_codclie = X("ALL_CODCLIE")
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If Not cli_llave.EOF Then
            loListItem.ListSubItems.Add Key:="Cliente", Text:=Trim(cli_llave("CLI_NOMBRE"))
        Else
            loListItem.ListSubItems.Add Key:="Cliente", Text:=""
        End If
        loListItem.ListSubItems.Add Key:="Moneda", Text:=X("ALL_MONEDA_CLI")
        loListItem.ListSubItems.Add Key:="------", Text:=0 'Format(X("ALL_IMPORTE"), "0.00")
        loListItem.ListSubItems.Add Key:="NumDoc", Text:="0" 'X("ALL_NUMDOC")
        loListItem.ListSubItems.Add Key:="NumOper", Text:=X("ALL_NUMOPER2")
        loListItem.ListSubItems.Add Key:="CodClie", Text:=X("ALL_CODCLIE")
        loListItem.ListSubItems.Add Key:="TCambio", Text:=X("ALL_tipo_cambio")
        loListItem.ListSubItems.Add Key:="CodVen", Text:=X("ALL_codven")
        loListItem.ListSubItems.Add Key:="FechaDia", Text:=X("ALL_FECHA_DIA")
        tNumOper2 = X("ALL_NUMOPER2")
OTRO:
        X.MoveNext
    Loop
    If Icount2 > 1 Then
        lvwCanjes.ListItems(Icount2).ListSubItems(3).Text = Format(dImporte, "0.00")
    ElseIf Icount2 = 1 Then
        lvwCanjes.ListItems(Icount2).ListSubItems(3).Text = Format(dImporte, "0.00")
    End If
    
End Sub
Private Sub LoadDocumentos(ByVal sNumOper As Integer)
Dim sw_totcanje As Currency
Dim sNumFac As String
Dim sNumSer As String
Dim sFBG As String
Dim sImporte As Double
Dim sMoneda As String
 Print sFechaDia
    sDocumentos = ""
    archi = "SELECT * FROM ALLOG WHERE ALL_FECHA_DIA = '" & Format(sFechaDia, "dd/mm/yyyy") & "' AND ALL_CP = '" & sCP & "' and all_codclie=" & sCodClie & " AND (ALL_TIPDOC='FA' OR ALL_TIPDOC='CC' OR ALL_TIPDOC='LE' OR ALL_TIPDOC='RE' OR ALL_TIPDOC='RF') AND ALL_CODTRA=1455 AND ALL_FLAG_EXT<>'E' and ALL_SIGNO_CAR = -1 AND ALL_CODCIA='" & LK_CODCIA & "' and ALL_NUMOPER2=" & sNumOper

    Set PSX = CN.CreateQuery("", archi)
    Set X = PSX.OpenResultset(rdOpenForwardOnly)
    X.Requery
    lvwDocumentos.ListItems.Clear
    lvwLetras(0).ListItems.Clear
    sw_totcanje = 0
    Do While Not X.EOF
        SQ_OPER = 1
        pu_cp = sCP
        pu_codclie = X("all_codclie")
        pu_codcia = LK_CODCIA
        PUB_SERDOC = X("all_serdoc")
        PUB_NUMDOC = X("all_numdoc")
        PUB_TIPDOC = X("all_tipdoc")
        LEER_CAR_LLAVE
        If Not car_llave.EOF Then
            sw_totcanje = sw_totcanje + Val(car_llave("car_imp_ini"))
            sImporte = car_llave("car_imp_ini")
            sMoneda = car_llave("car_moneda")
            If car_llave("CAR_TIPDOC") = "LE" Or car_llave("CAR_TIPDOC") = "RE" Then
                sFBG = car_llave("CAR_TIPDOC")
                sNumFac = car_llave("car_num_CHEQUE")
                sNumSer = " "
            ElseIf car_llave("CAR_CP") = "C" Then
            ElseIf car_llave("CAR_CP") = "P" Then
                sFBG = car_llave("car_fbg")
                sNumFac = car_llave("car_numfac_C")
                sNumSer = car_llave("car_numser_C")
            End If
        End If
        Set loListItem = lvwDocumentos.ListItems.Add(, "n" & X("all_NUMOPER") & X("all_FECHA_dia"), sFBG & " / " & sNumSer & " - " & sNumFac & " /Int " & car_llave!car_NUMSER & "-" & car_llave!car_NUMFAC)
        sDocumentos = sDocumentos & sFBG & " / " & sNumSer & " - " & sNumFac & "  "
        loListItem.ListSubItems.Add Key:="Fecha", Text:=car_llave("car_FECHA_SUNAT")  ' X("all_FECHA_sunat")
        loListItem.ListSubItems.Add Key:="TipDoc", Text:=X("all_TIPDOC")
        loListItem.ListSubItems.Add Key:="Moneda", Text:=sMoneda
        loListItem.ListSubItems.Add Key:="Importe", Text:=Format(X("all_IMPORTE_amort"), "0.00")
        loListItem.ListSubItems.Add Key:="ImporteIni", Text:=Format(sImporte, "0.00")
        loListItem.ListSubItems.Add Key:="CodClie", Text:=X("all_CODCLIE")
        loListItem.ListSubItems.Add Key:="NumDoc", Text:=X("all_NUMDOC")
        loListItem.ListSubItems.Add Key:="NumOper", Text:=X("all_NUMOper")
        X.MoveNext
    Loop
    lbltotcanje.Caption = Format(sw_totcanje, "##,##0.00")
End Sub
Private Sub LoadLetras()
Dim iRow As Integer
Dim sSaldoLetra As Double
Dim sw_salcanje As Currency
    'sTotalDocumento = 0
    sSaldoDocumento = 0
    archi = "SELECT * FROM ALLOG WHERE ALL_CP = '" & sCP & "' and all_codclie=" & sCodClie & " and ALL_TIPDOC IN('LE','RE','RF','FR')  AND ALL_FLAG_EXT<>'E' AND ALL_SIGNO_CAR=1 AND ALL_NUMOPER2=" & WS_NUM_OPER2 & " AND ALL_CODCIA='" & LK_CODCIA & "'  and all_fecha_dia = '" & sFechaDia & "'"
    Set PSX = CN.CreateQuery("", archi)
    Set X = PSX.OpenResultset(rdOpenForwardOnly)
    X.Requery
    lvwLetras(0).ListItems.Clear
    lvwLetras(1).ListItems.Clear
    lvwLetras(2).ListItems.Clear
    sw_salcanje = 0
    Do While Not X.EOF
        iRow = iRow + 1
        SQ_OPER = 1
        pu_cp = sCP
        pu_codclie = X("all_codclie")
        pu_codcia = LK_CODCIA
        PUB_SERDOC = X("all_serdoc")
        PUB_TIPDOC = X("all_TIPDOC")
        PUB_FECHA = X("ALL_FECHA_DIA")
        PUB_NUM_OPER = X("ALL_NUMOPER")
        LEER_CAA_LLAVE
        
        If Not caa_histo.EOF Then
            Set loListItem = lvwLetras(0).ListItems.Add(, "n" & caa_histo("caa_NUMDOC"), caa_histo("caa_TIPDOC") & " " & caa_histo("caa_NUM_CHEQUE") & "        .")
            Set loListItem1 = lvwLetras(1).ListItems.Add(, "n" & caa_histo("caa_NUMDOC"), caa_histo("caa_TIPDOC") & " " & caa_histo("caa_NUM_CHEQUE") & "        . ")
            Set loListItem2 = lvwLetras(2).ListItems.Add(, "n" & caa_histo("caa_NUMDOC"), caa_histo("caa_TIPDOC") & " " & caa_histo("caa_NUM_CHEQUE") & "        .")
            
            loListItem1.ListSubItems.Add Key:="Fecha", Text:=caa_histo("CAA_FECHA_COBRO")
            loListItem1.ListSubItems.Add Key:="Dias", Text:=DateDiff("d", caa_histo("CAA_FECHA_COBRO"), caa_histo("caa_fecha_VCTO"))
            
            PUB_SERDOC = caa_histo("CAA_SERDOC")
            PUB_NUMDOC = caa_histo("CAA_NUMDOC")
            LEER_CAR_LLAVE
            loListItem.ListSubItems.Add Key:="Fecha", Text:=car_llave("CAR_FECHA_VCTO_ORIG")
            loListItem1.ListSubItems.Add Key:="FechaVcto", Text:=car_llave("CAR_FECHA_VCTO_ORIG")
            loListItem2.ListSubItems.Add Key:="Fecha", Text:=car_llave("CAR_FECHA_VCTO_ORIG")
            If CDate(car_llave("CAR_FECHA_VCTO_ORIG")) < CDate(LK_FECHA_DIA) And car_llave("CAR_IMPORTE") <> 0 Then
                'LETRAS VENCIDAS QUE NO SE TERMINA DE AMORTIZAR
                lvwLetras(0).ListItems(iRow).ForeColor = QBColor(12)
                lvwLetras(0).ListItems(iRow).Bold = True
                lvwLetras(0).ListItems(iRow).ListSubItems(1).ForeColor = QBColor(12)
                lvwLetras(0).ListItems(iRow).ListSubItems(1).Bold = True
            ElseIf car_llave("car_IMPORTE") = 0 Then
                'LESTRAS CANCELADAS
                lvwLetras(0).ListItems(iRow).ForeColor = QBColor(9)
                lvwLetras(0).ListItems(iRow).Bold = True
                lvwLetras(0).ListItems(iRow).ListSubItems(1).ForeColor = QBColor(9)
                lvwLetras(0).ListItems(iRow).ListSubItems(1).Bold = True
            Else
                'OTRO CASO
                lvwLetras(0).ListItems(iRow).ForeColor = QBColor(2) '2 verde , 9 azul
                lvwLetras(0).ListItems(iRow).Bold = True
                lvwLetras(0).ListItems(iRow).ListSubItems(1).ForeColor = QBColor(2)
                lvwLetras(0).ListItems(iRow).ListSubItems(1).Bold = True
            End If
            
            
            If Not car_llave.EOF Then
                loListItem.ListSubItems.Add Key:="Moneda", Text:=car_llave("CAR_MONEDA")
                loListItem1.ListSubItems.Add Key:="Moneda", Text:=car_llave("CAR_MONEDA")
                loListItem2.ListSubItems.Add Key:="Moneda", Text:=car_llave("CAR_MONEDA")
                sSaldoDocumento = sSaldoDocumento + car_llave("CAR_IMPORTE")
                sw_salcanje = sw_salcanje + Val(car_llave("CAR_IMPORTE"))
                sSaldoLetra = car_llave("CAR_IMPORTE")
                SQ_OPER = 1
                pu_codcia = LK_CODCIA
                PUB_CODBAN = car_llave("CAR_CODBAN")
                LEER_CCM_LLAVE
                If ccm_llave.EOF Then
                loListItem.ListSubItems.Add Key:="Banco", Text:=""
                loListItem1.ListSubItems.Add Key:="Banco", Text:=""
                loListItem2.ListSubItems.Add Key:="Banco", Text:=""
                Else
                loListItem.ListSubItems.Add Key:="Banco", Text:=ccm_llave("CCM_NOMBRE")
                loListItem1.ListSubItems.Add Key:="Banco", Text:=ccm_llave("CCM_NOMBRE")
                loListItem2.ListSubItems.Add Key:="Banco", Text:=ccm_llave("CCM_NOMBRE")
                End If
                SQ_OPER = 1
                PUB_TIPREG = 160
                PUB_NUMTAB = Val(Nulo_Valors(car_llave("CAR_VOUCHER")))
                PUB_CODCIA = "00"
                LEER_TAB_LLAVE
                If Not tab_llave.EOF Then
                loListItem.ListSubItems.Add Key:="Situacion", Text:=Trim(tab_llave("TAB_NOMLARGO"))
                loListItem1.ListSubItems.Add Key:="Situacion", Text:=Trim(tab_llave("TAB_NOMLARGO"))
                loListItem2.ListSubItems.Add Key:="Situacion", Text:=Trim(tab_llave("TAB_NOMLARGO"))
                Else
                loListItem.ListSubItems.Add Key:="Situacion", Text:=""
                loListItem1.ListSubItems.Add Key:="Situacion", Text:=""
                loListItem2.ListSubItems.Add Key:="Situacion", Text:=""
                End If
            End If
            loListItem.ListSubItems.Add Key:="Importe", Text:=caa_histo("caa_IMPORTE")
            loListItem1.ListSubItems.Add Key:="Importe", Text:=caa_histo("caa_IMPORTE")
            loListItem1.ListSubItems.Add Key:="ImporteIni", Text:=Format(sSaldoLetra, "#0.00")
            loListItem2.ListSubItems.Add Key:="Importe", Text:=caa_histo("caa_IMPORTE")
            loListItem.ListSubItems.Add Key:="NumDoc", Text:=caa_histo("caa_numdoc")
            loListItem1.ListSubItems.Add Key:="NumDoc", Text:=caa_histo("caa_numdoc")
            loListItem2.ListSubItems.Add Key:="NumDoc", Text:=caa_histo("caa_numdoc")
            loListItem.ListSubItems.Add Key:="NumOper", Text:=X("all_numoper")
            loListItem1.ListSubItems.Add Key:="NumOper", Text:=X("all_numoper")
            loListItem2.ListSubItems.Add Key:="NumOper", Text:=X("all_numoper")
            loListItem.ListSubItems.Add Key:="FechaDia", Text:=X("all_fecha_dia")
            loListItem1.ListSubItems.Add Key:="FechaDia", Text:=X("all_fecha_dia")
            loListItem2.ListSubItems.Add Key:="FechaDia", Text:=X("all_fecha_dia")
            loListItem1.ListSubItems.Add Key:="CAA_SALDO_CAR", Text:=caa_histo("caa_SALDO_CAR")
            loListItem1.ListSubItems.Add Key:="caa_TOTAL", Text:=caa_histo("caa_TOTAL")
            loListItem1.ListSubItems.Add Key:="CodUniBko", Text:=0 ' Nulo_Valors(car_llave("CAR_CODUNIBKO"))
            loListItem1.ListSubItems.Add Key:="ImporteOriginal", Text:=caa_histo("caa_IMPORTE")
            loListItem.ListSubItems.Add Key:="FechaEmi", Text:=X("all_fecha_dia")
            'OJO
            loListItem1.ListSubItems.Add Key:="SerDoc", Text:=X("all_serdoc")
            loListItem2.ListSubItems.Add Key:="SerDoc", Text:=X("all_serdoc")
            loListItem.ListSubItems.Add Key:="SerDoc", Text:=X("all_serdoc")
            'CAR_IMPORTE=CAA_IMPORTE
        End If
        X.MoveNext
    Loop
    lblsalcanje.Caption = Format(sw_salcanje, "#,##0.00")
    CalculaSubTotal
End Sub
Private Sub FormatLvw()

    lvwCanjes.ColumnHeaders.Add , , "Fecha", 1200
    lvwCanjes.ColumnHeaders.Add , , "Cliente", 3500
    lvwCanjes.ColumnHeaders.Add , , "Moneda", 800
    lvwCanjes.ColumnHeaders.Add , , "Importe", 0
    lvwCanjes.ColumnHeaders.Add , , "NumDoc", 0
    lvwCanjes.ColumnHeaders.Add , , "NumOper2", 0
    lvwCanjes.ColumnHeaders.Add , , "TipoCambio", 0
    lvwCanjes.ColumnHeaders.Add , , "CodVen", 0
    lvwCanjes.ColumnHeaders.Add , , "FechaDia", 0
    
    lvwDocumentos.ColumnHeaders.Add , , "Documento", 1800
    lvwDocumentos.ColumnHeaders.Add , , "Fecha", 1200
    lvwDocumentos.ColumnHeaders.Add , , "Tipo", 800
    lvwDocumentos.ColumnHeaders.Add , , "Moneda", 800
    lvwDocumentos.ColumnHeaders.Add , , "Imp. Canj.", 1500
    lvwDocumentos.ColumnHeaders.Add , , "Imp. Orig", 1500
    lvwDocumentos.ColumnHeaders.Add , , "Codclie", 0
    lvwDocumentos.ColumnHeaders.Add , , "NumDoc", 0
    lvwDocumentos.ColumnHeaders.Add , , "NumOper", 0
    lvwDocumentos.ColumnHeaders.Add , , "SerDoc", 0
    
    lvwLetras(0).ColumnHeaders.Add , , "Documento", 2000
    lvwLetras(0).ColumnHeaders.Add , , "Vencimiento", 1200
    lvwLetras(0).ColumnHeaders.Add , , "Moneda", 800
    lvwLetras(0).ColumnHeaders.Add , , "Banck", 2200
    lvwLetras(0).ColumnHeaders.Add , , "Situacion", 1200
    lvwLetras(0).ColumnHeaders.Add , , "Importe", 1500
    lvwLetras(0).ColumnHeaders.Add , , "NumDoc", 0
    lvwLetras(0).ColumnHeaders.Add , , "NumOper", 0
    lvwLetras(0).ColumnHeaders.Add , , "FechaDia", 0
    lvwLetras(0).ColumnHeaders.Add , , "SerDoc", 0
        
    lvwLetras(1).ColumnHeaders.Add , , "Documento", 2000
    lvwLetras(1).ColumnHeaders.Add , , "FechaEmi", 1200
    lvwLetras(1).ColumnHeaders.Add , , "Dias", 800
    lvwLetras(1).ColumnHeaders.Add , , "Vencimiento", 1200
    lvwLetras(1).ColumnHeaders.Add , , "Moneda", 800
    lvwLetras(1).ColumnHeaders.Add , , "Banck", 2200
    lvwLetras(1).ColumnHeaders.Add , , "Situacion", 1200
    lvwLetras(1).ColumnHeaders.Add , , "Importe", 1500
    lvwLetras(1).ColumnHeaders.Add , , "Saldo", 1500
    lvwLetras(1).ColumnHeaders.Add , , "NumDoc", 0
    lvwLetras(1).ColumnHeaders.Add , , "NumOper", 0
    lvwLetras(1).ColumnHeaders.Add , , "FechaDia", 0
    lvwLetras(1).ColumnHeaders.Add , , "SALDOCAR", 0
    lvwLetras(1).ColumnHeaders.Add , , "TOTAL", 0
    lvwLetras(1).ColumnHeaders.Add , , "Num-Uni-Banc", 1500
    lvwLetras(1).ColumnHeaders.Add , , "ImporteOriginal", 0
    lvwLetras(1).ColumnHeaders.Add , , "SerDoc", 0
    
    lvwLetras(2).ColumnHeaders.Add , , "Documento", 1200
    lvwLetras(2).ColumnHeaders.Add , , "Vencimiento", 0
    lvwLetras(2).ColumnHeaders.Add , , "Moneda", 800
    lvwLetras(2).ColumnHeaders.Add , , "Banck", 0
    lvwLetras(2).ColumnHeaders.Add , , "Situacion", 0
    lvwLetras(2).ColumnHeaders.Add , , "Importe", 1500
    lvwLetras(2).ColumnHeaders.Add , , "NumDoc", 0
    lvwLetras(2).ColumnHeaders.Add , , "NumOper", 0
    lvwLetras(2).ColumnHeaders.Add , , "FechaDia", 0
    lvwLetras(2).ColumnHeaders.Add , , "FechaEmi", 0
    lvwLetras(2).ColumnHeaders.Add , , "SerDoc", 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmletras = Nothing
End Sub

Private Sub lvwCanjes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwCanjes.Sorted = True
    lvwCanjes.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub lvwCanjes_ItemClick(ByVal Item As MSComctlLib.ListItem)

    WS_NUM_OPER2 = Val(Item.ListSubItems(5).Text)
    sCodClie = Val(Item.ListSubItems(6).Text)
    sFechaDia = Item.ListSubItems(9).Text
    
    Call LoadDocumentos(WS_NUM_OPER2)
    Call LoadLetras
    txtCambio = Item.ListSubItems(7).Text
    BuscaInCbo cbovendedor, Val(Item.ListSubItems(8).Text)
    If Item.ListSubItems(2).Text = "S" Then
        cbomoneda.ListIndex = 0
        MonedaOriginal = "S"
    Else
        cbomoneda.ListIndex = 1
        MonedaOriginal = "D"
    End If
    txtFecEmiGrl = Item.Text
    FlagMoney = 0
    
    sTotalDocumento = Val(Item.ListSubItems(3).Text)
    lblTotal.Caption = Format(sTotalDocumento, "#0.00")
End Sub

Private Sub lvwLetras_DblClick(Index As Integer)
    If Index = 0 Then
        DoEvents
        fraLetra.Visible = True
        fraDocumentos.Visible = False
        FlagMoney = 1
    End If
End Sub

Private Sub lvwLetras_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)

    If Index = 2 Or Index = 0 Then
        RowActual = Item.Index
        lblDocumento.Caption = sDocumentos
        lblMoneda = Item.ListSubItems(2).Text
        If lblMoneda.Caption = "S" Then
            lblMoneda.Caption = "S/."
        ElseIf lblMoneda.Caption = "D" Then
            lblMoneda.Caption = "US $"
        End If
        PUB_FECHA = Item.ListSubItems(8).Text ' X("ALL_FECHA_DIA")
        PUB_NUM_OPER = Item.ListSubItems(7).Text ' X("ALL_NUMOPER")
        sNumDocLetra = Val(Item.ListSubItems(6).Text)
        
        If Index = 0 Then sSerDocLetra = Val(Item.ListSubItems(9).Text)
        If Index = 2 Then sSerDocLetra = Val(Item.ListSubItems(9).Text)
        
        SQ_OPER = 1
        pu_cp = sCP
        pu_codclie = sCodClie
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If Not cli_llave.EOF Then
            lblCliente.Caption = cli_llave("cli_nombre")
            lblRuc.Caption = cli_llave("cli_ruc_esposo")
            lblDireccion.Caption = cli_llave("cli_casa_direc")
            sConyugue = cli_llave("cli_nombre_empresa")
        End If
        SQ_OPER = 1
        pu_cp = sCP
        pu_codclie = sCodClie
        pu_codcia = LK_CODCIA
        PUB_SERDOC = sSerDocLetra
        PUB_NUMDOC = sNumDocLetra
        PUB_TIPDOC = Left(Item.Text, 2)
        LEER_CAR_LLAVE
        If Not car_llave.EOF Then
            txtFechaVcto.Text = Format(car_llave("car_fecha_vcto"), "dd/mm/yyyy")
            txtFechaEmi.Value = car_llave("car_fecha_sunat")
            txtdias.Text = DateDiff("d", car_llave("car_fecha_sunat"), car_llave("car_fecha_vcto"))
            CountDias.Value = IIf(txtdias.Text > 0, txtdias.Text, 1)
            txtImporte.Text = Format(car_llave("car_imp_ini"), "#0.00")
            txtSaldos.Text = Format(car_llave("car_importe"), "#0.00")
            txtAmort.Text = Format(car_llave("car_imp_ini") - car_llave("car_importe"), "#0.00")
            'nuevo control
            BuscaInCbo cboSituacion, Val(Nulo_Valors(car_llave("car_voucher")))   'situacion
            BuscaInCbo cboBanco, Val(car_llave("car_codban"))   'situacion
            txtObservaciones.Text = car_llave("car_concepto")  'observacion
            chkProtestada.Value = Val(Nulo_Valors(car_llave("car_placa"))) 'protestada
            chkAceptada.Value = Val(car_llave("car_precio")) 'aceptada
            txtCodUniBko.Text = Nulo_Valors(car_llave("CAR_CODUNIBKO"))
            'If Not IsNull(car_llave("car_fecha_entrega")) Then
            'If Format(car_llave("car_fecha_entrega"), "dd/mm/yyyy") <= CDate("01/01/1900") Or IsNull(car_llave("car_fecha_entrega")) Then
            'Else
            '    txtFechaEntrega.Text = car_llave("car_fecha_entrega") 'fechaentrega
            'End If
            'If Format(car_llave("car_fecha_devo"), "dd/mm/yyyy") <= CDate("01/01/1900") Or IsNull(car_llave("car_fecha_devo")) Then
            'Else
            '    txtFechaDevolucion.Text = car_llave("car_fecha_devo") 'fechaentrega
            'End If
   
            lblSaldo.Caption = Format(sSaldoDocumento, "#0.00")
            sNumOper1 = car_llave("car_numoper")
        End If
    ElseIf Index = 1 Then
        RowActual1 = Item.Index
    End If

End Sub

'Private Sub lvwDocumentos_ItemClick(ByVal Item As MSComctlLib.ListItem)
'Dim NumDoc As Integer
'
'    NumDoc = Val(Item.ListSubItems(6).Text)
'    If NumDoc = 0 Then Exit Sub
'    Call LoadLetras(NumDoc, sCodClie)
'    lblDocumento.Caption = Item.Text
'End Sub


Private Sub txtCambio_KeyPress(KeyAscii As Integer)
    KeyAscii = vNumeric(KeyAscii)
End Sub

'===========================================================================
'============Para Presentacion de Datos de  Cliente o Proveedor=============
Private Sub txtcp_Cancel()
    txtCP.TEXTO = ""
    lblCP = ""
End Sub
Private Sub txtcp_GetRegistros(ByVal oKeyFind As Variant)
Dim sSql As String
    sSql = "SELECT 'Razon Social de la Empresa'=CLI_NOMBRE ,'Codigo'=CLI_CODCLIE FROM Clientes WHERE Cli_Codcia= '" & LK_CODCIA & "' AND Cli_CP = '" & sCP & "' AND Cli_Nombre LIKE '" & oKeyFind & "%' ORDER BY Cli_Nombre"
    txtCP.TypeFind = NameField
    txtCP.SetRecordset = OpenSQLForwardOnly(sSql)
End Sub
Private Sub txtcp_GotFocus()
    txtCP.ZOrder 0
End Sub
Private Sub txtcp_ShowData(ByVal oKey As Variant)
    SQ_OPER = 1
    pu_cp = sCP
    pu_codclie = Val(oKey)
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If Not cli_llave.EOF Then
        'Call FormatLblDato(txtCP, lblCP)
        lblCP.Caption = Trim(UCase(cli_llave("Cli_Nombre"))) & " - RUC: " & cli_llave("CLI_RUC_ESPOSO")
        'Call NextCampo(txtCP.TabIndex, Me)
        PUB_CODCLIE = oKey
        PUB_RUC = cli_llave("CLI_RUC_ESPOSO")
    End If
End Sub

Private Function Consistencias() As Boolean
    Consistencias = True
    If txtFechaEntrega.Text <> "__/__/____" Then
        If Not IsDate(txtFechaEntrega.Text) Then
            Consistencias = False
            MsgBox "Fecha Incorrecta", vbInformation, Pub_Titulo
            txtFechaEntrega.SetFocus
            Exit Function
        End If
    End If
    
    If txtFechaDevolucion.Text <> "__/__/____" Then
        If Not IsDate(txtFechaDevolucion.Text) Then
            Consistencias = False
            MsgBox "Fecha Incorrecta", vbInformation, Pub_Titulo
            txtFechaDevolucion.SetFocus
            Exit Function
        End If
    End If
End Function

Private Sub LoadVendedores()
    archi = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = '" & LK_CODCIA & "'"
    Set PSX = CN.CreateQuery("", archi)
    Set X = PSX.OpenResultset(rdOpenForwardOnly)
    X.Requery
    cbovendedor.Clear
    Do While Not X.EOF
        cbovendedor.AddItem X("VeM_Nombre") & Space(80) & X("veM_codven")
        X.MoveNext
    Loop
End Sub

Private Sub txtFecEmiGrl_Change()
Dim I As Integer
Dim sDias As Integer
    For I = 1 To lvwLetras(1).ListItems.count
        lvwLetras(1).ListItems(I).ListSubItems(1).Text = txtFecEmiGrl.Value
        sDias = Val(lvwLetras(1).ListItems(I).ListSubItems(2).Text)
        lvwLetras(1).ListItems(I).ListSubItems(3).Text = DateAdd("d", sDias, txtFecEmiGrl.Value)
    Next I
End Sub

Private Sub txtFechaDevolucion1_DateDblClick(ByVal DateDblClicked As Date)
    txtFechaDevolucion.Text = txtFechaDevolucion1.Value
    txtFechaDevolucion1.Visible = False
End Sub

Private Sub txtFechaDevolucion1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then txtFechaDevolucion1.Visible = False
End Sub

Private Sub txtFechaEmi_Change()
    txtFechaVcto.Text = DateAdd("d", Val(txtdias.Text), txtFechaEmi.Value)
End Sub

Private Sub txtFechaEntrega1_DateDblClick(ByVal DateDblClicked As Date)
    txtFechaEntrega.Text = txtFechaEntrega1.Value
    txtFechaEntrega1.Visible = False
End Sub

Private Sub txtFechaEntrega1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then txtFechaEntrega1.Visible = False
End Sub

Private Sub CalculaSubTotal()
Dim I As Integer
Dim sSubTotal As Double
    For I = 1 To lvwLetras(2).ListItems.count
        sSubTotal = sSubTotal + lvwLetras(2).ListItems(I).ListSubItems(5).Text
    Next I
    txtSubTotal.Text = Format(sSubTotal, "0.00")
End Sub

Private Sub txtImporte_Change()
    txtSaldos.Text = Format(Val(txtImporte.Text) - Val(txtAmort.Text), "0.00")
End Sub

Private Sub cmdDistImpLet_Click()
Dim I As Integer
Dim TotalCanje As Double
Dim TotalCanjeTmp As Double

    For I = 1 To lvwLetras(2).ListItems.count
        If I <> RowActual Then
            TotalCanjeTmp = TotalCanjeTmp + Val(lvwLetras(2).ListItems(I).ListSubItems(5).Text)
        End If
    Next I
    txtImporte.Text = Format(sTotalDocumento - TotalCanjeTmp, "0.00")
    lvwLetras(2).ListItems(RowActual).ListSubItems(5).Text = Format(sTotalDocumento - TotalCanjeTmp, "0.00")
    CalculaSubTotal
End Sub
Private Sub ClearLetras()
    txtFechaVcto.Text = "__/__/____"
    txtdias.Text = 0
    CountDias.Value = 1
    txtImporte.Text = "0.00"
    txtSaldos.Text = "0.00"
    txtAmort.Text = "0.00"
    cboSituacion.ListIndex = -1
    cboBanco.ListIndex = -1
    txtObservaciones.Text = ""
    chkProtestada.Value = 0
    chkAceptada.Value = 0
    txtFechaEntrega.Text = "__/__/____"
    txtFechaDevolucion.Text = "__/__/____"
    txtCodUniBko.Text = ""
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Frmcospro 
   Caption         =   "Tipo de Cambio y Costeo de Articulos"
   ClientHeight    =   7455
   ClientLeft      =   2565
   ClientTop       =   1920
   ClientWidth     =   11880
   Icon            =   "FrmCostos.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11880
   Begin VB.Frame fradoccompra 
      Caption         =   "Estados de Documentos de Compra :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   36
      Top             =   3600
      Visible         =   0   'False
      Width           =   11895
      Begin VB.CommandButton cmdbp 
         Caption         =   "Buscar Prod."
         Height          =   435
         Left            =   5520
         TabIndex        =   46
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdcanceldoc 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10200
         Picture         =   "FrmCostos.frx":0442
         TabIndex        =   44
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton CmdEstado 
         Caption         =   "Estado_________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8760
         Picture         =   "FrmCostos.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton CmdCosteoDoc 
         Caption         =   "Costear_________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7320
         Picture         =   "FrmCostos.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbdocu 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   240
         Width           =   3495
      End
      Begin MSFlexGridLib.MSFlexGrid grid_doc 
         Height          =   2325
         Left            =   240
         TabIndex        =   37
         ToolTipText     =   "[Enter] = para Editar"
         Top             =   720
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4101
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         BackColorBkg    =   8421376
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.Label lblvar 
         Caption         =   "% Variaciï¿½n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4440
         TabIndex        =   45
         ToolTipText     =   "Cambiar % (TipReg 37 )"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Opciones :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4440
         TabIndex        =   41
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "Docm.:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fralista 
      Caption         =   "Lista Detalle :"
      Height          =   3615
      Left            =   9720
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton cmdesc 
         Caption         =   "Retornar"
         Height          =   555
         Left            =   4320
         TabIndex        =   23
         Top             =   2730
         Width           =   915
      End
      Begin VB.ListBox lstdoc 
         Height          =   645
         Left            =   270
         TabIndex        =   20
         Top             =   2715
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ListBox lstcosto 
         Height          =   2790
         Left            =   240
         TabIndex        =   19
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdactpre 
      Caption         =   "&Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   9330
      Picture         =   "FrmCostos.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6840
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   345
      Left            =   1470
      TabIndex        =   32
      Top             =   5670
      Visible         =   0   'False
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
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
      Left            =   11010
      Picture         =   "FrmCostos.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   435
   End
   Begin VB.Frame fracambio 
      Caption         =   "Productos con mucha variaciï¿½n :"
      Height          =   3195
      Left            =   -30
      TabIndex        =   25
      Top             =   3570
      Visible         =   0   'False
      Width           =   11895
      Begin RichTextLib.RichTextBox TEXTOVARpre 
         Height          =   375
         Left            =   3600
         TabIndex        =   31
         Top             =   510
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
         MultiLine       =   0   'False
         TextRTF         =   $"FrmCostos.frx":1694
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
         TabIndex        =   28
         Top             =   6030
         Width           =   2265
      End
      Begin VB.CommandButton cmdRET 
         Caption         =   "Retornar"
         Height          =   315
         Left            =   5220
         TabIndex        =   27
         Top             =   6000
         Width           =   1785
      End
      Begin MSFlexGridLib.MSFlexGrid gridcabe 
         Height          =   2925
         Left            =   30
         TabIndex        =   26
         ToolTipText     =   "[Enter] = para Editar"
         Top             =   210
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   5159
         _Version        =   393216
         Rows            =   1
         Cols            =   10
         FixedCols       =   5
         BackColorBkg    =   8421376
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
      Height          =   3555
      Left            =   90
      TabIndex        =   7
      Top             =   0
      Width           =   5505
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
         MultiLine       =   0   'False
         TextRTF         =   $"FrmCostos.frx":1722
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
      Height          =   3555
      Left            =   5760
      TabIndex        =   4
      Top             =   0
      Width           =   5985
      Begin VB.CommandButton cmdsec 
         Caption         =   ">>"
         Height          =   255
         Left            =   1080
         TabIndex        =   35
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox chetodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chefam 
         Caption         =   "Solo una Familia"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chefac 
         Caption         =   "Costear por Nro. Interno"
         Height          =   495
         Left            =   240
         TabIndex        =   50
         Top             =   720
         Width           =   1545
      End
      Begin VB.TextBox txtserie 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   49
         Text            =   "0"
         Top             =   870
         Width           =   525
      End
      Begin VB.TextBox txtnumero 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4050
         TabIndex        =   48
         Text            =   "0"
         Top             =   870
         Width           =   945
      End
      Begin VB.ComboBox tipmov 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdestados 
         Caption         =   "Estados de Doc. de Compras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   4560
         Picture         =   "FrmCostos.frx":17B0
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2520
         Width           =   1185
      End
      Begin VB.ComboBox TLINEA 
         Height          =   315
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1200
         Width           =   2415
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
         TabIndex        =   29
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
         Left            =   360
         TabIndex        =   24
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtava 
         Height          =   285
         Left            =   5280
         TabIndex        =   17
         Text            =   "0"
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TART 
         Height          =   288
         Left            =   4800
         TabIndex        =   12
         Top             =   360
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
         Left            =   1890
         TabIndex        =   11
         Top             =   300
         Width           =   1695
      End
      Begin MSComctlLib.ProgressBar ProgBar 
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
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
         Height          =   435
         Left            =   1920
         Picture         =   "FrmCostos.frx":1BF2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2520
         Width           =   1815
      End
      Begin MSComctlLib.ProgressBar P1 
         Height          =   195
         Left            =   600
         TabIndex        =   55
         Top             =   1560
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar P2 
         Height          =   195
         Left            =   600
         TabIndex        =   56
         Top             =   1800
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblm 
         Caption         =   "..."
         Height          =   375
         Left            =   840
         TabIndex        =   57
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Serie :"
         Height          =   195
         Index           =   0
         Left            =   3330
         TabIndex        =   53
         Top             =   660
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Numero:"
         Height          =   195
         Index           =   1
         Left            =   4050
         TabIndex        =   52
         Top             =   660
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Tip. Mov."
         Height          =   195
         Index           =   2
         Left            =   1980
         TabIndex        =   51
         Top             =   600
         Width           =   1005
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
         TabIndex        =   30
         Top             =   3060
         Width           =   3795
      End
      Begin VB.Label LART 
         Caption         =   "Codigo Alterno:"
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   360
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
         Left            =   720
         TabIndex        =   6
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
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
      Left            =   120
      TabIndex        =   16
      Top             =   7200
      Width           =   9135
   End
End
Attribute VB_Name = "Frmcospro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim actllave_precio As rdoResultset
Dim PSACTLLAVE As rdoQuery

Dim new_precio As rdoResultset
Dim PSNEWLLAVE As rdoQuery


Dim temporal
Dim temfecha
Dim flag_costeo As String * 1
Dim flag_costo_directo As String * 1
Dim WONL_INICIO As String

Private Sub chefac_Click()
If chefac.Value = 1 Then
    txtserie.Enabled = True
    txtnumero.Enabled = True
    tipmov.Enabled = True
    tipmov.ListIndex = 0
    tipmov.SetFocus
    SendKeysSeguro VK_UP, True
Else
    txtserie.Enabled = False
    txtnumero.Enabled = False
    tipmov.Enabled = False
End If
End Sub



Private Sub chefam_Click()
If chefam.Value = 1 Then
TLINEA.Visible = True
Else

TLINEA.Visible = False
End If
End Sub

Private Sub cmbdocu_KeyPress(KeyAscii As Integer)
Dim flin As Integer
Dim rs_regcos As rdoResultset
If KeyAscii <> 13 Then Exit Sub

grid_doc.Cols = 16
grid_doc.Rows = 2
grid_doc.ColWidth(0) = 500 ' tipmov
grid_doc.ColWidth(1) = 900
grid_doc.ColWidth(2) = 500
grid_doc.ColWidth(3) = 900
grid_doc.ColWidth(4) = 500
grid_doc.ColWidth(5) = 900
grid_doc.ColWidth(6) = 2000
grid_doc.ColWidth(7) = 800
grid_doc.ColWidth(8) = 0
grid_doc.ColWidth(9) = 500
grid_doc.ColWidth(10) = 900
grid_doc.ColWidth(11) = 900
grid_doc.ColWidth(12) = 900
grid_doc.ColWidth(13) = 1500
grid_doc.ColWidth(14) = 900
grid_doc.ColWidth(15) = 1100


grid_doc.TextMatrix(0, 0) = "TipMov"
grid_doc.TextMatrix(0, 1) = "Fecha"
grid_doc.TextMatrix(0, 2) = "Serie"
grid_doc.TextMatrix(0, 3) = "Numero"
grid_doc.TextMatrix(0, 4) = "Sec"
grid_doc.TextMatrix(0, 5) = "Codigo"
grid_doc.TextMatrix(0, 6) = "Producto"
grid_doc.TextMatrix(0, 7) = "U.M"
grid_doc.TextMatrix(0, 8) = "Cantidad"
grid_doc.TextMatrix(0, 9) = "Estado"
grid_doc.TextMatrix(0, 10) = "P.Compra"
grid_doc.TextMatrix(0, 11) = "C.Actual"
grid_doc.TextMatrix(0, 12) = "C.Anterior"
grid_doc.TextMatrix(0, 13) = "Proveedor"
grid_doc.TextMatrix(0, 14) = "Usuario"
grid_doc.TextMatrix(0, 15) = "Hora"

grid_doc.Rows = 1

pub_cadena = "SELECT * FROM TABESTADOS WHERE TAE_CODCIA = '" & LK_CODCIA & "' AND TAE_TIPMOV = '" & Val(Left(cmbdocu.Text, 2)) & "'  AND TAE_NUMSER = '" & Val(Mid(cmbdocu.Text, 23, 2)) & "' AND TAE_NUMFAC = '" & Val(Right(cmbdocu.Text, 8)) & "' "
Set rs_regcos = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
Do Until rs_regcos.EOF
grid_doc.Rows = grid_doc.Rows + 1
        grid_doc.TextMatrix(grid_doc.Rows - 1, 0) = rs_regcos!tae_tipmov
        grid_doc.TextMatrix(grid_doc.Rows - 1, 1) = Format(rs_regcos!tae_fecha_compra, "dd/nm/yy")
        grid_doc.TextMatrix(grid_doc.Rows - 1, 2) = rs_regcos!TAE_NUMSER
        grid_doc.TextMatrix(grid_doc.Rows - 1, 3) = rs_regcos!tae_numfac
        grid_doc.TextMatrix(grid_doc.Rows - 1, 4) = rs_regcos!tae_numsec
        grid_doc.TextMatrix(grid_doc.Rows - 1, 5) = rs_regcos!tae_codart
        SQ_OPER = 1
        pu_codcia = LK_CODCIA
        PUB_KEY = rs_regcos!tae_codart
        LEER_ART_LLAVE
        If art_LLAVE.EOF Then
        Else
        grid_doc.TextMatrix(grid_doc.Rows - 1, 6) = art_LLAVE!art_nombre
        End If
        grid_doc.TextMatrix(grid_doc.Rows - 1, 7) = rs_regcos!tae_descri
        grid_doc.TextMatrix(grid_doc.Rows - 1, 8) = rs_regcos!tae_cantidad
        grid_doc.TextMatrix(grid_doc.Rows - 1, 9) = rs_regcos!tae_estado
        grid_doc.TextMatrix(grid_doc.Rows - 1, 10) = Format(rs_regcos!tae_precio, "0.00")
        grid_doc.TextMatrix(grid_doc.Rows - 1, 11) = Format(rs_regcos!TAE_COSPRO, "0.00")
        grid_doc.TextMatrix(grid_doc.Rows - 1, 12) = Format(rs_regcos!TAE_COSPRO_ANT, "0.00")
        SQ_OPER = 1
        pu_cp = "P"
        pu_codclie = rs_regcos!tae_codclie
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_llave.EOF Then
        Else
        grid_doc.TextMatrix(grid_doc.Rows - 1, 13) = cli_llave!cli_nombre
        End If
        grid_doc.TextMatrix(grid_doc.Rows - 1, 14) = rs_regcos!tae_codusu
        grid_doc.TextMatrix(grid_doc.Rows - 1, 15) = Trim(rs_regcos!tae_hora)
rs_regcos.MoveNext
Loop
End Sub

Private Sub cmdactpre_Click()
'If LK_CODCIA = "01" Or LK_CODCIA = "30" Then
'Else
'  MsgBox "No Permite Cambiar en esta Compañia", 48, Pub_Titulo
'  Exit Sub
'End If
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
   Exit Sub
End If

If flag_cambio_pre <> "A" Then
  MsgBox "No tiene Acceso !!!", 48, Pub_Titulo
  Exit Sub
End If


If gridcabe.Rows <= 1 Then Exit Sub

pub_mensaje = "Esta seguro que desea Actalizar las Lista de Precios ï¿½ Continuar? ..."
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
If gridcabe.Rows <> 1 Then
pbMin = 0
pbMax = gridcabe.Rows - 1
pbValue = 0
pb.Visible = True
DoEvents
End If

For fila = 1 To gridcabe.Rows - 1
  pbValue = pbValue + 1
  DoEvents
  PUB_CODCIA = LK_CODCIA
  GoSub ACTALIZA_PRE_POR
  If LK_CODCIA = "05" Or LK_CODCIA = "09" Or LK_CODCIA = "50" Then
  Else
    If LK_CODCIA <> "30" Then  ' no en la de prueba
      PUB_CODCIA = "03"
      GoSub ACTALIZA_PRE_POR
    End If
  End If
Next fila
pb.Visible = False
DoEvents
MsgBox "Lista de Precios a sido cambiado con exito.", 48, Pub_Titulo

Exit Sub

ACTALIZA_PRE_POR:

    SQ_OPER = 1
    pu_codcia = PUB_CODCIA
    PUB_CODART = Val(gridcabe.TextMatrix(fila, 0))
    PUB_SECUEN = Val(gridcabe.TextMatrix(fila, 23))
    LEER_PRE_LLAVE
    If pre_llave.EOF Then
      MsgBox "Codigo no Actualizo ", 48, Pub_Titulo
      GoTo pasa_cod
    End If
    
    pre_llave.Edit
    pre_llave!PRE_PRE1 = Val(gridcabe.TextMatrix(fila, 7))
    pre_llave!PRE_POR1 = Val(gridcabe.TextMatrix(fila, 5))
    
    pre_llave!PRE_PRE2 = Val(gridcabe.TextMatrix(fila, 10))
    pre_llave!PRE_POR2 = Val(gridcabe.TextMatrix(fila, 8))
    
    pre_llave!PRE_PRE3 = Val(gridcabe.TextMatrix(fila, 13))
    pre_llave!PRE_POR3 = Val(gridcabe.TextMatrix(fila, 11))
    
    pre_llave!PRE_PRE4 = Val(gridcabe.TextMatrix(fila, 16))
    pre_llave!PRE_POR4 = Val(gridcabe.TextMatrix(fila, 14))
    
    pre_llave!PRE_PRE5 = Val(gridcabe.TextMatrix(fila, 19))
    pre_llave!PRE_POR5 = Val(gridcabe.TextMatrix(fila, 17))
    
    pre_llave!PRE_PRE6 = Val(gridcabe.TextMatrix(fila, 22))
    pre_llave!PRE_POR6 = Val(gridcabe.TextMatrix(fila, 20))
    If (Val(pre_llave!PRE_PRE1) + Val(pre_llave!PRE_PRE2) + Val(pre_llave!PRE_PRE3) + Val(pre_llave!PRE_PRE4) + Val(pre_llave!PRE_PRE5)) = 0 Then
       MsgBox "Producto : " & Trim(gridcabe.TextMatrix(fila, 1)) & " en " & Trim(pre_llave!PRE_UNIDAD) & Chr(13) & " Tienes por lo menos 5 Listas de Precios en 0 , Verificarlo ", 48, Pub_Titulo
    End If
    pre_llave.Update
    
pasa_cod:
    
Return
    
End Sub

Private Sub cmdbp_Click()
Dim rs_regcos As rdoResultset
Dim wcod As String
wcod = InputBox("Codigo interno de Producto:", "Buscar Codigo", 0)
If wcod = "" Then Exit Sub
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_KEY = wcod
LEER_ART_LLAVE
If art_LLAVE.EOF Then
  MsgBox "Codigo no Existe", 48, Pub_Titulo
  Exit Sub
End If

grid_doc.Rows = 1
cmbdocu.Clear
pub_cadena = "SELECT  DISTINCT TAE_CODCIA, TAE_TIPMOV ,TAE_FECHA_COMPRA, TAE_NUMSER,TAE_NUMFAC FROM TABESTADOS WHERE TAE_CODCIA = '" & LK_CODCIA & "' AND TAE_ESTADO IN ('B','C') and tae_codart = " & wcod & ""
Set rs_regcos = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
Do Until rs_regcos.EOF
    cmbdocu.AddItem Format(rs_regcos!tae_tipmov, "00") & " - " & Format(rs_regcos!tae_fecha_compra, "dd/mm/yyyy") & " Nro : " & Format(rs_regcos!TAE_NUMSER, "0") & "  - " & Format(rs_regcos!tae_numfac, "00000000")
rs_regcos.MoveNext
Loop
If cmbdocu.ListCount > 0 Then cmbdocu.ListIndex = 0
SQ_OPER = 1
PUB_CODCIA = "00"
PUB_TIPREG = 37
PUB_NUMTAB = 0
LEER_TAB_LLAVE
If tab_llave.EOF Then
lblvar.Caption = "0% Variaciï¿½n(B)"
Else
lblvar.Caption = Trim(tab_llave!TAB_NOMLARGO) & "% Variaciï¿½n(B)"
End If

fradoccompra.Visible = True
cmbdocu.SetFocus

End Sub

Private Sub cmdcanceldoc_Click()
fradoccompra.Visible = False
End Sub

Private Sub CmdCosteoDoc_Click()
pub_mensaje = "Costear Directo Continuar? ..."
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If

chefac.Value = 1
tipmov.ListIndex = 0
txtfecha.Text = Format((grid_doc.TextMatrix(1, 1)), "dd/mm/yyyy")
txtserie.Text = Trim(grid_doc.TextMatrix(1, 2))
txtnumero.Text = Trim(grid_doc.TextMatrix(1, 3))
flag_costo_directo = "A"
cmdcanceldoc_Click
Command1_Click
flag_costo_directo = ""


End Sub

Private Sub cmdesc_Click()
fralista.Visible = False
End Sub

Private Sub CmdEstado_Click()
Dim wvalor As String
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
   Exit Sub
End If


wvalor = InputBox("Forzar al Estado (A/B/C).!!! : ", "Cambiar Estado", "")
If wvalor = "" Then Exit Sub
If wvalor = "A" Or wvalor = "B" Or wvalor = "C" Then
Else
  MsgBox "Estado invalido, A o B o C", 48, Pub_Titulo
  Exit Sub
End If



pub_cadena = "UPDATE TABESTADOS SET TAE_ESTADO = '" & wvalor & "' WHERE TAE_CODCIA = '" & LK_CODCIA & "' AND TAE_TIPMOV = '" & Left(cmbdocu.Text, 2) & "'  AND TAE_NUMSER = '" & Val(Mid(cmbdocu.Text, 23, 2)) & "' AND TAE_NUMFAC = '" & Val(Right(cmbdocu.Text, 8)) & "' "
CN.Execute pub_cadena, rdExecDirect
MsgBox "Estado Cambiado", 48, Pub_Titulo
cmdcanceldoc_Click

End Sub

Private Sub cmdestados_Click()
Dim rs_regcos As rdoResultset
grid_doc.Rows = 1
cmbdocu.Clear
pub_cadena = "SELECT  DISTINCT TAE_CODCIA, TAE_TIPMOV ,TAE_FECHA_COMPRA, TAE_NUMSER,TAE_NUMFAC FROM TABESTADOS WHERE TAE_CODCIA = '" & LK_CODCIA & "' AND TAE_ESTADO IN ('B','C')"
Set rs_regcos = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
Do Until rs_regcos.EOF
    cmbdocu.AddItem Format(rs_regcos!tae_tipmov, "00") & " - " & Format(rs_regcos!tae_fecha_compra, "dd/mm/yyyy") & " Nro : " & Format(rs_regcos!TAE_NUMSER, "0") & "  - " & Format(rs_regcos!tae_numfac, "00000000")
rs_regcos.MoveNext
Loop
If cmbdocu.ListCount > 0 Then cmbdocu.ListIndex = 0
SQ_OPER = 1
PUB_CODCIA = "00"
PUB_TIPREG = 37
PUB_NUMTAB = 0
LEER_TAB_LLAVE
If tab_llave.EOF Then
lblvar.Caption = "0% Variaciï¿½n(B)"
Else
lblvar.Caption = Trim(tab_llave!TAB_NOMLARGO) & "% Variaciï¿½n(B)"
End If

fradoccompra.Visible = True
cmbdocu.SetFocus
End Sub


Private Sub cmdRET_Click()
fracambio.Visible = False
End Sub

Private Sub cmdsec_Click()
Dim I As Integer
Dim Y As Integer
WONL_INICIO = InputBox("Iniciar con Numero de Orden Nï¿½ ", "Inicializar", 0)
If WONL_INICIO = "" Then Exit Sub
Command1_Click

Exit Sub
pub_mensaje = "Continuar con el Costeo en Serie !!! ...   ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
Dim ps_cosprouno  As rdoResultset
flag_costeo = "A"
pub_cadena = "SELECT ART_ALTERNO, ART_NOMBRE FROM ARTI WHERE ART_CODCIA = '" & LK_CODCIA & "' AND ART_ALTERNO >= '" & Trim(TART.Text) & "' ORDER BY ART_ALTERNO"
Set ps_cosprouno = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)

Do Until ps_cosprouno.EOF
  
  CHEART.Value = 1
  cmdsec.Caption = ps_cosprouno.AbsolutePosition & " / " & ps_cosprouno.RowCount
  TART.Text = ps_cosprouno!ART_alterno
  DoEvents
  Command1_Click
 ps_cosprouno.MoveNext
Loop
flag_costeo = ""
MsgBox "TERMINO OK.", 48, Pub_Titulo
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
grid_doc.Visible = False
DoEvents
Dim WCODART As Currency
Dim wFLAG_ADDPRE As String
Dim wol_codcia As String
If cheactpre.Value = 1 Then
wFLAG_ADDPRE = "A"
Else
wFLAG_ADDPRE = ""
End If
If LK_CODCIA = "09" Or LK_CODCIA = "30" Or LK_CODCIA = "50" Then
   'Costeo_Chepen ' solo pï¿½ra la botica chepen
   wol_codcia = "('" & LK_CODCIA & "')"
ElseIf LK_CODCIA = "01" Or LK_CODCIA = "20" Then
   wol_codcia = "('01','02','04','08','20')"
ElseIf LK_CODCIA = "03" Or LK_CODCIA = "05" Or LK_CODCIA = "07" Then
   wol_codcia = "('03','05','07')"
ElseIf LK_CODCIA = "10" Or LK_CODCIA = "25" Or LK_CODCIA = "27" Or LK_CODCIA = "28" Or LK_CODCIA = "29" Then
   wol_codcia = "('10','25','27','28','29','15')"
Else
   wol_codcia = "('" & LK_CODCIA & "')"
    ' wol_codcia = "('01','02','03','04','07','08','10','20')"
End If
    pub_mensaje = "Confirmar Costeo , Continuar ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
   ''Costeo_Gen ' para la principal
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
        CosteoOnLine "P", wFLAG_ADDPRE, wol_codcia, "A", "(" & WCODART & ")", 0, Format(txtfecha.Text, "dd/mm/yyyy"), 0, 0, " ", 0, P1, P2, lblm, 0
   ElseIf chefac.Value = 1 Then
    PSFAR_LLAVE(0) = Val(Left(tipmov.Text, 3))
    PSFAR_LLAVE(1) = LK_CODCIA
    PSFAR_LLAVE(2) = Val(txtserie.Text)
    PSFAR_LLAVE(3) = " "
    PSFAR_LLAVE(4) = Val(txtnumero.Text)
    far_llave.Requery
    If far_llave.EOF Then
         MsgBox "Este Documentono Existe", 48, Pub_Titulo
         ProgBar.Visible = False
         Command1.Enabled = True
         Exit Sub
    Else
         txtfecha.Text = Format(far_llave!FAR_fecha_compra, "dd/mm/yyyy")
    End If
     CosteoOnLine "D", wFLAG_ADDPRE, wol_codcia, "A", "(0)", 0, Format(txtfecha.Text, "dd/mm/yyyy"), Val(Left(tipmov.Text, 3)), Val(txtserie.Text), " ", Val(txtnumero.Text), P1, P2, lblm, 0
   ElseIf chefam.Value = 1 Then
     CosteoOnLine "L", wFLAG_ADDPRE, wol_codcia, "A", "(0)", Val(Right(TLINEA.Text, 10)), Format(txtfecha.Text, "dd/mm/yyyy"), 0, 0, " ", 0, P1, P2, lblm, 0
   ElseIf chetodos.Value = 1 Then
     CosteoOnLine "T", wFLAG_ADDPRE, wol_codcia, "A", "(0)", 45, "01/01/2008", 20, 0, " ", 88, P1, P2, lblm, Val(WONL_INICIO)
   End If
 If wFLAG_ADDPRE = "A" Then
  cabeprecios
  Muestra_Preicios 0
 End If
 
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Unload Frmcospro
End Sub


Private Sub CHEART_Click()
If CHEART.Value = 1 Then
  TART.Visible = True
  LART.Visible = True
  TLINEA.Visible = True
  TART.SetFocus
Else
  TART.Visible = False
  LART.Visible = False
  TLINEA.Visible = False
End If
End Sub

Private Sub Form_Activate()
If txtfecha.Visible Then
 txtfecha_KeyPress 13
 Azul2 txtfecha, txtfecha
End If
End Sub

Private Sub Form_Load()
flag_costo_directo = ""
pub_cadena = "SELECT * FROM NEWPRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ? "
Set PSNEWLLAVE = CN.CreateQuery("", pub_cadena)
PSNEWLLAVE(0) = 0
PSNEWLLAVE(1) = 0
Set new_precio = PSNEWLLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

CenterMe Frmcospro
If LK_CODUSU = "ADMIN" Then TLINEA.Visible = True
'LblFecha.Caption = Format(LK_FECHA_DIA, "dddd, d mmmm yyyy")
'empresa.Caption = Trim(par_llave!PAR_NOMBRE)
txtfecha.Text = "01/" & Format(DatePart("m", LK_FECHA_DIA), "00") & "/" & DatePart("yyyy", LK_FECHA_DIA)
txtfecha.Mask = "##/##/####"
txtfecha.Visible = True
txtfecha.TabIndex = 0
Muestra_tc txtfecha.Text
temfecha = LK_FECHA_DIA
If LK_CODUSU = "ADMIN" Then
   txtava.Visible = True
End If
fechafinal.Visible = True
fechafinal.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
tipmov.Clear
tipmov.AddItem "020 - Compras"
tipmov.AddItem "006 - Ing. Ajuste"
tipmov.AddItem "101 - Recep. Merca."
tipmov.AddItem "093 - Transf. Paqt."
tipmov.Enabled = False
SQ_OPER = 2
PUB_TIPREG = 122
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
TLINEA.AddItem "<Todos>"
Do Until tab_mayor.EOF
   TLINEA.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
   tab_mayor.MoveNext
Loop
flag_costeo = ""
'End If
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


If gridcabe.COL = 7 Or gridcabe.COL = 10 Or gridcabe.COL = 13 Or gridcabe.COL = 16 Or gridcabe.COL = 19 Or gridcabe.COL = 22 Then
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


Private Sub Label8_Click()

End Sub

Private Sub Label8_DblClick()
    
End Sub

Private Sub lblvar_DblClick()
Dim wvalor As String
SQ_OPER = 1
PUB_CODCIA = "00"
PUB_TIPREG = 37
PUB_NUMTAB = 0
LEER_TAB_LLAVE
If tab_llave.EOF Then
wvalor = 0
Else
wvalor = Val(Trim(tab_llave!TAB_NOMLARGO))
End If
wvalor = InputBox("Cambiar Estado B para El Porcentaje de Variaciï¿½n del Costo : %", "Cambiar Porcentaje", wvalor)
If wvalor = "" Then Exit Sub

If tab_llave.EOF Then
tab_llave.AddNew
tab_llave!tab_codcia = "00"
tab_llave!tab_tipreg = PUB_TIPREG
tab_llave!TAB_NUMTAB = PUB_NUMTAB
tab_llave!TAB_NOMLARGO = wvalor
Else
tab_llave.Edit
tab_llave!TAB_NOMLARGO = wvalor
End If
tab_llave.Update



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


VALOR = (Val(TEXTOVARpre.Text) * 100) / Val(gridcabe.TextMatrix(gridcabe.Row, 4)) - 100
If VALOR = -100 Then VALOR = 0
gridcabe.CellForeColor = QBColor(9)
gridcabe.TextMatrix(gridcabe.Row, gridcabe.COL - 2) = Format(VALOR, "0.000")




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
gridigv.TextMatrix(0, 1) = "Publicaciï¿½n"
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

Private Sub tipmov_Click()
Azul txtserie, txtserie
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

WRES = DateDiff("M", wfecha_ini, Frmcospro.fechafinal.Text) + 1
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

Public Sub Costeo_Gen()
Dim wvalorVari As Currency
Dim wvardif As Currency
Dim rs_regcos As rdoResultset
Dim wx_suma_cia01 As Currency
Dim wx_suma_cia03 As Currency
Dim wx_suma_cia07 As Currency
Dim wsa_botica As Currency
Dim wcan_botica As Currency
Dim WS_SALDO_ALM As Currency
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
Dim Far_Cost_alm As rdoResultset
Dim PSFAR_COSTO_ALM As rdoQuery

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
WS_SALDO_ALM = 0
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

If flag_costeo = "A" Or flag_costo_directo = "A" Then
Else
    pub_mensaje = "Proceso de Costeo de Articulos de la Fecha : " & Chr(13) & " " & Format(txtfecha.Text, "dd/mm/yyyy") & " al " & Format(LK_FECHA_DIA, "dd/mm/yyyy") & Chr(13) & "ï¿½ Continuar? ..."
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
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

 ' jala el porc. de alaterta de variacion de costo
 '============
SQ_OPER = 1
PUB_CODCIA = "00"
PUB_TIPREG = 37
PUB_NUMTAB = 0
LEER_TAB_LLAVE
If tab_llave.EOF Then
wvalorVari = 0
Else
wvalorVari = Val(Trim(tab_llave!TAB_NOMLARGO))
End If


pub_cadena = "DELETE FROM NEWPRECIOS WHERE PRE_CODCIA = '" & LK_CODCIA & "'"
CN.Execute pub_cadena, rdExecDirect

'pub_cadena = "SELECT FAR_COSTEO, FAR_COSTEO_REAL, FAR_PRECIO FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_NUMGUIA = ? AND FAR_TIPMOV = 101  AND FAR_ESTADO <> 'E' and far_transito <> 'P' "
'Set PSFAR_ENVIO = CN.CreateQuery("", pub_cadena)
'PSFAR_ENVIO(0) = 0
'PSFAR_ENVIO(1) = 0
'PSFAR_ENVIO(2) = 0
'Set Far_envio = PSFAR_ENVIO.OpenResultset(rdOpenKeyset, rdConcurValues)

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

pub_cadena = "SELECT FAR_FECHA_COMPRA, FAR_FECHA_CAN,FAR_COSPRO_SUP, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_STOCK, FAR_COSPRO FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA < ? AND FAR_CODART = ? and far_estado <> 'E' and far_estado2 <> 'L' and far_transito <> 'P' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_TIPMOV, FAR_NUMOPER2"
Set PSFAR_COSTO2 = CN.CreateQuery("", pub_cadena)
PSFAR_COSTO2(0) = 0
PSFAR_COSTO2(1) = LK_FECHA_DIA
PSFAR_COSTO2(2) = 0
Set Far_Cost2 = PSFAR_COSTO2.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT FAR_FECHA_COMPRA, FAR_FECHA_CAN,FAR_COSPRO_SUP, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_STOCK, FAR_COSPRO FROM FACART WHERE FAR_CODCIA IN (?) AND FAR_FECHA_COMPRA < ? AND FAR_CODART = ? and far_estado <>'E' and far_estado2 <>'L' and far_transito <> 'P' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_TIPMOV, FAR_NUMOPER2"
Set PSFAR_COSTO_ALM = CN.CreateQuery("", pub_cadena)
PSFAR_COSTO_ALM(0) = 0
PSFAR_COSTO_ALM(1) = LK_FECHA_DIA
PSFAR_COSTO_ALM(2) = 0
Set Far_Cost_alm = PSFAR_COSTO_ALM.OpenResultset(rdOpenKeyset, rdConcurValues)



'pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_FECHA_COMPRA >= ? AND  FAR_FECHA_COMPRA <= ?   and far_estado<>'E' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_NUMOPER2"


pub_cadena = "SELECT FAR_CODCIA, FAR_NUM_LOTE,FAR_PRECIO_NETO,FAR_COSPRO_ANT,FAR_COSPRO_SUP, FAR_STOCK , FAR_TIPO_CAMBIO, FAR_COSTEO_REAL, FAR_NUMGUIA, FAR_NUMSER, FAR_NUMFAC, FAR_OTRA_CIA, FAR_CODART, FAR_COSPRO, FAR_FLETE, FAR_BRUTO, FAR_DESCTO, FAR_GASTOS, FAR_TOT_DESCTO, FAR_SIGNO_ARM, FAR_CANTIDAD, FAR_TIPMOV, FAR_MONEDA, FAR_FECHA_COMPRA, FAR_FECHA_CAN, FAR_COSTEO, FAR_ESTADO, FAR_PRECIO, FAR_EQUIV, FAR_NUMSEC  FROM FACART WHERE FAR_CODCIA IN (?,?,?) AND FAR_CODART = ? AND FAR_FECHA_COMPRA >= ? AND  FAR_FECHA_COMPRA <= ?   and far_estado <>'E' and far_estado2 <>'L' and far_transito <> 'P' ORDER BY FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC ,FAR_TIPMOV, FAR_NUMOPER2,FAR_CODCIA"
Set PSFAR_COSTO = CN.CreateQuery("", pub_cadena)
If LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07" Then
PSFAR_COSTO(0) = "01"
PSFAR_COSTO(1) = "03"
PSFAR_COSTO(2) = "07"
Else
PSFAR_COSTO(0) = LK_CODCIA
PSFAR_COSTO(1) = "-1"
PSFAR_COSTO(2) = "-1"
End If
PSFAR_COSTO(3) = 0
PSFAR_COSTO(4) = 0
PSFAR_COSTO(5) = LK_FECHA_DIA
Set Far_Cost = PSFAR_COSTO.OpenResultset(rdOpenKeyset, rdConcurValues)

'ï¿½DistinctRow
'DISTINCTROW
If chefac.Value = 1 Then
   PSFAR_LLAVE(0) = Val(Left(tipmov.Text, 3))
   PSFAR_LLAVE(1) = LK_CODCIA
   PSFAR_LLAVE(2) = Val(txtserie.Text)
   PSFAR_LLAVE(3) = " "
   PSFAR_LLAVE(4) = Val(txtnumero.Text)
   far_llave.Requery
   If far_llave.EOF Then
        MsgBox "Este Documentono Existe", 48, Pub_Titulo
        ProgBar.Visible = False
        Command1.Enabled = True
        Exit Sub
   Else
      If Format(txtfecha, "dd/mm/yyyy") <> Format(far_llave!FAR_fecha_compra, "dd/mm/yyyy") Then
        MsgBox "La Fecha de Emisión del Documento no Coincide con la Fecha Inicial del Costeo " & Chr(13) & "Fecha de Emisión del Documento : " & Format(far_llave!FAR_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo
        txtfecha.Text = Format(far_llave!FAR_fecha_compra, "dd/mm/yyyy")
        ProgBar.Visible = False
        Command1.Enabled = True
        Exit Sub
      End If
   End If
 pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ?  AND  FAR_ESTADO <> 'E' AND  FAR_ESTADO <> 'L' AND ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' )  and far_tipmov = " & Val(Left(tipmov.Text, 3)) & " and far_numser = " & Val(txtserie.Text) & " and far_numfac = " & txtnumero.Text & " ORDER BY FAR_CODART"
Else
 pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ?  AND  FAR_ESTADO <> 'E' AND FAR_ESTADO <> 'L' AND FAR_CODART <> 0  ORDER BY FAR_CODART " ' (FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' )  ORDER BY FAR_CODART"
End If
If CHEART.Value = 1 Then
  pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA IN ('01','07','03') AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND FAR_CODART = " & WCODART & "  AND  ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' ) AND FAR_ESTADO <> 'E' AND FAR_ESTADO2 <> 'L' AND FAR_CODART <> 0 ORDER BY FAR_CODART" ' AND ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' ) ORDER BY FAR_CODART"
End If
If Val(Right(TLINEA.Text, 8)) <> 0 Then
  WLINEA = Val(Right(TLINEA.Text, 8))
  pub_cadena = "SELECT Distinct FAR_CODART FROM FACART, ARTI WHERE (FAR_CODART = ART_KEY) AND (FAR_CODCIA = ART_CODCIA) AND FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND ARTI.ART_FAMILIA = " & WLINEA & "  AND  FAR_ESTADO <> 'E' AND FAR_ESTADO2 <> 'L' AND FAR_CODART <> 0 ORDER BY FAR_CODART" ' AND ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' ) ORDER BY FAR_CODART"
End If
' TEMPORAL
''''pub_cadena = "SELECT Distinct FAR_CODART FROM FACART, ARTI WHERE (ART_CODCIA = FAR_CODCIA) AND (FAR_CODART = ART_KEY) AND  FAR_TIPMOV = 6 AND FAR_NUM_LOTE <> 10 AND FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND ( FAR_COSTEO_REAL = 'A'  OR FAR_FECHA_COMPRA <> FAR_FECHA  OR FAR_MONEDA= 'D' )  AND FAR_ESTADO<>'E'  ORDER BY FAR_CODART"
Set PSFAR_COMPR = CN.CreateQuery("", pub_cadena)
If CHEART.Value = 1 Then
 PSFAR_COMPR(0) = Date
 PSFAR_COMPR(1) = Date
Else
 PSFAR_COMPR(0) = LK_CODCIA
 PSFAR_COMPR(1) = Date
 PSFAR_COMPR(2) = Date
End If
Set Far_Compr = PSFAR_COMPR.OpenResultset(rdOpenKeyset, rdConcurValues)
If CHEART.Value = 1 Then
    PSFAR_COMPR.rdoParameters(0) = CDate(txtfecha.Text)
    If fechafinal.Visible Then
     PSFAR_COMPR.rdoParameters(1) = fechafinal.Text
    Else
     PSFAR_COMPR.rdoParameters(1) = LK_FECHA_DIA
    End If

Else
    PSFAR_COMPR.rdoParameters(0) = LK_CODCIA
    PSFAR_COMPR.rdoParameters(1) = CDate(txtfecha.Text)
    If fechafinal.Visible Then
     PSFAR_COMPR.rdoParameters(2) = fechafinal.Text
    Else
     PSFAR_COMPR.rdoParameters(2) = LK_FECHA_DIA
    End If
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
   Unload Frmcospro
   Unload FORMGEN
   Exit Sub
End If
'OPEN_LOG "Proceso de Costeo"
gridcabe.Visible = False
POR(2).Caption = "Actualizando Costos. . ."
DoEvents
wCODARTI = Far_Compr!far_codart

flagcito = "A"
WPASA_COSPRO_ANT = 0
WCOSPRO_SUP = 0
Do Until Far_Compr.EOF
'   If Far_Compr!far_codart = 1504505 Then Stop
   If Val(txtava.Text) <> 0 Then
      If Far_Compr.AbsolutePosition < Val(txtava.Text) Then GoTo OTRO_ART
   End If
   ProgBar.Value = Far_Compr.AbsolutePosition
   POR(2).Caption = "Costeando... " & ProgBar.Value & " de " & ProgBar.max
   DoEvents
   
  ' PSFAR_COSTO.rdoParameters(0) = LK_CODCIA
  
   PSFAR_COSTO.rdoParameters(3) = Far_Compr!far_codart
   PSFAR_COSTO.rdoParameters(4) = CDate(txtfecha.Text)
   If fechafinal.Visible Then
     PSFAR_COSTO.rdoParameters(5) = fechafinal.Text
   Else
     PSFAR_COSTO.rdoParameters(5) = LK_FECHA_DIA
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
    Print Far_Cost!FAR_CODCIA
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
   'If chefac.Value = 1 Then
        'pub_cadena = "select TOP 1 FAR_COSPRO FROM FACART WHERE FAR_CODCIA IN ('01','03','07') AND " & _
        '    " FAR_CODART = " & Far_Cost!far_codart & " AND FAR_FECHA_COMPRA < '" & Format(txtfecha.Text, "dd/mm/yyyy") & "' " & _
        '    "and far_estado <>'E' and far_estado2 <>'L' and far_transito <> 'P' " & _
        '    "ORDER BY FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC ,FAR_TIPMOV, FAR_NUMOPER2,FAR_CODCIA"
        'Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
        'If Not X.EOF Then
        'WCOSPRO_SUP = Nulo_Valor0(X!FAR_COSPRO)
        'ws_cospro = Nulo_Valor0(X!FAR_COSPRO)
        'Else
        '  WCOSPRO_SUP = 0
        '  ws_cospro = 0
        'End If
    'Else
      WCOSPRO_SUP = Nulo_Valor0(Far_Cost!far_cospro_sup)
'      Print Far_Cost.RowCount
      ws_cospro = Nulo_Valor0(Far_Cost!far_cospro_sup)
    'End If
   Else
      WCOSPRO_SUP = Nulo_Valor0(Far_Cost!FAR_COSPRO)
      ws_cospro = Nulo_Valor0(Far_Cost!FAR_COSPRO)
   End If
   wx_suma_cia03 = 0
   If LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07" Then
           If LK_CODCIA = "01" Then
             PSFAR_COSTO2(0) = "03"
             PSFAR_COSTO2(1) = Format(txtfecha.Text, "dd/mm/yyyy")
             PSFAR_COSTO2(2) = Far_Cost!far_codart
             Far_Cost2.Requery
             wx_suma_cia03 = 0
             Far_Cost2.MoveLast
             If Far_Cost2.EOF Then
               wx_suma_cia03 = 0
             Else
               wx_suma_cia03 = Far_Cost2!FAR_STOCK
             End If
             PSFAR_COSTO2(0) = "07"
             PSFAR_COSTO2(1) = Format(txtfecha.Text, "dd/mm/yyyy")
             PSFAR_COSTO2(2) = Far_Cost!far_codart
             Far_Cost2.Requery
             wx_suma_cia07 = 0
             Far_Cost2.MoveLast
             If Far_Cost2.EOF Then
               wx_suma_cia07 = 0
             Else
               wx_suma_cia07 = Far_Cost2!FAR_STOCK
             End If
           ElseIf LK_CODCIA = "03" Then
             PSFAR_COSTO2(0) = "01"
             PSFAR_COSTO2(1) = Format(txtfecha.Text, "dd/mm/yyyy")
             PSFAR_COSTO2(2) = Far_Cost!far_codart
             Far_Cost2.Requery
             wx_suma_cia01 = 0
             Far_Cost2.MoveLast
             If Far_Cost2.EOF Then
                wx_suma_cia01 = 0
             Else
                wx_suma_cia01 = Far_Cost2!FAR_STOCK
                WCOSPRO_SUP = Far_Cost2!FAR_COSPRO 'COSTO DE LA 01 SIEMPRE
                ws_cospro = Far_Cost2!FAR_COSPRO
             End If
             PSFAR_COSTO2(0) = "07"
             PSFAR_COSTO2(1) = Format(txtfecha.Text, "dd/mm/yyyy")
             PSFAR_COSTO2(2) = Far_Cost!far_codart
             Far_Cost2.Requery
             wx_suma_cia07 = 0
             Far_Cost2.MoveLast
             If Far_Cost2.EOF Then
               wx_suma_cia07 = 0
             Else
               wx_suma_cia07 = Far_Cost2!FAR_STOCK
             End If
           Else
             PSFAR_COSTO2(0) = "01"
             PSFAR_COSTO2(1) = Format(txtfecha.Text, "dd/mm/yyyy")
             PSFAR_COSTO2(2) = Far_Cost!far_codart
             Far_Cost2.Requery
             wx_suma_cia01 = 0
             Far_Cost2.MoveLast
             'Print Far_Cost2!far_fecha_compra
             
             If Far_Cost2.EOF Then
                wx_suma_cia01 = 0
             Else
                wx_suma_cia01 = Far_Cost2!FAR_STOCK
                WCOSPRO_SUP = Far_Cost2!FAR_COSPRO 'COSTO DE LA 01 SIEMPRE
                ws_cospro = Far_Cost2!FAR_COSPRO
             End If
             PSFAR_COSTO2(0) = "03"
             PSFAR_COSTO2(1) = Format(txtfecha.Text, "dd/mm/yyyy")
             PSFAR_COSTO2(2) = Far_Cost!far_codart
             Far_Cost2.Requery
             wx_suma_cia03 = 0
             Far_Cost2.MoveLast
             If Far_Cost2.EOF Then
               wx_suma_cia03 = 0
             Else
               wx_suma_cia03 = Far_Cost2!FAR_STOCK
             End If
           End If
   End If
      
   PSFAR_COSTO2(0) = LK_CODCIA
   PSFAR_COSTO2(1) = Format(txtfecha.Text, "dd/mm/yyyy")
   PSFAR_COSTO2(2) = Far_Cost!far_codart
   Far_Cost2.Requery
   WS_SALDO_ALM = 0
'   If Far_Cost!far_codart = 891882 Then Stop
   Far_Cost2.MoveLast
   If Far_Cost2.EOF Then
     WS_SALDO = 0
   Else
     WS_SALDO = Far_Cost2!FAR_STOCK
     If LK_CODCIA = "01" Then
       wx_suma_cia01 = Far_Cost2!FAR_STOCK
       If WS_SALDO = 0 And ws_cospro <> 0 Then
       Else
        WCOSPRO_SUP = Far_Cost2!FAR_COSPRO
        ws_cospro = Far_Cost2!FAR_COSPRO
      End If
     ElseIf LK_CODCIA = "03" Then
       wx_suma_cia03 = Far_Cost2!FAR_STOCK
     ElseIf LK_CODCIA = "07" Then
       wx_suma_cia07 = Far_Cost2!FAR_STOCK
     End If
     'If LK_CODCIA = "01" Then 'COSTO DE LA 01 SIEMPRE
      
     'End If
     If LK_CODCIA = "01" Then
       WS_SALDO = WS_SALDO + wx_suma_cia03
       WS_SALDO = WS_SALDO + wx_suma_cia07
     ElseIf LK_CODCIA = "03" Then
       WS_SALDO = WS_SALDO + wx_suma_cia01
       WS_SALDO = WS_SALDO + wx_suma_cia07
     ElseIf LK_CODCIA = "07" Then
       WS_SALDO = WS_SALDO + wx_suma_cia01
       WS_SALDO = WS_SALDO + wx_suma_cia03
     End If
   End If
   
   
   '''' WS_SALDO = WS_SALDO + Val(WS_SALDO_ALM)
   Do Until Far_Cost.EOF
      'If Far_Cost!far_codart = 938575 And Far_Cost!far_numfac = 7776 And Far_Cost!far_tipmov = 20 Then Stop
    '  If Far_Cost!far_numfac = 2155 Then Stop
'      If Far_Cost!far_codart = 1504505 Then Stop
      
'       Print Far_Cost!FAR_codcia, Far_Cost!FAR_fecha_compra
      ProgBar.Value = Far_Cost.AbsolutePosition
      DoEvents
      WS_TIPO_CAMBIO = 1
      If (LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07") And (Far_Cost!FAR_tipmov = 100 Or Far_Cost!FAR_tipmov = 101) Then
       If (Far_Cost!far_otra_cia) = "10" Then
            WS_SALDO = WS_SALDO + (Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad)
       End If
      Else
          WS_SALDO = WS_SALDO + (Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad)
      End If
      If LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07" Then   ' multi costo la cia
        If Far_Cost!FAR_CODCIA = "01" Then
         wx_suma_cia01 = wx_suma_cia01 + (Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad)
        ElseIf Far_Cost!FAR_CODCIA = "03" Then
         wx_suma_cia03 = wx_suma_cia03 + (Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad)
        ElseIf Far_Cost!FAR_CODCIA = "07" Then
         wx_suma_cia07 = wx_suma_cia07 + (Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad)
        End If
      End If
      If Far_Cost!FAR_MONEDA = "D" Then WS_TIPO_CAMBIO = JALAR(Far_Cost!FAR_fecha_compra, Far_Cost!FAR_fecha_compra)
      Far_Cost.Edit
      
      If Far_Cost!FAR_COSTEO = "A" And Far_Cost!far_estado <> "E" And Far_Cost!far_signo_arm = 1 And Far_Cost!FAR_tipmov = 101 And (LK_CODCIA = "04" Or LK_CODCIA = "10") Then
       '  ws_cospro = Format(Far_Cost!FAR_PRECIO / Far_Cost!FAR_equiv, "0.0000")
       '  GoTo NADA
      End If
      
      If Far_Cost!FAR_COSTEO = "A" And Far_Cost!far_estado <> "E" And Far_Cost!far_signo_arm = 1 Then
       wcan_botica = 0
        If (LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07") And chefac.Value = 1 Then
            WS_SALDO_ALM = 0
            PSFAR_COSTO_ALM(0) = "02"
            PSFAR_COSTO_ALM(1) = Far_Cost!FAR_fecha_compra
            PSFAR_COSTO_ALM(2) = Far_Cost!far_codart
            Far_Cost_alm.Requery
            Far_Cost_alm.MoveLast
            If Not Far_Cost_alm.EOF Then WS_SALDO_ALM = WS_SALDO_ALM + Val(Far_Cost_alm!FAR_STOCK)
            PSFAR_COSTO_ALM(0) = "04"
            PSFAR_COSTO_ALM(1) = Far_Cost!FAR_fecha_compra
            PSFAR_COSTO_ALM(2) = Far_Cost!far_codart
            Far_Cost_alm.Requery
            Far_Cost_alm.MoveLast
            If Not Far_Cost_alm.EOF Then WS_SALDO_ALM = WS_SALDO_ALM + Val(Far_Cost_alm!FAR_STOCK)
            PSFAR_COSTO_ALM(0) = "08"
            PSFAR_COSTO_ALM(1) = Far_Cost!FAR_fecha_compra
            PSFAR_COSTO_ALM(2) = Far_Cost!far_codart
            Far_Cost_alm.Requery
            Far_Cost_alm.MoveLast
            If Not Far_Cost_alm.EOF Then WS_SALDO_ALM = WS_SALDO_ALM + Val(Far_Cost_alm!FAR_STOCK)
         End If '
'         Print Far_Cost!far_tipmov + Far_Cost!far_numser + Far_Cost!FAR_numfac
         WPASA_COSPRO_ANT = ws_cospro
         wk_equiv = Far_Cost!FAR_equiv
         If LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07" Then
           If chefac.Value <> 1 Then
             ws_valor_ant = (((wx_suma_cia01 + wx_suma_cia03 + wx_suma_cia07) - Far_Cost!FAR_cantidad) * ws_cospro)
           Else
             ws_valor_ant = (((WS_SALDO + WS_SALDO_ALM) - Far_Cost!FAR_cantidad) * ws_cospro)
           End If
         Else
           ws_valor_ant = (WS_SALDO - Far_Cost!FAR_cantidad) * ws_cospro
         End If
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
         Far_Cost!far_precio_neto = ws_compra
         ultimo_costo = Far_Cost!FAR_PRECIO
         ultimo_cambio = Far_Cost!FAR_tipo_cambio
         If ws_valor_ant < 0 And LK_EMP = "3AA" Then
            ws_valor_ant = 0
         End If
         'Print Far_Cost!FAR_codcia
         'Print wx_suma_cia01 + wx_suma_cia03
         
         If LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07" Then
            If WS_SALDO <> 0 Then
              If chefac.Value <> 1 Then
                 ws_cospro = (ws_valor_ant + ws_compra) / (wx_suma_cia01 + wx_suma_cia03 + wx_suma_cia07)
              Else
                 ws_cospro = (ws_valor_ant + ws_compra) / (WS_SALDO + WS_SALDO_ALM)
              End If
            Else
                ws_cospro = ((ws_valor_ant + ws_compra) / Far_Cost!FAR_cantidad)
            End If
         Else
           If WS_SALDO <> 0 Then
               ws_cospro = (ws_valor_ant + ws_compra) / (WS_SALDO)
           Else
              ws_cospro = ((ws_valor_ant + ws_compra) / Far_Cost!FAR_cantidad)
           End If
         End If
  '       If ws_cospro < 0 Then Stop
          GoSub reg_hiscos
         If Far_Cost!FAR_COSPRO <> ws_cospro Then
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
         End If
      End If
NADA:
   'If Val(Far_Cost!FAR_COSPRO) <> Val(ws_cospro) Then Stop
      Far_Cost!far_cospro_sup = WCOSPRO_SUP
      Far_Cost!FAR_COSPRO = ws_cospro
      Far_Cost!FAR_COSPRO_ANT = WPASA_COSPRO_ANT
      
      WCOSPRO_SUP = Far_Cost!FAR_COSPRO
      If WS_TIPO_CAMBIO <> 1 Then Far_Cost!FAR_tipo_cambio = WS_TIPO_CAMBIO
      If Val(Far_Cost!FAR_STOCK) <> Val(WS_SALDO) Then
          Print "ACTUALIZA FAR_STOCK"
      End If
      If LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07" Then
         If Far_Cost!FAR_CODCIA = "01" Then
            Far_Cost!FAR_STOCK = wx_suma_cia01
         ElseIf Far_Cost!FAR_CODCIA = "03" Then
            Far_Cost!FAR_STOCK = wx_suma_cia03
         ElseIf Far_Cost!FAR_CODCIA = "07" Then
            Far_Cost!FAR_STOCK = wx_suma_cia07
         End If
      Else
        Far_Cost!FAR_STOCK = WS_SALDO
      End If
   
      Far_Cost.Update
      Far_Cost.MoveNext
Loop
        POR(2).Caption = "Actualizando Costos de Articulo. . ."
        If flag_costeo = "A" Then
        Else
          actualizar_precio art_LLAVE!art_key, ws_cospro, arm_llave!ARM_COSPRO
        End If
        If LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07" Then
            SQ_OPER = 1
            pu_codcia = "02"
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update
            pu_codcia = "03"
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update
            pu_codcia = "07"
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update
            pu_codcia = "01"
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update
            
            pu_codcia = "04"
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update
            
            pu_codcia = "08"
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update
            
            pu_codcia = LK_CODCIA
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update

        End If
        
        arm_llave.Edit
        arm_llave!ARM_COSPRO = ws_cospro
        If ultimo_costo <> 0 Then
          arm_llave!ARM_COSTO_ULT = redondea(ultimo_costo)
        End If
        arm_llave.Update
        
        ' TEMPORAL
        '----------
        If chelives.Value = 1 Then
          ACTLIVES = "PRE_COSTO = (" & ws_cospro & "* PRE_EQUIV)"
          pub_cadena = "UPDATE PRECIOS SET " & ACTLIVES & " WHERE  PRE_CODCIA in ('03','07') AND PRE_CODART = " & art_LLAVE!art_key
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

If flag_costeo = "A" Then
Else
 cabeprecios
 Muestra_Preicios 0
 fralista.Visible = True
End If

        


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
 
 'MsgBox " P r o c e s o   d e   C o s t e o ", 48, Pub_Titulo ' & Chr(13) & WTEXTO, vbInformation, Pub_Titulo
 
'End If
POR(2).Caption = ""
If flag_costeo = "A" Then
Else
 MsgBox "Proceso Terminado.", vbInformation, Pub_Titulo
End If
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
       Unload Frmcospro
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
Exit Sub
' REGISTRA HISTORIAL DE COSTOS
'=============================
reg_hiscos:

pub_cadena = "SELECT * FROM TABESTADOS WHERE TAE_CODCIA = '" & Far_Cost!FAR_CODCIA & "' AND TAE_TIPMOV = '" & Far_Cost!FAR_tipmov & "'  AND TAE_NUMSER = '" & Far_Cost!far_numser & "' AND TAE_NUMFAC = '" & Far_Cost!far_numfac & "' AND TAE_NUMSEC = '" & Far_Cost!far_numsec & "'"
Set rs_regcos = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not rs_regcos.EOF Then
    rs_regcos.Edit
    rs_regcos!tae_codusu = LK_CODUSU
    rs_regcos!tae_hora = Format(Now, "hh:mm:ss AMPM")
    rs_regcos!TAE_COSPRO = ws_cospro
    rs_regcos!TAE_COSPRO_ANT = Far_Cost!FAR_COSPRO
    If Val(Far_Cost!FAR_COSPRO) <> 0 Then
      wvardif = 100 - Abs((ws_cospro * 100) / Val(Far_Cost!FAR_COSPRO))
      If Val(Abs(wvardif)) >= Val(wvalorVari) Then
      rs_regcos!tae_estado = "B"
      Else
      rs_regcos!tae_estado = "A"
      End If
    Else
      rs_regcos!tae_estado = "B"
    End If
    rs_regcos.Update
End If

Return


End Sub

Public Sub Costeo_Chepen()
Dim wx_suma_cia01 As Currency
Dim wx_suma_cia03 As Currency
Dim wsa_botica As Currency
Dim wcan_botica As Currency
Dim WS_SALDO_ALM As Currency
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
Dim Far_Cost_alm As rdoResultset
Dim PSFAR_COSTO_ALM As rdoQuery

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
WS_SALDO_ALM = 0
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

If flag_costeo = "A" Then
Else
    pub_mensaje = "Proceso de Costeo de Articulos de la Fecha : " & Chr(13) & " " & Format(txtfecha.Text, "dd/mm/yyyy") & " al " & Format(LK_FECHA_DIA, "dd/mm/yyyy") & Chr(13) & "ï¿½ Continuar? ..."
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
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

pub_cadena = "SELECT FAR_FECHA_COMPRA, FAR_FECHA_CAN,FAR_COSPRO_SUP, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_STOCK, FAR_COSPRO FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA < ? AND FAR_CODART = ? and far_estado <> 'E' and far_estado2 <> 'L' and far_transito <> 'P' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_TIPMOV, FAR_NUMOPER2"
Set PSFAR_COSTO2 = CN.CreateQuery("", pub_cadena)
PSFAR_COSTO2(0) = 0
PSFAR_COSTO2(1) = LK_FECHA_DIA
PSFAR_COSTO2(2) = 0
Set Far_Cost2 = PSFAR_COSTO2.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT FAR_FECHA_COMPRA, FAR_FECHA_CAN,FAR_COSPRO_SUP, FAR_CANTIDAD, FAR_SIGNO_ARM, FAR_STOCK, FAR_COSPRO FROM FACART WHERE FAR_CODCIA IN (?) AND FAR_FECHA_COMPRA < ? AND FAR_CODART = ? and far_estado <>'E' and far_estado2 <>'L' and far_transito <> 'P' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_TIPMOV, FAR_NUMOPER2"
Set PSFAR_COSTO_ALM = CN.CreateQuery("", pub_cadena)
PSFAR_COSTO_ALM(0) = 0
PSFAR_COSTO_ALM(1) = LK_FECHA_DIA
PSFAR_COSTO_ALM(2) = 0
Set Far_Cost_alm = PSFAR_COSTO_ALM.OpenResultset(rdOpenKeyset, rdConcurValues)



'pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_FECHA_COMPRA >= ? AND  FAR_FECHA_COMPRA <= ?   and far_estado<>'E' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_NUMOPER2"


pub_cadena = "SELECT FAR_CODCIA, FAR_NUM_LOTE,FAR_PRECIO_NETO,FAR_COSPRO_ANT,FAR_COSPRO_SUP, FAR_STOCK , FAR_TIPO_CAMBIO, FAR_COSTEO_REAL, FAR_NUMGUIA, FAR_NUMSER, FAR_NUMFAC, FAR_OTRA_CIA, FAR_CODART, FAR_COSPRO, FAR_FLETE, FAR_BRUTO, FAR_DESCTO, FAR_GASTOS, FAR_TOT_DESCTO, FAR_SIGNO_ARM, FAR_CANTIDAD, FAR_TIPMOV, FAR_MONEDA, FAR_FECHA_COMPRA, FAR_FECHA_CAN, FAR_COSTEO, FAR_ESTADO, FAR_PRECIO, FAR_EQUIV  FROM FACART WHERE FAR_CODCIA IN (?,?) AND FAR_CODART = ? AND FAR_FECHA_COMPRA >= ? AND  FAR_FECHA_COMPRA <= ?   and far_estado <>'E' and far_estado2 <>'L' and far_transito <> 'P' ORDER BY FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC ,FAR_TIPMOV, FAR_NUMOPER2,FAR_CODCIA"
Set PSFAR_COSTO = CN.CreateQuery("", pub_cadena)
If LK_CODCIA = "01" Or LK_CODCIA = "03" Then
PSFAR_COSTO(0) = "01"
PSFAR_COSTO(1) = "03"
Else
PSFAR_COSTO(0) = LK_CODCIA
PSFAR_COSTO(1) = "-1"
End If
PSFAR_COSTO(2) = 0
PSFAR_COSTO(3) = LK_FECHA_DIA
PSFAR_COSTO(4) = LK_FECHA_DIA
Set Far_Cost = PSFAR_COSTO.OpenResultset(rdOpenKeyset, rdConcurValues)

'ï¿½DistinctRow
'DISTINCTROW
If chefac.Value = 1 Then
   PSFAR_LLAVE(0) = Val(Left(tipmov.Text, 3))
   PSFAR_LLAVE(1) = LK_CODCIA
   PSFAR_LLAVE(2) = Val(txtserie.Text)
   PSFAR_LLAVE(3) = " "
   PSFAR_LLAVE(4) = Val(txtnumero.Text)
   far_llave.Requery
   If far_llave.EOF Then
        MsgBox "Este Documentono Existe", 48, Pub_Titulo
        ProgBar.Visible = False
        Command1.Enabled = True
        Exit Sub
   Else
      If Format(txtfecha, "dd/mm/yyyy") <> Format(far_llave!FAR_fecha_compra, "dd/mm/yyyy") Then
        MsgBox "La Fecha de Emisión del Documento no Coincide con la Fecha Inicial del Costeo " & Chr(13) & "Fecha de Emisión del Documento : " & Format(far_llave!FAR_fecha_compra, "dd/mm/yyyy"), 48, Pub_Titulo
        txtfecha.Text = Format(far_llave!FAR_fecha_compra, "dd/mm/yyyy")
        ProgBar.Visible = False
        Command1.Enabled = True
        Exit Sub
      End If
   End If
 pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ?  AND  FAR_ESTADO <> 'E' AND  FAR_ESTADO <> 'L' AND ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' )  and far_tipmov = " & Val(Left(tipmov.Text, 3)) & " and far_numser = " & Val(txtserie.Text) & " and far_numfac = " & txtnumero.Text & " ORDER BY FAR_CODART"
Else
 pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ?  AND  FAR_ESTADO <> 'E' AND FAR_ESTADO <> 'L' AND FAR_CODART <> 0  ORDER BY FAR_CODART " ' (FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' )  ORDER BY FAR_CODART"
End If
If CHEART.Value = 1 Then
  pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND FAR_CODART = " & WCODART & "  AND  FAR_ESTADO <> 'E' AND FAR_ESTADO2 <> 'L' AND FAR_CODART <> 0 ORDER BY FAR_CODART" ' AND ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' ) ORDER BY FAR_CODART"
End If
If Val(Right(TLINEA.Text, 8)) <> 0 Then
  WLINEA = Val(Right(TLINEA.Text, 8))
  pub_cadena = "SELECT Distinct FAR_CODART FROM FACART, ARTI WHERE (FAR_CODART = ART_KEY) AND (FAR_CODCIA = ART_CODCIA) AND FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND ARTI.ART_FAMILIA = " & WLINEA & "  AND  FAR_ESTADO <> 'E' AND FAR_ESTADO2 <> 'L' AND FAR_CODART <> 0 ORDER BY FAR_CODART" ' AND ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' ) ORDER BY FAR_CODART"
End If
' TEMPORAL
''''pub_cadena = "SELECT Distinct FAR_CODART FROM FACART, ARTI WHERE (ART_CODCIA = FAR_CODCIA) AND (FAR_CODART = ART_KEY) AND  FAR_TIPMOV = 6 AND FAR_NUM_LOTE <> 10 AND FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND ( FAR_COSTEO_REAL = 'A'  OR FAR_FECHA_COMPRA <> FAR_FECHA  OR FAR_MONEDA= 'D' )  AND FAR_ESTADO<>'E'  ORDER BY FAR_CODART"
Set PSFAR_COMPR = CN.CreateQuery("", pub_cadena)
PSFAR_COMPR(0) = LK_CODCIA
PSFAR_COMPR(1) = Date
PSFAR_COMPR(2) = Date
Set Far_Compr = PSFAR_COMPR.OpenResultset(rdOpenKeyset, rdConcurValues)
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
   Unload Frmcospro
   Unload FORMGEN
   Exit Sub
End If
'OPEN_LOG "Proceso de Costeo"
gridcabe.Visible = False
POR(2).Caption = "Actualizando Costos. . ."
DoEvents
wCODARTI = Far_Compr!far_codart

flagcito = "A"
WPASA_COSPRO_ANT = 0
WCOSPRO_SUP = 0
Do Until Far_Compr.EOF
'   If Far_Compr!far_codart = 938575 Then Stop
   If Val(txtava.Text) <> 0 Then
      If Far_Compr.AbsolutePosition < Val(txtava.Text) Then GoTo OTRO_ART
   End If
   ProgBar.Value = Far_Compr.AbsolutePosition
   POR(2).Caption = "Costeando... " & ProgBar.Value & " de " & ProgBar.max
   DoEvents
   
  ' PSFAR_COSTO.rdoParameters(0) = LK_CODCIA
   PSFAR_COSTO.rdoParameters(2) = Far_Compr!far_codart
   PSFAR_COSTO.rdoParameters(3) = CDate(txtfecha.Text)
   If fechafinal.Visible Then
     PSFAR_COSTO.rdoParameters(4) = fechafinal.Text
   Else
     PSFAR_COSTO.rdoParameters(4) = LK_FECHA_DIA
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
      'WCOSPRO_SUP = Nulo_Valor0(Far_Cost!FAR_COSPRO_SUP)
      'ws_cospro = Nulo_Valor0(Far_Cost!FAR_COSPRO_SUP)
   Else
      WCOSPRO_SUP = Nulo_Valor0(Far_Cost!FAR_COSPRO)
      ws_cospro = Nulo_Valor0(Far_Cost!FAR_COSPRO)
   End If
   wx_suma_cia03 = 0
'   If Far_Cost!far_codart = 1096119 Then Stop
   If LK_CODCIA = "01" Or LK_CODCIA = "03" Then
           If LK_CODCIA = "01" Then
             PSFAR_COSTO2(0) = "03"
           Else
             PSFAR_COSTO2(0) = "01"
           End If
           PSFAR_COSTO2(1) = Format(txtfecha.Text, "dd/mm/yyyy")
           PSFAR_COSTO2(2) = Far_Cost!far_codart
           Far_Cost2.Requery
           If LK_CODCIA = "01" Then
               wx_suma_cia03 = 0
            '   If Far_Cost!far_codart = 1338647 Then Stop
               Far_Cost2.MoveLast
               If Far_Cost2.EOF Then
                 wx_suma_cia03 = 0
               Else
                 wx_suma_cia03 = Far_Cost2!FAR_STOCK
               End If
           Else
               wx_suma_cia01 = 0
            '   If Far_Cost!far_codart = 1338647 Then Stop
               Far_Cost2.MoveLast
               If Far_Cost2.EOF Then
                 wx_suma_cia01 = 0
               Else
                 wx_suma_cia01 = Far_Cost2!FAR_STOCK
                 WCOSPRO_SUP = Far_Cost2!FAR_COSPRO 'COSTO DE LA 01 SIEMPRE
                 ws_cospro = Far_Cost2!FAR_COSPRO
               End If
           End If
   End If
      
   PSFAR_COSTO2(0) = LK_CODCIA
   PSFAR_COSTO2(1) = Format(txtfecha.Text, "dd/mm/yyyy")
   PSFAR_COSTO2(2) = Far_Cost!far_codart
   Far_Cost2.Requery
   WS_SALDO_ALM = 0
'   If Far_Cost!far_codart = 1338647 Then Stop
   Far_Cost2.MoveLast
   If Far_Cost2.EOF Then
     WS_SALDO = 0
   Else
     WS_SALDO = Far_Cost2!FAR_STOCK
     If LK_CODCIA = "01" Then
       wx_suma_cia01 = Far_Cost2!FAR_STOCK
       If WS_SALDO = 0 And ws_cospro <> 0 Then
       Else
        WCOSPRO_SUP = Far_Cost2!FAR_COSPRO
        ws_cospro = Far_Cost2!FAR_COSPRO
      End If
     Else
       wx_suma_cia03 = Far_Cost2!FAR_STOCK
     End If
     'If LK_CODCIA = "01" Then 'COSTO DE LA 01 SIEMPRE
      
     'End If
     If LK_CODCIA = "01" Then
       WS_SALDO = WS_SALDO + wx_suma_cia03
     Else
       WS_SALDO = WS_SALDO '+ wx_suma_cia01
     End If
   End If
   
   
   '''' WS_SALDO = WS_SALDO + Val(WS_SALDO_ALM)
   Do Until Far_Cost.EOF
      'If Far_Cost!far_codart = 938575 And Far_Cost!far_numfac = 7776 And Far_Cost!far_tipmov = 20 Then Stop
      ProgBar.Value = Far_Cost.AbsolutePosition
      DoEvents
      WS_TIPO_CAMBIO = 1
      If (LK_CODCIA = "01" Or LK_CODCIA = "03") And (Far_Cost!FAR_tipmov = 100 Or Far_Cost!FAR_tipmov = 101) Then
       If (Far_Cost!far_otra_cia) = "10" Then
            WS_SALDO = WS_SALDO + (Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad)
       End If
      Else
          WS_SALDO = WS_SALDO + (Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad)
      End If
      If LK_CODCIA = "01" Or LK_CODCIA = "03" Then  ' multi costo la cia
        If Far_Cost!FAR_CODCIA = "01" Then
         wx_suma_cia01 = wx_suma_cia01 + (Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad)
        ElseIf Far_Cost!FAR_CODCIA = "03" Then
         wx_suma_cia03 = wx_suma_cia03 + (Far_Cost!far_signo_arm * Far_Cost!FAR_cantidad)
        End If
      End If
      If Far_Cost!FAR_MONEDA = "D" Then WS_TIPO_CAMBIO = JALAR(Far_Cost!FAR_fecha_compra, Far_Cost!FAR_fecha_compra)
      Far_Cost.Edit
      
      If Far_Cost!FAR_COSTEO = "A" And Far_Cost!far_estado <> "E" And Far_Cost!far_signo_arm = 1 And Far_Cost!FAR_tipmov = 101 And (LK_CODCIA = "04" Or LK_CODCIA = "10") Then
         ws_cospro = Format(Far_Cost!FAR_PRECIO / Far_Cost!FAR_equiv, "0.0000")
         GoTo NADA
      End If
      
      If Far_Cost!FAR_COSTEO = "A" And Far_Cost!far_estado <> "E" And Far_Cost!far_signo_arm = 1 Then
       wcan_botica = 0
        If (LK_CODCIA = "01" Or LK_CODCIA = "03") And chefac.Value = 1 Then
            WS_SALDO_ALM = 0
            PSFAR_COSTO_ALM(0) = "02"
            PSFAR_COSTO_ALM(1) = Far_Cost!FAR_fecha_compra
            PSFAR_COSTO_ALM(2) = Far_Cost!far_codart
            Far_Cost_alm.Requery
            Far_Cost_alm.MoveLast
            If Not Far_Cost_alm.EOF Then WS_SALDO_ALM = WS_SALDO_ALM + Val(Far_Cost_alm!FAR_STOCK)
            PSFAR_COSTO_ALM(0) = "04"
            PSFAR_COSTO_ALM(1) = Far_Cost!FAR_fecha_compra
            PSFAR_COSTO_ALM(2) = Far_Cost!far_codart
            Far_Cost_alm.Requery
            Far_Cost_alm.MoveLast
            If Not Far_Cost_alm.EOF Then WS_SALDO_ALM = WS_SALDO_ALM + Val(Far_Cost_alm!FAR_STOCK)
            PSFAR_COSTO_ALM(0) = "08"
            PSFAR_COSTO_ALM(1) = Far_Cost!FAR_fecha_compra
            PSFAR_COSTO_ALM(2) = Far_Cost!far_codart
            Far_Cost_alm.Requery
            Far_Cost_alm.MoveLast
            If Not Far_Cost_alm.EOF Then WS_SALDO_ALM = WS_SALDO_ALM + Val(Far_Cost_alm!FAR_STOCK)
         End If
'         Print Far_Cost!far_tipmov + Far_Cost!far_numser + Far_Cost!FAR_numfac
         WPASA_COSPRO_ANT = ws_cospro
         wk_equiv = Far_Cost!FAR_equiv
         If LK_CODCIA = "01" Or LK_CODCIA = "03" Then
            'ws_valor_ant = (((WS_SALDO + WS_SALDO_ALM) - Far_Cost!far_cantidad) * ws_cospro)
            ws_valor_ant = (((WS_SALDO + WS_SALDO_ALM) - Far_Cost!FAR_cantidad) * ws_cospro)
         Else
           ws_valor_ant = (WS_SALDO - Far_Cost!FAR_cantidad) * ws_cospro
         End If
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
         Far_Cost!far_precio_neto = ws_compra
         ultimo_costo = Far_Cost!FAR_PRECIO
         ultimo_cambio = Far_Cost!FAR_tipo_cambio
         If ws_valor_ant < 0 And LK_EMP = "3AA" Then
            ws_valor_ant = 0
         End If
         Print Far_Cost!FAR_CODCIA
         If WS_SALDO <> 0 Then
            ws_cospro = (ws_valor_ant + ws_compra) / (WS_SALDO) '+ WS_SALDO_ALM)
         Else
            ws_cospro = ((ws_valor_ant + ws_compra) / Far_Cost!FAR_cantidad)
         End If
         If Far_Cost!FAR_COSPRO <> ws_cospro Then
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
         End If
      End If
NADA:
   'If Val(Far_Cost!FAR_COSPRO) <> Val(ws_cospro) Then Stop
      Far_Cost!far_cospro_sup = WCOSPRO_SUP
      Far_Cost!FAR_COSPRO = ws_cospro
      Far_Cost!FAR_COSPRO_ANT = WPASA_COSPRO_ANT
      
      WCOSPRO_SUP = Far_Cost!FAR_COSPRO
      If WS_TIPO_CAMBIO <> 1 Then Far_Cost!FAR_tipo_cambio = WS_TIPO_CAMBIO
      If Val(Far_Cost!FAR_STOCK) <> Val(WS_SALDO) Then
          Print "ACTUALIZA FAR_STOCK"
      End If
      If LK_CODCIA = "01" Or LK_CODCIA = "03" Then
         If Far_Cost!FAR_CODCIA = "01" Then
            Far_Cost!FAR_STOCK = wx_suma_cia01
         ElseIf Far_Cost!FAR_CODCIA = "03" Then
            Far_Cost!FAR_STOCK = wx_suma_cia03
         End If
      Else
        Far_Cost!FAR_STOCK = WS_SALDO
      End If
   
      Far_Cost.Update
      Far_Cost.MoveNext
Loop
        POR(2).Caption = "Actualizando Costos de Articulo. . ."
        If flag_costeo = "A" Then
        Else
          actualizar_precio art_LLAVE!art_key, ws_cospro, arm_llave!ARM_COSPRO
        End If
        If LK_CODCIA = "01" Or LK_CODCIA = "03" Then
            SQ_OPER = 1
            pu_codcia = "02"
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update
            pu_codcia = "03"
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update
            
            pu_codcia = "01"
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update
            
            pu_codcia = "04"
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update
            
            pu_codcia = "08"
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update
            
            pu_codcia = LK_CODCIA
            PUB_CODART = art_LLAVE!art_key
            LEER_ARM_LLAVE
            arm_llave.Edit
            arm_llave!ARM_COSPRO = ws_cospro
            arm_llave.Update

        End If
        
        arm_llave.Edit
        arm_llave!ARM_COSPRO = ws_cospro
        If ultimo_costo <> 0 Then
          arm_llave!ARM_COSTO_ULT = redondea(ultimo_costo)
        End If
        arm_llave.Update
        
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

If flag_costeo = "A" Then
Else
 cabeprecios
 Muestra_Preicios 0
 fralista.Visible = True
End If

        


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
 
 'MsgBox " P r o c e s o   d e   C o s t e o ", 48, Pub_Titulo ' & Chr(13) & WTEXTO, vbInformation, Pub_Titulo
 
'End If
POR(2).Caption = ""
If flag_costeo = "A" Then
Else
 MsgBox "Proceso Terminado.", vbInformation, Pub_Titulo
End If
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
       Unload Frmcospro
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


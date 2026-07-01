VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRegStock 
   Caption         =   "Inventario Automatica"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fralotes 
      Caption         =   "Informacion de Lote"
      Height          =   5535
      Left            =   5280
      TabIndex        =   30
      Tag             =   "9999"
      Top             =   480
      Visible         =   0   'False
      Width           =   5700
      Begin VB.CommandButton cmdltcancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   2190
         TabIndex        =   34
         Tag             =   "9999"
         Top             =   5070
         Width           =   1200
      End
      Begin VB.CommandButton cmdltaceptar 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   645
         TabIndex        =   33
         Tag             =   "9999"
         Top             =   5085
         Width           =   1455
      End
      Begin VB.TextBox textolote 
         Height          =   315
         Left            =   3345
         TabIndex        =   32
         Top             =   4845
         Visible         =   0   'False
         Width           =   780
      End
      Begin MSFlexGridLib.MSFlexGrid det_lot 
         Height          =   4020
         Left            =   5040
         TabIndex        =   31
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
      Begin MSFlexGridLib.MSFlexGrid gridlt 
         Height          =   3645
         Left            =   90
         TabIndex        =   35
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
         TabIndex        =   43
         Tag             =   "9999"
         Top             =   240
         Width           =   960
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
         TabIndex        =   42
         Tag             =   "9999"
         Top             =   255
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
         TabIndex        =   41
         Tag             =   "9999"
         Top             =   570
         Width           =   4350
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
         TabIndex        =   40
         Tag             =   "9999"
         Top             =   570
         Width           =   1170
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
         TabIndex        =   39
         Tag             =   "9999"
         Top             =   960
         Width           =   1590
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
         TabIndex        =   38
         Tag             =   "9999"
         Top             =   960
         Width           =   1785
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
         TabIndex        =   37
         Tag             =   "9999"
         Top             =   960
         Width           =   1275
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
         TabIndex        =   36
         Tag             =   "9999"
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame fraultimos 
      Caption         =   "Ultimos Ajustes Realizados :"
      Height          =   735
      Left            =   240
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ComboBox cmbajustes 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   360
         Width           =   3375
      End
   End
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   10935
      Begin VB.ComboBox cmbfami 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   4935
      End
      Begin VB.CommandButton cmdmostrar 
         Caption         =   "Mostrar Lista.."
         Height          =   375
         Left            =   6720
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "La Cantidades  se Actualiza automaticamente "
         Height          =   615
         Left            =   9240
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblra 
         Caption         =   "Fecha:(F1=Buscar)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblra 
         Caption         =   "Selecione Lines :"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle de Productos :"
      Height          =   4935
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   11055
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   3720
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
         MultiLine       =   0   'False
         TextRTF         =   $"frmRegStock.frx":0000
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
      Begin MSFlexGridLib.MSFlexGrid gridarti 
         Height          =   4605
         Left            =   120
         TabIndex        =   3
         Tag             =   "9999"
         Top             =   240
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   8123
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "                                                              Opciones de Registro de Stock :"
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   6120
      Width           =   11055
      Begin VB.CommandButton cmbcal 
         Caption         =   "< Calcular."
         Height          =   375
         Left            =   2880
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdRegoper 
         Caption         =   "&Generar  Registro de Diferencias"
         Height          =   495
         Left            =   4200
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblfechareg 
         BackColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   7080
         TabIndex        =   25
         Top             =   480
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Proc:"
         Height          =   255
         Index           =   4
         Left            =   7080
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblvaltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Left            =   1440
         TabIndex        =   23
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Resultado Valor  :"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label lblvalsob 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label lblvalfal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.00"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1440
         TabIndex        =   20
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Sobrante :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Faltante :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label lbldocingresos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0-0"
         Height          =   195
         Left            =   9960
         TabIndex        =   15
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lbldocsalidas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0-0"
         Height          =   195
         Left            =   9960
         TabIndex        =   14
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label5 
         Caption         =   "Doc. Ajuste de Ingresos :"
         Height          =   255
         Left            =   8160
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Doc. Ajuste de Salidas   :"
         Height          =   255
         Index           =   0
         Left            =   8160
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblreg 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Se Genero el Registro."
         Height          =   435
         Left            =   6120
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Label Label1 
      Caption         =   "REGISTRO AUTOMATICO DE INVENTARIO "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmRegStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_ev0 As Currency
Dim loc_ev1 As Currency
Dim loc_ev_activa As Currency
Dim loc_ev2 As Currency
Dim ParaLot_count As Integer
Dim ParaLot_codlot(100) As String
Dim ParaLot_lotcant(100) As Currency
Dim ParaLot_fechalot(100) As String


Private Sub cmbajustes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 txtfecha.Text = cmbajustes.Text
 fraultimos.Visible = False
 txtfecha_KeyPress 13
End If
If KeyAscii = 27 Then
  fraultimos.Visible = False
End If
End Sub

Private Sub cmbajustes_LostFocus()
  fraultimos.Visible = False
End Sub

Private Sub cmbcal_Click()
If ProBar.Visible = False Then sum_total
End Sub

Private Sub cmbfami_Click()
gridarti.Clear
gridarti.Cols = 12
gridarti.TextMatrix(0, 0) = "Item"
gridarti.TextMatrix(0, 1) = "Descripcion del Producto"
gridarti.TextMatrix(0, 2) = "Unid."
gridarti.TextMatrix(0, 2) = "Activa"
gridarti.TextMatrix(0, 3) = "St.Sist" ' SISTEMA GUARDADO
gridarti.TextMatrix(1, 3) = "Cajas." ' SISTEMA GUARDADO
gridarti.TextMatrix(0, 4) = "St.Fisic"
gridarti.TextMatrix(1, 4) = "Cajas"
gridarti.TextMatrix(0, 5) = "Stock "
gridarti.TextMatrix(1, 5) = "Difer."

gridarti.TextMatrix(0, 6) = "St.Sist" ' SISTEMA GUARDADO
gridarti.TextMatrix(1, 6) = "Unidad"  ' SISTEMA GUARDADO
gridarti.TextMatrix(0, 7) = "St.Fisic"
gridarti.TextMatrix(1, 7) = "Unidad"
gridarti.TextMatrix(0, 8) = "Stock "
gridarti.TextMatrix(1, 8) = "Difer."

gridarti.ColWidth(0) = 300
gridarti.ColWidth(1) = 3500
gridarti.ColWidth(2) = 600
gridarti.ColWidth(3) = 800
gridarti.ColWidth(4) = 800
gridarti.ColWidth(5) = 800
gridarti.ColWidth(6) = 800
gridarti.ColWidth(7) = 800
gridarti.ColWidth(8) = 800
gridarti.ColWidth(9) = 0
gridarti.ColWidth(10) = 0
gridarti.ColWidth(11) = 0

gridarti.Rows = 2

End Sub

Private Sub cmdmostrar_Click()
Dim wflag_reg As String
Dim ws_count As Currency
Dim ps_autoinv As rdoResultset
Dim xpinta As Integer
Dim xfil  As Integer
Dim WARTI As rdoQuery
Dim wllave_arti As rdoResultset

Dim WBOT As rdoQuery
Dim wllave_bot As rdoResultset

Dim wvalor As Integer
Dim CADE_ROT As String
If Val(Right(cmbfami.Text, 8)) = 0 Then
  MsgBox "Selecione la Familia / Muestre colocando la fecha y pulse Enter", 48, Pub_Titulo
  Exit Sub
End If
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


pub_cadena = "SELECT ARM_COSPRO, ART_STOCK_MAX , PRE_COSTO_ANT, ART_KEY , art_alterno, art_nombre, PRE_EQUIV ,PRE_PRE6 , PRE_UNIDAD, ARM_STOCK , ((ARM_STOCK/PRE_EQUIV) - (ART_STOCK_MAX/PRE_EQUIV)) AS DIF , ART_PLANCHA FROM ARTI, articulo, precios WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART ) AND PRE_FLAG_UNIDAD = 'A' AND ART_CODCIA = ? AND ART_FAMILIA = ?  AND ART_SITUACION <> '1' ORDER BY ART_NOMBRE"
Set WARTI = CN.CreateQuery("", pub_cadena)
WARTI(0) = LK_CODCIA
WARTI(1) = Val(Right(cmbfami.Text, 8))
Set wllave_arti = WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)
wllave_arti.Requery
If Not wllave_arti.EOF Then
 ProBar.Min = 0
 ProBar.max = wllave_arti.RowCount
 ProBar.Value = 0
End If
ProBar.Visible = True
gridarti.Visible = False
DoEvents
gridarti.Clear
gridarti.Cols = 14
gridarti.TextMatrix(0, 0) = "Item"
gridarti.TextMatrix(0, 1) = "Descripcion del Producto"
gridarti.TextMatrix(0, 2) = "Unid."
gridarti.TextMatrix(0, 2) = "Activa"
gridarti.TextMatrix(0, 3) = "St.Sist" ' SISTEMA GUARDADO
gridarti.TextMatrix(1, 3) = "Activa." ' SISTEMA GUARDADO
gridarti.TextMatrix(0, 4) = "St.Fisic"
gridarti.TextMatrix(1, 4) = "Cajas"
gridarti.TextMatrix(0, 5) = "Stock "
gridarti.TextMatrix(1, 5) = "Difer."

gridarti.TextMatrix(0, 6) = "St.Sist" ' SISTEMA GUARDADO
gridarti.TextMatrix(1, 6) = "Minima"  ' SISTEMA GUARDADO
gridarti.TextMatrix(0, 7) = "St.Fisic"
gridarti.TextMatrix(1, 7) = "Minima"
gridarti.TextMatrix(0, 8) = "Stock "
gridarti.TextMatrix(1, 8) = "Difer."
gridarti.TextMatrix(0, 12) = "Valor "
gridarti.TextMatrix(1, 12) = "Ajuste."

gridarti.ColWidth(0) = 300
gridarti.ColWidth(1) = 3500
gridarti.ColWidth(2) = 600
gridarti.ColWidth(3) = 800
gridarti.ColWidth(4) = 800
gridarti.ColWidth(5) = 800
gridarti.ColWidth(6) = 800
gridarti.ColWidth(7) = 800
gridarti.ColWidth(8) = 800
gridarti.ColWidth(9) = 0
gridarti.ColWidth(10) = 0
gridarti.ColWidth(11) = 0
gridarti.ColWidth(12) = 1200
gridarti.ColWidth(13) = 0  '  Para almacenar la equiv Activa

gridarti.Rows = 2


xfil = 1
Do Until wllave_arti.EOF
  ProBar.Value = ProBar.Value + 1
  DoEvents
  ' sec 0
  pub_cadena = "SELECT * FROM AUTOINV WHERE INV_CODCIA = '" & LK_CODCIA & "' AND INV_FECHA = '" & Format(txtfecha.Text, "dd/mm/yyyy") & "' AND INV_CODART = " & wllave_arti!art_key
  Set ps_autoinv = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
  If ps_autoinv.EOF Then
    MsgBox "Producto se a agregado despues del cuadro de stock, No entra para el ajuste." & Chr(13) & Trim(wllave_arti!ART_NOMBRE), 48, Pub_Titulo
    GoTo sale_otro
  End If
  xfil = xfil + 1
  gridarti.Rows = gridarti.Rows + 1
  gridarti.RowHeight(gridarti.Rows - 1) = 285
'  lisarti.AddItem wllave_arti!art_alterno & " " & wllave_arti!ART_NOMBRE & String(120, " ") & wllave_arti!ART_KEY
  gridarti.TextMatrix(xfil, 0) = Format(xfil - 1, "00")
  gridarti.TextMatrix(xfil, 1) = Trim(wllave_arti!ART_NOMBRE)
  gridarti.TextMatrix(xfil, 2) = Trim(wllave_arti!pre_unidad) & "(" & (Format(wllave_arti!pre_equiv, "0")) & ")"
  
  ' sec 1
  ws_count = cant_equiv(Val(wllave_arti!art_key))
  gridarti.Row = xfil
  gridarti.COL = 3
  gridarti.CellForeColor = QBColor(1)
  gridarti.COL = 6
  gridarti.CellForeColor = QBColor(1)
  gridarti.COL = 5
  gridarti.CellForeColor = vbRed
  gridarti.COL = 8
  gridarti.CellForeColor = vbRed
  
  
 ' If ws_count = 1 Or ws_count = 3 Then
'    gridarti.Visible = True
    gridarti.TextMatrix(xfil, 3) = Format(Int(ps_autoinv!inv_stockactual / wllave_arti!pre_equiv), "0.00")
    gridarti.TextMatrix(xfil, 4) = Format(ps_autoinv!INV_F_CAJA, "0.00")
 ' ElseIf ws_count = 2 Then
 '   gridarti.TextMatrix(xfil, 3) = Format(Int(ps_autoinv!inv_stockactual / loc_ev1), "0.00")
    gridarti.TextMatrix(xfil, 6) = Format(Val(ps_autoinv!inv_stockactual) - (Int(Val(ps_autoinv!inv_stockactual) / wllave_arti!pre_equiv) * wllave_arti!pre_equiv), "0.00")
    
    'gridarti.TextMatrix(xfil, 4) = Format(Int(ps_autoinv!INV_F_CAJA), "0.00")
    gridarti.TextMatrix(xfil, 7) = Format(Int(ps_autoinv!INV_F_UNIDAD), "0.00")
  'End If
  gridarti.TextMatrix(xfil, 9) = ws_count
  gridarti.TextMatrix(xfil, 10) = wllave_arti!art_key
  gridarti.TextMatrix(xfil, 11) = txtfecha.Text
  gridarti.TextMatrix(xfil, 13) = wllave_arti!pre_equiv
  ESTABLECE_DIF xfil
  If Not IsNull(ps_autoinv!INV_REG_FECHA) Then
    mst_fecha = Format(ps_autoinv!INV_REG_FECHA, "dd/mm/yyyy")
    mst_salidas = Format(ps_autoinv!INV_REG_NS5, "0") & " - " & Format(ps_autoinv!INV_REG_NF5, "0")
    mst_ingresos = Format(ps_autoinv!INV_REG_NS6, "0") & " - " & Format(ps_autoinv!INV_REG_NF6, "0")
  End If
  wflag_reg = ps_autoinv!INV_FLAG_REG
sale_otro:
wllave_arti.MoveNext
Loop
lbldocsalidas.Caption = mst_salidas
lbldocingresos.Caption = mst_ingresos
lblfechareg.Caption = mst_fecha
If Val(wflag_reg) = 1 Then
lblreg.Visible = True
Else
lblreg.Visible = False
End If
sum_total
On Error GoTo sale
gridarti.Visible = True
If gridarti.Rows > 2 Then
  gridarti.Row = 2
  gridarti.COL = 4
  gridarti.SetFocus
End If
ProBar.Visible = False
DoEvents
sale:
End Sub


Private Sub cmdRegoper_Click()
cmdRegoper.Enabled = False
Dim ws_numoper6 As Integer
Dim ws_numoper5 As Integer
Dim f_reg_tipmov5 As String * 1
Dim f_reg_tipmov6 As String * 1
Dim flag_reg As String * 1
Dim WS_CODTRA As Integer
Dim ws_linea As String
Dim WS_TOTAL As Currency
Dim WS_SUBTOTAL As Currency
Dim ws_descri As String
Dim ws_equiv As Currency
Dim WDIF_CAJ As Currency
Dim WDIF_UNI As Currency
Dim ws_count As Currency
Dim WS_NUM_MAS  As Currency
Dim WS_NUM_MENOS As Currency
cmbcal_Click
If lblreg.Visible = True Then
    cmdRegoper.Enabled = True
    MsgBox "OJO!! ya genero el proceso !!!!" & Chr(13) & "Necesita clave de Gerencia", 48, Pub_Titulo
    LK_ACCESO_REPORT = ""
    Load frmclave2
    Screen.MousePointer = 0
    frmclave2.Show 1
    If LK_ACCESO_REPORT <> "A" Then
       Exit Sub
    End If
End If
pub_mensaje = "Generar el Registro de Regularización Automatico de Inventario...   ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
  cmdRegoper.Enabled = True
  Exit Sub
End If



'*******************
' FACART
'*******************
On Error GoTo error_fatal
pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)

WS_CODTRA = 2403
SQ_OPER = 3
PU_TIPMOV = 6  ' AUMENTO
pu_codcia = LK_CODCIA
PU_NUMSER = 99
PU_FBG = " "
LEER_FAR_LLAVE
If Not far_menor.EOF Then
   WS_NUM_MAS = Val(far_menor!far_numfac) + 1
Else
   WS_NUM_MAS = 1
End If
SQ_OPER = 3
PU_TIPMOV = 5  ' DESMINUCION
pu_codcia = LK_CODCIA
PU_NUMSER = 88
PU_FBG = " "
LEER_FAR_LLAVE
If Not far_menor.EOF Then
   WS_NUM_MENOS = Val(far_menor!far_numfac) + 1
Else
  WS_NUM_MENOS = 1
End If
ws_linea = Trim(Left(cmbfami.Text, 15))


CN.Execute "Begin Transaction", rdExecDirect

GoSub manda_numero

PUB_NUM_OPER = PUB_NUM_OPER_XXX
ws_numoper6 = PUB_NUM_OPER
ws_numoper5 = PUB_NUM_OPER + 1


WS_NUMSEC = 0
flag_reg = ""
f_reg_tipmov6 = ""
f_reg_tipmov5 = ""
 
For fila = 2 To gridarti.Rows - 1
  ws_count = cant_equiv(Val(gridarti.TextMatrix(fila, 10)))
 ' If Val(gridarti.TextMatrix(fila, 9)) = 1 Or Val(gridarti.TextMatrix(fila, 9)) > 2 Then
 '   WS_CANTIDAD = Abs(Val(gridarti.TextMatrix(fila, 5)))
 '   If Val(gridarti.TextMatrix(fila, 5)) > 0 Then
 '    pub_signo_arm = 1
 '   ElseIf Val(gridarti.TextMatrix(fila, 5)) < 0 Then
 '    pub_signo_arm = -1
 '   End If
 ' ElseIf Val(gridarti.TextMatrix(fila, 9)) = 2 Then
    WDIF_CAJ = Val(gridarti.TextMatrix(fila, 5)) * Val(gridarti.TextMatrix(fila, 13)) ' CONVERTIDO TODO EN UNIDADES
    WDIF_UNI = Val(gridarti.TextMatrix(fila, 8))
    WS_CANTIDAD = WDIF_CAJ + WDIF_UNI ' TODO EN UNIDADES
    If WS_CANTIDAD > 0 Then
      pub_signo_arm = 1
    ElseIf WS_CANTIDAD < 0 Then
      pub_signo_arm = -1
    End If
    WS_CANTIDAD = Abs(WS_CANTIDAD)
  'End If
  If pub_signo_arm = 1 Then
    PUB_NUMSER = 99
    PUB_TIPMOV = 6
    PUB_NUMFAC = WS_NUM_MAS
    WS_TOTAL = Abs(Format(Val(lblvalsob.Caption), "0.00")) '
    PUB_NUM_OPER = ws_numoper6
  Else
    WS_TOTAL = Abs(Format(Val(lblvalfal.Caption), "0.00")) '
    PUB_NUMSER = 88
    PUB_TIPMOV = 5
    PUB_NUM_OPER = ws_numoper5
    PUB_NUMFAC = WS_NUM_MENOS
  End If
  If WS_CANTIDAD = 0 Then GoTo salta_registrof
  If PUB_TIPMOV = 6 Then f_reg_tipmov6 = "A"
  If PUB_TIPMOV = 5 Then f_reg_tipmov5 = "A"
  WS_SUBTOTAL = Abs(Val(gridarti.TextMatrix(fila, 12)))
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_CODART = Val(gridarti.TextMatrix(fila, 10))
  LEER_ARM_LLAVE
  arm_llave.Edit
  arm_llave!arm_stock = arm_llave!arm_stock + (WS_CANTIDAD * pub_signo_arm)
  arm_llave.Update
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_SECUEN = 0
  LEER_PRE_LLAVE
  If pre_llave.EOF Then
   ws_descri = ""
  Else
   ws_descri = pre_llave!pre_unidad
  End If
  ws_equiv = 1
  GoSub REG_FACART
  flag_reg = "A"
salta_registrof:
 ' MsgBox gridarti.TextMatrix(fila, 1) & Chr(13) & "SIGNO: " & pub_signo_arm & "  CANTIDAD : " & WS_CANTIDAD & " EQUIVS : " & loc_ev0 & " " & loc_ev1 & " " & loc_ev2
Next fila
If f_reg_tipmov5 = "A" Then
 pub_cadena = "UPDATE AUTOINV SET INV_REG_FECHA = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "' , INV_REG_NS5 = 88 , INV_REG_NF5 = " & WS_NUM_MENOS & " , INV_FLAG_REG = 1 WHERE INV_CODCIA = '" & LK_CODCIA & "'  AND INV_FAMILIA = " & Trim(Right(cmbfami.Text, 8)) & "  AND INV_FECHA = '" & Format(txtfecha.Text, "dd/mm/yyyy") & "'"
 CN.Execute pub_cadena, rdExecDirect
 If flag_reg = "A" Then ' SOLO CUANDO SE REGISTRE EN FACART
    WS_TOTAL = Abs(Format(Val(lblvalfal.Caption), "0.00")) '
    PUB_NUMSER = 88
    PUB_TIPMOV = 5
    PUB_NUM_OPER = ws_numoper5
    PUB_NUMFAC = WS_NUM_MENOS
    pub_signo_arm = -1
    GoSub REG_ALLOG
  End If
End If
If f_reg_tipmov6 = "A" Then
  pub_cadena = "UPDATE AUTOINV SET INV_REG_FECHA = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "' , INV_REG_NS6 = 99 , INV_REG_NF6= " & WS_NUM_MAS & " , INV_FLAG_REG = 1 WHERE INV_CODCIA = '" & LK_CODCIA & "'  AND INV_FAMILIA = " & Trim(Right(cmbfami.Text, 8)) & "  AND INV_FECHA = '" & Format(txtfecha.Text, "dd/mm/yyyy") & "'"
  CN.Execute pub_cadena, rdExecDirect
  If flag_reg = "A" Then ' SOLO CUANDO SE REGISTRE EN FACART
    PUB_NUMSER = 99
    PUB_TIPMOV = 6
    PUB_NUMFAC = WS_NUM_MAS
    WS_TOTAL = Abs(Format(Val(lblvalsob.Caption), "0.00")) '
    PUB_NUM_OPER = ws_numoper6
    pub_signo_arm = 1
    GoSub REG_ALLOG
  End If
End If
      

CN.Execute "Commit Transaction", rdExecDirect
con_llave.Close
cmdRegoper.Enabled = True
If flag_reg = "A" Then ' SOLO CUANDO SE REGISTRE EN FACART
   cmdmostrar_Click
   MsgBox "Registro Automatico Efectuado", 48, Pub_Titulo
Else
   MsgBox "No se encontro ninguna diferencia.", 48, Pub_Titulo
End If

Exit Sub

'_____________
REG_FACART:
'-------------
      far_llave.AddNew
      far_llave!FAR_TIPMOV = PUB_TIPMOV
      far_llave!FAR_CODCIA = LK_CODCIA
      far_llave!far_cod_sunat = 0 'Val(Right(i_codsunat.Text, 5))
      far_llave!far_numser = PUB_NUMSER
      far_llave!FAR_CODVEN = 0
      far_llave!far_numfac = PUB_NUMFAC
      WS_NUMSEC = WS_NUMSEC + 1
      far_llave!FAR_NUMSEC = WS_NUMSEC
      far_llave!FAR_STOCK = 0
      far_llave!far_codart = PUB_CODART
      far_llave!far_cantidad = WS_CANTIDAD
      far_llave!far_cantidad_p = WS_CANTIDAD
       ' AFECTANDO A LOTES
       '******************
      PUB_CODLOT = BUSCA_LOTE_AFECTAR(Val(PUB_CODART), pub_signo_arm, Val(WS_CANTIDAD))     ' BUSCA UN LOTE PARA AFECTAR EL AJUSTE
      PSLOT_LLAVE(0) = LK_CODCIA
      PSLOT_LLAVE(1) = arm_llave!ARM_CODART
      PSLOT_LLAVE(2) = PUB_CODLOT
      lot_llave.Requery
      If lot_llave.EOF Then
       lot_llave.AddNew
       lot_llave!LOT_CODCIA = LK_CODCIA
       lot_llave!LOT_NROLOTE = PUB_CODLOT
       lot_llave!lot_codart = arm_llave!ARM_CODART
       lot_llave!lot_fecha_vcto = LK_FECHA_DIA
       lot_llave!LOT_SALDOS = 0
      Else
       lot_llave.Edit
      End If
      lot_llave!lot_codclie = 0
      lot_llave!LOT_SALDOS = lot_llave!LOT_SALDOS + (WS_CANTIDAD * pub_signo_arm)
      lot_llave.Update
      
      far_llave!far_codlot = PUB_CODLOT
      far_llave!FAR_PRECIO = arm_llave!ARM_COSPRO
      far_llave!FAR_equiv = ws_equiv
      far_llave!far_descri = ws_descri 'Trim(xl.Cells(WS_FILA, 6))
      far_llave!far_PESO = 0
      far_llave!far_signo_car = 0
      far_llave!far_signo_arm = pub_signo_arm
      far_llave!far_codclie = 0
      far_llave!FAR_MONEDA = "S"
      far_llave!FAR_EX_IGV = 0
      far_llave!FAR_cp = " "
      far_llave!FAR_fecha_compra = LK_FECHA_DIA
      far_llave!far_estado = "N"
      far_llave!FAR_ESTADO2 = "N"
      far_llave!FAR_COSPRO = arm_llave!ARM_COSPRO
      far_llave!FAR_COSPRO_ANT = 0
      far_llave!far_fbg = " "
      far_llave!far_impto = 0
      far_llave!FAR_TOT_FLETE = 0
      far_llave!FAR_FLETE = 0
      far_llave!FAR_DESCTO = 0
      far_llave!FAR_TOT_DESCTO = 0
      far_llave!FAR_GASTOS = 0
      far_llave!far_bruto = WS_TOTAL
      far_llave!FAR_NUMDOC = 0
      far_llave!far_numguia = 0
      far_llave!far_serguia = 0
      far_llave!FAR_pordescto1 = 0
      far_llave!FAR_costeo = ""
      far_llave!FAR_COSTEO_REAL = ""
      far_llave!FAR_tipo_cambio = 1
      far_llave!FAR_DIAS = 0
      far_llave!FAR_fecha = LK_FECHA_DIA
      far_llave!FAR_NUMSER_C = 0
      far_llave!FAR_NUMFAC_C = 0
      far_llave!FAR_NUMOPER = PUB_NUM_OPER
      far_llave!FAR_NUMOPER2 = PUB_NUM_OPER
      far_llave!far_precio_neto = 0
      far_llave!far_otra_cia = " "
      far_llave!far_transito = " "
      far_llave!far_subtra = "Ajuste Automtico"
      far_llave!far_otra_cia = " "
      far_llave!far_transito = " "
      far_llave!far_JABAS = 0
      far_llave!far_UNIDADES = 0
      far_llave!far_mortal = 0
      far_llave!far_num_precio = 0
      far_llave!FAR_ORDEN_UNIDADES = 0
      far_llave!FAR_SUBTOTAL = WS_SUBTOTAL
      far_llave!far_turno = 0
      far_llave!far_concepto = " "
      far_llave!far_concepto = "Aj. Autom. - : " & ws_linea
      far_llave!far_codusu = LK_CODUSU
      far_llave!far_hora = Format(Now, "hh:mm:ss AMPM")
      If PUB_TIPMOV = 5 Then
        far_llave!FAR_NUM_LOTE = 1
      Else
        far_llave!FAR_NUM_LOTE = 0
      End If

      far_llave!FAR_PEDSER = 0
      far_llave!FAR_PEDFAC = 0
      far_llave!far_pedsec = 0
      far_llave!FAR_TIPDOC = ""
      far_llave!FAR_fecha_can = LK_FECHA_DIA
      far_llave!FAR_fecha_pro = LK_FECHA_DIA
      far_llave!far_key_dircli = 0
      far_llave.Update
Return

'*******************
' REGISTRO DE ALLOG
'*******************



'*******************
REG_ALLOG:
'*******************
    
all_llave.AddNew
all_llave!ALL_NUMOPER = PUB_NUM_OPER
all_llave!ALL_NUMOPER2 = PUB_NUM_OPER
all_llave!all_CODCIA = LK_CODCIA
all_llave!ALL_codtra = WS_CODTRA
all_llave!all_flag_ext = "N"
all_llave!ALL_CODCLIE = pu_codclie
all_llave!ALL_CODART = 0
all_llave!ALL_IMPORTE_amort = 0
all_llave!all_codusu = LK_CODUSU
all_llave!ALL_FBG = ""
all_llave!ALL_CODVEN = 0
all_llave!ALL_IMPORTE = WS_TOTAL ' llave_rep01!par_saldo_caja_hoy
all_llave!ALL_IMPORTE_DOLL = 0 'llave_rep01!PAR_SALDO_CAJA_D_HOY
all_llave!ALL_NUMDOC = 0
all_llave!ALL_CP = " "
all_llave!ALL_TIPDOC = ""
all_llave!all_numfac_c = 0
all_llave!all_numser_c = 0
all_llave!all_codban = 0
all_llave!all_concepto = "Ajuste Automtico"
all_llave!all_chenum = 0
all_llave!ALL_FECHA_DIA = LK_FECHA_DIA
all_llave!ALL_FECHA_SUNAT = LK_FECHA_DIA
all_llave!ALL_FECHA_VCTO = LK_FECHA_DIA
all_llave!ALL_CANTIDAD = WS_CANTIDAD
all_llave!ALL_NUMSER = PUB_NUMSER
all_llave!all_numfac = PUB_NUMFAC
all_llave!all_neto = 0
all_llave!ALL_BRUTO = 0
all_llave!ALL_tipmov = PUB_TIPMOV
all_llave!ALL_IMPTO = 0
all_llave!ALL_flete = 0
all_llave!ALL_HORA = Now
all_llave!ALL_DESCTO = 0
all_llave!ALL_GASTOS = 0
all_llave!ALL_PRECIO = 0
all_llave!ALL_MONEDA_CLI = ""
all_llave!ALL_moneda_ccm = ""
all_llave!ALL_MONEDA_CAJA = ""
all_llave!all_SECUENCIA = 0
all_llave!ALL_SIGNO_CAR = 0
all_llave!ALL_signo_caja = 0
all_llave!ALL_SIGNO_CCM = 0
all_llave!all_sIGNO_ARM = pub_signo_arm
all_llave!all_chenum = 0
all_llave!ALL_CHESEC = 0
all_llave!ALL_CHESER = 0
all_llave!ALL_SUBTRA = ""
all_llave!ALL_TIPO_BLOQ_ACT = 0
all_llave!ALL_TIPO_BLOQ_ANT = 0
all_llave!all_codtra_ext = 0
all_llave!ALL_TIPO_CAMBIO = 0
all_llave!ALL_RUC = 0
all_llave!ALL_IMPG1 = 0
all_llave!ALL_IMPG2 = 0
all_llave!ALL_GASTOS = 0
all_llave.Update
Return
Exit Sub
error_fatal:
CN.Execute "Rollback Transaction", rdExecDirect
con_llave.Close
MsgBox Err.Description

Exit Sub

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
Return

End Sub

Private Sub Form_Load()
txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtfecha.Mask = "##/##/####"
If IsDate(txtfecha.Text) Then
  RUTINA_LLENA_FAM CDate(txtfecha.Text)
End If
CenterMe frmRegStock
End Sub

Public Sub RUTINA_LLENA_FAM(ws_fecha As Date)
Dim WS_CADEN As String
Dim sp_familia As rdoResultset
Dim sp_cheingreso As rdoResultset

pub_cadena = "SELECT distinct inv_familia FROM AUTOINV WHERE INV_CODCIA = '" & LK_CODCIA & "' and inv_fecha = '" & Format(ws_fecha, "dd/mm/yyyy") & "'"
Set sp_familia = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
sp_familia.Requery
ProBar.Visible = True

If Not sp_familia.EOF Then
  ProBar.max = sp_familia.RowCount
  ProBar.Min = 0
  ProBar.Value = 0
End If
cmbfami.Clear
Do Until sp_familia.EOF
  ProBar.Value = sp_familia.AbsolutePosition
  DoEvents
  PUB_TIPREG = 122
  PUB_CODCIA = LK_CODCIA
  PUB_NUMTAB = sp_familia!INV_FAMILIA
  SQ_OPER = 1
  LEER_TAB_LLAVE
  If Not tab_llave.EOF Then
    pub_cadena = "SELECT SUM(INV_F_CAJA) AS RESCAJA , SUM(INV_F_UNIDAD) AS RESUNIDAD, MAX(INV_FLAG_REG) AS VAL_FLAG FROM AUTOINV WHERE INV_CODCIA = '" & LK_CODCIA & "' and inv_fecha = '" & Format(ws_fecha, "dd/mm/yyyy") & "' AND INV_FAMILIA = " & PUB_NUMTAB
    Set sp_cheingreso = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    WS_CADEN = ""
    If Not sp_cheingreso.EOF Then
      If Val(Nulo_Valor0(sp_cheingreso!resunidad)) <> 0 Or Val(Nulo_Valor0(sp_cheingreso!resCAJA)) <> 0 Then
         WS_CADEN = "Ok, Ingresado"
      End If
      If Val(Nulo_Valor0(sp_cheingreso!VAL_FLAG)) <> 0 Then
       WS_CADEN = WS_CADEN + "  -  Registrado "
      End If
    End If
    cmbfami.AddItem tab_llave!tab_nomlargo & " " & WS_CADEN & String(80, " ") & Trim(CStr(tab_llave!tab_numtab))
  End If
  sp_familia.MoveNext
Loop
ProBar.Visible = False

End Sub


Private Sub gridarti_GotFocus()
TEXTOVAR.Visible = False
End Sub

Private Sub gridarti_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Dim car As String
Static CONS
'If KeyAscii <> 13 Then Exit Sub
' solo numeros
   If gridarti.COL = 4 Or gridarti.COL = 7 Then
   Else
   Exit Sub
   End If
   If gridarti.COL = 7 Then
     'If Val(gridarti.TextMatrix(gridarti.Row, 9)) = 1 Or Val(gridarti.TextMatrix(gridarti.Row, 9)) > 2 Then
     If Val(gridarti.TextMatrix(gridarti.Row, 13)) = 1 Then
       MsgBox "No procede este producto solo registra unidades minimoas", 48, Pub_Titulo
       Exit Sub
     End If
   End If
   
   
   car = Chr(KeyAscii)
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." And car <> "-" Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If
    TEXTOVAR.Left = gridarti.Left + gridarti.CellLeft
    TEXTOVAR.Width = gridarti.CellWidth
    TEXTOVAR.Height = gridarti.CellHeight
    TEXTOVAR.Top = gridarti.Top + gridarti.CellTop
    TEXTOVAR.Text = gridarti.TextMatrix(gridarti.Row, gridarti.COL)
    TEXTOVAR.Visible = True
    Azul3 TEXTOVAR, TEXTOVAR
    TEXTOVAR.Text = Trim(car)
    TEXTOVAR.SelStart = Len(TEXTOVAR)
    TEXTOVAR.SetFocus
'End If
End Sub

Private Sub gridarti_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim sale As Boolean
Dim Wsec

'If WMODO = "C" Then Exit Sub

'If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
'End If
If KeyCode = 45 Then
  CONFIGAR_LOTES LK_CODCIA, Val(gridarti.TextMatrix(gridarti.Row, 10)), Val(gridarti.TextMatrix(gridarti.Row, 4)), Left(gridarti.TextMatrix(gridarti.Row, 5), 10), Val(Right(gridarti.TextMatrix(gridarti.Row, 5), 8)), gridarti.Row
  PASA_BOT_ACEPTAR
  Exit Sub
End If


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
TEXTOVAR.Visible = False
End Sub


Private Sub textovar_Change()
gridarti.Text = Format(TEXTOVAR.Text, "0.00")
End Sub

Private Sub TEXTOVAR_GotFocus()
 temporal = gridarti.TextMatrix(gridarti.Row, gridarti.COL)
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
Dim ps_autoinv As rdoResultset
If KeyAscii = 27 Then
  TEXTOVAR.Text = temporal
  TEXTOVAR.Visible = False
  gridarti.SetFocus
  Exit Sub
End If
If gridarti.COL = 4 Then Consistencias gridarti, TEXTOVAR, KeyAscii
If gridarti.COL = 7 Then Consistencias gridarti, TEXTOVAR, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If
' actualizar datos
'================

pub_cadena = "SELECT * FROM AUTOINV WHERE INV_CODCIA = '" & LK_CODCIA & "' AND INV_FECHA = '" & Format(gridarti.TextMatrix(gridarti.Row, 11), "dd/mm/yyyy") & "' AND INV_CODART = " & Val(gridarti.TextMatrix(gridarti.Row, 10))
Set ps_autoinv = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not ps_autoinv.EOF Then
  ps_autoinv.Edit
  ps_autoinv!INV_F_CAJA = Val(gridarti.TextMatrix(gridarti.Row, 4))
  ps_autoinv!INV_F_UNIDAD = Val(gridarti.TextMatrix(gridarti.Row, 7))
  ps_autoinv.Update
End If
ESTABLECE_DIF gridarti.Row

'================
If gridarti.Row >= gridarti.Rows - 1 And gridarti.COL <> 4 Then
 ESTABLECE_DIF gridarti.Row
 gridarti.SetFocus
 TEXTOVAR.Visible = False
Else
  If gridarti.COL = 4 Then
     If Val(gridarti.TextMatrix(gridarti.Row, 9)) = 1 Then
     'If Val(gridarti.TextMatrix(gridarti.Row, 7)) = 1 Then
     '  GoTo sala_lin
     End If
     gridarti.COL = 7
  ElseIf gridarti.COL = 7 Then
sala_lin:
   If gridarti.Row >= gridarti.Rows - 1 Then
   Else
     gridarti.Row = gridarti.Row + 1
     gridarti.COL = 4
   End If
  End If
  ESTABLECE_DIF gridarti.Row
  gridarti.SetFocus
  TEXTOVAR.Visible = False
End If

fin:

End Sub


Private Sub txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If IsDate(txtfecha.Text) Then
    RUTINA_LLENA_FAM CDate(txtfecha.Text)
  End If
End If
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


Public Function cant_equiv(wsp_codart As Currency) As Integer
Dim rs_precios As rdoResultset
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND PRE_CODART = " & wsp_codart & " "
Set rs_precios = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
cant_equiv = rs_precios.RowCount
loc_ev0 = rs_precios!pre_equiv
If rs_precios!pre_FLAG_UNIDAD = "A" Then loc_ev_activa = loc_ev0
rs_precios.MoveNext
If rs_precios.EOF Then GoTo SALIR
loc_ev1 = rs_precios!pre_equiv
If rs_precios!pre_FLAG_UNIDAD = "A" Then loc_ev_activa = loc_ev1
rs_precios.MoveNext
If rs_precios.EOF Then GoTo SALIR
loc_ev2 = rs_precios!pre_equiv
If rs_precios!pre_FLAG_UNIDAD = "A" Then loc_ev_activa = loc_ev2
SALIR:
End Function

Public Sub ESTABLECE_DIF(wfila)

 If (Val(gridarti.TextMatrix(wfila, 4)) - Val(gridarti.TextMatrix(wfila, 3))) = 0 Then
  gridarti.TextMatrix(wfila, 5) = ""
 Else
  gridarti.TextMatrix(wfila, 5) = Val(gridarti.TextMatrix(wfila, 4)) - Val(gridarti.TextMatrix(wfila, 3))
 End If
 If (Val(gridarti.TextMatrix(wfila, 7)) - Val(gridarti.TextMatrix(wfila, 6))) = 0 Then
   gridarti.TextMatrix(wfila, 8) = ""
 Else
   gridarti.TextMatrix(wfila, 8) = Val(gridarti.TextMatrix(wfila, 7)) - Val(gridarti.TextMatrix(wfila, 6))
 End If

End Sub

Public Function BUSCA_LOTE_AFECTAR(WS_CODART As Currency, WS_SIGNO As Integer, wlot_cantidad As Currency) As String
  PSLOT_LLAVE2(0) = LK_CODCIA
  PSLOT_LLAVE2(1) = WS_CODART
  lot_llave2.Requery
  Do Until lot_llave2.EOF
   If WS_SIGNO = 1 Then
      If lot_llave2!LOT_SALDOS <= 0 Then
         BUSCA_LOTE_AFECTAR = lot_llave2!LOT_NROLOTE
         GoTo listo
      End If
   Else
      If lot_llave2!LOT_SALDOS > 0 And Val(lot_llave2!LOT_SALDOS) >= Val(wlot_cantidad) Then
         BUSCA_LOTE_AFECTAR = lot_llave2!LOT_NROLOTE
         GoTo listo
      End If
   End If
   lot_llave2.MoveNext
  Loop
 BUSCA_LOTE_AFECTAR = "(*)"
listo:
End Function

Public Sub sum_total()
Dim WS_TOTAL_MENOS  As Currency
Dim WS_TOTAL_MAS  As Currency
Dim WS_TOTAL As Currency
WS_TOTAL_MENOS = 0
WS_TOTAL_MAS = 0
On Error GoTo sigue
ProBar.max = gridarti.Rows - 1
ProBar.Value = 2
ProBar.Min = 2

ProBar.Visible = True
DoEvents
For fila = 2 To gridarti.Rows - 1
 ProBar.Value = fila
'  If fila = 8 Then Stop
 DoEvents
 ws_count = cant_equiv(Val(gridarti.TextMatrix(fila, 10)))
  If Val(gridarti.TextMatrix(fila, 9)) = 1 Or Val(gridarti.TextMatrix(fila, 9)) > 2 Then
    WS_CANTIDAD = Abs(Val(gridarti.TextMatrix(fila, 5)))
    If Val(gridarti.TextMatrix(fila, 5)) > 0 Then
     pub_signo_arm = 1
    ElseIf Val(gridarti.TextMatrix(fila, 5)) < 0 Then
     pub_signo_arm = -1
    End If
  ElseIf Val(gridarti.TextMatrix(fila, 9)) = 2 Then
    WDIF_CAJ = Val(gridarti.TextMatrix(fila, 5)) * loc_ev_activa 'loc_ev1 ' CONVERTIDO TODO EN UNIDADES
    WDIF_UNI = Val(gridarti.TextMatrix(fila, 8))
    WS_CANTIDAD = WDIF_CAJ + WDIF_UNI ' TODO EN UNIDADES
    If WS_CANTIDAD > 0 Then
      pub_signo_arm = 1
    ElseIf WS_CANTIDAD < 0 Then
      pub_signo_arm = -1
    End If
    WS_CANTIDAD = Abs(WS_CANTIDAD)
  End If
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_CODART = Val(gridarti.TextMatrix(fila, 10))
  LEER_ARM_LLAVE
  gridarti.TextMatrix(fila, 12) = Format(Val(arm_llave!ARM_COSPRO) * (WS_CANTIDAD * pub_signo_arm), "0.000")
  If pub_signo_arm = -1 Then
    WS_TOTAL_MENOS = WS_TOTAL_MENOS + Val(Format(Val(arm_llave!ARM_COSPRO) * (WS_CANTIDAD * pub_signo_arm), "0.00"))
  Else
    WS_TOTAL_MAS = WS_TOTAL_MAS + Val(Format(Val(arm_llave!ARM_COSPRO) * (WS_CANTIDAD * pub_signo_arm), "0.00"))
  End If
  WS_TOTAL = WS_TOTAL + Val(Format(Val(arm_llave!ARM_COSPRO) * (WS_CANTIDAD * pub_signo_arm), "0.00"))
Next fila
ProBar.Visible = False
lblvalfal.Caption = Format(WS_TOTAL_MENOS, "0.00")
lblvalsob.Caption = Format(WS_TOTAL_MAS, "0.00")
lblvaltotal.Caption = Format(WS_TOTAL, "0.00")
Exit Sub
sigue:
Resume Next
End Sub

Private Sub txtfecha_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wult As String
If KeyCode = 112 Then
    wult = InputBox("Ultimos :", "Ultimas Ajustes de Inventarios", "3")
    If wult = "" Then Exit Sub
    If Val(wult) <= 0 Then Exit Sub
    pub_cadena = "SELECT  top " & wult & " inv_fecha FROM AUTOINV WHERE INV_CODCIA = '" & LK_CODCIA & "'  group by inv_fecha  order by inv_fecha desc"
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    Do Until X.EOF
        cmbajustes.AddItem Format(X!inv_fecha, "dd/mm/yyyy")
    X.MoveNext
    Loop
    fraultimos.Visible = True
    If cmbajustes.ListCount > 0 Then cmbajustes.ListIndex = 0
    
    cmbajustes.SetFocus
End If

End Sub
Public Sub CONFIGAR_LOTES(LT_CODCIA As String, LT_CODART As Currency, LT_CANTIDAD As Currency, LT_DESCRIP As String, LT_EQUIV As Integer, FILA_GRID As Integer, Optional NOMUESTRA)

Dim loc_temp As Currency
Dim flagCJ As String

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
   gridlt.TextMatrix(gridlt.Rows - 1, 0) = Trim(det_lot.TextMatrix(xcta, 2)) ' NRO LOTE
   gridlt.TextMatrix(gridlt.Rows - 1, 1) = Trim(det_lot.TextMatrix(xcta, 7))  ' DESCRIP UNIDAD
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
' chequear si
flagCJ = ""
If Val(FORMGEN.grid_fac.TextMatrix(FILA_GRID, 50)) <> 1 Then GoTo IR_NORMALENTE
xcta = 1
loc_temp = LT_CANTIDAD
Do Until JALA_LOTE.EOF
   If LT_CANTIDAD <= 0 Then
       LT_SALDO = 0
       LT_CANTIDAD = 0
       GoSub muestra
       GoTo OTROCJ
   End If
   If Val(JALA_LOTE!LOT_SALDOS) >= Val(LT_CANTIDAD * LT_EQUIV) Then
     LT_CANTIDAD = LT_CANTIDAD * LT_EQUIV
     LT_SALDO = LT_CANTIDAD
     flagCJ = "A"
     GoSub muestra
     LT_CANTIDAD = 0
'     Exit Do
   Else
     LT_SALDO = 0
     LT_CANTIDAD = 0
     GoSub muestra
     LT_CANTIDAD = loc_temp
   End If
OTROCJ:
JALA_LOTE.MoveNext
Loop

If flagCJ = "A" Then
'  ' ESPECIAL
'   GoSub muestra
  GoTo mues
End If
IR_NORMALENTE:
gridlt.Rows = 2
'NORMALMENTE
JALA_LOTE.MoveFirst
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
   If LT_CANTIDAD <= JALA_LOTE!LOT_SALDOS Then
       LT_SALDO = LT_CANTIDAD
       LT_CANTIDAD = 0
       GoSub muestra
   Else
      If JALA_LOTE.RowCount = JALA_LOTE.AbsolutePosition Then
          LT_SALDO = LT_CANTIDAD  'JALA_LOTE!LOT_SALDOS
          LT_CANTIDAD = 0
      Else
          LT_SALDO = JALA_LOTE!LOT_SALDOS
        ' LT_CANTIDAD = 0
       End If
       GoSub muestra
       LT_CANTIDAD = LT_CANTIDAD - JALA_LOTE!LOT_SALDOS
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
  lblltnom.Caption = Trim(art_LLAVE!ART_NOMBRE)
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
 gridlt.TextMatrix(xcta, 0) = Trim(JALA_LOTE!LOT_NROLOTE)
 gridlt.TextMatrix(xcta, 1) = Trim(LT_DESCRIP)
 gridlt.TextMatrix(xcta, 2) = Format(Val(JALA_LOTE!LOT_SALDOS) / Val(Right(LT_EQUIV, 8)), "0.0000")
 gridlt.TextMatrix(xcta, 3) = Format(Val(LT_SALDO) / Val(Right(LT_EQUIV, 8)), "0.0000")
 gridlt.TextMatrix(xcta, 4) = Format(JALA_LOTE!lot_fecha_vcto, "dd/mm/yyyy")
 gridlt.TextMatrix(xcta, 9) = JALA_LOTE!lot_codart
 gridlt.TextMatrix(xcta, 10) = FILA_GRID
 gridlt.TextMatrix(xcta, 11) = LT_EQUIV
 gridlt.TextMatrix(xcta, 12) = Trim(JALA_LOTE!LOT_NROLOTE)

Return
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


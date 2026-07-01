VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F6E4F630-E903-11D5-8BB9-0080AD40A177}#1.18#0"; "OSControlsUser.ocx"
Begin VB.Form frmconsulotesxarti 
   Caption         =   "Consulta de Lotes por Articulo"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   ControlBox      =   0   'False
   Icon            =   "frmconsulotesxarti.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin VB.Frame fra_detalle 
      BackColor       =   &H80000004&
      Caption         =   "DETALLES DEL FACART"
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
      Height          =   7305
      Left            =   2040
      TabIndex        =   21
      Top             =   30
      Visible         =   0   'False
      Width           =   11535
      Begin VB.Frame fra_nromovimientos 
         BackColor       =   &H80000004&
         Height          =   555
         Left            =   165
         TabIndex        =   36
         Top             =   915
         Width           =   3240
         Begin VB.ComboBox cmb_nroitems 
            Height          =   315
            ItemData        =   "frmconsulotesxarti.frx":000C
            Left            =   840
            List            =   "frmconsulotesxarti.frx":0022
            TabIndex        =   24
            Text            =   "5"
            Top             =   165
            Width           =   765
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Movimientos"
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
            Index           =   6
            Left            =   1725
            TabIndex        =   38
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Ultimos "
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
            Index           =   5
            Left            =   60
            TabIndex        =   37
            Top             =   210
            Width           =   705
         End
      End
      Begin VB.CommandButton cmdocultar 
         Cancel          =   -1  'True
         Caption         =   "Volver"
         Height          =   330
         Left            =   9945
         TabIndex        =   29
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdfinddetallefecha 
         Caption         =   "Consultar"
         Height          =   330
         Left            =   8730
         TabIndex        =   27
         Top             =   360
         Width           =   1185
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000004&
         Height          =   570
         Left            =   165
         TabIndex        =   35
         Top             =   255
         Width           =   5430
         Begin VB.OptionButton opt_order 
            BackColor       =   &H80000004&
            Caption         =   "Mostrar por Rango de Fechas"
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   1
            Left            =   2745
            TabIndex        =   23
            Top             =   225
            Width           =   2535
         End
         Begin VB.OptionButton opt_order 
            BackColor       =   &H80000004&
            Caption         =   "Mostrar Ultimos Movimientos"
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   22
            Top             =   225
            Value           =   -1  'True
            Width           =   2370
         End
      End
      Begin VB.Frame fra_fecha 
         BackColor       =   &H80000004&
         Height          =   555
         Left            =   165
         TabIndex        =   32
         Top             =   915
         Visible         =   0   'False
         Width           =   4905
         Begin OSControlsUser.ctlMaskEdBox txtfechac1 
            Height          =   255
            Left            =   1920
            TabIndex        =   25
            Top             =   210
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   450
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
         Begin OSControlsUser.ctlMaskEdBox txtfechac2 
            Height          =   255
            Left            =   3660
            TabIndex        =   26
            Top             =   210
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   450
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "al"
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
            Index           =   4
            Left            =   3195
            TabIndex        =   34
            Top             =   225
            Width           =   225
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Compra del "
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
            Index           =   3
            Left            =   75
            TabIndex        =   33
            Top             =   210
            Width           =   1845
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd_detalle 
         Height          =   4740
         Left            =   150
         TabIndex        =   28
         Top             =   1605
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   8361
         _Version        =   393216
         ForeColor       =   0
         Rows            =   3
         Cols            =   11
         FixedRows       =   2
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483634
         GridColor       =   -2147483647
         GridColorFixed  =   8388608
         GridLinesFixed  =   1
         GridLinesUnpopulated=   3
         AllowUserResizing=   1
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   11
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ce&rrar"
      Height          =   420
      Left            =   120
      TabIndex        =   42
      Top             =   6570
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   585
      Left            =   6225
      TabIndex        =   39
      Top             =   255
      Width           =   2445
      Begin VB.OptionButton Option1 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   41
         Top             =   225
         Width           =   885
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   1050
         TabIndex        =   40
         Top             =   225
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin OSControlsUser.OSFindItem txtfindarti 
      Height          =   285
      Left            =   135
      TabIndex        =   7
      Top             =   1065
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      Locked          =   0   'False
   End
   Begin VB.Frame fraFindArt 
      Caption         =   "ARTICULO"
      ForeColor       =   &H00808000&
      Height          =   720
      Left            =   45
      TabIndex        =   30
      Top             =   810
      Visible         =   0   'False
      Width           =   7770
      Begin VB.Label txtarticulo 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1275
         TabIndex        =   31
         Top             =   240
         Width           =   6195
      End
   End
   Begin VB.Frame Frame2 
      Height          =   435
      Left            =   8115
      TabIndex        =   19
      Top             =   1050
      Width           =   3465
      Begin MSComctlLib.ProgressBar BPro 
         Height          =   270
         Left            =   30
         TabIndex        =   20
         Top             =   120
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame fra_lotes 
      Height          =   4905
      Left            =   5160
      TabIndex        =   17
      Top             =   1605
      Width           =   6375
      Begin VB.TextBox txtNroLote 
         Height          =   285
         Left            =   1290
         TabIndex        =   8
         Top             =   180
         Width           =   1290
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd_lotes 
         Height          =   4335
         Left            =   120
         TabIndex        =   9
         Top             =   525
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   8
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin VB.Label Label1 
         Caption         =   "Nro. Lote"
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
         Index           =   2
         Left            =   105
         TabIndex        =   18
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.Frame fratipofindart 
      Height          =   615
      Left            =   45
      TabIndex        =   16
      Top             =   105
      Width           =   5025
      Begin VB.OptionButton optTipoBusqueda 
         Caption         =   "Buscar por Articulo"
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
         Index           =   1
         Left            =   2415
         TabIndex        =   1
         Top             =   255
         Width           =   2070
      End
      Begin VB.OptionButton optTipoBusqueda 
         Caption         =   "Buscar por Familias"
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
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   255
         Value           =   -1  'True
         Width           =   2100
      End
   End
   Begin VB.Frame frafechaingreso 
      Height          =   615
      Left            =   5160
      TabIndex        =   13
      Top             =   105
      Width           =   5100
      Begin OSControlsUser.ctlMaskEdBox txtfecha1 
         Height          =   255
         Left            =   1965
         TabIndex        =   2
         Top             =   255
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   450
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
      Begin OSControlsUser.ctlMaskEdBox txtfecha2 
         Height          =   255
         Left            =   3705
         TabIndex        =   3
         Top             =   255
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   450
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
      Begin VB.Label Label1 
         Caption         =   "Fecha de Ingreso del "
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
         Left            =   120
         TabIndex        =   15
         Top             =   255
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "al"
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
         Index           =   1
         Left            =   3240
         TabIndex        =   14
         Top             =   270
         Width           =   225
      End
   End
   Begin VB.Frame frafindfam 
      Height          =   5700
      Left            =   45
      TabIndex        =   10
      Top             =   825
      Width           =   5025
      Begin OSControlsUser.ctlCombo cmbfamilias 
         Height          =   315
         Left            =   1215
         TabIndex        =   4
         Top             =   270
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
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
         Text            =   ""
         List0           =   "Combo1"
         ListIndex       =   -1
         List0           =   "Combo1"
      End
      Begin OSControlsUser.ctlCombo cmbsubfamilias 
         Height          =   315
         Left            =   1230
         TabIndex        =   5
         Top             =   675
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
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
         Text            =   ""
         List0           =   "Combo1"
         ListIndex       =   -1
         List0           =   "Combo1"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd_articulos 
         Height          =   4455
         Left            =   105
         TabIndex        =   6
         Top             =   1155
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   4
         GridLinesUnpopulated=   3
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Label Label2 
         Caption         =   "SubFamilia"
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
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   12
         Top             =   690
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Familia"
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
         Height          =   240
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Top             =   360
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmconsulotesxarti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim L_CODART As Long
Dim L_FAMILIA As Integer
Dim L_SUBFAMILIA As Integer
Dim L_UNIDAD As String
Dim L_EQUIV As Integer
Dim L_CODLOT As String
Dim L_SALDOLOTE As Double
Dim Query_NroMovOrFechas As Integer

Dim SQLFind As String
Dim SQLWhere As String

Private Sub cmbfamilias_Click()
    cmbfamilias_KeyPress (13)
End Sub

Private Sub cmbfamilias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        L_FAMILIA = cmbfamilias.ItemData(cmbfamilias.ListIndex)
        Load_SubFamilias (L_FAMILIA)
        SetGrid
        SetGridLotes
    End If
End Sub

Private Sub cmbsubfamilias_Click()
    cmbsubfamilias_KeyPress (13)
End Sub

Private Sub cmbsubfamilias_KeyPress(KeyAscii As Integer)
    If cmbsubfamilias.ListIndex = -1 Then cmbsubfamilias.ListIndex = 0
    If KeyAscii = 13 Then
        L_SUBFAMILIA = cmbsubfamilias.ItemData(cmbsubfamilias.ListIndex)
        Load_Articulos
    End If
End Sub

Private Sub cmdfinddetallefecha_Click()
    L_SALDOLOTE = Val(Format(grd_lotes.TextMatrix(CLng(grd_lotes.Row), 6), "0.00"))
    If Query_NroMovOrFechas = 2 Then
        If ValFechasC Then
            LoadDetalleFacart 2, L_CODLOT, txtfechac1.Text, txtfechac2.Text
        End If
    Else
        LoadDetalleFacart 1, L_CODLOT
    End If
End Sub

Private Sub cmdocultar_Click()
    fra_detalle.Visible = False
    opt_order.Item(0).Value = True
    txtfechac1.Text = "__/__/____"
    txtfechac2.Text = "__/__/____"
    SetGridDetalle
    Dis_EnabledFrames (True)
    grd_lotes.SetFocus
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Query_NroMovOrFechas = 1 'top registros
    Load_Familias
    SQLFind = "ARTI.ART_NOMBRE as Descripción,ARTI.ARt_alterno as Codigo,"
    SQLWhere = "ARTI.art_nombre like '"
    txtfindarti.TypeFind = NameField
End Sub
Private Sub Load_Lotes(ByVal CodArt As Long, ByVal fecha1 As String, ByVal fecha2 As String)
On Error GoTo ErrHandler
Dim SQL As String
Dim RS As ADODB.Recordset
Dim i As Integer
Dim sWhereFecha As String

    If txtfecha1.Text = txtfecha2.Text And txtfecha1.Text = "__/__/____" Then
        sWhereFecha = ""
    ElseIf txtfecha1.Text <> "__/__/____" And txtfecha2.Text <> "__/__/____" Then
        sWhereFecha = " AND LOTE.LOT_FECHA_ING >= ? AND LOTE.LOT_FECHA_ING <= ? "
    End If
    
    SQL = "SELECT LOTE.LOT_NROLOTE, LOTE.LOT_FECHA_ING, LOTE.LOT_FECHA_VCTO, "
 SQL = SQL + "LOTE.LOT_INGRESO, LOTE.LOT_SALIDAS, LOTE.LOT_SALDOS,CLIENTES.CLI_NOMBRE FROM LOTE "
 SQL = SQL + "LEFT JOIN CLIENTES ON "
SQL = SQL + "LOTE.LOT_CODCIA =CLIENTES.CLI_CODCIA AND "
SQL = SQL + "LOTE.LOT_CODCLIE = CLIENTES.CLI_CODCLIE " & _
          "WHERE LOTE.LOT_CODCIA = ? AND LOTE.LOT_CODART = ? " & sWhereFecha & _
          "ORDER BY LOTE.LOT_FECHA_ING DESC"
    If ValFechas Then
        LoadUnidadActiva
        If sWhereFecha <> "" Then
            Set RS = SQLStaticRecordset(SQL, LK_CODCIA, CodArt, fecha1, fecha2)
        Else
            Set RS = SQLStaticRecordset(SQL, LK_CODCIA, CodArt)
        End If
        SetGridLotes
        grd_lotes.Rows = IIf(RS.RecordCount = 0, 2, RS.RecordCount + 1)
        Do While Not RS.EOF
            i = i + 1
            grd_lotes.TextMatrix(i, 1) = RS("LOT_NROLOTE")
            grd_lotes.TextMatrix(i, 2) = IIf(IsNull(RS("Lot_Fecha_Vcto")), "", RS("Lot_Fecha_Vcto"))
            grd_lotes.TextMatrix(i, 3) = IIf(IsNull(RS("Lot_Fecha_Ing")), "", RS("Lot_Fecha_Ing"))
            grd_lotes.TextMatrix(i, 4) = Format(IIf(IsNull(RS("Lot_Ingreso")), 0, RS("Lot_Ingreso")) / L_EQUIV, "#######,##0.000")
            grd_lotes.TextMatrix(i, 5) = Format(IIf(IsNull(RS("Lot_Salidas")), 0, RS("Lot_Salidas")) / L_EQUIV, "#######,##0.000")
            grd_lotes.TextMatrix(i, 6) = Format(IIf(IsNull(RS("Lot_Saldos")), 0, RS("Lot_Saldos")) / L_EQUIV, "#######,##0.000")
            grd_lotes.TextMatrix(i, 7) = IIf(IsNull(RS("CLI_NOMBRE")), "", RS("CLI_NOMBRE"))
            RS.MoveNext
        Loop
        RS.Close
        Set RS = Nothing
    End If
    txtNroLote.Text = ""
Exit Sub
ErrHandler:
    RS.Close
    Set RS = Nothing
End Sub

Private Function LoadLoteOne(ByVal NROLOTE As String) As Boolean
On Error GoTo ErrHandler
Dim RS As ADODB.Recordset
Dim SQL As String

    LoadLoteOne = False
    SQL = "SELECT LOTE.LOT_NROLOTE, LOTE.LOT_FECHA_ING, LOTE.LOT_FECHA_VCTO, LOTE.LOT_INGRESO, LOTE.LOT_SALIDAS, LOTE.LOT_SALDOS,CLIENTES.CLI_NOMBRE " & _
          "FROM CLIENTES INNER JOIN LOTE ON CLIENTES.CLI_CODCIA = LOTE.LOT_CODCIA AND CLIENTES.CLI_CODCLIE = LOTE.LOT_CODCLIE " & _
          "WHERE LOTE.LOT_CODCIA = ? AND LOTE.LOT_CODART = ? AND LOTE.LOT_NROLOTE = ?"
          
    Set RS = SQLStaticRecordset(SQL, LK_CODCIA, L_CODART, NROLOTE)
    If Not RS.EOF Then
        SetGridLotes
        grd_lotes.TextMatrix(1, 1) = RS("LOT_NROLOTE")
        grd_lotes.TextMatrix(1, 2) = RS("Lot_Fecha_Vcto")
        grd_lotes.TextMatrix(1, 3) = RS("Lot_Fecha_Ing")
        grd_lotes.TextMatrix(1, 4) = RS("Lot_Ingreso")
        grd_lotes.TextMatrix(1, 5) = RS("Lot_Salidas")
        grd_lotes.TextMatrix(1, 6) = RS("Lot_Saldos")
        grd_lotes.TextMatrix(1, 7) = RS("CLI_NOMBRE")
        RS.Close
        Set RS = Nothing
        LoadLoteOne = True
    Else
        MsgBox "No Existe Lote", vbInformation
        txtNroLote.Text = ""
        txtNroLote.SetFocus
    End If
Exit Function
ErrHandler:
    RS.Close
    LoadLoteOne = False
    Set RS = Nothing
End Function
Private Sub LoadDetalleFacart(ByVal TipoConsul As Integer, ByVal NROLOTE As String, Optional FECHA_C1 As String, Optional FECHA_C2 As String)
On Error GoTo ErrHandler

Dim SQL As String
Dim sWhere As String
Dim i As Integer
Dim nRows As Integer
Dim TotCant As Double
Dim TotPrec As Currency
Dim RSDetalleFacart As ADODB.Recordset
Dim NroReg As Integer
Dim NroTop  As String
Dim SUMRES As Integer
Dim iCount As Integer
Dim pSalida As Currency, pIngreso As Currency, PRECIO As Currency
Dim cSaldo As Double

    SQL = "SELECT FACART.* , CLIENTES.CLI_NOMBRE " & _
          "FROM FACART LEFT OUTER JOIN CLIENTES ON FACART.FAR_CODCIA = CLIENTES.CLI_CODCIA AND FACART.FAR_CODCLIE = CLIENTES.CLI_CODCLIE "
    
    If Query_NroMovOrFechas = 1 Then
        sWhere = "WHERE (FACART.FAR_CODCIA = ?) AND (FACART.FAR_CODART = ?) AND (FACART.FAR_ESTADO <> 'E') AND (FACART.FAR_CODLOT = ?) " & _
                 "ORDER BY FAR_FECHA DESC, FAR_NUMOPER DESC"
        Set RSDetalleFacart = SQLStaticRecordset(SQL + sWhere, LK_CODCIA, L_CODART, NROLOTE)
        SUMRES = -1
        NroReg = IIf(RSDetalleFacart.RecordCount <= cmb_nroitems.Text, RSDetalleFacart.RecordCount, cmb_nroitems.Text)
        iCount = NroReg + 3
        cSaldo = L_SALDOLOTE
    Else
        sWhere = "WHERE (FACART.FAR_CODCIA = ?) AND (FACART.FAR_CODART = ?) AND (FACART.FAR_ESTADO <> 'E') AND (FACART.FAR_FECHA_COMPRA >= ?) AND (FACART.FAR_FECHA_COMPRA <= ?) AND (FACART.FAR_CODLOT = ?) " & _
                 "ORDER BY FAR_FECHA ,FAR_NUMOPER"
        Set RSDetalleFacart = SQLStaticRecordset(SQL + sWhere, LK_CODCIA, L_CODART, FECHA_C1, FECHA_C2, NROLOTE)
        SUMRES = 1
        NroReg = RSDetalleFacart.RecordCount
        iCount = 2
        cSaldo = L_SALDOLOTE - GetStockFecha(FECHA_C2)
    End If
    
    SetGridDetalle
    If RSDetalleFacart.EOF Then
        MsgBox "No Existen Movimientos"
        Exit Sub
    End If
    grd_detalle.Rows = NroReg + 3
    BPro.max = IIf(NroReg = 0, 1, NroReg)
    i = 1
    Do While (Not RSDetalleFacart.EOF) And (NroReg >= i)
        i = i + 1
        iCount = iCount + SUMRES
        BPro.Value = i - 1
        grd_detalle.TextMatrix(iCount, 0) = RSDetalleFacart("Far_Fecha")
        grd_detalle.TextMatrix(iCount, 1) = RSDetalleFacart("Far_Fecha_Compra") 'Format(RSDetalleFacart(1) / L_EQUIV, "######,##0.000")
        grd_detalle.TextMatrix(iCount, 2) = TipoDocumento(RSDetalleFacart("Far_TipMov"), RSDetalleFacart("Far_CP"), RSDetalleFacart("Far_FBG"), RSDetalleFacart("Far_NumFac"), RSDetalleFacart("Far_NumFac_C"), RSDetalleFacart("Far_NumSer_C"), RSDetalleFacart("Far_NumSer"), RSDetalleFacart("Far_NumGuia"), RSDetalleFacart("Far_TipDoc"))
        grd_detalle.TextMatrix(iCount, 3) = RSDetalleFacart("Far_SubTra")
        grd_detalle.TextMatrix(iCount, 4) = RSDetalleFacart("Far_Concepto")
            PRECIO = RSDetalleFacart!FAR_COSPRO
        If (RSDetalleFacart!FAR_TIPMOV = 20 Or RSDetalleFacart!FAR_TIPMOV = 101) And RSDetalleFacart!far_estado <> "E" Then
            PRECIO = RSDetalleFacart!far_precio_neto / RSDetalleFacart!far_cantidad_p
        End If
        If RSDetalleFacart!far_signo_arm = 1 Then
        grd_detalle.TextMatrix(iCount, 5) = Format(RSDetalleFacart!far_cantidad_p / L_EQUIV, "#####.00")
            pIngreso = grd_detalle.TextMatrix(iCount, 5) * PRECIO
        grd_detalle.TextMatrix(iCount, 8) = Format(pIngreso, "#####.00")
        ElseIf RSDetalleFacart!far_signo_arm = -1 Then
        grd_detalle.TextMatrix(iCount, 6) = Format(RSDetalleFacart!far_cantidad_p / L_EQUIV, "#####.00")
            pSalida = (RSDetalleFacart!far_cantidad_p) * PRECIO
        grd_detalle.TextMatrix(iCount, 9) = Format(pSalida, "#####.00")
        End If
        grd_detalle.TextMatrix(iCount, 10) = RSDetalleFacart("FAR_CODUSU")
        RSDetalleFacart.MoveNext
    Loop
    RSDetalleFacart.Close
    Set RSDetalleFacart = Nothing
    GetSaldos (cSaldo)
    BPro.Value = 0
Exit Sub
ErrHandler:
    MsgBox Err.Description
    RSDetalleFacart.Close
    Set RSDetalleFacart = Nothing
End Sub

Private Sub Load_Articulos()
On Error GoTo ErrHandler
Dim RS As ADODB.Recordset
Dim SQL As String
Dim i As Integer
Dim nRows As Integer

    SQL = "SELECT ARTI.ART_KEY, ARTI.ART_NOMBRE, PRECIOS.PRE_UNIDAD FROM ARTI INNER JOIN PRECIOS ON dbo.ARTI.ART_CODCIA = dbo.PRECIOS.PRE_CODCIA AND dbo.ARTI.ART_KEY = dbo.PRECIOS.PRE_CODART WHERE Art_CodCia = ? AND Art_SubFam = ? AND PRECIOS.PRE_FLAG_UNIDAD = 'A' ORDER BY Art_Nombre"
    Set RS = SQLStaticRecordset(SQL, LK_CODCIA, L_SUBFAMILIA)
    SetGrid
    nRows = RS.RecordCount
    grd_articulos.Rows = IIf(nRows = 0, 2, RS.RecordCount + 1)
    BPro.max = nRows
    Do While Not RS.EOF
        i = i + 1
        BPro.Value = i
        grd_articulos.TextMatrix(i, 1) = RS("Art_Key")
        grd_articulos.TextMatrix(i, 2) = RS("PRE_UNIDAD")
        grd_articulos.TextMatrix(i, 3) = RS("Art_Nombre")
        RS.MoveNext
    Loop
    RS.Close
    Set RS = Nothing
    BPro.Value = 0
Exit Sub
ErrHandler:
    RS.Close
    Set RS = Nothing
End Sub

Private Sub LoadArtOne(ByVal sKey As Variant)
On Error GoTo ErrHandler
Dim RS As ADODB.Recordset
Dim SQL As String
    
    SQL = "SELECT ARTI.ART_NOMBRE, ARTICULO.ARM_CODART, ARTICULO.ARM_STOCK FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA " & _
          "WHERE ARTICULO.ARM_CODCIA = ? AND ARTI.ART_ALTERNO = ?"
    Set RS = SQLStaticRecordset(SQL, LK_CODCIA, sKey)
    
    If Not RS.EOF Then
        L_CODART = RS("ARM_CODART")
        txtarticulo.Caption = RS("ART_NOMBRE")
        RS.Close
        Set RS = Nothing
        Load_Lotes L_CODART, txtfecha1.Text, txtfecha2.Text
    Else
        MsgBox "Articulo no Existe"
        txtfindarti.TEXTO = ""
    End If
Exit Sub
ErrHandler:
    RS.Close
    Set RS = Nothing
End Sub

Private Sub Load_Familias()
Dim RS As ADODB.Recordset
Dim SQL As String
On Error GoTo ErrHandler
    SQL = "SELECT Tab_NumTab, Tab_NomLargo FROM tablas WHERE tab_tipreg = ? AND TAB_CODCIA = ? ORDER BY Tab_NomLargo"
    Set RS = SQLStaticRecordset(SQL, 122, LK_CODCIA)
    cmbfamilias.Clear
    cmbfamilias.AddItem "-- SELECCIONE UNA FAMILIA --"
    cmbfamilias.ItemData(0) = -1
    Do While Not RS.EOF
        DoEvents
        cmbfamilias.AddItem RS("Tab_NomLargo")
        cmbfamilias.ItemData(cmbfamilias.NewIndex) = RS("Tab_NumTab")
        RS.MoveNext
    Loop
    cmbfamilias.ListIndex = 0
    RS.Close
    Set RS = Nothing
    Exit Sub
    RS.Close
    Set RS = Nothing
Exit Sub
ErrHandler:
    RS.Close
    Set RS = Nothing
End Sub

Private Sub Load_SubFamilias(ByVal FAMILIA As String)
Dim RS As ADODB.Recordset
Dim SQL As String
On Error GoTo ErrHandler
    SQL = "SELECT Tab_NumTab, Tab_NomLargo FROM tablas WHERE tab_tipreg=? AND tab_codart = ? AND TAB_CODCIA = ? ORDER BY Tab_NomLargo"
    Set RS = SQLStaticRecordset(SQL, 123, FAMILIA, LK_CODCIA)
    cmbsubfamilias.Clear
    cmbsubfamilias.AddItem "-- SELECCIONE UNA SUBFAMILIA --"
    cmbsubfamilias.ItemData(0) = -1
    Do While Not RS.EOF
        DoEvents
        cmbsubfamilias.AddItem RS("Tab_NomLargo")
        cmbsubfamilias.ItemData(cmbsubfamilias.NewIndex) = RS("Tab_NumTab")
        RS.MoveNext
    Loop
    cmbsubfamilias.ListIndex = 0
    RS.Close
    Set RS = Nothing
Exit Sub
ErrHandler:
    RS.Close
    Set RS = Nothing
End Sub
Private Sub LoadUnidadActiva()
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo ErrHandler
    SQL = "SELECT PRE_UNIDAD,PRE_EQUIV FROM PRECIOS WHERE PRE_CODCIA=? AND PRE_CODART=? AND PRE_FLAG_UNIDAD='A'"
    Set RS = SQLStaticRecordset(SQL, LK_CODCIA, L_CODART)
    If Not RS.EOF Then
        L_UNIDAD = RS("PRE_UNIDAD")
        L_EQUIV = RS("PRE_EQUIV")
    End If
    RS.Close
    Set RS = Nothing
Exit Sub
ErrHandler:
    RS.Close
    Set RS = Nothing
End Sub
Private Function GetStockFecha(ByVal fecha As String) As Double
Dim RS As ADODB.Recordset
Dim SQL As String
On Error GoTo Handler
SQL = "SELECT sum(far_signo_arm*far_cantidad_p) AS StockFecha,far_codart FROM FACART " & _
      "WHERE FAR_CODCIA = ? AND far_codart= ? and far_estado<>'E' AND FAR_FECHA_COMPRA > ? AND FAR_CODLOT = ? " & _
      "Group BY far_codart"
        Set RS = SQLStaticRecordset(SQL, LK_CODCIA, L_CODART, fecha, L_CODLOT)
        If Not RS.EOF Then
            GetStockFecha = RS("StockFecha")
        Else
            GetStockFecha = 0
        End If
        RS.Close
        Set RS = Nothing
Exit Function
Handler:
    RS.Close
    Set RS = Nothing
    GetStockFecha = "0"
End Function
Private Sub GetSaldos(ByVal Saldo As Double)
Dim i As Integer
Dim iRows As Integer
Dim wSaldo As Double
    iRows = grd_detalle.Rows - 1
    wSaldo = Saldo
    Do Until iRows = 1
        grd_detalle.TextMatrix(iRows, 7) = Format(wSaldo, "#####.00")
        wSaldo = wSaldo - CDbl("0" + grd_detalle.TextMatrix(iRows, 5)) + CDbl("0" + grd_detalle.TextMatrix(iRows, 6))
        iRows = iRows - 1
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmconsulotesxarti = Nothing
End Sub

Private Sub grd_articulos_DblClick()
    grd_articulos_KeyPress (13)
End Sub

Private Sub grd_articulos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If grd_articulos.TextMatrix(grd_articulos.Row, 1) <> "" Then
            L_CODART = grd_articulos.TextMatrix(grd_articulos.Row, 1)
            Load_Lotes L_CODART, txtfecha1.Text, txtfecha2.Text
        End If
    End If
End Sub

Private Sub grd_articulos_RowColChange()
    If grd_articulos.TextMatrix(grd_articulos.Row, 1) <> "" Then
        L_CODART = grd_articulos.TextMatrix(grd_articulos.Row, 1)
    End If
End Sub

Private Sub grd_lotes_DblClick()
    grd_lotes_KeyPress (13)
End Sub

Private Sub grd_lotes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If L_CODART = 0 Then
            Exit Sub
        End If
        If grd_lotes.TextMatrix(grd_lotes.Row, 1) <> "" Then
            L_CODLOT = grd_lotes.TextMatrix(grd_lotes.Row, 1)
            fra_detalle.Visible = True
            fra_detalle.Caption = "DETALLE DE FACART DEL ARTICULO : " + Trim(grd_articulos.TextMatrix(grd_articulos.Row, 3)) + " -- LOTE NRO.: " + grd_lotes.TextMatrix(grd_lotes.Row, 1)
            fra_detalle.Top = 150
            fra_detalle.Left = 45
            opt_order.Item(0).SetFocus
            Dis_EnabledFrames (False)
        End If
    End If
End Sub

Private Sub opt_order_Click(Index As Integer)
    If Index = 0 Then
        fra_nromovimientos.Visible = True
        fra_fecha.Visible = False
        Query_NroMovOrFechas = 1 'top registros
    Else
        fra_nromovimientos.Visible = False
        fra_fecha.Visible = True
        Query_NroMovOrFechas = 2 'por fechas
    End If
    SetGridDetalle
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        SQLFind = "ARTI.ART_alterno AS Codigo , ARTI.ART_NOMBRE as Descripción,"
        SQLWhere = "ARTI.art_alterno like '"
        txtfindarti.TypeFind = Key
    Else
        SQLFind = "ARTI.ART_NOMBRE as Descripción, ARTI.art_alterno AS Codigo,"
        SQLWhere = "ARTI.ART_NOMBRE LIKE '"
        txtfindarti.TypeFind = NameField
    End If
End Sub

Private Sub optTipoBusqueda_Click(Index As Integer)
    If Index = 0 Then
        frafindfam.Visible = True
        fraFindArt.Visible = False
        txtarticulo.Visible = False
        txtfindarti.Visible = False
        cmbfamilias.SetFocus
        fra_lotes.Left = 5160
        fra_lotes.Top = 1605
    Else
        frafindfam.Visible = False
        fraFindArt.Visible = True
        txtarticulo.Visible = True
        txtfindarti.Visible = True
        txtfindarti.SetFocus
        fra_lotes.Left = 45
        fra_lotes.Top = 1605
    End If
    txtNroLote.Text = ""
    txtfindarti.TEXTO = ""
    txtarticulo.Caption = ""
    SetGridLotes
    SetGrid
End Sub

Private Sub txtfindarti_GetRegistros(ByVal oKeyFind As Variant)
Dim RS As ADODB.Recordset
Dim SQL As String
    'FALTA MODIFICAR COMPONENTE
    SQL = "SELECT " + SQLFind + " ARTICULO.ARM_STOCK FROM ARTI INNER JOIN ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA AND ART_CODCIA = '" & LK_CODCIA & "' " & _
          "WHERE " & SQLWhere & oKeyFind & "%' ORDER BY ARTI.ART_NOMBRE"
    Set RS = SQLStaticRecordset(SQL, LK_CODCIA)
    txtfindarti.SetRecordset = RS
    'RS.Close
    Set RS = Nothing
End Sub

Private Sub txtfindarti_ShowData(ByVal oKey As Variant)
    LoadArtOne (oKey)

    
End Sub

Private Sub txtNroLote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtNroLote.Text <> "" Then
            If LoadLoteOne(txtNroLote.Text) Then
                grd_lotes.SetFocus
            End If
        Else
            grd_lotes.SetFocus
        End If
    End If
End Sub
Private Function TipoDocumento(ByVal TIPOMOV As Integer, ByVal CP As String, ByVal fbg As String, ByVal NUMFAC As String, _
        ByVal NUMFAC_C As String, ByVal NUMSER_C As String, ByVal NUMSER As String, ByVal NUMGUIA As String, ByVal TIPDOC As String) As String

 If TIPOMOV = 20 Or TIPOMOV = 180 Then
    If NUMFAC_C <> 0 Then
       TipoDocumento = "Fact." & NUMSER_C & "-" & NUMFAC_C
    Else
       TipoDocumento = "Guia " & NUMGUIA
    End If
    'ws_nombre = Left(Trim(cli_llave!CLI_NOMBRE), 25)
 ElseIf TIPOMOV = 10 Then
    If fbg = "F" Then
       TipoDocumento = "Fact." & NUMSER & "-" & NUMFAC
    ElseIf fbg = "B" Then
       TipoDocumento = "Bol. " & NUMSER & "-" & NUMFAC
    ElseIf fbg = "G" Then
      'If LK_FLAG_GRIFO = "A" Then
      '  TipoDocumento = TIPDOC & " " & NUMSER & "-" & NUMFAC
      'Else
        TipoDocumento = "Guia " & NUMSER & "-" & NUMFAC
     ' End If
    ElseIf fbg = "P" Then
     'If LK_FLAG_GRIFO = "A" Then
     '   TipoDocumento = TIPDOC & " " & NUMSER & "-" & NUMFAC
     ' Else
     '  TipoDocumento = "Guia: " & NUMSER & "-" & NUMFAC
     ' End If
    ElseIf fbg = "X" Then
        TipoDocumento = "Doc.: " & NUMSER & "-" & NUMFAC
    Else
       TipoDocumento = TIPDOC & NUMSER & "-" & NUMFAC
    End If
 ElseIf TIPOMOV = 97 Or TIPOMOV = 98 Then
    If CP = "P" Then
      If TIPOMOV = 97 Then
       TipoDocumento = "N.Cred." & NUMSER_C & "-" & NUMFAC_C
      Else
       TipoDocumento = "N.Deb." & NUMSER_C & "-" & NUMFAC_C
      End If
    Else
      If TIPOMOV = 97 Then
       TipoDocumento = "N.Cred." & NUMSER & "-" & NUMFAC
      Else
       TipoDocumento = "N.Deb." & NUMSER & "-" & NUMFAC
      End If
    End If
 Else
    TipoDocumento = "Guia " & NUMSER & "-" & NUMFAC
 End If
End Function
Private Sub SetGrid()
    grd_articulos.Clear
    grd_articulos.Rows = 2
    grd_articulos.FormatString = "<|CodArt|Unidad|Descripción"
    grd_articulos.ColWidth(0) = 200
    grd_articulos.ColWidth(1) = 800
    grd_articulos.ColWidth(2) = 800
    grd_articulos.ColWidth(3) = 4000
End Sub
Private Sub SetGridLotes()
    grd_lotes.Clear
    grd_lotes.Rows = 2
    grd_lotes.FormatString = "<|NroLote|FechaVenc|FecIngreso|Ingreso|Salidas|Saldos|Proveedor"
    grd_lotes.ColWidth(0) = 200
    grd_lotes.ColWidth(1) = 1000
    grd_lotes.ColWidth(2) = 1000
    grd_lotes.ColWidth(3) = 1000
    grd_lotes.ColWidth(4) = 800
    grd_lotes.ColWidth(5) = 800
    grd_lotes.ColWidth(6) = 800
    grd_lotes.ColWidth(7) = 2000
    grd_lotes.ColWidth(8) = 4000
End Sub
Private Sub SetGridDetalle()
    grd_detalle.Clear
    grd_detalle.Rows = 3
    grd_detalle.FixedRows = 2
    grd_detalle.FormatString = "<Fec.Proc|Fec.Emis.|Documento|Concepto|Descripción|Ingreso|Salida|Saldo|Valor(Ing)|Valor(Sal)|Usuario"
    grd_detalle.TextMatrix(2, 4) = "Saldo Inicial"
    grd_detalle.ColWidth(0) = 1000
    grd_detalle.ColWidth(1) = 1000
    grd_detalle.ColWidth(2) = 1000
    grd_detalle.ColWidth(3) = 1700
    grd_detalle.ColWidth(4) = 2400
    grd_detalle.ColWidth(5) = 800
    grd_detalle.ColWidth(6) = 800
    grd_detalle.ColWidth(7) = 800
    grd_detalle.ColWidth(8) = 800
    grd_detalle.ColWidth(9) = 800
    grd_detalle.ColWidth(10) = 2000
End Sub
Private Sub Dis_EnabledFrames(ByVal sBool As Boolean)
    fra_lotes.Enabled = sBool
    fraFindArt.Enabled = sBool
    frafindfam.Enabled = sBool
    frafechaingreso.Enabled = sBool
    fratipofindart.Enabled = sBool
    txtfindarti.Enabled = sBool
End Sub
Private Function ValFechas() As Boolean
ValFechas = True
    If txtfecha1.Text > txtfecha2 Then
        txtfecha1.SetFocus
        ValFechas = False
        MsgBox "Fecha de Inicio debe ser menor a la Fecha Final"
        Exit Function
    End If
    If (IsDate(txtfecha1.Text) And Not IsDate(txtfecha2.Text)) Then
        MsgBox "Error en la Fecha Fin"
        txtfecha2.SetFocus
        ValFechas = False
        Exit Function
    End If
    If (Not IsDate(txtfecha1.Text) And IsDate(txtfecha2.Text)) Then
        MsgBox "Error en la Fecha de Inicio"
        ValFechas = False
        txtfecha1.SetFocus
        Exit Function
    End If
End Function
Private Function ValFechasC() As Boolean
ValFechasC = True
    If txtfechac1.Text > txtfechac2 Then
        txtfechac1.SetFocus
        ValFechasC = False
        MsgBox "Fecha de Inicio de debe ser menor a la Fecha Final"
        Exit Function
    End If
    If (IsDate(txtfechac1.Text) And Not IsDate(txtfechac2.Text)) Then
        MsgBox "Error en la Fecha Fin"
        txtfechac2.SetFocus
        ValFechasC = False
        Exit Function
    End If
    If (Not IsDate(txtfechac1.Text) And IsDate(txtfechac2.Text)) Then
        MsgBox "Error en la Fecha de Inicio"
        ValFechasC = False
        txtfechac1.SetFocus
        Exit Function
    End If
    If txtfechac1.Text = txtfechac2.Text And txtfecha1.Text = "__/__/____" Then
        ValFechasC = False
        'txtfechac1.SetFocus
        MsgBox "Error en las Fechas. Ingrese Fechas"
        Exit Function
    End If
End Function

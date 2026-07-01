VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmControl 
   Caption         =   "Documentos Anexos"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbfamilia 
      Height          =   315
      Left            =   8400
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmdsigno 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmControl.frx":0000
      Left            =   840
      List            =   "FrmControl.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1080
      Width           =   3135
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
      Left            =   3720
      Picture         =   "FrmControl.frx":002B
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6270
      Width           =   1005
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   735
      Left            =   4560
      TabIndex        =   30
      Top             =   6000
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
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Relacionar"
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
      Picture         =   "FrmControl.frx":0175
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6270
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle de Descuentos "
      Height          =   1635
      Left            =   0
      TabIndex        =   14
      Top             =   1440
      Width           =   11145
      Begin VB.ComboBox cmblinea 
         Height          =   315
         Left            =   2640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   480
         Width           =   1755
      End
      Begin VB.CheckBox cheliq 
         Caption         =   "Anexo Liquidado."
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
         Height          =   375
         Left            =   150
         TabIndex        =   36
         Top             =   990
         Width           =   1665
      End
      Begin VB.ComboBox cmbtipo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   480
         Width           =   2355
      End
      Begin VB.TextBox txtcantidad 
         Height          =   285
         Left            =   8910
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtprecio 
         Height          =   285
         Left            =   9990
         TabIndex        =   19
         Top             =   480
         Width           =   1035
      End
      Begin VB.TextBox i_codart2 
         Height          =   285
         Left            =   4590
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox unidad 
         Height          =   315
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   480
         Width           =   1065
      End
      Begin VB.TextBox txttotal 
         Height          =   285
         Left            =   7200
         TabIndex        =   16
         Top             =   1170
         Width           =   1035
      End
      Begin VB.TextBox txtglosa 
         Height          =   285
         Left            =   1980
         TabIndex        =   15
         Top             =   1170
         Width           =   4905
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Linea :"
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   38
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label i_nomarti 
         Height          =   375
         Left            =   5490
         TabIndex        =   31
         Top             =   420
         Width           =   2145
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Total S/."
         Height          =   255
         Index           =   4
         Left            =   7230
         TabIndex        =   28
         Top             =   930
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Tipo de Descuento :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Producto:"
         Height          =   255
         Index           =   1
         Left            =   4650
         TabIndex        =   26
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Cantidad"
         Height          =   255
         Index           =   2
         Left            =   8970
         TabIndex        =   25
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Precio"
         Height          =   255
         Index           =   3
         Left            =   10080
         TabIndex        =   24
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Unidad"
         Height          =   255
         Index           =   5
         Left            =   7770
         TabIndex        =   23
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Glosa de Descuento : "
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   22
         Top             =   930
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Relacion de Documento:"
      Height          =   1125
      Left            =   30
      TabIndex        =   2
      Top             =   -30
      Width           =   11145
      Begin VB.CommandButton cmdImp 
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
         Height          =   750
         Left            =   9600
         Picture         =   "FrmControl.frx":05B7
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Proveedor :"
         Height          =   255
         Index           =   10
         Left            =   90
         TabIndex        =   13
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Fecha : "
         Height          =   255
         Index           =   11
         Left            =   90
         TabIndex        =   12
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Factura : "
         Height          =   255
         Index           =   12
         Left            =   2760
         TabIndex        =   11
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label d_codigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   270
         Width           =   765
      End
      Begin VB.Label d_nombre 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2070
         TabIndex        =   9
         Top             =   270
         Width           =   5955
      End
      Begin VB.Label d_fecha 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1170
         TabIndex        =   8
         Top             =   690
         Width           =   1125
      End
      Begin VB.Label d_serie 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Top             =   690
         Width           =   435
      End
      Begin VB.Label d_numero 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4410
         TabIndex        =   6
         Top             =   690
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Interno : "
         Height          =   255
         Index           =   13
         Left            =   5640
         TabIndex        =   5
         Top             =   750
         Width           =   855
      End
      Begin VB.Label d_numser 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6570
         TabIndex        =   4
         Top             =   690
         Width           =   435
      End
      Begin VB.Label d_numfac 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7140
         TabIndex        =   3
         Top             =   690
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdidet 
      Height          =   2925
      Left            =   30
      TabIndex        =   0
      Top             =   3300
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5159
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedCols       =   0
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
   Begin VB.Label LBLESTADO 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   9480
      TabIndex        =   41
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label flag_signo 
      Caption         =   "1"
      Height          =   225
      Left            =   3180
      TabIndex        =   35
      Top             =   1110
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Opción :"
      Height          =   225
      Left            =   60
      TabIndex        =   34
      Top             =   1110
      Width           =   795
   End
   Begin VB.Label lblcontrol 
      Caption         =   "Detalle de Descuentos :"
      Height          =   255
      Index           =   8
      Left            =   30
      TabIndex        =   1
      Top             =   3030
      Width           =   3015
   End
End
Attribute VB_Name = "FrmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim PRE_RELCOMP As rdoQuery
Dim relcomp_llave  As rdoResultset
Dim loc_key As Integer
Dim PRE_UNIDADES As rdoQuery
Dim preuni_llave  As rdoResultset

Private Sub cheliq_Click()
txtglosa.SetFocus
End Sub

Private Sub cmbfamilia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grdidet.TextMatrix(grdidet.Row, 13) = Trim(Right(cmbfamilia.Text, 8))
    grdidet.TextMatrix(grdidet.Row, 14) = Left(cmbfamilia.Text, 10)
    cmbfamilia.Visible = False
End If
If KeyAscii = 27 Then cmbfamilia.Visible = False

End Sub

Private Sub cmblinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  i_codart2.SetFocus
End If
End Sub

Private Sub cmbtipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cmblinea.SetFocus
End If
End Sub

Private Sub cmbtipo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
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

PRE_RELCOMP(0) = LK_CODCIA
PRE_RELCOMP(1) = "P"
PRE_RELCOMP(2) = Val(d_codigo.Caption)
PRE_RELCOMP(3) = Val(d_numser.Caption)
PRE_RELCOMP(4) = Val(d_numfac.Caption)
relcomp_llave.Requery
If Not relcomp_llave.EOF Then
 pub_mensaje = "Reemplazar la infomación...   ¿Desea Continuar... ?"
 Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
 If Pub_Respuesta = vbNo Then
    Exit Sub
 End If
  pub_cadena = "DELETE FROM RELCOMPRA WHERE REL_CODCIA = '" & LK_CODCIA & "' AND REL_CP = 'P' AND REL_CODPRO = " & Trim(d_codigo.Caption) & " AND REL_NUMSER = " & Val(d_numser.Caption) & " AND REL_NUMFAC =  " & d_numfac.Caption & " "
 CN.Execute pub_cadena, rdExecDirect
End If
Screen.MousePointer = 13
For fila = 1 To grdidet.Rows - 1
    relcomp_llave.AddNew
    relcomp_llave!rel_codcia = LK_CODCIA
    relcomp_llave!rel_cp = "P"
    relcomp_llave!rel_codpro = Val(d_codigo.Caption)
    relcomp_llave!rel_numser = Val(d_numser.Caption)
    relcomp_llave!rel_numfac = Val(d_numfac.Caption)
    relcomp_llave!REL_NUMSEC = fila
    relcomp_llave!rel_FECHA_COMPRA = Trim(d_fecha.Caption)
    relcomp_llave!rel_numser_c = Val(d_serie.Caption)
    relcomp_llave!rel_numfac_c = Val(d_numero.Caption)
    relcomp_llave!rel_codtip = Val(Right(grdidet.TextMatrix(fila, 1), 8))
    relcomp_llave!rel_codart = Val(grdidet.TextMatrix(fila, 2))
    relcomp_llave!rel_descri = Trim(Left(grdidet.TextMatrix(fila, 4), 20))
    relcomp_llave!rel_equiv = Val(Trim(Right(grdidet.TextMatrix(fila, 4), 8)))
    relcomp_llave!rel_cantidad = Val(grdidet.TextMatrix(fila, 5)) * Val(relcomp_llave!rel_equiv)
    relcomp_llave!rel_total = Val(grdidet.TextMatrix(fila, 8))
    relcomp_llave!rel_precio = Val(grdidet.TextMatrix(fila, 6))
    relcomp_llave!rel_GLOSA = Trim((grdidet.TextMatrix(fila, 7)))
    relcomp_llave!rel_signo_carga = Val((grdidet.TextMatrix(fila, 9)))
    relcomp_llave!rel_FECHA = grdidet.TextMatrix(fila, 11)
    relcomp_llave!rel_LIQUIDO = cheliq.Value
    relcomp_llave!rel_linea = Val((grdidet.TextMatrix(fila, 13)))
    relcomp_llave!rel_codusu = Trim(grdidet.TextMatrix(fila, 15))
    relcomp_llave.Update
Next fila
Screen.MousePointer = 0
MsgBox "Relacion Efectuada.", 48, Pub_Titulo
Unload FrmControl
End Sub

Private Sub cmdCerrar_Click()
Unload FrmControl
End Sub

Private Sub cmdImp_Click()
Dim wser As String * 3
Dim WSRUTA As String
Dim wRuta As String
Dim rmoneda As String * 1

wRuta = PUB_RUTA_OTRO
FrmControl.Reportes.Connect = PUB_ODBC
FrmControl.Reportes.Destination = crptToWindow  '= crptToPrinter
FrmControl.Reportes.WindowLeft = 2
FrmControl.Reportes.WindowTop = 70
FrmControl.Reportes.WindowWidth = 635
FrmControl.Reportes.WindowHeight = 390
FrmControl.Reportes.Formulas(1) = ""
PUB_NETO = 0
PUB_FECHA = LK_FECHA_DIA
PU_NUMSER = Val(d_numser.Caption)
PU_NUMFAC = Val(d_numfac.Caption)
'FrmControl.Reportes.Formulas(1) = "SON_EFECTIVO=  'SON: " & CONVER_LETRAS(PUB_NETO, rmoneda) & "'"
FrmControl.Reportes.WindowTitle = "ORDEN DE COMPRA  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "0000000")
FrmControl.Reportes.ReportFileName = wRuta + "INFANEXO.RPT"
wser = PU_NUMSER
pub_cadena = "{RELCOMPRA.REL_CODCIA} = '" & LK_CODCIA & "' AND {RELCOMPRA.REL_CP} = 'P' AND {RELCOMPRA.REL_NUMSER}= " & wser & " AND {RELCOMPRA.REL_NUMFAC} = " & PU_NUMFAC
FrmControl.Reportes.SelectionFormula = pub_cadena
FrmControl.Reportes.WindowTitle = FrmControl.Reportes.WindowTitle & " Archivo: " & Trim(FrmControl.Reportes.ReportFileName)

On Error GoTo accion
FrmControl.Reportes.Action = 1
On Error GoTo 0

Exit Sub
accion:
MsgBox Err.Description, 48, Pub_Titulo
End Sub

Private Sub Form_Activate()
If Trim(LBLESTADO.Caption) <> "" Then Exit Sub

If LK_CODUSU = "ADMIN" Or LK_CODUSU = "OPER11" Or LK_CODUSU = "SUPER" Then
cheliq.Enabled = True
Else
cheliq.Enabled = False
End If
Screen.MousePointer = 0

'If fraanexos.Visible = True Then
' GoTo anexo_act
'End If

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
   LBLESTADO.Caption = "<Nuevo>"
   Exit Sub
End If
 LBLESTADO.Caption = "<Editar>"

cabe
cheliq.Tag = "X"
If Val(relcomp_llave!rel_LIQUIDO) = 9 Then
 cheliq.Value = 1
Else
 cheliq.Value = relcomp_llave!rel_LIQUIDO
End If
cheliq.Tag = ""
Do Until relcomp_llave.EOF
    grdidet.Rows = grdidet.Rows + 1
    PUB_TIPREG = 43
    PUB_NUMTAB = relcomp_llave!rel_codtip
    SQ_OPER = 1
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
       grdidet.TextMatrix(grdidet.Rows - 1, 1) = tab_llave!tab_nomlargo & String(80, " ") & tab_llave!tab_numtab
    End If
    grdidet.TextMatrix(grdidet.Rows - 1, 2) = relcomp_llave!rel_codart
    grdidet.TextMatrix(grdidet.Rows - 1, 0) = Format(relcomp_llave!rel_FECHA, "dd/mm/yy")
    If Val(relcomp_llave!rel_codart) <> 0 Then
     SQ_OPER = 1
     PUB_KEY = relcomp_llave!rel_codart
     pu_codcia = relcomp_llave!rel_codcia
     LEER_ART_LLAVE
     
     grdidet.TextMatrix(grdidet.Rows - 1, 3) = Trim(art_LLAVE!ART_NOMBRE)
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
    grdidet.TextMatrix(grdidet.Rows - 1, 13) = relcomp_llave!rel_linea
    grdidet.TextMatrix(grdidet.Rows - 1, 15) = Trim(relcomp_llave!rel_codusu)
    SQ_OPER = 1
    PUB_NUMTAB = relcomp_llave!rel_linea
    PUB_TIPREG = 122
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
       grdidet.TextMatrix(grdidet.Rows - 1, 14) = Trim(tab_llave!tab_nomlargo)
    End If
    
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
If Val(FrmControl.d_numfac.Tag) <> 0 Then
   FrmControl.d_numser.Caption = FrmControl.d_numser.Tag
   FrmControl.d_numfac.Caption = FrmControl.d_numfac.Tag
End If

Exit Sub
anexo_act:

pub_cadena = "SELECT * FROM RELCOMPRA WHERE REL_CODCIA = ? AND REL_CP = ? AND REL_CODPRO = ? AND " & archi & "  ORDER BY REL_NUMSER, REL_NUMFAC "
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
       GRID_MST.TextMatrix(GRID_MST.Rows - 1, 0) = tab_llave!tab_nomlargo & String(80, " ") & tab_llave!tab_numtab
    End If
    GRID_MST.TextMatrix(GRID_MST.Rows - 1, 1) = relcomp_llave!rel_codart
    If Val(relcomp_llave!rel_codart) <> 0 Then
     SQ_OPER = 1
     PUB_KEY = relcomp_llave!rel_codart
     pu_codcia = relcomp_llave!rel_codcia
     LEER_ART_LLAVE
     GRID_MST.TextMatrix(GRID_MST.Rows - 1, 2) = Trim(art_LLAVE!ART_NOMBRE)
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
CenterMe FrmControl
loc_key = 0
cabe
LLENADOS cmbTipo, 43
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ?  ORDER BY PRE_CODART"
Set PRE_UNIDADES = CN.CreateQuery("", pub_cadena)
PRE_UNIDADES(0) = 0
PRE_UNIDADES(1) = 0
Set preuni_llave = PRE_UNIDADES.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM RELCOMPRA WHERE REL_CODCIA = ? AND REL_CP = ? AND REL_CODPRO = ? AND REL_NUMSER = ? AND REL_NUMFAC = ?  ORDER BY REL_NUMSER, REL_NUMFAC "
Set PRE_RELCOMP = CN.CreateQuery("", pub_cadena)
PRE_RELCOMP(0) = 0
PRE_RELCOMP(1) = 0
PRE_RELCOMP(2) = 0
PRE_RELCOMP(3) = 0
PRE_RELCOMP(4) = 0
Set relcomp_llave = PRE_RELCOMP.OpenResultset(rdOpenKeyset, rdConcurValues)
LLENADOS cmblinea, 122
LLENADOS cmbfamilia, 122

End Sub


Private Sub grdidet_DblClick()
Dim wult  As String
If grdidet.COL = 0 Then
    wult = InputBox("Fecha Proceso:", "Cambio de Fecha", grdidet.TextMatrix(grdidet.Row, grdidet.COL))
    If wult = "" Then Exit Sub
    If Not IsDate(wult) Then
      MsgBox "Fecha Invalida.", 48, Pub_Titulo
      Exit Sub
    End If
    grdidet.TextMatrix(grdidet.Row, grdidet.COL) = wult
    grdidet.TextMatrix(grdidet.Row, 11) = wult
Else
cmbfamilia.Visible = True
If Val(grdidet.Text) <> 0 Then
    ASIGNA_INT cmbfamilia, grdidet.Text
End If
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

Private Sub grdidet_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo sale

If KeyCode = 46 Then
If grdidet.TextMatrix(grdidet.Row, 10) = "A" Then
    LK_ACCESO_REPORT = ""
    Load frmclave2
    Screen.MousePointer = 0
    frmclave2.Show 1
    If LK_ACCESO_REPORT <> "A" Then
        Exit Sub
    End If
End If

  If grdidet.Rows = 2 Then
    grdidet.Rows = 1
  ElseIf grdidet.Rows > 2 Then
    grdidet.RemoveItem grdidet.Row
  End If
End If
sale:
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
        PUB_CODART = art_llave_alt!ART_KEY
     Else
        SQ_OPER = 1

        PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
        pu_codcia = LK_CODCIA
        LEER_ART_LLAVE
        If art_LLAVE.EOF Then
           MsgBox "Codigo No Existe ...", 48, Pub_Titulo
           Exit Sub
        End If
        PUB_CODART = art_LLAVE!ART_KEY
        i_nomarti.Caption = art_LLAVE!ART_NOMBRE
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
    WCOD_ORIGINAL = art_LLAVE!ART_KEY
    i_nomarti.Caption = Trim(art_LLAVE!ART_NOMBRE)
    'i_codart2.text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
    PRE_UNIDADES(0) = LK_CODCIA
    PRE_UNIDADES(1) = art_LLAVE!ART_KEY
    preuni_llave.Requery
    unidad.Clear
    unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!pre_unidad & String(80, " ") & preuni_llave!pre_equiv
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
         unid_max = preuni_llave.AbsolutePosition - 1
'         LBLUNIDAD.Caption = Trim(preuni_llave!pre_UNIDAD)
      '   Exit Do
       End If
      preuni_llave.MoveNext
    Loop
    ListView1.Visible = False
    If unidad.ListCount > 0 Then unidad.ListIndex = unid_max
    unidad.SetFocus
    'If txtCampo1.Visible Then
    ' Azul2 txtCampo1, txtCampo1
    ' Exit Sub
    'End If
    'CmdProcesa.SetFocus
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
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     'i_codart2.text = Trim(art_llave_alt!ART_NOMBRE)
     If art_llave_alt!ART_MONEDA = "D" Then
      lblMoneda.Caption = "Moneda : US$."
     Else
      lblMoneda.Caption = "Moneda : S/ ."
     End If
     i_nomarti.Caption = Trim(art_llave_alt!ART_NOMBRE)
     ListView1.Visible = False
     
     PRE_UNIDADES(0) = LK_CODCIA
     PRE_UNIDADES(1) = art_llave_alt!ART_KEY
     preuni_llave.Requery
     unidad.Clear
     unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!pre_unidad & String(80, " ") & preuni_llave!pre_equiv
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
          unid_max = preuni_llave.AbsolutePosition - 1
          LBLUNIDAD.Caption = Trim(preuni_llave!pre_unidad)
          Exit Do
        End If
        preuni_llave.MoveNext
     Loop
     If unidad.ListCount > 0 Then unidad.ListIndex = unid_max
     unidad.SetFocus
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
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul i_codart2, i_codart2
        Exit Sub
      End If
      WCOD_ORIGINAL = art_LLAVE!ART_KEY
      i_nomarti.Caption = Trim(art_LLAVE!ART_NOMBRE)
    PRE_UNIDADES(0) = LK_CODCIA
    PRE_UNIDADES(1) = art_LLAVE!ART_KEY
    preuni_llave.Requery
    unidad.Clear
    unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!pre_unidad & String(80, " ") & preuni_llave!pre_equiv
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
         unid_max = preuni_llave.AbsolutePosition - 1
'         LBLUNIDAD.Caption = Trim(preuni_llave!pre_UNIDAD)
      '   Exit Do
       End If
      preuni_llave.MoveNext
    Loop
    ListView1.Visible = False
    If unidad.ListCount > 0 Then unidad.ListIndex = unid_max
     unidad.SetFocus
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
      archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
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
            cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
          End If
       Else
           cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
       End If
       tab_mayor.MoveNext
    Loop
End Sub

Private Sub txtcantidad_Change()
cal
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtprecio, txtprecio
End If
End Sub

Private Sub txtglosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Azul txttotal, txttotal
End If
End Sub

Private Sub txtprecio_Change()
cal
End Sub

Private Sub txtprecio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtglosa.SetFocus
End If
End Sub

Private Sub txttotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
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
grdidet.Cols = 16
grdidet.Rows = 1
grdidet.ColWidth(0) = 900 ' Fecha
grdidet.ColWidth(1) = 2000 ' tipo
grdidet.ColWidth(2) = 600 ' codigo internot
grdidet.ColWidth(3) = 1300  ' descrip prod
grdidet.ColWidth(4) = 600  ' unidad
grdidet.ColWidth(5) = 800 ' cantidad
grdidet.ColWidth(6) = 800 ' precio
grdidet.ColWidth(7) = 2800 ' glosa
grdidet.ColWidth(8) = 900 ' total
grdidet.ColWidth(9) = 0 ' Signo de cargo
grdidet.ColWidth(10) = 0 ' No Editar
grdidet.ColWidth(11) = 0 ' FECHA DE PROCESO
grdidet.ColWidth(12) = 0 ' FLAG DE CIERRE
grdidet.ColWidth(13) = 900 ' codigo de Linea
grdidet.ColWidth(14) = 900 ' descrip de Linea
grdidet.ColWidth(15) = 900 ' descrip de Linea

grdidet.TextMatrix(0, 0) = "Fecha"
grdidet.TextMatrix(0, 1) = "Tipo"
grdidet.TextMatrix(0, 2) = "Codigo"
grdidet.TextMatrix(0, 3) = "Descripción"
grdidet.TextMatrix(0, 4) = "Unidad"
grdidet.TextMatrix(0, 5) = "Cantidad"
grdidet.TextMatrix(0, 6) = "Precio"
grdidet.TextMatrix(0, 7) = "Glosa"
grdidet.TextMatrix(0, 8) = "Total"
grdidet.TextMatrix(0, 13) = "Cod.Linea"
grdidet.TextMatrix(0, 14) = "Des.Linea"
grdidet.TextMatrix(0, 15) = "Usuario"


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


GRID_MST.TextMatrix(0, 0) = "Tipo"
GRID_MST.TextMatrix(0, 1) = "Codigo"
GRID_MST.TextMatrix(0, 2) = "Descripción"
GRID_MST.TextMatrix(0, 3) = "Unidad"
GRID_MST.TextMatrix(0, 4) = "Cantidad"
GRID_MST.TextMatrix(0, 5) = "Precio"
GRID_MST.TextMatrix(0, 6) = "Glosa"
GRID_MST.TextMatrix(0, 7) = "Total"


End Sub

Public Sub add_registro()
If Trim(cmbTipo.Text) = "" Then
  MsgBox "Seleccionar Tipo de Documento", 48, Pub_Titulo
  cmbTipo.SetFocus
  Exit Sub
End If

If Trim(cmblinea.Text) = "" Then
  MsgBox "Seleccionar Linea de Productos ", 48, Pub_Titulo
  cmblinea.SetFocus
  Exit Sub
End If
If Len(Trim(txtglosa.Text)) <= 20 Then
  MsgBox "La glosa debe ser llenada por lo menos unos 20 Digitos.", 48, Pub_Titulo
  txtglosa.SetFocus
  Exit Sub
End If
If Val(txttotal.Text) = 0 Then
  MsgBox "El Valor Total, no puede estar en 0.00", 48, Pub_Titulo
  Azul txttotal, txttotal
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
grdidet.TextMatrix(grdidet.Rows - 1, 8) = Val(txttotal.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 9) = Trim(Left(cmdsigno.Text, 2))
grdidet.TextMatrix(grdidet.Rows - 1, 11) = LK_FECHA_DIA
grdidet.TextMatrix(grdidet.Rows - 1, 12) = cheliq.Value
grdidet.TextMatrix(grdidet.Rows - 1, 13) = Trim(Right(cmblinea.Text, 8))
grdidet.TextMatrix(grdidet.Rows - 1, 15) = LK_CODUSU

If Trim(Left(cmdsigno.Text, 2)) = -1 Then
  grdidet.Row = grdidet.Rows - 1
  grdidet.COL = 1
  grdidet.CellBackColor = QBColor(14)
End If


i_codart2.Text = ""
i_nomarti.Caption = ""
unidad.Clear
txtcantidad.Text = ""
txtprecio.Text = ""
txtglosa.Text = ""
txttotal.Text = ""
cmbTipo.SetFocus
End Sub

Public Sub cal()
 txttotal.Text = Format(Val(txtprecio.Text) * Val(txtcantidad.Text), "0.00")
End Sub

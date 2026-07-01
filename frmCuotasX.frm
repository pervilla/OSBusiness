VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCuotas 
   Caption         =   "Definición de Cuotas."
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView1 
      Height          =   495
      Left            =   6405
      TabIndex        =   21
      Top             =   6345
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Frame fratipo 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   11475
      Begin VB.ComboBox ano 
         Height          =   315
         Left            =   8520
         TabIndex        =   5
         Text            =   "ano"
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox mes 
         Height          =   315
         Left            =   6495
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   225
         Width           =   1575
      End
      Begin VB.ComboBox Cmbtipos 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label fecha2 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   10035
         TabIndex        =   20
         Top             =   420
         Width           =   1290
      End
      Begin VB.Label fecha1 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   10050
         TabIndex        =   19
         Top             =   135
         Width           =   1290
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Año:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8040
         TabIndex        =   18
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Mes:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   5880
         TabIndex        =   17
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccione Tipo de Cuota :"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6615
      Left            =   120
      TabIndex        =   11
      Top             =   705
      Width           =   11655
      Begin MSComctlLib.ProgressBar pro 
         Height          =   210
         Left            =   1425
         TabIndex        =   23
         Top             =   5235
         Visible         =   0   'False
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   6240
         TabIndex        =   16
         Top             =   1890
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmCuotas.frx":0000
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
      Begin VB.ComboBox cmbdivi 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   4485
      End
      Begin VB.TextBox Txt_key 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         MaxLength       =   10
         TabIndex        =   2
         Top             =   705
         Width           =   855
      End
      Begin VB.CommandButton cmdava 
         Caption         =   "Mostrar Avance"
         Height          =   435
         Left            =   2400
         TabIndex        =   8
         Top             =   5520
         Width           =   2415
      End
      Begin VB.CommandButton graba 
         Caption         =   "&Grabar"
         Height          =   435
         Left            =   585
         TabIndex        =   7
         Top             =   5535
         Width           =   1335
      End
      Begin VB.CommandButton Cmdclose 
         Caption         =   "Ce&rrar"
         Height          =   435
         Left            =   9480
         TabIndex        =   9
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdmostrar 
         Caption         =   "&Mostrar"
         Height          =   510
         Left            =   5820
         TabIndex        =   3
         Top             =   375
         Width           =   1050
      End
      Begin MSFlexGridLib.MSFlexGrid gridigv 
         Height          =   3795
         Left            =   150
         TabIndex        =   6
         ToolTipText     =   "[Enter] = para Editar"
         Top             =   1410
         Width           =   11370
         _ExtentX        =   20055
         _ExtentY        =   6694
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         BackColorBkg    =   8421376
      End
      Begin VB.Label lblpro 
         Caption         =   "Procesando ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   90
         TabIndex        =   24
         Top             =   5265
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblven 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1860
         TabIndex        =   22
         Top             =   705
         Width           =   3540
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   -165
         X2              =   11655
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   11640
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "División :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor:"
         Height          =   420
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de Cuotas :"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   1245
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loc_key As Integer
Dim temporal
Dim loc_tipo As Integer
Dim PSCUO_LLAVE As rdoQuery
Dim cuo_rep01 As rdoResultset
Dim PSCUO_VENDEDOR As rdoQuery
Dim cuo_vendedor As rdoResultset
Dim PSCUO_SIGUE As rdoQuery
Dim cuo_sigue As rdoResultset




Private Sub ano_Change()
cal_fecha
End Sub

Private Sub ano_Click()
cal_fecha
End Sub

Private Sub cmbdivi_Click()
gridigv.Rows = 2
End Sub

Private Sub cmbdivi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul Txt_key, Txt_key
End If
End Sub

Private Sub Cmbtipos_Click()
loc_tipo = Val(Left(Cmbtipos.Text, 2))
cmbdivi.Enabled = True
cmbdivi.BackColor = QBColor(15)
Txt_key.Enabled = True
Txt_key.BackColor = QBColor(15)
Select Case loc_tipo
Case 1
  cmbdivi.ListIndex = -1
  cmbdivi.BackColor = QBColor(7)
  Txt_key.Text = ""
  Txt_key.BackColor = QBColor(7)
  cmbdivi.Enabled = False
  Txt_key.Enabled = False
  cmdmostrar.SetFocus
Case 2
   cmbdivi.SetFocus
Case 3
  cmbdivi.ListIndex = -1
  cmbdivi.Enabled = False
  cmbdivi.BackColor = QBColor(7)
  Txt_key.SetFocus
Case 4
  cmbdivi.ListIndex = -1
  cmbdivi.Enabled = False
  cmbdivi.BackColor = QBColor(7)
  Txt_key.BackColor = QBColor(7)
  Txt_key.Text = ""
  Txt_key.Enabled = False
   cmdmostrar.SetFocus
Case 5
  Txt_key.Text = ""
  Txt_key.Enabled = False
  Txt_key.BackColor = QBColor(7)
  cmbdivi.SetFocus
Case 6
  cmbdivi.ListIndex = -1
  cmbdivi.BackColor = QBColor(7)
  Txt_key.Text = ""
  Txt_key.BackColor = QBColor(7)
  cmbdivi.Enabled = False
  Txt_key.Enabled = False
  cmdmostrar.SetFocus
End Select
gridigv.Rows = 2
End Sub

Private Sub cmdClose_Click()
Unload frmCuotas
End Sub

Private Sub cmdmostrar_Click()
Dim PSCUO_VENDEDOR As rdoQuery
Dim cuo_vendedor As rdoResultset
If Not IsDate(fecha1.Caption) Then
 MsgBox "Verificar Configuración Regional .. no procde las Fechas Nº1", 48, Pub_Titulo
 Exit Sub
End If
If Not IsDate(fecha2.Caption) Then
 MsgBox "Verificar Configuración Regional .. no procde las Fechas Nº1", 48, Pub_Titulo
 Exit Sub
End If
DoEvents
lblpro.Visible = True
pro.Visible = True
DoEvents

Select Case loc_tipo
Case 1
    cabe
    pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? ORDER BY VEM_CODVEN"
    Set PSCUO_VENDEDOR = CN.CreateQuery("", pub_cadena)
    PSCUO_VENDEDOR(0) = LK_CODCIA
    Set cuo_vendedor = PSCUO_VENDEDOR.OpenResultset(rdOpenKeyset, rdConcurValues)
    pub_cadena = "SELECT * FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? AND CUO_CODVEN = ? "
    Set PSCUO_LLAVE = CN.CreateQuery("", pub_cadena)
    PSCUO_LLAVE(0) = 0
    PSCUO_LLAVE(1) = 0
    PSCUO_LLAVE(2) = 0
    PSCUO_LLAVE(3) = 0
    PSCUO_LLAVE(4) = 0
    Set cuo_rep01 = PSCUO_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
    If Not cuo_vendedor.EOF Then
      pro.Min = 0
      pro.max = cuo_vendedor.RowCount
      pro.Value = 0
    End If
    Do Until cuo_vendedor.EOF
            pro.Value = pro.Value + 1
            gridigv.Rows = gridigv.Rows + 1
            gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!VEM_CODVEN
            gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!VEM_NOMBRE
            PSCUO_LLAVE(0) = LK_CODCIA
            PSCUO_LLAVE(1) = loc_tipo
            PSCUO_LLAVE(2) = fecha1.Caption
            PSCUO_LLAVE(3) = fecha2.Caption
            PSCUO_LLAVE(4) = cuo_vendedor!VEM_CODVEN
            cuo_rep01.Requery
            If Not cuo_rep01.EOF Then
             gridigv.TextMatrix(gridigv.Rows - 1, 3) = cuo_rep01!CUO_CANTIDAD / cuo_rep01!CUO_EQUIV
             gridigv.TextMatrix(gridigv.Rows - 1, 5) = cuo_rep01!CUO_MONTO
            End If
            
           cuo_vendedor.MoveNext
     Loop

Case 2
    cabe
    If Val(Txt_key.Text) = 0 Then
      MsgBox "Seleccionar un Vendedor", 48, Pub_Titulo
      Txt_key.SetFocus
      GoTo fin
    End If
    If Trim(cmbdivi.Text) = "" Then
      MsgBox "Seleccionar un División", 48, Pub_Titulo
      cmbdivi.SetFocus
      SendKeys "%{UP}"
      GoTo fin
    End If
    '  pub_cadena = "SELECT * FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? "
    pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_FAMILIA = ? ORDER BY ART_ALTERNO"
    Set PSCUO_VENDEDOR = CN.CreateQuery("", pub_cadena)
    PSCUO_VENDEDOR(0) = LK_CODCIA
    PSCUO_VENDEDOR(1) = Val(Right(cmbdivi.Text, 6))
    Set cuo_vendedor = PSCUO_VENDEDOR.OpenResultset(rdOpenKeyset, rdConcurValues)

    pub_cadena = "SELECT * FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? " & _
    "AND CUO_CODVEN = ? and CUO_DIVISION = ? AND CUO_CODART = ?"
    Set PSCUO_LLAVE = CN.CreateQuery("", pub_cadena)
    PSCUO_LLAVE(0) = 0
    PSCUO_LLAVE(1) = 0
    PSCUO_LLAVE(2) = Date
    PSCUO_LLAVE(3) = Date
    PSCUO_LLAVE(4) = 0
    PSCUO_LLAVE(5) = 0
    PSCUO_LLAVE(6) = 0
    Set cuo_rep01 = PSCUO_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
    If Not cuo_vendedor.EOF Then
      pro.Min = 0
      pro.max = cuo_vendedor.RowCount
      pro.Value = 0
    End If
    Do Until cuo_vendedor.EOF
         pro.Value = pro.Value + 1
         gridigv.Rows = gridigv.Rows + 1
         gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!art_alterno
         gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!art_nombre
         gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_vendedor!ART_KEY
         PSCUO_LLAVE(0) = LK_CODCIA
         PSCUO_LLAVE(1) = loc_tipo
         PSCUO_LLAVE(2) = fecha1.Caption
         PSCUO_LLAVE(3) = fecha2.Caption
         PSCUO_LLAVE(4) = Val(Txt_key.Text)
         PSCUO_LLAVE(5) = Val(Right(cmbdivi.Text, 6))
         PSCUO_LLAVE(6) = cuo_vendedor!ART_KEY
         cuo_rep01.Requery
         If Not cuo_rep01.EOF Then
          gridigv.TextMatrix(gridigv.Rows - 1, 2) = cuo_rep01!CUO_DESCRIP
          gridigv.TextMatrix(gridigv.Rows - 1, 3) = cuo_rep01!CUO_CANTIDAD / cuo_rep01!CUO_EQUIV
          gridigv.TextMatrix(gridigv.Rows - 1, 5) = cuo_rep01!CUO_MONTO
          gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_rep01!CUO_CODART
         End If
         
         cuo_vendedor.MoveNext
      Loop
    
Case 3
cabe
    If Val(Txt_key.Text) = 0 Then
      MsgBox "Seleccionar un Vendedor", 48, Pub_Titulo
      Txt_key.SetFocus
      GoTo fin
    End If
    pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ?  ORDER BY TAB_NOMLARGO"
    Set PSCUO_VENDEDOR = CN.CreateQuery("", pub_cadena)
    PSCUO_VENDEDOR(0) = LK_CODCIA
    PSCUO_VENDEDOR(1) = 122
    Set cuo_vendedor = PSCUO_VENDEDOR.OpenResultset(rdOpenKeyset, rdConcurValues)

    pub_cadena = "SELECT * FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? " & _
    "AND CUO_CODVEN = ? and CUO_DIVISION = ? "
    Set PSCUO_LLAVE = CN.CreateQuery("", pub_cadena)
    PSCUO_LLAVE(0) = 0
    PSCUO_LLAVE(1) = 0
    PSCUO_LLAVE(2) = Date
    PSCUO_LLAVE(3) = Date
    PSCUO_LLAVE(4) = 0
    PSCUO_LLAVE(5) = 0
    Set cuo_rep01 = PSCUO_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
     
    If Not cuo_vendedor.EOF Then
      pro.Min = 0
      pro.max = cuo_vendedor.RowCount
      pro.Value = 0
    End If
    Do Until cuo_vendedor.EOF
         pro.Value = pro.Value + 1
         gridigv.Rows = gridigv.Rows + 1
         gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!TAB_NUMTAB
         gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!tab_NOMLARGO
         gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_vendedor!TAB_NUMTAB
         PSCUO_LLAVE(0) = LK_CODCIA
         PSCUO_LLAVE(1) = loc_tipo
         PSCUO_LLAVE(2) = fecha1.Caption
         PSCUO_LLAVE(3) = fecha2.Caption
         PSCUO_LLAVE(4) = Val(Txt_key.Text)
         PSCUO_LLAVE(5) = Val(cuo_vendedor!TAB_NUMTAB)
         cuo_rep01.Requery
         If Not cuo_rep01.EOF Then
           gridigv.TextMatrix(gridigv.Rows - 1, 2) = cuo_rep01!CUO_DESCRIP
           gridigv.TextMatrix(gridigv.Rows - 1, 3) = cuo_rep01!CUO_CANTIDAD / cuo_rep01!CUO_EQUIV
           gridigv.TextMatrix(gridigv.Rows - 1, 5) = cuo_rep01!CUO_MONTO
           gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_rep01!CUO_CODART
         End If
         
         cuo_vendedor.MoveNext
      Loop

Case 4 ' CUOTA POR DIVISION
cabe
    
    pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ?  ORDER BY TAB_NOMLARGO"
    Set PSCUO_VENDEDOR = CN.CreateQuery("", pub_cadena)
    PSCUO_VENDEDOR(0) = LK_CODCIA
    PSCUO_VENDEDOR(1) = 122
    Set cuo_vendedor = PSCUO_VENDEDOR.OpenResultset(rdOpenKeyset, rdConcurValues)

    pub_cadena = "SELECT * FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? " & _
    "AND CUO_DIVISION = ? "
    Set PSCUO_LLAVE = CN.CreateQuery("", pub_cadena)
    PSCUO_LLAVE(0) = 0
    PSCUO_LLAVE(1) = 0
    PSCUO_LLAVE(2) = Date
    PSCUO_LLAVE(3) = Date
    PSCUO_LLAVE(4) = 0
    Set cuo_rep01 = PSCUO_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
     
    If Not cuo_vendedor.EOF Then
      pro.Min = 0
      pro.max = cuo_vendedor.RowCount
      pro.Value = 0
    End If
    Do Until cuo_vendedor.EOF
        pro.Value = pro.Value + 1
         gridigv.Rows = gridigv.Rows + 1
         gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!TAB_NUMTAB
         gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!tab_NOMLARGO
         gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_vendedor!TAB_NUMTAB
         PSCUO_LLAVE(0) = LK_CODCIA
         PSCUO_LLAVE(1) = loc_tipo
         PSCUO_LLAVE(2) = fecha1.Caption
         PSCUO_LLAVE(3) = fecha2.Caption
         PSCUO_LLAVE(4) = Val(cuo_vendedor!TAB_NUMTAB)
         cuo_rep01.Requery
         If Not cuo_rep01.EOF Then
           gridigv.TextMatrix(gridigv.Rows - 1, 2) = cuo_rep01!CUO_DESCRIP
           gridigv.TextMatrix(gridigv.Rows - 1, 3) = cuo_rep01!CUO_CANTIDAD / cuo_rep01!CUO_EQUIV
           gridigv.TextMatrix(gridigv.Rows - 1, 5) = cuo_rep01!CUO_MONTO
           gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_rep01!CUO_CODART
         End If
         
         cuo_vendedor.MoveNext
      Loop

Case 5  ' CUOTA POR ARTICULOS
    cabe
    If Trim(cmbdivi.Text) = "" Then
      MsgBox "Seleccionar un División", 48, Pub_Titulo
      cmbdivi.SetFocus
      SendKeys "%{UP}"
      GoTo fin
      Exit Sub
    End If
    '  pub_cadena = "SELECT * FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? "
    pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_FAMILIA = ? ORDER BY ART_ALTERNO"
    Set PSCUO_VENDEDOR = CN.CreateQuery("", pub_cadena)
    PSCUO_VENDEDOR(0) = LK_CODCIA
    PSCUO_VENDEDOR(1) = Val(Right(cmbdivi.Text, 6))
    Set cuo_vendedor = PSCUO_VENDEDOR.OpenResultset(rdOpenKeyset, rdConcurValues)

    pub_cadena = "SELECT * FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? " & _
    "AND  CUO_DIVISION = ? AND CUO_CODART = ?"
    Set PSCUO_LLAVE = CN.CreateQuery("", pub_cadena)
    PSCUO_LLAVE(0) = 0
    PSCUO_LLAVE(1) = 0
    PSCUO_LLAVE(2) = Date
    PSCUO_LLAVE(3) = Date
    PSCUO_LLAVE(4) = 0
    PSCUO_LLAVE(5) = 0
    Set cuo_rep01 = PSCUO_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
     
    If Not cuo_vendedor.EOF Then
       pro.Min = 0
       pro.max = cuo_vendedor.RowCount
       pro.Value = 0
     End If
     Do Until cuo_vendedor.EOF
         pro.Value = pro.Value + 1
         gridigv.Rows = gridigv.Rows + 1
         gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!art_alterno
         gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!art_nombre
         gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_vendedor!ART_KEY
         PSCUO_LLAVE(0) = LK_CODCIA
         PSCUO_LLAVE(1) = loc_tipo
         PSCUO_LLAVE(2) = fecha1.Caption
         PSCUO_LLAVE(3) = fecha2.Caption
         PSCUO_LLAVE(4) = Val(Right(cmbdivi.Text, 6))
         PSCUO_LLAVE(5) = cuo_vendedor!ART_KEY
         cuo_rep01.Requery
         If Not cuo_rep01.EOF Then
          gridigv.TextMatrix(gridigv.Rows - 1, 2) = cuo_rep01!CUO_DESCRIP
          gridigv.TextMatrix(gridigv.Rows - 1, 3) = cuo_rep01!CUO_CANTIDAD / cuo_rep01!CUO_EQUIV
          gridigv.TextMatrix(gridigv.Rows - 1, 5) = cuo_rep01!CUO_MONTO
          gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_rep01!CUO_CODART
         End If
         
         cuo_vendedor.MoveNext
      Loop
Case 6 ' CUOTA POR EMPRESA
    cabe
    pub_cadena = "SELECT * FROM PARGEN WHERE PAR_CODCIA = ? ORDER BY PAR_CODCIA"
    Set PSCUO_VENDEDOR = CN.CreateQuery("", pub_cadena)
    PSCUO_VENDEDOR(0) = LK_CODCIA
    Set cuo_vendedor = PSCUO_VENDEDOR.OpenResultset(rdOpenKeyset, rdConcurValues)

    pub_cadena = "SELECT * FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? "
    Set PSCUO_LLAVE = CN.CreateQuery("", pub_cadena)
    PSCUO_LLAVE(0) = 0
    PSCUO_LLAVE(1) = 0
    PSCUO_LLAVE(2) = 0
    PSCUO_LLAVE(3) = 0
    Set cuo_rep01 = PSCUO_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
    If Not cuo_vendedor.EOF Then
      pro.Min = 0
      pro.max = cuo_vendedor.RowCount
      pro.Value = 0
    End If
    Do Until cuo_vendedor.EOF
            pro.Value = pro.Value + 1
            gridigv.Rows = gridigv.Rows + 1
            gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!PAR_CODCIA
            gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!PAR_NOMBRE
            PSCUO_LLAVE(0) = LK_CODCIA
            PSCUO_LLAVE(1) = loc_tipo
            PSCUO_LLAVE(2) = fecha1.Caption
            PSCUO_LLAVE(3) = fecha2.Caption
            cuo_rep01.Requery
            If Not cuo_rep01.EOF Then
             gridigv.TextMatrix(gridigv.Rows - 1, 3) = cuo_rep01!CUO_CANTIDAD / cuo_rep01!CUO_EQUIV
             gridigv.TextMatrix(gridigv.Rows - 1, 5) = cuo_rep01!CUO_MONTO
            End If
           cuo_vendedor.MoveNext
     Loop

End Select
fin:
DoEvents
lblpro.Visible = False
pro.Visible = False
DoEvents
End Sub


Private Sub Form_Load()
cmbdivi.ListIndex = -1
cmbdivi.BackColor = QBColor(7)
Txt_key.Text = ""
Txt_key.BackColor = QBColor(7)
cmbdivi.Enabled = False
Txt_key.Enabled = False

pub_cadena = "SELECT CUO_NUMSEC FROM CUOTAS WHERE CUO_CODCIA = ? AND CUO_TIPO = ? AND CUO_FECHA1 = ? AND CUO_FECHA2 = ? ORDER BY CUO_NUMSEC"
Set PSCUO_SIGUE = CN.CreateQuery("", pub_cadena)
PSCUO_SIGUE(0) = ""
PSCUO_SIGUE(1) = 0
PSCUO_SIGUE(2) = 0
PSCUO_SIGUE(3) = 0
Set cuo_sigue = PSCUO_SIGUE.OpenResultset(rdOpenKeyset, rdConcurValues)
    
Cmbtipos.AddItem "01 - Cuota por Vendedores"
Cmbtipos.AddItem "02 - Cuota por Vendedores y Articulos"
Cmbtipos.AddItem "03 - Cuota por Vendedores y Divisiones"
Cmbtipos.AddItem "04 - Cuota por Divisiones"
Cmbtipos.AddItem "05 - Cuota por Articulos"
Cmbtipos.AddItem "06 - Cuota por Empresa"
LLENA_GRUPOS cmbdivi, 122
mes.AddItem "Enero" & String(80, " ") & "01"
mes.AddItem "Febrero" & String(80, " ") & "02"
mes.AddItem "Marzo" & String(80, " ") & "03"
mes.AddItem "Abril" & String(80, " ") & "04"
mes.AddItem "Mayo" & String(80, " ") & "05"
mes.AddItem "Junio" & String(80, " ") & "06"
mes.AddItem "Julio" & String(80, " ") & "07"
mes.AddItem "Agosto" & String(80, " ") & "08"
mes.AddItem "Setiembre" & String(80, " ") & "09"
mes.AddItem "Octubre" & String(80, " ") & "10"
mes.AddItem "Noviembre" & String(80, " ") & "11"
mes.AddItem "Diciembre" & String(80, " ") & "12"
ano.AddItem "2002"
ano.AddItem "2003"
ano.AddItem "2004"
ano.AddItem "2005"
ano.AddItem "2006"
For fila = 1 To 12
  If Right(mes.List(fila), 2) = Format(LK_FECHA_DIA, "mm") Then mes.ListIndex = fila
Next fila
ano.Text = Format(LK_FECHA_DIA, "yyyy")

End Sub

Private Sub i_nomarti_Click()

End Sub

Public Sub LLENA_GRUPOS(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
    
End Sub


Private Sub graba_Click()
Dim WNUMSEC As Integer
If gridigv.Rows <= 2 Then
 Exit Sub
End If
If loc_tipo = 1 Then
 pub_cadena = "DELETE CUOTAS WHERE CUO_CODCIA = '" & LK_CODCIA & "' AND CUO_TIPO= " & loc_tipo & " AND CUO_MES = " & Right(mes.Text, 2) & " AND CUO_ANO= " & ano.Text & ""
ElseIf loc_tipo = 2 Then
 pub_cadena = "DELETE CUOTAS WHERE CUO_CODCIA = '" & LK_CODCIA & "' AND CUO_TIPO= " & loc_tipo & " AND CUO_MES = " & Right(mes.Text, 2) & " AND CUO_ANO= " & ano.Text & " AND CUO_CODVEN = " & Trim(Txt_key.Text) & " AND CUO_DIVISION = " & Trim(Right(cmbdivi.Text, 8)) & ""
ElseIf loc_tipo = 3 Then
 pub_cadena = "DELETE CUOTAS WHERE CUO_CODCIA = '" & LK_CODCIA & "' AND CUO_TIPO= " & loc_tipo & " AND CUO_MES = " & Right(mes.Text, 2) & " AND CUO_ANO= " & ano.Text & " AND CUO_CODVEN = " & Trim(Txt_key.Text) & ""
ElseIf loc_tipo = 4 Then
 pub_cadena = "DELETE CUOTAS WHERE CUO_CODCIA = '" & LK_CODCIA & "' AND CUO_TIPO= " & loc_tipo & " AND CUO_MES = " & Right(mes.Text, 2) & " AND CUO_ANO= " & ano.Text & ""
ElseIf loc_tipo = 5 Then
 pub_cadena = "DELETE CUOTAS WHERE CUO_CODCIA = '" & LK_CODCIA & "' AND CUO_TIPO= " & loc_tipo & " AND CUO_MES = " & Right(mes.Text, 2) & " AND CUO_ANO= " & ano.Text & " AND CUO_DIVISION = " & Trim(Right(cmbdivi.Text, 8)) & ""
ElseIf loc_tipo = 6 Then
 pub_cadena = "DELETE CUOTAS WHERE CUO_CODCIA = '" & LK_CODCIA & "' AND CUO_TIPO= " & loc_tipo & " AND CUO_MES = " & Right(mes.Text, 2) & " AND CUO_ANO= " & ano.Text & ""
End If
CN.Execute pub_cadena, rdExecDirect
PSCUO_SIGUE(0) = LK_CODCIA
PSCUO_SIGUE(1) = loc_tipo
PSCUO_SIGUE(2) = fecha1.Caption
PSCUO_SIGUE(3) = fecha2.Caption
cuo_sigue.Requery
If cuo_sigue.EOF Then
WNUMSEC = 0
Else
cuo_sigue.MoveLast
WNUMSEC = Val(cuo_sigue!CUO_NUMSEC)
End If

 For fila = 2 To gridigv.Rows - 1
  WNUMSEC = WNUMSEC + 1
  cuo_rep01.AddNew
  cuo_rep01!CUO_CODCIA = LK_CODCIA
  cuo_rep01!CUO_TIPO = loc_tipo
  cuo_rep01!CUO_FECHA1 = fecha1.Caption
  cuo_rep01!CUO_FECHA2 = fecha2.Caption
  cuo_rep01!CUO_NUMSEC = WNUMSEC
  cuo_rep01!CUO_DESCRIP = gridigv.TextMatrix(fila, 2)
  cuo_rep01!CUO_EQUIV = 1
  cuo_rep01!CUO_CANTIDAD = Val(gridigv.TextMatrix(fila, 3))
  cuo_rep01!CUO_MONTO = Val(gridigv.TextMatrix(fila, 5))
  cuo_rep01!CUO_CODVEN = 0
  If loc_tipo = 1 Then
   cuo_rep01!CUO_CODVEN = gridigv.TextMatrix(fila, 0)
  ElseIf loc_tipo = 2 Or loc_tipo = 3 Then
   cuo_rep01!CUO_CODVEN = Val(Txt_key.Text)
  End If
  cuo_rep01!CUO_CODART = Val(gridigv.TextMatrix(fila, 7))
  If loc_tipo = 3 Or loc_tipo = 4 Then
    cuo_rep01!CUO_DIVISION = Val(gridigv.TextMatrix(fila, 0))
  Else
    cuo_rep01!CUO_DIVISION = Val(Trim(Right(cmbdivi.Text, 8)))
  End If
  cuo_rep01!CUO_MES = Right(mes.Text, 2)
  cuo_rep01!CUO_ANO = ano.Text
  cuo_rep01.Update
Next fila
  MsgBox "Datos Actualizados", 48, Pub_Titulo
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
If gridigv.COL <= 2 Then Exit Sub
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
Dim SALE As Boolean
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


Private Sub mes_Click()
cal_fecha
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

'PUB_CAL_INI = gridigv.TextMatrix(gridigv.Row, 2)
'PUB_CAL_FIN = gridigv.TextMatrix(gridigv.Row, 2)
'pu_codcia = LK_CODCIA
'SQ_OPER = 1
'PUB_CODCIA = LK_CODCIA
'LEER_CAL_LLAVE
'cal_llave.Edit
'If gridigv.COL = 4 Then
'   cal_llave!cal_tipo_cambio = Val(TEXTOVAR.Text)
'End If
'If gridigv.COL = 1 Then
'   cal_llave!CAL_TC_MERCA = Val(TEXTOVAR.Text)
'   If Format(LK_FECHA_DIA, "dd/mm/yyyy") = Format(gridigv.TextMatrix(gridigv.Row, 0), "dd/mm/yyyy") Then
'      LK_TIPO_CAMBIO = Val(TEXTOVAR.Text)
'      'MDIForm1.StatusBar1.Panels(3).Text = "T.C.= S/. " + Format(LK_TIPO_CAMBIO, "0.0000")
'   End If
'End If
'If gridigv.COL = 5 Then
'   cal_llave!cal_tc_ingre = Val(TEXTOVAR.Text)
'End If
'If gridigv.COL = 6 Then
'   cal_llave!cal_tc_salid = Val(TEXTOVAR.Text)
'End If
'
'cal_llave.Update
'If gridigv.Row >= gridigv.Rows - 1 Then
'Else
'  gridigv.Row = gridigv.Row + 1
'End If
gridigv.SetFocus
TEXTOVAR.Visible = False

fin:

End Sub

Public Sub cabe()
gridigv.Clear
gridigv.Cols = 10
gridigv.Rows = 2
gridigv.ColWidth(0) = 800
gridigv.ColWidth(1) = 2500
gridigv.ColWidth(2) = 1000
gridigv.ColWidth(3) = 1000
gridigv.ColWidth(4) = 1000
gridigv.ColWidth(5) = 1000
gridigv.ColWidth(6) = 1000
gridigv.ColWidth(7) = 0
gridigv.ColWidth(8) = 0
gridigv.ColWidth(9) = 0


gridigv.TextMatrix(0, 0) = "Codigo"
gridigv.TextMatrix(1, 0) = ""
gridigv.TextMatrix(0, 1) = "Descripcion"
gridigv.TextMatrix(1, 1) = ""

gridigv.TextMatrix(0, 2) = "  "
gridigv.TextMatrix(1, 2) = "U.M"
gridigv.TextMatrix(0, 3) = "Cuota "
gridigv.TextMatrix(1, 3) = "Cantidad"

gridigv.TextMatrix(0, 4) = "Avance "
gridigv.TextMatrix(1, 4) = "Unidades"
gridigv.TextMatrix(0, 5) = "Cuota"
gridigv.TextMatrix(1, 5) = "Valor"
gridigv.TextMatrix(0, 6) = "Avance"
gridigv.TextMatrix(1, 6) = "Valor"

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


Public Sub cal_fecha()
Dim wdia As String
fecha1.Caption = "01/" & Right(mes.Text, 2) & "/" & ano.Text
If Right(mes.Text, 2) = "01" Then
   wdia = "31"
ElseIf Right(mes.Text, 2) = "02" Then
   If (Val(ano.Text) Mod 4) <> 0 Then
     wdia = "28"
   Else
     wdia = "29"
   End If
ElseIf Right(mes.Text, 2) = "03" Then
   wdia = "31"
ElseIf Right(mes.Text, 2) = "04" Then
   wdia = "30"
ElseIf Right(mes.Text, 2) = "05" Then
   wdia = "31"
ElseIf Right(mes.Text, 2) = "06" Then
   wdia = "30"
ElseIf Right(mes.Text, 2) = "07" Then
   wdia = "31"
ElseIf Right(mes.Text, 2) = "08" Then
   wdia = "31"
ElseIf Right(mes.Text, 2) = "09" Then
   wdia = "30"
ElseIf Right(mes.Text, 2) = "10" Then
   wdia = "31"
ElseIf Right(mes.Text, 2) = "11" Then
   wdia = "30"
ElseIf Right(mes.Text, 2) = "12" Then
   wdia = "31"
End If
fecha2.Caption = wdia & "/" & Right(mes.Text, 2) & "/" & ano.Text
gridigv.Rows = 2
End Sub

Private Sub Txt_key_Change()
If Txt_key.Text = "" Then
   lblven.Caption = ""
   gridigv.Rows = 2
End If
End Sub

Private Sub txt_key_GotFocus()
 Azul Txt_key, Txt_key
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
  Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
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
pu_codclie = Val(Txt_key.Text)
If Len(Txt_key.Text) = 0 Or Txt_key.Locked Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(Txt_key.Text) = True Then
   loc_key = 0
   On Error GoTo mucho
   PUB_CODVEN = Val(Txt_key.Text)
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_VEN_LLAVE
   On Error GoTo 0
   If ven_llave.EOF Then
     Azul Txt_key, Txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Txt_key.SetFocus
     GoTo fin
   End If
   lblven.Caption = Trim(ven_llave!VEM_NOMBRE)
   ListView1.Visible = False
   Screen.MousePointer = 0
   cmdmostrar.SetFocus
   Exit Sub
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(Txt_key.Text)) = Left(VALOR, Len(Trim(Txt_key.Text))) Then
   Else
      Exit Sub
   End If
   Txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
   PUB_CODVEN = Val(Txt_key.Text)
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_VEN_LLAVE
   On Error GoTo 0
   If ven_llave.EOF Then
     Azul Txt_key, Txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Txt_key.SetFocus
     GoTo fin
   End If
   lblven.Caption = Trim(ven_llave!VEM_NOMBRE)
   ListView1.Visible = False
   cmdmostrar.SetFocus
End If
dale:
mucho:
ListView1.Visible = False
fin:
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
If Len(Txt_key.Text) = 0 Or Txt_key.Locked = True Or IsNumeric(Txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(Txt_key.Text) = 1 Then
    var = Asc(Txt_key.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    numarchi = 9
    archi = "SELECT * FROM VEMAEST WHERE  VEM_CODCIA = '" & LK_CODCIA & "' AND VEM_NOMBRE BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY VEM_NOMBRE"
    PROC_LISVIEW ListView1
    loc_key = 1
    If ListView1.Visible = False Then
        loc_key = 0
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
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


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCuotas 
   Caption         =   "Definición de Cuotas."
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fratipo 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   11655
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
         TabIndex        =   19
         Top             =   420
         Width           =   1290
      End
      Begin VB.Label fecha1 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   10050
         TabIndex        =   18
         Top             =   135
         Width           =   1290
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Año:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8040
         TabIndex        =   17
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Mes:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   5880
         TabIndex        =   16
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccione Tipo de Cuota :"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8520
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   11655
      Begin VB.CommandButton Cmdclose 
         Caption         =   "&Cerrar"
         Height          =   435
         Left            =   9960
         TabIndex        =   8
         Top             =   7800
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   495
         Left            =   4920
         TabIndex        =   45
         Top             =   7680
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
      Begin VB.Frame fcopia 
         Caption         =   "Copia Datos"
         Height          =   1230
         Left            =   6255
         TabIndex        =   25
         Top             =   5880
         Visible         =   0   'False
         Width           =   3660
         Begin VB.CommandButton Command3 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2265
            TabIndex        =   31
            Top             =   705
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Copiar"
            Height          =   375
            Left            =   2265
            TabIndex        =   30
            Top             =   210
            Width           =   1095
         End
         Begin MSMask.MaskEdBox mefecha1 
            Height          =   345
            Left            =   960
            TabIndex        =   28
            ToolTipText     =   "Ingresar ""mm/yyyy"""
            Top             =   270
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   7
            Mask            =   "##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mefecha2 
            Height          =   330
            Left            =   960
            TabIndex        =   29
            ToolTipText     =   "Ingresar ""mm/yyyy"""
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   7
            Mask            =   "##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Caption         =   "Al Mes "
            Height          =   255
            Left            =   150
            TabIndex        =   27
            Top             =   765
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Del Mes "
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   345
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Reportar."
         Height          =   2175
         Left            =   0
         TabIndex        =   33
         Top             =   6240
         Width           =   11655
         Begin VB.CommandButton cmdava 
            Caption         =   "Reportar Avance"
            Height          =   540
            Left            =   9945
            TabIndex        =   42
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox txtcodart 
            Height          =   315
            Left            =   120
            TabIndex        =   41
            Top             =   630
            Width           =   1230
         End
         Begin VB.Frame flista 
            Height          =   1875
            Left            =   1530
            TabIndex        =   38
            Top             =   240
            Visible         =   0   'False
            Width           =   3705
            Begin VB.ListBox ltarti 
               Height          =   1230
               Left            =   105
               TabIndex        =   40
               Top             =   510
               Width           =   3495
            End
            Begin VB.CommandButton cmdcancelar 
               Caption         =   "Cancelar"
               Height          =   285
               Left            =   2235
               TabIndex        =   39
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.TextBox txtcodven 
            Height          =   315
            Left            =   5355
            TabIndex        =   37
            Top             =   570
            Width           =   1260
         End
         Begin VB.Frame flista2 
            Height          =   1875
            Left            =   6765
            TabIndex        =   34
            Top             =   285
            Visible         =   0   'False
            Width           =   2835
            Begin VB.ListBox ltvend 
               Height          =   1230
               Left            =   45
               TabIndex        =   36
               Top             =   525
               Width           =   2685
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Cancelar"
               Height          =   300
               Left            =   1425
               TabIndex        =   35
               Top             =   180
               Width           =   1245
            End
         End
         Begin VB.Label Label6 
            Caption         =   "Cod. Articulo"
            Height          =   300
            Left            =   210
            TabIndex        =   44
            Top             =   375
            Width           =   1005
         End
         Begin VB.Label Label7 
            Caption         =   "Cod. Vendedor"
            Height          =   330
            Left            =   5355
            TabIndex        =   43
            Top             =   300
            Width           =   1245
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Copiar Datos."
         Height          =   465
         Left            =   9960
         TabIndex        =   32
         Top             =   5640
         Width           =   1350
      End
      Begin VB.OptionButton OPP 
         Caption         =   "Por Cantidades Minimas"
         Height          =   390
         Index           =   0
         Left            =   1935
         TabIndex        =   24
         Top             =   5835
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.OptionButton OPP 
         Caption         =   "Valor Monetario"
         Height          =   375
         Index           =   1
         Left            =   3570
         TabIndex        =   23
         Top             =   5865
         Visible         =   0   'False
         Width           =   1875
      End
      Begin MSComctlLib.ProgressBar pro 
         Height          =   210
         Left            =   1230
         TabIndex        =   21
         Top             =   5235
         Visible         =   0   'False
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   420
         Left            =   6240
         TabIndex        =   15
         Top             =   1890
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   741
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
         Sorted          =   -1  'True
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
      Begin VB.CommandButton graba 
         Caption         =   "&Grabar"
         Height          =   435
         Left            =   285
         TabIndex        =   7
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton cmdmostrar 
         Caption         =   "&Mostrar"
         Height          =   510
         Left            =   5820
         TabIndex        =   3
         Top             =   360
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
         AllowUserResizing=   1
      End
      Begin VB.Label LBLCODVEN 
         Caption         =   "."
         Height          =   375
         Left            =   7080
         TabIndex        =   48
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lbltotal 
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
         Left            =   6840
         TabIndex        =   47
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Total S/.="
         Height          =   255
         Left            =   6000
         TabIndex        =   46
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   11400
         Y1              =   6120
         Y2              =   6120
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
         TabIndex        =   22
         Top             =   5265
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblven 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1860
         TabIndex        =   20
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
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor:"
         Height          =   420
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de Cuotas :"
         Height          =   195
         Left            =   195
         TabIndex        =   12
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
Dim fec1 As Date
Dim fec2 As Date

Public Sub val_fecha(fecha As String)
Dim wdia As String
fec1 = "01/" & fecha
If Left(fecha, 2) = "01" Then
    wdia = "31"
ElseIf Left(fecha, 2) = "02" Then
    If (Val(Right(Str(fecha), 4)) Mod 4) <> 0 Then
        wdia = "28"
    Else
        wdia = "29"
    End If
ElseIf Left(fecha, 2) = "03" Then
    wdia = "31"
ElseIf Left(fecha, 2) = "04" Then
    wdia = "30"
ElseIf Left(fecha, 2) = "05" Then
    wdia = "31"
ElseIf Left(fecha, 2) = "06" Then
    wdia = "30"
ElseIf Left(fecha, 2) = "07" Then
    wdia = "31"
ElseIf Left(fecha, 2) = "08" Then
    wdia = "31"
ElseIf Left(fecha, 2) = "09" Then
    wdia = "30"
ElseIf Left(fecha, 2) = "10" Then
    wdia = "31"
ElseIf Left(fecha, 2) = "11" Then
    wdia = "30"
ElseIf Left(fecha, 2) = "12" Then
    wdia = "31"
End If
    fec2 = wdia & "/" & fecha


End Sub


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

Private Sub cmdava_Click()
cmdava.Enabled = False

Dim WVAR_CANTI  As Currency
Dim WVAR_CUOTA  As Currency

Dim VAR_CANTI  As Currency
Dim VAR_CUOTA  As Currency

Dim uni_canti As Currency
Dim uni_cuota As Currency
Dim val_canti As Currency
Dim val_cuota As Currency

Dim canti_cli As Currency
Dim suma_cli As Currency

Dim val_cob As Currency
Dim sum_cob As Currency

Dim llave_pre As rdoResultset
Dim llave_art As rdoResultset
Dim llave_cli As rdoResultset
Dim llave_fami As rdoResultset
Dim llave_ven As rdoResultset
Dim llave_cuota As rdoResultset
Dim llave_data As rdoResultset
Dim sum_cantidad As Currency
Dim sum_cuota As Currency
Dim sum_por As Currency
Dim cadena As String
Dim WPAS As String
Dim ban1 As Integer
Dim ban2 As Integer
Dim condi As String
Dim codarti As String
Dim codven As String
Dim condi2 As String
Dim xl As Object
Dim C As Integer
Dim F As Long
Dim d As Long
Dim b As Long
Dim nven As Long
Dim idr As Integer

Dim total_cuota  As Currency
Dim total_canti As Currency
Dim total_cuota2 As Currency
Dim total_canti2 As Currency
Dim total_cli As Currency
Dim total_cob As Currency
Dim lista_vend As String
Dim lista_vend2 As String

If OPP(0).Value Then
idr = 0  ' X CANTIDADES MINIMAS
Else
idr = 1  ' X VALOR MONETARIO
End If
If ltvend.ListCount <> 1 Then
    cmdava.Enabled = True
    MsgBox "Digitar Solo un Codigo alterno de Vendedor."
    Exit Sub

End If
If Trim(Cmbtipos.Text) = "" Then
    cmdava.Enabled = True
    MsgBox "Seleccionar el Tipo de Cuota"
    Exit Sub
End If
If Trim(cmbdivi.Text) = "" Then
    cmdava.Enabled = True
    MsgBox "Seleccionar La Linea de Productos"
    Exit Sub
End If
'COND. PARA SELECCIONAR ART. DE LA LISTA
For C = 0 To ltarti.ListCount - 1
    DoEvents
    codarti = codarti & "," & Trim(Right(ltarti.List(C), 10))
Next
If codarti <> "" Then
    condi = "  art_key in ( " & Right(Trim(codarti), Len(codarti) - 1) & " )  "
End If
'COD. PARA SELECTCIONAR VEND. DE LA LISTA
codven = ""
cadena = "select * from VEMAEST WHERE VEM_ALTERNO = " & Val(Left(ltvend.List(0), 2))
Set X = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
Do Until X.EOF
 DoEvents
 codven = codven & "," & Trim(X!VEM_CODVEN)
 X.MoveNext
Loop
'For C = 0 To ltvend.ListCount - 1
'    DoEvents
'    codven = codven & "," & Trim(Right(ltvend.List(C), 10))
'Next
If codven <> "" Then
    condi2 = "  VEM_CODVEN in ( " & Right(Trim(codven), Len(codven) - 1) & " )  "
    lista_vend = Right(Trim(codven), Len(codven) - 1)
    
End If


DoEvents
GoSub WEXCEL
xl.Cells(1, 1) = Trim(GEN!GEN_NOMBRE)
If idr = 0 Then
xl.Cells(2, 1) = "AVANCE  X VENDEDOR POR CANTIDADES MINIMAS"
Else
xl.Cells(2, 1) = "AVANCE  X VENDEDOR POR VALOR MONETARIO S/."
End If
xl.Cells(3, 1) = "DEL : " & fecha1.Caption & "  AL : " & fecha2.Caption

F = 7

If Trim(condi2) <> "" Then
    cadena = "select top 1 cuotas.cuo_codven,vemaest.vem_nombre from vemaest , cuotas where vemaest.vem_codcia=cuotas.cuo_codcia and vemaest.vem_alterno=cuotas.cuo_codven  and cuotas.cuo_fecha1 >= '" & Format(fecha1, "dd/mm/yyyy") & "' and cuotas.cuo_fecha1 <= '" & Format(fecha2, "dd/mm/yyyy") & "' and cuotas.cuo_codcia='" & LK_CODCIA & "'  AND  VEM_ALTERNO = " & Val(Left(ltvend.List(0), 2)) & " "     '" & condi2 & "  "
    
Else
    cadena = "select distinct(cuotas.cuo_codven),vemaest.vem_nombre from vemaest , cuotas where vemaest.vem_codcia=cuotas.cuo_codcia and vemaest.vem_codven=cuotas.cuo_codven  and cuotas.cuo_fecha1 >= '" & Format(fecha1, "dd/mm/yyyy") & "' and cuotas.cuo_fecha1 <= '" & Format(fecha2, "dd/mm/yyyy") & "' and cuotas.cuo_codcia='" & LK_CODCIA & "' "
End If
'cadena = "select * from vemaest where vem_estado=0 and vem_codcia='" & LK_CODCIA & "' and vem_estado=0"
Set llave_ven = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)



ban1 = 0
ban2 = 0

nven = llave_ven.RowCount
nven = (nven * 3) + 3
If Not llave_ven.EOF Then

pro.Min = 0
pro.Value = 0
pro.Max = llave_ven.RowCount



pro.Visible = True
lblpro.Visible = True
DoEvents
End If
Do Until llave_ven.EOF
    pro.Value = pro.Value + 1
    DoEvents
    
    If Trim(condi2) <> "" Then
        xl.Cells(F, 1) = "Vendedores : " & lista_vend
    Else
        xl.Cells(F, 1) = Trim(llave_ven!cuo_codven) & "-" & Trim(llave_ven!VEM_NOMBRE)
        lista_vend = Trim(llave_ven!cuo_codven)
    End If
    F = F + 1
    'cadena = "select * from tablas where tab_codcia='" & LK_CODCIA & "' and tab_tipreg=121 and tab_codart= " & Val(Right(cmbdivi.Text, 10)) & ""
    cadena = "select * from tablas where tab_codcia='" & LK_CODCIA & "' and tab_tipreg=122 and tab_numtab= " & Val(Right(cmbdivi.Text, 10)) & ""
    Set llave_fami = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
    
    Do Until llave_fami.EOF
        DoEvents
               
        
'        If Trim(condi) = "" Then
'            cadena = "SELECT * FROM ARTI, PRECIOS WHERE ART_CODCIA = '" & LK_CODCIA & "'  AND ART_FAMILIA = " & Val(Right(cmbdivi.Text, 10)) & " AND ART_CALIDAD = 1 AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART) AND (PRE_EQUIV = 1) and art_familia=" & Val(Right(cmbdivi.Text, 10)) & "  and art_subfam = " & Val(llave_fami!TAB_NUMTAB) & "       ORDER BY ART_NOMBRE" 'and art_situacion <> 1
'        Else
'            cadena = "SELECT * FROM ARTI, PRECIOS WHERE ART_CODCIA = '" & LK_CODCIA & "'  AND ART_FAMILIA = " & Val(Right(cmbdivi.Text, 10)) & " AND ART_CALIDAD = 1 AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART) AND (PRE_EQUIV = 1) and " & condi & "   and art_familia=" & Val(Right(cmbdivi.Text, 10)) & "  and art_subfam = " & Val(llave_fami!TAB_NUMTAB) & "    ORDER BY ART_NOMBRE"  'and art_situacion <> 1
'        End If
        
        If Trim(condi) = "" Then
            cadena = "SELECT * FROM ARTI, PRECIOS WHERE ART_CODCIA = '" & LK_CODCIA & "'  AND ART_FAMILIA = " & Val(Right(cmbdivi.Text, 10)) & " AND ART_CALIDAD = 1 AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART) AND (PRE_FLAG_UNIDAD = 'A') and art_situacion <> '1' and art_familia = " & Val(Right(cmbdivi.Text, 10)) & "      ORDER BY ART_NOMBRE" 'and art_situacion <> 1
        Else
            cadena = "SELECT * FROM ARTI, PRECIOS WHERE ART_CODCIA = '" & LK_CODCIA & "'  AND ART_FAMILIA = " & Val(Right(cmbdivi.Text, 10)) & " AND ART_CALIDAD = 1 AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART) AND (PRE_FLAG_UNIDAD = 'A') and art_situacion <> '1' and " & condi & "   and art_familia=" & Val(Right(cmbdivi.Text, 10)) & "   ORDER BY ART_NOMBRE"  'and art_situacion <> 1
        End If
    
        Set llave_art = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
        
        If llave_art.RowCount = 0 Then
            GoTo pasafami
        End If
        
        xl.Cells(F, 1) = "Familia :"
        xl.Cells(F, 2) = llave_fami!TAB_NOMLARGO
        F = F + 1
        Do Until llave_art.EOF
            DoEvents

            If ban2 = 0 Then
                xl.Cells(F, 1) = Trim(llave_art!ART_alterno)
                xl.Cells(F, 2) = Trim(llave_art!art_nombre)
                xl.Cells(F, 3) = Trim(llave_art!PRE_UNIDAD)

            End If
            
            'cadena = "select count(DISTINCT(FAR_codclie)) as cantidad  from FACART where FAR_codcia='" & LK_CODCIA & "' and FAR_fecha >= '" & Format(fecha1, "dd/mm/yyyy") & "' and FAR_fecha <= '" & Format(fecha2, "dd/mm/yyyy") & "' and FAR_codven=" & llave_ven!CUO_CODVEN & "   and  FAR_codart=" & llave_art!ART_KEY & " AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10  AND FAR_CODART<>0 AND FAR_ESTADO2 <> 'L'"
            cadena = "select count(DISTINCT(FAR_codclie)) as cantidad  from FACART where FAR_codcia='" & LK_CODCIA & "' and FAR_fecha >= '" & Format(fecha1, "dd/mm/yyyy") & "' and FAR_fecha <= '" & Format(fecha2, "dd/mm/yyyy") & "' and FAR_codven in " & "(" & lista_vend & ")" & "   and  FAR_codart=" & llave_art!art_key & " AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10  AND FAR_CODART<>0 AND FAR_ESTADO2 <> 'L'"
            Set llave_cli = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
            canti_cli = Nulo_Valor0(llave_cli!cantidad)
            
            'If Trim(llave_art!ART_alterno) = 301 Then Stop
                        
            'cadena = "select *  from cuotas  where cuo_codcia='" & LK_CODCIA & "' and cuo_codart=" & llave_art!ART_KEY & " and cuo_fecha1 >= '" & Format(fecha1, "dd/mm/yyyy") & "' and cuo_fecha2 <= '" & Format(fecha2, "dd/mm/yyyy") & "' and cuo_codven=" & llave_ven!cuo_codven & "   order by cuo_codart " Val(Left(ltvend.List(0), 2))
            'cadena = "select *  from cuotas  where cuo_codcia='" & LK_CODCIA & "' and cuo_codart=" & llave_art!art_key & " and cuo_fecha1 >= '" & Format(fecha1, "dd/mm/yyyy") & "' and cuo_fecha2 <= '" & Format(fecha2, "dd/mm/yyyy") & "' and cuo_codven in " & "(" & lista_vend & ")" & "   order by cuo_codart "
            cadena = "select *  from cuotas  where cuo_codcia='" & LK_CODCIA & "' and cuo_codart=" & llave_art!art_key & " and cuo_fecha1 >= '" & Format(fecha1, "dd/mm/yyyy") & "' and cuo_fecha2 <= '" & Format(fecha2, "dd/mm/yyyy") & "' and cuo_codven in " & "(" & Val(Left(ltvend.List(0), 2)) & ")" & "   order by cuo_codart "
            Set llave_cuota = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
            If llave_cuota.EOF Then
                VAR_CANTI = 0
                VAR_CUOTA = 0
                val_cob = 0
            Else
                VAR_CANTI = Val(llave_cuota!CUO_CANTIDAD) / Val(llave_art!PRE_EQUIV)
'*******************************************************
                cadena = "select * from precios where pre_codcia='" & LK_CODCIA & "' and pre_codart= " & llave_art!art_key & " AND PRE_FLAG_UNIDAD = 'A'"
                Set llave_pre = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
                VAR_CUOTA = Val(llave_pre!PRE_PRE1) * Nulo_Valor0(VAR_CANTI)
'********************************************************

                'VAR_CUOTA = Val(llave_cuota!CUO_MONTO)
                val_cob = Nulo_Valor0(llave_cuota!cuo_cobertura)
            End If
            ' PARA VENTAS
            ' X CANTIDADES MIN.
            'cadena = "select sum(FAR_cantidad) as cantidad  from FACART where FAR_codcia='" & LK_CODCIA & "' and FAR_fecha >= '" & Format(fecha1, "dd/mm/yyyy") & "' and FAR_fecha <= '" & Format(fecha2, "dd/mm/yyyy") & "' and FAR_codven=" & llave_ven!cuo_codven & "   and  FAR_codart=" & llave_art!ART_KEY & " AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10  AND FAR_CODART<>0  AND FAR_ESTADO2 <> 'L'"
            cadena = "select sum(FAR_cantidad) as cantidad  from FACART where FAR_codcia='" & LK_CODCIA & "' and FAR_fecha >= '" & Format(fecha1, "dd/mm/yyyy") & "' and FAR_fecha <= '" & Format(fecha2, "dd/mm/yyyy") & "' and FAR_codven in " & "(" & lista_vend & ")" & "   and  FAR_codart=" & llave_art!art_key & " AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10  AND FAR_CODART<>0  AND FAR_ESTADO2 <> 'L'"
            Set llave_data = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
            If llave_data.EOF Then
            Else
                WVAR_CANTI = WVAR_CANTI + (Val(Nulo_Valor0(llave_data!cantidad)) / Val(llave_art!PRE_EQUIV))
            End If
            ' X VALOR MONETARIO
            'cadena = "select sum( (FAR_cantidad/FAR_EQUIV) * FAR_PRECIO) as cantidad  from FACART where FAR_codcia='" & LK_CODCIA & "' and FAR_fecha >= '" & Format(fecha1, "dd/mm/yyyy") & "' and FAR_fecha <= '" & Format(fecha2, "dd/mm/yyyy") & "' and FAR_codven=" & llave_ven!cuo_codven & "   and  FAR_codart=" & llave_art!ART_KEY & " AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10  AND FAR_CODART<>0 "
            cadena = "select sum( (FAR_cantidad/FAR_EQUIV) * FAR_PRECIO) as cantidad  from FACART where FAR_codcia='" & LK_CODCIA & "' and FAR_fecha >= '" & Format(fecha1, "dd/mm/yyyy") & "' and FAR_fecha <= '" & Format(fecha2, "dd/mm/yyyy") & "' and FAR_codven in " & "(" & lista_vend & ")" & "   and  FAR_codart=" & llave_art!art_key & " AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10  AND FAR_CODART<>0 AND FAR_ESTADO2 <> 'L' "
            Set llave_data = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
            If llave_data.EOF Then
            Else
                WVAR_CUOTA = WVAR_CUOTA + Val(Nulo_Valor0(llave_data!cantidad))
            End If
            ' MENOS NOTA DE CREDITO
            'cadena = "select sum(FAR_cantidad/FAR_EQUIV) as cantidad  from FACART where FAR_codcia='" & LK_CODCIA & "' and FAR_fecha >= '" & Format(fecha1, "dd/mm/yyyy") & "' and FAR_fecha <= '" & Format(fecha2, "dd/mm/yyyy") & "' and FAR_codven=" & llave_ven!cuo_codven & "   and  FAR_codart=" & llave_art!ART_KEY & " AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 97 AND FAR_CODART<>0 "
            cadena = "select sum(FAR_cantidad/FAR_EQUIV) as cantidad  from FACART where FAR_codcia='" & LK_CODCIA & "' and FAR_fecha >= '" & Format(fecha1, "dd/mm/yyyy") & "' and FAR_fecha <= '" & Format(fecha2, "dd/mm/yyyy") & "' and FAR_codven in " & "(" & lista_vend & ")" & "   and  FAR_codart=" & llave_art!art_key & " AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 97 AND FAR_CODART <> 0 AND FAR_ESTADO2 <> 'L' "
            Set llave_data = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
            If llave_data.EOF Then
            Else
                WVAR_CANTI = WVAR_CANTI + (Val(Nulo_Valor0(llave_data!cantidad)) * -1)
            End If
            'cadena = "select  sum( (FAR_cantidad/FAR_EQUIV) * FAR_PRECIO) as cantidad  from FACART where FAR_codcia='" & LK_CODCIA & "' and FAR_fecha >= '" & Format(fecha1, "dd/mm/yyyy") & "' and FAR_fecha <= '" & Format(fecha2, "dd/mm/yyyy") & "' and FAR_codven=" & llave_ven!cuo_codven & "   and  FAR_codart=" & llave_art!ART_KEY & " AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 97 AND FAR_CODART<>0 "
            cadena = "select  sum( (FAR_cantidad/FAR_EQUIV) * FAR_PRECIO) as cantidad  from FACART where FAR_codcia='" & LK_CODCIA & "' and FAR_fecha >= '" & Format(fecha1, "dd/mm/yyyy") & "' and FAR_fecha <= '" & Format(fecha2, "dd/mm/yyyy") & "' and FAR_codven in " & "(" & lista_vend & ")" & "   and  FAR_codart=" & llave_art!art_key & " AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 97 AND FAR_CODART <> 0 AND FAR_ESTADO2 <> 'L' "
            Set llave_data = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
            If llave_data.EOF Then
            Else
                WVAR_CUOTA = WVAR_CUOTA + (Val(Nulo_Valor0(llave_data!cantidad)) * -1)
            End If
                   
                   
                   
            xl.Cells(F, 4) = Val(VAR_CANTI)
            xl.Cells(F, 5) = Val(WVAR_CANTI)
            xl.Cells(F, 7) = Val(VAR_CUOTA)
            xl.Cells(F, 8) = Val(WVAR_CUOTA)
            If Val(VAR_CANTI) = 0 Then
            xl.Cells(F, 4) = 0
                
            xl.Cells(F, 6) = 0
            Else
            xl.Cells(F, 6) = Val(Val(WVAR_CANTI) * 100) / Val(VAR_CANTI)
            End If
            If Val(VAR_CUOTA) = 0 Then
            xl.Cells(F, 7) = 0
                
            xl.Cells(F, 9) = 0
            Else
            xl.Cells(F, 9) = Val(Val(WVAR_CUOTA) * 100) / Val(VAR_CUOTA)
            End If
            xl.Cells(F, 10) = Val(val_cob)
        
            xl.Cells(F, 11) = Val(canti_cli)
                    
            If Val(val_cob) = 0 Then
                xl.Cells(F, 12) = 0
            Else
                xl.Cells(F, 12) = Val(Val(canti_cli) * 100) / Val(val_cob)
            End If
            uni_canti = uni_canti + Val(WVAR_CANTI)
            uni_cuota = uni_cuota + Val(VAR_CANTI)
        
            val_canti = val_canti + Val(WVAR_CUOTA)
            val_cuota = val_cuota + Val(VAR_CUOTA)
            
            suma_cli = suma_cli + canti_cli
            
            sum_cob = sum_cob + val_cob
            
            llave_art.MoveNext
            VAR_CANTI = 0
            VAR_CUOTA = 0
            WVAR_CANTI = 0
            WVAR_CUOTA = 0
            canti_cli = 0
            val_cob = 0
            
            F = F + 1
        Loop
    
        'F = F + 1
    
        xl.Cells(F, 2) = "TOTAL FAMILIA :"
    
        xl.Cells(F, 4) = Val(uni_cuota)
        xl.Cells(F, 5) = Val(uni_canti)
        xl.Cells(F, 7) = Val(val_cuota)
        xl.Cells(F, 8) = Val(val_canti)
        
        xl.Cells(F, 11) = Val(suma_cli)
        xl.Cells(F, 10) = Val(sum_cob)
        
        total_cuota = total_cuota + Val(uni_cuota)
        total_canti = total_canti + Val(uni_canti)
        total_cuota2 = total_cuota2 + Val(val_cuota)
        total_canti2 = total_canti2 + Val(val_canti)
        total_cli = total_cli + Val(suma_cli)
        total_cob = total_cob + Val(sum_cob)
        
        uni_cuota = 0
        uni_canti = 0
        val_cuota = 0
        val_canti = 0
        suma_cli = 0
        sum_cob = 0
        
        F = F + 2
pasafami:
        llave_fami.MoveNext
    Loop
    F = F + 1
    llave_ven.MoveNext
Loop
F = F + 1
xl.Cells(F, 2) = "TOTAL FAMILIA :"
xl.Cells(F, 4) = Val(total_cuota)
xl.Cells(F, 5) = Val(total_canti)
xl.Cells(F, 7) = Val(total_cuota2)
xl.Cells(F, 8) = Val(total_canti2)

xl.Cells(F, 11) = Val(total_cli)
xl.Cells(F, 10) = Val(total_cob)

xl.Application.Visible = True
pro.Visible = False
lblpro.Visible = False
DoEvents
cmdava.Enabled = True
Set xl = Nothing
Screen.MousePointer = 0
Exit Sub

WEXCEL:
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  'FrmImp2.lblProceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
  DoEvents
  WPAS = ""
  'WPAS = PUB_RUTA_OTRO + "CAJA_DET.xls"
  'DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) & "\ADMIN\STANDAR\REPCUOTA2.xls", 0, True, 4, WPAS, WPAS

  'xl.Workbooks.Open , "C:\ADMIN\HERTISA\CAJA_DET.XLS", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

CANCELA:
Exit Sub


End Sub

Private Sub cmdcancelar_Click()
ltarti.Clear
flista.Visible = False
End Sub

Private Sub cmdClose_Click()
Unload Me

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
      pro.Max = cuo_vendedor.RowCount
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
   ' cuota por Lineas x Vendedor
   '*****************************
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
    pub_cadena = "SELECT * FROM ARTI, PRECIOS WHERE ART_CODCIA = ? AND ART_FAMILIA = ? AND ART_CALIDAD = 1 AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART) AND (PRE_FLAG_UNIDAD = 'A') AND ART_SITUACION <> '1'  ORDER BY ART_NOMBRE"
'    pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = 123 AND TAB_CODART = ? "
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
      pro.Max = cuo_vendedor.RowCount
      pro.Value = 0
    End If
    Do Until cuo_vendedor.EOF
         pro.Value = pro.Value + 1
         gridigv.Rows = gridigv.Rows + 1
         gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!ART_alterno
         gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!art_nombre
         gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_vendedor!art_key
         gridigv.TextMatrix(gridigv.Rows - 1, 2) = cuo_vendedor!PRE_UNIDAD
         gridigv.TextMatrix(gridigv.Rows - 1, 10) = 1 'cuo_vendedor!PRE_EQUIV
         PSCUO_LLAVE(0) = LK_CODCIA
         PSCUO_LLAVE(1) = loc_tipo
         PSCUO_LLAVE(2) = fecha1.Caption
         PSCUO_LLAVE(3) = fecha2.Caption
         PSCUO_LLAVE(4) = Val(Txt_key.Text)
         PSCUO_LLAVE(5) = Val(Right(cmbdivi.Text, 6))
'         PSCUO_LLAVE(6) = cuo_vendedor!TAB_NUMTAB
         PSCUO_LLAVE(6) = cuo_vendedor!art_key
         cuo_rep01.Requery
         If Not cuo_rep01.EOF Then
           gridigv.TextMatrix(gridigv.Rows - 1, 3) = cuo_rep01!CUO_CANTIDAD / cuo_rep01!CUO_EQUIV
           gridigv.TextMatrix(gridigv.Rows - 1, 5) = cuo_rep01!CUO_MONTO
           gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_rep01!cuo_codart
           gridigv.TextMatrix(gridigv.Rows - 1, 6) = Nulo_Valor0(cuo_rep01!cuo_cobertura)
         End If
         gridigv.TextMatrix(gridigv.Rows - 1, 11) = cuo_vendedor!PRE_EQUIV
         gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_vendedor!art_key
         
         cuo_vendedor.MoveNext
         
      Loop
    sumaval
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
      pro.Max = cuo_vendedor.RowCount
      pro.Value = 0
    End If
    Do Until cuo_vendedor.EOF
         pro.Value = pro.Value + 1
         gridigv.Rows = gridigv.Rows + 1
         gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!TAB_NUMTAB
         gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!TAB_NOMLARGO
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
           gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_rep01!cuo_codart
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
      pro.Max = cuo_vendedor.RowCount
      pro.Value = 0
    End If
    Do Until cuo_vendedor.EOF
        pro.Value = pro.Value + 1
         gridigv.Rows = gridigv.Rows + 1
         gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!TAB_NUMTAB
         gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!TAB_NOMLARGO
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
           gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_rep01!cuo_codart
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
    'pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_FAMILIA = ? ORDER BY ART_ALTERNO"
    pub_cadena = "SELECT * FROM ARTI, PRECIOS WHERE (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART) AND  ART_CODCIA = ? AND ART_FAMILIA = ? and ART_CALIDAD = 1 AND (PRE_EQUIV = 1)  ORDER BY ART_NOMBRE"
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
       pro.Max = cuo_vendedor.RowCount
       pro.Value = 0
     End If
     Do Until cuo_vendedor.EOF
         pro.Value = pro.Value + 1
         gridigv.Rows = gridigv.Rows + 1
         gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!ART_alterno
         gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!art_nombre
         gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_vendedor!art_key
         gridigv.TextMatrix(gridigv.Rows - 1, 2) = cuo_vendedor!PRE_UNIDAD
         PSCUO_LLAVE(0) = LK_CODCIA
         PSCUO_LLAVE(1) = loc_tipo
         PSCUO_LLAVE(2) = fecha1.Caption
         PSCUO_LLAVE(3) = fecha2.Caption
         PSCUO_LLAVE(4) = Val(Right(cmbdivi.Text, 6))
         PSCUO_LLAVE(5) = cuo_vendedor!art_key
         cuo_rep01.Requery
         If Not cuo_rep01.EOF Then
          gridigv.TextMatrix(gridigv.Rows - 1, 2) = cuo_vendedor!PRE_UNIDAD
          gridigv.TextMatrix(gridigv.Rows - 1, 3) = cuo_rep01!CUO_CANTIDAD / cuo_rep01!CUO_EQUIV
          gridigv.TextMatrix(gridigv.Rows - 1, 5) = cuo_rep01!CUO_MONTO
          gridigv.TextMatrix(gridigv.Rows - 1, 7) = cuo_rep01!cuo_codart
          gridigv.TextMatrix(gridigv.Rows - 1, 6) = cuo_rep01!cuo_cobertura
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
      pro.Max = cuo_vendedor.RowCount
      pro.Value = 0
    End If
    Do Until cuo_vendedor.EOF
            pro.Value = pro.Value + 1
            gridigv.Rows = gridigv.Rows + 1
            gridigv.TextMatrix(gridigv.Rows - 1, 0) = cuo_vendedor!par_codcia
            gridigv.TextMatrix(gridigv.Rows - 1, 1) = cuo_vendedor!par_nombre
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


Private Sub Command1_Click()
ltvend.Clear
flista2.Visible = False
End Sub

Private Sub Command2_Click()
Dim fecha As String
Dim w_copia1 As Date
Dim w_copia2 As Date
Dim w_fecha1 As Date
Dim w_fecha2 As Date
Dim cadena As String
Dim llave_busca As rdoResultset

If IsDate(Format(mefecha1.Text, "mm/yyyy")) = False Then
    MsgBox "Fecha Invalida", 48, Pub_Titulo
    Azul2 Me.fecha1, mefecha1
    Exit Sub
End If
If IsDate(Format(mefecha2.Text, "mm/yyyy")) = False Then
    MsgBox "Fecha Invalida", 48, Pub_Titulo
    Azul2 Me.fecha2, mefecha2
    Exit Sub
End If

Call val_fecha(mefecha1.Text)
w_copia1 = fec1
w_copia2 = fec2
Call val_fecha(mefecha2.Text)
w_fecha1 = fec1
w_fecha2 = fec2




cadena = "select * into temp_cuotas from cuotas where cuo_fecha1='" & Format(w_copia1, "dd/mm/yyyy") & "' and cuo_fecha2 = '" & Format(w_copia2, "dd/mm/yyyy") & "' and cuo_codcia='" & LK_CODCIA & "' "
CN.Execute cadena

cadena = "select * from cuotas where cuo_fecha1='" & Format(w_fecha1, "dd/mm/yyyy") & "' and cuo_fecha2 = '" & Format(w_fecha2, "dd/mm/yyyy") & "' and cuo_codcia='" & LK_CODCIA & "'"
Set llave_busca = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)

If Not (llave_busca.EOF) = True Then
    Pub_Respuesta = MsgBox(" ¿Desea Remplazar Los Datos Existemtes  ... ?", Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then Exit Sub
    cadena = "delete from cuotas where cuo_fecha1='" & Format(w_fecha1, "dd/mm/yyyy") & "' and cuo_fecha2 = '" & Format(w_fecha2, "dd/mm/yyyy") & "' and cuo_codcia='" & LK_CODCIA & "' "
    CN.Execute cadena
Else
End If
cadena = "update temp_cuotas set cuo_fecha1='" & Format(w_fecha1, "dd/mm/yyyy") & "' , cuo_fecha2='" & Format(w_fecha2, "dd/mm/yyyy") & "' , cuo_mes = " & Format(fec1, "mm") & " "
CN.Execute cadena
cadena = "insert into  cuotas  select * from  temp_cuotas "
CN.Execute cadena
cadena = "drop table temp_cuotas"
CN.Execute cadena
MsgBox "Registros Copiados "

End Sub

Private Sub Command3_Click()
fcopia.Visible = False

End Sub

Private Sub Command4_Click()
fcopia.Visible = True
End Sub

Private Sub Form_Load()
CenterMe Me
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
    
'Cmbtipos.AddItem "01 - Cuota por Vendedores"
Cmbtipos.AddItem "02 - Cuota por Vendedores y Articulos"
'Cmbtipos.AddItem "03 - Cuota por Vendedores y Divisiones"
'Cmbtipos.AddItem "04 - Cuota por Divisiones"
'Cmbtipos.AddItem "05 - Cuota por Articulos"
'Cmbtipos.AddItem "06 - Cuota por Empresa"
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

PS_GEN(0) = 0
GEN.Requery
'Pub_Titulo = Trim(GEN!GEN_NOMBRE)
'PUB_RUTA_REPORTE = Trim(Nulo_Valors(GEN!gen_ruta_reportes))
PUB_RUTA_OTRO = Trim(Nulo_Valors(GEN!gen_ruta_otros))


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
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
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
  cuo_rep01!CUO_EQUIV = Val(gridigv.TextMatrix(fila, 11))
  cuo_rep01!CUO_CANTIDAD = Val(gridigv.TextMatrix(fila, 3)) * Val(gridigv.TextMatrix(fila, 11))
  cuo_rep01!CUO_MONTO = Val(gridigv.TextMatrix(fila, 5))
  cuo_rep01!cuo_cobertura = Val(gridigv.TextMatrix(fila, 6))
  
  cuo_rep01!cuo_codven = 0
  If loc_tipo = 1 Then
   cuo_rep01!cuo_codven = gridigv.TextMatrix(fila, 0)
  ElseIf loc_tipo = 2 Or loc_tipo = 3 Then
   cuo_rep01!cuo_codven = Val(Txt_key.Text)
  End If
  cuo_rep01!cuo_codart = Val(gridigv.TextMatrix(fila, 7))
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
If gridigv.COL <= 2 Or gridigv.COL = 5 Then Exit Sub
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


Private Sub ltarti_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If ltarti.ListCount <> 0 Then
        If ltarti.ListIndex = -1 Then Exit Sub
        ltarti.RemoveItem (ltarti.ListIndex)
        If ltarti.ListCount = 0 Then flista.Visible = False
        
    End If
End If
End Sub

Private Sub ltvend_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If ltvend.ListCount <> 0 Then
        If ltvend.ListIndex = -1 Then Exit Sub
        ltvend.RemoveItem (ltvend.ListIndex)
        If ltvend.ListCount = 0 Then flista2.Visible = False
        
    End If
End If
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

Dim cadena As String
Dim CODLISTA As Integer
Dim VPRECIO As Currency

cadena = "SELECT  * FROM VEMAEST WHERE VEM_ALTERNO = " & Val(Txt_key.Text) & " AND VEM_CODCIA = '" & LK_CODCIA & "'"
Set X = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
CODLISTA = 2 'Val(X!VEM_AM_LISTA)
cadena = "select * from precios where pre_codcia='" & LK_CODCIA & "' and pre_codart= " & Val(gridigv.TextMatrix(gridigv.Row, 7)) & " AND PRE_FLAG_UNIDAD = 'A'"
Set X = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurValues)
VPRECIO = 0
If CODLISTA = 1 Then
VPRECIO = X!PRE_PRE1
ElseIf CODLISTA = 2 Then
VPRECIO = X!PRE_PRE2
ElseIf CODLISTA = 3 Then
VPRECIO = X!PRE_PRE3
ElseIf CODLISTA = 4 Then
VPRECIO = X!PRE_PRE4
ElseIf CODLISTA = 5 Then
VPRECIO = X!PRE_PRE5
ElseIf CODLISTA = 6 Then
VPRECIO = X!PRE_PRE6
ElseIf CODLISTA = 7 Then
VPRECIO = X!PRE_PRE7
ElseIf CODLISTA = 8 Then
VPRECIO = X!PRE_PRE8
ElseIf CODLISTA = 9 Then
VPRECIO = X!PRE_PRE9
ElseIf CODLISTA = 10 Then
VPRECIO = X!PRE_PRECHI1
ElseIf CODLISTA = 11 Then
VPRECIO = X!PRE_PREC1
ElseIf CODLISTA = 12 Then
VPRECIO = X!PRE_PREC2
End If
gridigv.TextMatrix(gridigv.Row, 5) = Val(VPRECIO) * Val(gridigv.TextMatrix(gridigv.Row, 3))
sumaval

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
gridigv.Cols = 12
gridigv.Rows = 2
gridigv.ColWidth(0) = 800
gridigv.ColWidth(1) = 4500
gridigv.ColWidth(2) = 1000
gridigv.ColWidth(3) = 1000
gridigv.ColWidth(4) = 0
gridigv.ColWidth(5) = 1000
gridigv.ColWidth(6) = 1000
gridigv.ColWidth(7) = 0
gridigv.ColWidth(8) = 0
gridigv.ColWidth(9) = 0
gridigv.ColWidth(10) = 0
gridigv.ColWidth(11) = 0 ' EQUIVALENCIA



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
gridigv.TextMatrix(0, 6) = "Cobertura"
gridigv.TextMatrix(1, 6) = "Objetivo"

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
Dim CADE_VENDEDORES As String
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
   pub_cadena = "SELECT  * FROM VEMAEST WHERE VEM_ALTERNO = " & Val(Txt_key.Text) & " AND VEM_CODCIA = '" & LK_CODCIA & "'"
   Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
   CADE_VENDEDORES = ""
   Do Until X.EOF
     CADE_VENDEDORES = CADE_VENDEDORES + Format(X!VEM_CODVEN, "00") + ", "
     X.MoveNext
   Loop
   LBLCODVEN.Caption = "Codigos Relacionados: " & CADE_VENDEDORES
   X.MoveFirst
   If X.EOF Then
     Azul Txt_key, Txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Txt_key.SetFocus
     GoTo fin
   End If
   lblven.Caption = Trim(X!VEM_NOMBRE)
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
   pub_cadena = "SELECT  * FROM VEMAEST WHERE VEM_ALTERNO = " & Val(Txt_key.Text) & " AND VEM_CODCIA = '" & LK_CODCIA & "'"
   Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
   On Error GoTo 0
   If X.EOF Then
     Azul Txt_key, Txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Txt_key.SetFocus
     GoTo fin
   End If
   lblven.Caption = Trim(X!VEM_NOMBRE)
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

Private Sub txtcodart_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim llave_busca3 As rdoResultset
    Dim cade As String
    Dim wart As String
    Dim arti As String
    Dim C As Integer
    If Trim(Cmbtipos.Text) = "" Then
        MsgBox "Seleccionar el Tipo de Cuota"
        Exit Sub
    End If
    If Trim(cmbdivi.Text) = "" Then
        MsgBox "Seleccionar La Linea de Productos"
        Exit Sub
    End If
    cade = "select * from arti where art_alterno='" & Trim(txtcodart.Text) & "' and art_codcia='" & LK_CODCIA & "' and art_familia=" & Val(Right(cmbdivi, 10)) & " "
    Set llave_busca3 = CN.OpenResultset(cade, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    If llave_busca3.EOF Then
        MsgBox "EL ARTICULO NO EXISTE ", 48, Pub_Titulo
        txtcodart.SetFocus
        Exit Sub
    Else
        wart = Trim(Str(llave_busca3!art_key))
        For C = 0 To ltarti.ListCount - 1
            DoEvents
            arti = Trim(Right(ltarti.List(C), 10))
            If arti = wart Then
                MsgBox "EL ARTICULO YA EXISTE EN LA LISTA"
                Exit Sub
            End If
        Next
        'ltarti.Visible = True
        flista.Visible = True
        ltarti.AddItem (Trim(llave_busca3!ART_alterno) & " - " & Trim(llave_busca3!art_nombre)) & Space(60) & Trim(llave_busca3!art_key)
        
        
        txtcodart.Text = ""
        
'        Me.cmdprocesar.SetFocus
    End If
    
End If
End Sub

Private Sub txtcodven_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim llave_busca3 As rdoResultset
    Dim cade As String
    Dim wvend As Integer
    Dim vend As Integer
    Dim C As Integer
    If Trim(Cmbtipos.Text) = "" Then
        MsgBox "Seleccionar el Tipo de Cuota"
        Exit Sub
    End If
    If Trim(cmbdivi.Text) = "" Then
        MsgBox "Seleccionar La Linea de Productos"
        Exit Sub
    End If
    cade = "select * from vemaest where vem_alterno ='" & Trim(txtcodven.Text) & "' and vem_codcia='" & LK_CODCIA & "' "
    Set llave_busca3 = CN.OpenResultset(cade, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    If llave_busca3.EOF Then
        MsgBox "EL VENDEDOR NO EXISTE ", 48, Pub_Titulo
        txtcodart.SetFocus
        Exit Sub
    Else
        wvend = Val(llave_busca3!VEM_ALTERNO)
        For C = 0 To ltvend.ListCount - 1
            DoEvents
            vend = Trim(Right(ltvend.List(C), 10))
            If vend = wvend Then
                MsgBox "EL ARTICULO YA EXISTE EN LA LISTA"
                Exit Sub
            End If
        Next
        'ltarti.Visible = True
        flista2.Visible = True
        ltvend.AddItem (Trim(llave_busca3!VEM_ALTERNO) & " - " & Trim(llave_busca3!VEM_NOMBRE)) & Space(60) & Trim(llave_busca3!VEM_CODVEN)
        txtcodven.Text = ""
'        Me.cmdprocesar.SetFocus
    End If
End If
End Sub

Public Sub sumaval()
Dim wsum As Currency
Dim WF As Integer
wsum = 0
For WF = 1 To gridigv.Rows - 1
wsum = wsum + Val(gridigv.TextMatrix(WF, 5))
Next WF
lbltotal.Caption = Format(wsum, "##,##0.00")
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresup 
   Caption         =   "Establecer Presupuestos de Pedidos por Grupo"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Retornar"
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
      Left            =   10800
      Picture         =   "frmPlanS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Generar:"
      ForeColor       =   &H00808000&
      Height          =   1215
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11775
      Begin VB.OptionButton opperiodo 
         Caption         =   "Reemplazar Solo Promedio Sugerido y Costos."
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   22
         Top             =   600
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton opperiodo 
         Caption         =   "Reemplazar Todo el Periodo"
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin MSComctlLib.ProgressBar pbp 
         Height          =   225
         Left            =   7440
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.TextBox txtperiodop 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Text            =   "000000"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdgenerarpresup 
         Caption         =   "Procesar Periodo."
         Height          =   375
         Left            =   7560
         TabIndex        =   17
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lley 
         Caption         =   "-Captura Valores Actuales de Maximos."
         ForeColor       =   &H00808000&
         Height          =   435
         Index           =   1
         Left            =   2520
         TabIndex        =   29
         Top             =   360
         Width           =   2160
      End
      Begin VB.Label lley 
         AutoSize        =   -1  'True
         Caption         =   "-Genera Todos los Grupos"
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Periodo Presupuestal :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Tag             =   "9999"
         Top             =   360
         Width           =   2445
      End
   End
   Begin VB.Frame fraListar 
      Caption         =   "Opciones:"
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
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   12075
      Begin VB.CommandButton cmdcopy 
         Caption         =   "Copiar al Siguiente Mes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   31
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cmbgrupo 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdmostrar 
         Caption         =   "Mostrar Datos"
         Height          =   735
         Left            =   7800
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox fami 
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   720
         Width           =   5175
      End
      Begin VB.ComboBox cmbper 
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox cheprecios 
         Caption         =   "Lista de Precios Opciones 1 "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10560
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chepedmax 
         Caption         =   "Mostrar Actuzalicion de Canitdad Max. e Pedidos"
         Height          =   375
         Left            =   10440
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ListBox lisarti 
         Height          =   510
         Left            =   2160
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   2910
         Visible         =   0   'False
         Width           =   4695
      End
      Begin MSComctlLib.ProgressBar probar 
         Height          =   225
         Left            =   2400
         TabIndex        =   1
         Top             =   4440
         Visible         =   0   'False
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin RichTextLib.RichTextBox TEXTOVARPRO 
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   5190
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"frmPlanS.frx":014A
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
         Height          =   5445
         Left            =   120
         TabIndex        =   4
         Tag             =   "9999"
         Top             =   1320
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   9604
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         FixedCols       =   0
         BackColorBkg    =   8421376
         FocusRect       =   2
         HighLight       =   2
         AllowUserResizing=   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblfecha 
         AutoSize        =   -1  'True
         Caption         =   "."
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label lfami 
         Caption         =   "Grupo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   27
         Tag             =   "9999"
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Periodo Presupuestal :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Tag             =   "9999"
         Top             =   240
         Width           =   2805
      End
      Begin VB.Label Label1 
         Caption         =   "Valor Mercaderia Maxima :"
         Height          =   375
         Index           =   1
         Left            =   8160
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblvalMax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   9480
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label lblvalMin 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   10440
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Valor Mercaderia Minima :"
         Height          =   375
         Index           =   0
         Left            =   10440
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label larti 
         Caption         =   "Productos :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Tag             =   "9999"
         Top             =   1080
         Width           =   1605
      End
      Begin VB.Label lfami 
         Caption         =   "Linea :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Tag             =   "9999"
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label Label8 
         Caption         =   "Un Momento... procesando información."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   3240
         TabIndex        =   5
         Top             =   1800
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmPresup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim loc_fecha1 As Date
Dim loc_fecha2 As Date
Dim wciasg1  As String
Dim wciasg2  As String
Dim temporal As String

Private Sub cheprecios_Click()
If cheprecios.Value = 1 Then
 lisarti.Clear
 Screen.MousePointer = 11
 'LLENA_ARTI -99, 0
Screen.MousePointer = 0
End If
End Sub

Private Sub cmbper_GotFocus()
Dim rsaddreg  As rdoResultset

cmbper.Clear
pub_cadena = "select  distinct ppe_mes , ppe_ano from  presup_ped where ppe_codcia = '00' and ppe_ano = " & Format(LK_FECHA_DIA, "yyyy") & " group by ppe_mes , ppe_ano"
Set rsaddreg = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
Do Until rsaddreg.EOF
    cmbper.AddItem Format(rsaddreg!ppe_mes, "00") & Format(rsaddreg!ppe_ano, "0000")
rsaddreg.MoveNext
Loop
If cmbper.ListCount > 0 Then cmbper.ListIndex = 0
End Sub

Private Sub cmbper_LostFocus()
Dim wano
Dim wmes
Dim wpo_periodo
Dim wfec1  As Date
Dim wdias
Dim wresp
wpo_periodo = cmbper.Text
wano = Right(wpo_periodo, 4) - 1
wmes = Left(wpo_periodo, 2)
loc_fecha1 = "01/" & wmes & "/" & wano
If wmes = 1 Then wdias = 31
If wmes = 2 Then
   wresp = Val(Val(wano / 4)) - Val(Int(Val(wano / 4)))
   If Val(wresp) = 0 Then
     wdias = 29
   Else
     wdias = 28
   End If
End If
If wmes = 3 Then wdias = 31
If wmes = 4 Then wdias = 30
If wmes = 5 Then wdias = 31
If wmes = 6 Then wdias = 30
If wmes = 7 Then wdias = 31
If wmes = 8 Then wdias = 31
If wmes = 9 Then wdias = 30
If wmes = 10 Then wdias = 31
If wmes = 11 Then wdias = 30
If wmes = 12 Then wdias = 31
wfec1 = wdias & "/" & wmes & "/" & wano
loc_fecha2 = wfec1

End Sub

Private Sub cmdcancel_Click()
Unload frmPresup
End Sub

Private Sub cmdcopy_Click()
Dim wmes As Integer
Dim wmesreep As Integer
Dim wano  As Integer

pub_mensaje = "Reemplazar la infomación del Siguiente MES...   ¿Desea Continuar... ?"
 Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
 If Pub_Respuesta = vbNo Then
    Exit Sub
 End If
 
wmes = Left(cmbper.Text, 2)
wmesreep = Val(Left(cmbper.Text, 2)) + 1
If wmesreep > 12 Then
  MsgBox "No Procede dentro del periodo", 48, Pub_Titulo
  Exit Sub
End If
wano = Right(cmbper.Text, 4)
'pub_cadena = "select  * from presup_ped where ppe_codcia = '00' And ppe_mes = " & wmes & " And ppe_ano = " & wano
'Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
pub_cadena = "Delete  from presup_ped where ppe_codcia = '00' And ppe_mes = " & wmesreep & " And ppe_ano = " & wano
CN.Execute pub_cadena

pub_cadena = "select * into presup_pedTEMPO from presup_ped where ppe_codcia = '00' And ppe_mes = " & wmes & " And ppe_ano = " & wano
CN.Execute pub_cadena
pub_cadena = "update presup_pedTEMPO set ppe_mes = " & wmesreep
CN.Execute pub_cadena
pub_cadena = "insert into presup_ped select * from presup_pedTEMPO"
CN.Execute pub_cadena
pub_cadena = "drop table presup_pedTEMPO"
CN.Execute pub_cadena
MsgBox "Registros Copiados", 48, Pub_Titulo
End Sub

Private Sub cmdgenerarpresup_Click()
Dim rsaddreg  As rdoResultset
Dim rscos As rdoResultset
Dim pbpMin As Integer, pbpMax As Integer, pbpValue As Integer

pub_cadena = "select  * from presup_ped where  ppe_codcia = '99'"
Set rsaddreg = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)

pub_cadena = "select  art_key,sum(art_stock_max)  as vmax from arti where art_codcia in (" & wciasg1 & ") and art_situacion <> '1' group by art_key"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not X.EOF Then
pbpMax = X.RowCount
pbpMin = 0
pbpValue = 0
pbp.Visible = True
End If
Do Until X.EOF
pbpValue = pbpValue + 1
DoEvents

pub_cadena = "select  arm_cospro , arm_costo_ult  from articulo where arm_codcia = '01' and arm_codart = " & X!art_key
Set rscos = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If opperiodo(0).Value Then
rsaddreg.AddNew
rsaddreg!PPE_CODCIA = "00"
rsaddreg!ppe_grupo = 1
rsaddreg!ppe_mes = Left(txtperiodop.Text, 2)
rsaddreg!ppe_ano = Right(txtperiodop.Text, 4)
rsaddreg!ppe_codart = X!art_key
rsaddreg!PPE_MAX_PRESUP = 0
rsaddreg!PPE_MAX_SUGERIDO = X!vmax
If rscos.EOF Then
rsaddreg!PPE_COSPRO = 0
rsaddreg!PPE_COSTO_ULT = 0
Else
rsaddreg!PPE_COSPRO = rscos!ARM_COSPRO
rsaddreg!PPE_COSTO_ULT = rscos!ARM_COSTO_ULT
End If
Else
    pub_cadena = "select  * from presup_ped where  ppe_codcia = '00' and ppe_grupo = 1  and ppe_codart = " & X!art_key & ""
    Set rsaddreg = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    rsaddreg.Edit
    If rscos.EOF Then
    rsaddreg!PPE_COSPRO = 0
    rsaddreg!PPE_COSTO_ULT = 0
    Else
    rsaddreg!PPE_COSPRO = rscos!ARM_COSPRO
    rsaddreg!PPE_COSTO_ULT = rscos!ARM_COSTO_ULT
    End If
    rsaddreg!PPE_MAX_SUGERIDO = X!vmax
End If
rsaddreg.Update
X.MoveNext
Loop

pub_cadena = "select  art_key,sum(art_stock_max)  as vmax from arti where art_codcia in (" & wciasg2 & ") and art_situacion <> '1' group by art_key"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not X.EOF Then
pbpMax = X.RowCount
pbpMin = 0
pbpValue = 0
pbp.Visible = True
End If
Do Until X.EOF
pbpValue = pbpValue + 1
DoEvents
pub_cadena = "select  arm_cospro , arm_costo_ult  from articulo where arm_codcia = '01' and arm_codart = " & X!art_key
Set rscos = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If opperiodo(0).Value Then
   
    rsaddreg.AddNew
    rsaddreg!PPE_CODCIA = "00"
    rsaddreg!ppe_grupo = 2
    rsaddreg!ppe_mes = Left(txtperiodop.Text, 2)
    rsaddreg!ppe_ano = Right(txtperiodop.Text, 4)
    rsaddreg!ppe_codart = X!art_key
    rsaddreg!PPE_MAX_PRESUP = 0
    rsaddreg!PPE_MAX_SUGERIDO = X!vmax
    If rscos.EOF Then
    rsaddreg!PPE_COSPRO = 0
    rsaddreg!PPE_COSTO_ULT = 0
    Else
    rsaddreg!PPE_COSPRO = rscos!ARM_COSPRO
    rsaddreg!PPE_COSTO_ULT = rscos!ARM_COSTO_ULT
    End If
Else
    pub_cadena = "select  * from presup_ped where  ppe_codcia = '00' and ppe_grupo = 2  and ppe_codart = " & X!art_key & ""
    Set rsaddreg = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    rsaddreg.Edit
    If rscos.EOF Then
    rsaddreg!PPE_COSPRO = 0
    rsaddreg!PPE_COSTO_ULT = 0
    Else
    rsaddreg!PPE_COSPRO = rscos!ARM_COSPRO
    rsaddreg!PPE_COSTO_ULT = rscos!ARM_COSTO_ULT
    End If
    rsaddreg!PPE_MAX_SUGERIDO = X!vmax
End If

rsaddreg.Update
X.MoveNext
Loop

pbp.Visible = False
MsgBox "Proceso Terminado ", 48, Pub_Titulo

End Sub




Private Sub cmdmostrar_Click()
Dim wpos As Integer
Dim WFAMI2 As Integer
 lisarti.Clear
 LLENA_ARTI Val(Right(fami.Text, 6)), 0, Trim(cmbper.Text), Val(Left(cmbgrupo.Text, 1))

End Sub

Private Sub Form_Load()
CenterMe Me
PUB_CODCIA = LK_CODCIA
LLENADOS fami, 122
wciasg1 = "'01','10','03','07','05'"
wciasg2 = "'03','07'"
cmbgrupo.Clear
cmbgrupo.AddItem "1 - Grupo Lives " & wciasg1
'cmbgrupo.AddItem "2 - Boticas " & wciasg2
txtperiodop.Text = Format(LK_FECHA_DIA, "mm") & Format(LK_FECHA_DIA, "yyyy")
End Sub
Public Sub LLENADOS(cont As ComboBox, tip As Integer)
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

Private Sub gridarti_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub

If gridarti.COL <> 7 And gridarti.COL <> 6 And gridarti.COL <> 9 And gridarti.COL <> 10 Then Exit Sub
'veri_suma
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


If Left(gridarti.TextMatrix(gridarti.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
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
   gridarti.SetFocus
   End If
End If
End If
Exit Sub



End Sub

Private Sub TEXTOVARPRO_Change()
gridarti.Text = Format(TEXTOVARPRO.Text, "0.00")
End Sub

Private Sub TEXTOVARPRO_GotFocus()
 temporal = gridarti.TextMatrix(gridarti.Row, gridarti.COL)
End Sub

Private Sub TEXTOVARPRO_KeyPress(KeyAscii As Integer)
Dim WPERIODO As String
If KeyAscii = 27 Then
  TEXTOVARPRO.Text = temporal
  TEXTOVARPRO.Visible = False
  gridarti.SetFocus
  Exit Sub
End If
'If gridarti.COL = 1 Then Consistencias gridarti, TEXTOVARPRO, KeyAscii
If gridarti.COL = 9 Or gridarti.COL = 10 Then Consistencias gridarti, TEXTOVARPRO, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If
If gridarti.COL = 1 Or gridarti.COL = 4 Then
  If Val(TEXTOVARPRO.Text) > 99 Then
    Azul3 TEXTOVARPRO, TEXTOVARPRO
    Exit Sub
  End If
End If
WPERIODO = Trim(cmbper.Text)

' COLOCAR EDICION
'------------------
pub_cadena = "select  * from presup_ped where ppe_codcia = '00' and ppe_codart = " & gridarti.TextMatrix(gridarti.Row, 14) & " and ppe_mes = " & Trim(Left(WPERIODO, 2)) & " and ppe_ano =  " & Trim(Right(WPERIODO, 4)) & " and ppe_grupo = " & Val(Left(cmbgrupo.Text, 1))
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not X.EOF Then
 X.Edit
 If gridarti.COL = 6 Then
 X!PPE_MAX_PRESUP = Val(TEXTOVARPRO.Text) * Val(gridarti.TextMatrix(gridarti.Row, 15))
 ElseIf gridarti.COL = 7 Then
 X!PPE_COSTO_SUGERIDO = Format(Val(TEXTOVARPRO.Text) / Val(gridarti.TextMatrix(gridarti.Row, 15)), "0.0000")
 End If
 X.Update
End If

gridarti.TextMatrix(gridarti.Row, 8) = Format(Val(gridarti.TextMatrix(gridarti.Row, 6)) - Val(gridarti.TextMatrix(gridarti.Row, 5)), "0.00")
gridarti.TextMatrix(gridarti.Row, 12) = Format(Val(gridarti.TextMatrix(gridarti.Row, 6)) * Val(gridarti.TextMatrix(gridarti.Row, 4)), "0.00")
mstmaxp:
sum_costo_linea
pasa_alto:
If gridarti.COL = 9 Then
   gridarti.COL = 10
   gridarti.SetFocus
    TEXTOVARPRO.Visible = False
   Exit Sub
Else
  gridarti.COL = 9
End If
If gridarti.Row >= gridarti.Rows - 1 Then
Else
  gridarti.Row = gridarti.Row + 1
End If
pasa_alto2:
gridarti.SetFocus
TEXTOVARPRO.Visible = False


fin:

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

Public Sub LLENA_ARTI(wfami As Integer, WSUBFAMI As Integer, WPERIODO As String, wgrupo As Integer)
Dim cadecia As String
Dim RSSTOCKOT As rdoResultset
If probar.Visible = True Then Exit Sub
Dim xpinta As Integer
Dim xfil  As Integer
Dim WARTI As rdoQuery
Dim wllave_arti As rdoResultset

Dim WBOT As rdoQuery
Dim wllave_bot As rdoResultset

Dim wvalor As Integer
Dim CADE_ROT As String

If Left(cmbgrupo.Text, 1) = 1 Then
    cadecia = wciasg1
Else
    cadecia = wciasg2
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

pub_cadena = "select  ppe_codart,art_nombre , pre_unidad ,pre_equiv from presup_ped,arti,precios  where (ppe_codcia = '00' and art_codcia = '01') and " & _
   " (art_codcia = pre_codcia) and ART_SITUACION <> '1' AND (art_key = pre_codart) and (pre_flag_unidad = 'A') and (ppe_codart = art_key) and art_familia = " & Val(Right(fami.Text, 10)) & " and ppe_mes = " & Trim(Left(WPERIODO, 2)) & " and ppe_ano =  " & Trim(Right(WPERIODO, 4)) & " and ppe_grupo = " & wgrupo
Set WARTI = CN.CreateQuery("", pub_cadena)
Set wllave_arti = WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)
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
'If cheprecios.Value = 0 Then
 gridarti.Cols = 17
'Else
' gridarti.Cols = 13
'End If
gridarti.TextMatrix(0, 0) = "Item"
gridarti.TextMatrix(0, 1) = "Descripcion del Producto"
gridarti.TextMatrix(0, 2) = "Unid."
gridarti.TextMatrix(0, 2) = "Activa"
gridarti.TextMatrix(0, 3) = "Costo"
gridarti.TextMatrix(1, 3) = "Repos."
gridarti.TextMatrix(0, 4) = "Costo"
gridarti.TextMatrix(1, 4) = "Ptromd"

gridarti.TextMatrix(0, 5) = "Pedido"
gridarti.TextMatrix(1, 5) = "Sugerido"
gridarti.TextMatrix(0, 6) = "Pedido "
gridarti.TextMatrix(1, 6) = "Digitado"
gridarti.TextMatrix(0, 7) = "Costo "
gridarti.TextMatrix(1, 7) = "Digitado."
gridarti.TextMatrix(0, 8) = "Pedido "
gridarti.TextMatrix(1, 8) = "Difer."

gridarti.TextMatrix(0, 11) = "Valor"
gridarti.TextMatrix(1, 11) = "Sugerido"
gridarti.TextMatrix(0, 12) = "Valor"
gridarti.TextMatrix(1, 12) = "Digitado"
gridarti.TextMatrix(0, 16) = "Periodo: " & Format(loc_fecha2, "yyyy")
gridarti.TextMatrix(1, 16) = "Vta.Ant"
gridarti.ColWidth(0) = 300
gridarti.ColWidth(1) = 2900
gridarti.ColWidth(2) = 500
gridarti.ColWidth(3) = 700
gridarti.ColWidth(4) = 700
gridarti.ColWidth(5) = 900
gridarti.ColWidth(6) = 900
gridarti.ColWidth(7) = 900
gridarti.ColWidth(8) = 900
gridarti.ColWidth(9) = 0
gridarti.ColWidth(10) = 0
gridarti.ColWidth(11) = 900 ' valor min .
gridarti.ColWidth(12) = 900 ' valor Max
gridarti.ColWidth(13) = 0
gridarti.ColWidth(14) = 0
gridarti.ColWidth(15) = 0
gridarti.ColWidth(16) = 1000
PASACABE:
gridarti.Rows = 2


xfil = 1
Do Until wllave_arti.EOF
  probar.Value = probar.Value + 1
  DoEvents
  xfil = xfil + 1
  gridarti.Rows = gridarti.Rows + 1
  gridarti.RowHeight(gridarti.Rows - 1) = 285

  gridarti.TextMatrix(xfil, 0) = Format(xfil - 1, "00")
  gridarti.TextMatrix(xfil, 1) = Trim(wllave_arti!art_nombre)
  gridarti.TextMatrix(xfil, 2) = Trim(wllave_arti!PRE_UNIDAD)
  
  pub_cadena = "select  * from presup_ped where ppe_codcia = '00' and ppe_codart = " & wllave_arti!ppe_codart & " and ppe_mes = " & Trim(Left(WPERIODO, 2)) & " and ppe_ano =  " & Trim(Right(WPERIODO, 4)) & " and ppe_grupo = " & wgrupo
  Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
  gridarti.TextMatrix(xfil, 3) = Format(X!PPE_COSTO_ULT, "0.00")
  gridarti.TextMatrix(xfil, 4) = Format(X!PPE_COSPRO * wllave_arti!PRE_EQUIV, "0.00")
  gridarti.TextMatrix(xfil, 5) = Format((Val(X!PPE_MAX_SUGERIDO)) / Val(wllave_arti!PRE_EQUIV), "0.00")
  gridarti.TextMatrix(xfil, 6) = Format((Val(X!PPE_MAX_PRESUP)) / Val(wllave_arti!PRE_EQUIV), "0.00")
  gridarti.TextMatrix(xfil, 7) = Format(X!PPE_COSTO_SUGERIDO * wllave_arti!PRE_EQUIV, "0.00")
  gridarti.TextMatrix(xfil, 8) = Format(Val(gridarti.TextMatrix(xfil, 6)) - Val(gridarti.TextMatrix(xfil, 5)), "0.00")
  
  'If Val(gridarti.TextMatrix(xfil, 7)) < 0 Then
  '    gridarti.Row = xfil
  '    gridarti.COL = 7
  '    gridarti.CellForeColor = vbRed
  'End If
  'gridarti.TextMatrix(xfil, 8) = Format((wllave_bot!arm_stock) / wllave_arti!pre_equiv, "0.00")
  'If Val(wllave_arti!art_stockreal_min) <> 0 Then
  '  gridarti.TextMatrix(xfil, 9) = Format(Val(wllave_arti!art_stockreal_min) / Val(wllave_arti!pre_equiv), "0.00")
    gridarti.TextMatrix(xfil, 11) = Format(Val(gridarti.TextMatrix(xfil, 5)) * Val(gridarti.TextMatrix(xfil, 4)), "0.00")
  'End If
  'If Val(wllave_arti!art_stockreal_max) <> 0 Then
  '  gridarti.TextMatrix(xfil, 10) = Format(Val(wllave_arti!art_stockreal_max) / Val(wllave_arti!pre_equiv), "0.00")
    gridarti.TextMatrix(xfil, 12) = Format(Val(gridarti.TextMatrix(xfil, 6)) * Val(gridarti.TextMatrix(xfil, 4)), "0.00")
  'End If
   
mstmaxp:
    'gridarti.TextMatrix(xfil, 13) = wllave_arti!art_alterno
    gridarti.TextMatrix(xfil, 14) = wllave_arti!ppe_codart
    gridarti.TextMatrix(xfil, 15) = wllave_arti!PRE_EQUIV
    pub_cadena = "SELECT SUM(FAR_CANTIDAD) AS venta " & _
    " FROM FACART,CLIENTES WHERE (FAR_CODCIA = CLI_CODCIA) AND (FAR_CP = CLI_CP) AND " & _
    " (FAR_CODCLIE = CLI_CODCLIE) AND (CLI_LETRA <> '1') AND " & _
    " (FAR_CP = 'C') AND  FAR_CODCIA in (" & cadecia & ") AND  FAR_CP = 'C' AND " & _
    "FAR_FECHA_COMPRA >= '" & Format(loc_fecha1, "dd/mm/yyyy") & "' AND FAR_FECHA_COMPRA <= '" & Format(loc_fecha2, "dd/mm/yyyy") & "' " & _
    "AND FAR_ESTADO <> 'E' AND FAR_ESTADO2 <> 'L' AND FAR_TIPMOV = 10 AND FAR_CODART = " & wllave_arti!ppe_codart
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    If Not X.EOF Then
        gridarti.TextMatrix(xfil, 16) = Format((Val(Nulo_Valor0(X!VENTA))) / Val(wllave_arti!PRE_EQUIV), "0.00")
    Else
        gridarti.TextMatrix(xfil, 16) = 0
    End If
     
sale_otro:
wllave_arti.MoveNext
Loop
sum_costo_linea
gridarti.Visible = True
gridarti.SetFocus
probar.Visible = False
DoEvents
End Sub


Public Sub sum_costo_linea()
Dim xpinta  As Integer
Dim sum_smin As Currency
Dim sum_smax As Currency
sum_smin = 0
sum_smax = 0
For xpinta = 0 To gridarti.Rows - 1
   sum_smin = sum_smin + Val(gridarti.TextMatrix(xpinta, 11))
   sum_smax = sum_smax + Val(gridarti.TextMatrix(xpinta, 12))
Next xpinta
lblvalMin.Caption = Format(sum_smin, "#,###,##0.00")
lblvalMax.Caption = Format(sum_smax, "#,###,##0.00")
End Sub

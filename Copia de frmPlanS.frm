VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlanS 
   Caption         =   "Plantilla de Stock Minimos y Maximos"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraListar 
      Caption         =   "Filtrar :"
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
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11475
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
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chepedmax 
         Caption         =   "Mostrar Actuzalicion de Canitdad Max. e Pedidos"
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ListBox fami 
         Height          =   1230
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   4095
      End
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
         Height          =   510
         Left            =   9960
         Picture         =   "frmPlanS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5640
         Width           =   1035
      End
      Begin VB.ListBox lisarti 
         Height          =   510
         Left            =   1440
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   2910
         Visible         =   0   'False
         Width           =   6135
      End
      Begin MSComctlLib.ProgressBar probar 
         Height          =   225
         Left            =   750
         TabIndex        =   1
         Top             =   4410
         Visible         =   0   'False
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin RichTextLib.RichTextBox TEXTOVARPRO 
         Height          =   375
         Left            =   1050
         TabIndex        =   2
         Top             =   5190
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
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
         Height          =   3525
         Left            =   120
         TabIndex        =   5
         Tag             =   "9999"
         Top             =   2040
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   6218
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Valor Mercaderia Maxima :"
         Height          =   375
         Index           =   1
         Left            =   8160
         TabIndex        =   13
         Top             =   1440
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
         TabIndex        =   12
         Top             =   1560
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
         Left            =   9480
         TabIndex        =   11
         Top             =   960
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Valor Mercaderia Minima :"
         Height          =   375
         Index           =   0
         Left            =   8160
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label larti 
         Caption         =   "Articulos"
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
         Left            =   240
         TabIndex        =   8
         Tag             =   "9999"
         Top             =   1800
         Width           =   1125
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
         Left            =   240
         TabIndex        =   7
         Tag             =   "9999"
         Top             =   240
         Width           =   1125
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
         Height          =   405
         Left            =   1650
         TabIndex        =   6
         Top             =   3960
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmPlanS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temporal As String

Private Sub cheprecios_Click()
If cheprecios.Value = 1 Then
 lisarti.Clear
 Screen.MousePointer = 11
 LLENA_ARTI -99, 0
Screen.MousePointer = 0
End If
End Sub

Private Sub cmdcancel_Click()
Unload frmPlanS
End Sub

Private Sub fami_Click()
Dim wpos As Integer
Dim WFAMI2 As Integer
 lisarti.Clear
 LLENA_ARTI Val(Right(fami.Text, 6)), 0

End Sub


Private Sub Form_Load()
PUB_CODCIA = LK_CODCIA
LLENADOS fami, 122
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

Private Sub gridarti_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub

If gridarti.COL <> 6 And gridarti.COL <> 9 And gridarti.COL <> 10 Then Exit Sub
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



' COLOCAR EDICION
'------------------
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_KEY = gridarti.TextMatrix(gridarti.Row, 14)
LEER_ART_LLAVE
If chepedmax.Value = 1 Then
    pub_cadena = "update articulo set arm_stock2 = " & Val(gridarti.TextMatrix(gridarti.Row, 9)) * Val(gridarti.TextMatrix(gridarti.Row, 15)) & " where arm_codart in (" & PUB_KEY & ") and arm_codcia in('03','05','07','09') "
    CN.Execute pub_cadena
    GoTo mstmaxp
End If

If cheprecios.Value = 1 Then
   pub_cadena = "SELECT PRE_OP1 FROM PRECIOS WHERE PRE_CODCIA   = '" & LK_CODCIA & "' AND PRE_CODART = " & PUB_KEY & " AND PRE_EQUIV = " & Val(gridarti.TextMatrix(gridarti.Row, 15))
   Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   If Not X.EOF Then
   X.Edit
   X!PRE_OP1 = Val(gridarti.TextMatrix(gridarti.Row, 6))
   X.Update
   End If
    If gridarti.Row >= gridarti.Rows - 1 Then
    Else
    gridarti.Row = gridarti.Row + 1
    End If

   GoTo pasa_alto2
End If

If Not art_LLAVE.EOF Then
 art_LLAVE.Edit
 art_LLAVE!art_stockreal_min = Val(gridarti.TextMatrix(gridarti.Row, 9)) * Val(gridarti.TextMatrix(gridarti.Row, 15))
 art_LLAVE!art_stockreal_max = Val(gridarti.TextMatrix(gridarti.Row, 10)) * Val(gridarti.TextMatrix(gridarti.Row, 15))
 art_LLAVE.Update
End If
If Val(art_LLAVE!art_stockreal_min) <> 0 Then
   gridarti.TextMatrix(gridarti.Row, 11) = Format((Val(art_LLAVE!art_stockreal_min) / Val(gridarti.TextMatrix(gridarti.Row, 15)) * Val(gridarti.TextMatrix(gridarti.Row, 4))), "0.00")
Else
  gridarti.TextMatrix(gridarti.Row, 11) = ""
End If
If Val(art_LLAVE!art_stockreal_max) <> 0 Then
  gridarti.TextMatrix(gridarti.Row, 12) = Format((Val(art_LLAVE!art_stockreal_max) / Val(gridarti.TextMatrix(gridarti.Row, 15))) * Val(gridarti.TextMatrix(gridarti.Row, 4)), "0.00")
Else
  gridarti.TextMatrix(gridarti.Row, 12) = ""
End If
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

Public Sub LLENA_ARTI(wfami As Integer, WSUBFAMI As Integer)
Dim RSSTOCKOT As rdoResultset
If ProBar.Visible = True Then Exit Sub
Dim xpinta As Integer
Dim xfil  As Integer
Dim WARTI As rdoQuery
Dim wllave_arti As rdoResultset

Dim WBOT As rdoQuery
Dim wllave_bot As rdoResultset

Dim wvalor As Integer
Dim CADE_ROT As String

  

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
If cheprecios.Value = 1 Then
  pub_cadena = "SELECT ARM_COSPRO, ART_STOCK_MAX , PRE_COSTO_ANT , ART_KEY , art_alterno, art_nombre, PRE_EQUIV , PRE_UNIDAD,pre_op1,pre_pre2, art_stockreal_min, art_stockreal_max, ARM_STOCK  , ((ARM_STOCK / PRE_EQUIV) - (ART_STOCK_MAX / PRE_EQUIV)) AS DIF , ART_PLANCHA FROM ARTI, articulo, precios WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART ) AND PRE_FLAG_UNIDAD = 'A' AND ART_CODCIA = ? AND ART_ORDEN = 1 AND ART_CALIDAD = 1 and ART_FLAG_STOCK = 'M' ORDER BY art_nombre " 'DIF ASC   "
  wvalor = 0
  
Else
If wfami = 0 Then Exit Sub
'If wfami <> 0 Then
  pub_cadena = "SELECT ARM_COSPRO, ART_STOCK_MAX , PRE_COSTO_ANT , ART_KEY , art_alterno, art_nombre, PRE_EQUIV , PRE_UNIDAD, art_stockreal_min, art_stockreal_max, ARM_STOCK  , ((ARM_STOCK / PRE_EQUIV) - (ART_STOCK_MAX / PRE_EQUIV)) AS DIF , ART_PLANCHA FROM ARTI, articulo, precios WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART ) AND PRE_FLAG_UNIDAD = 'A' AND ART_CODCIA = ? AND ART_FAMILIA = ? AND ART_CALIDAD = 1 and ART_FLAG_STOCK = 'M' " & CADE_ROT & " ORDER BY art_nombre " 'DIF ASC   "
  wvalor = wfami
End If
'ElseIf WSUBFAMI <> 0 Then
'  pub_cadena = "SELECT ARM_COSPRO, ART_STOCK_MAX ,   PRE_COSTO_ANT, ART_KEY , art_alterno, art_nombre, PRE_EQUIV , PRE_UNIDAD, ARM_STOCK , ((ARM_STOCK/PRE_EQUIV) - (ART_STOCK_MAX/PRE_EQUIV)) AS DIF , ART_PLANCHA FROM ARTI, articulo, precios WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART ) AND PRE_FLAG_UNIDAD = 'A' AND ART_CODCIA = ? AND ART_SUBFAM = ? AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M'  " & CADE_ROT & " ORDER BY DIF ASC "
'  wvalor = WSUBFAMI
'End If
Set WARTI = CN.CreateQuery("", pub_cadena)
WARTI(0) = " "
If wvalor <> 0 Then WARTI(1) = 0
Set wllave_arti = WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)
WARTI(0) = LK_CODCIA
If wvalor <> 0 Then WARTI(1) = wvalor
wllave_arti.Requery
lisarti.Clear
If wllave_arti.EOF Then
  lisarti.Clear
  Exit Sub
End If
ProBar.Value = 0
ProBar.Min = 0
ProBar.max = wllave_arti.RowCount
ProBar.Visible = True
gridarti.Visible = False
DoEvents
gridarti.Clear
'If cheprecios.Value = 0 Then
 gridarti.Cols = 16
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
gridarti.TextMatrix(0, 5) = "Stock "
gridarti.TextMatrix(1, 5) = "Actual "

gridarti.TextMatrix(0, 6) = "Stock "
gridarti.TextMatrix(1, 6) = "Maximo"
If cheprecios.Value = 1 Then
gridarti.TextMatrix(0, 6) = "Precio"
gridarti.TextMatrix(1, 6) = "Opcional"
gridarti.ColWidth(1) = 2800
gridarti.TextMatrix(0, 8) = "Lista "
gridarti.TextMatrix(1, 8) = "Pre 2."

gridarti.ColWidth(7) = 0
gridarti.ColWidth(8) = 1000
gridarti.ColWidth(9) = 0
gridarti.ColWidth(10) = 0

gridarti.ColWidth(11) = 0 ' valor min .
gridarti.ColWidth(12) = 0 ' valor Max

gridarti.ColWidth(13) = 0
gridarti.ColWidth(14) = 0
gridarti.ColWidth(15) = 0
GoTo PASACABE:
End If
gridarti.TextMatrix(0, 7) = "Stock "
gridarti.TextMatrix(1, 7) = "Difer."

gridarti.TextMatrix(0, 8) = "Stock "
gridarti.TextMatrix(1, 8) = "Botica"
'If cheprecios.Value = 0 Then
  gridarti.TextMatrix(0, 9) = "Stock"
  gridarti.TextMatrix(1, 9) = "Minimo"
  gridarti.TextMatrix(0, 10) = "Stock"
  gridarti.TextMatrix(1, 10) = "Maximo"

  gridarti.TextMatrix(0, 11) = "Valor"
  gridarti.TextMatrix(1, 11) = "Minimo"
  gridarti.TextMatrix(0, 12) = "Valor"
  gridarti.TextMatrix(1, 12) = "Maximo"

'Else
'  gridarti.TextMatrix(0, 9) = "Precio a "
'  gridarti.TextMatrix(1, 9) = "Cotizar"
'  gridarti.TextMatrix(1, 12) = "Oferta"
'  gridarti.ColWidth(12) = 2000
'End If

gridarti.ColWidth(0) = 300
gridarti.ColWidth(1) = 3200
gridarti.ColWidth(2) = 700
gridarti.ColWidth(3) = 0
gridarti.ColWidth(4) = 700
gridarti.ColWidth(5) = 700
gridarti.ColWidth(6) = 700
gridarti.ColWidth(7) = 700
gridarti.ColWidth(8) = 0
gridarti.ColWidth(9) = 900
gridarti.ColWidth(10) = 900

gridarti.ColWidth(11) = 950 ' valor min .
gridarti.ColWidth(12) = 950 ' valor Max

gridarti.ColWidth(13) = 0
gridarti.ColWidth(14) = 0
gridarti.ColWidth(15) = 0
If chepedmax.Value = 1 Then
    gridarti.TextMatrix(0, 9) = "Max."
    gridarti.TextMatrix(1, 9) = "Pedido"
 
    gridarti.ColWidth(9) = 1200
    gridarti.ColWidth(10) = 0
    gridarti.ColWidth(11) = 0
    gridarti.ColWidth(12) = 0
End If
PASACABE:
gridarti.Rows = 2



xfil = 1
Do Until wllave_arti.EOF
  ProBar.Value = ProBar.Value + 1
  DoEvents
  xfil = xfil + 1
  gridarti.Rows = gridarti.Rows + 1
  gridarti.RowHeight(gridarti.Rows - 1) = 285
  lisarti.AddItem wllave_arti!art_alterno & " " & wllave_arti!ART_NOMBRE & String(120, " ") & wllave_arti!ART_KEY
  If LK_CODCIA = "01" Then
    WBOT(0) = "04"
  Else
   WBOT(0) = LK_CODCIA
  End If
  WBOT(1) = wllave_arti!ART_KEY
  wllave_bot.Requery
  If wllave_bot.EOF Then
     MsgBox "Avisar el Codigo no Existe (SUBALMACEN): " & wllave_arti!ART_KEY, 48, Pub_Titulo
     GoTo sale_otro
  End If
'  If Val(wllave_arti!PRE_COSTO_ANT) = 1610.7 Then Stop
  gridarti.TextMatrix(xfil, 0) = Format(xfil - 1, "00")
  gridarti.TextMatrix(xfil, 1) = Trim(wllave_arti!ART_NOMBRE)
  gridarti.TextMatrix(xfil, 2) = Trim(wllave_arti!pre_unidad)
  gridarti.TextMatrix(xfil, 3) = Format(wllave_arti!PRE_cosTO_ant, "0.00")
  gridarti.TextMatrix(xfil, 4) = Format(wllave_arti!ARM_COSPRO * wllave_arti!pre_equiv, "0.00")
  
  If cheprecios.Value = 1 Then
    pub_cadena = "SELECT SUM(ARM_STOCK) AS STOCK FROM ARTICULO WHERE ARM_CODCIA IN ('02','04','08','07') AND ARM_CODART = " & wllave_arti!ART_KEY
    Set RSSTOCKOT = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)   ' rdConcurReadOnly) ', rdConcurLock)
    gridarti.TextMatrix(xfil, 6) = Format(Nulo_Valor0(wllave_arti!PRE_OP1), "0.00")
    gridarti.TextMatrix(xfil, 5) = Format((Val(wllave_arti!arm_stock) + Val(RSSTOCKOT!stock)) / Val(wllave_arti!pre_equiv), "0.00")
    gridarti.TextMatrix(xfil, 8) = Trim(wllave_arti!pre_pre2)
  GoTo mstmaxp
  End If
  gridarti.TextMatrix(xfil, 5) = Format((Val(wllave_arti!arm_stock) + Val(wllave_bot!arm_stock)) / Val(wllave_arti!pre_equiv), "0.00")
  gridarti.TextMatrix(xfil, 6) = Format(wllave_arti!ART_STOCK_MAX / wllave_arti!pre_equiv, "0.00")
  
  gridarti.TextMatrix(xfil, 7) = Format(Val(gridarti.TextMatrix(xfil, 5)) - Val(gridarti.TextMatrix(xfil, 6)), "0.00")
  If Val(gridarti.TextMatrix(xfil, 7)) < 0 Then
      gridarti.Row = xfil
      gridarti.COL = 7
      gridarti.CellForeColor = vbRed
  End If
  gridarti.Row = xfil
  gridarti.COL = 6
  gridarti.CellForeColor = vbBlue
  If LK_CODCIA = "01" Then
    WBOT(0) = "03"
  Else
    WBOT(0) = LK_CODCIA
  End If
  WBOT(1) = wllave_arti!ART_KEY
  wllave_bot.Requery
  If wllave_bot.EOF Then
     MsgBox "Avisar el Codigo no Existe (BOTICA): " & wllave_arti!ART_KEY, 48, Pub_Titulo
     GoTo sale_otro
  End If
  If chepedmax.Value = 1 Then
    gridarti.TextMatrix(xfil, 9) = Format((wllave_bot!arm_stock2) / wllave_arti!pre_equiv, "0.00")
  GoTo mstmaxp
  End If
  gridarti.TextMatrix(xfil, 8) = Format((wllave_bot!arm_stock) / wllave_arti!pre_equiv, "0.00")
  If Val(wllave_arti!art_stockreal_min) <> 0 Then
    gridarti.TextMatrix(xfil, 9) = Format(Val(wllave_arti!art_stockreal_min) / Val(wllave_arti!pre_equiv), "0.00")
    gridarti.TextMatrix(xfil, 11) = Format(Val(gridarti.TextMatrix(xfil, 9)) * Val(gridarti.TextMatrix(xfil, 4)), "0.00")
  End If
  If Val(wllave_arti!art_stockreal_max) <> 0 Then
    gridarti.TextMatrix(xfil, 10) = Format(Val(wllave_arti!art_stockreal_max) / Val(wllave_arti!pre_equiv), "0.00")
    gridarti.TextMatrix(xfil, 12) = Format(Val(gridarti.TextMatrix(xfil, 10)) * Val(gridarti.TextMatrix(xfil, 4)), "0.00")
  End If
   
mstmaxp:
  gridarti.TextMatrix(xfil, 13) = wllave_arti!art_alterno
  gridarti.TextMatrix(xfil, 14) = wllave_arti!ART_KEY
  gridarti.TextMatrix(xfil, 15) = wllave_arti!pre_equiv
  
  If wllave_arti!art_plancha = 1 Then
    gridarti.Row = xfil
    For xpinta = 0 To gridarti.Cols - 1
       gridarti.COL = xpinta
       gridarti.CellForeColor = vbRed
    Next xpinta
  End If
sale_otro:
wllave_arti.MoveNext
Loop
sum_costo_linea
gridarti.Visible = True
gridarti.SetFocus
ProBar.Visible = False
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

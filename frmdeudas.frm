VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmdeudas 
   Caption         =   "Deudas Programadas."
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   ForeColor       =   &H00C0FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdias 
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      Text            =   "0"
      Top             =   240
      Width           =   495
   End
   Begin VB.Frame fradetalle 
      BackColor       =   &H00808000&
      Caption         =   "Detalle de Doicumentos por Pagar : "
      Height          =   3885
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   9465
      Begin MSFlexGridLib.MSFlexGrid griddet 
         Height          =   2655
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   4683
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   8421376
         Appearance      =   0
      End
      Begin VB.Label lbltotdia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   285
         Left            =   5130
         TabIndex        =   16
         Top             =   3300
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         Caption         =   "Total del Dia : "
         Height          =   255
         Index           =   3
         Left            =   3990
         TabIndex        =   15
         Top             =   3300
         Width           =   1155
      End
      Begin VB.Label lblfecha 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/2003"
         Height          =   285
         Left            =   1590
         TabIndex        =   13
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         Caption         =   "Fecha de Vcto :"
         Height          =   225
         Left            =   150
         TabIndex        =   12
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdmostrar 
      Caption         =   "Mostrar"
      Height          =   525
      Left            =   6000
      TabIndex        =   8
      Top             =   360
      Width           =   1515
   End
   Begin VB.CommandButton cmdreto 
      Caption         =   "&Retornar"
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4005
      Left            =   120
      TabIndex        =   0
      Top             =   1170
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7064
      _Version        =   393216
      Rows            =   1
      Cols            =   4
   End
   Begin MSMask.MaskEdBox txtCampo1 
      Height          =   285
      Left            =   1110
      TabIndex        =   3
      Top             =   60
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
   Begin MSMask.MaskEdBox txtCampo2 
      Height          =   285
      Left            =   1110
      TabIndex        =   4
      Top             =   420
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
   Begin VB.Label ldif 
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
      Height          =   285
      Left            =   5520
      TabIndex        =   24
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Diferencia :"
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   23
      Top             =   5640
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Exceso:"
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   22
      Top             =   5280
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "Disponible :"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   21
      Top             =   5280
      Width           =   915
   End
   Begin VB.Label lrojo 
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6720
      TabIndex        =   20
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lazul 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4440
      TabIndex        =   19
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Ver por:"
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   18
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblsuma 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Total General Deuda :"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   9
      Top             =   5250
      Width           =   1635
   End
   Begin VB.Label lblmonto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   285
      Left            =   1590
      TabIndex        =   7
      Top             =   810
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Disponible por Día :"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Final :"
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Inicial :"
      Height          =   255
      Left            =   30
      TabIndex        =   1
      Top             =   90
      Width           =   1095
   End
End
Attribute VB_Name = "frmdeudas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmostrar_Click()
fradetalle.Visible = False
txtCampo2_KeyPress 13
End Sub

Private Sub cmdreto_Click()
Unload frmdeudas
End Sub

Private Sub Form_Activate()
txtCampo2_KeyPress 13
End Sub

Private Sub Form_Load()
CenterMe frmdeudas
txtCampo1.Mask = "##/##/####"
txtCampo2.Mask = "##/##/####"
End Sub


Private Sub grid1_DblClick()
Dim acu_suma As Currency
Dim WS_CIA As String
Dim WS_CIAALLOG As String
If (LK_CODCIA = "25" Or LK_CODCIA = "27" Or LK_CODCIA = "28" Or LK_CODCIA = "29") Then
 WS_CIA = "CAR_CODCIA IN ('01','07','05','03','25','27','28','29') AND "
 WS_CIAALLOG = "ALL_CODCIA IN ('01','07','05','03','25','27','28','29') AND "
Else
 WS_CIA = "CAR_CODCIA IN ('" & LK_CODCIA & "') AND "
 WS_CIAALLOG = "ALL_CODCIA IN ('01','07','05','03','25','27','28','29') AND "
End If
    

If Not IsDate(grid1.TextMatrix(grid1.Row, 1)) Then Exit Sub

pub_cadena = "SELECT CAR_CODCIA , CAR_CODCLIE, CAR_TIPDOC, CAR_SERDOC, CAR_NUMDOC , CAR_NUMSER_C, CAR_NUMFAC_C ,CAR_NUMGUIA,  CAR_NUMSER , CAR_NUMFAC , CAR_CONCEPTO , CAR_FECHA_VCTO_ORIG, CAR_FECHA_SUNAT,car_num_cheque,  SUM(CAR_IMPORTE) AS TOTAL FROM CARTERA WHERE " & WS_CIA & "  CAR_CP = ? AND CAR_FECHA_VCTO = ? AND CAR_IMPORTE <> 0 GROUP BY CAR_CODCIA,CAR_CODCLIE ,CAR_TIPDOC , CAR_SERDOC,CAR_NUMDOC,CAR_NUMSER_C, CAR_NUMFAC_C , CAR_NUMGUIA ,CAR_NUMSER , CAR_NUMFAC ,  CAR_CONCEPTO, CAR_FECHA_VCTO_ORIG, CAR_FECHA_SUNAT,car_num_cheque"
Set PSCAR_TOPE = CN.CreateQuery("", pub_cadena)
PSCAR_TOPE(0) = ""
PSCAR_TOPE(1) = Date
Set car_tope = PSCAR_TOPE.OpenResultset(rdOpenKeyset, rdConcurValues)
PSCAR_TOPE(0) = "P"
PSCAR_TOPE(1) = grid1.TextMatrix(grid1.Row, 1)
car_tope.Requery

lblfecha.Caption = grid1.TextMatrix(grid1.Row, 1)
griddet.Rows = 1
griddet.Cols = 8

griddet.TextMatrix(0, 0) = "Fec.Emis"
griddet.TextMatrix(0, 1) = "TD"
griddet.TextMatrix(0, 2) = "Serie"
griddet.TextMatrix(0, 3) = "Numero"
griddet.TextMatrix(0, 4) = "Concepto"
griddet.TextMatrix(0, 5) = "Saldo"
griddet.TextMatrix(0, 6) = "F.V.Orig"
griddet.TextMatrix(0, 7) = "Proveedor"


griddet.ColWidth(0) = 1200
griddet.ColWidth(1) = 400
griddet.ColWidth(2) = 500
griddet.ColWidth(3) = 800
griddet.ColWidth(4) = 2300
griddet.ColWidth(5) = 1200
griddet.ColWidth(6) = 800
griddet.ColWidth(7) = 2500
griddet.Rows = 1
acu_suma = 0
Do Until car_tope.EOF
If Trim(car_tope!CAR_TIPDOC) = "LE" Or Trim(car_tope!CAR_TIPDOC) = "LR" Then
Else
GoTo AVAN
End If
 griddet.Rows = griddet.Rows + 1
 
 griddet.RowHeight(griddet.Rows - 1) = 285
 griddet.TextMatrix(griddet.Rows - 1, 0) = car_tope!car_codcia & " " & Format(car_tope!car_fecha_sunat, "dd/mm/yy")
 griddet.TextMatrix(griddet.Rows - 1, 1) = car_tope!CAR_TIPDOC
 griddet.TextMatrix(griddet.Rows - 1, 2) = car_tope!car_numser_c
 If Left(car_tope!CAR_TIPDOC, 1) = "L" Then
   griddet.TextMatrix(griddet.Rows - 1, 3) = car_tope!car_num_cheque
 ElseIf car_tope!car_NUMFAC_C = 0 Then
   griddet.TextMatrix(griddet.Rows - 1, 3) = car_tope!car_NUMGUIA
 Else
   griddet.TextMatrix(griddet.Rows - 1, 3) = car_tope!car_NUMFAC_C
 End If
 griddet.TextMatrix(griddet.Rows - 1, 4) = car_tope!car_concepto
 
 griddet.TextMatrix(griddet.Rows - 1, 5) = Format(car_tope!total, "##,##0.00")
 acu_suma = acu_suma + Val(car_tope!total)
 griddet.TextMatrix(griddet.Rows - 1, 6) = Format(car_tope!car_fecha_VCTO_orig, "dd/mm/yy")
 
 SQ_OPER = 1
 pu_codclie = car_tope!CAR_codclie
 pu_cp = "P"
 pu_codcia = car_tope!car_codcia
 LEER_CLI_LLAVE
 If cli_llave.EOF Then
   griddet.TextMatrix(griddet.Rows - 1, 7) = "No Encontrado"
 Else
   griddet.TextMatrix(griddet.Rows - 1, 7) = Trim(cli_llave!cli_nombre)
 End If
 
AVAN:
 car_tope.MoveNext
Loop


pub_cadena = "SELECT * FROM ALLOG WHERE " & WS_CIAALLOG & "  ALL_FECHA_VCTO = '" & Format(grid1.TextMatrix(grid1.Row, 1), "dd/mm/yyyy") & "' AND ALL_CODTRA = 2760 AND ALL_FLAG_EXT <> 'E' "
Set ps_lelr = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
Do Until ps_lelr.EOF
'If Trim(ps_lelr!car_TIPDOC) = "LE" Or Trim(ps_lelr!car_TIPDOC) = "LR" Then
'Else
'GoTo AVAN2
'End If
 griddet.Rows = griddet.Rows + 1
 
 griddet.RowHeight(griddet.Rows - 1) = 285
 griddet.TextMatrix(griddet.Rows - 1, 0) = ps_lelr!ALL_CODCIA & " " & Format(ps_lelr!ALL_FECHA_SUNAT, "dd/mm/yy")
 griddet.TextMatrix(griddet.Rows - 1, 1) = ps_lelr!ALL_TIPDOC
 griddet.TextMatrix(griddet.Rows - 1, 2) = ps_lelr!all_numser_c
 griddet.TextMatrix(griddet.Rows - 1, 3) = ps_lelr!all_numfac
 griddet.TextMatrix(griddet.Rows - 1, 4) = "Temporal - " & Trim(ps_lelr!all_concepto)
 griddet.TextMatrix(griddet.Rows - 1, 5) = Format(ps_lelr!ALL_IMPORTE_AMORT, "##,##0.00")
 acu_suma = acu_suma + Val(ps_lelr!ALL_IMPORTE_AMORT)
 griddet.TextMatrix(griddet.Rows - 1, 6) = Format(ps_lelr!ALL_FECHA_VCTO, "dd/mm/yy")
 
 SQ_OPER = 1
 pu_codclie = ps_lelr!ALL_CODCLIE
 pu_cp = "P"
 pu_codcia = ps_lelr!ALL_CODCIA
 LEER_CLI_LLAVE
 If cli_llave.EOF Then
   griddet.TextMatrix(griddet.Rows - 1, 7) = "No Encontrado"
 Else
   griddet.TextMatrix(griddet.Rows - 1, 7) = Trim(cli_llave!cli_nombre)
 End If
'AVAN2:
 ps_lelr.MoveNext
Loop


lbltotdia.Caption = Format(acu_suma, "##,##0.00")
fradetalle.Visible = True
griddet.SetFocus

End Sub

Private Sub grid1_GotFocus()
fradetalle.Visible = False
End Sub

Private Sub griddet_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 fradetalle.Visible = False
 grid1.SetFocus
End If
End Sub

Private Sub txtCampo1_GotFocus()
fradetalle.Visible = False
End Sub

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Azul2 txtCampo2, txtCampo2
End If
End Sub

Private Sub txtCampo2_GotFocus()
fradetalle.Visible = False
End Sub

Private Sub txtCampo2_KeyPress(KeyAscii As Integer)
Dim acu_azul As Currency
Dim acu_rojo As Currency
Dim acu_suma As Currency
Dim wfecha As Date
If KeyAscii <> 13 Then Exit Sub
' ----------------
If Not IsDate(txtCampo1.Text) Then
  MsgBox "Fecha Invalida"
 Exit Sub
End If
If Not IsDate(txtCampo2.Text) Then
  MsgBox "Fecha Invalida"
 Exit Sub
End If

wcant = DateDiff("d", txtCampo1.Text, txtCampo2.Text) + 1
grid1.Rows = 1
grid1.Cols = 5

grid1.TextMatrix(0, 0) = "Ord"
grid1.TextMatrix(0, 1) = "Fecha Vcto"
grid1.TextMatrix(0, 2) = "Dia"
grid1.TextMatrix(0, 3) = "Deuda Prog."
grid1.TextMatrix(0, 4) = "Disponible"

grid1.ColWidth(0) = 600
grid1.ColWidth(1) = 1200
grid1.ColWidth(2) = 600
grid1.ColWidth(3) = 1200
grid1.ColWidth(4) = 1200

wfecha = DateAdd("d", -1, txtCampo1.Text)
'If LK_CODTRA = 1455 Then
'lblmonto.Caption = Format(par_llave!PAR_SALDO_CAJA_D_AYER, "#,##0.00")
'Else
'lblmonto.Caption = Format(par_llave!PAR_TASA_VENTA, "#,##0.00")
par_llave.Requery
lblmonto.Caption = Format(par_llave!PAR_SALDO_CAJA_HOY, "#,##0.00")

'End If
acu_azul = 0
acu_rojo = 0
acu_suma = 0
grid1.Visible = False
DoEvents
For fila = 1 To wcant
  wfecha = DateAdd("d", 1, wfecha)
  grid1.Rows = grid1.Rows + 1
  grid1.RowHeight(grid1.Rows - 1) = 285
  grid1.TextMatrix(grid1.Rows - 1, 0) = Format(fila, "00")
  grid1.TextMatrix(grid1.Rows - 1, 1) = wfecha
  grid1.TextMatrix(grid1.Rows - 1, 2) = UCase(Format(wfecha, "ddd"))
  
  'If LK_CODTRA = 1455 Then
  'ws_deuda = DeudaDia("P", wfecha, 1)
  'Else
  ws_deuda = DeudaDia("P", wfecha)
  'End If
  acu_suma = acu_suma + Val(ws_deuda)
  grid1.TextMatrix(grid1.Rows - 1, 3) = Format(ws_deuda, "#,##0.00")
  grid1.TextMatrix(grid1.Rows - 1, 4) = Format(Val(par_llave!PAR_SALDO_CAJA_HOY) - ws_deuda, "#,##0.00")
  If Left(UCase(Format(wfecha, "ddd")), 1) = "D" Then
   grid1.TextMatrix(grid1.Rows - 1, 4) = 0
   grid1.Row = grid1.Rows - 1
   grid1.COL = 1
   grid1.CellForeColor = vbRed
  End If
  If Val(Format(grid1.TextMatrix(grid1.Rows - 1, 4), "0.00")) < 0 Then
   grid1.Row = grid1.Rows - 1
   grid1.COL = 4
   grid1.CellForeColor = vbRed
   acu_rojo = acu_rojo + Val(Format(grid1.TextMatrix(grid1.Rows - 1, 4), "0.00"))
  Else
   grid1.Row = grid1.Rows - 1
   grid1.COL = 4
   grid1.CellForeColor = vbBlue
   acu_azul = acu_azul + Val(Format(grid1.TextMatrix(grid1.Rows - 1, 4), "0.00"))
  End If
  
  
Next fila
grid1.Visible = True
DoEvents
lblsuma.Caption = Format(acu_suma, "#,##0.00")
lazul.Caption = Format(acu_azul, "#,##0.00")
lrojo.Caption = Format(acu_rojo, "#,##0.00")
ldif.Caption = Format(acu_azul + acu_rojo, "#,##0.00")

grid1.SetFocus

End Sub

Private Sub txtdias_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCampo2.Text = DateAdd("d", Val(txtdias.Text) - 1, CDate(txtCampo1.Text))
    cmdmostrar_Click
End If
End Sub

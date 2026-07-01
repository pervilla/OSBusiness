VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmprogra 
   Caption         =   "Programación"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fradetalle 
      Caption         =   "Detalle del Día:"
      Height          =   4455
      Left            =   1560
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   8295
      Begin MSFlexGridLib.MSFlexGrid griddet 
         Height          =   3615
         Left            =   120
         TabIndex        =   14
         Tag             =   "9999"
         Top             =   600
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         BackColorFixed  =   4210688
         ForeColorFixed  =   16777215
         GridColorFixed  =   8421376
         FocusRect       =   2
         HighLight       =   2
         GridLines       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   3
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
      Begin VB.Label Label1 
         Caption         =   "Total :"
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbltotdia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   255
         Left            =   5760
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lfecha 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbldet 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.ComboBox cmbcp 
      Height          =   315
      ItemData        =   "frmprogra.frx":0000
      Left            =   120
      List            =   "frmprogra.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.ListBox lsttipdoc 
      Height          =   1635
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdmostrar 
      Caption         =   "Mostrar"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid gridp 
      Height          =   6375
      Left            =   1320
      TabIndex        =   0
      Tag             =   "9999"
      Top             =   240
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      BackColorFixed  =   4210688
      ForeColorFixed  =   16777215
      GridColorFixed  =   8421376
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
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
   Begin MSMask.MaskEdBox txtCampo1 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   3000
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
      Left            =   0
      TabIndex        =   4
      Top             =   3600
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
   Begin VB.Label lblprog 
      Caption         =   "Cli/Pro"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblprog 
      Caption         =   "TipDoc:"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblprog 
      Caption         =   "Hasta:"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblprog 
      Caption         =   "Desde :"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblprog 
      Caption         =   "Detalle "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmprogra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmostrar_Click()
If CDate(txtCampo1.Text) > CDate(txtCampo2.Text) Then
 MsgBox "Fechas Invalidas", 48, Pub_Titulo
 Exit Sub
End If

Dim COL As Integer
Dim ws_tipdocs As String
Dim WE_FECHA As Date
Dim wdias As Integer
ws_tipdocs = ""
gridp.Clear
gridp.Cols = 2
gridp.Rows = 2
For fila = 0 To lsttipdoc.ListCount - 1
  If lsttipdoc.Selected(fila) Then
     gridp.Cols = gridp.Cols + 1
     gridp.TextMatrix(0, gridp.Cols - 1) = lsttipdoc.List(fila)
     gridp.ColWidth(gridp.Cols - 1) = 900
     ws_tipdocs = ws_tipdocs + "'" & Left(lsttipdoc.List(fila), 2) + "', "
  End If
Next fila
If ws_tipdocs <> "" Then
ws_tipdocs = "(" & Mid(Trim(ws_tipdocs), 1, Len(ws_tipdocs) - 2) & ")"
End If
gridp.TextMatrix(0, 0) = "Fecha"
gridp.TextMatrix(0, 1) = "Día"
gridp.ColWidth(0) = 900
gridp.ColWidth(1) = 500
wdias = CDate(txtCampo2.Text) - CDate(txtCampo1.Text) + 1
WE_FECHA = CDate(txtCampo1.Text)
gridp.Rows = 2 + wdias
For fila = 2 To gridp.Rows - 1
    gridp.RowHeight(fila) = 285
    gridp.TextMatrix(fila, 0) = Format(DateAdd("d", fila - 2, WE_FECHA), "dd/mm/yyyy")
    If Weekday(gridp.TextMatrix(fila, 0)) = 1 Then
       gridp.COL = 1
       gridp.Row = fila
       gridp.CellForeColor = vbRed
    End If
    gridp.TextMatrix(fila, 1) = UCase(Format(DateAdd("d", fila - 2, WE_FECHA), "ddd"))
Next fila
PUB_CP = cmbcp.Text
For fila = 2 To gridp.Rows - 1
   For COL = 2 To gridp.Cols - 1
        PUB_TIPDOC = Trim(gridp.TextMatrix(0, COL))
        WE_FECHA = CDate(gridp.TextMatrix(fila, 0))
        pub_cadena = "SELECT SUM(CAR_IMPORTE) AS TOTAL FROM CARTERA WHERE " & _
            "CAR_CP = '" & PUB_CP & "' AND CAR_CODCIA = '" & LK_CODCIA & "' AND CAR_TIPDOC = '" & PUB_TIPDOC & "' AND CAR_FECHA_VCTO = '" & WE_FECHA & "' AND CAR_IMPORTE <> 0"
        Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
        If Not X.EOF Then
          gridp.TextMatrix(fila, COL) = Format(Nulo_Valor0(X!total), "##,##0.00")
        End If
   Next COL
Next fila
gridp.Cols = gridp.Cols + 1
gridp.TextMatrix(0, gridp.Cols - 1) = "Total"
gridp.ColWidth(gridp.Cols - 1) = 1200
If gridp.Cols <= 3 Then GoTo sale
For fila = 2 To gridp.Rows - 1
    pub_cadena = "SELECT SUM(CAR_IMPORTE) AS TOTAL FROM CARTERA WHERE " & _
            "CAR_CP = '" & PUB_CP & "' AND CAR_CODCIA = '" & LK_CODCIA & "' AND CAR_TIPDOC in " & ws_tipdocs & "  AND CAR_FECHA_VCTO = '" & Format(gridp.TextMatrix(fila, 0), "dd/mm/yyyy") & "'   AND CAR_IMPORTE <> 0"
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    If Not X.EOF Then
      gridp.TextMatrix(fila, gridp.Cols - 1) = Format(Nulo_Valor0(X!total), "##,##0.00")
      gridp.COL = gridp.Cols - 1
      gridp.Row = fila
      gridp.CellFontBold = True
    End If
    
Next fila
sale:

End Sub

Private Sub Form_Activate()
cmbcp.ListIndex = 0
End Sub

Private Sub Form_Load()
CenterMe frmprogra
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtCampo1.Mask = "##/##/####"
 txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtCampo2.Mask = "##/##/####"
PUB_TIPREG = 8
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
lsttipdoc.ToolTipText = "TAB_TIPREG = 8"
lsttipdoc.Clear
Do Until tab_mayor.EOF
   lsttipdoc.AddItem tab_mayor!tab_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
   tab_mayor.MoveNext
Loop
cmbcp.AddItem "P"
cmbcp.AddItem "C"

End Sub

Public Sub det_dia(wcp As String, WDIA As Date, WTIPDOC As String)

Dim acu_suma As Currency
If Not IsDate(WDIA) Then Exit Sub
PSCAR_TOPE(0) = LK_CODCIA
PSCAR_TOPE(1) = wcp
PSCAR_TOPE(2) = WDIA
car_tope.Requery

lfecha.Caption = WDIA
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


griddet.ColWidth(0) = 800
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
 If Trim(WTIPDOC) <> Trim(car_tope!car_TIPDOC) Then GoTo pasa
 griddet.Rows = griddet.Rows + 1
 griddet.RowHeight(griddet.Rows - 1) = 285
 griddet.TextMatrix(griddet.Rows - 1, 0) = Format(car_tope!car_fecha_sunat, "dd/mm/yy")
 griddet.TextMatrix(griddet.Rows - 1, 1) = car_tope!car_TIPDOC
 griddet.TextMatrix(griddet.Rows - 1, 2) = car_tope!car_numser_c
 If car_tope!car_NUMFAC_C = 0 Then
   griddet.TextMatrix(griddet.Rows - 1, 3) = car_tope!car_NUMGUIA
 Else
   griddet.TextMatrix(griddet.Rows - 1, 3) = car_tope!car_NUMFAC_C
 End If
 If Left(car_tope!car_TIPDOC, 1) = "L" And Trim(car_tope!car_num_cheque) <> "" Then
    griddet.TextMatrix(griddet.Rows - 1, 2) = car_tope!car_numser_c
    griddet.TextMatrix(griddet.Rows - 1, 3) = Trim(car_tope!car_num_cheque)
 ElseIf Left(car_tope!car_TIPDOC, 1) = "L" And Trim(car_tope!car_num_cheque) = "" Then
    griddet.TextMatrix(griddet.Rows - 1, 2) = car_tope!car_numser_c
    griddet.TextMatrix(griddet.Rows - 1, 3) = Trim(car_tope!car_NUMFAC_C)
 ElseIf (Left(car_tope!car_TIPDOC, 2) = "CC" Or Left(car_tope!car_TIPDOC, 2) = "FA") And car_tope!car_NUMFAC_C = 0 Then
    griddet.TextMatrix(griddet.Rows - 1, 2) = car_tope!CAR_NUMSER
    griddet.TextMatrix(griddet.Rows - 1, 3) = car_tope!CAR_NUMFAC
 End If
 griddet.TextMatrix(griddet.Rows - 1, 4) = car_tope!car_concepto
 
 griddet.TextMatrix(griddet.Rows - 1, 5) = Format(car_tope!total, "##,##0.00")
 acu_suma = acu_suma + Val(car_tope!total)
 griddet.TextMatrix(griddet.Rows - 1, 6) = Format(car_tope!car_fecha_VCTO_orig, "dd/mm/yy")
 
 SQ_OPER = 1
 pu_codclie = car_tope!CAR_CODCLIE
 pu_cp = wcp
 pu_codcia = LK_CODCIA
 LEER_CLI_LLAVE
 If cli_llave.EOF Then
   griddet.TextMatrix(griddet.Rows - 1, 7) = "No Encontrado"
 Else
   griddet.TextMatrix(griddet.Rows - 1, 7) = Trim(cli_llave!CLI_NOMBRE)
 End If
 
pasa:
 car_tope.MoveNext
Loop
lbltotdia.Caption = Format(acu_suma, "##,##0.00")
fradetalle.Visible = True
griddet.SetFocus


End Sub

Private Sub griddet_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  fradetalle.Visible = False
End If
End Sub

Private Sub gridp_DblClick()
det_dia cmbcp.Text, gridp.TextMatrix(gridp.Row, 0), Left(gridp.TextMatrix(0, gridp.COL), 2)
End Sub

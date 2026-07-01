VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmmorosidad 
   Caption         =   "Analisis de Morosidad"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid gridp 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Tag             =   "9999"
      Top             =   960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7858
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
   Begin VB.Label lmpro 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lmmin 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lmmax 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblmor 
      Caption         =   "Morosidad Maxima:"
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
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   9
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblmor 
      Caption         =   "Morosidad Minima:"
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
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   8
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblmor 
      Caption         =   "Promedio Morosidad:"
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
      Height          =   375
      Index           =   3
      Left            =   5640
      TabIndex        =   7
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblnombre 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label lblcodi 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblmor 
      Caption         =   "Detalle de Pagos."
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
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblmor 
      Caption         =   "Nombre :"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblmor 
      Caption         =   "Codigo:"
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
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblm 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmmorosidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim wcompras As String
If pu_codclie = 0 Then Exit Sub
wcompras = InputBox("Porcentaje de sus Compras de Clientes ", "Ultmas Compras", 50)
If wcompras = "" Then Exit Sub
If Val(wcompras) = 0 Then Exit Sub
cade
Muestra_analis pu_codclie, pu_cp, Val(wcompras)
lblcodi.Caption = pu_codclie
SQ_OPER = 1
pu_codcia = LK_CODCIA
LEER_CLI_LLAVE
lblnombre.Caption = Trim(cli_llave!cli_nombre)
CenterMe frmmorosidad
End Sub

Public Sub cade()
gridp.Cols = 9
gridp.Rows = 2
gridp.TextMatrix(0, 0) = "Fec.Emis"
gridp.TextMatrix(0, 1) = "Documento"
gridp.ColWidth(1) = 1400
gridp.TextMatrix(0, 2) = "Fec.V.O."
gridp.TextMatrix(0, 3) = "Fec.V.PAG"
gridp.TextMatrix(0, 4) = "Imp.Orig."
gridp.TextMatrix(0, 5) = "Amortiz"
gridp.TextMatrix(0, 6) = "Saldo"
gridp.TextMatrix(0, 7) = "D.MORS"
gridp.TextMatrix(0, 8) = "Porc(%)"
End Sub

Public Sub Muestra_analis(wf_codclie As Currency, wf_cp As String, wf_compras As Integer)
Dim sum_dm As Currency
Dim rs_docum As rdoResultset
Dim W_SALDO As Currency
Dim wmin As Integer
Dim wmax As Integer
Dim wprom As Integer
Dim wcon As Integer
Screen.MousePointer = 13
If wf_compras = 0 Then Exit Sub
gridp.Rows = 1
pub_cadena = "SELECT  * FROM CARTERA WHERE CAR_CODCIA = '" & LK_CODCIA & "' AND CAR_CP = '" & wf_cp & "' AND CAR_CODCLIE = " & wf_codclie & "  AND CAR_IMPORTE <> 0 AND CAR_TIPDOC IN ('FA', 'CC')"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
Do Until X.EOF
        gridp.Rows = gridp.Rows + 1
        gridp.TextMatrix(gridp.Rows - 1, 0) = Format(X!car_fecha_sunat, "dd/mm/yyyy")
        gridp.TextMatrix(gridp.Rows - 1, 1) = X!car_TIPDOC & " " & X!car_fbg & " /. " & X!car_NUMSER & " - " & X!car_NUMFAC
        gridp.TextMatrix(gridp.Rows - 1, 2) = Format(X!car_fecha_VCTO_orig, "dd/mm/yyyy")
        gridp.TextMatrix(gridp.Rows - 1, 3) = "" 'Format(X!CAA_FECHA_COBRO, "dd/mm/yyyy")
        gridp.TextMatrix(gridp.Rows - 1, 4) = Format(X!CAR_IMP_INI, "##,##0.00")
        gridp.TextMatrix(gridp.Rows - 1, 5) = "" ' Format(Abs(X!CAA_IMPORTE), "##,##0.00")
        gridp.TextMatrix(gridp.Rows - 1, 6) = Format(X!car_importe, "##,##0.00")
        gridp.TextMatrix(gridp.Rows - 1, 7) = Val(LK_FECHA_DIA - X!CAr_FECHA_VCTO)
        gridp.TextMatrix(gridp.Rows - 1, 8) = "Pendiente."
X.MoveNext
Loop


pub_cadena = "SELECT TOP " & wf_compras & " * FROM CARTERA WHERE  " & _
" CAR_CODCIA = '" & LK_CODCIA & "' AND CAR_TIPMOV = 10 AND CAR_TIPDOC in ('FA','CC') AND " & _
" CAR_CODCLIE = '" & wf_codclie & "' AND CAR_CP = '" & wf_cp & "' ORDER BY CAR_CODCIA,CAR_FECHA_SUNAT DESC " ''AND CAR_FECHA_INGR >= '" & Format(wf_fecha_ini, "dd/mm/yyyy") & "'"
Set rs_docum = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
gridp.Rows = gridp.Rows + 1
wmin = 30000
wmax = -30000
sum_dm = 0
wcon = 0
Do Until rs_docum.EOF
    pub_cadena = "SELECT * FROM CARACU WHERE CAA_CODCIA = '" & rs_docum!car_codcia & "' AND CAA_TIPDOC = '" & rs_docum!car_TIPDOC & "' AND " & _
     " CAA_ESTADO <> 'E' AND CAA_SIGNO_CAR = -1 AND CAA_CODCLIE = '" & rs_docum!CAR_CODCLIE & "'  AND CAA_CP = '" & rs_docum!CAR_cp & "' AND CAA_SERDOC = " & rs_docum!car_serdoc & " AND CAA_NUMDOC = " & rs_docum!car_NUMDOC & " ORDER BY CAA_FECHA  DESC "
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    fila = 0
    W_SALDO = Val(rs_docum!CAR_IMP_INI)
    Do Until X.EOF
        gridp.Rows = gridp.Rows + 1
        gridp.TextMatrix(gridp.Rows - 1, 0) = Format(rs_docum!car_fecha_sunat, "dd/mm/yyyy")
        gridp.TextMatrix(gridp.Rows - 1, 1) = rs_docum!car_TIPDOC & " " & rs_docum!car_fbg & " /. " & rs_docum!car_NUMSER & " - " & rs_docum!car_NUMFAC
        gridp.TextMatrix(gridp.Rows - 1, 2) = Format(rs_docum!car_fecha_VCTO_orig, "dd/mm/yyyy")
        gridp.TextMatrix(gridp.Rows - 1, 3) = Format(X!CAA_FECHA_COBRO, "dd/mm/yyyy")
        gridp.TextMatrix(gridp.Rows - 1, 4) = Format(rs_docum!CAR_IMP_INI, "##,##0.00")
        gridp.TextMatrix(gridp.Rows - 1, 5) = Format(Abs(X!CAA_IMPORTE), "##,##0.00")
        W_SALDO = W_SALDO + Val(X!CAA_IMPORTE)
        gridp.TextMatrix(gridp.Rows - 1, 6) = Format(W_SALDO, "##,##0.00")
        gridp.TextMatrix(gridp.Rows - 1, 7) = Val(X!CAA_FECHA_COBRO - rs_docum!car_fecha_VCTO_orig)
        If wmin > Val(gridp.TextMatrix(gridp.Rows - 1, 7)) Then
            wmin = Val(gridp.TextMatrix(gridp.Rows - 1, 7))
        End If
        If Val(gridp.TextMatrix(gridp.Rows - 1, 7)) > wmax Then
            wmax = Val(gridp.TextMatrix(gridp.Rows - 1, 7))
        End If
        If Val(gridp.TextMatrix(gridp.Rows - 1, 7)) >= 1 Then
          wcon = wcon + 1
          sum_dm = sum_dm + Val(gridp.TextMatrix(gridp.Rows - 1, 7))
        End If
        If Val(rs_docum!CAR_IMP_INI) <> 0 Then gridp.TextMatrix(gridp.Rows - 1, 8) = Format(Val(Abs(X!CAA_IMPORTE)) * 100 / Val(rs_docum!CAR_IMP_INI), "0.00")
    X.MoveNext
    Loop
     If gridp.Rows - 1 <> 0 Then
     gridp.Row = gridp.Rows - 1
     gridp.COL = 6
     gridp.CellBackColor = vbYellow
    End If
rs_docum.MoveNext
DoEvents
Loop
If wcon <> 0 Then wprom = Format(sum_dm / wcon, "0")
Screen.MousePointer = 0
lmmin.Caption = wmin
lmmax.Caption = wmax
lmpro.Caption = wprom


End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FrmMotAnul 
   Caption         =   "Documento :"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox cheerror 
      Caption         =   "Anular por Error de Impresión y Otros de Documento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   360
      Width           =   4335
   End
   Begin VB.Frame framotivos 
      Caption         =   "Motivos . "
      Height          =   615
      Left            =   6600
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ComboBox cmbmotivo 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      Picture         =   "FrmMotAnul.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9735
      Begin MSFlexGridLib.MSFlexGrid gridfac 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5741
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
      End
   End
   Begin VB.Label lblmonto 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Monto : "
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblcliente 
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "FrmMotAnul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Jalar_doc(wsfec As String, wscia As String, wsno As Currency, WSFBG As String, wsnumser As Currency, wsnumfac As Currency)
pub_cadena = "SELECT * FROM facart WHERE FAR_FECHA = '" & Format(wsfec, "dd/mm/yyyy") & "' AND FAR_NUMOPER = " & wsno & " AND FAR_CODCIA = '" & wscia & "' AND FAR_ESTADO <> 'E' AND FAR_FBG = '" & WSFBG & "' AND FAR_NUMSER = " & wsnumser & " AND FAR_NUMFAC = " & wsnumfac & " ORDER BY  FAR_NUMSEC"
Set far_menor3 = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If far_menor3.EOF Then
  Exit Sub
End If
gridfac.Cols = 10
gridfac.Rows = 1
gridfac.ColWidth(0) = 900
gridfac.ColWidth(1) = 300
gridfac.ColWidth(2) = 400
gridfac.ColWidth(3) = 800
gridfac.ColWidth(4) = 2500
gridfac.ColWidth(5) = 1200
gridfac.ColWidth(6) = 3000
gridfac.ColWidth(7) = 0
gridfac.ColWidth(8) = 0
gridfac.ColWidth(9) = 0

gridfac.TextMatrix(0, 0) = "FECHA"
gridfac.TextMatrix(0, 1) = "TIPO"
gridfac.TextMatrix(0, 2) = "SERIE"
gridfac.TextMatrix(0, 3) = "NUMERO"
gridfac.TextMatrix(0, 4) = "PRODUCTO"
gridfac.TextMatrix(0, 5) = "MONTO"
gridfac.TextMatrix(0, 6) = "GLOSA - MOTIVO"
gridfac.TextMatrix(0, 7) = 0
gridfac.TextMatrix(0, 8) = 0
gridfac.TextMatrix(0, 9) = 0
SQ_OPER = 1
pu_cp = "C"
pu_codclie = far_menor3!far_codclie
pu_codcia = LK_CODCIA
LEER_CLI_LLAVE
If cli_llave.EOF Then
  lblcliente.Caption = far_menor3!far_cliente
Else
  lblcliente.Caption = cli_llave!cli_nombre
End If


lblmonto.Caption = Format(Val(far_menor3!far_bruto) + Val(far_menor3!far_impto), "#,##0.00")
Do Until far_menor3.EOF
gridfac.Rows = gridfac.Rows + 1
gridfac.TextMatrix(gridfac.Rows - 1, 0) = Format(far_menor3!far_fecha, "dd/mm/yy")
gridfac.TextMatrix(gridfac.Rows - 1, 1) = Trim(far_menor3!far_fbg)
gridfac.TextMatrix(gridfac.Rows - 1, 2) = Trim(far_menor3!far_numser)
gridfac.TextMatrix(gridfac.Rows - 1, 3) = Trim(far_menor3!far_numfac)
SQ_OPER = 1
PUB_KEY = far_menor3!far_codart
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
If art_LLAVE.EOF Then
 gridfac.TextMatrix(gridfac.Rows - 1, 4) = ""
Else
 gridfac.TextMatrix(gridfac.Rows - 1, 4) = Trim(art_LLAVE!ART_NOMBRE)
End If

gridfac.TextMatrix(gridfac.Rows - 1, 5) = Format(Val(far_menor3!far_subtotal), "0.00") ' + Val(far_menor3!far_impto), "0.00")
gridfac.TextMatrix(gridfac.Rows - 1, 6) = ""
gridfac.TextMatrix(gridfac.Rows - 1, 7) = far_menor3!far_codart ' COD PRODUCTO
gridfac.TextMatrix(gridfac.Rows - 1, 8) = far_menor3!far_codclie ' COD CLIENTE
gridfac.TextMatrix(gridfac.Rows - 1, 9) = far_menor3!far_numsec

far_menor3.MoveNext
Loop

End Sub


Private Sub cheerror_Click()
If cheerror.Value = 1 Then
    For fila = 1 To gridfac.Rows - 1
    gridfac.TextMatrix(fila, 6) = "Errores Propios del Documento   "
    Next fila
Else
    For fila = 1 To gridfac.Rows - 1
    gridfac.TextMatrix(fila, 6) = ""
    Next fila
End If
End Sub

Private Sub cmbmotivo_Click()
 cmbmotivo_KeyPress 13
End Sub

Private Sub cmbmotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  framotivos.Visible = False
  gridfac.SetFocus
End If
If KeyAscii = 13 Then
  gridfac.Text = cmbmotivo.Text
  framotivos.Visible = False
  gridfac.SetFocus
End If
End Sub

Private Sub cmbmotivo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = cmbmotivo.ListIndex
PUB_TIPREG = 74
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
LLENA_GRUPOS cmbmotivo, 74
cmbmotivo.SetFocus
SendKeys "%{up}"



End Sub

Private Sub cmdclose_Click()
Dim wflag As String
If gridfac.Rows <= 2 And gridfac.Cols >= 3 Then
Else
FrmMotAnul.Hide
Exit Sub
End If
wflag = ""
For fila = 1 To gridfac.Rows - 1
  If Trim(gridfac.TextMatrix(fila, 6)) <> "" Then
     wflag = "A"
  End If
Next fila
If Trim(wflag) = "" Then
   MsgBox "Seleccione algún motivo por que se esta Anulando", 48, Pub_Titulo
   Exit Sub
End If
FrmMotAnul.Hide
End Sub

Private Sub Form_Load()
 CenterMe Me
Jalar_doc Str(PUB_FECHA), LK_CODCIA, Val(pub_numoper), PUB_FBG, Val(PU_NUMSER), PU_NUMFAC
FrmMotAnul.Caption = "Anular el Documento : " & PUB_FBG & " /. " & PU_NUMSER & "  - " & PU_NUMFAC & " "
End Sub

Private Sub gridfac_DblClick()
If cheerror.Value = 1 Then
  MsgBox "Desmarcar la opción de errores de impresión", 48, Pub_Titulo
  cheerror.SetFocus
  Exit Sub
End If
If gridfac.COL = 6 Then
  LLENA_GRUPOS cmbmotivo, 74
  ASIGNA_INT cmbmotivo, Val(Trim(Right(gridfac.Text, 10)))
  framotivos.Visible = True
  cmbmotivo.SetFocus
  SendKeys "%{up}"
End If

End Sub
Public Sub LLENA_GRUPOS(cont As ComboBox, tip As Integer)
    PUB_TIPREG = tip
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_nomlargo & String(80, " ") & tab_mayor!tab_numtab
        tab_mayor.MoveNext
    Loop
    
End Sub

Public Sub ASIGNA_INT(WCONTROL As ComboBox, txt As Currency)
For fila = 0 To WCONTROL.ListCount - 1
    If Val(Trim(Right(WCONTROL.List(fila), 8))) = txt Then
        WCONTROL.ListIndex = fila
        Exit Sub
    End If
Next fila
End Sub

Private Sub gridfac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
gridfac_DblClick
End If
End Sub

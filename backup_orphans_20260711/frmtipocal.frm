VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtipocal 
   Caption         =   "Procesando Calculo de tipo de Cliente"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Ce&rrar"
      Height          =   735
      Left            =   3840
      TabIndex        =   12
      Top             =   3600
      Width           =   615
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdactva 
      Caption         =   "Actulizar Valores."
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtdias 
      Height          =   285
      Left            =   3000
      TabIndex        =   8
      Text            =   "0"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox tp2 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox tp1 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdprocesar 
      Caption         =   "Procesar Distribución de Tipo de Cliente"
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
   End
   Begin VB.ComboBox cmdtipo 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lbltipo 
      Caption         =   "Día(s)."
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   9
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lbltipo 
      Caption         =   "Ademas su morosidad no exeda a :"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lbltipo 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label lbltipo 
      Caption         =   "Aquellos que  sus Ventas al Credito y Contado  esta dentro de S/.:"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Cliente."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmtipocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdactva_Click()

SQ_OPER = 1
PUB_NUMTAB = Trim(Right(cmdtipo.Text, 8))
PUB_TIPREG = 230
LEER_TAB_LLAVE
If tab_llave.EOF Then
Else
    tab_llave.Edit
    tab_llave!tab_codart = Val(tp1.Text)
    tab_llave!tab_codclie = Val(tp2.Text)
    tab_llave!tab_contable2 = Val(txtdias.Text)
    tab_llave.Update
End If


End Sub

Private Sub cmdclose_Click()
Unload frmtipocal
End Sub


Private Sub cmdprocesar_Click()
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset

pub_cadena = "SELECT * FROM CLIENTES WHERE FROM CLI_CODCIA = ? AND CLI_CP = 'C' AND CLI_CODCLIE <> 1 "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = LK_CODCIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
Do Until llave_rep01.EOF
    pub_cadena = "SELECT * FROM VENTASCLI FROM FAR_CODCIA = '" & LK_CODCIA & "' AND FAR_CODCLIE = " & llave_rep01!CLI_CODCLIE & " "
    Set sp_venta = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    ws_total = sp_venta!TOTAL
    
llave_rep01.MoveNext
Loop
End Sub

Private Sub cmdtipo_Click()

SQ_OPER = 1
PUB_NUMTAB = Trim(Right(cmdtipo.Text, 8))
PUB_TIPREG = 230
LEER_TAB_LLAVE
If tab_llave.EOF Then
Else
    tp1.Text = Val(Nulo_Valor0(tab_llave!tab_codart))
    tp2.Text = Val(Nulo_Valor0(tab_llave!tab_codclie))
    txtdias.Text = Val(Nulo_Valor0(tab_llave!tab_contable2))
End If
Azul tp1, tp1
End Sub

Private Sub Form_Load()
LLENA_GRUPOS cmdtipo, 230
End Sub
Public Sub LLENA_GRUPOS(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    PUB_CODCIA = LK_CODCIA
    If tip = 66 Then
    PUB_CODCIA = "00"
    End If
    
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    If PUB_TIPREG = 333 Then lisdescto.Clear
    Do Until tab_mayor.EOF
         cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
        If PUB_TIPREG = 333 Then lisdescto.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
    
End Sub

Private Sub tp1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul tp2, tp2
End If
End Sub

Private Sub tp2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtdias, txtdias
End If
End Sub

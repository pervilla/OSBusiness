VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmConProAnul 
   Caption         =   "Consulta de Productos Anulados - Motivos"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pb 
      Height          =   4215
      Left            =   6600
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   7435
      _Version        =   393216
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdmst 
      Caption         =   "Mostar Productos por Verificar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6360
      Picture         =   "frmConProAnul.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   435
   End
   Begin MSFlexGridLib.MSFlexGrid Gridmst 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7435
      _Version        =   393216
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Doble click = Quitar de Lista."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Doble click = Quitar de Lista."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmConProAnul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub muestraGrid()
Gridmst.Cols = 8
Gridmst.Rows = 1
Gridmst.ColWidth(0) = 900
Gridmst.ColWidth(1) = 2500
Gridmst.ColWidth(2) = 1200
Gridmst.ColWidth(3) = 1000
Gridmst.ColWidth(4) = 0
Gridmst.ColWidth(5) = 0
Gridmst.ColWidth(6) = 0
Gridmst.ColWidth(7) = 0
Gridmst.TextMatrix(0, 0) = "Codigo"
Gridmst.TextMatrix(0, 1) = "Producto"
Gridmst.TextMatrix(0, 2) = "Documento"
Gridmst.TextMatrix(0, 3) = "Motivo"
pub_cadena = "SELECT  * FROM TABMOTIVOS, ARTI WHERE RAN_CODCIA = '" & LK_CODCIA & "' AND (RAN_CODART = ART_KEY) AND (RAN_CODCIA = ART_CODCIA) AND RAN_ESTADO <> 'R' AND RAN_CODMOT <> 0 ORDER BY ART_NOMBRE"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not X.EOF Then
  pb.Visible = True
  pbValue = 0
  pbMin = 0
  pbMax = X.RowCount
End If

SQ_OPER = 1
Do Until X.EOF
  pbValue = pbValue + 1
  DoEvents
  Gridmst.Rows = Gridmst.Rows + 1
  PUB_KEY = X!ran_codart
  pu_codcia = LK_CODCIA
  LEER_ART_LLAVE
  Gridmst.TextMatrix(Gridmst.Rows - 1, 0) = art_LLAVE!art_key
  Gridmst.TextMatrix(Gridmst.Rows - 1, 1) = Trim(art_LLAVE!art_nombre)
  Gridmst.TextMatrix(Gridmst.Rows - 1, 2) = X!ran_fbg & " / " & X!ran_numser & X!ran_numFAC
  PUB_TIPREG = 74
  PUB_NUMTAB = Val(X!ran_codmot)
  PUB_CODCIA = LK_CODCIA
  LEER_TAB_LLAVE
  Gridmst.TextMatrix(Gridmst.Rows - 1, 3) = Trim(tab_llave!TAB_NOMLARGO)
  Gridmst.TextMatrix(Gridmst.Rows - 1, 4) = X!ran_fbg
  Gridmst.TextMatrix(Gridmst.Rows - 1, 5) = X!ran_numser
  Gridmst.TextMatrix(Gridmst.Rows - 1, 6) = X!ran_numFAC
  Gridmst.TextMatrix(Gridmst.Rows - 1, 7) = X!ran_numSEC
  X.MoveNext
Loop
pb.Visible = False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdmst_Click()
muestraGrid
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterMe Me
End Sub

Private Sub Gridmst_DblClick()
If Val(Gridmst.TextMatrix(Gridmst.Row, 0)) = 0 Then Exit Sub
pub_mensaje = "Quitar de le lista  ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If

pub_cadena = "UPDATE TABMOTIVOS SET RAN_ESTADO = 'R' WHERE RAN_CODCIA = '" & LK_CODCIA & "' AND RAN_FBG = '" & Gridmst.TextMatrix(Gridmst.Row, 4) & "' AND RAN_NUMSER = '" & Gridmst.TextMatrix(Gridmst.Row, 5) & "' AND RAN_NUMFAC = '" & Gridmst.TextMatrix(Gridmst.Row, 6) & "' AND RAN_NUMSEC = '" & Gridmst.TextMatrix(Gridmst.Row, 7) & "' "
CN.Execute pub_cadena
MsgBox "Producto Quitado", 48, Pub_Titulo
muestraGrid
End Sub

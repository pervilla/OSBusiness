VERSION 5.00
Begin VB.Form frmrela 
   Caption         =   "Relación de Usuarios con Lineas"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdveri 
      Caption         =   "Verificar Lineas sin Usuarios"
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ComboBox cmdlista 
      Height          =   315
      Left            =   3450
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1110
      Width           =   2355
   End
   Begin VB.CommandButton CMDADD 
      Caption         =   "<< Adicionar"
      Height          =   345
      Left            =   3450
      TabIndex        =   5
      Top             =   1530
      Width           =   1245
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   3390
      Picture         =   "frmrela.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3990
      Width           =   1125
   End
   Begin VB.ComboBox cmduser 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   300
      Width           =   4455
   End
   Begin VB.ListBox lislineas 
      Height          =   3765
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   930
      Width           =   3045
   End
   Begin VB.Label Label1 
      Caption         =   "Para Quitar una Linea a un Usuario : seleccione un linea y pulse <DEL o SUPR>"
      Height          =   765
      Index           =   3
      Left            =   3330
      TabIndex        =   8
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Agregar una Linea  :"
      Height          =   255
      Index           =   2
      Left            =   3450
      TabIndex        =   7
      Top             =   870
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Lista de Linea relacionado al Usuario  :"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   660
      Width           =   3045
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario del Sistema :"
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   3
      Top             =   30
      Width           =   4815
   End
End
Attribute VB_Name = "frmrela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDADD_Click()
Dim codi As Integer
If Trim(cmdlista.text) = "" Then Exit Sub
codi = Val(Right(cmdlista.text, 8))
Dim llave_USER As rdoResultset
pub_cadena = "SELECT * FROM RELUSER WHERE REL_CODCIA = '" & LK_CODCIA & "' and rel_fami = " & codi
Set llave_USER = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
llave_USER.Requery
If Not llave_USER.EOF Then
 MsgBox "La Linea esta Asignado a uin Usuario : " & llave_USER!REL_CODUSU, 48, Pub_Titulo
 Exit Sub
End If
llave_USER.AddNew
llave_USER!rel_codcia = LK_CODCIA
llave_USER!REL_CODUSU = Trim(Left(cmduser.text, 10))
llave_USER!rel_fami = codi
llave_USER.Update
cmduser_Click

End Sub

Private Sub cmduser_Click()


Dim llave_USER As rdoResultset

pub_cadena = "SELECT * FROM RELUSER WHERE REL_CODCIA = '" & LK_CODCIA & "' AND REL_CODUSU = '" & Trim(Left(cmduser.text, 10)) & "'"
Set llave_USER = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
llave_USER.Requery
lislineas.Clear
Do Until llave_USER.EOF
  SQ_OPER = 1
  PUB_CODCIA = LK_CODCIA
  PUB_TIPREG = 122
  PUB_NUMTAB = Val(llave_USER!rel_fami)
  LEER_TAB_LLAVE
  If tab_llave.EOF Then
  Else
    lislineas.AddItem Trim(tab_llave!tab_nomlargo) & String(100, " ") & tab_llave!tab_numtab
  End If
llave_USER.MoveNext
Loop

End Sub

Private Sub cmdveri_Click()
Dim codi As Integer
Dim texto_ver As String
codi = Val(Right(cmdlista.text, 8))
Dim llave_USER As rdoResultset

SQ_OPER = 2
PUB_TIPREG = 122
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
cmdveri.Caption = ".."
texto_ver = "Lineas sin asiganar a un Usuario: " & Chr(13)
Do Until tab_mayor.EOF
  cmdveri.Caption = tab_mayor.AbsolutePosition & "/" & tab_mayor.RowCount
  DoEvents
  pub_cadena = "SELECT * FROM RELUSER WHERE REL_CODCIA = '" & LK_CODCIA & "' and rel_fami = " & tab_mayor!tab_numtab
  Set llave_USER = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
  llave_USER.Requery
  If llave_USER.EOF Then
   texto_ver = texto_ver + Trim(tab_mayor!tab_nomlargo) & ", "
  End If
  tab_mayor.MoveNext
Loop
cmdveri.Caption = "Verificar Lineas sin Usuarios"
MsgBox texto_ver


End Sub

Private Sub Form_Load()

CenterMe frmrela

Dim ws_user As String * 10
usu.Requery

Do Until usu.EOF
 ws_user = Trim(usu!usu_key)
 cmduser.AddItem ws_user & " " & usu!usu_nombre
 usu.MoveNext
Loop
SQ_OPER = 2
PUB_CODCIA = LK_CODCIA
PUB_TIPREG = 122
LEER_TAB_LLAVE
cmdlista.Clear
Do Until tab_mayor.EOF
  cmdlista.AddItem tab_mayor!tab_nomlargo & String(100, " ") & tab_mayor!tab_numtab
tab_mayor.MoveNext
Loop

End Sub

Private Sub lislineas_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 46 Then Exit Sub
Dim codi As Integer
If Trim(lislineas.text) = "" Then Exit Sub
codi = Val(Right(lislineas.text, 8))
Dim llave_USER As rdoResultset
pub_cadena = "SELECT * FROM RELUSER WHERE REL_CODCIA = '" & LK_CODCIA & "' and rel_fami = " & codi
Set llave_USER = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
llave_USER.Requery
If llave_USER.EOF Then
   MsgBox "ya fue quitada", 48, Pub_Titulo
   Exit Sub
End If
llave_USER.Delete
cmduser_Click

End Sub

Private Sub SALIR_Click()
Unload frmrela
End Sub

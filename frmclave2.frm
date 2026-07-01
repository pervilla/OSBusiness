VERSION 5.00
Begin VB.Form frmclave2 
   Caption         =   "Acceso Especial."
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FRACLAVE2 
      Caption         =   "Acceso a Reportes de Gerencia : "
      Height          =   735
      Left            =   420
      TabIndex        =   0
      Top             =   -30
      Width           =   3255
      Begin VB.TextBox txtclave2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   570
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmclave2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim j As Integer
Dim I As Integer
j = 0
For I = 1 To lk_OTROS_Count
   If Val(lk_OTROS(I)) = 25 Then 'Muestra Clave de Gerencia
     j = 1
   End If
Next I


LK_ACCESO_REPORT = ""
LK_ACCESO_EXTORNO = ""
'If Trim(frmclave2.Caption) = "CLAVE" Then Exit Sub

If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Then
Else
  If j = 0 Then Unload frmclave2

End If
End Sub

Private Sub Form_Load()
CenterMe frmclave2
End Sub


Private Sub txtclave2_KeyPress(KeyAscii As Integer)
LK_ACCESO_REPORT = ""
LK_ACCESO_EXTORNO = ""
If KeyAscii = 27 Then GoTo fin
If KeyAscii <> 13 Then Exit Sub

'*----------------------*
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Then
  LK_ACCESO_REPORT = "A"
  LK_ACCESO_EXTORNO = "A"
  GoTo fin
End If

If Trim(frmclave2.Caption) = "CLAVE" Then
    pub_cadena = "select usu_key , USU_CLAVE from usuarios where usu_clave = '" & Trim(txtclave2.Text) & "' and usu_otros like '%.28%'"
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    If Not X.EOF Then
      frmclave2.Caption = "Usuario : " & Trim(X!usu_key)
      LK_ACCESO_REPORT = "A"
      LK_ACCESO_EXTORNO = "A"
      GoTo fin
    End If
    MsgBox "Clave Incorrecta", 48, Pub_Titulo
    Exit Sub
End If
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
  PSUSU_LLAVE(0) = LK_CODUSU
Else
  PSUSU_LLAVE(0) = LK_CODUSU
End If
usu_llave.Requery

If LK_CLAVE_GERENCIA = Trim(txtclave2.Text) Then
   LK_ACCESO_REPORT = "A"
   LK_ACCESO_EXTORNO = "A"
Else
   MsgBox "Clave Incorrecta", 48, Pub_Titulo
End If
'*----------------------*
fin:
Unload frmclave2
End Sub


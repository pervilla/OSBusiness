VERSION 5.00
Begin VB.Form frmprecios 
   BackColor       =   &H00808000&
   Caption         =   "Datos de Precios Otros."
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   FillColor       =   &H00404000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbcomp 
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1560
      Width           =   2685
   End
   Begin VB.ComboBox cmdopcion 
      ForeColor       =   &H00000080&
      Height          =   315
      ItemData        =   "frmprecios.frx":0000
      Left            =   1590
      List            =   "frmprecios.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   3855
   End
   Begin VB.ComboBox unidad 
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1050
      Width           =   2205
   End
   Begin VB.TextBox txtprove 
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
      Height          =   300
      Left            =   1620
      TabIndex        =   3
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox txtprecio 
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
      Height          =   300
      Left            =   1620
      TabIndex        =   2
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label codast 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   135
      Left            =   5160
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. F.C.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4320
      TabIndex        =   14
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lblcantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Competencia :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Opcion : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   90
      TabIndex        =   10
      Top             =   150
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Presentación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1050
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   8
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Label lblpres 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4020
      TabIndex        =   6
      Top             =   1020
      Width           =   1365
   End
   Begin VB.Label lblnombre 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1590
      TabIndex        =   5
      Top             =   660
      Width           =   4035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   690
      Width           =   1185
   End
End
Attribute VB_Name = "frmprecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbcomp_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload frmprecios
  Exit Sub
End If

If KeyAscii = 13 Then
  Azul txtprecio, txtprecio
End If
End Sub

Private Sub cmdopcion_Click()
If Val(Left(cmdopcion.Text, 1)) = 1 Then
  Label1(2).Caption = "Precio : "
  Label1(3).Caption = "Observaciro n :"
  frmprecios.Caption = "Datos de Precios Otros."
  cmbcomp.Enabled = True
Else
  Label1(2).Caption = "Cantidad Falt. : "
  Label1(3).Caption = "Observaciro n :"
  frmprecios.Caption = "Cotiza/ No Vendido."
  unidad.ListIndex = unidad.ListCount - 1
  cmbcomp.Enabled = False
End If

End Sub

Private Sub cmdopcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   unidad.SetFocus
End If
End Sub

Private Sub Form_Activate()
cmdopcion.ListIndex = 0
If cmbcomp.ListCount > 0 Then cmbcomp.ListIndex = 0
End Sub

Private Sub Form_Load()
CenterMe frmprecios
PUB_CODCIA = LK_CODCIA
LLENADOS cmbcomp, 68
End Sub


Private Sub txtprecio_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload frmprecios
  Exit Sub
End If

SOLO_DECIMAL txtprecio, KeyAscii
If KeyAscii = 13 Then
   txtprove.SetFocus
End If
End Sub

Private Sub txtprove_KeyPress(KeyAscii As Integer)
Dim ps_verarti As rdoResultset
Dim wfecha As Date
wfecha = DateAdd("D", -30, LK_FECHA_DIA)
If KeyAscii = 27 Then
  Unload frmprecios
  Exit Sub
End If
If KeyAscii = 13 Then
  If Val(lblcantidad.Caption) = 0 Then GoTo pasa
  pub_cadena = "SELECT * FROM HISPRE WHERE HPR_CODCIA =  '" & LK_CODCIA & "' AND HPR_CODART = " & PUB_KEY & "  AND HPR_FECHA >= '" & Format(wfecha, "dd/mm/yyyy") & "' and HPR_CODCOMP = " & Val(codast.Caption)
  Set ps_verarti = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
  If Not ps_verarti.EOF Then
  Else
     If Val(lblcantidad.Caption) <= 1 And Val(codast.Caption) <> Val(Right(cmbcomp.Text, 8)) Then
      MsgBox "Solo Falta por Ingresar el Precio para el Distribuidor con * ", 48, Pub_Titulo
      cmbcomp.SetFocus
      Exit Sub
     End If
  End If
pasa:
  frmprecios.Hide
  Exit Sub
End If

End Sub

Private Sub unidad_Click()
frmprecios.lblpres.Caption = Left(unidad.Text, 30)
End Sub

Private Sub unidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload frmprecios
  Exit Sub
End If
If KeyAscii = 13 Then
  If cmbcomp.Enabled Then
    cmbcomp.SetFocus
    SendKeysSeguro VK_UP, True
  Else
    txtprecio.SetFocus
  End If
End If

End Sub
Public Sub LLENADOS(cont As ComboBox, tip As Integer)
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    codast.Caption = ""
    Do Until tab_mayor.EOF
     If LK_CODUSU = "BOT-OPER20" Or LK_CODUSU = "OPER20" Or LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Or Val(lblcantidad.Caption) = 0 Then
       cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
       If Left(tab_mayor!TAB_NOMLARGO, 1) = "*" Then
            codast.Caption = tab_mayor!TAB_NUMTAB
       End If
     Else
       If Left(tab_mayor!TAB_NOMLARGO, 1) = "*" Then
          cont.AddItem tab_mayor!TAB_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
          codast.Caption = tab_mayor!TAB_NUMTAB
       End If
     End If
     tab_mayor.MoveNext
    Loop
End Sub


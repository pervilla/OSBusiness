VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBusq 
   Caption         =   "Busqueda Avanzadas ..."
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_key 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.ComboBox cmdTipos 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2940
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   5186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   128
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblb 
      Caption         =   "Buscar :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblb 
      Caption         =   "Tipo:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmBusq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
cmdTipos.Clear
cmdTipos.AddItem "Sintomas"
cmdTipos.AddItem "Composición"
cmdTipos.ListIndex = 0
End Sub

Private Sub txt_key_GotFocus()
 If ListView1.Visible Then
  ListView1.Visible = False
 End If
 txt_key.Text = ""
 'frmBusq.F14.Visible = False
End Sub
Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_key.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  txt_key.SelStart = Len(txt_key.Text)
  DoEvents
fin:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As MSComctlLib.ListItem

If KeyAscii = 27 Then
 txt_key.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo ERROR_CODIGO
  pu_codclie = Val(txt_key.Text)
On Error GoTo 0
If Len(txt_key.Text) = 0 Then
   Exit Sub
End If

If pu_codclie <> 0 And IsNumeric(txt_key.Text) = True Then
   LOC_OPER = 1
   PUB_CODCIA = LK_CODCIA
   On Error GoTo ERROR_CODIGO
    PUB_KEY = pu_codclie
    LEER_LOC
   On Error GoTo 0
   If artloc_llave.EOF Then
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Azul txt_key, txt_key
     GoTo fin
   Else
     If pu_codclie = 1 Then
       MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
       Azul txt_key, txt_key
       GoTo fin
     End If
     LLENA_ARTI 1
     BLOQUEA_TEXT frmBusq.txt_key
     frmBusq.CmdModificar.SetFocus
     BLOQUEA_TEXT txtnombre
     cmdCancelar.Enabled = True
   End If
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(txt_key.Text)) = Left(VALOR, Len(Trim(txt_key.Text))) Then
   Else
      Exit Sub
   End If
   LLENA_ARTI 0
   BLOQUEA_TEXT frmBusq.txt_key
   frmBusq.CmdModificar.SetFocus
   BLOQUEA_TEXT txtnombre
   cmdCancelar.Enabled = True
End If
dale:
ListView1.Visible = False
fin:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul txt_key, txt_key
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
Dim ws_codcia As String * 2
If Len(txt_key.Text) = 0 Or IsNumeric(txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(txt_key.Text) = 1 Then
    If txt_key.Text = "" Then txt_key.Text = " "
    var = Asc(txt_key.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    ElseIf var = 58 Then
       var = "A"
    Else
       var = Chr(var)
    End If
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      ws_codcia = "00"
    End If
    numarchi = 0
    
    archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK ,PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_SITUACION <> '1' AND ART_KEY  <> 1 and ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & txt_key.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
    VER_BUSQ ListView1, 1000
    loc_key = 0
    If ListView1.Visible Then
    loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As Object ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(txt_key.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If


End Sub

Public Sub VER_BUSQ(LV1 As Object, Optional wmax, Optional sp_ban)
'On Error GoTo SALE
Dim wmaximo As Integer
Dim itmX As Object '  ListItem
Dim ws_bandera As String
If Not IsMissing(wmax) Then wmaximo = wmax Else wmaximo = 2000
If Not IsMissing(sp_ban) Then ws_bandera = sp_ban Else ws_bandera = ""

Set X = CN.OpenResultset(archi, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)

LV1.ListItems.Clear
LV1.ColumnHeaders.Clear
If X.EOF Then LV1.Visible = False: Exit Sub
'MsgBox X.RowCount
LV1.Top = 1800
'LV1.Left = 3300
'LV1.Width = 6500
LV1.Height = 3200
LV1.Width = 8000
LV1.Left = 2000
'LV1.Visible = True
 LV1.ColumnHeaders.Add 1, , "Descripción", 3800
 LV1.Width = 11500
 LV1.Left = 300
 LV1.ColumnHeaders.Add 2, , "Cod.", 600
 LV1.ColumnHeaders.Add 3, , "Dirección", 3800
 LV1.ColumnHeaders.Add 4, , "Zona", 1500
Do Until X.EOF
   Set itmX = LV1.ListItems.Add(, , Trim(CStr(X.rdoColumns(3))))
   
'   If numarchi = 3 Then itmX.SubItems(1) = Trim(CStr(X.rdoColumns(2))) Else: itmX.SubItems(1) = Trim(CStr(X.rdoColumns(0)))
   
   itmX.SubItems(2) = Trim(CStr(X.rdoColumns(4))) + " # " + Trim(CStr(X.rdoColumns(6))):  itmX.SubItems(3) = Trim(CStr(X!tab_nomlargo))
SIGUE_ART:
   itmX.Tag = X.AbsolutePosition
   X.MoveNext
Loop
LV1.ToolTipText = "Encontrados : " & itmX.Tag & "/" & X.RowCount & " Muestra un Maximo de: " & wmaximo
LV1.Visible = True
'DoEvents
Exit Sub
sale:
MsgBox Err.Description
'Resume Next
If Err.Number = 40002 Then Exit Sub Else MsgBox Err.Description, 48, Pub_Titulo

End Sub


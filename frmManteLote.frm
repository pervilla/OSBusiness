VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmManteLote 
   Caption         =   "Mantenimiento de Lotes"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdactlote 
      Caption         =   "Actulizar Productos con Lotes Vencidos"
      Height          =   675
      Left            =   5640
      Picture         =   "frmManteLote.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox TEXTOVARpre 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   12632064
      BorderStyle     =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"frmManteLote.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   735
      Left            =   3720
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   128
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmbcerrar 
      Caption         =   "Ce&rrar"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdgraba 
      Caption         =   "Actualizar"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   5160
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid gridcabe 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5530
      _Version        =   393216
   End
   Begin VB.TextBox i_codart2 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblsumlote 
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
      Left            =   1200
      TabIndex        =   11
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblstock 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label LBLLT 
      Alignment       =   2  'Center
      Caption         =   "Stock Actual = "
      ForeColor       =   &H00404000&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label LBLUNIDAD 
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label i_nomarti 
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label LBLLT 
      Caption         =   "Producto :"
      ForeColor       =   &H00404000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmManteLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temporal
Dim loc_key As Integer
Dim preuni_llave As rdoResultset
Dim PRE_UNIDADES As rdoQuery


Private Sub cmbcerrar_Click()
Unload frmManteLote
End Sub



Private Sub cmdactlote_Click()
Dim canti As Currency
Dim ps_contar As rdoResultset
Dim wmeses As String
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
    Exit Sub
End If

wmeses = InputBox("Marcar los Lotes que Vencen hasta la fecha de :.... ", "Marcar los proximos a vencer", LK_FECHA_DIA)
If wmeses = "" Then Exit Sub
If Not IsDate(wmeses) Then Exit Sub
Screen.MousePointer = 13
DoEvents
ws_fecha = Format(wmeses, "dd/mm/yyyy") '  Format(LK_FECHA_DIA, "dd") & "/" & Format(LK_FECHA_DIA, "mm") + wmeses & "/" & Format(LK_FECHA_DIA, "yyyy")

pub_cadena = "UPDATE ARTI SET ART_MARCA = 0 WHERE ART_CODCIA = '" & LK_CODCIA & "' AND ART_MARCA = 1 "
CN.Execute pub_cadena, rdExecDirect


pub_cadena = "UPDATE ARTI SET ART_MARCA = 1 WHERE ART_CODCIA = '" & LK_CODCIA & "' AND ART_KEY IN (SELECT LOT_CODART FROM LOTE WHERE LOT_CODCIA = '" & LK_CODCIA & "' AND LOT_SALDOS <> 0 AND LOT_FECHA_VCTO <= '" & ws_fecha & "' AND LOT_NROLOTE <> '(*)')"
CN.Execute pub_cadena, rdExecDirect

pub_cadena = "SELECT COUNT(ART_CODCIA) AS CANTIDAD FROM ARTI  WHERE ART_CODCIA = '" & LK_CODCIA & "' AND ART_MARCA = 1 "
Set ps_contar = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
canti = 0
If Not ps_contar.EOF Then
canti = Val(Nulo_Valor0(ps_contar!cantidad))
End If

  
Screen.MousePointer = 0
MsgBox "Actlizado , OK" & Chr(13) & "Marcados : " & canti, 48, Pub_Titulo
  
End Sub

Private Sub cmdgraba_Click()
Dim wsum As Currency
For fila = 1 To gridcabe.Rows - 1
   If Trim(gridcabe.TextMatrix(fila, 1)) <> "" Then
       If Not IsDate(Trim(gridcabe.TextMatrix(fila, 3))) Then
        MsgBox "Verficar Fecha de Vcto no procede", 48, Pub_Titulo
        gridcabe.SetFocus
        Exit Sub
       End If
       wsum = wsum + Val(gridcabe.TextMatrix(fila, 2))
   End If
Next fila

If Val(wsum) <> Val(lblstock.Caption) Then
   MsgBox "Los Importes No Cuadran, Verificar", 48, Pub_Titulo
   Exit Sub
End If


pub_cadena = "DELETE LOTE WHERE LOT_CODCIA = '" & LK_CODCIA & "' AND LOT_CODART = " & i_codart2.Text
CN.Execute pub_cadena, rdExecDirect

For fila = 1 To gridcabe.Rows - 1
   If Trim(gridcabe.TextMatrix(fila, 1)) <> "" Then
       lot_llave.AddNew
       lot_llave!LOT_CODCIA = LK_CODCIA
       lot_llave!LOT_NROLOTE = Trim(gridcabe.TextMatrix(fila, 1))
       lot_llave!lot_codart = Val(gridcabe.TextMatrix(fila, 4))
       lot_llave!lot_fecha_vcto = Trim(gridcabe.TextMatrix(fila, 3))
       lot_llave!LOT_SALDOS = Val(gridcabe.TextMatrix(fila, 2))
       lot_llave!lot_equiv = 1
       lot_llave!lot_fecha_ing = LK_FECHA_DIA
       lot_llave!lot_codclie = 0
       lot_llave!lot_numfac = 0
       lot_llave!lot_numguia = 0
       lot_llave!LOT_CODUSU = Trim(gridcabe.TextMatrix(fila, 7))
       lot_llave.Update
   End If
Next fila
i_codart2_KeyPress 13
MsgBox "Lotes Actualizados", 48, Pub_Titulo
End Sub

Private Sub Form_Load()
CenterMe frmManteLote
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ?  ORDER BY PRE_CODART"
Set PRE_UNIDADES = CN.CreateQuery("", pub_cadena)
PRE_UNIDADES(0) = 0
PRE_UNIDADES(1) = 0
Set preuni_llave = PRE_UNIDADES.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

End Sub

Private Sub gridcabe_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii = 27 Then
cabe
Azul i_codart2, i_codart2

End If
If KeyAscii <> 13 Then Exit Sub
If gridcabe.COL = 7 Then Exit Sub

'If gridcabe.COL = 7 Or gridcabe.COL = 10 Or gridcabe.COL = 13 Or gridcabe.COL = 16 Or gridcabe.COL = 19 Or gridcabe.COL = 22 Then
'Else
'Exit Sub
'End If
    TEXTOVARpre.Left = gridcabe.Left + gridcabe.CellLeft
    TEXTOVARpre.Width = gridcabe.CellWidth
    TEXTOVARpre.Height = gridcabe.CellHeight
    TEXTOVARpre.Top = gridcabe.Top + gridcabe.CellTop
    TEXTOVARpre.Text = gridcabe.TextMatrix(gridcabe.Row, gridcabe.COL)
    TEXTOVARpre.Visible = True
    Azul3 TEXTOVARpre, TEXTOVARpre
    TEXTOVARpre.SetFocus
'End If
End Sub

Private Sub gridcabe_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim sale As Boolean
Dim Wsec
If KeyCode = 46 Then
  If gridcabe.Rows <= 2 Then Exit Sub
  gridcabe.RemoveItem gridcabe.Row
End If
If KeyCode = 45 Then
  gridcabe.Rows = gridcabe.Rows + 1
  gridcabe.TextMatrix(gridcabe.Rows - 1, 4) = i_codart2.Text
End If

If Left(gridcabe.TextMatrix(gridcabe.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
  tC = gridcabe.COL
  For fila = 1 To gridcabe.Cols - 1
      gridcabe.COL = fila
      If gridcabe.CellBackColor = QBColor(12) Then
         gridcabe.CellBackColor = QBColor(15)
         gridcabe.TextMatrix(gridcabe.Row, 9) = "9"
      Else
         gridcabe.CellBackColor = QBColor(12)
         gridcabe.TextMatrix(gridcabe.Row, 9) = "-1"
      End If
  Next fila
  gridcabe.COL = tC
  gridcabe.SetFocus
  Exit Sub
End If
If KeyCode = 45 Then
    Wsec = Wsec + 1
    If Trim(gridcabe.TextMatrix(gridcabe.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(gridcabe.TextMatrix(gridcabe.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(gridcabe.TextMatrix(gridcabe.Row, 4)) = 0 And Val(gridcabe.TextMatrix(gridcabe.Row, 5)) = 0 Then Exit Sub
    gridcabe.AddItem "", gridcabe.Row + 1
    gridcabe.TextMatrix(gridcabe.Row + 1, 0) = "MAN. " & Format(gridcabe.TextMatrix(gridcabe.Row, 10), "dd/mm/yyyy")
    gridcabe.TextMatrix(gridcabe.Row + 1, 6) = Wsec
    gridcabe.TextMatrix(gridcabe.Row + 1, 8) = gridcabe.TextMatrix(gridcabe.Row, 8)
    gridcabe.TextMatrix(gridcabe.Row + 1, 3) = gridcabe.TextMatrix(gridcabe.Row, 3)
    gridcabe.TextMatrix(gridcabe.Row + 1, 7) = gridcabe.TextMatrix(gridcabe.Row, 7)
    gridcabe.TextMatrix(gridcabe.Row + 1, 10) = gridcabe.TextMatrix(gridcabe.Row, 10)
    gridcabe.TextMatrix(gridcabe.Row + 1, 11) = "8"
    gridcabe.Row = gridcabe.Row + 1
    gridcabe.COL = 1
    gridcabe.SetFocus
End If
Exit Sub
If KeyCode = 46 Then
If gridcabe.Rows <= 3 Then
Else
   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(gridcabe.TextMatrix(gridcabe.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     gridcabe.SetFocus
     Exit Sub
   Else
     gridcabe.RowHeight(gridcabe.Row) = 1
     gridcabe.Row = gridcabe.Row + 1
     gridcabe.SetFocus
   End If
End If
End If

Exit Sub



End Sub



Private Sub i_codart2_Change()
If i_codart2.Text = "" Then
 lblsumlote.Caption = ""
 lblstock.Caption = ""
 i_nomarti.Caption = ""
  VAR_ACTIVAR = 0
  
  LBLUNIDAD.Caption = ""
End If

End Sub
Private Sub i_codart2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codart2.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
'  LISTVIEW1.Visible = False
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
'  KeyCode = 0
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  i_codart2.SelStart = Len(i_codart2.Text)
'  If Frame4.Visible = True Then
     If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        SQ_OPER = 3
        pu_alterno = i_codart2.Text
        pu_codcia = LK_CODCIA
        LEER_ART_LLAVE
        If art_llave_alt.EOF Then
           MsgBox "Codigo No Existe ...", 48, Pub_Titulo
           Exit Sub
        End If
        PUB_CODART = art_llave_alt!art_key
     Else
        SQ_OPER = 1

        PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
        pu_codcia = LK_CODCIA
        LEER_ART_LLAVE
        If art_LLAVE.EOF Then
           MsgBox "Codigo No Existe ...", 48, Pub_Titulo
           Exit Sub
        End If
        PUB_CODART = art_LLAVE!art_key
        i_nomarti.Caption = art_LLAVE!ART_NOMBRE
     End If

     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     LEER_ARM_LLAVE
     'If Not arm_llave.EOF Then txtsaldo.Text = arm_llave!arm_stock
     'If Frame4.Visible = False Then Frame4.Visible = True
 ' End If
  
  DoEvents
fin:

End Sub
Private Sub i_codart2_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim i, car
Dim itmFound As MSComctlLib.ListItem
Dim unid_max As Integer
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
cabe
 ListView1.Visible = False
 i_codart2.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
Else
 On Error GoTo mucho
 PUB_KEY = Val(i_codart2.Text)
 On Error GoTo 0
 If Len(i_codart2.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(i_codart2.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = i_codart2.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       GoTo fin
    End If
    WCOD_ORIGINAL = art_LLAVE!art_key
    i_nomarti.Caption = Trim(art_LLAVE!ART_NOMBRE)
    PRE_UNIDADES(0) = LK_CODCIA
    PRE_UNIDADES(1) = art_LLAVE!art_key
    preuni_llave.Requery
    Do Until preuni_llave.EOF
       If preuni_llave!pre_equiv = 1 Then
           LBLUNIDAD.Caption = Trim(preuni_llave!pre_unidad)
       End If
      preuni_llave.MoveNext
    Loop
    ListView1.Visible = False
    Muestra_lote art_LLAVE!art_key
    Exit Sub
Else
  If ListView1.Visible = False And VAR_ACTIVAR <> 99 And i_codart2.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = i_codart2.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       Exit Sub
     End If
     WCOD_ORIGINAL = art_llave_alt!art_key
     i_nomarti.Caption = Trim(art_llave_alt!ART_NOMBRE)
     ListView1.Visible = False
     PRE_UNIDADES(0) = LK_CODCIA
     PRE_UNIDADES(1) = art_llave_alt!art_key
     preuni_llave.Requery
     Do Until preuni_llave.EOF
       If preuni_llave!pre_equiv = 1 Then
          LBLUNIDAD.Caption = Trim(preuni_llave!pre_unidad)
          Exit Do
        End If
        preuni_llave.MoveNext
     Loop
     
     Exit Sub
  Else
    If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
    If Trim(UCase(i_codart2.Text)) = Left(VALOR, Len(Trim(i_codart2.Text))) And Len(Trim(i_codart2.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       PUB_KEY = i_codart2.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul i_codart2, i_codart2
        Exit Sub
      End If
      WCOD_ORIGINAL = art_LLAVE!art_key
      i_nomarti.Caption = Trim(art_LLAVE!ART_NOMBRE)
    PRE_UNIDADES(0) = LK_CODCIA
    PRE_UNIDADES(1) = art_LLAVE!art_key
    preuni_llave.Requery
    Do Until preuni_llave.EOF
       If preuni_llave!pre_equiv = 1 Then
         LBLUNIDAD.Caption = Trim(preuni_llave!pre_unidad)
         Exit Do
       End If
      preuni_llave.MoveNext
    Loop
      ListView1.Visible = False
      ListView1.Visible = False
      i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      ListView1.Visible = False
      Muestra_lote art_LLAVE!art_key
      
      Exit Sub
    Else
      Exit Sub
    End If
    
  End If
End If
dale:
ListView1.Visible = False
fin:
mucho:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul i_codart2, i_codart2
  

End Sub

Private Sub i_codart2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
Dim ws_codcia As String * 2
If KeyCode = 13 Then Exit Sub
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  If Len(i_codart2.Text) = 0 Or i_codart2.Text = "" Then
    ListView1.Visible = False
    Exit Sub
  End If
  If i_codart2.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf i_codart2.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(i_codart2.Text, 1) = "*" Then
   i_codart2.Text = Mid(i_codart2.Text, 2, Len(i_codart2.Text))
   i_codart2.SelStart = Len(i_codart2.Text)
  End If
Else
 If Len(i_codart2.Text) = 0 Or IsNumeric(i_codart2.Text) = True Then
   ListView1.Visible = False
   Exit Sub
 End If
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(i_codart2.Text) = 1 Then
    If i_codart2.Text = "" Then Exit Sub
    var = Asc(i_codart2.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      ws_codcia = "00"
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      If chedes.Value = 1 Then
        archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
      Else
        archi = "SELECT VART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & ws_codcia & "' AND  ART_SITUACION <> 1 AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
      End If
    Else
      numarchi = 0
      'If chedes.Value = 1 Then
       archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      'Else
      ' archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_SITUACION <> 1 AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      'End If
    End If
   ' If Len(I_CODART2.text) > 1 And ListView1.ListItems.count = 0 Then
   ' Else
     PROC_LISVIEW ListView1
   ' End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(i_codart2.Text), lvwText, , lvwPartial)
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


Public Sub cabe()
gridcabe.Cols = 8
gridcabe.Rows = 1
gridcabe.Clear
gridcabe.TextMatrix(0, 0) = "Item"
gridcabe.TextMatrix(0, 1) = "Cod_Lote"
gridcabe.TextMatrix(0, 2) = "Saldo"
gridcabe.TextMatrix(0, 3) = "Fec. Vcto"
gridcabe.TextMatrix(0, 7) = "Usuario"

gridcabe.ColWidth(0) = 800
gridcabe.ColWidth(1) = 1200
gridcabe.ColWidth(2) = 1000
gridcabe.ColWidth(3) = 1200
gridcabe.ColWidth(4) = 0
gridcabe.ColWidth(5) = 0
gridcabe.ColWidth(6) = 0
gridcabe.ColWidth(7) = 1200
End Sub

Public Sub Muestra_lote(w_codart As Currency)
Dim SUMA_SALDOS As Currency
SQ_OPER = 1
PUB_CODART = w_codart
pu_codcia = LK_CODCIA
LEER_ARM_LLAVE
lblstock.Caption = arm_llave!arm_stock
  cabe
  PSLOT_LLAVE2(0) = LK_CODCIA
  PSLOT_LLAVE2(1) = w_codart
  lot_llave2.Requery
  SUMA_SALDOS = 0
  Do Until lot_llave2.EOF
    gridcabe.Rows = gridcabe.Rows + 1
    gridcabe.TextMatrix(gridcabe.Rows - 1, 0) = Format(lot_llave2.AbsolutePosition, "00")
    gridcabe.TextMatrix(gridcabe.Rows - 1, 1) = lot_llave2!LOT_NROLOTE '  "Cod_Lote"
    gridcabe.TextMatrix(gridcabe.Rows - 1, 2) = lot_llave2!LOT_SALDOS
    gridcabe.TextMatrix(gridcabe.Rows - 1, 3) = lot_llave2!lot_fecha_vcto
    gridcabe.TextMatrix(gridcabe.Rows - 1, 4) = w_codart
    gridcabe.TextMatrix(gridcabe.Rows - 1, 5) = lot_llave2!LOT_NROLOTE
    gridcabe.TextMatrix(gridcabe.Rows - 1, 7) = Trim(lot_llave2!LOT_CODUSU)
    
    SUMA_SALDOS = SUMA_SALDOS + Val(lot_llave2!LOT_SALDOS)
    
  lot_llave2.MoveNext
  Loop
  If SUMA_SALDOS <> Val(lblstock.Caption) Then
    MsgBox "Verificar Suma de Lotes no es igual que el Stock", 48, Pub_Titulo
  End If
  sum_lotes
gridcabe.SetFocus

End Sub


Private Sub textovarpre_Change()
gridcabe.Text = TEXTOVARpre.Text
gridcabe.TextMatrix(gridcabe.Row, 7) = LK_CODUSU
sum_lotes
End Sub

Private Sub textovarpre_GotFocus()
 temporal = gridcabe.TextMatrix(gridcabe.Row, gridcabe.COL)
End Sub

Private Sub textovarpre_KeyPress(KeyAscii As Integer)
Dim VALOR As Currency
If KeyAscii = 27 Then
  TEXTOVARpre.Text = temporal
  TEXTOVARpre.Visible = False
  gridcabe.SetFocus
  Exit Sub
End If
'If gridcabe.COL = 1 Then Consistencias gridcabe, TEXTOVARpre, KeyAscii
'If gridcabe.COL = 4 Then Consistencias gridcabe, TEXTOVARpre, KeyAscii
'Consistencias gridcabe, TEXTOVARpre, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If

' AQUI CALCULO

If gridcabe.Row >= gridcabe.Rows - 1 Then
Else
  gridcabe.Row = gridcabe.Row + 1
End If
gridcabe.SetFocus
TEXTOVARpre.Visible = False

fin:

End Sub


Public Sub sum_lotes()
Dim wsum As Currency

For fila = 1 To gridcabe.Rows - 1
   If Trim(gridcabe.TextMatrix(fila, 1)) <> "" Then
       wsum = wsum + Val(gridcabe.TextMatrix(fila, 2))
   End If
Next fila
lblsumlote.Caption = "Suma Lote = " & Format(wsum, "0.000")

End Sub

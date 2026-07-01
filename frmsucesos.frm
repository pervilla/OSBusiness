VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsucesos 
   Caption         =   "Sucesos en Tiempo Real"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView ListView1 
      Height          =   735
      Left            =   3480
      TabIndex        =   13
      Top             =   4920
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
   Begin VB.Frame frmprod 
      Caption         =   "Producto :"
      Height          =   1245
      Left            =   90
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton cmdmostrar 
         Caption         =   "Mostrar"
         Height          =   315
         Left            =   5160
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtcanp 
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox i_codart2 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Lista de Competencia ==>"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Cantidad de Precios a Cotizar:"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label i_nomarti 
         Height          =   495
         Left            =   1260
         TabIndex        =   14
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.ListBox LSTPROD 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.TextBox txtdetalle 
      Height          =   1545
      Left            =   1380
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2790
      Width           =   3765
   End
   Begin VB.ListBox lstgrupos 
      Height          =   1620
      Left            =   1380
      TabIndex        =   0
      Top             =   900
      Width           =   3735
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   135
      Left            =   1410
      TabIndex        =   2
      Top             =   4470
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Esc = Salir "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2040
      TabIndex        =   9
      Top             =   5070
      Width           =   2445
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "F1 = Grabar y Salir "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2010
      TabIndex        =   8
      Top             =   4770
      Width           =   2445
   End
   Begin VB.Label Label3 
      Caption         =   "Usuario :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   225
      Left            =   1380
      TabIndex        =   7
      Top             =   330
      Width           =   885
   End
   Begin VB.Label lbluser 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2310
      TabIndex        =   6
      Top             =   330
      Width           =   2205
   End
   Begin VB.Label Label2 
      Caption         =   "Detalle del Suceso :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   2550
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Grupo de Suceso : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   315
      Left            =   1410
      TabIndex        =   4
      Top             =   690
      Width           =   1695
   End
   Begin VB.Label lblsuc 
      Alignment       =   2  'Center
      Caption         =   "Registro de Sucesos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1380
      TabIndex        =   3
      Top             =   -30
      Width           =   5265
   End
End
Attribute VB_Name = "frmsucesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WCOD_ORIGINAL  As Currency
Dim loc_key  As Integer
Dim rela_con As rdoResultset
Dim loc_new_llave  As rdoResultset
Dim loc_conllave As rdoResultset
Dim PSCONPRECIO As rdoQuery

Dim PSNEW_PRECIO As rdoQuery

Private Sub cmdmostrar_Click()
Dim wpos As Integer
PUB_TIPREG = 68
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "Competencia-  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
frmprod.Visible = False

llena_datos
LSTPROD.SetFocus


End Sub

Private Sub Form_Activate()
If Left(lblsuc.Caption, 2) = "Co" Then
  llena_datos
  Exit Sub
End If
lstgrupos.SetFocus
lbluser.Caption = LK_CODUSU
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Left(lblsuc.Caption, 2) = "Co" Then Exit Sub
If KeyCode = 112 Then
 graba_suceso
 MsgBox "Suceso Enviado. Ok.", 48, Pub_Titulo
 Unload frmsucesos
 Exit Sub
End If
If KeyCode = 27 Then
 Unload frmsucesos
 Exit Sub
End If

End Sub


Public Sub graba_suceso()
Dim ws_numero As Integer
Dim det_suc As rdoResultset

pub_cadena = "SELECT * FROM SUCESOS WHERE SUC_CODCIA = '" & LK_CODCIA & "' AND SUC_CODUSU = '" & LK_CODUSU & "' ORDER BY SUC_CODCIA, SUC_CODUSU, SUC_NUMERO"
Set det_suc = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
det_suc.Requery
If det_suc.EOF Then
  ws_numero = 1
Else
  det_suc.MoveLast
  ws_numero = Val(det_suc!SUC_NUMERO) + 1
End If
det_suc.AddNew
det_suc!SUC_CODCIA = LK_CODCIA
det_suc!SUC_CODUSU = LK_CODUSU
det_suc!SUC_NUMERO = ws_numero
det_suc!SUC_CODSUC = Val(Trim(Right(lstgrupos.Text, 8)))
det_suc!SUC_DETALLE = Trim(txtdetalle.Text)
det_suc!SUC_FECHA = LK_FECHA_DIA
det_suc!SUC_HORA = Format(Now, "hh:mm:ss AMPM")
det_suc.Update
  

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   If Left(lblsuc.Caption, 2) = "Co" Then
     Unload frmsucesos
     Exit Sub
   End If
End If
End Sub

Private Sub Form_Load()
loc_key = 0
CenterMe frmsucesos
SQ_OPER = 2
PUB_TIPREG = 44
PUB_CODCIA = "00"
LEER_TAB_LLAVE
lstgrupos.Clear
Do Until tab_mayor.EOF
  lstgrupos.AddItem tab_mayor!tab_nomlargo & String(80, " ") & tab_mayor!tab_numtab
  tab_mayor.MoveNext
Loop

pub_cadena = "SELECT * FROM HISPRE WHERE HPR_CODCIA = ? AND HPR_FECHA = ? ORDER BY HPR_NUMOPER DESC"
Set PSNEW_PRECIO = CN.CreateQuery("", pub_cadena)
PSNEW_PRECIO(0) = 0
PSNEW_PRECIO(1) = Date
PSNEW_PRECIO.MaxRows = 1
Set loc_new_llave = PSNEW_PRECIO.OpenResultset(rdOpenForwardOnly, rdConcurValues)
  
pub_cadena = "SELECT * FROM CONSULART WHERE CON_CODCIA = ? AND CON_CODART = ? "
Set PSCONPRECIO = CN.CreateQuery("", pub_cadena)
PSCONPRECIO(0) = 0
PSCONPRECIO(1) = 0
Set loc_conllave = PSCONPRECIO.OpenResultset(rdOpenKeyset, rdConcurValues)
  
   
End Sub
Private Sub i_codart2_Change()
If i_codart2.Text = "" Then
  i_nomarti.Caption = ""
  VAR_ACTIVAR = 0
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
 ListView1.Visible = False
 i_codart2.Text = ""
 frmprod.Visible = False
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
    ListView1.Visible = False
    GoTo codigo_colo
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
     'i_codart2.text = Trim(art_llave_alt!ART_NOMBRE)
     
     i_nomarti.Caption = Trim(art_llave_alt!ART_NOMBRE)
     ListView1.Visible = False
     GoTo codigo_colo
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
      ListView1.Visible = False
      GoTo codigo_colo
    End If
  End If
End If
dale:
ListView1.Visible = False
fin:
mucho:
Exit Sub
codigo_colo:
Azul txtcanp, txtcanp
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
      archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
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


Private Sub lstgrupos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtdetalle.SetFocus
End If
End Sub

Private Sub LSTPROD_DblClick()

If Val(Right(Trim(LSTPROD.Text), 8)) = 0 Then
     MsgBox "Seleccione un codigo de Articulo", 48, Pub_Titulo
     Exit Sub
End If
PUB_KEY = Val(Right(Trim(LSTPROD.Text), 8))
SQ_OPER = 2
pu_codcia = LK_CODCIA
PUB_CODART = PUB_KEY
LEER_PRE_LLAVE
If pre_mayor.EOF Then GoTo PASA_PRE
  Load frmprecios
  frmprecios.unidad.Clear
  Do Until pre_mayor.EOF
     frmprecios.unidad.AddItem pre_mayor!pre_unidad & String(80, " ") & pre_mayor!pre_equiv
     pre_mayor.MoveNext
  Loop
  frmprecios.unidad.ListIndex = 0
  frmprecios.lblnombre.Caption = Trim(LSTPROD.Text)
  frmprecios.txtprecio.Text = ""
  frmprecios.txtprove.Text = ""
  frmprecios.Show 1
  loc_secuencia = 0
  If Val(frmprecios.txtprecio.Text) <> 0 Then
    PSNEW_PRECIO(0) = LK_CODCIA
    PSNEW_PRECIO(1) = LK_FECHA_DIA
    loc_new_llave.Requery
    If loc_new_llave.EOF Then
      loc_secuencia = 1
    Else
      loc_secuencia = Val(loc_new_llave!hpr_numoper) + 1
    End If
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = PUB_KEY
    PUB_SECUEN = frmprecios.unidad.ListIndex
    LEER_PRE_LLAVE
    If pre_llave.EOF Then GoTo PASA_PRE
    
    loc_new_llave.AddNew
    loc_new_llave!hpr_codcia = LK_CODCIA
    loc_new_llave!hpr_FECHA = LK_FECHA_DIA
    loc_new_llave!hpr_numoper = loc_secuencia
    loc_new_llave!hpr_codart = PUB_KEY
    loc_new_llave!hpr_TIPMOV = Val(Left(frmprecios.cmdopcion.Text, 1))
    loc_new_llave!hpr_precio = Val(frmprecios.txtprecio.Text)
    loc_new_llave!hpr_empresa = Trim(frmprecios.txtprove.Text)
    loc_new_llave!hpr_codusu = LK_CODUSU
    loc_new_llave!hpr_unidad = Left(frmprecios.unidad.Text, 20)
    loc_new_llave!hpr_EQUIV = Val(Right(frmprecios.unidad.Text, 12))
    loc_new_llave!hpr_pre1 = pre_llave!PRE_PRE1
    loc_new_llave!hpr_pre2 = pre_llave!PRE_PRE2
    loc_new_llave!hpr_pre3 = pre_llave!PRE_PRE3
    loc_new_llave!hpr_pre4 = pre_llave!PRE_PRE4
    loc_new_llave!hpr_pre5 = pre_llave!PRE_PRE5
    loc_new_llave.Update
    
    PSCONPRECIO(0) = LK_CODCIA
    PSCONPRECIO(1) = PUB_KEY
    loc_conllave.Requery
    If loc_conllave.EOF Then
      MsgBox "No se logro Quitar", 48, Pub_Titulo
    Else
      loc_conllave.Delete
    End If
    Unload frmprecios
    llena_datos
  End If
PASA_PRE:
  
End Sub

Private Sub LSTPROD_KeyUp(KeyCode As Integer, Shift As Integer)
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
Else
  Exit Sub
End If
If KeyCode = 45 Then
  frmprod.Visible = True
  i_codart2.Text = ""
  i_nomarti.Caption = ""
  i_codart2.SetFocus
End If
End Sub

Public Sub llena_datos()
Dim WS_FAM As Integer
Dim nomart As String * 60
pub_cadena = "SELECT * FROM CONSULART, ARTI,TABLAS WHERE (ART_CODCIA = TAB_CODCIA) AND (ART_FAMILIA = TAB_NUMTAB) AND TAB_TIPREG = 122 AND (CON_CODCIA = ART_CODCIA) AND (CON_CODART = ART_KEY) AND CON_CODCIA = '" & LK_CODCIA & "' ORDER BY TAB_NOMLARGO , ART_FAMILIA, ART_NOMBRE"
Set rela_con = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
rela_con.Requery
LSTPROD.Clear
If rela_con.EOF Then Exit Sub
WS_FAM = rela_con!art_familia
LSTPROD.AddItem "======= " & rela_con!tab_nomlargo
Do Until rela_con.EOF
   If WS_FAM <> rela_con!art_familia Then
      LSTPROD.AddItem "======= " & rela_con!tab_nomlargo
      WS_FAM = rela_con!art_familia
   End If
   nomart = Trim(rela_con!ART_NOMBRE)
   If Val(rela_con!con_cantidad) >= 50 Then
     LSTPROD.AddItem nomart & "  !!! Urgente !!! " & String(120, " ") & rela_con!con_codart
   Else
     LSTPROD.AddItem nomart & String(120, " ") & rela_con!con_codart
   End If
  rela_con.MoveNext
Loop
End Sub

Private Sub txtcanp_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Val(txtcanp.Text) <= 1 Then
 MsgBox "Como minimo es 1 precio", 48, Pub_Titulo
 Exit Sub
End If


PSCONPRECIO(0) = LK_CODCIA
PSCONPRECIO(1) = WCOD_ORIGINAL
loc_conllave.Requery
If loc_conllave.EOF Then
   loc_conllave.AddNew
   loc_conllave!con_codcia = LK_CODCIA
   loc_conllave!con_codart = art_LLAVE!art_key
   loc_conllave!con_flag = ""
   loc_conllave!con_cantidad = Val(txtcanp.Text)
   loc_conllave.Update
   If LK_CODCIA = "01" Then
   PSCONPRECIO(0) = "03"
   PSCONPRECIO(1) = WCOD_ORIGINAL
   loc_conllave.Requery
   If loc_conllave.EOF Then
    loc_conllave.AddNew
    loc_conllave!con_codcia = "03"
    loc_conllave!con_codart = art_LLAVE!art_key
    loc_conllave!con_flag = ""
    loc_conllave!con_cantidad = Val(txtcanp.Text)
    loc_conllave.Update
   End If
   End If
Else
  MsgBox " Ya Existe en la Lista ", 48, Pub_Titulo
End If
frmprod.Visible = False

llena_datos
LSTPROD.SetFocus


End Sub

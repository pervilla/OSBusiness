VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCal 
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11370
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   5295
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9340
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   8421376
      GridLines       =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox textovar 
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid gr 
      Height          =   3255
      Left            =   9120
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   5741
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   8421376
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox t_vuelto 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox t_total 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblacc 
      Alignment       =   2  'Center
      Caption         =   "Modo Lectura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   0
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   9240
      X2              =   11160
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label lblprocesados 
      Caption         =   "Doc. Procesados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7200
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblpendientes 
      Caption         =   "Doc. Pendientes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Un Momento..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   1680
      TabIndex        =   14
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label lblfecha 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2000"
      Height          =   285
      Left            =   840
      TabIndex        =   13
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   8400
      X2              =   11280
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label importer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9240
      TabIndex        =   10
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lop 
      AutoSize        =   -1  'True
      Caption         =   "Vuelto :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   8280
      TabIndex        =   9
      Top             =   5280
      Width           =   780
   End
   Begin VB.Label lop 
      AutoSize        =   -1  'True
      Caption         =   "Efectivo Total: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   8400
      TabIndex        =   8
      Top             =   4560
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label lop 
      AutoSize        =   -1  'True
      Caption         =   "Recibido :"
      Height          =   195
      Index           =   1
      Left            =   8400
      TabIndex        =   7
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lop 
      AutoSize        =   -1  'True
      Caption         =   "Total Venta :"
      Height          =   195
      Index           =   0
      Left            =   8400
      TabIndex        =   4
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F1 = Salir,      F5 = Actualizar,    F8 = Cobrar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   300
      Left            =   480
      TabIndex        =   2
      Top             =   5760
      Width           =   5385
   End
   Begin VB.Label lbldocu 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8400
      TabIndex        =   1
      Top             =   120
      Width           =   2880
   End
End
Attribute VB_Name = "FrmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag_actualiza As String * 1
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
  Unload FrmCal
End If

If KeyCode = vbKeyF8 And lblacc.Visible = False Then Tecla_F8

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then Tecla_F5
End Sub

Private Sub Form_Load()
  If LK_CODUSU <> "ADMIN" Then
      MsgBox "Solo el usuario ADMIN puede acceder a esta opcion", 48, Pub_Titulo
      Unload FrmCal
      Exit Sub
  End If
  flag_actualiza = ""
  FrmCal.Left = (Screen.Width - FrmCal.Width) / 2
  FrmCal.Top = (Screen.Height - FrmCal.Height) / 2
  gr.Cols = 2
  gr.Rows = 2
  gr.ColWidth(0) = 300
  gr.ColWidth(1) = 1000
  gr.TextMatrix(0, 0) = "+/-"
  gr.TextMatrix(0, 1) = "Importe"
  llena_lista
  lblfecha.Caption = Format(LK_FECHA_DIA, "dd/mm/yyyy")
  

End Sub

Private Sub gr_Click()
TEXTOVAR.Left = gr.Left + gr.CellLeft
TEXTOVAR.Width = gr.CellWidth
TEXTOVAR.Height = gr.CellHeight
TEXTOVAR.Top = gr.Top + gr.CellTop
TEXTOVAR.Tag = gr.TextMatrix(gr.Row, gr.COL)
TEXTOVAR.Text = gr.TextMatrix(gr.Row, gr.COL)
End Sub

Private Sub gr_EnterCell()
TEXTOVAR.Left = gr.Left + gr.CellLeft
TEXTOVAR.Width = gr.CellWidth
TEXTOVAR.Height = gr.CellHeight
TEXTOVAR.Top = gr.Top + gr.CellTop
TEXTOVAR.Tag = gr.TextMatrix(gr.Row, gr.COL)
TEXTOVAR.Text = gr.TextMatrix(gr.Row, gr.COL)

End Sub

Private Sub gr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
  If gr.Rows <= 2 Then
    gr.Text = ""
    Exit Sub
  End If
  gr.RemoveItem gr.Row
End If

End Sub

Private Sub gr_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Exit Sub
If gr.COL = 0 Then
  'textovar.MaxLength = 1
  If KeyAscii = 45 Or KeyAscii = 43 Then
  Else
    TEXTOVAR.Visible = False
    KeyAscii = 0
    Exit Sub
  End If
Else
'textovar.MaxLength = 0
End If
TEXTOVAR.Visible = True
If KeyAscii <> 13 Then TEXTOVAR.Text = Chr(KeyAscii) 'gr.TextMatrix(gr.Row, gr.Col)
TEXTOVAR.SelStart = Len(TEXTOVAR.Text)
TEXTOVAR.SetFocus
If gr.COL = 0 Then
    If KeyAscii = 45 Or KeyAscii = 43 Then
     textovar_KeyPress 13
    Else
      KeyAscii = 0
    End If
End If

End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim F As Integer
Dim fil_r As Integer

If KeyCode <> 96 Then Exit Sub
If grid.Row = 0 Then Exit Sub
fil_r = grid.COL
 For F = 0 To grid.Cols - 1
   grid.Row = grid.Row
   grid.COL = F
   If grid.CellBackColor = QBColor(7) Then
      grid.CellBackColor = QBColor(15)
      grid.TextMatrix(grid.Row, 5) = 0
   Else
      grid.TextMatrix(grid.Row, 5) = 1
      grid.CellBackColor = QBColor(7)
   End If
 Next F
grid.COL = fil_r
suma_select
calcu
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
Dim flag_sel As String
Dim F As Integer
If KeyAscii = 13 Then
 flag_sel = ""
 For F = 0 To grid.Rows - 1
      If Val(grid.TextMatrix(F, 5)) = 1 Then
         flag_sel = "A"
      End If
 Next F
    If flag_sel <> "A" Then
        MsgBox "Seleccione uno de la lista con la tecla 0 (cero)", 48, Pub_Titulo
        Exit Sub
    End If
 
    suma_select
    calcu
    'pub_mensaje = "Imprimir Documento(s) selecionado ?"
    'Pub_Respuesta = MsgBox(pub_mensaje, vbYesNoCancel + vbDefaultButton3 + vbInformation, Pub_Titulo)
    'If Pub_Respuesta = vbCancel Then
    '   Exit Sub
    'End If
    'If Pub_Respuesta = vbYes Then
    '   Imprime_lista_doc
    'End If

    gr.Rows = 2
    gr.Row = gr.Rows - 1
    gr.COL = 0
    gr.SetFocus
End If
End Sub

Private Sub textovar_Change()
gr.TextMatrix(gr.Row, gr.COL) = TEXTOVAR.Text
calcu
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
If gr.COL = 1 Then
   SOLO_DECIMAL TEXTOVAR, KeyAscii
End If
If KeyAscii = 27 Then
gr.TextMatrix(gr.Row, gr.COL) = TEXTOVAR.Tag
TEXTOVAR.Text = TEXTOVAR.Tag
TEXTOVAR.Visible = False
gr.SetFocus
Exit Sub
End If

If KeyAscii = 13 Then
'gr.TextMatrix(gr.Row, gr.Col) = textovar.Text
 If gr.COL = 1 And (gr.Row = gr.Rows - 1) Then
   gr.Rows = gr.Rows + 1
   gr.Row = gr.Row + 1
   gr.COL = 0
   GoTo sa
 ElseIf gr.COL = 1 Then
   gr.Row = gr.Row + 1
   gr.COL = 0
   GoTo sa
 ElseIf gr.COL = 0 Then
   gr.COL = 1
   GoTo sa
 End If
sa:
 TEXTOVAR.Visible = False
 gr.SetFocus
End If

End Sub

Public Sub calcu()
Dim I As Integer
Dim wtotal As Currency
Dim wsi As Integer
wtotal = 0
wsi = 0
For I = 1 To gr.Rows - 1
If gr.TextMatrix(I, 0) = "-" Then
 wsi = -1
Else
 wsi = 1
End If
wtotal = wtotal + (wsi * Val(gr.TextMatrix(I, 1)))
Next I
If (Val(Format(t_total.Text, "0.00")) - wtotal) < 0 Then
 t_vuelto.ForeColor = QBColor(12)
Else
 t_vuelto.ForeColor = QBColor(0)
End If
t_vuelto.Text = Format(Val(Format(t_total.Text, "0.00")) - wtotal, "#,##0.00")
importer.Caption = Format(wtotal, "#,##0.00")
End Sub

Public Sub cabe()
grid.Clear
grid.Rows = 1
grid.Cols = 7
grid.TextMatrix(0, 0) = "/"
grid.TextMatrix(0, 1) = "Serie"
grid.TextMatrix(0, 2) = "Nro."
grid.TextMatrix(0, 3) = "Cliente"
grid.TextMatrix(0, 4) = "Importe"
grid.TextMatrix(0, 6) = "Situación"
grid.RowHeight(0) = 500
grid.ColWidth(0) = 300
grid.ColWidth(1) = 500
grid.ColWidth(2) = 1000
grid.ColWidth(3) = 2400
grid.ColWidth(4) = 1050
grid.ColWidth(5) = 0
grid.ColWidth(6) = 900

End Sub


Public Sub llena_lista()
Dim rs_psvta As rdoResultset
cabe

pub_cadena = "select Distinct far_codcia,far_fbg ,far_numser,far_numfac,far_codclie,far_cliente,far_bruto,far_impto," & _
            " far_numoper from facart where far_codcia = '" & LK_CODCIA & "' and far_tipmov = 10 and far_estado <> 'E' and far_fecha = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "' " & _
            " and far_signo_car = 0 and far_oc <> 'X' order by  far_numoper "
Set rs_psvta = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)

Do Until rs_psvta.EOF
 grid.Rows = grid.Rows + 1
  grid.TextMatrix(grid.Rows - 1, 0) = rs_psvta!far_fbg
  grid.TextMatrix(grid.Rows - 1, 1) = rs_psvta!far_numser
  grid.TextMatrix(grid.Rows - 1, 2) = rs_psvta!far_numfac
  If rs_psvta!far_codclie <> 1 Then
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = rs_psvta!far_codclie
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If cli_llave.EOF Then
    Else
     grid.TextMatrix(grid.Rows - 1, 3) = cli_llave!cli_nombre
    End If
  Else
    grid.TextMatrix(grid.Rows - 1, 3) = rs_psvta!far_cliente
  End If
  grid.TextMatrix(grid.Rows - 1, 4) = Format(Val(rs_psvta!far_bruto) + Val(rs_psvta!far_impto), "##,##0.00")
rs_psvta.MoveNext
Loop
flag_actualiza = ""
End Sub

Public Sub suma_select()
Dim sum_efectivo As Currency
Dim F As Integer
Dim fil_r As Integer

fil_r = grid.COL
sum_efectivo = 0
 For F = 0 To grid.Rows - 1
      If Val(grid.TextMatrix(F, 5)) = 1 Then
         sum_efectivo = sum_efectivo + Val(Format(grid.TextMatrix(F, 4), "0.00"))
      End If
 Next F
lbldocu.Caption = "Total Seleccion de Documentos"
t_total.Text = Format(sum_efectivo, "0.00")
End Sub

Public Sub Tecla_F8()
Dim F As Integer

grid.Visible = False
DoEvents
 For F = 0 To grid.Rows - 1
     If Val(grid.TextMatrix(F, 5)) = 1 Then
         pub_cadena = "update facart set far_oc = 'X' where far_codcia= '" & LK_CODCIA & "' and far_fbg = '" & Trim(grid.TextMatrix(F, 0)) & "' and far_numser = '" & grid.TextMatrix(F, 1) & "' and far_numfac= " & Trim(grid.TextMatrix(F, 2))
         CN.Execute pub_cadena, rdExecDirect
     End If
 Next F
 grid.Visible = True
 gr.Rows = 2
 gr.TextMatrix(1, 0) = ""
 gr.TextMatrix(1, 1) = ""
 
Tecla_F5
End Sub
Public Sub Tecla_F5()
Dim fil_r As Integer

If flag_actualiza <> "" Then
   flag_actualiza = "A"
Else
   Exit Sub
End If
grid.Visible = False
DoEvents
llena_lista
grid.Visible = True
If grid.Rows <= 1 Then
Else
 grid.Row = 1
 grid.COL = 1
End If
suma_select
calcu

grid.SetFocus

flag_actualiza = ""

End Sub


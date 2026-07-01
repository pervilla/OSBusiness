VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmIncentivos 
   Caption         =   "Opciones de Inventivos."
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "Consultar "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   48
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox cmdsigno 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmIncentivos.frx":0000
      Left            =   10560
      List            =   "FrmIncentivos.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCerrar 
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
      Left            =   4560
      Picture         =   "FrmIncentivos.frx":002B
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6000
      Width           =   1300
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   375
      Left            =   6240
      TabIndex        =   37
      Top             =   5160
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
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
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Grabar"
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
      Left            =   480
      Picture         =   "FrmIncentivos.frx":0175
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      Width           =   1300
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle de Incentivos por Tipo"
      Height          =   3315
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   11025
      Begin VB.CommandButton cmdline 
         BackColor       =   &H00C0C000&
         Caption         =   "Insertar salto de linea"
         Height          =   255
         Left            =   3120
         MaskColor       =   &H00808000&
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox cmblinea 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   1875
      End
      Begin VB.ComboBox unidad 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   9840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   1065
      End
      Begin VB.OptionButton optipo 
         Caption         =   "Para Operadores."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   840
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.OptionButton optipo 
         Caption         =   "Para Clientes."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   600
         Width           =   1515
      End
      Begin VB.CheckBox cheliq 
         Caption         =   "Anexo Liquidado."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   10200
         TabIndex        =   44
         Top             =   1440
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtglosa 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   1245
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1920
         Width           =   10545
      End
      Begin VB.ComboBox cmbtipo 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2235
      End
      Begin VB.TextBox txtcantidad 
         Height          =   285
         Left            =   10080
         TabIndex        =   29
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtprecio 
         Height          =   285
         Left            =   10200
         TabIndex        =   28
         Top             =   1080
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox i_codart2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   6510
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txttotal 
         Height          =   285
         Left            =   9960
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfechavcto 
         Height          =   285
         Left            =   4320
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfechaini 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lblcondi 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6000
         TabIndex        =   49
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Linea."
         Height          =   255
         Index           =   15
         Left            =   4680
         TabIndex        =   46
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Fecha Incial :"
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   45
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Fecha Vcto :"
         Height          =   255
         Index           =   9
         Left            =   3120
         TabIndex        =   43
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Fecha Emision:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label i_nomarti 
         Height          =   375
         Left            =   7440
         TabIndex        =   38
         Top             =   780
         Width           =   2265
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Total S/."
         Height          =   255
         Index           =   4
         Left            =   9960
         TabIndex        =   36
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Tipo de Incentivo"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   35
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Producto:"
         Height          =   255
         Index           =   1
         Left            =   6570
         TabIndex        =   34
         Top             =   630
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Cantidad"
         Height          =   255
         Index           =   2
         Left            =   9960
         TabIndex        =   33
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Precio"
         Height          =   255
         Index           =   3
         Left            =   10560
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Unidad"
         Height          =   255
         Index           =   5
         Left            =   9840
         TabIndex        =   31
         Top             =   630
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Descripción del Incentivo "
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Width           =   3225
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Relacion de Documento:"
      Height          =   645
      Left            =   30
      TabIndex        =   14
      Top             =   -30
      Width           =   9345
      Begin VB.Label lblcontrol 
         Caption         =   "Proveedor :"
         Height          =   255
         Index           =   10
         Left            =   90
         TabIndex        =   25
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Fecha : "
         Height          =   255
         Index           =   11
         Left            =   90
         TabIndex        =   24
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Factura : "
         Height          =   255
         Index           =   12
         Left            =   2760
         TabIndex        =   23
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label d_codigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   22
         Top             =   270
         Width           =   765
      End
      Begin VB.Label d_nombre 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2070
         TabIndex        =   21
         Top             =   270
         Width           =   5955
      End
      Begin VB.Label d_fecha 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1170
         TabIndex        =   20
         Top             =   690
         Width           =   1125
      End
      Begin VB.Label d_serie 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3840
         TabIndex        =   19
         Top             =   690
         Width           =   435
      End
      Begin VB.Label d_numero 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4410
         TabIndex        =   18
         Top             =   690
         Width           =   855
      End
      Begin VB.Label lblcontrol 
         Caption         =   "Interno : "
         Height          =   255
         Index           =   13
         Left            =   5640
         TabIndex        =   17
         Top             =   750
         Width           =   855
      End
      Begin VB.Label d_numser 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6570
         TabIndex        =   16
         Top             =   690
         Width           =   435
      End
      Begin VB.Label d_numfac 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7140
         TabIndex        =   15
         Top             =   690
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdidet 
      Height          =   2325
      Left            =   30
      TabIndex        =   10
      Top             =   3600
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   4101
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label flag_signo 
      Caption         =   "1"
      Height          =   225
      Left            =   3180
      TabIndex        =   41
      Top             =   1110
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Opción :"
      Height          =   225
      Left            =   4380
      TabIndex        =   40
      Top             =   750
      Width           =   795
   End
   Begin VB.Label lblcontrol 
      Caption         =   "Detalle de Incentivos :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   3015
   End
End
Attribute VB_Name = "FrmIncentivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim PRE_RELCOMP As rdoQuery
Dim relcomp_llave  As rdoResultset
Dim loc_key As Integer
Dim PRE_UNIDADES As rdoQuery
Dim preuni_llave  As rdoResultset

Private Sub cheliq_Click()
txtglosa.SetFocus
End Sub

Private Sub cmblinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   i_codart2.SetFocus
End If

End Sub

Private Sub cmbtipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cmblinea.SetFocus
End If
End Sub

Private Sub cmbtipo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = cmbtipo.ListIndex
PUB_TIPREG = Mid(cmbtipo.ToolTipText, 13, Len(cmbtipo.ToolTipText))
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
LLENADOS cmbtipo, 49
cmbtipo.SetFocus
SendKeys "%{up}"


End Sub

Private Sub cmdAgregar_Click()
Dim ws_numero As Integer
PRE_RELCOMP(0) = LK_CODCIA
PRE_RELCOMP(1) = txtfecha.Text
relcomp_llave.Requery
ws_numero = 0
If Not relcomp_llave.EOF Then
 'pub_mensaje = "Reemplazar la infomación...   ¿Desea Continuar... ?"
 'Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
 'If Pub_Respuesta = vbNo Then
 '   Exit Sub
 'End If
 relcomp_llave.MoveLast
 ws_numero = Val(relcomp_llave!min_numero)
 If lblcondi.Caption = "NUEVO" Then
 Else
  pub_cadena = "DELETE FROM MAEINCENTIVO WHERE MIN_CODCIA = '" & LK_CODCIA & "'  AND MIN_FECHA = '" & Format(txtfecha.Text, "dd/mm/yyyy") & "'"
  CN.Execute pub_cadena, rdExecDirect
 End If
End If
Screen.MousePointer = 13
For fila = 1 To grdidet.Rows - 1
   ws_numero = ws_numero + 1
    relcomp_llave.AddNew
    relcomp_llave!min_codcia = LK_CODCIA
    relcomp_llave!min_FECHA = Trim(txtfecha.Text)
    relcomp_llave!min_numero = ws_numero
    relcomp_llave!min_codtip = Val(Right(grdidet.TextMatrix(fila, 1), 8))
    relcomp_llave!min_codart = Val(grdidet.TextMatrix(fila, 2))
    relcomp_llave!min_descri = Trim(Left(grdidet.TextMatrix(fila, 4), 20))
    relcomp_llave!min_equiv = Val(Trim(Right(grdidet.TextMatrix(fila, 4), 8)))
    relcomp_llave!min_total = Val(grdidet.TextMatrix(fila, 8))
    relcomp_llave!min_fecha_ini = (grdidet.TextMatrix(fila, 5))
    relcomp_llave!min_fecha_fin = (grdidet.TextMatrix(fila, 6))
    relcomp_llave!min_GLOSA = Trim((grdidet.TextMatrix(fila, 7)))
    relcomp_llave!min_tipo = Trim((grdidet.TextMatrix(fila, 12)))
    relcomp_llave!min_linea = Val(Trim(Right(grdidet.TextMatrix(fila, 8), 8)))
    relcomp_llave!min_estado = Val(Trim((grdidet.TextMatrix(fila, 13))))
    relcomp_llave.Update
Next fila
lblcondi.Caption = ""
Screen.MousePointer = 0
MsgBox "Datos Registrados", 48, Pub_Titulo
Unload FrmIncentivos
End Sub

Private Sub cmdCerrar_Click()
Unload FrmIncentivos
End Sub


Private Sub CmdConsultar_Click()
Frame2.Enabled = True
lblcondi.Caption = "MODIFICAR"
PRE_RELCOMP(0) = LK_CODCIA
PRE_RELCOMP(1) = txtfecha.Text
relcomp_llave.Requery
If relcomp_llave.EOF Then
  MsgBox "No Existe datos ", 48, Pub_Titulo
  Exit Sub
End If
If cmbtipo.ListCount > 0 Then cmbtipo.ListIndex = 0
If cmblinea.ListCount > 0 Then cmblinea.ListIndex = 0


cabe
'cheliq.Value = relcomp_llave!rel_LIQUIDO
Do Until relcomp_llave.EOF
    grdidet.Rows = grdidet.Rows + 1
    PUB_TIPREG = 49
    PUB_NUMTAB = relcomp_llave!min_codtip
    SQ_OPER = 1
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
       grdidet.TextMatrix(grdidet.Rows - 1, 1) = tab_llave!tab_nomlargo & String(80, " ") & tab_llave!tab_numtab
    End If
    grdidet.TextMatrix(grdidet.Rows - 1, 2) = relcomp_llave!min_codart
    grdidet.TextMatrix(grdidet.Rows - 1, 0) = Format(relcomp_llave!min_FECHA, "dd/mm/yy")
    If Val(relcomp_llave!min_codart) <> 0 Then
     SQ_OPER = 1
     PUB_KEY = relcomp_llave!min_codart
     pu_codcia = relcomp_llave!min_codcia
     LEER_ART_LLAVE
     
     grdidet.TextMatrix(grdidet.Rows - 1, 3) = Trim(art_LLAVE!ART_NOMBRE)
     grdidet.TextMatrix(grdidet.Rows - 1, 4) = relcomp_llave!min_descri & String(80, " ") & relcomp_llave!min_equiv
     End If
     grdidet.TextMatrix(grdidet.Rows - 1, 5) = Format(relcomp_llave!min_fecha_ini, "dd/mm/yyyy")
     grdidet.TextMatrix(grdidet.Rows - 1, 6) = Format(relcomp_llave!min_fecha_fin, "dd/mm/yyyy")
     grdidet.TextMatrix(grdidet.Rows - 1, 7) = Trim(relcomp_llave!min_GLOSA)
     grdidet.TextMatrix(grdidet.Rows - 1, 8) = relcomp_llave!min_linea
     grdidet.TextMatrix(grdidet.Rows - 1, 11) = relcomp_llave!min_FECHA
     grdidet.TextMatrix(grdidet.Rows - 1, 12) = relcomp_llave!min_tipo
     grdidet.TextMatrix(grdidet.Rows - 1, 13) = Val(relcomp_llave!min_estado)

 relcomp_llave.MoveNext
Loop
If Val(flag_signo.Caption) = 1 Then
   cmdsigno.ListIndex = 0
ElseIf Val(flag_signo.Caption) = -1 Then
   cmdsigno.ListIndex = 1
End If

Exit Sub

End Sub

Private Sub cmdline_Click()
txtglosa.Text = txtglosa.Text + Chr(13)
txtglosa.SelStart = Len(Trim(txtglosa.Text))
txtglosa.SetFocus
End Sub

Private Sub cmdnuevo_Click()
Frame2.Enabled = True
lblcondi.Caption = "NUEVO"

End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0

'If fraanexos.Visible = True Then
' GoTo anexo_act
'End If

'PRE_RELCOMP(0) = LK_CODCIA
'PRE_RELCOMP(1) = "P"
'PRE_RELCOMP(2) = Val(d_codigo.Caption)
'PRE_RELCOMP(3) = Val(d_numser.Caption)
'PRE_RELCOMP(4) = Val(d_numfac.Caption)
'relcomp_llave.Requery
'If relcomp_llave.EOF Then
'If Val(flag_signo.Caption) = 1 Then
'   cmdsigno.ListIndex = 0
'ElseIf Val(flag_signo.Caption) = -1 Then
'   cmdsigno.ListIndex = 1
'End If
If cmbtipo.ListCount > 0 Then cmbtipo.ListIndex = 0
If cmblinea.ListCount > 0 Then cmblinea.ListIndex = 0

Exit Sub
'End If


cabe
cheliq.Value = relcomp_llave!rel_LIQUIDO
Do Until relcomp_llave.EOF
    grdidet.Rows = grdidet.Rows + 1
    PUB_TIPREG = 43
    PUB_NUMTAB = relcomp_llave!rel_codtip
    SQ_OPER = 1
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
       grdidet.TextMatrix(grdidet.Rows - 1, 1) = tab_llave!tab_nomlargo & String(80, " ") & tab_llave!tab_numtab
    End If
    grdidet.TextMatrix(grdidet.Rows - 1, 2) = relcomp_llave!rel_codart
    grdidet.TextMatrix(grdidet.Rows - 1, 0) = Format(relcomp_llave!rel_FECHA, "dd/mm/yy")
    If Val(relcomp_llave!rel_codart) <> 0 Then
     SQ_OPER = 1
     PUB_KEY = relcomp_llave!rel_codart
     pu_codcia = relcomp_llave!rel_codcia
     LEER_ART_LLAVE
     
     grdidet.TextMatrix(grdidet.Rows - 1, 3) = Trim(art_LLAVE!ART_NOMBRE)
     grdidet.TextMatrix(grdidet.Rows - 1, 4) = relcomp_llave!rel_descri & String(80, " ") & relcomp_llave!rel_equiv
     grdidet.TextMatrix(grdidet.Rows - 1, 5) = Format(relcomp_llave!rel_cantidad / relcomp_llave!rel_equiv, "0.00")
     grdidet.TextMatrix(grdidet.Rows - 1, 6) = relcomp_llave!rel_precio
    End If
    grdidet.TextMatrix(grdidet.Rows - 1, 7) = Trim(relcomp_llave!rel_GLOSA)
    grdidet.TextMatrix(grdidet.Rows - 1, 8) = relcomp_llave!rel_total
    grdidet.TextMatrix(grdidet.Rows - 1, 9) = relcomp_llave!rel_signo_carga
    grdidet.TextMatrix(grdidet.Rows - 1, 10) = "A"
    grdidet.TextMatrix(grdidet.Rows - 1, 11) = relcomp_llave!rel_FECHA
    grdidet.TextMatrix(grdidet.Rows - 1, 12) = relcomp_llave!rel_LIQUIDO
    If Val(relcomp_llave!rel_signo_carga) = -1 Then
      grdidet.Row = grdidet.Rows - 1
      grdidet.COL = 1
      grdidet.CellBackColor = QBColor(14)
    End If
 relcomp_llave.MoveNext
Loop
If Val(flag_signo.Caption) = 1 Then
   cmdsigno.ListIndex = 0
ElseIf Val(flag_signo.Caption) = -1 Then
   cmdsigno.ListIndex = 1
End If

Exit Sub
anexo_act:

pub_cadena = "SELECT * FROM MAEINCENTIVO WHERE MIN_CODCIA = ? MIN_FECHA = ? ORDER BY MIN_NUMFAC "
Set PRE_RELCOMP = CN.CreateQuery("", pub_cadena)
PRE_RELCOMP(0) = 0
PRE_RELCOMP(1) = 0
Set relcomp_llave = PRE_RELCOMP.OpenResultset(rdOpenKeyset, rdConcurValues)
PRE_RELCOMP(0) = LK_CODCIA
PRE_RELCOMP(1) = txtfecha.Text
relcomp_llave.Requery

CABE_MUESTRA
Do Until relcomp_llave.EOF
    GRID_MST.Rows = GRID_MST.Rows + 1
    PUB_TIPREG = 43
    PUB_NUMTAB = relcomp_llave!rel_codtip
    SQ_OPER = 1
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
       GRID_MST.TextMatrix(GRID_MST.Rows - 1, 0) = tab_llave!tab_nomlargo & String(80, " ") & tab_llave!tab_numtab
    End If
    GRID_MST.TextMatrix(GRID_MST.Rows - 1, 1) = relcomp_llave!rel_codart
    If Val(relcomp_llave!rel_codart) <> 0 Then
     SQ_OPER = 1
     PUB_KEY = relcomp_llave!rel_codart
     pu_codcia = relcomp_llave!rel_codcia
     LEER_ART_LLAVE
     GRID_MST.TextMatrix(GRID_MST.Rows - 1, 2) = Trim(art_LLAVE!ART_NOMBRE)
     GRID_MST.TextMatrix(GRID_MST.Rows - 1, 3) = relcomp_llave!rel_descri & String(80, " ") & relcomp_llave!rel_equiv
     GRID_MST.TextMatrix(GRID_MST.Rows - 1, 4) = Format(relcomp_llave!rel_cantidad / relcomp_llave!rel_equiv, "0.00")
     GRID_MST.TextMatrix(GRID_MST.Rows - 1, 5) = relcomp_llave!rel_precio
    End If
    GRID_MST.TextMatrix(GRID_MST.Rows - 1, 6) = Trim(relcomp_llave!rel_GLOSA)
    GRID_MST.TextMatrix(GRID_MST.Rows - 1, 7) = relcomp_llave!rel_total

 relcomp_llave.MoveNext
Loop
Exit Sub



End Sub

Private Sub Form_Load()
CenterMe FrmIncentivos
loc_key = 0
cabe
PUB_CODCIA = LK_CODCIA
LLENADOS cmbtipo, 49
LLENADOS cmblinea, 122
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ?  ORDER BY PRE_CODART"
Set PRE_UNIDADES = CN.CreateQuery("", pub_cadena)
PRE_UNIDADES(0) = 0
PRE_UNIDADES(1) = 0
Set preuni_llave = PRE_UNIDADES.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM MAEINCENTIVO WHERE MIN_CODCIA = ? AND MIN_FECHA = ? ORDER BY MIN_NUMERO "
Set PRE_RELCOMP = CN.CreateQuery("", pub_cadena)
PRE_RELCOMP(0) = 0
PRE_RELCOMP(1) = 0
Set relcomp_llave = PRE_RELCOMP.OpenResultset(rdOpenKeyset, rdConcurValues)
txtfecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtfecha.Mask = "##/##/####"
txtFechaVcto.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtFechaVcto.Mask = "##/##/####"
txtfechaini.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtfechaini.Mask = "##/##/####"
lblcondi.Caption = ""
lblcondi.Caption = "NUEVO"
End Sub


Private Sub grdidet_DblClick()
If grdidet.COL = 13 Then
    If Val(grdidet.TextMatrix(grdidet.Row, grdidet.COL)) = 1 Then
      pub_mensaje = "Desea Activar Nuevamente el Incentivo"
    Else
      pub_mensaje = "Desea DesActivar el Incentivo"
    End If
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
    If Val(grdidet.TextMatrix(grdidet.Row, grdidet.COL)) = 1 Then
    grdidet.TextMatrix(grdidet.Row, grdidet.COL) = 0
    Else
    grdidet.TextMatrix(grdidet.Row, grdidet.COL) = 1
    End If
End If
End Sub

Private Sub grdidet_KeyUp(KeyCode As Integer, Shift As Integer)

If grdidet.TextMatrix(grdidet.Row, 10) = "A" Then Exit Sub
If KeyCode = 46 Then
    LK_ACCESO_REPORT = ""
    Load frmclave2
    Screen.MousePointer = 0
    frmclave2.Show 1
    If LK_ACCESO_REPORT <> "A" Then
        Exit Sub
    End If

  If grdidet.Rows = 2 Then
    grdidet.Rows = 1
  ElseIf grdidet.Rows > 2 Then
    grdidet.RemoveItem grdidet.Row
  End If
End If
End Sub

Private Sub i_codart2_Change()
If i_codart2.Text = "" Then
  i_nomarti.Caption = ""
  VAR_ACTIVAR = 0
End If

End Sub

Private Sub i_codart2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
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
Dim itmFound As ListItem
Dim unid_max As Integer
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView1.Visible = False
 i_codart2.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
If i_codart2.Text = "" Then
  Azul2 txtfechaini, txtfechaini
  Exit Sub
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
    'i_codart2.text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
    PRE_UNIDADES(0) = LK_CODCIA
    PRE_UNIDADES(1) = art_LLAVE!art_key
    preuni_llave.Requery
    unidad.Clear
    unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!pre_unidad & String(80, " ") & preuni_llave!pre_equiv
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
         unid_max = preuni_llave.AbsolutePosition - 1
'         LBLUNIDAD.Caption = Trim(preuni_llave!pre_UNIDAD)
      '   Exit Do
       End If
      preuni_llave.MoveNext
    Loop
    ListView1.Visible = False
    If unidad.ListCount > 0 Then unidad.ListIndex = unid_max
    unidad.SetFocus
    'If txtCampo1.Visible Then
    ' Azul2 txtCampo1, txtCampo1
    ' Exit Sub
    'End If
    'CmdProcesa.SetFocus
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
     If art_llave_alt!ART_MONEDA = "D" Then
      lblMoneda.Caption = "Moneda : US$."
     Else
      lblMoneda.Caption = "Moneda : S/ ."
     End If
     i_nomarti.Caption = Trim(art_llave_alt!ART_NOMBRE)
     ListView1.Visible = False
     
     PRE_UNIDADES(0) = LK_CODCIA
     PRE_UNIDADES(1) = art_llave_alt!art_key
     preuni_llave.Requery
     unidad.Clear
     unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!pre_unidad & String(80, " ") & preuni_llave!pre_equiv
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
          unid_max = preuni_llave.AbsolutePosition - 1
          LBLUNIDAD.Caption = Trim(preuni_llave!pre_unidad)
          Exit Do
        End If
        preuni_llave.MoveNext
     Loop
     If unidad.ListCount > 0 Then unidad.ListIndex = unid_max
     unidad.SetFocus
     'If Frame2.Visible And i_maximo.Visible Then i_maximo.SetFocus
     'If Frame1.Visible And txtCampo1.Visible Then Azul2 txtCampo1, txtCampo1
     
     'CmdProcesa.SetFocus
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
    unidad.Clear
    unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!pre_unidad & String(80, " ") & preuni_llave!pre_equiv
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
         unid_max = preuni_llave.AbsolutePosition - 1
'         LBLUNIDAD.Caption = Trim(preuni_llave!pre_UNIDAD)
      '   Exit Do
       End If
      preuni_llave.MoveNext
    Loop
    ListView1.Visible = False
    If unidad.ListCount > 0 Then unidad.ListIndex = unid_max
     unidad.SetFocus
      ListView1.Visible = False
      If unidad.ListCount > 0 Then unidad.ListIndex = unid_max
      i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      ListView1.Visible = False
      'If Frame2.Visible And i_maximo.Visible Then i_maximo.SetFocus
      'If Frame1.Visible And txtCampo1.Visible Then Azul2 txtCampo1, txtCampo1
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
      archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      If Val(Trim(Right(cmblinea.Text, 8))) = 0 Then
        archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      Else
        archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_FAMILIA = " & Trim(Right(cmblinea.Text, 8)) & " AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      End If
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
Dim itmFound As ListItem    ' Variable FoundItem.
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

Public Sub LLENADOS(cont As ComboBox, tip As Integer)
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    If tip = 122 Then
      cont.AddItem "..."
    End If
    Do Until tab_mayor.EOF
       If PUB_TIPREG = 35 And LK_EMP = "PAR" Then
          If Val(tab_mayor!tab_codart) = WW_CODVEN Then
            cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
          End If
       Else
           cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!tab_numtab
       End If
       tab_mayor.MoveNext
    Loop
End Sub

Private Sub txtcantidad_Change()
cal
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul txtprecio, txtprecio
End If
End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbtipo.SetFocus
End If
End Sub

Private Sub txtfechaini_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul2 txtFechaVcto, txtFechaVcto
End If

End Sub

Private Sub txtfechavcto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtglosa.SetFocus
End If
End Sub

Private Sub txtglosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
add_registro
End If
End Sub

Private Sub txtprecio_Change()
cal
End Sub

Private Sub txtprecio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtglosa.SetFocus
End If
End Sub

Private Sub txttotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
add_registro
End If
End Sub

Private Sub unidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul2 txtfechaini, txtfechaini
End If
End Sub
Public Sub cabe()
grdidet.Clear
grdidet.Cols = 14
grdidet.Rows = 1
grdidet.ColWidth(0) = 900 ' Fecha
grdidet.ColWidth(1) = 2000 ' tipo
grdidet.ColWidth(2) = 600 ' codigo internot
grdidet.ColWidth(3) = 1300  ' descrip prod
grdidet.ColWidth(4) = 800  ' unidad
grdidet.ColWidth(5) = 1000 ' cantidad
grdidet.ColWidth(6) = 1000 ' precio
grdidet.ColWidth(7) = 2800 ' glosa
grdidet.ColWidth(8) = 900 ' total
grdidet.ColWidth(9) = 0 ' Signo de cargo
grdidet.ColWidth(10) = 0 ' No Editar
grdidet.ColWidth(11) = 0 ' FECHA DE PROCESO
grdidet.ColWidth(12) = 900 ' FLAG DE CIERRE

grdidet.TextMatrix(0, 0) = "Fecha"
grdidet.TextMatrix(0, 1) = "Tipo"
grdidet.TextMatrix(0, 2) = "Codigo"
grdidet.TextMatrix(0, 3) = "Descripción"
grdidet.TextMatrix(0, 4) = "Unidad"
grdidet.TextMatrix(0, 5) = "Fec.Ini."
grdidet.TextMatrix(0, 6) = "Fec.Fin."
grdidet.TextMatrix(0, 7) = "Glosa"
grdidet.TextMatrix(0, 8) = "Linea"
grdidet.TextMatrix(0, 12) = "Tipo"
grdidet.TextMatrix(0, 13) = "(1-0)Act-DesAc."


End Sub

Public Sub CABE_MUESTRA()

GRID_MST.Clear
GRID_MST.Cols = 8
GRID_MST.Rows = 1
GRID_MST.ColWidth(0) = 1000 ' tipo
GRID_MST.ColWidth(1) = 900 ' codigo internot
GRID_MST.ColWidth(2) = 1300  ' descrip prod
GRID_MST.ColWidth(3) = 900  ' unidad
GRID_MST.ColWidth(4) = 900 ' cantidad
GRID_MST.ColWidth(5) = 900 ' precio
GRID_MST.ColWidth(6) = 2500 ' glosa
GRID_MST.ColWidth(7) = 900 ' total


GRID_MST.TextMatrix(0, 0) = "Tipo"
GRID_MST.TextMatrix(0, 1) = "Codigo"
GRID_MST.TextMatrix(0, 2) = "Descripción"
GRID_MST.TextMatrix(0, 3) = "Unidad"
GRID_MST.TextMatrix(0, 4) = "Cantidad"
GRID_MST.TextMatrix(0, 5) = "Precio"
GRID_MST.TextMatrix(0, 6) = "Glosa"
GRID_MST.TextMatrix(0, 7) = "Total"


End Sub

Public Sub add_registro()
If Trim(cmbtipo.Text) = "" Then
  MsgBox "Seleccionar Tipo de Documento", 48, Pub_Titulo
  cmbtipo.SetFocus
  Exit Sub
End If


grdidet.Rows = grdidet.Rows + 1
grdidet.TextMatrix(grdidet.Rows - 1, 0) = Format(LK_FECHA_DIA, "dd/mm/yy")
grdidet.TextMatrix(grdidet.Rows - 1, 1) = Trim(cmbtipo.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 2) = Trim(i_codart2.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 3) = Trim(i_nomarti.Caption)
grdidet.TextMatrix(grdidet.Rows - 1, 4) = Trim(unidad.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 5) = txtfechaini.Text
grdidet.TextMatrix(grdidet.Rows - 1, 6) = txtFechaVcto.Text
grdidet.TextMatrix(grdidet.Rows - 1, 7) = Trim(txtglosa.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 8) = Trim(cmblinea.Text)
grdidet.TextMatrix(grdidet.Rows - 1, 9) = Trim(Left(cmdsigno.Text, 2))
grdidet.TextMatrix(grdidet.Rows - 1, 11) = LK_FECHA_DIA
If optipo(0).Value = True Then
 grdidet.TextMatrix(grdidet.Rows - 1, 12) = 0 ' CLIENTE
Else
 grdidet.TextMatrix(grdidet.Rows - 1, 12) = 1 ' PRODUCTO
End If


i_codart2.Text = ""
i_nomarti.Caption = ""
unidad.Clear
txtcantidad.Text = ""
txtprecio.Text = ""
txtglosa.Text = ""
txttotal.Text = ""
cmbtipo.SetFocus
End Sub

Public Sub cal()
 txttotal.Text = Format(Val(txtprecio.Text) * Val(txtcantidad.Text), "0.00")
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEquivArti 
   Caption         =   "Actualizar Datos de Productos por Linea"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProBar 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   10695
      Begin VB.ComboBox cmbfami 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   4935
      End
      Begin VB.CommandButton cmdmostrar 
         Caption         =   "Mostrar Lista.."
         Height          =   375
         Left            =   5640
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "La Cantidades  se Actualiza automaticamente "
         Height          =   495
         Left            =   8280
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblra 
         Caption         =   "Selecione Lines :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle de Productos :"
      Height          =   4935
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   10695
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
         MultiLine       =   0   'False
         TextRTF         =   $"frmequivarti.frx":0000
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
      Begin MSFlexGridLib.MSFlexGrid gridarti 
         Height          =   4605
         Left            =   120
         TabIndex        =   2
         Tag             =   "9999"
         Top             =   240
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   8123
         _Version        =   393216
         Rows            =   3
         Cols            =   3
         FixedRows       =   2
         FixedCols       =   2
         BackColorBkg    =   8421376
         FocusRect       =   2
         HighLight       =   2
         AllowUserResizing=   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Ingresos de Partes de Unidades Minimas."
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmEquivArti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_ev0 As Currency
Dim loc_ev1 As Currency
Dim loc_ev2 As Currency

Private Sub cmbcal_Click()
If ProBar.Visible = False Then sum_total
End Sub

Private Sub cmbfami_Click()
gridarti.Clear
gridarti.Cols = 13
gridarti.TextMatrix(0, 0) = "Item"
gridarti.TextMatrix(0, 1) = "Descripcion del Producto"
gridarti.TextMatrix(0, 2) = "Unid."
gridarti.TextMatrix(0, 2) = "Activa"
gridarti.TextMatrix(0, 3) = "Equiv" ' SISTEMA GUARDADO
gridarti.TextMatrix(1, 3) = "Activa." ' SISTEMA GUARDADO
gridarti.TextMatrix(0, 4) = "Cantidad"
gridarti.TextMatrix(1, 4) = "p'Unidad"
gridarti.ColWidth(0) = 300
gridarti.ColWidth(1) = 3500
gridarti.ColWidth(2) = 600
gridarti.ColWidth(3) = 800
gridarti.ColWidth(4) = 800
gridarti.ColWidth(5) = 0
gridarti.ColWidth(6) = 0
gridarti.ColWidth(7) = 0
gridarti.ColWidth(8) = 0
gridarti.ColWidth(9) = 0
gridarti.ColWidth(10) = 0
gridarti.ColWidth(11) = 0
gridarti.ColWidth(0) = 300
gridarti.ColWidth(1) = 3500
gridarti.ColWidth(2) = 600
gridarti.ColWidth(3) = 800
gridarti.ColWidth(4) = 800
gridarti.ColWidth(5) = 0
gridarti.ColWidth(6) = 0
gridarti.ColWidth(7) = 0
gridarti.ColWidth(8) = 0
gridarti.ColWidth(9) = 0
gridarti.ColWidth(10) = 0
gridarti.ColWidth(11) = 0

gridarti.Rows = 2

End Sub

Private Sub cmdmostrar_Click()
Dim wflag_reg As String
Dim ws_count As Currency
Dim ps_autoinv As rdoResultset
Dim xpinta As Integer
Dim xfil  As Integer
Dim WARTI As rdoQuery
Dim wllave_arti As rdoResultset

Dim WBOT As rdoQuery
Dim wllave_bot As rdoResultset

Dim wvalor As Integer
Dim CADE_ROT As String
If Val(Right(cmbfami.Text, 8)) = 0 Then
  MsgBox "Selecione la Familia / Muestre colocando la fecha y pulse Enter", 48, Pub_Titulo
  Exit Sub
End If
pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODCIA = ? AND ARM_CODART = ? "
Set WBOT = CN.CreateQuery("", pub_cadena)
WBOT(0) = 0
WBOT(1) = 0
Set wllave_bot = WBOT.OpenResultset(rdOpenKeyset, rdConcurValues)
If LK_CODCIA = "01" Then
 WBOT(0) = "04"
Else
 WBOT(0) = LK_CODCIA
End If
CADE_ROT = ""


pub_cadena = "SELECT ARM_COSPRO, ART_STOCK_MAX , PRE_COSTO_ANT, ART_KEY , ART_MARGEN, PRE_EQUIV,art_alterno, art_nombre, PRE_EQUIV ,PRE_PRE6 , PRE_UNIDAD, ARM_STOCK , ((ARM_STOCK/PRE_EQUIV) - (ART_STOCK_MAX/PRE_EQUIV)) AS DIF , ART_PLANCHA FROM ARTI, articulo, precios WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART ) AND PRE_FLAG_UNIDAD = 'A' AND ART_CODCIA = ? AND ART_FAMILIA = ?  AND ART_SITUACION <> '1' ORDER BY ART_NOMBRE"
Set WARTI = CN.CreateQuery("", pub_cadena)
WARTI(0) = LK_CODCIA
WARTI(1) = Val(Right(cmbfami.Text, 8))
Set wllave_arti = WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)
wllave_arti.Requery
If Not wllave_arti.EOF Then
 ProBar.Min = 0
 ProBar.max = wllave_arti.RowCount
 ProBar.Value = 0
End If
ProBar.Visible = True
gridarti.Visible = False
DoEvents
gridarti.Clear
gridarti.Cols = 13
gridarti.TextMatrix(0, 0) = "Item"
gridarti.TextMatrix(0, 1) = "Descripcion del Producto"
gridarti.TextMatrix(0, 2) = "Unid."
gridarti.TextMatrix(0, 2) = "Activa"
gridarti.TextMatrix(0, 3) = "Equiv" ' SISTEMA GUARDADO
gridarti.TextMatrix(1, 3) = "Activa." ' SISTEMA GUARDADO
gridarti.TextMatrix(0, 4) = "Cantidad"
gridarti.TextMatrix(1, 4) = "p'Unidad"
gridarti.ColWidth(0) = 300
gridarti.ColWidth(1) = 3500
gridarti.ColWidth(2) = 600
gridarti.ColWidth(3) = 800
gridarti.ColWidth(4) = 800
gridarti.ColWidth(5) = 0
gridarti.ColWidth(6) = 0
gridarti.ColWidth(7) = 0
gridarti.ColWidth(8) = 0
gridarti.ColWidth(9) = 0
gridarti.ColWidth(10) = 0
gridarti.ColWidth(11) = 0
gridarti.ColWidth(12) = 0

gridarti.Rows = 2

xfil = 1
Do Until wllave_arti.EOF
  ProBar.Value = ProBar.Value + 1
  DoEvents
  xfil = xfil + 1
  gridarti.Rows = gridarti.Rows + 1
  gridarti.RowHeight(gridarti.Rows - 1) = 285
'  lisarti.AddItem wllave_arti!art_alterno & " " & wllave_arti!ART_NOMBRE & String(120, " ") & wllave_arti!ART_KEY
  gridarti.TextMatrix(xfil, 0) = Format(xfil - 1, "00")
  gridarti.TextMatrix(xfil, 1) = Trim(wllave_arti!ART_NOMBRE)
  gridarti.TextMatrix(xfil, 2) = Trim(wllave_arti!pre_unidad)
  gridarti.Row = xfil
  gridarti.COL = 3
  gridarti.CellForeColor = QBColor(1)
  gridarti.TextMatrix(xfil, 3) = Val(wllave_arti!pre_equiv)
  gridarti.TextMatrix(xfil, 4) = Val(wllave_arti!ART_MARGEN)
  gridarti.TextMatrix(xfil, 10) = Val(wllave_arti!art_key)
  
  
 
sale_otro:
wllave_arti.MoveNext
Loop

On Error GoTo SALE
gridarti.Visible = True
If gridarti.Rows > 2 Then
  gridarti.Row = 2
  gridarti.COL = 4
  gridarti.SetFocus
End If
ProBar.Visible = False
DoEvents
SALE:
End Sub


Private Sub cmdRegoper_Click()
cmdRegoper.Enabled = False
Dim ws_numoper6 As Integer
Dim ws_numoper5 As Integer
Dim f_reg_tipmov5 As String * 1
Dim f_reg_tipmov6 As String * 1
Dim flag_reg As String * 1
Dim WS_CODTRA As Integer
Dim ws_linea As String
Dim WS_TOTAL As Currency
Dim WS_SUBTOTAL As Currency
Dim ws_descri As String
Dim ws_equiv As Currency
Dim WDIF_CAJ As Currency
Dim WDIF_UNI As Currency
Dim ws_count As Currency
Dim WS_NUM_MAS  As Currency
Dim WS_NUM_MENOS As Currency
cmbcal_Click
If lblreg.Visible = True Then
    cmdRegoper.Enabled = True
    MsgBox "OJO!! ya genero el proceso !!!!" & Chr(13) & "Necesita clave de Gerencia", 48, Pub_Titulo
    LK_ACCESO_REPORT = ""
    Load frmclave2
    Screen.MousePointer = 0
    frmclave2.Show 1
    If LK_ACCESO_REPORT <> "A" Then
       Exit Sub
    End If
End If
pub_mensaje = "Generar el Registro de Regularización Automatico de Inventario...   ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
  cmdRegoper.Enabled = True
  Exit Sub
End If



'*******************
' FACART
'*******************
On Error GoTo error_fatal
pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)

WS_CODTRA = 2403
SQ_OPER = 3
PU_TIPMOV = 6  ' AUMENTO
pu_codcia = LK_CODCIA
PU_NUMSER = 99
PU_FBG = " "
LEER_FAR_LLAVE
If Not far_menor.EOF Then
   WS_NUM_MAS = Val(far_menor!far_numfac) + 1
Else
   WS_NUM_MAS = 1
End If
SQ_OPER = 3
PU_TIPMOV = 5  ' DESMINUCION
pu_codcia = LK_CODCIA
PU_NUMSER = 88
PU_FBG = " "
LEER_FAR_LLAVE
If Not far_menor.EOF Then
   WS_NUM_MENOS = Val(far_menor!far_numfac) + 1
Else
  WS_NUM_MENOS = 1
End If
ws_linea = Trim(Left(cmbfami.Text, 15))


CN.Execute "Begin Transaction", rdExecDirect

GoSub manda_numero

PUB_NUM_OPER = PUB_NUM_OPER_XXX
ws_numoper6 = PUB_NUM_OPER
ws_numoper5 = PUB_NUM_OPER + 1


WS_NUMSEC = 0
flag_reg = ""
f_reg_tipmov6 = ""
f_reg_tipmov5 = ""
 
For fila = 2 To gridarti.Rows - 1
  ws_count = cant_equiv(Val(gridarti.TextMatrix(fila, 10)))
  If Val(gridarti.TextMatrix(fila, 9)) = 1 Or Val(gridarti.TextMatrix(fila, 9)) > 2 Then
    WS_CANTIDAD = Abs(Val(gridarti.TextMatrix(fila, 5)))
    If Val(gridarti.TextMatrix(fila, 5)) > 0 Then
     pub_signo_arm = 1
    ElseIf Val(gridarti.TextMatrix(fila, 5)) < 0 Then
     pub_signo_arm = -1
    End If
  ElseIf Val(gridarti.TextMatrix(fila, 9)) = 2 Then
    WDIF_CAJ = Val(gridarti.TextMatrix(fila, 5)) * loc_ev1 ' CONVERTIDO TODO EN UNIDADES
    WDIF_UNI = Val(gridarti.TextMatrix(fila, 8))
    WS_CANTIDAD = WDIF_CAJ + WDIF_UNI ' TODO EN UNIDADES
    If WS_CANTIDAD > 0 Then
      pub_signo_arm = 1
    ElseIf WS_CANTIDAD < 0 Then
      pub_signo_arm = -1
    End If
    WS_CANTIDAD = Abs(WS_CANTIDAD)
  End If
  If pub_signo_arm = 1 Then
    PUB_NUMSER = 99
    PUB_TIPMOV = 6
    PUB_NUMFAC = WS_NUM_MAS
    WS_TOTAL = Abs(Format(Val(lblvalsob.Caption), "0.00")) '
    PUB_NUM_OPER = ws_numoper6
  Else
    WS_TOTAL = Abs(Format(Val(lblvalfal.Caption), "0.00")) '
    PUB_NUMSER = 88
    PUB_TIPMOV = 5
    PUB_NUM_OPER = ws_numoper5
    PUB_NUMFAC = WS_NUM_MENOS
  End If
  If WS_CANTIDAD = 0 Then GoTo salta_registrof
  If PUB_TIPMOV = 6 Then f_reg_tipmov6 = "A"
  If PUB_TIPMOV = 5 Then f_reg_tipmov5 = "A"
  WS_SUBTOTAL = Abs(Val(gridarti.TextMatrix(fila, 12)))
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_CODART = Val(gridarti.TextMatrix(fila, 10))
  LEER_ARM_LLAVE
  arm_llave.Edit
  arm_llave!arm_stock = arm_llave!arm_stock + (WS_CANTIDAD * pub_signo_arm)
  arm_llave.Update
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_SECUEN = 0
  LEER_PRE_LLAVE
  If pre_llave.EOF Then
   ws_descri = ""
  Else
   ws_descri = pre_llave!pre_unidad
  End If
  ws_equiv = 1
  GoSub REG_FACART
  flag_reg = "A"
salta_registrof:
 ' MsgBox gridarti.TextMatrix(fila, 1) & Chr(13) & "SIGNO: " & pub_signo_arm & "  CANTIDAD : " & WS_CANTIDAD & " EQUIVS : " & loc_ev0 & " " & loc_ev1 & " " & loc_ev2
Next fila
If f_reg_tipmov5 = "A" Then
 pub_cadena = "UPDATE AUTOINV SET INV_REG_FECHA = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "' , INV_REG_NS5 = 88 , INV_REG_NF5 = " & WS_NUM_MENOS & " , INV_FLAG_REG = 1 WHERE INV_CODCIA = '" & LK_CODCIA & "'  AND INV_FAMILIA = " & Trim(Right(cmbfami.Text, 8)) & "  AND INV_FECHA = '" & Format(txtfecha.Text, "dd/mm/yyyy") & "'"
 CN.Execute pub_cadena, rdExecDirect
 If flag_reg = "A" Then ' SOLO CUANDO SE REGISTRE EN FACART
    WS_TOTAL = Abs(Format(Val(lblvalfal.Caption), "0.00")) '
    PUB_NUMSER = 88
    PUB_TIPMOV = 5
    PUB_NUM_OPER = ws_numoper5
    PUB_NUMFAC = WS_NUM_MENOS
    pub_signo_arm = -1
    GoSub REG_ALLOG
  End If
End If
If f_reg_tipmov6 = "A" Then
  pub_cadena = "UPDATE AUTOINV SET INV_REG_FECHA = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "' , INV_REG_NS6 = 99 , INV_REG_NF6= " & WS_NUM_MAS & " , INV_FLAG_REG = 1 WHERE INV_CODCIA = '" & LK_CODCIA & "'  AND INV_FAMILIA = " & Trim(Right(cmbfami.Text, 8)) & "  AND INV_FECHA = '" & Format(txtfecha.Text, "dd/mm/yyyy") & "'"
  CN.Execute pub_cadena, rdExecDirect
  If flag_reg = "A" Then ' SOLO CUANDO SE REGISTRE EN FACART
    PUB_NUMSER = 99
    PUB_TIPMOV = 6
    PUB_NUMFAC = WS_NUM_MAS
    WS_TOTAL = Abs(Format(Val(lblvalsob.Caption), "0.00")) '
    PUB_NUM_OPER = ws_numoper6
    pub_signo_arm = 1
    GoSub REG_ALLOG
  End If
End If
      

CN.Execute "Commit Transaction", rdExecDirect
con_llave.Close
cmdRegoper.Enabled = True
If flag_reg = "A" Then ' SOLO CUANDO SE REGISTRE EN FACART
   cmdmostrar_Click
   MsgBox "Registro Automatico Efectuado", 48, Pub_Titulo
Else
   MsgBox "No se encontro ninguna diferencia.", 48, Pub_Titulo
End If

Exit Sub

'_____________
REG_FACART:
'-------------
      far_llave.AddNew
      far_llave!FAR_TIPMOV = PUB_TIPMOV
      far_llave!FAR_CODCIA = LK_CODCIA
      far_llave!far_cod_sunat = 0 'Val(Right(i_codsunat.Text, 5))
      far_llave!far_numser = PUB_NUMSER
      far_llave!FAR_CODVEN = 0
      far_llave!far_numfac = PUB_NUMFAC
      WS_NUMSEC = WS_NUMSEC + 1
      far_llave!FAR_NUMSEC = WS_NUMSEC
      far_llave!FAR_STOCK = 0
      far_llave!far_codart = PUB_CODART
      far_llave!far_cantidad = WS_CANTIDAD
      far_llave!far_cantidad_p = WS_CANTIDAD
       ' AFECTANDO A LOTES
       '******************
      PUB_CODLOT = BUSCA_LOTE_AFECTAR(Val(PUB_CODART), pub_signo_arm, Val(WS_CANTIDAD))     ' BUSCA UN LOTE PARA AFECTAR EL AJUSTE
      PSLOT_LLAVE(0) = LK_CODCIA
      PSLOT_LLAVE(1) = arm_llave!ARM_CODART
      PSLOT_LLAVE(2) = PUB_CODLOT
      lot_llave.Requery
      If lot_llave.EOF Then
       lot_llave.AddNew
       lot_llave!LOT_CODCIA = LK_CODCIA
       lot_llave!LOT_NROLOTE = PUB_CODLOT
       lot_llave!lot_codart = arm_llave!ARM_CODART
       lot_llave!lot_fecha_vcto = LK_FECHA_DIA
       lot_llave!LOT_SALDOS = 0
      Else
       lot_llave.Edit
      End If
      lot_llave!lot_codclie = 0
      lot_llave!LOT_SALDOS = lot_llave!LOT_SALDOS + (WS_CANTIDAD * pub_signo_arm)
      lot_llave.Update
      
      far_llave!far_codlot = PUB_CODLOT
      far_llave!FAR_PRECIO = arm_llave!ARM_COSPRO
      far_llave!FAR_equiv = ws_equiv
      far_llave!far_descri = ws_descri 'Trim(xl.Cells(WS_FILA, 6))
      far_llave!far_PESO = 0
      far_llave!far_signo_car = 0
      far_llave!far_signo_arm = pub_signo_arm
      far_llave!far_codclie = 0
      far_llave!FAR_MONEDA = "S"
      far_llave!FAR_EX_IGV = 0
      far_llave!FAR_cp = " "
      far_llave!FAR_fecha_compra = LK_FECHA_DIA
      far_llave!far_estado = "N"
      far_llave!FAR_ESTADO2 = "N"
      far_llave!FAR_COSPRO = arm_llave!ARM_COSPRO
      far_llave!FAR_COSPRO_ANT = 0
      far_llave!far_fbg = " "
      far_llave!far_impto = 0
      far_llave!FAR_TOT_FLETE = 0
      far_llave!FAR_FLETE = 0
      far_llave!FAR_DESCTO = 0
      far_llave!FAR_TOT_DESCTO = 0
      far_llave!FAR_GASTOS = 0
      far_llave!far_bruto = WS_TOTAL
      far_llave!FAR_NUMDOC = 0
      far_llave!far_numguia = 0
      far_llave!far_serguia = 0
      far_llave!FAR_pordescto1 = 0
      far_llave!FAR_costeo = ""
      far_llave!FAR_COSTEO_REAL = ""
      far_llave!FAR_tipo_cambio = 1
      far_llave!FAR_DIAS = 0
      far_llave!FAR_fecha = LK_FECHA_DIA
      far_llave!FAR_NUMSER_C = 0
      far_llave!FAR_NUMFAC_C = 0
      far_llave!FAR_NUMOPER = PUB_NUM_OPER
      far_llave!FAR_NUMOPER2 = PUB_NUM_OPER
      far_llave!far_precio_neto = 0
      far_llave!far_otra_cia = " "
      far_llave!far_transito = " "
      far_llave!far_subtra = "Ajuste Automtico"
      far_llave!far_otra_cia = " "
      far_llave!far_transito = " "
      far_llave!far_JABAS = 0
      far_llave!far_UNIDADES = 0
      far_llave!far_mortal = 0
      far_llave!far_num_precio = 0
      far_llave!FAR_ORDEN_UNIDADES = 0
      far_llave!FAR_SUBTOTAL = WS_SUBTOTAL
      far_llave!far_turno = 0
      far_llave!far_concepto = " "
      far_llave!far_concepto = "Aj. Autom. - : " & ws_linea
      far_llave!far_codusu = LK_CODUSU
      far_llave!far_hora = Format(Now, "hh:mm:ss AMPM")
      If PUB_TIPMOV = 5 Then
        far_llave!FAR_NUM_LOTE = 1
      Else
        far_llave!FAR_NUM_LOTE = 0
      End If

      far_llave!FAR_PEDSER = 0
      far_llave!FAR_PEDFAC = 0
      far_llave!far_pedsec = 0
      far_llave!FAR_TIPDOC = ""
      far_llave!FAR_fecha_can = LK_FECHA_DIA
      far_llave!FAR_fecha_pro = LK_FECHA_DIA
      far_llave!far_key_dircli = 0
      far_llave.Update
Return

'*******************
' REGISTRO DE ALLOG
'*******************



'*******************
REG_ALLOG:
'*******************
    
all_llave.AddNew
all_llave!ALL_NUMOPER = PUB_NUM_OPER
all_llave!ALL_NUMOPER2 = PUB_NUM_OPER
all_llave!all_CODCIA = LK_CODCIA
all_llave!ALL_codtra = WS_CODTRA
all_llave!all_flag_ext = "N"
all_llave!ALL_CODCLIE = pu_codclie
all_llave!ALL_CODART = 0
all_llave!ALL_IMPORTE_amort = 0
all_llave!all_codusu = LK_CODUSU
all_llave!ALL_FBG = ""
all_llave!ALL_CODVEN = 0
all_llave!ALL_IMPORTE = WS_TOTAL ' llave_rep01!par_saldo_caja_hoy
all_llave!ALL_IMPORTE_DOLL = 0 'llave_rep01!PAR_SALDO_CAJA_D_HOY
all_llave!ALL_NUMDOC = 0
all_llave!ALL_CP = " "
all_llave!ALL_TIPDOC = ""
all_llave!all_numfac_c = 0
all_llave!all_numser_c = 0
all_llave!all_codban = 0
all_llave!all_concepto = WDOCU
all_llave!all_chenum = 0
all_llave!ALL_FECHA_DIA = LK_FECHA_DIA
all_llave!ALL_FECHA_SUNAT = LK_FECHA_DIA
all_llave!ALL_FECHA_VCTO = LK_FECHA_DIA
all_llave!ALL_CANTIDAD = WS_CANTIDAD
all_llave!ALL_NUMSER = PUB_NUMSER
all_llave!all_numfac = PUB_NUMFAC
all_llave!all_neto = 0
all_llave!ALL_BRUTO = 0
all_llave!ALL_tipmov = PUB_TIPMOV
all_llave!ALL_IMPTO = 0
all_llave!ALL_flete = 0
all_llave!ALL_HORA = Now
all_llave!ALL_DESCTO = 0
all_llave!ALL_GASTOS = 0
all_llave!ALL_PRECIO = 0
all_llave!ALL_MONEDA_CLI = ""
all_llave!ALL_moneda_ccm = ""
all_llave!ALL_MONEDA_CAJA = ""
all_llave!all_SECUENCIA = 0
all_llave!ALL_SIGNO_CAR = 0
all_llave!ALL_signo_caja = 0
all_llave!ALL_SIGNO_CCM = 0
all_llave!all_sIGNO_ARM = pub_signo_arm
all_llave!all_chenum = 0
all_llave!ALL_CHESEC = 0
all_llave!ALL_CHESER = 0
all_llave!ALL_SUBTRA = ""
all_llave!ALL_TIPO_BLOQ_ACT = 0
all_llave!ALL_TIPO_BLOQ_ANT = 0
all_llave!all_codtra_ext = 0
all_llave!ALL_TIPO_CAMBIO = 0
all_llave!ALL_RUC = 0
all_llave!ALL_IMPG1 = 0
all_llave!ALL_IMPG2 = 0
all_llave!ALL_GASTOS = 0
all_llave.Update
Return
Exit Sub
error_fatal:
CN.Execute "Rollback Transaction", rdExecDirect
con_llave.Close
MsgBox Err.Description

Exit Sub

manda_numero:
SQ_OPER = 2
PUB_FECHA = LK_FECHA_DIA
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
If all_menor.EOF = False Then
   PUB_NUM_OPER_XXX = all_menor!ALL_NUMOPER
Else
   PUB_NUM_OPER_XXX = 0
End If
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
Return

End Sub

Private Sub Form_Load()
RUTINA_LLENA_FAM

CenterMe frmEquivArti
End Sub

Public Sub RUTINA_LLENA_FAM()
Dim WS_CADEN As String

PUB_TIPREG = 122
PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_TAB_LLAVE
cmbfami.Clear
Do Until tab_mayor.EOF
    cmbfami.AddItem tab_mayor!tab_nomlargo & " " & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
  tab_mayor.MoveNext
Loop
End Sub


Private Sub gridarti_GotFocus()
TEXTOVAR.Visible = False
End Sub

Private Sub gridarti_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Dim car As String
Static CONS
'If KeyAscii <> 13 Then Exit Sub
' solo numeros
   If gridarti.COL = 4 Or gridarti.COL = 7 Then
   Else
   Exit Sub
   End If
   If gridarti.COL = 7 Then
     If Val(gridarti.TextMatrix(gridarti.Row, 9)) = 1 Or Val(gridarti.TextMatrix(gridarti.Row, 9)) > 2 Then
       MsgBox "No procede este producto solo registra unidades minimoas", 48, Pub_Titulo
       Exit Sub
     End If
   End If
   
   
   car = Chr(KeyAscii)
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." And car <> "-" Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If
    TEXTOVAR.Left = gridarti.Left + gridarti.CellLeft
    TEXTOVAR.Width = gridarti.CellWidth
    TEXTOVAR.Height = gridarti.CellHeight
    TEXTOVAR.Top = gridarti.Top + gridarti.CellTop
    TEXTOVAR.Text = gridarti.TextMatrix(gridarti.Row, gridarti.COL)
    TEXTOVAR.Visible = True
    Azul3 TEXTOVAR, TEXTOVAR
    TEXTOVAR.Text = Trim(car)
    TEXTOVAR.SelStart = Len(TEXTOVAR)
    TEXTOVAR.SetFocus
'End If
End Sub

Private Sub gridarti_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean
Dim Wsec

'If WMODO = "C" Then Exit Sub

'If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
'End If


If Left(gridarti.TextMatrix(gridarti.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
  'If WMODO <> "C" Then Exit Sub
  tC = gridarti.COL
  For fila = 1 To gridarti.Cols - 1
      gridarti.COL = fila
      If gridarti.CellBackColor = QBColor(12) Then
         gridarti.CellBackColor = QBColor(15)
         gridarti.TextMatrix(gridarti.Row, 9) = "9"
      Else
         gridarti.CellBackColor = QBColor(12)
         gridarti.TextMatrix(gridarti.Row, 9) = "-1"
      End If
  Next fila
  gridarti.COL = tC
  gridarti.SetFocus
  Exit Sub
End If
If KeyCode = 45 Then
    Wsec = Wsec + 1
    If Trim(gridarti.TextMatrix(gridarti.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(gridarti.TextMatrix(gridarti.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(gridarti.TextMatrix(gridarti.Row, 4)) = 0 And Val(gridarti.TextMatrix(gridarti.Row, 5)) = 0 Then Exit Sub
    gridarti.AddItem "", gridarti.Row + 1
    gridarti.TextMatrix(gridarti.Row + 1, 0) = "MAN. " & Format(gridarti.TextMatrix(gridarti.Row, 10), "dd/mm/yyyy")
    gridarti.TextMatrix(gridarti.Row + 1, 6) = Wsec
    gridarti.TextMatrix(gridarti.Row + 1, 8) = gridarti.TextMatrix(gridarti.Row, 8)
    gridarti.TextMatrix(gridarti.Row + 1, 3) = gridarti.TextMatrix(gridarti.Row, 3)
    gridarti.TextMatrix(gridarti.Row + 1, 7) = gridarti.TextMatrix(gridarti.Row, 7)
    gridarti.TextMatrix(gridarti.Row + 1, 10) = gridarti.TextMatrix(gridarti.Row, 10)
    gridarti.TextMatrix(gridarti.Row + 1, 11) = "8"
    gridarti.Row = gridarti.Row + 1
    gridarti.COL = 1
    gridarti.SetFocus
End If
Exit Sub
If KeyCode = 46 Then
If gridarti.Rows <= 3 Then
Else
   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(gridarti.TextMatrix(gridarti.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     gridarti.SetFocus
     Exit Sub
   Else
     gridarti.RowHeight(gridarti.Row) = 1
     gridarti.Row = gridarti.Row + 1
    
   'gridarti.RemoveItem (gridarti.Row)
   'gridarti.Refresh
   gridarti.SetFocus
   End If
End If
End If
'gridarti.SetFocus
Exit Sub



End Sub

Private Sub gridarti_Scroll()
TEXTOVAR.Visible = False
End Sub


Private Sub textovar_Change()
gridarti.Text = Format(TEXTOVAR.Text, "0")
End Sub

Private Sub TEXTOVAR_GotFocus()
 temporal = gridarti.TextMatrix(gridarti.Row, gridarti.COL)
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
Dim wval As Integer
Dim ps_autoinv As rdoResultset
If KeyAscii = 27 Then
  TEXTOVAR.Text = temporal
  TEXTOVAR.Visible = False
  gridarti.SetFocus
  Exit Sub
End If
If gridarti.COL = 4 Then Consistencias gridarti, TEXTOVAR, KeyAscii
If gridarti.COL = 7 Then Consistencias gridarti, TEXTOVAR, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If
' actualizar datos
'================
' A TODAS LAS CIAS RELACIONADAS
'******************************

pub_cadena = "SELECT ART_KEY,ART_MARGEN FROM ARTI WHERE ART_CODCIA = '" & LK_CODCIA & "' AND ART_KEY = " & Val(gridarti.TextMatrix(gridarti.Row, 10))
Set ps_autoinv = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not ps_autoinv.EOF Then
    If Val(gridarti.TextMatrix(gridarti.Row, 4)) <> 0 Then
       SQ_OPER = 2
       PUB_CODART = ps_autoinv!art_key
       PUB_CODCIA = LK_CODCIA
       LEER_PRE_LLAVE
       wval = 1
       Do Until pre_mayor.EOF
          wval = pre_mayor!pre_equiv
          pre_mayor.MoveNext
       Loop

        WRES = Int(wval / Val(gridarti.TextMatrix(gridarti.Row, 4)))
        If (WRES - Val(Format(wval / Val(gridarti.TextMatrix(gridarti.Row, 4)), "0.00"))) <> 0 Then
          MsgBox "El Factor para la Unidades Minimas no corresponde a la unidad Mayor ", 48, Pub_Titulo
          TEXTOVAR.Text = Val(ps_autoinv!ART_MARGEN)
          Azul3 TEXTOVAR, TEXTOVAR
          Exit Sub
        End If
        ps_autoinv.Edit
        ps_autoinv!ART_MARGEN = Val(gridarti.TextMatrix(gridarti.Row, 4))
        ps_autoinv.Update
     End If
 
End If
    
    

'================
If gridarti.Row >= gridarti.Rows - 1 And gridarti.COL <> 4 Then
 gridarti.SetFocus
 TEXTOVAR.Visible = False
Else
  If gridarti.COL = 4 Then
       GoTo sala_lin
  ElseIf gridarti.COL = 7 Then
sala_lin:
   If gridarti.Row >= gridarti.Rows - 1 Then
   Else
     gridarti.Row = gridarti.Row + 1
     gridarti.COL = 4
   End If
  End If
  gridarti.SetFocus
  TEXTOVAR.Visible = False
End If

fin:

End Sub


Private Sub txtfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If IsDate(txtfecha.Text) Then
    RUTINA_LLENA_FAM
  End If
End If
End Sub
Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As RichTextBox, wsKeyAscii As Integer)
  Static VALOR
  Dim car As String
 ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If

End Sub


Public Function cant_equiv(wsp_codart As Currency) As Integer
Dim rs_precios As rdoResultset
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND PRE_CODART = " & wsp_codart & " "
Set rs_precios = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
cant_equiv = rs_precios.RowCount
loc_ev0 = rs_precios!pre_equiv
rs_precios.MoveNext
If rs_precios.EOF Then GoTo SALIR
loc_ev1 = rs_precios!pre_equiv
rs_precios.MoveNext
If rs_precios.EOF Then GoTo SALIR
loc_ev2 = rs_precios!pre_equiv
SALIR:
End Function

Public Sub ESTABLECE_DIF(wfila)

 If (Val(gridarti.TextMatrix(wfila, 4)) - Val(gridarti.TextMatrix(wfila, 3))) = 0 Then
  gridarti.TextMatrix(wfila, 5) = ""
 Else
  gridarti.TextMatrix(wfila, 5) = Val(gridarti.TextMatrix(wfila, 4)) - Val(gridarti.TextMatrix(wfila, 3))
 End If
 If (Val(gridarti.TextMatrix(wfila, 7)) - Val(gridarti.TextMatrix(wfila, 6))) = 0 Then
   gridarti.TextMatrix(wfila, 8) = ""
 Else
   gridarti.TextMatrix(wfila, 8) = Val(gridarti.TextMatrix(wfila, 7)) - Val(gridarti.TextMatrix(wfila, 6))
 End If

End Sub

Public Function BUSCA_LOTE_AFECTAR(WS_CODART As Currency, WS_SIGNO As Integer, wlot_cantidad As Currency) As String
  PSLOT_LLAVE2(0) = LK_CODCIA
  PSLOT_LLAVE2(1) = WS_CODART
  lot_llave2.Requery
  Do Until lot_llave2.EOF
   If WS_SIGNO = 1 Then
      If lot_llave2!LOT_SALDOS <= 0 Then
         BUSCA_LOTE_AFECTAR = lot_llave2!LOT_NROLOTE
         GoTo listo
      End If
   Else
      If lot_llave2!LOT_SALDOS > 0 And Val(lot_llave2!LOT_SALDOS) >= Val(wlot_cantidad) Then
         BUSCA_LOTE_AFECTAR = lot_llave2!LOT_NROLOTE
         GoTo listo
      End If
   End If
   lot_llave2.MoveNext
  Loop
 BUSCA_LOTE_AFECTAR = "(*)"
listo:
End Function

Public Sub sum_total()
Dim WS_TOTAL_MENOS  As Currency
Dim WS_TOTAL_MAS  As Currency
Dim WS_TOTAL As Currency
WS_TOTAL_MENOS = 0
WS_TOTAL_MAS = 0
On Error GoTo sigue
ProBar.max = gridarti.Rows - 1
ProBar.Value = 2
ProBar.Min = 2

ProBar.Visible = True
DoEvents
For fila = 2 To gridarti.Rows - 1
 ProBar.Value = fila
 DoEvents
 ws_count = cant_equiv(Val(gridarti.TextMatrix(fila, 10)))
  If Val(gridarti.TextMatrix(fila, 9)) = 1 Or Val(gridarti.TextMatrix(fila, 9)) > 2 Then
    WS_CANTIDAD = Abs(Val(gridarti.TextMatrix(fila, 5)))
    If Val(gridarti.TextMatrix(fila, 5)) > 0 Then
     pub_signo_arm = 1
    ElseIf Val(gridarti.TextMatrix(fila, 5)) < 0 Then
     pub_signo_arm = -1
    End If
  ElseIf Val(gridarti.TextMatrix(fila, 9)) = 2 Then
    WDIF_CAJ = Val(gridarti.TextMatrix(fila, 5)) * loc_ev1 ' CONVERTIDO TODO EN UNIDADES
    WDIF_UNI = Val(gridarti.TextMatrix(fila, 8))
    WS_CANTIDAD = WDIF_CAJ + WDIF_UNI ' TODO EN UNIDADES
    If WS_CANTIDAD > 0 Then
      pub_signo_arm = 1
    ElseIf WS_CANTIDAD < 0 Then
      pub_signo_arm = -1
    End If
    WS_CANTIDAD = Abs(WS_CANTIDAD)
  End If
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_CODART = Val(gridarti.TextMatrix(fila, 10))
  LEER_ARM_LLAVE
  gridarti.TextMatrix(fila, 12) = Format(Val(arm_llave!ARM_COSPRO) * (WS_CANTIDAD * pub_signo_arm), "0.000")
  If pub_signo_arm = -1 Then
    WS_TOTAL_MENOS = WS_TOTAL_MENOS + Val(Format(Val(arm_llave!ARM_COSPRO) * (WS_CANTIDAD * pub_signo_arm), "0.00"))
  Else
    WS_TOTAL_MAS = WS_TOTAL_MAS + Val(Format(Val(arm_llave!ARM_COSPRO) * (WS_CANTIDAD * pub_signo_arm), "0.00"))
  End If
  WS_TOTAL = WS_TOTAL + Val(Format(Val(arm_llave!ARM_COSPRO) * (WS_CANTIDAD * pub_signo_arm), "0.00"))
Next fila
ProBar.Visible = False
lblvalfal.Caption = Format(WS_TOTAL_MENOS, "0.00")
lblvalsob.Caption = Format(WS_TOTAL_MAS, "0.00")
lblvaltotal.Caption = Format(WS_TOTAL, "0.00")
Exit Sub
sigue:
Resume Next
End Sub

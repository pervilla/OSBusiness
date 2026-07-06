VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PRODIA 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Actualizacion de Fechas"
   ClientHeight    =   5895
   ClientLeft      =   300
   ClientTop       =   1770
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frmprodia.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   3885
   Begin VB.Frame Frame3 
      Caption         =   "Situaci�n de Operaciones de la Compa�ia: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   3615
      Begin VB.PictureBox poperativo 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   2760
         Picture         =   "Frmprodia.frx":0442
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox pbloqueado 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   2760
         Picture         =   "Frmprodia.frx":0884
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.OptionButton option1 
         Caption         =   "Operativo."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton option1 
         Caption         =   "Bloqueado."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cerrar Operaciones del D�a."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   720
      Picture         =   "Frmprodia.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1200
      Picture         =   "Frmprodia.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
      Begin VB.ListBox EMP 
         BackColor       =   &H00808000&
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
         Height          =   840
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Cierre de Compa�ia(s) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   3480
      Top             =   3840
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "OSBusiness"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   5640
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "D�a:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label LblFecha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label POR 
      AutoSize        =   -1  'True
      Caption         =   "Cerrando Operaciones Diarias..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label lblcierre 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESO DE  FECHA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   600
      TabIndex        =   0
      Top             =   45
      Width           =   2595
   End
End
Attribute VB_Name = "PRODIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WS_SALDO As Currency
Dim WS_SALDO_D As Currency
Dim PSPRE_MAYOR2  As rdoQuery
Dim pre_mayor2 As rdoResultset
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim stock_llave As rdoResultset
Dim PSST_LLAVE As rdoQuery
Dim chedef_llave As rdoResultset
Dim PSCHE_DEF  As rdoQuery






Public Sub REPO_CAJA_GEN2(ww_codcia As String)
Dim ww_moneda
WS_SALDO = 0
WS_SALDO_D = 0

PUB_FECHA = LK_FECHA_DIA
pu_codcia = ww_codcia
SQ_OPER = 1
LEER_ALL_LLAVE
If all_llave.EOF Then Exit Sub
PUB_CODCIA = ww_codcia
LEER_PAR_LLAVE


   WS_SALDO = Nulo_Valor0(par_llave!PAR_SALDO_CAJA_ayer)
   WS_SALDO_D = Nulo_Valor0(par_llave!PAR_SALDO_CAJA_D_AYER)

Do Until all_llave.EOF
   If all_llave!ALL_signo_caja = 0 Then GoTo OTRO
   
   
   If all_llave!all_flag_ext = "E" Then GoTo OTRO
   
   If all_llave!ALL_SIGNO_CAR = 0 And all_llave!ALL_tipmov = 0 Then
      WS_IMPORTE = all_llave!ALL_IMPORTE
   Else
      WS_IMPORTE = all_llave!ALL_IMPORTE_AMORT
   End If
   
   
   If Trim(all_llave!ALL_moneda_ccm) <> " " And Val(all_llave!all_codban) <> 0 Then
      ww_moneda = all_llave!ALL_moneda_ccm
   ElseIf Trim(all_llave!ALL_MONEDA_CLI) <> " " And Val(all_llave!ALL_CODCLIE) <> 0 Then
      ww_moneda = all_llave!ALL_MONEDA_CLI
   ElseIf Trim(all_llave!ALL_MONEDA_CAJA) <> " " Then
      ww_moneda = all_llave!ALL_MONEDA_CAJA
   End If
   
   If ww_moneda = "S" Then
   If all_llave!ALL_signo_caja = 1 Then
      WS_SALDO = WS_SALDO + WS_IMPORTE
   Else
      WS_SALDO = WS_SALDO - WS_IMPORTE
   End If
   End If
   If ww_moneda = "D" Then
   If all_llave!ALL_signo_caja = 1 Then
      WS_SALDO_D = WS_SALDO_D + WS_IMPORTE
   Else
      WS_SALDO_D = WS_SALDO_D - WS_IMPORTE
   End If
   End If
OTRO:
  all_llave.MoveNext
  Loop
  
par_llave.Edit
par_llave!PAR_SALDO_CAJA_HOY = WS_SALDO
'par_llave!PAR_SALDO_CAJA_D_HOY = WS_SALDO_D
par_llave.Update



End Sub

Private Sub Command1_Click()
Dim WS_CAJSALDO_C1 As Currency
Dim WS_CAJSALDO_C2 As Currency
Dim WS_CAJSALDO_C3 As Currency

Dim CONTADOR As Integer
Dim ww_ult_oper As Currency
Dim WS_SALDO_S As Currency
Dim WS_SALDO_D As Currency

Dim WS_CODTRA As Integer
Dim WW_FECHA As Date
Dim WS_SALDO2 As Currency
Dim WS_BLOQ1, WS_BLOQ2 As String
Dim ws_saldo_caa As Currency
Dim WS_MONEDA As String * 1
Dim ww_dias As Integer
Dim wcodven As Integer
Dim WDOCU As String
Dim xcuenta As Integer
Dim ws_codcia As String * 2
ww_dias = 0
wcodven = 0
ws_saldo_caa = 0
If GEN!gen_cierre_todas = 0 Then
    pub_mensaje = "CIERRE DEL DIA DE LA EMPRESA : " & Trim(par_llave!par_nombre) & " �Desea Continuar... ?"
Else
    pub_mensaje = "CIERRE DEL DIA DE EMPRESA(S) !!! ...   �Desea Continuar... ?"
End If
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
If par_llave!PAR_FECHA_DIA <> LK_FECHA_DIA Then
   MsgBox "!La Fecha del Sistema ha cambiado, ya se cerro el d�a de la compa�ia ..!, el sietma se cerrar� ", 48, Pub_Titulo
   End
End If

If Nulo_Valor0(GEN!GEN_TASA_VENTA) = 99 Then
   CANCEL_CH
End If


WDOCU = "Inicio de Operaciones "
Dim COS, CAA, PRE, CCMM As rdoResultset
Dim PSCAR, PSCOS, PSCCMM As rdoQuery

'pub_cadena = "SELECT FAR_COSTEO_REAL FROM FACART WHERE FAR_CODCIA = ? AND FAR_COSTEO_REAL = 'A'"
'Set PSCOS = CN.CreateQuery("", pub_cadena)
'PSCOS(0) = LK_CODCIA
'Set COS = PSCOS.OpenResultset(rdOpenKeyset, rdConcurValues)

'pub_cadena = "SELECT * FROM PRECIOS "
'Set PRE = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CCMAEST"
Set CCMM = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CODCIA = '" & LK_CODCIA & "' AND CAR_IMPORTE > 0 AND CAR_CP='C'  AND  CAR_FECHA_VCTO_ORIG <= ? "
Set PSCAR = CN.CreateQuery("", pub_cadena)
PSCAR(0) = LK_FECHA_DIA
Set car = PSCAR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? "
Set PSPRE_MAYOR2 = CN.CreateQuery("", pub_cadena)
PSPRE_MAYOR2(0) = LK_CODCIA
Set pre_mayor2 = PSPRE_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)


Timer1.Enabled = False
lblcierre.Visible = True
Command1.Enabled = False

ProgBar.Visible = True
DoEvents
POR(2).Visible = True
DoEvents
CONTADOR = 0
WS_BLOQ1 = ""
WS_BLOQ2 = ""

CN.Execute "BEGIN TRANSACTION", rdExecDirect

CN.Execute "UPDATE CCMAEST SET CCM_SAL_ANTERIOR = CCM_SALDO", rdExecDirect
If LK_EMP_PTO = "A" Then
    xcuenta = 1
    For fila = 1 To 30
      ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
      If Trim(ws_codcia) = "" Then Exit For
      GoSub cierra_cia
      PS_REP01(0) = ws_codcia
      llave_rep01.Requery
      GoSub PASA_POR_CIAS
      xcuenta = xcuenta + 2
    Next fila
Else
     ws_codcia = LK_CODCIA
     If Trim(par_llave!par_ruta_imp) <> "" Then
         If LK_CODCIA <> "01" And LK_CODCIA <> "10" And LK_CODCIA <> "07" And LK_CODCIA <> "05" And LK_CODCIA <> "09" And LK_CODCIA <> "50" And LK_CODCIA <> "25" And LK_CODCIA <> "26" And LK_CODCIA <> "27" And LK_CODCIA <> "28" And LK_CODCIA <> "29" Then GoSub cierra_cia
     End If
     
     PS_REP01(0) = LK_CODCIA
     llave_rep01.Requery
     GoSub PASA_POR_CIAS
End If

CN.Execute "commit TRANSACTION", rdExecDirect
 
MsgBox "Proceso de Cierre Terminado Satisfactoriamente... Ahora el Sistema se Cerrara...", 48, Pub_Titulo
End
Exit Sub


PASA_POR_CIAS:


   llave_rep01.Requery
   SQ_OPER = 1
   PUB_CODCIA = llave_rep01!par_codcia
   PUB_CAL_INI = llave_rep01!PAR_FECHA_DIA
   PUB_CAL_FIN = DateAdd("m", 1, llave_rep01!PAR_FECHA_DIA)
   LEER_CAL_LLAVE 1
   If cal_llave.EOF Then
      MsgBox "NO PUEDE SER.... ERROR GRAVE EN CAL"
      GoTo SALIDA_ERROR
   End If
   If Trim(par_llave!par_ruta_imp) <> "" Then
        If LK_CODCIA <> "01" And LK_CODCIA <> "07" And LK_CODCIA <> "05" And LK_CODCIA <> "10" And LK_CODCIA <> "09" And LK_CODCIA <> "50" And LK_CODCIA <> "11" And LK_CODCIA <> "25" And LK_CODCIA <> "26" And LK_CODCIA <> "27" And LK_CODCIA <> "28" And LK_CODCIA <> "29" Then GoTo solo_cierre
   End If
   
   
   GoSub manda_numero
   
   WS_CODTRA = 9999
   SQ_OPER = 1
   PUB_NUMTAB = 0
   PUB_CODCIA = llave_rep01!par_codcia
   PUB_TIPREG = 1000
   LEER_TAB_LLAVE
   If Val(tab_llave!TAB_NOMLARGO) <> Val(PUB_NUM_OPER_XXX - 1) Then
       MsgBox "Falta Procesar el Reporte de Caja en SOLES (S/.).", 48, Pub_Titulo
      'FREDY 23/11/2002
      If PUB_CODCIA <> "04" Then
       GoTo SALIDA_ERROR
      Else
      End If
   End If
   WS_SALDO_S = 0
   WS_CAJSALDO_C1 = 0
   WS_CAJSALDO_C2 = 0
   WS_CAJSALDO_C3 = 0
   'ICA
   If LK_FLAG_GRIFO = "A" Or LK_EMP = "3AA" Or LK_EMP = "PIU" Or LK_EMP = "PAR" Then
   Else
    WS_SALDO_S = tab_llave!TAB_contable2
   End If
   SQ_OPER = 1
   PUB_CODCIA = llave_rep01!par_codcia
   PUB_TIPREG = 320
   PUB_NUMTAB = 1
   LEER_TAB_LLAVE
   If Val(tab_llave!tab_codclie) <> Val(PUB_NUM_OPER_XXX - 1) Then
       MsgBox "Falta Procesar el Reporte de Caja en SOLES (S/.). " & Chr(13) & "Codigo de Caja N� : " & Format(PUB_NUMTAB, "00"), 48, Pub_Titulo
       GoTo SALIDA_ERROR
  End If
  WS_CAJSALDO_C1 = Val(tab_llave!TAB_contable2)
   
   PUB_NUMTAB = 2
   LEER_TAB_LLAVE
   If Val(tab_llave!tab_codclie) <> Val(PUB_NUM_OPER_XXX - 1) Then
       MsgBox "Falta Procesar el Reporte de Caja en SOLES (S/.). " & Chr(13) & "Codigo de Caja N� : " & Format(PUB_NUMTAB, "00"), 48, Pub_Titulo
       GoTo SALIDA_ERROR
   End If
   WS_CAJSALDO_C2 = Val(tab_llave!TAB_contable2)
   
   PUB_NUMTAB = 3
   LEER_TAB_LLAVE
   If Val(tab_llave!tab_codclie) <> Val(PUB_NUM_OPER_XXX - 1) Then
       MsgBox "Falta Procesar el Reporte de Caja en SOLES (S/.). " & Chr(13) & "Codigo de Caja N� : " & Format(PUB_NUMTAB, "00"), 48, Pub_Titulo
       GoTo SALIDA_ERROR
   End If
   WS_CAJSALDO_C3 = Val(tab_llave!TAB_contable2)
   If (Val(WS_CAJSALDO_C1) + Val(WS_CAJSALDO_C2) + Val(WS_CAJSALDO_C3)) <> WS_SALDO_S Then
      MsgBox "Falta Revisar o Procesar las Cajas no cuadran las Internas en SOLES (S/.). " & Chr(13) & "Codigo de Caja N� : " & Format(PUB_NUMTAB, "00"), 48, Pub_Titulo
      GoTo SALIDA_ERROR
   End If
   
   
'   SQ_OPER = 1
'   PUB_NUMTAB = 0
'   PUB_CODCIA = llave_rep01!PAR_CODCIA
'   PUB_TIPREG = 1001
'   LEER_TAB_LLAVE
'   If Not tab_llave.EOF Then
'    If Val(tab_llave!tab_nomlargo) <> Val(PUB_NUM_OPER_XXX - 1) Then
'       MsgBox "Falta Procesar el Reporte de Caja en DOLARES (US$/).", 48, Pub_Titulo
'       'FREDY 23/11/2002
'      If PUB_CODCIA <> "04" Then
'       GoTo SALIDA_ERROR
'      Else
'      End If
'    End If
'   End If
   WS_SALDO_D = 0
   'ICA
   If LK_FLAG_GRIFO = "A" Or LK_EMP = "3AA" Or LK_EMP = "PIU" Or LK_EMP = "PAR" Then
   Else
    If Not tab_llave.EOF Then WS_SALDO_D = tab_llave!TAB_contable2
   End If
solo_cierre:
   cal_llave.Edit
   cal_llave!CAL_INDICE = 3
   cal_llave.Update
   CONTADOR = 1
   cal_llave.MoveNext
   Do Until cal_llave!CAL_LABORABLE = "S" Or cal_llave.EOF
      CONTADOR = CONTADOR + 1
      cal_llave.MoveNext
   Loop
   If cal_llave.EOF Then
      MsgBox "Falta las fechas .... "
      GoTo SALIDA_ERROR
   End If
   cal_llave.Edit
   cal_llave!CAL_INDICE = 1
   cal_llave.Update
   'agfkn �asghd�ahg�khagfds�kafdsga sgfda sdfgasfd
   GoSub manda_numero
   If Trim(par_llave!par_ruta_imp) <> "" Then
      If LK_CODCIA <> "01" And LK_CODCIA <> "07" And LK_CODCIA <> "10" And LK_CODCIA <> "05" And LK_CODCIA <> "09" And LK_CODCIA <> "50" And LK_CODCIA <> "11" And LK_CODCIA <> "25" And LK_CODCIA <> "26" And LK_CODCIA <> "27" And LK_CODCIA <> "28" And LK_CODCIA <> "29" Then GoTo salta_grifo
   End If
   
   GoSub GRABA_ALLOG
   PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
   GoSub PROCESA_CAR
   ProgBar.Min = 0
   CONTADOR = 0
   PSPRE_MAYOR2.rdoParameters(0) = llave_rep01!par_codcia
   pre_mayor2.Requery
   CONTADOR = 0
   If Not pre_mayor2.EOF Then ProgBar.max = pre_mayor2.RowCount
   Do Until pre_mayor2.EOF
         pre_mayor2.Edit
         pre_mayor2!PRE_cosTO_ant = Nulo_Valor0(pre_mayor2!PRE_COSTO)
         pre_mayor2.Update
         pre_mayor2.MoveNext
         CONTADOR = CONTADOR + 1
         ProgBar.Value = CONTADOR
         DoEvents
   Loop
   ' PROCESO DE ACTUALIZACION DE COSTOS
   '-----------------------------------
     
    pub_cadena = "SELECT CAR_FECHA_VCTO FROM CARTERA WHERE CAR_CODCIA = '" & LK_CODCIA & "' AND CAR_IMPORTE > 0 AND CAR_CP='P'  AND  CAR_FECHA_VCTO <= ? "
    Set PSCAR = CN.CreateQuery("", pub_cadena)
    PSCAR(0) = LK_FECHA_DIA
    Set car = PSCAR.OpenResultset(rdOpenKeyset, rdConcurValues)
    car.Requery
    ProgBar.Min = 0
    If Not car.EOF Then ProgBar.max = car.RowCount
    ProgBar.Value = 0
    Do Until car.EOF
       ProgBar.Value = ProgBar.Value + 1
       car.Edit
       car!car_fecha_vcto = cal_llave!cal_fecha
       car.Update
      car.MoveNext
    Loop

   If LK_FLAG_GRIFO <> "A" Then GoTo salta_grifo
   SQ_OPER = 2
   PUB_KEY = 0
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   Do Until art_mayor.EOF
     'art_mayor.Edit
     'art_mayor!art_cash = Val(Nulo_Valor0(art_mayor!art_cash)) + Val(Nulo_Valor0(art_mayor!art_margen))
     'art_mayor.Update
     art_mayor.MoveNext
   Loop
salta_grifo:
llave_rep01.Requery
llave_rep01.Edit
llave_rep01!par_flag_cierre = 0
'llave_rep01!PAR_SALDO_CAJA_ayer = Nulo_Valor0(llave_rep01!par_saldo_caja_hoy)
'llave_rep01!PAR_SALDO_CAJA_D_ayer = Nulo_Valor0(llave_rep01!PAR_SALDO_CAJA_D_HOY)
llave_rep01!PAR_FECHA_DIA = cal_llave!cal_fecha
llave_rep01.Update

Return



PROCESA_CAR:

pub_cadena = "UPDATE CLIENTES SET CLI_AUTOAVALUO =  '' where cli_codcia = '" & LK_CODCIA & "' and cli_cp = 'C' and CLI_AUTOAVALUO <> ''"
CN.Execute pub_cadena, rdExecDirect

pub_cadena = "UPDATE CLIENTES SET CLI_CUENTA_CONTAB =  '' where cli_codcia = '" & LK_CODCIA & "' and cli_cp = 'C' and CLI_CUENTA_CONTAB <> ''"
CN.Execute pub_cadena, rdExecDirect

pub_cadena = "UPDATE CLIENTES SET CLI_regpub1 =  '' where cli_codcia = '" & LK_CODCIA & "' and cli_cp = 'C' and CLI_regpub1 <> ''"
CN.Execute pub_cadena, rdExecDirect
pub_cadena = "UPDATE CLIENTES SET CLI_limcre2 =  0 where cli_codcia = '" & LK_CODCIA & "' and cli_cp = 'C' and CLI_limcre2 <> 0"
CN.Execute pub_cadena, rdExecDirect




If Nulo_Valor0(par_llave!PAR_DIAS_LARGE) = 0 Then Return

pub_cadena = "UPDATE CLIENTES SET CLI_TIPO_BLOQ1 = ' ' WHERE CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_CP = 'C' AND CLI_TIPO_BLOQ1 =  '1' "
CN.Execute pub_cadena, rdExecDirect


WW_FECHA = DateAdd("d", cal_llave!cal_fecha, Nulo_Valor0(par_llave!PAR_DIAS_LARGE) * -1)

PSCAR.rdoParameters(0) = WW_FECHA
car.Requery
Do Until car.EOF
   SQ_OPER = 1
   pu_codclie = car!CAR_codclie
   pu_codcia = car!car_codcia
   pu_cp = "C"
   LEER_CLI_LLAVE
   ww_dias = DateDiff("d", car!car_fecha_VCTO_orig, cal_llave!cal_fecha)
   If car!car_NUMFAC <> 0 Then
     WDOCU = car!car_fbg & " / " & Format(car!car_NUMSER, "000") & " - " & Format(car!car_NUMFAC, "00000000")
   Else
     WDOCU = car!CAR_TIPDOC
     If car!CAR_TIPDOC = "CH" Then
        WDOCU = car!CAR_TIPDOC & " / " & Format(car!car_num_cheque, "00000000")
     End If
   End If
   wcodven = car!CAR_codven
   ws_saldo_caa = car!car_importe
   If Nulo_Valors(cli_llave!CLI_TIPO_BLOQ1) <> "1" Then
      WS_BLOQ1 = cli_llave!CLI_TIPO_BLOQ1 & Nulo_Valors(cli_llave!CLI_TIPO_BLOQ2) & (cli_llave!CLI_TIPO_BLOQ3) & Nulo_Valors(cli_llave!CLI_TIPO_BLOQ4)
      cli_llave.Edit
      cli_llave!CLI_TIPO_BLOQ1 = "1"
      cli_llave.Update
      WS_BLOQ2 = cli_llave!CLI_TIPO_BLOQ1 & Nulo_Valors(cli_llave!CLI_TIPO_BLOQ2) & (cli_llave!CLI_TIPO_BLOQ3) & Nulo_Valors(cli_llave!CLI_TIPO_BLOQ4)
      WS_CODTRA = 2582
      GoSub GRABA_ALLOG
      PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
   End If
   car.MoveNext
Loop



Return


manda_numero:
SQ_OPER = 2
PUB_FECHA = cal_llave!cal_fecha
pu_codcia = llave_rep01!par_codcia
LEER_ALL_LLAVE
If all_menor.EOF = False Then
   PUB_NUM_OPER_XXX = all_menor!ALL_NUMOPER
Else
   PUB_NUM_OPER_XXX = 0
End If
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
Return

cierra_cia:
      PS_REP01(0) = ws_codcia
      llave_rep01.Requery
    If llave_rep01!par_flag_cierre <> 9 Then
       MsgBox "!!! Falta realizar Cierre de Operaciones ..." & llave_rep01!par_codcia & "-" & llave_rep01!par_nombre
       GoTo SALIDA_ERROR
       Exit Sub
    End If
'    If llave_rep01!par_flag_costos <> 9 Then
'       MsgBox "!!! Falta procesar Costos  del Dia ...En " & llave_rep01!PAR_CODCIA & " - " & Trim(llave_rep01!PAR_NOMBRE), 48, Pub_Titulo
'       GoTo SALIDA_ERROR
'    End If
    
    If LK_EMP = "HER" Or LK_FLAG_GRIFO = "A" Or LK_EMP = "3AA" Or LK_EMP = "PIU" Or LK_EMP = "PAR" Or (LK_ICA = "A") Then
       GoTo PASE_CAJA
    End If
    If LK_FLAG_SOS = "A" Then GoTo PASE_CAJA
    SQ_OPER = 1
    pu_codcia = llave_rep01!par_codcia
    PUB_FECHA = LK_FECHA_DIA
    LEER_ALL_LLAVE
    all_llave.MoveLast
    If all_llave.EOF Then
       MsgBox "El Sistema Esta Cerrando sin Movimientos ... !!!", 48, Pub_Titulo
       GoTo PASE_CAJA
    End If
    SQ_OPER = 1
    PUB_NUMTAB = 0
    PUB_CODCIA = llave_rep01!par_codcia
    PUB_TIPREG = 1000
    LEER_TAB_LLAVE
    If tab_llave.EOF Then
       MsgBox "!!! Falta procesar Caja Soles  del Dia ...En " & llave_rep01!par_codcia & " - " & Trim(llave_rep01!par_nombre), 48, Pub_Titulo
       GoTo SALIDA_ERROR
       Exit Sub
    Else
      If ((Val(all_llave!ALL_NUMOPER) <> Val(tab_llave!TAB_NOMLARGO)) Or (CDate(tab_llave!tab_nomcorto) <> LK_FECHA_DIA)) And Val(tab_llave!TAB_NOMLARGO) <> 0 Then
       MsgBox "!!! Falta procesar Caja Soles   del Dia ...Ir a Contabilidad " & llave_rep01!par_codcia & " - " & Trim(llave_rep01!par_nombre), 48, Pub_Titulo
       GoTo SALIDA_ERROR
       Exit Sub
      Else
        WS_SALDO_S = tab_llave!TAB_contable2
      End If
    End If
    
    SQ_OPER = 1
    PUB_NUMTAB = 0
    PUB_CODCIA = llave_rep01!par_codcia
    PUB_TIPREG = 1001
    LEER_TAB_LLAVE
    If tab_llave.EOF Then
       MsgBox "!!! Falta procesar Caja  Dollares (U$$) del Dia ...En " & llave_rep01!par_codcia & " - " & Trim(llave_rep01!par_nombre), 48, Pub_Titulo
       GoTo SALIDA_ERROR
       Exit Sub
    Else
      If (Val(all_llave!ALL_NUMOPER) <> Val(tab_llave!TAB_NOMLARGO) Or tab_llave!tab_nomcorto <> LK_FECHA_DIA) And Val(tab_llave!TAB_NOMLARGO) <> 0 Then
       MsgBox "!!! Falta procesar Caja  Dollares (U$$) del Dia ...En " & llave_rep01!par_codcia & " - " & Trim(llave_rep01!par_nombre), 48, Pub_Titulo
       GoTo SALIDA_ERROR
       Exit Sub
      Else
      WS_SALDO_D = tab_llave!TAB_contable2
      End If
    End If
    
    llave_rep01.Edit
    llave_rep01!PAR_SALDO_CAJA_HOY = WS_SALDO_S
    'llave_rep01!PAR_SALDO_CAJA_D_HOY = WS_SALDO_D
    llave_rep01.Update
    
PASE_CAJA:
' AGREGE ACV
GoTo PASO
    'PSCOS.rdoParameters(0) = llave_rep01!PAR_CODCIA
    'COS.Requery
    'Do Until COS.EOF
    '    COS.Edit
    '    COS!FAR_COSTEO_REAL = " "
    '    COS.Update
    '    COS.MoveNext
    'Loop
PASO:
Return


GRABA_ALLOG:
all_llave.AddNew
all_llave!ALL_NUMOPER = PUB_NUM_OPER_XXX

all_llave!ALL_CODCIA = llave_rep01!par_codcia
all_llave!ALL_CODTRA = WS_CODTRA
all_llave!all_flag_ext = "E"
all_llave!ALL_CODCLIE = pu_codclie
all_llave!ALL_CODART = 0
all_llave!ALL_IMPORTE_AMORT = 0
all_llave!all_codusu = LK_CODUSU
all_llave!ALL_FBG = ""
all_llave!ALL_CODVEN = wcodven
all_llave!ALL_IMPORTE = WS_SALDO_S ' llave_rep01!par_saldo_caja_hoy
all_llave!ALL_IMPORTE_DOLL = WS_SALDO_D 'llave_rep01!PAR_SALDO_CAJA_D_HOY
all_llave!ALL_NUMDOC = 0
all_llave!ALL_CP = pu_cp
all_llave!ALL_TIPDOC = ""
all_llave!all_numfac_c = 0
all_llave!all_numser_c = 0
all_llave!all_codban = 0
all_llave!all_concepto = WDOCU
all_llave!all_chenum = 0
all_llave!ALL_FECHA_DIA = cal_llave!cal_fecha
all_llave!ALL_FECHA_SUNAT = cal_llave!cal_fecha
all_llave!ALL_FECHA_VCTO = cal_llave!cal_fecha
all_llave!ALL_CANTIDAD = ww_dias
all_llave!ALL_NUMSER = 0
all_llave!all_numfac = 0
all_llave!all_neto = 0
all_llave!ALL_BRUTO = ws_saldo_caa
all_llave!ALL_tipmov = 0
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
all_llave!all_sIGNO_ARM = 0
all_llave!all_chenum = 0
all_llave!ALL_CHESEC = 0
all_llave!ALL_CHESER = 0
all_llave!ALL_SUBTRA = ""
all_llave!ALL_TIPO_BLOQ_ACT = WS_BLOQ2
all_llave!ALL_TIPO_BLOQ_ANT = WS_BLOQ1
all_llave!all_codtra_ext = 0
all_llave!ALL_TIPO_CAMBIO = 0
all_llave!ALL_RUC = 0
all_llave!ALL_IMPG1 = WS_CAJSALDO_C1
all_llave!ALL_IMPG2 = WS_CAJSALDO_C2
all_llave!ALL_GASTOS = WS_CAJSALDO_C3

all_llave.Update
Return

SALIDA_ERROR:
SALIR:
CN.Execute "ROLLBACK TRANSACTION", rdExecDirect
Unload PRODIA
Exit Sub
fin:
MsgBox Err.Description
Resume Next
End Sub

Private Sub Command4_Click()
Unload PRODIA
End Sub

Private Sub Form_Activate()
Dim xcuenta  As Integer
Dim ws_codcia As Integer
If LK_EMP_PTO = "A" Then
   If LK_CODCIA <> "00" Then
     Screen.MousePointer = 0
     MsgBox "Se Encuentra en punto de Venta. El Cierre del d�a se ejecuta en la Compa�ia Central ", 48, Pub_Titulo
     Unload PRODIA
     Exit Sub
   End If
'   Label3.Caption = "ALMACEN Y PUNTOS DE VENTAS"
   Frame3 = "Estado de las Operaciones de  Empresas es: "
   option1(0).Caption = "D i s p o n i b l e s"
   option1(1).Caption = "C e r r a d a s"
   xcuenta = 1
   For fila = 1 To 30
      ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
      If Trim(ws_codcia) = "" Then Exit For
      PS_REP01(0) = ws_codcia
      llave_rep01.Requery
      EMP.AddItem Trim(llave_rep01!par_codcia) + " " + Trim(llave_rep01!par_nombre)
      xcuenta = xcuenta + 2
   Next fila
   lblfecha.Caption = Format(LK_FECHA_DIA, "dddd, d mmmm yyyy")
   EMP.ListIndex = 0
   SQ_OPER = 1
   PUB_CODCIA = LK_CODCIA
   LEER_PAR_LLAVE
   If par_llave!par_flag_cierre = 9 Then
      option1(1).Value = True
      option1(1).ForeColor = QBColor(12)
      option1(0).ForeColor = QBColor(0)
   Else
      option1(0).Value = True
      option1(0).ForeColor = QBColor(2)
      option1(1).ForeColor = QBColor(0)
   End If
Else
'   Label3.Caption = "C O M P A � I A"
   SQ_OPER = 1
   PUB_CODCIA = LK_CODCIA
   LEER_PAR_LLAVE
   EMP.AddItem Trim(par_llave!par_codcia) + " " + Trim(par_llave!par_nombre)
   lblfecha.Caption = Format(LK_FECHA_DIA, "dddd, dd Mmmm yyyy")
   EMP.ListIndex = 0
   If par_llave!par_flag_cierre = 9 Then
      option1(1).Value = True
      option1(1).ForeColor = QBColor(12)
      option1(0).ForeColor = QBColor(0)
      pbloqueado.Visible = True
      poperativo.Visible = False
   Else
      option1(0).Value = True
      option1(0).ForeColor = QBColor(2)
      option1(1).ForeColor = QBColor(0)
      pbloqueado.Visible = False
      poperativo.Visible = True
   End If
End If

If option1(0).Visible Then
'  Option1(0).SetFocus
End If

End Sub

Private Sub Form_DblClick()
CANCEL_CH
End Sub

Private Sub Form_Load()
Dim ws_codcia As String * 2
Dim xcuenta As Integer
Dim PSFAR_TRANS As rdoQuery
Dim FAR_TRANS As rdoResultset
CenterMe PRODIA
pub_cadena = ""
If LK_CODCIA = "10" Or LK_CODCIA = "09" Or LK_CODCIA = "50" Then GoTo pasa_10
'pub_cadena = "SELECT FAR_CODCIA  FROM facart WHERE  FAR_TRANSITO = 'T' AND FAR_ESTADO <> 'E' AND FAR_OTRA_CIA = ? AND FAR_FECHA >= ? AND FAR_ESTADO2 <> 'L' ORDER BY FAR_OTRA_CIA, FAR_FBG, FAR_NUMSER, FAR_NUMFAC,FAR_NUMSEC"
'Set PSFAR_TRANS = CN.CreateQuery("", pub_cadena)
'PSFAR_TRANS.rdoParameters(0) = "00"
'PSFAR_TRANS.rdoParameters(1) = Date
'PSFAR_TRANS.MaxRows = 1
'Set FAR_TRANS = PSFAR_TRANS.OpenResultset(rdOpenKeyset, rdConcurValues)
'PSFAR_TRANS.rdoParameters(0) = LK_CODCIA
'PSFAR_TRANS.rdoParameters(1) = Format(DateAdd("d", -30, LK_FECHA_DIA), "dd/mm/yyyy")
'FAR_TRANS.Requery

pub_cadena = "SELECT FAR_CODCIA  FROM facart WHERE  FAR_TRANSITO = 'T' AND FAR_ESTADO <> 'E' AND FAR_OTRA_CIA = '" & LK_CODCIA & "' AND FAR_FECHA >= '" & Format(DateAdd("d", -30, LK_FECHA_DIA), "dd/mm/yyyy") & "' AND FAR_ESTADO2 <> 'L' ORDER BY FAR_OTRA_CIA, FAR_FBG, FAR_NUMSER, FAR_NUMFAC,FAR_NUMSEC"
Set FAR_TRANS = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
 
If Not FAR_TRANS.EOF Then
    Screen.MousePointer = 0
    If Trim(LK_CODUSU) = "ADMIN" Or Trim(LK_CODUSU) = "SUPER" Then
       MsgBox "Existe Mercaderia en Transito, Verificar con las Operaci�nes de Recepci�n de Mercaderia, Tiene el acceso!!!!"
    Else
       MsgBox "Existe Mercaderia en Transito, Verificar con las Operaci�nes de Recepci�n de Mercaderia, No Procede,"
       Unload PRODIA
       Exit Sub
    End If
    'pub_mensaje = "�Desea Continuar de Todas Maneras ... ?"
    'Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    'If Pub_Respuesta = vbNo Then
    '   Unload PRODIA
    '   Exit Sub
    'End If
End If

pub_cadena = "SELECT ALL_CODCIA  FROM ALLOG WHERE  ALL_FECHA_DIA = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "' and ALL_NUMSER_C = '" & LK_CODCIA & "' AND ALL_NUMFAC_C = '9'"
Set FAR_TRANS = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
 
If Not FAR_TRANS.EOF Then
    Screen.MousePointer = 0
    'If Trim(LK_CODUSU) = "ADMIN" Or Trim(LK_CODUSU) = "SUPER" Then
    '   MsgBox "Existe Mercaderia en Transito, Verificar con las Operaci�nes de Recepci�n de Mercaderia, Tiene el acceso!!!!"
    'Else
       MsgBox "Existe Efectivo Transito, Verificar con las Operaci�nes de Recepci�n de Efectivo Operaci�n: 5375, No Procede,"
       Unload PRODIA
       Exit Sub
    'End If
End If


pasa_10:
If LK_FLAG_GRIFO = "A" Then
 pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ?  AND TAB_NOMCORTO = ? AND TAB_CODCLIE = ? AND TAB_CODART = ?  ORDER BY TAB_NUMTAB "
 Set PSST_LLAVE = CN.CreateQuery("", pub_cadena)
 PSST_LLAVE(0) = 0
 PSST_LLAVE(1) = 0
 PSST_LLAVE(2) = 0
 PSST_LLAVE(3) = 0
 PSST_LLAVE(4) = 0
 Set stock_llave = PSST_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
End If


'Frame3 = "Estado de Operaciones :"
'Option1(0).Caption = "Di s p o n i b l e"
'Option1(1).Caption = "C e r r a d o"
If LK_EMP_PTO = "A" Then
    xcuenta = 1
    For fila = 1 To 30
       ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
       If Trim(ws_codcia) = "" Then Exit For
        pub_cadena = pub_cadena + " FFF_CODCIA = '" & ws_codcia & "' OR "
       xcuenta = xcuenta + 2
    Next fila
    If pub_cadena = "" Then
       MsgBox "!!!! Verificar esta Activado puntos de Ventas pero no existe declaracion de Cias!!!!", 48, Pub_Titulo
       Unload PRODIA
       Exit Sub
    End If
End If

Timer1.Enabled = True

'Dim chedef_llave As rdoResultset
'Dim PSCHE_DEF  As rdoQuery''

'pub_cadena = "SELECT  CHE_ESTADO FROM CHEQUES WHERE CHE_FECHA_COBRO = ? AND CHE_ESTADO = 'T' ORDER BY CHE_FECHA_COBRO"
'Set PS_REP01 = CN.CreateQuery("", pub_cadena)
'PS_REP01(0) = LK_CODCIA
'Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
pub_cadena = "SELECT ALL_CODCIA  FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ? AND ALL_CODTRA = 2401"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
If llave_rep01.EOF Then
   MsgBox "OJO !!! El Dia esta sin Operaciones de Venta ", 48, Pub_Titulo
   pub_mensaje = "El dia esta sin operaciones de Venta !!! ...   �Desea Continuar... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then
     Screen.MousePointer = 0
     Unload PRODIA
     Exit Sub
   End If
End If

pub_cadena = "SELECT  * FROM PARGEN WHERE PAR_CODCIA = ?  order by par_codcia"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = LK_CODCIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

Exit Sub

If Nulo_Valor0(GEN!gen_cierre_todas) = 1 Or Nulo_Valor0(GEN!gen_cierre_todas) = 3 Then
   Frame2.Visible = False
   'Frame1.Left = Frame1.Left - 1500
Else
   lblfecha.Caption = "Fecha Actual : " & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy")
'   Label3.Caption = Trim(par_llave!PAR_NOMBRE)
   SQ_OPER = 1
   PUB_CODCIA = LK_CODCIA
   LEER_PAR_LLAVE
   If par_llave!par_flag_cierre = 9 Then
      option1(1).Value = True
   Else
      option1(0).Value = True
   End If
End If

If Nulo_Valor0(GEN!gen_cierre_todas) = 1 Or Nulo_Valor0(GEN!gen_cierre_todas) = 3 Then
End If

End Sub


Private Sub Option1_Click(Index As Integer)
If option1(0).Visible = False Then GoTo fin

For fila = 0 To EMP.ListCount - 1
   EMP.ListIndex = fila
   If Trim(Left(EMP.Text, 2)) = "" Then
     MsgBox "Seleccione de la lista una Compa�ia.", 48, Pub_Titulo
     EMP.SetFocus
     Exit Sub
   End If
   PS_REP01(0) = Left(EMP.Text, 2)
   llave_rep01.Requery
   llave_rep01.Edit
   If Index = 1 Then
      llave_rep01!par_flag_cierre = 9
      option1(1).Value = True
      option1(1).ForeColor = QBColor(12)
      option1(0).ForeColor = QBColor(0)
      poperativo.Visible = False
      pbloqueado.Visible = True
   Else
      llave_rep01!par_flag_cierre = 0
      option1(0).Value = True
      option1(1).ForeColor = QBColor(0)
      option1(0).ForeColor = QBColor(2)
      poperativo.Visible = True
      pbloqueado.Visible = False
   End If
  llave_rep01.Update
Next fila
EMP.ListIndex = 0
Command1.SetFocus
Exit Sub
   PS_REP01(0) = Left(EMP.Text, 2)
   llave_rep01.Requery
   If llave_rep01!par_flag_cierre = 9 Then
      option1(1).Value = True
      option1(1).ForeColor = QBColor(12)
      option1(0).ForeColor = QBColor(0)
   Else
      option1(0).Value = True
      option1(1).ForeColor = QBColor(0)
      option1(0).ForeColor = QBColor(12)
   End If

Exit Sub
fin:
End Sub


Private Sub Timer1_Timer()
lblcierre.Visible = Not lblcierre.Visible
End Sub

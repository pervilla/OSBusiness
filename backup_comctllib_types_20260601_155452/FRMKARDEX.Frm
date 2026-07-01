VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form KARDEX 
   Caption         =   "Consulta de Kardex de Articulo"
   ClientHeight    =   6255
   ClientLeft      =   210
   ClientTop       =   885
   ClientWidth     =   8955
   ControlBox      =   0   'False
   ForeColor       =   &H00808000&
   Icon            =   "FRMKARDEX.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6255
   ScaleWidth      =   8955
   WindowState     =   2  'Maximized
   Begin VB.CheckBox checias 
      BackColor       =   &H00808000&
      Caption         =   "Todas las Cias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   10560
      TabIndex        =   40
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox chedes 
      BackColor       =   &H00808000&
      Caption         =   "Incluir Desactivados"
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   10500
      TabIndex        =   39
      Top             =   2280
      Width           =   1275
   End
   Begin VB.TextBox NJ 
      Height          =   285
      Left            =   10440
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Chequear Saldos de todos los Articulos"
      Height          =   675
      Left            =   10500
      TabIndex        =   36
      Top             =   1530
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808000&
      Caption         =   "Recalculo de Costos"
      Height          =   375
      Left            =   10560
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Chequea si existen saldos iniciales"
      Height          =   705
      Left            =   10530
      TabIndex        =   34
      Top             =   690
      Width           =   1245
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   735
      Left            =   5280
      TabIndex        =   12
      Top             =   3720
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
   Begin MSFlexGridLib.MSFlexGrid GridK 
      Height          =   4815
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   3
      BackColorBkg    =   8421376
      HighLight       =   0
      AllowUserResizing=   1
   End
   Begin VB.Frame Consulta1 
      Height          =   1935
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10335
      Begin VB.Frame todos 
         Caption         =   "Kardex de Articulo"
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   10335
         Begin VB.OptionButton Option1 
            Caption         =   "Mostrar solo Saldos Actual"
            Enabled         =   0   'False
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
            Height          =   195
            Index           =   2
            Left            =   6720
            TabIndex        =   2
            Top             =   240
            Width           =   2655
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Mostar por Rango de Fechas"
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
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   1
            Top             =   240
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Mostar Ultimos Movim."
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
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   2415
         End
      End
      Begin VB.CheckBox cheext 
         Caption         =   "Visualizar Extonados"
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   4680
         TabIndex        =   22
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Articulo :"
         ForeColor       =   &H00808000&
         Height          =   735
         Left            =   0
         TabIndex        =   27
         Top             =   480
         Width           =   10335
         Begin VB.ComboBox unidad 
            Height          =   315
            Left            =   7800
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox i_codart2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   8
            TabIndex        =   3
            Top             =   320
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Unidad Activa :"
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
            Height          =   435
            Left            =   6840
            TabIndex        =   30
            Top             =   180
            Width           =   735
            WordWrap        =   -1  'True
         End
         Begin VB.Label LBLUNIDAD 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   7800
            TabIndex        =   29
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label i_nomarti 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   405
            Left            =   1560
            TabIndex        =   28
            Top             =   255
            Width           =   4995
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   615
         Left            =   3000
         TabIndex        =   15
         Top             =   1215
         Visible         =   0   'False
         Width           =   4455
         Begin MSMask.MaskEdBox txtCampo2 
            Height          =   285
            Left            =   3000
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCampo1 
            Height          =   285
            Left            =   1320
            TabIndex        =   5
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fechas "
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
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label4 
            Caption         =   "al"
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
            Height          =   255
            Left            =   2640
            TabIndex        =   24
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "del"
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
            Height          =   255
            Left            =   840
            TabIndex        =   23
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   615
         Left            =   6720
         TabIndex        =   19
         Top             =   1200
         Width           =   3615
         Begin VB.TextBox txtsaldo 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   320
            Left            =   1800
            TabIndex        =   20
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Stock Actual :"
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
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   0
         TabIndex        =   16
         Top             =   1200
         Width           =   4455
         Begin VB.ComboBox i_maximo 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000080&
            Height          =   315
            ItemData        =   "FRMKARDEX.frx":0442
            Left            =   960
            List            =   "FRMKARDEX.frx":0458
            TabIndex        =   4
            Text            =   "Combo1"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Movimientos"
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
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Ultimos "
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label lblmoneda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7440
         TabIndex        =   21
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   10560
      Picture         =   "FRMKARDEX.frx":0475
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Width           =   1185
   End
   Begin VB.CommandButton CmdProcesa 
      Caption         =   "Ver &kardex"
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
      Left            =   10560
      Picture         =   "FRMKARDEX.frx":077F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1185
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
      Left            =   10560
      Picture         =   "FRMKARDEX.frx":0BC1
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   1185
   End
   Begin MSComDlg.CommonDialog dialogo 
      Left            =   360
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
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
      Left            =   10440
      TabIndex        =   25
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Index           =   5
      Left            =   10440
      TabIndex        =   26
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblmensaje 
      Caption         =   "Procesando Información ..."
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
      Left            =   1800
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "KARDEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ff_flag As String * 1
Dim VAR_ACTIVAR As Integer
Dim X2 As rdoResultset
Dim SW As Integer
Dim ww_flag As String * 1
Dim wfecha As String
Dim PUB_OPERA As String
Dim docu As String
Dim ACTIVA_EQUIV As Currency
Dim llave1  As String
Dim loc_key As Integer
Public PSFAR_PEDIDOS As rdoQuery
Public far_pedidos As rdoResultset

Public PSFAR_CODART As rdoQuery
Public far_codart As rdoResultset
Dim PRE_UNIDADES As rdoQuery
Dim preuni_llave  As rdoResultset
Dim ART_LLAVES As rdoResultset
Dim psart As rdoQuery
Dim EQUIV_UNI() As Currency
Dim WCOD_ORIGINAL As Currency
Dim PS_VE2 As rdoQuery
Dim VE2_LLAVE As rdoResultset



 


Private Sub Command2_Click()
ff_flag = "A"
Dim Xx As rdoQuery
Dim ver_xx As rdoResultset
pub_mensaje = "Solo el Administrador del Sistema Sabe esta Opción , si no lo conoce no procesar, puede durar varios minutos ....¿ Continuar proceso ? . .."
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If

pub_cadena = "SELECT * FROM ARTI WHERE ART_KEY <> 0 AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 ORDER BY ART_ALTERNO"
Set Xx = CN.CreateQuery("", pub_cadena)
Set ver_xx = Xx.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
ver_xx.Requery
pub_cadena = ""
Do Until ver_xx.EOF
  i_codart2.Text = Trim(ver_xx!ART_alterno)
  i_codart2_KeyPress 13
  CmdProcesa_Click
'  MsgBox " PARA O SIGUE"
  ver_xx.MoveNext
Loop
ff_flag = ""
MsgBox pub_cadena
End Sub

Private Sub cheext_Click()
If option1(0).Value Then
Option1_Click 0
ElseIf option1(0).Value Then
Option1_Click 1
Else
Option1_Click 2
End If
End Sub

Private Sub CmdImprimir_Click()
Dim I, j
Dim wranF, wran1, wran2
Dim LETRAS(24) As String * 1
If GridK.Rows < 1 Or GridK.Visible = False Then
 MsgBox "Nuestre una Consulta.", 48, Pub_Titulo
  Exit Sub
End If
Dim xl As Object
On Error GoTo FINTODO
Screen.MousePointer = 11
GoSub WEXCEL
pub_cadena = ""

xl.Cells(4, 1) = "PRODUCTO : " & Trim(i_codart2.Text) & " - " & Trim(KARDEX.i_nomarti.Caption) & " > " & Trim(LBLUNIDAD.Caption)
'xl.Cells(5, 1) = "FECHA INICIAL : " & Trim(i_fecha.Text)
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
xl.Cells(2, 1) = "KARDEX DE PRODUCTOS"
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
For I = 0 To GridK.Rows - 1
  For j = 0 To 12
     If GridK.TextMatrix(I, j) = "" Then
       xl.Cells(I + 7, j + 1) = " "
     Else
       If j = 0 And I > 0 Then
         xl.Cells(I + 7, j + 1) = "'" & Format((GridK.TextMatrix(I, j)), "dd/mm/yy")
       ElseIf j = 1 And I > 0 Then
         xl.Cells(I + 7, j + 1) = "'" & Format((GridK.TextMatrix(I, j)), "dd/mm/yy")
       Else
         xl.Cells(I + 7, j + 1) = CStr(GridK.TextMatrix(I, j))
       End If
     End If
  Next j
  'xl.Application.Visible = True
  If I = 1 Then xl.Cells(I + 7, 4) = "Saldo Incial: "
Next I
GoSub LETRAS

For j = 8 To 10
  wranF = Trim(LETRAS(j)) & I + 1 + 6
  wran1 = Trim(LETRAS(j)) & 8
  wran2 = Trim(LETRAS(j)) & I + 6
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
Next
wranF = "H" & I + 7 & ":" & "L" & I + 7
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3

wranF = "A8:" & "L8"
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 3
xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
xl.Cells(2, 1) = "KADEX DE PRODUCTOS"
xl.Cells(3, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
xl.DisplayAlerts = False
xl.Worksheets(1).Protect PUB_CLAVE
xl.Application.Visible = True
Set xl = Nothing
Screen.MousePointer = 0
Exit Sub

WEXCEL:
  Dim wsfile1
'  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  xl.Workbooks.Open Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\KARDEX.xls", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

LETRAS:

LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
Return

FINTODO:
 'MsgBox ERR"Reintente Nuevamente ..", 48, Pub_Titulo
 MsgBox "Reintente Nuevamente ..", 48, Pub_Titulo
End Sub

Private Sub CmdProcesa_Click()
Dim wol_codcia As String
'On Error GoTo osERRsale
'*** KARDEX   ***
Dim WISLA
Dim WWFECHA1 As Date
Dim wwfecha2 As Date
Dim WRES
Dim ENTRADA
Dim ws_meses As Currency
Dim fx As Integer
Dim WWW_COSPRO As Currency
Dim vdocum As String * 90
Dim WSALDO2 As Currency
Dim WS_TOT As Currency
Dim TAB_MESES(12) As String * 12
Dim ws_nombre As String * 25
Dim WS_ANO As Integer
Dim WT_SAL01 As Currency
Dim WS_MES As Integer
Dim TT_MES As Integer
Dim cIngreso As Currency
Dim cSalida As Currency
Dim cSaldo As Currency
Dim fila1 As Integer
Dim PRECIO As Currency
Dim cuenta As String * 6
Dim articulo As String * 6
Dim WS_SAL01 As Currency
Dim WS_ING01 As Currency
Dim FINAL As String * 1
Dim Num_Fac As String * 5
Dim Num_Ser As String * 5
Dim WW_CODART As Long
Dim wFAR_CODART
Dim wSaldo As Currency
Dim Band As String * 1
Dim xCODCIA As String
Dim ws_equiv As Currency
Dim xcuenta As Integer
Dim wsFECHA1 As Date
Dim wcap
If option1(0).Value = True Then
  Option1_Click 0
Else
  Option1_Click 1
End If
otra_vez:
WWW_COSPRO = 0
If KARDEX.Check1.Value = 1 Then
  wcap = InputBox("Costo a Reeplazar:")
  If wcap = "" Then Exit Sub
  If Val(wcap) = 0 Then Exit Sub
  WWW_COSPRO = Val(wcap)
End If
If option1(1).Value = True Then
    If Not IsDate(txtCampo1.Text) Then
      MsgBox "Fecha no procede..", 48, Pub_Titulo
      Azul2 txtCampo1, txtCampo1
      Exit Sub
    End If
    If Not IsDate(txtCampo2.Text) Then
      MsgBox "Fecha no procede..", 48, Pub_Titulo
      Azul2 txtCampo2, txtCampo2
      Exit Sub
    End If
End If

If option1(0).Value = True Then
   'PSFAR_CODART.MaxRows = Val(i_maximo.Text)
Else
   PSFAR_CODART.MaxRows = 999999
End If

Set far_codart = PSFAR_CODART.OpenResultset(rdOpenKeyset, rdConcurValues)



'If Not IsDate(i_fecha.text) Then
'  MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
'  Azul i_fecha, i_fecha
' Exit Sub
'End If
'If DatePart("yyyy", i_fecha.text) < 1900 Then Exit Sub
'If i_mes.ListIndex > 0 And i_ano.ListIndex > 0 Then
WWFECHA1 = txtCampo1.Text
wwfecha2 = txtCampo2.Text
'End If
KARDEX.GridK.ColWidth(8) = 900
KARDEX.GridK.ColWidth(9) = 900 '
KARDEX.GridK.ColWidth(10) = 900
KARDEX.GridK.ColWidth(11) = 900



xCODCIA = LK_CODCIA

If Left(i_codart2.Text, 1) = "N" Then
End If

If i_codart2 = "" Then
    Screen.MousePointer = 0
    MsgBox "Verificar Datos .!", 48, Pub_Titulo
    KARDEX.i_codart2.SetFocus
    Exit Sub
End If
GridK.Visible = False
Screen.MousePointer = 11
If LK_FLAG_ALTERNO = "A" Then
  SQ_OPER = 3
  PUB_CODCIA = LK_CODCIA
  pu_alterno = Trim(i_codart2.Text)
  LEER_ART_LLAVE
  If art_llave_alt.EOF Then
    Screen.MousePointer = 0
    MsgBox "Codigo NO Existe ....", 48, Pub_Titulo
    i_codart2.SetFocus
    Exit Sub
  End If
  WCOD_ORIGINAL = art_llave_alt!art_key
  ENTRADA = WCOD_ORIGINAL
Else
  ENTRADA = Val(i_codart2.Text)
End If

 SQ_OPER = 1
 pu_codcia = LK_CODCIA
 PUB_KEY = Val(ENTRADA)
 LEER_ART_LLAVE
 If art_LLAVE.EOF Then
   Screen.MousePointer = 0
   MsgBox "Codigo NO Existe ....", 48, Pub_Titulo
   i_codart2.SetFocus
   Exit Sub
 End If
 
 SQ_OPER = 1
 pu_codcia = LK_CODCIA
 PUB_CODART = ENTRADA
 LEER_ARM_LLAVE
 If arm_llave.EOF Then
   Screen.MousePointer = 0
   MsgBox "Codigo NO Existe en Articulos , Verificar....", 48, Pub_Titulo
   i_codart2.SetFocus
   Exit Sub
 End If
 Frame4.Visible = True
FINAL = "S"
  If i_codart2.Text <> "" Then
   
    If option1(0).Value = True Then
       PSFAR_CODART.rdoParameters(0) = ENTRADA
       PSFAR_CODART.rdoParameters(1) = Format(par_llave!PAR_FECHA_DIA, "dd/mm/yyyy")
       PSFAR_CODART.rdoParameters(2) = xCODCIA

        If checias.Value = 0 Then GoTo go_una
        If LK_CODCIA = "09" Or LK_CODCIA = "50" Then
           'Costeo_Chepen ' solo pàra la botica chepen
go_una:
           wol_codcia = "('" & LK_CODCIA & "')"
        Else
           wol_codcia = "('25','27','28','29','10')"
        End If


       If cheext.Value = 1 Then
          ' pub_cadena = "SELECT TOP " & Val(i_maximo.Text) & " FAR_TRANSITO, FAR_FLAG_SO, FAR_TIPDOC, FAR_FECHA_COMPRA,  far_moneda, far_precio_neto, far_subtra, far_numSER_c, far_numfac_c , far_numGUIA, FAR_DIAS, FAR_SERGUIA, FAR_CODUSU,FAR_CODART, FAR_CANTIDAD, FAR_JABAS, FAR_CONCEPTO, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMOPER2, FAR_COSPRO,FAR_PRECIO, FAR_EQUIV, FAR_FECHA, FAR_STOCK, FAR_ESTADO,FAR_CP, FAR_CODCLIE, FAR_TIPMOV, FAR_OTRA_CIA, far_SIGNO_aRM ,FAR_codven, far_costeo , far_hora FROM facart WHERE FAR_CODART = " & ENTRADA & " AND FAR_FECHA <= '" & Format(par_llave!PAR_FECHA_DIA, "dd/mm/yyyy") & "' AND FAR_CODCIA = '" & LK_CODCIA & "'  AND FAR_ESTADO <> 'M' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA DESC, FAR_SIGNO_ARM ,FAR_TIPMOV DESC, FAR_NUMOPER2 DESC "
          pub_cadena = "SELECT TOP " & Val(i_maximo.Text) & "  * FROM facart WHERE FAR_CODART = " & ENTRADA & " AND FAR_FECHA <= '" & Format(par_llave!PAR_FECHA_DIA, "dd/mm/yyyy") & "' AND FAR_CODCIA IN " & wol_codcia & "  AND FAR_ESTADO <> 'M' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L' ORDER BY FAR_FECHA_COMPRA DESC, FAR_SIGNO_ARM ,FAR_TIPMOV DESC, FAR_NUMOPER2 DESC , FAR_CODCIA DESC"
       Else
          ' pub_cadena = "SELECT TOP " & Val(i_maximo.Text) & " FAR_TRANSITO, FAR_FLAG_SO, FAR_TIPDOC, FAR_FECHA_COMPRA,  far_moneda, far_precio_neto, far_subtra, far_numSER_c, far_numfac_c , far_numGUIA, FAR_DIAS, FAR_SERGUIA, FAR_CODUSU,FAR_CODART, FAR_CANTIDAD, FAR_JABAS, FAR_CONCEPTO, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMOPER2, FAR_COSPRO,FAR_PRECIO, FAR_EQUIV, FAR_FECHA, FAR_STOCK, FAR_ESTADO,FAR_CP, FAR_CODCLIE, FAR_TIPMOV, FAR_OTRA_CIA, far_SIGNO_aRM ,FAR_codven, far_costeo , far_hora FROM facart WHERE FAR_CODART = " & ENTRADA & " AND FAR_FECHA <= '" & Format(par_llave!PAR_FECHA_DIA, "dd/mm/yyyy") & "' AND FAR_CODCIA = '" & LK_CODCIA & "'  AND FAR_ESTADO <> 'E'  AND FAR_ESTADO <> 'M' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA DESC, FAR_SIGNO_ARM ,FAR_TIPMOV DESC, FAR_NUMOPER2 DESC "
          pub_cadena = "SELECT TOP " & Val(i_maximo.Text) & " *  FROM facart WHERE FAR_CODART = " & ENTRADA & " AND FAR_FECHA <= '" & Format(par_llave!PAR_FECHA_DIA, "dd/mm/yyyy") & "' AND FAR_CODCIA IN " & wol_codcia & "  AND FAR_ESTADO <> 'E'  AND FAR_ESTADO <> 'M' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L' ORDER BY FAR_FECHA_COMPRA DESC, FAR_SIGNO_ARM ,FAR_TIPMOV DESC, FAR_NUMOPER2 DESC , FAR_CODCIA DESC"
       End If
       
       Set far_codart = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
  
       'far_codart.Requery
    Else
       PSFAR_CODART.rdoParameters(0) = ENTRADA
       PSFAR_CODART.rdoParameters(1) = Format(WWFECHA1, "dd/mm/yyyy")
       PSFAR_CODART.rdoParameters(2) = Format(wwfecha2, "dd/mm/yyyy")
       PSFAR_CODART.rdoParameters(3) = xCODCIA
       far_codart.Requery
    End If
    
    FINAL = "N"
    If far_codart.EOF Then
       Screen.MousePointer = 0
       If Val(arm_llave!arm_stock) <> 0 Then
  ''        MsgBox "No hay movimientos a partir de esa fecha pero si hay STOCK ", 48, Pub_Titulo
       Else
''          MsgBox "No existen movimientos ...", 48, Pub_Titulo
       End If
       Exit Sub
    End If
    If option1(1).Value = True Then
       far_codart.MoveFirst
    Else
       far_codart.MoveLast
    End If

    pbValue = 0
    pbMin = 0
    pbMax = far_codart.RowCount
    LblMensaje.Visible = True
    pb.Visible = True
    wFAR_CODART = far_codart
    Band = "A"
  End If
cSaldo = 0
If FINAL = "N" Then
  KARDEX.GridK.Clear
  cabe
End If

PRE_UNIDADES(0) = LK_CODCIA
PRE_UNIDADES(1) = wFAR_CODART!far_codart
preuni_llave.Requery
ws_equiv = 0
If preuni_llave.EOF Then
   Screen.MousePointer = 0
   MsgBox "Unidad de Articulo no Esta Definido..", 48, Pub_Titulo
   Exit Sub
End If
ReDim EQUIV_UNI(preuni_llave.RowCount - 1)
xcuenta = 0
GoTo salta_unidad
  Do Until preuni_llave.EOF
    EQUIV_UNI(xcuenta) = preuni_llave!PRE_EQUIV
    If preuni_llave!pre_FLAG_UNIDAD = "A" Then
      ws_equiv = preuni_llave!PRE_EQUIV
      ACTIVA_EQUIV = preuni_llave!PRE_EQUIV
      LBLUNIDAD.Caption = Trim(preuni_llave!PRE_UNIDAD)
    End If
    xcuenta = xcuenta + 1
    preuni_llave.MoveNext
  Loop
salta_unidad:
ws_equiv = Val(Trim(Right(unidad.Text, 8)))
ACTIVA_EQUIV = ws_equiv
LBLUNIDAD.Caption = Trim(unidad.Text)

If LK_EMP = "HER" Then
 wSaldo = (wFAR_CODART!far_signo_arm * (wFAR_CODART!FAR_cantidad / ws_equiv)) + (wFAR_CODART!far_signo_arm * (wFAR_CODART!FAR_cantidad / ws_equiv) * -1)
Else
 wSaldo = (wFAR_CODART!FAR_STOCK / ws_equiv) + (wFAR_CODART!far_signo_arm * (wFAR_CODART!FAR_cantidad / ws_equiv) * -1)
End If

cSaldo = wSaldo * Nulo_Valor0(arm_llave!ARM_COSPRO)
If LK_FLAG_SOS = "A" Then NJ.Text = "A"
fila1 = 1

Do Until FINAL = "S"
  If Trim(NJ.Text) <> "" Then
     If Trim(NJ.Text) = "." And Trim(wFAR_CODART!FAR_FLAG_SO) <> "" Then GoTo pasa
     If Trim(NJ.Text) = "S" And Trim(wFAR_CODART!FAR_FLAG_SO) = "" Then GoTo pasa
  End If

  fila1 = fila1 + 1
  GridK.Rows = fila1 + 1
  GridK.RowHeight(GridK.Rows - 1) = 285
  GridK.Row = fila1
  KARDEX.GridK.TextMatrix(fila1, 0) = " | " & wFAR_CODART!FAR_CODCIA & " |   " & Format(wFAR_CODART!far_fecha, "dd/mm/yy")
  KARDEX.GridK.TextMatrix(fila1, 1) = Format(wFAR_CODART!FAR_fecha_compra, "dd/mm/yy")
  ws_nombre = wFAR_CODART!far_concepto
  KARDEX.GridK.TextMatrix(fila1, 2) = ws_nombre
  If wFAR_CODART!far_estado = "E" Then
     KARDEX.GridK.COL = 1
     KARDEX.GridK.CellBackColor = vb3DLight
  End If
  If wFAR_CODART!FAR_cp = "C" Then
    pu_cp = wFAR_CODART!FAR_cp
    pu_codclie = wFAR_CODART!far_codclie
    SQ_OPER = 1
    pu_codcia = wFAR_CODART!FAR_CODCIA ' xCODCIA
    LEER_CLI_LLAVE
    KARDEX.GridK.TextMatrix(fila1, 15) = cli_llave(3)
    ws_nombre = Left(Trim(cli_llave!cli_nombre), 25)
    If Band = "D" Then
       KARDEX.i_nomarti.Caption = Trim(cli_llave(3))
       PUB_CODART = Val(wFAR_CODART!far_codart)
       PSART_LLAVE.rdoParameters(0) = PUB_CODART
       art_LLAVE.Requery
       KARDEX.GridK.TextMatrix(fila1, 2) = art_LLAVE(2)
    End If
  ElseIf wFAR_CODART!FAR_cp = "P" Then
    pu_cp = wFAR_CODART!FAR_cp
    pu_codclie = wFAR_CODART!far_codclie
    SQ_OPER = 1
    pu_codcia = wFAR_CODART!FAR_CODCIA ' xCODCIA
    LEER_CLI_LLAVE
    If cli_llave.EOF = False Then
       KARDEX.GridK.TextMatrix(fila1, 15) = cli_llave(3)
       ws_nombre = Left(Trim(cli_llave!cli_nombre), 25)
    Else
       KARDEX.GridK.TextMatrix(fila1, 15) = " "
    End If
 End If
  
 PRECIO = wFAR_CODART!FAR_COSPRO
 If (wFAR_CODART!FAR_tipmov = 20 Or wFAR_CODART!FAR_tipmov = 101 Or (wFAR_CODART!FAR_tipmov = 5 And wFAR_CODART!FAR_COSTEO = "A")) And wFAR_CODART!far_estado <> "E" Then
    PRECIO = Nulo_Valor0(wFAR_CODART!FAR_PRECIO) / wFAR_CODART!FAR_cantidad
 End If
   
 If wFAR_CODART!far_signo_arm = 1 Then
    KARDEX.GridK.TextMatrix(fila1, 5) = Format(wFAR_CODART!far_JABAS, "0")
    KARDEX.GridK.TextMatrix(fila1, 8) = Format(wFAR_CODART!FAR_cantidad / ws_equiv, "#####.00")
    cIngreso = KARDEX.GridK.TextMatrix(fila1, 8) * PRECIO
    KARDEX.GridK.TextMatrix(fila1, 11) = Format(cIngreso, "#####.00")
'    GridK.Visible = True
    cSaldo = cSaldo + cIngreso
    wSaldo = wSaldo + (wFAR_CODART!FAR_cantidad / ws_equiv)
 ElseIf wFAR_CODART!far_signo_arm = -1 Then
    KARDEX.GridK.TextMatrix(fila1, 6) = Format(wFAR_CODART!far_JABAS, "0")
    KARDEX.GridK.TextMatrix(fila1, 9) = Format(wFAR_CODART!FAR_cantidad / ws_equiv, "#####.00")
    cSalida = (wFAR_CODART!FAR_cantidad) * PRECIO
    KARDEX.GridK.TextMatrix(fila1, 12) = Format(cSalida, "#####.00")
    cSaldo = cSaldo - cSalida
    wSaldo = wSaldo - (wFAR_CODART!FAR_cantidad / ws_equiv)
    If LK_FLAG_GRIFO <> "A" Then
'     SQ_OPER = 1
'     PUB_CODCIA = xCODCIA
     PUB_CODVEN = Nulo_Valor0(wFAR_CODART!FAR_CODVEN)
'     LEER_VEN_LLAVE
'     If Not ven_llave.EOF Then
          KARDEX.GridK.TextMatrix(fila1, 18) = PUB_CODVEN ' ven_llave!VEM_NOMBRE
'     End If
    Else
      SQ_OPER = 1
      PUB_CODCIA = wFAR_CODART!FAR_CODCIA ' xCODCIA
      If LK_EMP <> "HER" Then
       PUB_CODVEN = Nulo_Valor0(wFAR_CODART!far_ISLA)
      Else
       PUB_CODVEN = wFAR_CODART!FAR_CODVEN
      End If
      LEER_VEN_LLAVE
      WISLA = " "
      If Not ven_llave.EOF Then
          WISLA = ven_llave!VEM_NOMBRE
      End If
      SQ_OPER = 1
      PUB_TIPREG = 2103
      PUB_CODCIA = LK_CODCIA
      PUB_NUMTAB = wFAR_CODART!FAR_CODVEN
      LEER_TAB_LLAVE
      WRES = " "
      If Not tab_llave.EOF Then
        WRES = Trim(tab_llave!TAB_NOMLARGO)
      End If
      KARDEX.GridK.TextMatrix(fila1, 18) = WISLA & " / " & WRES
    End If
 End If

 
 GridK.COL = 10
  If wFAR_CODART!far_signo_arm = 1 Then
  ' cSaldo = WSALDO * Nulo_Valor0(arm_llave!ARM_COSPRO)
  End If

 KARDEX.GridK.TextMatrix(fila1, 13) = Format(cSaldo, "#####.0000")
 GridK.COL = 11
 KARDEX.GridK.TextMatrix(fila1, 14) = Format(wFAR_CODART!FAR_COSPRO * ws_equiv, "####.0000")
 ' GridK.Visible = True
 GridK.COL = 7
 If wSaldo < 0 Then
    'GridK.CellBackColor = QBColor(12)
 End If
 
 KARDEX.GridK.TextMatrix(fila1, 10) = Format((wSaldo), "#####.00")  'Format((wFAR_CODART!far_stock / WS_EQUIV), "#####.00")
 If redondea(wSaldo) <> redondea((wFAR_CODART!FAR_STOCK / ws_equiv)) Then
    wSaldo = wFAR_CODART!FAR_STOCK
'    MsgBox "REVISAR SALDO "
 End If
 
' GridK.COL = 11
' KARDEX.GridK.TextMatrix(fila1, 14) = Format(wFAR_CODART!FAR_COSPRO, "####.0000")
 GridK.COL = 12
 If wFAR_CODART!FAR_tipmov = 20 Or wFAR_CODART!FAR_COSTEO = "A" Then
    ' KARDEX.GridK.TextMatrix(fila1, 15) = Format(wFAR_CODART!far_precio_neto / (wFAR_CODART!far_cantidad), "####.0000")
    'KARDEX.GridK.TextMatrix(fila1, 15) = Format(wFAR_CODART!far_precio / (wFAR_CODART!FAR_equiv), "####.0000")
    KARDEX.GridK.TextMatrix(fila1, 15) = Format(wFAR_CODART!FAR_PRECIO * (wFAR_CODART!FAR_cantidad / wFAR_CODART!FAR_equiv), "####.0000")

 Else
    KARDEX.GridK.TextMatrix(fila1, 15) = Format(wFAR_CODART!FAR_PRECIO / wFAR_CODART!FAR_equiv, "####.0000")
    If wFAR_CODART!FAR_MONEDA = "S" Then
       KARDEX.GridK.TextMatrix(fila1, 15) = "S/. " & KARDEX.GridK.TextMatrix(fila1, 15)
    ElseIf wFAR_CODART!FAR_MONEDA = "D" Then
       KARDEX.GridK.TextMatrix(fila1, 15) = "US$." & KARDEX.GridK.TextMatrix(fila1, 15)
    End If
   
 End If
 KARDEX.GridK.TextMatrix(fila1, 19) = Trim(wFAR_CODART!FAR_DIAS)
 KARDEX.GridK.TextMatrix(fila1, 20) = Nulo_Valors(wFAR_CODART!far_fbg) & ":" & Trim(wFAR_CODART!far_numser) & "-" & Trim(wFAR_CODART!far_numfac)
 KARDEX.GridK.TextMatrix(fila1, 21) = Nulo_Valors(wFAR_CODART!far_codusu) & " " & (wFAR_CODART!far_hora)

 If wFAR_CODART!FAR_tipmov = 20 Or wFAR_CODART!FAR_tipmov = 180 Then
    If Val(Nulo_Valor0(wFAR_CODART!FAR_NUMFAC_C)) <> 0 Then
       vdocum = "Fact." & wFAR_CODART!FAR_NUMSER_C & "-" & wFAR_CODART!FAR_NUMFAC_C
    Else
       vdocum = "Guia " & wFAR_CODART!far_numguia
    End If
    ws_nombre = Left(Trim(cli_llave!cli_nombre), 25)
 ElseIf wFAR_CODART!FAR_tipmov = 10 Then
    If wFAR_CODART!far_fbg = "F" Then
       vdocum = "Fact." & wFAR_CODART!far_numser & "-" & wFAR_CODART!far_numfac
    ElseIf wFAR_CODART!far_fbg = "B" Then
       vdocum = "Bol. " & wFAR_CODART!far_numser & "-" & wFAR_CODART!far_numfac
    ElseIf wFAR_CODART!far_fbg = "G" Then
      If LK_FLAG_GRIFO = "A" Then
        vdocum = wFAR_CODART!FAR_TIPDOC & " " & wFAR_CODART!far_numser & "-" & wFAR_CODART!far_numfac
      Else
        vdocum = "Guia " & wFAR_CODART!far_numser & "-" & wFAR_CODART!far_numfac
      End If
    ElseIf wFAR_CODART!far_fbg = "P" Then
     If LK_FLAG_GRIFO = "A" Then
        vdocum = wFAR_CODART!FAR_TIPDOC & " " & wFAR_CODART!far_numser & "-" & wFAR_CODART!far_numfac
      Else
       vdocum = "Guia: " & wFAR_CODART!far_numser & "-" & wFAR_CODART!far_numfac
      End If
    ElseIf wFAR_CODART!far_fbg = "X" Then
        vdocum = "Doc.: " & wFAR_CODART!far_numser & "-" & wFAR_CODART!far_numfac
    Else
       vdocum = wFAR_CODART!FAR_TIPDOC & wFAR_CODART!far_numser & "-" & wFAR_CODART!far_numfac
    End If
 ElseIf wFAR_CODART!FAR_tipmov = 97 Or wFAR_CODART!FAR_tipmov = 98 Then
    If wFAR_CODART!FAR_cp = "P" Then
      If wFAR_CODART!FAR_tipmov = 97 Then
       vdocum = "N.Cred." & wFAR_CODART!FAR_NUMSER_C & "-" & wFAR_CODART!FAR_NUMFAC_C
      Else
       vdocum = "N.Deb." & wFAR_CODART!FAR_NUMSER_C & "-" & wFAR_CODART!FAR_NUMFAC_C
      End If
    Else
      If wFAR_CODART!FAR_tipmov = 97 Then
       vdocum = "N.Cred." & wFAR_CODART!far_numser & "-" & wFAR_CODART!far_numfac
      Else
       vdocum = "N.Deb." & wFAR_CODART!far_numser & "-" & wFAR_CODART!far_numfac
      End If
    End If
 Else
    vdocum = "Guia " & wFAR_CODART!far_numser & "-" & wFAR_CODART!far_numfac
 End If
 
 GridK.COL = 8
 GridK.CellAlignment = 7
 GridK.COL = 9
 GridK.CellAlignment = 7
 GridK.COL = 10
 GridK.CellAlignment = 7
 GridK.COL = 11
 GridK.CellAlignment = 7
 GridK.COL = 12
 GridK.CellAlignment = 7
 GridK.COL = 13
 GridK.CellAlignment = 7
 GridK.COL = 14
 GridK.CellAlignment = 7
 GridK.COL = 15
 GridK.CellAlignment = 7
  
 KARDEX.GridK.TextMatrix(fila1, 2) = vdocum

If wFAR_CODART!FAR_tipmov = 10 And Val(wFAR_CODART!far_numguia) <> 0 Then
    KARDEX.GridK.TextMatrix(fila1, 3) = " G. Rems.:" & wFAR_CODART!far_serguia & "-" & wFAR_CODART!far_numguia & "  / " & Nulo_Valors(wFAR_CODART!far_subtra)
Else
    KARDEX.GridK.TextMatrix(fila1, 3) = Nulo_Valors(wFAR_CODART!far_subtra)
End If
If LK_FLAG_GRIFO = "A" Then
     KARDEX.GridK.TextMatrix(fila1, 3) = "Turno: " & Str(wFAR_CODART!far_turno) + " " + KARDEX.GridK.TextMatrix(fila1, 3)
End If
 
  
 KARDEX.GridK.TextMatrix(fila1, 4) = ws_nombre
 If Val(far_codart!FAR_STOCK) <> Val(wSaldo) Then
 '   MsgBox "Falta procesar costeo"
 '  far_codart.Edit
  ' far_codart!FAR_STOCK = WSALDO
 '  far_codart.Update
 End If
   
If Check1.Visible = True And Check1.Value = 1 And WWW_COSPRO <> 0 Then
   far_codart.Edit
   far_codart!FAR_COSPRO = WWW_COSPRO
   far_codart.Update
End If
PUB_COSPRO = far_codart!FAR_COSPRO

pasa:
    If option1(1).Value = True Then
       far_codart.MoveNext
       If far_codart.EOF Then
           FINAL = "S"
       Else
          If i_maximo.Visible Then
            If fila1 > i_maximo.Text Then
               pub_mensaje = "Hay más registros... desea ver mas Registros ?"
               Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
               If Pub_Respuesta = vbYes Then   ' El usuario eligió
                  i_maximo.Text = i_maximo.Text + 100
               Else
                 FINAL = "S"
               End If
            End If
          End If
       End If
    Else
       far_codart.MovePrevious
'       pbValue = fila1
       If far_codart.BOF Then FINAL = "S"
    End If

    
    
ws_nombre = ""
Loop
PSFAR_PEDIDOS(0) = LK_CODCIA
PSFAR_PEDIDOS(1) = ENTRADA
far_pedidos.Requery
If far_pedidos.EOF Then GoTo salta_ped
Do Until far_pedidos.EOF
 fila1 = fila1 + 1
 GridK.Rows = fila1 + 1
 GridK.RowHeight(GridK.Rows - 1) = 285
 GridK.Row = fila1
 pu_cp = "C"
 pu_codclie = far_pedidos!PED_CODCLIE
 SQ_OPER = 1
 pu_codcia = LK_CODCIA
 LEER_CLI_LLAVE
 If cli_llave.EOF = False Then
      KARDEX.GridK.TextMatrix(fila1, 4) = Trim(cli_llave(3))
 Else
       KARDEX.GridK.TextMatrix(fila1, 4) = " "
 End If
 KARDEX.GridK.TextMatrix(fila1, 0) = Format(far_pedidos!PED_FECHA, "dd/mm/yy")
 KARDEX.GridK.TextMatrix(fila1, 1) = Format(far_pedidos!PED_FECHA, "dd/mm/yy")
 KARDEX.GridK.TextMatrix(fila1, 2) = "Ped. " & Format(far_pedidos!PED_NUMFAC, "000000")
 KARDEX.GridK.TextMatrix(fila1, 3) = "Cod.Vend: " & Format(far_pedidos!ped_codven, "00")
 
 KARDEX.GridK.TextMatrix(fila1, 9) = Format(far_pedidos!PED_CANTIDAD / ws_equiv, "0.00")
 GridK.COL = 9
 GridK.CellAlignment = 7
 far_pedidos.MoveNext
Loop
  
  
  

salta_ped:

If redondea(wSaldo) <> redondea((arm_llave!arm_stock / ws_equiv)) Then
'   MsgBox "AVISAR A COMPUTO..."
End If
'If LK_FLAG_GRIFO = "A" Then
Dim wst As Currency
pub_cadena = "SELECT SUM(ARM_STOCK) as stock FROM ARTICULO WHERE ARM_CODART = " & WCOD_ORIGINAL & " AND ARM_CODCIA IN " & wol_codcia
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If X.EOF Then
wst = 0
Else
wst = Val(Nulo_Valor0(X!stock))
End If
    
txtsaldo.Text = ""

If option1(0).Value = True Or option1(1).Value = True Then
   wSaldo = wst / ws_equiv
   WSALDO2 = 0 'arm_llave!arm_stock2 / ws_equiv
 
   If NJ.Text = "." Then
      wSaldo = wst / ws_equiv
      WSALDO2 = 0 ' arm_llave!ARM_SALDO_N2 / ws_equiv
   ElseIf NJ.Text = "S" Then
      wSaldo = wst / ws_equiv
      WSALDO2 = 0 'arm_llave!ARM_SALDO_S2 / WS_EQUIV
   End If
   txtsaldo.Text = wSaldo
End If

cSaldo = wSaldo * PUB_COSPRO
fila1 = GridK.Rows - 1
Do Until fila1 = 0
   'If fila1 = 3 Then Stop
    If fila1 = 1 Then wSaldo = Format(wSaldo, "0.00")
    KARDEX.GridK.TextMatrix(fila1, 7) = WSALDO2
    KARDEX.GridK.TextMatrix(fila1, 10) = wSaldo
    If cSaldo < 0 And fila1 = 1 Then cSaldo = 0
    KARDEX.GridK.TextMatrix(fila1, 13) = cSaldo
    wSaldo = wSaldo - Val(KARDEX.GridK.TextMatrix(fila1, 8)) + Val(KARDEX.GridK.TextMatrix(fila1, 9))
    cSaldo = cSaldo - Val(KARDEX.GridK.TextMatrix(fila1, 11)) + Val(KARDEX.GridK.TextMatrix(fila1, 12))
    WSALDO2 = WSALDO2 - Val(KARDEX.GridK.TextMatrix(fila1, 5)) + Val(KARDEX.GridK.TextMatrix(fila1, 6))
    fila1 = fila1 - 1
Loop
'End If
KARDEX.GridK.TextMatrix(1, 4) = "Saldo Inicial "

GoTo SALTARIN









'KARDEX.Nombre.Visible = True

SALTARIN:

If Check1.Visible = True And Check1.Value = 1 And WWW_COSPRO <> 0 Then
 'SQ_OPER = 1
 'pu_codcia = LK_CODCIA
 'PUB_CODART = ENTRADA
 'LEER_ARM_LLAVE
 'If arm_llave.EOF Then
 '  Screen.MousePointer = 0
 '  MsgBox "Codigo NO Existe en Articulos , Verificar....", 48, Pub_Titulo
 '  i_codart2.SetFocus
 '  Exit Sub
 'End If
 'arm_llave.Edit
 'arm_llave!ARM_COSPRO = WWW_COSPRO
 'arm_llave.Update
 Check1.Value = 0
End If

LblMensaje.Visible = False
pb.Visible = False
DoEvents
GridK.Visible = True
KARDEX.GridK.COL = 1
KARDEX.GridK.Row = 1
GridK.SetFocus
KARDEX.CmdProcesa.Enabled = True
If ff_flag = "A" Then
 If (KARDEX.GridK.TextMatrix(1, 10)) <> 0 Then
   pub_cadena = pub_cadena + Str(ENTRADA) + " "
 End If
End If

sigue:

Screen.MousePointer = 1
Exit Sub
osERRsale:
  MsgBox "Verificar, " & Err.Description, 48, Pub_Titulo
  Screen.MousePointer = 0

End Sub

Private Sub Command1_Click()
Dim Xx As rdoQuery
Dim ver_xx As rdoResultset

pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = '" & LK_CODCIA & "'AND ART_CALIDAD = 1 ORDER BY ART_ALTERNO"
Set Xx = CN.CreateQuery("", pub_cadena)
Set ver_xx = Xx.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
ver_xx.Requery
Do Until ver_xx.EOF
  i_codart2.Text = Trim(ver_xx!ART_alterno)
  i_codart2_KeyPress 13
  CmdProcesa_Click
'  MsgBox " PARA O SIGUE"
  ver_xx.MoveNext
Loop



End Sub

Private Sub Form_Activate()
'ICA
If LK_EMP = "PIU" Or LK_EMP = "HER" Then
   cheext.Value = 0
End If
If i_codart2.Visible Then Azul i_codart2, i_codart2
End Sub

Private Sub Form_Load()
Dim fech As String
Dim wwano As Integer
pub_cadena = "SELECT PED_FECHA,PED_NUMSER ,PED_NUMFAC, PED_CONDI,PED_CODVEN , PED_CODART , PED_CANTIDAD, PED_EQUIV , PED_CODCLIE FROM PEDIDOS " & _
"WHERE PED_CODCIA = ? AND PED_TIPMOV = 201 AND PED_CODART = ? AND PED_SITUACION not in ('P','E')  "
Set PSFAR_PEDIDOS = CN.CreateQuery("", pub_cadena)
PSFAR_PEDIDOS(0) = 0
PSFAR_PEDIDOS(1) = 0
Set far_pedidos = PSFAR_PEDIDOS.OpenResultset(rdOpenKeyset, rdConcurReadOnly)



' ICA
txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtCampo1.Mask = "##/##/####"
txtCampo1.Visible = True
txtCampo2.Mask = "##/##/####"
txtCampo2.Visible = True


If LK_FLAG_SOS = "A" Then NJ.Visible = False

If LK_CODUSU = "ADMIN" Then Check1.Visible = True
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE
If par_llave!par_flag_costos <> 9 Then
  MsgBox "Hay Cambios en Mercaderia, se suguiere hacer Proceso de Costeo", 48, Pub_Titulo
End If
KARDEX.GridK.Cols = 22
KARDEX.GridK.ColWidth(0) = 1300
KARDEX.GridK.ColWidth(1) = 900
KARDEX.GridK.ColWidth(2) = 1200
KARDEX.GridK.ColWidth(3) = 1100
KARDEX.GridK.ColWidth(4) = 2100
'If lk_jabas = "A" Then
'   KARDEX.GridK.ColWidth(5) = 700
'   KARDEX.GridK.ColWidth(6) = 700 '
'   KARDEX.GridK.ColWidth(7) = 700
'Else
   KARDEX.GridK.ColWidth(5) = 0
   KARDEX.GridK.ColWidth(6) = 0  '
   KARDEX.GridK.ColWidth(7) = 0
'End If


KARDEX.GridK.ColWidth(8) = 1000
KARDEX.GridK.ColWidth(9) = 1000 '
KARDEX.GridK.ColWidth(10) = 1000
KARDEX.GridK.ColWidth(11) = 1000
KARDEX.GridK.ColWidth(12) = 1000
KARDEX.GridK.ColWidth(13) = 1000
KARDEX.GridK.ColWidth(14) = 1200
KARDEX.GridK.ColWidth(15) = 1000
KARDEX.GridK.ColWidth(16) = 0
KARDEX.GridK.ColWidth(17) = 0
KARDEX.GridK.ColWidth(18) = 2000
KARDEX.GridK.ColWidth(19) = 900
KARDEX.GridK.ColWidth(21) = 2200

i_maximo.ListIndex = 0
cabe
SW = 0
UNICO = ""
ACV_TEXTO = "MOVIMIENTO DE KARDEX"
Dim VARX As String
KARDEX.GridK.Visible = False
Dim r As Integer
If LK_FLAG_GRIFO = "A" Then

End If
  

pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ?  ORDER BY PRE_CODART"
Set PRE_UNIDADES = CN.CreateQuery("", pub_cadena)
PRE_UNIDADES(0) = 0
PRE_UNIDADES(1) = 0
Set preuni_llave = PRE_UNIDADES.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
'Fecha


i_codart2.MaxLength = 20
Option1_Click (0)
'wwano = DatePart("yyyy", LK_FECHA_DIA)
'wwano = wwano - 15
'i_ano.Clear
'Do Until wwano = DatePart("yyyy", LK_FECHA_DIA) + 1
'   i_ano.AddItem wwano
'   wwano = wwano + 1
'Loop
'i_ano.ListIndex = i_ano.ListCount - 1
'i_mes.ListIndex = DatePart("m", LK_FECHA_DIA)

End Sub

Private Sub GRIDK_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 i_codart2.SetFocus
 Exit Sub
End If

If KeyAscii <> 13 Then
 Exit Sub
End If

Dim w_equiv As Currency
Dim w_valor As String
Dim w_valor2 As String
Dim wentero
Dim wdecimal
Dim wmensa
Dim ASA
If Trim(GridK.Text) = "" Then
 Exit Sub
End If
w_equiv = ACTIVA_EQUIV
If GridK.COL = 3 Or GridK.COL = 2 Then
 ASA = w_equiv * Val(GridK.Text)
 w_valor = Format(ASA, "######")
 MsgBox "Unidades  : " & w_valor
 GridK.SetFocus
End If
'GridK.text = "0.83"
If GridK.COL = 4 Then
   w_equiv = Val(GridK.Text) Mod ACTIVA_EQUIV
   If InStr(1, GridK.Text, ".") <> 0 Then
     wdecimal = Mid(GridK.Text, InStr(1, GridK.Text, "."), Len(GridK.Text))
     wentero = Mid(GridK.Text, 1, InStr(1, GridK.Text, ".") - 1)
     w_equiv = Int(wdecimal * ACTIVA_EQUIV)
     w_valor = wentero
     wmensa = w_valor & " " & Trim(LBLUNIDAD.Caption) & "   y " & w_equiv & " Unidades "
     MsgBox wmensa
   Else
     w_valor = w_equiv * ACTIVA_EQUIV
     MsgBox w_valor & " Unidades  "
   End If
 GridK.SetFocus
End If

End Sub

Private Sub i_codart2_Change()
If i_codart2.Text = "" Then
 i_nomarti.Caption = ""
  VAR_ACTIVAR = 0
  KARDEX.GridK.Visible = False
  LBLUNIDAD.Caption = ""
  txtsaldo.Text = ""
End If

End Sub

Private Sub i_codart2_GotFocus()

KARDEX.GridK.Visible = False

'i_codart2.Text = ""
'i_nomarti.Caption = ""
'lblmoneda.Caption = ""
'txtsaldo.Text = ""

End Sub
Private Sub i_codart2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If KeyCode = 113 Then
  If option1(0).Value Then
    option1(0).Value = False
    option1(1).Value = True
    option1(2).Value = False
'    Option1_Click 1
    Exit Sub
  End If
  If option1(1).Value Then
    option1(0).Value = False
    option1(1).Value = False
    option1(2).Value = True
    Exit Sub
 '   Option1_Click 2
  End If
  If option1(2).Value Then
    option1(0).Value = True
    option1(1).Value = False
    option1(2).Value = False
    Exit Sub
  '  Option1_Click 0
  End If
End If
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
  If Frame4.Visible = True Then
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
        i_nomarti.Caption = art_LLAVE!art_nombre
     End If

     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     LEER_ARM_LLAVE
     If Not arm_llave.EOF Then txtsaldo.Text = arm_llave!arm_stock
     If Frame4.Visible = False Then Frame4.Visible = True
  End If
  
  DoEvents
fin:

End Sub
Private Sub i_codart2_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I, car
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
    i_nomarti.Caption = Trim(art_LLAVE!art_nombre)
    'i_codart2.text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
    If art_LLAVE!ART_MONEDA = "D" Then
      lblMoneda.Caption = "Moneda : US$."
    Else
      lblMoneda.Caption = "Moneda : S/ ."
    End If
    PRE_UNIDADES(0) = LK_CODCIA
    PRE_UNIDADES(1) = art_LLAVE!art_key
    preuni_llave.Requery
    unidad.Clear
    unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!PRE_UNIDAD & String(80, " ") & preuni_llave!PRE_EQUIV
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
         unid_max = preuni_llave.AbsolutePosition - 1
         LBLUNIDAD.Caption = Trim(preuni_llave!PRE_UNIDAD)
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
     i_nomarti.Caption = Trim(art_llave_alt!art_nombre)
     ListView1.Visible = False
     If txtCampo1.Visible Then
       Azul2 txtCampo1, txtCampo1
       Exit Sub
     End If
     PRE_UNIDADES(0) = LK_CODCIA
     PRE_UNIDADES(1) = art_llave_alt!art_key
     preuni_llave.Requery
     unidad.Clear
     unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!PRE_UNIDAD & String(80, " ") & preuni_llave!PRE_EQUIV
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
          unid_max = preuni_llave.AbsolutePosition - 1
          LBLUNIDAD.Caption = Trim(preuni_llave!PRE_UNIDAD)
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
      i_nomarti.Caption = Trim(art_LLAVE!art_nombre)
      If art_LLAVE!ART_MONEDA = "D" Then
        lblMoneda.Caption = "Moneda : US$."
      Else
        lblMoneda.Caption = "Moneda : S/ ."
      End If
    PRE_UNIDADES(0) = LK_CODCIA
    PRE_UNIDADES(1) = art_LLAVE!art_key
    preuni_llave.Requery
    unidad.Clear
    unid_max = 0
    Do Until preuni_llave.EOF
       unidad.AddItem preuni_llave!PRE_UNIDAD & String(80, " ") & preuni_llave!PRE_EQUIV
       If preuni_llave!pre_FLAG_UNIDAD = "A" Then
         unid_max = preuni_llave.AbsolutePosition - 1
         LBLUNIDAD.Caption = Trim(preuni_llave!PRE_UNIDAD)
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
      If chedes.Value = 1 Then
        archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
      Else
        archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CODCIA = '" & ws_codcia & "' AND  ART_SITUACION <> 1 AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
      End If
    Else
      numarchi = 0
      If chedes.Value = 1 Then
       archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      Else
       archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND  ART_CODCIA = '" & ws_codcia & "' AND ART_SITUACION <> 1 AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
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


Private Sub i_fecha_GotFocus()
'Azul2 i_fecha, i_fecha
GridK.Visible = False
CmdProcesa.Enabled = True
End Sub

Private Sub i_maximo_Click()
GridK.Visible = False
CmdProcesa.Enabled = True
End Sub

Private Sub i_maximo_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then i_codart2.SetFocus
If KeyAscii <> 13 Then Exit Sub
If CmdProcesa.Enabled And CmdProcesa.Visible Then CmdProcesa_Click
End Sub

Private Sub ListView1_DblClick()
 loc_key = ListView1.SelectedItem.Index
 i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 i_codart2_KeyPress 13
End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.Item(loc_key).Selected = True
 ListView1.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView1.Visible = False
 i_codart2.Text = ""
 i_codart2.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
ListView1_DblClick
End Sub

Private Sub ListView1_LostFocus()
ListView1.Visible = False
End Sub

Private Sub NJ_Change()
If Trim(NJ.Text) <> "" Then i_maximo.ListIndex = 3
End Sub

Private Sub Option1_Click(Index As Integer)
ListView1.Visible = False
'i_codart2.Text = ""
txtsaldo.Text = ""

If option1(0).Value = True Then
'   txtsaldo.Text = ""
   Frame2.Visible = True
   Frame1.Visible = False
   'Frame4.Visible = False
'   Frame2.Top = todos.Top + 600
'   Frame2.Left = todos.Left + 2300
  
   
   
   If cheext.Value = 1 Then
     'pub_cadena = "SELECT FAR_TRANSITO , FAR_FLAG_SO,  FAR_TIPDOC, FAR_FECHA_COMPRA, far_moneda, far_precio_neto, far_subtra, far_numSER_c, far_numfac_c , far_numGUIA, FAR_SERGUIA, FAR_DIAS , FAR_CODUSU,FAR_CODART, FAR_CANTIDAD, FAR_JABAS, FAR_CONCEPTO, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMOPER2, FAR_COSPRO,FAR_PRECIO, FAR_EQUIV, FAR_FECHA, FAR_STOCK, FAR_ESTADO,FAR_CP, FAR_CODCLIE, FAR_TIPMOV, FAR_OTRA_CIA, far_SIGNO_aRM ,FAR_codven, far_costeo, far_hora  FROM facart WHERE FAR_CODART = ? AND FAR_FECHA <= ? AND FAR_CODCIA = ? AND FAR_ESTADO <> 'M' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA DESC, FAR_SIGNO_ARM ,FAR_TIPMOV DESC, FAR_NUMOPER2 DESC "
     pub_cadena = "SELECT TOP " & Val(i_maximo.Text) & " FAR_TRANSITO, FAR_FLAG_SO, FAR_TIPDOC, FAR_FECHA_COMPRA,  far_moneda, far_precio_neto, far_subtra, far_numSER_c, far_numfac_c , far_numGUIA, FAR_DIAS, FAR_SERGUIA, FAR_CODUSU,FAR_CODART, FAR_CANTIDAD, FAR_JABAS, FAR_CONCEPTO, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMOPER2, FAR_COSPRO,FAR_PRECIO, FAR_EQUIV, FAR_FECHA, FAR_STOCK, FAR_ESTADO,FAR_CP, FAR_CODCLIE, FAR_TIPMOV, FAR_OTRA_CIA, far_SIGNO_aRM ,FAR_codven, far_costeo , far_hora FROM facart WHERE FAR_CODART = ? AND FAR_FECHA <= ? AND FAR_CODCIA = ?  AND FAR_ESTADO <> 'M' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA DESC, FAR_SIGNO_ARM ,FAR_TIPMOV DESC, FAR_NUMOPER2 DESC "
   Else
      pub_cadena = "SELECT TOP " & Val(i_maximo.Text) & " FAR_TRANSITO, FAR_FLAG_SO, FAR_TIPDOC, FAR_FECHA_COMPRA,  far_moneda, far_precio_neto, far_subtra, far_numSER_c, far_numfac_c , far_numGUIA, FAR_DIAS, FAR_SERGUIA, FAR_CODUSU,FAR_CODART, FAR_CANTIDAD, FAR_JABAS, FAR_CONCEPTO, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMOPER2, FAR_COSPRO,FAR_PRECIO, FAR_EQUIV, FAR_FECHA, FAR_STOCK, FAR_ESTADO,FAR_CP, FAR_CODCLIE, FAR_TIPMOV, FAR_OTRA_CIA, far_SIGNO_aRM ,FAR_codven, far_costeo , far_hora FROM facart WHERE FAR_CODART = ? AND FAR_FECHA <= ? AND FAR_CODCIA = ? AND FAR_ESTADO <> 'E'  AND FAR_ESTADO <> 'M' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA DESC, FAR_SIGNO_ARM ,FAR_TIPMOV DESC, FAR_NUMOPER2 DESC "
   End If
   If LK_FLAG_SOS = "A" Then
      If cheext.Value = 1 Then
         pub_cadena = "SELECT FAR_TRANSITO, FAR_FLAG_SO, FAR_TIPDOC, FAR_FECHA_COMPRA, far_moneda, far_precio_neto, far_subtra, far_numSER_c, far_numfac_c , far_numGUIA, FAR_SERGUIA, FAR_DIAS , FAR_CODUSU,FAR_CODART, FAR_CANTIDAD, FAR_JABAS, FAR_CONCEPTO, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMOPER2, FAR_COSPRO,FAR_PRECIO, FAR_EQUIV, FAR_FECHA, FAR_STOCK, FAR_ESTADO,FAR_CP, FAR_CODCLIE, FAR_TIPMOV, FAR_OTRA_CIA, far_SIGNO_aRM ,FAR_codven, far_costeo  FROM facart WHERE FAR_CODART = ? AND FAR_FECHA_COMPRA <= ? AND FAR_CODCIA = ? AND FAR_ESTADO <> 'M' AND FAR_FLAG_SO = 'A' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L'  ORDER BY FAR_CODART, FAR_FECHA_COMPRA desc , FAR_NUMOPER2 desc, FAR_NUMSEC DESC "
      Else
         pub_cadena = "SELECT FAR_TRANSITO, FAR_FLAG_SO,  FAR_TIPDOC, FAR_FECHA_COMPRA, far_moneda, far_precio_neto, far_subtra, far_numSER_c, far_numfac_c , far_numGUIA, FAR_DIAS, FAR_SERGUIA, FAR_CODUSU,FAR_CODART, FAR_CANTIDAD, FAR_JABAS, FAR_CONCEPTO, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMOPER2, FAR_COSPRO,FAR_PRECIO, FAR_EQUIV, FAR_FECHA, FAR_STOCK, FAR_ESTADO,FAR_CP, FAR_CODCLIE, FAR_TIPMOV, FAR_OTRA_CIA, far_SIGNO_aRM ,FAR_codven, far_costeo  FROM facart WHERE FAR_CODART = ? AND FAR_FECHA_COMPRA <= ? AND FAR_CODCIA = ? AND FAR_ESTADO <> 'E'  AND FAR_ESTADO <> 'M'  AND FAR_FLAG_SO = 'A' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L'  ORDER BY FAR_CODART, FAR_FECHA_COMPRA desc , FAR_NUMOPER2 desc, FAR_NUMSEC DESC "
      End If
   End If
   
   Set PSFAR_CODART = CN.CreateQuery("", pub_cadena)
   PSFAR_CODART(0) = 0
   PSFAR_CODART(1) = LK_FECHA_DIA
   PSFAR_CODART(2) = LK_FECHA_DIA
   If Frame4.Visible = True Then i_codart2.SetFocus
ElseIf option1(1).Value = True Then
   If cheext.Value = 1 Then
      pub_cadena = "SELECT FAR_TRANSITO, FAR_FLAG_SO, FAR_TIPDOC, FAR_FECHA_COMPRA, far_moneda,far_precio_neto, far_subtra, far_numSER_c, far_numfac_c , far_numGUIA, FAR_DIAS , FAR_CODUSU, FAR_SERGUIA, FAR_CODUSU, FAR_HORA,FAR_CODART, FAR_CANTIDAD, FAR_JABAS, FAR_CONCEPTO, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMOPER2, FAR_COSPRO,FAR_PRECIO, FAR_EQUIV, FAR_FECHA, FAR_STOCK, FAR_ESTADO,FAR_CP, FAR_CODCLIE, FAR_TIPMOV, FAR_OTRA_CIA, far_SIGNO_aRM ,FAR_codven, far_costeo  FROM facart WHERE FAR_CODART = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ?  AND FAR_CODCIA = ? AND FAR_ESTADO <> 'M' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L'  ORDER BY FAR_CODART, FAR_FECHA asc , FAR_NUMOPER  asc , FAR_NUMSEC "
   Else
      pub_cadena = "SELECT FAR_TRANSITO, FAR_FLAG_SO, FAR_TIPDOC, FAR_FECHA_COMPRA, far_moneda,far_precio_neto, far_subtra, far_numSER_c, far_numfac_c , far_numGUIA, FAR_DIAS , FAR_CODUSU,FAR_CODART, FAR_CODUSU, FAR_HORA, FAR_CANTIDAD,  FAR_JABAS, FAR_SERGUIA, FAR_CONCEPTO, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMOPER2, FAR_COSPRO,FAR_PRECIO, FAR_EQUIV, FAR_FECHA, FAR_STOCK, FAR_ESTADO,FAR_CP, FAR_CODCLIE, FAR_TIPMOV, FAR_OTRA_CIA, far_SIGNO_aRM ,FAR_codven, far_costeo  FROM facart WHERE FAR_CODART = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_CODCIA = ? AND FAR_ESTADO <> 'E' AND FAR_ESTADO <> 'M' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L'  ORDER BY FAR_CODART, FAR_FECHA asc , FAR_NUMOPER  asc , FAR_NUMSEC "
   End If
   If LK_FLAG_SOS = "A" Then
      If cheext.Value = 1 Then
         pub_cadena = "SELECT FAR_TRANSITO, FAR_FLAG_SO, FAR_TIPDOC, FAR_FECHA_COMPRA, far_moneda,far_precio_neto, far_subtra, far_numSER_c, far_numfac_c , far_numGUIA, FAR_DIAS , FAR_CODUSU, FAR_SERGUIA, FAR_CODART, FAR_CANTIDAD, FAR_JABAS, FAR_CONCEPTO, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMOPER2, FAR_COSPRO,FAR_PRECIO, FAR_EQUIV, FAR_FECHA, FAR_STOCK, FAR_ESTADO,FAR_CP, FAR_CODCLIE, FAR_TIPMOV, FAR_OTRA_CIA, far_SIGNO_aRM ,FAR_codven, far_costeo  FROM facart WHERE FAR_CODART = ? AND FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA < ? AND  FAR_CODCIA = ? AND FAR_ESTADO <> 'M' AND FAR_FLAG_SO = 'A' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L'  ORDER BY FAR_CODART, FAR_FECHA_COMPRA asc , FAR_NUMOPER2  asc , FAR_NUMSEC "
      Else
         pub_cadena = "SELECT FAR_TRANSITO, FAR_FLAG_SO, FAR_TIPDOC, FAR_FECHA_COMPRA, far_moneda,far_precio_neto, far_subtra, far_numSER_c, far_numfac_c , far_numGUIA, FAR_DIAS , FAR_CODUSU,FAR_CODART, FAR_CANTIDAD,  FAR_JABAS, FAR_SERGUIA, FAR_CONCEPTO, FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMOPER2, FAR_COSPRO,FAR_PRECIO, FAR_EQUIV, FAR_FECHA, FAR_STOCK, FAR_ESTADO,FAR_CP, FAR_CODCLIE, FAR_TIPMOV, FAR_OTRA_CIA, far_SIGNO_aRM ,FAR_codven, far_costeo  FROM facart WHERE FAR_CODART = ? AND FAR_FECHA_COMPRA >= ? AND  FAR_FECHA_COMPRA < ? AND FAR_CODCIA = ? AND FAR_ESTADO <> 'E' AND FAR_ESTADO <> 'M' AND FAR_FLAG_SO = 'A' And FAR_TRANSITO <> 'P' AND FAR_ESTADO2 <> 'L'  ORDER BY FAR_CODART, FAR_FECHA asc , FAR_NUMOPER2  asc , FAR_NUMSEC "
      End If
   End If
   
   Set PSFAR_CODART = CN.CreateQuery("", pub_cadena)
   PSFAR_CODART(0) = 0
   PSFAR_CODART(1) = LK_FECHA_DIA
   PSFAR_CODART(2) = LK_FECHA_DIA
   PSFAR_CODART(3) = "00"
   
   Frame1.Visible = True
   Frame2.Visible = False
   'Frame4.Visible = False
'   Frame1.Top = todos.Top
   ListView1.Visible = False
   i_maximo.ListIndex = 3
   'i_mes.SetFocus
ElseIf option1(2).Value = True Then
   Frame1.Visible = False
   Frame2.Visible = False
   Frame4.Visible = True
   GridK.Clear
   GridK.Rows = 2
 '  i_codart2.Text = ""
End If
   Azul i_codart2, i_codart2
End Sub

Private Sub salir_Click()
 Unload KARDEX
End Sub

Public Sub cabe()
KARDEX.GridK.ColAlignment(2) = 0
 KARDEX.GridK.ColAlignment(3) = 0
 KARDEX.GridK.ColAlignment(4) = 1
 KARDEX.GridK.ColAlignment(5) = 1
 KARDEX.GridK.ColAlignment(6) = 1
 KARDEX.GridK.ColAlignment(7) = 1
 KARDEX.GridK.ColAlignment(8) = 1
 KARDEX.GridK.ColAlignment(9) = 1
 KARDEX.GridK.ColAlignment(11) = 1
 KARDEX.GridK.ColAlignment(11) = 1

KARDEX.GridK.TextMatrix(0, 0) = "Fec.Proc"
KARDEX.GridK.TextMatrix(0, 1) = "Fec.Emis."
KARDEX.GridK.TextMatrix(0, 2) = "Documento"
KARDEX.GridK.TextMatrix(0, 3) = "Concepto"
KARDEX.GridK.TextMatrix(0, 4) = "Descripción"

KARDEX.GridK.TextMatrix(0, 5) = "Ingreso"
KARDEX.GridK.TextMatrix(0, 6) = "Salida"
KARDEX.GridK.TextMatrix(0, 7) = "Saldo"


KARDEX.GridK.TextMatrix(0, 8) = "Ingreso"
KARDEX.GridK.TextMatrix(0, 9) = "Salida"
KARDEX.GridK.TextMatrix(0, 10) = "Saldo"
GridK.COL = 7
GridK.Row = 0
GridK.CellBackColor = QBColor(14)
KARDEX.GridK.TextMatrix(0, 11) = "Valor(Ing.)"
KARDEX.GridK.TextMatrix(0, 12) = "Valor(Sal.)"
KARDEX.GridK.TextMatrix(0, 13) = "Valor(Saldo)"
KARDEX.GridK.TextMatrix(0, 14) = "Cost.Prom."
KARDEX.GridK.TextMatrix(0, 15) = "Precio"
KARDEX.GridK.TextMatrix(0, 17) = "Nombre"
KARDEX.GridK.TextMatrix(0, 18) = "Vendedor"
KARDEX.GridK.TextMatrix(0, 19) = "Dias"
KARDEX.GridK.TextMatrix(0, 20) = "N.Kardex"
KARDEX.GridK.TextMatrix(0, 21) = "USUARIO"


End Sub

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If txtCampo2.Visible Then Azul2 txtCampo2, txtCampo2
 End If
End Sub

Private Sub txtCampo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If CmdProcesa.Visible Then CmdProcesa_Click
End If
End Sub

Private Sub unidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Frame2.Visible And i_maximo.Visible Then i_maximo.SetFocus
    If Frame1.Visible And txtCampo1.Visible Then Azul2 txtCampo1, txtCampo1
End If
End Sub

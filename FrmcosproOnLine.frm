VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmcosproOnLine 
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcl 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin MSComctlLib.ProgressBar P1 
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar P2 
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblm 
      Caption         =   "..."
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label ol_fecha 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01/01/2006"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label ol_numfac 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label ol_numser 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Documento :"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "FrmcosproOnLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcl_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Dim wol_codcia As String
Dim wfechE As Date
wfechE = Format(ol_fecha.Caption, "dd/mm/yyyy")
If LK_CODCIA = "09" Or LK_CODCIA = "30" Or LK_CODCIA = "50" Then
   'Costeo_Chepen ' solo pàra la botica chepen
   wol_codcia = "('" & LK_CODCIA & "')"
ElseIf LK_CODCIA = "01" Or LK_CODCIA = "20" Then
   wol_codcia = "('01')"
ElseIf LK_CODCIA = "03" Or LK_CODCIA = "05" Or LK_CODCIA = "07" Then
   wol_codcia = "('03','05','07')"
ElseIf LK_CODCIA = "10" Or LK_CODCIA = "25" Or LK_CODCIA = "27" Or LK_CODCIA = "28" Or LK_CODCIA = "29" Then
   wol_codcia = "('10','25','27','28','29','15')"
Else
   wol_codcia = "('" & LK_CODCIA & "')"
End If


'If LK_CODCIA = "05" Or LK_CODCIA = "09" Then
'   wol_codcia = "('" & LK_CODCIA & "')"
'Else
'   wol_codcia = "('01','02','03','04','07','08','10','20')"
'End If
  'CosteoOnLine "D", "", "", "A", "(0)", 0, LK_FECHA_DIA, 20, 0, " ", 0, P1, P1, lblm, 0
 CosteoOnLine "D", " ", wol_codcia, "A", "(0)", 0, wfechE, 20, Val(FrmcosproOnLine.ol_numser.Caption), " ", Val(FrmcosproOnLine.ol_numfac.Caption), P1, P2, lblm, 0
 MsgBox "Documento Costeado", 48, Pub_Titulo
 Unload Me
End Sub

Private Sub Form_Load()
CenterMe FrmcosproOnLine
End Sub

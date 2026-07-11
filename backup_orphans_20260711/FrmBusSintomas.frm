VERSION 5.00
Begin VB.Form FrmBusSintomas 
   Caption         =   "Busqueda ."
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraindica 
      Caption         =   "Información :"
      Height          =   3735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtIndica 
         Height          =   1005
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtDosis 
         Height          =   1005
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txtOtros 
         Height          =   885
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2760
         Width           =   4215
      End
      Begin VB.Label lblnombre 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label lblmae 
         Alignment       =   1  'Right Justify
         Caption         =   "Indicaciones: "
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
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblmae 
         Alignment       =   1  'Right Justify
         Caption         =   "Dosis: "
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
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblmae 
         Alignment       =   1  'Right Justify
         Caption         =   "Otros: "
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
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   1095
      End
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
      Left            =   5880
      Picture         =   "FrmBusSintomas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   945
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Aceptar"
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
      Left            =   5880
      Picture         =   "FrmBusSintomas.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   945
   End
   Begin VB.Frame frabus 
      Caption         =   "Busqueda."
      ForeColor       =   &H00800000&
      Height          =   3735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtcodisc 
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   3495
      End
      Begin VB.ListBox lstbusq 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800000&
         Height          =   2985
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar Cuadros Clinicos :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmBusSintomas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
If Val(Right(lstbusq.Text, 10)) = 0 Then
  MsgBox "Seleccionar Cuadro Clinico", 48, Pub_Titulo
  Exit Sub
End If
pu_sintoma = Val(Right(lstbusq.Text, 10))
Unload Me
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If fraindica.Visible Then
    pub_cadena = "SELECT  * FROM BUSQ_tab  where bta_tipo = 2  and bta_numtab = " & Val(fraindica.Tag)
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    If Not X.EOF Then
    txtIndica.Text = Trim(X!BTA_INDIC)
    txtDosis.Text = Trim(X!BTA_DOSIS)
    txtOtros.Text = Trim(X!BTA_OSB)
    End If
End If
End Sub

Private Sub Form_Load()
CenterMe Me
End Sub

Private Sub txtcodisc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If IsNumeric(txtcodisc.Text) Then
Else
    pub_cadena = "SELECT * FROM BUSQ_tab WHERE BTA_CODCIA = '00' AND BTA_TIPO = 1 and BTA_TITULO LIKE '%" & Trim(txtcodisc.Text) & "%'"
    lstbusq.Clear
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    Do Until X.EOF
    lstbusq.AddItem Trim(X!BTA_TITULO) & String(150, " ") & X!BTA_NUMTAB
    X.MoveNext
    Loop
    If lstbusq.ListCount > 0 Then
       frabus.Visible = True
       lstbusq.SetFocus
    Else
    Azul txtcodisc, txtcodisc
    End If
End If

End Sub

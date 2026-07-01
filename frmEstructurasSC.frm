VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEstructurasSC 
   Caption         =   "Estructuras de Sintomas / Composiciones"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Definir Estructuras"
      TabPicture(0)   =   "frmEstructurasSC.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frabus"
      Tab(0).Control(1)=   "txtcodisc"
      Tab(0).Control(2)=   "lstcom"
      Tab(0).Control(3)=   "cmdAgregar"
      Tab(0).Control(4)=   "cmdCerrar"
      Tab(0).Control(5)=   "fralineas"
      Tab(0).Control(6)=   "lstsin"
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(8)=   "lstprod"
      Tab(0).Control(9)=   "lblmae(10)"
      Tab(0).Control(10)=   "lblmae(9)"
      Tab(0).Control(11)=   "lblmae(8)"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Sintomas / Composiciones"
      TabPicture(1)   =   "frmEstructurasSC.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblmae(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblmae(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblmae(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblmae(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblmae(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblmae(6)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblmae(7)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtkey"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtTitulo"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtBusqueda"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtIndica"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtDosis"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtOtros"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Frame3"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      Begin VB.Frame frabus 
         Caption         =   "Busqueda."
         ForeColor       =   &H00800000&
         Height          =   3375
         Left            =   -67080
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstbusq 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00800000&
            Height          =   2985
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.TextBox txtcodisc 
         Height          =   285
         Left            =   -67680
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ListBox lstcom 
         Height          =   450
         Left            =   -69360
         TabIndex        =   30
         Top             =   4440
         Width           =   4815
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   10095
         Begin VB.ComboBox cmbTipo 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label lblmae 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo :"
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
            Index           =   1
            Left            =   360
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   16
         Top             =   3360
         Width           =   10095
         Begin VB.CommandButton CmdCancelSC 
            Caption         =   "&Cancelar"
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
            Left            =   2760
            Picture         =   "frmEstructurasSC.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   240
            Width           =   945
         End
         Begin VB.CommandButton CmdAceptaSC 
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
            Left            =   840
            Picture         =   "frmEstructurasSC.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   240
            Width           =   945
         End
         Begin VB.CommandButton cmdCloseSC 
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
            Left            =   8520
            Picture         =   "frmEstructurasSC.frx":08BC
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.TextBox txtOtros 
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   2880
         Width           =   8775
      End
      Begin VB.TextBox txtDosis 
         Height          =   285
         Left            =   1440
         TabIndex        =   25
         Top             =   2520
         Width           =   8775
      End
      Begin VB.TextBox txtIndica 
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   2160
         Width           =   8775
      End
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   1800
         Width           =   4815
      End
      Begin VB.TextBox txtTitulo 
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtkey 
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Actualizar"
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
         Left            =   -69000
         Picture         =   "frmEstructurasSC.frx":0A06
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5040
         Width           =   1185
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
         Left            =   -66120
         Picture         =   "frmEstructurasSC.frx":0E48
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5040
         Width           =   945
      End
      Begin VB.Frame fralineas 
         Caption         =   "Lineas :"
         Height          =   615
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   5415
         Begin VB.ComboBox cmblineas 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.ListBox lstsin 
         Height          =   2400
         Left            =   -69360
         TabIndex        =   5
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipos :"
         Height          =   615
         Left            =   -69240
         TabIndex        =   2
         Top             =   360
         Width           =   4215
         Begin VB.OptionButton OpPert 
            Caption         =   "Sintomas"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OpPert 
            Caption         =   "Composiciones"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   3
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.ListBox lstprod 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4545
         Left            =   -74880
         TabIndex        =   1
         Top             =   1080
         Width           =   5415
      End
      Begin VB.Label lblmae 
         Caption         =   "Composiciones"
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
         Index           =   10
         Left            =   -69360
         TabIndex        =   34
         Top             =   4200
         Width           =   2775
      End
      Begin VB.Label lblmae 
         Caption         =   "Sintomas"
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
         Index           =   9
         Left            =   -69240
         TabIndex        =   33
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label lblmae 
         Alignment       =   1  'Right Justify
         Caption         =   "Adicionar Codigo: "
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
         Index           =   8
         Left            =   -69240
         TabIndex        =   31
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblmae 
         Caption         =   "Cada palabra separado por un espacio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   6480
         TabIndex        =   22
         Top             =   1800
         Width           =   3495
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
         Left            =   240
         TabIndex        =   15
         Top             =   2880
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
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   1095
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
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblmae 
         Alignment       =   1  'Right Justify
         Caption         =   "Buscar como: "
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
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblmae 
         Alignment       =   1  'Right Justify
         Caption         =   "Titulo: "
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
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblmae 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo: "
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
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmEstructurasSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmblineas_Click()
pub_cadena = "SELECT  ART_KEY , ART_NOMBRE FROM ARTI WHERE ART_CODCIA = '" & LK_CODCIA & "' AND ART_SITUACION <> '1' AND ART_FAMILIA = " & Val(Right(cmblineas.Text, 10)) & " ORDER BY ART_NOMBRE"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
lstprod.Clear
Do Until X.EOF
lstprod.AddItem X!ART_NOMBRE & String(80, " ") & X!ART_KEY
X.MoveNext
Loop
lstsin.Clear
lstcom.Clear
End Sub

Private Sub cmbTipo_Click()
pub_cadena = "SELECT TOP 1 BTA_NUMTAB FROM BUSQ_tab WHERE BTA_CODCIA = '00' AND BTA_TIPO = " & Val(Right(cmbTipo.Text, 10)) & "  ORDER BY BTA_NUMTAB DESC"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not X.EOF Then
  txtkey.Text = Val(X!BTA_NUMTAB) + 1
Else
  txtkey.Text = 1
End If
txtTitulo.SetFocus
End Sub

Private Sub CmdAceptaSC_Click()
If Trim(txtTitulo.Text) = "" Then
  MsgBox "Debe Ingresar Titulo", 48, Pub_Titulo
  Azul txtTitulo, txtTitulo
  Exit Sub
End If
If Trim(txtBusqueda.Text) = "" Then
  MsgBox "Debe Ingresar Busqueda", 48, Pub_Titulo
  Azul txtBusqueda, txtBusqueda
  Exit Sub
End If

pub_cadena = "SELECT * FROM BUSQ_tab WHERE BTA_CODCIA = '00' AND BTA_TIPO = " & Val(Right(cmbTipo.Text, 10)) & " and BTA_NUMTAB   = " & Val(txtkey.Text) & " "
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If X.EOF Then
X.AddNew
Else
X.Edit
End If
X!BTA_CODCIA = "00"
X!BTA_TIPO = Val(Right(cmbTipo.Text, 10))
X!BTA_NUMTAB = Val(txtkey.Text)
X!BTA_TITULO = txtTitulo.Text
X!BTA_busqueda = txtBusqueda.Text
X!BTA_indic = txtIndica.Text
X!BTA_dosis = txtDosis.Text
X!BTA_osb = txtOtros.Text
X.Update
MsgBox "Datos Actualizado", 48, Pub_Titulo
CmdCancelSC_Click

End Sub

Private Sub cmdAgregar_Click()
Dim RS_BUSQ As rdoResultset
'BUS_CODCIA BUS_CODART BUS_TIPO    BUS_NUMTAB
pub_cadena = "SELECT * FROM  BUSQ_ART WHERE BUS_CODCIA = '00' AND BUS_TIPO = 0"
Set RS_BUSQ = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
CN.Execute "DELETE FROM BUSQ_ART WHERE BUS_CODCIA = '00' AND BUS_CODART = " & Val(Trim(Right(lstprod.Text, 15)))
If lstsin.ListCount > 0 Then
    For fila = 0 To lstsin.ListCount - 1
        RS_BUSQ.AddNew
        RS_BUSQ!BUS_CODCIA = "00"
        RS_BUSQ!BUS_CODART = Val(Trim(Right(lstprod.Text, 15)))
        RS_BUSQ!BUS_TIPO = 1
        RS_BUSQ!BUS_NUMTAB = Val(Trim(Right(lstsin.List(fila), 15)))
        RS_BUSQ.Update
    Next fila
End If
If lstcom.ListCount > 0 Then
    For fila = 0 To lstcom.ListCount - 1
        RS_BUSQ.AddNew
        RS_BUSQ!BUS_CODCIA = "00"
        RS_BUSQ!BUS_CODART = Val(Trim(Right(lstprod.Text, 15)))
        RS_BUSQ!BUS_TIPO = 2
        RS_BUSQ!BUS_NUMTAB = Val(Trim(Right(lstcom.List(fila), 15)))
        RS_BUSQ.Update
    Next fila
End If
MsgBox "Actualizado Ok.", 48, Pub_Titulo

End Sub

Private Sub CmdCancelSC_Click()
txtkey.Text = ""
txtTitulo.Text = ""
txtBusqueda.Text = ""
txtIndica.Text = ""
txtDosis.Text = ""
txtOtros.Text = ""
cmbTipo_Click

End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdCloseSC_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterMe Me
LlenadoCbo cmblineas, 122
cmbTipo.AddItem "Sintomas" & String(80, " ") & 1
cmbTipo.AddItem "Composiciones" & String(80, " ") & 2
End Sub

Private Sub lstbusq_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   frabus.Visible = False
   txtcodisc.SetFocus
End If
If KeyAscii = 13 Then
   If OpPert(0).Value = True Then
        lstsin.AddItem lstbusq.Text
   Else
        lstcom.AddItem lstbusq.Text
   End If
   frabus.Visible = False
   Azul txtcodisc, txtcodisc
End If

End Sub

Private Sub lstcom_GotFocus()
frabus.Visible = False
End Sub

Private Sub lstcom_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And lstcom.ListIndex <> -1 Then
  lstcom.RemoveItem lstcom.ListIndex
End If

End Sub

Private Sub lstprod_Click()
Dim rs_tit As rdoResultset
' muestar datos
pub_cadena = "SELECT * FROM BUSQ_ART WHERE BUS_CODCIA = '00' AND BUS_TIPO = 1 and BUS_CODART = " & Val(Trim(Right(lstprod.Text, 15)))
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
lstsin.Clear
Do Until X.EOF
  pub_cadena = "SELECT * FROM BUSQ_tab WHERE BTA_CODCIA = '00' AND BTA_TIPO = " & X!BUS_TIPO & " and BTA_NUMTAB = " & X!BUS_NUMTAB
  Set rs_tit = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
  lstsin.AddItem rs_tit!BTA_TITULO & String(150, " ") & rs_tit!BTA_NUMTAB
  X.MoveNext
Loop
pub_cadena = "SELECT * FROM BUSQ_ART WHERE BUS_CODCIA = '00' AND BUS_TIPO = 2 and BUS_CODART = " & Val(Trim(Right(lstprod.Text, 15)))
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
lstcom.Clear
Do Until X.EOF
  pub_cadena = "SELECT * FROM BUSQ_tab WHERE BTA_CODCIA = '00' AND BTA_TIPO = " & X!BUS_TIPO & " and BTA_NUMTAB = " & X!BUS_NUMTAB
  Set rs_tit = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
  lstcom.AddItem rs_tit!BTA_TITULO & String(150, " ") & rs_tit!BTA_NUMTAB
  X.MoveNext
Loop

End Sub

Private Sub lstsin_GotFocus()
frabus.Visible = False
End Sub

Private Sub lstsin_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And lstsin.ListIndex <> -1 Then
  lstsin.RemoveItem lstsin.ListIndex
End If
End Sub

Private Sub OpPert_Click(Index As Integer)
txtcodisc.Text = ""
txtcodisc.SetFocus
End Sub


Private Sub txtcodisc_GotFocus()
frabus.Visible = False
End Sub

Private Sub txtcodisc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If IsNumeric(txtcodisc.Text) Then
Else
    If OpPert(0).Value = True Then ' cod 1
    pub_cadena = "SELECT * FROM BUSQ_tab WHERE BTA_CODCIA = '00' AND BTA_TIPO = 1 and BTA_TITULO LIKE '%" & Trim(txtcodisc.Text) & "%'"
    Else
    pub_cadena = "SELECT * FROM BUSQ_tab WHERE BTA_CODCIA = '00' AND BTA_TIPO = 2 and BTA_TITULO LIKE '%" & Trim(txtcodisc.Text) & "%'"
    End If
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

Private Sub txtkey_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

pub_cadena = "SELECT * FROM BUSQ_tab WHERE BTA_CODCIA = '00' AND BTA_TIPO = " & Val(Right(cmbTipo.Text, 10)) & " and BTA_NUMTAB   = " & Val(txtkey.Text) & " "
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If X.EOF Then
  MsgBox "Codgo no Existe", 48, Pub_Titulo
  Exit Sub
End If
txtTitulo.Text = Trim(X!BTA_TITULO)
txtBusqueda.Text = Trim(X!BTA_busqueda)
txtIndica.Text = Trim(X!BTA_indic)
txtDosis.Text = Trim(X!BTA_dosis)
txtOtros.Text = Trim(X!BTA_osb)
txtTitulo.SetFocus
End Sub

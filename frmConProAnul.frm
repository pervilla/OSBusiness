VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmConProAnul 
   Caption         =   "Control rapido de documentos anulados"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   4683
      _Version        =   393216
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   4683
      _Version        =   393216
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label lblInstruccion 
      Caption         =   "Doble Click o Enter = Ver detalle del documento  |  Esc en detalle = Regresar a lista"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   6960
      Width           =   11880
   End
   Begin VB.Label lblTitulo2 
      Caption         =   "Detalle del Documento (FACART)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   5000
   End
   Begin VB.Label lblTitulo1 
      Caption         =   "Documentos Cancelados (ALLOG)"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Top             =   620
      Width           =   5000
   End
   Begin VB.Label lblTitulo 
      Caption         =   "CONTROL RAPIDO DE DOCUMENTOS ANULADOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8000
   End
End
Attribute VB_Name = "frmConProAnul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    CenterMe Me
    ConfigurarGrid1
    CargarGrid1
End Sub

Private Sub ConfigurarGrid1()
    Grid1.Cols = 14
    Grid1.Rows = 2
    Grid1.FixedRows = 1

    Grid1.ColWidth(0) = 1100
    Grid1.ColWidth(1) = 1600
    Grid1.ColWidth(2) = 1600
    Grid1.ColWidth(3) = 2000
    Grid1.ColWidth(4) = 0
    Grid1.ColWidth(5) = 0
    Grid1.ColWidth(6) = 0
    Grid1.ColWidth(7) = 0
    Grid1.ColWidth(8) = 0
    Grid1.ColWidth(9) = 0
    Grid1.ColWidth(10) = 0
    Grid1.ColWidth(11) = 0
    Grid1.ColWidth(12) = 0
    Grid1.ColWidth(13) = 0

    Grid1.TextMatrix(0, 0) = "Fecha"
    Grid1.TextMatrix(0, 1) = "Usuario"
    Grid1.TextMatrix(0, 2) = "Cliente"
    Grid1.TextMatrix(0, 3) = "Documento"
    Grid1.TextMatrix(0, 4) = "CODCIA"
    Grid1.TextMatrix(0, 5) = "CODTRA"
    Grid1.TextMatrix(0, 6) = "SECUENCIA"
    Grid1.TextMatrix(0, 7) = "NUMOPER"
    Grid1.TextMatrix(0, 8) = "AUTOCON"
    Grid1.TextMatrix(0, 9) = "TIPMOV"
    Grid1.TextMatrix(0, 10) = "FBG"
    Grid1.TextMatrix(0, 11) = "NUMSER"
    Grid1.TextMatrix(0, 12) = "NUMFAC"
    Grid1.TextMatrix(0, 13) = "CODCLIE"
End Sub

Public Sub CargarGrid1()
    Dim X As rdoResultset
    Dim wcodclie As String

    Grid1.Rows = 2
    Grid1.TextMatrix(1, 0) = ""
    Grid1.TextMatrix(1, 1) = ""
    Grid1.TextMatrix(1, 2) = ""
    Grid1.TextMatrix(1, 3) = ""

    pub_cadena = "SELECT ALL_CODCIA, ALL_CODTRA, ALL_SECUENCIA, ALL_NUMOPER, ALL_FECHA_DIA, ALL_CODUSU, ALL_AUTOCON, ALL_CODCLIE, ALL_TIPMOV, ALL_FBG, ALL_NUMSER, ALL_NUMFAC FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FLAG_EXT = 'E' AND ALL_CODTRA NOT IN (9999, 2582, 1111) AND ALL_NUMOPER > 0 ORDER BY ALL_FECHA_DIA DESC, ALL_NUMOPER DESC"
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)

    If X.EOF Then
        Grid1.Rows = 2
        Exit Sub
    End If

    SQ_OPER = 1
    Do Until X.EOF
        Grid1.Rows = Grid1.Rows + 1

        Grid1.TextMatrix(Grid1.Rows - 1, 0) = Format(X!ALL_FECHA_DIA, "dd/mm/yyyy")
        Grid1.TextMatrix(Grid1.Rows - 1, 1) = Trim(X!ALL_CODUSU & "")
        wcodclie = Trim(X!ALL_CODCLIE & "")
        Grid1.TextMatrix(Grid1.Rows - 1, 13) = wcodclie
        Grid1.TextMatrix(Grid1.Rows - 1, 2) = wcodclie
        Grid1.TextMatrix(Grid1.Rows - 1, 3) = Trim(X!ALL_FBG & "") & " " & Trim(X!ALL_NUMSER & "") & "-" & Trim(X!ALL_NUMFAC & "")

        Grid1.TextMatrix(Grid1.Rows - 1, 4) = Trim(X!ALL_CODCIA & "")
        Grid1.TextMatrix(Grid1.Rows - 1, 5) = Trim(X!ALL_CODTRA & "")
        Grid1.TextMatrix(Grid1.Rows - 1, 6) = Trim(X!ALL_SECUENCIA & "")
        Grid1.TextMatrix(Grid1.Rows - 1, 7) = Trim(X!ALL_NUMOPER & "")
        Grid1.TextMatrix(Grid1.Rows - 1, 8) = Trim(X!ALL_AUTOCON & "")
        Grid1.TextMatrix(Grid1.Rows - 1, 9) = Trim(X!ALL_TIPMOV & "")
        Grid1.TextMatrix(Grid1.Rows - 1, 10) = Trim(X!ALL_FBG & "")
        Grid1.TextMatrix(Grid1.Rows - 1, 11) = Trim(X!ALL_NUMSER & "")
        Grid1.TextMatrix(Grid1.Rows - 1, 12) = Trim(X!ALL_NUMFAC & "")

        X.MoveNext
    Loop

    If Grid1.Rows > 2 Then
        Grid1.Row = 1
    End If
End Sub

Public Sub CargarGrid2(ByVal wsCodCia As String, ByVal wsTipMov As String, ByVal wsFbg As String, ByVal wsNumSer As String, ByVal wsNumFac As String)
    Dim X As rdoResultset
    Dim wdesprod As String

    Grid2.Cols = 10
    Grid2.Rows = 2
    Grid2.FixedRows = 1

    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1400
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 2800
    Grid2.ColWidth(4) = 1000
    Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 1200
    Grid2.ColWidth(7) = 0
    Grid2.ColWidth(8) = 0
    Grid2.ColWidth(9) = 0

    Grid2.TextMatrix(0, 0) = "Fecha"
    Grid2.TextMatrix(0, 1) = "Documento"
    Grid2.TextMatrix(0, 2) = "Codigo"
    Grid2.TextMatrix(0, 3) = "Producto"
    Grid2.TextMatrix(0, 4) = "Cantidad"
    Grid2.TextMatrix(0, 5) = "Precio"
    Grid2.TextMatrix(0, 6) = "Importe"
    Grid2.TextMatrix(0, 7) = "CODART"
    Grid2.TextMatrix(0, 8) = "CODCLIE"
    Grid2.TextMatrix(0, 9) = "NUMSEC"

    pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = '" & wsCodCia & "' AND FAR_TIPMOV = " & Val(wsTipMov) & " AND FAR_NUMSER = " & Val(wsNumSer) & " AND FAR_NUMFAC = " & Val(wsNumFac) & " AND FAR_ESTADO <> 'E' ORDER BY FAR_NUMSEC"
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)

    If X.EOF Then
        Grid2.Rows = 2
        MsgBox "No se encontro detalle FACART para este documento.", 48, Pub_Titulo
        Exit Sub
    End If

    SQ_OPER = 1
    Do Until X.EOF
        Grid2.Rows = Grid2.Rows + 1

        Grid2.TextMatrix(Grid2.Rows - 1, 0) = Format(X!FAR_FECHA, "dd/mm/yy")
        Grid2.TextMatrix(Grid2.Rows - 1, 1) = Trim(X!FAR_NUMSER & "") & "-" & Trim(X!FAR_NUMFAC & "")
        Grid2.TextMatrix(Grid2.Rows - 1, 2) = Trim(X!FAR_CODART & "")

        PUB_KEY = X!FAR_CODART
        pu_codcia = LK_CODCIA
        LEER_ART_LLAVE
        If art_LLAVE.EOF Then
            wdesprod = Trim(X!FAR_DESCRI & "")
        Else
            wdesprod = Trim(art_LLAVE!ART_NOMBRE)
        End If
        Grid2.TextMatrix(Grid2.Rows - 1, 3) = wdesprod

        Grid2.TextMatrix(Grid2.Rows - 1, 4) = Format(X!FAR_CANTIDAD, "0.00")
        Grid2.TextMatrix(Grid2.Rows - 1, 5) = Format(Val(X!FAR_PRECIO & ""), "0.00")
        Grid2.TextMatrix(Grid2.Rows - 1, 6) = Format(Val(X!FAR_SUBTOTAL & ""), "0.00")

        Grid2.TextMatrix(Grid2.Rows - 1, 7) = Trim(X!FAR_CODART & "")
        Grid2.TextMatrix(Grid2.Rows - 1, 8) = Trim(X!FAR_CODCLIE & "")
        Grid2.TextMatrix(Grid2.Rows - 1, 9) = Trim(X!FAR_NUMSEC & "")

        X.MoveNext
    Loop

    If Grid2.Rows > 2 Then
        Grid2.Row = 1
    End If
End Sub

Private Sub Grid1_DblClick()
    If Grid1.Row < 1 Then Exit Sub
    If Val(Grid1.TextMatrix(Grid1.Row, 7)) = 0 Then Exit Sub

    CargarGrid2 Grid1.TextMatrix(Grid1.Row, 4), Grid1.TextMatrix(Grid1.Row, 9), Grid1.TextMatrix(Grid1.Row, 10), Grid1.TextMatrix(Grid1.Row, 11), Grid1.TextMatrix(Grid1.Row, 12)
    Grid2.SetFocus
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Grid1_DblClick
    End If
End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Grid1.SetFocus
    End If
End Sub

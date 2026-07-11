VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmstockgen 
   Caption         =   "Actualizar Ofertas"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PB 
      Height          =   135
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
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
      Height          =   630
      Left            =   9840
      Picture         =   "frmstockGen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   825
   End
   Begin VB.CommandButton cmdmostrar 
      Caption         =   "&Mostrar"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox cmdfamilia 
      Height          =   315
      Left            =   720
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid gdprod 
      Height          =   4605
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8123
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label lbltitulo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Familia:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
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
      TabIndex        =   4
      Top             =   5520
      Width           =   10695
   End
End
Attribute VB_Name = "frmstockgen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Unload frmstockgen
End Sub

Private Sub cmdmostrar_Click()
Dim ARMA_FAMILIA As String
Dim ps_mstprod As rdoResultset
lbltitulo.Caption = "Actualizar: Flag de Oferta"
gdprod.Cols = 9
gdprod.Rows = 1
gdprod.FixedCols = 0
gdprod.ColWidth(0) = 600 ' item
gdprod.ColWidth(1) = 400 ' item
gdprod.ColWidth(2) = 900  ' codigo
gdprod.ColWidth(3) = 2800  'descrip
gdprod.ColWidth(4) = 800  ' unidad
gdprod.ColWidth(5) = 1000  ' STOCK
gdprod.ColWidth(6) = 800   ' pre5
gdprod.ColWidth(7) = 800   ' pre6
gdprod.ColWidth(8) = 1200   ' Stock General
gdprod.TextMatrix(0, 0) = "Item"
gdprod.TextMatrix(0, 1) = "Flag"
gdprod.TextMatrix(0, 2) = "Codigo"
gdprod.TextMatrix(0, 3) = "Descripción"
gdprod.TextMatrix(0, 4) = "Unidad"
gdprod.TextMatrix(0, 5) = "St.General"
gdprod.TextMatrix(0, 6) = "Pre 5 "
gdprod.TextMatrix(0, 7) = "Pre 6 "
gdprod.TextMatrix(0, 8) = ""
If Val(Right(cmdfamilia.Text, 8)) = 0 Then
  ARMA_FAMILIA = ""
Else
  ARMA_FAMILIA = "AND ART_FAMILIA = " & Trim(Right(cmdfamilia.Text, 8)) & " "
End If
pub_cadena = "SELECT ARTI.ART_KEY,ARTI.ART_POR6, ARTI.ART_NOMBRE, ARTI.ART_COSTO, ARTI.ART_PLANCHA, ARTI.ART_SITUACION, ARTI.ART_STOCKREAL_MAX, " & _
" ARTICULO.ARM_STOCK, ARTICULO.ARM_COSPRO,  PRECIOS.PRE_PRE5, PRECIOS.PRE_UNIDAD, PRECIOS.PRE_EQUIV, PRECIOS.PRE_FLAG_UNIDAD, PRECIOS.PRE_COSTO_ANT, PRECIOS.PRE_PRE6, " & _
" TABLAS.TAB_TIPREG, stock03.ARM_CODCIA, stock03.ARM_STOCK,stock08.ARM_STOCK, SUBALMACEN.ARM_CODCIA, SUBALMACEN.ARM_STOCK, DILVISAALMACEN.ARM_CODCIA, DILVISAALMACEN.ARM_STOCK , " & _
" Round ( (ARTICULO.ARM_STOCK + SUBALMACEN.ARM_STOCK + stock03.ARM_STOCK  +  stock08.ARM_STOCK  + ALMA07.ARM_STOCK + DILVISAALMACEN.ARM_STOCK) / PRECIOS.PRE_EQUIV,2) as stockGen , Round (ARTI.ART_STOCKREAL_MAX / PRECIOS.PRE_EQUIV,2) as StockMin , " & _
" ((Round ( (ARTICULO.ARM_STOCK + SUBALMACEN.ARM_STOCK+ DILVISAALMACEN.ARM_STOCK) / PRECIOS.PRE_EQUIV,2)) ) as dif " & _
" From   { oj (((((((BDATOS.dbo.ARTI ARTI INNER JOIN BDATOS.dbo.PRECIOS PRECIOS ON  ARTI.ART_CODCIA = PRECIOS.PRE_CODCIA AND  ARTI.ART_KEY = PRECIOS.PRE_CODART) " & _
"  INNER JOIN BDATOS.dbo.TABLAS TABLAS ON " & _
"        ARTI.ART_CODCIA = TABLAS.TAB_CODCIA AND " & _
"    ARTI.ART_FAMILIA = TABLAS.TAB_NUMTAB) " & _
"     INNER JOIN BDATOS.dbo.ARTICULO ARTICULO ON ARTI.ART_KEY = ARTICULO.ARM_CODART AND  ARTI.ART_CODCIA = ARTICULO.ARM_CODCIA)" & _
"     INNER JOIN BDATOS.dbo.ARTICULO stock03 ON  ARTICULO.ARM_CODART = stock03.ARM_CODART)  " & _
"     INNER JOIN BDATOS.dbo.ARTICULO stock08 ON  ARTICULO.ARM_CODART = stock08.ARM_CODART)  " & _
"     INNER JOIN BDATOS.dbo.ARTICULO SUBALMACEN ON  stock03.ARM_CODART = SUBALMACEN.ARM_CODART) " & _
"     INNER JOIN BDATOS.dbo.ARTICULO ALMA07 ON  SUBALMACEN.ARM_CODART = ALMA07.ARM_CODART) " & _
"     INNER JOIN Bdatos.dbo.ARTICULO DILVISAALMACEN ON SUBALMACEN.ARM_CODART = DILVISAALMACEN.ARM_CODART} " & _
" Where    ARTI.ART_SITUACION <> '1' AND     " & _
"    PRECIOS.PRE_FLAG_UNIDAD = 'A' AND     TABLAS.TAB_TIPREG = 122 AND " & _
"   stock03.ARM_CODCIA = '03' AND stock08.ARM_CODCIA = '08'  and   stock03.ARM_CODCIA = '03' AND     stock03.ARM_CODCIA = '03' AND " & _
"    SUBALMACEN.ARM_CODCIA = '04' AND  ALMA07.ARM_CODCIA = '07' AND   SUBALMACEN.ARM_CODCIA = '04' AND      DILVISAALMACEN.ARM_CODCIA = '02' AND  " & _
" ARTI.ART_CODCIA = '" & LK_CODCIA & "' " & ARMA_FAMILIA & "  order by stockGen desc "
Set ps_mstprod = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not ps_mstprod.EOF Then
    pb.Visible = True
    pbMin = 0
    pbMax = ps_mstprod.RowCount
    pbValue = 0
    gdprod.Visible = False
    DoEvents
End If
Do Until ps_mstprod.EOF
    gdprod.Rows = gdprod.Rows + 1
    pbValue = ps_mstprod.AbsolutePosition
    DoEvents
    gdprod.TextMatrix(gdprod.Rows - 1, 0) = ps_mstprod.AbsolutePosition
    gdprod.TextMatrix(gdprod.Rows - 1, 1) = Format(ps_mstprod!ART_POR6, "#")
    gdprod.TextMatrix(gdprod.Rows - 1, 2) = Format(ps_mstprod!art_key, "0")
    gdprod.TextMatrix(gdprod.Rows - 1, 3) = Trim(ps_mstprod!art_nombre)
    gdprod.TextMatrix(gdprod.Rows - 1, 4) = Trim(ps_mstprod!PRE_UNIDAD)
    gdprod.TextMatrix(gdprod.Rows - 1, 5) = Format(ps_mstprod!stockGen, "#,##0.00")
    
    gdprod.TextMatrix(gdprod.Rows - 1, 6) = Format(ps_mstprod!PRE_PRE5, "#,##0.00")
    gdprod.TextMatrix(gdprod.Rows - 1, 7) = Format(ps_mstprod!PRE_PRE6, "#,##0.00")
    gdprod.TextMatrix(gdprod.Rows - 1, 8) = ""
    If Val(ps_mstprod!ART_POR6) = 1 Then
        gdprod.COL = 1
        gdprod.Row = gdprod.Rows - 1
        gdprod.CellBackColor = vbYellow
    End If
'    vbBlack
    
ps_mstprod.MoveNext
Loop
pb.Visible = False
gdprod.Visible = True
DoEvents
  
  

End Sub

Private Sub Form_Load()
CenterMe frmstockgen
    PUB_TIPREG = 122
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cmdfamilia.Clear
    cmdfamilia.AddItem "(..) Todos" & String(80, " ")
    Do Until tab_mayor.EOF
        cmdfamilia.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
        tab_mayor.MoveNext
    Loop

End Sub

Private Sub gdprod_DblClick()
If gdprod.COL <> 1 Then Exit Sub
SQ_OPER = 1
PUB_KEY = gdprod.TextMatrix(gdprod.Row, 2)
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
If art_LLAVE.EOF Then
 MsgBox "Intente Nuevamente", 48, Pub_Titulo
 Exit Sub
End If
art_LLAVE.Edit
If Val(gdprod.TextMatrix(gdprod.Row, 1)) = 1 Then
   art_LLAVE!ART_POR6 = 0
Else
   art_LLAVE!ART_POR6 = 1
End If
gdprod.TextMatrix(gdprod.Row, 1) = Format(art_LLAVE!ART_POR6, "#")
If Val(art_LLAVE!ART_POR6) = 1 Then
    gdprod.COL = 1
    gdprod.Row = gdprod.Row
    gdprod.CellBackColor = vbYellow
Else
    gdprod.COL = 1
    gdprod.Row = gdprod.Row
    gdprod.CellBackColor = vbBlack
End If

art_LLAVE.Update
End Sub

Private Sub gdprod_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 32 Then Exit Sub
gdprod_DblClick

End Sub

Private Sub lblbarraos_DblClick()
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
        Me.Move (Screen.Width - (Label3(5).Left + Label3(5).Width + 240)) / 2, _
                (Screen.Height - (Label3(5).Height + 540)) / 2, _
                Label3(5).Left + Label3(5).Width + 240, _
                Label3(5).Height + 540
    Else
        Me.WindowState = vbMaximized
    End If
End Sub

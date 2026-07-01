VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmpreciocia 
   Caption         =   "Actulizar por Bloque e Independiente."
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Mostrar Precios"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid grid_unid 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "[INSERT] Agrega, [DEL] Quitar"
      Top             =   720
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   4
      FixedCols       =   3
      BackColorBkg    =   8421376
      Enabled         =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin VB.Label Label2 
      Caption         =   "Producto:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblproducto 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmpreciocia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim wp_codcia  As String
Dim st_costo As rdoResultset
grid_unid.Cols = 21
grid_unid.ColWidth(0) = 1400 ' cia
grid_unid.ColWidth(1) = 1000 ' unidad
grid_unid.ColWidth(2) = 1000 ' Costo
grid_unid.ColWidth(3) = 1000 ' precio 1
grid_unid.ColWidth(4) = 0
grid_unid.ColWidth(5) = 1000 ' precio 2
grid_unid.ColWidth(6) = 0
grid_unid.ColWidth(7) = 1000 ' precio 3
grid_unid.ColWidth(8) = 0
grid_unid.ColWidth(9) = 1000 ' precio 4
grid_unid.ColWidth(10) = 0
grid_unid.ColWidth(11) = 1000 ' precio 5
grid_unid.ColWidth(12) = 0
grid_unid.ColWidth(13) = 1000  ' precio 6
grid_unid.ColWidth(14) = 0
grid_unid.ColWidth(15) = 0
grid_unid.ColWidth(16) = 0
grid_unid.ColWidth(17) = 0
grid_unid.ColWidth(18) = 0 ' codcia
grid_unid.ColWidth(19) = 0 ' codigo producto
grid_unid.ColWidth(20) = 0 ' secuencia
grid_unid.TextMatrix(0, 0) = "Compañia"
grid_unid.TextMatrix(0, 1) = "Present"
grid_unid.TextMatrix(0, 2) = "Costo"
grid_unid.TextMatrix(0, 3) = "Precio 1"
grid_unid.TextMatrix(0, 5) = "Precio 2"
grid_unid.TextMatrix(0, 7) = "Precio 3"
grid_unid.TextMatrix(0, 9) = "Precio 4"
grid_unid.TextMatrix(0, 11) = "Precio 5"
grid_unid.TextMatrix(0, 13) = "Precio 6"
grid_unid.TextMatrix(0, 18) = ""
grid_unid.TextMatrix(0, 19) = ""
grid_unid.TextMatrix(0, 20) = ""
grid_unid.Rows = 1

fila = 0
pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODCIA = '25' AND ARM_CODART = " & PUB_CODART
Set st_costo = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)
wp_codcia = "25"
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = '" & wp_codcia & "' AND PRE_CODART = " & PUB_CODART
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
PSPAR_MULTI(0) = wp_codcia
par_multi.Requery
Do Until X.EOF
    grid_unid.Rows = grid_unid.Rows + 1
    grid_unid.RowHeight(grid_unid.Rows - 1) = 328
    fila = fila + 1
    grid_unid.TextMatrix(fila, 0) = "'" & Trim(par_multi!par_codcia) & " " & Trim(par_multi!par_nombre)
    grid_unid.TextMatrix(fila, 1) = Trim(X!PRE_UNIDAD)
    grid_unid.TextMatrix(fila, 2) = Format(Val(st_costo!ARM_COSPRO) * Val(X!PRE_EQUIV), "0.0000")
    grid_unid.TextMatrix(fila, 3) = Format(X!PRE_PRE1, "0.0000")
    grid_unid.TextMatrix(fila, 5) = Format(X!PRE_PRE2, "0.0000")
    grid_unid.TextMatrix(fila, 7) = Format(X!PRE_PRE3, "0.0000")
    grid_unid.TextMatrix(fila, 9) = Format(X!PRE_PRE4, "0.0000")
    grid_unid.TextMatrix(fila, 11) = Format(X!PRE_PRE5, "0.0000")
    grid_unid.TextMatrix(fila, 13) = Format(X!PRE_PRE6, "0.0000")
    grid_unid.TextMatrix(fila, 18) = X!PRE_CODCIA
    grid_unid.TextMatrix(fila, 18) = X!PRE_CODART
    grid_unid.TextMatrix(fila, 18) = X!PRE_SECUENCIA
    
X.MoveNext
Loop
grid_unid.Rows = grid_unid.Rows + 1
fila = fila + 1

wp_codcia = "27"
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = '" & wp_codcia & "' AND PRE_CODART = " & PUB_CODART
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
PSPAR_MULTI(0) = wp_codcia
par_multi.Requery
Do Until X.EOF
    grid_unid.Rows = grid_unid.Rows + 1
    grid_unid.RowHeight(grid_unid.Rows - 1) = 328
    fila = fila + 1
    grid_unid.TextMatrix(fila, 0) = "'" & Trim(par_multi!par_codcia) & " " & Trim(par_multi!par_nombre)
    grid_unid.TextMatrix(fila, 1) = Trim(X!PRE_UNIDAD)
    grid_unid.TextMatrix(fila, 2) = Format(Val(st_costo!ARM_COSPRO) * Val(X!PRE_EQUIV), "0.0000")
    grid_unid.TextMatrix(fila, 3) = Format(X!PRE_PRE1, "0.0000")
    grid_unid.TextMatrix(fila, 5) = Format(X!PRE_PRE2, "0.0000")
    grid_unid.TextMatrix(fila, 7) = Format(X!PRE_PRE3, "0.0000")
    grid_unid.TextMatrix(fila, 9) = Format(X!PRE_PRE4, "0.0000")
    grid_unid.TextMatrix(fila, 11) = Format(X!PRE_PRE5, "0.0000")
    grid_unid.TextMatrix(fila, 13) = Format(X!PRE_PRE6, "0.0000")
    grid_unid.TextMatrix(fila, 18) = X!PRE_CODCIA
    grid_unid.TextMatrix(fila, 18) = X!PRE_CODART
    grid_unid.TextMatrix(fila, 18) = X!PRE_SECUENCIA
    
X.MoveNext
Loop



grid_unid.Rows = grid_unid.Rows + 1
fila = fila + 1
wp_codcia = "28"
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = '" & wp_codcia & "' AND PRE_CODART = " & PUB_CODART
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
PSPAR_MULTI(0) = wp_codcia
par_multi.Requery
Do Until X.EOF
    grid_unid.Rows = grid_unid.Rows + 1
    grid_unid.RowHeight(grid_unid.Rows - 1) = 328
    fila = fila + 1
    grid_unid.TextMatrix(fila, 0) = "'" & Trim(par_multi!par_codcia) & " " & Trim(par_multi!par_nombre)
    grid_unid.TextMatrix(fila, 1) = Trim(X!PRE_UNIDAD)
    grid_unid.TextMatrix(fila, 2) = Format(Val(st_costo!ARM_COSPRO) * Val(X!PRE_EQUIV), "0.0000")
    grid_unid.TextMatrix(fila, 3) = Format(X!PRE_PRE1, "0.0000")
    grid_unid.TextMatrix(fila, 5) = Format(X!PRE_PRE2, "0.0000")
    grid_unid.TextMatrix(fila, 7) = Format(X!PRE_PRE3, "0.0000")
    grid_unid.TextMatrix(fila, 9) = Format(X!PRE_PRE4, "0.0000")
    grid_unid.TextMatrix(fila, 11) = Format(X!PRE_PRE5, "0.0000")
    grid_unid.TextMatrix(fila, 13) = Format(X!PRE_PRE6, "0.0000")
    grid_unid.TextMatrix(fila, 18) = X!PRE_CODCIA
    grid_unid.TextMatrix(fila, 18) = X!PRE_CODART
    grid_unid.TextMatrix(fila, 18) = X!PRE_SECUENCIA
    
X.MoveNext
Loop


grid_unid.Rows = grid_unid.Rows + 1
fila = fila + 1
wp_codcia = "29"
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = '" & wp_codcia & "' AND PRE_CODART = " & PUB_CODART
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
PSPAR_MULTI(0) = wp_codcia
par_multi.Requery
Do Until X.EOF
    grid_unid.Rows = grid_unid.Rows + 1
    grid_unid.RowHeight(grid_unid.Rows - 1) = 328
    fila = fila + 1
    grid_unid.TextMatrix(fila, 0) = "'" & Trim(par_multi!par_codcia) & " " & Trim(par_multi!par_nombre)
    grid_unid.TextMatrix(fila, 1) = Trim(X!PRE_UNIDAD)
    grid_unid.TextMatrix(fila, 2) = Format(Val(st_costo!ARM_COSPRO) * Val(X!PRE_EQUIV), "0.0000")
    grid_unid.TextMatrix(fila, 3) = Format(X!PRE_PRE1, "0.0000")
    grid_unid.TextMatrix(fila, 5) = Format(X!PRE_PRE2, "0.0000")
    grid_unid.TextMatrix(fila, 7) = Format(X!PRE_PRE3, "0.0000")
    grid_unid.TextMatrix(fila, 9) = Format(X!PRE_PRE4, "0.0000")
    grid_unid.TextMatrix(fila, 11) = Format(X!PRE_PRE5, "0.0000")
    grid_unid.TextMatrix(fila, 13) = Format(X!PRE_PRE6, "0.0000")
    grid_unid.TextMatrix(fila, 18) = X!PRE_CODCIA
    grid_unid.TextMatrix(fila, 18) = X!PRE_CODART
    grid_unid.TextMatrix(fila, 18) = X!PRE_SECUENCIA
    
X.MoveNext
Loop



grid_unid.Rows = grid_unid.Rows + 1
fila = fila + 1

wp_codcia = "09"
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = '" & wp_codcia & "' AND PRE_CODART = " & PUB_CODART
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
PSPAR_MULTI(0) = wp_codcia
par_multi.Requery
Do Until X.EOF
    grid_unid.Rows = grid_unid.Rows + 1
    grid_unid.RowHeight(grid_unid.Rows - 1) = 328
    fila = fila + 1
    grid_unid.TextMatrix(fila, 0) = "'" & Trim(par_multi!par_codcia) & " " & Trim(par_multi!par_nombre)
    grid_unid.TextMatrix(fila, 1) = Trim(X!PRE_UNIDAD)
    grid_unid.TextMatrix(fila, 2) = Format(Val(st_costo!ARM_COSPRO) * Val(X!PRE_EQUIV), "0.0000")
    grid_unid.TextMatrix(fila, 3) = Format(X!PRE_PRE1, "0.0000")
    grid_unid.TextMatrix(fila, 5) = Format(X!PRE_PRE2, "0.0000")
    grid_unid.TextMatrix(fila, 7) = Format(X!PRE_PRE3, "0.0000")
    grid_unid.TextMatrix(fila, 9) = Format(X!PRE_PRE4, "0.0000")
    grid_unid.TextMatrix(fila, 11) = Format(X!PRE_PRE5, "0.0000")
    grid_unid.TextMatrix(fila, 13) = Format(X!PRE_PRE6, "0.0000")
    grid_unid.TextMatrix(fila, 18) = X!PRE_CODCIA
    grid_unid.TextMatrix(fila, 18) = X!PRE_CODART
    grid_unid.TextMatrix(fila, 18) = X!PRE_SECUENCIA
    
X.MoveNext
Loop






End Sub

Private Sub Form_Load()
SQ_OPER = 1
PUB_KEY = PUB_CODART
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
lblproducto.Caption = art_LLAVE!art_key & " - " & Trim(art_LLAVE!art_nombre)

End Sub

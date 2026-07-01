VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmPreciosArti 
   Caption         =   "Actualizar en Bloque segun Sub Linea"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ListBox lstproductos 
      Height          =   2535
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Actulizar"
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
      Left            =   5280
      Picture         =   "frmPreciosArti.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1300
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
      Left            =   5520
      Picture         =   "frmPreciosArti.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Lista de Productos a Afectar :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmPreciosArti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Unload frmPreciosArti
End Sub

Private Sub cmdGrabar_Click()
Dim W_COSTO As Currency
Dim WS_POR1 As Currency
Dim WS_POR2 As Currency
Dim WS_POR3 As Currency
Dim WS_POR4 As Currency
Dim WS_POR5 As Currency
Dim WS_POR6 As Currency
Dim WS_POR7 As Currency
Dim WS_POR8 As Currency
Dim WS_POR9 As Currency
Dim I As Integer
Dim ps_preart As rdoResultset

pb.Visible = True
If lstproductos.ListCount > 0 Then
    pbValue = 0
    pbMin = 0
    pbMax = lstproductos.ListCount
    DoEvents
End If

For I = 0 To lstproductos.ListCount - 1
    pbValue = pbValue + 1
    DoEvents
    If lstproductos.Selected(I) = False Then GoTo pasa_act
    SQ_OPER = 1
    PUB_KEY = Right(lstproductos.List(I), 10)
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "No Actualizo, el producto " & lstproductos.Text
    Else
        SQ_OPER = 1
        PUB_CODART = PUB_KEY
        pu_codcia = LK_CODCIA
        LEER_ARM_LLAVE
        
        SQ_OPER = 2
        pu_codcia = LK_CODCIA
        ' el codigo del producto del maestro
        ' revisar pases de codigos ono actualiza
        
        PUB_CODART = Val(frmARTI.Txt_key.Text)
        LEER_PRE_LLAVE
        pre_mayor.Requery
        Do Until pre_mayor.EOF
            If LK_CODCIA = "25" Or LK_CODCIA = "27" Or LK_CODCIA = "28" Then
              pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA IN ('25','27','28') AND PRE_CODART = " & PUB_KEY & " AND PRE_SECUENCIA = " & pre_mayor!PRE_SECUENCIA
            Else
              pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND PRE_CODART = " & PUB_KEY & " AND PRE_SECUENCIA = " & pre_mayor!PRE_SECUENCIA
            End If
            Set ps_preart = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
            ps_preart.Requery
            If Not ps_preart.EOF Then
              Do Until ps_preart.EOF
                    W_COSTO = Format(Val(arm_llave!ARM_COSPRO) * ps_preart!PRE_EQUIV, "0.0000")
                    ps_preart.Edit
                    ps_preart!PRE_PRE1 = pre_mayor!PRE_PRE1
                    ps_preart!PRE_PRE2 = pre_mayor!PRE_PRE2
                    ps_preart!PRE_PRE3 = pre_mayor!PRE_PRE3
                    ps_preart!PRE_PRE4 = pre_mayor!PRE_PRE4
                    ps_preart!PRE_PRE5 = pre_mayor!PRE_PRE5
                    ps_preart!PRE_PRE5 = pre_mayor!PRE_PRE5
                    ps_preart!PRE_PRE6 = Nulo_Valor0(pre_mayor!PRE_PRE6)
                    ps_preart!PRE_PRE7 = Nulo_Valor0(pre_mayor!PRE_PRE7)
                    ps_preart!PRE_PRE8 = Nulo_Valor0(pre_mayor!PRE_PRE8)
                    ps_preart!PRE_PRE9 = Nulo_Valor0(pre_mayor!PRE_PRE9)
                    If W_COSTO <= 0 Then
                       MsgBox "OJO !!!!!!!!!!!!!!!!!!!!"
                       MsgBox "Producto no tiene Costo, Codigo: " & ps_preart!PRE_CODART, 48, Pub_Titulo
                       GoTo avanza_otro
                    End If
                    WS_POR1 = ((Val(ps_preart!PRE_PRE1) * 100) / W_COSTO) - 100
                    ps_preart!PRE_POR1 = WS_POR1
                    WS_POR2 = ((Val(ps_preart!PRE_PRE2) * 100) / W_COSTO) - 100
                    ps_preart!PRE_POR2 = WS_POR2
                    WS_POR3 = ((Val(ps_preart!PRE_PRE3) * 100) / W_COSTO) - 100
                    ps_preart!PRE_POR3 = WS_POR3
                    WS_POR4 = ((Val(ps_preart!PRE_PRE4) * 100) / W_COSTO) - 100
                    ps_preart!PRE_POR4 = WS_POR4
                    WS_POR5 = ((Val(ps_preart!PRE_PRE5) * 100) / W_COSTO) - 100
                    ps_preart!PRE_POR5 = WS_POR5
                    WS_POR6 = ((Val(ps_preart!PRE_PRE6) * 100) / W_COSTO) - 100
                    ps_preart!PRE_POR6 = WS_POR6
                    WS_POR7 = ((Val(ps_preart!PRE_PRE7) * 100) / W_COSTO) - 100
                    ps_preart!PRE_POR7 = WS_POR7
                    WS_POR8 = ((Val(ps_preart!PRE_PRE8) * 100) / W_COSTO) - 100
                    ps_preart!PRE_POR8 = WS_POR8
                    WS_POR9 = ((Val(ps_preart!PRE_PRE9) * 100) / W_COSTO) - 100
                    ps_preart!PRE_POR9 = WS_POR9
avanza_otro:
                    ps_preart!PRE_PRECHI1 = Nulo_Valor0(pre_mayor!PRE_PRECHI1)
                    ps_preart!PRE_PREC1 = Nulo_Valor0(pre_mayor!PRE_PREC1)
                    ps_preart!PRE_PREC2 = Nulo_Valor0(pre_mayor!PRE_PREC2)
                    ps_preart.Update
               ps_preart.MoveNext
              Loop
            End If
         pre_mayor.MoveNext
        Loop
    End If
pasa_act:
Next I
pb.Visible = False
MsgBox "Actualizado!!!", 48, Pub_Titulo
Unload frmPreciosArti
End Sub

Private Sub Form_Load()
Dim ps_prodrel As rdoResultset
CenterMe frmPreciosArti

pub_cadena = "SELECT ART_NOMBRE, ART_KEY FROM ARTI WHERE ART_CODCIA = '" & LK_CODCIA & "' AND ART_FAMILIA = " & Val(Trim(Right(frmARTI.art_familia.Text, 8))) & " AND ART_SUBFAM = " & Val(Trim(Right(frmARTI.art_subfam.Text, 8))) & " AND ART_KEY <> " & Val(frmARTI.Txt_key.Text) & " ORDER BY ART_NOMBRE"
Set ps_prodrel = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)

Do Until ps_prodrel.EOF
   lstproductos.AddItem ps_prodrel!art_nombre & String(80, " ") & ps_prodrel!art_key
   lstproductos.Selected(lstproductos.ListCount - 1) = True
ps_prodrel.MoveNext
Loop
  
End Sub

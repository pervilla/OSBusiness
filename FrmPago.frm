VERSION 5.00
Begin VB.Form FrmPago 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cobro de Venta"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   930
      Left            =   5160
      Picture         =   "FrmPago.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   1800
   End
   Begin VB.CommandButton cmdSoloGrabar 
      Caption         =   "Solo Grabar"
      Height          =   1170
      Left            =   5160
      Picture         =   "FrmPago.frx":4888A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   1800
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Grabar e Imprimir"
      Height          =   1170
      Left            =   3000
      Picture         =   "FrmPago.frx":54110
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   2160
   End
   Begin VB.Frame fraMixto 
      BackColor       =   &H80000004&
      Caption         =   "Pago Mixto"
      Height          =   3495
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3960
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1500
      End
      Begin VB.TextBox txtRecibido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2280
         TabIndex        =   15
         Top             =   780
         Width           =   1500
      End
      Begin VB.TextBox txtVuelto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2400
         Width           =   1500
      End
      Begin VB.TextBox txtTarjeta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   2280
         TabIndex        =   7
         Top             =   1860
         Width           =   1500
      End
      Begin VB.TextBox txtQR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   2280
         TabIndex        =   6
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label lblTotCap 
         BackColor       =   &H80000004&
         Caption         =   "TOTAL A PAGAR :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   360
         TabIndex        =   18
         Top             =   390
         Width           =   2000
      End
      Begin VB.Label lblRecibido 
         BackColor       =   &H80000004&
         Caption         =   "Con cuanto paga :"
         Height          =   285
         Left            =   360
         TabIndex        =   16
         Top             =   971
         Width           =   2000
      End
      Begin VB.Label lblVuelto 
         BackColor       =   &H80000004&
         Caption         =   "Vuelto :"
         Height          =   285
         Left            =   360
         TabIndex        =   14
         Top             =   2580
         Width           =   2000
      End
      Begin VB.Label lblTarj 
         BackColor       =   &H80000004&
         Caption         =   "Tarjeta :"
         Height          =   285
         Left            =   360
         TabIndex        =   12
         Top             =   2043
         Width           =   2000
      End
      Begin VB.Label lblQr 
         BackColor       =   &H80000004&
         Caption         =   "QR (Yape/Plin) :"
         Height          =   285
         Left            =   360
         TabIndex        =   11
         Top             =   1507
         Width           =   2000
      End
   End
   Begin VB.Frame fraMetodo 
      BackColor       =   &H80000004&
      Caption         =   "Metodo de Pago"
      Height          =   5700
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2880
      Begin VB.CommandButton cmdMedio 
         BackColor       =   &H00C0C0C0&
         Height          =   1330
         Index           =   3
         Left            =   120
         Picture         =   "FrmPago.frx":64F39
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4200
         Width           =   2600
      End
      Begin VB.CommandButton cmdMedio 
         BackColor       =   &H00C0C0C0&
         Height          =   1330
         Index           =   2
         Left            =   120
         Picture         =   "FrmPago.frx":685BC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1560
         Width           =   2600
      End
      Begin VB.CommandButton cmdMedio 
         BackColor       =   &H00C0C0C0&
         Height          =   1330
         Index           =   1
         Left            =   120
         Picture         =   "FrmPago.frx":6B994
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2880
         Width           =   2600
      End
      Begin VB.CommandButton cmdMedio 
         BackColor       =   &H00C0FFFF&
         Height          =   1330
         Index           =   0
         Left            =   120
         Picture         =   "FrmPago.frx":6F3DA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2600
      End
   End
End
Attribute VB_Name = "FrmPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalVenta As Currency
Public TieneImpresion As Boolean
Dim Iniciado As Boolean
Dim MetodoSel As Integer
Dim Mostrado As Boolean
Dim Ajustando As Boolean

Private Sub Form_Load()
Iniciado = False
MetodoSel = 0
Mostrado = False
CargaIconos
CenterMe FrmPago
End Sub

Private Sub CargaIconos()
On Error Resume Next
Dim wRuta As String
wRuta = App.Path & "\Imagenes\Pagos\"
If dir(wRuta & "efectivo.ico") <> "" Then cmdMedio(0).Picture = LoadPicture(wRuta & "efectivo.ico")
If dir(wRuta & "qr.ico") <> "" Then cmdMedio(1).Picture = LoadPicture(wRuta & "qr.ico")
If dir(wRuta & "tarjeta.ico") <> "" Then cmdMedio(2).Picture = LoadPicture(wRuta & "tarjeta.ico")
If dir(wRuta & "mixto.ico") <> "" Then cmdMedio(3).Picture = LoadPicture(wRuta & "mixto.ico")
End Sub

Public Sub Inicia(Monto As Currency, Impresion As Boolean)
TotalVenta = Monto
TieneImpresion = Impresion
txttotal.Text = Format(TotalVenta, "0.00")
txtVuelto.Text = "0.00"
txtRecibido.Text = ""
txtQR.Text = ""
txtTarjeta.Text = ""
cmdImprimir.Enabled = TieneImpresion
MetodoSel = 0
CambiaMetodo
Iniciado = True
Mostrado = False
End Sub

Private Sub Form_Activate()
If Iniciado And Not Mostrado Then
   Mostrado = True
   cmdMedio(0).SetFocus
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   cmdcancelar_Click
End If
End Sub

Private Sub CambiaMetodo()
Dim i As Integer
For i = 0 To 3
   If i = MetodoSel Then
      cmdMedio(i).BackColor = &HC0FFFF
   Else
      cmdMedio(i).BackColor = &HC0C0C0
   End If
Next i
fraMixto.Visible = True
lblRecibido.Visible = False
txtRecibido.Visible = False
lblQr.Visible = False
txtQR.Visible = False
lblTarj.Visible = False
txtTarjeta.Visible = False
txtVuelto.Text = "0.00"
Select Case MetodoSel
  Case 0
    fraMixto.Caption = "Pago en Efectivo"
    lblRecibido.Caption = "Con cuanto paga :"
    lblRecibido.Visible = True
    txtRecibido.Visible = True
  Case 1
    fraMixto.Caption = "Pago por QR"
  Case 2
    fraMixto.Caption = "Pago con Tarjeta"
  Case 3
    fraMixto.Caption = "Pago Mixto"
    lblRecibido.Caption = "Efectivo :"
    lblRecibido.Visible = True
    txtRecibido.Visible = True
    lblQr.Visible = True
    txtQR.Visible = True
    lblTarj.Visible = True
    txtTarjeta.Visible = True
    Ajustando = True
    txtQR.Text = ""
    txtTarjeta.Text = ""
    Ajustando = False
    AjustaMixto "R"
End Select
End Sub

Private Sub cmdMedio_Click(Index As Integer)
If MetodoSel = Index Then
   Select Case Index
     Case 0
       txtRecibido.SetFocus
     Case 1
       EmitirPago
     Case 2
       EmitirPago
     Case 3
       txtRecibido.SetFocus
   End Select
Else
   MetodoSel = Index
   CambiaMetodo
   cmdMedio(Index).SetFocus
End If
End Sub

Private Sub AjustaMixto(Origen As String)
If Not Iniciado Or Ajustando Then Exit Sub
Ajustando = True
Dim wSaldo As Currency
Select Case Origen
  Case "R"
    wSaldo = TotalVenta - Val(txtRecibido.Text) - Val(txtTarjeta.Text)
    If wSaldo < 0 Then wSaldo = 0
    txtQR.Text = Format(wSaldo, "0.00")
  Case "T"
    wSaldo = TotalVenta - Val(txtRecibido.Text) - Val(txtTarjeta.Text)
    If wSaldo < 0 Then wSaldo = 0
    txtQR.Text = Format(wSaldo, "0.00")
  Case "Q"
    wSaldo = TotalVenta - Val(txtRecibido.Text) - Val(txtQR.Text)
    If wSaldo < 0 Then wSaldo = 0
    txtTarjeta.Text = Format(wSaldo, "0.00")
End Select
CalculaVueltoMixto
Ajustando = False
End Sub

Private Sub CalculaVueltoMixto()
Dim wfalta As Currency
wfalta = TotalVenta - Val(txtQR.Text) - Val(txtTarjeta.Text)
If wfalta < 0 Then wfalta = 0
If Val(txtRecibido.Text) > wfalta Then
   txtVuelto.Text = Format(Val(txtRecibido.Text) - wfalta, "0.00")
Else
   txtVuelto.Text = "0.00"
End If
End Sub

Private Sub txtRecibido_Change()
If Not Iniciado Or Ajustando Then Exit Sub
If MetodoSel = 3 Then
   AjustaMixto "R"
Else
   If TotalVenta > 0 And Val(txtRecibido.Text) >= TotalVenta Then
      txtVuelto.Text = Format(Val(txtRecibido.Text) - TotalVenta, "0.00")
   Else
      txtVuelto.Text = "0.00"
   End If
End If
End Sub

Private Sub txtRecibido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If MetodoSel = 3 Then
      txtQR.SetFocus
   Else
      If Trim(txtRecibido.Text) = "" Then txtRecibido.Text = Format(TotalVenta, "0.00")
      If Val(txtRecibido.Text) < TotalVenta Then
         MsgBox "Monto recibido es menor al total de la venta.", vbExclamation, Pub_Titulo
         Exit Sub
      End If
      txtVuelto.Text = Format(Val(txtRecibido.Text) - TotalVenta, "0.00")
      EmitirPago
   End If
Else
   SOLO_DECIMAL txtRecibido, KeyAscii
End If
End Sub

Private Sub txtQR_Change()
If Not Iniciado Or Ajustando Then Exit Sub
AjustaMixto "Q"
End Sub

Private Sub txtTarjeta_Change()
If Not Iniciado Or Ajustando Then Exit Sub
AjustaMixto "T"
End Sub

Private Sub txtQR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Val(txtRecibido.Text) + Val(txtQR.Text) + Val(txtTarjeta.Text) < TotalVenta Then
      MsgBox "Los montos del pago mixto no cubren el total.", vbExclamation, Pub_Titulo
      Exit Sub
   End If
   CalculaVueltoMixto
   EmitirPago
Else
   SOLO_DECIMAL txtQR, KeyAscii
End If
End Sub

Private Sub txtTarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Val(txtRecibido.Text) + Val(txtQR.Text) + Val(txtTarjeta.Text) < TotalVenta Then
      MsgBox "Los montos del pago mixto no cubren el total.", vbExclamation, Pub_Titulo
      Exit Sub
   End If
   CalculaVueltoMixto
   EmitirPago
Else
   SOLO_DECIMAL txtTarjeta, KeyAscii
End If
End Sub

Private Sub CmdImprimir_Click()
GrabaPago vbYes
End Sub

Private Sub cmdSoloGrabar_Click()
GrabaPago vbNo
End Sub

Private Sub cmdcancelar_Click()
PUB_RESP_COBRO = vbCancel
Unload FrmPago
End Sub

Private Sub EmitirPago()
If cmdImprimir.Enabled Then
   CmdImprimir_Click
Else
   cmdSoloGrabar_Click
End If
End Sub

Private Sub GrabaPago(wresp As Integer)
Dim wfalta As Currency
If MetodoSel = 0 Then
   If Trim(txtRecibido.Text) = "" Then txtRecibido.Text = Format(TotalVenta, "0.00")
   If Val(txtRecibido.Text) < TotalVenta Then
      MsgBox "Monto recibido es menor al total de la venta.", vbExclamation, Pub_Titulo
      Exit Sub
   End If
ElseIf MetodoSel = 3 Then
   If Val(txtRecibido.Text) + Val(txtQR.Text) + Val(txtTarjeta.Text) < TotalVenta Then
      MsgBox "Los montos del pago mixto no cubren el total.", vbExclamation, Pub_Titulo
      Exit Sub
   End If
End If
PUB_RESP_COBRO = wresp
PUB_PAG_TOTAL = TotalVenta
Select Case MetodoSel
  Case 0
    PUB_PAG_MEDIO = "E"
    PUB_PAG_EFECTIVO = Val(txtRecibido.Text)
    PUB_PAG_QR = 0
    PUB_PAG_TARJETA = 0
    PUB_PAG_VUELTO = Val(txtRecibido.Text) - TotalVenta
  Case 1
    PUB_PAG_MEDIO = "Q"
    PUB_PAG_EFECTIVO = 0
    PUB_PAG_QR = TotalVenta
    PUB_PAG_TARJETA = 0
    PUB_PAG_VUELTO = 0
  Case 2
    PUB_PAG_MEDIO = "T"
    PUB_PAG_EFECTIVO = 0
    PUB_PAG_QR = 0
    PUB_PAG_TARJETA = TotalVenta
    PUB_PAG_VUELTO = 0
  Case 3
    PUB_PAG_MEDIO = "M"
    PUB_PAG_EFECTIVO = Val(txtRecibido.Text)
    PUB_PAG_QR = Val(txtQR.Text)
    PUB_PAG_TARJETA = Val(txtTarjeta.Text)
    wfalta = TotalVenta - Val(txtQR.Text) - Val(txtTarjeta.Text)
    If wfalta < 0 Then wfalta = 0
    If Val(txtRecibido.Text) > wfalta Then
       PUB_PAG_VUELTO = Val(txtRecibido.Text) - wfalta
    Else
       PUB_PAG_VUELTO = 0
    End If
End Select
Unload FrmPago
End Sub


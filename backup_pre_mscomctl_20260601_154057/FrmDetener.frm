VERSION 5.00
Begin VB.Form FrmDetener 
   Caption         =   "Detener Impresiones"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   2970
   StartUpPosition =   3  'Windows Default
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
      Left            =   840
      Picture         =   "FrmDetener.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2895
      Begin VB.OptionButton opopcion 
         BackColor       =   &H00C0C000&
         Caption         =   "&Operativo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton opopcion 
         BackColor       =   &H000000FF&
         Caption         =   "&Detener"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   0
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.ComboBox cmdimpresoras 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione Impresora:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "FrmDetener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload FrmDetener
End Sub

Private Sub cmdCerrar_Click()
Unload FrmDetener
End Sub

Private Sub cmdimpresoras_Click()
    SQ_OPER = 1
    PUB_TIPREG = 72
    PUB_NUMTAB = Val(Left(cmdimpresoras.Text, 1))
    PUB_CODCIA = LK_CODCIA
    LEER_TAB_LLAVE
    opopcion(0).Value = False
    opopcion(1).Value = False
    If Not tab_llave.EOF Then
      opopcion(tab_llave!tab_codart).Value = True
    End If

End Sub

Private Sub Form_Activate()
Dim j As Integer
Dim i As Integer
j = 0
For i = 1 To lk_OTROS_Count
   If Val(lk_OTROS(i)) = 20 Then 'modificar Venta
     j = 1
   End If
Next i
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Then
Else
 If j = 0 Then
   Unload FrmDetener
   Exit Sub
 End If
End If
End Sub

Private Sub Form_Load()
CenterMe FrmDetener
Dim P As Printer
fila = 0
For Each P In Printers
    cmdimpresoras.AddItem P.DeviceName & String(80, " ") & fila
    fila = fila + 1
Next P

End Sub

Private Sub opopcion_Click(Index As Integer)
    If Trim(cmdimpresoras.Text) = "" Then
        opopcion(0).Value = False
        opopcion(1).Value = False
        Exit Sub
    End If
    SQ_OPER = 1
    PUB_TIPREG = 72
    PUB_NUMTAB = Val(Left(cmdimpresoras.Text, 1))
    PUB_CODCIA = LK_CODCIA
    LEER_TAB_LLAVE
    If tab_llave.EOF Then
    tab_llave.AddNew
    Else
    tab_llave.Edit
    End If
    tab_llave!tab_codcia = LK_CODCIA
    tab_llave!tab_tipreg = 72
    tab_llave!tab_numtab = PUB_NUMTAB
    tab_llave!tab_nomlargo = ""
    tab_llave!tab_nomcorto = ""
    tab_llave!tab_codart = Index
    tab_llave.Update
End Sub

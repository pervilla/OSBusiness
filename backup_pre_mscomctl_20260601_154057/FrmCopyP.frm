VERSION 5.00
Begin VB.Form FrmCopyP 
   Caption         =   "Copia de Presupuestos."
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3060
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tAnoD 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "0000"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox tmesD 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "00"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Copiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox tanoO 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "0000"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox tmesO 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "00"
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Origen Presupuesto: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Destino Presupuesto: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Año :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   9
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Mes :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Año :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Mes :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccionar :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FrmCopyP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim WPERIODO As String
Dim VANO1 As Integer
Dim VANO2 As Integer
Dim Vmes1 As Integer
Dim Vmes2 As Integer

VANO1 = Format(LK_FECHA_DIA, "yyyy") - 3
VANO2 = Format(LK_FECHA_DIA, "yyyy") + 3
Vmes1 = 1
Vmes2 = 12

If Val(tmesO.Text) >= Vmes1 And Val(tmesO.Text) <= Vmes2 Then
Else
  MsgBox "Verificar Rango del  M E S  en el Periodo de Origen" & Chr(13) & "Rango entre : " & Vmes1 & " y " & Vmes2, 48, Pub_Titulo
  Azul tmesO, tmesO
  Exit Sub
End If

If Val(tanoO.Text) >= VANO1 And Val(tanoO.Text) <= VANO2 Then
Else
  MsgBox "Verificar Rango de A Ñ O el Periodo de Origen" & Chr(13) & "Rango entre : " & VANO1 & " y " & VANO2, 48, Pub_Titulo
  Azul tanoO, tanoO
  Exit Sub
End If

If Val(tmesD.Text) >= Vmes1 And Val(tmesD.Text) <= Vmes2 Then
Else
  Azul tmesD, tmesD
  MsgBox "Verificar Rango del M E S en el Periodo de destino" & Chr(13) & "Rango entre : " & Vmes1 & " y " & Vmes2, 48, Pub_Titulo
  Exit Sub
End If

If Val(tAnoD.Text) >= VANO1 And Val(tAnoD.Text) <= VANO2 Then
Else
  MsgBox "Verificar Rango de A Ñ O  en el Periodo de destino" & Chr(13) & "Rango entre : " & VANO1 & " y " & VANO2, 48, Pub_Titulo
  Azul tAnoD, tAnoD
  Exit Sub
End If




WPERIODO = tmesD.Text & tAnoD.Text
pub_cadena = "SELECT TOP 1 * FROM MPRESUP WHERE MPR_PERIODO = '" & WPERIODO & "'"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not X.EOF Then
   MsgBox "VERIFICAR INFOMRACION !!!!!!", vbCritical, Pub_Titulo
    pub_mensaje = "Existe Registros en el periodo que va a Reemplzar ¿Desea Continuar... ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
Else
    pub_mensaje = "Desea Copiar al Periodo : " & WPERIODO & " ¿Desea Continuar... ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
End If
cmdAceptar.Enabled = False
WPERIODO = tmesO.Text & tanoO.Text
pub_cadena = "SELECT * INTO MPRESUP_COPY  FROM MPRESUP WHERE MPR_PERIODO = '" & WPERIODO & "' "
CN.Execute pub_cadena

WPERIODO = tmesD.Text & tAnoD.Text

pub_cadena = "UPDATE MPRESUP_COPY SET MPR_PERIODO = '" & WPERIODO & "' "
CN.Execute pub_cadena

pub_cadena = "DELETE FROM MPRESUP WHERE MPR_PERIODO = '" & WPERIODO & "' "
CN.Execute pub_cadena

pub_cadena = "INSERT INTO MPRESUP SELECT * FROM MPRESUP_COPY "
CN.Execute pub_cadena

pub_cadena = "DROP TABLE  MPRESUP_COPY "
CN.Execute pub_cadena
cmdAceptar.Enabled = True
MsgBox "Registros Copiados.", 48, Pub_Titulo
Unload FrmCopyP
  

End Sub

Private Sub Form_Load()
CenterMe Me

tmesD.Text = Format(LK_FECHA_DIA, "mm")
tAnoD.Text = Format(LK_FECHA_DIA, "yyyy")
If Val(tmesD.Text) = 1 Then
tanoO.Text = Format(DateAdd("yyyy", -1, LK_FECHA_DIA), "yyyy")
Else
tanoO.Text = Format(LK_FECHA_DIA, "yyyy")
End If
tmesO.Text = Format(DateAdd("m", -1, LK_FECHA_DIA), "mm")

If LK_CODUSU = "ADMIN" Then
tmesD.Locked = False
tAnoD.Locked = False
tmesO.Locked = False
tanoO.Locked = False
End If

End Sub

Private Sub tmesD_LostFocus()
tmesD.Text = Format(Val(tmesD.Text), "00")
End Sub

Private Sub tmesO_LostFocus()
tmesO.Text = Format(Val(tmesO.Text), "00")
End Sub

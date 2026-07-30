VERSION 5.00
Begin VB.Form Splash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6600
   ClientLeft      =   825
   ClientTop       =   1155
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Splash.frx":57E2
   ScaleHeight     =   6600
   ScaleWidth      =   4965
   WhatsThisHelp   =   -1  'True
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Autorizado a:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Empresa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   3540
   End
   Begin VB.Label lblporcentaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0%..."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   2310
      TabIndex        =   1
      Top             =   3120
      Width           =   525
   End
   Begin VB.Label LblMensa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciando..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   4965
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iStatusBarWidth As Integer

Private Sub Form_Click()
' Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo sale
If KeyCode = 120 Then
    Splash.WindowState = 1
    Splash.Caption = "Proceso Abortado!!! . . ."
    Screen.MousePointer = 0
    EN.Close
    CN.Close
    Pub_ConnAdo.Close
    Screen.MousePointer = 0
    End
End If
Exit Sub
sale:
End
End Sub
Private Sub Form_Load()
IniciarLogDesarrollo True
LogInfo "Splash: Inicio Form_Load"

CenterMe Splash
Dim wflag_bloq As String * 1
Dim success%
Dim PB
PB = Chr(10) & Chr(13) & Chr(10) & Chr(13)
'On Error GoTo SALE
Screen.MousePointer = 11
If App.PrevInstance Then
  pub_mensaje = App.Path & " " & "Software"
  pub_mensaje = pub_mensaje & PB & "Posiblemente la Aplicacion este cargada o no ha sido cerrada Correctamente "
  pub_mensaje = pub_mensaje & PB & "Debe Cerrar todos los Programas e Iniciar la seccion como Usuario Distinto ..."
  MsgBox pub_mensaje, vbCritical, "Software"
  Screen.MousePointer = 0
  End
End If

Pub_Titulo = "Inversiones San Martin"
LK_CODCIA = ""
LK_CODUSU = ""
If Nulo_Valor0(PUB_FLAG) = 0 Then
  wflag_bloq = ""
  If dir("C:\WINDOWS\System\Sisacv", vbDirectory) <> "" Then
    wflag_bloq = "A"
  End If
  If dir("C:\Winnt\System\Sisacv", vbDirectory) <> "" Then
    wflag_bloq = "A"
  End If
  If wflag_bloq <> "A" Then
    MsgBox "Equipo: MicroProcesador No Identificado..." & Chr(13) & "- Esta copia del Ejecutable no procede - No tiene licencia de uso", vbCritical, "Proveedor del Software - Celular: (044)687304"
    End
  End If

  Splash.Show
  'Success% = SetWindowPos(Splash.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  DoEvents
  LogInfo "Splash: Antes de CONEXION_GEN"
  CONEXION_GEN
  LogInfo "Splash: Despues de CONEXION_GEN"
End If
PUB_FLAG = 0
DoEvents
'Success% = SetWindowPos(Splash.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
DoEvents
LogInfo "Splash: Antes de Load MDIForm1"
Load MDIForm1
LogInfo "Splash: Despues de Load MDIForm1"
Screen.MousePointer = 0
Exit Sub
sale:
 Screen.MousePointer = 0
 MsgBox Err.Description, 48, "pub_titulo"
End

End Sub



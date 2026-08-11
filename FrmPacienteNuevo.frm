VERSION 5.00
Begin VB.Form FrmPacienteNuevo 
   Caption         =   "Nuevo Paciente"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   450
      Left            =   4140
      TabIndex        =   11
      Top             =   2220
      Width           =   1500
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   450
      Left            =   240
      TabIndex        =   10
      Top             =   2220
      Width           =   1500
   End
   Begin VB.TextBox txtFechaNac 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   1500
   End
   Begin VB.TextBox txtDireccion 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Width           =   4200
   End
   Begin VB.TextBox txtTelefono 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   2400
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   4200
   End
   Begin VB.TextBox txtDNI 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2400
   End
   Begin VB.CommandButton cmdFactiliza 
      Caption         =   "Factiliza"
      Height          =   360
      Left            =   4620
      TabIndex        =   12
      Top             =   225
      Width           =   1500
   End
   Begin VB.Label lbl1 
      Caption         =   "DNI / RUC:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   1800
   End
   Begin VB.Label lbl2 
      Caption         =   "Nombre completo *:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   645
      Width           =   1800
   End
   Begin VB.Label lbl3 
      Caption         =   "Telefono:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1005
      Width           =   1800
   End
   Begin VB.Label lbl4 
      Caption         =   "Direccion:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1365
      Width           =   1800
   End
   Begin VB.Label lbl5 
      Caption         =   "Fecha nacimiento (dd/mm/aaaa):"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1725
      Width           =   2700
   End
End
Attribute VB_Name = "FrmPacienteNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PacienteId As Long
Public ClienteId As Long
Public Nombre As String

Private Sub Form_Load()
    PacienteId = 0
    ClienteId = 0
    Nombre = ""
End Sub

Private Sub cmdCancelar_Click()
    Me.Hide
End Sub

Private Sub cmdFactiliza_Click()
    On Error GoTo ErrorHTTP
    Dim ws_dni As String
    Dim ws_tipo As String
    Dim ws_url As String
    Dim ws_response As String
    Dim http As Object
    Dim ws_nombre As String
    Dim ws_direccion As String
    Dim ws_fec As String
    Dim ws_estado As String
    Dim ws_token As String

    ws_dni = Trim(txtDNI.Text)
    If Len(ws_dni) <> 8 And Len(ws_dni) <> 11 Then
        MsgBox "Ingrese un DNI (8 digitos) o RUC (11 digitos) para consultar.", vbExclamation, Pub_Titulo
        Exit Sub
    End If
    If Len(ws_dni) = 8 Then
        ws_tipo = "DNI"
    Else
        ws_tipo = "RUC"
    End If

    ws_token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiI1NTgiLCJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL3dzLzIwMDgvMDYvaWRlbnRpdHkvY2xhaW1zL3JvbGUiOiJjb25zdWx0b3IifQ.K8PwFsfNIpIl2ve0KJ2F08JZYLdGaBEx6_PvMRCm_Mw"

    Screen.MousePointer = 11
    DoEvents

    If ws_tipo = "DNI" Then
        ws_url = "https://api.factiliza.com/v1/dni/info/" & ws_dni
    Else
        ws_url = "https://api.factiliza.com/v1/ruc/info/" & ws_dni
    End If

    Set http = CreateObject("Microsoft.XMLHTTP")
    http.Open "GET", ws_url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & ws_token
    http.send

    If http.Status <> 200 Then
        Screen.MousePointer = 0
        MsgBox "Error al consultar Factiliza API:" & Chr(13) & "HTTP Status: " & http.Status, vbExclamation, Pub_Titulo
        Set http = Nothing
        Exit Sub
    End If

    ws_response = http.responseText
    Set http = Nothing
    Screen.MousePointer = 0

    If ws_tipo = "DNI" Then
        ws_nombre = ExtraerJSON(ws_response, "nombre_completo")
        ws_direccion = ExtraerJSON(ws_response, "direccion_completa")
        ws_fec = ExtraerJSON(ws_response, "fecha_nacimiento")
        If ws_nombre = "" Then
            ws_nombre = Trim(ExtraerJSON(ws_response, "nombres") & " " & ExtraerJSON(ws_response, "apellido_paterno") & " " & ExtraerJSON(ws_response, "apellido_materno"))
        End If
        If ws_direccion = "" Then ws_direccion = ExtraerJSON(ws_response, "direccion")
    Else
        ws_nombre = ExtraerJSON(ws_response, "nombre_o_razon_social")
        ws_direccion = ExtraerJSON(ws_response, "direccion_completa")
        ws_estado = ExtraerJSON(ws_response, "estado")
    End If

    If ws_nombre = "" Then
        MsgBox "No se encontro informacion para el " & ws_tipo & ": " & ws_dni, vbExclamation, Pub_Titulo
        Exit Sub
    End If

    If ws_tipo = "RUC" And Len(ws_estado) > 0 Then
        If UCase(ws_estado) <> "ACTIVO" And UCase(ws_estado) <> "HABIDO" Then
            MsgBox "El RUC " & ws_dni & " esta en estado: " & ws_estado & Chr(13) & _
                   "Verifique la informacion antes de continuar.", vbExclamation, Pub_Titulo
        End If
    End If

    txtNombre.Text = ws_nombre
    If Len(ws_direccion) > 0 Then txtDireccion.Text = ws_direccion
    If Len(ws_fec) > 0 Then txtFechaNac.Text = ws_fec

    MsgBox ws_tipo & " importado desde Factiliza:" & Chr(13) & _
           ws_tipo & ": " & ws_dni & Chr(13) & _
           "Nombre: " & ws_nombre & Chr(13) & _
           "Direccion: " & ws_direccion, vbInformation, Pub_Titulo
    Exit Sub

ErrorHTTP:
    Screen.MousePointer = 0
    On Error Resume Next
    Set http = Nothing
    On Error GoTo 0
    MsgBox "Error de conexion con Factiliza API:" & Chr(13) & Err.Description, vbExclamation, Pub_Titulo
End Sub

Private Function ExtraerJSON(ByVal jsonStr As String, ByVal key As String) As String
    Dim pos_key As Long
    Dim pos1 As Long
    Dim pos2 As Long
    ExtraerJSON = ""
    pos_key = InStr(1, jsonStr, """" & key & """")
    If pos_key = 0 Then Exit Function
    pos1 = InStr(pos_key, jsonStr, ":")
    If pos1 = 0 Then Exit Function
    pos1 = InStr(pos1, jsonStr, """")
    If pos1 = 0 Then Exit Function
    pos2 = InStr(pos1 + 1, jsonStr, """")
    If pos2 > pos1 + 1 Then
        ExtraerJSON = Mid(jsonStr, pos1 + 1, pos2 - pos1 - 1)
    End If
End Function

Private Sub cmdGuardar_Click()
    On Error GoTo EH
    Dim wNombre As String
    Dim wDNI As String
    Dim wTel As String
    Dim wDir As String
    Dim wFec As String
    Dim wFecSql As String
    Dim wCod As Long
    Dim wCia As Long
    Dim wSql As String
    Dim rs As rdoResultset

    wNombre = Trim(UCase(txtNombre.Text))
    If Len(wNombre) < 3 Then
        MsgBox "Ingrese el nombre completo del paciente.", vbExclamation, Pub_Titulo
        Exit Sub
    End If

    wDNI = Trim(txtDNI.Text)
    wTel = Trim(txtTelefono.Text)
    wDir = Trim(txtDireccion.Text)
    wFec = Trim(txtFechaNac.Text)

    wFecSql = "NULL"
    If Len(wFec) > 0 Then
        If IsDate(wFec) Then
            wFecSql = "'" & Format(CDate(wFec), "dd/mm/yyyy") & "'"
        Else
            MsgBox "Fecha de nacimiento invalida. Use el formato dd/mm/aaaa.", vbExclamation, Pub_Titulo
            Exit Sub
        End If
    End If

    wCia = Val(LK_CODCIA)
    If wCia = 0 Then wCia = 25

    Screen.MousePointer = 11
    DoEvents

    Set rs = CN.OpenResultset("SELECT ISNULL(MAX(CLI_CODCLIE), 0) + 1 AS nuevo FROM CLIENTES WHERE CLI_CP = 'C' AND CLI_CODCIA = " & wCia, rdOpenKeyset, rdConcurReadOnly)
    wCod = Val(Nulo_Valor0(rs!nuevo))
    rs.Close
    Set rs = Nothing

    wSql = "INSERT INTO CLIENTES (" & _
           "CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_NOMBRE_ESPOSO, CLI_NOMBRE_ESPOSA, CLI_NOMBRE_EMPRESA, " & _
           "CLI_123, CLI_TELEF1, CLI_TELEF2, CLI_CASA_DIREC, CLI_CASA_NUM, CLI_CASA_ZONA, CLI_CASA_SUBZONA, " & _
           "CLI_TRAB_DIREC, CLI_TRAB_NUM, CLI_TRAB_ZONA, CLI_TRAB_SUBZONA, CLI_TRAB_PROV, " & _
           "CLI_RUC_ESPOSO, CLI_RUC_ESPOSA, CLI_RUC_EMPRESA, " & _
           "CLI_CASA1, CLI_CASA2, CLI_REGPUB1, CLI_REGPUB2, CLI_AUTOAVALUO, CLI_PRENDA, " & _
           "CLI_AUTO1, CLI_AUTO2, CLI_IGV_INCLUIDO, CLI_OTRO_CONTR, CLI_LETRA, CLI_LIMCRE, CLI_FECHA_FAC, " & _
           "CLI_TIPO_BLOQ1, CLI_TIPO_BLOQ2, CLI_TIPO_BLOQ3, CLI_TIPO_BLOQ4, CLI_DET_TOT, " & _
           "CLI_NOM_LET1, CLI_NOM_LET2, CLI_GRUPO, CLI_SUBGRUPO, CLI_DIVISION, CLI_ESTADO, CLI_MONEDA, " & _
           "CLI_CODART, CLI_NUCLEO, CLI_CUENTA_CONTAB, CLI_CIA_REF, CLI_PORDESCTO, CLI_SALDO, CLI_PRECIOS, " & _
           "CLI_DIA_VISITA, CLI_ZONA_NEW, CLI_PROGRAMADO, CLI_LUGAR_CASA, CLI_LUGAR_TRAB, CLI_CUENTA_CONTAB2, " & _
           "CLI_DIAS_CRED, CLI_DIAS_FAC, CLI_CUENTA_CONTAB22, CLI_LIMCRE2, CLI_TIPO, CLI_FECHAHORA, " & _
           "CLI_CIARELA, CLI_MARCAID, CLI_TIPOCLI, CLI_FECHA_NAC, CLI_HISTORIA) VALUES (" & _
           wCod & ", " & wCia & ", 'C', '" & SQLC(wNombre) & "', '" & SQLC(wNombre) & "', '', '', " & _
           "1, '" & SQLC(wTel) & "', '', '" & SQLC(Left(wDir, 120)) & "', 0, 0, 0, " & _
           "'" & SQLC(Left(wDir, 30)) & "', 0, 0, 0, 0, " & _
           "'" & SQLC(IIf(Len(wDNI) = 11, wDNI, "")) & "', '" & SQLC(IIf(Len(wDNI) = 8, wDNI, "")) & "', '', " & _
           "'', '', '', '', '', '', " & _
           "'07', '', '', 1, 0, 0, 0, " & _
           "1, '', '', '', '', " & _
           "'', '', 1, 0, 0, 'A', 'S', " & _
           "'', '', '', '', 0, 0, '', " & _
           "3, 0, '', 1, 1, 1, " & _
           "0, 2, '', 0, '', '" & Format(Date, "yy/mm/dd") & " " & Format(Time, "hh:mm AM/PM") & " CM', " & _
           "'', '', 7, " & wFecSql & ", 0)"
    CN.Execute wSql, rdExecDirect

    wSql = "INSERT INTO DIRCLI (CODCIA, CODCLI, CP, DIREC, DIRCOMP, REF, CLI_LUGAR_TRAB, CLI_TRAB_ZONA, CLI_CASA_SUBZONA, CLI_TRAB_SUBZONA, NUMERO) VALUES (" & _
           wCia & ", " & wCod & ", 'C', '" & SQLC(Left(wDir, 60)) & "', '" & SQLC(Left(wDir, 100)) & "', '', 0, 0, 0, 0, 0)"
    CN.Execute wSql, rdExecDirect

    CN.Execute "INSERT INTO CM_PACIENTES (cliente_id, estado) VALUES (" & wCod & ", 1)", rdExecDirect
    Set rs = CN.OpenResultset("SELECT @@IDENTITY AS nuevo_id", rdOpenKeyset, rdConcurReadOnly)
    If Not rs.EOF Then PacienteId = Val(rs!nuevo_id)
    rs.Close
    Set rs = Nothing

    ClienteId = wCod
    Nombre = wNombre

    Screen.MousePointer = 0
    Me.Hide
    Exit Sub
EH:
    Screen.MousePointer = 0
    MsgBox "Error al guardar el paciente: " & Err.Description, vbExclamation, Pub_Titulo
End Sub

Private Function SQLC(ByVal wValor As String) As String
    SQLC = Replace(wValor, "'", "''")
End Function

VERSION 5.00
Begin VB.Form FrmPacienteNuevo 
   Caption         =   "Nuevo Paciente"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   Icon            =   "FrmPacienteNuevo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   915
      Left            =   4920
      Picture         =   "FrmPacienteNuevo.frx":35FD6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2220
      Width           =   1500
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   915
      Left            =   3360
      Picture         =   "FrmPacienteNuevo.frx":7E860
      Style           =   1  'Graphical
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
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   4620
      Picture         =   "FrmPacienteNuevo.frx":8A0E6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   105
      Width           =   540
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
      Caption         =   "Nacimiento (dd/mm/aaaa):"
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

Private Sub cmdcancelar_Click()
    Me.Hide
End Sub

Private Sub cmdFactiliza_Click()
    On Error GoTo ErrorHTTP
    Dim ws_dni As String
    Dim WS_TIPO As String
    Dim ws_url As String
    Dim ws_response As String
    Dim http As Object
    Dim ws_nombre As String
    Dim ws_direccion As String
    Dim ws_fec As String
    Dim WS_ESTADO As String
    Dim ws_token As String
    Dim rsLocal As rdoResultset
    Dim wCiaL As Long
    Dim wCodExist As Long
    Dim wMsgExist As String

    ws_dni = Trim(txtDNI.Text)
    If Len(ws_dni) <> 8 And Len(ws_dni) <> 11 Then
        MsgBox "Ingrese un DNI (8 digitos) o RUC (11 digitos) para consultar.", vbExclamation, Pub_Titulo
        Exit Sub
    End If
    If Len(ws_dni) = 8 Then
        WS_TIPO = "DNI"
    Else
        WS_TIPO = "RUC"
    End If

    ' Comprobar en BD local antes de consultar la API
    wCiaL = Val(LK_CODCIA)
    If wCiaL = 0 Then wCiaL = 25
    If WS_TIPO = "DNI" Then
        Set rsLocal = CN.OpenResultset("SELECT TOP 1 CLI_CODCLIE, CLI_NOMBRE, CLI_TELEF1, CLI_CASA_DIREC, CLI_FECHA_NAC FROM CLIENTES WHERE CLI_RUC_ESPOSA = '" & ws_dni & "' AND CLI_CODCIA = " & wCiaL, rdOpenKeyset, rdConcurReadOnly)
    Else
        Set rsLocal = CN.OpenResultset("SELECT TOP 1 CLI_CODCLIE, CLI_NOMBRE, CLI_TELEF1, CLI_CASA_DIREC, CLI_FECHA_NAC FROM CLIENTES WHERE CLI_RUC_ESPOSO = '" & ws_dni & "' AND CLI_CODCIA = " & wCiaL, rdOpenKeyset, rdConcurReadOnly)
    End If
    If Not rsLocal.EOF Then
        wCodExist = Val(rsLocal!cli_codclie)
        ClienteId = wCodExist
        txtNombre.Text = Trim(Nulo_Valors(rsLocal!CLI_NOMBRE))
        txtTelefono.Text = Trim(Nulo_Valors(rsLocal!cli_telef1))
        txtDireccion.Text = Trim(Nulo_Valors(rsLocal!CLI_CASA_DIREC))
        If IsDate(Nulo_Valors(rsLocal!CLI_FECHA_NAC)) Then
            txtFechaNac.Text = Format(CDate(Nulo_Valors(rsLocal!CLI_FECHA_NAC)), "dd/mm/yyyy")
        Else
            txtFechaNac.Text = ""
        End If
        rsLocal.Close
        Set rsLocal = Nothing
        Set rsLocal = CN.OpenResultset("SELECT id FROM CM_PACIENTES WHERE cliente_id = " & wCodExist, rdOpenKeyset, rdConcurReadOnly)
        If Not rsLocal.EOF Then
            PacienteId = Val(rsLocal!ID)
            wMsgExist = "El " & WS_TIPO & " " & ws_dni & " ya existe en el sistema (Cliente Nro " & wCodExist & ")." & Chr(13) & _
                        "Es paciente registrado del consultorio. Se cargaron sus datos para revision." & Chr(13) & Chr(13) & _
                        "Si es el mismo paciente, use GUARDAR para actualizar, o seleccionelo desde la lista del modulo de citas."
        Else
            PacienteId = 0
            wMsgExist = "El " & WS_TIPO & " " & ws_dni & " ya existe en el sistema (Cliente Nro " & wCodExist & ")" & Chr(13) & _
                        "pero aun no es paciente del consultorio." & Chr(13) & Chr(13) & _
                        "Puede usar GUARDAR para registrarlo como paciente sin duplicar el cliente."
        End If
        rsLocal.Close
        Set rsLocal = Nothing
        MsgBox wMsgExist, vbInformation, Pub_Titulo
        Exit Sub
    End If
    rsLocal.Close
    Set rsLocal = Nothing

    ws_token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiI1NTgiLCJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL3dzLzIwMDgvMDYvaWRlbnRpdHkvY2xhaW1zL3JvbGUiOiJjb25zdWx0b3IifQ.K8PwFsfNIpIl2ve0KJ2F08JZYLdGaBEx6_PvMRCm_Mw"

    Screen.MousePointer = 11
    DoEvents

    If WS_TIPO = "DNI" Then
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

    If WS_TIPO = "DNI" Then
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
        WS_ESTADO = ExtraerJSON(ws_response, "estado")
    End If

    If ws_nombre = "" Then
        MsgBox "No se encontro informacion para el " & WS_TIPO & ": " & ws_dni, vbExclamation, Pub_Titulo
        Exit Sub
    End If

    If WS_TIPO = "RUC" And Len(WS_ESTADO) > 0 Then
        If UCase(WS_ESTADO) <> "ACTIVO" And UCase(WS_ESTADO) <> "HABIDO" Then
            MsgBox "El RUC " & ws_dni & " esta en estado: " & WS_ESTADO & Chr(13) & _
                   "Verifique la informacion antes de continuar.", vbExclamation, Pub_Titulo
        End If
    End If

    ' Nuevo paciente desde la API: asegurar que no arrastre un cliente existente
    ClienteId = 0
    PacienteId = 0
    txtNombre.Text = ws_nombre
    If Len(ws_direccion) > 0 Then txtDireccion.Text = ws_direccion
    If Len(ws_fec) > 0 Then txtFechaNac.Text = ws_fec

    MsgBox WS_TIPO & " importado desde Factiliza:" & Chr(13) & _
           WS_TIPO & ": " & ws_dni & Chr(13) & _
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
    Dim wnombre As String
    Dim wDNI As String
    Dim wTel As String
    Dim wDir As String
    Dim wFec As String
    Dim wFecSql As String
    Dim wcod As Long
    Dim wCia As Long
    Dim wSql As String
    Dim rs As rdoResultset

    wnombre = Trim(UCase(txtNombre.Text))
    If Len(wnombre) < 3 Then
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

    ' Determinar si el DNI/RUC ya existe en CLIENTES (evita duplicados)
    ClienteId = 0
    If Len(wDNI) = 8 Then
        Set rs = CN.OpenResultset("SELECT TOP 1 CLI_CODCLIE FROM CLIENTES WHERE CLI_RUC_ESPOSA = '" & SQLC(wDNI) & "' AND CLI_CODCIA = " & wCia, rdOpenKeyset, rdConcurReadOnly)
        If Not rs.EOF Then ClienteId = Val(rs!cli_codclie)
        rs.Close
        Set rs = Nothing
    ElseIf Len(wDNI) = 11 Then
        Set rs = CN.OpenResultset("SELECT TOP 1 CLI_CODCLIE FROM CLIENTES WHERE CLI_RUC_ESPOSO = '" & SQLC(wDNI) & "' AND CLI_CODCIA = " & wCia, rdOpenKeyset, rdConcurReadOnly)
        If Not rs.EOF Then ClienteId = Val(rs!cli_codclie)
        rs.Close
        Set rs = Nothing
    End If

    If ClienteId > 0 Then
        ' Ya existe: actualizar CLIENTES
        wcod = ClienteId
        wSql = "UPDATE CLIENTES SET CLI_NOMBRE = '" & SQLC(wnombre) & "', CLI_NOMBRE_ESPOSO = '" & SQLC(wnombre) & "', " & _
               "CLI_TELEF1 = '" & SQLC(wTel) & "', CLI_CASA_DIREC = '" & SQLC(Left(wDir, 120)) & "', " & _
               "CLI_TRAB_DIREC = '" & SQLC(Left(wDir, 30)) & "', " & _
               "CLI_RUC_ESPOSO = '" & SQLC(IIf(Len(wDNI) = 11, wDNI, "")) & "', CLI_RUC_ESPOSA = '" & SQLC(IIf(Len(wDNI) = 8, wDNI, "")) & "', " & _
               "CLI_FECHA_NAC = " & wFecSql & " WHERE CLI_CODCLIE = " & wcod & " AND CLI_CODCIA = " & wCia
        CN.Execute wSql, rdExecDirect
    Else
        ' Nuevo cliente
        Set rs = CN.OpenResultset("SELECT ISNULL(MAX(CLI_CODCLIE), 0) + 1 AS nuevo FROM CLIENTES WHERE CLI_CP = 'C' AND CLI_CODCIA = " & wCia, rdOpenKeyset, rdConcurReadOnly)
        wcod = Val(Nulo_Valor0(rs!nuevo))
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
               wcod & ", " & wCia & ", 'C', '" & SQLC(wnombre) & "', '" & SQLC(wnombre) & "', '', '', " & _
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
               wCia & ", " & wcod & ", 'C', '" & SQLC(Left(wDir, 60)) & "', '" & SQLC(Left(wDir, 100)) & "', '', 0, 0, 0, 0, 0)"
        CN.Execute wSql, rdExecDirect
    End If

    ' Asegurar que exista como paciente CM_PACIENTES
    Set rs = CN.OpenResultset("SELECT id FROM CM_PACIENTES WHERE cliente_id = " & wcod, rdOpenKeyset, rdConcurReadOnly)
    If Not rs.EOF Then
        PacienteId = Val(rs!ID)
        rs.Close
        Set rs = Nothing
    Else
        rs.Close
        Set rs = Nothing
        CN.Execute "INSERT INTO CM_PACIENTES (cliente_id, estado) VALUES (" & wcod & ", 1)", rdExecDirect
        Set rs = CN.OpenResultset("SELECT @@IDENTITY AS nuevo_id", rdOpenKeyset, rdConcurReadOnly)
        If Not rs.EOF Then PacienteId = Val(rs!nuevo_id)
        rs.Close
        Set rs = Nothing
    End If

    ClienteId = wcod
    Nombre = wnombre

    Screen.MousePointer = 0
    Me.Hide
    Exit Sub
EH:
    Screen.MousePointer = 0
    MsgBox "Error al guardar el paciente: " & Err.Description, vbExclamation, Pub_Titulo
End Sub

Private Function SQLC(ByVal wvalor As String) As String
    SQLC = Replace(wvalor, "'", "''")
End Function

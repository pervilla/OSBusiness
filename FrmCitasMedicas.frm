VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCitasMedicas 
   Caption         =   "Reservar Cita Medica - Consultorio"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   9855
   Begin MSFlexGridLib.MSFlexGrid gridCamp 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      FixedCols       =   0
      BackColorBkg    =   16384
      AllowBigSelection=   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "PACIENTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   2745
      Width           =   9615
      Begin MSFlexGridLib.MSFlexGrid gridPac 
         Height          =   1575
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16384
         AllowBigSelection=   0   'False
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   480
         Left            =   4560
         Picture         =   "FrmCitasMedicas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   1215
      End
      Begin VB.TextBox txtBuscar 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   270
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "PAGO / COBRO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      TabIndex        =   6
      Top             =   5100
      Width           =   9615
      Begin VB.OptionButton optReserva 
         Caption         =   "Solo Reservar (paga el dia)"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   2850
      End
      Begin VB.OptionButton optPagado 
         Caption         =   "Cobrar y Reservar (deja pagado)"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   3150
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         Caption         =   "MONTO: S/ 0.00"
         Height          =   300
         Left            =   6400
         TabIndex        =   9
         Top             =   225
         Width           =   2900
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "COMPROBANTE DEL DIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   5865
      Width           =   9615
      Begin VB.OptionButton optBoleta 
         Caption         =   "Boleta"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   245
         Value           =   -1  'True
         Width           =   800
      End
      Begin VB.OptionButton optFactura 
         Caption         =   "Factura"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   245
         Width           =   900
      End
      Begin VB.Label lblComp 
         Caption         =   "Se emitira el dia de la cita"
         Height          =   255
         Left            =   2250
         TabIndex        =   13
         Top             =   245
         Width           =   3000
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   840
      Left            =   7800
      Picture         =   "FrmCitasMedicas.frx":170C0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6615
      Width           =   1815
   End
   Begin VB.CommandButton cmdReservar 
      Caption         =   "Reservar Cita"
      Height          =   840
      Left            =   3300
      Picture         =   "FrmCitasMedicas.frx":5F94A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6615
      Width           =   1800
   End
   Begin VB.CommandButton cmdNuevoPaciente 
      Caption         =   "Nuevo Paciente"
      Height          =   840
      Left            =   120
      Picture         =   "FrmCitasMedicas.frx":A547C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6615
      Width           =   1700
   End
   Begin VB.Label lblCamp 
      Caption         =   "CAMPANAS / HORARIOS DISPONIBLES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   4000
   End
   Begin VB.Label lblMsg 
      Height          =   255
      Left            =   5200
      TabIndex        =   15
      Top             =   6690
      Width           =   2450
   End
End
Attribute VB_Name = "FrmCitasMedicas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mHorarioId As Long
Private mPrecio As Currency
Private mCodArtServicio As Long
Private mFechaCita As String
Private mFechaCitaDate As Date
Private mHoraCita As String
Private mMedico As String
Private mEspecialidad As String
Private mCuposLibres As Integer
Private mPacienteId As Long
Private mClienteId As Long
Private mPacNombre As String
Private mPacDNI As String

Private Sub Form_Load()
    Caption = "Reservar Cita Medica - Consultorio"
    optReserva.Value = True
    optBoleta.Value = True
    Frame3.Enabled = False
    ConfigurarGrids
    LimpiarSeleccion
    CargarCampanias
    BuscarPacientes
End Sub

Private Sub ConfigurarGrids()
    With gridCamp
        .TextMatrix(0, 0) = "id"
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Inicio"
        .TextMatrix(0, 3) = "Fin"
        .TextMatrix(0, 4) = "Medico"
        .TextMatrix(0, 5) = "Especialidad"
        .TextMatrix(0, 6) = "Cupos"
        .TextMatrix(0, 7) = "Precio"
        .TextMatrix(0, 8) = "art"
        .TextMatrix(0, 9) = "fec"
        .TextMatrix(0, 10) = "hora"
        .ColWidth(0) = 0
        .ColWidth(1) = 900
        .ColWidth(2) = 700
        .ColWidth(3) = 700
        .ColWidth(4) = 2300
        .ColWidth(5) = 1800
        .ColWidth(6) = 600
        .ColWidth(7) = 800
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
    End With
    With gridPac
        .TextMatrix(0, 0) = "pid"
        .TextMatrix(0, 1) = "cid"
        .TextMatrix(0, 2) = "Nombre"
        .TextMatrix(0, 3) = "DNI"
        .TextMatrix(0, 4) = "Telefono"
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 5200
        .ColWidth(3) = 1100
        .ColWidth(4) = 1300
    End With
End Sub

Private Sub LimpiarSeleccion()
    mHorarioId = 0
    mPrecio = 0
    mCodArtServicio = 0
    mFechaCita = ""
    mFechaCitaDate = #1/1/1900#
    mHoraCita = ""
    mMedico = ""
    mEspecialidad = ""
    mCuposLibres = 0
    mPacienteId = 0
    mClienteId = 0
    mPacNombre = ""
    mPacDNI = ""
    lblMonto.Caption = "MONTO: S/ 0.00"
End Sub

Private Sub CargarCampanias()
    On Error GoTo EH
    Dim rs As rdoResultset
    Dim i As Integer
    Dim wCia As Long
    Dim wPrecio As Currency
    Dim wSql As String

    wCia = Val(LK_CODCIA)
    If wCia = 0 Then wCia = 25

    wSql = "SELECT h.id, h.fecha_especifica, h.hora_inicio, h.hora_fin, " & _
           " h.cupos_totales, h.cupos_ocupados, h.cod_art_servicio, " & _
           " ISNULL(m.apellidos + ' ' + m.nombres, '') AS medico, " & _
           " ISNULL(m.especialidad, '') AS especialidad, " & _
           " ISNULL(p.PRE_PRE1, 0) AS precio " & _
           " FROM CM_MEDICOS_HORARIOS h " & _
           " LEFT JOIN CM_MEDICOS m ON m.id = h.medico_id " & _
           " LEFT JOIN PRECIOS p ON p.PRE_CODART = h.cod_art_servicio AND p.PRE_FLAG_UNIDAD = 'A' AND p.PRE_CODCIA = " & wCia & " " & _
           " WHERE h.estado = 1 AND h.fecha_especifica >= CONVERT(date, GETDATE()) " & _
           " ORDER BY h.fecha_especifica, h.hora_inicio"

    gridCamp.Rows = 1
    Set rs = CN.OpenResultset(wSql, rdOpenKeyset, rdConcurReadOnly)
    Do Until rs.EOF
        gridCamp.Rows = gridCamp.Rows + 1
        i = gridCamp.Rows - 1
        wPrecio = Val(Nulo_Valor0(rs!PRECIO))
        If wPrecio = 0 Then wPrecio = 50
        gridCamp.TextMatrix(i, 0) = rs!ID
        gridCamp.TextMatrix(i, 1) = Format(Nulo_Valor0(rs!fecha_especifica), "dd/mm/yyyy")
        gridCamp.TextMatrix(i, 2) = Format(rs!hora_inicio, "hh:mm")
        gridCamp.TextMatrix(i, 3) = Format(rs!hora_fin, "hh:mm")
        gridCamp.TextMatrix(i, 4) = Trim(Nulo_Valors(rs!medico))
        gridCamp.TextMatrix(i, 5) = Trim(Nulo_Valors(rs!especialidad))
        gridCamp.TextMatrix(i, 6) = Val(rs!cupos_totales) - Val(rs!cupos_ocupados)
        gridCamp.TextMatrix(i, 7) = Format(wPrecio, "0.00")
        gridCamp.TextMatrix(i, 8) = Val(Nulo_Valor0(rs!cod_art_servicio))
        gridCamp.TextMatrix(i, 9) = Format(Nulo_Valor0(rs!fecha_especifica), "yyyy-mm-dd")
        gridCamp.TextMatrix(i, 10) = Format(rs!hora_inicio, "hh:mm")
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Exit Sub
EH:
    MsgBox "Error al cargar campanas: " & Err.Description, vbExclamation, Pub_Titulo
End Sub

Private Sub BuscarPacientes()
    On Error GoTo EH
    Dim rs As rdoResultset
    Dim wBusqueda As String
    Dim wSql As String
    Dim i As Integer

    wBusqueda = Trim(txtBuscar.Text)
    gridPac.Rows = 1

    If Len(wBusqueda) < 2 Then
        wSql = "SELECT TOP 200 p.id AS paciente_id, p.cliente_id, c.CLI_NOMBRE, " & _
               " ISNULL(c.CLI_RUC_ESPOSA, '') AS DNI, ISNULL(c.CLI_TELEF1, '') AS telf " & _
               " FROM CM_PACIENTES p INNER JOIN CLIENTES c ON c.CLI_CODCLIE = p.cliente_id " & _
               " WHERE c.CLI_CP = 'C' " & _
               " ORDER BY p.id DESC"
    Else
        wBusqueda = Replace(wBusqueda, "'", "''")
        wSql = "SELECT p.id AS paciente_id, p.cliente_id, c.CLI_NOMBRE, " & _
               " ISNULL(c.CLI_RUC_ESPOSA, '') AS DNI, ISNULL(c.CLI_TELEF1, '') AS telf " & _
               " FROM CM_PACIENTES p INNER JOIN CLIENTES c ON c.CLI_CODCLIE = p.cliente_id " & _
               " WHERE c.CLI_CP = 'C' AND (c.CLI_NOMBRE LIKE '%" & wBusqueda & "%' OR c.CLI_RUC_ESPOSA LIKE '%" & wBusqueda & "%' OR c.CLI_RUC_ESPOSO LIKE '%" & wBusqueda & "%') " & _
               " ORDER BY p.id DESC"
    End If

    Set rs = CN.OpenResultset(wSql, rdOpenKeyset, rdConcurReadOnly)
    Do Until rs.EOF
        gridPac.Rows = gridPac.Rows + 1
        i = gridPac.Rows - 1
        gridPac.TextMatrix(i, 0) = rs!paciente_id
        gridPac.TextMatrix(i, 1) = rs!cliente_id
        gridPac.TextMatrix(i, 2) = Trim(rs!CLI_NOMBRE)
        gridPac.TextMatrix(i, 3) = Trim(Nulo_Valors(rs!DNI))
        gridPac.TextMatrix(i, 4) = Trim(Nulo_Valors(rs!telf))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If gridPac.Rows = 1 And Len(wBusqueda) > 0 Then
        wSql = "SELECT 0 AS paciente_id, c.CLI_CODCLIE AS cliente_id, c.CLI_NOMBRE, " & _
               " ISNULL(c.CLI_RUC_ESPOSA, '') AS DNI, ISNULL(c.CLI_TELEF1, '') AS telf " & _
               " FROM CLIENTES c " & _
               " WHERE c.CLI_CP = 'C' AND NOT EXISTS (SELECT 1 FROM CM_PACIENTES p2 WHERE p2.cliente_id = c.CLI_CODCLIE) " & _
               " AND (c.CLI_NOMBRE LIKE '%" & wBusqueda & "%' OR c.CLI_RUC_ESPOSA LIKE '%" & wBusqueda & "%' OR c.CLI_RUC_ESPOSO LIKE '%" & wBusqueda & "%') " & _
               " ORDER BY c.CLI_CODCLIE DESC"
        Set rs = CN.OpenResultset(wSql, rdOpenKeyset, rdConcurReadOnly)
        Do Until rs.EOF
            gridPac.Rows = gridPac.Rows + 1
            i = gridPac.Rows - 1
            gridPac.TextMatrix(i, 0) = 0
            gridPac.TextMatrix(i, 1) = rs!cliente_id
            gridPac.TextMatrix(i, 2) = Trim(rs!CLI_NOMBRE) & "  (no registrado)"
            gridPac.TextMatrix(i, 3) = Trim(Nulo_Valors(rs!DNI))
            gridPac.TextMatrix(i, 4) = Trim(Nulo_Valors(rs!telf))
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If

    Exit Sub
EH:
    MsgBox "Error al buscar pacientes: " & Err.Description, vbExclamation, Pub_Titulo
End Sub

Private Sub gridCamp_Click()
    Dim i As Integer
    i = gridCamp.Row
    If i <= 0 Then Exit Sub
    If Val(gridCamp.TextMatrix(i, 6)) <= 0 Then Exit Sub
    mHorarioId = Val(gridCamp.TextMatrix(i, 0))
    mPrecio = Val(gridCamp.TextMatrix(i, 7))
    If mPrecio = 0 Then mPrecio = 50
    mCodArtServicio = Val(gridCamp.TextMatrix(i, 8))
    mFechaCita = gridCamp.TextMatrix(i, 9)
    mHoraCita = gridCamp.TextMatrix(i, 10)
    mMedico = gridCamp.TextMatrix(i, 4)
    mEspecialidad = gridCamp.TextMatrix(i, 5)
    mCuposLibres = Val(gridCamp.TextMatrix(i, 6))
    mFechaCitaDate = FechaDeSql(mFechaCita)
    lblMonto.Caption = "MONTO: S/ " & Format(mPrecio, "0.00")
End Sub

Private Sub gridPac_Click()
    Dim i As Integer
    i = gridPac.Row
    If i <= 0 Then Exit Sub
    mPacienteId = Val(gridPac.TextMatrix(i, 0))
    mClienteId = Val(gridPac.TextMatrix(i, 1))
    mPacNombre = gridPac.TextMatrix(i, 2)
    mPacDNI = gridPac.TextMatrix(i, 3)
    If Len(mPacDNI) = 11 Then
        optFactura.Value = True
    Else
        optBoleta.Value = True
    End If
End Sub

Private Function FechaDeSql(ByVal wFec As String) As Date
    Dim wY As String, wM As String, wD As String
    If Len(wFec) >= 10 Then
        wY = Left(wFec, 4)
        wM = Mid(wFec, 6, 2)
        wD = Mid(wFec, 9, 2)
        FechaDeSql = DateSerial(Val(wY), Val(wM), Val(wD))
    Else
        FechaDeSql = LK_FECHA_DIA
    End If
End Function

Private Function FechaSql() As String
    FechaSql = Format(mFechaCitaDate, "dd/mm/yyyy")
End Function

Private Sub cmdBuscar_Click()
    BuscarPacientes
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscarPacientes
    End If
End Sub

Private Sub optReserva_Click()
    Frame3.Enabled = False
End Sub

Private Sub optPagado_Click()
    Frame3.Enabled = True
End Sub

Private Sub cmdNuevoPaciente_Click()
    Load FrmPacienteNuevo
    FrmPacienteNuevo.Show 1
    If FrmPacienteNuevo.ClienteId > 0 Then
        mPacienteId = FrmPacienteNuevo.PacienteId
        mClienteId = FrmPacienteNuevo.ClienteId
        mPacNombre = FrmPacienteNuevo.Nombre
        mPacDNI = ""
        txtBuscar.Text = ""
        BuscarPacientes
        lblMsg.Caption = "Paciente registrado: " & FrmPacienteNuevo.Nombre
    End If
    Unload FrmPacienteNuevo
End Sub

Private Sub cmdReservar_Click()
    GuardarCita optPagado.Value
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub GuardarCita(ByVal pagar As Boolean)
    On Error GoTo EH
    Dim rs As rdoResultset
    Dim wSql As String
    Dim nCitaId As Long
    Dim wLocal As String
    Dim wTipoComp As String
    Dim wObs As String
    Dim wTicket As String

    If mHorarioId <= 0 Then
        MsgBox "Seleccione una campana / horario de la lista.", vbExclamation, Pub_Titulo
        Exit Sub
    End If
    If mClienteId <= 0 Then
        MsgBox "Seleccione o registre un paciente.", vbExclamation, Pub_Titulo
        Exit Sub
    End If
    If mCuposLibres <= 0 Then
        MsgBox "El horario seleccionado ya no tiene cupos disponibles.", vbExclamation, Pub_Titulo
        Exit Sub
    End If

    Screen.MousePointer = 11
    DoEvents

    If mPacienteId <= 0 Then
        CN.Execute "INSERT INTO CM_PACIENTES (cliente_id, estado) VALUES (" & mClienteId & ", 1)", rdExecDirect
        Set rs = CN.OpenResultset("SELECT @@IDENTITY AS nuevo_id", rdOpenKeyset, rdConcurReadOnly)
        If Not rs.EOF Then mPacienteId = Val(rs!nuevo_id)
        rs.Close
        Set rs = Nothing
    End If

    wLocal = Format(Val(pu_codCaja), "00")
    If wLocal = "00" Then wLocal = "01"

    ' Generar ticket y observaciones antes de insertar la cita
    wTicket = ""
    If pagar Then
        If optFactura.Value Then wTipoComp = "FACTURA" Else wTipoComp = "BOLETA"
        Set rs = CN.OpenResultset("SELECT ISNULL(MAX(id),0)+1 AS nxt FROM CM_PAGOS", rdOpenKeyset, rdConcurReadOnly)
        If Not rs.EOF Then wTicket = "TKT-" & Format(Val(rs!nxt), "000000")
        rs.Close
        Set rs = Nothing
        wObs = "PAGADO - TICKET " & wTicket & " - CANJE POR " & wTipoComp & " EL DIA DE LA CITA"
    Else
        wTipoComp = ""
        wObs = "PAGO PENDIENTE"
    End If

    wSql = "INSERT INTO CM_CITAS (paciente_id, horario_id, cliente_id, estado, orden, orden_atencion, " & _
           " total, saldo, fecha, hora, observaciones, local_origen, created_at) VALUES (" & _
           mPacienteId & ", " & mHorarioId & ", " & mClienteId & ", " & IIf(pagar, 1, 0) & ", 0, 0, " & _
           IIf(pagar, mPrecio, 0) & ", " & IIf(pagar, 0, mPrecio) & ", " & _
           "'" & FechaSql() & "', '" & mHoraCita & ":00', " & _
           "'" & Replace(wObs, "'", "''") & "', '" & wLocal & "', GETDATE())"
    CN.Execute wSql, rdExecDirect

    Set rs = CN.OpenResultset("SELECT @@IDENTITY AS nuevo_id", rdOpenKeyset, rdConcurReadOnly)
    If Not rs.EOF Then nCitaId = Val(rs!nuevo_id)
    rs.Close
    Set rs = Nothing

    ' Registrar el pago en CM_PAGOS con el numero de ticket (solo si cobra)
    If pagar Then
        wSql = "INSERT INTO CM_PAGOS (cita_id, monto, forma_pago, local_pago, usuario_cajero, ticket_nro, estado) VALUES (" & _
               nCitaId & ", " & Replace(CStr(mPrecio), ",", ".") & ", 'EFECTIVO', '" & wLocal & "', " & _
               "'CAJERO', '" & wTicket & "', 1)"
        CN.Execute wSql, rdExecDirect
    End If

    ' Actualizar cupos (incluye estado 4 Pendiente)
    wSql = "UPDATE CM_MEDICOS_HORARIOS SET cupos_ocupados = (SELECT COUNT(*) FROM CM_CITAS " & _
           " WHERE horario_id = " & mHorarioId & " AND estado IN (0,1,2,4)) WHERE id = " & mHorarioId
    CN.Execute wSql, rdExecDirect

    Screen.MousePointer = 0

    If pagar Then
        MsgBox "Cita registrada y pagada. Ticket: " & wTicket & ". Se imprimira la constancia para canjear por su " & wTipoComp & " el dia de la cita.", vbInformation, Pub_Titulo
        ImprimirTicket wTicket, wTipoComp
    Else
        MsgBox "Reserva confirmada. Pago pendiente de S/ " & Format(mPrecio, "0.00") & " para el dia de la cita.", vbInformation, Pub_Titulo
    End If

    RecargarTodo
    Exit Sub
EH:
    Screen.MousePointer = 0
    MsgBox "Error al guardar la cita: " & Err.Description, vbExclamation, Pub_Titulo
End Sub

Private Sub RecargarTodo()
    LimpiarSeleccion
    CargarCampanias
    If gridPac.Rows > 1 Then gridPac.Row = 1
End Sub

Private Sub ImprimirTicket(ByVal wTicket As String, ByVal wTipoComp As String)
    On Error GoTo EH
    Dim wEmp As String

    wEmp = "MEDINAFARMA"
    On Error Resume Next
    If Not GEN.EOF Then wEmp = Trim(GEN!GEN_NOMBRE)
    On Error GoTo EH

    Printer.Font.Name = "Courier New"
    Printer.Font.Size = 9
    Printer.Font.Bold = True
    Printer.Print "      " & wEmp
    Printer.Font.Bold = False
    Printer.Print "--------------------------------"
    Printer.Font.Bold = True
    Printer.Print "   TICKET DE RESERVA DE CITA"
    Printer.Font.Bold = False
    Printer.Print "--------------------------------"
    Printer.Print "TICKET Nro      : " & wTicket
    Printer.Print "EMISION         : " & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm")
    Printer.Print "--------------------------------"
    Printer.Print "PACIENTE        : " & mPacNombre
    If Len(mPacDNI) > 0 Then Printer.Print "DNI             : " & mPacDNI
    Printer.Print "MEDICO          : " & mMedico
    If Len(mEspecialidad) > 0 Then Printer.Print "ESPECIALIDAD    : " & mEspecialidad
    Printer.Print "FECHA DE CITA   : " & Format(mFechaCitaDate, "dd/mm/yyyy")
    Printer.Print "HORA DE CITA    : " & mHoraCita
    Printer.Print "--------------------------------"
    Printer.Print "MONTO PAGADO    : S/ " & Format(mPrecio, "0.00")
    Printer.Print "A CANJEAR POR   : " & wTipoComp
    Printer.Print "--------------------------------"
    Printer.Print ""
    Printer.Font.Bold = True
    Printer.Print " CONSERVE ESTE TICKET Y"
    Printer.Print " PRESENTELO EL DIA DE SU CITA"
    Printer.Print " PARA CANJEARLO POR SU " & wTipoComp
    Printer.Font.Bold = False
    Printer.Print ""
    Printer.Print "--------------------------------"
    Printer.EndDoc
    Exit Sub
EH:
    MsgBox "No se pudo imprimir el ticket: " & Err.Description, vbExclamation, Pub_Titulo
End Sub

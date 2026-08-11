# INSTRUCCIONES - Módulo Citas con Tickets (VB6)

## Contexto
El sistema web (MedinaFarma / CodeIgniter) ya maneja la reserva de citas médicas con este flujo:

1. **Reserva sin pago**: se inserta en `CM_CITAS` con `estado=0` (Inscrito), `saldo=monto`.
2. **Cobrar y reservar**: se inserta en `CM_CITAS` con `estado=1` (Confirmado/Pagado) **y además se registra el pago en `CM_PAGOS`** con un número de ticket (`TKT-000001`).
3. El ticket impreso es la constancia para canjear por el comprobante (Boleta/Factura) el día de la cita.
4. Estados de cita: `0`=Inscrito, `1`=Confirmado, `4`=Pendiente (exámenes), `2`=Atendido, `3`=Anulado.

## Objetivo
Actualizar `FrmCitasMedicas.frm` para que la reserva hecha en VB6 **registre el pago en `CM_PAGOS`** (igual que el web), imprima el ticket con el número `TKT-xxxxxx`, y que los reportes del web (`cmCitas/reporte`) y el botón de tickets del web reflejen esas reservas.

---

## Cambio 1: `GuardarCita` - registrar `CM_PAGOS` al cobrar

Reemplazar la sub `GuardarCita(ByVal pagar As Boolean)` completa por esta versión:

```vb
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

    If pagar Then
        If optFactura.Value Then wTipoComp = "FACTURA" Else wTipoComp = "BOLETA"
    Else
        wTipoComp = ""
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

    ' =====================================================
    ' NUEVO: registrar el pago en CM_PAGOS con ticket TKT-
    ' =====================================================
    If pagar Then
        Set rs = CN.OpenResultset("SELECT ISNULL(MAX(id),0)+1 AS nxt FROM CM_PAGOS", rdOpenKeyset, rdConcurReadOnly)
        wTicket = "TKT-" & Format(Val(rs!nxt), "000000")
        rs.Close
        Set rs = Nothing

        wSql = "INSERT INTO CM_PAGOS (cita_id, monto, forma_pago, local_pago, usuario_cajero, ticket_nro, estado) VALUES (" & _
               nCitaId & ", " & Replace(CStr(mPrecio), ",", ".") & ", 'EFECTIVO', '" & wLocal & "', " & _
               "'CAJERO', '" & wTicket & "', 1)"
        CN.Execute wSql, rdExecDirect

        wObs = "PAGADO - TICKET " & wTicket & " - CANJE POR " & wTipoComp & " EL DIA DE LA CITA"
        CN.Execute "UPDATE CM_CITAS SET observaciones = '" & Replace(wObs, "'", "''") & "' WHERE id = " & nCitaId, rdExecDirect
    Else
        wObs = "PAGO PENDIENTE"
    End If
    ' =====================================================

    ' Actualizar cupos: incluye el nuevo estado 4 (Pendiente)
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
```

> Nota: El `INSERT INTO CM_CITAS` original usa la variable `wObs` ANTES de asignarla. En la versión nueva se asigna `wObs` después del insert (para incluir el ticket) y luego se actualiza con un `UPDATE`. Si se prefiere, se puede calcular `wTicket` antes del INSERT y usarlo directamente.

---

## Cambio 2: `ImprimirTicket` - usar el número de ticket

Reemplazar la sub `ImprimirTicket` por:

```vb
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
```

---

## Cambio 3: `BuscarPacientes` - filtrar solo clientes `CLI_CP='C'`

En las dos consultas de `BuscarPacientes` (con y sin búsqueda) agregar el filtro:

```vb
" FROM CM_PACIENTES p INNER JOIN CLIENTES c ON c.CLI_CODCLIE = p.cliente_id " & _
" WHERE c.CLI_CP = 'C' " & _
```

En la consulta con `WHERE (c.CLI_NOMBRE LIKE ...)`, el filtro queda:

```vb
" WHERE c.CLI_CP = 'C' AND (c.CLI_NOMBRE LIKE '%" & wBusqueda & "%' OR c.CLI_RUC_ESPOSA LIKE '%" & wBusqueda & "%' OR c.CLI_RUC_ESPOSO LIKE '%" & wBusqueda & "%') " & _
```

---

## Cambio 4: Cupos con estado 4

Ya incluido en `GuardarCita` (Cambio 1): `estado IN (0,1,2,4)`. Si hay otros lugares que actualicen `cupos_ocupados`, aplicar el mismo cambio.

---

## Verificación
1. Compilar el proyecto VB6.
2. Abrir `FrmCitasMedicas`, elegir campaña + paciente, marcar "Cobrar y Reservar".
3. Verificar que se imprime el ticket con número `TKT-xxxxxx`.
4. En el web: `cmCitas/reporte` debe mostrar ese pago, y en `cmCitas/listado` el botón de Tickets de esa cita debe mostrar el `ticket_nro`.

## Tablas compartidas (misma BD `BDATOS`)
- `CM_CITAS` : la cita (estado 0/1/4/2/3)
- `CM_PAGOS` : el pago (monto, forma_pago, local_pago, ticket_nro, estado 1=Pagado, 2=Comprobante emitido, 3=Anulado)
- `CM_MEDICOS_HORARIOS` : campañas/horarios
- `CM_PACIENTES` / `CLIENTES` : pacientes

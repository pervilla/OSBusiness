# APARIENCIA - Módulo de Reserva de Citas (para replicar en VB6)

Este documento describe la apariencia de la reserva de citas en el sistema web,
para que `FrmCitasMedicas` se vea lo más parecido posible dentro de lo que permite VB6.

---

## 1. Modal "Reservar Cupo" (web)

El web usa un **modal** (ventana emergente) con estas zonas, de arriba a abajo:

### Encabezado
- Barra de color **azul** (`#007bff`) ocupando todo el ancho.
- Título en **blanco y negrita**: `Reservar Cupo - Dr(a). {apellidos}`.
- A la derecha un botón **"X"** para cerrar.

### Cuerpo (fondo blanco, secciones numeradas con separador)
1. **1. Seleccionar Campaña** (solo en el listado): un combobox con opciones tipo `22/08/2026 - APELLIDOS, NOMBRE`.
2. **2. Seleccionar Paciente**: un campo editable que autocompleta (busca por nombre o DNI) + un botón verde **"+ Registrar Nuevo Paciente"**.
3. **3. Tipo de Comprobante**: radios en línea:
   - `Boleta (Serie BC11)`
   - `Factura (Serie FC11)`
   - `Guía (Serie ...)`
4. **4. Servicios Adicionales (opcional)**: lista multiselección.

### Pie
- Botón gris **"Cancelar"**
- Botón azul claro **"Solo Reservar"** (sin pago)
- Botón **verde** **"Cobrar y Reservar"** (destacado)

---

## 2. Tarjetas de campaña (web `/cmCitas`)

Cada campaña se muestra como una **tarjeta**:
- Borde superior **azul**, esquinas redondeadas, sombra suave.
- Fecha en un **badge azul** (ej. `22/08/2026`).
- Cupos libres en un **badge verde** (disponibles) o **rojo** (lleno).
- Círculo con icono de médico + nombre **`Dr(a). APELLIDOS`** en negrita.
- Especialidad y nombre del servicio debajo.
- Franja con **Inicio | Fin | Inscritos**.
- Botón **verde** "Reservar Cita" al pie.

---

## 3. Qué se puede replicar en VB6

### Sí (fácil)
| Elemento web | Equivalente VB6 |
|--------------|-----------------|
| Encabezado azul | `Label` con `BackColor = &H00FF8000&` (azul), texto blanco negrita |
| Secciones numeradas | `Frame` con `Caption = "1. Seleccionar Paciente"` |
| Separadores | `Frame` vacío delgado o `Line` |
| Radios Boleta/Factura/Guía | `OptionButton` (ya existen) |
| Botón verde destacado | `CommandButton` con `Style = 1 - Graphical` y `BackColor = &H0000C000&` |
| Botón gris Cancelar | `CommandButton` con `BackColor = &H00C0C0C0&` |
| Texto DNI/Monto | `Label` con `Font.Bold = True` |

### Parcial (se puede aproximar)
| Elemento web | Aproximación VB6 |
|--------------|------------------|
| Combobox autocompleta (select2) | `TextBox` + `MSFlexGrid` filtrando al escribir (ya está implementado en `BuscarPacientes`) |
| Servicios multiselección | `ListBox` con `MultiSelect = 1 - Simple` (o varios `CheckBox`) |
| Tarjetas de campaña | `MSFlexGrid` (ya está) con colores de fila o `Frame`s apilados |
| Iconos (Font Awesome) | Imágenes `.ico`/`.bmp` pequeñas en los botones (`Picture`) o texto simbólico |

### No se puede (limitación de VB6)
- Esquinas redondeadas / sombras (CSS). Se aproxima con `Frame`s.
- Autocompletado tipo select2 real. Se usa TextBox + grid.
- Iconografía de fuentes. Se usan imágenes o texto.

---

## 4. Sugerencias concretas para que se vea parecido

1. **Franja azul superior**: agregar un `Label` ancho con `BackColor` azul y texto blanco negrita con el título, en lugar de depender solo del `Caption` del formulario.
2. **Monto resaltado**: `lblMonto` con texto verde/negrita (`Font.Bold = True`, `ForeColor = &H0000C000&`).
3. **Botón "Cobrar y Reservar" verde**: `cmdReservar` con `Style = 1` y `BackColor = &H0000C000&`, texto blanco.
4. **Botón "Solo Reservar"**: en vez de un solo `cmdReservar`, usar dos botones separados:
   - `cmdSoloReservar` (azul claro `&H00FFFFC0&`)
   - `cmdCobrarReservar` (verde)
   Cada uno llama a `GuardarCita False/True`.
5. **Secciones numeradas**: cambiar el `Caption` de los `Frame`s a `"1. PACIENTE"`, `"2. PAGO / COBRO"`, `"3. COMPROBANTE DEL DIA"` para imitar los pasos del web.
6. **Mensaje del comprobante**: `lblComp` mostrar la serie real, ej. `Se emitira el dia de la cita (Serie BC11)`.

---

## 5. Verificación visual
- Abrir `FrmCitasMedicas` y comparar con `http://localhost:8080/cmCitas/listado` → botón "Nueva Cita".
- Los elementos y el orden deben coincidir: Campaña → Paciente → Tipo de comprobante → (Servicios) → botones.

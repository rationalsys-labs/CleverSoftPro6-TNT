# Control de cargos

Abra **Operaciones → Control de Cargos**. La pantalla presenta bandas y acciones a la izquierda, documentos y ficha en el centro, y contadores e información a la derecha. Trabaja con los documentos recibidos mediante el importador central.

## Listas de trabajo

Existen cinco bandas numeradas, una personalizada, una de notas de crédito y otra de reparto a otros. Cada banda conserva su lista mientras la ventana permanece abierta. Son listas de trabajo del módulo, no una clasificación de importación/exportación de guías.

Seleccione la banda y el tipo FT/BV/NC/ND, escriba el **número completo** del documento importado y pulse Enter. Cada banda admite hasta 200 documentos, sin duplicados. La banda de notas de crédito admite únicamente NC. Reparto otros muestra además contacto, dirección, distrito y teléfono de entrega disponibles en el cliente.

Supr retira la fila seleccionada de la lista. **Vaciar banda** pide confirmación y solo descarta esa lista. F5 recarga las versiones y estados actuales de sus documentos. Doble clic o Enter en una fila muestra su historial. Al cerrar con listas cargadas se advierte que se descartarán; los registros guardados se conservan.

## Generar cargos

Pulse **Generar cargos**, seleccione un estado activo y escriba la fecha en formato `dd/mm/aaaa`. No se aceptan fechas futuras ni una gestión anterior a la emisión de un documento.

- **Registrar y ver reporte:** confirma todos los documentos de la banda en una sola transacción y abre su reporte.
- **Solo registrar:** confirma el mismo registro sin abrir la vista del reporte.
- **Vista previa:** muestra los documentos sin escribir estados ni reservar números. Se identifica como «Sin registrar».
- **Cancelar:** regresa a la banda.

Los documentos anulados externamente no se registran: retírelos de la banda. Las versiones de cada documento se verifican al confirmar. Si otro proceso los cambió, se rechaza el lote completo y se conserva la lista para actualizarla con F5. Un error no consume números ni deja documentos parcialmente registrados. El número definitivo y el correlativo del estado se reservan al guardar; el número mostrado en el panel es orientativo.

La gestión actual se obtiene del último evento activo por fecha e identificador. No se modifica la situación externa, los importes, el saldo ni los servicios facturados. Si el guardado se confirmó y falló la apertura del reporte o la actualización visual, se indica expresamente que no debe repetirse el guardado.

## Registros, anulación y reimpresión

**Ver / reimprimir** y **Vigencia de registros** abren los últimos 50 registros o permiten buscar uno por su número. La selección muestra sus documentos. La reimpresión conserva la fotografía guardada de cliente, importes, servicios y datos de entrega; cambios posteriores del maestro no alteran esa fotografía.

**Anular registro** pide confirmación, retira sus eventos de la gestión vigente y recalcula el estado de los documentos. No anula las facturas ni borra su historial. **Reactivar registro** vuelve a incorporar esos eventos, también con confirmación. Reactivar un registro antiguo no desplaza un evento de fecha posterior. Ambos cambios verifican la versión del registro y guardan usuario, fecha y vigencia en auditoría. Una reimpresión de un registro retirado lleva el aviso `[ANULADO]`.

Los reportes emplean Data Report de VB6 y presentan totales separados por moneda. **Exportar a Excel** entrega la banda actual; **Regenerar en Excel**, el detalle original del registro. Se abre un libro nuevo para que el usuario lo revise y guarde. Los importes mantienen valor numérico y los textos se escriben como texto, sin ejecutar fórmulas contenidas en nombres o direcciones. No se exportan identificadores internos ni versiones.

Se requiere el permiso 19 y el grupo Operaciones. Los estados disponibles proceden de `Int_StatusFac`. El esquema agrega `RegistrosCargos`, `DetallesCargos`, `EventosRegistrosCargos`, `SeriesCargos` y el estado inicial necesario para recalcular la gestión; los eventos de documentos continúan en `DbHisFac`.

Las pruebas incluyen la vista previa del DataReport y la exportación real a Excel. No certifican una impresión física en una impresora concreta.

# Diario y Compilatorio RP

## Diario: pizarra de trabajo

La pantalla presenta herramientas a la izquierda, estadísticas, grilla de trabajo, ficha desplegable y paneles inferiores de información/historial. **Consultar ficha** o Enter despliega los datos sin editar; **Asignar / modificar** habilita la edición. **Volver a la pizarra completa** cierra la ficha. Los eventos se introducen mediante un diálogo con selector de estado y comentario.

Abra **Cobranzas → Control de RP → Control Diario RP**. Las guías provienen del importador central; no hay un importador propio ni alta manual de guías en esta pantalla.

Busque por número o cliente, seleccione el grupo superior y active **Solo con salida** si corresponde. El selector se carga desde `Int_DivRp`: presenta el grupo general y únicamente los clientes separados expresamente. **El general excluye a esos clientes y a las guías sin asignar**. **Por asignar** permite trabajar estas últimas. No se carga todo el catálogo de clientes como grupos. Cada fila identifica guía, producto, gestión, salida, total y fecha. Enter consulta la ficha. Doble clic o espacio marca/desmarca la fila para una operación de lote. Buscar o actualizar limpia las marcas anteriores.

## Administrar grupos

El botón **Administrar grupos**, en el panel lateral gris del Diario, abre el mantenimiento. Cada grupo separado corresponde a un cliente existente; puede ser Local o Importado. Esta operación no edita el maestro del cliente.

1. Escriba parte del nombre y pulse **F9**, o **Buscar (F9)**.
2. La ventana presenta hasta 100 clientes activos coincidentes. Puede precisar el texto. Seleccione la fila completa y pulse Enter, haga doble clic o pulse Seleccionar.
3. Pulse **Agregar grupo**. Puede preparar varios cambios antes de guardar. No se admite repetir el mismo cliente.
4. Pulse **Guardar cambios** para confirmar el lote. Un error conserva los cambios pendientes y no deja modificaciones parciales.

**Eliminar grupo** advierte que, al guardar, todas las guías de ese cliente en el Diario volverán al grupo general. No borra guías ni cambia cliente, estados, asignaciones o historial. El retiro también queda pendiente hasta Guardar. El grupo general no se puede eliminar. **Cancelar cambios** o cerrar solicita confirmación antes de descartar el borrador.

Al volver al Diario se actualizan el selector y sus filas. La selección anterior se conserva si el grupo sigue existiendo; si fue eliminado, se regresa al general. Las guías futuras del cliente siguen la misma configuración, sin reasignaciones manuales por guía. El mantenimiento usa el permiso de Diario RP y comprueba al usuario también en SQL.

**Modificar ficha** permite asignar un cliente activo, moneda, pesos, neto, recargo, IGV, total e información. No modifica los datos del maestro de Clientes ni el número de guía. El total debe coincidir con neto + recargo + IGV; no se impone una tasa fiscal por suposición. **Guardar cambios** confirma la edición; **Cancelar**, Escape o cerrar solicitan descarte. Un error conserva el borrador. Cambios concurrentes se rechazan y exigen recargar.

**Activar/desactivar salida** cambia la marca operativa que permite preparar una cartilla. No significa que exista una factura externa. Los documentos RP 15D se cargan con salida activa; las RP 15N/48N, inactiva.

**Registrar evento** presenta los códigos y descripciones disponibles. Seleccione el estado e introduzca un comentario; se guarda el estado actual y su evento histórico. **Ver historial** muestra eventos y fechas.

## Confirmación y cierre

Marque las filas del lote y pulse **Confirmar cartilla**. Todas deben tener salida activa, cliente asignado, moneda y total positivo. Se reserva una sola planilla para todo el lote, se confirma su contenido en Compilatorio y se retiran las filas del Diario. La operación completa es transaccional.

Se conserva la etiqueta histórica **Facturado** del compilatorio para esta confirmación. Representa la preparación de la cartilla para facturación externa; no demuestra la existencia de una factura importada. Los documentos externos y sus servicios se cargan por separado.

**Cerrar gestión** es una vía diferente, para guías con estado COT, OTR o TNT. Conserva el resultado como **Finalizado** en Compilatorio y lo retira del Diario, sin generar una cartilla de facturación.

## Compilatorio

**Listado del compilatorio** presenta las guías confirmadas; **Estadísticas del Diario RP** abre un panel independiente con contadores y detalle de cada grupo. Los botones junto a los contadores filtran ese detalle. Estos datos son de consulta y no cambian la gestión.

Abra **Cobranzas → Control de RP → Compilatorio de Guías RP**. Presenta registros confirmados/finalizados; los registros pendientes que se prepararon al importar no se presentan como trabajo terminado.

Busque y consulte la ficha o el historial. **Ver / reimprimir cartilla** abre el Data Report original adaptado a los datos confirmados. Consultar o reimprimir no vuelve a guardar la operación ni consume otro correlativo. Cada cartilla conserva una copia de sus datos al confirmarse.

**Reactivar al Diario** solicita confirmación, devuelve la guía a la pizarra con salida desactivada y conserva los datos de su ficha. Limpia la confirmación actual del Compilatorio y conserva historial y cartillas anteriores. No crea otra guía. La importación repetida nunca realiza esta reactivación por sí misma.

F5 actualiza listas y datos. Las opciones respetan los permisos independientes de Diario y Compilatorio. Solo existe CRUD de clientes Locales en su mantenimiento; asignar una cuenta Importada a una guía RP no cambia su procedencia ni permite editar su maestro.

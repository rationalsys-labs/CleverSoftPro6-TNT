# Solicitudes Queries

Ruta: **Cobranzas → Solicitud de Queries**. El menú reúne Corrección de documentos y anulaciones, Regeneración de Queries emitidas y Finalizar Queries emitidas. Cada módulo permite trabajar con las solicitudes, reportes e historial del expediente.

## Corrección de documentos y anulaciones

Seleccione FT o BV, escriba el número del documento importado y presione **Enter**. Se admite `0099000001` o el número completo `FT-0099000001`; el tipo seleccionado determina el prefijo. Al cargar se muestra el número completo. Un prefijo distinto del tipo seleccionado se rechaza. La ficha muestra el cliente, moneda, importe y fecha; **Información de documento** consulta sus guías y servicios. Las solicitudes anteriores aparecen en la grilla superior; doble clic permite reimprimirlas. El botón junto a esa grilla alterna el historial con las guías seleccionadas.

La solicitud permite anulación directa, anulación con NC, descuento parcial o descuento al 100%. El descuento parcial debe ser positivo y menor que el importe del documento. Seleccione responsable, servicio y motivo: datos del cliente, tarifas, servicio, condiciones de facturación u otros. Los motivos de la categoría seleccionada se guardan con el expediente. Las propuestas de razón social, RUC, dirección y teléfono se incluyen en el reporte.

Puede asociar hasta seis guías del documento: escriba el número y presione Enter o el botón de ingreso; **Supr** retira la guía de la lista. La selección de liberar/no liberar guía expresa la instrucción solicitada al sistema origen. No modifica los amarres de facturación.

Las tres fechas usan **dd/mm/aaaa**; inicio ≤ visto ≤ solicitud ≤ hoy, y la solicitud no puede preceder a la emisión. El texto adicional admite 255 caracteres. **Guardar y ver solicitud** solicita confirmación, asigna un correlativo global `Q000000001` y abre el formato elegido. Otra pulsación reimprime el mismo expediente. **Nueva solicitud** inicia otro registro; confirma el descarte si había una solicitud sin guardar.

**Configurar casos Otros** permite agregar casos y darles baja lógica. Las solicitudes existentes conservan la descripción original. Los catálogos de responsables y servicios se toman de `IntRespQuerie` e `IntServFac` activos.

## Regeneración de Queries emitidas

El administrador consulta solicitudes abiertas, cerradas y anuladas. **Activar filtros** permite buscar por solicitud, documento o cliente, usuario, estado y rango de fechas; las fechas son opcionales. Se muestran hasta 200 registros ordenados del más reciente al más antiguo. Las estadísticas corresponden a los registros visibles. La búsqueda interpreta literalmente apóstrofes y símbolos.

La selección abarca toda la fila. Enter/doble clic o **Ver/reimprimir** abre el reporte guardado. **Solución e historial** consulta el cierre, la NC, los reemplazos y la auditoría. **Exportar lista a Excel** exporta los resultados actuales con sus estados y soluciones, preservando los números de documento como texto.

**Regenerar en masivo** agrega la fila a una banda sin duplicados. Con la banda abierta, Enter/doble clic agrega más solicitudes; Supr retira una fila. **Generar** abre los reportes según el formato registrado; una banda mixta genera ambas vistas. La banda admite hasta 200 solicitudes y se conserva al cambiar los filtros. **Limpiar** la vacía y cierra el panel.

Solo se puede **anular una solicitud abierta y vigente**, con motivo obligatorio y confirmación. Quedan registrados autor y fecha. No se reactiva ni se elimina el expediente. La anulación de la solicitud no anula la factura externa.

## Finalizar Queries emitidas

La pizarra presenta solicitudes abiertas y vigentes, con búsqueda por solicitud/documento/cliente (Enter o F9, o Actualizar). La selección abre el panel de anulación directa o el de NC. El doble clic consulta el reporte original.

En resoluciones con NC, escriba el número y presione Enter, o use **F9 / …** para elegir una nota importada. La búsqueda devuelve hasta 100 notas compatibles con el cliente, moneda y referencia de la solicitud, excluyendo notas ya usadas en otro cierre. La fila se selecciona completa y se acepta con Enter, doble clic o Aceptar. El importe aplicado debe ser positivo y no superar la NC ni lo solicitado; una anulación con NC o descuento al 100% exige cubrir todo el importe del documento. Se conserva el importe confirmado en la solución. El exceso mostrado es informativo.

Los documentos FT/BV de reemplazo son opcionales. Se agregan con número completo y Enter; deben existir, estar vigentes, ser del mismo cliente y moneda y diferir del documento original. Supr retira un documento de la lista. Se guardan sus referencias y una fotografía de sus datos.

Es obligatoria una descripción de la **solución confirmada por el sistema origen** (hasta 1000 caracteres). **Finalizar Query** solicita confirmación; registra el cierre, usuario, fecha, NC y reemplazos, y retira el caso de los pendientes. La solución queda accesible en el administrador de solicitudes cerradas. No se permite un segundo cierre ni anular una solicitud cerrada.

## Propiedad de los datos y persistencia

La 6.0 gestiona expedientes de corrección sobre documentos recibidos de otro sistema. No emite documentos de reemplazo ni NC. Crear, anular o finalizar una Query no modifica los importes, saldos, vigencia ni referencias económicas de `DBEmiFac`/`DbNCredito`, tampoco aplica NC contablemente, edita clientes importados, borra guías ni libera sus vínculos. Las propuestas y soluciones se conservan como instrucciones y evidencia de gestión.

`DbQuerie` mantiene la identidad histórica y agrega versiones de fila, actores, solicitud original y solución. `SeriesQueries` controla la numeración; `HistorialQueries` registra creación/anulación/cierre; `GuiasQuery` asocia las guías; `DocumentosResolucionQuery` conserva los reemplazos; `OtrosCasosQuery` mantiene motivos con baja lógica. Las FK conservan las referencias. La solicitud original y las fotografías de su reporte no se actualizan al cerrar el caso.

La escritura verifica permisos en SQL, incluidos los grupos **Cobranzas** y **Solicitud de Queries**, además del módulo correspondiente (5, 6 o 7). Se verifican versiones de factura, Query, NC y reemplazos para rechazar información desactualizada. La numeración y el guardado son transaccionales; una clave por solicitud evita duplicados al reintentar una operación ya confirmada. Se impide abrir simultáneamente el mismo tipo de solicitud y motivo para la misma factura, y utilizar una NC en más de un cierre.

Instalación: `BaseDatos/Esquema.sql`, incluido en el instalador de bases nuevas. Los scripts finales se destinan a una base nueva. Compatibilidad SQL Server 2008 (nivel 100). Los reportes originales se mantienen en `RepQueries` y `RepQueriesModelo`; la impresión física depende de la configuración de la impresora y no se ha certificado.

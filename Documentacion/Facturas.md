# Listado de facturas y documentos emitidos

Abra **Cobranzas → Listado de Facturas Emitidas**. La pantalla organiza la información en cuatro pestañas nativas de VB6, con los colores del sistema. Consulta los documentos que llegan por el importador central; no emite ni modifica sus datos externos.

- **Listado de documentos:** FT, BV, NC y ND. El selector limita el tipo; el texto busca por número, cliente o código. La búsqueda es literal, incluida cualquier comilla o comodín. Devuelve hasta 200 coincidencias, ordenadas por emisión; escriba más caracteres para precisar.
- **Datos e históricos:** número completo del documento, cliente, código, situación externa, gestión actual, fecha, guías y courier disponible. Los campos son cajas de texto bloqueadas: permiten seleccionar y copiar con Ctrl+C, sin editar ni borrar su contenido. Presenta el historial de estados con usuario y eventos retirados, además del historial de responsables existente. Ambos historiales utilizan los iconos incluidos en el formulario.
- **Notas relacionadas:** notas de crédito/débito y su documento de referencia. La selección resalta toda la fila; Enter o doble clic vuelve al listado de documentos con el tipo y número de esa nota filtrados de forma exacta. Las solicitudes y soluciones se gestionan desde Queries.
- **Guías y servicios:** todas las líneas vinculadas al documento, con origen, envío, producto y gestión de la guía.

Seleccione una fila completa. Enter o doble clic abre sus datos; **Listado de documentos** vuelve a la consulta. F5 actualiza. Escape vuelve al listado o cierra la ventana si ya está en él. La etiqueta Buscar, su caja y el selector de tipo solo son visibles en Listado de documentos. Las otras tres pestañas aprovechan ese espacio ampliando su contenido hacia arriba.

El saldo de una nota de crédito puede estar vacío: no se deduce ni se recalcula. Los importes y la situación externa se conservan como fueron importados. El estado de gestión proviene del historial de Cargos. Se respeta el permiso del módulo 11 y del grupo Cobranzas.

Las tablas principales son `DBEmiFac`, `DbNCredito`, `DbServGui`, `DbGuias`, `DbHisFac` y `DbHisCourier`. Las consultas usan sus identificadores relacionados, no una coincidencia del número entre tablas distintas.

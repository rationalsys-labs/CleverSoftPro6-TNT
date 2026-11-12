# Mantenimiento de clientes

La ficha y la barra inferior distinguen **Local** e **Importado**. Crear genera clientes locales con código automático `CS` y ocho dígitos, reservado al guardar. El código es inmutable. Modificar y Eliminar solo están disponibles para Locales; las cuentas Importadas se consultan y se actualizan desde el [importador central](Importacion.md). La demo incluye clientes Locales e Importados; DEMO0001 y EXTDEMO001 permiten consultar ambas procedencias.

Abra **Operaciones → Mantenimiento de Clientes**. El usuario debe tener acceso a Clientes y al grupo Operaciones.

**General** muestra clientes por código, razón social, RUC/DNI, dirección y estado. Puede buscar por razón social, RUC/DNI, dirección, teléfono o contacto. Cambiar el criterio vuelve a aplicar el texto escrito. Los caracteres especiales se buscan como datos.

Cada consulta muestra un máximo de **100 clientes**, también al filtrar. Escriba más caracteres para precisar la búsqueda; no hay paginación. La barra inferior indica las coincidencias encontradas y las filas mostradas. Código, identificación y estado tienen ancho suficiente para su contenido; razón social y dirección comparten el espacio restante. Al seleccionar con el mouse se resalta toda la fila. **Detalle del Cliente** presenta la ficha seleccionada. F5 recarga la consulta y los catálogos.

Con el foco en la grilla, pulse **Enter** sobre la fila seleccionada para abrir su detalle. Sin un cliente seleccionado no se abre ninguna ficha.

Mientras se muestra el detalle, el criterio de búsqueda, el texto del filtro e **Incluir eliminados** permanecen deshabilitados. Al regresar a **General** vuelven a habilitarse.

## Crear y modificar

Pulse **Crear** para una ficha nueva o seleccione una fila y pulse **Modificar**. El código local se asigna al guardar. Complete nombre o razón social, tipo de persona, identificación, ambas direcciones y territorio. El DNI exige 8 dígitos; el RUC, 11, con inicio 10 para persona natural y 20 para jurídica. También se admiten carné de extranjería y pasaporte para personas naturales, e identificación fiscal extranjera para jurídicas. Estas son validaciones de formato, sin consulta a registros externos. El recargo se expresa como porcentaje numérico no negativo.

En la dirección principal elija **Nacional** o **Internacional**. Nacional utiliza Perú y exige departamento, provincia y distrito; las listas dependen de la selección anterior. Internacional exige país y deshabilita las listas locales. La dirección de calle se completa en ambos casos.

La **dirección de entrega siempre es nacional (Perú)**: complete departamento, provincia, distrito y dirección de calle. No tiene selector de país ni radios nacional/internacional. Puede ser distinta de la ubicación principal, incluso si el cliente es internacional. Una ficha nueva propone Lima en ambas direcciones y Nacional para la principal.

Seleccione nivel (Regular, Bueno, Top o Premium), cobranza (TNT o Periferia), comprobante (Boleta o Factura) y moneda (Soles o Dólares). Son catálogos relacionados con el cliente. El territorio comercial es independiente de la ubicación geográfica.

Cada cliente tiene un territorio del catálogo, seleccionado en la lista. Para un cliente nuevo se exige territorio activo. Una ficha histórica puede conservar su territorio anterior aunque este ya no esté activo.

**Guardar** solicita confirmación con el código y el nombre del cliente. Un error conserva los datos para corregirlos. Si otro usuario modificó o eliminó el registro después de abrir la ficha, se rechaza la escritura para evitar sobrescribir su cambio.

**Cancelar** y Escape permiten abandonar la edición, previa confirmación de descarte. Rechazar el descarte conserva el borrador. Al cerrar el formulario durante una edición se solicita la misma confirmación.

## Eliminar

Seleccione un cliente activo y pulse **Eliminar**. La confirmación identifica código y razón social y explica que se realizará una baja lógica.

La baja solo cambia el estado del cliente. No borra ni modifica documentos, guías, comentarios u otras referencias a ese cliente. El registro queda eliminado y no tiene opción de reactivación ni modificación. Su código permanece reservado y no puede usarse para otro cliente.

Active **Incluir eliminados**, junto al criterio de búsqueda, para consultar también esas fichas. Las acciones Modificar y Eliminar permanecen deshabilitadas al seleccionar un eliminado. **Salir**, al pie del panel lateral, cierra el formulario.

Territorios, niveles de cliente, tipos de cobranza, monedas y tipos de documento se administran desde [Mantenimiento de Catálogos](Catalogos.md), en Configuración. Las nuevas asignaciones utilizan opciones activas; una ficha existente puede conservar su asignación anterior eliminada. Las nuevas monedas están disponibles al cargar las opciones de la ficha. El comprobante preferido sigue siendo Boleta o Factura.

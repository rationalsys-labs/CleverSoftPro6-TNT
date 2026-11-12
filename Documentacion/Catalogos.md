# Mantenimiento de catálogos

Abra **Configuración → Mantenimiento de Catálogos** con un usuario autorizado. La cuenta de demostración **DEMO** tiene permiso; **OPER** no lo tiene.

La lista superior permite elegir **Territorios**, **Niveles de cliente**, **Tipos de cobranza**, **Monedas** o **Tipos de documento**. Escriba en Buscar para filtrar la descripción o símbolo. **Incluir eliminados** permite consultar las bajas. F5 actualiza la lista.

Seleccione una fila para consultar su ficha. La grilla muestra código, descripción, abreviatura cuando corresponde y estado. La barra inferior indica modo de trabajo, código seleccionado y cantidad de registros.

## Crear y modificar

Pulse **Crear**, complete la descripción y confirme **Guardar**. En Monedas se registra el símbolo o denominación que mostrará la aplicación. El código numérico se asigna al guardar; no se reutilizan códigos de registros eliminados.

Para tipos de documento se exige además una abreviatura de dos letras o números, por ejemplo BV o FT. Una abreviatura ya existente se conserva al modificar para no cambiar la identificación de los documentos relacionados.

**Modificar** requiere seleccionar un registro activo. Territorios admite hasta 150 caracteres; los demás catálogos, 50. No se permiten descripciones repetidas dentro del mismo catálogo, incluidas las eliminadas, ni abreviaturas de documento repetidas.

Durante la edición se bloquean la selección del catálogo y los filtros. **Cancelar**, Escape o cerrar la ventana solicitan confirmación antes de descartar. Rechazarla conserva el borrador. Los errores de guardado también conservan lo escrito. Si otra sesión cambió o eliminó el registro, se rechaza la modificación y se solicita recargar.

Si el guardado se confirmó y falla la recarga, el mensaje identifica el código guardado: actualice con F5 sin repetir la escritura.

## Eliminar

Seleccione un registro activo y pulse **Eliminar**. La confirmación muestra catálogo, código y descripción. La baja afecta únicamente al registro del catálogo: conserva clientes, documentos y demás referencias. No hay reactivación ni modificación de eliminados.

## Relación con Clientes

Los catálogos activos se ofrecen para nuevas asignaciones. Un cliente puede conservar y consultar una asignación anterior aunque su catálogo se haya dado de baja. Al crear una ficha se cargan nuevamente las opciones vigentes.

Las monedas nuevas pueden seleccionarse y guardarse en Clientes. Su comprobante preferido continúa limitado a **Boleta o Factura**, según la regla de este mantenimiento; los otros tipos de documento pertenecen a las operaciones que los utilizan.

Países, departamentos, provincias, distritos, tipos de persona y tipos de identificación son catálogos fijos y no aparecen en esta pantalla.

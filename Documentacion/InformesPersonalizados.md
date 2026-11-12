# Informes personalizados

Ruta: **Informes → Informes personalizados** o el botón de informes de la barra principal. El generador permite construir un informe con los datos de Clever SoftPro 6.0.

## Uso

1. Seleccione el origen de datos.
2. Marque en el árbol las columnas que desea incluir. El resultado conserva el orden del catálogo de columnas.
3. Complete hasta cuatro filtros: columna, condición y valor. Los filtros se habilitan sucesivamente. Las columnas de filtro se eligen entre las marcadas.
4. Seleccione **Y** cuando deban cumplirse ambas condiciones u **O** cuando baste cualquiera. Se evalúan de izquierda a derecha: A O B Y C equivale a `(A O B) Y C`.
5. Use **Ver resultado**, **Generar Excel**, **Guardar Informe** (texto) o **Imprimir Informe**. Limpiar filtros conserva las columnas elegidas.

Las doce condiciones son igual, distinto, mayor, mayor o igual, menor, menor o igual, comienza/no comienza, termina/no termina y contiene/no contiene. Los caracteres `%`, `_` y `[` se buscan literalmente en las condiciones de texto.

Las fechas admiten `dd/mm/aaaa` o `aaaa-mm-dd`, con hora `hh:mm:ss` opcional. Una fecha sin hora comprende el día completo. Los valores numéricos se validan como números; se conservan los tipos de fecha e importe al exportar a Excel. Los códigos y números de documento permanecen como texto para conservar ceros iniciales, y los textos de origen no se ejecutan como fórmulas.

Los límites numéricos usan notación decimal y admiten hasta seis decimales, con el separador local de Windows en la pantalla. Si se supera esa precisión, se informa el error y se conserva el filtro para corregirlo; no se redondea ni se genera un resultado aproximado. La notación científica debe escribirse como decimal dentro de esa precisión.

Se admiten hasta 10000 filas por resultado. Si se supera el límite, se solicita precisar los filtros y no se entrega un resultado parcial. La vista permite desplazarse por filas y columnas. No se agregan totales que mezclen monedas diferentes.

## Orígenes y unidad de fila

| Origen | Una fila corresponde a | Permiso de origen |
| --- | --- | --- |
| Documentos emitidos | Un comprobante FT/BV/ND o NC | Listado de documentos |
| Clientes | Un cliente local o importado, incluido su estado | Mantenimiento de Clientes |
| Guías de importación | Una guía; incluye cantidad de servicios facturados | Diario RP |
| Diario RP | Un registro del diario | Diario RP |
| Compilatorio RP | Un registro confirmado o finalizado; excluye pendientes y reactivados, como su pantalla | Compilatorio RP |
| Queries | Una solicitud | Consulta/administración de Queries |
| Casos de quejas | Un caso, con su último estado y cantidad de gestiones | Quejas y Reclamos |
| Cheques rechazados | Un expediente | Cheques Rechazados |
| Órdenes de recojo | Una orden, con cantidad de documentos y último resultado de envío | Recojos |
| Cartas de cobranza | Una carta, con cantidades de emisiones y documentos | Cartas |
| Devoluciones | Un expediente, con cantidades de documentos y guías | Devoluciones |

Los detalles relacionados se cuentan o se consulta su último evento; no se unen indiscriminadamente al conjunto principal. Una carta con dos emisiones sigue siendo una fila, y un caso con varias gestiones tampoco se multiplica. Los informes no modifican clientes, guías, comprobantes, saldos ni gestiones.

## Permisos

Se requiere **Informes personalizados**, identificador nuevo 40 y código `000118`, más el permiso del origen y sus grupos superiores. SQL valida nuevamente la cuenta activa y todos esos permisos. El selector solo ofrece orígenes autorizados. Los administradores activos disponen de los orígenes vigentes; un operador debe recibir los permisos mediante Usuarios o su perfil. Conceder Informes personalizados por sí solo no concede acceso a otros módulos.

El contrato de login, Usuarios y Perfiles utiliza 40 posiciones. Se mantienen los identificadores previos. No se reutilizan Contados ni Complementos. Las asignaciones históricas se conservan; sus opciones retiradas quedan no vigentes y fuera del menú, la barra y los árboles de permisos.

## Salidas

- **Excel:** libro nuevo con las columnas elegidas, filtros automáticos, tipos numéricos y fechas. El usuario decide dónde guardarlo desde Excel.
- **Texto:** archivo Unicode UTF-16 con encabezados y columnas separadas por tabulaciones. Se usa el diálogo nativo Guardar; cancelar no genera archivo. El archivo anterior solo se reemplaza después de terminar correctamente la escritura temporal.
- **Impresión:** confirmación con cantidad de registros y salida mediante el objeto Printer de VB6. Presenta cada registro y divide líneas y páginas según el espacio disponible.

Guardar el archivo del resultado no guarda una plantilla reutilizable de selección y filtros. Las salidas utilizan la misma definición de consulta; cada acción vuelve a consultar los datos vigentes.

## Instalación

El esquema final y los catálogos de informes se instalan desde cero mediante la [guía de instalación](../BaseDatos/LEEME.md). No se requieren migraciones históricas. Incluye once vistas, el catálogo de orígenes, las consultas de generación y el permiso 40.

La consulta solo permite campos del origen validado. Los valores de filtro se parametrizan; no se acepta SQL libre del usuario. Los módulos son `src/MdiInformes.bas`, `src/FrmInformes.frm` y `src/FrmVistaInforme.frm`. Se conserva VB6, ADO, MSComctl/TreeView y MSFlexGrid. El diálogo de archivo usa la API de Windows, sin añadir COMDLG32.OCX.

Las consultas, filtros, selección de columnas y exportación se verificaron con datos demo. La impresión física depende de la impresora de destino.

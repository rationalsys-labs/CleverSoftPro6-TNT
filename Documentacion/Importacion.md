# Importación de información

Abra **Operaciones → Importación de información** con una cuenta autorizada. El sistema recibe datos del sistema externo; no emite facturas. El alcance es exclusivamente servicios de importación.

Use [Plantilla vacía](../Ejemplos/Plantilla_Importacion.xlsx) o [Ejemplo ficticio](../Ejemplos/Importacion_Demo.xlsx). Cada hoja tiene una fila de encabezados y los registros desde la segunda fila. No agregue títulos, subtotales, notas ni fórmulas dentro de los datos. Puede retirar las hojas que no vaya a cargar. Se admiten libros `.xls` y `.xlsx`, con Excel instalado, hasta 50.000 filas por hoja.

## Validar y confirmar

La pantalla conserva la organización del importador 5.0: árbol de proceso, barra de progreso, estadísticas y grilla inferior. El selector **Data a actualizar** cambia la vista previa entre Clientes, Guías, Documentos y Servicios; la validación y confirmación incluyen todas las hojas del libro. La vista previa presenta hasta 500 filas por hoja; el lote se valida completo. **Detalle de incidencias** permite consultar el resultado del proceso.

1. Seleccione el archivo mediante **Examinar**.
2. Pulse **Validar archivo**. Se comprueban las relaciones y las reglas de guardado sin confirmar escrituras ni consumir correlativos.
3. Ante un error, revise la hoja y fila indicadas, corrija el libro y vuelva a validar.
4. Pulse **Confirmar carga**. Se procesa la información validada en una sola transacción. Si falla una fila, no se confirma el lote.

Las dependencias se procesan en este orden: clientes, guías, documentos y servicios. Dentro de Documentos, coloque las facturas/boletas antes que sus notas. Una referencia también puede existir previamente en la base.

Los códigos de cliente, identificaciones, números de guía, documentos y referencias deben ser **celdas de texto**, conservando sus ceros iniciales. Una cadena que comience con `=` debe introducirse como texto literal; las fórmulas se rechazan. Fechas: celdas de fecha Excel o texto `AAAA-MM-DD`. Importes: celdas numéricas o texto numérico con el separador local de Windows, sin símbolos monetarios. No se ejecutan macros ni se actualizan vínculos externos.

## Clientes

La plantilla contiene el código externo `Asient`, razón social `CLIENT`, identificación `RUCDNI`, dirección `DIRECC`, teléfono `TELEF1`, contactos, entrega, recargo y los códigos de catálogo requeridos por la ficha actual.

| Campos | Uso |
|---|---|
| `IdTerritorio`, `IdNivel`, `IdCobranza` | Códigos de los catálogos de Clientes. |
| `IdComprobante` | 1: boleta; 2: factura. |
| `IdMoneda` | Código de moneda activa; inicialmente 1: soles, 2: dólares. |
| `IdTipoPersona` | 1: natural; 2: jurídica. |
| `IdTipoIdentificacion` | 1: DNI; 2: RUC; 3: carné; 4: pasaporte; 5: identificación fiscal extranjera. |
| `EsInternacional` | 0: dirección principal nacional; 1: internacional. |
| `IdDistrito`, `IdPais` | Solo distrito para nacional, o país extranjero para internacional. |
| `IdDistritoEntrega` | Distrito peruano de entrega, siempre obligatorio. |

Los códigos geográficos corresponden al catálogo histórico instalado; el ejemplo usa distrito 150101. No se deducen ubicaciones ambiguas por nombre ni se asigna Lima por defecto al importar.

Los clientes cargados quedan identificados como **Importados**. Las siguientes cargas pueden actualizar sus datos, conservando código, identidad y referencias. No se modifican ni se dan de baja desde Clientes. Un código que pertenece a un cliente Local produce conflicto: no se fusionan cuentas por RUC o razón social.

## Guías

`NumeroGuia`, `Origen`, `FechaEnvio`, `Consignatario`, `Gestion`, `Producto`.

El número, de hasta 10 caracteres, es único e inmutable. Gestión: **RP** o **Sender**. Productos: **15D** (documentos), **15N** y **48N** (muestras). El origen debe existir en el catálogo. LIM queda fuera del alcance; PEO y PEI también se rechazan por quedar fuera de los orígenes admitidos para esta carga.

Una guía idéntica ya existente no se duplica. Si sus datos difieren, se informa el conflicto y no se reinicia su gestión. Las RP generan Diario y registro pendiente de Compilatorio; las Sender permanecen como guías para su relación con servicios facturados. RP 15D inicia con salida activa; las otras RP, sin salida.

## Documentos y servicios

Documentos: `Tipo`, `Documento`, `Emision`, `Vencimiento`, `Cliente`, `Moneda`, `Subtotal`, `Recargo`, `IGV`, `Total`, `Situacion`, `Referencia`.

Tipos admitidos: **FT, BV, NC y ND**. El número completo lleva el prefijo, por ejemplo `FT-0099000001`, y admite hasta 14 caracteres. Situación: `ACT` o `ANL`. Cliente y moneda deben existir y estar activos. NC y ND requieren en Referencia una factura/boleta del mismo cliente. Los importes son los proporcionados por el sistema externo: no se calcula el total ni se emite un documento aquí.

Servicios: `Documento`, `Linea`, `NumeroGuia`, `Servicio`.

La línea es un entero positivo estable dentro del documento. Permite varias guías y servicios por documento. Documento y guía deben existir. Repetir la misma línea con los mismos datos no duplica; cambiar su contenido requiere revisión y se rechaza la carga.

Las longitudes se comprueban antes de convertir o buscar referencias: Tipo admite 2 caracteres; Documento y Referencia, 14; Cliente, 20; NumeroGuia, 10; Servicio, 50. Un identificador demasiado largo no se recorta ni se vincula por su prefijo. El error indica hoja, fila y campo; toda la carga se revierte, incluidos sus correlativos. La misma validación protege los demás textos del libro según el tamaño de su campo.

`ANL` identifica un comprobante anulado por el sistema emisor. Se conserva para consulta, pero no permite registrar Cargos, tanto en FT/BV/ND como en NC.

Los documentos existentes idénticos se conservan. Si una carga pretende cambiar sus datos, se informa el conflicto; no se sobrescriben saldos ni estados de cobranza. La ausencia de registros en un archivo no provoca bajas. La cartilla RP es una preparación de trabajo para facturación externa; su confirmación no crea la factura importada.

Cuando una carga se confirma, se registra usuario, archivo, fecha y cantidad de filas procesadas. No repita una escritura por un fallo posterior de presentación: vuelva a consultar el resultado. Una nueva validación es necesaria para leer cambios posteriores del archivo.

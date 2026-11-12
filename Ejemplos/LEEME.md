# Libros de importación

Las cuatro hojas admitidas son `Clientes`, `Guias`, `Documentos` y `Servicios`. Los encabezados deben conservar su nombre y posición, en la primera fila; los datos comienzan en la segunda. No añadir títulos ni instrucciones antes del encabezado. Una carga puede contener una o varias hojas.

| Libro | Uso |
| --- | --- |
| Plantilla_Clientes.xlsx | Maestro de clientes del sistema externo. |
| Plantilla_Guias.xlsx | Guías de importación RP/Sender, productos 15D, 15N y 48N. |
| Plantilla_Documentos.xlsx | Facturas FT, boletas BV, notas NC y ND; las notas requieren Referencia. |
| Plantilla_Servicios.xlsx | Amarre Documento + Linea + NumeroGuia + Servicio. |
| Plantilla_Importacion.xlsx | Las cuatro hojas vacías en un mismo libro. |
| Importacion_Demo.xlsx | Ejemplo de la carga inicial ya incluida en la demo; permite verificar reimportación idéntica. |
| Ejemplo_NuevaCarga.xlsx | Trece filas de un cliente, tres guías, cuatro comprobantes y cinco servicios nuevos. |

El ejemplo nuevo usa EXTNUEVO01, guías 0000088001–0000088003 y documentos con serie 0088. Se validó con el lector VB6/Excel real y la demo final, sin confirmar cambios. Conservar códigos y documentos como **texto**, incluidos sus ceros iniciales. Las fechas ISO `aaaa-mm-dd` también pueden ser texto. Los importes usan el separador numérico de Windows; no se admiten fórmulas.

Primero validar, revisar las incidencias y después confirmar. El importador procesa clientes, guías, documentos y servicios en ese orden dentro de la carga. Si se usa cada plantilla por separado hay que respetar ese orden y cargar antes las referencias de una nota. No existe carga independiente del Diario ni del Compilatorio RP.

La [guía de importación](../Documentacion/Importacion.md) documenta todos los campos, identificadores de catálogos y límites. El [recorrido demo](../BaseDatos/Demo.md) propone casos de consulta, continuación y rechazo. Para repetir una nueva carga confirmada con los mismos datos se comprueba idempotencia; para empezar otra demostración limpia se instala otra base nueva.
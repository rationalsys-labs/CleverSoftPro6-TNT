# Demostración de Clever SoftPro 6.0

Instalar desde cero con `-DatosDemo`. Datos y expedientes ficticios. Usuario DEMO, clave 123456, permite recorrer todos los módulos; OPER conserva el perfil de Operaciones.

| Área | Registros |
| --- | ---: |
| Clientes | 38 |
| Guias | 6 |
| Documentos | 21 |
| Servicios | 7 |
| Diario RP | 3 |
| Compilatorio confirmado | 1 |
| Queries | 6 |
| Casos quejas | 2 |
| Gestiones quejas | 4 |
| Cheques | 3 |
| Recojos | 2 |
| Cartas | 2 |
| Devoluciones | 3 |
| Cargos | 2 |

## Recorrido

1. Clientes: buscar DEMO0001 (local) y EXTDEMO001 (importado). DEMO0007 está de baja y conserva sus referencias. Crear un local genera su código CS; el importado admite consulta.
2. Catálogos: consultar niveles, territorios, monedas, documentos y cobranza. El nivel 4 está de baja en la demo y permanece en asignaciones anteriores.
3. Diario: guías RP pendientes de asignar y con salida; productos 15D, 15N y 48N. Hay dos clientes separados por grupos. La guía 0000000011 está confirmada en Compilatorio y tiene cartilla; reactivar conserva esa cartilla.
4. Documentos: FT-0099000001 tiene dos servicios y notas NC/ND relacionadas. FT-DEMO0008 está anulado externamente y Cargos lo identifica. Todos son datos ficticios; el sistema consulta los comprobantes y no los emite.
5. Cargos: consultar los dos registros, sus detalles y eventos de anulación/reactivación; utilizar documentos de otra banda para generar un cargo nuevo.
6. Queries: seis solicitudes con las cuatro clases de gestión, soluciones por NC, reemplazo, expediente anulado y caso pendiente. Los números y fechas se consultan en la grilla; las notas ya utilizadas no vuelven a ofrecerse.
7. Quejas: dos casos (uno cerrado CCR y otro abierto), con cuatro gestiones para consulta y continuación.
8. Cheques: tres expedientes, abiertos, cerrados y anulados. Recojos: modalidades cheque y factura con documento asociado e historial de intentos de correo ficticios.
9. Cartas: carta de facturas con dos emisiones y carta cerrada correspondiente a un cheque rechazado. Devoluciones: con guía, sin documentos y destinatario manual; estados abierto, cerrado y anulado.
10. Informes: once orígenes con datos. Elegir columnas, filtrar y exportar. El Compilatorio excluye pendientes y los filtros no redondean límites con más de seis decimales.
11. Importación: Ejemplo_NuevaCarga.xlsx contiene trece filas nuevas y válidas. Importacion_Demo.xlsx reproduce la carga inicial y permite comprobar la idempotencia. Cambiar un identificador a uno inexistente, alargar una guía, duplicar una línea con datos distintos o introducir fórmulas permite comprobar rechazos sin escribir parcialmente.

Las operaciones demo son manuales: volver a instalar con otro nombre permite repetirlas desde el estado inicial. El envío SMTP está deshabilitado por compatibilidad; los resultados del historial son datos ficticios, no pruebas de entrega. Las pruebas de versión obsoleta requieren abrir dos ediciones de un registro y guardar una primero. No se incluyen ejecutores de tests en el paquete.

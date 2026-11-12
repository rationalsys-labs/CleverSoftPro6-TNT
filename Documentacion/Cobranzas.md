# Quejas, cheques, recojos, cartas y devoluciones

Los cinco módulos permiten registrar expedientes, consultar sus eventos y emitir los formatos de cobranza mediante formularios y Data Report de VB6. El envío SMTP está deshabilitado por compatibilidad.

## Operación

- **Quejas:** cliente en seguimiento → casos → gestiones. F9 selecciona un cliente existente. Caso y primer evento se guardan juntos; cada gestión conserva contacto, teléfono, comentario y estado. `CCR` cierra el caso e impide eventos posteriores. Configuración permite agregar tipos/estados y darlos de baja con Supr, conservando su historia; CCR no admite baja.
- **Cheques rechazados:** cliente con F9, banco, número, giro, motivo, moneda, monto y gastos. El total se recalcula automáticamente. Registrar guarda expediente y genera carta; finalizar/anular exige motivo. Se rechaza el mismo cliente/banco/número/giro. Reimprimir y exportar conservan el expediente.
- **Recojos:** modalidades cheque y factura, destinatarios y mensaje editables. Las facturas FT/BV importadas se agregan por número y Enter; Supr retira una asociación antes de guardar. Registrar y enviar guarda una única orden y muestra «Envío SMTP deshabilitado por compatibilidad.». Reenviar muestra el mismo aviso, sin generar intentos de transporte. El historial conserva los registros existentes.
- **Cartas:** primera carta para cheque rechazado o facturas pendientes; cliente, dirección, contacto/cargo, referencia, moneda, monto y documentos asociados. La segunda emisión corresponde a facturas pendientes y se registra una sola vez. Cada emisión conserva su contenido y configuración institucional. Ver / reimprimir no crea otra emisión. Solicitar el aviso al guardar, tras la segunda emisión o al reenviar muestra «Envío SMTP deshabilitado por compatibilidad.». Se mantienen los tres modelos de carta originales.
- **Devoluciones:** expediente de devolución de dinero de servicios de importación, motivo, forma, banco/cuenta, moneda, monto, contacto, cancelaciones/abonos y carta. Hasta seis documentos y seis guías de diez caracteres. Las guías deben existir y, si se indican documentos, estar relacionadas con ellos. Sin documentos exige marcarlo expresamente. Destinatario manual no crea ni asigna un cliente ficticio. Finalizar registra el resultado externo; anular conserva el historial y se identifica en el papel.

Los listados devuelven hasta 200 filas y seleccionan la fila completa. F9 devuelve hasta 100 clientes activos por nombre, código o RUC, tanto locales como importados. Los formularios usan el centrado común. Los reportes permiten vista previa, reimpresión e impresión; la salida física depende de la impresora instalada.

## Correo

El registro de órdenes y cartas conserva sus destinatarios y mensajes. Al solicitar el envío se muestra «Envío SMTP deshabilitado por compatibilidad.». No se transmite correo, no se inicia ningún programa externo y no se agrega una solicitud pendiente ni un resultado Enviado. Los historiales demo contienen únicamente ejemplos ficticios.

El administrador mantiene la información institucional, los parámetros de correo y los destinos generales o por territorio desde Configuración. Esos datos se conservan aunque el transporte esté deshabilitado.

## Persistencia y permisos

`Esquema.sql` mantiene las tablas `DbQuejas_Cli/Cas/Ges`, `DbCheques`, `DbRecojos`, `DbsCartas` y `DbDevol`; no conecta estos formularios al modelo paralelo `Mdb*`. Agrega versiones, solicitud original, actores y cierre/anulación lógica.

`SeriesCobranzas` reserva CAS/GES/CCR/RCJ/CRT/CDV dentro de cada transacción. `ClavesCobranzas` evita duplicaciones por reintentos. `HistorialCobranzas` registra actor, fecha y resultado. `DocumentosCobranzas`, `GuiasDevolucion` y `EmisionesCarta` usan FK numéricas y conservan fotografías. La segunda carta toma la configuración institucional de su emisión sin alterar la primera.

SQL comprueba cuenta activa y permiso de módulo 12–16 junto con su grupo Cobranzas. Solo un administrador guarda configuración; los demás la consultan sin recuperar la contraseña. Las escrituras comprueban versiones de clientes, documentos y expedientes. Las bajas conservan las referencias existentes.

Estos módulos no emiten facturas, aplican notas, efectúan pagos ni cambian saldos, importes, clientes o amarres de guías del sistema emisor. Las confirmaciones económicas se registran como gestión local; los comprobantes siguen entrando por el importador central.

## Compilación y validación

Compilar Clever.vbp con VB6 o ejecutar Compilar.ps1. Se genera únicamente Clever.exe. La instalación utiliza el esquema consolidado de SQL Server 2008 y uno de los archivos de datos.

El registro, cierre, reportes e historial de los expedientes permanecen disponibles. La impresión física depende de la impresora instalada.

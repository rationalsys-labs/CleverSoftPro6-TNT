# Inicio de Clever SoftPro 6.0

Clever SoftPro 6.0 fue concebido y programado en **2011** por **Juan Pablo Aguilar Campos** para **Correo Privado S.A.C., nombre comercial TNT**. Trabaja con datos del sistema principal de la empresa para apoyar a Cobranzas, Operaciones, Importaciones y Facturación.

## Acceso

El arranque presenta la imagen de Clever SoftPro durante cinco segundos, muestra el login y, tras validar la cuenta, abre la ventana principal MDI.

Seleccione su usuario e indique la clave. **Enter** acepta y **Escape** cancela. **F5** actualiza los usuarios y permite reintentar una conexión fallida. El usuario de la aplicación es independiente de la cuenta Windows utilizada para conectar a SQL Server.

La conexión utiliza **SQL Server 2008**, base `CleverSoftPro6` en `localhost\SQLEXPRESS`, proveedor SQLOLEDB.1 y autenticación Windows. La cuenta debe estar activa y su clave se valida mediante MD5. En una instalación con datos demo se puede acceder con **DEMO / 123456** o **OPER / 123456**.

Los permisos se cargan al iniciar sesión y se comprueban con sus grupos superiores antes de abrir cada módulo. Un permiso ausente o inválido no concede acceso. Al salir de la sesión se limpian el usuario y sus permisos. La administración de usuarios y perfiles corresponde al administrador.

## Ventana principal

Los menús y la barra ofrecen las funciones autorizadas para la cuenta. El sistema reúne el importador central, Diario y Compilatorio RP, documentos y cargos, los tres módulos de Queries, expedientes de cobranza, informes personalizados y mantenimientos.

En **Configuración** se encuentran Catálogos, Perfiles de acceso y Control de Usuarios. Los tipos de documento se mantienen en Catálogos; los formatos de importación se describen en su [guía](Importacion.md).

**Informes → Informes personalizados** permite elegir datos, columnas y filtros. Véase [Informes personalizados](InformesPersonalizados.md).

En **Ayuda**, Contenido presenta una guía breve de uso. Acerca de Clever SoftPro muestra la versión interna del ejecutable, su propósito, tecnologías, autor y correo de contacto.

Las acciones de correo de Recojos y Cartas conservan el registro de los expedientes y muestran **«Envío SMTP deshabilitado por compatibilidad.»** al solicitar un envío.

## Posición de las ventanas

Los módulos se abren centrados en el área de trabajo del MDI. Los diálogos se centran sobre su ventana propietaria y respetan el área disponible del monitor. Las cartillas RP también se centran. Si un formulario supera el tamaño de su contenedor, se conserva accesible su borde superior izquierdo.

Volver a solicitar un módulo desde el menú aplica el centrado; cambiar de foco conserva la posición que haya elegido el usuario.

## Recursos y componentes

Los recursos gráficos de presentación, cabecera del login y fondo MDI se conservan en `Imagenes/`. Las imágenes utilizadas por los formularios están incorporadas en sus recursos FRX. La versión interna de la aplicación es 6.0.1.

Para abrir y compilar el proyecto se necesita Visual Basic 6.0 y sus componentes COM de 32 bits registrados, incluidos los controles indicados por `Clever.vbp`. HookMenu debe estar registrado en el equipo de destino; no se distribuye un OCX de terceros sin su licencia.

`src/ChameleonButton.ctl` conserva el aviso de Gonchuki y sus condiciones originales. `src/MD5.cls` conserva los créditos de su implementación. Los avisos de terceros permanecen en sus fuentes.

## Organización VB6

- `src/ModInicio.bas`: inicio de la aplicación.
- `src/frmMain.frm`: presentación.
- `src/FrmLogin.frm`: acceso de usuarios.
- `src/ModSesion.bas`: sesión, permisos y apertura de módulos.
- `src/MdiConect.bas`: conexión ADO y autenticación de la aplicación.
- `src/MdiPri.frm`: ventana principal, menús, barra y ayuda.
- `src/ModVentanas.bas`: centrado común de formularios y reportes.

Para preparar una instalación nueva, consulte la [guía SQL](../BaseDatos/LEEME.md) y el [recorrido de demostración](../BaseDatos/Demo.md).

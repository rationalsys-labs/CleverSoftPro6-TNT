# Perfiles de acceso

Abra **Configuración → Perfiles de acceso** con un administrador. Este mantenimiento reúne el nombre del perfil y su conjunto de permisos. Sustituye las opciones separadas de plantillas y niveles de acceso del menú anterior.

La grilla muestra código, nombre, estado y usuarios vinculados, incluidas las cuentas inactivas que conservan esa relación. Busque por nombre o active **Incluir eliminados**. Enter o doble clic abren la ficha. **Volver al listado** regresa a la grilla. F5 actualiza la consulta.

## Crear y modificar

Pulse **Crear perfil**, escriba un nombre y marque sus permisos en el árbol. Por ejemplo, puede crear «Operaciones» y marcar Mantenimiento de Clientes; se activa automáticamente el grupo del que depende. Las opciones retiradas del sistema no se ofrecen en el árbol de permisos.

El nombre admite hasta 50 caracteres y no puede repetirse, incluso si el perfil anterior está eliminado. Personalizado, Administrador y Acceso completo son nombres reservados. El código numérico se asigna al guardar.

Archivo y Salir se conservan para permitir cerrar sesión. La gestión de usuarios y perfiles pertenece al administrador y no se concede mediante un perfil de operador.

**Guardar** pide confirmación e informa cuántos usuarios están vinculados. Nombre y permisos se guardan juntos. Modificar un perfil afecta a todos los usuarios que siguen vinculados, al iniciar sesión nuevamente; no afecta a quienes tienen accesos personalizados.

**Cancelar**, Escape o cerrar la ficha solicitan confirmación antes de descartar el borrador. Rechazarla conserva lo escrito. Un error también conserva el borrador. Si otro administrador modificó el perfil, se exige recargar. En consulta se puede recorrer el árbol, sin cambiar sus casillas.

Si se confirmó el guardado y falló la recarga, el mensaje identifica el código guardado: pulse F5 sin repetir la operación.

## Asignar y personalizar

En Usuarios, **Tipo de usuario** distingue Administrador y Operador. Para el operador, **Perfil de acceso** permite elegir un perfil activo o **Personalizado**.

Al elegir un perfil se muestran sus accesos. Guarde la ficha para confirmar la asignación. Si cambia una casilla mientras la cuenta usa un perfil compartido, se solicita confirmar el paso a Personalizado. Se toman sus permisos actuales como punto de partida; los cambios quedan solo para esa cuenta y no modifican el perfil original. Cancelar restaura la asignación anterior.

El administrador mantiene acceso completo y no requiere un perfil compartido. Si un perfil cambia mientras se edita una cuenta vinculada, es necesario recargar para evitar guardar una personalización basada en permisos anteriores.

## Eliminar

Seleccione un perfil activo y pulse **Eliminar perfil**. La confirmación identifica código, nombre y usuarios vinculados. La baja es lógica y afecta únicamente al perfil: conserva sus permisos y los vínculos de usuarios existentes. No hay reactivación ni reutilización de su nombre.

Un perfil eliminado deja de ofrecerse para nuevas asignaciones. Una cuenta que ya lo utilizaba conserva sus accesos y puede mantenerlo; su ficha lo identifica como eliminado. Si se necesitan accesos distintos para esa cuenta, puede pasar a Personalizado o elegir otro perfil activo.

La base demo incluye **Operaciones**, vinculado a **OPER** y preparado con sus permisos anteriores. **DEMO** mantiene el tipo Administrador con acceso completo. Puede crear otros perfiles y asignarlos expresamente. F5 actualiza el listado de perfiles.

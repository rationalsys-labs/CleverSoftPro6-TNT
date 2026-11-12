# Usuarios y permisos

Abra **Configuración → Control de Usuarios → Administrar Usuarios**. Solo puede hacerlo un administrador autorizado. En la demostración utilice **DEMO**, clave **123456**. La baja se realiza desde la ficha de la cuenta.

La pantalla abre con el listado de usuarios. La ficha conserva la organización del panel 5.0: **Datos Generales** y **Accesos de Usuario**, dentro del mismo contenedor, con acciones laterales y Guardar/Cancelar al pie.

## Consultar y seleccionar

Busque por usuario, nombres o apellidos; **Incluir eliminados** muestra también las cuentas dadas de baja. La grilla muestra código, usuario, nombres y apellidos, perfil y estado. Al seleccionar se resalta toda la fila.

Pulse **Enter**, haga doble clic o use **Detalle del usuario** para abrir la ficha seleccionada. Una búsqueda sin resultados limpia la selección y no permite abrir una cuenta anterior. **Volver al listado** regresa a la grilla y habilita los filtros.

Dentro de la ficha, Datos Generales muestra usuario, nombres, apellidos, correo, perfil y estado. Accesos de Usuario muestra el árbol de menús. En consulta puede desplazarse, expandir ramas y seleccionar nodos sin modificar sus casillas. La barra inferior indica modo de trabajo y cuenta seleccionada. F5 actualiza el listado y la cuenta seleccionada en consulta.

## Crear y modificar

Pulse **Crear usuario**. Complete usuario de acceso, nombres, apellidos, tipo de usuario, clave y confirmación. El correo es opcional. Una cuenta nueva admite de uno a cinco caracteres de usuario, letras o números; se guarda en mayúsculas. Los usuarios eliminados conservan su nombre y no se reutilizan.

El **Administrador** tiene todos los accesos de menús. Para un **Operador**, elija un [Perfil de acceso](Perfiles.md) o Personalizado. Un perfil compartido proporciona sus permisos; en Personalizado se guardan accesos individuales. Al cambiar una casilla de una cuenta vinculada se pide confirmar el paso a Personalizado, sin modificar el perfil compartido. Marcar una opción activa sus grupos superiores; desmarcar un grupo desmarca sus opciones. Archivo y Salir se conservan para permitir cerrar sesión. La administración de usuarios y perfiles queda reservada al administrador.

El árbol presenta las opciones vigentes. Las opciones retiradas del menú no conceden acceso ni se ofrecen para nuevas asignaciones.

**Modificar usuario** permite cambiar nombres, apellidos, correo, tipo, perfil de acceso y permisos individuales. El nombre de acceso existente se conserva. Los cambios del árbol permanecen en el borrador hasta pulsar **Guardar** y confirmar. Datos, asignación y permisos se guardan como una sola operación: un error no debe dejar cambios parcialmente aplicados.

**Cancelar**, Escape o cerrar la ventana solicitan confirmación antes de descartar una edición. Rechazarla conserva lo escrito. Cancelar regresa al listado y restaura los datos y permisos de la fila seleccionada. Un error de guardado también conserva el borrador. Si otro administrador cambió la cuenta, se rechaza la versión anterior y se solicita recargar. Al guardar, la grilla se actualiza y conserva seleccionada la cuenta guardada.

Si el guardado se confirmó y falla la recarga, el mensaje indica el código guardado: seleccione nuevamente esa cuenta sin repetir la escritura.

Los permisos y el perfil de otras sesiones se cargan al iniciar sesión nuevamente. El mantenimiento no cierra por la fuerza las aplicaciones de otros usuarios.

## Cambiar clave

Seleccione una cuenta activa y pulse **Cambiar clave**. Complete la clave nueva y su confirmación, y guarde. Esta operación no modifica sus datos ni permisos. No se muestra la contraseña anterior ni su resumen MD5.

Las claves de aplicación se conservan mediante MD5, como se estableció para este proyecto histórico. Son independientes de la autenticación de Windows utilizada para conectar con SQL Server.

## Eliminar usuario

Seleccione una cuenta activa y pulse **Eliminar usuario**. La confirmación identifica usuario, nombres y apellidos. La baja bloquea los siguientes inicios de sesión; conserva la cuenta, sus permisos y las referencias históricas. No hay eliminación física ni reactivación desde el módulo.

El administrador no puede dar de baja su propia cuenta ni quitarse la administración durante su sesión. El sistema conserva al menos un administrador activo, incluso ante operaciones simultáneas. La cuenta de sesión sí puede cambiar su clave.

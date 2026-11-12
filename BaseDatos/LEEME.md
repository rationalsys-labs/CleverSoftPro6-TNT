# Instalación de la base de datos

El paquete instala el estado final de Clever SoftPro 6.0 desde una base nueva, sin respaldos privados ni migraciones históricas.

| Archivo | Contenido |
| --- | --- |
| Esquema.sql | 122 tablas con PK, FK, restricciones e índices; vistas, funciones y procedimientos finales. |
| DatosBase.sql | INSERT de catálogos y configuración inicial, sin cuentas ni expedientes. |
| DatosDemo.sql | INSERT completos de catálogos, cuentas y expedientes ficticios de todos los módulos. Reemplaza la elección DatosBase; no se ejecutan ambos. |
| Instalar.ps1 | Crea una base nueva y aplica esquema y datos en una transacción. |
| Demo.md | Recorrido y escenarios disponibles. |

Desde la raíz, con Windows autorizado en SQL Server:

```powershell
& .\BaseDatos\Instalar.ps1 -Base CleverSoftPro6 -DatosDemo
```

Sin `-DatosDemo` se instalan catálogos sin cuentas de aplicación; no es la opción para recorrer la demo. El primer administrador de una instalación sin demo debe aprovisionarlo su responsable SQL; no se incorpora una cuenta oculta. El ejecutable utiliza el nombre fijo CleverSoftPro6 en localhost\SQLEXPRESS.

El parámetro opcional `-CarpetaDatos` indica una carpeta accesible por el servicio SQL Server. Si se omite se utilizan las ubicaciones predeterminadas de la instancia. La base o sus MDF/LDF no deben existir: no se sobrescriben. Ante un error de esquema/datos se revierte la transacción y se conserva la base vacía para diagnóstico.

La configuración utiliza intercalación Modern_Spanish_CI_AS, nivel de compatibilidad 100, recuperación SIMPLE y PAGE_VERIFY CHECKSUM. Los scripts se destinan a SQL Server 2008 y utilizan su gramática T-SQL.

La demo tiene DEMO y OPER con clave 123456, datos enteramente ficticios y envío SMTP deshabilitado por compatibilidad. `DatosDemo.sql` es una instantánea reproducible de INSERT con identidades conservadas; rowversion se genera nuevamente en SQL. Los historiales y fotografías XML conservan la información del expediente. La carga no emite correos ni llama a programas externos.

Para ejecutar SQL manualmente debe seleccionarse una base vacía y aplicar Esquema.sql seguido de uno de los archivos de datos en una sola transacción. Se recomienda el instalador para evitar ejecutar en una base incorrecta. No aplicar estos scripts sobre una instalación con registros existentes.
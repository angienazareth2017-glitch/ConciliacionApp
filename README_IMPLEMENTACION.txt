TESORIQ PRO | Persistencia y acceso
==================================

Instalación de prueba
---------------------
1. Instala las dependencias con python -m pip install -r requirements.txt.
2. Crea un proyecto PostgreSQL dedicado o un entorno de prueba en Supabase.
   Obtén la conexión de mantenimiento desde Connect. Para un entorno IPv4
   puedes usar el pooler en modo Session. Usa TLS con verificación del
   certificado y del nombre del servidor según las instrucciones del proveedor.
3. En una terminal privada ejecuta una sola de estas opciones:

      python preparar_bd.py --crear-admin mi_usuario
      python preparar_bd.py --migrar-sqlite /ruta/privada/copia.db

   El comando pide la conexión de mantenimiento sin mostrarla. Si creas el
   administrador, pide también su contraseña. El segundo comando conserva
   usuarios, contraseñas, roles, estado, fechas e historiales de un SQLite que
   incluya un administrador activo. Usa una copia consistente obtenida por el
   mecanismo de backup de SQLite. Si el origen está activo en modo WAL,
   copiar solo el archivo .db puede omitir transacciones recientes.
   Prepara el destino antes de habilitarlo en Streamlit. Guarda el respaldo
   original en privado. El historial JSON antiguo requiere asignación manual
   después de verificar el propietario de cada registro.

Conexión de la aplicación
-------------------------
Con una cuenta de mantenimiento, crea un rol dedicado para la aplicación y
concédele estos permisos. En psql, asigna su contraseña con
\password tesoriq_app para evitar incluirla en archivos o comandos SQL.

    CREATE ROLE tesoriq_app LOGIN;
    GRANT USAGE ON SCHEMA tesoriq TO tesoriq_app;
    GRANT SELECT, INSERT ON tesoriq.usuarios TO tesoriq_app;
    GRANT UPDATE (activo) ON tesoriq.usuarios TO tesoriq_app;
    GRANT SELECT, INSERT, DELETE ON tesoriq.historial TO tesoriq_app;
    GRANT UPDATE (payload) ON tesoriq.historial TO tesoriq_app;
    GRANT USAGE ON ALL SEQUENCES IN SCHEMA tesoriq TO tesoriq_app;

En la configuración de la app de prueba de Streamlit Cloud, añade en Secrets:

    TESORIQ_DATABASE_URL = "conexión PostgreSQL del rol tesoriq_app"

Para probar localmente, puedes usar .streamlit/secrets.toml, excluido por Git,
o la variable de entorno TESORIQ_DATABASE_URL. Ejecuta:

    streamlit run aplicacion.py

La app conserva el acceso por cuenta y los historiales individuales. El
administrador usa el menú Usuarios para crear y desactivar cuentas. Configura
una app de prueba para la rama de cambios y comprueba el inicio de sesión y
la permanencia de los datos tras reiniciar antes de actualizar producción.
El proyecto anterior y su copia SQLite deben quedar disponibles hasta validar
el traspaso. Si el origen sigue en uso, coordina una pausa de escrituras y
toma la copia definitiva durante esa ventana.

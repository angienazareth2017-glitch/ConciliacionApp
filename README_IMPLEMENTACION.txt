TESORIQ - Actualización multiusuario y controles
================================================

Archivos principales
- aplicacion.py: aplicación Streamlit actualizada.
- conciliar.py: motor de conciliación original, sin cambios funcionales.
- tasas.py: módulo de tasas original.
- seguridad.py: autenticación, usuarios y almacenamiento aislado de historiales.
- requirements.txt: dependencias.

Primer inicio
1. Instala dependencias: pip install -r requirements.txt
2. Ejecuta: streamlit run aplicacion.py
3. En la primera apertura aparecerá el formulario para crear el administrador.
4. Inicia sesión con el usuario y contraseña que acabas de configurar.
5. En la barra lateral, el administrador tendrá la opción Usuarios para crear otras cuentas.

Seguridad y aislamiento
- Cada cuenta tiene un ID interno independiente.
- Las contraseñas se guardan como hashes PBKDF2-HMAC-SHA256 con salt individual.
- Cada historial se consulta y actualiza filtrando por el ID del usuario autenticado.
- Los usuarios normales no pueden acceder al módulo de administración.
- La sesión se cierra manualmente o tras 30 minutos de inactividad.
- El usuario no administrador no puede activar/desactivar otras cuentas.

Base de datos
- Por defecto se crea tesoriq.db en el directorio de ejecución.
- Se puede cambiar la ubicación mediante la variable de entorno TESORIQ_DB_PATH.
- La base de datos SQLite es adecuada para pruebas y despliegues de una sola instancia con disco persistente.
- IMPORTANTE: Streamlit Community Cloud no debe tratarse como almacenamiento persistente garantizado. Antes de usar datos reales o lanzar un servicio comercial, configura un volumen persistente o migra el repositorio de datos a PostgreSQL administrado. No publiques ni compartas tesoriq.db.

Cambios funcionales incluidos
- Verificación de saldos iniciales opcional: compara saldo banco vs. saldo mayor cuando se activa y se confirma la captura; muestra diferencia y tolerancia, sin impedir la conciliación.
- Clasificación automática de egresos financieros y comisiones bancarias mediante patrones de descripción.
- La columna Naturaleza Contable del consolidado permite ajustar manualmente la clasificación.
- Refuerzo de contraste de fuentes y controles en CSS.

Notas de migración
- El archivo historial_conciliaciones.json anterior no se importa automáticamente porque no contiene propietario por registro. Importarlo a una cuenta elegida podría exponer datos de otra persona. Si ese historial contiene información de una sola persona, puede realizarse una migración explícita después de confirmar a qué cuenta pertenece.

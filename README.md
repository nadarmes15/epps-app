# Sistema web de consulta de renovación de EPPs

Versión preparada para Render con persistencia real de datos usando **Render Postgres** o cualquier base PostgreSQL compatible.

## Variables recomendadas en Render

- `SECRET_KEY`: clave privada para Flask
- `ADMIN_USERNAME`: usuario del administrador
- `ADMIN_PASSWORD`: contraseña inicial del administrador
- `DATABASE_URL`: cadena de conexión de Render Postgres

## Build y Start Command

Build Command:

```bash
pip install -r requirements.txt
```

Start Command:

```bash
gunicorn app:app
```

## Cómo dejar los datos persistentes

1. Crea una base de datos **PostgreSQL** en Render.
2. Copia la variable `External Database URL` o `Internal Database URL`.
3. En tu Web Service, agrega esa URL como variable `DATABASE_URL`.
4. Redeploy de la aplicación.
5. Ingresa al panel admin y vuelve a cargar tu Excel.

## Nota importante

En Render Free, los archivos locales se borran cuando el servicio se reinicia. Por eso esta versión ya **no depende de guardar el Excel cargado en disco**: lo importa a la base de datos y luego usa la base para las consultas por DNI.

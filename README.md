# Sistema web de consulta de renovación de EPPs - versión sin pandas

Esta versión elimina `pandas` para evitar errores de despliegue en Render.

## Ejecutar localmente

```bash
py -3.13 -m pip install -r requirements.txt
py -3.13 seed_from_excel.py
py -3.13 app.py
```

## Render

- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app`

## Archivos soportados

- `.xlsx`
- `.xlsm`
- `.xls`

## Acceso administrador inicial

- Usuario: `admin`
- Contraseña: `Admin123*`

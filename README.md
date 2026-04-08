# Sistema web de consulta de renovación de EPPs - Transportes Libertad

## Qué incluye
- Portal del trabajador para consultar por DNI.
- Panel de administrador con acceso privado.
- Carga de archivos Excel `.xlsx`, `.xlsm` y `.xls`.
- Detección automática de encabezado, incluso si el archivo tiene filas vacías o títulos arriba.
- Estado visual automático: `VIGENTE`, `POR VENCER` o `VENCIDO`.
- Exportación del reporte actual a Excel.
- Logo integrado de Transportes Libertad.

## Campos esperados
- NRO. DNI
- NOMBRE AUXILIAR
- DESCRIPCION
- CANTIDAD
- FECHA DE MOVIMIENTO
- FECHA DE RENOVACION
- ESTADO

## Instalación
```bash
pip install -r requirements.txt
python seed_from_excel.py
python app.py
```

Luego abre:
```text
http://127.0.0.1:5000
```

## Credenciales iniciales
- Usuario: `admin`
- Contraseña: `Admin123*`

## Recomendación importante
Después del primer ingreso al panel administrador, cambia la contraseña.

## Rutas principales
- `/` → consulta del trabajador
- `/admin/login` → acceso del administrador
- `/admin` → panel administrador

## Cómo publicar en internet
Puedes subir este proyecto a un servidor Windows o Linux con Python.
También puedes publicarlo en Render, Railway o un VPS con Nginx.

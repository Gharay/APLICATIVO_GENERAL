# Gestión Catastral - Aplicativo Local

Este proyecto es una aplicación local de Flask diseñada para gestionar documentos e información catastral. Está pensada para usarse en un equipo local con PostgreSQL.

## Requisitos previos

1. Python 3.12 o superior
2. PostgreSQL instalado
3. Un editor de texto o IDE (opcional)
4. Acceso a PowerShell o terminal de Windows

## Archivos principales

- `app.py` — entrada principal de la aplicación Flask
- `requirements.txt` — dependencias Python
- `backend/` — lógica de rutas y exportación de Excel/Word
- `database/models.py` — modelos SQLAlchemy
- `templates/` — archivos HTML
- `static/` — recursos CSS/JS
- `run.bat` — script para iniciar la app en Windows

## Paso 1: Preparar la base de datos

1. Instala PostgreSQL.
2. Crea una base de datos llamada `sic_catastral`.
3. Anota el usuario, la contraseña y el puerto usados.

Ejemplo con `psql`:

```sql
CREATE DATABASE sic_catastral;
```

## Paso 2: Crear y activar el entorno virtual

Abre PowerShell en la carpeta del proyecto y ejecuta:

```powershell
python -m venv env
.\env\Scripts\Activate.ps1
```

Si el comando de activación no funciona, ejecuta previamente:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
```

## Paso 3: Instalar dependencias

Con el entorno virtual activado, instala las dependencias:

```powershell
pip install -r requirements.txt
```

## Paso 4: Configurar la conexión a PostgreSQL

La aplicación usa variables de entorno para la conexión. Crea un archivo `.env` en la carpeta raíz del proyecto con estos valores:

```env
DB_USER=postgres
DB_PASSWORD=52360
DB_HOST=localhost
DB_PORT=5433
DB_NAME=sic_catastral
```

Ajusta `DB_PASSWORD` y `DB_PORT` según tu instalación local.

> Si no se define `.env`, el proyecto usa valores predeterminados:
> - usuario: `postgres`
> - contraseña: `52360`
> - host: `localhost`
> - puerto: `5433`
> - base de datos: `sic_catastral`

## Paso 5: Ejecutar la aplicación

Con el entorno virtual activo y la base de datos disponible, ejecuta:

```powershell
python app.py
```

También puedes usar el script de Windows:

```powershell
.\run.bat
```

La aplicación abrirá el navegador en:

```
http://localhost:5000
```

## Nota importante

- Si tu PostgreSQL está escuchando en el puerto `5432`, cambia `DB_PORT=5432` en `.env`.
- Si tu usuario o contraseña de PostgreSQL son diferentes, actualiza `DB_USER` y `DB_PASSWORD`.

## Consejos para compartir el proyecto

- No incluyas la carpeta `env/` cuando compartas el proyecto.
- Incluye `requirements.txt`, `migrations/`, `backend/`, `database/`, `templates/` y `static/`.
- El receptor debe crear su propio entorno virtual y su propia base de datos.

## Problemas comunes

- Error de conexión a la base de datos: revisa que PostgreSQL esté ejecutándose y que los datos en `.env` sean correctos.
- Error `ModuleNotFoundError`: asegura que el entorno virtual está activado y que instalaste `requirements.txt`.
- Error de permisos en PowerShell: ejecuta `Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned`.

# Gestión Catastral - Aplicativo Local

Este proyecto es una aplicación local de Flask diseñada para gestionar documentos e información catastral. Está pensada para usarse en un equipo local con PostgreSQL.

## Requisitos previos

1. Python 3.12 o superior
2. PostgreSQL instalado y un gestor de DB como PgAdmin o DBeaver
3. Un editor de texto o IDE (opcional)
4. Acceso a PowerShell o terminal de Windows

## Archivos principales

- `app.py` — entrada principal de la aplicación Flask
- `requirements.txt` — dependencias Python
- `backend/` — lógica de rutas y exportación de Excel/Word
- `database/models.py` — modelos SQLAlchemy
- `templates/` — archivos HTML
- `static/` — recursos CSS/JS

# RECOMENDACIÓN: Usar el aplicativo desde VISUAL STUDIO CODE con el fin de verificar la funcionalidad.

## Descarga toda la carpeta desde el repositorio
## Guardala en el dispositivo local y en una ruta no tan larga
## Abre VSCode y abre la carpeta del APLICATIVO

## Paso 1: Preparar la base de datos

1. Instala PostgreSQL.
2. Crea una base de datos llamada `sic_catastral`.
3. Anota el usuario, la contraseña y el puerto usados por tu postgres.

Ejemplo con `psql`:

```sql
CREATE DATABASE sic_catastral;
```

## Paso 2: Crear y activar el entorno virtual

En terminal:

python -m venv env
.\env\Scripts\Activate.ps1
```

Si el comando de activación no funciona, ejecuta previamente:


Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
```

## Paso 3: Instalar dependencias

Con el entorno virtual activado, instala las dependencias:


pip install -r requirements.txt
```

## Paso 4: Configurar la conexión a PostgreSQL

Ajusta en el archivo *app.py* la linea:

app.config['SQLALCHEMY_DATABASE_URI'] = (
    'postgresql://postgres:XXXX@localhost:5433/sic_catastral'
)

Donde:

XXXX debe ser la constraseña de tu postgres
5433 corresponde al puesto de tu postgres, reemplazar por el correspondiente

## Paso 5: Crear tablas y relaciones de la base de datos, para ello se debe ejecutar el comando en el terminal de VSCode:

flask db upgrade

## Paso 6: Ejecutar la aplicación

Con el entorno virtual activo y la base de datos disponible, ejecuta:


python app.py
```
La aplicación abrirá el navegador en:

```
http://localhost:5000
```


## Consejos para compartir el proyecto

- Por lo general se ejecuta de forma automatica el entorno virtual, en caso de que no, se debe ejecutar de forma manual, con el entorno virtual activo, ahi si se debe ejecutar el app.
- En el terminal podras ver si hubo algun tipo de error al momento de ejecutarlo.
- SE RECOMIENDA IR LLENANDO LA INFORMACIÓN EN LOS FORMULARIOS E IR GUARDANDO SEGUIDO.
- En el tema de colindantes se deben crear y guardar antes de cargar documentos adjuntos.
- Los documentos adjuntos tanto para el predio objeto de estudio y sus colindantes se guardan en la carpeta del aplicativo /static/uploads
- Una vez termines de utilizar el aplicativo y guardar todo, en la terminal de VSCode, usar CTRL + C para dejar de ejecutar el aplicativo.
- No incluyas la carpeta `env/` cuando compartas el proyecto.
- Incluye `requirements.txt`, `migrations/`, `backend/`, `database/`, `templates/` y `static/`.
- El receptor debe crear su propio entorno virtual y su propia base de datos.

## Problemas comunes

- Error de conexión a la base de datos: revisa que PostgreSQL esté ejecutándose y que los datos en `.env` sean correctos. Verifica desde el administrado de tareas y/o los servicios de windows que se este ejecutando el servicio de Postgres
- Error `ModuleNotFoundError`: asegura que el entorno virtual está activado y que instalaste `requirements.txt`.
- Error de permisos en PowerShell: ejecuta `Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned`.

@echo off
REM ================================
REM Crear estructura base del proyecto
REM ================================

REM Archivos raíz
echo. > app.py
echo. > config.py
echo. > requirements.txt
echo. > run.bat

REM ================================
REM Carpeta database
REM ================================
mkdir database
echo. > database\__init__.py
echo. > database\models.py
echo. > database\db_handler.py
echo. > database\sic.db

REM ================================
REM Carpeta backend
REM ================================
mkdir backend
echo. > backend\__init__.py
echo. > backend\routes.py
echo. > backend\export_word.py
echo. > backend\export_excel.py
echo. > backend\file_handler.py
echo. > backend\utils.py

REM ================================
REM Carpeta templates
REM ================================
mkdir templates
echo. > templates\index.html
echo. > templates\form_predio.html
echo. > templates\base.html

REM ================================
REM Carpeta static
REM ================================
mkdir static
mkdir static\css
mkdir static\js
mkdir static\uploads
mkdir static\uploads\IDXXX

echo. > static\css\style.css
echo. > static\js\main.js
echo. > static\js\colindantes.js
echo. > static\js\ajax_handlers.js

REM ================================
REM Carpeta plantillas
REM ================================
mkdir plantillas
echo. > plantillas\informe_template.docx
echo. > plantillas\README.txt

REM ================================
echo.
echo Estructura de carpetas y archivos creada correctamente.
pause

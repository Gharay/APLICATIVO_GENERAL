@echo off
if not exist backups mkdir backups
set FECHA=%date:~-4,4%%date:~-7,2%%date:~0,2%
set PGPASSWORD=52360

REM Backup en formato SQL (texto plano, legible)
pg_dump -U postgres -h localhost -p 5433 sic_catastral > backups\sic_%FECHA%.sql

REM Backup en formato custom comprimido (restaurable con pg_restore o pgAdmin)
pg_dump -U postgres -h localhost -p 5433 -Fc sic_catastral -f backups\sic_%FECHA%.backup

echo Backups completados:
echo   - backups\sic_%FECHA%.sql
echo   - backups\sic_%FECHA%.backup
pause
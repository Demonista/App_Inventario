@echo off
REM Script para compilar app Flask con PyInstaller

REM Cambia al directorio donde está tu app.py
cd /d %~dp0

REM Eliminar carpetas anteriores de compilación
rmdir /s /q build
rmdir /s /q dist
del /q app.spec

REM Compilar la app
pyinstaller --onefile --noconsole --add-data "templates;templates" --add-data "static;static" --add-data "uploads;uploads" app.py

REM Esperar a que termine y abrir carpeta con EXE
if exist dist\app.exe (
    echo Compilación completada con éxito.
    start dist
) else (
    echo ERROR: No se encontró el ejecutable. Revisa errores anteriores.
)

pause

@echo off
REM ==========================================================
REM  Excel Merger - lanzador
REM ==========================================================
REM  Doble clic para ejecutar el JAR sin abrir la terminal.
REM  El .bat se posiciona en su propia carpeta, asi las rutas
REM  relativas del config.properties (input/, output/...) funcionan
REM  aunque lo lances desde un acceso directo o desde otra carpeta.
REM ==========================================================

setlocal

REM Situarse en la carpeta donde esta este .bat
cd /d "%~dp0"

REM Buscar el JAR ejecutable (permite cambios de version sin editar el .bat)
set "JAR="
for %%f in ("target\excel-merger-*-jar-with-dependencies.jar") do set "JAR=%%f"

if not defined JAR (
    echo.
    echo [ERROR] No se encuentra el JAR en la carpeta 'target'.
    echo         Ejecuta antes: mvn clean package
    echo.
    pause
    exit /b 1
)

echo.
echo Ejecutando: %JAR%
echo.

java -jar "%JAR%"
set EXITCODE=%ERRORLEVEL%

echo.
if %EXITCODE% NEQ 0 (
    echo [ERROR] El programa ha terminado con codigo %EXITCODE%.
) else (
    echo [OK] Proceso finalizado correctamente.
)

echo.
pause
endlocal
exit /b %EXITCODE%

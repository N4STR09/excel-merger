@echo off
REM ==========================================================
REM  Excel Merger - lanzador
REM ==========================================================
REM  Doble clic para ejecutar el JAR sin abrir la terminal.
REM  El .bat se posiciona en su propia carpeta, asi las rutas
REM  relativas del config.properties (input/, output/...) funcionan
REM  aunque lo lances desde un acceso directo o desde otra carpeta.
REM
REM  Configs por entorno:
REM    run.bat                      -> usa config.properties (default del JAR)
REM    run.bat contabilidad         -> usa config-contabilidad.properties
REM    run.bat mi-cfg.properties    -> pasa la ruta/nombre tal cual al JAR
REM
REM  Si el fichero resuelto no existe en la carpeta del .bat, el
REM  lanzador aborta con codigo de salida 2 antes de arrancar Java.
REM ==========================================================

setlocal

REM Situarse en la carpeta donde esta este .bat
cd /d "%~dp0"

REM --- Resolver el fichero de configuracion a partir del primer argumento ---
set "CONFIG_ARG="
if "%~1"=="" goto afterConfig

call :resolveConfig "%~1"

if not exist "%CONFIG_ARG%" (
    echo.
    echo [ERROR] No se encuentra el fichero de configuracion: %CONFIG_ARG%
    echo         Buscado en: %CD%\%CONFIG_ARG%
    echo         Si es un entorno nuevo, crealo copiando 'config.properties'
    echo         como plantilla. Por ejemplo:
    echo           copy config.properties %CONFIG_ARG%
    echo.
    pause
    exit /b 2
)

:afterConfig

REM Buscar el JAR ejecutable (permite cambios de version sin editar el .bat)
set "JAR="
for %%f in ("target\excel-merger-*-jar-with-dependencies.jar") do set "JAR=%%f"

if not defined JAR (
    echo.
    echo [ERROR] No se encuentra el JAR en la carpeta 'target'.
    if exist "mvnw.cmd" (
        echo         Ejecuta antes: mvnw.cmd clean package
    ) else (
        echo         Ejecuta antes: mvn clean package
    )
    echo.
    pause
    exit /b 1
)

echo.
if defined CONFIG_ARG (
    echo Ejecutando: %JAR% con config '%CONFIG_ARG%'
) else (
    echo Ejecutando: %JAR%
)
echo.

if defined CONFIG_ARG (
    java -jar "%JAR%" "%CONFIG_ARG%"
) else (
    java -jar "%JAR%"
)
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

REM ==========================================================
REM Subrutina: resuelve el argumento a una ruta de .properties
REM   %~1 = argumento del usuario (p. ej. 'contabilidad' o 'cfg.properties')
REM   Deja el resultado en la variable CONFIG_ARG:
REM     - Si el argumento ya termina en '.properties', se usa tal cual.
REM     - Si no, se construye 'config-<arg>.properties'.
REM ==========================================================
:resolveConfig
set "ARG=%~1"
if /I "%ARG:~-11%"==".properties" (
    set "CONFIG_ARG=%ARG%"
) else (
    set "CONFIG_ARG=config-%ARG%.properties"
)
goto :eof

@echo off
REM ==========================================================
REM  Excel Merger - lanzador (wrapper fino)
REM ==========================================================
REM  Reenvia los argumentos al JAR sin interpretarlos:
REM    run.bat              -> interfaz web local (abre navegador)
REM    run.bat --cli        -> menu interactivo de terminal
REM    run.bat --merge      -> fusion directa (exit codes 0-4)
REM    run.bat --compare    -> comprobador directo (exit 0-4)
REM    run.bat --help       -> ayuda
REM
REM  El proceso se lanza con CWD en esta carpeta, de modo que el
REM  config por defecto sigue siendo 'config.properties' de aqui.
REM
REM  El pause solo ocurre al arrancar SIN argumentos (doble clic),
REM  para que con argumentos el exit code sea util en scripts.
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
    if exist "mvnw.cmd" (
        echo         Ejecuta antes: mvnw.cmd clean package
    ) else (
        echo         Ejecuta antes: mvn clean package
    )
    echo.
    if "%~1"=="" pause
    endlocal
    exit /b 1
)

java -jar "%JAR%" %*
set EXITCODE=%ERRORLEVEL%

if "%~1"=="" (
    echo.
    if %EXITCODE% EQU 0 (
        echo [OK] Proceso finalizado correctamente.
    ) else (
        echo [ERROR] El programa ha terminado con codigo %EXITCODE%.
    )
    echo.
    pause
)

endlocal & exit /b %EXITCODE%

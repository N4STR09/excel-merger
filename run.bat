@echo off
REM ==========================================================
REM  Excel Merger v3.1.0 - lanzador
REM ==========================================================
REM  Doble clic para ejecutar el JAR. Al arrancar, aparece un
REM  menu interactivo con tres opciones:
REM    1) Fusion de Excel
REM    2) Comprobador de discrepancias contra CSV (v3.1.0)
REM    3) Salir
REM  El config a usar es siempre 'config.properties' del
REM  directorio del .bat.
REM
REM  v3.0.0 BREAKING (sigue vigente): ya NO se aceptan argumentos.
REM  Las antiguas formas:
REM    run.bat                      (default config) -> SIGUE FUNCIONANDO
REM    run.bat contabilidad         -> YA NO. Renombra el config
REM                                    'config-contabilidad.properties'
REM                                    a 'config.properties' o copialo.
REM    run.bat mi-cfg.properties    -> YA NO. Mismo workaround.
REM
REM  Para entornos multiples mantener varios 'config-<entorno>.properties'
REM  y copiar el deseado a 'config.properties' antes de lanzar.
REM ==========================================================

setlocal

REM Situarse en la carpeta donde esta este .bat
cd /d "%~dp0"

REM Aviso si el usuario ha pasado algun argumento (v2.7.1 lo hacia)
if not "%~1"=="" (
    echo.
    echo [AVISO] Desde v3.0.0 no se aceptan argumentos en linea de comandos.
    echo         El argumento '%~1' sera ignorado por el JAR.
    echo         Si necesitas un config alternativo, copialo como
    echo         'config.properties' antes de lanzar.
    echo.
)

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

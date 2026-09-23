@echo off
REM ==========================================================
REM  Excel Merger - empaquetado portable (jpackage)
REM ==========================================================
REM  Genera una imagen de aplicacion autocontenida (JDK embebido)
REM  y la comprime en un .zip portatil. NUNCA crea instalador
REM  (decision v4.0.0: solo app-image + zip).
REM
REM  Requisitos: JDK con jpackage (14+) y mvn en el PATH.
REM
REM  IMPORTANTE: lo gestiona el propio script, pero tenlo en cuenta:
REM  los ficheros que genera jpackage en target\dist salen de SOLO
REM  LECTURA y una copia abierta (ExcelMerger.exe, o el .jar lanzado
REM  con doble clic = javaw) bloquea logs. Cualquiera de los dos hace
REM  fallar 'mvn clean' a mitad y deja target\ sin runtime ni zip:
REM  el .exe sobrevive pero ya no arranca.
REM
REM  Salida:
REM    target\dist\ExcelMerger\ExcelMerger.exe  (imagen autocontenida)
REM    target\excel-merger-win64-portable.zip   (portatil)
REM ==========================================================

setlocal
cd /d "%~dp0"

REM Copias abiertas = 'mvn clean' fallaria a mitad (exe, o .jar doble-clickeado = javaw).
taskkill /F /IM ExcelMerger.exe >nul 2>nul
if not errorlevel 1 echo [AVISO] Cerrada una instancia de ExcelMerger que estaba en ejecucion.
powershell -NoProfile -NonInteractive -Command "Get-CimInstance Win32_Process -Filter \"Name='javaw.exe'\" | Where-Object CommandLine -like '*jar-with-dependencies*' | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }" >nul 2>nul
REM Igual con un 'java -jar' de terminal (java.exe, no javaw): bloquearia
REM 'mvn clean' igual que el doble clic. No toca el JVM de Maven (otra CommandLine).
powershell -NoProfile -NonInteractive -Command "Get-CimInstance Win32_Process -Filter \"Name='java.exe'\" | Where-Object CommandLine -like '*jar-with-dependencies*' | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }" >nul 2>nul

REM La imagen jpackage es de solo lectura: sin quitar el atributo,
REM 'mvn clean' muere a mitad (borra runtime y zip y se para en el .exe).
attrib -R -S -H "target\dist\*.*" /S /D >nul 2>nul

REM Perfil de la instalacion actual (token de web.properties = URL
REM 'estable', y ajustes de web-settings.properties). Vive fuera de
REM target porque el 'mvn clean' de abajo borra target entero.
set "PROFILE_BACKUP=%TEMP%\excelmerger-pkg-profile"
if exist "target\dist\ExcelMerger\web.properties" (
    mkdir "%PROFILE_BACKUP%" 2>nul
    copy /y "target\dist\ExcelMerger\web.properties" "%PROFILE_BACKUP%\" >nul
)
if exist "target\dist\ExcelMerger\web-settings.properties" (
    mkdir "%PROFILE_BACKUP%" 2>nul
    copy /y "target\dist\ExcelMerger\web-settings.properties" "%PROFILE_BACKUP%\" >nul
)

REM jpackage: del PATH o del JDK indicado por JAVA_HOME
set "JP=jpackage"
where jpackage >nul 2>nul
if errorlevel 1 (
    if defined JAVA_HOME (
        set "JP=%JAVA_HOME%\bin\jpackage.exe"
    ) else (
        echo [ERROR] jpackage no esta en el PATH ni existe JAVA_HOME.
        if not defined EXCELMERGER_NO_PAUSE pause
        exit /b 1
    )
)

echo [1/4] Build de Maven (clean package)...
call mvn -q clean package
if errorlevel 1 (
    echo [ERROR] El build de Maven ha fallado.
    echo         Suele ser una instancia abierta: cierra ExcelMerger o el
    echo         .jar si se abrio con doble clic, y reintenta.
    if not defined EXCELMERGER_NO_PAUSE pause
    exit /b 1
)

set "JAR="
for %%f in ("target\excel-merger-*-jar-with-dependencies.jar") do set "JAR=%%f"
if not defined JAR (
    echo [ERROR] No se encuentra el JAR con dependencias en 'target'.
    if not defined EXCELMERGER_NO_PAUSE pause
    exit /b 1
)
for %%f in ("%JAR%") do set "JARNAME=%%~nxf"

REM Entrada limpia con solo el JAR: si --input fuera 'target', la imagen
REM escribiria dentro de la misma carpeta de la que lee (target\dist).
rmdir /s /q "target\stage" 2>nul
mkdir "target\stage"
copy /y "%JAR%" "target\stage\" >nul

rmdir /s /q "target\dist" 2>nul

echo [2/4] jpackage --type app-image...
"%JP%" --type app-image --name "ExcelMerger" --input "target\stage" --main-jar "%JARNAME%" --main-class "com.excelmerger.Main" --dest "target\dist" --win-console
if errorlevel 1 (
    echo [ERROR] jpackage ha fallado.
    if not defined EXCELMERGER_NO_PAUSE pause
    exit /b 1
)

REM Carpetas de trabajo DENTRO de la imagen: el usuario las tiene ya
REM extraidas, antes de arrancar nada. El LEEME.txt ademas garantiza
REM que Compress-Archive no descarte carpetas vacias (bug conocido).
echo [3/4] Carpetas input y output con su LEEME...
mkdir "target\dist\ExcelMerger\input" 2>nul
mkdir "target\dist\ExcelMerger\output" 2>nul
> "target\dist\ExcelMerger\input\LEEME.txt" echo Excel Merger - carpeta de entrada
>> "target\dist\ExcelMerger\input\LEEME.txt" echo --------------------------------------------------
>> "target\dist\ExcelMerger\input\LEEME.txt" echo Pon aqui tus ficheros Excel de entrada, 2 o 3 ficheros .xlsx.
>> "target\dist\ExcelMerger\input\LEEME.txt" echo No hace falta renombrarlos: se identifican por su contenido.
>> "target\dist\ExcelMerger\input\LEEME.txt" echo La pagina web te muestra esta misma ruta.
>> "target\dist\ExcelMerger\input\LEEME.txt" echo Este archivo no se tiene en cuenta; puedes borrarlo.
> "target\dist\ExcelMerger\output\LEEME.txt" echo Excel Merger - carpeta de salida
>> "target\dist\ExcelMerger\output\LEEME.txt" echo --------------------------------------------------
>> "target\dist\ExcelMerger\output\LEEME.txt" echo Aqui se guarda el resultado de la fusion: resultado_fusion.xlsx
>> "target\dist\ExcelMerger\output\LEEME.txt" echo y las discrepancias de la opcion 2. La pagina web te muestra
>> "target\dist\ExcelMerger\output\LEEME.txt" echo esta misma ruta.
>> "target\dist\ExcelMerger\output\LEEME.txt" echo Este archivo no se tiene en cuenta; puedes borrarlo.
if errorlevel 1 (
    echo [ERROR] No se pudieron crear las carpetas de trabajo.
    if not defined EXCELMERGER_NO_PAUSE pause
    exit /b 1
)

REM Restaura el perfil conservado (token y ajustes) en la imagen nueva,
REM y lo hace antes del zip para que el portatil lo incluya tambien.
if exist "%PROFILE_BACKUP%\web.properties" (
    copy /y "%PROFILE_BACKUP%\web.properties" "target\dist\ExcelMerger\" >nul
)
if exist "%PROFILE_BACKUP%\web-settings.properties" (
    copy /y "%PROFILE_BACKUP%\web-settings.properties" "target\dist\ExcelMerger\" >nul
)
rmdir /s /q "%PROFILE_BACKUP%" 2>nul

echo [4/4] Zip portatil...
powershell -NoProfile -Command "Compress-Archive -Force -Path 'target\dist\ExcelMerger' -DestinationPath 'target\excel-merger-win64-portable.zip'"
if errorlevel 1 (
    echo [ERROR] No se pudo crear el zip.
    if not defined EXCELMERGER_NO_PAUSE pause
    exit /b 1
)

echo.
echo [OK] Imagen: target\dist\ExcelMerger\ExcelMerger.exe
echo [OK] Zip:    target\excel-merger-win64-portable.zip
echo.
endlocal
exit /b 0

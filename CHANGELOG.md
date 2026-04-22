# Changelog

## [1.4.0] — Sesión D: nuevas funcionalidades

### Añadido

- **`--dry-run`** (`Main`, `ExcelMerger`): nueva flag de CLI que ejecuta el pipeline completo (validación de config, detección de perfiles, construcción de hojas Resultado/Equipos/derivadas en memoria, detección de apps sin mapeo, cabeceras no encontradas, etc.) pero **NO escribe el Excel de salida y NO mueve el anterior a `history/`**. Útil para validar la configuración antes de un cierre mensual. El chequeo de lock `~$` sobre el output sí se mantiene para avisar cuanto antes. Nuevo constructor sobrecargado `ExcelMerger(ConfigLoader, RunReport, boolean dryRun)`; el constructor de 2 args delega con `dryRun=false` (retrocompatible, los 125 tests existentes siguen verdes sin tocar).
- **Hoja `_Avisos` en el Excel resultado** (clase nueva `AvisosSheetBuilder`): opt-in con `report.inExcel=true`. Vuelca todos los warnings acumulados en el `RunReport` (apps sin mapeo, cabeceras no encontradas, perfiles sin match, etc.) a una hoja extra del Excel, con cabecera `Categoria | Mensaje`. Oculta por defecto. La hoja no se crea si no hay warnings (evita hojas vacías ruidosas) ni si `report.inExcel=false`. Colisión con hoja existente: se omite y se registra un warning `HOJA`. Se construye igual en dry-run (en memoria), aunque luego no se escriba a disco.
- **Configs por entorno** (`run.bat`): el lanzador acepta un alias opcional como primer argumento y lo resuelve a `config-<alias>.properties`. Por ejemplo, `run.bat contabilidad` carga `config-contabilidad.properties`. Si el argumento ya termina en `.properties` (case-insensitive), se pasa tal cual al JAR. Si el fichero resuelto no existe en la carpeta del `.bat`, el lanzador aborta con `exit /b 2` y mensaje claro **antes** de arrancar la JVM. Sin argumento: comportamiento actual (el JAR usa `config.properties` por defecto).

### Configuración nueva

Solo se añaden claves nuevas; la sintaxis del `config.properties` y las claves existentes no cambian:

- `report.inExcel=false` (default; opt-in explícito).
- `report.sheetName=_Avisos` (nombre de la hoja de avisos).
- `report.hidden=true` (la hoja se crea oculta).

### Tests

- `ExcelMergerIntegrationTest` (+7): 5 de `--dry-run` (no crea fichero, no mueve a history, detecta apps sin mapeo igualmente, falla si output lockeado, backward-compat del constructor de 2 args) y 2 de `_Avisos` (aparece con `inExcel=true` y contiene la app `ZZ`, no aparece con el default `false`).
- `AvisosSheetBuilderTest` (+8, clase nueva): no-op si `inExcel=false`, no-op si no hay warnings, volcado ordenado con cabecera `Categoria|Mensaje`, oculta por defecto, visible con `hidden=false`, nombre personalizable vía `report.sheetName`, colisión con hoja existente + warning `HOJA`, registro en `RunReport.sheets()`.
- `ConfigLoaderTest` (+2): carga `config-<env>.properties` por ruta externa; `ConfigurationException` si el fichero no existe. Blindan el contrato que usa `run.bat`.
- Total: 125 → **142 tests** (+17). Cobertura ≥70% mantenida.

### Sin cambios

- Exit codes del JAR (0-4).
- Case-sensitive de `{col:X}` en `MesSheetBuilder` (decisión heredada de [1.2.1], sin modificar).
- Sintaxis del `config.properties` (solo claves añadidas).
- CLI del JAR: no se añaden flags nuevas más allá de `--dry-run`. La resolución de entornos vive en `run.bat`, el JAR sigue recibiendo una ruta `.properties` como siempre.

### `.gitignore`

- Añadidos patrones `input/*.xlsx`, `input/*.xls`, `output/**/*.xlsx`, `output/**/*.xls` para evitar subir al repositorio Excel con datos reales (peticiones, horas, matrículas...). Los fixtures sintéticos de `src/test/resources/fixtures/` siguen versionándose (necesarios para los tests).

## [1.3.1] — Rename MES → Resultado y condición Funcion en SUMIFS

### Cambiado
- **Hoja `MES` renombrada a `Resultado`** en `config.properties` (principal y fallback). Los prefijos de propiedades internas siguen siendo `mes.*` (convención del código). El nombre visible en el Excel resultado cambia a `Resultado`.
- **Nueva columna `Funcion` en el perfil `Cierre`**, posicionada entre `Matricula` y `Account`. Añadida tanto en la lista de `profile.Cierre.detect.headers` como en los fixtures `cierre.xlsx` (cabecera en fila 2, valor en cada imputación).
- **SUMIFS de la columna `Jira` ampliado** con el par `Funcion:Funcion`. Ahora la suma cruza tres criterios contra Cierre:
  - `Component Name` ⟷ `Peticion`
  - `Matricula` ⟷ `Recurso`
  - `Funcion` ⟷ `Funcion` *(nuevo)*

### Fixtures
- `gen_fixtures.py`: añadida columna `Funcion` a `CIERRE_HEADERS` y a las 16 filas de imputación. 15 filas tienen `Funcion=Dev`; la fila `PROJ-3` (P-001/M-1001) tiene `Funcion=Sup` deliberadamente, para que el SUMIFS la filtre y los tests puedan verificar que la restricción funciona.

### Tests
- `ExcelMergerIntegrationTest`: 7 apariciones de `"MES"` actualizadas a `"Resultado"`.
- `test-config.properties`: `mes.sheetName=Resultado` y `mes.col.4.match` ampliado con `Funcion:Funcion`.
- **2 tests nuevos** (`ExcelMergerIntegrationTest`):
  - `mesColJiraFormulaSumIfsIncluyeCondicionFuncion`: verifica que la fórmula SUMIFS tiene 4 referencias a `Cierre` (1 `sum_range` + 3 criterios).
  - `mesColJiraSumIfsFiltraPorFuncion`: evalúa numéricamente la fórmula para P-001/M-1001 y comprueba que el resultado es **5** (3+2 de las filas con `Funcion=Dev`), no 9 (que incluiría la fila `Sup`).
- `mesColJiraContieneFormulaSumIfsConReferenciaACierre`: actualizada la aserción `$O:$O` → `$P:$P` porque la columna `Hours` de `cierre.xlsx` se desplaza una posición al insertar `Funcion`.
- Total: 123 → **125 tests**.

### Sin cambios
- Resto de referencias a "MES" en tests aislados (`MesSheetBuilderTest`, `RunReportTest`) que construyen sus propios configs autocontenidos: intencionales, validan el builder de forma abstracta.
- Defaults `config.get("mes.sheetName", "MES")` en `ConfigValidator` y `MesSheetBuilder`: en runtime real la clave siempre está definida, el default no se alcanza. Se dejan como están.

## [1.3.0] — Refactor arquitectónico

### Objetivo
Mejorar la mantenibilidad del código sin cambiar el comportamiento externo ni el contenido del Excel resultado. Los tests existentes siguen siendo la red de seguridad; ninguna fase modifica su semántica salvo el tipo de excepción asertado en los 7 tests de Fase 2.

### Paquetes nuevos y clases movidas

**`com.excelmerger.util`** — utilidades POI compartidas, antes duplicadas en cada builder.
- `PoiUtils`: `cellAsString`, `copyCellValue`, `isBlank`, `findColumnIndex`, `columnLetter`, `detectHeaderRow`, `quoteSheetName`, `countColumns`.
- `StyleFactory`: `header(Workbook)`, `title(Workbook)`.

**`com.excelmerger.exception`** — jerarquía de excepciones tipadas (todas `RuntimeException`).
- `ExcelMergerException` (base abstracta).
- `ConfigurationException` → exit 2. Config ausente, ilegible o con errores en `strictValidation=true`.
- `InputValidationException` → exit 3 (nuevo). Directorio de entrada o ficheros Excel mal; locks `~$` sobre inputs.
- `OutputException` → exit 4 (nuevo). Lock `~$` sobre output, `overwrite=false` con output existente, fallo de escritura o backup.
- `MergeException` → exit 1. Fallo genérico durante la fusión.

**`com.excelmerger.sheet.column`** — strategy pattern para las columnas MES.
- `MesColumnStrategy` (interfaz) + `AbstractMesColumnStrategy` (base con `disabled`, `greenIfPositive`, `name` y template method para `writeCell`).
- `CopyColumnStrategy`, `SumIfsColumnStrategy`, `FormulaColumnStrategy`, `EmptyColumnStrategy`.
- `MesColumnStrategyFactory.fromConfig(...)`: devuelve la estrategia correcta o `EmptyColumnStrategy` si el tipo es inválido o la configuración está incompleta.
- `VlookupLink`: extraída de `MesSheetBuilder` a su propia clase pública (antes inner class).

**`com.excelmerger.io`** — colaboradores de E/S extraídos de `ExcelMerger`.
- `InputFileDetector`: `findExcelFiles(String)`, `validateExcelFiles(List, boolean, RunReport)`, `hasExcelExtension(String)`.
- `FileLockDetector`: `assertNotLocked(File)`, `openForRead(File)`, `looksLikeLocked(Throwable)`.
- `OutputManager`: `assertOutputWritable(String)`, `prepareOutputFile(String, boolean, boolean)`, `backupOutput(Path)`, `writeResult(Workbook, String)`.
- `SheetCopier`: `copySheet(Sheet, Sheet, Workbook)`, `copyRow(...)`, `countColumns(Sheet)` (delega en `PoiUtils` para evitar duplicación nueva).

### Fases del refactor

**Fase 1 — Utilidades comunes.** `createHeaderStyle` (3 copias idénticas en `LookupSheetBuilder`, `DerivedSheetBuilder`, `MesSheetBuilder`) y `copyCellValue` (2 copias en `ExcelMerger` y `MesSheetBuilder`) eliminadas. `findColumnIndex`, `columnLetter`, `detectHeaderRow`, `quote`, `isBlank`, `cellAsString`, `countColumns` centralizadas en `PoiUtils`. `DerivedSheetBuilder` conserva su `cellToString` propio porque aplica `trim()` + fechas, semántica distinta de `PoiUtils.cellAsString`. `FileProfileResolver.cellToString` también se queda en su sitio por la misma razón. Sin cambios de comportamiento.

**Fase 2 — Excepciones tipadas.** `ExcelMerger.merge()` deja de declarar `throws IOException` y lanza `InputValidationException` / `OutputException` / `MergeException` según el punto de fallo. `ConfigLoader` deja de declarar `throws IOException` y lanza `ConfigurationException`. `Main.java` añade catches tipados que mapean a los nuevos exit codes (2, 3, 4). `--help` actualizado. Los 7 tests que asertaban `IOException` actualizados al tipo correspondiente:
- `ConfigLoaderTest.classpathFallbackLanzaSiNoEstaNiEnDiscoNiEnClasspath` → `ConfigurationException`.
- `ExcelMergerIntegrationTest.lanzaSiInputDirectoryNoExiste` → `InputValidationException`.
- `ExcelMergerIntegrationTest.lanzaSiInputDirectorioSinDosExcel` → `InputValidationException`.
- `ExcelMergerIntegrationTest.strictTwoFilesTrueAbortaConTresExcel` → `InputValidationException`.
- `ExcelMergerIntegrationTest.overwriteFalseAbortaSiOutputYaExiste` → `OutputException`.
- `ExcelMergerIntegrationTest.lockDeInputAbortaConMensajeClaro` → `InputValidationException`.
- `ExcelMergerIntegrationTest.lockDeOutputAbortaAntesDeTocarInputs` → `OutputException`.

**Fase 3 — Strategy pattern en MES.** La clase interna `MesColumn` de `MesSheetBuilder` (con su `switch` en `writeCell` y su `switch` en `preValidate`) desaparece, reemplazada por 4 strategies + factory en el nuevo paquete. `MesSheetBuilder` queda como orquestador: lee config → crea lista de strategies → orquesta `preValidate` → itera filas llamando a `writeCell`. `MesSheetBuilder` pasa de **696 → 325 líneas** (−53%). La detección de VLOOKUP sin mapeo sigue en el orquestador y usa `MesColumnStrategy.formulaTemplate()` (Optional) para acceder a la plantilla solo en las columnas FORMULA.

**Fase 4 — Colaboradores de `ExcelMerger`.** Los 4 colaboradores del paquete `com.excelmerger.io` se instancian en el constructor de `ExcelMerger`. `merge()` pasa a ser una secuencia de llamadas: `inputDetector.findExcelFiles` → `inputDetector.validateExcelFiles` → `lockDetector.assertNotLocked` (×2) → `outputManager.assertOutputWritable` → `outputManager.prepareOutputFile` → `sheetCopier.copySheet` → builders (lookup, MES, derived) → `outputManager.writeResult`. `ExcelMerger` pasa de **570 → 264 líneas** (−54%). `mergeSheetsSeparate` y `mergeAppendRows` se mantienen como métodos privados del orquestador (no se extrae strategy para el modo por quedar bajo el presupuesto de riesgo de esta sesión).

**Fase 5 — Limpieza final.** `VlookupLink` extraída a su propia clase en `com.excelmerger.sheet.column`. Versión subida a 1.3.0 en `Main.APP_VERSION`. Sin `@SuppressWarnings` residuales ni TODOs abiertos.

### Decisiones conservadas

- **`MesSheetBuilder` sigue resolviendo `{col:X}` de forma case-sensitive**. Mantenido como feature consciente (permite distinguir columnas `Real` y `REAL` cuando ambas están declaradas), no uniformado al case-insensitive del resto del proyecto. Los dos tests que lo blindan (`placeholderColDistingueEntreRealYREAL` y `placeholderColNoMatcheaPorCasePermisivo`) siguen pasando sin tocar. Si se quiere unificar en un futuro release, el punto único de cambio es el `Map<String,Integer> mesColIndexByName` en `MesSheetBuilder.build()`: cambiar `LinkedHashMap` por `TreeMap(String.CASE_INSENSITIVE_ORDER)` y actualizar los dos tests.
- **`mergeSheetsSeparate` y `mergeAppendRows` permanecen inline** en `ExcelMerger` como métodos privados, no se convierten en strategies de modo de merge. La ganancia era limitada y el riesgo de la sesión subía; aplazable a una futura v1.4 si se añade un tercer modo.

### Código eliminado
- `MesSheetBuilder.MesColumn` (inner class, ~275 líneas): sustituida por las 4 strategies.
- `MesSheetBuilder.VlookupLink` (inner class): movida al paquete strategy.
- `ExcelMerger.findExcelFiles`, `validateExcelFiles`, `hasExcelExtension`, `assertNotLocked`, `assertOutputWritable`, `openForRead`, `writeResult`, `prepareOutputFile`, `backupOutput`, `looksLikeLocked`, `copySheet`, `copyRow`, `copyCellValue`, `countColumns`: movidos al paquete `io` o `util`.
- Métodos `createHeaderStyle` y `createTitleStyle` duplicados en 3 builders: reducidos a una única `StyleFactory`.

### Sin cambios
- Firma y sintaxis del `config.properties` del usuario.
- Firma de la CLI (flags `--help`, `--version`, códigos de salida 0/1/2 preexistentes siguen funcionando).
- Contenido del Excel resultado (hojas, orden, valores, fórmulas generadas).

## [1.2.1] - Red de seguridad de tests

### Añadido
- **Suite de tests con JUnit 5 + AssertJ** (`src/test/java/com/excelmerger/`):
  - `ConfigLoaderTest` (10): carga desde disco y fallback a classpath, defaults, coerciones `boolean`/`int`, trim, lectura UTF-8.
  - `RunReportTest` (15): acumulación de hojas/warnings en orden, copia defensiva de `registerLookupKeys`, inmutabilidad de `warnings()`/`getLookupKeys()`, contenido de `formatSummary()`.
  - `ConfigValidatorTest` (28): detecta config incompletos, tipos MES inválidos, placeholders `{col:X}` que no casan, SUMIFS con hojas/cabeceras desconocidas, lookups sin data, derivadas mal definidas. Verifica acumulación de varios errores e idempotencia de `validate()`.
  - `FileProfileResolverTest` (15): `safeSheetName` (null, caracteres prohibidos, truncado a 31), `hasProfiles`, matching case-insensitive y por subcadena, `headerRow` configurable, criterio `cellValue`, first-match wins.
  - `LookupSheetBuilderTest` (12): construcción de hoja con cabeceras/datos, separador configurable, `hidden`, entradas malformadas/claves vacías, registro de claves en `RunReport`, colisión con hoja existente.
  - `MesSheetBuilderTest` (14): resolución de placeholders `{col:X}` y `{colLetter:X}`, distinción `{col:Real}` vs `{col:REAL}` cuando ambas columnas están definidas, `=` inicial se elimina, skip de filas con ancla vacía, warnings por columna ancla/hoja origen inexistentes.
  - `DerivedSheetBuilderTest` (12): hojas tipo `FORMULAS` (texto/número/fórmula), `AGGREGATION` con SUM (+ fila TOTAL) / AVG / MIN, warnings por tipo desconocido, sourceSheet inexistente y colisión.
  - `ExcelMergerIntegrationTest` (17): merge end-to-end sobre fixtures reales, verifica 4 hojas de salida (Cierre, Extraccion, Equipos, MES), conteo de filas, fórmulas SUMIFS con referencia a Cierre, `{col:Jira}*1.2` resuelto a `D2*1.2`, detección de apps sin mapeo (ZZ), validación de entradas (dir inexistente, <2 archivos, `strictTwoFiles`), backup a `history/`, `overwrite=false`, locks `~$` en input y output, modo `APPEND_ROWS`.
  - Total: 123 tests.
- **Fixtures Excel versionados** en `src/test/resources/fixtures/`: `extraccion.xlsx` (14 peticiones válidas + 1 fila con `Peticion` vacía para probar el skip) y `cierre.xlsx` (1 fila de metadatos + 16 imputaciones con totales conocidos para los SUMIFS). Script `gen_fixtures.py` incluido en la raíz del proyecto para regenerarlos.
- **`TestFixtures.java`**: utilidad compartida con `copyFixturesTo(Path)`, `renderTestConfig(...)` (sustituye los placeholders `${TEST_INPUT_DIR}` y `${TEST_OUTPUT_FILE}` en `test-config.properties`), `buildRealisticConfig(Path)` para el happy-path, y `configFromProperties(...)` / `configFromPairs(...)` para tests unitarios.
- **Configuración Maven para tests**: JUnit 5.10.2 vía BOM, AssertJ 3.25.3, Surefire 3.2.5 y JaCoCo 0.8.14. Cobertura mínima del **70% INSTRUCTION a nivel de bundle**, excluyendo `com/excelmerger/Main.class` (contiene `System.exit()` y lógica de CLI difícil de cubrir unitariamente sin un runner externo).

### Descubierto (no arreglado en esta sesión — revisar en Sesión B)
- **`MesSheetBuilder` resuelve `{col:X}` de forma case-sensitive.** Si el config declara una columna `Real` y una fórmula referencia `{col:REAL}`, la columna consumidora se deshabilita silenciosamente (celda en blanco + warning `FORMULA`). El código lo documenta como intencional (permite distinguir `Real` vs `REAL` cuando ambas están definidas), pero es incoherente con el resto del proyecto, donde el matching de cabeceras en `FileProfileResolver` y `MesSheetBuilder.findColumnIndex` es case-insensitive. Documentado por `MesSheetBuilderTest.placeholderColNoMatcheaPorCasePermisivo` (fija el comportamiento actual) y `MesSheetBuilderTest.placeholderColDistingueEntreRealYREAL` (verifica que ambas se distinguen cuando están declaradas simultáneamente). Decisión pendiente: ¿feature (mantener) o bug (uniformar a case-insensitive)?

### Corregido tras primera ejecución en Windows
- `TestFixtures.configFromProperties` escribía el `.properties` serializando a mano (`key=value\n`). En Windows, las rutas con backslashes (`C:\Users\...`) se leían como `C:UsersERLANT~1...` porque `Properties.load` interpreta `\U`, `\n`, `\A`, etc. como secuencias de escape. Cambiado a `Properties.store(Writer, null)` que escapa correctamente. Afectaba a 8 tests de `ExcelMergerIntegrationTest`.
- `ExcelMergerIntegrationTest.mesColJiraContieneFormulaSumIfsConReferenciaACierre` verificaba que la fórmula SUMIFS contuviese el literal `"Hours"`, pero POI registra las referencias por letra de columna (`$O:$O`), no por nombre de cabecera. Asserción ajustada.
- `MesSheetBuilderTest.hojaMesYaExistenteSeOmite` esperaba `getLastRowNum()==0` para una hoja recién creada sin filas. En POI 5 el valor correcto es `-1`; cambiado a `getPhysicalNumberOfRows()==0`, que es más semántico y no depende de la convención de POI.
- JaCoCo subido de 0.8.12 a **0.8.14**. La 0.8.12 es de marzo 2024 y no soporta el bytecode de Java 25 (class file major version 69); fallaba con `Unsupported class file major version 69` durante la fase `report`. La 0.8.14 (agosto 2025) añade soporte oficial para Java 23–25.

## [1.2.0]

### Añadido
- **Resumen final de ejecución** (`RunReport`): bloque con tiempo total, hojas generadas (con número de filas) y warnings categorizados (`PERFIL`, `CABECERA`, `FORMULA`, `LOOKUP`, `HOJA`, `CONFIG`). Se vuelca tanto en consola como en `logs/excel-merger.log`, incluso en caso de error.
- **Detección de apps sin mapeo**: para cada fórmula `VLOOKUP({col:X}, Hoja!...)` que apunte a una hoja de lookup configurada, se recorre la hoja MES comparando los valores de la columna key con las claves del lookup. Se añade un único warning por hoja con el listado (hasta 30 valores, truncando con `...` si hay más).
- **Detección de archivos bloqueados**: pre-check del lock `~$<nombre>` y apertura de prueba antes de comenzar el proceso. Si un input o el output está abierto en Excel, se aborta con el mensaje `Cierra 'X' antes de ejecutar` en vez de fallar con un `IOException` críptico de Windows.
- **Backup del output**: nueva propiedad `output.backup=false` (default). Si se pone a `true` y el archivo de salida ya existe, se mueve antes de sobreescribirlo a `<parent>/history/<base>_yyyy-MM-dd_HHmmss.<ext>` (se crea la carpeta `history` si no existe, con fallback anti-colisión si el mismo timestamp ya existe).
- **Validación estricta del config** (`ConfigValidator`): antes de abrir ningún Excel se valida el config completo (perfiles, tipos de columna MES, placeholders `{col:X}`, referencias a hojas, lookups, derivadas, campos requeridos por tipo). Nueva propiedad `config.strictValidation=true` (default): si hay errores, aborta con código de salida `2` y los lista todos. Con `config.strictValidation=false` los errores se degradan a warnings y la ejecución continúa (comportamiento anterior).
- Nuevo código de salida `2` documentado en `--help` para config inválido.

### Cambiado
- Los builders (`LookupSheetBuilder`, `MesSheetBuilder`, `DerivedSheetBuilder`, `ExcelMerger`) ahora reciben un `RunReport` por constructor y registran ahí las hojas creadas y los problemas detectados (cabecera no encontrada, hoja inexistente, SUMIFS mal configurado, etc.), en vez de solo loguearlos silenciosamente.
- `MesSheetBuilder` hace una pre-validación de cada columna al principio: si detecta un fallo, la marca como *disabled* y avisa UNA vez (antes emitía un `log.warn` por fila afectada).

### Arreglado
- `DerivedSheetBuilder` ya no lanza `IllegalArgumentException` si falta `sourceSheet`, `groupByColumn` o `valueColumn` en una hoja `AGGREGATION`: registra un warning y la omite limpiamente.

## [1.1.0]

### Añadido
- Lanzador `run.bat` para Windows con búsqueda automática del JAR.
- Documentación actualizada (README completo).
- `CHANGELOG.md` para trazabilidad de cambios.
- `.gitattributes` para finales de línea consistentes.
- Logging con Logback: salida dual (consola + `logs/excel-merger.log` con rotación diaria).
- Soporte de argumentos `--help`, `--version` en Main.

### Cambiado
- Java: se fija la versión a **JDK 25** en el `pom.xml`.
- Dependencia de logging: `slf4j-simple` → `logback-classic`.
- Todas las clases migradas a SLF4J (niveles `info`/`warn`/`error`).

### Arreglado
- `pom.xml`: version bump a 1.1.0.

## [1.0.0] - Estado inicial funcional

### Añadido
- Fusión de dos Excel con detección por contenido (no por nombre de archivo).
- Soporte para perfiles (`profile.<id>.detect.*`).
- Hoja `MES` configurable con tipos de columna: `COPY`, `SUMIFS`, `FORMULA`, `EMPTY`.
- Placeholders `{col:Nombre}` y `{colLetter:X}` en fórmulas.
- Formato condicional "verde si ≥ 0" (`greenIfPositive`).
- Hojas de lookup con VLOOKUP (`lookup.<id>.*`), opcionalmente ocultas.
- Empaquetado con `maven-shade-plugin` respetando providers de Apache POI.
- Configuración UTF-8 en `config.properties`.

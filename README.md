# Excel Merger

Herramienta Java para fusionar dos ficheros Excel en uno solo, con:

- Detección automática de cada archivo por su contenido (no por su nombre).
- Generación de una hoja resumen (`Resultado`) con columnas copia, fórmulas (SUMIFS, VLOOKUP…) y formato condicional.
- Tablas de mapeo auxiliares (lookups) embebidas en el propio Excel resultado.
- Configuración totalmente externa en `config.properties`, editable sin recompilar.

## Estructura del proyecto

```
excel-merger/
├── pom.xml                      # Configuración Maven
├── config.properties            # ⚙️ Configuración editable
├── run.bat                      # Lanzador Windows (doble clic)
├── input/                       # 📂 Coloca aquí los dos Excel de entrada
├── output/                      # Aquí se genera el resultado
├── src/main/
│   ├── java/com/excelmerger/
│   │   ├── Main.java                    # Punto de entrada
│   │   ├── ConfigLoader.java            # Carga del config en UTF-8
│   │   ├── ConfigValidator.java         # Validación previa del config
│   │   ├── RunReport.java               # Acumula hojas/warnings del run
│   │   ├── ExcelMerger.java             # Orquestador de la fusión
│   │   ├── FileProfileResolver.java     # Identifica cada Excel por su contenido
│   │   ├── MesSheetBuilder.java         # Construye la hoja Resultado
│   │   ├── SummarySheetBuilder.java     # Hoja Resumen (sumatorio por matrícula)
│   │   ├── DerivedSheetBuilder.java     # Hojas derivadas (fórmulas/agregación)
│   │   └── LookupSheetBuilder.java      # Tablas de mapeo estáticas
│   └── resources/
│       └── config.properties            # Fallback empaquetado en el JAR
├── src/test/
│   ├── java/com/excelmerger/            # Suite JUnit 5 (9 clases, 156 tests)
│   └── resources/
│       ├── test-config.properties       # Config con placeholders para tests
│       └── fixtures/                    # extraccion.xlsx, cierre.xlsx
└── gen_fixtures.py                      # Regenera los fixtures con openpyxl
```

## Requisitos

- **Java 25** (Oracle JDK o cualquier distribución compatible).
- Maven 3.9+ **opcional**: el proyecto incluye Maven Wrapper (`mvnw` / `mvnw.cmd`), así que no hace falta tener Maven instalado. La primera ejecución del wrapper descarga Maven 3.9.9 a `~/.m2/wrapper/dists/`.

## Compilación

Con el wrapper (recomendado, no requiere Maven instalado):

```bash
# Windows
mvnw.cmd clean verify

# Linux / macOS
./mvnw clean verify
```

O bien, si ya tienes Maven en el sistema:

```bash
mvn clean package
```

Genera `target/excel-merger-1.5.0-jar-with-dependencies.jar`.

## Tests

Desde v1.2.1 el proyecto lleva una suite JUnit 5 con red de seguridad sobre los tipos de la Sesión A (`ConfigValidator`, `RunReport`, detección de locks, backup del output) y el resto de builders (`MesSheetBuilder`, `DerivedSheetBuilder`, `LookupSheetBuilder`, `FileProfileResolver`, `ExcelMerger`).

### Ejecutar la suite

```bash
# Solo tests unitarios + integración (rápido)
mvnw.cmd test

# Tests + reporte y gate de cobertura JaCoCo (≥70% INSTRUCTION)
# + spotless:check + spotbugs:check + pmd:check + checkstyle:check
mvnw.cmd clean verify
```

El reporte HTML de cobertura queda en `target/site/jacoco/index.html`. Si el gate del 70% falla, `mvnw clean verify` sale con error indicando la diferencia.

### Calidad de código (v1.5.0)

El `verify` aplica además cuatro gates automáticos:

- **Spotless**: invariantes básicos (UTF-8, LF, sin whitespace al final de línea, newline final). Auto-formatea con `mvnw.cmd spotless:apply`.
- **SpotBugs**: umbral `High`, falsos positivos documentados en `spotbugs-exclude.xml`.
- **PMD**: rulesets `java-basic`, `java-design`, `java-unusedcode`, con reglas relajadas para `src/test/java/**`.
- **Checkstyle**: `google_checks.xml` con supresiones para abreviaturas del dominio (SUMIFS, VLOOKUP, PDCL), `LineLength=140`, y `JavadocMethod` como warning.

### Organización

- **9 clases en `src/test/java/com/excelmerger/`**, 142 tests: `ConfigLoaderTest`, `ConfigValidatorTest`, `RunReportTest`, `FileProfileResolverTest`, `LookupSheetBuilderTest`, `MesSheetBuilderTest`, `DerivedSheetBuilderTest`, `AvisosSheetBuilderTest`, `ExcelMergerIntegrationTest`.
- **`TestFixtures.java`** es la utilidad compartida: copia los fixtures a un `@TempDir`, renderiza `test-config.properties` sustituyendo `${TEST_INPUT_DIR}` / `${TEST_OUTPUT_FILE}`, y ofrece `configFromProperties(...)` para tests unitarios que no tocan disco.
- Todos los tests usan `@TempDir`; no hay efectos colaterales fuera del directorio temporal de cada caso.

### Pruebas manuales del `run.bat` (v1.4.0)

La lógica de resolución de alias de entorno vive en batch, fuera de la red de seguridad JUnit. Para validarla en Windows, tres casos rápidos desde `cmd` en la carpeta del proyecto:

| Caso                           | Comando                         | Resultado esperado                                                                 |
| ------------------------------ | ------------------------------- | ---------------------------------------------------------------------------------- |
| Sin argumento                  | `run.bat`                       | El JAR arranca con `config.properties`. Log: `Ejecutando: ...jar`.                 |
| Alias con fichero existente    | `run.bat contabilidad`          | Si existe `config-contabilidad.properties`, el JAR arranca con él. Log: `Ejecutando: ...jar con config 'config-contabilidad.properties'`. |
| Alias con fichero inexistente  | `run.bat pre`                   | Sin `config-pre.properties`: el `.bat` imprime `[ERROR] No se encuentra el fichero de configuracion: config-pre.properties` y sale con `exit /b 2` sin arrancar Java. |
| Ruta `.properties` explícita   | `run.bat otra-carpeta\x.properties` | Se pasa tal cual al JAR (misma regla: case-insensitive sobre el sufijo). |

### Fixtures

`src/test/resources/fixtures/` contiene dos `.xlsx` versionados:

- **`extraccion.xlsx`** — 14 peticiones válidas + 1 fila con `Peticion` vacía (para probar el skip por ancla vacía).
- **`cierre.xlsx`** — 1 fila de metadatos + cabeceras en fila 2 + 16 imputaciones con totales conocidos (por ejemplo, `P-001` + `M-1001` suma 9 horas; así los tests de `SUMIFS` pueden asertar valores exactos).

Para regenerarlos desde cero (por ejemplo si cambian las cabeceras esperadas en los perfiles):

```bash
pip install openpyxl
python3 gen_fixtures.py
```

### Cobertura

El `pom.xml` configura JaCoCo 0.8.14 con umbral **70% INSTRUCTION a nivel de bundle**, excluyendo únicamente `com/excelmerger/Main.class` (contiene `System.exit()` y lógica de CLI que requeriría un runner externo para cubrirse). Si en alguna iteración el umbral falla en una clase concreta por ramas difíciles (p. ej. la heurística `looksLikeLocked` de `ExcelMerger` en Linux), la recomendación es **añadir exclusiones granulares** en el bloque de JaCoCo, no bajar el umbral global.

## Ejecución

### Windows (recomendado)

Doble clic en `run.bat`. Desde una terminal de Windows también puedes pasarle un **alias de entorno** como primer argumento:

```bat
run.bat                     :: usa config.properties (default)
run.bat contabilidad        :: usa config-contabilidad.properties
run.bat mi-cfg.properties   :: pasa esta ruta tal cual al JAR
```

La regla que aplica el `.bat`:
- Sin argumento → arranca el JAR sin pasar ningún config (el JAR usará `config.properties` del directorio actual).
- Argumento que termina en `.properties` → se pasa tal cual.
- Cualquier otro argumento → se resuelve a `config-<arg>.properties` en la carpeta del `.bat`.

Si el fichero resuelto no existe, el lanzador aborta con **exit code 2** y muestra la ruta buscada, **antes** de arrancar la JVM.

### Línea de comandos

```bash
java -jar target/excel-merger-1.5.0-jar-with-dependencies.jar
```

Opcionalmente se puede pasar un config alternativo:

```bash
java -jar target/excel-merger-1.5.0-jar-with-dependencies.jar mi-config.properties
```

### Dry-run (v1.4.0)

Ejecuta el pipeline completo (validación, detección de perfiles, apps sin mapeo, cabeceras, etc.) pero **no escribe el Excel de salida y no mueve el anterior a `history/`**. Útil antes de un cierre mensual para validar la configuración sin tocar el output real:

```bash
java -jar target/excel-merger-1.5.0-jar-with-dependencies.jar --dry-run
```

En el log final verás `PROCESO FINALIZADO OK (DRY-RUN, N ms)`. Los warnings (apps sin mapeo, perfiles sin match...) aparecen igualmente en el resumen. Los chequeos de lock `~$` sobre el output sí se mantienen: si el fichero está abierto en Excel, te lo dice ya.

## Flujo del programa

1. Se detectan los ficheros Excel en `input.directory`.
2. Cada uno se identifica por contenido contra los **perfiles** definidos en el config.
3. Se fusionan ambos en el Excel de salida respetando los nombres de hoja del perfil.
4. Se construyen las hojas de lookup (tablas auxiliares).
5. Se construye la hoja `Resultado` con sus columnas y fórmulas.
6. Se recalculan todas las fórmulas para que los valores se vean al abrir.

## Configuración

Todo se define en `config.properties`. Las secciones principales:

### Entrada y salida

```properties
input.directory=input
# v2.2.0: rango [min,max] de ficheros Excel aceptados. Reemplaza al
# obsoleto input.strictTwoFiles. Por defecto admite 2 (Cierre +
# Extraccion) o 3 (con el fichero opcional de Deuda, ver mas abajo).
input.strictMinFiles=2
input.strictMaxFiles=3
output.file=output/resultado_fusion.xlsx
output.overwrite=true

# Novedad v1.2.0: backup del output anterior si existe
output.backup=false
```

La clave legada `input.strictTwoFiles` sigue funcionando para compatibilidad (v2.2.0 la lee y mapea automáticamente a `strictMinFiles=2,strictMaxFiles=2`, emitiendo un warning `CONFIG` de deprecación). Si configuras ambas familias, manda la nueva.

Si `output.backup=true` y el archivo de salida ya existe, antes de sobreescribirlo se mueve a `<carpeta del output>/history/<nombre>_yyyy-MM-dd_HHmmss.xlsx` (la carpeta `history` se crea si no existe).

### Modos de generación (`output.mode`, v2.3.0)

La clave `output.mode` controla **qué hojas se generan** en el libro de salida. Acepta tres valores (estricto, case-sensitive, en minúsculas):

| Modo | Hojas generadas |
| --- | --- |
| `cierre` (**default**) | `Cierre`, `Extraccion`, `Deuda` (si hay 3er fichero), `Equipos` (oculta), `Resultado`, `Resumen`. Es el comportamiento histórico de v2.2.0. |
| `responsables` | `Cierre`, `Extraccion`, `Equipos` (oculta), `Resultado`, y **N hojas vacías** (una por cada responsable distinto que aparezca en `Resultado.Res. Tecnico`). **NO** genera `Deuda` (la copia del input se omite) ni `Resumen`. |
| `completo` | Suma de los dos: todas las hojas de `cierre` + las hojas por responsable. |

```properties
# Default si la clave esta ausente o vacia: cierre.
output.mode=cierre
```

#### Detección de responsables (modos `responsables` y `completo`)

- **Trim + case-insensitive**: valores como `tresp1@x`, `TRESP1@x` y `  tresp1@x  ` colapsan en una **única** hoja, igual que cómo `summary.byResponsible` ya normaliza responsables en v1.8.0+.
- **Nombre canónico**: el primer literal visto en orden de filas de `Resultado`. Si en `Resultado.Res. Tecnico` aparece primero `tresp1@x` (P-001) y después `TRESP1@x` (P-015), la hoja se llama `tresp1@x` y la celda `A1` muestra `tresp1@x`.
- **Sin filtros ni exclusiones**: cualquier valor no-vacío genera hoja. Valores vacíos o sólo espacios se ignoran.
- **Orden alfabético** con `Collator es_ES PRIMARY`: tildes y mayúsculas se tratan como un humano espera; salida determinista entre ejecuciones.

#### Saneo de nombres de hoja

Excel limita los nombres de hoja a 31 caracteres y prohíbe `\ / ? * [ ] :`. Si un responsable se llama `Juan Carlos / Pérez-García`:

1. Caracteres prohibidos → `_`. Truncado a 31 chars.
2. Si el saneo modifica el nombre → warning categoría `RESPONSABLE` en el resumen final y en la hoja `_Avisos` (si está habilitada).
3. Si tras sanear hay colisión con otra hoja → sufijo `_2`, `_3`, ... + warning `RESPONSABLE`.

La celda `A1` siempre muestra el **nombre canónico original** (sin sanear), aplicando el estilo `title` (negrita, 14 pt). El nombre de la pestaña es el saneado.

#### Cabecera de cada hoja por responsable

La celda `A1` contiene el **nombre canónico del responsable** (con estilo título: negrita, 14 pt). A partir de v2.4.0, debajo se generan dos tablas pivot.

#### Tablas pivot Petición × Matrícula (v2.4.0)

Por defecto cada hoja de responsable contiene, debajo de `A1`, **dos tablas pivot SUMIFS** apiladas verticalmente:

1. **Horas imputadas (Jira) por Petición × Matrícula.**
2. **REAL por Petición × Matrícula.**

Las tablas son fórmulas vivas contra `Resultado`, filtradas por el responsable cuyo nombre figura en `A1`. Las peticiones y matrículas que aparecen son únicamente las que ese responsable tiene en `Resultado` (no todas las del libro), por consistencia con la Tabla 2 de Resumen.

```
Excel row     Contenido (con responsable de 2 peticiones × 2 matrículas)
─────────     ──────────────────────────────────────────────────────────────
1             tresp1@x                                       (estilo título)
2             — vacía —
3             [merged] Horas imputadas (Jira) por Petición × Matrícula
4             Petición  | M-1001  | M-1002  | Total
5             P-001     | =SUMIFS | =SUMIFS | =SUM(B5:C5)
6             P-002     | =SUMIFS | =SUMIFS | =SUM(B6:C6)
7             Total     | =SUM    | =SUM    | =SUM(D5:D6)
8, 9          — gap (configurable, default 2 filas) —
10            [merged] REAL por Petición × Matrícula
11            Petición  | M-1001  | M-1002  | Total
12, 13, 14    … (idéntica estructura, columna REAL)
```

Las pivots se controlan con cuatro claves nuevas en `config.properties`:

```properties
# Habilita/deshabilita las tablas pivot. Default true (opt-out).
responsables.tables.enabled=true

# Títulos (literal en la fila merged).
responsables.tables.jiraTitle=Horas imputadas (Jira) por Petición × Matrícula
responsables.tables.realTitle=REAL por Petición × Matrícula

# Filas en blanco entre las dos tablas. Admite 0. Default 2.
responsables.tables.gapRows=2
```

El rango de los SUMIFS se acota con la clave existente `summary.sumifsMaxRow` (compartida con la hoja `Resumen`). Para preservar el comportamiento exacto de v2.3.0 (hoja con solo `A1`), añadir `responsables.tables.enabled=false`.

#### Ejemplos de uso

```bash
# Cierre completo del mes (comportamiento histórico):
java -jar excel-merger.jar config.properties

# Solo Resultado + plantillas vacías por responsable para repartir:
# (en el config: output.mode=responsables)
java -jar excel-merger.jar config-responsables.properties

# Todo en un único libro: cierre + plantillas por responsable:
# (en el config: output.mode=completo)
java -jar excel-merger.jar config-completo.properties
```

#### Comportamiento de `Deuda` en modo `responsables`

Si el usuario aporta `deuda.xlsx` en el directorio de entrada y configura `output.mode=responsables`, el fichero se detecta normalmente pero **su hoja no se copia al libro de salida** (el motor lo omite con un warning `CONFIG`). Las fórmulas `PDCL + Deuda` en `Resultado` se degradan automáticamente al modo "sin Deuda", igual que en v2.2.0 cuando el usuario no aporta el 3er fichero.

#### Validación

Si `output.mode` tiene un valor no permitido (`Cierre` con mayúscula, `foo`, ...), `ConfigValidator` lo reporta como error listando los 3 válidos. Con `config.strictValidation=true` (default) el run aborta con exit code 2; con `strictValidation=false` el motor cae a `cierre` defensivamente.

### Perfiles (detección por contenido)

Cada perfil define las cabeceras que deben encontrarse para considerar que un Excel es de ese tipo.

> **v2.0.0 — cambio de nombres de perfil**: hasta la versión 1.8.1, el perfil que detectaba las peticiones del ERP se llamaba `Extraccion`, y el perfil del export de Jira se llamaba `Cierre`. Estos nombres eran contraintuitivos respecto a los nombres habituales de los ficheros de entrada. En 2.0.0 se intercambian: el perfil de peticiones ERP pasa a llamarse `Cierre` y el perfil de Jira pasa a llamarse `Extraccion`. Ver [CHANGELOG 2.0.0](CHANGELOG.md) para la guía de migración de un `config.properties` heredado de 1.x.

```properties
profiles=Cierre,Extraccion,Deuda

# Perfil Cierre: peticiones del ERP, cabeceras en fila 1.
profile.Cierre.sheetName=Cierre
profile.Cierre.detect.headerRow=1
profile.Cierre.detect.headers=Peticion,Titulo,Estado,Recurso
profile.Cierre.detect.minMatches=4

# Perfil Extraccion: export de Jira, cabeceras en fila 2
# (la 1 suele ser una fila de metadatos/título).
profile.Extraccion.sheetName=Extraccion
profile.Extraccion.detect.headerRow=2
profile.Extraccion.detect.headers=Project Key,Issue Key,Hours
profile.Extraccion.detect.minMatches=3

# Perfil Deuda (v2.2.0): OPCIONAL. Ver seccion "Archivo de Deuda opcional".
profile.Deuda.sheetName=Deuda
profile.Deuda.detect.headerRow=1
profile.Deuda.detect.headers=Peticion,Matricula,Funcion,Horas
profile.Deuda.detect.minMatches=4
```

Criterios soportados:
- `detect.headers` — lista de cabeceras (case-insensitive, match por "contains").
- `detect.minMatches` — cuántas deben coincidir.
- `detect.cellValue.<REF>` — opcional, fuerza que una celda concreta contenga un valor.
- `detect.sheetIndex` — opcional, qué hoja analizar (default 0).

#### Normalización al copiar: `asText.columns` y `trim.columns`

Al copiar de los Excels origen al workbook resultado, el perfil puede normalizar columnas concretas. Dos claves opt-in:

```properties
# v1.6.2: columnas que se fuerzan a STRING al copiar (aunque vengan NUMERIC
# en el origen). Útil para cruces por SUMIFS donde un lado viene numérico
# y el otro textual: el SUMIFS no casa si los tipos difieren.
profile.Cierre.asText.columns=Peticion,Recurso,Usuario_Resp_Tecnico

# v1.8.1: columnas a las que se aplica trim() tras el cast a STRING. Útil
# para exports ERP que alinean códigos con padding de espacios
# ("MG002   "). El SUMIFS de Excel es case-insensitive pero NO
# trim-insensitive: "MG002" no casa contra "MG002   ". Sin trim, el bug
# aparece en cualquier hoja aguas abajo que haga SUMIFS contra esta
# columna (p. ej. la segunda tabla de Resumen).
#
# Regla: una columna se trima solo si está también en asText.columns
# (el trim es una capa sobre la rama STRING del cast). Si se declara en
# trim pero no en asText, warning `CONFIG` en runtime y se ignora. Si se
# declara trim.columns sin ningún asText.columns, error de validación
# duro.
profile.Cierre.trim.columns=Recurso,Usuario_Resp_Tecnico
profile.Extraccion.trim.columns=Matricula
```

### Archivo de Deuda opcional (v2.2.0)

A partir de v2.2.0, el programa acepta un **tercer fichero Excel opcional** en `input/` que aporta horas de deuda por `(Peticion, Matricula, Funcion)`. Cuando está presente, la columna `PDCL + Deuda` de la hoja `Resultado` deja de ser igual a `PDCL` y pasa a sumar las horas cruzadas desde la nueva hoja `Deuda`. Cuando no está, el comportamiento es **idéntico** a v2.1.0 (`PDCL + Deuda == PDCL`, sin warnings).

**Cabeceras esperadas** (fila 1, nombre del fichero irrelevante; detección por contenido):

| `Peticion`    | `Matricula` | `Funcion` | `Horas` |
|---------------|-------------|-----------|---------|
| P-001         | M-1001      | Dev       | 5       |
| P-002         | M-1002      | Dev       | 3       |
| P-001         | M-1001      | Dev       | 2       |

**Cómo se usa la suma.** La columna `PDCL + Deuda` de `Resultado` se convierte en:

```
{col:PDCL} + IFERROR(SUMIFS(Deuda[Horas]; Deuda[Peticion]=Resultado[Petición];
                            Deuda[Matricula]=Resultado[Matrícula];
                            Deuda[Funcion]=Resultado[Funcion]); 0)
```

- La **clave de cruce** son las tres columnas `Peticion + Matricula + Funcion`. `Funcion` en el fichero Deuda se cruza contra la columna `Funcion` de `Resultado` (el "rol" `Dev`/`Sup`/...). **No confundir con `Res. Tecnico`** de `Resultado`, que es el nombre del responsable-persona (copiado de `Usuario_Resp_Tecnico` de Cierre) y no se usa como clave de cruce con Deuda.
- Si una fila de `Resultado` no tiene entrada correspondiente en Deuda, el SUMIFS devuelve 0 y `PDCL + Deuda == PDCL` para esa fila.
- Varias filas de Deuda con la misma clave se **agregan**: con las filas del ejemplo, la fila P-001/M-1001/Dev de `Resultado` suma `5 + 2 = 7h` de deuda.
- El `IFERROR(...,0)` es una defensa: si la hoja `Deuda` no existe en el libro (porque el fichero no se aportó), el SUMIFS devuelve `#REF!` y se convierte en 0 — por tanto la fórmula equivale a solo `PDCL`.
- Las letras de columna de la hoja Deuda (`$A/$B/$C/$D`) se **resuelven en runtime** leyendo la cabecera, no se hardcodean. Si el Excel de Deuda viene con columnas en otro orden o columnas extra intercaladas, la fórmula sigue apuntando a la columna correcta.

**Normalización defensiva.** El perfil Deuda aplica `asText.columns=Peticion,Matricula` y `trim.columns=Matricula`, igual que Cierre y Extraccion, para evitar los mismatches de tipo numérico/textual (v1.6.2) y de padding de espacios (v1.8.1). `Funcion` deliberadamente no se normaliza; si tu fichero de Deuda llega con `"DEV  "` con espacios añade `Funcion` a ambas listas.

**Comportamiento cuando no se aporta el fichero.**

- El libro de salida NO tendrá hoja `Deuda`.
- La fórmula de `PDCL + Deuda` será solo `={col:PDCL}` (misma que v2.1.0).
- `PDCL + Deuda == PDCL` para todas las filas → la regla `redIfNotEqualTo=PDCL` no pinta nada de rojo (identidad).
- **No se emite ningún warning**. El degradado es silencioso por diseño: si tu empresa todavía no usa el fichero de Deuda, los `_Avisos` quedarían contaminados de warnings sin valor.

**Posición en el libro de salida.** `Cierre, Extraccion, [Deuda si existe,] Equipos (oculta), Resultado, Resumen, [_Avisos si opt-in]`. Orden garantizado por `merge.profileOrder` (default `Cierre,Extraccion,Deuda`) independientemente del orden alfabético de los ficheros en `input/`.

### Hoja Resultado

Estructura fija de columnas definida en el config. Cada columna tiene un tipo:

- **COPY** — copia directa de una columna de la hoja origen.
- **SUMIFS** — suma condicional cruzando con otra hoja.
- **FORMULA** — fórmula libre con placeholders `{col:NombreColumnaMES}` y `{colLetter:X}`.
- **FORMULA_PLUS_SUMIFS** (v2.2.0) — fórmula con una parte base `{col:X}` y un SUMIFS opcional que se concatena solo si la hoja remota existe. Se usa para la columna `PDCL + Deuda`. Campos: `baseFormula`, `from`, `sum`, `match` (con la misma sintaxis que `SUMIFS`).
- **EMPTY** — celda vacía (placeholder).

Modificadores opcionales por columna:

- `greenIfPositive=true` — formato condicional: fondo verde si la celda es ≥ 0.
- `fill=<COLOR>` — fondo sólido permanente para toda la columna. Valores: `LIGHT_GREEN` (`#E2EFDA`), `MEDIUM_GREEN` (`#C6E0B4`), `LIGHT_BLUE` (`#DDEBF7`), `LIGHT_YELLOW` (`#FFF2CC`), `LIGHT_RED` (`#FCE4D6`), `LIGHT_LAVENDER` (`#E4DFEC`). Nombres desconocidos generan un warning `CONFIG` y la columna se escribe sin fill.
- `redIfNotEqualTo=<NombreColumna>` — formato condicional: pinta la celda de rojo claro cuando su valor difiere del de la celda homóloga en la columna referenciada (misma fila). Útil para resaltar diferencias entre dos columnas relacionadas.

Ejemplo:

```properties
mes.enabled=true
mes.sheetName=Resultado
mes.sourceSheet=Cierre
mes.sourceHeaderRow=1
mes.anchorColumn=Peticion

mes.col.1.name=Petición
mes.col.1.type=COPY
mes.col.1.from=Peticion

# v2.1.0: columna Funcion (AN, DI, PR, OT, IN, RE, SC, TE, ...) copiada
# desde Cierre, inmediatamente despues de Matricula. Si el valor original
# es "-", se copia tal cual.
mes.col.7.name=Funcion
mes.col.7.type=COPY
mes.col.7.from=Funcion

mes.col.10.name=Jira
mes.col.10.type=SUMIFS
mes.col.10.from=Extraccion
mes.col.10.sum=Hours
mes.col.10.match=Component Name:Peticion,Matricula:Recurso,Funcion:Funcion

mes.col.11.name=REAL
mes.col.11.type=FORMULA
mes.col.11.formula={col:Jira}*1.2

mes.col.12.name=PDCL
mes.col.12.type=FORMULA
mes.col.12.formula={col:Jira}*1.2
mes.col.12.fill=LIGHT_GREEN

mes.col.13.name=PDCL + Deuda
mes.col.13.type=FORMULA
mes.col.13.formula={col:PDCL}
mes.col.13.fill=MEDIUM_GREEN
mes.col.13.redIfNotEqualTo=PDCL
```

En el ejemplo anterior, `PDCL` aparece sobre fondo verde muy claro y `PDCL + Deuda` sobre fondo verde un tono más oscuro; además, cualquier celda de `PDCL + Deuda` cuyo valor no coincida con el de `PDCL` en la misma fila se pinta en rojo claro, como alerta visual de modificaciones manuales.

**Columna `Funcion` (v2.1.0)**: tras `Matrícula`, se añade una columna `Funcion` que se copia tal cual desde `Cierre.Funcion`. En `Cierre`, una misma matrícula puede aparecer en varias peticiones con funciones distintas (`AN`, `DI`, `PR`, `OT`, `IN`, `RE`, `SC`, `TE`, ...); cada petición es su propia fila en `Resultado`, y cada fila expone su propia función. No hay agregación ni concatenación: si la matrícula `M-1001` tiene 3 peticiones con funciones `AN`, `DI` y `PR`, aparecen 3 filas en `Resultado` con esa matrícula, una por función. El literal `"-"` del origen se preserva sin normalizar.

### Hojas de lookup

Tablas estáticas embebidas en el Excel para usarlas con VLOOKUP:

```properties
lookup.sheets=Equipos

lookup.Equipos.header1=App
lookup.Equipos.header2=Equipo
lookup.Equipos.hidden=true
lookup.Equipos.data=\
  DF:Iker,\
  HE:Jon,\
  EW:JAVA
```

Luego se referencia desde una columna FORMULA:

```properties
mes.col.3.name=Equipo
mes.col.3.type=FORMULA
mes.col.3.formula=IFERROR(VLOOKUP({col:Aplicación},Equipos!$A:$B,2,FALSE),"")
```

### Hoja Resumen (v1.6.0)

Tabla adicional con el sumatorio por matrícula de varias columnas de `Resultado`. Se construye tras `MesSheetBuilder`, de modo que las fórmulas `SUMIFS` pueden referenciar las columnas calculadas (`PDCL`, `PDCL + Deuda`, etc.).

Layout generado:

- Fila 1: título `Resumen por <matriculaColumn>` con merge sobre todo el ancho (estilo negro/blanco, negrita).
- Fila 3: cabecera (nombre de la columna clave + cada columna de valor, estilo gris).
- Filas 4..N: una por cada matrícula única detectada en `Resultado`. Se emiten fórmulas `SUMIFS` con rangos acotados (`I2:I10000`) para que Apache POI pueda evaluarlas en los tests.
- Fila N+1: `Total` + `SUM(4:last)` por cada columna (estilo gris, bordes medium, negrita).

```properties
# Habilita la hoja (opt-in explícito)
summary.enabled=true

# Nombre de la hoja generada
summary.sheetName=Resumen

# Hoja de donde se suman los valores (normalmente la hoja Resultado)
summary.sumSheet=Resultado

# Nombre de la columna clave. Debe coincidir con alguno de mes.col.N.name
summary.matriculaColumn=Matrícula

# Columnas cuyos valores se suman por matrícula (CSV).
# Los nombres deben coincidir con mes.col.N.name. Las columnas no
# encontradas se omiten con un warning, sin abortar la generación.
summary.valueColumns=Jira,REAL,PDCL,PDCL + Deuda

# Tope de fila para los rangos SUMIFS. El cuaderno original usa columnas
# completas (I:I); aquí se acota para que POI pueda evaluarlas en tests.
# Cambia este valor si alguna extracción mensual supera 10 000 filas.
summary.sumifsMaxRow=10000
```

Las matrículas se auto-descubren leyendo la columna clave en `sumSheet`. Se incluyen todos los valores no vacíos: las numéricas (`99641`, `99642`...) se ordenan ascendentemente, y después se listan las no numéricas (`-`, `Sin Matricula`, etc.) en orden alfabético. Los vacíos/nulos se filtran.

El validador estricto comprueba que `summary.sheetName` no colisione con ninguna otra hoja del libro, que `summary.sumSheet` sea una hoja conocida, y que `summary.valueColumns` no esté vacío. Si alguna columna listada en `valueColumns` no existe en `sumSheet`, no se aborta: se emite un warning `CABECERA` en el `RunReport` y se sigue con las que sí existen.

#### Segunda tabla: matriz Matrícula × Responsable (v1.8.0)

Opcional, opt-in. Añade una **segunda tabla** debajo de la primera, en la misma hoja `Resumen`, cruzando matrículas (filas) con responsables técnicos (columnas) y mostrando el `PDCL` de cada par. Utilidad: la primera tabla responde "¿cuántas horas ha metido cada matrícula?"; la segunda responde "¿qué responsables están imputando en cada matrícula?".

Layout generado (ejemplo con 3 matrículas y 3 responsables):

|              | TRESP1@X | TRESP2@X | TRESP3@X | **Total** |
|--------------|----------|----------|----------|-----------|
| **99641**    | 10.8     | 0.0      | 0.0      | **10.8**  |
| **99642**    | 0.0      | 6.0      | 0.0      | **6.0**   |
| **-**        | 0.0      | 0.0      | 1.8      | **1.8**   |
| **Total**    | **10.8** | **6.0**  | **1.8**  | **18.6**  |

Puntos clave:

- La fila `Total` sumada por columna y la columna `Total` sumada por fila se cruzan en la esquina inferior derecha con el **gran total**, que debe coincidir con el total `PDCL` de la primera tabla. Si no coinciden hay algo mal en los datos (útil como sanity check al abrir el Excel).
- Los responsables se normalizan a MAYÚSCULAS (`trim()` + `toUpperCase(Locale.ROOT)`) al descubrir los códigos únicos de la columna `Res. Tecnico`. Códigos como `tresp1@x`, `TRESP1@x` y ` Tresp1@X ` colapsan en una única columna `TRESP1@X` de la cabecera. Para que el `SUMIFS` también **sume correctamente** todas esas variantes (no solo las agrupe en la cabecera), los valores del origen tienen que llegar a `Resultado` sin padding de espacios — por eso en 1.8.1 se añadió la clave `profile.*.trim.columns` que aplica `trim()` a la columna `Usuario_Resp_Tecnico` en la capa de copia. Ver la sección "Perfiles — `trim.columns`" más abajo para el detalle. Excel `SUMIFS` es case-insensitive sobre texto, así que la normalización a MAYÚSCULAS en cabecera y la tolerancia de casing en el criterio se complementan.
- El orden es alfabético puro por el código normalizado.
- Los ceros se dejan como `0` numérico (la mayoría de celdas lo serán: un responsable típico no toca todas las matrículas).

```properties
# Habilita la segunda tabla (opt-in). Si false o ausente, Resumen queda
# como en 1.7.1.
summary.byResponsible.enabled=true

# Columna de Resultado usada como agrupador de columnas.
# Debe coincidir con uno de los mes.col.N.name.
summary.byResponsible.column=Res. Tecnico

# Columna de valor a sumar (una sola, no lista).
# Debe coincidir con uno de los mes.col.N.name.
summary.byResponsible.valueColumn=PDCL

# Título visible de la segunda tabla (merge sobre el ancho).
summary.byResponsible.title=Totales Peticiones por Responsables Matrículas

# Número de filas en blanco entre la tabla de matrículas y esta.
# Admite 0 (tablas pegadas). Default 2.
summary.byResponsible.gapRows=2
```

Restricciones:

- `summary.byResponsible.enabled=true` requiere `summary.enabled=true`. Si se configura la segunda tabla sin la primera, el validador aborta con un error explícito (la segunda tabla se renderiza **dentro** de la hoja `Resumen`; no tiene sentido sola).
- Si la columna configurada en `summary.byResponsible.column` o `summary.byResponsible.valueColumn` no existe en `Resultado`, la segunda tabla se omite con un warning `CABECERA`, pero la primera queda intacta.

## Robustez y diagnóstico (v1.2.0)

### Validación estricta del config

Antes de abrir ningún Excel, la aplicación valida el `config.properties` al completo:

- Perfiles incompletos (sin criterios de detección).
- Columnas de Resultado con tipo inválido o campos requeridos ausentes.
- Placeholders `{col:X}` que no coinciden con ninguna columna de Resultado.
- Referencias a hojas inexistentes (SUMIFS `from`, AGGREGATION `sourceSheet`).
- Lookups vacíos, derivadas con tipo desconocido, etc.

Comportamiento configurable:

```properties
# Por defecto: si hay errores, aborta con código de salida 2 y los lista.
config.strictValidation=true
```

Si se pone a `false`, los errores se degradan a warnings (aparecen en el resumen final) y la ejecución continúa como en versiones anteriores.

### Hoja de avisos en el Excel (v1.4.0)

Todos los warnings acumulados durante la ejecución (apps sin mapeo, cabeceras no encontradas, perfiles sin match, entradas de lookup malformadas...) se pueden volcar a una hoja extra del Excel resultado para revisarlos sin abrir el log. Es **opt-in**:

```properties
# Si true, añade una hoja '_Avisos' al Excel resultado con los warnings.
# Si no hay warnings, la hoja no se crea.
report.inExcel=false

# Nombre de la hoja (default: _Avisos)
report.sheetName=_Avisos

# Si true (default), la hoja se crea oculta.
# Se ve en Excel con clic derecho sobre una pestaña > Mostrar...
report.hidden=true
```

La hoja tiene dos columnas (`Categoria` y `Mensaje`) y un warning por fila, en el orden en que se produjeron. Las categorías habituales son `PERFIL`, `CABECERA`, `FORMULA`, `LOOKUP`, `HOJA` y `CONFIG`. En modo `--dry-run` la hoja también se construye (en memoria), aunque no se escriba a disco.

### Códigos de salida

Desde la v1.3.0 el CLI devuelve códigos de salida tipados según la clase de excepción. Los códigos 0, 1 y 2 son retrocompatibles con versiones anteriores; 3 y 4 son nuevos.

| Código | Significado                                                                  | Excepción asociada          |
| ------ | ---------------------------------------------------------------------------- | --------------------------- |
| `0`    | Ejecución correcta                                                           | —                           |
| `1`    | Error genérico en tiempo de ejecución                                        | `MergeException` u otra     |
| `2`    | Configuración inválida o no cargable (o `config.strictValidation=true` con errores) | `ConfigurationException` |
| `3`    | Entrada inválida: directorio o ficheros Excel mal, lock sobre un input       | `InputValidationException`  |
| `4`    | Salida inválida: lock sobre el output, `overwrite=false` con output existente, fallo de escritura/backup | `OutputException` |

Todas las excepciones heredan de `com.excelmerger.exception.ExcelMergerException` (runtime). Scripts que integren el merger pueden ramificar por código; por ejemplo:

```bash
java -jar excel-merger.jar
code=$?
case $code in
  0) echo "OK" ;;
  2) echo "Revisa el config.properties" ;;
  3) echo "Revisa el directorio de entrada" ;;
  4) echo "Cierra el Excel de salida o cambia output.overwrite" ;;
  *) echo "Error inesperado (exit $code)" ;;
esac
```

### Detección de archivos bloqueados

Antes de empezar se comprueba que los Excel de entrada y el de salida no estén abiertos en Excel (busca el lock `~$<nombre>` y hace una apertura de prueba). Si alguno está abierto, se aborta con un mensaje claro:

```
Cierra 'Extraccion.xlsx' antes de ejecutar (parece abierto en Excel).
```

### Resumen final de ejecución

Al terminar (con éxito o con error) se vuelca un bloque tipo:

```
====================================
   RESUMEN DE EJECUCION
====================================
Tiempo total: 2345 ms
Hojas generadas (4):
  - Extraccion (122 filas)
  - Cierre (2500 filas)
  - Equipos (34 filas)
  - Resultado (122 filas)
Warnings (2):
  - [LOOKUP] 2 valor(es) de 'Aplicación' sin mapeo en 'Equipos': XY, WZ
  - [CABECERA] Columna 'Unknown' no encontrada en 'Extraccion' (usada por COPY 'Foo').
====================================
```

Categorías de warnings:

| Categoría | Qué indica |
|-----------|------------|
| `PERFIL`  | Un archivo no coincide con ningún perfil, o ambos coinciden con el mismo. |
| `CABECERA`| Una cabecera esperada no está en la hoja origen. |
| `FORMULA` | Un placeholder `{col:X}` no coincide, o falla el recalculo de derivadas. |
| `LOOKUP`  | Entradas malformadas, o valores que no tienen mapeo en una hoja de lookup. |
| `HOJA`    | Se referencia una hoja que no existe en el libro resultado. |
| `CONFIG`  | Otros problemas de configuración detectados en runtime. |

### Detección de apps sin mapeo en VLOOKUP

Para cada fórmula `VLOOKUP({col:X}, Hoja!...)` que apunte a una hoja de **lookup configurado**, la aplicación recorre la hoja Resultado comparando los valores de la columna `X` con las claves del lookup. Cualquier valor que no tenga mapeo se agrega en un único warning por hoja de lookup (hasta 30 valores de ejemplo).

Útil para detectar códigos de aplicación nuevos que aún no están dados de alta en la tabla `Equipos`.

## Troubleshooting

**"Your InputStream was neither an OLE2 stream, nor an OOXML stream"**
El JAR se construyó mal. Comprueba que el `pom.xml` usa `maven-shade-plugin` y regenera con `mvn clean package`.

**"El archivo ya existe y output.overwrite=false"**
Cierra el Excel si lo tienes abierto, o pon `output.overwrite=true` (o `output.backup=true` para conservar una copia).

**"Cierra 'X.xlsx' antes de ejecutar (parece abierto en Excel)"**
Tienes abierto el archivo indicado en Excel. Ciérralo y vuelve a lanzar.

**"CONFIGURACION INVALIDA ({N} error(es))" seguido de exit 2**
El config tiene errores que impiden arrancar. Revisa la lista y corrige, o ponlo temporalmente en `config.strictValidation=false` para ver el detalle sin abortar.

**"No coincide con ningún perfil"**
Las cabeceras del archivo no corresponden con las esperadas. Revisa `detect.headers` del perfil para que coincida con las del archivo.

**Se ignoran filas con la columna ancla vacía.**
Es el comportamiento por defecto en Resultado. Si tu Peticion puede ser vacía y quieres mantener la fila, cambia `mes.anchorColumn` a otra columna siempre informada.

## Carga del fichero de configuración

La aplicación busca `config.properties` en este orden:
1. **Fichero externo** en el directorio de ejecución (recomendado).
2. **Recurso interno** empaquetado en el JAR (fallback).

El fichero se lee en **UTF-8**, admite acentos directamente.

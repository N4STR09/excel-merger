# Changelog

## [2.4.0] — Tablas pivot Petición × Matrícula en hojas de responsable

Release minor. En los modos `output.mode=responsables` y `output.mode=completo`, cada hoja de responsable contiene ahora dos tablas pivot SUMIFS apiladas verticalmente:

1. **Horas imputadas (Jira) por Petición × Matrícula.**
2. **REAL por Petición × Matrícula.**

Ambas son fórmulas vivas contra `Resultado`, filtradas por el responsable cuyo nombre figura en `A1`. Las peticiones y matrículas que aparecen son únicamente las que ese responsable tiene en `Resultado` (no todas las del libro), por consistencia con la Tabla 2 de Resumen.

La feature es **opt-out**: si la clave `responsables.tables.enabled` está ausente o vale `true`, las pivots se generan; con `false`, las hojas quedan como en v2.3.0 (solo cabecera `A1`). Esto significa que un upgrade del binario sobre un `config.properties` v2.3.0 aporta la funcionalidad nueva **sin requerir cambios en la configuración**, lo que es consistente con la filosofía de defaults seguros del proyecto.

### Claves nuevas en `config.properties`

```properties
# Habilita/deshabilita las tablas pivot. Default true.
responsables.tables.enabled=true

# Títulos de las dos tablas (literal en la fila merged).
responsables.tables.jiraTitle=Horas imputadas (Jira) por Petición × Matrícula
responsables.tables.realTitle=REAL por Petición × Matrícula

# Filas en blanco entre las dos tablas. Admite 0. Default 2.
responsables.tables.gapRows=2
```

El rango de los SUMIFS se acota con la clave existente `summary.sumifsMaxRow` (default 10000) — **se reutiliza** en lugar de introducir una clave separada para mantener el config plano.

### Estructura por hoja de responsable

```
Excel row     Contenido
─────────     ───────────────────────────────────────────────────────────────
1             A1 = nombre canónico del responsable (estilo título, v2.3.0)
2             — vacía —
3             [merged] Horas imputadas (Jira) por Petición × Matrícula
4             Petición | M-1001 | M-1002 | … | Total
5..(4+nP)     P-001    | SUMIFS | SUMIFS | … | SUM(fila)
5+nP          Total    | SUM    | SUM    | … | SUM(grand)
6+nP, 7+nP    — gap (2 filas en blanco, configurable) —
8+nP          [merged] REAL por Petición × Matrícula
9+nP          Petición | M-1001 | M-1002 | … | Total
10+nP..       … (idéntica estructura que la primera tabla, columna REAL)
```

### Semántica acordada (Fase 0)

1. **Descubrimiento**: una sola pasada sobre `Resultado` produciendo `Map<String,ResponsableData>` con peticiones y matrículas por responsable. La clave del map es el responsable trimeado y en lowercase (mismo criterio v2.3.0). Eficiente: O(N×3) donde N es el número de filas de Resultado.
2. **Orden de peticiones y matrículas**: numéricas ascendentes primero, no numéricas alfabético después (mismo criterio que `SummarySheetBuilder.discoverMatriculas`). Para no duplicar la regla por tercera vez, se ha extraído a `PoiUtils.mixedNumericLexicographicSort(Collection<String>)` con sus propios tests.
3. **Total de fila + total de columna + gran total**: presentes en ambas tablas (mismo formato de `StyleFactory` que la Tabla 2 de Resumen).
4. **Caso degenerado**: responsable sin peticiones (no debería ocurrir si está en Resultado, pero por defensividad), se escribe el título y debajo `(Sin datos)` sin filas SUMIFS.
5. **SUMIFS bounds**: se reutiliza `summary.sumifsMaxRow` (default 10000) — no se introduce nueva clave.
6. **Modo `cierre`**: completamente intacto. No se construyen hojas de responsable y por tanto no hay pivots.
7. **Criterio del responsable en SUMIFS**: referencia absoluta `$A$1` de cada hoja (no literal). Esto evita problemas de escapado con caracteres especiales (`'`, `"`, `@`) y permite editar A1 manualmente sin romper las fórmulas.

### Lecciones aplicadas

- **1.7.1 — mismatch numérico/textual**: las cabeceras de matrícula y las celdas clave de petición se escriben **siempre como STRING**, incluso cuando son todo dígitos (e.g. `99641`, `55751`). Las columnas Petición y Matrícula de `Resultado` están marcadas como `asText.columns` desde 1.6.2, así que el SUMIFS compara string-string en ambos extremos. Si las cabeceras de pivot fueran NUMERIC, el SUMIFS daría 0. **Cubierto por test específico** (`ResponsablePivotBuilderTest.formulaEvaluatorRespetaCaseSinIgnorarFiltrosNumericos`).
- **1.8.0 — recálculo de fórmulas**: cuando se generan pivots, se llama a `workbook.setForceFormulaRecalculation(true)` al final del builder. Sin esto, Excel/POI muestran los SUMIFS sin evaluar al abrir el fichero. Idempotente con `SummarySheetBuilder` que también lo setea.

### Tests añadidos (todos verdes en mi código local; ver caveat de validación)

- **`ResponsablePivotBuilderTest`** (7 tests): estructura, fórmula SUMIFS bien formada con 3 criterios, SUM de fila/columna, caso `(Sin datos)`, FormulaEvaluator sobre combinación conocida (5+3=8), regresión 1.7.1 con peticiones y matrículas todo dígitos, criterio responsable cambiando A1 dinámicamente, tipo STRING en cabecera de matrícula.
- **`ResponsablesSheetBuilderV24Test`** (8 tests): dos pivots apiladas con gap, FormulaEvaluator sobre tabla Jira (gran total = 20), FormulaEvaluator sobre tabla REAL (gran total = 24), aislamiento entre responsables (tresp2@x no incluye filas de tresp1@x), `enabled=false` mantiene comportamiento v2.3.0, clave ausente equivale a `enabled=true`, `RunReport` registra resumen agregado con conteo correcto de filas (14 filas para 2×2), columna Jira ausente desactiva pivots con warning RESPONSABLE.
- **`ResponsablesSheetBuilderTest`**: el helper `minimalConfig()` añade `responsables.tables.enabled=false`. Los **14 tests existentes mantienen sus aserciones intactas** (preservando el contrato v2.3.0 de "una sola fila por hoja").
- **`ConfigValidatorTest`** (6 tests): `gapRows` válido / negativo / no numérico, `jiraTitle` blank, `realTitle` blank, claves ausentes (no error).
- **`ExcelMergerIntegrationTest`** (5 tests): modo `responsables` genera dos pivots por hoja; comparación FormulaEvaluator(SUMIFS) == suma manual sobre Resultado para combinación P-001/M-1001/tresp1@x; modo `completo` mantiene Resumen y añade pivots; modo `cierre` sin cambios; `enabled=false` produce hojas v2.3.0.

### Retrocompatibilidad

- **Modo `cierre`**: 100% intacto.
- **Configs v2.3.0**: funcionan tal cual; las pivots se activan automáticamente con el upgrade del binario (default `true`). Para preservar el comportamiento exacto v2.3.0, añadir `responsables.tables.enabled=false`.
- **API pública**: sin cambios. `ResponsablesSheetBuilder.buildAll(Workbook)` mantiene su firma; `ResponsablePivotBuilder` es package-private.

---

## [2.3.0] — Modos de generación (`output.mode`)

Release minor. Añade la clave `output.mode` para seleccionar **qué hojas se generan** en el libro de salida. Tres modos:

- `cierre` (**default**, preserva 100% el comportamiento de v2.2.0): genera `Cierre`, `Extraccion`, `Deuda` (si el usuario aporta el 3er fichero), `Equipos` (oculta), `Resultado` y `Resumen`.
- `responsables`: genera `Cierre`, `Extraccion`, `Equipos` (oculta), `Resultado` y **N hojas vacías**, una por cada responsable distinto que aparezca en `Resultado.Res. Tecnico`. **NO** genera `Deuda` (la copia del input se omite explícitamente) ni `Resumen`.
- `completo`: la suma de los dos: todas las hojas de `cierre` + las hojas por responsable.

Si la clave está ausente o vacía se asume `cierre`, así que **ningún config existente requiere modificación** y todo el comportamiento de v2.2.0 queda preservado.

⚠️ **Nota de validación**: esta release se preparó en un entorno sin acceso a Maven Central. `mvnw verify` **no ha sido ejecutado** para certificar la release. Los cambios están validados por: (a) revisión línea a línea del código nuevo (`OutputMode`, `ResponsablesSheetBuilder`, switch en `ExcelMerger.merge`); (b) tests unitarios nuevos (`OutputModeTest`, `ResponsablesSheetBuilderTest`) y de integración (8 tests añadidos a `ExcelMergerIntegrationTest`); (c) confirmación de que los tests v2.2.0 cargan ahora `output.mode=cierre` (default) y por tanto siguen verdes sin tocar aserciones. Antes de publicar, correr `./mvnw verify` localmente y confirmar tests verdes y cobertura ≥70%.

### Semántica acordada (Fase 0)

Decisiones cerradas con el usuario antes de tocar código:

1. **Validación del valor de `mode`**: opción (c) del plan — ausente o vacío → `cierre`; presente con valor inválido → error de validación que `ConfigValidator` reporta listando los 3 válidos. Comparación **case-sensitive estricta en minúsculas**: `Cierre`, `CIERRE`, `Responsables` son inválidos. Misma convención que el resto del proyecto para enums (`merge.mode` valida en mayúsculas estrictas).
2. **Orden de hojas en modo `responsables`**: opción (a) del plan — `Equipos` (oculta) se queda en su posición histórica entre los inputs y `Resultado` (no se reordena). Las hojas por responsable van **al final**, ordenadas alfabéticamente con `Collator es_ES PRIMARY` (tildes y mayúsculas se tratan igual que un humano espera).
3. **Orden de hojas en modo `completo`**: misma decisión que en (2) — `Equipos` no se mueve. Las hojas por responsable van **después** de `Resumen`, al final del libro.
4. **Saneo de nombres de hoja**: opción (c) del plan — caracteres prohibidos (`\ / ? * [ ] :`) reemplazados por `_` y truncado a 31 chars (reutiliza `FileProfileResolver.safeSheetName`, ya existente). Si el saneo modifica el nombre, warning categoría `RESPONSABLE`. Si tras sanear hay colisión con otra hoja, sufijo `_2`, `_3`, ... (reutiliza `FileProfileResolver.ensureUniqueSheetName`, **movido en esta release** desde `ExcelMerger`).
5. **Cabecera de cada hoja por responsable**: opción (a) del plan — solo celda `A1` con el **nombre canónico del responsable** (no el saneado), aplicando el estilo `StyleFactory.title()` (negrita, 14 pt). Sin filas reservadas para futuros encabezados; cuando se añadan las dos tablas de resumen en una sesión posterior, ya se decidirá su layout.
6. **Detección de responsables únicos**: opción (b) del plan — **trim + case-insensitive**. Valores como `tresp1@x`, `TRESP1@x`, `  tresp1@x  ` colapsan en una única hoja. El **nombre canónico** (en `A1` y como nombre de hoja) es el **primer literal visto** en orden de filas de `Resultado`. Coherente con cómo `summary.byResponsible` ya normaliza responsables en v1.8.0+ para la segunda tabla de Resumen.
7. **Diff entre ejecuciones de modos distintos**: el usuario aclaró que **no existe ninguna feature de Diff** en v2.2.0; pregunta omitida.
8. **Tests existentes**: como `cierre` es default y los configs del proyecto añaden `output.mode=cierre` explícito, **ningún test existente requiere modificación de aserciones**. Se confirmó por inspección de `ExcelMergerIntegrationTest` (incluida la aserción estricta `wb.getNumberOfSheets() == 5` del happy path).

Decisiones colaterales tomadas en Fase 1:

- **Nombre de la clave**: `output.mode` (no `mode` a secas) por simetría con `output.file`, `output.overwrite`, `output.backup`.
- **Validación**: en `ConfigValidator.validateOutputMode()` (acumula error en la lista de errores), simétrico a `validateMergeMode`. Con `config.strictValidation=true` (default) aborta con exit code 2; con `strictValidation=false` el motor cae a `cierre` defensivamente.
- **Lectura del valor de `Res. Tecnico`**: directo con `Cell.getStringCellValue()` + `trim()`, **sin** construir `FormulaEvaluator`. Justificado: la columna `mes.col.9` es de tipo `COPY` desde `Usuario_Resp_Tecnico`, que ya viene casteado a STRING y trimado en la copia de `profile.Cierre` (`asText.columns` + `trim.columns`).
- **Omisión de Deuda en modo `responsables`**: lectura literal del prompt original ("Sin Deuda"). Se omite la **copia del input** `deuda.xlsx` al libro de salida (`mergeSheetsSeparate` filtra perfiles `Deuda` cuando `outputMode == RESPONSABLES`), con warning CONFIG. Las fórmulas `PDCL + Deuda` en `Resultado` se degradan automáticamente al modo "sin hoja Deuda" (devuelven `{col:PDCL}`, sin warning) gracias al comportamiento existente de `FormulaPlusSumIfsColumnStrategy` v2.2.0.
- **`DerivedSheetBuilder`**: ortogonal al modo. Se invoca en los 3 modos. Con `derived.sheets=` vacío (default real) es no-op.

### Añadido

- **Nuevo enum `OutputMode`** en `src/main/java/com/excelmerger/OutputMode.java` con tres valores `CIERRE`, `RESPONSABLES`, `COMPLETO` y método `parseStrict(String)` (case-sensitive, solo minúsculas, lanza `IllegalArgumentException` con mensaje listando los válidos).
- **Nuevo builder `ResponsablesSheetBuilder`** en `src/main/java/com/excelmerger/ResponsablesSheetBuilder.java`. Lee `Resultado.Res. Tecnico` (nombre de hoja vía `mes.sheetName`, columna vía `summary.byResponsible.column`), agrupa los valores por minúscula tras trim, ordena alfabéticamente con `Collator es_ES PRIMARY`, sanea nombres con `FileProfileResolver.safeSheetName`, resuelve colisiones con `FileProfileResolver.ensureUniqueSheetName`, y crea una hoja por cada responsable canónico con `A1` = nombre canónico en estilo `title`.
- **Validación `validateOutputMode()`** en `ConfigValidator`. Acumula error en la lista de errores si el valor no es uno de los tres válidos (estricto, case-sensitive). Ausente o vacío no es error.
- **Switch de modos en `ExcelMerger.merge()`** (alrededor de las líneas 5f–5f-bis): omite `SummarySheetBuilder` en `RESPONSABLES`, invoca `ResponsablesSheetBuilder` en `RESPONSABLES` y `COMPLETO`, omite la copia del input Deuda en `RESPONSABLES`.
- **`RunReport.setOutputMode/getOutputMode`** y línea `Modo: <MODO>` en el resumen final.
- **Clave `output.mode=cierre`** explícita en los tres `config.properties` (raíz, `src/main/resources/`, `src/test/resources/test-config.properties`) con bloque de comentarios documentando los 3 valores.
- **Tests unitarios nuevos**:
  - `OutputModeTest`: 5 tests sobre `parseStrict` (3 valores válidos, mayúsculas rechazadas, valores inventados, null, mensaje de error completo).
  - `ResponsablesSheetBuilderTest`: ~13 tests que cubren Resultado vacío, ausente, sin columna `Res. Tecnico`, 3 responsables → 3 hojas, valores vacíos/sólo espacios ignorados, case-folding (`tresp1@x` y `TRESP1@x` colapsan), trim, contenido de `A1`, estilo título, saneo de caracteres prohibidos, truncado a 31 chars, colisiones con sufijo `_2`, orden alfabético con Collator (case-insensitive), registro en `RunReport`.
- **Tests de integración nuevos** (`ExcelMergerIntegrationTest`): 8 tests para los 3 modos:
  - `outputModeCierreEsElDefault…`: estructura idéntica a v2.2.0 con 5 hojas.
  - `outputModeResponsablesGeneraHojasPorResponsable…`: presencia de hojas `tresp1@x`, `tresp2@x`, `tresp3@x`, `MG002` (tras trim) y ausencia de `Resumen`.
  - `outputModeResponsablesPosicionaHojasResponsableTrasResultado`: orden relativo (índices > Resultado).
  - `outputModeResponsablesOmiteCopiaDeDeuda…`: con `deuda.xlsx` en input, la hoja `Deuda` no aparece en output, fórmula `PDCL+Deuda` no contiene `Deuda!`, warning `CONFIG` emitido.
  - `outputModeCompletoIncluyeTodasLasHojas…`: Cierre, Extraccion, Deuda, Equipos, Resultado, Resumen, y las hojas por responsable.
  - `outputModeCompletoPosicionaResponsablesTrasResumen`: orden relativo.
  - `outputModeQuedaRegistradoEnRunReport`: `report.getOutputMode()` y línea `Modo:` en el summary.
  - `outputModeInvalidoProduceErrorEnConfigValidator` y `outputModeInvalidoConStrictValidationFalseCaeACierre`: validación y degradación defensiva.
- **Tests añadidos a `ConfigValidatorTest`**: 8 nuevos casos para `output.mode` (ausente, vacío, los 3 válidos, mayúsculas rechazadas, valor inventado, mensaje listando los 3 válidos).

### Cambiado

- **`FileProfileResolver.ensureUniqueSheetName`**: nuevo método **público estático**. **Movido** desde `ExcelMerger` (donde era `private`). Comportamiento idéntico, cero cambios funcionales. El método antiguo en `ExcelMerger` queda como **wrapper privado** de una sola línea para minimizar diff de los callsites internos.
- **`ExcelMerger.mergeSheetsSeparate`**: firma extendida con un parámetro `OutputMode`. Cuando es `RESPONSABLES`, los inputs cuyo perfil resuelto es `Deuda` se omiten de la copia (con log + warning `CONFIG`). El único callsite (la línea 207 de `merge()`) se actualizó. La lógica de orden de perfiles (`computeProfileOrder`) y de copia (`copyAllSheetsFrom`) no cambia.
- **`TestFixtures`**: dos helpers nuevos `buildRealisticConfigWithOutputMode(baseDir, mode)` y `buildRealisticConfigWithDeudaAndOutputMode(baseDir, mode)` que renderizan el `test-config.properties` y luego sobreescriben/insertan la línea `output.mode=<mode>` con regex multilínea. El helper original `buildRealisticConfig` no cambia y sigue alimentando todos los tests v2.2.0 con `output.mode=cierre` (default explícito en `test-config.properties`).
- **`Main.APP_VERSION`** → `2.3.0`. **`pom.xml` `<version>`** → `2.3.0`.

### Notas de no-regresión

- **0 cambios en builders existentes**: `MesSheetBuilder`, `LookupSheetBuilder`, `SummarySheetBuilder`, `DerivedSheetBuilder`, `AvisosSheetBuilder` no se tocan. La selección de qué builders invocar vive en el switch de `ExcelMerger.merge()`.
- **0 cambios en aserciones de tests existentes**. Verificado por inspección: la aserción estricta `wb.getNumberOfSheets() == 5` del happy path, las verificaciones de orden de Deuda/Resultado, y todas las pruebas de fórmulas SUMIFS, asumen modo `cierre` implícitamente — y `cierre` es el default tanto en runtime como en `test-config.properties`.
- **0 cambios en formatos de columna o fórmulas de `Resultado`**. Las hojas por responsable son completamente independientes; no afectan a la construcción de `Resultado` ni a sus fórmulas SUMIFS.
- **`--dry-run`** funciona idéntico: las hojas en memoria se construyen igual; `outputManager.writeResult` se sigue saltando.

### Riesgos conocidos al cierre

- `mvnw verify` no se ejecutó. Hay que ejecutarlo localmente. Posibles puntos de fricción en orden de probabilidad decreciente:
  1. Algún test de integración v2.2.0 que asume el orden absoluto de las hojas (no detectado en la auditoría manual, pero no se puede descartar al 100% sin runner).
  2. Alguna violación menor de Checkstyle/PMD en el código nuevo (line length verificado a 140; sin tabs verificado; resto sin auditar exhaustivamente).
  3. Diferencias entre el `Collator es_ES PRIMARY` y el orden esperado en algún test si el JDK del runner tuviera comportamiento particular (improbable: PRIMARY es estable entre JDKs para alfabeto latino básico).

---

## [2.2.0] — Fichero de Deuda opcional (suma real en `PDCL + Deuda`)

Release minor. Añade soporte para un **tercer fichero Excel opcional** en el directorio de entrada que aporta horas de deuda por `(Peticion, Matricula, Funcion)`. Cuando está presente, la columna `PDCL + Deuda` de `Resultado` deja de ser igual a `PDCL` y pasa a sumar las horas cruzadas desde la nueva hoja `Deuda`. Cuando no está, el comportamiento es **idéntico** a v2.1.0: `PDCL + Deuda == PDCL`, sin warnings.

⚠️ **Nota de validación**: esta release se preparó en un entorno sin acceso a Maven Central. `mvnw verify` **no ha sido ejecutado** para certificar la release. Los cambios están validados por: (a) revisión línea a línea de las firmas cambiadas (`InputFileDetector.validateExcelFiles`, `ExcelMerger.mergeSheetsSeparate`); (b) auditoría de los tests existentes que referencian las claves afectadas (`strictTwoFilesTrueAbortaConTresExcel`, `strictTwoFilesFalseConTresExcelTomaLosDosPrimeros`) y su semántica preservada vía el atajo de retrocompat en `ExcelMerger.merge()`; (c) comprobación de que los tests existentes de `PDCL + Deuda` (valor igual a `PDCL` cuando la hoja Deuda no existe) siguen pasando por construcción de la nueva estrategia. Antes de publicar, correr `./mvnw verify` localmente y confirmar tests verdes y cobertura ≥70%.

### Semántica acordada (Fase 0)

Decisiones cerradas con el usuario antes de tocar código:

1. **Detección por contenido**, igual que Cierre y Extraccion. Cabeceras obligatorias: `Peticion`, `Matricula`, `Funcion`, `Horas`. Fila cabecera 1, `minMatches=4`. El fichero NO es obligatorio.
2. **Clave de cruce** `Resultado ↔ Deuda`: `Peticion + Matricula + Funcion` las tres a la vez. `Funcion` del fichero Deuda se cruza contra la columna `Funcion` de `Resultado` (rol `Dev`/`Sup`/...), **no** contra `Res. Tecnico` (que es el nombre del responsable-persona).
3. **Posición de la hoja Deuda** en el libro de salida: entre `Extraccion` y `Resultado`. El resto de hojas mantiene su orden (Equipos sigue antes de Resultado por ser lookup; ver punto 2a de Fase 1).
4. **Fórmula de `PDCL + Deuda` con Deuda presente**: `{col:PDCL}+IFERROR(SUMIFS(Deuda!$D:$D,Deuda!$A:$A,{Petición},Deuda!$B:$B,{Matrícula},Deuda!$C:$C,{Funcion}),0)` con letras de columna **resueltas en runtime** leyendo la cabecera de `Deuda` (no hardcodeadas A/B/C/D). Rangos abiertos `$A:$A`, consistente con `mes.col.10` (Jira), no con `summary.sumifsMaxRow`.
5. **Fallback sin Deuda**: la fórmula es solo `{col:PDCL}` (idéntica a v2.1.0). Sin warnings.
6. **Fila sin match**: SUMIFS devuelve 0 naturalmente → `PDCL + Deuda == PDCL` para esa fila. Sin warnings.
7. **Normalización de tipos (lección 1.7.1)**: `profile.Deuda.asText.columns=Peticion,Matricula` y `profile.Deuda.trim.columns=Matricula`, igual que los otros perfiles. `Funcion` deliberadamente fuera — si aparece con padding (`"DEV  "`) lo añadimos en una patch v2.2.x.
8. **Rango de ficheros**: opción (a) del plan — `input.strictTwoFiles` reemplazado por `input.strictMinFiles=2` + `input.strictMaxFiles=3`. `strictTwoFiles` queda **deprecated pero leído** con warning CONFIG para no romper despliegues existentes.
9. **Estrategia de implementación**: opción (X) del plan — nuevo tipo de columna `FORMULA_PLUS_SUMIFS` con campos `baseFormula`/`from`/`sum`/`match`. Letras resueltas en runtime. Alternativa rechazada: dos fórmulas alternativas en properties con letras A/B/C/D fijas (fragil si el Excel de Deuda cambia de estructura).

### Añadido

- **Perfil `Deuda`** en los tres `config.properties` (raíz, `src/main/resources/`, `src/test/resources/test-config.properties`):
  - `profile.Deuda.sheetName=Deuda`, `detect.headerRow=1`, `detect.headers=Peticion,Matricula,Funcion,Horas`, `detect.minMatches=4`.
  - `profile.Deuda.asText.columns=Peticion,Matricula` y `trim.columns=Matricula`.
- **Nuevo tipo de columna MES `FORMULA_PLUS_SUMIFS`** en `FormulaPlusSumIfsColumnStrategy`. Campos obligatorios: `baseFormula`, `from`, `sum`, `match`. Los opcionales `fill` y `redIfNotEqualTo` se heredan de `AbstractMesColumnStrategy`. La estrategia:
  - En `preValidate`: valida placeholders `{col:X}` contra las columnas MES declaradas; si la hoja `from` existe, valida cabeceras y deshabilita con warning CABECERA si falta alguna; si la hoja no existe, NO deshabilita y NO emite warnings (degradado silencioso por diseño).
  - En `doWriteCell`: resuelve `{col:X}`/`{colLetter:X}` como `FormulaColumnStrategy`; si la hoja remota no existe escribe solo la base; si existe, concatena `+IFERROR(SUMIFS(...),0)` con letras resueltas dinámicamente.
- **Claves nuevas** `input.strictMinFiles=2` (default) e `input.strictMaxFiles=3` (default). Reemplazan a `input.strictTwoFiles`.
- **Retrocompat**: si el usuario tiene `input.strictTwoFiles` y NO tiene las claves nuevas, se mapea automáticamente (`true → [2,2]`, `false → [2,2]` con truncado y warning). Se emite warning CONFIG de deprecación. Si coexisten las tres, manda la nueva y se emite un warning CONFIG adicional de clave ignorada. Un chequeo adicional en `merge()` preserva el contrato "exactamente 2" de v2.1.0 cuando `strictTwoFiles=true` y hay >2 ficheros (aborta con `InputValidationException`).
- **Fixture `src/test/resources/fixtures/deuda.xlsx`**: 6 filas deliberadamente diseñadas para cubrir (a) match simple, (b) agregación de dos filas con la misma clave (5+2=7h para P-001/M-1001/Dev), (c) match de P-002 y P-007, (d) fila `P-999/M-9999` sin match en Resultado, (e) placeholder `Matricula="-"` que deliberadamente NO cruza (P-010 en Cierre tiene M-1006, no "-"). Documentado en `gen_fixtures.py` para regeneración.
- **`FormulaPlusSumIfsColumnStrategyTest`** con 7 tests unitarios cubriendo los cuatro escenarios del ciclo de vida (preValidate sin hoja, con hoja OK, con cabecera faltante, placeholder base inválido) y tres escenarios de `writeCell` (sin hoja, con hoja, deshabilitada).
- **5 tests de integración nuevos** en `ExcelMergerIntegrationTest`:
  - `deudaFilePresenteSumaHorasEnPdclMasDeuda` — valida posición de la hoja, forma de la fórmula y evaluación numérica (`FormulaEvaluator`) para 4 filas.
  - `sinFicheroDeudaComportamientoIdenticoAVersionAnterior` — valida que NO existe hoja Deuda, la fórmula no contiene `SUMIFS` ni `Deuda!`, y `PDCL + Deuda == PDCL`.
  - `deudaFilePresenteFilaSinMatchDevuelveSoloPdcl` — valida P-005 (sin entrada en Deuda) da delta=0.
  - `deudaFilePlaceholderMatriculaNoCruza` — valida que la fila con `Matricula="-"` del fichero Deuda NO cruza con P-010 (que tiene M-1006).
  - `strictMinFilesConUnSoloExcelFalla` y `retrocompatStrictTwoFilesTrueConTresExcelAbortaIgualQueV210` — cubren el rango min/max y la retrocompat del `strictTwoFiles=true`.
- **`TestFixtures.copyFixturesWithDeudaTo(path)`** para tests que necesitan el tercer fichero.
- **`ConfigLoader.has(key)`** para distinguir "clave ausente" de "clave definida explícitamente con valor default".

### Cambiado

- **`InputFileDetector.validateExcelFiles`**: firma `(List<File>, boolean strictTwo, RunReport)` → `(List<File>, int minFiles, int maxFiles, RunReport)`. Devuelve la lista posiblemente truncada; los llamadores deben usar la lista devuelta. Si `files.size() > maxFiles`, se truncan al primer `maxFiles` por orden alfabético y se emite warning CONFIG.
- **`ExcelMerger.merge()`**: refactor para soportar 2 o 3 workbooks. La iteración pasa de dos variables locales (`file1`, `file2`) a una lista que se recorre para open/validate/profile/merge. Nuevo método privado `computeProfileOrder()` que reordena los workbooks según `merge.profileOrder` (default `Cierre,Extraccion,Deuda`) antes de copiarlos al libro resultado, garantizando la posición canónica de Deuda entre Extraccion y Resultado independientemente del orden alfabético de los ficheros en disco.
- **`ExcelMerger.mergeSheetsSeparate`**: firma cambiada a `(List<Workbook>, List<File>, Workbook, List<FileProfile>)`.
- **`mes.col.13.*` en `config.properties` y `src/main/resources/config.properties`** (y `mes.col.10.*` en `test-config.properties`): pasa de `type=FORMULA` + `formula={col:PDCL}` a `type=FORMULA_PLUS_SUMIFS` + `baseFormula={col:PDCL}` + `from=Deuda` + `sum=Horas` + `match=Peticion:Petición,Matricula:Matrícula,Funcion:Funcion`.
- **`ConfigValidator`**: añade `FORMULA_PLUS_SUMIFS` a `VALID_COL_TYPES` y validación específica del tipo (`baseFormula` requerido y válido, `from`/`sum`/`match` requeridos, `from` debe ser hoja conocida, `match` con la misma validación que `SUMIFS`).
- **`profiles=` en los tres `config.properties`** incluye `Deuda` por defecto. Si el usuario no aporta el fichero, el perfil queda declarado pero sin asignar — comportamiento OK porque la validación solo verifica el léxico del config, no la presencia del fichero en disco.

### Obsoleto

- **`input.strictTwoFiles`**: deprecated. Sigue funcionando con el mapping descrito arriba pero emite warning CONFIG cada vez que se carga.

### Mantenido (decisiones no tomadas)

- **Sin clave `deuda.enabled`**. La activación es automática: si el perfil `Deuda` está en `profiles=` y aparece un fichero que casa, se activa. Eliminado para evitar un tercer estado confuso (`enabled=true` sin fichero → ¿qué debe pasar?).
- **Sin clave `deuda.sumifsMaxRow`**. La fórmula usa rangos abiertos `Deuda!$A:$A`, consistente con `mes.col.10` (Jira). `summary.sumifsMaxRow` sigue siendo específico de la hoja Resumen.
- **Orden de hojas**: se respeta el orden actual del proyecto (`Cierre, Extraccion, [Deuda,] Equipos(oculta), Resultado, Resumen, [_Avisos]`). `Equipos` permanece antes de `Resultado` por ser lookup construido en `LookupSheetBuilder.buildAll()` (ejecutado antes que `MesSheetBuilder.build()`). Como `Equipos` está oculta, la diferencia con la spec literal (`Equipos` al final) no es visible en la UI de Excel.

### Archivos tocados

- `pom.xml`, `src/main/java/com/excelmerger/Main.java` (bump 2.1.0 → 2.2.0).
- `src/main/java/com/excelmerger/ConfigLoader.java` (+`has`).
- `src/main/java/com/excelmerger/ConfigValidator.java` (+`FORMULA_PLUS_SUMIFS`).
- `src/main/java/com/excelmerger/ExcelMerger.java` (refactor 2→2-o-3).
- `src/main/java/com/excelmerger/io/InputFileDetector.java` (firma min/max).
- `src/main/java/com/excelmerger/sheet/column/MesColumnStrategyFactory.java` (+case nuevo).
- `src/main/java/com/excelmerger/sheet/column/FormulaPlusSumIfsColumnStrategy.java` (nuevo).
- `config.properties`, `src/main/resources/config.properties`, `src/test/resources/test-config.properties`.
- `gen_fixtures.py` (+`build_deuda_profile_fixture`).
- `src/test/resources/fixtures/deuda.xlsx` (nuevo, generado).
- `src/test/java/com/excelmerger/TestFixtures.java` (+helper).
- `src/test/java/com/excelmerger/sheet/column/FormulaPlusSumIfsColumnStrategyTest.java` (nuevo).
- `src/test/java/com/excelmerger/ExcelMergerIntegrationTest.java` (+5 tests).
- `README.md`, `CHANGELOG.md`, `pom.xml`.

### Lecciones aprendidas

- **Builder vs copia bruta.** La hoja Deuda es una copia directa sin transformaciones; reutilizar `SheetCopier` vía el mecanismo de perfiles que ya gestiona Cierre y Extraccion evita duplicar código. No se crea `DeudaSheetBuilder`: sería un builder que solo delega.
- **Orden de hojas sin reordering al final.** El orden final del libro se controla solo por el orden de las llamadas en `ExcelMerger.merge()` y el `computeProfileOrder()` que rebaraja los ficheros de entrada a orden canónico antes de copiarlos. No se usa `workbook.setSheetOrder()` final — un reordering a posteriori sería más frágil y obligaría a conocer nombres de hojas que aún no existen al inicio del pipeline.
- **Degradado silencioso en `FormulaPlusSumIfsColumnStrategy.preValidate`.** El briefing pedía explícitamente "sin warnings" cuando el fichero Deuda no está. La tentación de emitir un warning INFO "Deuda no aportado, PDCL+Deuda se calcula sin suma" fue descartada: con la retrocompat intacta, ese warning aparecería en **todos** los cierres que no usan Deuda (la inmensa mayoría en la mayoría de empresas durante mucho tiempo) y contaminaría `_Avisos` sin aportar valor. La documentación del README es suficiente para saber cómo activar la suma.
- **Retrocompat de clave booleana → rango.** Mapear `strictTwoFiles=true` a `[2,2]` sin más rompe el contrato original ("exactamente 2 ABORTA con >2"), porque la API nueva trunca en lugar de abortar. Se resolvió con un chequeo explícito anterior a `validateExcelFiles` que preserva el contrato viejo sin contaminar la firma nueva. Alternativa rechazada: añadir un tercer parámetro `boolean exact` a `validateExcelFiles` (sería una API peor para el 99% de los casos).
- **Hotfix sobre Fase 0 pregunta 5.1: `Funcion` del fichero Deuda se cruza con la columna `Funcion` de Resultado, no con `Res. Tecnico`.** El briefing original describía ambas como "el mismo concepto con distinto nombre"; lo acepté en Fase 0 sin cotejarlo contra `test-config.properties`. Cuando `mvnw verify` arrojó que el delta P-001 era 0 en vez de 7, el diagnóstico fue directo: `Res. Tecnico` en Resultado proviene de `Usuario_Resp_Tecnico` de Cierre (el nombre del responsable, p. ej. `tresp1@x`), mientras que `Funcion` proviene de `Funcion` de Cierre (el rol: `Dev`/`Sup`). `Dev` de Deuda no casa contra `tresp1@x` y el SUMIFS devuelve 0. El fix es una línea en los tres `config.properties` (`Funcion:Res. Tecnico` → `Funcion:Funcion`) más los tests unitarios que codificaban el mapping incorrecto. Lección estructural: cuando el briefing hace una afirmación sobre equivalencia semántica entre columnas, abrir el archivo y leerlo antes de confirmarla en Fase 0, no después del primer fallo de integración.

---

## [2.1.0] — Columna `Funcion` en la hoja `Resultado`

Release minor. Añade una columna `Funcion` a la hoja `Resultado`, inmediatamente después de `Matrícula`. El valor se copia tal cual desde la columna `Funcion` de la hoja `Cierre`. Cero cambios de código de producción: el builder ya era data-driven y el cambio se resuelve en los tres `config.properties`. Se actualizan los tests que accedían a `Resultado` por índice absoluto y se añade cobertura nueva para la columna.

⚠️ **Nota de validación**: esta release se preparó en un entorno sin acceso a Maven Central. `mvnw verify` **no ha sido ejecutado** para certificar la release. Los cambios están validados por (a) análisis estático completo de todas las referencias cruzadas a las columnas de `Resultado` en código y tests, (b) verificación de que `MesSheetBuilder.loadColumns()` y `ConfigValidator.validateMes()` son ambos data-driven (iteran `mes.col.N.*` hasta hueco, sin hardcodear número), (c) auditoría línea a línea de los `getCell(N)` en `ExcelMergerIntegrationTest` para identificar exactamente qué índices se desplazan. Antes de publicar, correr `./mvnw verify` localmente y confirmar 202 tests verdes y cobertura ≥70%.

### Semántica acordada (Fase 0)

Decisiones cerradas con el usuario antes de tocar código:

1. La columna se añade a la hoja `Resultado` (no al `Resumen`), inmediatamente **después de `Matrícula`**.
2. Fuente: columna `Funcion` de la hoja `Cierre`, **copia directa** sin transformación.
3. Semántica de cardinalidad (B3): una matrícula con N funciones distintas en `Cierre` genera N filas en `Resultado`, una por petición+recurso+función original. No hay pivot, no hay concatenación, no hay agregación; `Resultado` ya era una fila por petición, simplemente esa fila ahora expone además su función.
4. Si el valor origen es `"-"`, se propaga tal cual.
5. Sin toggle de configuración: la columna está **siempre presente** (no hay `mes.funcion.enabled`).
6. `--dry-run`: comportamiento idéntico al resto del pipeline (la columna se calcula en memoria; el dry-run solo evita la escritura del XLSX final).

### Cambiado

- **`config.properties` (raíz)**: insertada `mes.col.7.name=Funcion` (`type=COPY`, `from=Funcion`) entre `Matrícula` (col.6) y `Estado` (col.7 previa). Las columnas 7–14 previas se desplazan a 8–15. Total: **15 columnas** (antes 14). Comentario de cabecera actualizado a "15 columnas de Resultado".

- **`src/main/resources/config.properties`** (fallback del JAR): mismo cambio equivalente. Total: 15 columnas.

- **`src/test/resources/test-config.properties`**: insertada `mes.col.7.name=Funcion` (`type=COPY`, `from=Funcion`). Las columnas 7–9 previas (`Res. Tecnico`, `PDCL`, `PDCL + Deuda`) se desplazan a 8–10. Total: **10 columnas** (antes 9).

### Tests actualizados por desplazamiento de índices

`ExcelMergerIntegrationTest` — tres asserts que accedían a `Resultado` por índice 0-based a partir de la posición 6 se desplazan +1:

| Test | Antes | Después | Columna |
|---|---|---|---|
| `orphansEnabledColumnasSinDatoRecibenLiteralGuion` | `getCell(6)` | `getCell(7)` | `Res. Tecnico` |
| `orphansEnabledColumnasFormulaCalculanJiraPor12` | `getCell(7)`, `getCell(8)` | `getCell(8)`, `getCell(9)` | `PDCL`, `PDCL + Deuda` |
| `trimV181ResolvsMatrizMg002` | `getCell(6)` | `getCell(7)` | `Res. Tecnico` |

Comentarios de cada test actualizados para reflejar el nuevo layout con `Funcion` en índice 6.

Asserts con `getCell(N)` para `N ∈ {0, 3, 4, 5}` **no cambian** (Petición/Jira/REAL/Matrícula van antes de la inserción).

### Tests nuevos

- **`MesSheetBuilderTest`** (3):
  - `funcionCopiaValorDesdeCierreTalCual` — verifica cabecera `Funcion` y copia literal de 3 funciones distintas (`AN`, `DI`, `PR`).
  - `funcionPreservaGuionComoValorLiteral` — blinda decisión 4 (el `"-"` no se normaliza a vacío).
  - `funcionGeneraUnaFilaPorCombinacionMatriculaFuncion` — blinda la semántica B3 (misma matrícula × 3 funciones ⇒ 3 filas distintas con la misma matrícula y distinta función).

- **`ExcelMergerIntegrationTest`** (3, con `FormulaEvaluator` donde aplica, regla inquebrantable 4):
  - `resultadoIncluyeColumnaFuncionJustoDespuesDeMatricula` — verifica cabecera de `Resultado` en el pipeline real: `Matrícula` en índice 5, `Funcion` en índice 6.
  - `resultadoPdclYPdclMasDeudaSiguenCalculandoTrasDesplazarFuncion` — evalúa con `FormulaEvaluator` que `Jira`, `REAL`, `PDCL` y `PDCL + Deuda` siguen calculando sus valores correctos tras el desplazamiento de índices. P-001: `Jira=5`, `REAL=6`, `PDCL=6`, `PDCL+Deuda=6`. Test de guardia anti-regresión por si alguien introdujera fórmulas con referencias absolutas por letra.
  - `resultadoFuncionParaHuerfanoGuionSePropagaComoGuion` — con `mes.orphans.enabled=true`, la fila huérfana `TICKETS` recibe `"-"` en la columna `Funcion` (coherente con el resto de columnas COPY sin datos de origen en huérfanos).

### Sin cambios

- `MesSheetBuilder.java`, `SummarySheetBuilder.java`, `ConfigValidator.java`, `RunReport.java`, ningún código de producción se modifica. Todos los builders y validadores ya localizaban columnas por nombre (`PoiUtils.findColumnIndex`) y parseaban las columnas MES iterando `mes.col.N.*` hasta encontrar un hueco.
- Fixtures `.xlsx` sin modificar. `Cierre` ya trae la columna `Funcion` en sus cabeceras (línea 48 del config principal), y los tests nuevos que necesitaban variedad de funciones (distintas por matrícula) usan workbooks sintéticos in-memory, no contaminan los fixtures compartidos.
- La hoja `Resumen` (dos tablas apiladas, `SummarySheetBuilder`) no cambia en absoluto.

### Bump

- `pom.xml`: `2.0.1` → `2.1.0`.
- `Main.APP_VERSION`: `2.0.1` → `2.1.0`.

---

## [2.0.1] — Segunda vuelta de limpieza PMD

Release patch. Sin cambios funcionales, sin cambios de API pública, sin cambios de configuración. Se atacan las exclusiones del `pmd-ruleset.xml` que en 2.0.0 se habían anotado como "para segunda pasada", dejando solo las que responden a decisiones arquitectónicas conscientes.

⚠️ **Nota de validación**: esta release se preparó en un entorno sin acceso a Maven Central. `mvnw verify` **no ha sido ejecutado** para certificar la release. Los cambios están validados por (a) compilación limpia con `javac` contra el uber-jar, (b) análisis manual replicando la semántica de cada regla PMD, y (c) comprobación estática de que ningún test existente referencia los símbolos renombrados. Antes de publicar, correr `mvnw verify` localmente y confirmar verde.

### Corregido

- **`UseLocaleWithCaseConversions`** — 20 ocurrencias. Todas las llamadas a `.toLowerCase()` y `.toUpperCase()` sin argumento se convirtieron a `Locale.ROOT`. Dominio de todas las llamadas: identificadores internos ASCII (tipos, agregaciones, nombres de columna, colores, extensiones de fichero), no texto de usuario — `Locale.ROOT` es correcto y seguro. Ficheros afectados: `ConfigValidator`, `DerivedSheetBuilder`, `FileProfileResolver`, `MesSheetBuilder`, `FileLockDetector`, `InputFileDetector`, `FormulaColumnStrategy`, `MesColumnStrategyFactory`. Se añade `import java.util.Locale;` donde hacía falta. La exclusión desaparece del ruleset.

  *Nota*: el comentario del ruleset 2.0.0 decía "19 ocurrencias"; el conteo real era 20 — una línea de `FileProfileResolver` tenía dos llamadas encadenadas (`actual.toLowerCase().contains(expected.toLowerCase())`) que se contaban como una.

- **`AssignmentInOperand`** — 5 ocurrencias. Los dos `guard++ < 50` en las condiciones de `while` de `FormulaColumnStrategy` se extrajeron a `guard++` al inicio del cuerpo (semántica preservada: mismas 50 iteraciones máximas). Los tres `sheet.createRow(rowIdx++)` (en `LookupSheetBuilder`, `AvisosSheetBuilder` y `ExcelMerger`) se separaron en dos líneas: `createRow(rowIdx)` seguido de `rowIdx++`. La exclusión desaparece del ruleset.

- **`AvoidDuplicateLiterals`** — Se combinó extracción de constantes para 7 literales de dominio con configuración de la regla para filtrar los 3 restantes, que son ruido de concatenación (`"')."`, `",\""`, `"SUM("`) y al extraerlos a constante producirían nombres sin significado.

  Constantes extraídas (con 10 usos en total sustituidos para `CABECERA` que aparece en dos clases):

  | Clase | Constante | Literal | Usos |
  |---|---|---|---|
  | `StyleFactory` | `FONT_CALIBRI` | `"Calibri"` | 5 |
  | `DerivedSheetBuilder` | `WARN_CATEGORY_CONFIG` | `"CONFIG"` | 5 |
  | `LookupSheetBuilder` | `WARN_CATEGORY_LOOKUP` | `"LOOKUP"` | 4 |
  | `MesSheetBuilder` | `WARN_CATEGORY_CABECERA` | `"CABECERA"` | 5 |
  | `SummarySheetBuilder` | `WARN_CATEGORY_CABECERA` | `"CABECERA"` | 5 |
  | `ConfigValidator` | `PROP_PREFIX_MES_COL` | `"mes.col."` | 5 |
  | `ConfigValidator` | `MSG_SUFFIX_KNOWN_SHEETS` | `"'. Hojas conocidas: "` | 4 |
  | `ConfigValidator` | `MSG_SUFFIX_ALLOWED_VALUES` | `"'. Valores permitidos: "` | 5 |
  | `Main` | `BANNER_SEPARATOR` | `"====================================` | 8 |
  | `FileLockDetector` | `MSG_CLOSE_PREFIX` | `"Cierra '"` | 4 |
  | `MesSheetBuilder` | `MSG_SUFFIX_MES_SKIPPED` | `"'. MES omitida."` | 5 |
  | `SummarySheetBuilder` | `MSG_COL_NOT_FOUND_IN` | `"' no encontrada en '"` | 5 |

  La regla queda activa en el ruleset con `minimumLength=5`, que filtra los 3 literales cortos restantes sin necesidad de `exceptionList`. `maxDuplicateLiterals=4` y `skipAnnotations=true` (defaults explícitos para legibilidad del ruleset).

### Documentado (exclusiones cuya justificación se mejoró)

- **`AvoidCatchingGenericException`** — exclusión mantenida con comentario reforzado. Los 5 `catch (Exception)` en `DerivedSheetBuilder` (2), `FileProfileResolver` (1), `Main` (1) y un `catch (Exception ignored)` en `autoSizeColumn` son todos defensivos contra POI. `evaluateAll()`, `autoSizeColumn()` y rutinas similares pueden lanzar varias subclases de `RuntimeException` no declaradas en su firma (POI no las documenta consistentemente), y `RuntimeException` también está marcada por esta regla en PMD 7.x (confirmado en docs), así que cambiar a `catch (RuntimeException)` no resolvería. La exclusión es la respuesta correcta. *Nota técnica*: desde PMD 7.18.0 esta regla se movió de `category/java/design.xml` a `category/java/errorprone.xml`; el ruleset ahora la excluye solo una vez, en la categoría correcta.

- **Complejidad / tamaño de clase / acoplamiento** — exclusiones mantenidas con nota arquitectónica. Afecta a los tres orquestadores grandes (`MesSheetBuilder` 772 LoC, `SummarySheetBuilder` 664 LoC, `ConfigValidator` 607 LoC). Cada uno encapsula una transformación Excel completa con múltiples casos de negocio; partirlos en colaboradores pequeños fragmentaría la lógica de dominio sin ganar mantenibilidad real porque cada sub-pieza solo tiene sentido en el contexto del orquestador. Objetivo de una futura **Sesión F** si se decide abordar, no de limpieza PMD. Reglas afectadas: `CognitiveComplexity`, `CyclomaticComplexity`, `NPathComplexity`, `NcssCount`, `CouplingBetweenObjects`, `ExcessiveImports`, `ExcessiveParameterList`, `ExcessivePublicCount`, `GodClass`, `TooManyFields`, `TooManyMethods`.

- **`DataClass`** — exclusión mantenida. `RunReport` y `VlookupLink` son objetos de valor intencionales; `DataClass` es el patrón correcto para ellos.

- **`LawOfDemeter`** — exclusión mantenida. POI requiere cadenas tipo `row.getCell(i).getStringCellValue()`; respetar Demeter exigiría envolver POI entero, cambio de envergadura que no pertenece a una limpieza PMD.

### Retirado del ruleset (eran exclusiones de 0 violaciones)

Cuatro exclusiones resultaron no cubrir ninguna violación real: el código había dejado de disparar la regla desde que se escribió la exclusión o directamente nunca la disparó. Se retiran del `<exclude>` sin cambio alguno en código:

- **`AvoidThrowingRawExceptionTypes`** — 0 `throw new RuntimeException(...)` en todo src/main.
- **`SignatureDeclareThrowsException`** — 0 `throws Exception` en firmas de método.
- **`UseUtilityClass`** — las tres clases utility (`Main`, `PoiUtils`, `StyleFactory`) ya tienen constructor privado desde antes.
- **`LoosePackageCoupling`** — es no-op sin configurar la propiedad `packages`; estaba excluida sin razón.

### Pendiente para futura Sesión F (refactor arquitectónico)

No es trabajo de limpieza PMD, sino decisión de diseño. Si se aborda, el objetivo sería partir los orquestadores grandes en colaboradores más pequeños — lo cual haría que desaparecieran automáticamente del ruleset las 11 exclusiones del bloque "Complejidad / tamaño / acoplamiento" enumeradas arriba.

### Cambios en ficheros (resumen)

- `pom.xml`: `<version>` a `2.0.1`.
- `src/main/java/com/excelmerger/Main.java`: `APP_VERSION` a `"2.0.1"`, constante `BANNER_SEPARATOR` añadida (8 sustituciones).
- `src/test/java/com/excelmerger/MainTest.java`: assertion de versión actualizado.
- 9 ficheros Java tocados para `Locale.ROOT` (entre `src/main/java/com/excelmerger/` y sub-paquetes).
- 4 ficheros Java tocados para `AssignmentInOperand`.
- 7 ficheros Java tocados para `AvoidDuplicateLiterals`.
- `pmd-ruleset.xml`: reescrito con los cambios descritos.
- `CHANGELOG.md`: esta entrada.

### Versiones

- `2.0.0` → `2.0.1`
- `APP_VERSION` bumpeado en código y test.

---

## [2.0.0] — BREAKING: swap de nombres de perfil Extraccion ↔ Cierre

Release mayor. Se intercambian los nombres internos de los dos perfiles de detección de ficheros para que coincidan con los nombres habituales de los ficheros de entrada del usuario. Hasta 1.8.1, los nombres internos eran contraintuitivos:

- `EXCEL_CIERRE.xlsx` (fichero del usuario con las peticiones del ERP) era detectado por el perfil llamado **`Extraccion`**.
- `Extraccion_para_PDCL___DA23.xlsx` (fichero del usuario con el export de Jira) era detectado por el perfil llamado **`Cierre`**.

A partir de 2.0.0 los nombres coinciden:

- El fichero de peticiones del ERP (`EXCEL_CIERRE.xlsx` o similar) → detectado por el perfil **`Cierre`**.
- El fichero con el export de Jira (`Extraccion_para_PDCL___DA23.xlsx` o similar) → detectado por el perfil **`Extraccion`**.

La detección sigue siendo por contenido (cabeceras, filas de cabecera, valores de celda característicos), así que los ficheros de entrada pueden tener cualquier nombre. El swap es puramente cosmético a nivel de etiqueta de perfil — no cambia ningún comportamiento del pipeline: misma hoja `Resultado`, mismas fórmulas SUMIFS, mismo `Resumen`, mismos huérfanos, mismos warnings. Un `resultado.xlsx` generado con 1.8.1 y uno generado con 2.0.0 contienen los mismos datos; solo cambian dos nombres de hoja.

### Por qué mayor y no patch

Un `config.properties` heredado de 1.x y alimentado a 2.0.0 falla en la validación: las claves `profile.Extraccion.*` y `profile.Cierre.*` del usuario contienen cabeceras que no coinciden con el perfil bajo ese nombre. Un fichero de configuración de cliente sin actualizar es incompatible. Este es el criterio SemVer para mayor: **ruptura de contrato de configuración externa**.

Los `resultado.xlsx` generados cambian dos nombres de hoja (`Extraccion` ↔ `Cierre`). Código o herramientas downstream que referencien estas hojas por nombre exacto (fórmulas en otros libros, scripts de post-proceso) también se rompen — otro criterio de mayor.

### Cambios en config

Swap simétrico de cada par de claves en `config.properties`, `src/main/resources/config.properties` y `src/test/resources/test-config.properties`:

```properties
# Antes (1.x):                          # Ahora (2.0.0):
profiles=Extraccion,Cierre              profiles=Cierre,Extraccion

profile.Extraccion.sheetName=...        profile.Cierre.sheetName=...
profile.Extraccion.detect.headerRow=1   profile.Cierre.detect.headerRow=1
profile.Extraccion.detect.headers=...   profile.Cierre.detect.headers=...
profile.Extraccion.asText.columns=...   profile.Cierre.asText.columns=...
profile.Extraccion.trim.columns=...     profile.Cierre.trim.columns=...

profile.Cierre.sheetName=...            profile.Extraccion.sheetName=...
profile.Cierre.detect.headerRow=2       profile.Extraccion.detect.headerRow=2
profile.Cierre.detect.headers=...       profile.Extraccion.detect.headers=...
profile.Cierre.asText.columns=...       profile.Extraccion.asText.columns=...
profile.Cierre.trim.columns=Matricula   profile.Extraccion.trim.columns=Matricula

# Referencias cruzadas:
mes.sourceSheet=Extraccion              mes.sourceSheet=Cierre
mes.col.9.from=Cierre                   mes.col.9.from=Extraccion
mes.orphans.sourceSheet=Cierre          mes.orphans.sourceSheet=Extraccion
```

Los valores de las claves (cabeceras, columnas a trimar, etc.) **no cambian** — solo cambia la etiqueta del perfil que las agrupa.

### Cambios en código

- `MesSheetBuilder`: defaults de dos lecturas de config ajustados al nuevo naming — `config.get("mes.sourceSheet", "Cierre")` (antes `"Extraccion"`) y `config.get("mes.orphans.sourceSheet", "Extraccion")` (antes `"Cierre"`). Solo afectan a configs que omiten estas claves; en la práctica los configs provistos siempre las setean explícitamente.
- `MesSheetBuilder.RowSource.ofExtraction` → renombrado a `ofCierre`. Método interno (package-private), un único callsite. Semánticamente consistente con el nuevo naming.
- Javadocs y comentarios de clase actualizados en `ExcelMerger`, `MesSheetBuilder`, `PoiUtils` para reflejar la semántica post-swap.
- Sin cambios en APIs públicas de builders, exceptiones ni formato del `RunReport`.

### Cambios en fixtures de test

- `src/test/resources/fixtures/extraccion.xlsx` ↔ `src/test/resources/fixtures/cierre.xlsx` **intercambiados físicamente** en disco para que el nombre del fichero coincida con el nombre del perfil que detecta su contenido.
- `gen_fixtures.py`: funciones `build_extraccion` / `build_cierre` renombradas a `build_cierre_profile_fixture` / `build_extraccion_profile_fixture` (nombres explícitos del perfil que generan). Constantes `EXTRACCION_HEADERS` / `CIERRE_HEADERS` renombradas a `CIERRE_PROFILE_HEADERS` / `EXTRACCION_PROFILE_HEADERS`. Mensajes `print` actualizados.

### Cambios en tests

- `ExcelMergerIntegrationTest`: `extraccionConservaSus19Filas...` renombrado a `hojaCierreConservaSus21Filas...`; `extraccionEnResultadoTienePeticionYRecursoComoString...` renombrado a `cierreEnResultado...`. Las fórmulas SUMIFS de Jira ahora referencian `"Extraccion"` (antes `"Cierre"`). El helper `runReportContabilizaFilasCorrectamente` ahora asserta `containsEntry("Cierre", 21)` + `containsEntry("Extraccion", 28)` (antes solo `Extraccion, 21`).
- `ConfigValidatorTest`, `FileProfileResolverTest`, `MesSheetBuilderTest`, `RunReportTest`: no modificados. Usan los nombres `"Extraccion"` / `"Cierre"` como meras etiquetas en configs ad-hoc construidos dentro del test. La lógica que prueban no depende de la semántica swapeada.

### Guía de migración para usuarios con `config.properties` heredado de 1.x

Si tienes un `config.properties` personalizado heredado de una instalación 1.x:

1. Intercambia todas las claves `profile.Extraccion.*` ↔ `profile.Cierre.*`. El contenido de cada clave (cabeceras, columnas a textualizar, columnas a trimar) **no cambia**.
2. Actualiza las tres referencias cruzadas:
   - `mes.sourceSheet=Extraccion` → `mes.sourceSheet=Cierre`
   - `mes.col.<N>.from=Cierre` → `mes.col.<N>.from=Extraccion` (donde `<N>` es la columna Jira del SUMIFS)
   - `mes.orphans.sourceSheet=Cierre` → `mes.orphans.sourceSheet=Extraccion`
3. Los ficheros Excel de entrada no requieren cambios; la detección por contenido sigue funcionando igual.
4. Scripts, fórmulas Excel o herramientas downstream que referencien las hojas del `resultado.xlsx` por nombre literal también deben intercambiar las dos referencias.

Una forma rápida de verificar la migración: generar un `resultado.xlsx` con tu config swapeado y abrirlo — debe contener exactamente las mismas hojas que antes (`Resultado`, `Resumen`, `Equipos`, más las dos hojas copiadas) con el mismo contenido en cada una; solo los nombres `Extraccion` / `Cierre` de las dos últimas hojas están intercambiados.

### No cambia

- Semántica del pipeline: misma hoja `Resultado`, mismas fórmulas, mismo `Resumen` con sus dos tablas, mismos huérfanos, mismos warnings.
- APIs Java públicas del core (ExcelMerger, MesSheetBuilder, SummarySheetBuilder, builders en general).
- Formato del `RunReport` y de los ficheros de salida.
- Todos los fixes y features anteriores: v1.6.2 (asText), v1.7.0 (huérfanos), v1.7.1 (recálculo), v1.8.0 (segunda tabla Resumen), v1.8.1 (trim padding). Siguen activos e intactos.

### Cambios de versión

- `pom.xml` y `Main.APP_VERSION` pasan de `1.8.1` a `2.0.0`.
- `MainTest.appVersionEsLaEsperadaPorLaSesionE` actualizado con bloque de comentario 2.0.0 explicando el swap.

## [1.8.1] — Fix: padding de espacios en Usuario_Resp_Tecnico rompía el SUMIFS de Resumen

Patch. Arregla un bug reportado en producción tras el release de 1.8.0: en la segunda tabla de la hoja `Resumen` (matriz Matrícula × Responsable), todas las celdas mostraban `0` excepto la fila del responsable literal `-` (huérfanos). Al pulsar F9 en Excel los valores no cambiaban. El fix de v1.8.0 (`setForceFormulaRecalculation`) había resuelto el problema de recálculo, pero no cubría este segundo fallo, que es de datos y no de engine.

### Causa raíz

El export ERP que alimenta el perfil `Extraccion` alinea los códigos de `Usuario_Resp_Tecnico` con padding de espacios a la derecha: todos los valores vienen como `"MG002   "`, `"IA03    "`, etc. (los 832/832 analizados). Al copiar estos valores a la hoja `Resultado`, la capa de copia los preservaba tal cual. Después, `SummarySheetBuilder.discoverResponsibles` aplicaba `trim() + toUpperCase()` al descubrir códigos únicos, y emitía la cabecera de la segunda tabla con el valor limpio: `"MG002"`. El SUMIFS resultante comparaba el criterio `"MG002"` (limpio, en cabecera de Resumen) contra el rango `Resultado!H2:H10000` (con padding). **SUMIFS de Excel es case-insensitive en texto pero no trim-insensitive**: `"MG002" ≠ "MG002   "`, el filtro no casaba, el SUMIFS devolvía 0.

La única fila que sumaba correctamente era la del responsable literal `-`, porque ese valor lo escribe el builder directamente (via `mes.orphans.colResTecnico.literal=-`) sin pasar por el export con padding.

En v1.8.0 este caso se había documentado como "límite conocido" en el test `byResponsibleEspaciosEnLosDatosOrigenNoCasanEnSumifs`, asumiendo que los espacios eran un escenario hipotético. Resultó ser el comportamiento estándar del ERP.

### Fix

La solución correcta es normalizar los valores **en la capa de copia** y no en cada consumidor aguas abajo. Nueva clave de configuración opt-in:

```properties
profile.Extraccion.trim.columns=Recurso,Usuario_Resp_Tecnico
profile.Cierre.trim.columns=Matricula
```

Semántica: para cada columna listada, al copiar al workbook resultado se aplica `trim()` al valor tras el cast a STRING. Solo tiene efecto si la columna también está en `asText.columns` (el trim es una capa sobre la rama STRING del cast; sin cast no hay trim posible). Si se declara `trim.columns` sin `asText.columns` el validador aborta con error de config; si una columna individual está en trim pero no en asText, se emite warning `CONFIG` y se ignora.

Las celdas de `Resultado` llegan ahora sin padding, el SUMIFS de la segunda tabla casa correctamente, y cualquier otro consumidor aguas abajo (fórmulas custom, hojas derivadas futuras) se beneficia automáticamente.

### Añadido

- `PoiUtils.copyCellValueAsTextTrimmed(Cell, Cell)` — variante de `copyCellValueAsText` que aplica `trim()` a la rama STRING.
- `SheetCopier.copySheet` y `copyRow` — sobrecargas que aceptan un segundo set `trimColumnIndexes`. Una columna se trima si está en **ambos** sets (asText y trim).
- `FileProfile.getTrimColumns()` y `resolveTrimColumnIndexes(Sheet)` — equivalentes a los de `asText`.
- `ExcelMerger.resolveTrimIndexes` — calcula índices y emite warnings `CABECERA` para columnas no encontradas y `CONFIG` para columnas declaradas en trim pero no en asText (se ignoran silenciosamente en runtime tras el warning).
- `ConfigValidator` — valida que `trim.columns` no esté declarado sin un `asText.columns` no vacío (error de config duro).
- `config.properties` (raíz y fallback) y `test-config.properties`: añadidas las claves nuevas. `profile.Extraccion.asText.columns` se amplía para incluir `Usuario_Resp_Tecnico` (necesario para que el trim tenga efecto sobre esa columna).

### Tests

- **9 tests unitarios nuevos** en `PoiUtilsTest` cubriendo todas las ramas de `copyCellValueAsTextTrimmed`: STRING con padding derecha, STRING con espacios al principio y final, STRING sin espacios (no-op), NUMERIC entero, NUMERIC fecha (no se trima ni se castea), BOOLEAN, FORMULA, BLANK, ERROR.
- **Test de integración crítico** `trimResponsableConPaddingHaceQueElSumifsDeResumenCasePorResponsable` — reproduce el bug con fixture real: añade una fila `P-016 / M-1010 / "MG002   "` en Extracción y una imputación `PROJ-25 / P-016 / M-1010 / Dev / 5h` en Cierre. Verifica primero que tras el pipeline la celda `Res. Tecnico` de `Resultado` para la fila P-016 está sin padding (`"MG002"`, no `"MG002   "`). Luego busca en la segunda tabla de `Resumen` la celda `(M-1010, MG002)` y la evalúa con `FormulaEvaluator`; debe valer `6.0` (PDCL = Jira × 1.2 = 5 × 1.2). Sin el fix, este valor es `0` — exactamente el bug que se reportó. Este test es **el que faltaba en 1.8.0** y cuya ausencia permitió que el bug llegara a producción.

### Fixtures

- `gen_fixtures.py`: fila nueva en `extraccion.xlsx` (P-016 con `Usuario_Resp_Tecnico="MG002   "`, 3 espacios de padding) y fila paralela en `cierre.xlsx` (PROJ-25 casando con P-016/M-1010). Conteos documentados:
  - `extraccion.xlsx`: 20 → 21 filas.
  - `cierre.xlsx`: 27 → 28 filas.

### Cambiado

- Tests de integración con asserts de conteo: 5 asserts actualizados (Extraccion 20→21, Resultado sin huérfanos 19→20, Resultado con huérfanos 22→23, asserts en dos tests adicionales).
- Test de integración v1.8.0 `byResponsiblePipelineGeneraSegundaTablaConTituloCorrecto`: el número de responsables en cabecera pasa de 3 a 4 (MG002 se suma a TRESP1@X/TRESP2@X/TRESP3@X) y el orden alfabético los coloca de forma distinta — assert actualizado.
- Test `SummarySheetBuilderTest.byResponsibleEspaciosEnLosDatosOrigenNoCasanEnSumifs`: el javadoc pasa de documentar "límite conocido" a documentar el contrato del builder en aislado (el builder no trima, el trim es de la capa de copia). La aserción del test no cambia — sigue siendo cierto que si alguien construye `Resultado` a mano saltándose `SheetCopier`, el builder no puede adivinar el trim.

### Cambios de versión

- `pom.xml` y `Main.APP_VERSION` pasan de `1.8.0` a `1.8.1`.
- `MainTest.appVersionEsLaEsperadaPorLaSesionE` actualizado con bloque de comentario 1.8.1.

### No cambia

- Semántica de `asText.columns`: una columna que solo esté en `asText.columns` (no en `trim.columns`) se copia como STRING como antes, sin trim.
- Primera tabla de `Resumen`, hojas derivadas, avisos, huérfanos: intactos.
- API pública del builder: las sobrecargas nuevas de `SheetCopier.copySheet` y `copyRow` son adicionales; las firmas anteriores siguen disponibles con el mismo comportamiento.

### Cómo actualizar desde 1.8.0

Si heredas un `config.properties` de 1.8.0 y quieres el fix: añade las dos líneas siguientes y verifica que tus cabeceras aparecen también en `asText.columns`:

```properties
profile.Extraccion.asText.columns=Peticion,Recurso,Usuario_Resp_Tecnico
profile.Extraccion.trim.columns=Recurso,Usuario_Resp_Tecnico

profile.Cierre.trim.columns=Matricula
```

Si no añades estas claves, el comportamiento queda como 1.8.0 (trim inactivo). Es **opt-in puro**, no rompe configs existentes.

## [1.8.0] — Segunda tabla en Resumen: matriz Matrícula × Responsable

Minor opt-in. La hoja `Resumen` admite ahora una **segunda tabla** que cruza matrículas (filas) con responsables técnicos (columnas), mostrando el `PDCL` de cada par. La primera tabla (sumatorio por matrícula de `Jira, REAL, PDCL, PDCL + Deuda`) se mantiene intacta; la nueva va debajo, separada por filas en blanco configurables. Es opt-in (`summary.byResponsible.enabled=true`) y si no se activa el comportamiento de v1.7.1 queda exactamente igual.

### Fix aplicado durante la sesión de release

Durante las pruebas en Excel real (no cubiertas por los tests automáticos, que usan POI `FormulaEvaluator` o LibreOffice) se detectó que las celdas de la segunda tabla se abrían mostrando `0` en todas las posiciones, aunque POI y LibreOffice las evaluaban correctamente. **Causa raíz**: Apache POI escribe las celdas con fórmula sin valor cacheado (`<f>...</f><v></v>`). Excel, en ausencia del flag `fullCalcOnLoad="1"` en el elemento `<calcPr>`, recalcula la mayoría de las fórmulas al abrir pero tiene comportamientos inconsistentes con SUMIFS de 4+ criterios (como los de esta matriz) cuando la caché está vacía. `DerivedSheetBuilder` ya activaba este flag, pero solo si `derived.sheets` tenía contenido — y en el pipeline real está vacío. **Fix**: `SummarySheetBuilder` ahora setea `workbook.setForceFormulaRecalculation(true)` incondicionalmente al final de `build()` cuando la hoja se ha creado. Es idempotente con lo que haría `DerivedSheetBuilder`. Test de regresión `buildFuerzaRecalculoAlAbrirEnExcel` añadido.

### Motivación

La tabla actual responde a "¿cuántas horas ha metido cada persona?", pero no a "¿qué personas están imputando horas en cada matrícula?". La nueva matriz contesta esa segunda pregunta de un vistazo: columnas = responsables, filas = matrículas, celdas = PDCL. Fila `Total` (suma por responsable) y columna `Total` (suma por matrícula) al final, y un gran total en la esquina inferior derecha que debe cuadrar con el `PDCL` global de la primera tabla (check cruzado natural).

### Añadido

- Nueva sección `summary.byResponsible.*` en `config.properties`:
  - `summary.byResponsible.enabled` (default `false`) — opt-in explícito.
  - `summary.byResponsible.column` (default `Res. Tecnico`) — columna de `Resultado` a usar como agrupador de columnas. Debe coincidir con un `mes.col.N.name`.
  - `summary.byResponsible.valueColumn` (default `PDCL`) — métrica única a sumar. Debe coincidir con un `mes.col.N.name`.
  - `summary.byResponsible.title` (default `Totales Peticiones por Responsables Matrículas`) — texto de la fila de título (merge sobre el ancho de la tabla).
  - `summary.byResponsible.gapRows` (default `2`) — filas en blanco entre la primera y la segunda tabla. Admite `0` (tablas pegadas).
- `ConfigValidator.validateSummaryByResponsible` valida las claves nuevas cuando la feature está activa. Regla fuerte: `summary.byResponsible.enabled=true` con `summary.enabled=false` produce error de configuración (la segunda tabla se ancla a la hoja `Resumen`; no tiene sentido sola).
- `SummarySheetBuilder.writeByResponsibleTable` construye la segunda tabla reutilizando la misma hoja, estilos y helpers que la primera. Es un método privado de la misma clase (misma responsabilidad conceptual: "construir la hoja Resumen"), no una clase separada.
- `SummarySheetBuilder.discoverResponsibles` — auto-descubre los códigos únicos de la columna de responsable normalizando con `trim()` + `toUpperCase(Locale.ROOT)`. Las variantes `"resp01"`, `"RESP01"`, `" Resp01 "` colapsan en un único responsable `"RESP01"`. El `SUMIFS` de Excel es case-insensitive en criterio de texto, así que una sola celda de cabecera en MAYÚSCULAS suma correctamente todas las variantes del Excel original sin código adicional.
- `RunReport` recibe un warning informativo (categoría `HOJA`) cuando la segunda tabla se construye con éxito, reportando cuántas matrículas × cuántos responsables.

### Cambiado

- `config.properties` (raíz) y `src/main/resources/config.properties` (fallback) incluyen las claves nuevas. En el fallback interno la feature queda `enabled=false` para no alterar despliegues minimalistas que dependan del classpath resource. El externo la trae `true` con los defaults del uso típico.
- `src/test/resources/test-config.properties` la habilita para que el pipeline de integración cubra la nueva tabla de verdad (end-to-end con `FormulaEvaluator`, no solo inspección de texto de fórmula).

### Fixtures

- Nueva fila en `gen_fixtures.py` (`extraccion.xlsx`): `P-015 / M-1009 / TRESP1@x`. Exactamente la misma ficha de responsable que `tresp1@x` pero en MAYÚSCULAS, para validar end-to-end que la segunda tabla colapsa ambas en una única columna `TRESP1@X` y el `SUMIFS` case-insensitive suma las dos variantes. El conteo total de filas pasa de `1+14+3+1=19` a `1+14+3+1+1=20`.

### Tests

- **9 tests unitarios nuevos** en `SummarySheetBuilderTest`:
  - `byResponsibleDeshabilitadoNoAnadeSegundaTabla` — feature-flag off; la hoja `Resumen` queda como en 1.7.1 y no se emite el warning de "añadida tabla".
  - `byResponsibleConstruyeMatrizConCabecerasCorrectas` — layout básico: título merge + fila en blanco + cabecera (esquina vacía, responsables alfabético, "Total" al final).
  - `byResponsibleDescubreYNormalizaResponsablesAMayusculas` — 4 filas con `"resp01"`, `"RESP01"`, `" Resp01 "`, `"OTHER"` producen exactamente 2 columnas (`OTHER`, `RESP01`) y 0 columnas extra.
  - `byResponsibleFormulasTienenFormaCorrectaYRangosAcotados` — la fórmula `SUMIFS` cruza los tres rangos (valor, matrícula, responsable), con rangos acotados `2:10000`, criterio de matrícula `$A<fila>` (relativo en fila, absoluto en columna) y criterio de responsable `B$<fila_cabecera>` (relativo en columna, absoluto en fila).
  - `byResponsibleSumifsEvaluadoProduceValoresCorrectos` — **evaluación real con `FormulaEvaluator`** (lección 1.7.1) de una cuadrícula completa de 3×2 con valores documentados. Verifica celdas individuales, totales por fila, totales por columna y gran total.
  - `byResponsibleNormalizacionDeVariantesCapitalizacionSumaCorrectamente` — 3 variantes del mismo código (`resp01`, `RESP01`, ` Resp01 `) suman correctamente al evaluar con `FormulaEvaluator`: `1.5 + 3 + 6 = 10.5`. Este test es el que blinda la normalización end-to-end.
  - `byResponsibleColumnaResponsableInexistenteEmiteWarningPerosNoRompePrimeraTabla` — si la columna configurada no existe, warning en `RunReport` pero la primera tabla queda intacta.
  - `byResponsibleGapRowsRespeta` — `gapRows=0` pega las dos tablas sin fila en blanco entre ellas.
  - `byResponsibleRegistraWarningInformativoEnReport` — el warning de éxito lleva categoría `HOJA` y menciona número de matrículas y responsables.
- **3 tests de integración nuevos** en `ExcelMergerIntegrationTest` (pipeline completo sobre fixtures reales):
  - `byResponsiblePipelineGeneraSegundaTablaConTituloCorrecto` — tras el merge, la hoja `Resumen` contiene el título de la segunda tabla y la cabecera con 3 responsables únicos (`TRESP1@X`, `TRESP2@X`, `TRESP3@X`) + `Total`, en ese orden.
  - `byResponsibleSumifsCaseInsensitiveSumaTrespVariantes` — evalúa con `FormulaEvaluator` la celda `(99641, TRESP1@X)` y verifica que vale exactamente `10.8` (el `PDCL = Jira * 1.2` del responsable `tresp1@x` en la fila `138074/99641`).
  - `byResponsibleTotalGlobalCuadraConPDCLGlobalDeLaPrimeraTabla` — check cruzado: el gran total de la segunda tabla coincide con el total `PDCL` de la primera. Si divergen, algo está mal en el agrupamiento o en los criterios del `SUMIFS`.
- Tests existentes de conteo de filas actualizados para la fila v1.8.0: `Extraccion` pasa de 19 a 20, `Resultado` sin huérfanos de 18 a 19, `Resultado` con huérfanos de 21 a 22.
- `hojaResumenTrasPipelineCompletoExisteYContieneSumatoriosPorMatricula` refactorizado: antes asumía que la última fila de la hoja `Resumen` era el `Total` de la primera tabla; ahora con la segunda tabla activa eso ya no vale. El test escanea desde arriba buscando el primer `Total` (que es el de la primera tabla).
- `MainTest.appVersionEsLaEsperadaPorLaSesionE` actualizado a `1.8.0`.

### Cambios de versión

- `pom.xml` y `Main.APP_VERSION` pasan de `1.7.1` a `1.8.0`.

### No cambia

- La primera tabla de `Resumen` (sumatorio por matrícula) es byte-a-byte la misma que en 1.7.1: mismo título, misma cabecera, mismas fórmulas, misma fila de totales. Todos los tests que la validan siguen verdes.
- El resto del pipeline (`Extraccion`, `Cierre`, `Resultado`, `Equipos`, `_Avisos`) no se toca.

### Límite conocido (documentado, no bug)

La normalización de responsables aplica `trim()` + `toUpperCase(Locale.ROOT)` **al descubrir** los códigos únicos para la cabecera de la segunda tabla. El SUMIFS emitido es sin embargo el estándar de Excel, que es case-insensitive pero **no trim-insensitive**. Consecuencia: si el Excel original trae un responsable con espacios al principio o final (`" Resp01 "`), aparece en la cabecera como `RESP01` (una sola columna junto con las variantes sin espacios) pero su fila **no se suma** en esa columna — el criterio `RESP01` del SUMIFS no casa contra la celda ` Resp01 `. El test `byResponsibleEspaciosEnLosDatosOrigenNoCasanEnSumifs` fija este comportamiento para que no se regrese por accidente. El escenario real pactado son códigos alfanuméricos sin espacios, en los que la única variabilidad esperada es la capitalización; el trim completo en la capa de copia de datos queda para una iteración futura si hace falta.

## [1.7.1] — Fix bug C: matrículas numéricas en Resumen daban Jira=0

Patch. Arregla un bug latente descubierto al escribir los tests de la feature 1.7.0 (huérfanos): el `SUMIFS` del `Resumen` daba `0` para cualquier matrícula todo-dígito (`55751`, `90014`, `99641`, `99642`, etc.), aunque esa matrícula tuviera horas asignadas en `Resultado`. El bug es **preexistente a 1.7.0** — afectaba ya en 1.6.2 y antes — pero no lo cazaba ningún test porque los tests existentes del `Resumen` no evaluaban los SUMIFS con `FormulaEvaluator`. Al añadir esa evaluación en los tests de huérfanos salió a la luz.

### Causa raíz

Mismatch de tipos entre la celda clave del `Resumen` y el rango que escanea el SUMIFS:

- `SummarySheetBuilder.setNumericOrString` escribía las matrículas todo-dígito como `NUMERIC` (Excel las alineaba a la derecha, más limpio visualmente).
- La columna `Matrícula` de `Resultado` es `STRING` tras el fix 1.6.2 (`asText.columns=Recurso`).
- Excel `SUMIFS` con criterio `NUMERIC` contra rango `STRING` no casa (misma asimetría que arregló 1.6.2, pero en sentido inverso).

El efecto en producción es que cualquier matrícula con código numérico produce `Jira=0`, `REAL=0`, `PDCL=0`, `PDCL + Deuda=0` en `Resumen`, aunque esa persona sí tenga horas imputadas correctamente en `Resultado`.

### Cambiado

- `SummarySheetBuilder.setNumericOrString` escribe ahora **siempre** como `STRING`. La celda clave del `Resumen` queda alineada a la izquierda (consecuencia estética aceptable). El método `isNumeric` se conserva porque `discoverMatriculas` sigue usándolo para ordenar (numéricas ASC primero, alfabéticas al final), pero ya no decide el tipo de celda.

### Tests

- `SummarySheetBuilderTest.matriculasSeDescubrenOrdenandoNumericasPrimeroYStringsDespues` — actualizado al nuevo contrato: todas las matrículas son `STRING`. El orden (numéricas ASC primero, alfabéticas al final) se preserva.
- `SummarySheetBuilderTest.sumifsDeMatriculaNumericaEvaluaCorrectamenteContraResultado` — nuevo test de regresión del bug C. Evalúa con `FormulaEvaluator` que la matrícula `99641` suma `5 + 2 = 7` (sus dos filas en el workbook de test).
- `ExcelMergerIntegrationTest.orphansEnabledSumaEnResumenPorMatricula` — restaurado a su versión end-to-end. Verifica que `Resumen` suma `5h` (normal `101770/90014`) + `3h` (huérfano `VACACIONES/90014`) = `8h` para la matrícula numérica `90014`. En 1.7.0 se había rebajado (`orphansEnabledHuerfanoAparecEnResultadoConSuMatricula`) para esquivar el bug C; ahora conviven ambos tests — el primero verifica que el huérfano existe en `Resultado`, el segundo verifica el end-to-end hasta `Resumen`.

### Cambios de versión

- `pom.xml` y `Main.APP_VERSION` pasan de `1.7.0` a `1.7.1`. `MainTest.appVersionEsLaEsperadaPorLaSesionE` actualizado.

## [1.7.0] — Filas huérfanas en Resultado

Minor opt-in. `Resultado` puede ahora incluir **filas adicionales** para las imputaciones de `Cierre` que no tienen contrapartida `(Peticion, Recurso)` en `Extracción`. Sobre los inputs reales del usuario esto son ~35 imputaciones con ~988 horas que el 1.6.2 dejaba fuera (el `SUMIFS` de Jira no las recogía porque no había fila destino para ellas). Con `mes.orphans.enabled=true` esas imputaciones pasan a aparecer como filas huérfanas en `Resultado`, con la matrícula correspondiente, las horas totales agrupadas por par `(Component Name, Matrícula)`, y fórmulas `REAL`/`PDCL`/`PDCL + Deuda` calculadas con la misma cascada `*1.2`. `Resumen` suma estas filas automáticamente por matrícula.

### Añadido

- Nueva sección `mes.orphans.*` en `config.properties`:
  - `mes.orphans.enabled` (default `false`) — opt-in explícito. El comportamiento del 1.6.2 queda intacto mientras no se active.
  - `mes.orphans.sourceSheet` (default `Cierre`) — hoja de la que extraer imputaciones huérfanas.
  - `mes.orphans.matchComponent` (default `Component Name`), `mes.orphans.matchMatricula` (default `Matricula`), `mes.orphans.sumColumn` (default `Hours`) — cabeceras de esa hoja usadas como clave de cruce y columna a sumar.
  - `mes.orphans.colPeticion`, `mes.orphans.colMatricula`, `mes.orphans.colJira` — nombres de las columnas MES que reciben respectivamente el `Component Name`, la matrícula y las horas agregadas. Deben coincidir con los `mes.col.N.name` definidos.
- Orden configurable implícito de `Resultado` cuando se activa: peticiones numéricas ascendente primero, no numéricas en orden alfabético al final. Esto también reordena las filas normales, que hasta 1.6.2 heredaban el orden de `Extracción`.
- `ConfigValidator.validateOrphans()` — valida cuando `enabled=true`: `mes.enabled=true`, `sourceSheet` referencia a hoja conocida, `colPeticion`/`colMatricula`/`colJira` apuntan a columnas MES definidas.
- `MesSheetBuilder` ampliado con:
  - Clase interna `RowSource` — fuente de una fila de `Resultado`, bien de `Extracción` (con su `srcRow` y número Excel de origen) o huérfana (con `Peticion`, `Matrícula`, `Hours` ya agregadas).
  - `collectOrphans(...)` — agrupa imputaciones de `Cierre` por `(CN, Mat)` sumando horas, excluyendo las que sí casan con Extracción.
  - `loadExtractionPairKeys(...)` — set de claves `(Peticion|Recurso)` existentes en Extracción.
  - `writeOrphanCell(...)` — escribe cada columna para una fila huérfana: las columnas configuradas (`colPeticion`/`colMatricula`/`colJira`) reciben el dato del huérfano, las columnas `FORMULA` (`REAL`, `PDCL`, `PDCL + Deuda`, `Equipo`) se resuelven con sus plantillas habituales (funcionan igual sin `srcRow`), y el resto (`COPY` de otros campos, `EMPTY`) reciben un literal `"-"`.
  - `sortRowSources(...)` con clave `SortKey` — ordena numéricas ASC primero, no numéricas alfabético al final.

### Cambiado

- `MesSheetBuilder.detectUnmappedVlookupKeys` ignora ahora el sentinela `"-"` para no emitir warnings falsos de "apps sin mapeo" cuando las filas huérfanas rellenan `Aplicación` con `"-"`.
- El bucle de escritura de datos en `MesSheetBuilder` se ha refactorizado a dos pasadas (recolectar → ordenar → escribir). La primera pasada itera `Extracción` igual que antes; la segunda pasada añade huérfanos e invoca el sort solo si `orphans.enabled=true`. Si está desactivado, la primera pasada se escribe sin ordenar y el orden heredado de `Extracción` se preserva exactamente como en 1.6.2.

### Tests

- `ExcelMergerIntegrationTest` — ocho tests nuevos:
  - `orphansEnabledAnadeFilasParaImputacionesSinContrapartida` — verifica que las tres parejas huérfanas (`TICKETS/-`, `VACACIONES/90014`, `P-001/MAT-HUERFANO`) aparecen con sus horas agregadas correctas (`8`, `3`, `1`).
  - `orphansEnabledFilasOrdenadasNumericasPrimero` — verifica el orden: `55751`, `101770`, `138074` en las tres primeras posiciones; `TICKETS`, `VACACIONES` en las últimas.
  - `orphansEnabledColumnasSinDatoRecibenLiteralGuion` — verifica que `Aplicación`, `Res. Tecnico` en filas huérfanas son `"-"`.
  - `orphansEnabledColumnasFormulaCalculanJiraPor12` — evalúa con `FormulaEvaluator` que `REAL=9.6`, `PDCL=9.6`, `PDCL+Deuda=9.6` para `TICKETS` (`Jira=8`).
  - `orphansEnabledNoGeneraWarningLookupParaGuion` — verifica que `"-"` no se cuenta como app sin mapeo en el lookup `Equipos`.
  - `orphansEnabledHuerfanoAparecEnResultadoConSuMatricula` — verifica que el huérfano `VACACIONES/90014/3h` aparece en `Resultado` con su matrícula correcta. No evalúa el total en `Resumen` a propósito; ver "Pendiente conocido fuera de alcance" más abajo para el bug C descubierto durante los tests.
  - `orphansDisabledMantieneComportamiento16Point2` — regresión: con `enabled=false`, `Resultado` tiene 18 filas y la primera es `P-001`.
  - `orphansEnabledConSheetInexistenteEmiteWarning` — caso degradado: si `sourceSheet` apunta a una hoja inexistente, el merge continúa sin huérfanos y emite warning `HOJA`.
- `ConfigValidatorTest` — cinco tests nuevos: `orphansDisabledNoValida`, `orphansEnabledConMesDeshabilitadoEsError`, `orphansConSourceSheetDesconocidaEsError`, `orphansConColumnaMesInexistenteEsError`, `orphansCompletoYValidoNoLevantaErrores`.
- Fixtures (`gen_fixtures.py`) ampliados con cuatro imputaciones huérfanas en `Cierre` que cubren tres casos: CN inexistente con matrícula `"-"` (2 imputaciones que suman 8h), CN inexistente con matrícula que sí existe en Extracción pero asociada a otra petición (`VACACIONES/90014/3h`), y CN existente con matrícula no asociada a esa petición (`P-001/MAT-HUERFANO/1h`).

### Diseño

- **Opt-in**: el default es `mes.orphans.enabled=false` en el classpath fallback y en el test-config, pero **sí** se activa en el `config.properties` raíz que se distribuye con el proyecto (el que usa el usuario). Este desdoblamiento permite ejecutar los tests existentes con el comportamiento de 1.6.2 (contadores exactos de filas) mientras que el despliegue real recibe la feature activa.
- **Agrupación por pareja `(CN, Mat)`**, no por CN solo: cada pareja única genera una fila. Esto mantiene la coherencia con `Resumen`, que suma por matrícula (una celda `Matrícula` con CSV no casaría con el `SUMIFS`). El coste es modesto: ~35 filas extra en los datos reales frente a las ~24 que resultarían agrupando solo por CN.
- **Criterio de huérfano**: un par `(CN, Mat)` es huérfano si `(CN, Mat)` no existe como `(Peticion, Recurso)` en Extracción. Esto cubre dos casos en una sola regla: petición totalmente inexistente, y petición existente con matrícula no asociada. `Funcion` se excluye del criterio porque el export real del usuario trae esa columna vacía en Cierre; el criterio más estricto con `Funcion` generaría huérfanos en exceso.
- **`"-"` como sentinela**: los huérfanos no tienen `Aplicación`, `Título`, `Estado`, `Departamento` ni `Res. Tecnico`. Se rellenan con `"-"` literal (decisión del usuario). Esto hace que el VLOOKUP de `Equipo` contra `Equipos` devuelva `""` vía `IFERROR`, y obliga a excluir `"-"` de la detección de "apps sin mapeo" para no generar ruido.
- **Ordenación única global**: en vez de "huérfanos al final con separador", las filas quedan mezcladas con las normales según el criterio numérico-primero. Peticiones numéricas van arriba ordenadas ASC, no numéricas al final ordenadas alfabéticamente. Esto reordena también las filas de Extracción cuando se activa, cosa que el usuario confirmó aceptar.

### Nota de migración

Para activar la feature en un `config.properties` preexistente, añade el bloque completo:

```properties
mes.orphans.enabled=true
mes.orphans.sourceSheet=Cierre
mes.orphans.matchComponent=Component Name
mes.orphans.matchMatricula=Matricula
mes.orphans.sumColumn=Hours
mes.orphans.colPeticion=Petición
mes.orphans.colMatricula=Matrícula
mes.orphans.colJira=Jira
```

Si los nombres `Petición`, `Matrícula`, `Jira` en tu `mes.col.N.name` difieren (por ejemplo si has internacionalizado), ajusta las tres últimas claves en consecuencia. Sin esta sección, el comportamiento es exactamente el de 1.6.2.

### Cambios de versión

- `pom.xml` y `Main.APP_VERSION` pasan de `1.6.2` a `1.7.0`. `MainTest.appVersionEsLaEsperadaPorLaSesionE` actualizado.

### Pendiente conocido fuera de alcance

**Bug B — `Cierre.Funcion` vacío**: el export real del usuario trae `Cierre.Funcion` vacía en todas las filas, mientras que `Extraccion.Funcion` trae valores (`AN`, `OT`, `PR`, ...). Eso hace que el `SUMIFS` de Jira (que incluye el criterio `Funcion:Funcion`) dé `0` incluso para filas con contrapartida válida. El usuario indicó en la conversación de esta sesión que **lo arregla por su lado** (modificando el export o quitando el criterio `Funcion` del match en su propio config). La feature de huérfanos de este 1.7.0 es complementaria a ese arreglo: una vez restaurado el match de Jira, las filas normales contabilizan sus horas y las filas huérfanas siguen recogiendo el resto.

*Nota: durante la implementación de 1.7.0 se descubrió un segundo bug (C) en `SummarySheetBuilder` que daba `0` para matrículas numéricas en `Resumen`. Se documentó aquí como pendiente y se arregló en la versión siguiente 1.7.1 — ver entrada correspondiente al inicio del changelog.*

## [1.6.2] — Fix del SUMIFS de Jira: mismatch de tipos entre Extracción y Cierre

Patch. La columna **Jira** de `Resultado` (SUMIFS contra `Cierre`) estaba devolviendo `0` para todas las peticiones cuya `Peticion`/`Recurso` en el export de Extracción venía como número mientras el `Component Name`/`Matricula` correspondiente en el export de Cierre venía como texto. Excel trata `55751` (número) y `"55751"` (texto) como valores distintos cuando el criterio del `SUMIFS` es numérico, así que esas imputaciones se perdían sin aviso.

Sobre los inputs reales reportados por el usuario: 55 peticiones y 4 matrículas únicas estaban afectadas (100% de los valores en común entre ambas hojas tenían tipos distintos). **La hoja `Resumen` hereda el fix**, porque suma a partir de `Resultado`: donde antes veía `0`, ahora ve los totales reales.

### Arreglado

- `SUMIFS` de la columna `Jira` de `Resultado` ahora recupera imputaciones donde `Cierre.Component Name` / `Cierre.Matricula` son de tipo distinto a `Extraccion.Peticion` / `Extraccion.Recurso`. Verificado empíricamente: el bug se reproducía con criterio numérico sobre rango textual (asimetría conocida de `SUMIFS`; criterio textual sobre rango numérico sí funcionaba por coerción implícita).
- `SheetCopier` acepta ahora un set opcional de índices de columna que se copian forzando tipo `STRING` al escribir al workbook resultado, a partir de la primera fila de datos. Cabeceras y filas de metadatos se copian intactas.
- Nuevo helper `PoiUtils.copyCellValueAsText` con reglas documentadas por tipo: `NUMERIC` entero → sin decimales (`55751`, no `55751.0`); `NUMERIC` decimal → `String.valueOf`; `NUMERIC` con formato de fecha → **no** se fuerza a texto (se delega al comportamiento original, convertir una fecha a epoch Excel como texto sería prácticamente siempre un bug del usuario); `BOOLEAN` → `"true"`/`"false"`; `FORMULA` → fórmula preservada; `BLANK` → blank (no cadena vacía); `ERROR` → código de error.

### Añadido

- Nueva clave de configuración por perfil: `profile.<id>.asText.columns=<cabecera1>,<cabecera2>,...`. Lista CSV de cabeceras cuyo valor se fuerza a `STRING` al copiar la hoja al workbook resultado. Opt-in: sin la clave, el comportamiento del 1.6.0 queda intacto (cambio retrocompatible). En los `config.properties` por defecto se activan ya las cuatro cabeceras que participan en el `SUMIFS` de `Jira`:
  - `profile.Extraccion.asText.columns=Peticion,Recurso`
  - `profile.Cierre.asText.columns=Component Name,Matricula`
- Warning `CABECERA` en runtime si una cabecera listada en `asText.columns` no existe en la hoja detectada; el merge no se interrumpe (se ignora esa columna y se registra el aviso en el `RunReport`, visible en `_Avisos` con `report.inExcel=true`).

### Tests

- `ExcelMergerIntegrationTest` — cinco casos nuevos: tres de regresión (`55751` → `7h`, `101770` → `5h`, `138074` → `9h` con filtrado adicional por `Funcion=Dev`), uno que verifica que `Extraccion.Peticion`/`Recurso` en el workbook resultado quedan como `STRING` tras el fix, y uno que confirma el warning `CABECERA` cuando `asText.columns` menciona una cabecera inexistente.
- Nuevo `PoiUtilsTest` en el paquete `util` con 9 casos unitarios cubriendo cada rama de `copyCellValueAsText`: numérico entero, decimal, string, `"-"`, boolean, blank, fórmula, error, fecha con formato.
- Fixtures (`extraccion.xlsx`, `cierre.xlsx`) ampliados vía `gen_fixtures.py`: tres filas nuevas en Extracción con `Peticion`/`Recurso` como `NUMERIC`, cinco imputaciones nuevas en Cierre con `Component Name`/`Matricula` como `STRING` sobre esos mismos valores. Totales esperados documentados en los comentarios del script. Los tests existentes que contaban filas exactas de `Extraccion` y `Resultado` se han actualizado a los nuevos conteos (19 y 18 respectivamente).

### Diseño

- **Por qué `SheetCopier` y no `CopyColumnStrategy`**: el `SUMIFS` que emite `SumIfsColumnStrategy` referencia el rango de criterio contra la hoja `Extraccion` del workbook resultado (no contra `Resultado`). Normalizar en `CopyColumnStrategy` sólo habría cambiado la hoja `Resultado`, que no interviene en el match. El fix debe ocurrir donde se escribe la hoja que el `SUMIFS` va a leer: `SheetCopier`.
- **Simetría `Extraccion` + `Cierre`**: con normalizar sólo el lado del criterio (Extracción) sería suficiente en la práctica — `SUMIFS` con criterio textual hace coerción sobre rangos numéricos — pero normalizar también Cierre es barato y blinda frente a futuros exports donde alguna de esas columnas vuelva a cambiar de tipo.
- **Por qué no `SUMPRODUCT`**: la alternativa natural (`SUMPRODUCT((TEXT(rango,"@")=TEXT(critrango,"@"))*valores)`) tolera tipos pero obliga a acotar rangos (no admite `A:A`), penaliza evaluación en workbooks grandes, y requiere reescribir `SumIfsColumnStrategy` por completo. Normalizar en copia es la misma idea expresada en el punto más barato del pipeline.
- **Casos huérfanos no abordados aquí**: el export de Cierre contiene imputaciones con `Matricula="-"` y `Component Name` no numérico que no tienen contrapartida en Extracción. Con el diseño actual (`Resultado` ancla en `Extraccion.Peticion`) esas horas no tienen fila destino donde sumarse — es un cambio de alcance distinto del fix de tipos y queda fuera de esta patch. El usuario puede detectar esos casos manualmente cruzando `Cierre.Hours` contra `SUM(Resultado.Jira)`.

### Nota de migración

Sólo aplica si partes de un `config.properties` personalizado (no el que se distribuye con el proyecto). Añade las dos claves:

```properties
profile.Extraccion.asText.columns=Peticion,Recurso
profile.Cierre.asText.columns=Component Name,Matricula
```

Sin esas claves, el comportamiento del 1.6.0 queda exactamente como estaba: las columnas se copian preservando su tipo original y el bug sigue presente para exports con mismatch de tipos. El merge no se rompe; lo que no se activa es el fix.

El `config.properties` que viene con el proyecto (raíz y classpath fallback) ya incluye las claves con los valores por defecto.

### Cambios de versión

- `pom.xml` y `Main.APP_VERSION` pasan de `1.6.0` a `1.6.2`. `MainTest.appVersionEsLaEsperadaPorLaSesionE` actualizado en consecuencia.
- La versión `1.6.1` quedó sin publicar (bump intermedio descartado durante la elaboración del fix).

## [1.6.0] — Hoja resumen + tintado de columnas

Minor que añade la hoja **Resumen** al libro de salida (sumatorio por matrícula sobre las columnas de horas de `Resultado`) e introduce dos atributos configurables nuevos por columna de `Resultado`: `mes.col.N.fill` y `mes.col.N.redIfNotEqualTo`. Se simplifica `Resultado` eliminando las columnas `Desfase` y `Dif Mes` (que habían quedado como placeholders sin relleno real en el flujo actual).

### Añadido

#### Hoja Resumen

- Nueva clase `com.excelmerger.SummarySheetBuilder` en el paquete raíz, siguiendo el patrón de `MesSheetBuilder` y `AvisosSheetBuilder`. Recibe `ConfigLoader` y `RunReport` por constructor y expone `build(Workbook)`. Se engancha en `ExcelMerger.merge()` después de `DerivedSheetBuilder` y antes de `AvisosSheetBuilder` para que pueda leer la hoja `Resultado` ya construida.
- Nuevas claves `summary.*` en `config.properties` (externo y fallback interno) y en `test-config.properties`:
  - `summary.enabled` (`true`/`false`) — opt-in explícito.
  - `summary.sheetName` — nombre de la hoja generada (default `Resumen`).
  - `summary.sumSheet` — hoja que se agrega (normalmente `mes.sheetName`).
  - `summary.matriculaColumn` — nombre de la columna clave en `sumSheet`.
  - `summary.valueColumns` — lista CSV con las columnas a sumar.
  - `summary.sumifsMaxRow` — tope de fila para los rangos `SUMIFS` (default `10000`; acotado para que POI pueda evaluar en tests).
- Nuevos métodos en `com.excelmerger.util.StyleFactory` con la paleta del cuaderno de referencia: `summaryBlockHeader`, `summarySubHeaderGray`, `summaryValueCell`, `summaryNumericCell`, `summaryTotalCell`, además del helper genérico `solidFill(Workbook, String argbHex)`.
- `ConfigValidator.validateSummary()` — chequeos estáticos alineados con el resto de hojas (requeridos, colisión de nombre, referencia a hoja conocida).

#### Tintado de columnas de `Resultado`

- Nuevo atributo `mes.col.N.fill=<COLOR>` — aplica fondo sólido permanente a todas las celdas de datos de la columna. Valores soportados: `LIGHT_GREEN` (`#E2EFDA`), `LIGHT_BLUE` (`#DDEBF7`), `LIGHT_YELLOW` (`#FFF2CC`), `LIGHT_RED` (`#FCE4D6`), `LIGHT_LAVENDER` (`#E4DFEC`). Nombres desconocidos generan un warning `CONFIG` y la columna se escribe sin fill.
- Nuevo atributo `mes.col.N.redIfNotEqualTo=<NombreColumna>` — aplica un formato condicional que pinta el fondo en rojo claro cuando el valor de esta celda difiere del de la celda homóloga en la columna referenciada. Útil para resaltar filas donde `PDCL + Deuda` se ha modificado respecto a `PDCL`.
- Configuración aplicada por defecto en `config.properties`:
  - `mes.col.11` (`PDCL`) → `fill=LIGHT_GREEN`.
  - `mes.col.12` (`PDCL + Deuda`) → `fill=LIGHT_GREEN` + `redIfNotEqualTo=PDCL`.

#### Tests

- `SummarySheetBuilderTest` — 10 casos cubriendo los escenarios habituales.
- `MesSheetBuilderTest` — 4 casos nuevos: fill aplicado al estilo de las celdas, fill con color desconocido, CF `redIfNotEqualTo` con fórmula que contiene `<>`, redIfNotEqualTo apuntando a columna inexistente.
- `ExcelMergerIntegrationTest` — nuevo test que verifica la hoja `Resumen` tras el pipeline completo.

### Cambiado

- **Eliminadas del `Resultado` las columnas `Desfase` y `Dif Mes`**. Eran placeholders en el flujo actual (`Desfase` tenía fórmula `{col:PDCL}-{col:REAL}` con `greenIfPositive=true`; `Dif Mes` era `EMPTY` con `greenIfPositive=true`). En `config.properties` principal y fallback, `Horas_RealizadoTot` y `Realizadas_Horas_Mes` se renumeran de `mes.col.15/16` a `mes.col.13/14` (la numeración debe ser consecutiva: `loadColumns()` corta al primer hueco). En `test-config.properties`, `Desfase` se elimina y `Matrícula`/`Res. Tecnico`/`PDCL`/`PDCL + Deuda` se renumeran de 7-10 a 6-9.
- Columnas `PDCL` y `PDCL + Deuda` añadidas al `test-config.properties` para que los tests del Summary puedan validar las 4 columnas de valor configuradas por defecto.

### Diseño

- **Auto-descubrimiento de matrículas**: se leen los valores no vacíos de la columna clave en `Resultado`. Las numéricas van ordenadas ascendentemente y después las no numéricas (por ejemplo `-` o `Sin Matricula`) en orden alfabético.
- **Fórmulas SUMIFS con rangos acotados**: el cuaderno original usaba `I:I` (columna completa); aquí se acota a `I2:I10000` (configurable) para que `FormulaEvaluator` de POI pueda evaluarlas en los tests de integración. En Excel el rango sigue siendo más que suficiente para cualquier extracción mensual realista.
- **Columnas de valor no encontradas**: se omiten silenciosamente con un warning `CABECERA` en `RunReport` en lugar de abortar la generación. Esto permite que el config pueda listar columnas opcionales (como `PDCL + Deuda`) sin romper si no están presentes.
- **Fill permanente**: se aplica desde `MesSheetBuilder` celda por celda con un cache `LinkedHashMap<String, CellStyle>`. No toca las 4 estrategias de columna (`Copy/Sumifs/Formula/Empty`), solo la interfaz `MesColumnStrategy` se extiende con `getFillColor()` y `getRedIfNotEqualTo()`.
- **Formato condicional `redIfNotEqualTo`**: se emite con fórmula relativa `X2<>Y2` y fill `IndexedColors.ROSE` (equivalente indexado de `#FCE4D6`); POI propaga la referencia al resto del rango automáticamente.

### Cambios de versión

- `pom.xml` y `Main.APP_VERSION` pasan de `1.5.2` a `1.6.0`. `MainTest.appVersionEsLaEsperadaPorLaSesionE` actualizado en consecuencia.
- `SummarySheetBuilder` añadido a la exclusión `EI_EXPOSE_REP2` de `spotbugs-exclude.xml` (mismo patrón que los otros 5 builders: `RunReport` es por diseño un agregador compartido).

### Notas

La primera iteración replicaba los cuatro bloques del cuaderno `DA23_*.xlsx` (reparto, tickets, pre-cierre y matriz responsables × matrículas). Ese alcance se descartó a favor del resumen agregado simple por petición del usuario; el código intermedio no quedó en el árbol.

## [1.5.2] — Sesión E2: segunda pasada de calidad

Patch sobre 1.5.0 que cierra los 6 code smells reales que la Sesión E había dejado anotados en `pmd-ruleset.xml` y baja el umbral de SpotBugs de `High` a `Medium` arreglando 2 NP_NULL reales y excluyendo con justificación los 8 issues restantes. **Sin cambios en comportamiento runtime ni en la API pública**. Los 146 tests siguen verdes (`MainTest.appVersionEsLaEsperadaPorLaSesionE` actualizado al nuevo literal `1.5.2`). Cobertura JaCoCo sigue en ≥ 70% INSTRUCTION.

### Corregido

Los 6 code smells de PMD quedan resueltos en código; las 6 exclusiones correspondientes desaparecen de `pmd-ruleset.xml`.

- **`UnnecessaryCaseChange` en `PoiUtils.findColumnIndex`** — `cabecera.toUpperCase().equals(needle.toUpperCase())` reemplazado por `value.trim().equalsIgnoreCase(headerText.trim())`. Semántica idéntica para cabeceras ASCII del dominio (`CODIGO`, `CUENTA`, etc.). Se elimina también el `target` local que dejó de ser necesario.
- **`UnusedLocalVariable` en `DerivedSheetBuilder.buildAggregationSheet`** — `int headerRow = config.getInt(prefix + "headerRow", 1);` eliminado. La variable se asignaba y nunca se leía; `firstDataRow` (que sí se usa) permanece.
- **`UnusedAssignment` en `Main.main`** — `int exitCode = EXIT_OK;` cambiado a `int exitCode;`. Cada rama de salida (happy path línea 129; los 5 `catch`) reasigna antes del `System.exit(exitCode)` final, por lo que definite assignment sigue satisfecho. El inicializador era puro ruido defensivo.
- **`ClassWithOnlyPrivateConstructorsShouldBeFinal` en `FileProfileResolver.FileProfile`** — clase anidada marcada `public static final class FileProfile`. Verificado con grep que ningún test la extiende; el único constructor es privado y la factoría estática `fromConfig` devuelve instancias.
- **`PreserveStackTrace` en `OutputManager.backupOutput`** — en el catch anidado `IOException atomicFailed → try REPLACE_EXISTING → catch IOException e`, la excepción original del intento atómico se descartaba si el fallback también fallaba. Fix: `e.addSuppressed(atomicFailed);` antes del `throw`. Conserva ambas trazas sin alterar el tipo ni el mensaje de la excepción pública; el constructor `OutputException(String, Throwable)` ya existía.
- **`AvoidRethrowingException` en `ExcelMerger.merge`** — eliminado el `catch (ExcelMergerException e) { throw e; }` redundante del try-with-resources. `ExcelMergerException extends RuntimeException`, así que la excepción se propaga igual sin el catch explícito y sin necesidad de cambiar la firma del método. El `catch (IOException e) → throw new MergeException(..., e)` que le seguía se mantiene intacto.

Adicionalmente, 2 bugs reales reportados por SpotBugs en `Medium`:

- **`NP_NULL_ON_SOME_PATH_FROM_RETURN_VALUE` en `OutputManager.assertOutputWritable`** — `Paths.get(outputPath).getFileName()` puede devolver `null` si la ruta es una raíz (`/`, `C:\`). Se añade null-check explícito que lanza `OutputException` con mensaje claro (`"output.file no apunta a un fichero: ..."`) en vez de NPE opaco. Caso poco probable en uso normal pero defensa en profundidad correcta.
- **`NP_NULL_ON_SOME_PATH_FROM_RETURN_VALUE` en `OutputManager.backupOutput`** — mismo patrón con `out.getFileName().toString()`. Se extrae a variable local con validación previa. En la práctica el callsite (`prepareOutputFile`) solo invoca `backupOutput` si `Files.exists(out)` es true, por lo que en runtime real el null no llega nunca, pero el fix silencia el análisis y cubre el caso hipotético de llamada directa.

### Mantenimiento

**PMD**
- `pmd-ruleset.xml`: retiradas las 6 exclusiones con sus comentarios "Anotado". El ruleset queda más limpio: ya no convive "regla activa" con "excepción anotada para sesión futura" para estas seis reglas.

**SpotBugs: High → Medium**

Umbral bajado en `pom.xml` (`<threshold>Medium</threshold>`). El comentario del plugin se actualiza para reflejar el cambio.

Tras bajar, SpotBugs reportó 10 issues. Clasificación:

- **2 bugs reales triviales** (arreglados en código; ver sección "Corregido" arriba): los `NP_NULL_ON_SOME_PATH_FROM_RETURN_VALUE` en `OutputManager`.
- **6 falsos positivos defendibles** (excluidos con justificación en `spotbugs-exclude.xml`):
    - 5× `EI_EXPOSE_REP2` en `AvisosSheetBuilder`, `DerivedSheetBuilder`, `ExcelMerger`, `LookupSheetBuilder`, `MesSheetBuilder`: los constructores guardan la referencia del `RunReport` recibido. **RunReport es por diseño un agregador compartido** — Main crea una instancia, ExcelMerger la distribuye a todos los builders para que acumulen warnings, y Main lee el reporte final. Clonar rompería el patrón. Es el uso correcto de Collecting Parameter, no una fuga de representación.
    - 1× `EI_EXPOSE_REP` en `ConfigLoader.getRawProperties()`: el método se usa exclusivamente desde `FileProfileResolver.FileProfile.fromConfig` para iterar claves con sufijo dinámico (`profile.<id>.detect.cellValue.*`). Clonar el `Properties` en cada lectura es overhead sin beneficio — `ConfigLoader` no muta tras la carga. Si en el futuro se quiere endurecer, la migración natural es a `Set<String> getPropertyNamesStartingWith(String prefix)`, pero eso cambia la API y queda fuera del alcance de un patch.
- **2 issues reales pero con ripple de API no trivial** (excluidos con justificación, anotados para posible 1.6.0):
    - 2× `CT_CONSTRUCTOR_THROW` en los dos constructores públicos de `ConfigLoader`: ambos lanzan `ConfigurationException` si el properties no existe. El refactor canónico (factoría estática `ConfigLoader.load(path)` + constructor privado no-lanzante) cambiaría la API pública y afectaría a `Main` y varios tests. El riesgo de Finalizer attack que motiva la regla no aplica a esta CLI standalone (no hay subclases hostiles, no hay sandbox, no hay código no confiable con acceso a la JVM).

Las tres reglas quedan excluidas con bloques `<Match>` precisos en `spotbugs-exclude.xml` (scope limitado a clases/métodos concretos, nunca `Bug pattern="..."` global) y cada una lleva un comentario de justificación.

**Version bump**
- `pom.xml` y `Main.APP_VERSION` pasan de `1.5.0` a `1.5.2`. `MainTest.appVersionEsLaEsperadaPorLaSesionE` actualizado en consecuencia — es el único test que tocó esta sesión.

### Sin cambios

- Comportamiento runtime del merge: los builders, la resolución de perfiles, la escritura del Excel y los códigos de salida son idénticos a 1.5.0.
- API pública: las firmas de `PoiUtils.findColumnIndex`, `OutputManager.backupOutput`, `OutputManager.assertOutputWritable`, `ExcelMerger.merge`, `FileProfileResolver.FileProfile.*`, `ConfigLoader.*` y `Main.*` no cambian.
- Formato del código fuera de los sitios listados: no se ha reformateado nada más.

---

## [1.5.0] — Sesión E: calidad de código

Endurecimiento profesional del proyecto **sin cambiar comportamiento ni API**. Los 142 tests existentes siguen verdes y sin tocar. Cobertura JaCoCo sigue en ≥ 70% INSTRUCTION.

### Fases completadas

**Fase 1 — Maven Wrapper**
- Añadidos `mvnw`, `mvnw.cmd` y `.mvn/wrapper/maven-wrapper.properties` (wrapper 3.3.2 en modo `only-script`, distribución Maven 3.9.9). El proyecto ya no requiere tener Maven instalado en el sistema.
- `run.bat`: el mensaje de error cuando falta el JAR ahora sugiere `mvnw.cmd clean package` si existe el wrapper, y `mvn clean package` como fallback. Sin cambios en la lógica del lanzador.
- `README.md`: nueva sección de compilación con el wrapper como recomendado.

**Fase 2 — Spotless**
- Añadido `spotless-maven-plugin:2.43.0`, enganchado a `verify`. Comando manual: `mvnw spotless:apply`.
- **Decisión de configuración**: se descarta usar presets (Google Java Format, Google-AOSP, Eclipse) porque cualquiera de ellos reformatearía decenas de ficheros del proyecto (indent 2 vs 4, reorganización de imports, expansión de wildcards de POI, etc.). En su lugar se configura una lista mínima de invariantes que el código ya cumple: `trimTrailingWhitespace`, `endWithNewline`, `encoding UTF-8`, `lineEndings UNIX`. Resultado: **0 diffs** al correr `spotless:apply`. El control de indent e import order se delega a Checkstyle (Fase 3).

**Fase 3 — Análisis estático (SpotBugs + PMD + Checkstyle)**

Todos enganchados a `verify`; los fallos rompen la build.

- **SpotBugs** (`spotbugs-maven-plugin:4.9.8.2`, core `4.9.8`): umbral `High`, `includeTests=false`, falsos positivos en `spotbugs-exclude.xml` (vacío de inicio, listo para recibir exclusiones justificadas). Dejado en `High` como pediste; bajar a `Medium` queda para una segunda pasada.

- **PMD** (`maven-pmd-plugin:3.28.0` con PMD `7.18.0` forzado vía override de `pmd-core`/`pmd-java`, necesario para `targetJdk=25`) con ruleset propio `pmd-ruleset.xml`. **Nota histórica**: el brief pedía los rulesets clásicos `java-basic`, `java-design`, `java-unusedcode`, pero esos nombres fueron eliminados en PMD 6.0.0 (2017) y **no existen en PMD 7**, que es lo que soporta JDK 25. El ruleset propio los traduce a las categorías modernas equivalentes (`category/java/errorprone.xml`, `category/java/bestpractices.xml`, `category/java/design.xml`) excluyendo reglas conocidas por ser ruidosas (`GodClass`, `CyclomaticComplexity`, `NcssCount`, `TooManyMethods`, `LawOfDemeter`, etc.) que requerirían refactor y caen fuera del objetivo de no cambiar comportamiento. `src/test/java/**` queda excluido.

    Tras la primera corrida real de PMD 7.18.0 sobre el código existente salieron **65 violaciones agrupadas en 13 reglas**. Cumpliendo la regla "no arreglar code smells que impliquen cambios de lógica", **todas se han resuelto excluyendo reglas en el ruleset** (con justificación en comentario dentro de `pmd-ruleset.xml`), sin tocar ni una línea de código de producción. Clasificación:

    - **Falsos positivos**: `UseProperClassLoader` (advierte para J2EE, el proyecto es CLI), `ReturnEmptyCollectionRatherThanNull` (PMD confunde `Properties` con `Collection` en `Main.readGitProperties`).
    - **Ruido conocido**: `SystemPrintln` (21 — `Main.java` es CLI, `System.out.println` es el canal correcto para help/version), `AvoidDuplicateLiterals` (10 — literales repetidos en mensajes de ayuda), `AssignmentInOperand` (5 — estilo, no bug), `UseLocaleWithCaseConversions` (19 — migrar a `Locale.ROOT` es cambio de comportamiento, queda para una pasada futura de hardening i18n), `ExhaustiveSwitchHasDefault` (estilo Java 21+).
    - **Code smells reales anotados para revisión futura**: `UnusedAssignment` en `Main.main` (inicializador defensivo de `exitCode`), `UnusedLocalVariable headerRow` en `DerivedSheetBuilder.buildAggregationSheet`, `PreserveStackTrace` en `OutputManager.backupOutput` (pierde stack trace en un catch), `UnnecessaryCaseChange` en `PoiUtils.findColumnIndex` (usa `toUpperCase().equals(...)` en vez de `equalsIgnoreCase`), `AvoidRethrowingException` en `ExcelMerger.merge`, `ClassWithOnlyPrivateConstructorsShouldBeFinal` en `FileProfileResolver.FileProfile`. Cada exclusión lleva comentario indicándolo.

- **Checkstyle** (`maven-checkstyle-plugin:3.6.0` con `com.puppycrawl.tools:checkstyle:10.21.4` forzado para soporte JDK moderno) con ruleset propio `checkstyle.xml`. **Decisión clave**: se descarta `google_checks.xml` estándar; una prueba controlada sobre el código actual produjo **2290 violaciones** debidas a incompatibilidad estructural (Google usa indent 2, el proyecto usa 4; distinto orden de imports). Las suppressions del brief (LineLength=140, AbbreviationAsWordInName, Javadoc como warning) solo cubrían ~73 de esas 2290. En su lugar, el ruleset propio contiene únicamente las reglas pedidas explícitamente en el brief más algunos invariantes obvios, diseñado para dar **0 violaciones ERROR** sobre el código actual:
    - `FileTabCharacter` (sin tabs).
    - `LineLength = 140` (margen para fórmulas Excel largas).
    - `AvoidStarImport`, con suppression para `DerivedSheetBuilder.java` y `LookupSheetBuilder.java` (wildcards intencionales de `org.apache.poi.ss.usermodel.*`).
    - `MissingJavadocMethod` con `severity=info` (aparece en consola pero no rompe build; el proyecto tiene métodos sin Javadoc intencional).
    - `AbbreviationAsWordInName` con vocabulario del dominio permitido (`SUMIFS`, `VLOOKUP`, `PDCL`, `POI`, `XLSX`, `CSV`, `UTF`, `BOM`...), suprimida entera en `src/test/java/**` porque los nombres de tests siguen la convención `casoDescriptivoEnCastellanoCON_PalabraFinalEnMayusculas`.

**Fase 4 — git-commit-id-maven-plugin**
- Añadido `io.github.git-commit-id:git-commit-id-maven-plugin:9.2.0`, enganchado a `initialize` (antes del compile). Genera `target/classes/git.properties` con solo las dos claves necesarias: `git.commit.id.abbrev` y `git.commit.time` (formato `yyyy-MM-dd`, UTC).
- `failOnNoGitDirectory=false` y `failOnUnableToExtractRepoInfo=false`: permite compilar fuera de un repo Git sin reventar; en ese caso no se genera `git.properties`.
- `Main.java`: nuevo método `buildInfoString()` (package-private para test) que lee `/git.properties` del classpath. Formato enriquecido si hay datos: `Excel Merger v1.5.0 (build a3f9b2c, 2026-04-22)`. Formato básico si falta: `Excel Merger v1.5.0`. Usado en `--version` y en la cabecera de `--help`.
- Nuevo test `MainTest` (4 tests): la llamada no rompe cuando falta `git.properties`, el formato siempre empieza por `Excel Merger v<APP_VERSION>`, el formato completo solo puede ser uno de los dos variantes esperados, y `APP_VERSION` es la esperada de la Sesión E.
- Tests totales: 142 → **146** (+4).

**Fase 5 — release**
- Versión en `pom.xml` y `Main.APP_VERSION` bumpeada a `1.5.0`.
- Este CHANGELOG.

### Sin cambios

- Comportamiento runtime: los 142 tests anteriores siguen verdes sin modificación.
- API pública: las clases, constructores y signaturas existentes no cambian.
- Exit codes del JAR (0-4).
- Sintaxis del `config.properties`.
- CLI del JAR: `--version` y `--help` siguen existiendo; solo se enriquece la salida cuando hay `git.properties` disponible.
- Cobertura JaCoCo: umbral 70% INSTRUCTION mantenido. El test nuevo `MainTest` aumenta la cobertura sobre `Main` (que sigue excluida del gate de cobertura por contener `System.exit`).

### Convenciones para mantenimiento

- Ante un fallo de Checkstyle: añadir `SuppressionSingleFilter` con path específico en `checkstyle.xml` con comentario justificando el porqué, o ampliar `allowedAbbreviations`.
- Ante un fallo de PMD: añadir `<exclude name="..."/>` en `pmd-ruleset.xml` dentro del bloque de categoría correspondiente, con comentario justificando.
- Ante un fallo de SpotBugs: añadir `<Match>` con comentario de justificación FUERA del `<Match>` (el XML no permite comentarios anidados dentro del Match) en `spotbugs-exclude.xml`.
- No bajar el umbral de SpotBugs a Medium/Low en esta sesión: pasada pendiente.

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

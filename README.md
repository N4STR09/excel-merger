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
│   │   ├── DerivedSheetBuilder.java     # Hojas derivadas (fórmulas/agregación)
│   │   └── LookupSheetBuilder.java      # Tablas de mapeo estáticas
│   └── resources/
│       └── config.properties            # Fallback empaquetado en el JAR
├── src/test/
│   ├── java/com/excelmerger/            # Suite JUnit 5 (8 clases, 123 tests)
│   └── resources/
│       ├── test-config.properties       # Config con placeholders para tests
│       └── fixtures/                    # extraccion.xlsx, cierre.xlsx
└── gen_fixtures.py                      # Regenera los fixtures con openpyxl
```

## Requisitos

- **Java 25** (Oracle JDK o cualquier distribución compatible).
- Maven 3.6+ (solo para compilar).

## Compilación

```bash
mvn clean package
```

Genera `target/excel-merger-1.2.0-jar-with-dependencies.jar`.

## Tests

Desde v1.2.1 el proyecto lleva una suite JUnit 5 con red de seguridad sobre los tipos de la Sesión A (`ConfigValidator`, `RunReport`, detección de locks, backup del output) y el resto de builders (`MesSheetBuilder`, `DerivedSheetBuilder`, `LookupSheetBuilder`, `FileProfileResolver`, `ExcelMerger`).

### Ejecutar la suite

```bash
# Solo tests unitarios + integración (rápido)
mvn test

# Tests + reporte y gate de cobertura JaCoCo (≥70% INSTRUCTION)
mvn clean verify
```

El reporte HTML de cobertura queda en `target/site/jacoco/index.html`. Si el gate del 70% falla, `mvn verify` sale con error indicando la diferencia.

### Organización

- **8 clases en `src/test/java/com/excelmerger/`**, 123 tests: `ConfigLoaderTest`, `ConfigValidatorTest`, `RunReportTest`, `FileProfileResolverTest`, `LookupSheetBuilderTest`, `MesSheetBuilderTest`, `DerivedSheetBuilderTest`, `ExcelMergerIntegrationTest`.
- **`TestFixtures.java`** es la utilidad compartida: copia los fixtures a un `@TempDir`, renderiza `test-config.properties` sustituyendo `${TEST_INPUT_DIR}` / `${TEST_OUTPUT_FILE}`, y ofrece `configFromProperties(...)` para tests unitarios que no tocan disco.
- Todos los tests usan `@TempDir`; no hay efectos colaterales fuera del directorio temporal de cada caso.

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

El `pom.xml` configura JaCoCo 0.8.12 con umbral **70% INSTRUCTION a nivel de bundle**, excluyendo únicamente `com/excelmerger/Main.class` (contiene `System.exit()` y lógica de CLI que requeriría un runner externo para cubrirse). Si en alguna iteración el umbral falla en una clase concreta por ramas difíciles (p. ej. la heurística `looksLikeLocked` de `ExcelMerger` en Linux), la recomendación es **añadir exclusiones granulares** en el bloque de JaCoCo, no bajar el umbral global.

## Ejecución

### Windows (recomendado)

Doble clic en `run.bat`.

### Línea de comandos

```bash
java -jar target/excel-merger-1.2.0-jar-with-dependencies.jar
```

Opcionalmente se puede pasar un config alternativo:

```bash
java -jar target/excel-merger-1.2.0-jar-with-dependencies.jar mi-config.properties
```

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
input.strictTwoFiles=true
output.file=output/resultado_fusion.xlsx
output.overwrite=true

# Novedad v1.2.0: backup del output anterior si existe
output.backup=false
```

Si `output.backup=true` y el archivo de salida ya existe, antes de sobreescribirlo se mueve a `<carpeta del output>/history/<nombre>_yyyy-MM-dd_HHmmss.xlsx` (la carpeta `history` se crea si no existe).

### Perfiles (detección por contenido)

Cada perfil define las cabeceras que deben encontrarse para considerar que un Excel es de ese tipo.

```properties
profiles=Extraccion,Cierre

profile.Extraccion.sheetName=Extraccion
profile.Extraccion.detect.headerRow=1
profile.Extraccion.detect.headers=Peticion,Titulo,Estado,Recurso
profile.Extraccion.detect.minMatches=4

profile.Cierre.sheetName=Cierre
profile.Cierre.detect.headerRow=2
profile.Cierre.detect.headers=Project Key,Issue Key,Hours
profile.Cierre.detect.minMatches=3
```

Criterios soportados:
- `detect.headers` — lista de cabeceras (case-insensitive, match por "contains").
- `detect.minMatches` — cuántas deben coincidir.
- `detect.cellValue.<REF>` — opcional, fuerza que una celda concreta contenga un valor.
- `detect.sheetIndex` — opcional, qué hoja analizar (default 0).

### Hoja Resultado

Estructura fija de columnas definida en el config. Cada columna tiene un tipo:

- **COPY** — copia directa de una columna de la hoja origen.
- **SUMIFS** — suma condicional cruzando con otra hoja.
- **FORMULA** — fórmula libre con placeholders `{col:NombreColumnaMES}` y `{colLetter:X}`.
- **EMPTY** — celda vacía (placeholder).

Modificador:
- `greenIfPositive=true` — formato condicional: fondo verde si la celda es ≥ 0.

Ejemplo:

```properties
mes.enabled=true
mes.sheetName=Resultado
mes.sourceSheet=Extraccion
mes.sourceHeaderRow=1
mes.anchorColumn=Peticion

mes.col.1.name=Petición
mes.col.1.type=COPY
mes.col.1.from=Peticion

mes.col.9.name=Jira
mes.col.9.type=SUMIFS
mes.col.9.from=Cierre
mes.col.9.sum=Hours
mes.col.9.match=Component Name:Peticion,Matricula:Recurso

mes.col.10.name=REAL
mes.col.10.type=FORMULA
mes.col.10.formula={col:Jira}*1.2

mes.col.13.name=Desfase
mes.col.13.type=FORMULA
mes.col.13.formula={col:PDCL}-{col:REAL}
mes.col.13.greenIfPositive=true
```

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

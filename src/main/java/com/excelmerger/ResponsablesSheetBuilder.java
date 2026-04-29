package com.excelmerger;

import com.excelmerger.util.PoiUtils;
import com.excelmerger.util.StyleFactory;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.text.Collator;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/**
 * Construye una hoja por cada responsable distinto que aparezca en la
 * columna {@code Res. Tecnico} de la hoja {@code Resultado}.
 *
 * <p>Solo se invoca desde {@link ExcelMerger} cuando el modo de salida es
 * {@link OutputMode#RESPONSABLES} o {@link OutputMode#COMPLETO}. En modo
 * {@link OutputMode#CIERRE} no se llama y por tanto el comportamiento
 * v2.2.0 queda preservado al 100%.</p>
 *
 * <h2>v2.4.0 — Tablas pivot por hoja de responsable</h2>
 *
 * <p>Por defecto cada hoja de responsable contiene, ademas de la cabecera
 * A1 (presente desde v2.3.0), dos tablas pivot Petición × Matrícula
 * apiladas verticalmente. <b>Orden v2.7.0 (Modif 3):</b></p>
 * <ol>
 *   <li>PDCL por Petición × Matrícula <i>(antes: Facturar; v2.7.0)</i>.</li>
 *   <li>Horas imputadas (Jira) por Petición × Matrícula.</li>
 * </ol>
 *
 * <p>Las pivots son SUMIFS vivos contra Resultado, filtrados por el
 * responsable cuyo nombre figura en {@code A1}. Las peticiones y
 * matrículas que aparecen son únicamente las que ese responsable tiene
 * en Resultado (no todas las del libro), por consistencia con la Tabla 2
 * de Resumen.</p>
 *
 * <p>Las pivots se desactivan con {@code responsables.tables.enabled=false}
 * (default {@code true}); en ese caso las hojas quedan como en v2.3.0
 * (solo la cabecera A1).</p>
 *
 * <h3>Reglas de extracción del responsable (sin cambios desde v2.3.0)</h3>
 * <ul>
 *   <li><b>Trim + case-insensitive</b>: valores como {@code "tresp1@x"},
 *       {@code "TRESP1@x"} y {@code " tresp1@x "} colapsan en una unica
 *       hoja. Se usa la primera ocurrencia (en orden de filas de
 *       {@code Resultado}) como nombre canónico.</li>
 *   <li><b>Sin filtros ni exclusiones</b>: si el valor (tras trim) es
 *       no-vacio, genera hoja.</li>
 *   <li><b>Orden de creacion</b>: alfabetico por nombre canónico, con
 *       {@link Collator} {@code es_ES} y {@code PRIMARY strength}.</li>
 *   <li><b>Naming de hoja</b>: saneo {@code \\ / ? * [ ] :} → {@code _},
 *       trunca a 31 chars, sufijo {@code _2} en colisiones.</li>
 * </ul>
 *
 * <h3>Reglas de la pivot (v2.4.0)</h3>
 * <ul>
 *   <li><b>Orden de peticiones/matrículas</b>: numéricas ascendentes
 *       primero, no numéricas alfabético después
 *       ({@link PoiUtils#mixedNumericLexicographicSort}).</li>
 *   <li><b>SUMIFS</b>: 3 criterios, rangos acotados a
 *       {@code summary.sumifsMaxRow} (default 10000), criterio del
 *       responsable es {@code $A$1} de la hoja destino (referencia a
 *       celda, no literal — evita problemas de escapado).</li>
 *   <li><b>Recálculo</b>: si se generan pivots, se setea
 *       {@code workbook.setForceFormulaRecalculation(true)} para
 *       garantizar evaluación al abrir en Excel (lección 1.8.0).</li>
 * </ul>
 */
public class ResponsablesSheetBuilder {

    private static final Logger log = LoggerFactory.getLogger(ResponsablesSheetBuilder.class);

    private static final String WARN_CATEGORY = "RESPONSABLE";
    private static final String WARN_CATEGORY_HOJA = "HOJA";

    /** Nombres de columna en Resultado (alineados con config.properties). */
    private static final String COL_PETICION = "Petición";
    private static final String COL_MATRICULA = "Matrícula";
    private static final String COL_JIRA = "Jira";
    /**
     * v2.7.0 (Modif 3): la primera tabla pivot pasa a leer la columna PDCL
     * (mes.col.12, fórmula {@code {col:Jira}*1.2}) en lugar de Facturar
     * (mes.col.11). El usuario quiere ver la columna PDCL real en la pivot
     * principal de cada hoja de responsable. La constante COL_FACTURAR de
     * v2.4.0..v2.6.0 desaparece; el builder ya no consulta Facturar.
     */
    private static final String COL_PDCL = "PDCL";
    private static final String COL_RESPONSABLE = "Res. Tecnico";

    /** Defaults v2.7.0 — claves nuevas. */
    private static final String DEFAULT_JIRA_TITLE =
            "Horas imputadas (Jira) por Petición × Matrícula";
    /**
     * v2.7.0 (Modif 3): la primera tabla pasa a ser PDCL. Antes el default era
     * "Facturar por Petición × Matrícula" bajo la clave facturarTitle (v2.4.0..v2.6.0)
     * o "REAL por Petición × Matrícula" bajo realTitle (≤v2.5.1). Ambas claves
     * obsoletas son rechazadas con error en v2.7.0 — ver
     * {@link com.excelmerger.config.ResponsablesTablesConfigSection}.
     */
    private static final String DEFAULT_PDCL_TITLE =
            "PDCL por Petición × Matrícula";
    private static final int DEFAULT_GAP_ROWS = 2;

    private final ConfigLoader config;
    private final RunReport report;

    public ResponsablesSheetBuilder(ConfigLoader config, RunReport report) {
        this.config = config;
        this.report = report;
    }

    /**
     * Punto de entrada. Recorre Resultado, descubre los responsables únicos
     * y, para cada uno, crea su hoja con la cabecera A1 y (si las tablas
     * pivot están habilitadas) las dos tablas pivot.
     *
     * <p>v2.7.0 (Modif 1): devuelve la lista de nombres de hoja creados (en
     * el orden en que se crearon), para que el paso final de freeze top row
     * pueda excluirlos (decision Fase 0 P1: las hojas de responsable NO
     * tienen freeze). Si no se crea ninguna hoja, devuelve lista vacia
     * (nunca {@code null}).</p>
     */
    public List<String> buildAll(Workbook workbook) {
        List<String> created = new ArrayList<>();
        String resultadoName = config.get("mes.sheetName", "MES");
        String responsibleColumn = config.get("summary.byResponsible.column", COL_RESPONSABLE);

        Sheet resultado = workbook.getSheet(resultadoName);
        if (resultado == null) {
            log.warn("[Responsables] La hoja '{}' no existe; no se generan hojas por responsable.",
                    resultadoName);
            report.addWarning(WARN_CATEGORY,
                    "No se pudo construir hojas por responsable: hoja '"
                            + resultadoName + "' ausente.");
            return created;
        }
        Row header = resultado.getRow(0);
        if (header == null) {
            log.warn("[Responsables] La hoja '{}' no tiene cabecera; no se generan hojas por responsable.",
                    resultadoName);
            report.addWarning(WARN_CATEGORY,
                    "No se pudo construir hojas por responsable: hoja '"
                            + resultadoName + "' sin cabecera.");
            return created;
        }
        int respColIdx = PoiUtils.findColumnIndex(header, responsibleColumn);
        if (respColIdx < 0) {
            log.warn("[Responsables] Columna '{}' no encontrada en '{}'; no se generan hojas por responsable.",
                    responsibleColumn, resultadoName);
            report.addWarning(WARN_CATEGORY,
                    "No se pudo construir hojas por responsable: columna '"
                            + responsibleColumn + "' no encontrada en '" + resultadoName + "'.");
            return created;
        }
        if (resultado.getLastRowNum() < 1) {
            log.info("[Responsables] La hoja '{}' no tiene filas de datos; nada que generar.",
                    resultadoName);
            report.addWarning(WARN_CATEGORY,
                    "Hoja '" + resultadoName + "' sin filas de datos; 0 hojas por responsable generadas.");
            return created;
        }

        // v2.4.0 — descubrir, en una sola pasada, responsables + peticiones + matriculas
        // por responsable. Si las tablas estan deshabilitadas, lo mismo que v2.3.0
        // (solo necesitamos los responsables); pero hacer la pasada completa una
        // sola vez es trivial y evita un segundo recorrido de Resultado.
        int peticionColIdx = PoiUtils.findColumnIndex(header, COL_PETICION);
        int matriculaColIdx = PoiUtils.findColumnIndex(header, COL_MATRICULA);

        Map<String, ResponsableData> dataByLowerKey = scanResultado(
                resultado, respColIdx, peticionColIdx, matriculaColIdx);

        if (dataByLowerKey.isEmpty()) {
            log.info("[Responsables] No hay responsables no-vacios en '{}'; 0 hojas generadas.",
                    resultadoName);
            return created;
        }

        // Ordenar alfabeticamente por nombre canonico con Collator es_ES PRIMARY.
        List<ResponsableData> ordered = new ArrayList<>(dataByLowerKey.values());
        Collator collator = Collator.getInstance(Locale.of("es", "ES"));
        collator.setStrength(Collator.PRIMARY);
        Collections.sort(ordered, (a, b) -> collator.compare(a.canonical, b.canonical));

        log.info("[Responsables] {} responsable(s) distinto(s) detectado(s) en '{}'.",
                ordered.size(), resultadoName);

        // ¿Tablas pivot habilitadas? Default true (retrocompatibilidad: si la clave
        // no esta, se asume true para que un upgrade del binario aporte la feature
        // sin necesidad de tocar config.properties).
        boolean tablesEnabled = config.getBoolean("responsables.tables.enabled", true);

        // Resolver letras de columnas de Resultado UNA SOLA VEZ (mismo origen
        // para todas las tablas).
        ColumnLetters letters = resolveColumnLetters(header, resultadoName, tablesEnabled);

        CellStyle titleStyle = StyleFactory.title(workbook);
        int totalSumifs = 0;
        int totalTables = 0;

        for (ResponsableData rd : ordered) {
            String unique = createSheetForResponsable(workbook, rd, titleStyle);
            created.add(unique);
            Sheet sheet = workbook.getSheet(unique);

            int rowsInSheet = 1;

            // Tablas pivot (v2.4.0)
            if (tablesEnabled && letters != null) {
                int pivotsLastRow = writePivots(workbook, sheet, resultadoName, letters, rd);
                if (pivotsLastRow > 0) {
                    rowsInSheet = pivotsLastRow + 1; // 0-based -> count
                    totalTables += 2;
                    totalSumifs += rd.peticiones.size() * rd.matriculas.size() * 2;
                }
            }

            applyColumnWidths(sheet, tablesEnabled, rd);

            report.addSheet(unique, rowsInSheet);
            log.info("[Responsables] Hoja creada: '{}' (responsable canonico: '{}', filas: {})",
                    unique, rd.canonical, rowsInSheet);
        }

        if (totalTables > 0) {
            // Garantizar que Excel reevalúe las fórmulas al abrir el libro.
            // Idempotente con SummarySheetBuilder, que también lo setea.
            workbook.setForceFormulaRecalculation(true);
            report.addWarning(WARN_CATEGORY_HOJA,
                    "Responsables: " + ordered.size() + " hoja(s) con "
                            + totalTables + " tabla(s) pivot; "
                            + totalSumifs + " SUMIFS escritos.");
            log.info("[Responsables] Total: {} hojas con {} pivots y {} SUMIFS.",
                    ordered.size(), totalTables, totalSumifs);
        }
        return created;
    }

    // ==================================================================
    //  Creación de la hoja con cabecera A1
    // ==================================================================

    private String createSheetForResponsable(Workbook workbook, ResponsableData rd,
                                             CellStyle titleStyle) {
        String safe = FileProfileResolver.safeSheetName(rd.canonical);
        if (!safe.equals(rd.canonical)) {
            report.addWarning(WARN_CATEGORY,
                    "Nombre de responsable saneado para uso como nombre de hoja: '"
                            + rd.canonical + "' -> '" + safe + "'.");
        }
        String unique = FileProfileResolver.ensureUniqueSheetName(workbook, safe);
        if (!unique.equals(safe)) {
            report.addWarning(WARN_CATEGORY,
                    "Colision de nombre de hoja para responsable; resuelto con sufijo: '"
                            + safe + "' -> '" + unique + "'.");
        }

        Sheet sheet = workbook.createSheet(unique);
        Row headerRow = sheet.createRow(0);
        Cell a1 = headerRow.createCell(0);
        a1.setCellValue(rd.canonical);
        a1.setCellStyle(titleStyle);
        return unique;
    }

    // ==================================================================
    //  Pasada única sobre Resultado
    // ==================================================================

    /**
     * Recorre Resultado una sola vez y agrupa por responsable (clave normalizada
     * trim + lowercase). Para cada responsable acumula peticiones únicas y
     * matrículas únicas (ambas en {@link LinkedHashSet} para orden determinista
     * antes de la ordenación final).
     *
     * <p>Si las columnas Petición o Matrícula no están en Resultado (caso muy
     * raro), {@code peticionColIdx} y/o {@code matriculaColIdx} llegan a {@code -1}
     * y se ignoran las celdas correspondientes; las pivots saldrán como
     * "Sin datos" en {@link ResponsablePivotBuilder}.</p>
     */
    private Map<String, ResponsableData> scanResultado(Sheet resultado,
                                                       int respColIdx,
                                                       int peticionColIdx,
                                                       int matriculaColIdx) {
        Map<String, ResponsableData> out = new LinkedHashMap<>();
        for (int r = 1; r <= resultado.getLastRowNum(); r++) {
            Row row = resultado.getRow(r);
            if (row == null) continue;

            String responsableTrim = readTrimmedCell(row, respColIdx);
            if (responsableTrim == null) continue;

            String key = responsableTrim.toLowerCase(Locale.ROOT);
            ResponsableData data = out.computeIfAbsent(
                    key, k -> new ResponsableData(responsableTrim));

            if (peticionColIdx >= 0) {
                String pet = readTrimmedCell(row, peticionColIdx);
                if (pet != null) data.peticiones.add(pet);
            }
            if (matriculaColIdx >= 0) {
                String mat = readTrimmedCell(row, matriculaColIdx);
                if (mat != null) data.matriculas.add(mat);
            }
        }
        return out;
    }

    private static String readTrimmedCell(Row row, int colIdx) {
        if (colIdx < 0) return null;
        Cell cell = row.getCell(colIdx);
        if (cell == null) return null;
        String raw = PoiUtils.cellAsString(cell);
        if (raw == null) return null;
        String trimmed = raw.trim();
        return trimmed.isEmpty() ? null : trimmed;
    }

    // ==================================================================
    //  Resolución de letras de columna en Resultado
    // ==================================================================

    /**
     * Devuelve las letras Excel de las columnas Petición, Matrícula,
     * Res. Tecnico, Jira y PDCL en la hoja Resultado. Si alguna falta y
     * las tablas están habilitadas, registra un warning y devuelve {@code null}
     * (lo que desactiva las pivots para todas las hojas en esta ejecución).
     * Si las tablas están deshabilitadas, devuelve {@code null} sin warning.
     *
     * <p>v2.7.0 (Modif 3): la columna leida por la primera pivot pasa a ser
     * PDCL en lugar de Facturar.</p>
     */
    private ColumnLetters resolveColumnLetters(Row headerRow, String resultadoName,
                                                boolean tablesEnabled) {
        if (!tablesEnabled) return null;

        String petLetter = ResponsablePivotBuilder.letterOrNull(headerRow, COL_PETICION);
        String matLetter = ResponsablePivotBuilder.letterOrNull(headerRow, COL_MATRICULA);
        String respLetter = ResponsablePivotBuilder.letterOrNull(headerRow, COL_RESPONSABLE);
        String jiraLetter = ResponsablePivotBuilder.letterOrNull(headerRow, COL_JIRA);
        String pdclLetter = ResponsablePivotBuilder.letterOrNull(headerRow, COL_PDCL);

        List<String> missing = new ArrayList<>();
        if (petLetter == null) missing.add(COL_PETICION);
        if (matLetter == null) missing.add(COL_MATRICULA);
        if (respLetter == null) missing.add(COL_RESPONSABLE);
        if (jiraLetter == null) missing.add(COL_JIRA);
        if (pdclLetter == null) missing.add(COL_PDCL);
        if (!missing.isEmpty()) {
            String msg = "Tablas pivot por responsable omitidas: columna(s) ausente(s) en '"
                    + resultadoName + "': " + String.join(", ", missing) + ".";
            log.warn("[Responsables] {}", msg);
            report.addWarning(WARN_CATEGORY, msg);
            return null;
        }
        return new ColumnLetters(petLetter, matLetter, respLetter, jiraLetter, pdclLetter);
    }

    // ==================================================================
    //  Escritura de las dos tablas pivot por hoja
    // ==================================================================

    /**
     * Escribe las dos tablas pivot en {@code sheet} y devuelve el indice
     * 0-based de la última fila ocupada (para calcular cuántas filas
     * tiene la hoja en el {@link RunReport}).
     *
     * <p>v2.7.0 (Modif 3): el orden se invierte respecto a v2.4.0..v2.6.0.
     * Antes era Jira primero, Facturar segunda. Ahora es PDCL primero, Jira
     * segunda. La fuente de la primera tabla pasa de Facturar (mes.col.11)
     * a PDCL (mes.col.12). El gap entre tablas y el calculo de la fila
     * de inicio de la segunda tabla se mantienen identicos: la segunda
     * tabla empieza en {@code primera.lastRow0 + 1 + gapRows}.</p>
     */
    private int writePivots(Workbook wb, Sheet sheet, String resultadoName,
                            ColumnLetters letters, ResponsableData rd) {
        // Ordenar peticiones y matriculas con la politica unificada:
        // numericas ascendentes primero, no numericas alfabetico despues.
        List<String> peticionesSorted = PoiUtils.mixedNumericLexicographicSort(rd.peticiones);
        List<String> matriculasSorted = PoiUtils.mixedNumericLexicographicSort(rd.matriculas);

        int sumifsMaxRow = Math.max(2, config.getInt("summary.sumifsMaxRow", 10000));
        // v2.7.0: la clave era "responsables.tables.facturarTitle" hasta v2.6.0
        // y "responsables.tables.realTitle" hasta v2.5.1. Ambas se rechazan
        // con error en ResponsablesTablesConfigSection (sin alias retrocompat).
        String pdclTitle = config.get("responsables.tables.pdclTitle", DEFAULT_PDCL_TITLE);
        String jiraTitle = config.get("responsables.tables.jiraTitle", DEFAULT_JIRA_TITLE);
        int gapRows = Math.max(0,
                config.getInt("responsables.tables.gapRows", DEFAULT_GAP_ROWS));

        ResponsablePivotBuilder pivot = new ResponsablePivotBuilder(config);

        // Tabla 1 (PDCL) — empieza en fila 2 (0-based) = fila Excel 3.
        // v2.7.0 (Modif 3): antes la primera tabla era Jira; ahora es PDCL.
        ResponsablePivotBuilder.PivotInputs pdclInputs = new ResponsablePivotBuilder.PivotInputs(
                resultadoName,
                letters.peticion, letters.matricula, letters.responsable, letters.pdcl,
                pdclTitle,
                peticionesSorted, matriculasSorted,
                sumifsMaxRow);
        ResponsablePivotBuilder.PivotResult pdclResult =
                pivot.writeTable(wb, sheet, 2, pdclInputs);

        // Tabla 2 (Jira) — empieza tras la primera + gapRows filas en blanco.
        // v2.7.0 (Modif 3): antes esta era la primera tabla; ahora va segunda.
        int jiraStartRow0 = pdclResult.lastRow0 + 1 + gapRows;

        ResponsablePivotBuilder.PivotInputs jiraInputs = new ResponsablePivotBuilder.PivotInputs(
                resultadoName,
                letters.peticion, letters.matricula, letters.responsable, letters.jira,
                jiraTitle,
                peticionesSorted, matriculasSorted,
                sumifsMaxRow);
        ResponsablePivotBuilder.PivotResult jiraResult =
                pivot.writeTable(wb, sheet, jiraStartRow0, jiraInputs);

        return jiraResult.lastRow0;
    }

    // ==================================================================
    //  Anchuras de columna
    // ==================================================================

    private void applyColumnWidths(Sheet sheet, boolean tablesEnabled, ResponsableData rd) {
        if (!tablesEnabled || rd.matriculas.isEmpty()) {
            // v2.3.0 behavior: autosize de A solamente. Envuelto en try
            // porque autoSizeColumn requiere fonts disponibles en algunos
            // entornos de test sin display.
            try {
                sheet.autoSizeColumn(0);
            } catch (Exception ignored) {
                // ignorado; mantenemos compatibilidad de comportamiento v2.3.0
            }
            return;
        }
        // Mismas anchuras que SummarySheetBuilder:
        //   col 0 (clave Petición) = 20 chars
        //   col 1..n (matriculas)  = 14 chars
        //   col última (Total)     = 14 chars
        sheet.setColumnWidth(0, 20 * 256);
        int lastDataCol = 1 + rd.matriculas.size();
        for (int c = 1; c <= lastDataCol; c++) {
            sheet.setColumnWidth(c, 14 * 256);
        }
    }

    // ==================================================================
    //  Helpers / tipos internos
    // ==================================================================

    /**
     * Datos acumulados por responsable durante la pasada única sobre Resultado.
     * Mutable por construcción (los sets se rellenan a medida que se itera).
     */
    private static final class ResponsableData {
        final String canonical;
        final Set<String> peticiones = new LinkedHashSet<>();
        final Set<String> matriculas = new LinkedHashSet<>();

        ResponsableData(String canonical) {
            this.canonical = canonical;
        }
    }

    /**
     * Letras Excel de las 5 columnas que las pivots leen de Resultado.
     * v2.7.0 (Modif 3): el último campo se renombra de {@code facturar} a
     * {@code pdcl} acompañando el cambio de fuente de la primera pivot.
     */
    private static final class ColumnLetters {
        final String peticion;
        final String matricula;
        final String responsable;
        final String jira;
        final String pdcl;

        ColumnLetters(String peticion, String matricula, String responsable,
                      String jira, String pdcl) {
            this.peticion = peticion;
            this.matricula = matricula;
            this.responsable = responsable;
            this.jira = jira;
            this.pdcl = pdcl;
        }
    }
}

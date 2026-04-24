package com.excelmerger;

import com.excelmerger.util.PoiUtils;
import com.excelmerger.util.StyleFactory;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/**
 * Construye la hoja "Resumen" (v1.6.0): sumatorio por matrícula de las
 * columnas numéricas de {@code Resultado} (Jira, REAL, PDCL, PDCL + Deuda
 * por defecto).
 *
 * <p>Estructura de la hoja generada (tabla 1, por matrícula):</p>
 * <ul>
 *   <li>Fila 1: titulo "Resumen por Matrícula" con merge sobre todas las
 *       columnas.</li>
 *   <li>Fila 3: cabeceras ("Matrícula" + valores configurados en
 *       {@code summary.valueColumns}).</li>
 *   <li>Filas 4..N: una por cada matrícula única detectada en
 *       {@code Resultado}, con SUMIFS sobre dicha hoja.</li>
 *   <li>Fila N+1: totales por columna.</li>
 * </ul>
 *
 * <p>v1.8.0 — segunda tabla opcional en la misma hoja, debajo de la
 * primera (opt-in con {@code summary.byResponsible.enabled=true}): matriz
 * cruzada Matrícula × Responsable con el valor de una sola columna
 * configurable (PDCL por defecto).</p>
 *
 * <p>Las matrículas se auto-descubren leyendo la columna
 * {@code summary.matriculaColumn} (por defecto "Matrícula") de la hoja
 * {@code summary.sumSheet} (por defecto "Resultado"). Se incluyen todos
 * los valores no vacíos (numéricos o texto, como "-" o "Sin Matricula").</p>
 *
 * <p>Los SUMIFS se emiten con rangos acotados ({@code D2:D10000}) para
 * que Apache POI pueda evaluarlos en los tests de integración; el tope
 * es configurable con {@code summary.sumifsMaxRow}.</p>
 */
public class SummarySheetBuilder {

    private static final Logger log = LoggerFactory.getLogger(SummarySheetBuilder.class);

    /** Categoría de los avisos relacionados con detección de cabeceras. */
    private static final String WARN_CATEGORY_CABECERA = "CABECERA";

    /** Sufijo común de los mensajes "columna X no encontrada en Y". */
    private static final String MSG_COL_NOT_FOUND_IN = "' no encontrada en '";

    /** Nombre por defecto de la columna clave en Resultado. */
    private static final String DEFAULT_MATRICULA_COLUMN = "Matrícula";
    /** Columnas agregadas por defecto. */
    private static final String DEFAULT_VALUE_COLUMNS = "Jira,REAL,PDCL,PDCL + Deuda";
    /** Nombre de la hoja origen por defecto. */
    private static final String DEFAULT_SUM_SHEET = "Resultado";
    /** Nombre de la hoja generada por defecto. */
    private static final String DEFAULT_SHEET_NAME = "Resumen";

    // v1.8.0 — defaults de la segunda tabla (por responsable)
    /** Columna de Resultado usada como agrupador para la segunda tabla. */
    private static final String DEFAULT_BY_RESP_COLUMN = "Res. Tecnico";
    /** Columna de valor de la segunda tabla (una sola metrica). */
    private static final String DEFAULT_BY_RESP_VALUE = "PDCL";
    /** Titulo de la segunda tabla. */
    private static final String DEFAULT_BY_RESP_TITLE = "Totales Peticiones por Responsables Matrículas";
    /** Filas en blanco entre la primera y la segunda tabla. */
    private static final int DEFAULT_BY_RESP_GAP_ROWS = 2;

    // Layout (1-indexado)
    private static final int ROW_TITLE = 1;
    private static final int ROW_HEADER = 3;
    private static final int ROW_FIRST_DATA = 4;

    private final ConfigLoader config;
    private final RunReport report;

    public SummarySheetBuilder(ConfigLoader config, RunReport report) {
        this.config = config;
        this.report = report;
    }

    public void build(Workbook workbook) {
        if (!config.getBoolean("summary.enabled", false)) {
            log.info("Deshabilitado (summary.enabled=false).");
            return;
        }

        String sheetName = config.get("summary.sheetName", DEFAULT_SHEET_NAME);
        if (workbook.getSheet(sheetName) != null) {
            log.warn("Ya existe una hoja '{}'. Resumen omitido.", sheetName);
            report.addWarning("HOJA",
                    "Ya existe una hoja '" + sheetName + "'. Resumen omitido.");
            return;
        }

        String sumSheetName = config.get("summary.sumSheet", DEFAULT_SUM_SHEET);
        Sheet sumSheet = workbook.getSheet(sumSheetName);
        if (sumSheet == null) {
            log.warn("La hoja origen '{}' no existe. Resumen omitido.", sumSheetName);
            report.addWarning("HOJA",
                    "La hoja origen '" + sumSheetName + "' no existe. Resumen omitido.");
            return;
        }

        String matrColumnName = config.get("summary.matriculaColumn", DEFAULT_MATRICULA_COLUMN);
        List<String> valueColumnNames = parseCsv(config.get("summary.valueColumns", DEFAULT_VALUE_COLUMNS));
        if (valueColumnNames.isEmpty()) {
            log.warn("summary.valueColumns vacio; Resumen omitido.");
            report.addWarning("CONFIG",
                    "summary.valueColumns vacio. Resumen omitido.");
            return;
        }

        Row sumHeaderRow = sumSheet.getRow(0);
        if (sumHeaderRow == null) {
            log.warn("La hoja origen '{}' no tiene cabeceras. Resumen omitido.", sumSheetName);
            report.addWarning(WARN_CATEGORY_CABECERA,
                    "'" + sumSheetName + "' sin cabeceras. Resumen omitido.");
            return;
        }

        int matrColIdx = PoiUtils.findColumnIndex(sumHeaderRow, matrColumnName);
        if (matrColIdx < 0) {
            log.warn("Columna '{}' no encontrada en '{}'. Resumen omitido.",
                    matrColumnName, sumSheetName);
            report.addWarning(WARN_CATEGORY_CABECERA,
                    "Columna '" + matrColumnName + MSG_COL_NOT_FOUND_IN
                            + sumSheetName + "'. Resumen omitido.");
            return;
        }

        // Resolver letras de columna para cada valor configurado; si alguna
        // no existe, avisa y la quita de la lista (en vez de abortar toda la hoja).
        List<ValueColumn> valueColumns = new ArrayList<>();
        for (String name : valueColumnNames) {
            int idx = PoiUtils.findColumnIndex(sumHeaderRow, name);
            if (idx < 0) {
                log.warn("Columna de valor '{}' no encontrada en '{}'. Se omite de Resumen.",
                        name, sumSheetName);
                report.addWarning(WARN_CATEGORY_CABECERA,
                        "Columna de valor '" + name + MSG_COL_NOT_FOUND_IN
                                + sumSheetName + "'. Columna omitida de Resumen.");
                continue;
            }
            valueColumns.add(new ValueColumn(name, CellReference.convertNumToColString(idx)));
        }
        if (valueColumns.isEmpty()) {
            log.warn("Ninguna columna de valor localizada en '{}'. Resumen omitido.", sumSheetName);
            return;
        }

        List<String> matriculas = discoverMatriculas(sumSheet, matrColIdx);
        if (matriculas.isEmpty()) {
            log.info("No hay matriculas en '{}'; Resumen creado solo con cabecera.", sumSheetName);
        }

        Sheet sheet = workbook.createSheet(sheetName);
        writeSheet(workbook, sheet, sumSheetName, matrColumnName, matrColIdx,
                valueColumns, matriculas);

        // v1.8.0 — segunda tabla opcional (matriz Matrícula × Responsable)
        if (config.getBoolean("summary.byResponsible.enabled", false)) {
            int firstTableLastRow0 = sheet.getLastRowNum();
            int gapRows = Math.max(0,
                    config.getInt("summary.byResponsible.gapRows", DEFAULT_BY_RESP_GAP_ROWS));
            int startRow0 = firstTableLastRow0 + 1 + gapRows;
            int firstTableColsExcel = 1 + valueColumns.size();
            writeByResponsibleTable(workbook, sheet, startRow0, sumSheetName,
                    matrColIdx, matriculas, firstTableColsExcel);
        }

        int totalRows = sheet.getLastRowNum() + 1;
        report.addSheet(sheetName, totalRows);
        log.info("'{}' creada con {} matricula(s) y {} columna(s) de valor.",
                sheetName, matriculas.size(), valueColumns.size());

        // v1.8.0 — Forzar recalculo al abrir el fichero en Excel.
        //
        // Razon: POI escribe las celdas con fórmula sin valor cacheado
        // (<f>...</f><v></v>). LibreOffice y Google Sheets recalculan al
        // abrir igualmente. Excel tambien suele recalcular, pero en
        // SUMIFS con multiples criterios (como los de la segunda tabla,
        // 4 argumentos) hay escenarios reportados donde muestra 0 si el
        // workbook no lleva el flag fullCalcOnLoad=1 en calcPr.
        //
        // DerivedSheetBuilder ya setea este flag al final de su buildAll,
        // pero solo si derived.sheets no esta vacio. En el pipeline actual
        // derived.sheets=, asi que el flag no se ponia y solo las formulas
        // cacheadas implicitamente funcionaban. Lo seteamos aqui siempre
        // que se haya construido Resumen (idempotente: si DerivedSheet
        // tambien lo setea, da igual).
        workbook.setForceFormulaRecalculation(true);
    }

    // ==================================================================
    //  Construccion de la hoja
    // ==================================================================

    private void writeSheet(Workbook wb, Sheet sheet, String sumSheetName,
                            String matrColumnName, int matrColIdx,
                            List<ValueColumn> valueColumns, List<String> matriculas) {
        CellStyle titleStyle = StyleFactory.summaryBlockHeader(wb);
        CellStyle headerStyle = StyleFactory.summarySubHeaderGray(wb, true);
        CellStyle valueCellStyle = StyleFactory.summaryValueCell(wb);
        CellStyle numericCellStyle = StyleFactory.summaryNumericCell(wb);
        CellStyle totalsStyle = StyleFactory.summaryTotalCell(wb, true);

        int totalColsExcel = 1 + valueColumns.size();

        // Fila 1: titulo
        Row titleRow = sheet.createRow(ROW_TITLE - 1);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Resumen por " + matrColumnName);
        titleCell.setCellStyle(titleStyle);
        for (int c = 1; c < totalColsExcel; c++) {
            titleRow.createCell(c).setCellStyle(titleStyle);
        }
        if (totalColsExcel > 1) {
            sheet.addMergedRegion(new CellRangeAddress(
                    ROW_TITLE - 1, ROW_TITLE - 1, 0, totalColsExcel - 1));
        }

        // Fila 3: cabeceras
        Row headerRow = sheet.createRow(ROW_HEADER - 1);
        Cell hKey = headerRow.createCell(0);
        hKey.setCellValue(matrColumnName);
        hKey.setCellStyle(headerStyle);
        for (int i = 0; i < valueColumns.size(); i++) {
            Cell c = headerRow.createCell(1 + i);
            c.setCellValue(valueColumns.get(i).name);
            c.setCellStyle(headerStyle);
        }

        // Filas 4..N: datos (SUMIFS por matrícula)
        int maxRow = Math.max(2, config.getInt("summary.sumifsMaxRow", 10000));
        String quotedSum = PoiUtils.quoteSheetName(sumSheetName);
        String matrLetter = CellReference.convertNumToColString(matrColIdx);

        for (int i = 0; i < matriculas.size(); i++) {
            String matr = matriculas.get(i);
            int rowIdx0 = ROW_FIRST_DATA - 1 + i;
            Row r = sheet.createRow(rowIdx0);

            Cell matrCell = r.createCell(0);
            setNumericOrString(matrCell, matr);
            matrCell.setCellStyle(valueCellStyle);

            for (int j = 0; j < valueColumns.size(); j++) {
                ValueColumn vc = valueColumns.get(j);
                Cell c = r.createCell(1 + j);
                String formula = "SUMIFS("
                        + quotedSum + "!" + vc.letter + "2:" + vc.letter + maxRow + ","
                        + quotedSum + "!" + matrLetter + "2:" + matrLetter + maxRow + ","
                        + "$A" + (rowIdx0 + 1) + ")";
                c.setCellFormula(formula);
                c.setCellStyle(numericCellStyle);
            }
        }

        // Fila totales: "Total" + SUM de cada columna
        if (!matriculas.isEmpty()) {
            int totalsRowIdx0 = ROW_FIRST_DATA - 1 + matriculas.size();
            Row totalsRow = sheet.createRow(totalsRowIdx0);
            Cell totLabel = totalsRow.createCell(0);
            totLabel.setCellValue("Total");
            totLabel.setCellStyle(totalsStyle);

            int firstDataExcel = ROW_FIRST_DATA;
            int lastDataExcel = ROW_FIRST_DATA + matriculas.size() - 1;
            for (int j = 0; j < valueColumns.size(); j++) {
                String letter = CellReference.convertNumToColString(1 + j);
                Cell c = totalsRow.createCell(1 + j);
                c.setCellFormula("SUM(" + letter + firstDataExcel + ":" + letter + lastDataExcel + ")");
                c.setCellStyle(totalsStyle);
            }
        }

        // Ensanchar columnas para que se lea bien en Excel
        sheet.setColumnWidth(0, 20 * 256);
        for (int c = 1; c < totalColsExcel; c++) {
            sheet.setColumnWidth(c, 18 * 256);
        }
    }

    // ==================================================================
    //  v1.8.0 — Segunda tabla: matriz Matrícula × Responsable
    // ==================================================================

    /**
     * Construye la segunda tabla de Resumen, debajo de la primera, en
     * forma de matriz cruzada Matrícula (filas) × Responsable (columnas)
     * sumando la columna configurada en
     * {@code summary.byResponsible.valueColumn} (por defecto "PDCL").
     *
     * <p>Layout:</p>
     * <ul>
     *   <li>Fila {@code startRow0}: titulo (merge sobre todas las columnas).</li>
     *   <li>Fila {@code startRow0 + 2}: cabecera (esquina vacía +
     *       responsables en MAYUSCULAS + "Total").</li>
     *   <li>Filas siguientes: una por matricula; celdas = SUMIFS cruzando
     *       matricula (fila) × responsable (columna) sobre la misma hoja
     *       de datos que la primera tabla (Resultado).</li>
     *   <li>Fila final: "Total" + SUM por columna de responsable +
     *       gran total (SUM de la columna Total).</li>
     * </ul>
     *
     * <p>Normalizacion de responsables: se agrupa por codigo en
     * MAYUSCULAS ({@code Locale.ROOT}) tras {@code trim()}. El SUMIFS
     * de Excel es case-insensitive en criterios de texto, asi que una
     * unica celda clave en MAYUSCULAS suma correctamente todas las
     * variantes de capitalizacion del Excel original.</p>
     *
     * <p>Si la columna de responsable o la columna de valor configuradas
     * no existen en Resultado, se registra un warning y la tabla no se
     * genera (pero la primera tabla queda intacta).</p>
     */
    private void writeByResponsibleTable(Workbook wb, Sheet sheet, int startRow0,
                                         String sumSheetName, int matrColIdx,
                                         List<String> matriculas,
                                         int firstTableColsExcel) {
        Sheet sumSheet = wb.getSheet(sumSheetName);
        Row sumHeaderRow = sumSheet.getRow(0);

        String respColName = config.get("summary.byResponsible.column", DEFAULT_BY_RESP_COLUMN);
        String valueColName = config.get("summary.byResponsible.valueColumn", DEFAULT_BY_RESP_VALUE);
        String title = config.get("summary.byResponsible.title", DEFAULT_BY_RESP_TITLE);

        int respColIdx = PoiUtils.findColumnIndex(sumHeaderRow, respColName);
        if (respColIdx < 0) {
            log.warn("Columna '{}' no encontrada en '{}'. Segunda tabla omitida.",
                    respColName, sumSheetName);
            report.addWarning(WARN_CATEGORY_CABECERA,
                    "Columna '" + respColName + MSG_COL_NOT_FOUND_IN
                            + sumSheetName + "'. Tabla por responsable omitida.");
            return;
        }

        int valueColIdx = PoiUtils.findColumnIndex(sumHeaderRow, valueColName);
        if (valueColIdx < 0) {
            log.warn("Columna de valor '{}' no encontrada en '{}'. Segunda tabla omitida.",
                    valueColName, sumSheetName);
            report.addWarning(WARN_CATEGORY_CABECERA,
                    "Columna de valor '" + valueColName + MSG_COL_NOT_FOUND_IN
                            + sumSheetName + "'. Tabla por responsable omitida.");
            return;
        }

        List<String> responsibles = discoverResponsibles(sumSheet, respColIdx);
        if (responsibles.isEmpty()) {
            log.info("No hay responsables en '{}'; tabla por responsable omitida.", sumSheetName);
            report.addWarning("CONFIG",
                    "No se encontraron responsables en '" + sumSheetName
                            + "'. Tabla por responsable omitida.");
            return;
        }

        CellStyle titleStyle = StyleFactory.summaryBlockHeader(wb);
        CellStyle headerStyle = StyleFactory.summarySubHeaderGray(wb, true);
        CellStyle valueCellStyle = StyleFactory.summaryValueCell(wb);
        CellStyle numericCellStyle = StyleFactory.summaryNumericCell(wb);
        CellStyle totalsStyle = StyleFactory.summaryTotalCell(wb, true);

        // Nº de columnas de Excel que usa la tabla:
        //   1 (clave matricula) + N responsables + 1 (Total de fila)
        int totalColsExcel = 1 + responsibles.size() + 1;

        // Fila de titulo (merge)
        Row titleRow = sheet.createRow(startRow0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(title);
        titleCell.setCellStyle(titleStyle);
        for (int c = 1; c < totalColsExcel; c++) {
            titleRow.createCell(c).setCellStyle(titleStyle);
        }
        if (totalColsExcel > 1) {
            sheet.addMergedRegion(new CellRangeAddress(
                    startRow0, startRow0, 0, totalColsExcel - 1));
        }

        // Fila de cabeceras (separada del titulo por una fila en blanco,
        // igual que la primera tabla)
        int headerRow0 = startRow0 + 2;
        Row headerRow = sheet.createRow(headerRow0);
        // Esquina superior izquierda: vacia con estilo de cabecera
        Cell corner = headerRow.createCell(0);
        corner.setCellValue("");
        corner.setCellStyle(headerStyle);
        for (int i = 0; i < responsibles.size(); i++) {
            Cell c = headerRow.createCell(1 + i);
            c.setCellValue(responsibles.get(i));
            c.setCellStyle(headerStyle);
        }
        // Ultima columna: "Total"
        int totalColIdx0 = 1 + responsibles.size();
        Cell totalHeader = headerRow.createCell(totalColIdx0);
        totalHeader.setCellValue("Total");
        totalHeader.setCellStyle(headerStyle);

        // Filas de datos: una por matricula
        int firstDataRow0 = headerRow0 + 1;
        int maxRow = Math.max(2, config.getInt("summary.sumifsMaxRow", 10000));
        String quotedSum = PoiUtils.quoteSheetName(sumSheetName);
        String matrLetter = CellReference.convertNumToColString(matrColIdx);
        String respLetter = CellReference.convertNumToColString(respColIdx);
        String valueLetter = CellReference.convertNumToColString(valueColIdx);

        for (int i = 0; i < matriculas.size(); i++) {
            String matr = matriculas.get(i);
            int rowIdx0 = firstDataRow0 + i;
            Row r = sheet.createRow(rowIdx0);

            // Celda clave: matricula (siempre STRING por el mismo motivo
            // que en la primera tabla — ver setNumericOrString).
            Cell matrCell = r.createCell(0);
            setNumericOrString(matrCell, matr);
            matrCell.setCellStyle(valueCellStyle);

            // SUMIFS por cada responsable
            for (int j = 0; j < responsibles.size(); j++) {
                Cell c = r.createCell(1 + j);
                // La cabecera esta en fila Excel (headerRow0 + 1); la
                // referencia al responsable es absoluta en fila (fila de
                // cabecera fija) y relativa en columna (distinta por celda).
                String respHeaderRef = CellReference.convertNumToColString(1 + j)
                        + "$" + (headerRow0 + 1);
                String formula = "SUMIFS("
                        + quotedSum + "!" + valueLetter + "2:" + valueLetter + maxRow + ","
                        + quotedSum + "!" + matrLetter + "2:" + matrLetter + maxRow + ","
                        + "$A" + (rowIdx0 + 1) + ","
                        + quotedSum + "!" + respLetter + "2:" + respLetter + maxRow + ","
                        + respHeaderRef + ")";
                c.setCellFormula(formula);
                c.setCellStyle(numericCellStyle);
            }

            // Columna Total por fila: SUM sobre los responsables de esa fila
            Cell rowTotal = r.createCell(totalColIdx0);
            String firstRespLetter = CellReference.convertNumToColString(1);
            String lastRespLetter = CellReference.convertNumToColString(responsibles.size());
            rowTotal.setCellFormula("SUM(" + firstRespLetter + (rowIdx0 + 1)
                    + ":" + lastRespLetter + (rowIdx0 + 1) + ")");
            rowTotal.setCellStyle(totalsStyle);
        }

        // Fila final de totales por columna
        if (!matriculas.isEmpty()) {
            int totalsRow0 = firstDataRow0 + matriculas.size();
            Row totalsRow = sheet.createRow(totalsRow0);
            Cell totLabel = totalsRow.createCell(0);
            totLabel.setCellValue("Total");
            totLabel.setCellStyle(totalsStyle);

            int firstDataExcel = firstDataRow0 + 1;
            int lastDataExcel = firstDataRow0 + matriculas.size();
            // Total por cada responsable
            for (int j = 0; j < responsibles.size(); j++) {
                String letter = CellReference.convertNumToColString(1 + j);
                Cell c = totalsRow.createCell(1 + j);
                c.setCellFormula("SUM(" + letter + firstDataExcel
                        + ":" + letter + lastDataExcel + ")");
                c.setCellStyle(totalsStyle);
            }
            // Esquina inferior derecha: total global (SUM sobre la columna Total)
            Cell grandTotal = totalsRow.createCell(totalColIdx0);
            String totalLetter = CellReference.convertNumToColString(totalColIdx0);
            grandTotal.setCellFormula("SUM(" + totalLetter + firstDataExcel
                    + ":" + totalLetter + lastDataExcel + ")");
            grandTotal.setCellStyle(totalsStyle);
        }

        // Ancho de columnas: la primera tabla ya asigno 20 a col 0
        // (matricula) y 18 a las cols 1..firstTableColsExcel-1. No las
        // pisamos: solo ensanchamos las columnas nuevas a la derecha,
        // mas estrechas (14 unidades) porque las columnas de responsable
        // se acumulan y es facil tener muchas.
        for (int c = firstTableColsExcel; c <= totalColIdx0; c++) {
            sheet.setColumnWidth(c, 14 * 256);
        }

        report.addWarning("HOJA",
                "Resumen: anadida tabla '" + title + "' ("
                        + matriculas.size() + " matriculas x "
                        + responsibles.size() + " responsables).");
        log.info("Segunda tabla de Resumen anadida: {} matriculas x {} responsables.",
                matriculas.size(), responsibles.size());
    }

    // ==================================================================
    //  Auto-descubrimiento de matrículas en Resultado
    // ==================================================================

    /**
     * Recorre la columna de matriculas y devuelve los valores únicos no
     * vacios. Conserva el orden de aparición (LinkedHashSet) y después
     * ordena: primero las numéricas ascendentemente, después las no
     * numéricas en orden alfabético.
     */
    private static List<String> discoverMatriculas(Sheet sheet, int colIdx) {
        Set<String> seen = new LinkedHashSet<>();
        int last = sheet.getLastRowNum();
        for (int r = 1; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell c = row.getCell(colIdx);
            String s = cellToStringForKey(c);
            if (s == null) continue;
            s = s.trim();
            if (s.isEmpty()) continue;
            seen.add(s);
        }

        List<String> numeric = new ArrayList<>();
        List<String> others = new ArrayList<>();
        for (String s : seen) {
            if (isNumeric(s)) numeric.add(s);
            else others.add(s);
        }
        numeric.sort((a, b) -> {
            // Numerica por longitud y contenido: como todas son digit-only,
            // basta con comparar como long.
            return Long.compare(Long.parseLong(a), Long.parseLong(b));
        });
        others.sort(String::compareTo);

        List<String> out = new ArrayList<>(numeric.size() + others.size());
        out.addAll(numeric);
        out.addAll(others);
        return out;
    }

    /**
     * Recorre la columna de responsables y devuelve los codigos unicos
     * normalizados a MAYUSCULAS, ordenados alfabeticamente.
     *
     * <p>Normalizacion aplicada: {@code trim()} + {@code toUpperCase(Locale.ROOT)}.
     * Los duplicados por capitalizacion ("resp01", "RESP01", " Resp01 ")
     * se colapsan en una sola entrada ("RESP01") en la cabecera.</p>
     *
     * <p><b>Limite conocido sobre el SUMIFS emitido</b>: Excel SUMIFS es
     * case-insensitive en criterios de texto, asi que las variantes
     * "resp01" y "RESP01" del origen suman correctamente contra el
     * criterio "RESP01" puesto en la cabecera. Pero SUMIFS NO es
     * trim-insensitive: una variante " Resp01 " con espacios al principio
     * o final aparece como la misma cabecera (por el trim al descubrir),
     * pero su fila no se suma en esa columna porque el criterio "RESP01"
     * no casa contra " Resp01 ". El alcance pactado para 1.8.0 son
     * codigos alfanumericos sin espacios; el trim completo de los datos
     * origen queda para una iteracion futura.</p>
     */
    private static List<String> discoverResponsibles(Sheet sheet, int colIdx) {
        // LinkedHashMap para que, si aparecen dos variantes distintas,
        // nos quedemos con la primera en forma MAYUSCULAS (determinista).
        Map<String, String> normalizedToCanonical = new LinkedHashMap<>();
        int last = sheet.getLastRowNum();
        for (int r = 1; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell c = row.getCell(colIdx);
            String s = cellToStringForKey(c);
            if (s == null) continue;
            s = s.trim();
            if (s.isEmpty()) continue;
            String upper = s.toUpperCase(Locale.ROOT);
            normalizedToCanonical.putIfAbsent(upper, upper);
        }

        List<String> out = new ArrayList<>(normalizedToCanonical.values());
        out.sort(String::compareTo);
        return out;
    }

    /**
     * Obtiene el valor de la celda como cadena apta para usarse como
     * clave de matrícula. Para numéricas enteras devuelve "99641" en
     * lugar de "99641.0".
     */
    private static String cellToStringForKey(Cell cell) {
        if (cell == null) return null;
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            type = cell.getCachedFormulaResultType();
        }
        switch (type) {
            case STRING:  return cell.getStringCellValue();
            case NUMERIC:
                double d = cell.getNumericCellValue();
                return (d == (long) d) ? String.valueOf((long) d) : String.valueOf(d);
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case BLANK:   return null;
            default:      return null;
        }
    }

    // ==================================================================
    //  Helpers
    // ==================================================================

    private static boolean isNumeric(String s) {
        if (s == null || s.isEmpty()) return false;
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (ch < '0' || ch > '9') return false;
        }
        return true;
    }

    /**
     * Escribe el valor de una matrícula en la celda clave de Resumen.
     *
     * <p>Contrato crítico: <b>siempre</b> se escribe como STRING. Motivo
     * (v1.7.1): el SUMIFS que cada fila de valor lanza contra Resultado
     * compara esta celda clave contra la columna Matrícula de Resultado,
     * que tras el fix 1.6.2 ({@code asText.columns=Recurso}) es siempre
     * STRING. Si escribiéramos las matrículas todo-dígito como NUMERIC
     * (que era el comportamiento hasta 1.7.0 por estética — Excel las
     * alinea a la derecha), el SUMIFS compararía NUMERIC contra STRING
     * y daría 0 para todas las matrículas numéricas, mismo sintoma que
     * corrigió 1.6.2 pero en sentido inverso.</p>
     */
    private static void setNumericOrString(Cell cell, String value) {
        cell.setCellValue(value == null ? "" : value);
    }

    private static List<String> parseCsv(String raw) {
        List<String> out = new ArrayList<>();
        if (raw == null) return out;
        for (String part : raw.split(",")) {
            String t = part.trim();
            if (!t.isEmpty()) out.add(t);
        }
        // Itera para eliminar duplicados conservando orden
        Set<String> seen = new LinkedHashSet<>();
        Iterator<String> it = out.iterator();
        while (it.hasNext()) {
            if (!seen.add(it.next())) it.remove();
        }
        return out;
    }

    // ==================================================================
    //  Tipos internos
    // ==================================================================

    /** Par (nombre lógico, letra de columna en la hoja fuente). */
    private static final class ValueColumn {
        final String name;
        final String letter;
        ValueColumn(String name, String letter) {
            this.name = name;
            this.letter = letter;
        }
    }
}

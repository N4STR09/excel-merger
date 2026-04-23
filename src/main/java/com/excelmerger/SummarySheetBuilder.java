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
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/**
 * Construye la hoja "Resumen" (v1.6.0): sumatorio por matrícula de las
 * columnas numéricas de {@code Resultado} (Jira, REAL, PDCL, PDCL + Deuda
 * por defecto).
 *
 * <p>Estructura de la hoja generada:</p>
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

    /** Nombre por defecto de la columna clave en Resultado. */
    private static final String DEFAULT_MATRICULA_COLUMN = "Matrícula";
    /** Columnas agregadas por defecto. */
    private static final String DEFAULT_VALUE_COLUMNS = "Jira,REAL,PDCL,PDCL + Deuda";
    /** Nombre de la hoja origen por defecto. */
    private static final String DEFAULT_SUM_SHEET = "Resultado";
    /** Nombre de la hoja generada por defecto. */
    private static final String DEFAULT_SHEET_NAME = "Resumen";

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
            report.addWarning("CABECERA",
                    "'" + sumSheetName + "' sin cabeceras. Resumen omitido.");
            return;
        }

        int matrColIdx = PoiUtils.findColumnIndex(sumHeaderRow, matrColumnName);
        if (matrColIdx < 0) {
            log.warn("Columna '{}' no encontrada en '{}'. Resumen omitido.",
                    matrColumnName, sumSheetName);
            report.addWarning("CABECERA",
                    "Columna '" + matrColumnName + "' no encontrada en '"
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
                report.addWarning("CABECERA",
                        "Columna de valor '" + name + "' no encontrada en '"
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

        int totalRows = sheet.getLastRowNum() + 1;
        report.addSheet(sheetName, totalRows);
        log.info("'{}' creada con {} matricula(s) y {} columna(s) de valor.",
                sheetName, matriculas.size(), valueColumns.size());
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

    private static void setNumericOrString(Cell cell, String value) {
        if (value == null) {
            cell.setCellValue("");
            return;
        }
        if (isNumeric(value)) {
            cell.setCellValue(Long.parseLong(value));
        } else {
            cell.setCellValue(value);
        }
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

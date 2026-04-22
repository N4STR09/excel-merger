package com.excelmerger;

import com.excelmerger.util.PoiUtils;
import com.excelmerger.util.StyleFactory;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.util.LinkedHashSet;
import java.util.Properties;
import java.util.Set;
import java.util.TreeSet;

/**
 * Construye hojas adicionales ("derivadas") dentro del mismo libro resultado
 * de la fusion. Soporta dos tipos:
 *   FORMULAS    -> hoja con celdas libres (valores y formulas)
 *   AGGREGATION -> resumen agrupado generado automaticamente con formulas
 *                  SUMIF/AVERAGEIF/COUNTIF/MINIFS/MAXIFS.
 */
public class DerivedSheetBuilder {

    private static final Logger log = LoggerFactory.getLogger(DerivedSheetBuilder.class);

    public enum SheetType { FORMULAS, AGGREGATION }
    public enum AggFunction { SUM, AVG, COUNT, MIN, MAX }

    private final ConfigLoader config;
    private final RunReport report;

    public DerivedSheetBuilder(ConfigLoader config, RunReport report) {
        this.config = config;
        this.report = report;
    }

    /**
     * Lee la propiedad 'derived.sheets' y construye cada hoja declarada.
     */
    public void buildAll(Workbook workbook) {
        String list = config.get("derived.sheets", "").trim();
        if (list.isEmpty()) {
            log.info("No hay hojas derivadas configuradas.");
            return;
        }

        for (String rawId : list.split(",")) {
            String id = rawId.trim();
            if (id.isEmpty()) continue;
            buildOne(workbook, id);
        }

        // Al terminar, recalcular todas las formulas para que los valores
        // se vean ya calculados al abrir el Excel
        try {
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();
        } catch (Exception e) {
            log.info("{}", "[Derived] Aviso: no se pudieron evaluar todas las formulas: " + e.getMessage());
            report.addWarning("FORMULA",
                    "No se pudieron evaluar todas las formulas de hojas derivadas: " + e.getMessage());
        }

        // Forzar recalculo al abrir en Excel
        workbook.setForceFormulaRecalculation(true);
    }

    private void buildOne(Workbook workbook, String id) {
        String typeKey = "sheet." + id + ".type";
        String typeValue = config.get(typeKey, "FORMULAS");
        SheetType type;
        try {
            type = SheetType.valueOf(typeValue.toUpperCase());
        } catch (IllegalArgumentException e) {
            log.info("Tipo desconocido '" + typeValue + "' para hoja '" + id + "'. Omitida.");
            report.addWarning("CONFIG",
                    "Tipo '" + typeValue + "' desconocido para hoja derivada '" + id + "'. Omitida.");
            return;
        }

        if (workbook.getSheet(id) != null) {
            log.info("Aviso: ya existe una hoja con nombre '" + id + "'. Se omite la derivada.");
            report.addWarning("HOJA",
                    "Ya existe una hoja '" + id + "' en el libro; derivada omitida.");
            return;
        }

        Sheet sheet = workbook.createSheet(id);
        log.info("Creando hoja derivada '" + id + "' (" + type + ")");

        boolean built = false;
        switch (type) {
            case FORMULAS:
                buildFormulasSheet(sheet, id);
                built = true;
                break;
            case AGGREGATION:
                built = buildAggregationSheet(workbook, sheet, id);
                break;
        }

        autoSizeAll(sheet);

        if (built) {
            int rows = sheet.getLastRowNum() + 1;
            report.addSheet(id, rows);
        }
    }

    // --------------------------------------------------------------
    //  Tipo FORMULAS: celdas libres definidas una por una en el config
    // --------------------------------------------------------------
    private void buildFormulasSheet(Sheet sheet, String id) {
        String prefix = "sheet." + id + ".cell.";
        Properties props = config.getRawProperties();

        // Estilo de cabecera si existe una celda A1 tipo titulo
        CellStyle titleStyle = StyleFactory.title(sheet.getWorkbook());

        for (String key : new TreeSet<>(props.stringPropertyNames())) {
            if (!key.startsWith(prefix)) continue;
            String cellRef = key.substring(prefix.length()).trim();
            String rawValue = props.getProperty(key).trim();
            writeCell(sheet, id, cellRef, rawValue, titleStyle);
        }
    }

    private void writeCell(Sheet sheet, String id, String cellRef, String value, CellStyle titleStyle) {
        CellReference ref;
        try {
            ref = new CellReference(cellRef);
        } catch (Exception e) {
            log.info("{}", "[Derived] Referencia de celda invalida: " + cellRef);
            report.addWarning("CONFIG",
                    "Referencia de celda invalida '" + cellRef + "' en hoja derivada '" + id + "'.");
            return;
        }
        Row row = sheet.getRow(ref.getRow());
        if (row == null) row = sheet.createRow(ref.getRow());
        Cell cell = row.createCell(ref.getCol());

        if (value.startsWith("=")) {
            // Formula de Excel: quitamos el '=' inicial porque POI no lo quiere
            cell.setCellFormula(value.substring(1));
        } else {
            // Intentar guardar como numero si parece numerico; si no, como texto
            try {
                double numeric = Double.parseDouble(value);
                cell.setCellValue(numeric);
            } catch (NumberFormatException nfe) {
                cell.setCellValue(value);
            }
        }

        // Si es A1 y es texto, aplicar estilo de titulo
        if (ref.getRow() == 0 && ref.getCol() == 0 && !value.startsWith("=")) {
            try {
                Double.parseDouble(value);
            } catch (NumberFormatException nfe) {
                cell.setCellStyle(titleStyle);
            }
        }
    }

    // --------------------------------------------------------------
    //  Tipo AGGREGATION: resumen agrupado automatico
    // --------------------------------------------------------------
    private boolean buildAggregationSheet(Workbook workbook, Sheet target, String id) {
        String prefix = "sheet." + id + ".";

        String sourceSheetName = config.get(prefix + "sourceSheet", null);
        if (sourceSheetName == null || sourceSheetName.trim().isEmpty()) {
            report.addWarning("CONFIG",
                    "'" + prefix + "sourceSheet' requerido para AGGREGATION '" + id + "'. Omitida.");
            return false;
        }
        int firstDataRow = config.getInt(prefix + "firstDataRow", 2);  // 1-indexado
        String groupColRaw = config.get(prefix + "groupByColumn", null);
        String valueColRaw = config.get(prefix + "valueColumn", null);
        if (groupColRaw == null || valueColRaw == null) {
            report.addWarning("CONFIG",
                    "AGGREGATION '" + id + "' requiere 'groupByColumn' y 'valueColumn'. Omitida.");
            return false;
        }
        String groupCol = groupColRaw.trim().toUpperCase();
        String valueCol = valueColRaw.trim().toUpperCase();

        AggFunction agg;
        try {
            agg = AggFunction.valueOf(config.get(prefix + "aggregation", "SUM").toUpperCase());
        } catch (IllegalArgumentException e) {
            report.addWarning("CONFIG",
                    "Funcion de agregacion desconocida para '" + id + "': "
                            + config.get(prefix + "aggregation", "SUM") + ". Omitida.");
            return false;
        }
        String groupHeader = config.get(prefix + "groupHeader", "Grupo");
        String valueHeader = config.get(prefix + "valueHeader", agg.name());

        Sheet source = workbook.getSheet(sourceSheetName);
        if (source == null) {
            log.info("La hoja origen '" + sourceSheetName + "' no existe. Hoja omitida.");
            report.addWarning("HOJA",
                    "Hoja origen '" + sourceSheetName + "' no existe (AGGREGATION '" + id + "').");
            return false;
        }

        int groupColIdx = CellReference.convertColStringToIndex(groupCol);
        int lastRow = source.getLastRowNum() + 1; // 1-indexado para formulas
        if (lastRow < firstDataRow) {
            log.info("La hoja '" + sourceSheetName + "' no tiene datos. Hoja omitida.");
            report.addWarning("HOJA",
                    "Hoja origen '" + sourceSheetName + "' sin datos (AGGREGATION '" + id + "').");
            return false;
        }

        // 1. Recolectar grupos unicos en orden de aparicion
        Set<String> groups = new LinkedHashSet<>();
        for (int r = firstDataRow - 1; r <= source.getLastRowNum(); r++) {
            Row row = source.getRow(r);
            if (row == null) continue;
            Cell c = row.getCell(groupColIdx);
            if (c == null) continue;
            String key = cellToString(c);
            if (key != null && !key.isEmpty()) {
                groups.add(key);
            }
        }

        // 2. Escribir cabeceras
        CellStyle headerStyle = StyleFactory.header(workbook);
        Row header = target.createRow(0);
        Cell h1 = header.createCell(0); h1.setCellValue(groupHeader); h1.setCellStyle(headerStyle);
        Cell h2 = header.createCell(1); h2.setCellValue(valueHeader); h2.setCellStyle(headerStyle);

        // 3. Rango completo a usar en las formulas
        String groupRange = "'" + sourceSheetName + "'!$" + groupCol + "$" + firstDataRow + ":$" + groupCol + "$" + lastRow;
        String valueRange = "'" + sourceSheetName + "'!$" + valueCol + "$" + firstDataRow + ":$" + valueCol + "$" + lastRow;

        // 4. Una fila por grupo, con la formula correspondiente
        int rowIdx = 1;
        for (String group : groups) {
            Row r = target.createRow(rowIdx);
            String safeGroup = group.replace("\"", "\"\"");
            r.createCell(0).setCellValue(group);

            Cell valueCell = r.createCell(1);
            String formula = buildFormula(agg, groupRange, valueRange, safeGroup);
            valueCell.setCellFormula(formula);
            rowIdx++;
        }

        // 5. Fila de totales
        if (agg == AggFunction.SUM || agg == AggFunction.COUNT) {
            Row totalRow = target.createRow(rowIdx + 1);
            Cell lbl = totalRow.createCell(0);
            lbl.setCellValue("TOTAL");
            lbl.setCellStyle(headerStyle);
            Cell tot = totalRow.createCell(1);
            String totalFormula = (agg == AggFunction.COUNT)
                    ? "COUNTA(" + valueRange + ")"
                    : "SUM(" + valueRange + ")";
            tot.setCellFormula(totalFormula);
            tot.setCellStyle(headerStyle);
        }

        log.info("{}", groups.size() + " grupos generados con " + agg);
        return true;
    }

    private String buildFormula(AggFunction agg, String groupRange, String valueRange, String safeGroup) {
        switch (agg) {
            case SUM:
                return "SUMIF(" + groupRange + ",\"" + safeGroup + "\"," + valueRange + ")";
            case AVG:
                return "AVERAGEIF(" + groupRange + ",\"" + safeGroup + "\"," + valueRange + ")";
            case COUNT:
                return "COUNTIF(" + groupRange + ",\"" + safeGroup + "\")";
            case MIN:
                return "MINIFS(" + valueRange + "," + groupRange + ",\"" + safeGroup + "\")";
            case MAX:
                return "MAXIFS(" + valueRange + "," + groupRange + ",\"" + safeGroup + "\")";
            default:
                throw new IllegalArgumentException("Funcion no soportada: " + agg);
        }
    }

    // --------------------------------------------------------------
    //  Utilidades
    // --------------------------------------------------------------

    /**
     * Representacion en cadena de una celda para usar como clave de agrupacion.
     * Semantica propia de AGGREGATION: recorta espacios en STRING e interpreta
     * fechas en NUMERIC (difiere de {@link PoiUtils#cellAsString(Cell)}).
     */
    private String cellToString(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:  return cell.getStringCellValue().trim();
            case NUMERIC: return DateUtil.isCellDateFormatted(cell)
                    ? cell.getDateCellValue().toString()
                    : stripTrailingZero(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA: return cell.getCellFormula();
            default:      return "";
        }
    }

    private String stripTrailingZero(double d) {
        if (d == (long) d) return String.valueOf((long) d);
        return String.valueOf(d);
    }

    private void autoSizeAll(Sheet sheet) {
        int maxCol = PoiUtils.countColumns(sheet);
        for (int c = 0; c < maxCol; c++) {
            try { sheet.autoSizeColumn(c); } catch (Exception ignored) {}
        }
    }
}

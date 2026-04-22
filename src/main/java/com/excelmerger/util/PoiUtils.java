package com.excelmerger.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

/**
 * Utilidades comunes sobre el modelo de Apache POI.
 *
 * <p>Centraliza los helpers que antes estaban duplicados en varios builders
 * ({@code ExcelMerger}, {@code MesSheetBuilder}, {@code DerivedSheetBuilder},
 * {@code LookupSheetBuilder}). No tiene estado: todos los metodos son
 * estaticos y puros.</p>
 *
 * <p>Nota de semantica: {@link #cellAsString(Cell)} <b>no</b> recorta espacios
 * y <b>no</b> interpreta celdas numericas con formato de fecha. Los puntos del
 * codigo que necesitan ese comportamiento conservan su propio helper local
 * (p. ej. {@code DerivedSheetBuilder.cellToString} y
 * {@code FileProfileResolver.FileProfile.cellToString}).</p>
 */
public final class PoiUtils {

    private PoiUtils() {
        // Utility class
    }

    /**
     * Representacion en cadena de una celda para comparaciones y warnings.
     * Para NUMERIC formatea sin decimales si el valor es entero.
     * No recorta espacios. No interpreta fechas.
     * Devuelve {@code null} si la celda es {@code null}.
     */
    public static String cellAsString(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:  return cell.getStringCellValue();
            case NUMERIC:
                double d = cell.getNumericCellValue();
                return (d == (long) d) ? String.valueOf((long) d) : String.valueOf(d);
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA: return cell.getCellFormula();
            default:      return "";
        }
    }

    /**
     * Copia el valor de una celda origen en una celda destino, preservando el
     * tipo. No copia estilos. Para fechas usa {@link DateUtil#isCellDateFormatted}.
     */
    public static void copyCellValue(Cell source, Cell target) {
        switch (source.getCellType()) {
            case STRING:
                target.setCellValue(source.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(source)) {
                    target.setCellValue(source.getDateCellValue());
                } else {
                    target.setCellValue(source.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                target.setCellValue(source.getBooleanCellValue());
                break;
            case FORMULA:
                target.setCellFormula(source.getCellFormula());
                break;
            case BLANK:
                target.setBlank();
                break;
            case ERROR:
                target.setCellErrorValue(source.getErrorCellValue());
                break;
            default:
                target.setCellValue(source.toString());
        }
    }

    /**
     * {@code true} si la celda es {@code null}, de tipo BLANK, o una cadena
     * que solo contiene espacios.
     */
    public static boolean isBlank(Cell cell) {
        if (cell == null) return true;
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim().isEmpty();
            case BLANK:  return true;
            default:     return false;
        }
    }

    /**
     * Busca, en una fila de cabeceras, la columna cuyo texto (tras trim y
     * lowercase) coincide con {@code headerText}. Case-insensitive. Devuelve
     * el indice 0-based de la columna, o {@code -1} si no se encuentra.
     */
    public static int findColumnIndex(Row headerRow, String headerText) {
        if (headerRow == null || headerText == null) return -1;
        String target = headerText.trim().toLowerCase();
        for (int c = 0; c < headerRow.getLastCellNum(); c++) {
            Cell cell = headerRow.getCell(c);
            if (cell == null) continue;
            String value = cellAsString(cell);
            if (value != null && value.trim().toLowerCase().equals(target)) {
                return c;
            }
        }
        return -1;
    }

    /**
     * Devuelve la letra (A, B, ..., AA, AB, ...) de la columna cuyo valor de
     * cabecera coincide con {@code headerText} en la fila {@code headerRow0}
     * de la hoja. Devuelve {@code null} si no existe.
     */
    public static String columnLetter(Sheet sheet, int headerRow0, String headerText) {
        int idx = findColumnIndex(sheet.getRow(headerRow0), headerText);
        return idx < 0 ? null : CellReference.convertNumToColString(idx);
    }

    /**
     * Detecta heuristicamente la fila de cabeceras: entre las primeras 11
     * filas, devuelve el indice (0-based) de la que contiene mas celdas de
     * tipo STRING no vacias. Si no encuentra candidatas devuelve 0.
     */
    public static int detectHeaderRow(Sheet sheet) {
        int last = Math.min(sheet.getLastRowNum(), 10);
        int bestRow = 0;
        int bestCount = 0;
        for (int r = 0; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            int count = 0;
            for (int c = 0; c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                if (cell != null && cell.getCellType() == CellType.STRING
                        && !cell.getStringCellValue().trim().isEmpty()) {
                    count++;
                }
            }
            if (count > bestCount) {
                bestCount = count;
                bestRow = r;
            }
        }
        return bestRow;
    }

    /**
     * Devuelve el nombre de hoja listo para usar en una formula Excel.
     * Si solo contiene caracteres alfanumericos y guiones bajos, lo devuelve
     * tal cual. Si tiene espacios u otros caracteres especiales lo envuelve
     * en comillas simples y escapa las comillas internas.
     */
    public static String quoteSheetName(String sheetName) {
        if (sheetName.matches("[A-Za-z_][A-Za-z0-9_]*")) return sheetName;
        return "'" + sheetName.replace("'", "''") + "'";
    }

    /**
     * Devuelve el numero de columnas efectivas de la hoja, entendido como
     * el maximo {@code getLastCellNum()} entre todas sus filas no nulas.
     */
    public static int countColumns(Sheet sheet) {
        int max = 0;
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row != null && row.getLastCellNum() > max) {
                max = row.getLastCellNum();
            }
        }
        return max;
    }
}

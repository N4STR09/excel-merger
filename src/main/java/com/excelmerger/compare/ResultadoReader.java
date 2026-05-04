package com.excelmerger.compare;

import com.excelmerger.exception.InputValidationException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Lee {@code output/resultado_fusion.xlsx} y devuelve un mapa
 * {@code (Matricula, Petición, Funcion) -> PDCL+Deuda} para el comprobador
 * de discrepancias (Opcion 2 del menu, v3.1.0).
 *
 * <h2>Por que con FormulaEvaluator</h2>
 * <p>La columna {@code PDCL + Deuda} (col M, indice 12) en la hoja
 * {@code Resultado} contiene formulas tipo
 * {@code =L2 + IFERROR(SUMIFS(Deuda!$D:$D, ...), 0)}. Los valores
 * cacheados en el XLSX no son fiables: en pruebas reales muchas filas
 * tienen {@code null} en cache (el ultimo evaluador del archivo no las
 * recalculo). Sin {@link FormulaEvaluator} se leerian ceros falsos y se
 * generarian discrepancias inventadas. Por eso, igual que en
 * {@code EmptyRowFilter}, hacemos {@code evaluator.evaluateAll()} primero
 * y luego {@code evaluator.evaluate(cell)} por celda con try/catch.</p>
 *
 * <h2>Estructura esperada de la hoja Resultado</h2>
 * <pre>
 *   A: Petición   B: Aplicación   C: Equipo   D: Título   E: Departamento
 *   F: Matrícula  G: Funcion      H: Estado   I: Res. Tecnico  J: Jira
 *   K: Facturar   L: PDCL         M: PDCL + Deuda
 *   N: Horas_RealizadoTot   O: Horas_Mes
 * </pre>
 *
 * <h2>Politica ante claves duplicadas</h2>
 * <p>En teoria una matricula no deberia imputar la misma peticion con la
 * misma funcion dos veces, pero el Resultado puede contener duplicados (la
 * misma combinacion en mas de una fila). En ese caso se <b>suman</b> los
 * valores de {@code PDCL + Deuda}. Es el comportamiento conservador:
 * compara el total que el programa atribuye a esa combinacion contra el
 * total que el ERP imputa, sin perder filas.</p>
 */
public final class ResultadoReader {

    private static final Logger log = LoggerFactory.getLogger(ResultadoReader.class);

    /** Nombre de la hoja a leer. */
    public static final String SHEET_NAME = "Resultado";

    // Indices 0-based de las columnas que nos interesan.
    private static final int COL_PETICION  = 0;   // A
    private static final int COL_MATRICULA = 5;   // F
    private static final int COL_FUNCION   = 6;   // G
    private static final int COL_PDCL_DEUDA = 12; // M

    /** Cabeceras esperadas (validacion defensiva contra cambios de formato). */
    private static final String EXPECTED_HEADER_PETICION    = "Petición";
    private static final String EXPECTED_HEADER_MATRICULA   = "Matrícula";
    private static final String EXPECTED_HEADER_FUNCION     = "Funcion";
    private static final String EXPECTED_HEADER_PDCL_DEUDA  = "PDCL + Deuda";

    /** {@link DataFormatter} reutilizable para extraer celdas de texto. */
    private final DataFormatter dataFormatter = new DataFormatter();

    /**
     * Carga el Resultado desde el path indicado y devuelve el mapa de
     * cruces. El orden de insercion se preserva ({@link LinkedHashMap})
     * para facilitar la trazabilidad en logs.
     *
     * @throws InputValidationException si el fichero no existe, no se
     *         puede abrir, no contiene la hoja {@code Resultado} o sus
     *         cabeceras no son las esperadas.
     */
    public Map<DiscrepancyKey, Double> read(Path path) {
        if (path == null) {
            throw new InputValidationException("Path al Excel de fusion es null.");
        }
        if (!Files.exists(path)) {
            throw new InputValidationException(
                    "No se encuentra " + path + ". Ejecuta primero la Opcion 1 "
                            + "para generar el Excel de fusion.");
        }
        try (InputStream in = Files.newInputStream(path);
             Workbook wb = new XSSFWorkbook(in)) {
            return readFromWorkbook(wb, path.toString());
        } catch (IOException e) {
            throw new InputValidationException(
                    "Error leyendo " + path + ": " + e.getMessage(), e);
        }
    }

    /**
     * Variante para tests: lee directamente de un {@link Workbook} ya
     * cargado, evitando IO. {@code source} es solo para mensajes.
     */
    public Map<DiscrepancyKey, Double> read(Workbook wb, String source) {
        return readFromWorkbook(wb, source);
    }

    private Map<DiscrepancyKey, Double> readFromWorkbook(Workbook wb, String source) {
        Sheet sheet = wb.getSheet(SHEET_NAME);
        if (sheet == null) {
            throw new InputValidationException(
                    "El Excel " + source + " no contiene la hoja '" + SHEET_NAME + "'. "
                            + "Es probable que sea de una version anterior. Ejecuta "
                            + "la Opcion 1 para regenerarlo.");
        }
        validateHeaders(sheet, source);

        // Crear el evaluator y forzar la evaluacion de todas las formulas.
        // Sin esto, las celdas con cache nula devuelven null al evaluar
        // individualmente algunas SUMIFS cross-sheet. Mismo patron que
        // EmptyRowFilter.fillSheet (lecciones 1.7.x).
        FormulaEvaluator evaluator;
        try {
            evaluator = wb.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();
        } catch (RuntimeException e) {
            throw new InputValidationException(
                    "No se pudo evaluar las formulas de " + source
                            + ": " + e.getMessage(), e);
        }

        Map<DiscrepancyKey, Double> result = new LinkedHashMap<>();
        int lastRow = sheet.getLastRowNum();
        int rowsRead = 0;
        for (int r = 1; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            String peticion = readStringCell(row.getCell(COL_PETICION));
            String matricula = readStringCell(row.getCell(COL_MATRICULA));
            String funcion = readStringCell(row.getCell(COL_FUNCION));
            // Fila vacia / parcial: ignorar.
            if (peticion.isEmpty() && matricula.isEmpty() && funcion.isEmpty()) {
                continue;
            }
            if (peticion.isEmpty() || matricula.isEmpty() || funcion.isEmpty()) {
                log.warn("[ResultadoReader] {} fila {}: clave incompleta "
                        + "(peticion='{}', matricula='{}', funcion='{}'). Ignorada.",
                        source, r + 1, peticion, matricula, funcion);
                continue;
            }

            double pdclDeuda = evaluateNumericCell(row.getCell(COL_PDCL_DEUDA), evaluator);
            DiscrepancyKey key = new DiscrepancyKey(matricula, peticion, funcion);
            // Si ya existe la clave, sumamos (claves duplicadas en
            // Resultado se agregan).
            result.merge(key, pdclDeuda, Double::sum);
            rowsRead++;
        }
        log.info("[ResultadoReader] {}: {} filas con clave valida, {} entradas en el mapa.",
                source, rowsRead, result.size());
        return result;
    }

    /**
     * Verifica que las cabeceras de la fila 0 coinciden con las esperadas
     * en las posiciones que vamos a leer. Falla rapido si el formato del
     * Resultado ha cambiado, para evitar leer columnas equivocadas
     * silenciosamente.
     */
    private void validateHeaders(Sheet sheet, String source) {
        Row header = sheet.getRow(0);
        if (header == null) {
            throw new InputValidationException(
                    "Hoja '" + SHEET_NAME + "' de " + source + " sin cabeceras.");
        }
        checkHeader(header, COL_PETICION, EXPECTED_HEADER_PETICION, source);
        checkHeader(header, COL_MATRICULA, EXPECTED_HEADER_MATRICULA, source);
        checkHeader(header, COL_FUNCION, EXPECTED_HEADER_FUNCION, source);
        checkHeader(header, COL_PDCL_DEUDA, EXPECTED_HEADER_PDCL_DEUDA, source);
    }

    private void checkHeader(Row header, int col, String expected, String source) {
        String actual = readStringCell(header.getCell(col));
        if (!expected.equals(actual)) {
            throw new InputValidationException(
                    "Cabecera inesperada en hoja '" + SHEET_NAME + "' col "
                            + (col + 1) + " de " + source + ": esperada \""
                            + expected + "\", encontrada \"" + actual + "\".");
        }
    }

    /**
     * Lee una celda como string. Maneja celdas null, vacias, numericas
     * (formateadas con {@link DataFormatter}) o texto.
     */
    private String readStringCell(Cell cell) {
        if (cell == null) return "";
        // DataFormatter respeta el formato del Excel (Excel almacena como
        // texto las peticiones gracias al builder de la fusion). Devuelve
        // string vacia para BLANK.
        return dataFormatter.formatCellValue(cell).trim();
    }

    /**
     * Evalua una celda como numero (double). Tolera celdas null, blank,
     * formulas, errores y texto. Devuelve 0.0 ante incertidumbre.
     *
     * <p>Politica:</p>
     * <ul>
     *   <li>Celda null/blank/error -&gt; 0.0.</li>
     *   <li>Celda formula -&gt; evaluada con {@link FormulaEvaluator}.</li>
     *   <li>Celda numerica -&gt; valor directo.</li>
     *   <li>Celda string que parsea como double -&gt; ese valor.</li>
     *   <li>Celda string no parseable -&gt; 0.0 + warning log.</li>
     * </ul>
     */
    private static double evaluateNumericCell(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) return 0.0;
        CellType type = cell.getCellType();
        if (type == CellType.BLANK) return 0.0;
        if (type == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        if (type == CellType.STRING) {
            return parseStringAsDouble(cell.getStringCellValue());
        }
        if (type == CellType.FORMULA) {
            CellValue cv;
            try {
                cv = evaluator.evaluate(cell);
            } catch (RuntimeException e) {
                log.warn("[ResultadoReader] formula no evaluable en celda "
                        + "{}: {}. Asumido 0.0.",
                        cell.getAddress(), e.getMessage());
                return 0.0;
            }
            if (cv == null) return 0.0;
            switch (cv.getCellType()) {
                case NUMERIC:
                    return cv.getNumberValue();
                case STRING:
                    return parseStringAsDouble(cv.getStringValue());
                case BOOLEAN:
                    return cv.getBooleanValue() ? 1.0 : 0.0;
                case BLANK:
                case ERROR:
                default:
                    return 0.0;
            }
        }
        return 0.0;
    }

    private static double parseStringAsDouble(String s) {
        if (s == null) return 0.0;
        String t = s.trim();
        if (t.isEmpty()) return 0.0;
        try {
            return Double.parseDouble(t.replace(',', '.'));
        } catch (NumberFormatException e) {
            log.warn("[ResultadoReader] valor de texto no parseable como double: \"{}\". Asumido 0.0.", t);
            return 0.0;
        }
    }
}

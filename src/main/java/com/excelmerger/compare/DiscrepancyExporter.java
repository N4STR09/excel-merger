package com.excelmerger.compare;

import com.excelmerger.exception.OutputException;
import com.excelmerger.util.StyleFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * Escribe el Excel de salida del comprobador de discrepancias (Opcion 2
 * del menu, v3.1.0). Una hoja {@code Discrepancias} con cabeceras y una
 * fila por discrepancia.
 *
 * <h2>Estructura del Excel generado</h2>
 * <pre>
 *   A: Origen          (matricula del CSV de origen, p. ej. "90014")
 *   B: Tipo            (DIFERENCIA, SOLO_CSV, SOLO_RESULTADO)
 *   C: Petición
 *   D: Matrícula
 *   E: Función
 *   F: Realizado Horas (CSV)         vacio si SOLO_RESULTADO
 *   G: PDCL + Deuda    (Resultado)   vacio si SOLO_CSV
 *   H: Diferencia      (F - G)       vacio si SOLO_CSV o SOLO_RESULTADO
 * </pre>
 *
 * <p>Las celdas vacias se escriben como BLANK (no como 0) para que se
 * distingan visualmente del valor numerico cero, que tambien puede
 * aparecer legitimamente en F o G.</p>
 */
public final class DiscrepancyExporter {

    private static final Logger log = LoggerFactory.getLogger(DiscrepancyExporter.class);

    /** Nombre de la hoja del Excel de salida. */
    public static final String SHEET_NAME = "Discrepancias";

    /** Cabeceras (orden = orden de columnas en el Excel). */
    public static final List<String> HEADERS = List.of(
            "Origen", "Tipo", "Petición", "Matrícula", "Función",
            "Realizado Horas", "PDCL + Deuda", "Diferencia");

    /** Formato Excel para los valores numericos (2 decimales). */
    private static final String NUMERIC_FORMAT = "0.00";

    /**
     * Escribe el Excel de salida en {@code outPath}. Crea/sobrescribe.
     *
     * @throws OutputException si hay un error de IO al escribir.
     */
    public void write(Path outPath, List<Discrepancy> discrepancies) {
        if (outPath == null) {
            throw new OutputException("Path de salida null.");
        }
        try (Workbook wb = new XSSFWorkbook()) {
            buildWorkbook(wb, discrepancies);
            // Asegurarse de que el directorio padre existe.
            Path parent = outPath.getParent();
            if (parent != null) {
                Files.createDirectories(parent);
            }
            try (OutputStream os = Files.newOutputStream(outPath)) {
                wb.write(os);
            }
            log.info("[DiscrepancyExporter] Escrito {} con {} discrepancia(s).",
                    outPath, discrepancies.size());
        } catch (IOException e) {
            throw new OutputException(
                    "Error escribiendo Excel de discrepancias en " + outPath
                            + ": " + e.getMessage(), e);
        }
    }

    /**
     * Variante para tests: rellena un workbook ya creado en vez de
     * escribir a disco. El llamador es responsable de cerrar el workbook.
     */
    public void buildWorkbook(Workbook wb, List<Discrepancy> discrepancies) {
        Sheet sheet = wb.createSheet(SHEET_NAME);
        CellStyle headerStyle = StyleFactory.header(wb);
        CellStyle numberStyle = createNumberStyle(wb);

        // Fila 0: cabeceras.
        Row header = sheet.createRow(0);
        for (int i = 0; i < HEADERS.size(); i++) {
            Cell c = header.createCell(i);
            c.setCellValue(HEADERS.get(i));
            c.setCellStyle(headerStyle);
        }

        // Filas de datos.
        int r = 1;
        for (Discrepancy d : discrepancies) {
            Row row = sheet.createRow(r);
            r++;
            row.createCell(0).setCellValue(d.getOrigen() == null ? "" : d.getOrigen());
            row.createCell(1).setCellValue(d.getTipo().name());
            row.createCell(2).setCellValue(d.getPeticion() == null ? "" : d.getPeticion());
            row.createCell(3).setCellValue(d.getMatricula() == null ? "" : d.getMatricula());
            row.createCell(4).setCellValue(d.getFuncion() == null ? "" : d.getFuncion());

            writeNumericCellOrBlank(row, 5, d.getRealizadoHoras(), numberStyle);
            writeNumericCellOrBlank(row, 6, d.getPdclDeuda(), numberStyle);
            writeNumericCellOrBlank(row, 7, d.getDiferencia(), numberStyle);
        }

        // Auto-tamaño de columnas (cosmetico, ayuda a la legibilidad).
        for (int i = 0; i < HEADERS.size(); i++) {
            sheet.autoSizeColumn(i);
        }
        // Freeze de la fila de cabecera.
        sheet.createFreezePane(0, 1);
    }

    private CellStyle createNumberStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        DataFormat df = wb.createDataFormat();
        s.setDataFormat(df.getFormat(NUMERIC_FORMAT));
        return s;
    }

    /**
     * Escribe {@code value} en la columna {@code col} de {@code row}. Si
     * {@code value} es {@link Double#NaN}, deja la celda como BLANK.
     * Aplica el estilo numerico siempre.
     */
    private static void writeNumericCellOrBlank(Row row, int col, double value, CellStyle style) {
        Cell cell = row.createCell(col);
        if (Double.isNaN(value)) {
            // BLANK: no setear ningun valor. Aun asi aplicamos estilo
            // para que si en el futuro alguien edita la celda, herede el
            // formato numerico.
            cell.setCellStyle(style);
            return;
        }
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }
}

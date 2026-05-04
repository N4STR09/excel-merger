package com.excelmerger.compare;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/**
 * Helper para tests del paquete {@code compare}: construye un workbook en
 * memoria con la estructura que espera {@link ResultadoReader} (hoja
 * {@code Resultado} con las cabeceras correctas en las posiciones A, F,
 * G, M).
 *
 * <p>La columna {@code PDCL + Deuda} se rellena con una formula
 * {@code "=L<row> + 0"} (donde L es {@code PDCL}) para forzar que
 * {@link ResultadoReader} pase por la ruta {@code FormulaEvaluator}, igual
 * que en el Resultado real. Asi cubrimos en tests la lectura via
 * evaluador (regla inquebrantable 4).</p>
 */
final class ResultadoFixtureBuilder {

    private final List<Row3> rows = new ArrayList<>();

    private static final class Row3 {
        final String peticion;
        final String matricula;
        final String funcion;
        final double pdcl; // valor de la columna L (PDCL). PDCL+Deuda sera = L + 0.

        Row3(String peticion, String matricula, String funcion, double pdcl) {
            this.peticion = peticion;
            this.matricula = matricula;
            this.funcion = funcion;
            this.pdcl = pdcl;
        }
    }

    ResultadoFixtureBuilder add(String peticion, String matricula, String funcion, double pdclDeuda) {
        rows.add(new Row3(peticion, matricula, funcion, pdclDeuda));
        return this;
    }

    Workbook buildWorkbook() {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet(ResultadoReader.SHEET_NAME);

        // Cabeceras: igual que en el Resultado real.
        String[] headers = {"Petición", "Aplicación", "Equipo", "Título", "Departamento",
                "Matrícula", "Funcion", "Estado", "Res. Tecnico", "Jira",
                "Facturar", "PDCL", "PDCL + Deuda", "Horas_RealizadoTot", "Horas_Mes"};
        Row header = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            header.createCell(i).setCellValue(headers[i]);
        }

        for (int i = 0; i < rows.size(); i++) {
            Row3 r = rows.get(i);
            Row row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(r.peticion);  // A
            row.createCell(5).setCellValue(r.matricula); // F
            row.createCell(6).setCellValue(r.funcion);   // G
            row.createCell(11).setCellValue(r.pdcl);     // L (PDCL)
            // M (PDCL + Deuda) = formula =L<n> + 0. Forzamos formula
            // para que el test ejercite FormulaEvaluator.
            int rowNum = i + 2; // 1-based row number en Excel
            Cell cellM = row.createCell(12);
            cellM.setCellFormula("L" + rowNum + " + 0");
        }

        return wb;
    }

    /**
     * Escribe el workbook construido a un path temporal del tempDir
     * indicado. Cierra el workbook al terminar. Devuelve el path.
     */
    Path writeTo(Path tempDir, String fileName) throws IOException {
        Files.createDirectories(tempDir);
        Path out = tempDir.resolve(fileName);
        try (Workbook wb = buildWorkbook();
             OutputStream os = Files.newOutputStream(out)) {
            wb.write(os);
        }
        return out;
    }
}

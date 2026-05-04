package com.excelmerger.compare;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests unitarios de {@link DiscrepancyExporter}. v3.1.0.
 *
 * <p>Verifica la estructura del Excel generado: cabeceras correctas,
 * NaN -&gt; celda BLANK, freeze pane, formato numerico, conteo de filas.</p>
 */
class DiscrepancyExporterTest {

    @Test
    void escribeCabecerasEnFila0() throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {
            new DiscrepancyExporter().buildWorkbook(wb, Collections.emptyList());

            Sheet sheet = wb.getSheet(DiscrepancyExporter.SHEET_NAME);
            assertThat(sheet).isNotNull();
            Row header = sheet.getRow(0);
            for (int i = 0; i < DiscrepancyExporter.HEADERS.size(); i++) {
                assertThat(header.getCell(i).getStringCellValue())
                        .isEqualTo(DiscrepancyExporter.HEADERS.get(i));
            }
        }
    }

    @Test
    void escribeUnaFilaPorDiscrepancia() throws IOException {
        List<Discrepancy> discs = Arrays.asList(
                Discrepancy.difference("99001", new DiscrepancyKey("99001", "100002", "AN"), 10.0, 8.0),
                Discrepancy.onlyInCsv("99001", new DiscrepancyKey("99001", "100003", "OT"), 3.5),
                Discrepancy.onlyInResultado("99001", new DiscrepancyKey("99001", "100006", "OT"), 7.0));
        try (Workbook wb = new XSSFWorkbook()) {
            new DiscrepancyExporter().buildWorkbook(wb, discs);

            Sheet sheet = wb.getSheet(DiscrepancyExporter.SHEET_NAME);
            assertThat(sheet.getLastRowNum()).isEqualTo(3); // 0=header + 3 datos.
        }
    }

    @Test
    void diferenciaTieneTodasLasCeldasNumericasRellenas() throws IOException {
        Discrepancy d = Discrepancy.difference("99001", new DiscrepancyKey("99001", "100002", "AN"), 10.0, 8.0);
        try (Workbook wb = new XSSFWorkbook()) {
            new DiscrepancyExporter().buildWorkbook(wb, Collections.singletonList(d));
            Row row = wb.getSheet(DiscrepancyExporter.SHEET_NAME).getRow(1);

            assertThat(row.getCell(0).getStringCellValue()).isEqualTo("99001");
            assertThat(row.getCell(1).getStringCellValue()).isEqualTo("DIFERENCIA");
            assertThat(row.getCell(2).getStringCellValue()).isEqualTo("100002");
            assertThat(row.getCell(3).getStringCellValue()).isEqualTo("99001");
            assertThat(row.getCell(4).getStringCellValue()).isEqualTo("AN");
            assertThat(row.getCell(5).getNumericCellValue()).isEqualTo(10.0);
            assertThat(row.getCell(6).getNumericCellValue()).isEqualTo(8.0);
            assertThat(row.getCell(7).getNumericCellValue()).isEqualTo(2.0);
        }
    }

    @Test
    void soloCsvDejaPdclDeudaYDiferenciaComoBlank() throws IOException {
        Discrepancy d = Discrepancy.onlyInCsv("99001", new DiscrepancyKey("99001", "100003", "OT"), 3.5);
        try (Workbook wb = new XSSFWorkbook()) {
            new DiscrepancyExporter().buildWorkbook(wb, Collections.singletonList(d));
            Row row = wb.getSheet(DiscrepancyExporter.SHEET_NAME).getRow(1);

            // Realizado Horas (col 5) si tiene valor.
            assertThat(row.getCell(5).getNumericCellValue()).isEqualTo(3.5);
            // PDCL+Deuda (col 6) y Diferencia (col 7) son BLANK.
            assertThat(row.getCell(6).getCellType()).isEqualTo(CellType.BLANK);
            assertThat(row.getCell(7).getCellType()).isEqualTo(CellType.BLANK);
        }
    }

    @Test
    void soloResultadoDejaRealizadoYDiferenciaComoBlank() throws IOException {
        Discrepancy d = Discrepancy.onlyInResultado("99001", new DiscrepancyKey("99001", "100006", "OT"), 7.0);
        try (Workbook wb = new XSSFWorkbook()) {
            new DiscrepancyExporter().buildWorkbook(wb, Collections.singletonList(d));
            Row row = wb.getSheet(DiscrepancyExporter.SHEET_NAME).getRow(1);

            // Realizado Horas (col 5) y Diferencia (col 7) son BLANK.
            assertThat(row.getCell(5).getCellType()).isEqualTo(CellType.BLANK);
            // PDCL+Deuda (col 6) si.
            assertThat(row.getCell(6).getNumericCellValue()).isEqualTo(7.0);
            assertThat(row.getCell(7).getCellType()).isEqualTo(CellType.BLANK);
        }
    }

    @Test
    void elValorCeroEnRealizadoSeDistingueDeBlank() throws IOException {
        // Caso real importante: si CSV tiene "       ,0" (0.0), debe
        // escribirse como 0.0 (numerico), no como BLANK. Asi se distingue
        // visualmente del caso SOLO_RESULTADO.
        Discrepancy d = Discrepancy.difference("99001", new DiscrepancyKey("99001", "100004", "PR"), 0.0, 0.0);
        try (Workbook wb = new XSSFWorkbook()) {
            new DiscrepancyExporter().buildWorkbook(wb, Collections.singletonList(d));
            Row row = wb.getSheet(DiscrepancyExporter.SHEET_NAME).getRow(1);

            assertThat(row.getCell(5).getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(row.getCell(5).getNumericCellValue()).isEqualTo(0.0);
            assertThat(row.getCell(6).getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(row.getCell(6).getNumericCellValue()).isEqualTo(0.0);
        }
    }

    @Test
    void elFreezeDeCabeceraEstaActivado() throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {
            new DiscrepancyExporter().buildWorkbook(wb, Collections.emptyList());
            Sheet sheet = wb.getSheet(DiscrepancyExporter.SHEET_NAME);
            // PaneInformation: fila 1 congelada (rowSplit=1).
            assertThat(sheet.getPaneInformation()).isNotNull();
            // getHorizontalSplitTopRow devuelve short; promoverlo a int
            // evita el mismatch Short/Integer al asertar con AssertJ.
            assertThat((int) sheet.getPaneInformation().getHorizontalSplitTopRow()).isEqualTo(1);
        }
    }

    @Test
    void writeToDiskCreaElFichero(@TempDir Path tempDir) throws IOException {
        Path out = tempDir.resolve("subdir").resolve("disc.xlsx");
        new DiscrepancyExporter().write(out, Collections.singletonList(
                Discrepancy.onlyInCsv("99001", new DiscrepancyKey("99001", "100003", "OT"), 3.5)));
        assertThat(Files.exists(out)).isTrue();
        assertThat(Files.size(out)).isGreaterThan(0);

        // Reabrir y verificar que es un XLSX valido con la hoja esperada.
        try (InputStream in = Files.newInputStream(out);
             Workbook wb = new XSSFWorkbook(in)) {
            assertThat(wb.getSheet(DiscrepancyExporter.SHEET_NAME)).isNotNull();
        }
    }
}

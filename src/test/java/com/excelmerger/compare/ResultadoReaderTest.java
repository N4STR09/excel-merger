package com.excelmerger.compare;

import com.excelmerger.exception.InputValidationException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Tests unitarios de {@link ResultadoReader}. v3.1.0.
 *
 * <p>Cubre:</p>
 * <ul>
 *   <li>Lectura del Resultado fixture con celdas tipo formula
 *       ({@code =L<n> + 0}) — fuerza la ruta {@link org.apache.poi.ss.usermodel.FormulaEvaluator},
 *       que es la regla inquebrantable 4.</li>
 *   <li>Agregacion (suma) de claves duplicadas.</li>
 *   <li>Filas con clave incompleta: ignoradas con warning, no rompen el flujo.</li>
 *   <li>Errores: hoja Resultado ausente, cabecera inesperada, fichero ausente.</li>
 * </ul>
 */
class ResultadoReaderTest {

    @Test
    void leeFormulasViaFormulaEvaluator(@TempDir Path tempDir) throws IOException {
        ResultadoFixtureBuilder builder = new ResultadoFixtureBuilder()
                .add("100001", "99001", "RE", 5.0)
                .add("100002", "99001", "AN", 8.0)
                .add("100004", "99001", "PR", 0.0)
                .add("100005", "99001", "RE", 0.5);
        Path xlsx = builder.writeTo(tempDir, "resultado.xlsx");

        Map<DiscrepancyKey, Double> map = new ResultadoReader().read(xlsx);

        assertThat(map).hasSize(4);
        assertThat(map.get(new DiscrepancyKey("99001", "100001", "RE"))).isEqualTo(5.0);
        assertThat(map.get(new DiscrepancyKey("99001", "100002", "AN"))).isEqualTo(8.0);
        assertThat(map.get(new DiscrepancyKey("99001", "100004", "PR"))).isEqualTo(0.0);
        assertThat(map.get(new DiscrepancyKey("99001", "100005", "RE"))).isEqualTo(0.5);
    }

    @Test
    void agregaSumandoClavesDuplicadas(@TempDir Path tempDir) throws IOException {
        ResultadoFixtureBuilder builder = new ResultadoFixtureBuilder()
                .add("100001", "99001", "RE", 3.0)
                .add("100001", "99001", "RE", 2.0); // duplicada -> suma
        Path xlsx = builder.writeTo(tempDir, "resultado.xlsx");

        Map<DiscrepancyKey, Double> map = new ResultadoReader().read(xlsx);

        assertThat(map).hasSize(1);
        assertThat(map.get(new DiscrepancyKey("99001", "100001", "RE"))).isEqualTo(5.0);
    }

    @Test
    void ignoraFilasConClaveIncompleta(@TempDir Path tempDir) throws IOException {
        // Construyo a mano un workbook con una fila completa y otra con
        // matricula vacia.
        Path xlsx = tempDir.resolve("res.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet(ResultadoReader.SHEET_NAME);
            String[] headers = {"Petición", "Aplicación", "Equipo", "Título", "Departamento",
                    "Matrícula", "Funcion", "Estado", "Res. Tecnico", "Jira",
                    "Facturar", "PDCL", "PDCL + Deuda", "Horas_RealizadoTot", "Horas_Mes"};
            Row header = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
            }
            // Fila 1: completa.
            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("100001");
            r1.createCell(5).setCellValue("99001");
            r1.createCell(6).setCellValue("RE");
            r1.createCell(11).setCellValue(5.0);
            r1.createCell(12).setCellFormula("L2 + 0");
            // Fila 2: matricula vacia.
            Row r2 = sheet.createRow(2);
            r2.createCell(0).setCellValue("100002");
            r2.createCell(5).setCellValue("");      // matricula vacia
            r2.createCell(6).setCellValue("AN");
            r2.createCell(11).setCellValue(7.0);
            r2.createCell(12).setCellFormula("L3 + 0");
            // Fila 3: completamente vacia (todos los campos clave en blanco).
            sheet.createRow(3);
            try (OutputStream os = Files.newOutputStream(xlsx)) {
                wb.write(os);
            }
        }

        Map<DiscrepancyKey, Double> map = new ResultadoReader().read(xlsx);

        assertThat(map).hasSize(1);
        assertThat(map).containsKey(new DiscrepancyKey("99001", "100001", "RE"));
    }

    @Test
    void preservaOrdenDeInsercionEnElMapa(@TempDir Path tempDir) throws IOException {
        ResultadoFixtureBuilder builder = new ResultadoFixtureBuilder()
                .add("300", "99001", "RE", 1.0)
                .add("100", "99001", "RE", 2.0)
                .add("200", "99001", "RE", 3.0);
        Path xlsx = builder.writeTo(tempDir, "resultado.xlsx");

        Map<DiscrepancyKey, Double> map = new ResultadoReader().read(xlsx);

        // LinkedHashMap: orden de insercion.
        assertThat(map.keySet())
                .extracting(DiscrepancyKey::getPeticion)
                .containsExactly("300", "100", "200");
    }

    @Test
    void lanzaSiFicheroNoExiste(@TempDir Path tempDir) {
        Path missing = tempDir.resolve("nope.xlsx");
        assertThatThrownBy(() -> new ResultadoReader().read(missing))
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("Ejecuta primero la Opcion 1");
    }

    @Test
    void lanzaSiFaltaLaHojaResultado(@TempDir Path tempDir) throws IOException {
        Path xlsx = tempDir.resolve("noresultado.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            wb.createSheet("OtraHoja");
            try (OutputStream os = Files.newOutputStream(xlsx)) {
                wb.write(os);
            }
        }
        assertThatThrownBy(() -> new ResultadoReader().read(xlsx))
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("Resultado");
    }

    @Test
    void lanzaSiCabeceraInesperada(@TempDir Path tempDir) throws IOException {
        Path xlsx = tempDir.resolve("res.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet(ResultadoReader.SHEET_NAME);
            Row header = sheet.createRow(0);
            // Cabecera incorrecta en col A: ponemos "X" en vez de "Petición".
            header.createCell(0).setCellValue("X");
            header.createCell(5).setCellValue("Matrícula");
            header.createCell(6).setCellValue("Funcion");
            header.createCell(12).setCellValue("PDCL + Deuda");
            try (OutputStream os = Files.newOutputStream(xlsx)) {
                wb.write(os);
            }
        }
        assertThatThrownBy(() -> new ResultadoReader().read(xlsx))
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("Cabecera inesperada");
    }

    @Test
    void variantePorWorkbookEnMemoriaFunciona() {
        ResultadoFixtureBuilder builder = new ResultadoFixtureBuilder()
                .add("100001", "99001", "RE", 5.0);
        try (Workbook wb = builder.buildWorkbook()) {
            Map<DiscrepancyKey, Double> map = new ResultadoReader().read(wb, "in-memory");
            assertThat(map).hasSize(1);
        } catch (IOException e) {
            throw new AssertionError(e);
        }
    }
}

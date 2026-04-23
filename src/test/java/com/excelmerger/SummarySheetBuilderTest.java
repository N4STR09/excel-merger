package com.excelmerger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.Properties;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests de {@link SummarySheetBuilder} (v1.6.0). Construyen a mano un
 * {@link Workbook} con la hoja {@code Resultado} mínima en vez de correr
 * el pipeline completo: mantiene los casos aislados y permite cubrir
 * escenarios de error sin depender de fixtures reales.
 *
 * <p>Escenarios cubiertos:</p>
 * <ul>
 *   <li>summary.enabled=false: no crea la hoja.</li>
 *   <li>Colision con hoja existente: no crea la hoja + warning.</li>
 *   <li>Hoja Resultado inexistente: no crea + warning.</li>
 *   <li>Columna Matrícula inexistente en Resultado: no crea + warning.</li>
 *   <li>Columnas de valor parcialmente inexistentes: se omiten las faltantes + warning.</li>
 *   <li>Construccion basica: titulo, cabecera, datos y totales.</li>
 *   <li>Matriculas numericas y strings se mezclan (numericas primero).</li>
 *   <li>Formulas SUMIFS bien formadas con rangos acotados.</li>
 *   <li>Fila de totales con SUM(4:last).</li>
 *   <li>addSheet registra la hoja en RunReport.</li>
 * </ul>
 */
class SummarySheetBuilderTest {

    // ------------------------------------------------------------------
    //  Helpers de construccion
    // ------------------------------------------------------------------

    private static Properties baseConfig() {
        Properties p = new Properties();
        p.setProperty("summary.enabled", "true");
        p.setProperty("summary.sheetName", "Resumen");
        p.setProperty("summary.sumSheet", "Resultado");
        p.setProperty("summary.matriculaColumn", "Matrícula");
        p.setProperty("summary.valueColumns", "Jira,REAL,PDCL,PDCL + Deuda");
        p.setProperty("summary.sumifsMaxRow", "10000");
        return p;
    }

    /**
     * Monta un workbook con la hoja {@code Resultado} poblada:
     * cabeceras en fila 0, datos en filas 1..5.
     */
    private static Workbook buildResultadoWorkbook() {
        Workbook wb = new XSSFWorkbook();
        Sheet res = wb.createSheet("Resultado");
        writeRow(res, 0, "Petición", "Matrícula", "Jira", "REAL", "PDCL", "PDCL + Deuda");
        writeRow(res, 1, "P-001", "99641",       5.0,  6.0,  7.0,  7.5);
        writeRow(res, 2, "P-002", "99642",       3.0,  3.6,  4.0,  4.2);
        writeRow(res, 3, "P-003", "99641",       2.0,  2.4,  3.0,  3.0);
        writeRow(res, 4, "P-004", "-",           1.0,  1.2,  1.5,  1.5);
        writeRow(res, 5, "P-005", "Sin Matricula", 4.0, 4.8,  5.0,  5.0);
        return wb;
    }

    private static void writeRow(Sheet sheet, int rowIdx, Object... values) {
        Row r = sheet.createRow(rowIdx);
        for (int c = 0; c < values.length; c++) {
            Object v = values[c];
            Cell cell = r.createCell(c);
            if (v == null) continue;
            if (v instanceof Number) {
                cell.setCellValue(((Number) v).doubleValue());
            } else {
                cell.setCellValue(v.toString());
            }
        }
    }

    // ------------------------------------------------------------------
    //  Casos
    // ------------------------------------------------------------------

    @Test
    void resumenDeshabilitadoNoCreaHoja() throws Exception {
        Properties p = baseConfig();
        p.setProperty("summary.enabled", "false");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbook()) {
            new SummarySheetBuilder(cfg, report).build(wb);
            assertThat(wb.getSheet("Resumen")).isNull();
            assertThat(report.sheets()).doesNotContainKey("Resumen");
        }
    }

    @Test
    void hojaExistenteAvisaYNoSobrescribe() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfig());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbook()) {
            Sheet pre = wb.createSheet("Resumen");
            pre.createRow(0).createCell(0).setCellValue("manual");

            new SummarySheetBuilder(cfg, report).build(wb);

            assertThat(wb.getSheet("Resumen").getRow(0).getCell(0).getStringCellValue())
                    .isEqualTo("manual");
            assertThat(report.warnings())
                    .anyMatch(w -> "HOJA".equals(w.category) && w.message.contains("Resumen"));
        }
    }

    @Test
    void hojaResultadoInexistenteAvisaYNoCreaResumen() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfig());
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            assertThat(wb.getSheet("Resumen")).isNull();
            assertThat(report.warnings())
                    .anyMatch(w -> "HOJA".equals(w.category) && w.message.contains("Resultado"));
        }
    }

    @Test
    void columnaMatriculaInexistenteAvisaYNoCreaResumen() throws Exception {
        Properties p = baseConfig();
        p.setProperty("summary.matriculaColumn", "NoExiste");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbook()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            assertThat(wb.getSheet("Resumen")).isNull();
            assertThat(report.warnings())
                    .anyMatch(w -> "CABECERA".equals(w.category) && w.message.contains("NoExiste"));
        }
    }

    @Test
    void columnasDeValorParcialmenteInexistentesSeOmitenConWarning() throws Exception {
        Properties p = baseConfig();
        p.setProperty("summary.valueColumns", "Jira,NoExiste,PDCL");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbook()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            assertThat(resumen).isNotNull();
            // Cabeceras: "Matrícula", "Jira", "PDCL" (NoExiste omitida)
            Row hdr = resumen.getRow(2); // fila 3
            assertThat(hdr.getCell(0).getStringCellValue()).isEqualTo("Matrícula");
            assertThat(hdr.getCell(1).getStringCellValue()).isEqualTo("Jira");
            assertThat(hdr.getCell(2).getStringCellValue()).isEqualTo("PDCL");
            assertThat(hdr.getCell(3)).isNull();

            assertThat(report.warnings())
                    .anyMatch(w -> "CABECERA".equals(w.category) && w.message.contains("NoExiste"));
        }
    }

    @Test
    void construccionBasicaGeneraTituloCabeceraDatosYTotales() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfig());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbook()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            assertThat(resumen).isNotNull();
            // Fila 1 (idx 0): titulo
            assertThat(resumen.getRow(0).getCell(0).getStringCellValue())
                    .isEqualTo("Resumen por Matrícula");
            // Fila 3 (idx 2): cabeceras
            Row hdr = resumen.getRow(2);
            assertThat(hdr.getCell(0).getStringCellValue()).isEqualTo("Matrícula");
            assertThat(hdr.getCell(1).getStringCellValue()).isEqualTo("Jira");
            assertThat(hdr.getCell(2).getStringCellValue()).isEqualTo("REAL");
            assertThat(hdr.getCell(3).getStringCellValue()).isEqualTo("PDCL");
            assertThat(hdr.getCell(4).getStringCellValue()).isEqualTo("PDCL + Deuda");
        }
    }

    @Test
    void matriculasSeDescubrenOrdenandoNumericasPrimeroYStringsDespues() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfig());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbook()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            // Fila 4..7 (idx 3..6): 99641, 99642, -, Sin Matricula.
            //
            // v1.7.1: todas las matriculas se escriben como STRING (antes
            // las todo-digito se escribian como NUMERIC). El orden sigue
            // siendo el mismo: numericas ordenadas ascendente primero,
            // luego no numericas alfabetico. Lo que cambia es el tipo de
            // celda, necesario para que el SUMIFS case contra la columna
            // Matricula de Resultado (que es STRING por el fix 1.6.2).
            assertThat(resumen.getRow(3).getCell(0).getStringCellValue()).isEqualTo("99641");
            assertThat(resumen.getRow(4).getCell(0).getStringCellValue()).isEqualTo("99642");
            assertThat(resumen.getRow(5).getCell(0).getStringCellValue()).isEqualTo("-");
            assertThat(resumen.getRow(6).getCell(0).getStringCellValue()).isEqualTo("Sin Matricula");
        }
    }

    @Test
    void sumifsUsaRangosAcotadosSobreResultado() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfig());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbook()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            // Celda B4 (fila 4, col B): SUMIFS de Jira para la primera matrícula
            String formula = resumen.getRow(3).getCell(1).getCellFormula();
            assertThat(formula).startsWith("SUMIFS(");
            assertThat(formula).contains("Resultado!");
            assertThat(formula).contains("C2:C10000"); // Jira es la col C en Resultado
            assertThat(formula).contains("B2:B10000"); // Matrícula es la col B
            assertThat(formula).contains("$A4");
            assertThat(formula).doesNotContain(":C,");  // no column-complete
        }
    }

    @Test
    void filaTotalesApareceConSumDeCadaColumna() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfig());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbook()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            // 4 matriculas (99641, 99642, -, Sin Matricula) => totales en fila idx 7
            Row totals = resumen.getRow(7);
            assertThat(totals).isNotNull();
            assertThat(totals.getCell(0).getStringCellValue()).isEqualTo("Total");
            assertThat(totals.getCell(1).getCellFormula()).isEqualTo("SUM(B4:B7)");
            assertThat(totals.getCell(2).getCellFormula()).isEqualTo("SUM(C4:C7)");
            assertThat(totals.getCell(3).getCellFormula()).isEqualTo("SUM(D4:D7)");
            assertThat(totals.getCell(4).getCellFormula()).isEqualTo("SUM(E4:E7)");
        }
    }

    @Test
    void registraHojaEnRunReport() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfig());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbook()) {
            new SummarySheetBuilder(cfg, report).build(wb);
            assertThat(report.sheets()).containsKey("Resumen");
        }
    }

    @Test
    void sumifsDeMatriculaNumericaEvaluaCorrectamenteContraResultado() throws Exception {
        // Regresion v1.7.1 (bug C): antes del fix, las matriculas todo-digito
        // se escribian como NUMERIC en la columna clave del Resumen, mientras
        // que la columna Matricula de Resultado es STRING (por el fix 1.6.2).
        // El SUMIFS con criterio NUMERIC contra rango STRING no casa en Excel,
        // por lo que el total de Jira para una matricula numerica salia 0.
        //
        // Tras el fix, la celda clave es STRING y el SUMIFS suma correctamente.
        // En el workbook de test, la matricula 99641 aparece en las filas
        // P-001 (Jira=5) y P-003 (Jira=2), total esperado = 7.
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfig());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbook()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            // Fila 4 (idx 3): 99641. Verificamos tipo STRING y suma de Jira.
            Cell matrCell = resumen.getRow(3).getCell(0);
            assertThat(matrCell.getCellType())
                    .as("Celda clave de Resumen debe ser STRING para que el SUMIFS case")
                    .isEqualTo(CellType.STRING);
            assertThat(matrCell.getStringCellValue()).isEqualTo("99641");

            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            CellValue jiraTotal = evaluator.evaluate(resumen.getRow(3).getCell(1));
            assertThat(jiraTotal.getNumberValue())
                    .as("SUMIFS para matricula 99641 debe sumar 5 + 2 = 7 "
                            + "(regresion del bug C: NUMERIC vs STRING daba 0)")
                    .isEqualTo(7.0);
        }
    }
}

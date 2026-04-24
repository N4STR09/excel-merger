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

    // ==================================================================
    //  v1.8.0 — Tabla por responsable (matriz Matricula x Responsable)
    // ==================================================================

    /**
     * Añade las claves de la segunda tabla a una config base.
     */
    private static Properties baseConfigWithByResponsible() {
        Properties p = baseConfig();
        p.setProperty("summary.byResponsible.enabled", "true");
        p.setProperty("summary.byResponsible.column", "Res. Tecnico");
        p.setProperty("summary.byResponsible.valueColumn", "PDCL");
        p.setProperty("summary.byResponsible.title", "Totales Peticiones por Responsables Matrículas");
        p.setProperty("summary.byResponsible.gapRows", "2");
        return p;
    }

    /**
     * Workbook Resultado con columna Res. Tecnico adicional. Totales
     * documentados por (matricula, responsable, PDCL):
     *
     *   99641 / RESP_A -> P-001 (PDCL=7) + P-003 (PDCL=3)          = 10
     *   99641 / RESP_B ->                                           = 0
     *   99642 / RESP_A ->                                           = 0
     *   99642 / RESP_B -> P-002 (PDCL=4)                            = 4
     *   -     / RESP_A -> P-004 (PDCL=1.5)                          = 1.5
     *
     * Total por columna (responsable):
     *   RESP_A = 7 + 3 + 1.5 = 11.5
     *   RESP_B = 4             = 4
     *
     * Total por fila (matricula):
     *   99641 = 10
     *   99642 = 4
     *   -     = 1.5
     *
     * Gran total = 15.5 (y debe coincidir con el PDCL total de la tabla
     * por matricula).
     */
    private static Workbook buildResultadoWorkbookWithResponsible() {
        Workbook wb = new XSSFWorkbook();
        Sheet res = wb.createSheet("Resultado");
        writeRow(res, 0, "Petición", "Matrícula", "Jira", "REAL", "PDCL",
                "PDCL + Deuda", "Res. Tecnico");
        writeRow(res, 1, "P-001", "99641", 5.0, 6.0, 7.0, 7.5, "RESP_A");
        writeRow(res, 2, "P-002", "99642", 3.0, 3.6, 4.0, 4.2, "RESP_B");
        writeRow(res, 3, "P-003", "99641", 2.0, 2.4, 3.0, 3.0, "RESP_A");
        writeRow(res, 4, "P-004", "-",     1.0, 1.2, 1.5, 1.5, "RESP_A");
        return wb;
    }

    @Test
    void byResponsibleDeshabilitadoNoAnadeSegundaTabla() throws Exception {
        // Feature-flag off: la hoja Resumen se genera con la tabla de
        // matriculas como antes y nada mas (comportamiento v1.7.1).
        Properties p = baseConfig();
        p.setProperty("summary.byResponsible.enabled", "false");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbookWithResponsible()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            assertThat(resumen).isNotNull();

            // Tabla de matriculas: 4 matriculas (99641, 99642, -) = 3,
            // pero el workbook tiene 99641 repetida; total unicas = 3.
            // Fila de totales en rowIdx = 4-1 + 3 = 6 (fila 7 Excel).
            // Ninguna fila despues, porque la segunda tabla esta off.
            int last = resumen.getLastRowNum();
            Row totals = resumen.getRow(last);
            assertThat(totals.getCell(0).getStringCellValue())
                    .as("Con byResponsible=off, la ultima fila de Resumen "
                            + "debe ser la de totales de la primera tabla")
                    .isEqualTo("Total");
            assertThat(totals.getCell(1).getCellFormula()).startsWith("SUM(B4:");

            // Y no hay warning de "anadida tabla"
            assertThat(report.warnings())
                    .noneMatch(w -> w.message.contains("anadida tabla"));
        }
    }

    @Test
    void byResponsibleConstruyeMatrizConCabecerasCorrectas() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfigWithByResponsible());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbookWithResponsible()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            assertThat(resumen).isNotNull();

            // Primera tabla: titulo idx 0, cabeceras idx 2, 3 matriculas
            // idx 3..5, totales idx 6. Ultima fila primera tabla = 6.
            // gapRows=2 -> titulo segunda tabla en idx 6+1+2 = 9.
            // headerRow0 = 9 + 2 = 11.
            int titleRow0 = 9;
            int headerRow0 = 11;

            Row titleRow = resumen.getRow(titleRow0);
            assertThat(titleRow).isNotNull();
            assertThat(titleRow.getCell(0).getStringCellValue())
                    .isEqualTo("Totales Peticiones por Responsables Matrículas");

            Row headerRow = resumen.getRow(headerRow0);
            assertThat(headerRow).isNotNull();
            // Esquina superior izquierda: vacia.
            assertThat(headerRow.getCell(0).getStringCellValue()).isEmpty();
            // Responsables: RESP_A, RESP_B (alfabetico)
            assertThat(headerRow.getCell(1).getStringCellValue()).isEqualTo("RESP_A");
            assertThat(headerRow.getCell(2).getStringCellValue()).isEqualTo("RESP_B");
            // Ultima columna: Total
            assertThat(headerRow.getCell(3).getStringCellValue()).isEqualTo("Total");
        }
    }

    @Test
    void byResponsibleDescubreYNormalizaResponsablesAMayusculas() throws Exception {
        // El codigo del responsable en la fuente viene con capitalizacion
        // mixta ("resp01", "RESP01", " Resp01 "). La segunda tabla debe
        // colapsarlos en una sola columna "RESP01" (trim + toUpperCase).
        Workbook wb = new XSSFWorkbook();
        Sheet res = wb.createSheet("Resultado");
        writeRow(res, 0, "Petición", "Matrícula", "Jira", "REAL", "PDCL",
                "PDCL + Deuda", "Res. Tecnico");
        writeRow(res, 1, "P-001", "99641", 1.0, 1.2, 1.5, 1.5, "resp01");
        writeRow(res, 2, "P-002", "99641", 2.0, 2.4, 3.0, 3.0, "RESP01");
        writeRow(res, 3, "P-003", "99641", 4.0, 4.8, 6.0, 6.0, " Resp01 ");
        writeRow(res, 4, "P-004", "99641", 0.5, 0.6, 0.5, 0.5, "OTHER");

        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfigWithByResponsible());
        RunReport report = new RunReport();
        try {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            // Primera tabla: 1 matricula (99641) -> filas 3,4 (hdr) y 3 (data) y 4 (total)
            //   titulo=0, hdr=2, matr=3, total=4. Ultima fila=4.
            //   gapRows=2 -> titulo 2T = 4+1+2 = 7. hdr2T = 9. data = 10.
            int headerRow0 = 9;
            Row headerRow = resumen.getRow(headerRow0);
            // Deben ser exactamente 2 columnas de responsable: OTHER, RESP01
            // (alfabetico, en MAYUSCULAS) + 1 Total al final.
            assertThat(headerRow.getCell(0).getStringCellValue()).isEmpty();
            assertThat(headerRow.getCell(1).getStringCellValue()).isEqualTo("OTHER");
            assertThat(headerRow.getCell(2).getStringCellValue()).isEqualTo("RESP01");
            assertThat(headerRow.getCell(3).getStringCellValue()).isEqualTo("Total");
            // La fila de la 4a celda (cell4) no debe existir — solo hay 2 responsables.
            assertThat(headerRow.getCell(4)).isNull();
        } finally {
            wb.close();
        }
    }

    @Test
    void byResponsibleFormulasTienenFormaCorrectaYRangosAcotados() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfigWithByResponsible());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbookWithResponsible()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            // headerRow0 = 11 (ver test anterior). firstDataRow0 = 12.
            int firstDataRow0 = 12;
            Row firstDataRow = resumen.getRow(firstDataRow0);

            // Celda clave: matricula 99641 (primera, numerica mas chica)
            assertThat(firstDataRow.getCell(0).getStringCellValue()).isEqualTo("99641");

            // Celda SUMIFS para (99641, RESP_A): columna 1.
            // En el workbook del test Resultado tiene columnas:
            //   A=Peticion, B=Matricula, C=Jira, D=REAL, E=PDCL,
            //   F=PDCL+Deuda, G=Res. Tecnico.
            // Por tanto valueLetter=E, matrLetter=B, respLetter=G.
            // SUMIFS con rangos acotados 2:10000 y criterio:
            //   $A<rowIdx+1> para matricula (columna 1 de la propia fila de Resumen),
            //   B$<headerRow0+1> para responsable (cabecera RESP_A).
            String formula = firstDataRow.getCell(1).getCellFormula();
            assertThat(formula).startsWith("SUMIFS(");
            assertThat(formula).contains("Resultado!E2:E10000");   // rango de PDCL
            assertThat(formula).contains("Resultado!B2:B10000");   // rango de Matricula
            assertThat(formula).contains("Resultado!G2:G10000");   // rango de Res. Tecnico
            assertThat(formula).contains("$A" + (firstDataRow0 + 1)); // matricula de la fila
            assertThat(formula).contains("B$" + (11 + 1));          // cabecera RESP_A en B12

            // Y la columna Total por fila es SUM sobre los responsables de esa fila
            Cell rowTotal = firstDataRow.getCell(3);
            assertThat(rowTotal.getCellFormula())
                    .isEqualTo("SUM(B" + (firstDataRow0 + 1) + ":C" + (firstDataRow0 + 1) + ")");
        }
    }

    @Test
    void byResponsibleSumifsEvaluadoProduceValoresCorrectos() throws Exception {
        // Integracion de verdad: evalua las formulas con FormulaEvaluator
        // y comprueba los totales documentados en buildResultadoWorkbookWithResponsible().
        // Esta es la leccion de 1.7.1: inspeccionar el texto de la formula
        // no basta; hay que evaluarla.
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfigWithByResponsible());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbookWithResponsible()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            // Layout esperado:
            //   firstDataRow0 = 12 (99641), 13 (99642), 14 (-).
            //   Columna 1 = RESP_A, columna 2 = RESP_B, columna 3 = Total.
            //   Fila totales = 15.

            // 99641 / RESP_A = 7 + 3 = 10, 99641 / RESP_B = 0
            assertThat(evaluator.evaluate(resumen.getRow(12).getCell(1)).getNumberValue())
                    .as("99641 / RESP_A PDCL = 10").isEqualTo(10.0);
            assertThat(evaluator.evaluate(resumen.getRow(12).getCell(2)).getNumberValue())
                    .as("99641 / RESP_B PDCL = 0").isEqualTo(0.0);

            // 99642 / RESP_A = 0, 99642 / RESP_B = 4
            assertThat(evaluator.evaluate(resumen.getRow(13).getCell(1)).getNumberValue())
                    .as("99642 / RESP_A PDCL = 0").isEqualTo(0.0);
            assertThat(evaluator.evaluate(resumen.getRow(13).getCell(2)).getNumberValue())
                    .as("99642 / RESP_B PDCL = 4").isEqualTo(4.0);

            // - / RESP_A = 1.5, - / RESP_B = 0
            assertThat(evaluator.evaluate(resumen.getRow(14).getCell(1)).getNumberValue())
                    .as("- / RESP_A PDCL = 1.5").isEqualTo(1.5);
            assertThat(evaluator.evaluate(resumen.getRow(14).getCell(2)).getNumberValue())
                    .as("- / RESP_B PDCL = 0").isEqualTo(0.0);

            // Total por fila (columna Total)
            assertThat(evaluator.evaluate(resumen.getRow(12).getCell(3)).getNumberValue())
                    .as("99641 Total = 10").isEqualTo(10.0);
            assertThat(evaluator.evaluate(resumen.getRow(13).getCell(3)).getNumberValue())
                    .as("99642 Total = 4").isEqualTo(4.0);
            assertThat(evaluator.evaluate(resumen.getRow(14).getCell(3)).getNumberValue())
                    .as("- Total = 1.5").isEqualTo(1.5);

            // Fila Total (idx 15): total por columna responsable y gran total
            Row totalsRow = resumen.getRow(15);
            assertThat(totalsRow.getCell(0).getStringCellValue()).isEqualTo("Total");
            assertThat(evaluator.evaluate(totalsRow.getCell(1)).getNumberValue())
                    .as("Total RESP_A = 11.5").isEqualTo(11.5);
            assertThat(evaluator.evaluate(totalsRow.getCell(2)).getNumberValue())
                    .as("Total RESP_B = 4").isEqualTo(4.0);
            assertThat(evaluator.evaluate(totalsRow.getCell(3)).getNumberValue())
                    .as("Gran total = 15.5 (check cruzado con PDCL global)").isEqualTo(15.5);
        }
    }

    @Test
    void byResponsibleNormalizacionDeVariantesCapitalizacionSumaCorrectamente() throws Exception {
        // 2 variantes del mismo codigo con solo diferencia de capitalizacion
        // ("resp01", "RESP01") mas PDCL=1.5+3=4.5 para M=99641: la columna
        // "RESP01" de la segunda tabla debe evaluar a 4.5.
        //
        // Limite conocido: SUMIFS de Excel es case-insensitive en texto,
        // pero NO trim-insensitive. Una variante con espacios al principio
        // o final (" Resp01 ") sale como columna unica "RESP01" en la
        // cabecera (porque discoverResponsibles trima al descubrir), pero
        // su fila NO se suma en esa columna (el criterio SUMIFS es
        // "RESP01" y la celda en Resultado es " Resp01 ", lo cual no casa).
        // El escenario real pactado para 1.8.0 son codigos alfanumericos
        // sin espacios — solo diferencia de caso. El trim completo de
        // los datos origen queda fuera del alcance de esta iteracion.
        Workbook wb = new XSSFWorkbook();
        Sheet res = wb.createSheet("Resultado");
        writeRow(res, 0, "Petición", "Matrícula", "Jira", "REAL", "PDCL",
                "PDCL + Deuda", "Res. Tecnico");
        writeRow(res, 1, "P-001", "99641", 1.0, 1.2, 1.5, 1.5, "resp01");
        writeRow(res, 2, "P-002", "99641", 2.0, 2.4, 3.0, 3.0, "RESP01");

        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfigWithByResponsible());
        RunReport report = new RunReport();
        try {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            // Primera tabla: 1 matricula -> titulo=0,hdr=2,data=3,total=4.
            // gap=2 -> titulo 2T=7, hdr=9, data=10, total=11.
            // Celda 99641 / RESP01 = cell(1) de fila 10.
            assertThat(evaluator.evaluate(resumen.getRow(10).getCell(1)).getNumberValue())
                    .as("SUMIFS case-insensitive debe sumar las 2 variantes: "
                            + "resp01(1.5) + RESP01(3) = 4.5")
                    .isEqualTo(4.5);
        } finally {
            wb.close();
        }
    }

    @Test
    void byResponsibleEspaciosEnLosDatosOrigenNoCasanEnSumifs() throws Exception {
        // Caso limite documentado: si el Excel original trae un responsable
        // con espacios al principio/final (" Resp01 "), discoverResponsibles
        // lo normaliza y produce una unica columna "RESP01" en la cabecera
        // (porque el trim se aplica AL DESCUBRIR). Pero el SUMIFS que
        // despues emitimos compara "RESP01" contra la celda " Resp01 " de
        // Resultado, que tiene espacios, y NO casa (SUMIFS es
        // case-insensitive pero no trim-insensitive).
        //
        // Este test fija el comportamiento actual para no regresionarlo por
        // accidente si mañana alguien "arregla" discoverResponsibles. El
        // arreglo completo requiere trim en la capa de copia de datos,
        // fuera del alcance de 1.8.0.
        Workbook wb = new XSSFWorkbook();
        Sheet res = wb.createSheet("Resultado");
        writeRow(res, 0, "Petición", "Matrícula", "Jira", "REAL", "PDCL",
                "PDCL + Deuda", "Res. Tecnico");
        writeRow(res, 1, "P-001", "99641", 1.0, 1.2, 1.5, 1.5, "RESP01");
        writeRow(res, 2, "P-002", "99641", 4.0, 4.8, 6.0, 6.0, " Resp01 ");

        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfigWithByResponsible());
        RunReport report = new RunReport();
        try {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            // Solo 1 responsable descubierto (RESP01), no 2
            Row headerRow = resumen.getRow(9);
            assertThat(headerRow.getCell(1).getStringCellValue()).isEqualTo("RESP01");
            assertThat(headerRow.getCell(2).getStringCellValue()).isEqualTo("Total");

            // Y el SUMIFS para RESP01 solo suma la fila sin espacios (1.5),
            // NO la fila " Resp01 " (6.0). Total = 1.5, no 7.5.
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            assertThat(evaluator.evaluate(resumen.getRow(10).getCell(1)).getNumberValue())
                    .as("La fila con espacios no casa en SUMIFS; solo suma la variante exacta.")
                    .isEqualTo(1.5);
        } finally {
            wb.close();
        }
    }

    @Test
    void byResponsibleColumnaResponsableInexistenteEmiteWarningPerosNoRompePrimeraTabla() throws Exception {
        Properties p = baseConfigWithByResponsible();
        p.setProperty("summary.byResponsible.column", "NoExiste");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbookWithResponsible()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            // Primera tabla sigue ahi, intacta
            Sheet resumen = wb.getSheet("Resumen");
            assertThat(resumen).isNotNull();
            assertThat(resumen.getRow(0).getCell(0).getStringCellValue())
                    .isEqualTo("Resumen por Matrícula");
            // La fila "Total" de la primera tabla debe seguir siendo la ultima
            int last = resumen.getLastRowNum();
            assertThat(resumen.getRow(last).getCell(0).getStringCellValue())
                    .isEqualTo("Total");

            // Y warning registrado
            assertThat(report.warnings()).anyMatch(w ->
                    "CABECERA".equals(w.category) && w.message.contains("NoExiste"));
        }
    }

    @Test
    void byResponsibleGapRowsRespeta() throws Exception {
        // gapRows=0 debe pegar la segunda tabla justo despues de la primera
        Properties p = baseConfigWithByResponsible();
        p.setProperty("summary.byResponsible.gapRows", "0");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbookWithResponsible()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            Sheet resumen = wb.getSheet("Resumen");
            // Primera tabla: titulo=0, hdr=2, data=3..5, total=6.
            // Con gap=0 -> titulo 2T = 6+1+0 = 7.
            assertThat(resumen.getRow(7).getCell(0).getStringCellValue())
                    .isEqualTo("Totales Peticiones por Responsables Matrículas");
        }
    }

    @Test
    void byResponsibleRegistraWarningInformativoEnReport() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfigWithByResponsible());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbookWithResponsible()) {
            new SummarySheetBuilder(cfg, report).build(wb);

            assertThat(report.warnings()).anyMatch(w ->
                    "HOJA".equals(w.category)
                            && w.message.contains("anadida tabla")
                            && w.message.contains("matriculas")
                            && w.message.contains("responsables"));
        }
    }

    @Test
    void buildFuerzaRecalculoAlAbrirEnExcel() throws Exception {
        // Regresion de bug reportado en produccion (v1.8.0): los SUMIFS
        // de la segunda tabla mostraban 0 al abrir en Excel real aunque
        // POI y LibreOffice los evaluaban correctamente. Causa: POI
        // escribe celdas con fórmula sin valor cacheado, y Excel sin
        // fullCalcOnLoad=1 no siempre recalcula SUMIFS con 4 criterios.
        // El builder setea setForceFormulaRecalculation(true) para que
        // el calcPr salga con fullCalcOnLoad=1.
        ConfigLoader cfg = TestFixtures.configFromProperties(baseConfig());
        RunReport report = new RunReport();

        try (Workbook wb = buildResultadoWorkbook()) {
            assertThat(wb.getForceFormulaRecalculation())
                    .as("Precondicion: recien creado, el workbook no fuerza recalculo")
                    .isFalse();

            new SummarySheetBuilder(cfg, report).build(wb);

            assertThat(wb.getForceFormulaRecalculation())
                    .as("Tras construir Resumen, el workbook debe forzar recalculo "
                            + "al abrir (para que Excel real calcule los SUMIFS).")
                    .isTrue();
        }
    }
}

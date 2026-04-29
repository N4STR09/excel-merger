package com.excelmerger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Properties;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests v2.4.0 de {@link ResponsablesSheetBuilder} centrados en las
 * tablas pivot por hoja de responsable.
 *
 * <p><b>v2.7.0 (Modif 3)</b>: actualizado para reflejar el nuevo contrato:</p>
 * <ul>
 *   <li>La primera tabla pasa a ser PDCL (antes Jira).</li>
 *   <li>La segunda tabla pasa a ser Jira (antes Facturar).</li>
 *   <li>La columna fuente leida de Resultado pasa de Facturar a PDCL.
 *       En el fixture de tests el valor numerico de PDCL coincide con
 *       el que tenia Facturar (ambas formulas son {@code Jira*1.2}),
 *       asi que los valores esperados de la pivot no cambian.</li>
 * </ul>
 *
 * <p>Construye una hoja {@code Resultado} con cabeceras realistas
 * (Petición, Matrícula, Res. Tecnico, Jira, PDCL) y filas de datos
 * conocidas, y verifica:</p>
 * <ul>
 *   <li>Con {@code responsables.tables.enabled=true} (default), cada
 *       hoja de responsable contiene las dos pivots con la estructura
 *       descrita en la spec.</li>
 *   <li>Con {@code responsables.tables.enabled=false}, las hojas
 *       quedan como en v2.3.0 (solo cabecera A1).</li>
 *   <li>Combinaciones evaluadas con {@link FormulaEvaluator} dan los
 *       valores esperados.</li>
 *   <li>Si falta una columna en Resultado, las pivots se omiten para
 *       todas las hojas con un warning RESPONSABLE.</li>
 * </ul>
 */
class ResponsablesSheetBuilderV24Test {

    /**
     * Cabeceras de Resultado en orden A..E (las que el builder lee).
     * v2.7.0 (Modif 3): la columna E pasa de Facturar a PDCL.
     */
    private void writeResultadoHeader(Sheet sheet) {
        Row h = sheet.createRow(0);
        h.createCell(0).setCellValue("Petición");
        h.createCell(1).setCellValue("Matrícula");
        h.createCell(2).setCellValue("Res. Tecnico");
        h.createCell(3).setCellValue("Jira");
        h.createCell(4).setCellValue("PDCL");
    }

    private void addRow(Sheet sheet, String peticion, String matricula,
                        String responsable, double jira, double pdcl) {
        int next = sheet.getLastRowNum() + 1;
        Row r = sheet.createRow(next);
        r.createCell(0).setCellValue(peticion);
        r.createCell(1).setCellValue(matricula);
        r.createCell(2).setCellValue(responsable);
        r.createCell(3).setCellValue(jira);
        r.createCell(4).setCellValue(pdcl);
    }

    /**
     * Workbook minimo realista: 2 responsables, cada uno con 2 peticiones
     * y 2 matrículas, distribuidas para que el SUMIFS tenga cosas
     * distintas que sumar (no todo a 0).
     *
     * <p>v2.7.0 (Modif 3): los valores de la columna PDCL coinciden con
     * los que tenia Facturar en v2.4.0..v2.6.0 (ambas formulas son
     * {@code Jira*1.2}), por lo que los esperados numericos no cambian
     * — solo cambia el nombre de la columna y el orden de las tablas.</p>
     *
     * <pre>
     * tresp1@x:
     *   P-001 / M-1001 -> Jira=5  PDCL=6
     *   P-001 / M-1002 -> Jira=2  PDCL=2.4
     *   P-002 / M-1001 -> Jira=10 PDCL=12
     *   P-002 / M-1001 -> Jira=3  PDCL=3.6  (segunda imputacion misma combinacion)
     * tresp2@x:
     *   P-003 / M-1003 -> Jira=7  PDCL=8.4
     *   P-004 / M-1003 -> Jira=4  PDCL=4.8
     * </pre>
     *
     * <p>Esperados (tresp1@x, tabla Jira):</p>
     * <ul>
     *   <li>P-001 x M-1001 = 5</li>
     *   <li>P-001 x M-1002 = 2</li>
     *   <li>P-002 x M-1001 = 13 (10+3)</li>
     *   <li>P-002 x M-1002 = 0</li>
     *   <li>Total fila P-001 = 7, total fila P-002 = 13</li>
     *   <li>Total columna M-1001 = 18, total columna M-1002 = 2</li>
     *   <li>Gran total = 20</li>
     * </ul>
     */
    private Workbook buildRealisticWorkbook() {
        Workbook wb = new XSSFWorkbook();
        Sheet res = wb.createSheet("Resultado");
        writeResultadoHeader(res);
        addRow(res, "P-001", "M-1001", "tresp1@x",  5,  6);
        addRow(res, "P-001", "M-1002", "tresp1@x",  2,  2.4);
        addRow(res, "P-002", "M-1001", "tresp1@x", 10, 12);
        addRow(res, "P-002", "M-1001", "tresp1@x",  3,  3.6);
        addRow(res, "P-003", "M-1003", "tresp2@x",  7,  8.4);
        addRow(res, "P-004", "M-1003", "tresp2@x",  4,  4.8);
        return wb;
    }

    /** Configuración con tablas habilitadas (default v2.4.0). */
    private ConfigLoader configWithTablesEnabled() {
        Properties p = new Properties();
        p.setProperty("mes.sheetName", "Resultado");
        p.setProperty("summary.byResponsible.column", "Res. Tecnico");
        p.setProperty("responsables.tables.enabled", "true");
        p.setProperty("summary.sumifsMaxRow", "10000");
        return TestFixtures.configFromProperties(p);
    }

    /** Configuración con tablas deshabilitadas (comportamiento v2.3.0). */
    private ConfigLoader configWithTablesDisabled() {
        Properties p = new Properties();
        p.setProperty("mes.sheetName", "Resultado");
        p.setProperty("summary.byResponsible.column", "Res. Tecnico");
        p.setProperty("responsables.tables.enabled", "false");
        return TestFixtures.configFromProperties(p);
    }

    /** Configuración SIN la clave responsables.tables.enabled (debe asumirse true). */
    private ConfigLoader configDefault() {
        Properties p = new Properties();
        p.setProperty("mes.sheetName", "Resultado");
        p.setProperty("summary.byResponsible.column", "Res. Tecnico");
        return TestFixtures.configFromProperties(p);
    }

    // ==================================================================
    //  Estructura básica con tablas habilitadas (v2.7.0 — orden invertido)
    // ==================================================================

    @Test
    void hojaResponsableTieneDosTablasPivotApiladasConGap() throws java.io.IOException {
        try (Workbook wb = buildRealisticWorkbook()) {
            new ResponsablesSheetBuilder(configWithTablesEnabled(), new RunReport()).buildAll(wb);

            Sheet sheet = wb.getSheet("tresp1@x");
            assertThat(sheet).isNotNull();

            // Fila 1 (0-based 0): cabecera A1
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("tresp1@x");

            // v2.7.0 (Modif 3): primera tabla = PDCL.
            // Fila 3 (0-based 2): título tabla PDCL
            String t1 = sheet.getRow(2).getCell(0).getStringCellValue();
            assertThat(t1).as("primera tabla v2.7.0 = PDCL").contains("PDCL");
            assertThat(t1).as("primera tabla NO debe contener 'Jira'").doesNotContain("Jira");

            // Para tresp1@x: 2 peticiones x 2 matriculas
            // Layout v2.7.0:
            //   row 2 (Excel 3):  título PDCL  (antes Jira)
            //   row 3 (Excel 4):  cabecera (Petición | M-1001 | M-1002 | Total)
            //   row 4 (Excel 5):  P-001
            //   row 5 (Excel 6):  P-002
            //   row 6 (Excel 7):  Total
            //   row 7-8         : gap (2 filas en blanco)
            //   row 9 (Excel 10): título Jira  (antes Facturar)
            //   row 10..14      : cabecera+datos+total Jira

            // Title de la segunda tabla con gap=2: row = 6 + 1 + 2 = 9
            String t2 = sheet.getRow(9).getCell(0).getStringCellValue();
            assertThat(t2).as("segunda tabla v2.7.0 = Jira").contains("Jira");
            // Sanity: no debe quedar referencia a 'Facturar' en titulos.
            assertThat(t2).as("segunda tabla NO debe contener 'Facturar'").doesNotContain("Facturar");

            // Fila cabecera (row 10): Petición + matrículas + Total
            Row hdr2 = sheet.getRow(10);
            assertThat(hdr2.getCell(0).getStringCellValue()).isEqualTo("Petición");
            assertThat(hdr2.getCell(1).getStringCellValue()).isEqualTo("M-1001");
            assertThat(hdr2.getCell(2).getStringCellValue()).isEqualTo("M-1002");
            assertThat(hdr2.getCell(3).getStringCellValue()).isEqualTo("Total");
        }
    }

    @Test
    void formulaEvaluatorDevuelveValoresEsperadosEnPrimeraTablaPDCL()
            throws java.io.IOException {
        try (Workbook wb = buildRealisticWorkbook()) {
            new ResponsablesSheetBuilder(configWithTablesEnabled(), new RunReport()).buildAll(wb);

            Sheet sheet = wb.getSheet("tresp1@x");
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();

            // v2.7.0: la PRIMERA tabla es ahora PDCL (antes era Jira).
            // Layout: row 2 = titulo, row 3 = cabecera, row 4 = P-001,
            //         row 5 = P-002, row 6 = Total.
            //         col 1 = M-1001, col 2 = M-1002, col 3 = Total.
            //
            // Esperados PDCL = Jira * 1.2 segun fixture:
            //   P-001/M-1001 = 6
            //   P-001/M-1002 = 2.4
            //   P-002/M-1001 = 12 + 3.6 = 15.6
            //   Gran total   = 6 + 2.4 + 15.6 + 0 = 24

            assertThat(ev.evaluate(sheet.getRow(4).getCell(1)).getNumberValue())
                    .as("PDCL P-001 x M-1001").isEqualTo(6.0);
            assertThat(ev.evaluate(sheet.getRow(4).getCell(2)).getNumberValue())
                    .as("PDCL P-001 x M-1002").isEqualTo(2.4);
            assertThat(ev.evaluate(sheet.getRow(5).getCell(1)).getNumberValue())
                    .as("PDCL P-002 x M-1001").isEqualTo(15.6);
            assertThat(ev.evaluate(sheet.getRow(5).getCell(2)).getNumberValue())
                    .as("PDCL P-002 x M-1002 (sin filas)").isZero();

            // Gran total
            assertThat(ev.evaluate(sheet.getRow(6).getCell(3)).getNumberValue())
                    .as("Gran total PDCL").isEqualTo(24.0);
        }
    }

    @Test
    void formulaEvaluatorDevuelveValoresEsperadosEnSegundaTablaJira()
            throws java.io.IOException {
        try (Workbook wb = buildRealisticWorkbook()) {
            new ResponsablesSheetBuilder(configWithTablesEnabled(), new RunReport()).buildAll(wb);

            Sheet sheet = wb.getSheet("tresp1@x");
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();

            // v2.7.0: la SEGUNDA tabla es ahora Jira (antes era Facturar).
            // Con gap=2, empieza en row 9 (titulo), 10 (cabecera), 11 (P-001),
            // 12 (P-002), 13 (Total).
            //
            // Esperados Jira segun fixture:
            //   P-001/M-1001 = 5
            //   P-001/M-1002 = 2
            //   P-002/M-1001 = 10 + 3 = 13
            //   Gran total   = 5 + 2 + 13 + 0 = 20

            assertThat(ev.evaluate(sheet.getRow(11).getCell(1)).getNumberValue())
                    .as("Jira P-001 x M-1001").isEqualTo(5.0);
            assertThat(ev.evaluate(sheet.getRow(11).getCell(2)).getNumberValue())
                    .as("Jira P-001 x M-1002").isEqualTo(2.0);
            assertThat(ev.evaluate(sheet.getRow(12).getCell(1)).getNumberValue())
                    .as("Jira P-002 x M-1001").isEqualTo(13.0);
            assertThat(ev.evaluate(sheet.getRow(12).getCell(2)).getNumberValue())
                    .as("Jira P-002 x M-1002 (sin filas)").isZero();

            // Totales fila
            assertThat(ev.evaluate(sheet.getRow(11).getCell(3)).getNumberValue())
                    .as("Total fila P-001 (Jira)").isEqualTo(7.0);
            assertThat(ev.evaluate(sheet.getRow(12).getCell(3)).getNumberValue())
                    .as("Total fila P-002 (Jira)").isEqualTo(13.0);

            // Totales columna (row 13)
            assertThat(ev.evaluate(sheet.getRow(13).getCell(1)).getNumberValue())
                    .as("Total col M-1001 (Jira)").isEqualTo(18.0);
            assertThat(ev.evaluate(sheet.getRow(13).getCell(2)).getNumberValue())
                    .as("Total col M-1002 (Jira)").isEqualTo(2.0);

            // Gran total
            assertThat(ev.evaluate(sheet.getRow(13).getCell(3)).getNumberValue())
                    .as("Gran total Jira").isEqualTo(20.0);
        }
    }

    @Test
    void hojaTresp2NoIncluyePeticionesDeOtrosResponsables() throws java.io.IOException {
        try (Workbook wb = buildRealisticWorkbook()) {
            new ResponsablesSheetBuilder(configWithTablesEnabled(), new RunReport()).buildAll(wb);

            Sheet sheet = wb.getSheet("tresp2@x");
            assertThat(sheet).isNotNull();

            // tresp2@x tiene P-003 y P-004 con matrícula M-1003.
            // En la cabecera de la PRIMERA tabla (PDCL en v2.7.0, row 3) la
            // unica matricula es M-1003.
            Row hdrPDCL = sheet.getRow(3);
            assertThat(hdrPDCL.getCell(0).getStringCellValue()).isEqualTo("Petición");
            assertThat(hdrPDCL.getCell(1).getStringCellValue()).isEqualTo("M-1003");
            // col 2 es la columna Total (única matrícula -> 1 + 1)
            assertThat(hdrPDCL.getCell(2).getStringCellValue()).isEqualTo("Total");

            // Data: row 4 = P-003, row 5 = P-004, row 6 = Total
            assertThat(sheet.getRow(4).getCell(0).getStringCellValue()).isEqualTo("P-003");
            assertThat(sheet.getRow(5).getCell(0).getStringCellValue()).isEqualTo("P-004");

            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            // PDCL P-003 x M-1003 = 8.4
            assertThat(ev.evaluate(sheet.getRow(4).getCell(1)).getNumberValue()).isEqualTo(8.4);
            // PDCL P-004 x M-1003 = 4.8
            assertThat(ev.evaluate(sheet.getRow(5).getCell(1)).getNumberValue()).isEqualTo(4.8);
            // Gran total PDCL = 13.2
            assertThat(ev.evaluate(sheet.getRow(6).getCell(2)).getNumberValue()).isEqualTo(13.2);
        }
    }

    // ==================================================================
    //  Comportamiento v2.3.0 cuando enabled=false
    // ==================================================================

    @Test
    void conTablasDeshabilitadasHojaSoloTieneCabeceraA1() throws java.io.IOException {
        try (Workbook wb = buildRealisticWorkbook()) {
            new ResponsablesSheetBuilder(configWithTablesDisabled(), new RunReport()).buildAll(wb);

            Sheet sheet = wb.getSheet("tresp1@x");
            assertThat(sheet).isNotNull();
            // Solo fila 0 (A1).
            assertThat(sheet.getLastRowNum()).isZero();
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("tresp1@x");
        }
    }

    @Test
    void claveAusenteImplicaTablasHabilitadas() throws java.io.IOException {
        try (Workbook wb = buildRealisticWorkbook()) {
            new ResponsablesSheetBuilder(configDefault(), new RunReport()).buildAll(wb);

            Sheet sheet = wb.getSheet("tresp1@x");
            assertThat(sheet).isNotNull();
            // Hoja tiene más de una fila — las pivots se han generado.
            assertThat(sheet.getLastRowNum()).isGreaterThan(5);
            // v2.7.0: Fila 2 contiene el título de la primera pivot (PDCL).
            assertThat(sheet.getRow(2).getCell(0).getStringCellValue())
                    .contains("PDCL");
        }
    }

    // ==================================================================
    //  RunReport
    // ==================================================================

    @Test
    void runReportRegistraResumenDeTablasYSumifs() throws java.io.IOException {
        try (Workbook wb = buildRealisticWorkbook()) {
            RunReport report = new RunReport();
            new ResponsablesSheetBuilder(configWithTablesEnabled(), report).buildAll(wb);

            // Resumen final
            List<RunReport.Warning> hojaWarnings = report.warnings().stream()
                    .filter(w -> "HOJA".equals(w.category))
                    .toList();
            assertThat(hojaWarnings)
                    .anyMatch(w -> w.message.contains("Responsables")
                            && w.message.contains("hoja(s)")
                            && w.message.contains("tabla(s)")
                            && w.message.contains("SUMIFS"));

            // RunReport.sheets() debe tener el conteo real (no 1)
            // tresp1@x: 2 peticiones, 2 matriculas → 5 filas/tabla + gap 2 + 5 = 14 filas
            //   layout: 0(A1), 1(blank), 2(title1), 3(hdr1), 4-5(data1), 6(total1),
            //           7-8(gap), 9(title2), 10(hdr2), 11-12(data2), 13(total2) = 14 rows
            assertThat(report.sheets().get("tresp1@x"))
                    .as("filas en tresp1@x").isEqualTo(14);
        }
    }

    // ==================================================================
    //  Caso degenerado: columna ausente en Resultado
    // ==================================================================

    @Test
    void columnaPDCLAusenteEnResultadoOmitePivotsConWarning() throws java.io.IOException {
        // v2.7.0 (Modif 3): la columna fuente de la primera pivot es PDCL
        // (antes Facturar). Sin PDCL en Resultado, las pivots no deben
        // generarse y el warning debe mencionar la columna que falta.
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet res = wb.createSheet("Resultado");
            Row h = res.createRow(0);
            h.createCell(0).setCellValue("Petición");
            h.createCell(1).setCellValue("Matrícula");
            h.createCell(2).setCellValue("Res. Tecnico");
            h.createCell(3).setCellValue("Jira");
            // SIN PDCL — solo Jira.
            Row r1 = res.createRow(1);
            r1.createCell(0).setCellValue("P-001");
            r1.createCell(1).setCellValue("M-1001");
            r1.createCell(2).setCellValue("tresp1@x");
            r1.createCell(3).setCellValue(5.0);

            RunReport report = new RunReport();
            new ResponsablesSheetBuilder(configWithTablesEnabled(), report).buildAll(wb);

            Sheet sheet = wb.getSheet("tresp1@x");
            assertThat(sheet).isNotNull();
            // Sin pivots: solo cabecera A1.
            assertThat(sheet.getLastRowNum()).isZero();

            assertThat(report.warnings()).anyMatch(w ->
                    "RESPONSABLE".equals(w.category) && w.message.contains("omitidas")
                            && w.message.contains("PDCL"));
        }
    }

    // ==================================================================
    //  Tipo de fórmula
    // ==================================================================

    @Test
    void celdasDeDatosSonTipoFormula() throws java.io.IOException {
        try (Workbook wb = buildRealisticWorkbook()) {
            new ResponsablesSheetBuilder(configWithTablesEnabled(), new RunReport()).buildAll(wb);

            Sheet sheet = wb.getSheet("tresp1@x");
            // Celda dato (P-001 x M-1001) — primera tabla = PDCL (v2.7.0)
            Cell c = sheet.getRow(4).getCell(1);
            assertThat(c.getCellType()).isEqualTo(CellType.FORMULA);
            // Celda total
            Cell t = sheet.getRow(6).getCell(1);
            assertThat(t.getCellType()).isEqualTo(CellType.FORMULA);
        }
    }
}

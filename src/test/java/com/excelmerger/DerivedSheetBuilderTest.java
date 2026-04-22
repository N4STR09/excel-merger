package com.excelmerger;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.Properties;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests de {@link DerivedSheetBuilder}: cubre los dos tipos soportados
 * ({@code FORMULAS}, {@code AGGREGATION}), los casos de error (tipo
 * invalido, sourceSheet ausente, hoja ya existente) y el registro en
 * {@link RunReport}.
 */
class DerivedSheetBuilderTest {

    // ==================================================================
    //  Tipo FORMULAS
    // ==================================================================

    @Test
    void tipoFormulasEscribeCeldasConRefsYValores() throws Exception {
        Properties p = new Properties();
        p.setProperty("derived.sheets", "Resumen");
        p.setProperty("sheet.Resumen.type", "FORMULAS");
        p.setProperty("sheet.Resumen.cell.A1", "INFORME");
        p.setProperty("sheet.Resumen.cell.A2", "Total");
        p.setProperty("sheet.Resumen.cell.B2", "=SUM(1,2,3)");
        p.setProperty("sheet.Resumen.cell.A3", "42");       // numero
        p.setProperty("sheet.Resumen.cell.B3", "texto");    // texto

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new DerivedSheetBuilder(cfg, report).buildAll(wb);

            Sheet s = wb.getSheet("Resumen");
            assertThat(s).isNotNull();

            // A1 texto (titulo)
            assertThat(s.getRow(0).getCell(0).getStringCellValue()).isEqualTo("INFORME");
            // A2 texto
            assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("Total");
            // B2 formula
            assertThat(s.getRow(1).getCell(1).getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(s.getRow(1).getCell(1).getCellFormula()).isEqualTo("SUM(1,2,3)");
            // A3 numero (parseado como double)
            assertThat(s.getRow(2).getCell(0).getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(s.getRow(2).getCell(0).getNumericCellValue()).isEqualTo(42.0);
            // B3 texto
            assertThat(s.getRow(2).getCell(1).getStringCellValue()).isEqualTo("texto");
        }
    }

    @Test
    void tipoFormulasConRefInvalidaGeneraWarning() throws Exception {
        Properties p = new Properties();
        p.setProperty("derived.sheets", "X");
        p.setProperty("sheet.X.type", "FORMULAS");
        p.setProperty("sheet.X.cell.ZZZ999ABC", "algo"); // referencia absurda

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new DerivedSheetBuilder(cfg, report).buildAll(wb);

            // La hoja se crea pero la celda invalida no se escribe
            assertThat(wb.getSheet("X")).isNotNull();
        }

        assertThat(report.warnings()).anyMatch(w ->
                "CONFIG".equals(w.category) && w.message.contains("ZZZ999ABC"));
    }

    // ==================================================================
    //  Tipo AGGREGATION
    // ==================================================================

    private static Workbook workbookWithDataSheet() {
        Workbook wb = new XSSFWorkbook();
        Sheet s = wb.createSheet("Datos");
        // cabecera
        Row header = s.createRow(0);
        header.createCell(0).setCellValue("Grupo");
        header.createCell(1).setCellValue("Valor");
        // 3 filas, 2 grupos
        for (int i = 0; i < 3; i++) {
            Row r = s.createRow(i + 1);
            r.createCell(0).setCellValue(i < 2 ? "A" : "B");
            r.createCell(1).setCellValue(10 * (i + 1));
        }
        return wb;
    }

    @Test
    void aggregationSumEscribeFormulasSumifPorGrupo() throws Exception {
        Properties p = new Properties();
        p.setProperty("derived.sheets", "AggSum");
        p.setProperty("sheet.AggSum.type", "AGGREGATION");
        p.setProperty("sheet.AggSum.sourceSheet", "Datos");
        p.setProperty("sheet.AggSum.groupByColumn", "A");
        p.setProperty("sheet.AggSum.valueColumn", "B");
        p.setProperty("sheet.AggSum.aggregation", "SUM");
        p.setProperty("sheet.AggSum.groupHeader", "Grupo");
        p.setProperty("sheet.AggSum.valueHeader", "Total");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = workbookWithDataSheet()) {
            new DerivedSheetBuilder(cfg, report).buildAll(wb);

            Sheet s = wb.getSheet("AggSum");
            assertThat(s).isNotNull();
            // Cabecera
            assertThat(s.getRow(0).getCell(0).getStringCellValue()).isEqualTo("Grupo");
            assertThat(s.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Total");
            // Fila 1: grupo "A"
            assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("A");
            assertThat(s.getRow(1).getCell(1).getCellType()).isEqualTo(CellType.FORMULA);
            String formulaA = s.getRow(1).getCell(1).getCellFormula();
            assertThat(formulaA).startsWith("SUMIF(");
            assertThat(formulaA).contains("'Datos'!$A$2:$A$4");
            assertThat(formulaA).contains("\"A\"");
            assertThat(formulaA).contains("'Datos'!$B$2:$B$4");
            // Fila 2: grupo "B"
            assertThat(s.getRow(2).getCell(0).getStringCellValue()).isEqualTo("B");
        }
    }

    @Test
    void aggregationSumAnadeFilaDeTotal() throws Exception {
        Properties p = new Properties();
        p.setProperty("derived.sheets", "AggSum");
        p.setProperty("sheet.AggSum.type", "AGGREGATION");
        p.setProperty("sheet.AggSum.sourceSheet", "Datos");
        p.setProperty("sheet.AggSum.groupByColumn", "A");
        p.setProperty("sheet.AggSum.valueColumn", "B");
        p.setProperty("sheet.AggSum.aggregation", "SUM");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = workbookWithDataSheet()) {
            new DerivedSheetBuilder(cfg, report).buildAll(wb);

            Sheet s = wb.getSheet("AggSum");
            // Fila de TOTAL (se anade en rowIdx + 1, tras un hueco)
            Row totalRow = null;
            for (int r = 0; r <= s.getLastRowNum(); r++) {
                Row row = s.getRow(r);
                if (row == null) continue;
                if (row.getCell(0) != null && "TOTAL".equals(row.getCell(0).getStringCellValue())) {
                    totalRow = row;
                    break;
                }
            }
            assertThat(totalRow).as("Fila de TOTAL debe existir para SUM").isNotNull();
            assertThat(totalRow.getCell(1).getCellFormula()).startsWith("SUM(");
        }
    }

    @Test
    void aggregationAvgUsaAverageifYNoAnadeTotal() throws Exception {
        Properties p = new Properties();
        p.setProperty("derived.sheets", "AggAvg");
        p.setProperty("sheet.AggAvg.type", "AGGREGATION");
        p.setProperty("sheet.AggAvg.sourceSheet", "Datos");
        p.setProperty("sheet.AggAvg.groupByColumn", "A");
        p.setProperty("sheet.AggAvg.valueColumn", "B");
        p.setProperty("sheet.AggAvg.aggregation", "AVG");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = workbookWithDataSheet()) {
            new DerivedSheetBuilder(cfg, report).buildAll(wb);

            Sheet s = wb.getSheet("AggAvg");
            assertThat(s.getRow(1).getCell(1).getCellFormula()).startsWith("AVERAGEIF(");

            for (int r = 0; r <= s.getLastRowNum(); r++) {
                Row row = s.getRow(r);
                if (row == null || row.getCell(0) == null) continue;
                assertThat(row.getCell(0).getStringCellValue()).isNotEqualTo("TOTAL");
            }
        }
    }

    @Test
    void aggregationMinUsaMinifs() throws Exception {
        Properties p = new Properties();
        p.setProperty("derived.sheets", "AggMin");
        p.setProperty("sheet.AggMin.type", "AGGREGATION");
        p.setProperty("sheet.AggMin.sourceSheet", "Datos");
        p.setProperty("sheet.AggMin.groupByColumn", "A");
        p.setProperty("sheet.AggMin.valueColumn", "B");
        p.setProperty("sheet.AggMin.aggregation", "MIN");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = workbookWithDataSheet()) {
            new DerivedSheetBuilder(cfg, report).buildAll(wb);

            Sheet s = wb.getSheet("AggMin");
            assertThat(s.getRow(1).getCell(1).getCellFormula()).startsWith("MINIFS(");
        }
    }

    @Test
    void aggregationSourceSheetInexistenteOmiteLaHoja() throws Exception {
        Properties p = new Properties();
        p.setProperty("derived.sheets", "X");
        p.setProperty("sheet.X.type", "AGGREGATION");
        p.setProperty("sheet.X.sourceSheet", "NoExiste");
        p.setProperty("sheet.X.groupByColumn", "A");
        p.setProperty("sheet.X.valueColumn", "B");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = workbookWithDataSheet()) {
            new DerivedSheetBuilder(cfg, report).buildAll(wb);

            // La hoja 'X' se crea pero al fallar la construccion AGGREGATION,
            // no se registra en report.sheets() (ver flag 'built').
        }

        assertThat(report.warnings()).anyMatch(w ->
                "HOJA".equals(w.category) && w.message.contains("NoExiste"));
        assertThat(report.sheets()).doesNotContainKey("X");
    }

    @Test
    void aggregationSinSourceSheetGeneraWarningConfig() throws Exception {
        Properties p = new Properties();
        p.setProperty("derived.sheets", "X");
        p.setProperty("sheet.X.type", "AGGREGATION");
        // sin sourceSheet
        p.setProperty("sheet.X.groupByColumn", "A");
        p.setProperty("sheet.X.valueColumn", "B");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = workbookWithDataSheet()) {
            new DerivedSheetBuilder(cfg, report).buildAll(wb);
        }

        assertThat(report.warnings()).anyMatch(w ->
                "CONFIG".equals(w.category) && w.message.contains("sourceSheet"));
    }

    // ==================================================================
    //  Error handling
    // ==================================================================

    @Test
    void tipoDesconocidoGeneraWarningYOmiteLaHoja() throws Exception {
        Properties p = new Properties();
        p.setProperty("derived.sheets", "raro");
        p.setProperty("sheet.raro.type", "INVENTADO");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new DerivedSheetBuilder(cfg, report).buildAll(wb);

            assertThat(wb.getSheet("raro")).isNull();
        }

        assertThat(report.warnings()).anyMatch(w ->
                "CONFIG".equals(w.category) && w.message.contains("INVENTADO"));
    }

    @Test
    void hojaYaExistenteGeneraWarning() throws Exception {
        Properties p = new Properties();
        p.setProperty("derived.sheets", "X");
        p.setProperty("sheet.X.type", "FORMULAS");
        p.setProperty("sheet.X.cell.A1", "nuevo");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet pre = wb.createSheet("X");
            pre.createRow(0).createCell(0).setCellValue("original");

            new DerivedSheetBuilder(cfg, report).buildAll(wb);

            // La hoja original no se toca
            assertThat(wb.getSheet("X").getRow(0).getCell(0).getStringCellValue())
                    .isEqualTo("original");
        }

        assertThat(report.warnings()).anyMatch(w ->
                "HOJA".equals(w.category) && w.message.contains("X"));
    }

    @Test
    void derivedSheetsVacioNoHaceNada() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromPairs("derived.sheets", "");
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new DerivedSheetBuilder(cfg, report).buildAll(wb);
            assertThat(wb.getNumberOfSheets()).isZero();
        }

        assertThat(report.warnings()).isEmpty();
    }

    @Test
    void multiplesHojasDerivadasSeProcesanEnOrden() throws Exception {
        Properties p = new Properties();
        p.setProperty("derived.sheets", "A,B,C");
        p.setProperty("sheet.A.type", "FORMULAS");
        p.setProperty("sheet.A.cell.A1", "a");
        p.setProperty("sheet.B.type", "FORMULAS");
        p.setProperty("sheet.B.cell.A1", "b");
        p.setProperty("sheet.C.type", "FORMULAS");
        p.setProperty("sheet.C.cell.A1", "c");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new DerivedSheetBuilder(cfg, report).buildAll(wb);

            assertThat(wb.getSheetIndex("A")).isLessThan(wb.getSheetIndex("B"));
            assertThat(wb.getSheetIndex("B")).isLessThan(wb.getSheetIndex("C"));
        }
    }
}

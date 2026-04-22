package com.excelmerger;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.Properties;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests de {@link AvisosSheetBuilder}: vuelca los warnings del {@link RunReport}
 * a una hoja extra (por defecto {@code _Avisos}) dentro del Excel resultado.
 * Cubre:
 * <ul>
 *   <li>No se crea hoja si {@code report.inExcel=false}.</li>
 *   <li>No se crea hoja si no hay warnings (aunque este habilitada).</li>
 *   <li>Los warnings se vuelcan en orden con categoria y mensaje.</li>
 *   <li>La hoja queda oculta por defecto (configurable).</li>
 *   <li>El nombre de la hoja es configurable via {@code report.sheetName}.</li>
 *   <li>Si la hoja ya existe (colision), se omite y se anade warning HOJA.</li>
 *   <li>La hoja se registra en {@link RunReport#sheets()}.</li>
 * </ul>
 */
class AvisosSheetBuilderTest {

    private static Properties configInExcelTrue() {
        Properties p = new Properties();
        p.setProperty("report.inExcel", "true");
        return p;
    }

    @Test
    void noCreaHojaSiInExcelFalse() throws Exception {
        // Con warnings presentes pero report.inExcel=false, no debe crear la hoja.
        Properties p = new Properties();
        p.setProperty("report.inExcel", "false");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        RunReport report = new RunReport();
        report.addWarning("LOOKUP", "App 'ZZ' sin mapeo.");

        try (Workbook wb = new XSSFWorkbook()) {
            new AvisosSheetBuilder(cfg, report).build(wb);

            assertThat(wb.getSheet("_Avisos")).isNull();
            assertThat(wb.getNumberOfSheets()).isZero();
        }
    }

    @Test
    void noCreaHojaSiNoHayWarnings() throws Exception {
        // Con report.inExcel=true pero sin warnings, no debe crear hoja vacia.
        ConfigLoader cfg = TestFixtures.configFromProperties(configInExcelTrue());
        RunReport report = new RunReport();  // sin warnings

        try (Workbook wb = new XSSFWorkbook()) {
            new AvisosSheetBuilder(cfg, report).build(wb);

            assertThat(wb.getSheet("_Avisos")).isNull();
        }
    }

    @Test
    void volcaTodosLosWarningsEnOrden() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(configInExcelTrue());
        RunReport report = new RunReport();
        report.addWarning("LOOKUP", "App 'ZZ' sin mapeo.");
        report.addWarning("CABECERA", "Cabecera 'Foo' no encontrada en Extraccion.");
        report.addWarning("PERFIL", "'rarito.xlsx' no coincide con ningun perfil.");

        try (Workbook wb = new XSSFWorkbook()) {
            new AvisosSheetBuilder(cfg, report).build(wb);

            Sheet sheet = wb.getSheet("_Avisos");
            assertThat(sheet).isNotNull();

            // Cabecera en fila 0
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("Categoria");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Mensaje");

            // 3 filas de datos, en orden de insercion
            assertThat(sheet.getLastRowNum()).isEqualTo(3);

            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("LOOKUP");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue())
                    .isEqualTo("App 'ZZ' sin mapeo.");

            assertThat(sheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("CABECERA");
            assertThat(sheet.getRow(2).getCell(1).getStringCellValue())
                    .isEqualTo("Cabecera 'Foo' no encontrada en Extraccion.");

            assertThat(sheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("PERFIL");
            assertThat(sheet.getRow(3).getCell(1).getStringCellValue())
                    .isEqualTo("'rarito.xlsx' no coincide con ningun perfil.");
        }
    }

    @Test
    void hojaOcultaPorDefecto() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(configInExcelTrue());
        RunReport report = new RunReport();
        report.addWarning("LOOKUP", "un warning");

        try (Workbook wb = new XSSFWorkbook()) {
            new AvisosSheetBuilder(cfg, report).build(wb);

            int idx = wb.getSheetIndex("_Avisos");
            assertThat(idx).isGreaterThanOrEqualTo(0);
            assertThat(wb.isSheetHidden(idx)).isTrue();
        }
    }

    @Test
    void hojaVisibleSiHiddenFalse() throws Exception {
        Properties p = configInExcelTrue();
        p.setProperty("report.hidden", "false");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();
        report.addWarning("LOOKUP", "un warning");

        try (Workbook wb = new XSSFWorkbook()) {
            new AvisosSheetBuilder(cfg, report).build(wb);

            int idx = wb.getSheetIndex("_Avisos");
            assertThat(idx).isGreaterThanOrEqualTo(0);
            assertThat(wb.isSheetHidden(idx)).isFalse();
        }
    }

    @Test
    void nombreDeHojaPersonalizableConConfig() throws Exception {
        Properties p = configInExcelTrue();
        p.setProperty("report.sheetName", "Diagnostico");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();
        report.addWarning("LOOKUP", "un warning");

        try (Workbook wb = new XSSFWorkbook()) {
            new AvisosSheetBuilder(cfg, report).build(wb);

            assertThat(wb.getSheet("Diagnostico")).isNotNull();
            assertThat(wb.getSheet("_Avisos")).isNull();
        }
    }

    @Test
    void hojaAvisosYaExistenteSeOmiteConWarning() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(configInExcelTrue());
        RunReport report = new RunReport();
        report.addWarning("LOOKUP", "un warning previo");

        try (Workbook wb = new XSSFWorkbook()) {
            // Hoja pre-existente con el nombre del report
            Sheet pre = wb.createSheet("_Avisos");
            pre.createRow(0).createCell(0).setCellValue("contenido-previo");

            new AvisosSheetBuilder(cfg, report).build(wb);

            // La hoja original sigue intacta
            assertThat(wb.getSheet("_Avisos").getRow(0).getCell(0).getStringCellValue())
                    .isEqualTo("contenido-previo");
        }

        assertThat(report.warnings()).anyMatch(w ->
                "HOJA".equals(w.category) && w.message.contains("_Avisos"));
    }

    @Test
    void registraLaHojaEnElRunReportConNumeroDeFilas() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(configInExcelTrue());
        RunReport report = new RunReport();
        report.addWarning("LOOKUP", "a");
        report.addWarning("LOOKUP", "b");

        try (Workbook wb = new XSSFWorkbook()) {
            new AvisosSheetBuilder(cfg, report).build(wb);
        }

        // 1 cabecera + 2 warnings = 3 filas
        assertThat(report.sheets()).containsEntry("_Avisos", 3);
    }
}

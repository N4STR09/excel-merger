package com.excelmerger;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.Properties;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests de {@link LookupSheetBuilder}: construye las hojas de lookup tal y
 * como las referencia MesSheetBuilder para VLOOKUP. Covera:
 * <ul>
 *   <li>Hoja creada con cabeceras configuradas y una fila por entrada.</li>
 *   <li>Separador ':' por defecto y override via lookup.X.sep.</li>
 *   <li>Entradas mal formadas se saltan y generan warning.</li>
 *   <li>Entradas con clave vacia se saltan y generan warning.</li>
 *   <li>Opcion 'hidden' oculta la hoja.</li>
 *   <li>Las claves quedan registradas en el RunReport para deteccion de apps sin mapeo.</li>
 *   <li>Lookup sin data: hoja omitida + warning.</li>
 *   <li>Colision con hoja existente: lookup omitido + warning.</li>
 * </ul>
 */
class LookupSheetBuilderTest {

    private static Properties baseLookupConfig() {
        Properties p = new Properties();
        p.setProperty("lookup.sheets", "Equipos");
        p.setProperty("lookup.Equipos.header1", "App");
        p.setProperty("lookup.Equipos.header2", "Equipo");
        p.setProperty("lookup.Equipos.data", "DF:Iker,HE:Jon,EW:JAVA");
        return p;
    }

    @Test
    void construyeHojaConCabecerasYFilasDeDatos() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseLookupConfig());
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new LookupSheetBuilder(cfg, report).buildAll(wb);

            Sheet sheet = wb.getSheet("Equipos");
            assertThat(sheet).isNotNull();

            // Cabeceras en fila 0
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("App");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Equipo");

            // 3 filas de datos
            assertThat(sheet.getLastRowNum()).isEqualTo(3);
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("DF");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("Iker");
            assertThat(sheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("HE");
            assertThat(sheet.getRow(2).getCell(1).getStringCellValue()).isEqualTo("Jon");
            assertThat(sheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("EW");
            assertThat(sheet.getRow(3).getCell(1).getStringCellValue()).isEqualTo("JAVA");
        }
    }

    @Test
    void registraClavesEnRunReportParaDeteccionDeAppsSinMapeo() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseLookupConfig());
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new LookupSheetBuilder(cfg, report).buildAll(wb);
        }

        assertThat(report.hasLookup("Equipos")).isTrue();
        assertThat(report.getLookupKeys("Equipos")).containsExactly("DF", "HE", "EW");
    }

    @Test
    void addSheetRegistraLaHojaCreadaEnReport() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseLookupConfig());
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new LookupSheetBuilder(cfg, report).buildAll(wb);
        }

        // 1 cabecera + 3 filas = 4 filas totales
        assertThat(report.sheets()).containsEntry("Equipos", 4);
    }

    @Test
    void respetaFlagHidden() throws Exception {
        Properties p = baseLookupConfig();
        p.setProperty("lookup.Equipos.hidden", "true");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new LookupSheetBuilder(cfg, report).buildAll(wb);

            int idx = wb.getSheetIndex("Equipos");
            assertThat(wb.isSheetHidden(idx)).isTrue();
        }
    }

    @Test
    void porDefectoLaHojaEsVisible() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseLookupConfig());
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new LookupSheetBuilder(cfg, report).buildAll(wb);

            int idx = wb.getSheetIndex("Equipos");
            assertThat(wb.isSheetHidden(idx)).isFalse();
        }
    }

    @Test
    void separadorPersonalizadoFunciona() throws Exception {
        Properties p = new Properties();
        p.setProperty("lookup.sheets", "Mapa");
        p.setProperty("lookup.Mapa.sep", "->");
        p.setProperty("lookup.Mapa.data", "A->1,B->2,C->3");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new LookupSheetBuilder(cfg, report).buildAll(wb);

            Sheet sheet = wb.getSheet("Mapa");
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("A");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("1");
            assertThat(sheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("C");
        }
    }

    @Test
    void entradasSinSeparadorSeSaltanYGeneranWarning() throws Exception {
        Properties p = new Properties();
        p.setProperty("lookup.sheets", "Lk");
        p.setProperty("lookup.Lk.data", "A:1,no-separador,B:2");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new LookupSheetBuilder(cfg, report).buildAll(wb);

            Sheet sheet = wb.getSheet("Lk");
            // 2 filas validas (A y B), no 3
            assertThat(sheet.getLastRowNum()).isEqualTo(2);
        }

        assertThat(report.warnings()).anyMatch(w ->
                "LOOKUP".equals(w.category) && w.message.contains("no-separador"));
    }

    @Test
    void entradaConClaveVaciaSeSaltaYGeneraWarning() throws Exception {
        Properties p = new Properties();
        p.setProperty("lookup.sheets", "Lk");
        p.setProperty("lookup.Lk.data", "A:1,:valor-sin-clave,B:2");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new LookupSheetBuilder(cfg, report).buildAll(wb);

            Sheet sheet = wb.getSheet("Lk");
            assertThat(sheet.getLastRowNum()).isEqualTo(2);
        }

        assertThat(report.warnings()).anyMatch(w ->
                "LOOKUP".equals(w.category) && w.message.contains("clave vacia"));
    }

    @Test
    void lookupSinDataSeOmiteYGeneraWarning() throws Exception {
        Properties p = new Properties();
        p.setProperty("lookup.sheets", "Vacio");
        // sin lookup.Vacio.data
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new LookupSheetBuilder(cfg, report).buildAll(wb);

            assertThat(wb.getSheet("Vacio")).isNull();
        }

        assertThat(report.warnings()).anyMatch(w ->
                "LOOKUP".equals(w.category) && w.message.contains("sin datos"));
    }

    @Test
    void colisionConHojaExistenteOmiteElLookup() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromProperties(baseLookupConfig());
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            // Hoja pre-existente con el mismo nombre
            Sheet pre = wb.createSheet("Equipos");
            pre.createRow(0).createCell(0).setCellValue("hola");

            new LookupSheetBuilder(cfg, report).buildAll(wb);

            // La hoja original sigue, sin tocarse
            assertThat(wb.getSheet("Equipos").getRow(0).getCell(0).getStringCellValue())
                    .isEqualTo("hola");
        }

        assertThat(report.warnings()).anyMatch(w ->
                "LOOKUP".equals(w.category) && w.message.contains("Equipos"));
    }

    @Test
    void lookupSheetsVacioNoHaceNada() throws Exception {
        ConfigLoader cfg = TestFixtures.configFromPairs("lookup.sheets", "");
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new LookupSheetBuilder(cfg, report).buildAll(wb);
            assertThat(wb.getNumberOfSheets()).isZero();
        }

        assertThat(report.warnings()).isEmpty();
    }

    @Test
    void multiplesLookupsSeProcesanEnOrden() throws Exception {
        Properties p = new Properties();
        p.setProperty("lookup.sheets", "A,B");
        p.setProperty("lookup.A.data", "a1:x");
        p.setProperty("lookup.B.data", "b1:y");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = new XSSFWorkbook()) {
            new LookupSheetBuilder(cfg, report).buildAll(wb);

            assertThat(wb.getSheet("A")).isNotNull();
            assertThat(wb.getSheet("B")).isNotNull();
            assertThat(wb.getSheetIndex("A")).isLessThan(wb.getSheetIndex("B"));
        }
    }
}

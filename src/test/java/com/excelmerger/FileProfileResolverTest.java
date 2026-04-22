package com.excelmerger;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.Properties;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests de {@link FileProfileResolver} y de su utilidad estatica
 * {@link FileProfileResolver#safeSheetName(String)}.
 */
class FileProfileResolverTest {

    // ==================================================================
    //  safeSheetName()
    // ==================================================================

    @Test
    void safeSheetNameNullDevuelveSheetPorDefecto() {
        assertThat(FileProfileResolver.safeSheetName(null)).isEqualTo("Sheet");
    }

    @Test
    void safeSheetNameSustituyeCaracteresProhibidos() {
        String cleaned = FileProfileResolver.safeSheetName("a/b\\c*d?e[f]g:h");

        // Todos los especiales sustituidos por '_'
        assertThat(cleaned).isEqualTo("a_b_c_d_e_f_g_h");
    }

    @Test
    void safeSheetNameNoTocaNombresValidos() {
        assertThat(FileProfileResolver.safeSheetName("Datos_Fusionados"))
                .isEqualTo("Datos_Fusionados");
    }

    @Test
    void safeSheetNameTruncaA31Caracteres() {
        String original = "abcdefghijklmnopqrstuvwxyz0123456789"; // 36 chars
        String trimmed = FileProfileResolver.safeSheetName(original);

        assertThat(trimmed).hasSize(FileProfileResolver.MAX_SHEET_NAME_LEN);
        assertThat(trimmed).isEqualTo(original.substring(0, 31));
    }

    @Test
    void maxSheetNameLenEs31() {
        assertThat(FileProfileResolver.MAX_SHEET_NAME_LEN).isEqualTo(31);
    }

    // ==================================================================
    //  hasProfiles()
    // ==================================================================

    @Test
    void hasProfilesEsFalseSinConfigProfiles() {
        ConfigLoader cfg = TestFixtures.configFromPairs("profiles", "");

        FileProfileResolver resolver = new FileProfileResolver(cfg);

        assertThat(resolver.hasProfiles()).isFalse();
    }

    @Test
    void hasProfilesEsTrueConPerfilesDeclarados() {
        Properties p = new Properties();
        p.setProperty("profiles", "A");
        p.setProperty("profile.A.detect.headers", "x,y,z");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        FileProfileResolver resolver = new FileProfileResolver(cfg);

        assertThat(resolver.hasProfiles()).isTrue();
    }

    // ==================================================================
    //  resolve()  -  matching por cabeceras
    // ==================================================================

    @Test
    void resolveDevuelvePerfilSiLasCabecerasCoinciden() throws Exception {
        Properties p = new Properties();
        p.setProperty("profiles", "Extraccion");
        p.setProperty("profile.Extraccion.sheetName", "Extraccion");
        p.setProperty("profile.Extraccion.detect.headerRow", "1");
        p.setProperty("profile.Extraccion.detect.headers", "Peticion,Titulo,Estado,Recurso");
        p.setProperty("profile.Extraccion.detect.minMatches", "3");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Hoja1");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Peticion");
            header.createCell(1).setCellValue("Titulo");
            header.createCell(2).setCellValue("Estado");

            FileProfileResolver resolver = new FileProfileResolver(cfg);
            FileProfileResolver.FileProfile match = resolver.resolve(wb, new File("dummy.xlsx"));

            assertThat(match).isNotNull();
            assertThat(match.getId()).isEqualTo("Extraccion");
            assertThat(match.getSheetName()).isEqualTo("Extraccion");
        }
    }

    @Test
    void resolveDevuelveNullSiCabecerasInsuficientes() throws Exception {
        Properties p = new Properties();
        p.setProperty("profiles", "Extraccion");
        p.setProperty("profile.Extraccion.detect.headers", "A,B,C,D,E");
        p.setProperty("profile.Extraccion.detect.minMatches", "4");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Hoja1");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("A");
            header.createCell(1).setCellValue("B");
            // Solo 2 coincidencias, necesitamos 4

            FileProfileResolver resolver = new FileProfileResolver(cfg);
            FileProfileResolver.FileProfile match = resolver.resolve(wb, new File("dummy.xlsx"));

            assertThat(match).isNull();
        }
    }

    @Test
    void resolveCoincideCabecerasIgnorandoCase() throws Exception {
        Properties p = new Properties();
        p.setProperty("profiles", "P");
        p.setProperty("profile.P.sheetName", "P");
        p.setProperty("profile.P.detect.headers", "PETICION,TITULO,ESTADO");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Hoja1");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("peticion");
            header.createCell(1).setCellValue("titulo");
            header.createCell(2).setCellValue("estado");

            FileProfileResolver.FileProfile match =
                    new FileProfileResolver(cfg).resolve(wb, new File("x.xlsx"));

            assertThat(match).isNotNull();
            assertThat(match.getId()).isEqualTo("P");
        }
    }

    @Test
    void resolveDeteccionAdmiteCabecerasComoSubcadena() throws Exception {
        // El codigo real usa actualHeaders.stream().anyMatch(h -> h.contains(expected))
        // asi que "Peticion" debe casar con una cabecera real "Peticion BFA".
        Properties p = new Properties();
        p.setProperty("profiles", "P");
        p.setProperty("profile.P.detect.headers", "Peticion,Titulo");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Hoja1");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Peticion BFA");
            header.createCell(1).setCellValue("Titulo de la peticion");

            FileProfileResolver.FileProfile match =
                    new FileProfileResolver(cfg).resolve(wb, new File("x.xlsx"));

            assertThat(match).isNotNull();
        }
    }

    @Test
    void resolveUsaHeaderRowConfigurada() throws Exception {
        // Cabeceras en fila 2 (como en el perfil Cierre real).
        Properties p = new Properties();
        p.setProperty("profiles", "Cierre");
        p.setProperty("profile.Cierre.detect.headerRow", "2");
        p.setProperty("profile.Cierre.detect.headers", "Project Key,Issue Key,Hours");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Hoja1");
            Row title = sheet.createRow(0);
            title.createCell(0).setCellValue("EXPORT JIRA");
            Row header = sheet.createRow(1);
            header.createCell(0).setCellValue("Project Key");
            header.createCell(1).setCellValue("Issue Key");
            header.createCell(2).setCellValue("Hours");

            FileProfileResolver.FileProfile match =
                    new FileProfileResolver(cfg).resolve(wb, new File("x.xlsx"));

            assertThat(match).isNotNull();
            assertThat(match.getId()).isEqualTo("Cierre");
        }
    }

    @Test
    void resolveConCellValueAnadeCriterio() throws Exception {
        Properties p = new Properties();
        p.setProperty("profiles", "P");
        p.setProperty("profile.P.detect.headers", "A,B");
        p.setProperty("profile.P.detect.cellValue.A1", "EXPORT");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Hoja1");
            Row row0 = sheet.createRow(0);
            row0.createCell(0).setCellValue("EXPORT JIRA - REPORT");
            row0.createCell(1).setCellValue("A");
            // row 1 vacia a proposito

            FileProfileResolver.FileProfile match =
                    new FileProfileResolver(cfg).resolve(wb, new File("x.xlsx"));

            // Aunque "B" no este presente, basta con que "A" este y con el cellValue
            // El minMatches default es headers.size(), aqui 2, con solo 1 match... no deberia cuadrar
            // Probamos el efecto del cellValue cuando cabeceras OK:
            assertThat(match).isNull();
        }

        // Ahora con todas las cabeceras + cellValue -> match
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Hoja1");
            Row row0 = sheet.createRow(0);
            row0.createCell(0).setCellValue("EXPORT JIRA - REPORT");
            Row row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue("A");
            row1.createCell(1).setCellValue("B");

            // Ajustamos headerRow a 2 para esta parte del test
            Properties p2 = new Properties();
            p2.setProperty("profiles", "P");
            p2.setProperty("profile.P.detect.headerRow", "2");
            p2.setProperty("profile.P.detect.headers", "A,B");
            p2.setProperty("profile.P.detect.cellValue.A1", "EXPORT");
            ConfigLoader cfg2 = TestFixtures.configFromProperties(p2);

            FileProfileResolver.FileProfile match =
                    new FileProfileResolver(cfg2).resolve(wb, new File("x.xlsx"));

            assertThat(match).isNotNull();
        }
    }

    @Test
    void resolveDevuelveNullSiCellValueNoCoincide() throws Exception {
        Properties p = new Properties();
        p.setProperty("profiles", "P");
        p.setProperty("profile.P.detect.headers", "A,B");
        p.setProperty("profile.P.detect.cellValue.A1", "EXPECTED_TITLE");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Hoja1");
            Row row0 = sheet.createRow(0);
            row0.createCell(0).setCellValue("OTRO TITULO");
            row0.createCell(1).setCellValue("A");
            row0.createCell(2).setCellValue("B");

            FileProfileResolver.FileProfile match =
                    new FileProfileResolver(cfg).resolve(wb, new File("x.xlsx"));

            assertThat(match).isNull();
        }
    }

    @Test
    void resolveSeDetieneEnElPrimerMatchEnOrdenDeclaracion() throws Exception {
        Properties p = new Properties();
        p.setProperty("profiles", "Primero,Segundo");
        p.setProperty("profile.Primero.sheetName", "Primero");
        p.setProperty("profile.Primero.detect.headers", "X,Y");
        p.setProperty("profile.Segundo.sheetName", "Segundo");
        p.setProperty("profile.Segundo.detect.headers", "X,Y");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("H");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("X");
            header.createCell(1).setCellValue("Y");

            FileProfileResolver.FileProfile match =
                    new FileProfileResolver(cfg).resolve(wb, new File("x.xlsx"));

            // Ambos perfiles casan; se devuelve el PRIMERO de la lista "profiles"
            assertThat(match).isNotNull();
            assertThat(match.getId()).isEqualTo("Primero");
        }
    }
}

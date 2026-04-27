package com.excelmerger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests unitarios del builder de hojas por responsable (v2.3.0).
 *
 * <p>No usa fixtures Excel: construye un Workbook en memoria con la hoja
 * Resultado sintetica, que es la unica entrada que el builder consume.
 * Asi los tests son rapidos, deterministas y se centran en la logica
 * de extraccion (trim, case-folding, sanitizacion, colisiones, orden).</p>
 */
class ResponsablesSheetBuilderTest {

    /**
     * Construye un Workbook en memoria con una hoja {@code Resultado} cuya
     * cabecera incluye una columna {@code Res. Tecnico}, y rellena las celdas
     * de esa columna con los valores indicados (uno por fila de datos).
     */
    private Workbook buildWorkbookWithResultado(String... resTecnicoValues) {
        Workbook wb = new XSSFWorkbook();
        Sheet resultado = wb.createSheet("Resultado");
        Row header = resultado.createRow(0);
        header.createCell(0).setCellValue("Petición");
        header.createCell(1).setCellValue("Res. Tecnico");
        for (int i = 0; i < resTecnicoValues.length; i++) {
            Row row = resultado.createRow(i + 1);
            row.createCell(0).setCellValue("P-" + (1000 + i));
            String v = resTecnicoValues[i];
            if (v != null) {
                row.createCell(1).setCellValue(v);
            }
        }
        return wb;
    }

    /**
     * ConfigLoader minimo apto para alimentar al builder. Solo necesita
     * mes.sheetName (default "MES"; aqui apuntamos a "Resultado") y
     * summary.byResponsible.column (default "Res. Tecnico").
     */
    private ConfigLoader minimalConfig() {
        Properties p = new Properties();
        p.setProperty("mes.sheetName", "Resultado");
        p.setProperty("summary.byResponsible.column", "Res. Tecnico");
        return TestFixtures.configFromProperties(p);
    }

    private List<String> sheetNamesAfterResultado(Workbook wb) {
        // Devuelve los nombres de las hojas creadas, excluida la hoja
        // Resultado (que ya existia antes de invocar al builder).
        List<String> out = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetName(i);
            if (!"Resultado".equals(name)) {
                out.add(name);
            }
        }
        return out;
    }

    @Test
    void resultadoVacioGeneraCeroHojasYWarning() {
        Workbook wb = new XSSFWorkbook();
        Sheet resultado = wb.createSheet("Resultado");
        Row header = resultado.createRow(0);
        header.createCell(0).setCellValue("Petición");
        header.createCell(1).setCellValue("Res. Tecnico");
        // Sin filas de datos.

        RunReport report = new RunReport();
        new ResponsablesSheetBuilder(minimalConfig(), report).buildAll(wb);

        assertThat(sheetNamesAfterResultado(wb)).isEmpty();
        assertThat(report.warnings())
                .anyMatch(w -> "RESPONSABLE".equals(w.category)
                        && w.message.contains("sin filas de datos"));
    }

    @Test
    void hojaResultadoAusenteGeneraWarningYNoCreaHojas() {
        Workbook wb = new XSSFWorkbook();
        // Sin hoja Resultado.
        wb.createSheet("Otra");

        RunReport report = new RunReport();
        new ResponsablesSheetBuilder(minimalConfig(), report).buildAll(wb);

        assertThat(report.warnings())
                .anyMatch(w -> "RESPONSABLE".equals(w.category)
                        && w.message.contains("ausente"));
        assertThat(wb.getNumberOfSheets()).isEqualTo(1); // solo "Otra"
    }

    @Test
    void columnaResTecnicoAusenteGeneraWarningYNoCreaHojas() {
        Workbook wb = new XSSFWorkbook();
        Sheet resultado = wb.createSheet("Resultado");
        Row header = resultado.createRow(0);
        header.createCell(0).setCellValue("Petición");
        // Sin columna Res. Tecnico.
        Row row = resultado.createRow(1);
        row.createCell(0).setCellValue("P-001");

        RunReport report = new RunReport();
        new ResponsablesSheetBuilder(minimalConfig(), report).buildAll(wb);

        assertThat(sheetNamesAfterResultado(wb)).isEmpty();
        assertThat(report.warnings())
                .anyMatch(w -> "RESPONSABLE".equals(w.category)
                        && w.message.contains("Res. Tecnico"));
    }

    @Test
    void tresResponsablesDistintosGeneranTresHojas() {
        Workbook wb = buildWorkbookWithResultado("juan", "ana", "pedro");

        RunReport report = new RunReport();
        new ResponsablesSheetBuilder(minimalConfig(), report).buildAll(wb);

        // Ordenadas alfabeticamente por Collator es_ES PRIMARY.
        assertThat(sheetNamesAfterResultado(wb))
                .containsExactly("ana", "juan", "pedro");
    }

    @Test
    void valoresVaciosOSoloEspaciosSeIgnoran() {
        Workbook wb = buildWorkbookWithResultado("juan", "", "  ", null, "ana");

        RunReport report = new RunReport();
        new ResponsablesSheetBuilder(minimalConfig(), report).buildAll(wb);

        assertThat(sheetNamesAfterResultado(wb)).containsExactly("ana", "juan");
    }

    @Test
    void valoresConCaseDistintoColapsaEnUnaSolaHojaConPrimeraOcurrenciaComoCanonico() {
        // "tresp1@x" aparece primero, "TRESP1@x" despues.
        // Tras case-fold, son la misma clave; el nombre canonico debe ser
        // "tresp1@x" (el primero visto, P1003).
        Workbook wb = buildWorkbookWithResultado(
                "ana",         // P-1000
                "tresp1@x",    // P-1001
                "TRESP1@x",    // P-1002 (mismo case-folded que tresp1@x)
                "Tresp1@X");   // P-1003 (mismo)

        RunReport report = new RunReport();
        new ResponsablesSheetBuilder(minimalConfig(), report).buildAll(wb);

        // Solo 2 hojas: "ana" y "tresp1@x".
        assertThat(sheetNamesAfterResultado(wb)).containsExactly("ana", "tresp1@x");
    }

    @Test
    void trimAplicadoAValoresDeResTecnico() {
        // "  juan  " (con padding) y "juan" deben colapsar.
        Workbook wb = buildWorkbookWithResultado("juan", "  juan  ", "ana");

        RunReport report = new RunReport();
        new ResponsablesSheetBuilder(minimalConfig(), report).buildAll(wb);

        assertThat(sheetNamesAfterResultado(wb)).containsExactly("ana", "juan");
    }

    @Test
    void cabeceraA1DeCadaHojaContieneNombreCanonicoDelResponsable() {
        Workbook wb = buildWorkbookWithResultado("Juan", "Ana");

        RunReport report = new RunReport();
        new ResponsablesSheetBuilder(minimalConfig(), report).buildAll(wb);

        Sheet juan = wb.getSheet("Juan");
        Sheet ana = wb.getSheet("Ana");
        assertThat(juan).isNotNull();
        assertThat(ana).isNotNull();
        Cell juanA1 = juan.getRow(0).getCell(0);
        Cell anaA1 = ana.getRow(0).getCell(0);
        assertThat(juanA1.getStringCellValue()).isEqualTo("Juan");
        assertThat(anaA1.getStringCellValue()).isEqualTo("Ana");
    }

    @Test
    void cabeceraA1AplicaEstiloTitulo() {
        Workbook wb = buildWorkbookWithResultado("juan");
        new ResponsablesSheetBuilder(minimalConfig(), new RunReport()).buildAll(wb);

        Cell a1 = wb.getSheet("juan").getRow(0).getCell(0);
        // El estilo aplicado proviene de StyleFactory.title (bold, 14pt).
        // Verificamos via la altura de fuente, que es la huella mas
        // distintiva: title() pone 14pt frente al default de 11pt de
        // Excel/POI. Asi no dependemos de la API exacta de getFont*
        // (que cambia entre versiones de POI 4.x/5.x).
        org.apache.poi.ss.usermodel.Font font =
                ((org.apache.poi.xssf.usermodel.XSSFCellStyle) a1.getCellStyle()).getFont();
        assertThat(font.getBold()).as("A1 debe estar en negrita").isTrue();
        assertThat(font.getFontHeightInPoints())
                .as("A1 debe ser de 14pt").isEqualTo((short) 14);
    }

    @Test
    void nombreConCaracteresProhibidosSeSaneaYEmiteWarning() {
        // El caracter ':' esta prohibido en Excel, igual que / \ * ? [ ].
        Workbook wb = buildWorkbookWithResultado("Juan/Carlos");

        RunReport report = new RunReport();
        new ResponsablesSheetBuilder(minimalConfig(), report).buildAll(wb);

        // Hoja saneada: '/' -> '_'.
        assertThat(wb.getSheet("Juan_Carlos")).isNotNull();
        // Pero A1 contiene el nombre canonico ORIGINAL (antes de sanear).
        assertThat(wb.getSheet("Juan_Carlos").getRow(0).getCell(0).getStringCellValue())
                .isEqualTo("Juan/Carlos");
        // Y se ha emitido warning RESPONSABLE de saneo.
        assertThat(report.warnings())
                .anyMatch(w -> "RESPONSABLE".equals(w.category)
                        && w.message.contains("saneado")
                        && w.message.contains("Juan/Carlos")
                        && w.message.contains("Juan_Carlos"));
    }

    @Test
    void nombresLargosSeTruncan31Chars() {
        String veryLong = "Juan_Carlos_Manuel_Perez_Garcia_de_la_Vega";  // > 31 chars
        Workbook wb = buildWorkbookWithResultado(veryLong);

        new ResponsablesSheetBuilder(minimalConfig(), new RunReport()).buildAll(wb);

        // El builder no asume un nombre concreto; comprobamos que existe
        // una unica hoja nueva, con nombre <= 31 chars.
        List<String> nuevas = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String n = wb.getSheetName(i);
            if (!"Resultado".equals(n)) nuevas.add(n);
        }
        assertThat(nuevas).hasSize(1);
        assertThat(nuevas.get(0).length()).isLessThanOrEqualTo(31);
        // El nombre canonico (en A1) si conserva los 42 chars originales.
        assertThat(wb.getSheet(nuevas.get(0)).getRow(0).getCell(0).getStringCellValue())
                .isEqualTo(veryLong);
    }

    @Test
    void nombresQueColisionanTrasSaneoSeSufijanYWarning() {
        // Dos responsables con nombres distintos pero que tras truncado a 31
        // chars colapsan: el mismo prefijo de 31 chars con cola distinta.
        // Resp1: "Juan_Carlos_Manuel_Perez_Garcia_AAA" (35 chars; trunca a 31)
        // Resp2: "Juan_Carlos_Manuel_Perez_Garcia_BBB" (35 chars; trunca a 31)
        // Tras truncado, ambos = "Juan_Carlos_Manuel_Perez_Garcia" (31 chars).
        // Pero antes son distintos en case-folded (terminacion AAA vs BBB),
        // asi que llegan ambos al saneo.
        String r1 = "Juan_Carlos_Manuel_Perez_Garcia_AAA";
        String r2 = "Juan_Carlos_Manuel_Perez_Garcia_BBB";
        Workbook wb = buildWorkbookWithResultado(r1, r2);

        RunReport report = new RunReport();
        new ResponsablesSheetBuilder(minimalConfig(), report).buildAll(wb);

        // Una hoja con el nombre saneado base, y otra con sufijo _2.
        // Recolectamos lo creado.
        List<String> creadas = sheetNamesAfterResultado(wb);
        assertThat(creadas).hasSize(2);
        // Ambas tienen <= 31 chars.
        assertThat(creadas.get(0).length()).isLessThanOrEqualTo(31);
        assertThat(creadas.get(1).length()).isLessThanOrEqualTo(31);
        // Una de ellas tiene sufijo "_2".
        assertThat(creadas).anyMatch(n -> n.endsWith("_2"));
        // Y se emitio warning de colision.
        assertThat(report.warnings())
                .anyMatch(w -> "RESPONSABLE".equals(w.category)
                        && w.message.contains("Colision"));
    }

    @Test
    void ordenAlfabeticoEsDeterministaIndependienteDelOrdenEnResultado() {
        // Inserto en orden Z, A, M; espero salida A, M, Z.
        Workbook wb = buildWorkbookWithResultado("zoe", "ana", "marcos");
        new ResponsablesSheetBuilder(minimalConfig(), new RunReport()).buildAll(wb);
        assertThat(sheetNamesAfterResultado(wb)).containsExactly("ana", "marcos", "zoe");
    }

    @Test
    void ordenAlfabeticoIgnoraCaseEnLaComparacion() {
        // ZOE en MAYUS, ana en minus, Marcos capitalizado: deben ordenar
        // como humano: ana, Marcos, ZOE.
        Workbook wb = buildWorkbookWithResultado("ZOE", "ana", "Marcos");
        new ResponsablesSheetBuilder(minimalConfig(), new RunReport()).buildAll(wb);
        assertThat(sheetNamesAfterResultado(wb)).containsExactly("ana", "Marcos", "ZOE");
    }

    @Test
    void runReportRegistraLasHojasCreadasConUnaFila() {
        Workbook wb = buildWorkbookWithResultado("ana", "juan");
        RunReport report = new RunReport();
        new ResponsablesSheetBuilder(minimalConfig(), report).buildAll(wb);

        // Cada hoja tiene 1 fila (la cabecera A1).
        assertThat(report.sheets()).containsKeys("ana", "juan");
        assertThat(report.sheets().get("ana")).isEqualTo(1);
        assertThat(report.sheets().get("juan")).isEqualTo(1);
    }
}

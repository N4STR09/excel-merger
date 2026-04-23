package com.excelmerger;

import com.excelmerger.exception.InputValidationException;
import com.excelmerger.exception.OutputException;

import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Properties;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Tests de integracion de {@link ExcelMerger} sobre los fixtures reales
 * ({@code extraccion.xlsx}, {@code cierre.xlsx}). No van por Main para poder
 * instanciar directamente el merger y auditar el {@link RunReport}.
 */
class ExcelMergerIntegrationTest {

    // ==================================================================
    //  Happy path: merge completo con perfiles + lookup + MES
    // ==================================================================

    @Test
    void mergeCompletoProduceTodasLasHojasEsperadas(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        RunReport report = new RunReport();

        new ExcelMerger(cfg, report).merge();

        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        assertThat(output).exists();

        try (FileInputStream fis = new FileInputStream(output.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            // Esperamos 5 hojas: Cierre (de cierre.xlsx), Extraccion (de
            // extraccion.xlsx), Equipos (lookup), Resultado (MES) y
            // Resumen (sumatorio por matrícula, v1.6.0).
            assertThat(wb.getNumberOfSheets()).isEqualTo(5);
            assertThat(wb.getSheet("Cierre")).isNotNull();
            assertThat(wb.getSheet("Extraccion")).isNotNull();
            assertThat(wb.getSheet("Equipos")).isNotNull();
            assertThat(wb.getSheet("Resultado")).isNotNull();
            assertThat(wb.getSheet("Resumen")).isNotNull();
        }
    }

    @Test
    void extraccionConservaSus16FilasIncluidaLaDePeticionVacia(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet extraccion = wb.getSheet("Extraccion");
            // 1 cabecera + 14 filas de datos + 1 fila con Peticion vacia = 16 filas
            assertThat(extraccion.getLastRowNum() + 1).isEqualTo(16);
        }
    }

    @Test
    void mesGeneraUnaFilaPorPeticionNoVaciaDeExtraccion(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            // 1 cabecera + 14 peticiones validas (la 15 con Peticion="" se salta)
            assertThat(mes.getLastRowNum() + 1).isEqualTo(15);
            // Primera peticion
            assertThat(mes.getRow(1).getCell(0).getStringCellValue()).isEqualTo("P-001");
            // Ultima peticion (P-014)
            assertThat(mes.getRow(14).getCell(0).getStringCellValue()).isEqualTo("P-014");
        }
    }

    @Test
    void mesColJiraContieneFormulaSumIfsConReferenciaACierre(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            // Columna "Jira" es la 4 (index 3). Primera fila de datos.
            String formula = mes.getRow(1).getCell(3).getCellFormula();

            // POI normaliza la formula con referencias por LETRA DE COLUMNA,
            // no por el nombre "Hours". La columna Hours es la P (index 15)
            // en cierre.xlsx (ahora con Funcion insertada entre Matricula y Account)
            // y asi queda referenciada.
            assertThat(formula).startsWith("SUMIFS(");
            assertThat(formula).contains("Cierre");
            // La columna Hours es ahora la P (index 15), antes era la O (index 14)
            assertThat(formula).contains("$P:$P");
        }
    }

    @Test
    void mesColJiraFormulaSumIfsIncluyeCondicionFuncion(@TempDir Path tmp) throws IOException {
        // Tras el cambio 1.3.1, el SUMIFS de Jira añade un tercer par de criterios:
        // Funcion:Funcion. Verificamos que la formula contiene las 4 referencias
        // a Cierre (1 sum_range + 3 criterios).
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            String formula = mes.getRow(1).getCell(3).getCellFormula();

            // SUMIFS(sum_range, crit_range1, crit1, crit_range2, crit2, crit_range3, crit3)
            // -> 4 referencias "Cierre!..." en total (1 sum + 3 criterios).
            int count = (formula.length() - formula.replace("Cierre", "").length()) / "Cierre".length();
            assertThat(count).isEqualTo(4);
        }
    }

    @Test
    void mesColJiraSumIfsFiltraPorFuncion(@TempDir Path tmp) throws IOException {
        // Verifica numericamente que el SUMIFS filtra por Funcion:
        // P-001/M-1001 tiene en Cierre 3 imputaciones: PROJ-1 (Dev, 3h), PROJ-2 (Dev, 2h),
        // PROJ-3 (Sup, 4h). Como P-001 en Extraccion tiene Funcion=Dev, el SUMIFS debe
        // sumar solo las dos de Dev (5h), excluyendo la de Sup.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            // Fila 2 de MES (index 1) = primera peticion = P-001 / M-1001 / Dev
            CellValue value = evaluator.evaluate(mes.getRow(1).getCell(3));
            assertThat(value.getNumberValue()).isEqualTo(5.0);
        }
    }

    @Test
    void mesColRealUsaFormulaConReferenciaAColumnaJira(@TempDir Path tmp) throws IOException {
        // REAL es la columna 5 (index 4). Su plantilla es "{col:Jira}*1.2".
        // Jira es la columna 4 (index 3 -> letra D). Fila 2 de MES (primera de datos).
        // Por tanto la formula registrada debe ser exactamente "D2*1.2".
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            assertThat(mes.getRow(1).getCell(4).getCellFormula()).isEqualTo("D2*1.2");
            assertThat(mes.getRow(2).getCell(4).getCellFormula()).isEqualTo("D3*1.2");
        }
    }

    @Test
    void runReportRegistraLasCuatroHojas(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        RunReport report = new RunReport();

        new ExcelMerger(cfg, report).merge();

        assertThat(report.sheets().keySet())
                .contains("Cierre", "Extraccion", "Equipos", "Resultado");
    }

    @Test
    void runReportContabilizaFilasCorrectamente(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        RunReport report = new RunReport();

        new ExcelMerger(cfg, report).merge();

        // MES = 1 cabecera + 14 filas
        assertThat(report.sheets()).containsEntry("Resultado", 15);
        // Equipos = 1 cabecera + 10 entradas (de test-config.properties)
        assertThat(report.sheets()).containsEntry("Equipos", 11);
        // Extraccion = 1 cabecera + 14 + 1 (la de Peticion vacia se copia igual)
        assertThat(report.sheets()).containsEntry("Extraccion", 16);
    }

    @Test
    void appsSinMapeoEnVlookupGeneranWarningLOOKUP(@TempDir Path tmp) throws IOException {
        // El fixture tiene una fila con Aplicaci_Activi="ZZ" (P-011) que no esta
        // en el lookup Equipos. Debe levantar un warning LOOKUP con ese valor.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        RunReport report = new RunReport();

        new ExcelMerger(cfg, report).merge();

        assertThat(report.warnings()).anyMatch(w ->
                "LOOKUP".equals(w.category)
                        && w.message.contains("sin mapeo")
                        && w.message.contains("ZZ"));
    }

    // ==================================================================
    //  Hoja Resumen (v1.6.0)
    // ==================================================================

    @Test
    void hojaResumenTrasPipelineCompletoExisteYContieneSumatoriosPorMatricula(
            @TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resumen = wb.getSheet("Resumen");
            assertThat(resumen).isNotNull();

            // Fila 1 (idx 0): titulo
            assertThat(resumen.getRow(0).getCell(0).getStringCellValue())
                    .isEqualTo("Resumen por Matrícula");

            // Fila 3 (idx 2): cabeceras segun test-config
            assertThat(resumen.getRow(2).getCell(0).getStringCellValue()).isEqualTo("Matrícula");
            assertThat(resumen.getRow(2).getCell(1).getStringCellValue()).isEqualTo("Jira");
            assertThat(resumen.getRow(2).getCell(2).getStringCellValue()).isEqualTo("REAL");
            assertThat(resumen.getRow(2).getCell(3).getStringCellValue()).isEqualTo("PDCL");
            assertThat(resumen.getRow(2).getCell(4).getStringCellValue()).isEqualTo("PDCL + Deuda");

            // Al menos una fila de datos con SUMIFS sobre Resultado
            Row firstData = resumen.getRow(3);
            assertThat(firstData).isNotNull();
            String jiraFormula = firstData.getCell(1).getCellFormula();
            assertThat(jiraFormula).startsWith("SUMIFS(");
            assertThat(jiraFormula).contains("Resultado!");

            // Fila de totales al final con SUM(4:...)
            int last = resumen.getLastRowNum();
            Row totals = resumen.getRow(last);
            assertThat(totals.getCell(0).getStringCellValue()).isEqualTo("Total");
            assertThat(totals.getCell(1).getCellFormula()).startsWith("SUM(B4:");
        }
    }

    // ==================================================================
    //  Validacion de archivos de entrada
    // ==================================================================

    @Test
    void lanzaSiInputDirectoryNoExiste(@TempDir Path tmp) throws IOException {
        Properties p = new Properties();
        p.setProperty("input.directory", tmp.resolve("no-existe").toString());
        p.setProperty("output.file", tmp.resolve("out.xlsx").toString());
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        assertThatThrownBy(() -> new ExcelMerger(cfg, new RunReport()).merge())
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("no existe");
    }

    @Test
    void lanzaSiInputDirectorioSinDosExcel(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        Files.createDirectories(inputDir);
        // Solo 1 fichero
        Files.writeString(inputDir.resolve("solo.xlsx"), "dummy");

        Properties p = new Properties();
        p.setProperty("input.directory", inputDir.toString());
        p.setProperty("output.file", tmp.resolve("out.xlsx").toString());
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        assertThatThrownBy(() -> new ExcelMerger(cfg, new RunReport()).merge())
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("Se necesitan al menos 2");
    }

    @Test
    void strictTwoFilesTrueAbortaConTresExcel(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        TestFixtures.copyFixturesTo(inputDir);
        // Anadimos un tercer xlsx copiando uno de los fixtures
        Files.copy(inputDir.resolve("cierre.xlsx"), inputDir.resolve("tercero.xlsx"));

        Properties p = new Properties();
        p.setProperty("input.directory", inputDir.toString());
        p.setProperty("output.file", tmp.resolve("out.xlsx").toString());
        p.setProperty("input.strictTwoFiles", "true");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        assertThatThrownBy(() -> new ExcelMerger(cfg, new RunReport()).merge())
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("exactamente 2");
    }

    @Test
    void strictTwoFilesFalseConTresExcelTomaLosDosPrimeros(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        TestFixtures.copyFixturesTo(inputDir);
        // Copiamos un tercer archivo cuyo nombre empieza por 'z' para que
        // quede el ultimo alfabeticamente.
        Files.copy(inputDir.resolve("cierre.xlsx"), inputDir.resolve("zzz_extra.xlsx"));

        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(output.getParent());

        Properties p = new Properties();
        p.setProperty("input.directory", inputDir.toString());
        p.setProperty("output.file", output.toString());
        p.setProperty("input.strictTwoFiles", "false");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        RunReport report = new RunReport();
        new ExcelMerger(cfg, report).merge();

        assertThat(output).exists();
        // El warning CONFIG sobre "se encontraron N archivos" debe estar
        assertThat(report.warnings()).anyMatch(w ->
                "CONFIG".equals(w.category) && w.message.contains("3"));
    }

    // ==================================================================
    //  Backup y overwrite
    // ==================================================================

    @Test
    void backupMueveOutputAnteriorACarpetaHistory(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        Path output = tmp.resolve("output").resolve("resultado.xlsx");

        // Primera ejecucion: genera el output
        new ExcelMerger(cfg, new RunReport()).merge();
        assertThat(output).exists();

        // Segunda ejecucion con backup=true: debe mover el anterior a history/
        Path configFile = tmp.resolve("test-config.properties");
        String raw = Files.readString(configFile);
        Files.writeString(configFile, raw.replace("output.backup=false", "output.backup=true"));
        ConfigLoader cfg2 = new ConfigLoader(configFile.toString());

        new ExcelMerger(cfg2, new RunReport()).merge();

        Path historyDir = tmp.resolve("output").resolve("history");
        assertThat(historyDir).exists();
        assertThat(Files.list(historyDir).count())
                .as("Debe existir exactamente 1 backup en history/")
                .isEqualTo(1L);
    }

    @Test
    void overwriteFalseAbortaSiOutputYaExiste(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        TestFixtures.copyFixturesTo(inputDir);
        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(output.getParent());
        Files.writeString(output, "previo");  // output pre-existente

        Properties p = new Properties();
        p.setProperty("input.directory", inputDir.toString());
        p.setProperty("output.file", output.toString());
        p.setProperty("output.overwrite", "false");
        p.setProperty("output.backup", "false");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        assertThatThrownBy(() -> new ExcelMerger(cfg, new RunReport()).merge())
                .isInstanceOf(OutputException.class)
                .hasMessageContaining("ya existe");
    }

    @Test
    void lockDeInputAbortaConMensajeClaro(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        TestFixtures.copyFixturesTo(inputDir);
        // Simulamos que Excel tiene 'extraccion.xlsx' abierto dejando el lock
        Files.writeString(inputDir.resolve("~$extraccion.xlsx"), "lock");

        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        Properties p = new Properties();
        p.setProperty("input.directory", inputDir.toString());
        p.setProperty("output.file", output.toString());
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        assertThatThrownBy(() -> new ExcelMerger(cfg, new RunReport()).merge())
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("Cierra 'extraccion.xlsx'");
    }

    @Test
    void lockDeOutputAbortaAntesDeTocarInputs(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        TestFixtures.copyFixturesTo(inputDir);

        Path outputDir = tmp.resolve("output");
        Files.createDirectories(outputDir);
        Path output = outputDir.resolve("resultado.xlsx");
        Files.writeString(outputDir.resolve("~$resultado.xlsx"), "lock");

        Properties p = new Properties();
        p.setProperty("input.directory", inputDir.toString());
        p.setProperty("output.file", output.toString());
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        assertThatThrownBy(() -> new ExcelMerger(cfg, new RunReport()).merge())
                .isInstanceOf(OutputException.class)
                .hasMessageContaining("Cierra 'resultado.xlsx'");
    }

    // ==================================================================
    //  Modo APPEND_ROWS como alternativa
    // ==================================================================

    @Test
    void modoAppendRowsGeneraUnaUnicaHojaFusionada(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        TestFixtures.copyFixturesTo(inputDir);
        Path output = tmp.resolve("output").resolve("merged.xlsx");
        Files.createDirectories(output.getParent());

        Properties p = new Properties();
        p.setProperty("input.directory", inputDir.toString());
        p.setProperty("output.file", output.toString());
        p.setProperty("merge.mode", "APPEND_ROWS");
        p.setProperty("merge.resultSheetName", "Fusionado");
        p.setProperty("merge.copyStyles", "false");
        // En modo APPEND_ROWS los perfiles y MES no aportan nada; los omitimos
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(output.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            assertThat(wb.getSheet("Fusionado")).isNotNull();
            // No hay MES ni Equipos cuando mes.enabled y lookups no estan configurados
            assertThat(wb.getSheet("Resultado")).isNull();
        }
    }

    // ==================================================================
    //  Dry-run
    // ==================================================================

    @Test
    void dryRunNoCreaFicheroDeSalida(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        RunReport report = new RunReport();
        Path output = tmp.resolve("output").resolve("resultado.xlsx");

        new ExcelMerger(cfg, report, true).merge();

        // El fichero de salida NO debe existir
        assertThat(output).doesNotExist();

        // Pero el pipeline si ha corrido entero: las hojas figuran en el report
        assertThat(report.sheets())
                .containsKey("Resultado")
                .containsKey("Extraccion")
                .containsKey("Cierre")
                .containsKey("Equipos")
                .containsKey("Resumen");
    }

    @Test
    void dryRunNoMueveOutputPrevioAHistory(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        Path output = tmp.resolve("output").resolve("resultado.xlsx");

        // Crear un output previo con contenido distintivo
        Files.writeString(output, "contenido-previo-intocable");

        // Activar backup=true en el config para asegurar que, en run real, si
        // moveria el fichero a history/. En dry-run NO debe moverlo.
        Path configFile = tmp.resolve("test-config.properties");
        String raw = Files.readString(configFile);
        Files.writeString(configFile, raw.replace("output.backup=false", "output.backup=true"));
        ConfigLoader cfg2 = new ConfigLoader(configFile.toString());

        new ExcelMerger(cfg2, new RunReport(), true).merge();

        // El output original sigue intacto
        assertThat(Files.readString(output)).isEqualTo("contenido-previo-intocable");
        // Y no se ha creado la carpeta history/
        assertThat(tmp.resolve("output").resolve("history")).doesNotExist();
    }

    @Test
    void dryRunDetectaAppsSinMapeoIgualmente(@TempDir Path tmp) throws IOException {
        // Mismo fixture que appsSinMapeoEnVlookupGeneranWarningLOOKUP: la fila
        // P-011 tiene Aplicaci_Activi="ZZ" sin mapeo en Equipos. El dry-run
        // debe detectarlo igual (ese es justamente su valor: validar antes
        // del cierre mensual).
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        RunReport report = new RunReport();

        new ExcelMerger(cfg, report, true).merge();

        assertThat(report.warnings()).anyMatch(w ->
                "LOOKUP".equals(w.category)
                        && w.message.contains("sin mapeo")
                        && w.message.contains("ZZ"));
        // Y ademas no se ha escrito nada
        assertThat(tmp.resolve("output").resolve("resultado.xlsx")).doesNotExist();
    }

    @Test
    void dryRunSiOutputEstaLockeadoFalla(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        TestFixtures.copyFixturesTo(inputDir);

        Path outputDir = tmp.resolve("output");
        Files.createDirectories(outputDir);
        Path output = outputDir.resolve("resultado.xlsx");
        // Lock de Excel sobre el output
        Files.writeString(outputDir.resolve("~$resultado.xlsx"), "lock");

        Properties p = new Properties();
        p.setProperty("input.directory", inputDir.toString());
        p.setProperty("output.file", output.toString());
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        // En dry-run tambien queremos avisar: si el fichero esta abierto en
        // Excel el usuario necesita saberlo cuanto antes, no cuando haga el
        // run de verdad.
        assertThatThrownBy(() -> new ExcelMerger(cfg, new RunReport(), true).merge())
                .isInstanceOf(OutputException.class)
                .hasMessageContaining("Cierra 'resultado.xlsx'");
    }

    @Test
    void constructorPorDefectoMantieneComportamientoNoDryRun(@TempDir Path tmp) throws IOException {
        // Blindaje de backward-compat: el constructor de 2 args debe escribir
        // el Excel igual que antes de introducir --dry-run.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        Path output = tmp.resolve("output").resolve("resultado.xlsx");

        new ExcelMerger(cfg, new RunReport()).merge();

        assertThat(output).exists();
    }

    // ==================================================================
    //  Hoja _Avisos (report.inExcel)
    // ==================================================================

    @Test
    void hojaAvisosAparaceEnExcelSiHayAppsSinMapeo(@TempDir Path tmp) throws IOException {
        // El fixture tiene la app 'ZZ' sin mapeo; el merger genera warning LOOKUP.
        // Con report.inExcel=true la hoja _Avisos debe aparecer en el xlsx final.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        Path configFile = tmp.resolve("test-config.properties");
        String raw = Files.readString(configFile);
        Files.writeString(configFile, raw + System.lineSeparator() + "report.inExcel=true");
        ConfigLoader cfg2 = new ConfigLoader(configFile.toString());

        new ExcelMerger(cfg2, new RunReport()).merge();

        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        try (FileInputStream fis = new FileInputStream(output.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet avisos = wb.getSheet("_Avisos");
            assertThat(avisos).isNotNull();
            // Cabecera + al menos 1 warning
            assertThat(avisos.getLastRowNum()).isGreaterThanOrEqualTo(1);
            assertThat(avisos.getRow(0).getCell(0).getStringCellValue()).isEqualTo("Categoria");
            assertThat(avisos.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Mensaje");

            // Alguna fila contiene la app 'ZZ' sin mapeo
            boolean encontrado = false;
            for (int r = 1; r <= avisos.getLastRowNum(); r++) {
                String msg = avisos.getRow(r).getCell(1).getStringCellValue();
                if (msg.contains("ZZ")) { encontrado = true; break; }
            }
            assertThat(encontrado)
                    .as("La hoja _Avisos deberia contener una fila mencionando la app 'ZZ'")
                    .isTrue();

            // Por defecto, la hoja esta oculta
            assertThat(wb.isSheetHidden(wb.getSheetIndex("_Avisos"))).isTrue();
        }
    }

    @Test
    void hojaAvisosNoAparaceSiReportInExcelFalse(@TempDir Path tmp) throws IOException {
        // Default: report.inExcel=false -> no hay hoja _Avisos aunque haya warnings.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);

        new ExcelMerger(cfg, new RunReport()).merge();

        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        try (FileInputStream fis = new FileInputStream(output.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            assertThat(wb.getSheet("_Avisos")).isNull();
        }
    }
}

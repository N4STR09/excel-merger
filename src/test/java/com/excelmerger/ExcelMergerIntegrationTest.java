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
    void extraccionConservaSus19FilasIncluidaLaDePeticionVacia(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet extraccion = wb.getSheet("Extraccion");
            // 1 cabecera + 14 filas texto + 3 filas regresion v1.6.2 + 1 skip = 19
            assertThat(extraccion.getLastRowNum() + 1).isEqualTo(19);
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
            // 1 cabecera + 14 peticiones validas + 3 regresion v1.6.2 (la que
            // tiene Peticion="" se salta)
            assertThat(mes.getLastRowNum() + 1).isEqualTo(18);
            // Primera peticion (texto)
            assertThat(mes.getRow(1).getCell(0).getStringCellValue()).isEqualTo("P-001");
            // Ultima peticion de las filas "historicas" (index 14 = fila 15)
            assertThat(mes.getRow(14).getCell(0).getStringCellValue()).isEqualTo("P-014");
            // Primera fila de regresion v1.6.2: Peticion=55751 en Extraccion
            // venia como NUMERIC; tras asText.columns=Peticion,Recurso en el
            // perfil, en Extraccion del workbook resultado aparece como
            // STRING "55751", y CopyColumnStrategy la preserva como STRING
            // aqui. El SUMIFS por tanto casa con Cierre.Component Name="55751"
            // (tambien STRING) y recupera las horas.
            assertThat(mes.getRow(15).getCell(0).getStringCellValue()).isEqualTo("55751");
            assertThat(mes.getRow(17).getCell(0).getStringCellValue()).isEqualTo("138074");
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

    // ==================================================================
    //  Regresion v1.6.2: mismatch de tipos numerico/textual en SUMIFS
    // ==================================================================
    //
    // Los fixtures incluyen desde v1.6.2 tres peticiones con Peticion y
    // Recurso como NUMERIC en Extraccion (55751, 101770, 138074 / 99642,
    // 90014, 99641) y sus imputaciones correspondientes en Cierre con
    // Component Name y Matricula como STRING. Antes del fix, el SUMIFS
    // emitido para esas peticiones daba 0 por mismatch silencioso de
    // tipos (criterio numerico + rango textual no casa).
    //
    // Con el fix, el perfil declara asText.columns=Peticion,Recurso en
    // Extraccion (y Component Name,Matricula en Cierre), de forma que
    // SheetCopier normaliza ambos lados a STRING antes de que se evaluen
    // las formulas. El SUMIFS ve criterio y rango como texto y suma.

    @Test
    void sumifsRecuperaImputacionParaPeticionNumericaSimple(@TempDir Path tmp) throws IOException {
        // Regresion: Peticion=55751 (num en Extraccion) con una unica
        // imputacion en Cierre (PROJ-20, 7h, Dev). Esperado: Jira=7.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            // La primera fila de regresion es la 16 (1-based) -> index 15.
            // Columna Jira es index 3.
            Row row = mes.getRow(15);
            assertThat(row.getCell(0).getStringCellValue()).isEqualTo("55751");
            CellValue value = evaluator.evaluate(row.getCell(3));
            assertThat(value.getNumberValue())
                    .as("SUMIFS para Peticion=55751 (normalizada a STRING) debe recuperar PROJ-20")
                    .isEqualTo(7.0);
        }
    }

    @Test
    void sumifsRecuperaImputacionParaPeticionNumericaConVariasFilas(@TempDir Path tmp) throws IOException {
        // Regresion: Peticion=101770 (num). Dos imputaciones en Cierre
        // (PROJ-21 3h + PROJ-22 2h, ambas Dev). Esperado: Jira=5.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            Row row = mes.getRow(16); // segunda fila de regresion
            assertThat(row.getCell(0).getStringCellValue()).isEqualTo("101770");
            CellValue value = evaluator.evaluate(row.getCell(3));
            assertThat(value.getNumberValue())
                    .as("SUMIFS debe sumar PROJ-21 + PROJ-22 tras normalizar tipos")
                    .isEqualTo(5.0);
        }
    }

    @Test
    void sumifsRecuperaImputacionPeticionNumericaYSigueFiltrandoSup(@TempDir Path tmp) throws IOException {
        // Regresion combinada: Peticion=138074 (num). En Cierre hay
        // PROJ-23 (9h, Dev) y PROJ-24 (4h, Sup). Tras el fix, el SUMIFS
        // debe (1) casar pese al mismatch de tipo y (2) seguir filtrando
        // por Funcion=Dev. Esperado: Jira=9, no 13.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            Row row = mes.getRow(17); // tercera fila de regresion
            assertThat(row.getCell(0).getStringCellValue()).isEqualTo("138074");
            CellValue value = evaluator.evaluate(row.getCell(3));
            assertThat(value.getNumberValue())
                    .as("Jira debe contar PROJ-23 (Dev) y excluir PROJ-24 (Sup)")
                    .isEqualTo(9.0);
        }
    }

    @Test
    void extraccionEnResultadoTienePeticionYRecursoComoStringTrasAsText(@TempDir Path tmp) throws IOException {
        // Verifica directamente la garantia del fix: las celdas de
        // Peticion y Recurso en la hoja Extraccion del workbook
        // resultado son de tipo STRING, aunque en el fixture original
        // algunas son NUMERIC.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet ext = wb.getSheet("Extraccion");
            // Cabecera en fila 1 (index 0). Las filas de regresion son 16,
            // 17 y 18 (index 15, 16, 17). Peticion=col 0, Recurso=col 11.
            for (int r : new int[] {15, 16, 17}) {
                org.apache.poi.ss.usermodel.Cell pet = ext.getRow(r).getCell(0);
                org.apache.poi.ss.usermodel.Cell rec = ext.getRow(r).getCell(11);
                assertThat(pet.getCellType())
                        .as("Peticion en fila " + (r + 1) + " debe ser STRING tras asText")
                        .isEqualTo(org.apache.poi.ss.usermodel.CellType.STRING);
                assertThat(rec.getCellType())
                        .as("Recurso en fila " + (r + 1) + " debe ser STRING tras asText")
                        .isEqualTo(org.apache.poi.ss.usermodel.CellType.STRING);
            }
            // Las filas historicas (texto) deben seguir siendo STRING
            // (no las rompemos accidentalmente al forzar tipo).
            assertThat(ext.getRow(1).getCell(0).getCellType())
                    .isEqualTo(org.apache.poi.ss.usermodel.CellType.STRING);
            assertThat(ext.getRow(1).getCell(0).getStringCellValue()).isEqualTo("P-001");
        }
    }

    @Test
    void asTextColumnsConCabeceraInexistenteEmiteWarning(@TempDir Path tmp) throws IOException {
        // Si el usuario configura asText.columns con una cabecera que no
        // existe en la hoja del perfil, el merge no rompe: simplemente
        // se ignora esa columna y se emite un warning CABECERA en el
        // RunReport para que sea visible en la hoja _Avisos y en el log.
        Path inputDir = tmp.resolve("input");
        Path outputFile = tmp.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(outputFile.getParent());
        TestFixtures.copyFixturesTo(inputDir);
        Path cfgFile = TestFixtures.renderTestConfig(
                tmp.resolve("test-config.properties"), inputDir, outputFile);
        // Reemplazamos la clave por una que mezcla una cabecera real con
        // una inexistente.
        String content = Files.readString(cfgFile);
        content = content.replace(
                "profile.Extraccion.asText.columns=Peticion,Recurso",
                "profile.Extraccion.asText.columns=Peticion,ColumnaQueNoExiste");
        Files.writeString(cfgFile, content);

        ConfigLoader cfg = new ConfigLoader(cfgFile.toString());
        RunReport report = new RunReport();
        new ExcelMerger(cfg, report).merge();

        assertThat(report.warnings()).anyMatch(w ->
                "CABECERA".equals(w.category)
                        && w.message.contains("ColumnaQueNoExiste")
                        && w.message.contains("asText.columns"));
    }

    // ==================================================================
    //  Huerfanos (v1.7.0): filas de Resultado para imputaciones de Cierre
    //  sin contrapartida (Peticion, Recurso) en Extraccion.
    // ==================================================================
    //
    // Los fixtures incluyen desde v1.7.0 cuatro imputaciones huerfanas:
    //   (TICKETS, -)              2 imputaciones x 4h = 8h
    //   (VACACIONES, 90014)       1 imputacion de 3h    (90014 existe en
    //                                                    Extraccion, pero
    //                                                    asociada a 101770)
    //   (P-001, MAT-HUERFANO)     1 imputacion de 1h    (P-001 existe, pero
    //                                                    con M-1001 no con
    //                                                    MAT-HUERFANO)
    //
    // Con mes.orphans.enabled=false (default del test-config) ninguna de
    // estas imputaciones aparece en Resultado. Con enabled=true aparecen
    // como filas adicionales, ordenadas segun el criterio numerico-primero.

    private static ConfigLoader buildConfigWithOrphansEnabled(Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        Path outputFile = tmp.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(outputFile.getParent());
        TestFixtures.copyFixturesTo(inputDir);
        Path cfgFile = TestFixtures.renderTestConfig(
                tmp.resolve("test-config.properties"), inputDir, outputFile);
        String content = Files.readString(cfgFile);
        content = content.replace(
                "mes.orphans.enabled=false",
                "mes.orphans.enabled=true");
        Files.writeString(cfgFile, content);
        return new ConfigLoader(cfgFile.toString());
    }

    @Test
    void orphansEnabledAnadeFilasParaImputacionesSinContrapartida(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = buildConfigWithOrphansEnabled(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            // Antes: 1 cabecera + 14 texto + 3 regresion = 18 filas.
            // Ahora: + 3 huerfanos (TICKETS/-, VACACIONES/90014, P-001/MAT-HUERFANO) = 21.
            assertThat(mes.getLastRowNum() + 1).isEqualTo(21);

            // Construir mapa (Peticion, Matricula) -> fila para buscar los huerfanos.
            java.util.Map<String, Integer> rowByPair = new java.util.LinkedHashMap<>();
            for (int r = 1; r <= mes.getLastRowNum(); r++) {
                String pet = mes.getRow(r).getCell(0).getStringCellValue();
                String mat = mes.getRow(r).getCell(5).getStringCellValue();
                rowByPair.put(pet + "|" + mat, r);
            }

            assertThat(rowByPair).containsKey("TICKETS|-");
            assertThat(rowByPair).containsKey("VACACIONES|90014");
            assertThat(rowByPair).containsKey("P-001|MAT-HUERFANO");

            // Verificar horas agregadas en cada fila huerfana
            int rTickets = rowByPair.get("TICKETS|-");
            int rVac = rowByPair.get("VACACIONES|90014");
            int rP001H = rowByPair.get("P-001|MAT-HUERFANO");

            assertThat(mes.getRow(rTickets).getCell(3).getNumericCellValue()).isEqualTo(8.0);
            assertThat(mes.getRow(rVac).getCell(3).getNumericCellValue()).isEqualTo(3.0);
            assertThat(mes.getRow(rP001H).getCell(3).getNumericCellValue()).isEqualTo(1.0);
        }
    }

    @Test
    void orphansEnabledFilasOrdenadasNumericasPrimero(@TempDir Path tmp) throws IOException {
        // Orden: numericas ASC primero, no numericas alfabetico al final.
        // Peticiones numericas: 55751, 101770, 138074.
        // No numericas: MAT-HUERFANO no, eso es matricula; la Peticion de ese
        //               huerfano es "P-001" que ya existe. Las no numericas
        //               son: P-001..P-014, TICKETS, VACACIONES.
        ConfigLoader cfg = buildConfigWithOrphansEnabled(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            // Las tres primeras filas de datos deben ser las peticiones
            // numericas ordenadas ascendente.
            assertThat(mes.getRow(1).getCell(0).getStringCellValue()).isEqualTo("55751");
            assertThat(mes.getRow(2).getCell(0).getStringCellValue()).isEqualTo("101770");
            assertThat(mes.getRow(3).getCell(0).getStringCellValue()).isEqualTo("138074");

            // La ultima fila de datos debe ser VACACIONES (alfabeticamente
            // la mayor de las no numericas presentes).
            int last = mes.getLastRowNum();
            assertThat(mes.getRow(last).getCell(0).getStringCellValue()).isEqualTo("VACACIONES");
            // La anterior a la ultima debe ser TICKETS.
            assertThat(mes.getRow(last - 1).getCell(0).getStringCellValue()).isEqualTo("TICKETS");
        }
    }

    @Test
    void orphansEnabledColumnasSinDatoRecibenLiteralGuion(@TempDir Path tmp) throws IOException {
        // En las filas huerfanas, columnas como Aplicacion, Titulo, Departamento,
        // Estado, Res. Tecnico reciben un literal "-" porque no hay info de
        // Extraccion asociada.
        ConfigLoader cfg = buildConfigWithOrphansEnabled(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            // Buscar la fila TICKETS/- y verificar columnas "-"
            for (int r = 1; r <= mes.getLastRowNum(); r++) {
                String pet = mes.getRow(r).getCell(0).getStringCellValue();
                if (!"TICKETS".equals(pet)) continue;
                // Columnas MES en test-config:
                //  0=Petición, 1=Aplicación, 2=Equipo, 3=Jira, 4=REAL,
                //  5=Matrícula, 6=Res. Tecnico, 7=PDCL, 8=PDCL + Deuda.
                assertThat(mes.getRow(r).getCell(1).getStringCellValue()).isEqualTo("-");
                assertThat(mes.getRow(r).getCell(6).getStringCellValue()).isEqualTo("-");
                return;
            }
            throw new AssertionError("No se encontro fila TICKETS en Resultado");
        }
    }

    @Test
    void orphansEnabledColumnasFormulaCalculanJiraPor12(@TempDir Path tmp) throws IOException {
        // REAL, PDCL, PDCL+Deuda son columnas FORMULA con plantilla *1.2
        // (o referencia a PDCL). En huerfanos deben evaluar igual que en
        // filas normales: Jira*1.2 y la cascada.
        ConfigLoader cfg = buildConfigWithOrphansEnabled(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 1; r <= mes.getLastRowNum(); r++) {
                String pet = mes.getRow(r).getCell(0).getStringCellValue();
                if (!"TICKETS".equals(pet)) continue;
                // Jira=8.0, REAL=9.6, PDCL=9.6, PDCL+Deuda=9.6
                assertThat(mes.getRow(r).getCell(3).getNumericCellValue()).isEqualTo(8.0);
                CellValue real = evaluator.evaluate(mes.getRow(r).getCell(4));
                assertThat(real.getNumberValue()).isEqualTo(9.6);
                CellValue pdcl = evaluator.evaluate(mes.getRow(r).getCell(7));
                assertThat(pdcl.getNumberValue()).isEqualTo(9.6);
                CellValue pdclPlus = evaluator.evaluate(mes.getRow(r).getCell(8));
                assertThat(pdclPlus.getNumberValue()).isEqualTo(9.6);
                return;
            }
            throw new AssertionError("No se encontro fila TICKETS en Resultado");
        }
    }

    @Test
    void orphansEnabledNoGeneraWarningLookupParaGuion(@TempDir Path tmp) throws IOException {
        // El valor "-" puesto en Aplicacion de las filas huerfanas no debe
        // contabilizarse como "app sin mapeo" en el lookup Equipos. Solo
        // deben aparecer warnings LOOKUP para apps reales (ZZ en el fixture).
        ConfigLoader cfg = buildConfigWithOrphansEnabled(tmp);
        RunReport report = new RunReport();
        new ExcelMerger(cfg, report).merge();

        boolean guionEnWarnings = report.warnings().stream()
                .anyMatch(w -> "LOOKUP".equals(w.category) && w.message.contains("'-'"));
        assertThat(guionEnWarnings)
                .as("El sentinela '-' no debe considerarse app sin mapeo")
                .isFalse();
    }

    @Test
    void orphansEnabledHuerfanoAparecEnResultadoConSuMatricula(@TempDir Path tmp) throws IOException {
        // Verifica que la fila huerfana VACACIONES/90014/3h aparece en
        // Resultado con la matricula correcta y sus horas.
        ConfigLoader cfg = buildConfigWithOrphansEnabled(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            // Buscar fila de VACACIONES y verificar su matricula y horas
            for (int r = 1; r <= mes.getLastRowNum(); r++) {
                String pet = mes.getRow(r).getCell(0).getStringCellValue();
                if (!"VACACIONES".equals(pet)) continue;
                assertThat(mes.getRow(r).getCell(5).getStringCellValue())
                        .as("La matricula del huerfano VACACIONES debe ser '90014'")
                        .isEqualTo("90014");
                assertThat(mes.getRow(r).getCell(3).getNumericCellValue())
                        .as("Las horas del huerfano VACACIONES deben ser 3")
                        .isEqualTo(3.0);
                return;
            }
            throw new AssertionError("No se encontro fila VACACIONES en Resultado");
        }
    }

    @Test
    void orphansEnabledSumaEnResumenPorMatricula(@TempDir Path tmp) throws IOException {
        // End-to-end del principio "no perder ni un dato": las horas
        // huerfanas deben sumar en Resumen. La matricula 90014 aparece
        // en una fila normal (101770/90014/5h) y en una huerfana
        // (VACACIONES/90014/3h). Su total Jira en Resumen debe ser 8h.
        //
        // Tras el fix 1.7.1 del bug C (SummarySheetBuilder escribe todas
        // las matriculas como STRING), el SUMIFS casa correctamente
        // contra la columna Matricula de Resultado, que es STRING por
        // el fix 1.6.2.
        ConfigLoader cfg = buildConfigWithOrphansEnabled(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resumen = wb.getSheet("Resumen");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            // Buscar la fila con matricula 90014 y evaluar su Jira
            for (int r = 3; r <= resumen.getLastRowNum(); r++) {
                Row row = resumen.getRow(r);
                if (row == null) continue;
                org.apache.poi.ss.usermodel.Cell matCell = row.getCell(0);
                if (matCell == null) continue;
                if (matCell.getCellType() != org.apache.poi.ss.usermodel.CellType.STRING) continue;
                if (!"90014".equals(matCell.getStringCellValue())) continue;
                CellValue jiraTotal = evaluator.evaluate(row.getCell(1));
                assertThat(jiraTotal.getNumberValue())
                        .as("90014 debe sumar 5h (fila normal 101770) + 3h (huerfano VACACIONES) = 8h")
                        .isEqualTo(8.0);
                return;
            }
            throw new AssertionError("No se encontro fila 90014 en Resumen");
        }
    }

    @Test
    void orphansDisabledMantieneComportamiento16Point2(@TempDir Path tmp) throws IOException {
        // Test de regresion: con enabled=false (default del test-config)
        // Resultado tiene exactamente 18 filas y nada nuevo.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            assertThat(mes.getLastRowNum() + 1).isEqualTo(18);
            // Primera fila de datos sigue siendo P-001 (orden natural de
            // Extraccion, sin reordenar).
            assertThat(mes.getRow(1).getCell(0).getStringCellValue()).isEqualTo("P-001");
        }
    }

    @Test
    void orphansEnabledConSheetInexistenteEmiteWarning(@TempDir Path tmp) throws IOException {
        // Si mes.orphans.sourceSheet apunta a una hoja que no existe en
        // el workbook, el merge emite warning y continua sin huerfanos.
        Path inputDir = tmp.resolve("input");
        Path outputFile = tmp.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(outputFile.getParent());
        TestFixtures.copyFixturesTo(inputDir);
        Path cfgFile = TestFixtures.renderTestConfig(
                tmp.resolve("test-config.properties"), inputDir, outputFile);
        String content = Files.readString(cfgFile);
        content = content.replace("mes.orphans.enabled=false", "mes.orphans.enabled=true");
        content = content.replace("mes.orphans.sourceSheet=Cierre",
                "mes.orphans.sourceSheet=HojaQueNoExiste");
        Files.writeString(cfgFile, content);

        ConfigLoader cfg = new ConfigLoader(cfgFile.toString());
        RunReport report = new RunReport();
        new ExcelMerger(cfg, report).merge();

        assertThat(report.warnings()).anyMatch(w ->
                "HOJA".equals(w.category) && w.message.contains("HojaQueNoExiste"));
        // Y Resultado sigue teniendo 18 filas (sin huerfanos)
        try (FileInputStream fis = new FileInputStream(outputFile.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {
            assertThat(wb.getSheet("Resultado").getLastRowNum() + 1).isEqualTo(18);
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

        // MES = 1 cabecera + 14 filas texto + 3 filas regresion v1.6.2 = 18
        assertThat(report.sheets()).containsEntry("Resultado", 18);
        // Equipos = 1 cabecera + 10 entradas (de test-config.properties)
        assertThat(report.sheets()).containsEntry("Equipos", 11);
        // Extraccion = 1 cabecera + 14 texto + 3 regresion + 1 Peticion vacia = 19
        assertThat(report.sheets()).containsEntry("Extraccion", 19);
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

package com.excelmerger;

import com.excelmerger.config.ConfigValidator;
import com.excelmerger.exception.InputValidationException;
import com.excelmerger.exception.OutputException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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
    void hojaCierreConservaSus21FilasIncluidaLaDePeticionVacia(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            // v2.0.0: la hoja con las peticiones del ERP (Peticion, Recurso,
            // etc.) se llama ahora "Cierre" tras el swap de perfiles.
            Sheet cierre = wb.getSheet("Cierre");
            // 1 cabecera + 14 filas texto + 3 filas regresion v1.6.2 +
            // 1 fila regresion v1.8.0 (responsable MAYUS) +
            // 1 fila regresion v1.8.1 (responsable con padding) + 1 skip = 21
            assertThat(cierre.getLastRowNum() + 1).isEqualTo(21);
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
            // 1 cabecera + 14 peticiones validas + 3 regresion v1.6.2 +
            // 1 regresion v1.8.0 + 1 regresion v1.8.1 (la que tiene Peticion="" se salta)
            assertThat(mes.getLastRowNum() + 1).isEqualTo(20);
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
            // Fila de regresion v1.8.0: P-015 con responsable en MAYUSCULAS
            // (TRESP1@x). Verifica que el pipeline la copia tal cual, sin
            // tocar el caso; la normalizacion a MAYUSCULAS es cosa de la
            // tabla por responsable en Resumen, no de Resultado.
            assertThat(mes.getRow(18).getCell(0).getStringCellValue()).isEqualTo("P-015");
            // Fila de regresion v1.8.1: P-016 con Usuario_Resp_Tecnico="MG002   "
            // (padding). Tras el trim (profile.Extraccion.trim.columns) la
            // celda Res. Tecnico de Resultado debe venir ya sin espacios.
            assertThat(mes.getRow(19).getCell(0).getStringCellValue()).isEqualTo("P-016");
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
            // en la hoja con el export de Jira (Funcion insertada entre
            // Matricula y Account).
            // v2.0.0: tras el swap, la hoja con las imputaciones de Jira se
            // llama "Extraccion" (antes "Cierre").
            assertThat(formula).startsWith("SUMIFS(");
            assertThat(formula).contains("Extraccion");
            // La columna Hours es ahora la P (index 15), antes era la O (index 14)
            assertThat(formula).contains("$P:$P");
        }
    }

    @Test
    void mesColJiraFormulaSumIfsIncluyeCondicionFuncion(@TempDir Path tmp) throws IOException {
        // Tras el cambio 1.3.1, el SUMIFS de Jira añade un tercer par de criterios:
        // Funcion:Funcion. Verificamos que la formula contiene las 4 referencias
        // a la hoja con el export de Jira (1 sum_range + 3 criterios).
        // v2.0.0: tras el swap, esa hoja se llama "Extraccion" (antes "Cierre").
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            String formula = mes.getRow(1).getCell(3).getCellFormula();

            // SUMIFS(sum_range, crit_range1, crit1, crit_range2, crit2, crit_range3, crit3)
            // -> 4 referencias "Extraccion!..." en total (1 sum + 3 criterios).
            int count = (formula.length() - formula.replace("Extraccion", "").length())
                    / "Extraccion".length();
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
    void cierreEnResultadoTienePeticionYRecursoComoStringTrasAsText(@TempDir Path tmp) throws IOException {
        // Verifica directamente la garantia del fix: las celdas de
        // Peticion y Recurso en la hoja que contiene las peticiones
        // (v2.0.0: llamada "Cierre"; antes era "Extraccion") son de tipo
        // STRING, aunque en el fixture original algunas son NUMERIC.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet peticiones = wb.getSheet("Cierre");
            // Cabecera en fila 1 (index 0). Las filas de regresion son 16,
            // 17 y 18 (index 15, 16, 17). Peticion=col 0, Recurso=col 11.
            for (int r : new int[] {15, 16, 17}) {
                org.apache.poi.ss.usermodel.Cell pet = peticiones.getRow(r).getCell(0);
                org.apache.poi.ss.usermodel.Cell rec = peticiones.getRow(r).getCell(11);
                assertThat(pet.getCellType())
                        .as("Peticion en fila " + (r + 1) + " debe ser STRING tras asText")
                        .isEqualTo(org.apache.poi.ss.usermodel.CellType.STRING);
                assertThat(rec.getCellType())
                        .as("Recurso en fila " + (r + 1) + " debe ser STRING tras asText")
                        .isEqualTo(org.apache.poi.ss.usermodel.CellType.STRING);
            }
            // Las filas historicas (texto) deben seguir siendo STRING
            // (no las rompemos accidentalmente al forzar tipo).
            assertThat(peticiones.getRow(1).getCell(0).getCellType())
                    .isEqualTo(org.apache.poi.ss.usermodel.CellType.STRING);
            assertThat(peticiones.getRow(1).getCell(0).getStringCellValue()).isEqualTo("P-001");
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
        // v2.0.0: la clave es profile.Cierre.asText.columns (antes
        // profile.Extraccion.asText.columns), y su valor en test-config
        // incluye Usuario_Resp_Tecnico desde v1.8.1.
        String content = Files.readString(cfgFile);
        content = content.replace(
                "profile.Cierre.asText.columns=Peticion,Recurso,Usuario_Resp_Tecnico",
                "profile.Cierre.asText.columns=Peticion,ColumnaQueNoExiste");
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
    //  Huerfanos (v1.7.0): filas de Resultado para imputaciones del
    //  perfil Extraccion sin contrapartida (Peticion, Recurso) en el
    //  perfil Cierre. v2.0.0: swap de nombres de perfil.
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
            // Antes: 1 cabecera + 14 texto + 3 regresion + 1 v1.8.0 + 1 v1.8.1 = 20 filas.
            // Ahora: + 3 huerfanos (TICKETS/-, VACACIONES/90014, P-001/MAT-HUERFANO) = 23.
            assertThat(mes.getLastRowNum() + 1).isEqualTo(23);

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
                // Columnas MES en test-config (v2.1.0 con Funcion en posicion 7):
                //  0=Petición, 1=Aplicación, 2=Equipo, 3=Jira, 4=Facturar,
                //  5=Matrícula, 6=Funcion, 7=Res. Tecnico, 8=PDCL, 9=PDCL + Deuda.
                assertThat(mes.getRow(r).getCell(1).getStringCellValue()).isEqualTo("-");
                assertThat(mes.getRow(r).getCell(7).getStringCellValue()).isEqualTo("-");
                return;
            }
            throw new AssertionError("No se encontro fila TICKETS en Resultado");
        }
    }

    @Test
    void orphansEnabledColumnasFormulaCalculanJiraPor12(@TempDir Path tmp) throws IOException {
        // Facturar, PDCL, PDCL+Deuda son columnas FORMULA con plantilla *1.2
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
                // Indices (v2.1.0 con Funcion en 6): 3=Jira, 4=Facturar, 8=PDCL, 9=PDCL+Deuda.
                // Jira=8.0, Facturar=9.6, PDCL=9.6, PDCL+Deuda=9.6
                assertThat(mes.getRow(r).getCell(3).getNumericCellValue()).isEqualTo(8.0);
                CellValue real = evaluator.evaluate(mes.getRow(r).getCell(4));
                assertThat(real.getNumberValue()).isEqualTo(9.6);
                CellValue pdcl = evaluator.evaluate(mes.getRow(r).getCell(8));
                assertThat(pdcl.getNumberValue()).isEqualTo(9.6);
                CellValue pdclPlus = evaluator.evaluate(mes.getRow(r).getCell(9));
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
        // Resultado tiene exactamente 20 filas y nada nuevo.
        // (1 cabecera + 14 historicas + 3 regresion v1.6.2 + 1 v1.8.0 + 1 v1.8.1).
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet mes = wb.getSheet("Resultado");
            assertThat(mes.getLastRowNum() + 1).isEqualTo(20);
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
        // v2.0.0: tras el swap, mes.orphans.sourceSheet apunta a "Extraccion"
        // (antes "Cierre") — la hoja con imputaciones de Jira.
        content = content.replace("mes.orphans.sourceSheet=Extraccion",
                "mes.orphans.sourceSheet=HojaQueNoExiste");
        Files.writeString(cfgFile, content);

        ConfigLoader cfg = new ConfigLoader(cfgFile.toString());
        RunReport report = new RunReport();
        new ExcelMerger(cfg, report).merge();

        assertThat(report.warnings()).anyMatch(w ->
                "HOJA".equals(w.category) && w.message.contains("HojaQueNoExiste"));
        // Y Resultado sigue teniendo 20 filas (sin huerfanos)
        try (FileInputStream fis = new FileInputStream(outputFile.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {
            assertThat(wb.getSheet("Resultado").getLastRowNum() + 1).isEqualTo(20);
        }
    }

    @Test
    void mesColRealUsaFormulaConReferenciaAColumnaJira(@TempDir Path tmp) throws IOException {
        // Facturar es la columna 5 (index 4). Su plantilla es "{col:Jira}*1.2".
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

        // MES = 1 cabecera + 14 filas texto + 3 filas regresion v1.6.2 +
        //       1 regresion v1.8.0 (P-015) + 1 regresion v1.8.1 (P-016) = 20
        assertThat(report.sheets()).containsEntry("Resultado", 20);
        // Equipos = 1 cabecera + 10 entradas (de test-config.properties)
        assertThat(report.sheets()).containsEntry("Equipos", 11);
        // v2.0.0: la hoja con las peticiones del ERP (antes "Extraccion")
        // se llama ahora "Cierre".
        // Cierre = 1 cabecera + 14 texto + 3 regresion v1.6.2 +
        //          1 regresion v1.8.0 + 1 regresion v1.8.1 +
        //          1 Peticion vacia = 21
        assertThat(report.sheets()).containsEntry("Cierre", 21);
        // v2.0.0: la hoja con las imputaciones de Jira (antes "Cierre")
        // se llama ahora "Extraccion".
        // Extraccion = 1 meta + 1 cabecera + 16 historical + 5 regresion v1.6.2 +
        //              1 regresion v1.8.1 + 4 huerfanos v1.7.0 = 28
        assertThat(report.sheets()).containsEntry("Extraccion", 28);
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
            assertThat(resumen.getRow(2).getCell(2).getStringCellValue()).isEqualTo("Facturar");
            assertThat(resumen.getRow(2).getCell(3).getStringCellValue()).isEqualTo("PDCL");
            assertThat(resumen.getRow(2).getCell(4).getStringCellValue()).isEqualTo("PDCL + Deuda");

            // Al menos una fila de datos con SUMIFS sobre Resultado
            Row firstData = resumen.getRow(3);
            assertThat(firstData).isNotNull();
            String jiraFormula = firstData.getCell(1).getCellFormula();
            assertThat(jiraFormula).startsWith("SUMIFS(");
            assertThat(jiraFormula).contains("Resultado!");

            // Fila de totales de la PRIMERA tabla con SUM(B4:...).
            // Con v1.8.0 y summary.byResponsible.enabled=true, la hoja
            // Resumen tiene una segunda tabla debajo; por eso no podemos
            // mirar la ultima fila de la hoja entera. Buscamos la primera
            // fila "Total" desde la cabecera (fila 2) de arriba hacia abajo.
            int last = resumen.getLastRowNum();
            Row totals = null;
            for (int r = 3; r <= last; r++) {
                Row row = resumen.getRow(r);
                if (row == null) continue;
                org.apache.poi.ss.usermodel.Cell c0 = row.getCell(0);
                if (c0 != null
                        && c0.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING
                        && "Total".equals(c0.getStringCellValue())) {
                    totals = row;
                    break;
                }
            }
            assertThat(totals)
                    .as("Debe existir una fila 'Total' en la primera tabla de Resumen")
                    .isNotNull();
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

    // ==================================================================
    //  v1.8.0 — Segunda tabla en Resumen (matriz Matricula x Responsable)
    // ==================================================================

    @Test
    void byResponsiblePipelineGeneraSegundaTablaConTituloCorrecto(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resumen = wb.getSheet("Resumen");
            assertThat(resumen).isNotNull();

            // Localizar el titulo de la segunda tabla (no nos atamos a
            // una fila concreta para no fragilizarse ante cambios de fixtures).
            int titleRow0 = findRowByFirstCell(resumen,
                    "Totales Peticiones por Responsables Matrículas");
            assertThat(titleRow0)
                    .as("La segunda tabla con su titulo debe existir en Resumen")
                    .isGreaterThan(0);

            // Header 2 filas mas abajo: corner vacia, luego responsables, y Total al final.
            Row headerRow = resumen.getRow(titleRow0 + 2);
            assertThat(headerRow).isNotNull();
            // Esquina vacia
            assertThat(headerRow.getCell(0).getStringCellValue()).isEmpty();
            // 4 responsables normalizados a MAYUSCULAS en orden alfabetico:
            //   MG002    -> viene con padding "MG002   " (v1.8.1), trim lo limpia
            //   TRESP1@X -> union de tresp1@x (historicos) y TRESP1@x (v1.8.0)
            //   TRESP2@X -> tresp2@x (historicos)
            //   TRESP3@X -> tresp3@x (historicos)
            assertThat(headerRow.getCell(1).getStringCellValue()).isEqualTo("MG002");
            assertThat(headerRow.getCell(2).getStringCellValue()).isEqualTo("TRESP1@X");
            assertThat(headerRow.getCell(3).getStringCellValue()).isEqualTo("TRESP2@X");
            assertThat(headerRow.getCell(4).getStringCellValue()).isEqualTo("TRESP3@X");
            // Ultima columna: Total
            assertThat(headerRow.getCell(5).getStringCellValue()).isEqualTo("Total");
            // No hay 6a columna
            assertThat(headerRow.getCell(6)).isNull();
        }
    }

    @Test
    void byResponsibleSumifsCaseInsensitiveSumaTrespVariantes(@TempDir Path tmp) throws IOException {
        // Este test es el core de la leccion 1.7.1: evaluamos la formula
        // de verdad con FormulaEvaluator. La matricula 99641 aparece en
        // la fixture a traves de la fila de regresion 1.6.2 (138074/99641,
        // tresp1@x, Jira=9 -> PDCL=9*1.2=10.8). Debe ser 10.8 para la
        // celda (99641, TRESP1@X) de la segunda tabla.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resumen = wb.getSheet("Resumen");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            int titleRow0 = findRowByFirstCell(resumen,
                    "Totales Peticiones por Responsables Matrículas");
            int headerRow0 = titleRow0 + 2;
            int firstDataRow0 = headerRow0 + 1;

            // Buscar indice de columna del responsable TRESP1@X
            Row headerRow = resumen.getRow(headerRow0);
            int respColIdx = -1;
            for (int c = 1; c < headerRow.getLastCellNum(); c++) {
                Cell cell = headerRow.getCell(c);
                if (cell != null && "TRESP1@X".equals(cell.getStringCellValue())) {
                    respColIdx = c;
                    break;
                }
            }
            assertThat(respColIdx)
                    .as("Responsable TRESP1@X debe aparecer en la cabecera")
                    .isGreaterThan(0);

            // Buscar la fila de la matricula 99641
            int matrRowIdx = -1;
            int last = resumen.getLastRowNum();
            for (int r = firstDataRow0; r <= last; r++) {
                Row row = resumen.getRow(r);
                if (row == null) continue;
                Cell c0 = row.getCell(0);
                if (c0 == null) continue;
                if (c0.getCellType() != CellType.STRING) continue;
                if ("99641".equals(c0.getStringCellValue())) {
                    matrRowIdx = r;
                    break;
                }
            }
            assertThat(matrRowIdx)
                    .as("Matricula 99641 debe aparecer como fila en la segunda tabla")
                    .isGreaterThan(0);

            // Evaluar la celda (99641, TRESP1@X) y comprobar valor.
            // Resultado de la fixture: 138074/99641/tresp1@x tiene Jira=9
            // (via SUMIFS sobre Cierre: PROJ-23 con 9h/Dev). PDCL=9*1.2=10.8.
            // Al haber una unica imputacion con ese responsable/matricula,
            // la celda de la matriz debe ser exactamente 10.8.
            CellValue cellVal = evaluator.evaluate(
                    resumen.getRow(matrRowIdx).getCell(respColIdx));
            assertThat(cellVal.getNumberValue())
                    .as("(99641, TRESP1@X) PDCL = 9 (Jira) * 1.2 = 10.8")
                    .isCloseTo(10.8, org.assertj.core.data.Offset.offset(1e-9));
        }
    }

    @Test
    void byResponsibleTotalGlobalCuadraConPDCLGlobalDeLaPrimeraTabla(@TempDir Path tmp) throws IOException {
        // Check cruzado: el gran total de la segunda tabla (suma de
        // todos los PDCL, vista como sumatorio por responsable) debe
        // coincidir con el total PDCL de la primera tabla (sumatorio
        // por matricula).
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resumen = wb.getSheet("Resumen");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            // Encontrar la fila totals de la primera tabla (la primera
            // "Total" que aparece desde arriba).
            int firstTotalRow0 = findRowByFirstCell(resumen, "Total");
            assertThat(firstTotalRow0).isGreaterThan(0);
            // Primera tabla: columnas son Matricula, Jira, Facturar, PDCL, PDCL+Deuda.
            // PDCL es cell(3).
            double pdclGlobalTabla1 = evaluator.evaluate(
                    resumen.getRow(firstTotalRow0).getCell(3)).getNumberValue();

            // Encontrar la fila totals de la segunda tabla (la segunda "Total")
            int titleRow0 = findRowByFirstCell(resumen,
                    "Totales Peticiones por Responsables Matrículas");
            int last = resumen.getLastRowNum();
            int secondTotalRow0 = -1;
            for (int r = titleRow0 + 1; r <= last; r++) {
                Row row = resumen.getRow(r);
                if (row == null) continue;
                Cell c0 = row.getCell(0);
                if (c0 == null || c0.getCellType() != CellType.STRING) continue;
                if ("Total".equals(c0.getStringCellValue())) {
                    secondTotalRow0 = r;
                    break;
                }
            }
            assertThat(secondTotalRow0)
                    .as("La fila Total de la segunda tabla debe existir")
                    .isGreaterThan(titleRow0);

            // El gran total esta en la ultima columna ocupada de esa fila
            Row totalsRow2 = resumen.getRow(secondTotalRow0);
            int lastColIdx = totalsRow2.getLastCellNum() - 1;
            double grandTotal = evaluator.evaluate(totalsRow2.getCell(lastColIdx)).getNumberValue();

            assertThat(grandTotal)
                    .as("El gran total de la segunda tabla (PDCL via responsables) "
                            + "debe cuadrar con el total PDCL de la primera tabla "
                            + "(PDCL via matriculas). Ambos suman las mismas filas "
                            + "de Resultado, solo que agrupadas distinto.")
                    .isCloseTo(pdclGlobalTabla1,
                            org.assertj.core.data.Offset.offset(1e-6));
        }
    }

    // ==================================================================
    //  v1.8.1 — Trim del padding en Usuario_Resp_Tecnico
    // ==================================================================

    @Test
    void trimResponsableConPaddingHaceQueElSumifsDeResumenCasePorResponsable(
            @TempDir Path tmp) throws IOException {
        // Regresion v1.8.1 del bug reportado en produccion: el export real
        // del ERP trae Usuario_Resp_Tecnico con padding de espacios a la
        // derecha ("MG002   "). Sin el trim en la capa de copia, el SUMIFS
        // de la segunda tabla de Resumen comparaba la cabecera "MG002"
        // (normalizada, sin padding) contra la celda "MG002   " de
        // Resultado (con padding) y no casaba — devolvia 0.
        //
        // Con profile.Extraccion.trim.columns=Usuario_Resp_Tecnico activo,
        // la capa de copia aplica trim() al valor en Resultado, asi que
        // la celda queda "MG002" y el SUMIFS suma correctamente.
        //
        // Este test evalua la formula con FormulaEvaluator (no solo
        // inspecciona texto) — es precisamente el tipo de test que falto
        // en 1.8.0 y que permitio que el bug llegara a produccion.
        //
        // Fixture v1.8.1: Extraccion tiene P-016/M-1010 con responsable
        // "MG002   " (3 espacios). Cierre tiene PROJ-25 con
        // CN="P-016", Mat="M-1010", Funcion="Dev", Hours=5.
        // Jira esperado = 5, PDCL = 5 * 1.2 = 6.0.
        // En la matriz por responsable de Resumen, la celda
        // (M-1010, MG002) debe valer PDCL = 6.0.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            // Primero: verificar que el trim se aplicó. La celda
            // Res. Tecnico de la fila P-016 en Resultado debe ser "MG002"
            // exactamente, sin espacios.
            Sheet resultado = wb.getSheet("Resultado");
            int p016Row = -1;
            for (int r = 1; r <= resultado.getLastRowNum(); r++) {
                Cell c = resultado.getRow(r).getCell(0);
                if (c != null && c.getCellType() == CellType.STRING
                        && "P-016".equals(c.getStringCellValue())) {
                    p016Row = r;
                    break;
                }
            }
            assertThat(p016Row).as("Fila P-016 debe existir en Resultado").isGreaterThan(0);

            // Res. Tecnico es la col.8 del test-config (v2.1.0 con Funcion
            // en col.7). No col.9 del config raíz — el layout es distinto
            // porque test-config compacta columnas. col.8 -> indice 0-based = 7.
            Cell resTecnicoCell = resultado.getRow(p016Row).getCell(7);
            assertThat(resTecnicoCell.getCellType()).isEqualTo(CellType.STRING);
            assertThat(resTecnicoCell.getStringCellValue())
                    .as("Tras el trim, Res. Tecnico en Resultado no debe tener padding")
                    .isEqualTo("MG002");

            // Ahora: ir a Resumen y evaluar la celda (M-1010, MG002) de
            // la segunda tabla.
            Sheet resumen = wb.getSheet("Resumen");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            int titleRow0 = findRowByFirstCell(resumen,
                    "Totales Peticiones por Responsables Matrículas");
            int headerRow0 = titleRow0 + 2;

            // Buscar indice de columna del responsable MG002
            Row headerRow = resumen.getRow(headerRow0);
            int mg002ColIdx = -1;
            for (int c = 1; c < headerRow.getLastCellNum(); c++) {
                Cell cell = headerRow.getCell(c);
                if (cell != null && "MG002".equals(cell.getStringCellValue())) {
                    mg002ColIdx = c;
                    break;
                }
            }
            assertThat(mg002ColIdx)
                    .as("Responsable MG002 (tras el trim) debe aparecer en la cabecera")
                    .isGreaterThan(0);

            // Buscar fila de la matricula M-1010
            int mat1010Row = -1;
            int firstDataRow0 = headerRow0 + 1;
            for (int r = firstDataRow0; r <= resumen.getLastRowNum(); r++) {
                Row row = resumen.getRow(r);
                if (row == null) continue;
                Cell c0 = row.getCell(0);
                if (c0 == null) continue;
                if (c0.getCellType() != CellType.STRING) continue;
                if ("M-1010".equals(c0.getStringCellValue())) {
                    mat1010Row = r;
                    break;
                }
            }
            assertThat(mat1010Row)
                    .as("Matricula M-1010 (la nueva v1.8.1) debe aparecer como fila")
                    .isGreaterThan(0);

            // Evaluar la celda (M-1010, MG002) y verificar que vale 6.0.
            // Sin el fix del trim este valor seria 0 (bug reportado).
            CellValue cellVal = evaluator.evaluate(
                    resumen.getRow(mat1010Row).getCell(mg002ColIdx));
            assertThat(cellVal.getNumberValue())
                    .as("(M-1010, MG002) PDCL = 5 (Jira tras SUMIFS con Funcion=Dev) * 1.2 = 6.0. "
                            + "Si sale 0, el trim de Usuario_Resp_Tecnico no se aplico y el bug 1.8.1 vuelve.")
                    .isCloseTo(6.0, org.assertj.core.data.Offset.offset(1e-6));
        }
    }

    // ==================================================================
    //  v2.1.0 — columna Funcion en Resultado
    // ==================================================================

    @Test
    void resultadoIncluyeColumnaFuncionJustoDespuesDeMatricula(@TempDir Path tmp) throws IOException {
        // Regresion v2.1.0: la columna Funcion se inserta en el layout de
        // Resultado inmediatamente despues de Matricula. En test-config,
        // Matricula es col.6 (index 5) y Funcion es col.7 (index 6).
        // Verifica cabecera y al menos una fila con valor copiado de Cierre.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resultado = wb.getSheet("Resultado");

            // Cabecera
            Row header = resultado.getRow(0);
            assertThat(header.getCell(5).getStringCellValue()).isEqualTo("Matrícula");
            assertThat(header.getCell(6).getStringCellValue())
                    .as("Funcion debe estar en la posicion inmediatamente posterior a Matricula")
                    .isEqualTo("Funcion");

            // Al menos una fila trae el valor copiado desde Cierre.Funcion.
            // En el fixture realista (cierre.xlsx), todas las peticiones
            // P-001..P-016 tienen Funcion="Dev".
            assertThat(resultado.getRow(1).getCell(6).getStringCellValue())
                    .as("La primera fila de datos debe copiar la funcion de Cierre")
                    .isEqualTo("Dev");
        }
    }

    @Test
    void resultadoPdclYPdclMasDeudaSiguenCalculandoTrasDesplazarFuncion(@TempDir Path tmp) throws IOException {
        // Regresion v2.1.0 (regla inquebrantable 4: tests con FormulaEvaluator).
        // Al insertar Funcion en posicion 7 de test-config se desplazan los
        // indices de PDCL (antes col.8 -> ahora col.9, index 8) y PDCL+Deuda
        // (antes col.9 -> ahora col.10, index 9). Las formulas deben seguir
        // evaluando correctamente post-desplazamiento. Usamos P-001, que en
        // el fixture tiene Jira=5.0 (PROJ-1 3h + PROJ-2 2h = 5h, filtrado
        // Funcion=Dev deja fuera PROJ-3 Sup). Esperado: Facturar=6.0, PDCL=6.0,
        // PDCL+Deuda=6.0.
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resultado = wb.getSheet("Resultado");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            // P-001 es la primera fila de datos.
            Row row = resultado.getRow(1);
            assertThat(row.getCell(0).getStringCellValue()).isEqualTo("P-001");

            // Indices (v2.1.0): 3=Jira, 4=Facturar, 5=Matrícula, 6=Funcion,
            // 7=Res. Tecnico, 8=PDCL, 9=PDCL + Deuda.
            CellValue jira = evaluator.evaluate(row.getCell(3));
            assertThat(jira.getNumberValue())
                    .as("Jira SUMIFS filtrado por Funcion=Dev en P-001: PROJ-1 + PROJ-2 = 5h")
                    .isEqualTo(5.0);
            CellValue real = evaluator.evaluate(row.getCell(4));
            assertThat(real.getNumberValue()).isEqualTo(6.0);
            CellValue pdcl = evaluator.evaluate(row.getCell(8));
            assertThat(pdcl.getNumberValue())
                    .as("PDCL tras desplazarse al index 8 debe seguir evaluando Jira*1.2")
                    .isEqualTo(6.0);
            CellValue pdclPlus = evaluator.evaluate(row.getCell(9));
            assertThat(pdclPlus.getNumberValue())
                    .as("PDCL+Deuda tras desplazarse al index 9 debe seguir referenciando PDCL")
                    .isEqualTo(6.0);
        }
    }

    @Test
    void resultadoFuncionParaHuerfanoGuionSePropagaComoGuion(@TempDir Path tmp) throws IOException {
        // Decision fase 0 del usuario: cuando el valor de origen es "-",
        // la columna Funcion lo copia tal cual. El fixture tiene una fila
        // de cierre con todas las celdas a "-" (ultima fila). Esa fila
        // se omite en Resultado porque la ancla Peticion esta vacia, asi
        // que aqui lo comprobamos via el SUMIFS huerfano: la fila
        // TICKETS/- que introduce el modo orphans toma Funcion="-".
        ConfigLoader cfg = buildConfigWithOrphansEnabled(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resultado = wb.getSheet("Resultado");
            for (int r = 1; r <= resultado.getLastRowNum(); r++) {
                String pet = resultado.getRow(r).getCell(0).getStringCellValue();
                if (!"TICKETS".equals(pet)) continue;
                // Funcion es el index 6 en test-config v2.1.0. Los huerfanos
                // rellenan con "-" todas las columnas COPY sin dato de origen.
                assertThat(resultado.getRow(r).getCell(6).getStringCellValue())
                        .as("Funcion del huerfano TICKETS debe ser '-'")
                        .isEqualTo("-");
                return;
            }
            throw new AssertionError("No se encontro fila TICKETS en Resultado");
        }
    }

    // ==================================================================
    //  v2.2.0 — Tercer fichero opcional de Deuda (columna PDCL + Deuda)
    // ==================================================================

    /**
     * Con el fichero de Deuda presente, el libro de salida debe contener
     * la hoja "Deuda" entre Extraccion y Resultado, la fórmula de
     * "PDCL + Deuda" debe referenciar la hoja Deuda via SUMIFS y su
     * valor evaluado debe coincidir con PDCL + horas de deuda cruzadas.
     *
     * <p>Filas comprobadas (diseño del fixture deuda.xlsx):</p>
     * <ul>
     *   <li>P-001 / M-1001 / Dev: deuda=5+2=7, PDCL+Deuda=PDCL+7.</li>
     *   <li>P-002 / M-1002 / Dev: deuda=3, PDCL+Deuda=PDCL+3.</li>
     *   <li>P-007 / M-1001 / Dev: deuda=10, PDCL+Deuda=PDCL+10.</li>
     *   <li>P-003 / M-1003 / Dev: sin deuda, PDCL+Deuda=PDCL.</li>
     * </ul>
     */
    @Test
    void deudaFilePresenteSumaHorasEnPdclMasDeuda(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        TestFixtures.copyFixturesWithDeudaTo(inputDir);
        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(output.getParent());
        Path cfg = TestFixtures.renderTestConfig(
                tmp.resolve("test-config.properties"), inputDir, output);
        ConfigLoader cfgLoader = new ConfigLoader(cfg.toString());

        new ExcelMerger(cfgLoader, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(output.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            // La hoja Deuda esta presente en el libro.
            Sheet deuda = wb.getSheet("Deuda");
            assertThat(deuda).as("La hoja Deuda debe existir cuando se aporta el fichero").isNotNull();

            // Orden: Cierre, Extraccion, Deuda, Equipos, Resultado, Resumen.
            // Verificamos que Deuda viene ANTES que Resultado.
            int idxDeuda = wb.getSheetIndex("Deuda");
            int idxResultado = wb.getSheetIndex("Resultado");
            int idxExtraccion = wb.getSheetIndex("Extraccion");
            assertThat(idxDeuda)
                    .as("Deuda debe posicionarse entre Extraccion y Resultado")
                    .isGreaterThan(idxExtraccion)
                    .isLessThan(idxResultado);

            Sheet resultado = wb.getSheet("Resultado");
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            // Indices en test-config: 0=Petición, 3=Jira, 8=PDCL, 9=PDCL+Deuda.
            assertPdclPlusDeuda(resultado, evaluator, "P-001", 7.0);
            assertPdclPlusDeuda(resultado, evaluator, "P-002", 3.0);
            assertPdclPlusDeuda(resultado, evaluator, "P-007", 10.0);
            // P-003: no hay entrada en Deuda con esa clave -> delta=0
            assertPdclPlusDeuda(resultado, evaluator, "P-003", 0.0);

            // La formula escrita incluye la rama +IFERROR(SUMIFS(...),0).
            int rowP001 = findRowByFirstCell(resultado, "P-001");
            String formula = resultado.getRow(rowP001).getCell(9).getCellFormula();
            assertThat(formula)
                    .as("Con Deuda presente, PDCL+Deuda debe incluir +IFERROR(SUMIFS(...),0)")
                    .contains("+IFERROR(SUMIFS(")
                    .contains("Deuda!");
        }
    }

    /**
     * Sin el fichero de Deuda, el libro NO tiene hoja Deuda y la fórmula
     * de "PDCL + Deuda" se degrada silenciosamente a solo {@code {col:PDCL}}
     * — comportamiento idéntico a v2.1.0. Sin warnings nuevos.
     */
    @Test
    void sinFicheroDeudaComportamientoIdenticoAVersionAnterior(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        RunReport report = new RunReport();

        new ExcelMerger(cfg, report).merge();

        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        try (FileInputStream fis = new FileInputStream(output.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            assertThat(wb.getSheet("Deuda"))
                    .as("Sin fichero Deuda, la hoja Deuda NO debe existir en el libro")
                    .isNull();

            Sheet resultado = wb.getSheet("Resultado");
            int rowP001 = findRowByFirstCell(resultado, "P-001");
            String formula = resultado.getRow(rowP001).getCell(9).getCellFormula();
            assertThat(formula)
                    .as("Sin Deuda, PDCL+Deuda debe ser una referencia simple a la columna PDCL (sin SUMIFS)")
                    .doesNotContain("SUMIFS")
                    .doesNotContain("Deuda!");

            // Valor evaluado: PDCL + Deuda == PDCL.
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            double pdcl = ev.evaluate(resultado.getRow(rowP001).getCell(8)).getNumberValue();
            double pdclPlus = ev.evaluate(resultado.getRow(rowP001).getCell(9)).getNumberValue();
            assertThat(pdclPlus).isEqualTo(pdcl);

            // No debe aparecer warning nuevo relacionado con Deuda.
            assertThat(report.warnings())
                    .as("Sin fichero Deuda no se deben emitir warnings sobre la hoja/cabecera Deuda")
                    .noneMatch(w -> w.message.contains("Deuda"));
        }
    }

    /**
     * Fila sin match en Deuda: SUMIFS devuelve 0 naturalmente y por tanto
     * PDCL + Deuda == PDCL. El test verifica P-005 (existe en Cierre pero
     * no tiene entrada en deuda.xlsx).
     */
    @Test
    void deudaFilePresenteFilaSinMatchDevuelveSoloPdcl(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        TestFixtures.copyFixturesWithDeudaTo(inputDir);
        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(output.getParent());
        Path cfg = TestFixtures.renderTestConfig(
                tmp.resolve("test-config.properties"), inputDir, output);
        ConfigLoader cfgLoader = new ConfigLoader(cfg.toString());

        new ExcelMerger(cfgLoader, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(output.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resultado = wb.getSheet("Resultado");
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            assertPdclPlusDeuda(resultado, ev, "P-005", 0.0);
        }
    }

    /**
     * La fila F del fixture de Deuda tiene Matricula="-" para P-010.
     * En Resultado, P-010 tiene Matricula=M-1006, NO "-". El SUMIFS no
     * debe casar: delta=0 para esa fila, sin efecto colateral.
     */
    @Test
    void deudaFilePlaceholderMatriculaNoCruza(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        TestFixtures.copyFixturesWithDeudaTo(inputDir);
        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(output.getParent());
        Path cfg = TestFixtures.renderTestConfig(
                tmp.resolve("test-config.properties"), inputDir, output);
        ConfigLoader cfgLoader = new ConfigLoader(cfg.toString());

        new ExcelMerger(cfgLoader, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(output.toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resultado = wb.getSheet("Resultado");
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            assertPdclPlusDeuda(resultado, ev, "P-010", 0.0);
        }
    }

    /**
     * Verifica que {@code PDCL + Deuda} menos {@code PDCL} para la Peticion
     * indicada es igual al delta esperado.
     */
    private static void assertPdclPlusDeuda(Sheet resultado, FormulaEvaluator ev,
                                            String peticion, double expectedDelta) {
        int row = findRowByFirstCell(resultado, peticion);
        assertThat(row).as("Fila de " + peticion + " en Resultado").isGreaterThanOrEqualTo(0);
        double pdcl = ev.evaluate(resultado.getRow(row).getCell(8)).getNumberValue();
        double pdclPlus = ev.evaluate(resultado.getRow(row).getCell(9)).getNumberValue();
        assertThat(pdclPlus - pdcl)
                .as("Delta PDCL+Deuda - PDCL para " + peticion)
                .isEqualTo(expectedDelta);
    }

    // ==================================================================
    //  v2.2.0 — Rango [minFiles, maxFiles] y retrocompat strictTwoFiles
    // ==================================================================

    @Test
    void strictMinFilesConUnSoloExcelFalla(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        Files.createDirectories(inputDir);
        // Solo copiamos uno de los fixtures
        TestFixtures.copyFixturesTo(inputDir);
        Files.delete(inputDir.resolve("extraccion.xlsx"));

        Properties p = new Properties();
        p.setProperty("input.directory", inputDir.toString());
        p.setProperty("output.file", tmp.resolve("out.xlsx").toString());
        p.setProperty("input.strictMinFiles", "2");
        p.setProperty("input.strictMaxFiles", "3");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        assertThatThrownBy(() -> new ExcelMerger(cfg, new RunReport()).merge())
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("Se necesitan al menos 2");
    }

    /**
     * Con la clave legada {@code input.strictTwoFiles=true} y 3 ficheros,
     * la v2.2.0 debe preservar el contrato v2.1.0: abortar con "exactamente 2".
     */
    @Test
    void retrocompatStrictTwoFilesTrueConTresExcelAbortaIgualQueV210(@TempDir Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        TestFixtures.copyFixturesWithDeudaTo(inputDir);

        Properties p = new Properties();
        p.setProperty("input.directory", inputDir.toString());
        p.setProperty("output.file", tmp.resolve("out.xlsx").toString());
        p.setProperty("input.strictTwoFiles", "true");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);

        assertThatThrownBy(() -> new ExcelMerger(cfg, new RunReport()).merge())
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("exactamente 2");
    }

    /**
     * Busca la primera fila cuya celda(0) contiene exactamente el texto
     * dado. Devuelve -1 si no se encuentra.
     */
    private static int findRowByFirstCell(Sheet sheet, String text) {
        int last = sheet.getLastRowNum();
        for (int r = 0; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell c = row.getCell(0);
            if (c == null) continue;
            if (c.getCellType() != CellType.STRING) continue;
            if (text.equals(c.getStringCellValue())) return r;
        }
        return -1;
    }

    // ==================================================================
    //  v2.3.0 — Output mode (cierre / responsables / completo)
    // ==================================================================

    /**
     * El default ausente equivale a {@code cierre}: estructura identica al
     * happy path historico (ver {@link #mergeCompletoProduceTodasLasHojasEsperadas}).
     */
    @Test
    void outputModeCierreEsElDefaultYProduceLaMismaEstructuraQueV220(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithOutputMode(tmp, "cierre");
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {
            assertThat(wb.getNumberOfSheets()).isEqualTo(5);
            assertThat(wb.getSheet("Cierre")).isNotNull();
            assertThat(wb.getSheet("Extraccion")).isNotNull();
            assertThat(wb.getSheet("Equipos")).isNotNull();
            assertThat(wb.getSheet("Resultado")).isNotNull();
            assertThat(wb.getSheet("Resumen")).isNotNull();
        }
    }

    /**
     * En modo {@code responsables}, el output tiene Cierre, Extraccion,
     * Equipos (oculta), Resultado, y N hojas vacias por responsable
     * distinto de {@code Resultado.Res. Tecnico}. NO tiene Deuda ni Resumen.
     *
     * <p>Los fixtures generan estos responsables en Resultado: tresp1@x,
     * tresp2@x, tresp3@x, MG002 (sin padding tras trim de profile.Cierre),
     * y "TRESP1@x" en MAYUS que colapsa con tresp1@x via case-folding.
     * Total: 4 responsables distintos.</p>
     */
    @Test
    void outputModeResponsablesGeneraHojasPorResponsableYOmiteResumen(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithOutputMode(tmp, "responsables");
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            // Hojas que SI deben existir.
            assertThat(wb.getSheet("Cierre")).as("Cierre presente").isNotNull();
            assertThat(wb.getSheet("Extraccion")).as("Extraccion presente").isNotNull();
            assertThat(wb.getSheet("Equipos")).as("Equipos (lookup) presente").isNotNull();
            assertThat(wb.getSheet("Resultado")).as("Resultado presente").isNotNull();

            // Hojas que NO deben existir en modo responsables.
            assertThat(wb.getSheet("Resumen")).as("Resumen NO debe generarse").isNull();
            // Sin fichero Deuda en input por defecto (solo 2 ficheros), tampoco
            // habria Deuda en cierre. Aqui validamos solo la ausencia de Resumen.

            // Hojas por responsable: al menos las 4 esperadas.
            assertThat(wb.getSheet("tresp1@x")).as("hoja tresp1@x").isNotNull();
            assertThat(wb.getSheet("tresp2@x")).as("hoja tresp2@x").isNotNull();
            assertThat(wb.getSheet("tresp3@x")).as("hoja tresp3@x").isNotNull();
            assertThat(wb.getSheet("MG002")).as("hoja MG002 (tras trim)").isNotNull();

            // Cabecera A1 con el nombre canonico.
            Sheet juan = wb.getSheet("tresp1@x");
            assertThat(juan.getRow(0).getCell(0).getStringCellValue()).isEqualTo("tresp1@x");
        }
    }

    /**
     * En modo {@code responsables}, las hojas por responsable deben
     * posicionarse DESPUES de Resultado en el orden del libro. Equipos
     * (oculta) puede quedar entremedias por motivos de implementacion
     * (decision Fase 0 P2 opcion (a)).
     */
    @Test
    void outputModeResponsablesPosicionaHojasResponsableTrasResultado(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithOutputMode(tmp, "responsables");
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            int idxResultado = wb.getSheetIndex("Resultado");
            int idxTresp1 = wb.getSheetIndex("tresp1@x");
            int idxTresp2 = wb.getSheetIndex("tresp2@x");
            int idxTresp3 = wb.getSheetIndex("tresp3@x");
            int idxMG002 = wb.getSheetIndex("MG002");

            assertThat(idxTresp1).as("tresp1@x tras Resultado").isGreaterThan(idxResultado);
            assertThat(idxTresp2).as("tresp2@x tras Resultado").isGreaterThan(idxResultado);
            assertThat(idxTresp3).as("tresp3@x tras Resultado").isGreaterThan(idxResultado);
            assertThat(idxMG002).as("MG002 tras Resultado").isGreaterThan(idxResultado);
        }
    }

    /**
     * En modo {@code responsables} con fichero Deuda en el input, la hoja
     * Deuda NO se copia al libro de salida. La formula PDCL+Deuda en
     * Resultado se degrada al modo "sin Deuda" (igual que cuando el usuario
     * no aporta el 3er fichero), comportamiento heredado de v2.2.0.
     */
    @Test
    void outputModeResponsablesOmiteCopiaDeDeudaIncluso3FicherosEnInput(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithDeudaAndOutputMode(tmp, "responsables");
        RunReport report = new RunReport();
        new ExcelMerger(cfg, report).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            assertThat(wb.getSheet("Deuda")).as("Deuda NO debe estar en el output").isNull();
            assertThat(wb.getSheet("Resumen")).as("Resumen NO debe estar en el output").isNull();
            assertThat(wb.getSheet("Resultado")).isNotNull();

            // La formula PDCL+Deuda no debe contener referencia a Deuda.
            Sheet resultado = wb.getSheet("Resultado");
            int rowP001 = findRowByFirstCell(resultado, "P-001");
            // Indice 9 = PDCL + Deuda en test-config.
            String formula = resultado.getRow(rowP001).getCell(9).getCellFormula();
            assertThat(formula).doesNotContain("Deuda!");

            // Y un warning CONFIG sobre la omision.
            assertThat(report.warnings())
                    .anyMatch(w -> "CONFIG".equals(w.category)
                            && w.message.contains("RESPONSABLES")
                            && w.message.contains("Deuda"));
        }
    }

    /**
     * En modo {@code completo}, el output contiene todas las hojas del modo
     * cierre (incluida Deuda y Resumen) MAS las hojas por responsable.
     */
    @Test
    void outputModeCompletoIncluyeTodasLasHojasDeCierreYDeResponsables(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithDeudaAndOutputMode(tmp, "completo");
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            // Todas las hojas del modo cierre (con Deuda).
            assertThat(wb.getSheet("Cierre")).isNotNull();
            assertThat(wb.getSheet("Extraccion")).isNotNull();
            assertThat(wb.getSheet("Deuda")).as("Deuda SI debe estar en completo").isNotNull();
            assertThat(wb.getSheet("Equipos")).isNotNull();
            assertThat(wb.getSheet("Resultado")).isNotNull();
            assertThat(wb.getSheet("Resumen")).as("Resumen SI debe estar en completo").isNotNull();

            // Hojas por responsable.
            assertThat(wb.getSheet("tresp1@x")).isNotNull();
            assertThat(wb.getSheet("tresp2@x")).isNotNull();
            assertThat(wb.getSheet("tresp3@x")).isNotNull();
            assertThat(wb.getSheet("MG002")).isNotNull();
        }
    }

    /**
     * En modo {@code completo}, las hojas por responsable se colocan despues
     * de Resumen.
     */
    @Test
    void outputModeCompletoPosicionaResponsablesTrasResumen(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithDeudaAndOutputMode(tmp, "completo");
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            int idxResumen = wb.getSheetIndex("Resumen");
            assertThat(wb.getSheetIndex("tresp1@x")).isGreaterThan(idxResumen);
            assertThat(wb.getSheetIndex("tresp2@x")).isGreaterThan(idxResumen);
            assertThat(wb.getSheetIndex("tresp3@x")).isGreaterThan(idxResumen);
            assertThat(wb.getSheetIndex("MG002")).isGreaterThan(idxResumen);
        }
    }

    /**
     * El RunReport registra el modo efectivo usado.
     */
    @Test
    void outputModeQuedaRegistradoEnRunReport(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithOutputMode(tmp, "responsables");
        RunReport report = new RunReport();
        new ExcelMerger(cfg, report).merge();

        assertThat(report.getOutputMode()).isEqualTo(OutputMode.RESPONSABLES);
        assertThat(report.formatSummary(java.time.Duration.ofMillis(1)))
                .contains("Modo: RESPONSABLES");
    }

    /**
     * Modo invalido en config con strictValidation=true (default) aborta
     * el run via ConfigValidator antes de llegar a ExcelMerger. El test
     * unitario de ConfigValidator cubre el contenido del mensaje de error;
     * aqui solo verificamos que efectivamente ConfigValidator detecta el
     * error.
     */
    @Test
    void outputModeInvalidoProduceErrorEnConfigValidator(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithOutputMode(tmp, "TODO_VAL");

        java.util.List<String> errors = new ConfigValidator(cfg).validate();
        assertThat(errors)
                .anyMatch(e -> e.contains("output.mode")
                        && e.contains("TODO_VAL"));
    }

    /**
     * Con valor invalido y strictValidation=false, ExcelMerger cae a CIERRE
     * de forma defensiva (en lugar de lanzar excepcion en runtime).
     */
    @Test
    void outputModeInvalidoConStrictValidationFalseCaeACierre(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithOutputMode(tmp, "TODO_VAL");
        // Simulamos lo que hace Main cuando strictValidation=false: ignora
        // los errores del validador y deja correr al merger. ExcelMerger
        // debe caer a CIERRE.
        RunReport report = new RunReport();
        new ExcelMerger(cfg, report).merge();

        // Estructura de modo cierre (5 hojas).
        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {
            assertThat(wb.getNumberOfSheets()).isEqualTo(5);
            assertThat(wb.getSheet("Resumen")).isNotNull();
        }
        assertThat(report.getOutputMode()).isEqualTo(OutputMode.CIERRE);
    }

    // ==================================================================
    //  v2.4.0 — Tablas pivot por hoja de responsable
    // ==================================================================

    /**
     * En modo {@code responsables}, cada hoja de responsable debe contener
     * dos tablas pivot SUMIFS (Jira y Facturar) apiladas verticalmente, ademas
     * de la cabecera A1 (v2.3.0).
     *
     * <p>Verifica la estructura: titulo Jira, cabecera con matriculas,
     * datos, totales, gap, titulo Facturar, cabecera, datos, totales.</p>
     */
    @Test
    void v240ResponsablesGeneraDosTablasPivotPorHoja(@TempDir Path tmp) throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithOutputMode(tmp, "responsables");
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheet("tresp1@x");
            assertThat(sheet).as("hoja tresp1@x").isNotNull();

            // Cabecera A1
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("tresp1@x");

            // Titulo de la primera tabla en row 2 (0-based)
            String title1 = sheet.getRow(2).getCell(0).getStringCellValue();
            assertThat(title1).as("titulo tabla Jira").contains("Jira");

            // La cabecera de la tabla Jira debe estar en row 3 con "Petición"
            // como primer encabezado.
            assertThat(sheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("Petición");

            // Hay una segunda tabla en algún punto más abajo cuyo título
            // contiene "Facturar". La buscamos sin asumir índice exacto.
            int facturarTitleRow = -1;
            for (int r = 4; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell c0 = row.getCell(0);
                if (c0 == null) continue;
                if (c0.getCellType() == CellType.STRING
                        && c0.getStringCellValue().contains("Facturar")
                        && c0.getStringCellValue().contains("Petición")) {
                    facturarTitleRow = r;
                    break;
                }
            }
            assertThat(facturarTitleRow).as("titulo tabla Facturar encontrado").isPositive();
        }
    }

    /**
     * Verifica que las fórmulas SUMIFS de las pivots evalúan a los valores
     * esperados según los datos de los fixtures (lección 1.7.1).
     *
     * <p>Combinacion conocida usada: tresp1@x / P-001 / M-1001.
     * En cierre.xlsx hay una fila P-001/Recurso=M-1001/Responsable=tresp1@x
     * con Horas_RealizadoTot=12. Es la única ocurrencia de esa combinación
     * para tresp1@x.</p>
     *
     * <p>Como Jira y Facturar en Resultado se calculan a partir de cruces, el
     * valor exacto depende del fixture. Para no acoplarse a aritmética
     * exacta, este test verifica únicamente: (a) que el SUMIFS evalúa sin
     * error, (b) que el valor es no negativo, (c) que coincide con la
     * suma manual de las celdas correspondientes en Resultado filtradas
     * por responsable=tresp1@x, peticion=P-001, matricula=M-1001.</p>
     */
    @Test
    void v240ResponsablesFormulaEvaluatorSumifsCoincideConSumaManual(@TempDir Path tmp)
            throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithOutputMode(tmp, "responsables");
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheet("tresp1@x");
            assertThat(sheet).isNotNull();

            // Calcular la suma manual a partir de Resultado.
            Sheet resultado = wb.getSheet("Resultado");
            int colPet = findHeaderIndex(resultado, "Petición");
            int colMat = findHeaderIndex(resultado, "Matrícula");
            int colResp = findHeaderIndex(resultado, "Res. Tecnico");
            int colJira = findHeaderIndex(resultado, "Jira");
            assertThat(colPet).as("col Petición").isNotNegative();
            assertThat(colMat).as("col Matrícula").isNotNegative();
            assertThat(colResp).as("col Res. Tecnico").isNotNegative();
            assertThat(colJira).as("col Jira").isNotNegative();

            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();

            double manualSum = 0;
            for (int r = 1; r <= resultado.getLastRowNum(); r++) {
                Row row = resultado.getRow(r);
                if (row == null) continue;
                String pet = stringOf(row.getCell(colPet));
                String mat = stringOf(row.getCell(colMat));
                String resp = stringOf(row.getCell(colResp));
                if (!"P-001".equals(pet)) continue;
                if (!"M-1001".equals(mat)) continue;
                // SUMIFS de Excel es case-insensitive, asi que usamos
                // equalsIgnoreCase(trim) para replicar exactamente el mismo
                // comportamiento. NO usar contains: harìa match espurio.
                if (resp == null || !"tresp1@x".equalsIgnoreCase(resp.trim())) continue;
                Cell jc = row.getCell(colJira);
                if (jc == null) continue;
                CellValue cv = ev.evaluate(jc);
                if (cv != null && cv.getCellType() == CellType.NUMERIC) {
                    manualSum += cv.getNumberValue();
                }
            }

            // Localizar P-001 en la tabla Jira de tresp1@x y leer la
            // columna M-1001. Buscamos el row donde A.value == "P-001"
            // entre las filas de la primera pivot (que empieza en row 2
            // con titulo + cabecera + datos). La cabecera de matriculas
            // está en row 3.
            Row matrHeaderRow = sheet.getRow(3);
            int m1001Col = -1;
            for (int c = 1; c < matrHeaderRow.getLastCellNum(); c++) {
                Cell hc = matrHeaderRow.getCell(c);
                if (hc != null && "M-1001".equals(hc.getStringCellValue())) {
                    m1001Col = c;
                    break;
                }
            }
            assertThat(m1001Col).as("columna M-1001 en tabla Jira").isPositive();

            int p001Row = -1;
            for (int r = 4; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell c0 = row.getCell(0);
                if (c0 == null) continue;
                // Parar al llegar al "Total" o a la fila de gap
                if (c0.getCellType() == CellType.STRING
                        && "Total".equals(c0.getStringCellValue())) {
                    break;
                }
                if (c0.getCellType() == CellType.STRING
                        && "P-001".equals(c0.getStringCellValue())) {
                    p001Row = r;
                    break;
                }
            }
            assertThat(p001Row).as("fila P-001 en tabla Jira").isPositive();

            CellValue pivotValue = ev.evaluate(sheet.getRow(p001Row).getCell(m1001Col));
            assertThat(pivotValue.getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(pivotValue.getNumberValue())
                    .as("SUMIFS pivot tresp1@x P-001 x M-1001 == suma manual")
                    .isEqualTo(manualSum);
        }
    }

    /** Helper local para encontrar índice de columna por nombre de cabecera. */
    private static int findHeaderIndex(Sheet sheet, String headerName) {
        Row hdr = sheet.getRow(0);
        if (hdr == null) return -1;
        for (int c = 0; c < hdr.getLastCellNum(); c++) {
            Cell cell = hdr.getCell(c);
            if (cell == null) continue;
            if (cell.getCellType() == CellType.STRING
                    && headerName.equals(cell.getStringCellValue())) {
                return c;
            }
        }
        return -1;
    }

    /** Helper local que devuelve el contenido de una celda como String. */
    private static String stringOf(Cell cell) {
        if (cell == null) return null;
        if (cell.getCellType() == CellType.STRING) return cell.getStringCellValue();
        if (cell.getCellType() == CellType.NUMERIC) {
            double d = cell.getNumericCellValue();
            return (d == (long) d) ? String.valueOf((long) d) : String.valueOf(d);
        }
        return null;
    }

    /**
     * En modo {@code completo}, las hojas de responsable también deben tener
     * sus dos tablas pivot, sin afectar a la generación de Resumen.
     */
    @Test
    void v240CompletoTieneResumenYTablasPivotEnHojasResponsable(@TempDir Path tmp)
            throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithOutputMode(tmp, "completo");
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            assertThat(wb.getSheet("Resumen")).as("Resumen presente en completo").isNotNull();

            Sheet sheet = wb.getSheet("tresp1@x");
            assertThat(sheet).as("hoja tresp1@x").isNotNull();

            // Verifica presencia de las dos pivots: titulo Jira y titulo Facturar.
            String title1 = sheet.getRow(2).getCell(0).getStringCellValue();
            assertThat(title1).contains("Jira");

            // Buscar el segundo titulo más abajo con "Facturar"
            boolean foundFacturar = false;
            for (int r = 4; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell c0 = row.getCell(0);
                if (c0 == null || c0.getCellType() != CellType.STRING) continue;
                if (c0.getStringCellValue().contains("Facturar")
                        && c0.getStringCellValue().contains("Petición")) {
                    foundFacturar = true;
                    break;
                }
            }
            assertThat(foundFacturar).as("segunda tabla Facturar en hoja de responsable").isTrue();
        }
    }

    /**
     * En modo {@code cierre}, no se generan hojas de responsable y por tanto
     * tampoco tablas pivot. El comportamiento debe ser idéntico al de v2.3.0.
     */
    @Test
    void v240CierreNoTieneTablasPivotPorqueNoTieneHojasResponsable(@TempDir Path tmp)
            throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfigWithOutputMode(tmp, "cierre");
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            // En cierre, las hojas de responsable NO existen.
            assertThat(wb.getSheet("tresp1@x")).as("no hay hoja tresp1@x en cierre").isNull();
            assertThat(wb.getSheet("tresp2@x")).isNull();
            assertThat(wb.getSheet("tresp3@x")).isNull();
            // Pero Resumen y Resultado sí.
            assertThat(wb.getSheet("Resumen")).isNotNull();
            assertThat(wb.getSheet("Resultado")).isNotNull();
        }
    }

    /**
     * Si {@code responsables.tables.enabled=false}, las hojas de responsable
     * en modo responsables/completo deben quedar como en v2.3.0
     * (solo cabecera A1, sin pivots).
     */
    @Test
    void v240FlagOffMantieneHojasResponsableComoV230(@TempDir Path tmp) throws IOException {
        // Construir el config base con modo responsables y luego sobreescribir
        // la flag de tablas a false manualmente sobre el fichero ya renderizado.
        ConfigLoader baseCfg = TestFixtures.buildRealisticConfigWithOutputMode(tmp, "responsables");
        Path cfgPath = tmp.resolve("test-config.properties");
        String content = Files.readString(cfgPath);
        // Reemplazar la línea responsables.tables.enabled=true por =false
        content = content.replaceAll(
                "(?m)^responsables\\.tables\\.enabled=.*$",
                "responsables.tables.enabled=false");
        // Si no estaba presente, añadirla.
        if (!content.contains("responsables.tables.enabled=false")) {
            content += System.lineSeparator() + "responsables.tables.enabled=false";
        }
        Files.writeString(cfgPath, content);
        ConfigLoader cfg = new ConfigLoader(cfgPath.toString());
        // Aseguramos que el override realmente se aplicó
        assertThat(cfg.getBoolean("responsables.tables.enabled", true)).isFalse();

        // baseCfg ya no se usa: lo retenemos para mantener test legible
        assertThat(baseCfg).isNotNull();

        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = wb.getSheet("tresp1@x");
            assertThat(sheet).isNotNull();
            // Solo fila 0 (cabecera A1).
            assertThat(sheet.getLastRowNum()).as("solo cabecera A1, sin pivots").isZero();
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("tresp1@x");
        }
    }

    // ==================================================================
    //  v2.6.0 — Modif 1: la columna Horas_RealizadoTot de Resultado
    //           proviene de Cierre.Total_Horas_Realizadas_Recurso
    //           (no de Cierre.Horas_RealizadoTot, fuente hasta v2.5.1).
    // ==================================================================

    /**
     * v2.6.0 (Modif 1): verifica que la columna {@code Horas_RealizadoTot} de
     * Resultado contiene los valores de {@code Cierre.Total_Horas_Realizadas_Recurso}
     * y NO los de {@code Cierre.Horas_RealizadoTot}.
     *
     * <p>El fixture {@code cierre.xlsx} tiene valores diferenciables en ambas
     * columnas (p. ej. P-001: Horas_RealizadoTot=12 vs
     * Total_Horas_Realizadas_Recurso=20; P-002: 30 vs 35; etc.). El test
     * comprueba que en Resultado aparece el valor de la columna
     * <i>Total_Horas_Realizadas_Recurso</i>, lo que solo es posible si el
     * config aplica {@code mes.col.N.from=Total_Horas_Realizadas_Recurso}.</p>
     */
    @Test
    void v260HorasRealizadoTotProvieneDeTotalHorasRealizadasRecurso(@TempDir Path tmp)
            throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resultado = wb.getSheet("Resultado");
            assertThat(resultado).as("hoja Resultado").isNotNull();

            int colPet = findHeaderIndex(resultado, "Petición");
            int colHRT = findHeaderIndex(resultado, "Horas_RealizadoTot");
            assertThat(colPet).as("col Petición").isNotNegative();
            assertThat(colHRT).as("col Horas_RealizadoTot").isNotNegative();

            // Mapeo Peticion -> valor esperado tomado de la columna
            // Total_Horas_Realizadas_Recurso del fixture cierre.xlsx
            // (NO de Horas_RealizadoTot, que es la fuente vieja).
            // Datos del fixture (ver gen_fixtures.py):
            //   P-001: HRT=12  THRR=20   <- esperamos 20 (Modif 1)
            //   P-002: HRT=30  THRR=35   <- esperamos 35
            //   P-007: HRT=40  THRR=60   <- esperamos 60
            // Valores que descartarian la fuente vieja: 12, 30, 40 NO deben aparecer.
            java.util.Map<String, Double> esperados = new java.util.HashMap<>();
            esperados.put("P-001", 20.0);
            esperados.put("P-002", 35.0);
            esperados.put("P-007", 60.0);

            int comprobadas = 0;
            for (int r = 1; r <= resultado.getLastRowNum(); r++) {
                Row row = resultado.getRow(r);
                if (row == null) continue;
                Cell pc = row.getCell(colPet);
                if (pc == null || pc.getCellType() != CellType.STRING) continue;
                String pet = pc.getStringCellValue();
                Double exp = esperados.get(pet);
                if (exp == null) continue;
                Cell hrtCell = row.getCell(colHRT);
                assertThat(hrtCell).as("celda Horas_RealizadoTot fila " + pet).isNotNull();
                double actual = readNumericOrParse(hrtCell);
                assertThat(actual)
                        .as("Horas_RealizadoTot[" + pet + "] debe venir de "
                                + "Total_Horas_Realizadas_Recurso (Modif 1 v2.6.0), "
                                + "no de Horas_RealizadoTot (fuente vieja).")
                        .isEqualTo(exp);
                comprobadas++;
            }
            assertThat(comprobadas)
                    .as("se han verificado las 3 peticiones del fixture P-001/P-002/P-007")
                    .isEqualTo(3);
        }
    }

    // ==================================================================
    //  v2.6.0 — Modif 2: la columna Realizadas_Horas_Mes copia tal cual
    //           desde Cierre, sea cero o no.
    // ==================================================================

    /**
     * v2.6.0 (Modif 2): el diagnóstico de v2.6.0 concluyó que NO hay bug en
     * {@code CopyColumnStrategy}: la columna {@code Realizadas_Horas_Mes} sale
     * toda a 0 en el output del usuario porque su ERP rellena esa columna con
     * "0.00" en el 100% de las 832 filas del export real. Es realidad del ERP,
     * no bug del programa.
     *
     * <p>Este test ancla el contrato: la columna de Resultado debe coincidir
     * celda a celda con la columna correspondiente de Cierre, sea el valor
     * cero o no. El fixture sintético tiene ambos casos (P-001=5, P-003=0,
     * P-002=15, etc.).</p>
     */
    @Test
    void v260RealizadasHorasMesCopiaTalCualDesdeCierre(@TempDir Path tmp)
            throws IOException {
        ConfigLoader cfg = TestFixtures.buildRealisticConfig(tmp);
        new ExcelMerger(cfg, new RunReport()).merge();

        try (FileInputStream fis = new FileInputStream(
                tmp.resolve("output").resolve("resultado.xlsx").toFile());
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet resultado = wb.getSheet("Resultado");
            assertThat(resultado).as("hoja Resultado").isNotNull();

            int colPet = findHeaderIndex(resultado, "Petición");
            int colRHM = findHeaderIndex(resultado, "Realizadas_Horas_Mes");
            assertThat(colPet).as("col Petición").isNotNegative();
            assertThat(colRHM).as("col Realizadas_Horas_Mes").isNotNegative();

            // Datos del fixture cierre.xlsx (ver gen_fixtures.py):
            //   P-001 -> 5    (no-cero, comprueba que el COPY trae el valor)
            //   P-002 -> 15   (no-cero)
            //   P-003 -> 0    (cero legítimo, comprueba que tambien se copia)
            //   P-008 -> 0    (cero legítimo)
            //   P-007 -> 15
            // El test comprueba ambos casos (no-cero y cero) para confirmar
            // que CopyColumnStrategy hace exactamente lo que debe.
            java.util.Map<String, Double> esperados = new java.util.HashMap<>();
            esperados.put("P-001", 5.0);
            esperados.put("P-002", 15.0);
            esperados.put("P-003", 0.0);
            esperados.put("P-007", 15.0);
            esperados.put("P-008", 0.0);

            int comprobadas = 0;
            for (int r = 1; r <= resultado.getLastRowNum(); r++) {
                Row row = resultado.getRow(r);
                if (row == null) continue;
                Cell pc = row.getCell(colPet);
                if (pc == null || pc.getCellType() != CellType.STRING) continue;
                String pet = pc.getStringCellValue();
                Double exp = esperados.get(pet);
                if (exp == null) continue;
                Cell rhmCell = row.getCell(colRHM);
                assertThat(rhmCell).as("celda Realizadas_Horas_Mes fila " + pet).isNotNull();
                double actual = readNumericOrParse(rhmCell);
                assertThat(actual)
                        .as("Realizadas_Horas_Mes[" + pet + "] debe copiarse tal cual desde "
                                + "Cierre.Realizadas_Horas_Mes (sea cero o no).")
                        .isEqualTo(exp);
                comprobadas++;
            }
            assertThat(comprobadas)
                    .as("se han verificado las 5 peticiones del fixture")
                    .isEqualTo(5);
        }
    }

    /**
     * Helper local: devuelve el valor numérico de una celda. Acepta tanto
     * {@code NUMERIC} (caso natural) como {@code STRING} con un literal
     * parseable como double (caso ERP real, donde valores numéricos vienen
     * formateados como "12.00" en sharedStrings).
     */
    private static double readNumericOrParse(Cell cell) {
        if (cell == null) return Double.NaN;
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        if (cell.getCellType() == CellType.STRING) {
            String s = cell.getStringCellValue().trim();
            return Double.parseDouble(s);
        }
        if (cell.getCellType() == CellType.FORMULA) {
            // El cache de la fórmula puede tener el valor evaluado (lección 1.7.1):
            // si hay un FormulaEvaluator disponible, esto se evaluaría primero.
            // Aquí sólo apoyamos el cache. Las columnas COPY no son FORMULA, así
            // que esta rama es defensiva.
            return cell.getNumericCellValue();
        }
        return Double.NaN;
    }
}

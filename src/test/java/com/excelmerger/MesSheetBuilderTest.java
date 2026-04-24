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
 * Tests de {@link MesSheetBuilder} enfocados en la resolucion de
 * placeholders, la sensibilidad a mayusculas/minusculas en {@code {col:X}}
 * y el comportamiento de skip cuando la celda ancla esta vacia.
 *
 * <p>El builder no expone la logica de resolucion; se testea via el
 * comportamiento observable: construir el workbook y leer las formulas
 * que escribe en la hoja MES.</p>
 */
class MesSheetBuilderTest {

    /**
     * Monta un workbook con una hoja "Extraccion" que tiene cabeceras en la
     * fila 1 y datos a partir de la 2. {@code anchors} define que celda ira
     * en la columna ancla de cada fila: una cadena no vacia hace que la fila
     * se incluya en MES, una cadena vacia hace que se salte.
     */
    private static Workbook buildSourceWorkbook(String... anchors) {
        Workbook wb = new XSSFWorkbook();
        Sheet s = wb.createSheet("Extraccion");

        Row header = s.createRow(0);
        header.createCell(0).setCellValue("Peticion");
        header.createCell(1).setCellValue("Aplicaci_Activi");

        for (int i = 0; i < anchors.length; i++) {
            Row r = s.createRow(i + 1);
            r.createCell(0).setCellValue(anchors[i]);
            r.createCell(1).setCellValue("APP" + i);
        }
        return wb;
    }

    private static Properties baseMesConfig() {
        Properties p = new Properties();
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sheetName", "MES");
        p.setProperty("mes.sourceSheet", "Extraccion");
        p.setProperty("mes.sourceHeaderRow", "1");
        p.setProperty("mes.anchorColumn", "Peticion");
        return p;
    }

    // ==================================================================
    //  Skip de filas con ancla vacia
    // ==================================================================

    @Test
    void saltaFilasDeOrigenConAnclaVacia() throws Exception {
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1", "", "P-3", "P-4", "")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            assertThat(mes).isNotNull();

            // Cabecera + 3 filas utiles (P-1, P-3, P-4). Las dos con ancla vacia
            // se saltan.
            assertThat(mes.getLastRowNum()).isEqualTo(3);
            assertThat(mes.getRow(1).getCell(0).getStringCellValue()).isEqualTo("P-1");
            assertThat(mes.getRow(2).getCell(0).getStringCellValue()).isEqualTo("P-3");
            assertThat(mes.getRow(3).getCell(0).getStringCellValue()).isEqualTo("P-4");
        }
    }

    // ==================================================================
    //  Placeholder {col:Nombre} - resolucion basica
    // ==================================================================

    @Test
    void placeholderColReferenciaLetraDeColumnaMesYFilaActual() throws Exception {
        // MES tendra: col A = Peticion (COPY), col B = Valor (COPY), col C = Doble (FORMULA {col:Valor}*2)
        // Para la fila 2 de MES (primera fila de datos), {col:Valor} -> B2 y la formula queda "B2*2".
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");

        p.setProperty("mes.col.2.name", "Valor");
        p.setProperty("mes.col.2.type", "EMPTY");

        p.setProperty("mes.col.3.name", "Doble");
        p.setProperty("mes.col.3.type", "FORMULA");
        p.setProperty("mes.col.3.formula", "{col:Valor}*2");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1", "P-2")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            // Fila 2 (primera de datos): formula B2*2
            assertThat(mes.getRow(1).getCell(2).getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(mes.getRow(1).getCell(2).getCellFormula()).isEqualTo("B2*2");
            // Fila 3 (segunda de datos): formula B3*2
            assertThat(mes.getRow(2).getCell(2).getCellFormula()).isEqualTo("B3*2");
        }
    }

    @Test
    void placeholderColResuelveMultiplesOcurrencias() throws Exception {
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");

        p.setProperty("mes.col.2.name", "A");
        p.setProperty("mes.col.2.type", "EMPTY");
        p.setProperty("mes.col.3.name", "B");
        p.setProperty("mes.col.3.type", "EMPTY");

        p.setProperty("mes.col.4.name", "Resta");
        p.setProperty("mes.col.4.type", "FORMULA");
        p.setProperty("mes.col.4.formula", "{col:A}-{col:B}");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("X")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            // A esta en col B (indice 1), B esta en col C (indice 2)
            assertThat(mes.getRow(1).getCell(3).getCellFormula()).isEqualTo("B2-C2");
        }
    }

    // ==================================================================
    //  Placeholder {col:X} es CASE-SENSITIVE
    // ==================================================================

    @Test
    void placeholderColDistingueEntreRealYREAL() throws Exception {
        // Definimos DOS columnas que solo difieren en mayusculas: "Real" y "REAL".
        // Una formula referencia {col:Real} y otra {col:REAL}; cada una debe
        // apuntar a su columna correspondiente, no confundirse.
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");

        p.setProperty("mes.col.2.name", "Real");
        p.setProperty("mes.col.2.type", "EMPTY");

        p.setProperty("mes.col.3.name", "REAL");
        p.setProperty("mes.col.3.type", "EMPTY");

        p.setProperty("mes.col.4.name", "RefMinus");
        p.setProperty("mes.col.4.type", "FORMULA");
        p.setProperty("mes.col.4.formula", "{col:Real}+0");

        p.setProperty("mes.col.5.name", "RefMayus");
        p.setProperty("mes.col.5.type", "FORMULA");
        p.setProperty("mes.col.5.formula", "{col:REAL}+0");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            // Petición=A, Real=B, REAL=C, RefMinus=D, RefMayus=E
            // RefMinus usa {col:Real} -> B2
            assertThat(mes.getRow(1).getCell(3).getCellFormula()).isEqualTo("B2+0");
            // RefMayus usa {col:REAL} -> C2
            assertThat(mes.getRow(1).getCell(4).getCellFormula()).isEqualTo("C2+0");
        }
    }

    @Test
    void placeholderColNoMatcheaPorCasePermisivo() throws Exception {
        // Solo existe "Real" (sin REAL). Una formula con {col:REAL} no casa
        // -> la columna se queda disabled y la celda se escribe como BLANK,
        // ademas de anadir un warning FORMULA.
        //
        // NOTE: El codigo esta documentado como "CASE-SENSITIVE" a proposito
        // en MesSheetBuilder (ver comentario sobre colByName). El enunciado
        // de la Sesion C lo califica de "bug"; este test fija el
        // comportamiento ACTUAL para que cualquier cambio futuro sea
        // intencional. Ver CHANGELOG para la bandera de revision.
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");

        p.setProperty("mes.col.2.name", "Real");
        p.setProperty("mes.col.2.type", "EMPTY");

        p.setProperty("mes.col.3.name", "Ref");
        p.setProperty("mes.col.3.type", "FORMULA");
        p.setProperty("mes.col.3.formula", "{col:REAL}*2"); // mayusculas, no matcha

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            // La celda queda en blanco (columna disabled)
            assertThat(mes.getRow(1).getCell(2).getCellType()).isEqualTo(CellType.BLANK);
        }

        assertThat(report.warnings()).anyMatch(w ->
                "FORMULA".equals(w.category)
                        && w.message.contains("{col:REAL}"));
    }

    // ==================================================================
    //  Placeholder {colLetter:X}
    // ==================================================================

    @Test
    void placeholderColLetterUsaLaLetraLiteralYFilaActual() throws Exception {
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");

        p.setProperty("mes.col.2.name", "LetraDirecta");
        p.setProperty("mes.col.2.type", "FORMULA");
        // {colLetter:Z} -> Z<fila-actual>
        p.setProperty("mes.col.2.formula", "{colLetter:Z}*10");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1", "P-2")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            assertThat(mes.getRow(1).getCell(1).getCellFormula()).isEqualTo("Z2*10");
            assertThat(mes.getRow(2).getCell(1).getCellFormula()).isEqualTo("Z3*10");
        }
    }

    @Test
    void placeholderColLetterSeNormalizaAMayusculas() throws Exception {
        // El codigo hace letter.toUpperCase() -> "aa" acaba siendo "AA".
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");

        p.setProperty("mes.col.2.name", "F");
        p.setProperty("mes.col.2.type", "FORMULA");
        p.setProperty("mes.col.2.formula", "{colLetter:aa}+1");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("X")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            assertThat(mes.getRow(1).getCell(1).getCellFormula()).isEqualTo("AA2+1");
        }
    }

    // ==================================================================
    //  Combinacion de ambos placeholders en una sola formula
    // ==================================================================

    @Test
    void combinacionDePlaceholdersEnLaMismaFormula() throws Exception {
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");

        p.setProperty("mes.col.2.name", "Valor");
        p.setProperty("mes.col.2.type", "EMPTY");

        p.setProperty("mes.col.3.name", "Calc");
        p.setProperty("mes.col.3.type", "FORMULA");
        p.setProperty("mes.col.3.formula", "{col:Valor}+{colLetter:Z}");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            assertThat(mes.getRow(1).getCell(2).getCellFormula()).isEqualTo("B2+Z2");
        }
    }

    @Test
    void formulaConIgualInicialQuitaElIgual() throws Exception {
        // Si la plantilla empieza con '=' (costumbre humana), el builder
        // lo elimina porque POI no lo quiere al llamar a setCellFormula.
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");

        p.setProperty("mes.col.2.name", "A");
        p.setProperty("mes.col.2.type", "EMPTY");

        p.setProperty("mes.col.3.name", "B");
        p.setProperty("mes.col.3.type", "FORMULA");
        p.setProperty("mes.col.3.formula", "={col:A}*2");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("X")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            // La formula registrada NO lleva '=' al principio
            assertThat(mes.getRow(1).getCell(2).getCellFormula()).isEqualTo("B2*2");
        }
    }

    // ==================================================================
    //  Comportamientos de skip y omision
    // ==================================================================

    @Test
    void mesDeshabilitadoNoHaceNada() throws Exception {
        Properties p = baseMesConfig();
        p.setProperty("mes.enabled", "false");
        p.setProperty("mes.col.1.name", "X");
        p.setProperty("mes.col.1.type", "EMPTY");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1")) {
            new MesSheetBuilder(cfg, report).build(wb);

            assertThat(wb.getSheet("MES")).isNull();
            assertThat(report.warnings()).isEmpty();
        }
    }

    @Test
    void sourceSheetInexistenteGeneraWarningHOJA() throws Exception {
        Properties p = baseMesConfig();
        p.setProperty("mes.sourceSheet", "NoExiste");
        p.setProperty("mes.col.1.name", "X");
        p.setProperty("mes.col.1.type", "EMPTY");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1")) {
            new MesSheetBuilder(cfg, report).build(wb);

            assertThat(wb.getSheet("MES")).isNull();
        }

        assertThat(report.warnings()).anyMatch(w ->
                "HOJA".equals(w.category) && w.message.contains("NoExiste"));
    }

    @Test
    void columnaAnclaInexistenteGeneraWarningCABECERA() throws Exception {
        Properties p = baseMesConfig();
        p.setProperty("mes.anchorColumn", "NoExisteEstaCabecera");
        p.setProperty("mes.col.1.name", "X");
        p.setProperty("mes.col.1.type", "EMPTY");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1")) {
            new MesSheetBuilder(cfg, report).build(wb);

            assertThat(wb.getSheet("MES")).isNull();
        }

        assertThat(report.warnings()).anyMatch(w ->
                "CABECERA".equals(w.category)
                        && w.message.contains("NoExisteEstaCabecera"));
    }

    @Test
    void hojaMesYaExistenteSeOmite() throws Exception {
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "X");
        p.setProperty("mes.col.1.type", "EMPTY");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1")) {
            wb.createSheet("MES"); // pre-existente
            new MesSheetBuilder(cfg, report).build(wb);

            // La MES original (creada por el test) sigue sin filas fisicas
            Sheet mes = wb.getSheet("MES");
            assertThat(mes.getPhysicalNumberOfRows()).isZero();
        }

        assertThat(report.warnings()).anyMatch(w ->
                "HOJA".equals(w.category) && w.message.contains("MES"));
    }

    @Test
    void copyColumnaInexistenteDegradaAEmpty() throws Exception {
        // MesColumn.fromConfig degrada COPY sin 'from' a EMPTY directamente,
        // pero COPY con 'from' que no existe en la hoja origen produce una
        // celda BLANK (via preValidate disabled + writeCell blank) y un warning.
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "X");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "NoExiste");

        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            assertThat(mes.getRow(1).getCell(0).getCellType()).isEqualTo(CellType.BLANK);
        }

        assertThat(report.warnings()).anyMatch(w ->
                "CABECERA".equals(w.category) && w.message.contains("NoExiste"));
    }

    // ==================================================================
    //  v1.6.0: mes.col.N.fill y mes.col.N.redIfNotEqualTo
    // ==================================================================

    @Test
    void columnaConFillAplicaEstiloDeFondoAlasCeldasDeDatos() throws Exception {
        Properties p = baseMesConfig();
        // Columna 1: ancla (COPY de Peticion, obligatoria para que el loop no corte)
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");
        // Columna 2 con fill LIGHT_GREEN
        p.setProperty("mes.col.2.name", "PDCL");
        p.setProperty("mes.col.2.type", "FORMULA");
        p.setProperty("mes.col.2.formula", "{colLetter:A}2*2");
        p.setProperty("mes.col.2.fill", "LIGHT_GREEN");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1", "P-2")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            assertThat(mes).isNotNull();
            // La celda (fila 1, col 1) debe tener un estilo con fill solido
            org.apache.poi.ss.usermodel.CellStyle style = mes.getRow(1).getCell(1).getCellStyle();
            assertThat(style).isNotNull();
            assertThat(style.getFillPattern())
                    .isEqualTo(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);
            // Fila 2 tambien, mismo estilo (cacheado)
            org.apache.poi.ss.usermodel.CellStyle style2 = mes.getRow(2).getCell(1).getCellStyle();
            assertThat(style2.getFillPattern())
                    .isEqualTo(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);
        }
    }

    @Test
    void fillConColorDesconocidoGeneraWarningYNoAplicaEstilo() throws Exception {
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");
        p.setProperty("mes.col.2.name", "X");
        p.setProperty("mes.col.2.type", "EMPTY");
        p.setProperty("mes.col.2.fill", "RAINBOW");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1")) {
            new MesSheetBuilder(cfg, report).build(wb);
        }

        assertThat(report.warnings()).anyMatch(w ->
                "CONFIG".equals(w.category) && w.message.contains("RAINBOW"));
    }

    @Test
    void redIfNotEqualToAñadeFormatoCondicionalSobreLaColumna() throws Exception {
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");
        // Columna 2 = PDCL, columna 3 = PDCL Deuda con redIfNotEqualTo=PDCL
        p.setProperty("mes.col.2.name", "PDCL");
        p.setProperty("mes.col.2.type", "FORMULA");
        p.setProperty("mes.col.2.formula", "{colLetter:A}2");
        p.setProperty("mes.col.3.name", "PDCL Deuda");
        p.setProperty("mes.col.3.type", "FORMULA");
        p.setProperty("mes.col.3.formula", "{col:PDCL}");
        p.setProperty("mes.col.3.redIfNotEqualTo", "PDCL");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1", "P-2", "P-3")) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("MES");
            assertThat(mes).isNotNull();
            org.apache.poi.ss.usermodel.SheetConditionalFormatting scf =
                    mes.getSheetConditionalFormatting();
            // Tras construir la hoja, debe haber al menos 1 regla CF
            assertThat(scf.getNumConditionalFormattings()).isGreaterThanOrEqualTo(1);

            // Alguna de las reglas debe tener la formula "C2<>B2"
            boolean found = false;
            for (int i = 0; i < scf.getNumConditionalFormattings(); i++) {
                org.apache.poi.ss.usermodel.ConditionalFormatting cf =
                        scf.getConditionalFormattingAt(i);
                for (int r = 0; r < cf.getNumberOfRules(); r++) {
                    String f = cf.getRule(r).getFormula1();
                    if (f != null && f.contains("<>")) {
                        found = true;
                    }
                }
            }
            assertThat(found)
                    .as("Se esperaba una regla CF con operador '<>' en la formula")
                    .isTrue();
        }
    }

    @Test
    void redIfNotEqualToApuntandoAColumnaInexistenteGeneraWarning() throws Exception {
        Properties p = baseMesConfig();
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");
        p.setProperty("mes.col.2.name", "PDCL Deuda");
        p.setProperty("mes.col.2.type", "FORMULA");
        p.setProperty("mes.col.2.formula", "{colLetter:A}2");
        p.setProperty("mes.col.2.redIfNotEqualTo", "NoExiste");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWorkbook("P-1")) {
            new MesSheetBuilder(cfg, report).build(wb);
        }

        assertThat(report.warnings()).anyMatch(w ->
                "CONFIG".equals(w.category)
                        && w.message.contains("redIfNotEqualTo")
                        && w.message.contains("NoExiste"));
    }

    // ==================================================================
    //  v2.1.0 — columna Funcion (COPY desde Cierre.Funcion)
    // ==================================================================

    /**
     * Monta un workbook origen con las columnas minimas necesarias para
     * probar la semantica de la nueva columna Funcion: Peticion (ancla),
     * Recurso (Matricula), Funcion. Cada fila se pasa como un array
     * {peticion, recurso, funcion}. Una peticion vacia salta la fila.
     */
    private static Workbook buildSourceWithFuncion(String[]... rows) {
        Workbook wb = new XSSFWorkbook();
        Sheet s = wb.createSheet("Cierre");

        Row header = s.createRow(0);
        header.createCell(0).setCellValue("Peticion");
        header.createCell(1).setCellValue("Recurso");
        header.createCell(2).setCellValue("Funcion");

        for (int i = 0; i < rows.length; i++) {
            Row r = s.createRow(i + 1);
            r.createCell(0).setCellValue(rows[i][0]);
            r.createCell(1).setCellValue(rows[i][1]);
            r.createCell(2).setCellValue(rows[i][2]);
        }
        return wb;
    }

    private static Properties mesConfigWithFuncion() {
        Properties p = new Properties();
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sheetName", "Resultado");
        p.setProperty("mes.sourceSheet", "Cierre");
        p.setProperty("mes.sourceHeaderRow", "1");
        p.setProperty("mes.anchorColumn", "Peticion");
        // Layout minimo alineado con la v2.1.0: Peticion, Matricula, Funcion.
        // Matricula en col.2 y Funcion en col.3 — el contrato "Funcion justo
        // despues de Matricula" se testea aqui.
        p.setProperty("mes.col.1.name", "Petición");
        p.setProperty("mes.col.1.type", "COPY");
        p.setProperty("mes.col.1.from", "Peticion");
        p.setProperty("mes.col.2.name", "Matrícula");
        p.setProperty("mes.col.2.type", "COPY");
        p.setProperty("mes.col.2.from", "Recurso");
        p.setProperty("mes.col.3.name", "Funcion");
        p.setProperty("mes.col.3.type", "COPY");
        p.setProperty("mes.col.3.from", "Funcion");
        return p;
    }

    @Test
    void funcionCopiaValorDesdeCierreTalCual() throws Exception {
        // Caso basico: la celda Funcion de Resultado copia exactamente el
        // valor de Cierre.Funcion, sin transformaciones.
        Properties p = mesConfigWithFuncion();
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWithFuncion(
                new String[]{"P-1", "M-1001", "AN"},
                new String[]{"P-2", "M-1002", "DI"},
                new String[]{"P-3", "M-1003", "PR"})) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("Resultado");
            assertThat(mes).isNotNull();
            // Cabecera: col 0=Petición, 1=Matrícula, 2=Funcion
            assertThat(mes.getRow(0).getCell(2).getStringCellValue()).isEqualTo("Funcion");
            // Valores copiados tal cual
            assertThat(mes.getRow(1).getCell(2).getStringCellValue()).isEqualTo("AN");
            assertThat(mes.getRow(2).getCell(2).getStringCellValue()).isEqualTo("DI");
            assertThat(mes.getRow(3).getCell(2).getStringCellValue()).isEqualTo("PR");
        }
    }

    @Test
    void funcionPreservaGuionComoValorLiteral() throws Exception {
        // Decision del usuario (v2.1.0, fase 0, pregunta 4): si la celda
        // origen trae "-", la celda destino tambien trae "-". No hay
        // normalizacion a vacio ni filtrado.
        Properties p = mesConfigWithFuncion();
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWithFuncion(
                new String[]{"P-1", "M-1001", "AN"},
                new String[]{"P-2", "-",     "-"})) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("Resultado");
            assertThat(mes.getRow(2).getCell(1).getStringCellValue()).isEqualTo("-");
            assertThat(mes.getRow(2).getCell(2).getStringCellValue()).isEqualTo("-");
        }
    }

    @Test
    void funcionGeneraUnaFilaPorCombinacionMatriculaFuncion() throws Exception {
        // Semantica B3 confirmada en Fase 0: una misma matricula con N
        // funciones distintas en Cierre genera N filas en Resultado, porque
        // Resultado es una fila por peticion+recurso+funcion original de
        // Cierre. Este test lo blinda: misma matricula M-1001, tres
        // funciones (AN, DI, PR) -> tres filas distintas en Resultado, todas
        // con la misma matricula pero distinta funcion.
        Properties p = mesConfigWithFuncion();
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        RunReport report = new RunReport();

        try (Workbook wb = buildSourceWithFuncion(
                new String[]{"P-1", "M-1001", "AN"},
                new String[]{"P-2", "M-1001", "DI"},
                new String[]{"P-3", "M-1001", "PR"})) {
            new MesSheetBuilder(cfg, report).build(wb);

            Sheet mes = wb.getSheet("Resultado");
            // 1 cabecera + 3 filas de datos
            assertThat(mes.getLastRowNum()).isEqualTo(3);

            // Las 3 filas comparten matricula pero difieren en funcion
            assertThat(mes.getRow(1).getCell(1).getStringCellValue()).isEqualTo("M-1001");
            assertThat(mes.getRow(2).getCell(1).getStringCellValue()).isEqualTo("M-1001");
            assertThat(mes.getRow(3).getCell(1).getStringCellValue()).isEqualTo("M-1001");

            assertThat(mes.getRow(1).getCell(2).getStringCellValue()).isEqualTo("AN");
            assertThat(mes.getRow(2).getCell(2).getStringCellValue()).isEqualTo("DI");
            assertThat(mes.getRow(3).getCell(2).getStringCellValue()).isEqualTo("PR");
        }
    }
}

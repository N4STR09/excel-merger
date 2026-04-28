package com.excelmerger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Properties;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests unitarios del helper {@link ResponsablePivotBuilder} (v2.4.0).
 *
 * <p>Construyen a mano un workbook con una hoja {@code Resultado} mínima
 * + una hoja destino para el responsable, llaman al builder y verifican
 * estructura, fórmulas SUMIFS bien formadas, y evaluación con
 * {@link FormulaEvaluator} (lección 1.7.1: nunca confiar en una fórmula
 * sin haberla evaluado).</p>
 *
 * <p>Las columnas en Resultado se montan idénticas a las que produce
 * {@code MesSheetBuilder} en producción (orden y nombres):
 * Petición(A), Matrícula(B), Res. Tecnico(C), Jira(D), REAL(E).</p>
 */
class ResponsablePivotBuilderTest {

    /** Letras de columna usadas en TODOS los tests de esta clase. */
    private static final String LETTER_PETICION = "A";
    private static final String LETTER_MATRICULA = "B";
    private static final String LETTER_RESPONSABLE = "C";
    private static final String LETTER_JIRA = "D";
    private static final String LETTER_REAL = "E";

    private static final String JIRA_TITLE = "Horas imputadas (Jira) por Petición × Matrícula";

    /**
     * Crea un workbook con una hoja {@code Resultado} con cabeceras y filas
     * de datos para los tests de evaluación.
     */
    private Workbook newWorkbookWithResultado() {
        Workbook wb = new XSSFWorkbook();
        Sheet res = wb.createSheet("Resultado");
        // Cabecera
        Row h = res.createRow(0);
        h.createCell(0).setCellValue("Petición");
        h.createCell(1).setCellValue("Matrícula");
        h.createCell(2).setCellValue("Res. Tecnico");
        h.createCell(3).setCellValue("Jira");
        h.createCell(4).setCellValue("REAL");
        return wb;
    }

    /** Añade una fila de datos a Resultado (peticion+matricula como STRING, hours como NUMERIC). */
    private void addResultadoRow(Workbook wb, String peticion, String matricula,
                                 String responsable, double jira, double real) {
        Sheet res = wb.getSheet("Resultado");
        int next = res.getLastRowNum() + 1;
        Row r = res.createRow(next);
        r.createCell(0).setCellValue(peticion);
        r.createCell(1).setCellValue(matricula);
        r.createCell(2).setCellValue(responsable);
        r.createCell(3).setCellValue(jira);
        r.createCell(4).setCellValue(real);
    }

    /** ConfigLoader mínimo (no se usa en el helper salvo para construcción). */
    private ConfigLoader minimalConfig() {
        return TestFixtures.configFromProperties(new Properties());
    }

    /** Construye una hoja destino con A1 (responsable) ya rellena. */
    private Sheet createTargetSheet(Workbook wb, String responsableName) {
        Sheet sheet = wb.createSheet(responsableName);
        Row r = sheet.createRow(0);
        r.createCell(0).setCellValue(responsableName);
        return sheet;
    }

    private ResponsablePivotBuilder.PivotInputs jiraInputs(List<String> peticiones,
                                                            List<String> matriculas) {
        return new ResponsablePivotBuilder.PivotInputs(
                "Resultado",
                LETTER_PETICION, LETTER_MATRICULA, LETTER_RESPONSABLE, LETTER_JIRA,
                JIRA_TITLE,
                peticiones, matriculas,
                10000);
    }

    // ==================================================================
    //  Estructura básica
    // ==================================================================

    @Test
    void escribeCabeceraTituloMatriculasYTotal() throws java.io.IOException {
        try (Workbook wb = newWorkbookWithResultado()) {
            Sheet sheet = createTargetSheet(wb, "tresp1@x");
            ResponsablePivotBuilder builder = new ResponsablePivotBuilder(minimalConfig());

            ResponsablePivotBuilder.PivotResult result = builder.writeTable(
                    wb, sheet, 2,
                    jiraInputs(Arrays.asList("P-001", "P-002"), Arrays.asList("M-1001", "M-1002")));

            // Fila título (row 2 0-based = row 3 Excel)
            Row titleRow = sheet.getRow(2);
            assertThat(titleRow.getCell(0).getStringCellValue()).isEqualTo(JIRA_TITLE);

            // Fila cabecera (row 3 0-based)
            Row headerRow = sheet.getRow(3);
            assertThat(headerRow.getCell(0).getStringCellValue()).isEqualTo("Petición");
            assertThat(headerRow.getCell(1).getStringCellValue()).isEqualTo("M-1001");
            assertThat(headerRow.getCell(2).getStringCellValue()).isEqualTo("M-1002");
            // Total al final: índice = 1 + nM = 3
            assertThat(headerRow.getCell(3).getStringCellValue()).isEqualTo("Total");

            // 2 filas de datos (P-001, P-002)
            assertThat(sheet.getRow(4).getCell(0).getStringCellValue()).isEqualTo("P-001");
            assertThat(sheet.getRow(5).getCell(0).getStringCellValue()).isEqualTo("P-002");

            // Fila Total (row 6 0-based)
            Row totalsRow = sheet.getRow(6);
            assertThat(totalsRow.getCell(0).getStringCellValue()).isEqualTo("Total");

            // PivotResult bien calculado
            assertThat(result.firstRow0).isEqualTo(2);
            assertThat(result.lastRow0).isEqualTo(6);
            assertThat(result.sumifsCount).isEqualTo(4); // 2 peticiones x 2 matriculas
        }
    }

    @Test
    void formulaSumifsBienFormadaCon3Criterios() throws java.io.IOException {
        try (Workbook wb = newWorkbookWithResultado()) {
            Sheet sheet = createTargetSheet(wb, "tresp1@x");
            ResponsablePivotBuilder builder = new ResponsablePivotBuilder(minimalConfig());

            builder.writeTable(wb, sheet, 2,
                    jiraInputs(Collections.singletonList("P-001"),
                               Collections.singletonList("M-1001")));

            // La celda B5 (Excel) = row 4 col 1 (0-based) debe ser SUMIFS con 3 criterios.
            Cell c = sheet.getRow(4).getCell(1);
            assertThat(c.getCellType()).isEqualTo(CellType.FORMULA);

            String formula = c.getCellFormula();
            // Forma esperada (sin espacios extras):
            //   SUMIFS(Resultado!D2:D10000,
            //          Resultado!A2:A10000,$A5,
            //          Resultado!B2:B10000,B$4,
            //          Resultado!C2:C10000,$A$1)
            assertThat(formula).contains("SUMIFS(");
            assertThat(formula).contains("Resultado!D2:D10000");      // sum_range = Jira
            assertThat(formula).contains("Resultado!A2:A10000");      // peticion range
            assertThat(formula).contains("$A5");                      // peticion criterio (cell)
            assertThat(formula).contains("Resultado!B2:B10000");      // matricula range
            assertThat(formula).contains("B$4");                      // matricula criterio (header)
            assertThat(formula).contains("Resultado!C2:C10000");      // responsable range
            assertThat(formula).contains("$A$1");                     // responsable criterio (A1)
        }
    }

    @Test
    void formulaSumDeFilaYColumnaTotal() throws java.io.IOException {
        try (Workbook wb = newWorkbookWithResultado()) {
            Sheet sheet = createTargetSheet(wb, "tresp1@x");
            ResponsablePivotBuilder builder = new ResponsablePivotBuilder(minimalConfig());

            builder.writeTable(wb, sheet, 2,
                    jiraInputs(Arrays.asList("P-001", "P-002"),
                               Arrays.asList("M-1001", "M-1002")));

            // Total de fila P-001 (col 3 = "D") en row 4 (0-based)
            Cell rowTotal = sheet.getRow(4).getCell(3);
            assertThat(rowTotal.getCellFormula()).isEqualTo("SUM(B5:C5)");

            // Total de columna M-1001 en row 6 (0-based, fila Excel 7), col 1
            Cell colTotal = sheet.getRow(6).getCell(1);
            assertThat(colTotal.getCellFormula()).isEqualTo("SUM(B5:B6)");

            // Gran total en row 6 col 3
            Cell grandTotal = sheet.getRow(6).getCell(3);
            assertThat(grandTotal.getCellFormula()).isEqualTo("SUM(D5:D6)");
        }
    }

    // ==================================================================
    //  Caso degenerado
    // ==================================================================

    @Test
    void responsableSinPeticionesEscribeSoloTituloYSinDatos() throws java.io.IOException {
        try (Workbook wb = newWorkbookWithResultado()) {
            Sheet sheet = createTargetSheet(wb, "tresp1@x");
            ResponsablePivotBuilder builder = new ResponsablePivotBuilder(minimalConfig());

            ResponsablePivotBuilder.PivotResult result = builder.writeTable(
                    wb, sheet, 5,
                    jiraInputs(Collections.<String>emptyList(),
                               Collections.singletonList("M-1001")));

            // Fila título (row 5)
            assertThat(sheet.getRow(5).getCell(0).getStringCellValue()).isEqualTo(JIRA_TITLE);
            // Fila siguiente "(Sin datos)" (row 6)
            assertThat(sheet.getRow(6).getCell(0).getStringCellValue()).isEqualTo("(Sin datos)");

            assertThat(result.firstRow0).isEqualTo(5);
            assertThat(result.lastRow0).isEqualTo(6);
            assertThat(result.sumifsCount).isZero();
        }
    }

    // ==================================================================
    //  Evaluación con FormulaEvaluator (LECCIÓN 1.7.1)
    // ==================================================================

    @Test
    void formulaEvaluatorSumaCorrectaParaCombinacionConocida() throws java.io.IOException {
        try (Workbook wb = newWorkbookWithResultado()) {
            // Datos: tresp1@x tiene P-001/M-1001 con jira=5, P-001/M-1001 con jira=3
            //                            P-001/M-1002 con jira=2
            //                            P-002/M-1001 con jira=10
            // tresp2@x tiene P-001/M-1001 con jira=99 (NO debe contarse en pivot de tresp1@x)
            addResultadoRow(wb, "P-001", "M-1001", "tresp1@x",  5, 6);
            addResultadoRow(wb, "P-001", "M-1001", "tresp1@x",  3, 3.6);
            addResultadoRow(wb, "P-001", "M-1002", "tresp1@x",  2, 2.4);
            addResultadoRow(wb, "P-002", "M-1001", "tresp1@x", 10, 12);
            addResultadoRow(wb, "P-001", "M-1001", "tresp2@x", 99, 99); // OTRO responsable

            Sheet sheet = createTargetSheet(wb, "tresp1@x");
            ResponsablePivotBuilder builder = new ResponsablePivotBuilder(minimalConfig());

            builder.writeTable(wb, sheet, 2,
                    jiraInputs(Arrays.asList("P-001", "P-002"),
                               Arrays.asList("M-1001", "M-1002")));

            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            // P-001/M-1001/tresp1@x = 5 + 3 = 8
            CellValue v = evaluator.evaluate(sheet.getRow(4).getCell(1));
            assertThat(v.getNumberValue())
                    .as("P-001 x M-1001 filtrado por tresp1@x")
                    .isEqualTo(8.0);

            // P-001/M-1002/tresp1@x = 2
            assertThat(evaluator.evaluate(sheet.getRow(4).getCell(2)).getNumberValue())
                    .as("P-001 x M-1002").isEqualTo(2.0);

            // P-002/M-1001/tresp1@x = 10
            assertThat(evaluator.evaluate(sheet.getRow(5).getCell(1)).getNumberValue())
                    .as("P-002 x M-1001").isEqualTo(10.0);

            // P-002/M-1002/tresp1@x = 0 (no hay fila)
            assertThat(evaluator.evaluate(sheet.getRow(5).getCell(2)).getNumberValue())
                    .as("P-002 x M-1002 vacio").isZero();

            // Total de fila P-001 = 8 + 2 = 10
            assertThat(evaluator.evaluate(sheet.getRow(4).getCell(3)).getNumberValue())
                    .as("Total fila P-001").isEqualTo(10.0);

            // Total columna M-1001 = 8 + 10 = 18
            assertThat(evaluator.evaluate(sheet.getRow(6).getCell(1)).getNumberValue())
                    .as("Total columna M-1001").isEqualTo(18.0);

            // Gran total = 8 + 2 + 10 + 0 = 20
            assertThat(evaluator.evaluate(sheet.getRow(6).getCell(3)).getNumberValue())
                    .as("Gran total").isEqualTo(20.0);
        }
    }

    @Test
    void formulaEvaluatorRespetaCaseSinIgnorarFiltrosNumericos() throws java.io.IOException {
        // Combinación con matrícula y petición numéricas (regresión 1.7.1):
        // si la cabecera "55751" se escribiera como NUMERIC en lugar de STRING,
        // el SUMIFS daría 0 contra una columna asText. Verificamos que pasa.
        try (Workbook wb = newWorkbookWithResultado()) {
            // En Resultado, columna Petición es STRING (configurada como asText)
            addResultadoRow(wb, "55751", "99641", "tresp1@x", 7, 8.4);
            addResultadoRow(wb, "55751", "99641", "tresp1@x", 3, 3.6);
            addResultadoRow(wb, "55752", "99641", "tresp1@x", 1, 1.2);

            Sheet sheet = createTargetSheet(wb, "tresp1@x");
            ResponsablePivotBuilder builder = new ResponsablePivotBuilder(minimalConfig());

            builder.writeTable(wb, sheet, 2,
                    jiraInputs(Arrays.asList("55751", "55752"),
                               Collections.singletonList("99641")));

            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            // 55751 x 99641 = 7 + 3 = 10
            assertThat(evaluator.evaluate(sheet.getRow(4).getCell(1)).getNumberValue())
                    .as("Peticion 55751 x Matricula 99641 (numericas as STRING)")
                    .isEqualTo(10.0);

            // 55752 x 99641 = 1
            assertThat(evaluator.evaluate(sheet.getRow(5).getCell(1)).getNumberValue())
                    .as("Peticion 55752 x Matricula 99641")
                    .isEqualTo(1.0);
        }
    }

    @Test
    void respectaCriterioResponsableViaA1() throws java.io.IOException {
        // Si cambiamos el contenido de A1 después de generar la pivot, el SUMIFS
        // debe evaluar con el nuevo valor — confirmando que el criterio
        // realmente apunta a $A$1 y no se hardcodea como literal.
        try (Workbook wb = newWorkbookWithResultado()) {
            addResultadoRow(wb, "P-001", "M-1001", "tresp1@x", 5, 6);
            addResultadoRow(wb, "P-001", "M-1001", "tresp2@x", 9, 11);

            Sheet sheet = createTargetSheet(wb, "tresp1@x");
            ResponsablePivotBuilder builder = new ResponsablePivotBuilder(minimalConfig());
            builder.writeTable(wb, sheet, 2,
                    jiraInputs(Collections.singletonList("P-001"),
                               Collections.singletonList("M-1001")));

            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            // Con A1=tresp1@x, valor = 5
            assertThat(evaluator.evaluate(sheet.getRow(4).getCell(1)).getNumberValue())
                    .isEqualTo(5.0);

            // Cambiamos A1 a tresp2@x y reevaluamos
            sheet.getRow(0).getCell(0).setCellValue("tresp2@x");
            evaluator.clearAllCachedResultValues();
            assertThat(evaluator.evaluate(sheet.getRow(4).getCell(1)).getNumberValue())
                    .as("Cambiando A1 cambia el filtro -> 9")
                    .isEqualTo(9.0);
        }
    }

    @Test
    void columnaMatriculaSeEscribeComoString() throws java.io.IOException {
        // Lección 1.7.1: aunque la matrícula sea todo dígitos, la cabecera
        // se escribe como STRING porque la columna Matrícula de Resultado
        // es asText. Si se escribiera NUMERIC, SUMIFS daría 0.
        try (Workbook wb = newWorkbookWithResultado()) {
            Sheet sheet = createTargetSheet(wb, "tresp1@x");
            ResponsablePivotBuilder builder = new ResponsablePivotBuilder(minimalConfig());

            builder.writeTable(wb, sheet, 2,
                    jiraInputs(Arrays.asList("P-001"),
                               Arrays.asList("99641")));

            Cell matrHeader = sheet.getRow(3).getCell(1);
            assertThat(matrHeader.getCellType()).isEqualTo(CellType.STRING);
            assertThat(matrHeader.getStringCellValue()).isEqualTo("99641");
        }
    }
}

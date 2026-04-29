package com.excelmerger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests unitarios de {@link EmptyRowFilter} (v2.7.1).
 *
 * <p>Cubren tres aspectos:</p>
 * <ul>
 *   <li>Criterio "vale 0" para celdas no-formula (BLANK, NUMERIC=0,
 *       NUMERIC distinto, STRING vacio, STRING "-", STRING con texto).</li>
 *   <li>Traduccion de referencias de fila locales en formulas
 *       ({@code translateLocalRowRefs}).</li>
 *   <li>Filtrado integral: dado un sheet con cabeceras y filas
 *       construidas a mano (con valores literales numericos en las 5
 *       columnas), {@code apply} elimina las filas correctas y conserva
 *       las demas.</li>
 * </ul>
 *
 * <p>No usa fixtures Excel reales: todos los workbooks se construyen
 * en memoria (XSSF). Los tests del integration end-to-end estan en
 * {@link ExcelMergerIntegrationTest}.</p>
 */
class EmptyRowFilterTest {

    // ==================================================================
    //  Criterio "vale 0" - cellIsZeroForTest
    // ==================================================================

    @Test
    void blankCellSeConsideraCero() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("S");
            Row r = s.createRow(0);
            Cell blank = r.createCell(0);
            // setBlank explicitly via setCellValue("")? POI: createCell sin
            // valor queda como BLANK por defecto.
            assertThat(EmptyRowFilter.cellIsZeroForTest(blank)).isTrue();
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void cellNullSeConsideraCero() {
        assertThat(EmptyRowFilter.cellIsZeroForTest(null)).isTrue();
    }

    @Test
    void numericCeroSeConsideraCero() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("S");
            Row r = s.createRow(0);
            Cell c = r.createCell(0);
            c.setCellValue(0.0);
            assertThat(EmptyRowFilter.cellIsZeroForTest(c)).isTrue();
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void numericNoCeroNoEsCero() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("S");
            Row r = s.createRow(0);
            Cell c = r.createCell(0);
            c.setCellValue(0.001);
            assertThat(EmptyRowFilter.cellIsZeroForTest(c)).isFalse();

            Cell c2 = r.createCell(1);
            c2.setCellValue(-1.0);
            assertThat(EmptyRowFilter.cellIsZeroForTest(c2)).isFalse();
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void stringVaciaYGuionSeConsideranCero() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("S");
            Row r = s.createRow(0);

            Cell empty = r.createCell(0);
            empty.setCellValue("");
            assertThat(EmptyRowFilter.cellIsZeroForTest(empty)).isTrue();

            Cell spaces = r.createCell(1);
            spaces.setCellValue("   ");
            assertThat(EmptyRowFilter.cellIsZeroForTest(spaces)).isTrue();

            Cell guion = r.createCell(2);
            guion.setCellValue("-");
            assertThat(EmptyRowFilter.cellIsZeroForTest(guion)).isTrue();

            Cell guionConSpaces = r.createCell(3);
            guionConSpaces.setCellValue("  -  ");
            assertThat(EmptyRowFilter.cellIsZeroForTest(guionConSpaces)).isTrue();
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    /**
     * Regresion v2.7.1.1: el ERP real serializa la columna
     * {@code UltimaPrevision_Horas_Mes} como STRING formateada (ej.
     * {@code "0.00"}, {@code "20.00"}). El criterio "vale 0" debe
     * tratar como cero cualquier STRING que parsee numericamente a 0,
     * incluyendo coma decimal del locale es_ES.
     */
    @Test
    void stringNumericamenteCeroSeConsideraCero() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("S");
            Row r = s.createRow(0);

            String[] cerosVariados = {"0", "0.0", "0.00", "0,00", "0,0", " 0.00 ", "  0  "};
            for (int i = 0; i < cerosVariados.length; i++) {
                Cell c = r.createCell(i);
                c.setCellValue(cerosVariados[i]);
                assertThat(EmptyRowFilter.cellIsZeroForTest(c))
                        .as("STRING %s debe ser cero", cerosVariados[i])
                        .isTrue();
            }
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    /**
     * Regresion v2.7.1.1: una STRING que parsea numericamente a un valor
     * distinto de 0 NO debe contar como cero (no filtrar la fila).
     */
    @Test
    void stringNumericaNoCeroNoEsCero() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("S");
            Row r = s.createRow(0);

            String[] noCeros = {"0.5", "5.00", "20.00", "0,5", "-5", "1e-9", "0.0001"};
            for (int i = 0; i < noCeros.length; i++) {
                Cell c = r.createCell(i);
                c.setCellValue(noCeros[i]);
                assertThat(EmptyRowFilter.cellIsZeroForTest(c))
                        .as("STRING %s NO debe ser cero", noCeros[i])
                        .isFalse();
            }
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void stringConTextoNoEsCero() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("S");
            Row r = s.createRow(0);
            Cell c = r.createCell(0);
            c.setCellValue("Sin Matricula");
            assertThat(EmptyRowFilter.cellIsZeroForTest(c)).isFalse();
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    // ==================================================================
    //  Traduccion de referencias de fila locales
    // ==================================================================

    @Test
    void translateLocalRowRefsFormulaSimple() {
        // J5*1.2 con shift 5 -> 4 debe convertirse en J4*1.2
        String result = EmptyRowFilter.translateLocalRowRefs("J5*1.2", 5, 4);
        assertThat(result).isEqualTo("J4*1.2");
    }

    @Test
    void translateLocalRowRefsMultiplesReferencias() {
        // Caso real de FORMULA_PLUS_SUMIFS: la formula generada combina
        // baseFormula y SUMIFS. La parte local es la baseFormula (ej. K5).
        String result = EmptyRowFilter.translateLocalRowRefs("K5+SUMIFS(Deuda!$D:$D,Deuda!$A:$A,A5)", 5, 3);
        assertThat(result).isEqualTo("K3+SUMIFS(Deuda!$D:$D,Deuda!$A:$A,A3)");
    }

    @Test
    void translateLocalRowRefsSinHojaActualNoTocaPrefijos() {
        // Con la sobrecarga sin currentSheetName, ninguna referencia con
        // prefijo de hoja se traduce (sea o no la propia hoja). Es el
        // comportamiento usado en los tests unitarios donde no se conoce
        // el contexto de hoja.
        String result = EmptyRowFilter.translateLocalRowRefs("Resultado!J5+1", 5, 4);
        assertThat(result).isEqualTo("Resultado!J5+1");
    }

    @Test
    void translateLocalRowRefsRespetaReferenciasAOtraHoja() {
        // Cuando currentSheetName="Resultado", una referencia a "Cierre!" o
        // "Deuda!" NO debe tocarse: solo se traducen las que apuntan a la
        // propia hoja.
        String r1 = EmptyRowFilter.translateLocalRowRefs("Cierre!$A$5+1", 5, 4, "Resultado");
        assertThat(r1).isEqualTo("Cierre!$A$5+1");

        String r2 = EmptyRowFilter.translateLocalRowRefs("SUMIFS(Deuda!$D:$D,Deuda!$A:$A,A5)", 5, 3, "Resultado");
        // 'A5' (sin prefijo de hoja, local) SI se traduce. 'Deuda!' no.
        assertThat(r2).isEqualTo("SUMIFS(Deuda!$D:$D,Deuda!$A:$A,A3)");
    }

    /**
     * Regresion v2.7.1.1: el bug critico que rompia el output era que las
     * formulas FORMULA_PLUS_SUMIFS de PDCL+Deuda se construyen con
     * "Resultado!$A$<row>" y similares (referencias absolutas a la PROPIA
     * hoja). Tras el shift, esas referencias deben re-numerarse.
     */
    @Test
    void translateLocalRowRefsTraduceAutoReferenciasAbsolutas() {
        // Forma exacta de la formula generada por FormulaPlusSumIfsColumnStrategy
        String original = "L40+IFERROR(SUMIFS(Deuda!$D:$D,Deuda!$A:$A,Resultado!$A$66,"
                + "Deuda!$B:$B,Resultado!$F$66,Deuda!$C:$C,Resultado!$G$66),0)";
        // Si la fila origen 66 se mueve a la 40, las tres Resultado!$X$66
        // deben pasar a Resultado!$X$40. La parte L40 ya esta en su sitio
        // (la fila destino). Pero supongamos shift 66 -> 40: la formula
        // ORIGINAL tendria L66 (no L40); aqui simulamos la formula tal y como
        // estaba en la fila 66 antes del shift.
        String preShift = "L66+IFERROR(SUMIFS(Deuda!$D:$D,Deuda!$A:$A,Resultado!$A$66,"
                + "Deuda!$B:$B,Resultado!$F$66,Deuda!$C:$C,Resultado!$G$66),0)";
        String result = EmptyRowFilter.translateLocalRowRefs(preShift, 66, 40, "Resultado");
        String expected = "L40+IFERROR(SUMIFS(Deuda!$D:$D,Deuda!$A:$A,Resultado!$A$40,"
                + "Deuda!$B:$B,Resultado!$F$40,Deuda!$C:$C,Resultado!$G$40),0)";
        assertThat(result).isEqualTo(expected);
    }

    @Test
    void translateLocalRowRefsTraduceAutoReferenciasMixtas() {
        // Resultado!J5 (relativa), Resultado!$J$5 (absoluta), Resultado!J$5 (mixta)
        String r1 = EmptyRowFilter.translateLocalRowRefs("Resultado!J5+1", 5, 4, "Resultado");
        assertThat(r1).isEqualTo("Resultado!J4+1");

        String r2 = EmptyRowFilter.translateLocalRowRefs("Resultado!$J$5+1", 5, 4, "Resultado");
        assertThat(r2).isEqualTo("Resultado!$J$4+1");

        String r3 = EmptyRowFilter.translateLocalRowRefs("Resultado!J$5+1", 5, 4, "Resultado");
        assertThat(r3).isEqualTo("Resultado!J$4+1");
    }

    @Test
    void translateLocalRowRefsNoMatcheaPrefijosParecidos() {
        // "MiResultado!" no es la propia hoja "Resultado": no debe traducirse
        // (el char anterior 'i' es alfanumerico y descalifica).
        String result = EmptyRowFilter.translateLocalRowRefs(
                "MiResultado!$A$5+1", 5, 4, "Resultado");
        assertThat(result).isEqualTo("MiResultado!$A$5+1");
    }

    @Test
    void translateLocalRowRefsRespetaReferenciasAbsolutas() {
        // J$5 y $J$5 sin prefijo de hoja son referencias locales con fila
        // absoluta: NO se traducen (deliberadamente conservador).
        String r1 = EmptyRowFilter.translateLocalRowRefs("J$5*1.2", 5, 4);
        assertThat(r1).isEqualTo("J$5*1.2");

        String r2 = EmptyRowFilter.translateLocalRowRefs("$J$5*1.2", 5, 4);
        assertThat(r2).isEqualTo("$J$5*1.2");
    }

    @Test
    void translateLocalRowRefsNoMatcheaSubstrings() {
        // El texto "AB52" no debe traducirse aunque contenga "B5" pegado.
        // Probamos con src=5, dst=99: si AB52 se traduce mal, saldria AB992 o similar.
        String result = EmptyRowFilter.translateLocalRowRefs("AB52*2", 5, 99);
        assertThat(result).isEqualTo("AB52*2");
    }

    @Test
    void translateLocalRowRefsNoMatcheaCuandoSrcDifiere() {
        // Si la formula referencia la fila 5 pero estamos shifteando desde
        // la 8, no debe tocar nada.
        String result = EmptyRowFilter.translateLocalRowRefs("J5*1.2", 8, 4);
        assertThat(result).isEqualTo("J5*1.2");
    }

    @Test
    void translateLocalRowRefsIdempotenteSiSrcIgualDst() {
        String result = EmptyRowFilter.translateLocalRowRefs("J5*1.2", 5, 5);
        assertThat(result).isEqualTo("J5*1.2");
    }

    @Test
    void translateLocalRowRefsManejaNumerosGrandes() {
        // src=1234 -> dst=1230. La formula tiene J1234 y debe quedar J1230.
        // Importante: no debe matchear J1 (un substring de J1234).
        String result = EmptyRowFilter.translateLocalRowRefs("SUM(J1234,K1234)", 1234, 1230);
        assertThat(result).isEqualTo("SUM(J1230,K1230)");
    }

    // ==================================================================
    //  apply: filtrado integral
    // ==================================================================

    /**
     * Construye una hoja Resultado minimal con las 5 columnas + Petición
     * + Matrícula. Las 5 columnas se rellenan con valores LITERALES
     * numericos (no formulas) para que FormulaEvaluator no tenga que
     * resolver nada en este test (verificamos solo la mecanica de
     * filtrado y compactacion).
     *
     * <p>Estructura:</p>
     * <pre>
     * Petición | Matrícula | Jira | Facturar | PDCL | PDCL + Deuda | Horas_Mes
     * P-001    | M-1001    |  5.0 |  6.0     |  6.0 |  7.0         |  8.0
     * P-002    | M-1002    |  0.0 |  0.0     |  0.0 |  0.0         |  0.0    &lt;-- se filtra
     * P-003    | M-1003    |  3.0 |  3.6     |  3.6 |  3.6         |  4.0
     * P-004    | M-1004    |  0.0 |  0.0     |  0.0 |  0.0         |  0.0    &lt;-- se filtra
     * P-005    | M-1005    |  1.0 |  1.2     |  1.2 |  1.2         |  0.0
     * </pre>
     */
    private static Workbook buildSheetWithLiteralValues() {
        Workbook wb = new XSSFWorkbook();
        Sheet s = wb.createSheet("Resultado");
        // Cabeceras (orden y nombres exactos que busca EmptyRowFilter)
        Row h = s.createRow(0);
        h.createCell(0).setCellValue("Petición");
        h.createCell(1).setCellValue("Matrícula");
        h.createCell(2).setCellValue("Jira");
        h.createCell(3).setCellValue("Facturar");
        h.createCell(4).setCellValue("PDCL");
        h.createCell(5).setCellValue("PDCL + Deuda");
        h.createCell(6).setCellValue("Horas_Mes");

        addDataRow(s, 1, "P-001", "M-1001", 5.0, 6.0, 6.0, 7.0, 8.0);
        addDataRow(s, 2, "P-002", "M-1002", 0.0, 0.0, 0.0, 0.0, 0.0);
        addDataRow(s, 3, "P-003", "M-1003", 3.0, 3.6, 3.6, 3.6, 4.0);
        addDataRow(s, 4, "P-004", "M-1004", 0.0, 0.0, 0.0, 0.0, 0.0);
        addDataRow(s, 5, "P-005", "M-1005", 1.0, 1.2, 1.2, 1.2, 0.0);

        return wb;
    }

    private static void addDataRow(Sheet s, int rowIdx, String pet, String mat,
                                   double jira, double facturar, double pdcl,
                                   double pdclMasDeuda, double horasMes) {
        Row r = s.createRow(rowIdx);
        r.createCell(0).setCellValue(pet);
        r.createCell(1).setCellValue(mat);
        r.createCell(2).setCellValue(jira);
        r.createCell(3).setCellValue(facturar);
        r.createCell(4).setCellValue(pdcl);
        r.createCell(5).setCellValue(pdclMasDeuda);
        r.createCell(6).setCellValue(horasMes);
    }

    @Test
    void applyFiltraFilasConTodasLasCincoColumnasACero() {
        try (Workbook wb = buildSheetWithLiteralValues()) {
            Sheet s = wb.getSheet("Resultado");
            RunReport report = new RunReport();

            int removed = EmptyRowFilter.apply(wb, s, 5, report);

            assertThat(removed).isEqualTo(2);
            // Tras filtrar, deben quedar 3 filas de datos (P-001, P-003, P-005)
            // Numero total de filas (incluyendo cabecera) = 4. lastRowNum = 3.
            assertThat(s.getLastRowNum()).isEqualTo(3);
            // Verificar contenido en orden
            assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("P-001");
            assertThat(s.getRow(2).getCell(0).getStringCellValue()).isEqualTo("P-003");
            assertThat(s.getRow(3).getCell(0).getStringCellValue()).isEqualTo("P-005");
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void applySinFilasACeroNoElimina() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("Resultado");
            Row h = s.createRow(0);
            h.createCell(0).setCellValue("Petición");
            h.createCell(1).setCellValue("Matrícula");
            h.createCell(2).setCellValue("Jira");
            h.createCell(3).setCellValue("Facturar");
            h.createCell(4).setCellValue("PDCL");
            h.createCell(5).setCellValue("PDCL + Deuda");
            h.createCell(6).setCellValue("Horas_Mes");
            addDataRow(s, 1, "P-001", "M-1001", 5.0, 6.0, 6.0, 7.0, 8.0);
            addDataRow(s, 2, "P-002", "M-1002", 0.0, 0.0, 0.0, 0.0, 1.0); // Horas_Mes salva
            addDataRow(s, 3, "P-003", "M-1003", 0.0, 0.0, 0.0, 0.0, 0.0); // todas a 0

            RunReport report = new RunReport();
            int removed = EmptyRowFilter.apply(wb, s, 3, report);
            assertThat(removed).isEqualTo(1);
            // P-002 sobrevive porque Horas_Mes=1.0
            assertThat(s.getRow(2).getCell(0).getStringCellValue()).isEqualTo("P-002");
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void applyConColumnasFaltantesEmiteWarningYNoFiltra() {
        // Si falta alguna de las 5 columnas, el filtro debe abortar (no filtrar)
        // y registrar un warning CONFIG en el report.
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("Resultado");
            Row h = s.createRow(0);
            h.createCell(0).setCellValue("Petición");
            // OJO: aqui falta "PDCL + Deuda", "Horas_Mes" y "Facturar"
            h.createCell(1).setCellValue("Jira");
            h.createCell(2).setCellValue("PDCL");
            Row r1 = s.createRow(1);
            r1.createCell(0).setCellValue("P-001");
            r1.createCell(1).setCellValue(0.0);
            r1.createCell(2).setCellValue(0.0);

            RunReport report = new RunReport();
            int removed = EmptyRowFilter.apply(wb, s, 1, report);

            assertThat(removed).isZero();
            // El warning debe mencionar las columnas faltantes
            assertThat(report.warnings()).anyMatch(w ->
                    "CONFIG".equals(w.category)
                    && w.message.contains("Facturar")
                    && w.message.contains("PDCL + Deuda")
                    && w.message.contains("Horas_Mes"));
            // La fila P-001 sigue ahi
            assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("P-001");
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void applyConHojaSinDatosNoHaceNada() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("Resultado");
            Row h = s.createRow(0);
            h.createCell(0).setCellValue("Petición");
            h.createCell(1).setCellValue("Matrícula");
            h.createCell(2).setCellValue("Jira");
            h.createCell(3).setCellValue("Facturar");
            h.createCell(4).setCellValue("PDCL");
            h.createCell(5).setCellValue("PDCL + Deuda");
            h.createCell(6).setCellValue("Horas_Mes");

            RunReport report = new RunReport();
            int removed = EmptyRowFilter.apply(wb, s, 0, report);
            assertThat(removed).isZero();
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void applyEvaluaFormulasYFiltraSegunResultado() {
        // Verifica que con celdas FORMULA, FormulaEvaluator se invoca
        // y las filas se filtran segun el VALOR evaluado.
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("Resultado");
            Row h = s.createRow(0);
            h.createCell(0).setCellValue("Petición");
            h.createCell(1).setCellValue("Matrícula");
            h.createCell(2).setCellValue("Jira");
            h.createCell(3).setCellValue("Facturar");
            h.createCell(4).setCellValue("PDCL");
            h.createCell(5).setCellValue("PDCL + Deuda");
            h.createCell(6).setCellValue("Horas_Mes");

            // Fila 1: Jira=0 (numerico), pero Facturar = 0+1 (formula evaluada -> 1)
            // -> NO se filtra
            Row r1 = s.createRow(1);
            r1.createCell(0).setCellValue("P-001");
            r1.createCell(1).setCellValue("M-1001");
            r1.createCell(2).setCellValue(0.0);
            r1.createCell(3).setCellFormula("0+1");
            r1.createCell(4).setCellValue(0.0);
            r1.createCell(5).setCellValue(0.0);
            r1.createCell(6).setCellValue(0.0);

            // Fila 2: todas 0 incluidas las formulas -> se filtra
            Row r2 = s.createRow(2);
            r2.createCell(0).setCellValue("P-002");
            r2.createCell(1).setCellValue("M-1002");
            r2.createCell(2).setCellValue(0.0);
            r2.createCell(3).setCellFormula("0+0");
            r2.createCell(4).setCellFormula("C3*1.2"); // referencia local: row 3 (excel) = Jira fila 2
            r2.createCell(5).setCellValue(0.0);
            r2.createCell(6).setCellValue(0.0);

            RunReport report = new RunReport();
            int removed = EmptyRowFilter.apply(wb, s, 2, report);
            assertThat(removed).isEqualTo(1);
            assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("P-001");
            assertThat(s.getLastRowNum()).isEqualTo(1);
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void applyTraduceFormulasLocalesAlCompactar() {
        // Si una fila SUPERVIVIENTE estaba en Excel-row 4 y al filtrar
        // pasa a Excel-row 2, sus formulas con referencia local J4 deben
        // pasar a J2 para que sigan referenciando la misma fila relativa.
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("Resultado");
            Row h = s.createRow(0);
            h.createCell(0).setCellValue("Petición");
            h.createCell(1).setCellValue("Matrícula");
            h.createCell(2).setCellValue("Jira");
            h.createCell(3).setCellValue("Facturar");
            h.createCell(4).setCellValue("PDCL");
            h.createCell(5).setCellValue("PDCL + Deuda");
            h.createCell(6).setCellValue("Horas_Mes");

            // Fila 1 (excel-row 2): todo 0 -> se filtra
            Row r1 = s.createRow(1);
            r1.createCell(0).setCellValue("P-001");
            r1.createCell(1).setCellValue("M-1001");
            r1.createCell(2).setCellValue(0.0);
            r1.createCell(3).setCellValue(0.0);
            r1.createCell(4).setCellValue(0.0);
            r1.createCell(5).setCellValue(0.0);
            r1.createCell(6).setCellValue(0.0);

            // Fila 2 (excel-row 3): todo 0 -> se filtra
            Row r2 = s.createRow(2);
            r2.createCell(0).setCellValue("P-002");
            r2.createCell(1).setCellValue("M-1002");
            r2.createCell(2).setCellValue(0.0);
            r2.createCell(3).setCellValue(0.0);
            r2.createCell(4).setCellValue(0.0);
            r2.createCell(5).setCellValue(0.0);
            r2.createCell(6).setCellValue(0.0);

            // Fila 3 (excel-row 4): formula con ref local C4 (=Jira fila 3),
            // valor de Jira=5 -> Facturar=C4*1.2 -> 6.0. Tras filtrado, esta
            // fila debe pasar a excel-row 2 y la formula a "C2*1.2".
            Row r3 = s.createRow(3);
            r3.createCell(0).setCellValue("P-003");
            r3.createCell(1).setCellValue("M-1003");
            r3.createCell(2).setCellValue(5.0);
            r3.createCell(3).setCellFormula("C4*1.2");
            r3.createCell(4).setCellValue(6.0);
            r3.createCell(5).setCellValue(6.0);
            r3.createCell(6).setCellValue(0.0);

            RunReport report = new RunReport();
            int removed = EmptyRowFilter.apply(wb, s, 3, report);
            assertThat(removed).isEqualTo(2);

            // P-003 ahora esta en excel-row 2 (rowIdx0=1)
            Row survivor = s.getRow(1);
            assertThat(survivor.getCell(0).getStringCellValue()).isEqualTo("P-003");
            // La formula traducida debe ser C2*1.2
            assertThat(survivor.getCell(3).getCellFormula()).isEqualTo("C2*1.2");

            // Y la formula evaluada debe seguir dando 6.0 (5.0 * 1.2)
            org.apache.poi.ss.usermodel.FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            org.apache.poi.ss.usermodel.CellValue cv = ev.evaluate(survivor.getCell(3));
            assertThat(cv.getNumberValue()).isEqualTo(6.0);
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void applyConTodasLasFilasACeroEliminaTodas() {
        try (Workbook wb = buildSheetWithLiteralValues()) {
            Sheet s = wb.getSheet("Resultado");
            // Sobrescribir todas las filas de datos con ceros
            for (int r = 1; r <= 5; r++) {
                Row row = s.getRow(r);
                row.getCell(2).setCellValue(0.0);
                row.getCell(3).setCellValue(0.0);
                row.getCell(4).setCellValue(0.0);
                row.getCell(5).setCellValue(0.0);
                row.getCell(6).setCellValue(0.0);
            }
            RunReport report = new RunReport();
            int removed = EmptyRowFilter.apply(wb, s, 5, report);
            assertThat(removed).isEqualTo(5);
            // Solo queda la cabecera (fila 0)
            assertThat(s.getLastRowNum()).isZero();
            assertThat(s.getRow(0).getCell(0).getStringCellValue()).isEqualTo("Petición");
        } catch (java.io.IOException e) {
            throw new AssertionError(e);
        }
    }
}

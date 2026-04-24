package com.excelmerger.sheet.column;

import com.excelmerger.RunReport;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests unitarios de {@link FormulaPlusSumIfsColumnStrategy}.
 *
 * <p>Cubre los cuatro escenarios del ciclo de vida:</p>
 * <ul>
 *   <li>preValidate con hoja remota ausente: no deshabilita ni emite
 *       warnings (degradado silencioso por diseño).</li>
 *   <li>preValidate con hoja remota presente y cabeceras completas: no
 *       warnings.</li>
 *   <li>preValidate con hoja presente pero cabecera faltante: warning
 *       CABECERA y deshabilita la columna.</li>
 *   <li>doWriteCell con y sin hoja remota: escribe {@code base} sola o
 *       {@code base+IFERROR(SUMIFS(...),0)} respectivamente.</li>
 * </ul>
 *
 * <p>v2.2.0: nueva estrategia para la columna "PDCL + Deuda".</p>
 */
class FormulaPlusSumIfsColumnStrategyTest {

    private static List<String[]> deudaMatches() {
        return Arrays.asList(
                new String[]{"Peticion", "Petición"},
                new String[]{"Matricula", "Matrícula"},
                new String[]{"Funcion", "Funcion"}
        );
    }

    private static FormulaPlusSumIfsColumnStrategy buildStrategy() {
        return new FormulaPlusSumIfsColumnStrategy(
                "PDCL + Deuda", false, null, "PDCL",
                "{col:PDCL}", "Deuda", "Horas", deudaMatches());
    }

    /**
     * Crea una hoja "Resultado" minima: solo cabeceras en fila 0 y una
     * fila de datos en fila 1 con Peticion/Matricula/Res.Tecnico.
     */
    private static Sheet buildResultadoSheet(Workbook wb) {
        Sheet s = wb.createSheet("Resultado");
        Row header = s.createRow(0);
        header.createCell(0).setCellValue("Petición");
        header.createCell(1).setCellValue("Matrícula");
        header.createCell(2).setCellValue("Funcion");
        header.createCell(3).setCellValue("PDCL");
        header.createCell(4).setCellValue("PDCL + Deuda");
        Row r1 = s.createRow(1);
        r1.createCell(0).setCellValue("P-001");
        r1.createCell(1).setCellValue("M-100");
        r1.createCell(2).setCellValue("Dev");
        r1.createCell(3).setCellValue(10.0);
        return s;
    }

    /** Cabeceras completas de la hoja Deuda: Peticion, Matricula, Funcion, Horas. */
    private static void buildDeudaSheet(Workbook wb) {
        Sheet d = wb.createSheet("Deuda");
        Row h = d.createRow(0);
        h.createCell(0).setCellValue("Peticion");
        h.createCell(1).setCellValue("Matricula");
        h.createCell(2).setCellValue("Funcion");
        h.createCell(3).setCellValue("Horas");
    }

    private static Map<String, Integer> mesColIndex() {
        Map<String, Integer> m = new LinkedHashMap<>();
        m.put("Petición", 0);
        m.put("Matrícula", 1);
        m.put("Funcion", 2);
        m.put("PDCL", 3);
        m.put("PDCL + Deuda", 4);
        return m;
    }

    // =================================================================
    //  preValidate
    // =================================================================

    @Test
    void preValidateSinHojaRemotaNoDeshabilitaNiEmiteWarnings() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet source = buildResultadoSheet(wb);
            FormulaPlusSumIfsColumnStrategy strat = buildStrategy();
            RunReport report = new RunReport();

            strat.preValidate(source, 0, wb, mesColIndex(), report);

            assertThat(strat.isDisabled()).isFalse();
            assertThat(report.warnings()).isEmpty();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void preValidateConHojaRemotaYCabecerasOkNoWarnings() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet source = buildResultadoSheet(wb);
            buildDeudaSheet(wb);
            FormulaPlusSumIfsColumnStrategy strat = buildStrategy();
            RunReport report = new RunReport();

            strat.preValidate(source, 0, wb, mesColIndex(), report);

            assertThat(strat.isDisabled()).isFalse();
            assertThat(report.warnings()).isEmpty();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void preValidateConHojaRemotaYCabeceraRemotaFaltanteEmiteWarningYDeshabilita() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet source = buildResultadoSheet(wb);
            // Hoja Deuda sin la columna "Horas"
            Sheet d = wb.createSheet("Deuda");
            Row h = d.createRow(0);
            h.createCell(0).setCellValue("Peticion");
            h.createCell(1).setCellValue("Matricula");
            h.createCell(2).setCellValue("Funcion");
            // <-- sin Horas

            FormulaPlusSumIfsColumnStrategy strat = buildStrategy();
            RunReport report = new RunReport();

            strat.preValidate(source, 0, wb, mesColIndex(), report);

            assertThat(strat.isDisabled()).isTrue();
            assertThat(report.warnings()).anyMatch(w ->
                    "CABECERA".equals(w.category)
                            && w.message.contains("Horas")
                            && w.message.contains("Deuda"));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void preValidateConPlaceholderInvalidoEnBaseFormulaEmiteWarningYDeshabilita() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet source = buildResultadoSheet(wb);
            FormulaPlusSumIfsColumnStrategy strat = new FormulaPlusSumIfsColumnStrategy(
                    "PDCL + Deuda", false, null, null,
                    "{col:NoExiste}", "Deuda", "Horas", deudaMatches());
            RunReport report = new RunReport();

            strat.preValidate(source, 0, wb, mesColIndex(), report);

            assertThat(strat.isDisabled()).isTrue();
            assertThat(report.warnings()).anyMatch(w ->
                    "FORMULA".equals(w.category)
                            && w.message.contains("NoExiste"));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    // =================================================================
    //  doWriteCell via writeCell (plantilla final)
    // =================================================================

    @Test
    void writeCellSinHojaRemotaEscribeSoloBaseFormula() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet source = buildResultadoSheet(wb);
            FormulaPlusSumIfsColumnStrategy strat = buildStrategy();

            Row targetRow = source.getRow(1);
            Cell target = targetRow.createCell(4);
            strat.writeCell(target, targetRow, source, 0, wb, 2, mesColIndex(), 2);

            assertThat(target.getCellFormula()).isEqualTo("D2");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void writeCellConHojaRemotaEscribeBaseMasIferrorSumifs() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet source = buildResultadoSheet(wb);
            buildDeudaSheet(wb);
            FormulaPlusSumIfsColumnStrategy strat = buildStrategy();

            Row targetRow = source.getRow(1);
            Cell target = targetRow.createCell(4);
            strat.writeCell(target, targetRow, source, 0, wb, 2, mesColIndex(), 2);

            String formula = target.getCellFormula();
            // Base resuelta al principio: D2
            assertThat(formula).startsWith("D2+IFERROR(SUMIFS(");
            // Referencias a la hoja Deuda con rangos abiertos (consistente
            // con mes.col.10 actual), apuntando a la fila Resultado 2.
            assertThat(formula).contains("Deuda!$D:$D");
            assertThat(formula).contains("Deuda!$A:$A,Resultado!$A$2");
            assertThat(formula).contains("Deuda!$B:$B,Resultado!$B$2");
            assertThat(formula).contains("Deuda!$C:$C,Resultado!$C$2");
            assertThat(formula).endsWith(",0)");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void writeCellDeshabilitadaEscribeBlank() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet source = buildResultadoSheet(wb);
            FormulaPlusSumIfsColumnStrategy strat = new FormulaPlusSumIfsColumnStrategy(
                    "PDCL + Deuda", false, null, null,
                    "{col:NoExiste}", "Deuda", "Horas", deudaMatches());
            RunReport report = new RunReport();
            strat.preValidate(source, 0, wb, mesColIndex(), report);
            assertThat(strat.isDisabled()).isTrue();

            Row targetRow = source.getRow(1);
            Cell target = targetRow.createCell(4);
            strat.writeCell(target, targetRow, source, 0, wb, 2, mesColIndex(), 2);

            assertThat(target.getCellType())
                    .isEqualTo(org.apache.poi.ss.usermodel.CellType.BLANK);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}

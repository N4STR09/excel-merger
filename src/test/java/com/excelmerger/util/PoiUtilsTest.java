package com.excelmerger.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.Date;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests unitarios de los helpers "text-forcing" de {@link PoiUtils}
 * anadidos en v1.6.2.
 *
 * <p>El resto de {@link PoiUtils} se prueba indirectamente a traves de
 * los tests de integracion ({@code ExcelMergerIntegrationTest}, los
 * builders...). Esta clase se centra en el contrato de
 * {@link PoiUtils#copyCellValueAsText} para cubrir el fix del mismatch
 * numerico/textual en SUMIFS.</p>
 */
class PoiUtilsTest {

    @Test
    void copyCellValueAsTextConvierteNumericEnteroALaRepresentacionSinDecimales() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);

            Cell src = row.createCell(0);
            src.setCellValue(55751d);
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsText(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.STRING);
            assertThat(dst.getStringCellValue()).isEqualTo("55751");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextPreservaDecimalesSiLosTiene() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);

            Cell src = row.createCell(0);
            src.setCellValue(3.14d);
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsText(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.STRING);
            assertThat(dst.getStringCellValue()).isEqualTo("3.14");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextPreservaStringTalCual() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);

            Cell src = row.createCell(0);
            src.setCellValue("55751");
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsText(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.STRING);
            assertThat(dst.getStringCellValue()).isEqualTo("55751");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextPreservaGuionComoStringSinTocar() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);

            Cell src = row.createCell(0);
            src.setCellValue("-");
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsText(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.STRING);
            assertThat(dst.getStringCellValue()).isEqualTo("-");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextConvierteBooleanAtexto() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);

            Cell src = row.createCell(0);
            src.setCellValue(true);
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsText(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.STRING);
            assertThat(dst.getStringCellValue()).isEqualTo("true");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextCeldaBlankQuedaBlank() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);

            Cell src = row.createCell(0, CellType.BLANK);
            Cell dst = row.createCell(1);
            // Escribimos algo previo en dst para asegurar que el metodo la
            // "limpia" correctamente.
            dst.setCellValue("no deberia quedar");

            PoiUtils.copyCellValueAsText(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.BLANK);
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextFormulaSeCopiaSinEvaluar() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);

            Cell src = row.createCell(0);
            src.setCellFormula("1+2");
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsText(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(dst.getCellFormula()).isEqualTo("1+2");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextErrorSeCopiaComoError() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);

            Cell src = row.createCell(0);
            src.setCellErrorValue(FormulaError.NA.getCode());
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsText(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.ERROR);
            assertThat(dst.getErrorCellValue()).isEqualTo(FormulaError.NA.getCode());
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextNoFuerzaAFechaFormateada() {
        // Contrato explicito: las fechas no se serializan como epoch Excel.
        // Se copian como fecha (delegando en el mismo camino que
        // copyCellValue). Ver javadoc de PoiUtils.copyCellValueAsText.
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);

            CreationHelper helper = wb.getCreationHelper();
            DataFormat fmt = helper.createDataFormat();
            CellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(fmt.getFormat("yyyy-mm-dd"));

            Cell src = row.createCell(0);
            Date date = new Date(1713196800000L); // 2024-04-15
            src.setCellValue(date);
            src.setCellStyle(dateStyle);
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsText(src, dst);

            // Destino es NUMERIC (POI representa fechas como numeric con
            // formato), no STRING.
            assertThat(dst.getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(dst.getDateCellValue()).isEqualTo(date);
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    // ==================================================================
    //  v1.8.1 — copyCellValueAsTextTrimmed
    // ==================================================================

    @Test
    void copyCellValueAsTextTrimmedQuitaEspaciosDeSTRINGConPadding() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);
            Cell src = row.createCell(0);
            src.setCellValue("MG002   ");   // padding a la derecha, caso real de export ERP
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsTextTrimmed(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.STRING);
            assertThat(dst.getStringCellValue()).isEqualTo("MG002");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextTrimmedQuitaEspaciosAlPrincipioYFinal() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);
            Cell src = row.createCell(0);
            src.setCellValue("  RESP01  ");
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsTextTrimmed(src, dst);

            assertThat(dst.getStringCellValue()).isEqualTo("RESP01");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextTrimmedNoCambiaSTRINGSinEspacios() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);
            Cell src = row.createCell(0);
            src.setCellValue("RESP01");
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsTextTrimmed(src, dst);

            assertThat(dst.getStringCellValue()).isEqualTo("RESP01");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextTrimmedSobreNumericEnteroSerializaSinDecimales() {
        // Un NUMERIC entero convertido a string nunca lleva espacios, asi
        // que el trim es no-op. Verificamos que la semantica coincide con
        // copyCellValueAsText para esta rama.
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);
            Cell src = row.createCell(0);
            src.setCellValue(55751d);
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsTextTrimmed(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.STRING);
            assertThat(dst.getStringCellValue()).isEqualTo("55751");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextTrimmedSobreFechaNoLaFuerzaATexto() {
        // Misma regla que copyCellValueAsText: una fecha se copia como
        // fecha, no como texto (evita convertir 2024-04-15 a "45397").
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);
            CreationHelper ch = wb.getCreationHelper();
            DataFormat fmt = ch.createDataFormat();
            CellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(fmt.getFormat("yyyy-mm-dd"));

            Cell src = row.createCell(0);
            Date date = new Date(1713196800000L);
            src.setCellValue(date);
            src.setCellStyle(dateStyle);
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsTextTrimmed(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(dst.getDateCellValue()).isEqualTo(date);
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextTrimmedSobreBooleanSerializaComoTrueFalse() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);
            Cell src = row.createCell(0);
            src.setCellValue(true);
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsTextTrimmed(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.STRING);
            assertThat(dst.getStringCellValue()).isEqualTo("true");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextTrimmedPreservaFormulaSinEvaluar() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);
            Cell src = row.createCell(0);
            src.setCellFormula("1+1");
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsTextTrimmed(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(dst.getCellFormula()).isEqualTo("1+1");
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextTrimmedSobreBlankDejaCeldaBlank() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);
            Cell src = row.createCell(0);
            src.setBlank();
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsTextTrimmed(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.BLANK);
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void copyCellValueAsTextTrimmedSobreErrorCopiaElError() {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("t");
            Row row = sheet.createRow(0);
            Cell src = row.createCell(0);
            src.setCellErrorValue(FormulaError.DIV0.getCode());
            Cell dst = row.createCell(1);

            PoiUtils.copyCellValueAsTextTrimmed(src, dst);

            assertThat(dst.getCellType()).isEqualTo(CellType.ERROR);
            assertThat(dst.getErrorCellValue()).isEqualTo(FormulaError.DIV0.getCode());
        } catch (java.io.IOException e) {
            throw new RuntimeException(e);
        }
    }
}

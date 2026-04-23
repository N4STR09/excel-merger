package com.excelmerger.sheet.column;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

/**
 * Columna EMPTY: la celda se deja en blanco (placeholder para rellenar a mano).
 * Tambien es el fallback al que cae la factoria cuando el tipo indicado no
 * es valido o la configuracion esta incompleta.
 */
public final class EmptyColumnStrategy extends AbstractMesColumnStrategy {

    public EmptyColumnStrategy(String name, boolean greenIfPositive) {
        this(name, greenIfPositive, null, null);
    }

    public EmptyColumnStrategy(String name, boolean greenIfPositive,
                               String fillColor, String redIfNotEqualTo) {
        super(name, greenIfPositive, fillColor, redIfNotEqualTo);
    }

    @Override
    protected void doWriteCell(Cell target, Row srcRow, Sheet source, int sourceHeaderRow0,
                               Workbook workbook, int sourceExcelRow,
                               Map<String, Integer> mesColIndexByName, int mesExcelRow) {
        target.setBlank();
    }
}

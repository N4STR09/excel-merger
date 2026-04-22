package com.excelmerger.sheet.column;

import com.excelmerger.RunReport;
import com.excelmerger.util.PoiUtils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

/**
 * Columna COPY: copia directa del valor de una columna identificada por su
 * cabecera en la hoja origen.
 */
public final class CopyColumnStrategy extends AbstractMesColumnStrategy {

    private final String copyFromHeader;

    public CopyColumnStrategy(String name, boolean greenIfPositive, String copyFromHeader) {
        super(name, greenIfPositive);
        this.copyFromHeader = copyFromHeader;
    }

    @Override
    public void preValidate(Sheet source, int sourceHeaderRow0, Workbook workbook,
                            Map<String, Integer> mesColIndexByName, RunReport report) {
        int col = PoiUtils.findColumnIndex(source.getRow(sourceHeaderRow0), copyFromHeader);
        if (col < 0) {
            report.addWarning("CABECERA",
                    "Columna '" + copyFromHeader + "' no encontrada en '"
                            + source.getSheetName() + "' (usada por COPY '" + name + "').");
            disable();
        }
    }

    @Override
    protected void doWriteCell(Cell target, Row srcRow, Sheet source, int sourceHeaderRow0,
                               Workbook workbook, int sourceExcelRow,
                               Map<String, Integer> mesColIndexByName, int mesExcelRow) {
        int col = PoiUtils.findColumnIndex(source.getRow(sourceHeaderRow0), copyFromHeader);
        if (col < 0) {
            target.setBlank();
            return;
        }
        Cell src = srcRow.getCell(col);
        if (src == null) {
            target.setBlank();
            return;
        }
        PoiUtils.copyCellValue(src, target);
    }
}

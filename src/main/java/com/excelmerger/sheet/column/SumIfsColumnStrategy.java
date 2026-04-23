package com.excelmerger.sheet.column;

import com.excelmerger.RunReport;
import com.excelmerger.util.PoiUtils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * Columna SUMIFS: genera una formula {@code SUMIFS(...)} de Excel que suma
 * condicionalmente valores de una hoja remota, cruzando por una o mas
 * cabeceras.
 *
 * <p>La formula se construye para cada fila por separado porque la
 * referencia local a "fila del origen" depende de la fila actual de MES.</p>
 */
public final class SumIfsColumnStrategy extends AbstractMesColumnStrategy {

    private final String fromSheet;
    private final String sumHeader;
    private final List<String[]> matches;   // cada par: [remoteHeader, localHeader]

    public SumIfsColumnStrategy(String name, boolean greenIfPositive,
                                String fromSheet, String sumHeader, List<String[]> matches) {
        this(name, greenIfPositive, null, null, fromSheet, sumHeader, matches);
    }

    public SumIfsColumnStrategy(String name, boolean greenIfPositive, String fillColor,
                                String redIfNotEqualTo, String fromSheet, String sumHeader,
                                List<String[]> matches) {
        super(name, greenIfPositive, fillColor, redIfNotEqualTo);
        this.fromSheet = fromSheet;
        this.sumHeader = sumHeader;
        this.matches = Collections.unmodifiableList(matches);
    }

    @Override
    public void preValidate(Sheet source, int sourceHeaderRow0, Workbook workbook,
                            Map<String, Integer> mesColIndexByName, RunReport report) {
        Sheet remote = workbook.getSheet(fromSheet);
        if (remote == null) {
            report.addWarning("HOJA",
                    "Hoja '" + fromSheet + "' no existe (usada por SUMIFS '" + name + "').");
            disable();
            return;
        }
        int remoteHeaderRow0 = PoiUtils.detectHeaderRow(remote);
        if (PoiUtils.columnLetter(remote, remoteHeaderRow0, sumHeader) == null) {
            report.addWarning("CABECERA",
                    "Columna '" + sumHeader + "' no encontrada en '"
                            + fromSheet + "' (SUMIFS '" + name + "').");
            disable();
        }
        for (String[] m : matches) {
            String remoteHeader = m[0];
            String localHeader = m[1];
            if (PoiUtils.columnLetter(remote, remoteHeaderRow0, remoteHeader) == null) {
                report.addWarning("CABECERA",
                        "Columna '" + remoteHeader + "' no encontrada en '"
                                + fromSheet + "' (SUMIFS '" + name + "').");
                disable();
            }
            if (PoiUtils.columnLetter(source, sourceHeaderRow0, localHeader) == null) {
                report.addWarning("CABECERA",
                        "Columna '" + localHeader + "' no encontrada en '"
                                + source.getSheetName() + "' (SUMIFS '" + name + "').");
                disable();
            }
        }
    }

    @Override
    protected void doWriteCell(Cell target, Row srcRow, Sheet source, int sourceHeaderRow0,
                               Workbook workbook, int sourceExcelRow,
                               Map<String, Integer> mesColIndexByName, int mesExcelRow) {
        Sheet remote = workbook.getSheet(fromSheet);
        if (remote == null) { target.setBlank(); return; }

        int remoteHeaderRow0 = PoiUtils.detectHeaderRow(remote);
        String sumCol = PoiUtils.columnLetter(remote, remoteHeaderRow0, sumHeader);
        if (sumCol == null) { target.setBlank(); return; }

        StringBuilder f = new StringBuilder();
        f.append("SUMIFS(").append(PoiUtils.quoteSheetName(fromSheet))
                .append("!$").append(sumCol).append(":$").append(sumCol);

        for (String[] m : matches) {
            String remoteHeader = m[0];
            String localHeader = m[1];
            String remoteCol = PoiUtils.columnLetter(remote, remoteHeaderRow0, remoteHeader);
            String localCol = PoiUtils.columnLetter(source, sourceHeaderRow0, localHeader);
            if (remoteCol == null || localCol == null) {
                target.setBlank();
                return;
            }
            f.append(",").append(PoiUtils.quoteSheetName(fromSheet))
                    .append("!$").append(remoteCol).append(":$").append(remoteCol);
            f.append(",").append(PoiUtils.quoteSheetName(source.getSheetName()))
                    .append("!$").append(localCol).append("$").append(sourceExcelRow);
        }
        f.append(")");
        target.setCellFormula(f.toString());
    }
}

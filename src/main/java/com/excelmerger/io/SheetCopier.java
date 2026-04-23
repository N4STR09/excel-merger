package com.excelmerger.io;

import com.excelmerger.util.PoiUtils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * Copia hojas completas (o filas sueltas) entre workbooks de POI. Respeta
 * alturas de fila, anchos de columna, regiones fusionadas y, opcionalmente,
 * los estilos de celda (clonando cada estilo distinto una unica vez por
 * copia, para evitar el limite de estilos de Excel).
 *
 * <p>v1.6.2: las sobrecargas que aceptan {@code asTextColumnIndexes} copian
 * esas columnas como STRING (ver {@link PoiUtils#copyCellValueAsText}) a
 * partir de la fila indicada en {@code firstDataRow0}. Las filas por encima
 * (cabecera y metadatos) se copian sin transformar.</p>
 */
public final class SheetCopier {

    private final boolean copyStyles;

    /**
     * @param copyStyles si {@code true}, al copiar cada celda se clona su
     *                   estilo en el workbook destino. El resultado es mas
     *                   fiel visualmente pero aumenta el tamano del fichero.
     */
    public SheetCopier(boolean copyStyles) {
        this.copyStyles = copyStyles;
    }

    /**
     * Copia completa de {@code source} a {@code target}, incluyendo alturas
     * de fila, anchos de columna y regiones fusionadas. El cache de estilos
     * es local a esta invocacion.
     */
    public void copySheet(Sheet source, Sheet target, Workbook targetWb) {
        copySheet(source, target, targetWb, Collections.emptySet(), 0);
    }

    /**
     * Variante de {@link #copySheet(Sheet, Sheet, Workbook)} que fuerza a
     * STRING las celdas de {@code asTextColumnIndexes} a partir de la fila
     * {@code firstDataRow0} (0-based) inclusive. Las cabeceras y filas por
     * encima se copian intactas.
     *
     * @param asTextColumnIndexes indices 0-based de columnas cuyo valor se
     *                            fuerza a STRING. Puede estar vacio.
     * @param firstDataRow0       primera fila (0-based) a partir de la cual
     *                            se aplica la transformacion. Habitualmente,
     *                            {@code headerRow - 1 + 1 = headerRow}
     *                            (convertido de 1-based a 0-based).
     */
    public void copySheet(Sheet source, Sheet target, Workbook targetWb,
                          Set<Integer> asTextColumnIndexes, int firstDataRow0) {
        Map<Integer, CellStyle> styleCache = new HashMap<>();

        for (int r = 0; r <= source.getLastRowNum(); r++) {
            Row sourceRow = source.getRow(r);
            if (sourceRow == null) continue;
            Row targetRow = target.createRow(r);
            targetRow.setHeight(sourceRow.getHeight());
            boolean applyAsText = r >= firstDataRow0 && !asTextColumnIndexes.isEmpty();
            copyRow(sourceRow, targetRow, targetWb, styleCache,
                    applyAsText ? asTextColumnIndexes : Collections.emptySet());
        }

        int maxCols = PoiUtils.countColumns(source);
        for (int c = 0; c < maxCols; c++) {
            target.setColumnWidth(c, source.getColumnWidth(c));
        }

        for (int i = 0; i < source.getNumMergedRegions(); i++) {
            CellRangeAddress region = source.getMergedRegion(i);
            target.addMergedRegion(region);
        }
    }

    /**
     * Copia una fila individual. {@code styleCache} se reutiliza entre filas
     * para no crear duplicados de un mismo estilo en el workbook destino.
     */
    public void copyRow(Row sourceRow, Row targetRow, Workbook targetWb,
                        Map<Integer, CellStyle> styleCache) {
        copyRow(sourceRow, targetRow, targetWb, styleCache, Collections.emptySet());
    }

    /**
     * Variante que fuerza a STRING las celdas de {@code asTextColumnIndexes}
     * en esta fila. El resto de columnas se copian preservando tipo.
     */
    public void copyRow(Row sourceRow, Row targetRow, Workbook targetWb,
                        Map<Integer, CellStyle> styleCache,
                        Set<Integer> asTextColumnIndexes) {
        for (int c = 0; c < sourceRow.getLastCellNum(); c++) {
            Cell sourceCell = sourceRow.getCell(c);
            if (sourceCell == null) continue;
            Cell targetCell = targetRow.createCell(c);
            if (asTextColumnIndexes.contains(c)) {
                PoiUtils.copyCellValueAsText(sourceCell, targetCell);
            } else {
                PoiUtils.copyCellValue(sourceCell, targetCell);
            }

            if (copyStyles) {
                int styleHash = sourceCell.getCellStyle().hashCode();
                CellStyle newStyle = styleCache.get(styleHash);
                if (newStyle == null) {
                    newStyle = targetWb.createCellStyle();
                    newStyle.cloneStyleFrom(sourceCell.getCellStyle());
                    styleCache.put(styleHash, newStyle);
                }
                targetCell.setCellStyle(newStyle);
            }
        }
    }

    /**
     * Alias de {@link PoiUtils#countColumns(Sheet)}, expuesto aqui porque
     * ExcelMerger lo usa para calcular el rango de columnas en
     * {@code mergeAppendRows}.
     */
    public int countColumns(Sheet sheet) {
        return PoiUtils.countColumns(sheet);
    }
}

package com.excelmerger.sheet.column;

import com.excelmerger.RunReport;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

/**
 * Base comun de las {@link MesColumnStrategy}: gestiona nombre,
 * {@code greenIfPositive} y el flag {@code disabled}. Las subclases solo
 * tienen que implementar {@link #doWriteCell} (y opcionalmente
 * {@link #preValidate}).
 */
public abstract class AbstractMesColumnStrategy implements MesColumnStrategy {

    protected final String name;
    protected final boolean greenIfPositive;
    protected final String fillColor;
    protected final String redIfNotEqualTo;
    protected boolean disabled;

    protected AbstractMesColumnStrategy(String name, boolean greenIfPositive) {
        this(name, greenIfPositive, null, null);
    }

    protected AbstractMesColumnStrategy(String name, boolean greenIfPositive,
                                        String fillColor, String redIfNotEqualTo) {
        this.name = name;
        this.greenIfPositive = greenIfPositive;
        this.fillColor = fillColor;
        this.redIfNotEqualTo = redIfNotEqualTo;
    }

    @Override public final String getName() { return name; }
    @Override public final boolean isGreenIfPositive() { return greenIfPositive; }
    @Override public final String getFillColor() { return fillColor; }
    @Override public final String getRedIfNotEqualTo() { return redIfNotEqualTo; }
    @Override public final boolean isDisabled() { return disabled; }

    /** Marca la columna como disabled: sus celdas se escribiran en blanco. */
    protected final void disable() { this.disabled = true; }

    /**
     * Template method: si la columna esta disabled, escribe blank.
     * En caso contrario delega en {@link #doWriteCell}.
     */
    @Override
    public final void writeCell(Cell target, Row srcRow, Sheet source, int sourceHeaderRow0,
                                Workbook workbook, int sourceExcelRow,
                                Map<String, Integer> mesColIndexByName, int mesExcelRow) {
        if (disabled) {
            target.setBlank();
            return;
        }
        doWriteCell(target, srcRow, source, sourceHeaderRow0, workbook, sourceExcelRow,
                mesColIndexByName, mesExcelRow);
    }

    /** Logica real de escritura de la celda. */
    protected abstract void doWriteCell(Cell target, Row srcRow, Sheet source, int sourceHeaderRow0,
                                        Workbook workbook, int sourceExcelRow,
                                        Map<String, Integer> mesColIndexByName, int mesExcelRow);

    /** Implementacion por defecto: no hace nada. Subclases con referencias sobreescriben. */
    @Override
    public void preValidate(Sheet source, int sourceHeaderRow0, Workbook workbook,
                            Map<String, Integer> mesColIndexByName, RunReport report) {
        // no-op por defecto
    }
}

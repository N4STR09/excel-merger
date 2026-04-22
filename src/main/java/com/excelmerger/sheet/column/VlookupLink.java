package com.excelmerger.sheet.column;

/**
 * Enlace VLOOKUP resuelto a nivel de datos: relaciona una columna MES que
 * se usa como clave de busqueda con la hoja de lookup destino.
 *
 * <p>Poblado por {@code MesSheetBuilder} al escanear las plantillas de
 * {@link FormulaColumnStrategy}. Se usa para detectar "apps sin mapeo":
 * valores que aparecen en la columna-clave pero no estan en la hoja de
 * lookup correspondiente.</p>
 */
public final class VlookupLink {

    /** Nombre de la columna MES usada como clave de busqueda. */
    public final String sourceColName;

    /** Indice (0-based) de esa columna en la hoja MES. */
    public final int sourceColIdx;

    /** Nombre de la hoja de lookup destino (referenciada en el VLOOKUP). */
    public final String lookupSheet;

    public VlookupLink(String sourceColName, int sourceColIdx, String lookupSheet) {
        this.sourceColName = sourceColName;
        this.sourceColIdx = sourceColIdx;
        this.lookupSheet = lookupSheet;
    }
}

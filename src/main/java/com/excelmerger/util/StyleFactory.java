package com.excelmerger.util;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Fabrica de {@link CellStyle} reutilizables en varios builders.
 *
 * <p>Centraliza el estilo de cabecera (negrita, fondo gris claro, borde
 * inferior fino) y el estilo de titulo (negrita, tamano 14 pt), que antes
 * estaban duplicados en {@code LookupSheetBuilder}, {@code DerivedSheetBuilder}
 * y {@code MesSheetBuilder}.</p>
 *
 * <p>Cada invocacion crea un estilo nuevo dentro del workbook indicado (POI
 * no permite compartir {@code CellStyle} entre workbooks). No cachea: cada
 * llamada produce una instancia independiente.</p>
 */
public final class StyleFactory {

    private StyleFactory() {
        // Utility class
    }

    /**
     * Estilo tipico de cabecera: negrita, fondo gris 25 %, borde inferior fino.
     */
    public static CellStyle header(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        s.setFont(f);
        s.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setBorderBottom(BorderStyle.THIN);
        return s;
    }

    /**
     * Estilo de titulo: negrita, 14 pt, sin fondo ni bordes.
     */
    public static CellStyle title(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints((short) 14);
        s.setFont(f);
        return s;
    }
}

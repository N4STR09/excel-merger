package com.excelmerger.util;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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

    // ==================================================================
    //  Estilos de la hoja Resumen (v1.6.0)
    // ==================================================================
    //  La hoja Resumen reproduce la paleta del Excel de referencia:
    //  cabeceras de bloque con fondo negro y texto blanco, sub-cabeceras
    //  grises o naranjas, celdas de balance rojo/amarillo, y celdas de
    //  referencia lavanda para el pre-cierre. Los colores se fijan como
    //  ARGB exactos con XSSFColor; al usar XSSF la herramienta ya genera
    //  libros .xlsx, asi que el cast es seguro.
    //
    //  Cada metodo crea una nueva instancia (POI no permite compartir
    //  estilos entre workbooks y tampoco mutar uno ya aplicado sin
    //  efectos colaterales).
    // ==================================================================

    private static final String COLOR_BLACK        = "FF000000";
    private static final String COLOR_WHITE        = "FFFFFFFF";
    private static final String COLOR_GRAY_LIGHT   = "FFC0C0C0";

    /**
     * Cabecera de bloque (A2, A16, A20, G20, A33): fondo negro, texto
     * blanco en negrita, bordes thin.
     */
    public static CellStyle summaryBlockHeader(Workbook wb) {
        XSSFCellStyle s = (XSSFCellStyle) wb.createCellStyle();
        XSSFFont f = (XSSFFont) wb.createFont();
        f.setBold(true);
        f.setFontName("Calibri");
        f.setFontHeightInPoints((short) 10);
        f.setColor(argb(wb, COLOR_WHITE));
        s.setFont(f);
        s.setFillForegroundColor(argb(wb, COLOR_BLACK));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        applyThinBorders(s);
        s.setAlignment(HorizontalAlignment.LEFT);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    /**
     * Sub-cabecera gris (A21, A29, B34, C14, D3): fondo gris claro, texto
     * negro, bordes thin, negrita opcional controlada por {@code bold}.
     */
    public static CellStyle summarySubHeaderGray(Workbook wb, boolean bold) {
        return summaryFilledCell(wb, COLOR_GRAY_LIGHT, bold);
    }

    /**
     * Celda generica de datos de Resumen: Calibri 10, sin fill, bordes
     * thin. Alineacion izquierda por defecto.
     */
    public static CellStyle summaryValueCell(Workbook wb) {
        XSSFCellStyle s = (XSSFCellStyle) wb.createCellStyle();
        XSSFFont f = (XSSFFont) wb.createFont();
        f.setFontName("Calibri");
        f.setFontHeightInPoints((short) 10);
        s.setFont(f);
        applyThinBorders(s);
        s.setAlignment(HorizontalAlignment.LEFT);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    /**
     * Celda de totales (fila TOT H / Total en bloque 3): bordes medium,
     * negrita, formato numerico 0.0, opcionalmente con fill gris.
     */
    public static CellStyle summaryTotalCell(Workbook wb, boolean filled) {
        XSSFCellStyle s = (XSSFCellStyle) wb.createCellStyle();
        XSSFFont f = (XSSFFont) wb.createFont();
        f.setBold(true);
        f.setFontName("Calibri");
        f.setFontHeightInPoints((short) 10);
        s.setFont(f);
        if (filled) {
            s.setFillForegroundColor(argb(wb, COLOR_GRAY_LIGHT));
            s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        s.setBorderTop(BorderStyle.MEDIUM);
        s.setBorderBottom(BorderStyle.MEDIUM);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
        s.setDataFormat(wb.createDataFormat().getFormat("0.0"));
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    /**
     * Celda numerica con formato 0.0 y bordes thin (cuerpo del bloque 4
     * y celdas de horas en general).
     */
    public static CellStyle summaryNumericCell(Workbook wb) {
        XSSFCellStyle s = (XSSFCellStyle) wb.createCellStyle();
        XSSFFont f = (XSSFFont) wb.createFont();
        f.setFontName("Calibri");
        f.setFontHeightInPoints((short) 10);
        s.setFont(f);
        applyThinBorders(s);
        s.setDataFormat(wb.createDataFormat().getFormat("0.0"));
        s.setAlignment(HorizontalAlignment.RIGHT);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    // ------------------------------------------------------------------
    //  Helpers internos
    // ------------------------------------------------------------------

    private static CellStyle summaryFilledCell(Workbook wb, String fillHex, boolean bold) {
        XSSFCellStyle s = (XSSFCellStyle) wb.createCellStyle();
        XSSFFont f = (XSSFFont) wb.createFont();
        f.setBold(bold);
        f.setFontName("Calibri");
        f.setFontHeightInPoints((short) 10);
        s.setFont(f);
        s.setFillForegroundColor(argb(wb, fillHex));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        applyThinBorders(s);
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    private static void applyThinBorders(CellStyle s) {
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
    }

    /**
     * Crea un XSSFColor ARGB a partir de una cadena hexadecimal tipo
     * "FFRRGGBB". Requiere que el workbook sea XSSF (siempre lo es en
     * este proyecto: el resultado se escribe en .xlsx con
     * {@code new XSSFWorkbook()}).
     */
    private static XSSFColor argb(Workbook wb, String hex) {
        if (!(wb instanceof XSSFWorkbook)) {
            // Fallback defensivo: si alguien pasara HSSFWorkbook no
            // cascariamos, pero los colores se aproximarian a los
            // IndexedColors mas cercanos. En practica esto no se ejecuta.
            return new XSSFColor(new byte[] {
                    (byte) Integer.parseInt(hex.substring(0, 2), 16),
                    (byte) Integer.parseInt(hex.substring(2, 4), 16),
                    (byte) Integer.parseInt(hex.substring(4, 6), 16),
                    (byte) Integer.parseInt(hex.substring(6, 8), 16),
            }, null);
        }
        byte[] rgba = new byte[] {
                (byte) Integer.parseInt(hex.substring(0, 2), 16),
                (byte) Integer.parseInt(hex.substring(2, 4), 16),
                (byte) Integer.parseInt(hex.substring(4, 6), 16),
                (byte) Integer.parseInt(hex.substring(6, 8), 16),
        };
        return new XSSFColor(rgba, ((XSSFWorkbook) wb).getStylesSource().getIndexedColors());
    }
}

package com.excelmerger;

import com.excelmerger.util.StyleFactory;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * Vuelca los warnings acumulados en {@link RunReport} a una hoja extra
 * ({@code _Avisos} por defecto) dentro del Excel resultado. Utilidad: permite
 * revisar desde el propio Excel las apps sin mapeo, cabeceras no encontradas,
 * perfiles sin match, etc., sin tener que abrir el log.
 *
 * <p>Configuracion (todas las claves son opcionales):</p>
 * <ul>
 *   <li>{@code report.inExcel=true|false} (default {@code false}) — activa o
 *       desactiva la generacion de la hoja.</li>
 *   <li>{@code report.sheetName=_Avisos} (default {@code _Avisos}) — nombre
 *       de la hoja destino.</li>
 *   <li>{@code report.hidden=true|false} (default {@code true}) — si
 *       {@code true}, la hoja se crea oculta.</li>
 * </ul>
 *
 * <p>La hoja no se crea si no hay warnings (evita hojas vacias ruidosas) o si
 * {@code report.inExcel=false}. Si ya existe una hoja con ese nombre (colision
 * con un nombre que el usuario haya asignado a un lookup, por ejemplo), se
 * omite y se registra un warning {@code HOJA} en el propio report.</p>
 */
public class AvisosSheetBuilder {

    private static final Logger log = LoggerFactory.getLogger(AvisosSheetBuilder.class);

    private static final String DEFAULT_SHEET_NAME = "_Avisos";
    private static final String HEADER_CATEGORIA = "Categoria";
    private static final String HEADER_MENSAJE   = "Mensaje";

    private final ConfigLoader config;
    private final RunReport report;

    public AvisosSheetBuilder(ConfigLoader config, RunReport report) {
        this.config = config;
        this.report = report;
    }

    /**
     * Construye la hoja de avisos en el workbook indicado. No-op si
     * {@code report.inExcel=false} o si no hay warnings acumulados.
     */
    public void build(Workbook workbook) {
        boolean enabled = config.getBoolean("report.inExcel", false);
        if (!enabled) {
            return;
        }

        List<RunReport.Warning> warnings = report.warnings();
        if (warnings.isEmpty()) {
            log.info("[Avisos] No hay warnings; hoja de avisos omitida.");
            return;
        }

        String sheetName = config.get("report.sheetName", DEFAULT_SHEET_NAME);
        boolean hidden = config.getBoolean("report.hidden", true);

        if (workbook.getSheet(sheetName) != null) {
            log.warn("Ya existe una hoja '{}'; hoja de avisos omitida.", sheetName);
            report.addWarning("HOJA",
                    "Ya existe una hoja '" + sheetName + "' en el libro; hoja de avisos omitida.");
            return;
        }

        Sheet sheet = workbook.createSheet(sheetName);
        CellStyle headerStyle = StyleFactory.header(workbook);

        Row header = sheet.createRow(0);
        Cell h1 = header.createCell(0); h1.setCellValue(HEADER_CATEGORIA); h1.setCellStyle(headerStyle);
        Cell h2 = header.createCell(1); h2.setCellValue(HEADER_MENSAJE);   h2.setCellStyle(headerStyle);

        int rowIdx = 1;
        for (RunReport.Warning w : warnings) {
            Row r = sheet.createRow(rowIdx++);
            r.createCell(0).setCellValue(w.category);
            r.createCell(1).setCellValue(w.message);
        }

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);

        if (hidden) {
            int idx = workbook.getSheetIndex(sheet);
            if (idx >= 0) workbook.setSheetHidden(idx, true);
        }

        int totalRows = sheet.getLastRowNum() + 1;
        report.addSheet(sheetName, totalRows);
        log.info("Hoja '{}' creada con {} aviso(s){} ({} filas)",
                sheetName, warnings.size(), hidden ? " (oculta)" : "", totalRows);
    }
}

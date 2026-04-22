package com.excelmerger;

import com.excelmerger.util.StyleFactory;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/**
 * Construye "hojas de lookup": tablas estaticas de mapeo clave/valor que viven
 * dentro del Excel resultado y que otras hojas (p. ej. MES) pueden consultar
 * con VLOOKUP.
 *
 * Configuracion:
 *   lookup.sheets=Equipos,Otra
 *
 *   lookup.Equipos.header1=App
 *   lookup.Equipos.header2=Equipo
 *   lookup.Equipos.hidden=true          (opcional, default false)
 *   lookup.Equipos.data=\
 *     DF:Iker,\
 *     HE:Jon,\
 *     EW:JAVA
 *
 * Separador por defecto entre clave y valor: ':'  (configurable con lookup.X.sep)
 * Separador de entradas: coma
 */
public class LookupSheetBuilder {

    private static final Logger log = LoggerFactory.getLogger(LookupSheetBuilder.class);

    private final ConfigLoader config;
    private final RunReport report;

    public LookupSheetBuilder(ConfigLoader config, RunReport report) {
        this.config = config;
        this.report = report;
    }

    public void buildAll(Workbook workbook) {
        String list = config.get("lookup.sheets", "").trim();
        if (list.isEmpty()) return;

        for (String rawId : list.split(",")) {
            String id = rawId.trim();
            if (id.isEmpty()) continue;
            buildOne(workbook, id);
        }
    }

    private void buildOne(Workbook workbook, String id) {
        String prefix = "lookup." + id + ".";

        if (workbook.getSheet(id) != null) {
            log.info("Ya existe una hoja '" + id + "'. Se omite.");
            report.addWarning("LOOKUP",
                    "Ya existe una hoja '" + id + "' en el libro; lookup omitido.");
            return;
        }

        String header1 = config.get(prefix + "header1", "Key");
        String header2 = config.get(prefix + "header2", "Value");
        String sep = config.get(prefix + "sep", ":");
        boolean hidden = config.getBoolean(prefix + "hidden", false);
        String data = config.get(prefix + "data", "").trim();

        if (data.isEmpty()) {
            log.info("'" + id + "' no tiene datos (lookup." + id + ".data). Se omite.");
            report.addWarning("LOOKUP",
                    "Lookup '" + id + "' sin datos ('lookup." + id + ".data'). Omitido.");
            return;
        }

        Sheet sheet = workbook.createSheet(id);
        CellStyle headerStyle = StyleFactory.header(workbook);

        Row header = sheet.createRow(0);
        Cell h1 = header.createCell(0); h1.setCellValue(header1); h1.setCellStyle(headerStyle);
        Cell h2 = header.createCell(1); h2.setCellValue(header2); h2.setCellStyle(headerStyle);

        List<String[]> pairs = parsePairs(id, data, sep);
        Set<String> keys = new LinkedHashSet<>();
        int rowIdx = 1;
        for (String[] pair : pairs) {
            Row r = sheet.createRow(rowIdx++);
            r.createCell(0).setCellValue(pair[0]);
            r.createCell(1).setCellValue(pair[1]);
            keys.add(pair[0]);
        }

        // Registrar las claves de este lookup para la deteccion de apps sin mapeo
        report.registerLookupKeys(id, keys);

        for (int c = 0; c < 2; c++) {
            sheet.autoSizeColumn(c);
        }

        if (hidden) {
            int idx = workbook.getSheetIndex(sheet);
            if (idx >= 0) workbook.setSheetHidden(idx, true);
        }

        int totalRows = sheet.getLastRowNum() + 1;
        report.addSheet(id, totalRows);
        log.info("Hoja '{}' creada con {} entradas{} ({} filas)",
                id, pairs.size(), hidden ? " (oculta)" : "", totalRows);
    }

    private List<String[]> parsePairs(String id, String data, String sep) {
        List<String[]> out = new ArrayList<>();
        for (String raw : data.split(",")) {
            String entry = raw.trim();
            if (entry.isEmpty()) continue;
            int i = entry.indexOf(sep);
            if (i < 0) {
                log.info("[Lookup] Entrada sin separador '{}': {}", sep, entry);
                report.addWarning("LOOKUP",
                        "Entrada sin separador '" + sep + "' en '" + id + "': '" + entry + "'.");
                continue;
            }
            String key = entry.substring(0, i).trim();
            String value = entry.substring(i + sep.length()).trim();
            if (key.isEmpty()) {
                report.addWarning("LOOKUP",
                        "Entrada con clave vacia en '" + id + "': '" + entry + "'.");
                continue;
            }
            out.add(new String[]{key, value});
        }
        return out;
    }

}

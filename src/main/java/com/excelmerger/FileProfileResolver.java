package com.excelmerger;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * Resuelve el perfil de un archivo Excel analizando su CONTENIDO (cabeceras
 * de columna y, opcionalmente, valores en celdas concretas). Asi el nombre
 * del fichero es irrelevante: si la plantilla no cambia, el programa siempre
 * reconoce de que tipo de documento se trata.
 *
 * Sintaxis en el config:
 *   profiles=contabilidad,ventas
 *
 *   profile.contabilidad.sheetName=Contabilidad
 *   profile.contabilidad.detect.headerRow=1
 *   profile.contabilidad.detect.headers=Factura,Importe,IVA
 *   profile.contabilidad.detect.minMatches=2
 *   profile.contabilidad.detect.cellValue.A1=Libro Mayor    (opcional)
 *   profile.contabilidad.detect.sheetIndex=0                (opcional)
 */
public class FileProfileResolver {

    private static final Logger log = LoggerFactory.getLogger(FileProfileResolver.class);

    /** Nombre maximo permitido por Excel para una hoja. */
    public static final int MAX_SHEET_NAME_LEN = 31;

    private final List<FileProfile> profiles = new ArrayList<>();

    public FileProfileResolver(ConfigLoader config) {
        String list = config.get("profiles", "").trim();
        if (list.isEmpty()) return;

        for (String rawId : list.split(",")) {
            String id = rawId.trim();
            if (id.isEmpty()) continue;
            profiles.add(FileProfile.fromConfig(config, id));
        }
    }

    public boolean hasProfiles() {
        return !profiles.isEmpty();
    }

    /**
     * Inspecciona el workbook y devuelve el perfil que coincide, o null.
     * Ademas, guarda en el propio perfil el detalle de la deteccion para logs.
     */
    public FileProfile resolve(Workbook workbook, File sourceFile) {
        for (FileProfile p : profiles) {
            DetectionResult r = p.detect(workbook);
            if (r.matched) {
                log.info("'{}' -> perfil '{}' (cabeceras encontradas: {})",
                        sourceFile.getName(), p.id, r.matchedHeaders);
                return p;
            }
        }
        return null;
    }

    /**
     * Trunca y sanea un nombre para que sea valido como nombre de hoja en Excel.
     */
    public static String safeSheetName(String name) {
        if (name == null) return "Sheet";
        String cleaned = name.replaceAll("[\\\\/*?:\\[\\]]", "_");
        if (cleaned.length() > MAX_SHEET_NAME_LEN) {
            cleaned = cleaned.substring(0, MAX_SHEET_NAME_LEN);
        }
        return cleaned;
    }

    // ================================================================
    //  FileProfile: un perfil y su logica de deteccion
    // ================================================================
    public static class FileProfile {
        final String id;
        final String sheetName;
        final int sheetIndex;
        final int headerRow;           // 1-indexado
        final List<String> expectedHeaders;
        final int minMatches;
        final Map<String, String> expectedCellValues;  // ref -> valor esperado

        private FileProfile(String id, String sheetName, int sheetIndex,
                            int headerRow, List<String> expectedHeaders, int minMatches,
                            Map<String, String> expectedCellValues) {
            this.id = id;
            this.sheetName = sheetName;
            this.sheetIndex = sheetIndex;
            this.headerRow = headerRow;
            this.expectedHeaders = expectedHeaders;
            this.minMatches = minMatches;
            this.expectedCellValues = expectedCellValues;
        }

        static FileProfile fromConfig(ConfigLoader config, String id) {
            String prefix = "profile." + id + ".";
            String sheetName = config.get(prefix + "sheetName", id);
            int sheetIndex = config.getInt(prefix + "detect.sheetIndex", 0);
            int headerRow = config.getInt(prefix + "detect.headerRow", 1);

            String headersRaw = config.get(prefix + "detect.headers", "").trim();
            List<String> headers = headersRaw.isEmpty()
                    ? Collections.emptyList()
                    : Arrays.stream(headersRaw.split(","))
                            .map(String::trim)
                            .filter(s -> !s.isEmpty())
                            .collect(Collectors.toList());

            // Default: TODAS las cabeceras listadas deben coincidir
            int minMatches = config.getInt(prefix + "detect.minMatches", Math.max(1, headers.size()));

            // Buscar todas las propiedades de tipo detect.cellValue.<REF>
            Map<String, String> cellValues = new LinkedHashMap<>();
            String cellPrefix = prefix + "detect.cellValue.";
            Set<String> keys = config.getRawProperties().stringPropertyNames();
            for (String k : keys) {
                if (k.startsWith(cellPrefix)) {
                    String ref = k.substring(cellPrefix.length()).trim();
                    String expected = config.get(k).trim();
                    cellValues.put(ref, expected);
                }
            }

            return new FileProfile(id, sheetName, sheetIndex, headerRow,
                    headers, minMatches, cellValues);
        }

        /**
         * Examina el workbook y decide si coincide con este perfil.
         */
        DetectionResult detect(Workbook wb) {
            DetectionResult result = new DetectionResult();

            if (sheetIndex < 0 || sheetIndex >= wb.getNumberOfSheets()) {
                return result; // no match
            }
            Sheet sheet = wb.getSheetAt(sheetIndex);

            // 1. Comprobar cabeceras
            if (!expectedHeaders.isEmpty()) {
                List<String> actualHeaders = readRowAsStrings(sheet, headerRow - 1);
                int found = 0;
                for (String expected : expectedHeaders) {
                    String needle = expected.toLowerCase();
                    boolean present = actualHeaders.stream()
                            .anyMatch(h -> h.toLowerCase().contains(needle));
                    if (present) {
                        found++;
                        result.matchedHeaders.add(expected);
                    }
                }
                if (found < minMatches) {
                    return result; // no match
                }
            }

            // 2. Comprobar valores de celda concretos (si hay)
            for (Map.Entry<String, String> entry : expectedCellValues.entrySet()) {
                String actual = readCellAsString(sheet, entry.getKey());
                String expected = entry.getValue();
                if (actual == null || !actual.toLowerCase().contains(expected.toLowerCase())) {
                    return result; // no match
                }
                result.matchedCellValues.put(entry.getKey(), expected);
            }

            // Si llegamos aqui, todos los criterios se cumplen
            result.matched = true;
            return result;
        }

        // ------- helpers -------
        private List<String> readRowAsStrings(Sheet sheet, int rowIdx) {
            List<String> out = new ArrayList<>();
            if (rowIdx < 0) return out;
            Row row = sheet.getRow(rowIdx);
            if (row == null) return out;
            for (int c = 0; c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                String s = cellToString(cell);
                if (s != null) out.add(s.trim());
            }
            return out;
        }

        private String readCellAsString(Sheet sheet, String ref) {
            try {
                CellReference cr = new CellReference(ref);
                Row row = sheet.getRow(cr.getRow());
                if (row == null) return null;
                return cellToString(row.getCell(cr.getCol()));
            } catch (Exception e) {
                return null;
            }
        }

        private String cellToString(Cell cell) {
            if (cell == null) return null;
            switch (cell.getCellType()) {
                case STRING:  return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString();
                    }
                    double d = cell.getNumericCellValue();
                    return (d == (long) d) ? String.valueOf((long) d) : String.valueOf(d);
                case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
                case FORMULA: return cell.getCellFormula();
                default:      return "";
            }
        }

        public String getId() { return id; }
        public String getSheetName() { return sheetName; }
    }

    /** Info de diagnostico del intento de deteccion. */
    public static class DetectionResult {
        boolean matched = false;
        List<String> matchedHeaders = new ArrayList<>();
        Map<String, String> matchedCellValues = new LinkedHashMap<>();
    }
}

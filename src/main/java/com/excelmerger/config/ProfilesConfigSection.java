package com.excelmerger.config;

import com.excelmerger.ConfigLoader;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import static com.excelmerger.config.ValidationHelpers.isBlank;
import static com.excelmerger.config.ValidationHelpers.parseCsv;
import static com.excelmerger.config.ValidationHelpers.parsePositiveInt;

/**
 * Valida los bloques {@code profiles=}, {@code lookup.sheets=} y
 * {@code derived.sheets=} del config.
 *
 * <p>Las tres validaciones se agrupan aquí porque las tres construyen
 * conjuntos de identificadores de hojas que el orquestador combina luego
 * en {@code knownSheets}.</p>
 *
 * <p>Extraída de {@code ConfigValidator} durante el refactor de la
 * Sesión F (v2.5.0). Package-private.</p>
 */
final class ProfilesConfigSection {

    private final ConfigLoader config;
    private final List<String> errors;

    ProfilesConfigSection(ConfigLoader config, List<String> errors) {
        this.config = config;
        this.errors = errors;
    }

    /**
     * Valida cada perfil y devuelve el conjunto de sheetName declarados (para
     * resolver referencias a hojas desde otras secciones del config).
     */
    Set<String> validateProfiles() {
        Set<String> sheetNames = new LinkedHashSet<>();
        List<String> ids = parseCsv(config.get("profiles", ""));
        for (String id : ids) {
            String prefix = "profile." + id + ".";
            String sheetName = config.get(prefix + "sheetName", id);
            sheetNames.add(sheetName);

            // Al menos un criterio de deteccion: headers o cellValue.*
            String headersRaw = config.get(prefix + "detect.headers", "");
            List<String> headers = parseCsv(headersRaw);
            boolean hasCellValue = hasAnyKeyStartingWith(prefix + "detect.cellValue.");

            if (headers.isEmpty() && !hasCellValue) {
                errors.add("Perfil '" + id + "' incompleto: sin '" + prefix + "detect.headers' "
                        + "ni ninguna '" + prefix + "detect.cellValue.<REF>'.");
            }

            Integer headerRow = parsePositiveInt(config, prefix + "detect.headerRow", 1, errors);
            if (headerRow != null && headerRow < 1) {
                errors.add(prefix + "detect.headerRow debe ser >= 1 (actual: " + headerRow + ").");
            }

            Integer sheetIndex = parsePositiveInt(config, prefix + "detect.sheetIndex", 0, errors);
            if (sheetIndex != null && sheetIndex < 0) {
                errors.add(prefix + "detect.sheetIndex debe ser >= 0 (actual: " + sheetIndex + ").");
            }

            Integer minMatches = parsePositiveInt(config, prefix + "detect.minMatches", null, errors);
            if (minMatches != null && minMatches < 1) {
                errors.add(prefix + "detect.minMatches debe ser >= 1 (actual: " + minMatches + ").");
            }
            if (minMatches != null && !headers.isEmpty() && minMatches > headers.size()) {
                errors.add(prefix + "detect.minMatches (" + minMatches
                        + ") es mayor que el numero de cabeceras declaradas ("
                        + headers.size() + ").");
            }

            // v1.8.1: trim.columns solo tiene sentido junto con asText.columns
            // (el trim es una capa sobre la rama STRING del cast). Si el
            // usuario declara trim sin asText, es un error de configuracion
            // puro (nada que trimar). Cada columna suelta de trim que no
            // este en asText se trata como warning en runtime (ExcelMerger
            // ya lo emite), pero si la lista entera de asText esta vacia,
            // es error de config.
            String trimRaw = config.get(prefix + "trim.columns", "");
            List<String> trimCols = parseCsv(trimRaw);
            List<String> asTextCols = parseCsv(config.get(prefix + "asText.columns", ""));
            if (!trimCols.isEmpty() && asTextCols.isEmpty()) {
                errors.add(prefix + "trim.columns declarado (" + trimRaw
                        + ") pero '" + prefix + "asText.columns' vacio. "
                        + "El trim solo aplica a columnas casteadas a STRING; "
                        + "sin asText.columns no trima nada.");
            }
        }
        return sheetNames;
    }

    Set<String> validateLookups() {
        Set<String> ids = new LinkedHashSet<>(parseCsv(config.get("lookup.sheets", "")));
        for (String id : ids) {
            String data = config.get("lookup." + id + ".data", "");
            if (isBlank(data)) {
                errors.add("Lookup '" + id + "' sin datos: falta 'lookup." + id + ".data'.");
            }
        }
        return ids;
    }

    Set<String> collectDerivedIds() {
        return new LinkedHashSet<>(parseCsv(config.get("derived.sheets", "")));
    }

    private boolean hasAnyKeyStartingWith(String prefix) {
        for (String key : config.getRawProperties().stringPropertyNames()) {
            if (key.startsWith(prefix)) {
                return true;
            }
        }
        return false;
    }
}

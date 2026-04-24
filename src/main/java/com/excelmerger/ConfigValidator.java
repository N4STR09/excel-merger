package com.excelmerger;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Validacion estatica del config.properties ANTES de abrir ningun archivo.
 *
 * <p>Acumula TODOS los errores detectados y los devuelve. {@link Main} decide
 * si aborta con exit code 2 (strictValidation=true) o solo los registra como
 * warnings en el {@link RunReport} (strictValidation=false).</p>
 *
 * <p>Chequeos soportados:</p>
 * <ul>
 *   <li>Entrada/salida minimas.</li>
 *   <li>{@code merge.mode} valido.</li>
 *   <li>Perfiles: cada perfil listado en {@code profiles=} tiene al menos un
 *       criterio de deteccion ({@code detect.headers} o {@code detect.cellValue.*}).</li>
 *   <li>Columnas MES: tipo valido y campos obligatorios por tipo.</li>
 *   <li>Placeholders {@code {col:X}} de fórmulas MES referencian nombres existentes.</li>
 *   <li>Referencias a hojas: SUMIFS {@code from} y AGGREGATION {@code sourceSheet}
 *       deben apuntar a una hoja conocida (perfil, lookup, derivada o resultado de fusion).</li>
 *   <li>Lookups: cada hoja listada tiene datos.</li>
 *   <li>Derivadas: tipo valido y campos requeridos por tipo.</li>
 * </ul>
 */
public class ConfigValidator {

    private static final Set<String> VALID_COL_TYPES = new LinkedHashSet<>(
            Arrays.asList("COPY", "SUMIFS", "FORMULA", "EMPTY"));
    private static final Set<String> VALID_DERIVED_TYPES = new LinkedHashSet<>(
            Arrays.asList("FORMULAS", "AGGREGATION"));
    private static final Set<String> VALID_AGG_FUNCS = new LinkedHashSet<>(
            Arrays.asList("SUM", "AVG", "COUNT", "MIN", "MAX"));
    private static final Set<String> VALID_MERGE_MODES = new LinkedHashSet<>(
            Arrays.asList("SHEETS_SEPARATE", "APPEND_ROWS"));
    private static final Set<String> VALID_FILL_COLORS = new LinkedHashSet<>(
            Arrays.asList("LIGHT_GREEN", "MEDIUM_GREEN", "LIGHT_BLUE",
                    "LIGHT_YELLOW", "LIGHT_RED", "LIGHT_LAVENDER"));

    private static final Pattern PLACEHOLDER_COL = Pattern.compile("\\{col:([^}]+)\\}");
    private static final Pattern COLUMN_LETTER = Pattern.compile("[A-Za-z]+");

    private final ConfigLoader config;
    private final List<String> errors = new ArrayList<>();

    public ConfigValidator(ConfigLoader config) {
        this.config = config;
    }

    /**
     * Ejecuta todas las validaciones y devuelve la lista completa de errores.
     * Orden estable: se acumulan en el orden en el que se detectan.
     */
    public List<String> validate() {
        errors.clear();

        validateInputOutput();
        String mergeMode = validateMergeMode();

        Set<String> profileSheetNames = validateProfiles();
        Set<String> lookupIds = validateLookups();
        Set<String> derivedIds = collectDerivedIds();

        Set<String> knownSheets = buildKnownSheets(mergeMode, profileSheetNames, lookupIds, derivedIds);

        validateMes(knownSheets);
        validateDerived(knownSheets);
        validateSummary(knownSheets);
        validateOrphans(knownSheets);

        return Collections.unmodifiableList(new ArrayList<>(errors));
    }

    // ==================================================================
    //  Bloques de validacion
    // ==================================================================

    private void validateInputOutput() {
        if (isBlank(config.get("input.directory", ""))) {
            errors.add("input.directory: propiedad requerida.");
        }
        if (isBlank(config.get("output.file", ""))) {
            errors.add("output.file: propiedad requerida.");
        }
    }

    private String validateMergeMode() {
        String mode = config.get("merge.mode", "SHEETS_SEPARATE").toUpperCase();
        if (!VALID_MERGE_MODES.contains(mode)) {
            errors.add("merge.mode: valor invalido '" + mode + "'. Valores permitidos: "
                    + VALID_MERGE_MODES);
            return "SHEETS_SEPARATE";
        }
        return mode;
    }

    /**
     * Valida cada perfil y devuelve el conjunto de sheetName declarados (para
     * resolver referencias a hojas desde otras secciones del config).
     */
    private Set<String> validateProfiles() {
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

            Integer headerRow = parsePositiveInt(prefix + "detect.headerRow", 1);
            if (headerRow != null && headerRow < 1) {
                errors.add(prefix + "detect.headerRow debe ser >= 1 (actual: " + headerRow + ").");
            }

            Integer sheetIndex = parsePositiveInt(prefix + "detect.sheetIndex", 0);
            if (sheetIndex != null && sheetIndex < 0) {
                errors.add(prefix + "detect.sheetIndex debe ser >= 0 (actual: " + sheetIndex + ").");
            }

            Integer minMatches = parsePositiveInt(prefix + "detect.minMatches", null);
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

    private Set<String> validateLookups() {
        Set<String> ids = new LinkedHashSet<>(parseCsv(config.get("lookup.sheets", "")));
        for (String id : ids) {
            String data = config.get("lookup." + id + ".data", "");
            if (isBlank(data)) {
                errors.add("Lookup '" + id + "' sin datos: falta 'lookup." + id + ".data'.");
            }
        }
        return ids;
    }

    private Set<String> collectDerivedIds() {
        return new LinkedHashSet<>(parseCsv(config.get("derived.sheets", "")));
    }

    private Set<String> buildKnownSheets(String mergeMode,
                                         Set<String> profileSheetNames,
                                         Set<String> lookupIds,
                                         Set<String> derivedIds) {
        Set<String> known = new LinkedHashSet<>();
        if ("APPEND_ROWS".equals(mergeMode)) {
            known.add(config.get("merge.resultSheetName", "Datos_Fusionados"));
        } else {
            // En SHEETS_SEPARATE cada perfil aporta una hoja con su sheetName
            known.addAll(profileSheetNames);
        }
        known.addAll(lookupIds);
        known.addAll(derivedIds);
        // La propia hoja MES tambien es referenciable
        if (config.getBoolean("mes.enabled", false)) {
            known.add(config.get("mes.sheetName", "MES"));
        }
        return known;
    }

    private void validateMes(Set<String> knownSheets) {
        if (!config.getBoolean("mes.enabled", false)) {
            return;
        }
        String sourceSheet = config.get("mes.sourceSheet", "");
        if (isBlank(sourceSheet)) {
            errors.add("mes.sourceSheet: requerido cuando mes.enabled=true.");
        } else if (!knownSheets.contains(sourceSheet)) {
            errors.add("mes.sourceSheet: referencia a hoja desconocida '" + sourceSheet
                    + "'. Hojas conocidas: " + knownSheets + ".");
        }

        if (isBlank(config.get("mes.anchorColumn", ""))) {
            errors.add("mes.anchorColumn: requerido cuando mes.enabled=true.");
        }
        Integer srcHeader = parsePositiveInt("mes.sourceHeaderRow", 1);
        if (srcHeader != null && srcHeader < 1) {
            errors.add("mes.sourceHeaderRow debe ser >= 1 (actual: " + srcHeader + ").");
        }

        // Recorrer mes.col.N.* hasta el primer hueco (misma logica que MesSheetBuilder)
        List<MesColInfo> cols = new ArrayList<>();
        int i = 1;
        while (true) {
            String name = config.get("mes.col." + i + ".name", null);
            if (name == null || name.trim().isEmpty()) break;
            String typeRaw = config.get("mes.col." + i + ".type", "EMPTY");
            String type = typeRaw == null ? "" : typeRaw.trim().toUpperCase();
            cols.add(new MesColInfo(i, name.trim(), type));
            i++;
        }

        if (cols.isEmpty()) {
            errors.add("mes.enabled=true pero no se ha definido ninguna columna (mes.col.1.name).");
            return;
        }

        Set<String> colNames = new LinkedHashSet<>();
        for (MesColInfo c : cols) {
            if (!colNames.add(c.name)) {
                errors.add("mes.col." + c.idx + ".name: nombre duplicado '" + c.name
                        + "' (ya existe en otra columna MES).");
            }
        }

        for (MesColInfo c : cols) {
            String prefix = "mes.col." + c.idx + ".";
            if (!VALID_COL_TYPES.contains(c.type)) {
                errors.add(prefix + "type: valor invalido '" + c.type
                        + "'. Valores permitidos: " + VALID_COL_TYPES + ".");
                continue; // no tiene sentido seguir validando esta columna
            }
            switch (c.type) {
                case "COPY":
                    if (isBlank(config.get(prefix + "from", ""))) {
                        errors.add(prefix + "from: requerido para type=COPY ('" + c.name + "').");
                    }
                    break;
                case "SUMIFS":
                    String fromSheet = config.get(prefix + "from", "");
                    if (isBlank(fromSheet)) {
                        errors.add(prefix + "from: requerido para type=SUMIFS ('" + c.name + "').");
                    } else if (!knownSheets.contains(fromSheet)) {
                        errors.add(prefix + "from: referencia a hoja desconocida '" + fromSheet
                                + "' ('" + c.name + "'). Hojas conocidas: " + knownSheets + ".");
                    }
                    if (isBlank(config.get(prefix + "sum", ""))) {
                        errors.add(prefix + "sum: requerido para type=SUMIFS ('" + c.name + "').");
                    }
                    String matchRaw = config.get(prefix + "match", "");
                    if (isBlank(matchRaw)) {
                        errors.add(prefix + "match: requerido para type=SUMIFS ('" + c.name + "').");
                    } else {
                        validateSumifsMatch(prefix, c.name, matchRaw);
                    }
                    break;
                case "FORMULA":
                    String formula = config.get(prefix + "formula", "");
                    if (isBlank(formula)) {
                        errors.add(prefix + "formula: requerido para type=FORMULA ('" + c.name + "').");
                    } else {
                        validateFormulaPlaceholders(prefix, c.name, formula, colNames);
                    }
                    break;
                case "EMPTY":
                default:
                    // nada
            }

            // Validaciones transversales (aplicables a cualquier type)
            validateFill(prefix, c.name);
            validateRedIfNotEqualTo(prefix, c.name, colNames);
        }
    }

    private void validateFill(String prefix, String colName) {
        String fill = config.get(prefix + "fill", null);
        if (fill == null || fill.trim().isEmpty()) return;
        String normalized = fill.trim().toUpperCase();
        if (!VALID_FILL_COLORS.contains(normalized)) {
            errors.add(prefix + "fill: color desconocido '" + fill + "' en '" + colName
                    + "'. Valores permitidos: " + VALID_FILL_COLORS + ".");
        }
    }

    private void validateRedIfNotEqualTo(String prefix, String colName, Set<String> colNames) {
        String ref = config.get(prefix + "redIfNotEqualTo", null);
        if (ref == null || ref.trim().isEmpty()) return;
        String refTrim = ref.trim();
        if (!colNames.contains(refTrim)) {
            errors.add(prefix + "redIfNotEqualTo: columna '" + refTrim + "' no existe en MES (columna '"
                    + colName + "'). Columnas definidas: " + colNames + ".");
        } else if (refTrim.equals(colName)) {
            errors.add(prefix + "redIfNotEqualTo: no puede apuntar a si misma ('" + colName + "').");
        }
    }

    private void validateSumifsMatch(String prefix, String colName, String matchRaw) {
        String[] pairs = matchRaw.split(",");
        int validPairs = 0;
        for (String rawPair : pairs) {
            String pair = rawPair.trim();
            if (pair.isEmpty()) continue;
            int idx = pair.indexOf(':');
            if (idx <= 0 || idx == pair.length() - 1) {
                errors.add(prefix + "match: entrada malformada '" + pair
                        + "' en '" + colName + "' (formato esperado: cabeceraRemota:cabeceraLocal).");
                continue;
            }
            validPairs++;
        }
        if (validPairs == 0) {
            errors.add(prefix + "match: no contiene ninguna pareja valida en '" + colName + "'.");
        }
    }

    private void validateFormulaPlaceholders(String prefix, String colName,
                                             String formula, Set<String> allColNames) {
        Matcher m = PLACEHOLDER_COL.matcher(formula);
        while (m.find()) {
            String ref = m.group(1).trim();
            if (!allColNames.contains(ref)) {
                errors.add(prefix + "formula: placeholder {col:" + ref
                        + "} no coincide con ninguna columna MES (columna '" + colName + "').");
            }
        }
    }

    private void validateDerived(Set<String> knownSheets) {
        List<String> ids = parseCsv(config.get("derived.sheets", ""));
        for (String id : ids) {
            String prefix = "sheet." + id + ".";
            String typeRaw = config.get(prefix + "type", "FORMULAS");
            String type = typeRaw == null ? "" : typeRaw.trim().toUpperCase();
            if (!VALID_DERIVED_TYPES.contains(type)) {
                errors.add(prefix + "type: valor invalido '" + type
                        + "' para hoja derivada '" + id + "'. Valores permitidos: "
                        + VALID_DERIVED_TYPES + ".");
                continue;
            }
            if ("AGGREGATION".equals(type)) {
                String src = config.get(prefix + "sourceSheet", "");
                if (isBlank(src)) {
                    errors.add(prefix + "sourceSheet: requerido para type=AGGREGATION ('" + id + "').");
                } else if (!knownSheets.contains(src)) {
                    errors.add(prefix + "sourceSheet: referencia a hoja desconocida '" + src
                            + "' ('" + id + "'). Hojas conocidas: " + knownSheets + ".");
                }

                String group = config.get(prefix + "groupByColumn", "");
                if (isBlank(group)) {
                    errors.add(prefix + "groupByColumn: requerido para type=AGGREGATION ('" + id + "').");
                } else if (!COLUMN_LETTER.matcher(group.trim()).matches()) {
                    errors.add(prefix + "groupByColumn: debe ser una letra de columna Excel (A, B, AA...). "
                            + "Actual: '" + group + "'.");
                }

                String value = config.get(prefix + "valueColumn", "");
                if (isBlank(value)) {
                    errors.add(prefix + "valueColumn: requerido para type=AGGREGATION ('" + id + "').");
                } else if (!COLUMN_LETTER.matcher(value.trim()).matches()) {
                    errors.add(prefix + "valueColumn: debe ser una letra de columna Excel (A, B, AA...). "
                            + "Actual: '" + value + "'.");
                }

                String agg = config.get(prefix + "aggregation", "SUM");
                String aggNorm = agg == null ? "" : agg.trim().toUpperCase();
                if (!VALID_AGG_FUNCS.contains(aggNorm)) {
                    errors.add(prefix + "aggregation: valor invalido '" + agg + "' para '" + id
                            + "'. Valores permitidos: " + VALID_AGG_FUNCS + ".");
                }

                Integer headerRow = parsePositiveInt(prefix + "headerRow", 1);
                if (headerRow != null && headerRow < 1) {
                    errors.add(prefix + "headerRow debe ser >= 1 (actual: " + headerRow + ").");
                }
                Integer firstDataRow = parsePositiveInt(prefix + "firstDataRow", 2);
                if (firstDataRow != null && firstDataRow < 1) {
                    errors.add(prefix + "firstDataRow debe ser >= 1 (actual: " + firstDataRow + ").");
                }
            }
        }
    }

    private void validateSummary(Set<String> knownSheets) {
        if (!config.getBoolean("summary.enabled", false)) {
            return;
        }

        String sheetName = config.get("summary.sheetName", "Resumen");
        if (isBlank(sheetName)) {
            errors.add("summary.sheetName: valor requerido cuando summary.enabled=true.");
        } else if (knownSheets.contains(sheetName)) {
            errors.add("summary.sheetName: colisiona con otra hoja conocida '"
                    + sheetName + "'. Hojas conocidas: " + knownSheets + ".");
        }

        String sumSheet = config.get("summary.sumSheet", "");
        if (isBlank(sumSheet)) {
            errors.add("summary.sumSheet: valor requerido (normalmente mes.sheetName).");
        } else if (!knownSheets.contains(sumSheet)) {
            errors.add("summary.sumSheet: referencia a hoja desconocida '" + sumSheet
                    + "'. Hojas conocidas: " + knownSheets + ".");
        }

        if (isBlank(config.get("summary.matriculaColumn", ""))) {
            errors.add("summary.matriculaColumn: valor requerido.");
        }

        String valueCols = config.get("summary.valueColumns", "");
        if (isBlank(valueCols)) {
            errors.add("summary.valueColumns: valor requerido (lista separada por comas).");
        } else if (parseCsv(valueCols).isEmpty()) {
            errors.add("summary.valueColumns: lista vacia tras parsear ('" + valueCols + "').");
        }

        Integer maxRow = parsePositiveInt("summary.sumifsMaxRow", 10000);
        if (maxRow != null && maxRow < 1) {
            errors.add("summary.sumifsMaxRow debe ser >= 1 (actual: " + maxRow + ").");
        }

        // v1.8.0 — segunda tabla (matriz Matricula x Responsable)
        validateSummaryByResponsible();
    }

    /**
     * Valida las claves {@code summary.byResponsible.*} (v1.8.0).
     *
     * <p>La segunda tabla es opt-in ({@code summary.byResponsible.enabled}
     * default false). Si esta habilitada sin que {@code summary.enabled}
     * tambien lo este, se trata como error porque la segunda tabla se
     * ancla a la hoja Resumen y no tiene sentido sola.</p>
     */
    private void validateSummaryByResponsible() {
        boolean byRespEnabled = config.getBoolean("summary.byResponsible.enabled", false);
        if (!byRespEnabled) {
            return;
        }

        if (!config.getBoolean("summary.enabled", false)) {
            errors.add("summary.byResponsible.enabled=true requiere summary.enabled=true "
                    + "(la segunda tabla se renderiza dentro de la hoja Resumen).");
            // Seguimos validando el resto para dar feedback completo al usuario.
        }

        if (isBlank(config.get("summary.byResponsible.column", ""))) {
            errors.add("summary.byResponsible.column: valor requerido cuando "
                    + "summary.byResponsible.enabled=true.");
        }

        if (isBlank(config.get("summary.byResponsible.valueColumn", ""))) {
            errors.add("summary.byResponsible.valueColumn: valor requerido cuando "
                    + "summary.byResponsible.enabled=true.");
        }

        if (isBlank(config.get("summary.byResponsible.title", ""))) {
            errors.add("summary.byResponsible.title: valor requerido cuando "
                    + "summary.byResponsible.enabled=true.");
        }

        // gapRows admite 0 (tablas pegadas); por eso validamos distinto de parsePositiveInt.
        String gapRaw = config.get("summary.byResponsible.gapRows", null);
        if (gapRaw != null && !gapRaw.trim().isEmpty()) {
            try {
                int gap = Integer.parseInt(gapRaw.trim());
                if (gap < 0) {
                    errors.add("summary.byResponsible.gapRows debe ser >= 0 (actual: "
                            + gap + ").");
                }
            } catch (NumberFormatException e) {
                errors.add("summary.byResponsible.gapRows: valor no numerico '"
                        + gapRaw + "'.");
            }
        }
    }

    /**
     * Valida las claves de {@code mes.orphans.*} (v1.7.0). Solo se activa si
     * {@code mes.orphans.enabled=true}. Comprueba:
     * <ul>
     *   <li>La hoja declarada en {@code mes.orphans.sourceSheet} existe.</li>
     *   <li>Los nombres de columna MES ({@code colPeticion}, {@code colMatricula},
     *       {@code colJira}) referencian columnas definidas en {@code mes.col.N.name}.</li>
     *   <li>{@code mes.enabled=true} (no tiene sentido si MES esta apagada).</li>
     * </ul>
     * Las claves {@code matchComponent}, {@code matchMatricula}, {@code sumColumn}
     * no se validan aqui porque dependen del contenido del Excel en runtime;
     * el builder emite warnings si no las encuentra.
     */
    private void validateOrphans(Set<String> knownSheets) {
        if (!config.getBoolean("mes.orphans.enabled", false)) {
            return;
        }
        if (!config.getBoolean("mes.enabled", false)) {
            errors.add("mes.orphans.enabled=true pero mes.enabled=false; "
                    + "los huerfanos solo se emiten dentro de la hoja MES.");
            return;
        }

        String sourceSheet = config.get("mes.orphans.sourceSheet", "");
        if (isBlank(sourceSheet)) {
            errors.add("mes.orphans.sourceSheet: valor requerido cuando mes.orphans.enabled=true.");
        } else if (!knownSheets.contains(sourceSheet)) {
            errors.add("mes.orphans.sourceSheet: referencia a hoja desconocida '"
                    + sourceSheet + "'. Hojas conocidas: " + knownSheets + ".");
        }

        // Recolectar los nombres de columna MES para validar las claves col*
        Set<String> mesColNames = new LinkedHashSet<>();
        int i = 1;
        while (true) {
            String name = config.get("mes.col." + i + ".name", null);
            if (name == null || name.trim().isEmpty()) break;
            mesColNames.add(name.trim());
            i++;
        }

        String[][] colsToCheck = {
                {"mes.orphans.colPeticion",  "Petición"},
                {"mes.orphans.colMatricula", "Matrícula"},
                {"mes.orphans.colJira",      "Jira"},
        };
        for (String[] pair : colsToCheck) {
            String key = pair[0];
            String defaultVal = pair[1];
            String value = config.get(key, defaultVal);
            if (isBlank(value)) {
                errors.add(key + ": valor requerido cuando mes.orphans.enabled=true.");
            } else if (!mesColNames.contains(value)) {
                errors.add(key + ": '" + value + "' no coincide con ningun mes.col.N.name. "
                        + "Columnas MES definidas: " + mesColNames + ".");
            }
        }
    }

    // ==================================================================
    //  Helpers
    // ==================================================================

    private static List<String> parseCsv(String raw) {
        List<String> out = new ArrayList<>();
        if (raw == null) return out;
        for (String part : raw.split(",")) {
            String t = part.trim();
            if (!t.isEmpty()) out.add(t);
        }
        return out;
    }

    private boolean hasAnyKeyStartingWith(String prefix) {
        for (String key : config.getRawProperties().stringPropertyNames()) {
            if (key.startsWith(prefix)) return true;
        }
        return false;
    }

    /**
     * Parsea un entero. Si la clave no existe o esta vacia, devuelve defaultVal
     * (o null si defaultVal es null y la clave esta ausente). Si el valor
     * existe pero no es numerico, registra un error y devuelve defaultVal.
     */
    private Integer parsePositiveInt(String key, Integer defaultVal) {
        String raw = config.get(key, null);
        if (raw == null || raw.trim().isEmpty()) return defaultVal;
        try {
            return Integer.parseInt(raw.trim());
        } catch (NumberFormatException e) {
            errors.add(key + ": valor no numerico '" + raw + "'.");
            return defaultVal;
        }
    }

    private static boolean isBlank(String s) {
        return s == null || s.trim().isEmpty();
    }

    private static final class MesColInfo {
        final int idx;
        final String name;
        final String type;
        MesColInfo(int idx, String name, String type) {
            this.idx = idx;
            this.name = name;
            this.type = type;
        }
    }
}

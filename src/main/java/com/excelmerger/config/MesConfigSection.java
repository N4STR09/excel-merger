package com.excelmerger.config;

import com.excelmerger.ConfigLoader;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static com.excelmerger.config.ValidationHelpers.isBlank;
import static com.excelmerger.config.ValidationHelpers.parsePositiveInt;
import static com.excelmerger.config.ValidationMessages.MSG_SUFFIX_ALLOWED_VALUES;
import static com.excelmerger.config.ValidationMessages.MSG_SUFFIX_KNOWN_SHEETS;

/**
 * Valida el bloque {@code mes.*} del config: hoja MES, columnas MES (5 tipos),
 * placeholders de fórmula, fill y redIfNotEqualTo.
 *
 * <p>Solo actúa si {@code mes.enabled=true}. Acumula errores en la lista
 * compartida que recibe por constructor; no devuelve nada.</p>
 *
 * <p>Extraída de {@code ConfigValidator} durante el refactor de la
 * Sesión F (v2.5.0). Package-private.</p>
 */
final class MesConfigSection {

    private static final Set<String> VALID_COL_TYPES = new LinkedHashSet<>(
            Arrays.asList("COPY", "SUMIFS", "FORMULA", "FORMULA_PLUS_SUMIFS", "EMPTY"));
    private static final Set<String> VALID_FILL_COLORS = new LinkedHashSet<>(
            Arrays.asList("LIGHT_GREEN", "MEDIUM_GREEN", "LIGHT_BLUE",
                    "LIGHT_YELLOW", "LIGHT_RED", "LIGHT_LAVENDER"));

    private static final Pattern PLACEHOLDER_COL = Pattern.compile("\\{col:([^}]+)\\}");

    /** Prefijo de las propiedades de configuración de columnas MES. */
    private static final String PROP_PREFIX_MES_COL = "mes.col.";

    private final ConfigLoader config;
    private final List<String> errors;

    MesConfigSection(ConfigLoader config, List<String> errors) {
        this.config = config;
        this.errors = errors;
    }

    void validate(Set<String> knownSheets) {
        if (!config.getBoolean("mes.enabled", false)) {
            return;
        }
        String sourceSheet = config.get("mes.sourceSheet", "");
        if (isBlank(sourceSheet)) {
            errors.add("mes.sourceSheet: requerido cuando mes.enabled=true.");
        } else if (!knownSheets.contains(sourceSheet)) {
            errors.add("mes.sourceSheet: referencia a hoja desconocida '" + sourceSheet
                    + MSG_SUFFIX_KNOWN_SHEETS + knownSheets + ".");
        }

        if (isBlank(config.get("mes.anchorColumn", ""))) {
            errors.add("mes.anchorColumn: requerido cuando mes.enabled=true.");
        }
        Integer srcHeader = parsePositiveInt(config, "mes.sourceHeaderRow", 1, errors);
        if (srcHeader != null && srcHeader < 1) {
            errors.add("mes.sourceHeaderRow debe ser >= 1 (actual: " + srcHeader + ").");
        }

        // Recorrer mes.col.N.* hasta el primer hueco (misma logica que MesSheetBuilder)
        List<MesColInfo> cols = new ArrayList<>();
        int i = 1;
        while (true) {
            String name = config.get(PROP_PREFIX_MES_COL + i + ".name", null);
            if (name == null || name.trim().isEmpty()) break;
            String typeRaw = config.get(PROP_PREFIX_MES_COL + i + ".type", "EMPTY");
            String type = typeRaw == null ? "" : typeRaw.trim().toUpperCase(Locale.ROOT);
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
                errors.add(PROP_PREFIX_MES_COL + c.idx + ".name: nombre duplicado '" + c.name
                        + "' (ya existe en otra columna MES).");
            }
        }

        for (MesColInfo c : cols) {
            String prefix = PROP_PREFIX_MES_COL + c.idx + ".";
            if (!VALID_COL_TYPES.contains(c.type)) {
                errors.add(prefix + "type: valor invalido '" + c.type
                        + MSG_SUFFIX_ALLOWED_VALUES + VALID_COL_TYPES + ".");
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
                case "FORMULA_PLUS_SUMIFS":
                    String baseFormula = config.get(prefix + "baseFormula", "");
                    if (isBlank(baseFormula)) {
                        errors.add(prefix + "baseFormula: requerido para type=FORMULA_PLUS_SUMIFS ('"
                                + c.name + "').");
                    } else {
                        validateFormulaPlaceholders(prefix, c.name, baseFormula, colNames);
                    }
                    String fpsFromSheet = config.get(prefix + "from", "");
                    if (isBlank(fpsFromSheet)) {
                        errors.add(prefix + "from: requerido para type=FORMULA_PLUS_SUMIFS ('"
                                + c.name + "').");
                    } else if (!knownSheets.contains(fpsFromSheet)) {
                        errors.add(prefix + "from: referencia a hoja desconocida '" + fpsFromSheet
                                + "' ('" + c.name + "'). Hojas conocidas: " + knownSheets + ".");
                    }
                    if (isBlank(config.get(prefix + "sum", ""))) {
                        errors.add(prefix + "sum: requerido para type=FORMULA_PLUS_SUMIFS ('" + c.name + "').");
                    }
                    String fpsMatchRaw = config.get(prefix + "match", "");
                    if (isBlank(fpsMatchRaw)) {
                        errors.add(prefix + "match: requerido para type=FORMULA_PLUS_SUMIFS ('"
                                + c.name + "').");
                    } else {
                        validateSumifsMatch(prefix, c.name, fpsMatchRaw);
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
        String normalized = fill.trim().toUpperCase(Locale.ROOT);
        if (!VALID_FILL_COLORS.contains(normalized)) {
            errors.add(prefix + "fill: color desconocido '" + fill + "' en '" + colName
                    + MSG_SUFFIX_ALLOWED_VALUES + VALID_FILL_COLORS + ".");
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

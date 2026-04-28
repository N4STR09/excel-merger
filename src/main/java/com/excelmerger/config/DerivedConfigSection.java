package com.excelmerger.config;

import com.excelmerger.ConfigLoader;

import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.regex.Pattern;

import static com.excelmerger.config.ValidationHelpers.isBlank;
import static com.excelmerger.config.ValidationHelpers.parseCsv;
import static com.excelmerger.config.ValidationHelpers.parsePositiveInt;
import static com.excelmerger.config.ValidationMessages.MSG_SUFFIX_ALLOWED_VALUES;

/**
 * Valida el bloque {@code derived.sheets=} y las propiedades
 * {@code sheet.<id>.*} asociadas. Soporta dos tipos: {@code FORMULAS}
 * (sin validación de campos en config; el contenido se interpreta en
 * runtime) y {@code AGGREGATION} (sourceSheet, groupByColumn, valueColumn,
 * aggregation, headerRow, firstDataRow).
 *
 * <p>Extraída de {@code ConfigValidator} durante el refactor de la
 * Sesión F (v2.5.0). Package-private.</p>
 */
final class DerivedConfigSection {

    private static final Set<String> VALID_DERIVED_TYPES = new LinkedHashSet<>(
            Arrays.asList("FORMULAS", "AGGREGATION"));
    private static final Set<String> VALID_AGG_FUNCS = new LinkedHashSet<>(
            Arrays.asList("SUM", "AVG", "COUNT", "MIN", "MAX"));

    private static final Pattern COLUMN_LETTER = Pattern.compile("[A-Za-z]+");

    private final ConfigLoader config;
    private final List<String> errors;

    DerivedConfigSection(ConfigLoader config, List<String> errors) {
        this.config = config;
        this.errors = errors;
    }

    void validate(Set<String> knownSheets) {
        List<String> ids = parseCsv(config.get("derived.sheets", ""));
        for (String id : ids) {
            String prefix = "sheet." + id + ".";
            String typeRaw = config.get(prefix + "type", "FORMULAS");
            String type = typeRaw == null ? "" : typeRaw.trim().toUpperCase(Locale.ROOT);
            if (!VALID_DERIVED_TYPES.contains(type)) {
                errors.add(prefix + "type: valor invalido '" + type
                        + "' para hoja derivada '" + id + MSG_SUFFIX_ALLOWED_VALUES
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
                String aggNorm = agg == null ? "" : agg.trim().toUpperCase(Locale.ROOT);
                if (!VALID_AGG_FUNCS.contains(aggNorm)) {
                    errors.add(prefix + "aggregation: valor invalido '" + agg + "' para '" + id
                            + MSG_SUFFIX_ALLOWED_VALUES + VALID_AGG_FUNCS + ".");
                }

                Integer headerRow = parsePositiveInt(config, prefix + "headerRow", 1, errors);
                if (headerRow != null && headerRow < 1) {
                    errors.add(prefix + "headerRow debe ser >= 1 (actual: " + headerRow + ").");
                }
                Integer firstDataRow = parsePositiveInt(config, prefix + "firstDataRow", 2, errors);
                if (firstDataRow != null && firstDataRow < 1) {
                    errors.add(prefix + "firstDataRow debe ser >= 1 (actual: " + firstDataRow + ").");
                }
            }
        }
    }
}

package com.excelmerger.config;

import com.excelmerger.ConfigLoader;

import java.util.List;
import java.util.Set;

import static com.excelmerger.config.ValidationHelpers.isBlank;
import static com.excelmerger.config.ValidationHelpers.parseCsv;
import static com.excelmerger.config.ValidationHelpers.parsePositiveInt;
import static com.excelmerger.config.ValidationMessages.MSG_SUFFIX_KNOWN_SHEETS;

/**
 * Valida los bloques {@code summary.*} y {@code summary.byResponsible.*}
 * del config.
 *
 * <p>El orden de invocación interno (summary → byResponsible) preserva
 * el orden de aparición de errores del validador original. Las tablas
 * {@code responsables.tables.*} viven en su propia
 * {@link ResponsablesTablesConfigSection} para preservar el orden global
 * de errores entre orphans y responsables.tables.</p>
 *
 * <p>Extraída de {@code ConfigValidator} durante el refactor de la
 * Sesión F (v2.5.0). Package-private.</p>
 */
final class SummaryConfigSection {

    private final ConfigLoader config;
    private final List<String> errors;

    SummaryConfigSection(ConfigLoader config, List<String> errors) {
        this.config = config;
        this.errors = errors;
    }

    void validate(Set<String> knownSheets) {
        if (!config.getBoolean("summary.enabled", false)) {
            return;
        }

        String sheetName = config.get("summary.sheetName", "Resumen");
        if (isBlank(sheetName)) {
            errors.add("summary.sheetName: valor requerido cuando summary.enabled=true.");
        } else if (knownSheets.contains(sheetName)) {
            errors.add("summary.sheetName: colisiona con otra hoja conocida '"
                    + sheetName + MSG_SUFFIX_KNOWN_SHEETS + knownSheets + ".");
        }

        String sumSheet = config.get("summary.sumSheet", "");
        if (isBlank(sumSheet)) {
            errors.add("summary.sumSheet: valor requerido (normalmente mes.sheetName).");
        } else if (!knownSheets.contains(sumSheet)) {
            errors.add("summary.sumSheet: referencia a hoja desconocida '" + sumSheet
                    + MSG_SUFFIX_KNOWN_SHEETS + knownSheets + ".");
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

        Integer maxRow = parsePositiveInt(config, "summary.sumifsMaxRow", 10000, errors);
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
}

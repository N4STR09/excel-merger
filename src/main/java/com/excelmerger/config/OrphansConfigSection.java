package com.excelmerger.config;

import com.excelmerger.ConfigLoader;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import static com.excelmerger.config.ValidationHelpers.isBlank;
import static com.excelmerger.config.ValidationMessages.MSG_SUFFIX_KNOWN_SHEETS;

/**
 * Valida las claves de {@code mes.orphans.*} (v1.7.0). Solo se activa si
 * {@code mes.orphans.enabled=true}. Comprueba:
 * <ul>
 *   <li>La hoja declarada en {@code mes.orphans.sourceSheet} existe.</li>
 *   <li>Los nombres de columna MES ({@code colPeticion}, {@code colMatricula},
 *       {@code colJira}) referencian columnas definidas en {@code mes.col.N.name}.</li>
 *   <li>{@code mes.enabled=true} (no tiene sentido si MES esta apagada).</li>
 * </ul>
 *
 * <p>Las claves {@code matchComponent}, {@code matchMatricula}, {@code sumColumn}
 * no se validan aqui porque dependen del contenido del Excel en runtime;
 * el builder emite warnings si no las encuentra.</p>
 *
 * <p>Extraída de {@code ConfigValidator} durante el refactor de la
 * Sesión F (v2.5.0). Package-private.</p>
 */
final class OrphansConfigSection {

    /** Prefijo de las propiedades de configuración de columnas MES. */
    private static final String PROP_PREFIX_MES_COL = "mes.col.";

    private final ConfigLoader config;
    private final List<String> errors;

    OrphansConfigSection(ConfigLoader config, List<String> errors) {
        this.config = config;
        this.errors = errors;
    }

    void validate(Set<String> knownSheets) {
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
                    + sourceSheet + MSG_SUFFIX_KNOWN_SHEETS + knownSheets + ".");
        }

        // Recolectar los nombres de columna MES para validar las claves col*
        Set<String> mesColNames = new LinkedHashSet<>();
        int i = 1;
        while (true) {
            String name = config.get(PROP_PREFIX_MES_COL + i + ".name", null);
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
}

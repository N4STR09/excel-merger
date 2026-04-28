package com.excelmerger.config;

import com.excelmerger.ConfigLoader;

import java.util.List;

import static com.excelmerger.config.ValidationHelpers.isBlank;

/**
 * v2.4.0 — Valida las claves {@code responsables.tables.*}.
 *
 * <p>Las tablas pivot por responsable son opt-out: si la clave
 * {@code responsables.tables.enabled} no está, se asume {@code true}.
 * No tienen precondiciones cruzadas (las hojas de responsable se
 * generan según {@code OutputMode} y la existencia de Resultado se
 * valida en runtime).</p>
 *
 * <p>Validaciones:</p>
 * <ul>
 *   <li>{@code responsables.tables.gapRows}: si presente, entero >= 0.</li>
 *   <li>{@code responsables.tables.jiraTitle} y {@code .realTitle}: si
 *       están presentes, no pueden ser blancos.</li>
 * </ul>
 *
 * <p>No se valida que las columnas Petición/Matrícula/Res. Tecnico/Jira/REAL
 * existan en Resultado: ese chequeo es runtime y emite warning si faltan
 * (mismo patrón que el resto de los builders).</p>
 *
 * <p>Extraída de {@code ConfigValidator} durante el refactor de la
 * Sesión F (v2.5.0). Vive como sección independiente (no agrupada con
 * summary) para preservar el orden global de errores orphans →
 * responsables.tables del validador original. Package-private.</p>
 */
final class ResponsablesTablesConfigSection {

    private final ConfigLoader config;
    private final List<String> errors;

    ResponsablesTablesConfigSection(ConfigLoader config, List<String> errors) {
        this.config = config;
        this.errors = errors;
    }

    void validate() {
        // gapRows admite 0; validar como entero >= 0 si está presente.
        String gapRaw = config.get("responsables.tables.gapRows", null);
        if (gapRaw != null && !gapRaw.trim().isEmpty()) {
            try {
                int gap = Integer.parseInt(gapRaw.trim());
                if (gap < 0) {
                    errors.add("responsables.tables.gapRows debe ser >= 0 (actual: "
                            + gap + ").");
                }
            } catch (NumberFormatException e) {
                errors.add("responsables.tables.gapRows: valor no numerico '"
                        + gapRaw + "'.");
            }
        }

        // Títulos: si están explícitos, no pueden ser solo espacios.
        // Si no están, se usan los defaults del builder, así que no validamos.
        String jiraTitle = config.get("responsables.tables.jiraTitle", null);
        if (jiraTitle != null && isBlank(jiraTitle)) {
            errors.add("responsables.tables.jiraTitle: no puede estar vacio.");
        }
        String realTitle = config.get("responsables.tables.realTitle", null);
        if (realTitle != null && isBlank(realTitle)) {
            errors.add("responsables.tables.realTitle: no puede estar vacio.");
        }
    }
}

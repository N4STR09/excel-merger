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
 *   <li>{@code responsables.tables.jiraTitle} y {@code .pdclTitle}: si
 *       están presentes, no pueden ser blancos.</li>
 *   <li>v2.7.0: claves obsoletas {@code responsables.tables.facturarTitle}
 *       (v2.4.0..v2.6.0) y {@code responsables.tables.realTitle}
 *       (≤v2.5.1) son rechazadas con error explícito y mensaje de migración.
 *       Esto evita degradados silenciosos cuando un usuario actualiza el
 *       binario sin migrar el config.</li>
 * </ul>
 *
 * <p>No se valida que las columnas Petición/Matrícula/Res. Tecnico/Jira/PDCL
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
        // v2.7.0 (Modif 3): clave nueva pdclTitle (rename de facturarTitle).
        String pdclTitle = config.get("responsables.tables.pdclTitle", null);
        if (pdclTitle != null && isBlank(pdclTitle)) {
            errors.add("responsables.tables.pdclTitle: no puede estar vacio.");
        }

        // v2.7.0 (Modif 3): rechazo explicito de claves obsoletas.
        // Sin alias retrocompat — mantenerlos hace que el output salga con
        // el titulo viejo en silencio, lo opuesto a lo que el usuario pidio
        // al renombrar la fuente Facturar -> PDCL.
        if (config.has("responsables.tables.facturarTitle")) {
            errors.add("responsables.tables.facturarTitle: clave obsoleta en v2.7.0. "
                    + "Renombrala a 'responsables.tables.pdclTitle' (la pivot ahora lee "
                    + "la columna PDCL en lugar de Facturar — Modif 3).");
        }
        if (config.has("responsables.tables.realTitle")) {
            errors.add("responsables.tables.realTitle: clave obsoleta desde v2.6.0. "
                    + "Renombrala a 'responsables.tables.pdclTitle' en v2.7.0 "
                    + "(la columna paso de REAL a Facturar en v2.6.0 y a PDCL en v2.7.0).");
        }
    }
}

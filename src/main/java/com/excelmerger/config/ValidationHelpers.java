package com.excelmerger.config;

import com.excelmerger.ConfigLoader;

import java.util.ArrayList;
import java.util.List;

/**
 * Helpers de validación compartidos por las distintas {@code *ConfigSection}.
 * Extraídos del antiguo {@code ConfigValidator} monolítico durante el
 * refactor estructural de la Sesión F (v2.5.0).
 *
 * <p>Métodos estáticos sin estado. {@link #parsePositiveInt} es la única
 * excepción al principio "función pura": muta la lista {@code errors}
 * recibida por parámetro cuando el valor no es numérico, para preservar
 * exactamente el patrón "acumular errores en una sola lista" del validador
 * original.</p>
 */
final class ValidationHelpers {

    private ValidationHelpers() {
        // utility holder
    }

    /**
     * Parsea una lista CSV: separa por coma, hace trim y descarta entradas
     * vacías. Devuelve lista vacía si la entrada es {@code null}.
     */
    static List<String> parseCsv(String raw) {
        List<String> out = new ArrayList<>();
        if (raw == null) {
            return out;
        }
        for (String part : raw.split(",")) {
            String t = part.trim();
            if (!t.isEmpty()) {
                out.add(t);
            }
        }
        return out;
    }

    /**
     * Devuelve true si la cadena es {@code null} o solo espacios.
     */
    static boolean isBlank(String s) {
        return s == null || s.trim().isEmpty();
    }

    /**
     * Parsea un entero del config. Si la clave no existe o está vacía,
     * devuelve {@code defaultVal} (puede ser null si la clave es opcional
     * y no hay default razonable). Si el valor existe pero no es numérico,
     * registra un error en {@code errors} y devuelve {@code defaultVal}.
     */
    static Integer parsePositiveInt(ConfigLoader config, String key,
                                    Integer defaultVal, List<String> errors) {
        String raw = config.get(key, null);
        if (raw == null || raw.trim().isEmpty()) {
            return defaultVal;
        }
        try {
            return Integer.parseInt(raw.trim());
        } catch (NumberFormatException e) {
            errors.add(key + ": valor no numerico '" + raw + "'.");
            return defaultVal;
        }
    }
}

package com.excelmerger.config;

/**
 * Sufijos comunes de mensajes de error usados por múltiples
 * {@code *ConfigSection}. Extraídos del antiguo {@code ConfigValidator}
 * monolítico durante el refactor estructural de la Sesión F (v2.5.0).
 *
 * <p>Package-private: solo los validadores internos del paquete los usan.</p>
 */
final class ValidationMessages {

    /** Sufijo común de los mensajes que listan las hojas conocidas. */
    static final String MSG_SUFFIX_KNOWN_SHEETS = "'. Hojas conocidas: ";

    /** Sufijo común de los mensajes que listan valores válidos. */
    static final String MSG_SUFFIX_ALLOWED_VALUES = "'. Valores permitidos: ";

    private ValidationMessages() {
        // utility holder
    }
}

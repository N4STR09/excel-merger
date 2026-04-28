package com.excelmerger.config;

import com.excelmerger.ConfigLoader;

import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import static com.excelmerger.config.ValidationHelpers.isBlank;
import static com.excelmerger.config.ValidationMessages.MSG_SUFFIX_ALLOWED_VALUES;

/**
 * Valida las claves "globales" del config: {@code input.directory},
 * {@code output.file}, {@code merge.mode} y {@code output.mode}.
 *
 * <p>Extraída de {@code ConfigValidator} durante el refactor de la
 * Sesión F (v2.5.0). Los tres métodos públicos se invocan en el
 * orden histórico {@code validateInputOutput} → {@code validateMergeMode}
 * → {@code validateOutputMode} desde el orquestador.</p>
 *
 * <p>Package-private: solo {@link ConfigValidator} la instancia.</p>
 */
final class IoConfigSection {

    private static final Set<String> VALID_MERGE_MODES = new LinkedHashSet<>(
            Arrays.asList("SHEETS_SEPARATE", "APPEND_ROWS"));

    /** v2.3.0: valores permitidos para {@code output.mode}. Case-sensitive, minusculas. */
    private static final Set<String> VALID_OUTPUT_MODES = new LinkedHashSet<>(
            Arrays.asList("cierre", "responsables", "completo"));

    private final ConfigLoader config;
    private final List<String> errors;

    IoConfigSection(ConfigLoader config, List<String> errors) {
        this.config = config;
        this.errors = errors;
    }

    void validateInputOutput() {
        if (isBlank(config.get("input.directory", ""))) {
            errors.add("input.directory: propiedad requerida.");
        }
        if (isBlank(config.get("output.file", ""))) {
            errors.add("output.file: propiedad requerida.");
        }
    }

    String validateMergeMode() {
        String mode = config.get("merge.mode", "SHEETS_SEPARATE").toUpperCase(Locale.ROOT);
        if (!VALID_MERGE_MODES.contains(mode)) {
            errors.add("merge.mode: valor invalido '" + mode + MSG_SUFFIX_ALLOWED_VALUES
                    + VALID_MERGE_MODES);
            return "SHEETS_SEPARATE";
        }
        return mode;
    }

    /**
     * v2.3.0: valida la clave {@code output.mode}. Estricto, case-sensitive,
     * solo admite los valores literales {@code cierre}, {@code responsables},
     * {@code completo}. Si la clave esta ausente o vacia, no es error
     * (el default operativo es {@code cierre}, aplicado en {@code ExcelMerger}).
     */
    void validateOutputMode() {
        String raw = config.get("output.mode", "");
        if (raw.isEmpty()) {
            return;  // ausente o vacio -> default CIERRE, sin error
        }
        if (!VALID_OUTPUT_MODES.contains(raw)) {
            errors.add("output.mode: valor invalido '" + raw + MSG_SUFFIX_ALLOWED_VALUES
                    + VALID_OUTPUT_MODES + ".");
        }
    }
}

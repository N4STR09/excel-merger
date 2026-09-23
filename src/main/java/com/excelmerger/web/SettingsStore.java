package com.excelmerger.web;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Locale;
import java.util.Properties;

/**
 * Ajustes de salida configurables en runtime desde la interfaz web
 * (v4.1.0): tres claves que determinan que hojas genera la fusion.
 *
 * <p>Se persisten en {@value #FILE_NAME} del directorio de trabajo, junto
 * al {@code web.properties}: en la imagen portatil viajan con la carpeta
 * y en ejecuciones desde codigo fuente se ignoran en git. El fichero es
 * opcional y resiliente: si no existe, esta corrupto o contiene valores
 * invalidos, se parte de los valores por defecto sin que el arranque
 * falle.</p>
 *
 * <p>Las claves son las mismas que el motor lee de {@code config.properties}
 * y se aplican como overrides sobre la configuracion base
 * ({@link com.excelmerger.ConfigLoader#ConfigLoader(java.util.Properties)});
 * el fichero de configuracion no se toca. Por defecto se restaura el
 * comportamiento completo de v3.1.0 (la salida que tenia el usuario antes
 * de v4.0.0): modo {@code completo} + hoja Resumen + matriz por
 * responsable activas, sin necesidad de editar configuraciones a mano.</p>
 */
final class SettingsStore {

    /** Fichero de ajustes en el directorio de trabajo. */
    static final String FILE_NAME = "web-settings.properties";

    /** Claves motor que gestiona la interfaz (mismas que en config.properties). */
    static final String KEY_MODE = "output.mode";
    static final String KEY_SUMMARY = "summary.enabled";
    static final String KEY_BY_RESPONSIBLE = "summary.byResponsible.enabled";

    /** Valores validos de {@code output.mode} (case-sensitive, en minusculas). */
    static final String MODE_CIERRE = "cierre";
    static final String MODE_RESPONSABLES = "responsables";
    static final String MODE_COMPLETO = "completo";

    private static final Logger log = LoggerFactory.getLogger(SettingsStore.class);

    private static final String FILE_HEADER =
            "# Ajustes de salida de la interfaz web (se generan desde la pagina; no editar).";
    private static final String NL = "\n";
    private static final String JSON_MODE = "mode";
    private static final String JSON_SUMMARY = "summaryEnabled";
    private static final String JSON_BY_RESPONSIBLE = "byResponsibleEnabled";

    /** Default v4.1.0: restaura la salida completa de v3.1.0. */
    private static final String DEFAULT_MODE = MODE_COMPLETO;
    private static final boolean DEFAULT_SUMMARY = true;
    private static final boolean DEFAULT_BY_RESPONSIBLE = true;

    private final Path file;
    private String mode;
    private boolean summaryEnabled;
    private boolean byResponsibleEnabled;

    private SettingsStore(Path file, String mode, boolean summaryEnabled,
                          boolean byResponsibleEnabled) {
        this.file = file;
        this.mode = mode;
        this.summaryEnabled = summaryEnabled;
        this.byResponsibleEnabled = byResponsibleEnabled;
    }

    /**
     * Carga los ajustes del fichero por defecto del directorio de trabajo.
     *
     * @return ajustes validos siempre; ante cualquier problema, los
     *         valores por defecto.
     */
    static SettingsStore load() {
        return load(Paths.get(FILE_NAME));
    }

    /**
     * Carga los ajustes de un fichero concreto (usado por los tests).
     * Mismo contrato resiliente que {@link #load()}.
     *
     * @param file fichero de ajustes a leer.
     * @return ajustes validos siempre.
     */
    static SettingsStore load(Path file) {
        Properties props = new Properties();
        if (Files.isRegularFile(file)) {
            try (Reader reader = Files.newBufferedReader(file, StandardCharsets.UTF_8)) {
                props.load(reader);
            } catch (IOException e) {
                log.warn("No se pudieron leer los ajustes de salida ({}): {}; se usan los "
                        + "valores por defecto.", file, e.getMessage());
            }
        }
        String mode = normalizeMode(props.getProperty(KEY_MODE));
        boolean summary = parseBool(props.getProperty(KEY_SUMMARY), DEFAULT_SUMMARY);
        boolean byResponsible = parseBool(props.getProperty(KEY_BY_RESPONSIBLE),
                DEFAULT_BY_RESPONSIBLE);
        if (byResponsible && !summary) {
            summary = true;
        }
        return new SettingsStore(file.toAbsolutePath().normalize(), mode, summary, byResponsible);
    }

    /** Modo de salida en vigor: {@code cierre}, {@code responsables} o {@code completo}. */
    String mode() {
        return mode;
    }

    /** {@code true} si se genera la hoja Resumen. */
    boolean summaryEnabled() {
        return summaryEnabled;
    }

    /** {@code true} si el Resumen incluye la tabla por responsable. */
    boolean byResponsibleEnabled() {
        return byResponsibleEnabled;
    }

    /** Modos de salida ofertados por la interfaz, en orden de presentacion. */
    static List<String> modes() {
        return List.of(MODE_CIERRE, MODE_RESPONSABLES, MODE_COMPLETO);
    }

    /** {@code true} si {@code value} es un modo de salida valido. */
    static boolean isValidMode(String value) {
        return MODE_CIERRE.equals(value) || MODE_RESPONSABLES.equals(value)
                || MODE_COMPLETO.equals(value);
    }

    /**
     * Aplica y persiste un cambio de ajustes. Los campos {@code null}
     * mantienen el valor actual; la matriz por responsable obliga a activar
     * la hoja Resumen (regla del validador del motor, evitada en origen).
     *
     * @param mode               nuevo modo ({@code null} = sin cambio).
     * @param summaryEnabled     si se genera el Resumen ({@code null} = sin cambio).
     * @param byResponsibleEnabled si el Resumen incluye la tabla por responsable
     *                             ({@code null} = sin cambio).
     * @throws IllegalArgumentException si el modo no es valido.
     */
    void update(String mode, Boolean summaryEnabled, Boolean byResponsibleEnabled) {
        String nextMode = mode == null ? this.mode : requireValidMode(mode);
        boolean nextSummary = summaryEnabled == null ? this.summaryEnabled : summaryEnabled;
        boolean nextByResponsible = byResponsibleEnabled == null
                ? this.byResponsibleEnabled : byResponsibleEnabled;
        if (nextByResponsible && !nextSummary) {
            nextSummary = true;
        }
        SettingsStore updated = new SettingsStore(file, nextMode, nextSummary, nextByResponsible);
        updated.save();
        this.mode = updated.mode;
        this.summaryEnabled = updated.summaryEnabled;
        this.byResponsibleEnabled = updated.byResponsibleEnabled;
    }

    /**
     * Overrides para {@link com.excelmerger.ConfigLoader}: las tres claves
     * listas para aplicarse sobre la configuracion base en cada fusion.
     *
     * @return {@link Properties} con las claves motor en vigor.
     */
    Properties toOverrides() {
        Properties overrides = new Properties();
        overrides.setProperty(KEY_MODE, mode);
        overrides.setProperty(KEY_SUMMARY, Boolean.toString(summaryEnabled));
        overrides.setProperty(KEY_BY_RESPONSIBLE, Boolean.toString(byResponsibleEnabled));
        return overrides;
    }

    /** Serializa los ajustes en vigor como JSON para {@code GET/POST /api/settings}. */
    String toJson() {
        return "{\"" + JSON_MODE + "\":" + Json.quote(mode)
                + ",\"" + JSON_SUMMARY + "\":" + summaryEnabled
                + ",\"" + JSON_BY_RESPONSIBLE + "\":" + byResponsibleEnabled + "}";
    }

    /**
     * Persiste los ajustes en el fichero. Best effort, como
     * {@code WebSession}: un fallo de escritura se loguea y no tumba la
     * interfaz (la siguiente ejecucion usaria estos mismos valores en
     * memoria).
     */
    private void save() {
        StringBuilder sb = new StringBuilder(FILE_HEADER).append(NL);
        sb.append(KEY_MODE).append('=').append(mode).append(NL);
        sb.append(KEY_SUMMARY).append('=').append(summaryEnabled).append(NL);
        sb.append(KEY_BY_RESPONSIBLE).append('=').append(byResponsibleEnabled).append(NL);
        try {
            Files.writeString(file, sb, StandardCharsets.UTF_8);
        } catch (IOException e) {
            log.warn("No se pudieron guardar los ajustes de salida ({}): {}", file, e.getMessage());
        }
    }

    /** Normaliza un modo leido de fichero: invalido o ausente = default. */
    private static String normalizeMode(String raw) {
        if (raw == null) {
            return DEFAULT_MODE;
        }
        String value = raw.trim().toLowerCase(Locale.ROOT);
        return isValidMode(value) ? value : DEFAULT_MODE;
    }

    /** Valida un modo enviado por la interfaz; invalido = excepcion (400). */
    private static String requireValidMode(String value) {
        if (!isValidMode(value)) {
            throw new IllegalArgumentException(
                    "Modo de salida invalido '" + value + "': valores permitidos "
                            + String.join(", ", modes()) + ".");
        }
        return value;
    }

    /** Lee un booleano del fichero: ausente o no parseable = default. */
    private static boolean parseBool(String raw, boolean defaultValue) {
        if (raw == null) {
            return defaultValue;
        }
        String value = raw.trim().toLowerCase(Locale.ROOT);
        return "true".equals(value) || (!"false".equals(value) && defaultValue);
    }
}

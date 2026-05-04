package com.excelmerger;

import com.excelmerger.cli.InteractiveMenu;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

/**
 * Punto de entrada de la aplicacion Excel Merger.
 *
 * <p>v3.0.0: BREAKING CHANGE. La CLI argumentada de v2.7.1 se ha eliminado.
 * Al arrancar la aplicacion aparece SIEMPRE un menu interactivo. Las
 * antiguas flags ({@code --help}, {@code --version}, {@code --dry-run}) y
 * el argumento posicional {@code <configPath>} ya no se aceptan: cualquier
 * argumento pasado en la linea de comandos es ignorado silenciosamente.</p>
 *
 * <p>La logica de fusion (carga de config, validacion, merge, formateo del
 * banner final) vive ahora en {@link App}. {@code Main} se limita a
 * arrancar el menu interactivo y a propagar el exit code que este
 * devuelva.</p>
 *
 * <p>Para detalles del flujo dentro de cada opcion del menu, vease
 * {@link InteractiveMenu}.</p>
 *
 * <p>Codigos de salida (identicos a v2.7.1):</p>
 * <ul>
 *   <li>{@code 0} — ejecucion correcta (incluye salida via Opcion 3).</li>
 *   <li>{@code 1} — error en tiempo de ejecucion en la fusion.</li>
 *   <li>{@code 2} — configuracion invalida.</li>
 *   <li>{@code 3} — entrada invalida.</li>
 *   <li>{@code 4} — salida invalida.</li>
 * </ul>
 */
public final class Main {

    private static final Logger log = LoggerFactory.getLogger(Main.class);
    private static final String APP_NAME = "Excel Merger";
    public static final String APP_VERSION = "3.1.0";
    private static final String GIT_PROPERTIES_PATH = "/git.properties";

    private Main() {
        // Utility class
    }

    public static void main(String[] args) {
        // v3.0.0: los argumentos CLI se ignoran. El menu interactivo es
        // siempre el primer paso. No hay flags --help ni --version.
        int exitCode;
        try {
            exitCode = new InteractiveMenu().run();
        } catch (IOException e) {
            // Fallo al construir el Terminal de JLine. Caso muy raro
            // (entorno sin TTY ni stdin/stdout). Log y exit con 1.
            log.error("No se pudo inicializar el menu interactivo: {}", e.getMessage(), e);
            exitCode = App.EXIT_RUNTIME;
        }
        System.exit(exitCode);
    }

    /**
     * Devuelve la cadena de version mostrada (entre otros sitios) en el
     * banner ASCII al arrancar. Formato:
     * {@code "Excel Merger v3.0.0 (build a3f9b2c, 2026-04-22)"}. Si no hay
     * {@code git.properties} en el classpath o no contiene las claves
     * esperadas (compilacion fuera de un repo Git), cae limpiamente a
     * {@code "Excel Merger v3.0.0"}.
     */
    public static String buildInfoString() {
        Properties git = readGitProperties();
        if (git == null) {
            return APP_NAME + " v" + APP_VERSION;
        }
        String hash = git.getProperty("git.commit.id.abbrev");
        String date = git.getProperty("git.commit.time");
        if (hash == null || hash.isBlank() || date == null || date.isBlank()) {
            return APP_NAME + " v" + APP_VERSION;
        }
        return APP_NAME + " v" + APP_VERSION + " (build " + hash + ", " + date + ")";
    }

    /**
     * Carga {@code /git.properties} desde el classpath. Devuelve {@code null}
     * si el recurso no existe o no se puede leer (p. ej. compilacion fuera
     * de un repo Git con {@code failOnNoGitDirectory=false}).
     */
    private static Properties readGitProperties() {
        try (InputStream in = Main.class.getResourceAsStream(GIT_PROPERTIES_PATH)) {
            if (in == null) {
                return null;
            }
            Properties props = new Properties();
            props.load(in);
            return props;
        } catch (IOException e) {
            return null;
        }
    }
}

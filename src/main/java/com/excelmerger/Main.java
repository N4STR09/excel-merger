package com.excelmerger;

import com.excelmerger.cli.InteractiveMenu;
import com.excelmerger.compare.CompareRunner;
import com.excelmerger.web.WebLauncher;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.util.Properties;

/**
 * Punto de entrada de la aplicacion Excel Merger.
 *
 * <p>v4.0.0: deteccion de modos en la linea de comandos. Sin argumentos
 * arranca la interfaz web local, la experiencia por defecto que sustituye
 * al menu; el menu interactivo de terminal de v3.x sigue intacto detras
 * de {@code --cli}:</p>
 * <ul>
 *   <li>Sin argumentos — interfaz web local ({@link WebLauncher}).</li>
 *   <li>{@code --cli} — menu interactivo de terminal
 *       ({@link InteractiveMenu}), identico al de v3.x.</li>
 *   <li>{@code --merge} — fusion directa con {@code config.properties},
 *       mismo flujo que la Opcion 1 del menu, sin interaccion.</li>
 *   <li>{@code --compare} — comprobador de discrepancias directo, mismo
 *       flujo que la Opcion 2 del menu.</li>
 *   <li>{@code --help} / {@code -h} — ayuda por stdout, exit 0. Cualquier
 *       otro argumento: uso por stderr y exit 2.</li>
 * </ul>
 *
 * <p>v4.1.0: la interfaz web permite configurar en runtime que hojas
 * genera la fusion (modo de salida, hoja Resumen y tabla por responsable)
 * sin tocar {@code config.properties}; esos ajustes se aplican solo a las
 * ejecuciones de la interfaz web. Menu y linea de comandos siguen
 * cargando con su configuracion de siempre.</p>
 *
 * <p>La logica de fusion (carga de config, validacion, merge, formateo
 * del banner final) vive en {@link App}; el comprobador, en
 * {@link CompareRunner}; la interfaz web, en {@code com.excelmerger.web}.
 * {@code Main} solo decide el camino y propaga el exit code.</p>
 *
 * <p>Codigos de salida (identicos a v3.x):</p>
 * <ul>
 *   <li>{@code 0} — ejecucion correcta (Opcion 3 del menu o 'Salir' de
 *       la interfaz web).</li>
 *   <li>{@code 1} — error en tiempo de ejecucion en la fusion.</li>
 *   <li>{@code 2} — configuracion invalida (incluye argumento
 *       desconocido).</li>
 *   <li>{@code 3} — entrada invalida.</li>
 *   <li>{@code 4} — salida invalida.</li>
 * </ul>
 */
public final class Main {

    private static final Logger log = LoggerFactory.getLogger(Main.class);
    private static final String APP_NAME = "Excel Merger";
    public static final String APP_VERSION = "4.4.0";
    private static final String GIT_PROPERTIES_PATH = "/git.properties";

    private Main() {
        // Utility class
    }

    /**
     * Modo de arranque detectado en la linea de comandos (v4.0.0).
     */
    enum Mode {
        /** Sin argumentos: interfaz web local. */
        WEB,
        /** {@code --cli}: menu interactivo de terminal. */
        CLI,
        /** {@code --merge}: fusion directa. */
        MERGE,
        /** {@code --compare}: comprobador directo. */
        COMPARE,
        /** {@code --help} / {@code -h}: ayuda. */
        HELP,
        /** Argumento no reconocido. */
        UNKNOWN
    }

    /**
     * Decide el modo de arranque y lanza el flujo correspondiente,
     * propagando el exit code con {@code System.exit} (la interfaz web
     * decide el suyo internamente: 0 via su boton 'Salir', 1 si el
     * servidor no puede arrancar).
     *
     * @param args argumentos de la linea de comandos; su semantica esta
     *        descrita en el javadoc de la clase.
     * @throws InterruptedException si la interfaz web se interrumpe
     *         mientras espera (el flag lo restaura
     *         {@link WebLauncher#main(String[])}).
     */
    public static void main(String[] args) throws InterruptedException {
        Mode mode = parseMode(args);
        switch (mode) {
            case WEB -> WebLauncher.main(args);
            case CLI -> System.exit(runCli());
            case MERGE -> System.exit(App.run(null));
            case COMPARE -> System.exit(new CompareRunner().run());
            case HELP -> {
                printUsage(System.out);
                System.exit(App.EXIT_OK);
            }
            case UNKNOWN -> {
                System.err.println("Argumento desconocido: " + args[0]);
                printUsage(System.err);
                System.exit(App.EXIT_CONFIG);
            }
            default -> System.exit(App.EXIT_CONFIG);
        }
    }

    /**
     * Traduce los argumentos de la linea de comandos al modo de
     * arranque. Solo mira el primer argumento: los restantes se ignoran
     * (la fusion no acepta ruta de config alternativa desde v3.0.0).
     *
     * @param args argumentos; {@code null} o vacio equivale a
     *        {@link Mode#WEB}.
     * @return modo detectado; nunca {@code null}.
     */
    static Mode parseMode(String... args) {
        if (args == null || args.length == 0) {
            return Mode.WEB;
        }
        return switch (args[0]) {
            case "--cli" -> Mode.CLI;
            case "--merge" -> Mode.MERGE;
            case "--compare" -> Mode.COMPARE;
            case "--help", "-h" -> Mode.HELP;
            default -> Mode.UNKNOWN;
        };
    }

    /**
     * Menu interactivo de terminal, identico al de v3.x: cualquier fallo
     * de inicializacion de JLine (entorno sin TTY) se loguea y devuelve
     * {@link App#EXIT_RUNTIME}, igual que siempre.
     *
     * @return el exit code que devuelva el menu.
     */
    private static int runCli() {
        try {
            return new InteractiveMenu().run();
        } catch (IOException e) {
            // Fallo al construir el Terminal de JLine. Caso muy raro
            // (entorno sin TTY ni stdin/stdout). Log y exit con 1.
            log.error("No se pudo inicializar el menu interactivo: {}", e.getMessage(), e);
            return App.EXIT_RUNTIME;
        }
    }

    /**
     * Imprime la ayuda de modos. Textos nuevos de v4.0.0: la linea de
     * comandos es superficie de arranque (lo cambiante del contrato);
     * los textos de fusion y comparacion no se tocan.
     *
     * @param out stdout en {@code --help}; stderr cuando el argumento es
     *        desconocido.
     */
    static void printUsage(PrintStream out) {
        out.println("Uso: java -jar excel-merger.jar [modo]");
        out.println();
        out.println("Modos:");
        out.println("  (sin argumentos)  Interfaz web local (abre el navegador)");
        out.println("  --cli             Menu interactivo de terminal (como v3.x)");
        out.println("  --merge           Fusion directa con config.properties");
        out.println("  --compare         Comprobador de discrepancias directo");
        out.println("  --help, -h        Esta ayuda");
        out.println();
        out.println("--merge y --compare devuelven los exit codes 0-4.");
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

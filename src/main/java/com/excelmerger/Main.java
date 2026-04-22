package com.excelmerger;

import com.excelmerger.exception.ConfigurationException;
import com.excelmerger.exception.InputValidationException;
import com.excelmerger.exception.MergeException;
import com.excelmerger.exception.OutputException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.time.Duration;
import java.time.Instant;
import java.util.Arrays;
import java.util.List;

/**
 * Punto de entrada de la aplicacion Excel Merger.
 *
 * <p>Uso:</p>
 * <ul>
 *   <li>{@code java -jar excel-merger.jar} — usa {@code config.properties} por defecto.</li>
 *   <li>{@code java -jar excel-merger.jar mi-config.properties} — usa un config alternativo.</li>
 *   <li>{@code java -jar excel-merger.jar --help | -h} — muestra la ayuda.</li>
 *   <li>{@code java -jar excel-merger.jar --version | -v} — muestra la version.</li>
 *   <li>{@code java -jar excel-merger.jar --dry-run} — ejecuta el pipeline sin escribir
 *       el Excel de salida ni tocar el historial (util para validar antes de cierre).</li>
 * </ul>
 *
 * <p>Codigos de salida:</p>
 * <ul>
 *   <li>{@code 0} — ejecucion correcta.</li>
 *   <li>{@code 1} — error en tiempo de ejecucion ({@code MergeException} u otro no tipado).</li>
 *   <li>{@code 2} — {@code ConfigurationException}: config invalida o no cargable, o errores
 *       detectados con {@code config.strictValidation=true}.</li>
 *   <li>{@code 3} — {@code InputValidationException}: directorio de entrada o ficheros Excel mal.</li>
 *   <li>{@code 4} — {@code OutputException}: no se puede preparar o escribir el output.</li>
 * </ul>
 */
public final class Main {

    private static final Logger log = LoggerFactory.getLogger(Main.class);
    private static final String APP_NAME = "Excel Merger";
    private static final String APP_VERSION = "1.4.0";

    // Exit codes (mantenidos compatibles con versiones anteriores: 0, 1, 2)
    static final int EXIT_OK               = 0;
    static final int EXIT_RUNTIME          = 1;
    static final int EXIT_CONFIG           = 2;
    static final int EXIT_INPUT_INVALID    = 3;
    static final int EXIT_OUTPUT_INVALID   = 4;

    private Main() {
        // Utility class
    }

    public static void main(String[] args) {
        List<String> argList = Arrays.asList(args);

        if (argList.contains("--help") || argList.contains("-h")) {
            printHelp();
            System.exit(EXIT_OK);
        }
        if (argList.contains("--version") || argList.contains("-v")) {
            System.out.println(APP_NAME + " v" + APP_VERSION);
            System.exit(EXIT_OK);
        }

        boolean dryRun = argList.contains("--dry-run");

        String configPath = argList.stream()
                .filter(a -> !a.startsWith("-"))
                .findFirst()
                .orElse(null);

        Instant start = Instant.now();
        log.info("====================================");
        log.info("   {} v{}", APP_NAME, APP_VERSION);
        log.info("====================================");

        RunReport report = new RunReport();
        int exitCode = EXIT_OK;

        try {
            ConfigLoader config = (configPath != null)
                    ? new ConfigLoader(configPath)
                    : new ConfigLoader();

            // --- Validacion previa del config ---
            boolean strict = config.getBoolean("config.strictValidation", true);
            List<String> errors = new ConfigValidator(config).validate();
            if (!errors.isEmpty()) {
                if (strict) {
                    log.error("====================================");
                    log.error("   CONFIGURACION INVALIDA ({} error(es))", errors.size());
                    log.error("====================================");
                    for (String err : errors) {
                        log.error("  - {}", err);
                    }
                    log.error("Para avanzar igualmente con warnings en vez de abortar,");
                    log.error("anade a tu config.properties: config.strictValidation=false");
                    System.exit(EXIT_CONFIG);
                } else {
                    log.warn("Configuracion con {} aviso(s) (strictValidation=false):", errors.size());
                    for (String err : errors) {
                        log.warn("  - {}", err);
                        report.addWarning("CONFIG", err);
                    }
                }
            }

            // --- Merge ---
            ExcelMerger merger = new ExcelMerger(config, report, dryRun);
            merger.merge();

            Duration elapsed = Duration.between(start, Instant.now());
            log.info(report.formatSummary(elapsed));
            log.info("====================================");
            if (dryRun) {
                log.info("   PROCESO FINALIZADO OK (DRY-RUN, {} ms)", elapsed.toMillis());
                log.info("   [DRY-RUN] No se ha escrito el Excel de salida ni se ha movido ningun backup.");
            } else {
                log.info("   PROCESO FINALIZADO OK ({} ms)", elapsed.toMillis());
            }
            log.info("====================================");
            exitCode = EXIT_OK;

        } catch (ConfigurationException e) {
            exitCode = logAndExit(start, report, e, "CONFIGURACION", EXIT_CONFIG);
        } catch (InputValidationException e) {
            exitCode = logAndExit(start, report, e, "ENTRADA INVALIDA", EXIT_INPUT_INVALID);
        } catch (OutputException e) {
            exitCode = logAndExit(start, report, e, "OUTPUT INVALIDO", EXIT_OUTPUT_INVALID);
        } catch (MergeException e) {
            exitCode = logAndExit(start, report, e, "ERROR EN LA FUSION", EXIT_RUNTIME);
        } catch (Exception e) {
            exitCode = logAndExit(start, report, e, "ERROR EN LA EJECUCION", EXIT_RUNTIME);
        }

        System.exit(exitCode);
    }

    private static int logAndExit(Instant start, RunReport report, Exception e,
                                  String banner, int code) {
        Duration elapsed = Duration.between(start, Instant.now());
        log.error("====================================");
        log.error("   {}", banner);
        log.error("====================================");
        log.error("Mensaje: {}", e.getMessage(), e);
        log.info(report.formatSummary(elapsed));
        return code;
    }

    private static void printHelp() {
        System.out.println(APP_NAME + " v" + APP_VERSION);
        System.out.println();
        System.out.println("Uso: java -jar excel-merger.jar [opciones] [config.properties]");
        System.out.println();
        System.out.println("Opciones:");
        System.out.println("  --help, -h        Muestra esta ayuda");
        System.out.println("  --version, -v     Muestra la version");
        System.out.println("  --dry-run         Ejecuta el pipeline completo (validacion, deteccion de apps");
        System.out.println("                    sin mapeo, cabeceras, perfiles...) pero NO escribe el Excel");
        System.out.println("                    de salida ni mueve el anterior a history/.");
        System.out.println();
        System.out.println("Si no se indica un config, se usa 'config.properties' del directorio actual.");
        System.out.println();
        System.out.println("Codigos de salida:");
        System.out.println("  0   ejecucion correcta");
        System.out.println("  1   error en tiempo de ejecucion (MergeException u otro no tipado)");
        System.out.println("  2   configuracion invalida o no cargable (ConfigurationException,");
        System.out.println("      o config.strictValidation=true con errores)");
        System.out.println("  3   entrada invalida: directorio o ficheros Excel mal (InputValidationException)");
        System.out.println("  4   salida invalida: no se puede preparar o escribir el output (OutputException)");
    }
}

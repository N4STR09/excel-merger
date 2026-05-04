package com.excelmerger;

import com.excelmerger.config.ConfigValidator;
import com.excelmerger.exception.ConfigurationException;
import com.excelmerger.exception.InputValidationException;
import com.excelmerger.exception.MergeException;
import com.excelmerger.exception.OutputException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.time.Duration;
import java.time.Instant;
import java.util.List;

/**
 * Logica de la aplicacion Excel Merger, extraida de {@link Main} en v3.0.0
 * para que sea testeable sin pasar por {@code System.exit(...)} ni por el
 * menu interactivo.
 *
 * <p>Esta clase contiene exactamente el flujo que en v2.7.1 estaba dentro
 * de {@code Main.main(args)}: carga del config, validacion previa, merge
 * y formateo del banner final OK / KO. La unica diferencia es que en lugar
 * de invocar {@link System#exit(int)} devuelve el exit code como int para
 * que el caller decida que hacer (en produccion, {@link Main} hace exit;
 * en tests, se assertea el valor devuelto).</p>
 *
 * <p>Codigos de salida (identicos a v2.7.1):</p>
 * <ul>
 *   <li>{@code 0} — ejecucion correcta.</li>
 *   <li>{@code 1} — error en tiempo de ejecucion ({@link MergeException} u
 *       otro no tipado).</li>
 *   <li>{@code 2} — {@link ConfigurationException}: config invalida o no
 *       cargable, o errores con {@code config.strictValidation=true}.</li>
 *   <li>{@code 3} — {@link InputValidationException}: directorio o
 *       ficheros Excel mal.</li>
 *   <li>{@code 4} — {@link OutputException}: no se puede preparar o
 *       escribir el output.</li>
 * </ul>
 */
public final class App {

    private static final Logger log = LoggerFactory.getLogger(App.class);
    private static final String APP_NAME = "Excel Merger";
    /** Separador visual usado en los banners del log. */
    private static final String BANNER_SEPARATOR = "====================================";

    // Exit codes (mantenidos identicos a v2.7.1)
    public static final int EXIT_OK             = 0;
    public static final int EXIT_RUNTIME        = 1;
    public static final int EXIT_CONFIG         = 2;
    public static final int EXIT_INPUT_INVALID  = 3;
    public static final int EXIT_OUTPUT_INVALID = 4;

    private App() {
        // Utility class
    }

    /**
     * Ejecuta el flujo completo de fusion. Equivalente al cuerpo de
     * {@code Main.main(args)} en v2.7.1, pero sin {@code System.exit}.
     *
     * @param configPath ruta al fichero de configuracion. Si es {@code null},
     *                   se usa el default ({@code config.properties}).
     * @return exit code (0 = OK; 1-4 segun excepcion).
     */
    public static int run(String configPath) {
        Instant start = Instant.now();
        log.info(BANNER_SEPARATOR);
        log.info("   {} v{}", APP_NAME, Main.APP_VERSION);
        log.info(BANNER_SEPARATOR);

        RunReport report = new RunReport();

        try {
            ConfigLoader config = (configPath != null)
                    ? new ConfigLoader(configPath)
                    : new ConfigLoader();

            // v3.0.0: dry-run se lee del config en vez de la CLI desaparecida.
            boolean dryRun = config.getBoolean("output.dryRun", false);

            // --- Validacion previa del config ---
            boolean strict = config.getBoolean("config.strictValidation", true);
            List<String> errors = new ConfigValidator(config).validate();
            if (!errors.isEmpty()) {
                if (strict) {
                    log.error(BANNER_SEPARATOR);
                    log.error("   CONFIGURACION INVALIDA ({} error(es))", errors.size());
                    log.error(BANNER_SEPARATOR);
                    for (String err : errors) {
                        log.error("  - {}", err);
                    }
                    log.error("Para avanzar igualmente con warnings en vez de abortar,");
                    log.error("anade a tu config.properties: config.strictValidation=false");
                    return EXIT_CONFIG;
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
            log.info(BANNER_SEPARATOR);
            if (dryRun) {
                log.info("   PROCESO FINALIZADO OK (DRY-RUN, {} ms)", elapsed.toMillis());
                log.info("   [DRY-RUN] No se ha escrito el Excel de salida ni se ha movido ningun backup.");
            } else {
                log.info("   PROCESO FINALIZADO OK ({} ms)", elapsed.toMillis());
            }
            log.info(BANNER_SEPARATOR);
            return EXIT_OK;

        } catch (ConfigurationException e) {
            return logAndReturn(start, report, e, "CONFIGURACION", EXIT_CONFIG);
        } catch (InputValidationException e) {
            return logAndReturn(start, report, e, "ENTRADA INVALIDA", EXIT_INPUT_INVALID);
        } catch (OutputException e) {
            return logAndReturn(start, report, e, "OUTPUT INVALIDO", EXIT_OUTPUT_INVALID);
        } catch (MergeException e) {
            return logAndReturn(start, report, e, "ERROR EN LA FUSION", EXIT_RUNTIME);
        } catch (Exception e) {
            return logAndReturn(start, report, e, "ERROR EN LA EJECUCION", EXIT_RUNTIME);
        }
    }

    private static int logAndReturn(Instant start, RunReport report, Exception e,
                                    String banner, int code) {
        Duration elapsed = Duration.between(start, Instant.now());
        log.error(BANNER_SEPARATOR);
        log.error("   {}", banner);
        log.error(BANNER_SEPARATOR);
        log.error("Mensaje: {}", e.getMessage(), e);
        log.info(report.formatSummary(elapsed));
        return code;
    }
}

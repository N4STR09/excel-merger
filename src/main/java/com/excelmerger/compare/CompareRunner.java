package com.excelmerger.compare;

import com.excelmerger.App;
import com.excelmerger.exception.InputValidationException;
import com.excelmerger.exception.OutputException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Comparator;

/**
 * Orquesta el flujo completo de la Opcion 2 del menu: comprobador de
 * discrepancias contra CSV externos del ERP. v3.1.0.
 *
 * <h2>Flujo</h2>
 * <ol>
 *   <li>Localizar todos los {@code *.csv} (case-insensitive) en el
 *       directorio de input ({@code input/} por defecto). Si no hay
 *       ninguno, mostrar mensaje informativo y devolver
 *       {@link App#EXIT_OK} (vuelve al menu).</li>
 *   <li>Verificar que existe {@code output/resultado_fusion.xlsx}. Si no,
 *       mensaje "Ejecuta primero la Opcion 1" y devolver {@code EXIT_OK}.</li>
 *   <li>Leer el Resultado con {@link ResultadoReader} (incluye
 *       {@code FormulaEvaluator.evaluateAll()}).</li>
 *   <li>Para cada CSV, parsearlo con {@link CsvParser} y compararlo
 *       contra el Resultado con {@link DiscrepancyComparator}. Acumular
 *       todas las discrepancias.</li>
 *   <li>Escribir un unico Excel
 *       {@code output/discrepancias_<yyyy-MM-dd_HHmmss>.xlsx} con
 *       {@link DiscrepancyExporter}.</li>
 * </ol>
 *
 * <h2>Politica de errores</h2>
 * <p>Los casos "no hay nada que hacer" (no CSV, no Resultado) NO son
 * errores: se informa al usuario y se devuelve {@code EXIT_OK} para que
 * el menu vuelva a aparecer. Los errores reales (CSV mal formado,
 * Resultado corrupto, IO escritura) se loguean y devuelven
 * {@code EXIT_RUNTIME} (1) o {@code EXIT_INPUT_INVALID} (3) segun el caso,
 * para que {@code InteractiveMenu} pueda mostrar el banner de error
 * antes de volver al menu.</p>
 */
public final class CompareRunner {

    private static final Logger log = LoggerFactory.getLogger(CompareRunner.class);

    /** Patron de timestamp para el nombre del Excel de salida. */
    private static final DateTimeFormatter TIMESTAMP_FORMAT =
            DateTimeFormatter.ofPattern("yyyy-MM-dd_HHmmss");

    /** Path por defecto del Excel del Resultado (mismo que {@code output.file}). */
    static final Path DEFAULT_RESULTADO_PATH = Paths.get("output", "resultado_fusion.xlsx");

    /** Path por defecto del directorio de input. */
    static final Path DEFAULT_INPUT_DIR = Paths.get("input");

    /** Path por defecto del directorio de output. */
    static final Path DEFAULT_OUTPUT_DIR = Paths.get("output");

    private final Path inputDir;
    private final Path resultadoPath;
    private final Path outputDir;
    private final PrintStream out;

    /**
     * Constructor de produccion. Usa rutas por defecto y stdout UTF-8.
     */
    public CompareRunner() {
        this(DEFAULT_INPUT_DIR, DEFAULT_RESULTADO_PATH, DEFAULT_OUTPUT_DIR,
                new PrintStream(System.out, true, StandardCharsets.UTF_8));
    }

    /**
     * Constructor para tests: rutas inyectables y PrintStream capturable.
     *
     * <p>Visibilidad: <b>package-private</b> intencional. Mantenerlo fuera
     * del API publica evita el bug SpotBugs EI_EXPOSE_REP2 (almacenar un
     * {@link PrintStream} mutable recibido externamente). El paquete
     * {@code com.excelmerger.compare} contiene los tests del comparador,
     * que pueden invocarlo libremente. Mismo patron que
     * {@link com.excelmerger.cli.InteractiveMenu}, que tambien tiene el
     * constructor inyectable como package-private por la misma razon.</p>
     */
    CompareRunner(Path inputDir, Path resultadoPath, Path outputDir, PrintStream out) {
        this.inputDir = inputDir;
        this.resultadoPath = resultadoPath;
        this.outputDir = outputDir;
        this.out = out;
    }

    /**
     * Ejecuta el flujo completo. Devuelve un exit code compatible con
     * los de {@link App}.
     */
    public int run() {
        out.println();
        out.println("=== Comprobador de discrepancias contra CSV externos ===");
        out.println();

        // 1. Localizar CSVs.
        List<Path> csvs;
        try {
            csvs = listCsvFiles(inputDir);
        } catch (IOException e) {
            log.error("[CompareRunner] No se pudo listar el directorio de input: {}", e.getMessage(), e);
            out.println("[ERROR] No se pudo listar " + inputDir + ": " + e.getMessage());
            return App.EXIT_INPUT_INVALID;
        }
        if (csvs.isEmpty()) {
            out.println("[INFO] No se encontraron CSV en " + inputDir
                    + ". Coloca los exports del ERP (formato <matricula>.CSV) y vuelve a ejecutar.");
            return App.EXIT_OK;
        }
        out.println("Detectados " + csvs.size() + " fichero(s) CSV en " + inputDir + ":");
        for (Path c : csvs) {
            out.println("  - " + fileNameOf(c));
        }

        // 2. Verificar Resultado.
        if (!Files.exists(resultadoPath)) {
            out.println();
            out.println("[INFO] No se encuentra " + resultadoPath
                    + ". Ejecuta primero la Opcion 1 para generar el Excel de fusion.");
            return App.EXIT_OK;
        }

        // 3. Leer Resultado (FormulaEvaluator dentro).
        Map<DiscrepancyKey, Double> resultadoMap;
        try {
            resultadoMap = new ResultadoReader().read(resultadoPath);
        } catch (InputValidationException e) {
            log.error("[CompareRunner] Error leyendo Resultado: {}", e.getMessage(), e);
            out.println("[ERROR] " + e.getMessage());
            return App.EXIT_INPUT_INVALID;
        }
        out.println();
        out.println("Resultado cargado: " + resultadoMap.size() + " entradas (Matricula+Peticion+Funcion).");

        // 4. Procesar cada CSV.
        CsvParser parser = new CsvParser();
        DiscrepancyComparator comparator = new DiscrepancyComparator();
        List<Discrepancy> all = new ArrayList<>();
        for (Path csv : csvs) {
            String fileName = fileNameOf(csv);
            String origen = CsvParser.originFromFileName(fileName);
            out.println();
            out.println("Procesando " + fileName + " (origen=" + origen + ")...");
            try {
                List<CsvImputacion> imps = parser.parse(csv);
                List<Discrepancy> discs = comparator.compare(origen, imps, resultadoMap);
                all.addAll(discs);
                out.println("  -> " + imps.size() + " imputacion(es) leidas, "
                        + discs.size() + " discrepancia(s).");
            } catch (InputValidationException e) {
                log.error("[CompareRunner] Error procesando {}: {}", csv, e.getMessage(), e);
                out.println("  [ERROR] " + e.getMessage());
                // Continua con el resto de CSVs en lugar de abortar todo.
            }
        }

        all.sort(
            Comparator
                .comparing(Discrepancy::getPeticion, Comparator.nullsFirst(String::compareTo))
                .thenComparing(Discrepancy::getMatricula)
                .thenComparing(Discrepancy::getFuncion)
        );

        // 5. Escribir Excel de salida.
        Path outFile = outputDir.resolve("discrepancias_"
                + LocalDateTime.now().format(TIMESTAMP_FORMAT) + ".xlsx");
        try {
            new DiscrepancyExporter().write(outFile, all);
        } catch (OutputException e) {
            log.error("[CompareRunner] Error escribiendo Excel: {}", e.getMessage(), e);
            out.println();
            out.println("[ERROR] " + e.getMessage());
            return App.EXIT_OUTPUT_INVALID;
        }

        out.println();
        out.println("[OK] Excel de discrepancias generado: " + outFile);
        out.println("Total de discrepancias: " + all.size());
        return App.EXIT_OK;
    }

    /**
     * Lista los ficheros con extension {@code .csv} (case-insensitive)
     * presentes en {@code dir}. Si {@code dir} no existe, devuelve lista
     * vacia (caso "no hay nada que hacer", no es error).
     */
    static List<Path> listCsvFiles(Path dir) throws IOException {
        if (dir == null || !Files.exists(dir) || !Files.isDirectory(dir)) {
            return Collections.emptyList();
        }
        List<Path> result = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir)) {
            for (Path p : stream) {
                if (Files.isRegularFile(p)
                        && CsvParser.isCsvFileName(fileNameOf(p))) {
                    result.add(p);
                }
            }
        }
        // Orden alfabetico estable, case-insensitive (mismo criterio que
        // InputFileDetector usa para los Excel).
        result.sort((a, b) -> fileNameOf(a).toLowerCase(Locale.ROOT)
                .compareTo(fileNameOf(b).toLowerCase(Locale.ROOT)));
        return result;
    }

    /**
     * Devuelve el nombre del fichero del path como {@link String}, o cadena
     * vacia si {@link Path#getFileName()} es {@code null} (caso degenerado:
     * paths que representan una raiz como {@code /} en Linux o {@code C:\}
     * en Windows). Centraliza el null-check para evitar el bug SpotBugs
     * NP_NULL_ON_SOME_PATH_FROM_RETURN_VALUE.
     */
    private static String fileNameOf(Path p) {
        Path name = p.getFileName();
        return name == null ? "" : name.toString();
    }
}

package com.excelmerger;

import com.excelmerger.exception.MergeException;
import com.excelmerger.io.FileLockDetector;
import com.excelmerger.io.InputFileDetector;
import com.excelmerger.io.OutputManager;
import com.excelmerger.io.SheetCopier;
import com.excelmerger.util.PoiUtils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * Motor principal de fusion de ficheros Excel. <b>Orquestador puro</b> que
 * delega las responsabilidades en colaboradores especializados del paquete
 * {@link com.excelmerger.io} y en los builders de hojas.
 *
 * <p>La entrada se define como un DIRECTORIO: {@link InputFileDetector}
 * detecta automaticamente los dos primeros archivos .xlsx/.xls (orden
 * alfabetico). Soporta dos modos: {@code SHEETS_SEPARATE} (copia cada hoja
 * de ambos libros) y {@code APPEND_ROWS} (concatena las filas de la primera
 * hoja de cada libro).</p>
 *
 * <p>Modo <b>dry-run</b>: el constructor sobrecargado con {@code dryRun=true}
 * ejecuta todo el pipeline (validacion de config, deteccion de perfiles,
 * construccion de hojas en memoria, chequeo de apps sin mapeo) pero NO
 * escribe el Excel final ni mueve el anterior a {@code history/}. Util para
 * validar la configuracion antes de un cierre. Los chequeos de lock {@code ~$}
 * sobre el output si se mantienen: si el fichero esta abierto en Excel, se
 * avisa igualmente.</p>
 */
public class ExcelMerger {

    private static final Logger log = LoggerFactory.getLogger(ExcelMerger.class);

    public enum MergeMode {
        SHEETS_SEPARATE,
        APPEND_ROWS
    }

    private final ConfigLoader config;
    private final RunReport report;
    private final boolean dryRun;

    // Colaboradores
    private final InputFileDetector inputDetector;
    private final FileLockDetector lockDetector;
    private final OutputManager outputManager;
    private final SheetCopier sheetCopier;

    public ExcelMerger(ConfigLoader config, RunReport report) {
        this(config, report, false);
    }

    public ExcelMerger(ConfigLoader config, RunReport report, boolean dryRun) {
        this.config = config;
        this.report = report;
        this.dryRun = dryRun;
        this.inputDetector = new InputFileDetector();
        this.lockDetector = new FileLockDetector();
        this.outputManager = new OutputManager();
        this.sheetCopier = new SheetCopier(config.getBoolean("merge.copyStyles", true));
    }

    /** Clave de configuracion legada para la validacion estricta "exactamente 2 ficheros". */
    private static final String KEY_STRICT_TWO_FILES = "input.strictTwoFiles";
    /** Clave v2.2.0: minimo de ficheros Excel aceptados en el directorio de entrada. */
    private static final String KEY_STRICT_MIN_FILES = "input.strictMinFiles";
    /** Clave v2.2.0: maximo de ficheros Excel aceptados en el directorio de entrada. */
    private static final String KEY_STRICT_MAX_FILES = "input.strictMaxFiles";

    public void merge() {
        String inputDir = config.get("input.directory");
        String outputPath = config.get("output.file");
        MergeMode mode = MergeMode.valueOf(config.get("merge.mode", "SHEETS_SEPARATE"));
        boolean overwrite = config.getBoolean("output.overwrite", true);
        boolean backup = config.getBoolean("output.backup", false);

        // v2.2.0: resolver rango [minFiles, maxFiles] con retrocompat de
        // input.strictTwoFiles. Si se usa la clave legada, se emite warning
        // CONFIG y se mapea a min/max.
        int[] range = resolveFileCountRange();
        int minFiles = range[0];
        int maxFiles = range[1];

        // 1. Detectar y validar los ficheros de entrada (lista posiblemente truncada)
        List<File> excelFilesRaw = inputDetector.findExcelFiles(inputDir);
        // Retrocompat: input.strictTwoFiles=true era "exactamente 2" con
        // aborto si habia mas. La nueva API [min,max] trunca en exceso
        // con warning. Para preservar el contrato legado, cuando el usuario
        // usa explicitamente strictTwoFiles=true y hay >2 ficheros, abortamos
        // con el mismo mensaje de antes.
        if (config.has(KEY_STRICT_TWO_FILES)
                && !(config.has(KEY_STRICT_MIN_FILES) || config.has(KEY_STRICT_MAX_FILES))
                && config.getBoolean(KEY_STRICT_TWO_FILES, true)
                && excelFilesRaw.size() > 2) {
            throw new com.excelmerger.exception.InputValidationException(
                    "Se han encontrado " + excelFilesRaw.size() + " archivos Excel, pero se esperaban "
                            + "exactamente 2 (" + KEY_STRICT_TWO_FILES + "=true). Archivos detectados: "
                            + excelFilesRaw.stream().map(File::getName)
                                    .collect(Collectors.joining(", ")));
        }
        List<File> excelFiles = inputDetector.validateExcelFiles(excelFilesRaw, minFiles, maxFiles, report);

        // 2. Evitar que el archivo de salida coincida con uno de los de entrada
        Path outAbs = Paths.get(outputPath).toAbsolutePath().normalize();
        for (File f : excelFiles) {
            if (outAbs.equals(f.toPath().toAbsolutePath().normalize())) {
                throw new com.excelmerger.exception.InputValidationException(
                        "El archivo de salida no puede coincidir con uno de los archivos de entrada: " + outAbs);
            }
        }

        // 3. Chequeos tempranos de lock (inputs y output)
        for (File f : excelFiles) {
            lockDetector.assertNotLocked(f);
        }
        outputManager.assertOutputWritable(outputPath);

        // 4. Preparar el fichero de salida (backup, overwrite, crear dir).
        //    En dry-run no se toca el disco: ni se mueve el output previo al
        //    history/, ni se crea el directorio de salida. El chequeo de lock
        //    anterior si se mantiene: si el output esta abierto en Excel
        //    queremos avisar ya, tanto en run real como en dry-run.
        if (!dryRun) {
            outputManager.prepareOutputFile(outputPath, overwrite, backup);
        }

        log.info("Directorio de entrada: {}", Paths.get(inputDir).toAbsolutePath());
        int fileIdx = 1;
        for (File f : excelFiles) {
            log.info("Archivo {}: {}", fileIdx, f.getName());
            fileIdx++;
        }
        log.info("Modo de fusion: {}", mode);
        log.info("Archivo de salida: {}", outputPath);
        if (dryRun) {
            log.info("[DRY-RUN] Activado: el pipeline se ejecutara pero no se escribira el Excel final.");
        }

        // 5. Abrir archivos y ejecutar la fusion. Los workbooks/streams de
        //    entrada se encapsulan en InputWorkbooks (AutoCloseable) para que
        //    el try-with-resources se encargue del cierre en todos los
        //    caminos de salida, incluidas excepciones.
        try (Workbook result = new XSSFWorkbook();
             InputWorkbooks inputs = InputWorkbooks.open(excelFiles, lockDetector)) {

            List<Workbook> workbooks = inputs.workbooks();

            // 5a. Identificar cada archivo por su CONTENIDO
            FileProfileResolver resolver = new FileProfileResolver(config);
            List<FileProfileResolver.FileProfile> profiles = new ArrayList<>();
            if (resolver.hasProfiles()) {
                Set<String> seenIds = new HashSet<>();
                for (int i = 0; i < excelFiles.size(); i++) {
                    FileProfileResolver.FileProfile p = resolver.resolve(workbooks.get(i), excelFiles.get(i));
                    profiles.add(p);
                    if (p == null) {
                        log.warn("'{}' no coincide con ningun perfil. Se usaran los nombres de hoja originales.",
                                excelFiles.get(i).getName());
                        report.addWarning("PERFIL",
                                "'" + excelFiles.get(i).getName() + "' no coincide con ningun perfil.");
                    } else if (!seenIds.add(p.getId())) {
                        log.warn("Varios archivos coinciden con el perfil '{}'.", p.getId());
                        report.addWarning("PERFIL",
                                "Varios archivos coinciden con el perfil '" + p.getId() + "'.");
                    }
                }
            } else {
                // Sin perfiles configurados: rellenar con null (1 por fichero)
                // para que la iteracion indexada posterior sobre profiles
                // siga funcionando.
                while (profiles.size() < excelFiles.size()) {
                    profiles.add(null);
                }
            }

            // 5b. Fusionar segun modo
            if (mode == MergeMode.SHEETS_SEPARATE) {
                mergeSheetsSeparate(workbooks, excelFiles, result, profiles);
            } else {
                mergeAppendRows(workbooks.get(0), workbooks.get(1), result);
            }

            // 5c. Hojas de lookup (tablas estaticas para VLOOKUP)
            new LookupSheetBuilder(config, report).buildAll(result);

            // 5d. Hoja MES (peticiones del perfil Cierre + columnas calculadas).
            //     v2.0.0: tras el swap, la hoja con las peticiones se llama
            //     Cierre (antes Extraccion).
            new MesSheetBuilder(config, report).build(result);

            // 5e. Hojas derivadas (formulas / agregaciones)
            new DerivedSheetBuilder(config, report).buildAll(result);

            // 5f. Hoja "Resumen" (v1.6.0): panel de cierre mensual con
            //     SUMIFS sobre Resultado y la hoja del export de Jira.
            //     Se construye tras MES para poder referenciar la hoja
            //     Resultado por fórmula.
            //     v2.0.0: la hoja con las imputaciones de Jira se llama
            //     Extraccion (antes Cierre).
            new SummarySheetBuilder(config, report).build(result);

            // 5g. Hoja de avisos (opt-in con report.inExcel=true). Se construye
            //     despues del resto de builders para recoger TODOS los warnings
            //     acumulados (perfiles sin match, apps sin mapeo, cabeceras no
            //     encontradas, etc.). En dry-run se construye igual: queda en
            //     memoria pero no se escribe a disco.
            new AvisosSheetBuilder(config, report).build(result);

            // 5h. Escribir el resultado (omitido en dry-run)
            if (dryRun) {
                log.info("[DRY-RUN] Omitida escritura de '{}'.", outputPath);
            } else {
                outputManager.writeResult(result, outputPath);
            }

            log.info("Fusion completada correctamente{}.", dryRun ? " (dry-run, sin escritura)" : "");
        } catch (IOException e) {
            throw new MergeException("Fallo durante la fusion: " + e.getMessage(), e);
        }
    }

    /**
     * Bolsa AutoCloseable que agrupa los streams y workbooks de los ficheros
     * de entrada. Abrir via {@link #open(List, FileLockDetector)}; si la
     * apertura de cualquier fichero falla a mitad, los ya abiertos se
     * cierran antes de propagar la excepcion.
     *
     * <p>Existe para que el {@code merge()} pueda usar try-with-resources
     * sobre un numero variable de ficheros (2 o 3 en v2.2.0).</p>
     */
    private static final class InputWorkbooks implements AutoCloseable {
        private static final Logger LOG = LoggerFactory.getLogger(InputWorkbooks.class);
        private final List<FileInputStream> streams = new ArrayList<>();
        private final List<Workbook> workbooks = new ArrayList<>();

        private InputWorkbooks() {
            // instancia vacia; los elementos se agregan via open()
        }

        /**
         * Abre todos los ficheros indicados y los encapsula en una nueva
         * instancia. Si la apertura de algun fichero intermedio falla,
         * los recursos ya abiertos se cierran antes de propagar la
         * excepcion original, manteniendo la invariante "o todo abierto
         * o nada abierto".
         */
        static InputWorkbooks open(List<File> files, FileLockDetector lockDetector) throws IOException {
            InputWorkbooks bag = new InputWorkbooks();
            try {
                for (File f : files) {
                    bag.addFile(f, lockDetector);
                }
                return bag;
            } catch (IOException | RuntimeException e) {
                bag.close();
                throw e;
            }
        }

        @SuppressWarnings("PMD.CloseResource")
        private void addFile(File f, FileLockDetector lockDetector) throws IOException {
            // PMD CloseResource: falso positivo. El FileInputStream se
            // transfiere a this.streams y el Workbook a this.workbooks en
            // la misma llamada; ambos se cierran en close() vivo el ciclo
            // de InputWorkbooks (gestionado por el try-with-resources
            // externo en merge()).
            FileInputStream fis = lockDetector.openForRead(f);
            streams.add(fis);
            workbooks.add(WorkbookFactory.create(fis));
        }

        List<Workbook> workbooks() {
            return workbooks;
        }

        @Override
        @SuppressWarnings("PMD.CloseResource")
        public void close() {
            // PMD CloseResource: falso positivo. Las variables locales 'w' y
            // 'fis' son alias a elementos de las listas propias de la clase;
            // el close() se llama explicitamente sobre cada uno en este
            // mismo metodo. PMD no vincula el iterador del foreach con el
            // close() dentro del bucle.
            for (Workbook w : workbooks) {
                try {
                    w.close();
                } catch (IOException e) {
                    LOG.warn("No se pudo cerrar un workbook de entrada: {}", e.getMessage());
                }
            }
            for (FileInputStream fis : streams) {
                try {
                    fis.close();
                } catch (IOException e) {
                    LOG.warn("No se pudo cerrar un FileInputStream de entrada: {}", e.getMessage());
                }
            }
        }
    }

    /**
     * v2.2.0: resuelve el rango [min,max] de ficheros aceptados.
     *
     * <p>Precedencia:</p>
     * <ol>
     *   <li>Si {@code input.strictMinFiles} y/o {@code input.strictMaxFiles}
     *       estan presentes, se usan. Default: min=2, max=3.</li>
     *   <li>Si solo esta presente la clave legada {@code input.strictTwoFiles},
     *       se mapea: {@code true} -> [2,2]; {@code false} -> [2, Integer.MAX_VALUE].
     *       Se emite warning CONFIG de deprecacion.</li>
     *   <li>Si ambas familias estan presentes, mandan las nuevas y se ignora
     *       la legada con warning CONFIG.</li>
     * </ol>
     */
    private int[] resolveFileCountRange() {
        boolean hasNew = config.has(KEY_STRICT_MIN_FILES) || config.has(KEY_STRICT_MAX_FILES);
        boolean hasLegacy = config.has(KEY_STRICT_TWO_FILES);
        if (hasNew) {
            int min = config.getInt(KEY_STRICT_MIN_FILES, 2);
            int max = config.getInt(KEY_STRICT_MAX_FILES, 3);
            if (hasLegacy) {
                report.addWarning("CONFIG",
                        KEY_STRICT_TWO_FILES + " es obsoleto y se ignora porque estan presentes "
                                + KEY_STRICT_MIN_FILES + "/" + KEY_STRICT_MAX_FILES + ".");
            }
            return new int[]{min, max};
        }
        if (hasLegacy) {
            boolean strictTwo = config.getBoolean(KEY_STRICT_TWO_FILES, true);
            report.addWarning("CONFIG",
                    KEY_STRICT_TWO_FILES + " es obsoleto; usa " + KEY_STRICT_MIN_FILES + "=2"
                            + " e " + KEY_STRICT_MAX_FILES + "=" + (strictTwo ? "2" : "3") + ".");
            // Ambos ramos mapean a [2,2]: si hay >2 ficheros, se trunca
            // con warning. La rama strictTwo=true lleva ademas un chequeo
            // previo en merge() que aborta antes de llegar al truncado,
            // preservando el contrato "exactamente 2" de la v2.1.0.
            return new int[]{2, 2};
        }
        // Defaults v2.2.0
        return new int[]{2, 3};
    }

    // ==================================================================
    //  Modos de fusion (privados: logica del orchestrator)
    // ==================================================================

    /**
     * Copia todas las hojas de los workbooks de entrada al libro resultado.
     * Usa {@link SheetCopier} para la copia bruta; la asignacion de nombres
     * (perfil + unicidad) es competencia de este orquestador.
     *
     * <p>v2.2.0: acepta N workbooks (2 o 3) en vez de dos fijos. Reordena
     * los workbooks segun {@code merge.profileOrder} (por defecto
     * {@code Cierre,Extraccion,Deuda}) para que la hoja Deuda caiga entre
     * Extraccion y la hoja Resultado que escribira MesSheetBuilder.
     * Los workbooks sin perfil resuelto se copian al final respetando el
     * orden alfabetico de los ficheros.</p>
     */
    private void mergeSheetsSeparate(List<Workbook> workbooks, List<File> files, Workbook result,
                                     List<FileProfileResolver.FileProfile> profiles) {
        int[] order = computeProfileOrder(profiles);
        for (int i : order) {
            copyAllSheetsFrom(workbooks.get(i), result, profiles.get(i),
                    "archivo " + (i + 1) + " (" + files.get(i).getName() + ")");
        }
    }

    /**
     * Devuelve los indices de {@code profiles} ordenados segun el orden
     * canonico de perfiles {@code merge.profileOrder}. Los que no encajan en
     * ese orden (incluyendo perfiles null) van al final por orden original.
     */
    private int[] computeProfileOrder(List<FileProfileResolver.FileProfile> profiles) {
        String orderProp = config.get("merge.profileOrder", "Cierre,Extraccion,Deuda");
        List<String> canonical = new ArrayList<>();
        for (String id : orderProp.split(",")) {
            String trimmed = id.trim();
            if (!trimmed.isEmpty()) canonical.add(trimmed);
        }
        int n = profiles.size();
        int[] out = new int[n];
        boolean[] used = new boolean[n];
        int cursor = 0;
        for (String id : canonical) {
            for (int i = 0; i < n; i++) {
                if (!used[i] && profiles.get(i) != null && id.equals(profiles.get(i).getId())) {
                    out[cursor] = i;
                    cursor++;
                    used[i] = true;
                }
            }
        }
        for (int i = 0; i < n; i++) {
            if (!used[i]) {
                out[cursor] = i;
                cursor++;
            }
        }
        return out;
    }

    private void copyAllSheetsFrom(Workbook source, Workbook result,
                                   FileProfileResolver.FileProfile profile, String label) {
        Set<Integer> asTextIdx = resolveAsTextIndexes(source, profile);
        Set<Integer> trimIdx = resolveTrimIndexes(source, profile);
        int firstDataRow0 = profile == null ? 0 : profile.getHeaderRow();
        int count = source.getNumberOfSheets();
        for (int i = 0; i < count; i++) {
            Sheet src = source.getSheetAt(i);
            String desiredName = resolveTargetName(profile, src.getSheetName(), i, count);
            String finalName = ensureUniqueSheetName(result, desiredName);

            Sheet target = result.createSheet(finalName);
            boolean applyAsText = profile != null
                    && i == profile.sheetIndex
                    && !asTextIdx.isEmpty();
            if (applyAsText) {
                sheetCopier.copySheet(src, target, result, asTextIdx, trimIdx, firstDataRow0);
            } else {
                sheetCopier.copySheet(src, target, result);
            }
            int rows = src.getLastRowNum() + 1;
            report.addSheet(finalName, rows);
            log.info("Hoja copiada desde {}: '{}'  ->  '{}' ({} filas)",
                    label, src.getSheetName(), finalName, rows);
        }
    }

    /**
     * Calcula los indices 0-based de las columnas que deben copiarse a STRING
     * para este perfil, y emite warnings {@code CABECERA} para las cabeceras
     * declaradas en {@code asText.columns} que no se han encontrado en la
     * hoja principal del perfil. Devuelve un set vacio si no hay perfil o
     * no hay declaradas.
     */
    private Set<Integer> resolveAsTextIndexes(Workbook source,
                                              FileProfileResolver.FileProfile profile) {
        if (profile == null) {
            return Collections.emptySet();
        }
        List<String> declared = profile.getAsTextColumns();
        if (declared.isEmpty()) {
            return Collections.emptySet();
        }
        Sheet mainSrc = source.getSheetAt(profile.sheetIndex);
        Set<Integer> resolved = profile.resolveAsTextColumnIndexes(mainSrc);
        if (resolved.size() < declared.size()) {
            warnMissingAsTextHeaders(mainSrc, profile, declared);
        }
        return resolved;
    }

    private void warnMissingAsTextHeaders(Sheet mainSrc,
                                          FileProfileResolver.FileProfile profile,
                                          List<String> declared) {
        Row header = mainSrc.getRow(profile.getHeaderRow() - 1);
        for (String col : declared) {
            int idx = header == null ? -1 : PoiUtils.findColumnIndex(header, col);
            if (idx < 0) {
                report.addWarning("CABECERA",
                        "Columna '" + col + "' declarada en 'profile."
                                + profile.getId() + ".asText.columns' no se encontro en '"
                                + mainSrc.getSheetName() + "'; se ignora.");
            }
        }
    }

    /**
     * v1.8.1: calcula los indices 0-based de columnas a trim()ar. Emite
     * warnings para: (a) cabeceras de {@code trim.columns} no encontradas
     * en la hoja, (b) cabeceras de {@code trim.columns} que NO estan
     * tambien en {@code asText.columns} (el trim solo aplica a la rama
     * STRING del cast; sin asText no hay cast, y el trim es no-op).
     */
    private Set<Integer> resolveTrimIndexes(Workbook source,
                                            FileProfileResolver.FileProfile profile) {
        if (profile == null) {
            return Collections.emptySet();
        }
        List<String> declared = profile.getTrimColumns();
        if (declared.isEmpty()) {
            return Collections.emptySet();
        }
        Sheet mainSrc = source.getSheetAt(profile.sheetIndex);
        Set<Integer> resolved = profile.resolveTrimColumnIndexes(mainSrc);
        if (resolved.size() < declared.size()) {
            warnMissingTrimHeaders(mainSrc, profile, declared);
        }
        // Aviso si una columna esta en trim pero no en asText
        List<String> asTextDeclared = profile.getAsTextColumns();
        for (String col : declared) {
            if (!containsIgnoreCase(asTextDeclared, col)) {
                report.addWarning("CONFIG",
                        "Columna '" + col + "' declarada en 'profile."
                                + profile.getId() + ".trim.columns' pero no en "
                                + "'asText.columns'; el trim solo aplica a valores "
                                + "casteados a STRING. Se ignorara para esta columna.");
            }
        }
        // Filtrar a solo las que tambien estan en asText (interseccion real)
        Set<Integer> asTextResolved = profile.resolveAsTextColumnIndexes(mainSrc);
        Set<Integer> effective = new HashSet<>(resolved);
        effective.retainAll(asTextResolved);
        return effective;
    }

    private void warnMissingTrimHeaders(Sheet mainSrc,
                                        FileProfileResolver.FileProfile profile,
                                        List<String> declared) {
        Row header = mainSrc.getRow(profile.getHeaderRow() - 1);
        for (String col : declared) {
            int idx = header == null ? -1 : PoiUtils.findColumnIndex(header, col);
            if (idx < 0) {
                report.addWarning("CABECERA",
                        "Columna '" + col + "' declarada en 'profile."
                                + profile.getId() + ".trim.columns' no se encontro en '"
                                + mainSrc.getSheetName() + "'; se ignora.");
            }
        }
    }

    private static boolean containsIgnoreCase(List<String> list, String value) {
        for (String s : list) {
            if (s.equalsIgnoreCase(value)) return true;
        }
        return false;
    }

    private String resolveTargetName(FileProfileResolver.FileProfile profile,
                                     String originalName, int index, int total) {
        if (profile == null) {
            return FileProfileResolver.safeSheetName(originalName);
        }
        if (total == 1 || index == 0) {
            return FileProfileResolver.safeSheetName(profile.getSheetName());
        }
        return FileProfileResolver.safeSheetName(profile.getSheetName() + "_" + originalName);
    }

    private String ensureUniqueSheetName(Workbook wb, String base) {
        if (wb.getSheet(base) == null) return base;
        String root = base;
        int suffix = 2;
        while (true) {
            String candidate = root + "_" + suffix;
            if (candidate.length() > FileProfileResolver.MAX_SHEET_NAME_LEN) {
                int excess = candidate.length() - FileProfileResolver.MAX_SHEET_NAME_LEN;
                root = root.substring(0, root.length() - excess);
                candidate = root + "_" + suffix;
            }
            if (wb.getSheet(candidate) == null) return candidate;
            suffix++;
        }
    }

    /**
     * Combina filas de la primera hoja de ambos libros en una sola hoja del
     * resultado. Usa {@link SheetCopier} para cada fila.
     */
    private void mergeAppendRows(Workbook wb1, Workbook wb2, Workbook result) {
        String resultSheetName = config.get("merge.resultSheetName", "Datos_Fusionados");
        int headerRows = config.getInt("merge.headerRows", 1);
        boolean skipHeader2 = config.getBoolean("merge.skipHeaderFromSecondFile", true);

        Sheet source1 = wb1.getSheetAt(0);
        Sheet source2 = wb2.getSheetAt(0);
        Sheet target = result.createSheet(resultSheetName);

        sheetCopier.copySheet(source1, target, result);

        int nextRowIdx = source1.getLastRowNum() + 1;
        int startRow2 = skipHeader2 ? headerRows : 0;

        Map<Integer, org.apache.poi.ss.usermodel.CellStyle> styleCache = new HashMap<>();

        for (int r = startRow2; r <= source2.getLastRowNum(); r++) {
            Row sourceRow = source2.getRow(r);
            if (sourceRow == null) {
                nextRowIdx++;
                continue;
            }
            Row targetRow = target.createRow(nextRowIdx);
            nextRowIdx++;
            sheetCopier.copyRow(sourceRow, targetRow, result, styleCache);
        }

        int maxCols = Math.max(sheetCopier.countColumns(source1),
                sheetCopier.countColumns(source2));
        for (int c = 0; c < maxCols; c++) {
            target.autoSizeColumn(c);
        }

        int totalRows = target.getLastRowNum() + 1;
        report.addSheet(resultSheetName, totalRows);
        log.info("[Merger] Filas fusionadas en la hoja: {} ({} filas)", resultSheetName, totalRows);
    }
}

package com.excelmerger;

import com.excelmerger.exception.MergeException;
import com.excelmerger.io.FileLockDetector;
import com.excelmerger.io.InputFileDetector;
import com.excelmerger.io.OutputManager;
import com.excelmerger.io.SheetCopier;

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
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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

    public void merge() {
        String inputDir = config.get("input.directory");
        String outputPath = config.get("output.file");
        MergeMode mode = MergeMode.valueOf(config.get("merge.mode", "SHEETS_SEPARATE"));
        boolean overwrite = config.getBoolean("output.overwrite", true);
        boolean backup = config.getBoolean("output.backup", false);
        boolean strictTwo = config.getBoolean("input.strictTwoFiles", true);

        // 1. Detectar y validar los ficheros de entrada
        List<File> excelFiles = inputDetector.findExcelFiles(inputDir);
        inputDetector.validateExcelFiles(excelFiles, strictTwo, report);

        File file1 = excelFiles.get(0);
        File file2 = excelFiles.get(1);

        // 2. Evitar que el archivo de salida coincida con uno de los de entrada
        Path outAbs = Paths.get(outputPath).toAbsolutePath().normalize();
        if (outAbs.equals(file1.toPath().toAbsolutePath().normalize())
                || outAbs.equals(file2.toPath().toAbsolutePath().normalize())) {
            throw new com.excelmerger.exception.InputValidationException(
                    "El archivo de salida no puede coincidir con uno de los archivos de entrada: " + outAbs);
        }

        // 3. Chequeos tempranos de lock (inputs y output)
        lockDetector.assertNotLocked(file1);
        lockDetector.assertNotLocked(file2);
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
        log.info("Archivo 1: {}", file1.getName());
        log.info("Archivo 2: {}", file2.getName());
        log.info("Modo de fusion: {}", mode);
        log.info("Archivo de salida: {}", outputPath);
        if (dryRun) {
            log.info("[DRY-RUN] Activado: el pipeline se ejecutara pero no se escribira el Excel final.");
        }

        // 5. Abrir archivos y ejecutar la fusion
        try (FileInputStream fis1 = lockDetector.openForRead(file1);
             FileInputStream fis2 = lockDetector.openForRead(file2);
             Workbook wb1 = WorkbookFactory.create(fis1);
             Workbook wb2 = WorkbookFactory.create(fis2);
             Workbook result = new XSSFWorkbook()) {

            // 5a. Identificar cada archivo por su CONTENIDO
            FileProfileResolver resolver = new FileProfileResolver(config);
            FileProfileResolver.FileProfile profile1 = null;
            FileProfileResolver.FileProfile profile2 = null;
            if (resolver.hasProfiles()) {
                profile1 = resolver.resolve(wb1, file1);
                profile2 = resolver.resolve(wb2, file2);

                if (profile1 == null) {
                    log.warn("'{}' no coincide con ningun perfil. Se usaran los nombres de hoja originales.",
                            file1.getName());
                    report.addWarning("PERFIL",
                            "'" + file1.getName() + "' no coincide con ningun perfil.");
                }
                if (profile2 == null) {
                    log.warn("'{}' no coincide con ningun perfil. Se usaran los nombres de hoja originales.",
                            file2.getName());
                    report.addWarning("PERFIL",
                            "'" + file2.getName() + "' no coincide con ningun perfil.");
                }
                if (profile1 != null && profile2 != null && profile1.getId().equals(profile2.getId())) {
                    log.warn("Ambos archivos coinciden con el mismo perfil '{}'. Se diferenciaran con un sufijo numerico.",
                            profile1.getId());
                    report.addWarning("PERFIL",
                            "Ambos archivos coinciden con el mismo perfil '" + profile1.getId() + "'.");
                }
            }

            // 5b. Fusionar segun modo
            if (mode == MergeMode.SHEETS_SEPARATE) {
                mergeSheetsSeparate(wb1, wb2, result, profile1, profile2);
            } else {
                mergeAppendRows(wb1, wb2, result);
            }

            // 5c. Hojas de lookup (tablas estaticas para VLOOKUP)
            new LookupSheetBuilder(config, report).buildAll(result);

            // 5d. Hoja MES (Extraccion + columnas calculadas)
            new MesSheetBuilder(config, report).build(result);

            // 5e. Hojas derivadas (formulas / agregaciones)
            new DerivedSheetBuilder(config, report).buildAll(result);

            // 5f. Hoja de avisos (opt-in con report.inExcel=true). Se construye
            //     despues del resto de builders para recoger TODOS los warnings
            //     acumulados (perfiles sin match, apps sin mapeo, cabeceras no
            //     encontradas, etc.). En dry-run se construye igual: queda en
            //     memoria pero no se escribe a disco.
            new AvisosSheetBuilder(config, report).build(result);

            // 5g. Escribir el resultado (omitido en dry-run)
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

    // ==================================================================
    //  Modos de fusion (privados: logica del orchestrator)
    // ==================================================================

    /**
     * Copia todas las hojas de wb1 y wb2 al libro resultado. Usa
     * {@link SheetCopier} para la copia bruta; la asignacion de nombres
     * (perfil + unicidad) es competencia de este orquestador.
     */
    private void mergeSheetsSeparate(Workbook wb1, Workbook wb2, Workbook result,
                                     FileProfileResolver.FileProfile profile1,
                                     FileProfileResolver.FileProfile profile2) {
        copyAllSheetsFrom(wb1, result, profile1, "archivo 1");
        copyAllSheetsFrom(wb2, result, profile2, "archivo 2");
    }

    private void copyAllSheetsFrom(Workbook source, Workbook result,
                                   FileProfileResolver.FileProfile profile, String label) {
        int count = source.getNumberOfSheets();
        for (int i = 0; i < count; i++) {
            Sheet src = source.getSheetAt(i);
            String desiredName = resolveTargetName(profile, src.getSheetName(), i, count);
            String finalName = ensureUniqueSheetName(result, desiredName);

            Sheet target = result.createSheet(finalName);
            sheetCopier.copySheet(src, target, result);
            int rows = src.getLastRowNum() + 1;
            report.addSheet(finalName, rows);
            log.info("Hoja copiada desde {}: '{}'  ->  '{}' ({} filas)",
                    label, src.getSheetName(), finalName, rows);
        }
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
            Row targetRow = target.createRow(nextRowIdx++);
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

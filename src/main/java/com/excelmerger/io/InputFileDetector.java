package com.excelmerger.io;

import com.excelmerger.RunReport;
import com.excelmerger.exception.InputValidationException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.stream.Collectors;

/**
 * Detecta y valida los ficheros Excel presentes en un directorio de entrada.
 * Ignora ficheros temporales de Excel (~$...) y ocultos; ordena alfabeticamente.
 * Cualquier inconsistencia (directorio ausente, menos de 2 ficheros, modo
 * estricto con mas de 2) se traduce en {@link InputValidationException}.
 */
public final class InputFileDetector {

    private static final Logger log = LoggerFactory.getLogger(InputFileDetector.class);
    private static final List<String> EXCEL_EXTENSIONS = Arrays.asList(".xlsx", ".xls");

    /**
     * Busca todos los archivos con extension .xlsx o .xls en el directorio
     * dado, ordenados alfabeticamente por nombre en minusculas.
     *
     * @throws InputValidationException si el directorio no existe, no es un
     *         directorio, o no se puede listar su contenido.
     */
    public List<File> findExcelFiles(String inputDir) {
        Path dir = Paths.get(inputDir);
        if (!Files.exists(dir)) {
            throw new InputValidationException("El directorio de entrada no existe: " + dir.toAbsolutePath());
        }
        if (!Files.isDirectory(dir)) {
            throw new InputValidationException("La ruta de entrada no es un directorio: " + dir.toAbsolutePath());
        }

        File[] files = dir.toFile().listFiles();
        if (files == null) {
            throw new InputValidationException("No se pudo leer el directorio: " + dir.toAbsolutePath());
        }

        return Arrays.stream(files)
                .filter(File::isFile)
                .filter(f -> !f.getName().startsWith("~$"))
                .filter(f -> !f.getName().startsWith("."))
                .filter(f -> hasExcelExtension(f.getName()))
                .sorted(Comparator.comparing(f -> f.getName().toLowerCase(Locale.ROOT)))
                .collect(Collectors.toList());
    }

    /**
     * Verifica que haya al menos {@code minFiles} ficheros. Si hay mas de
     * {@code maxFiles}, trunca al primer {@code maxFiles} por orden
     * alfabetico y emite un warning {@code CONFIG} en el {@link RunReport}.
     *
     * <p>Devuelve la lista, posiblemente truncada. Los llamadores deben usar
     * la lista devuelta.</p>
     *
     * <p>v2.2.0: sustituye la version anterior con {@code strictTwo}
     * booleano. La retrocompat con la clave legada
     * {@code input.strictTwoFiles} se resuelve en el llamador, mapeando
     * a min/max.</p>
     *
     * @throws InputValidationException si {@code files.size() < minFiles}
     *         o si los parametros son incoherentes.
     */
    public List<File> validateExcelFiles(List<File> files, int minFiles, int maxFiles,
                                         RunReport report) {
        if (minFiles < 1 || maxFiles < minFiles) {
            throw new InputValidationException(
                    "Configuracion incoherente: input.strictMinFiles=" + minFiles
                            + ", input.strictMaxFiles=" + maxFiles
                            + " (se requiere 1 <= min <= max).");
        }
        if (files.size() < minFiles) {
            throw new InputValidationException(
                    "Se han encontrado " + files.size() + " archivos Excel en el directorio de entrada. "
                            + "Se necesitan al menos " + minFiles + ". Archivos detectados: " + joinNames(files));
        }
        if (files.size() > maxFiles) {
            log.warn("Se encontraron {} archivos Excel, se usaran los {} primeros por orden alfabetico.",
                    files.size(), maxFiles);
            report.addWarning("CONFIG",
                    "Se encontraron " + files.size() + " archivos Excel; se usaron los "
                            + maxFiles + " primeros.");
            return new java.util.ArrayList<>(files.subList(0, maxFiles));
        }
        return files;
    }

    public static boolean hasExcelExtension(String fileName) {
        String lower = fileName.toLowerCase(Locale.ROOT);
        return EXCEL_EXTENSIONS.stream().anyMatch(lower::endsWith);
    }

    private static String joinNames(List<File> files) {
        return files.stream().map(File::getName).collect(Collectors.joining(", "));
    }
}

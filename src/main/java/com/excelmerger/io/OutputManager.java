package com.excelmerger.io;

import com.excelmerger.exception.OutputException;

import org.apache.poi.ss.usermodel.Workbook;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

/**
 * Gestion del fichero de salida: detecta locks, aplica politica de overwrite
 * y backup, y escribe el workbook final. Todos los fallos se traducen en
 * {@link OutputException}.
 */
public final class OutputManager {

    private static final Logger log = LoggerFactory.getLogger(OutputManager.class);
    private static final DateTimeFormatter BACKUP_TS =
            DateTimeFormatter.ofPattern("yyyy-MM-dd_HHmmss");

    /**
     * Chequeo previo: si existe un lock {@code ~$...} para el fichero de
     * salida, aborta con mensaje claro antes de tocar los inputs.
     */
    public void assertOutputWritable(String outputPath) {
        Path out = Paths.get(outputPath);
        Path fileName = out.getFileName();
        if (fileName == null) {
            throw new OutputException("output.file no apunta a un fichero: '" + outputPath + "'");
        }
        Path parent = out.getParent();
        if (parent == null) parent = Paths.get(".");
        Path lock = parent.resolve("~$" + fileName.toString());
        if (Files.exists(lock)) {
            throw new OutputException("Cierra '" + fileName
                    + "' antes de ejecutar (parece abierto en Excel: existe el lock '"
                    + lock.getFileName() + "').");
        }
    }

    /**
     * Si el output ya existe:
     * <ul>
     *   <li>con {@code backup=true}, lo mueve a
     *       {@code <parent>/history/<base>_YYYY-MM-DD_HHMMSS.<ext>}</li>
     *   <li>si {@code overwrite=false} y no hay backup, aborta con {@link OutputException}</li>
     * </ul>
     * Ademas crea el directorio padre del output si no existe.
     */
    public void prepareOutputFile(String outputPath, boolean overwrite, boolean backup) {
        Path out = Paths.get(outputPath);
        if (Files.exists(out)) {
            if (backup) {
                Path backupPath = backupOutput(out);
                log.info("Backup del output anterior: {}", backupPath.toAbsolutePath());
            } else if (!overwrite) {
                throw new OutputException("El archivo de salida ya existe y output.overwrite=false: " + outputPath);
            }
        }
        Path parent = out.getParent();
        if (parent != null && !Files.exists(parent)) {
            try {
                Files.createDirectories(parent);
            } catch (IOException e) {
                throw new OutputException("No se pudo crear el directorio de salida '" + parent
                        + "': " + e.getMessage(), e);
            }
        }
    }

    /**
     * Mueve el output existente a {@code <parent>/history/<basename>_<ts>.<ext>} y
     * devuelve la ruta destino. Crea la carpeta 'history' si no existe.
     */
    public Path backupOutput(Path out) {
        Path outName = out.getFileName();
        if (outName == null) {
            throw new OutputException("No se puede hacer backup de un path sin nombre de fichero: '" + out + "'");
        }
        String fileName = outName.toString();
        int dot = fileName.lastIndexOf('.');
        String base = (dot > 0) ? fileName.substring(0, dot) : fileName;
        String ext = (dot > 0) ? fileName.substring(dot) : "";
        String ts = LocalDateTime.now().format(BACKUP_TS);

        Path parent = out.getParent();
        if (parent == null) parent = Paths.get(".");
        Path historyDir = parent.resolve("history");
        try {
            Files.createDirectories(historyDir);
        } catch (IOException e) {
            throw new OutputException("No se pudo crear la carpeta de history para backup: "
                    + e.getMessage(), e);
        }

        // En caso extremo de colision (mismo segundo), anade un sufijo numerico.
        Path target = historyDir.resolve(base + "_" + ts + ext);
        int suffix = 2;
        while (Files.exists(target)) {
            target = historyDir.resolve(base + "_" + ts + "_" + suffix + ext);
            suffix++;
        }
        try {
            Files.move(out, target, StandardCopyOption.ATOMIC_MOVE);
        } catch (IOException atomicFailed) {
            try {
                Files.move(out, target, StandardCopyOption.REPLACE_EXISTING);
            } catch (IOException e) {
                e.addSuppressed(atomicFailed);
                throw new OutputException("No se pudo mover el output a history: "
                        + e.getMessage(), e);
            }
        }
        return target;
    }

    /**
     * Escribe el libro resultado. Traduce fallos de escritura por bloqueo
     * ({@code ~$...}) a {@link OutputException} con mensaje claro.
     */
    public void writeResult(Workbook result, String outputPath) {
        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            result.write(fos);
        } catch (FileNotFoundException e) {
            if (FileLockDetector.looksLikeLocked(e)) {
                throw new OutputException("Cierra '" + Paths.get(outputPath).getFileName()
                        + "' antes de ejecutar (parece abierto en Excel).", e);
            }
            throw new OutputException("No se pudo escribir el fichero de salida '" + outputPath
                    + "': " + e.getMessage(), e);
        } catch (IOException e) {
            if (FileLockDetector.looksLikeLocked(e)) {
                throw new OutputException("Cierra '" + Paths.get(outputPath).getFileName()
                        + "' antes de ejecutar (parece abierto en Excel).", e);
            }
            throw new OutputException("No se pudo escribir el fichero de salida '" + outputPath
                    + "': " + e.getMessage(), e);
        }
    }
}

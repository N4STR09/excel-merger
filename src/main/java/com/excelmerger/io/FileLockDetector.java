package com.excelmerger.io;

import com.excelmerger.exception.InputValidationException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * Detecta si un fichero esta "bloqueado" en el sentido Excel: abierto por
 * otro proceso (en Windows suele dejar un fichero lock {@code ~$nombre} y
 * devolver errores de "used by another process" al intentar abrir).
 *
 * <p>Cualquier problema detectado para ficheros de <b>entrada</b> se traduce
 * en {@link InputValidationException}. Para el fichero de <b>salida</b>, la
 * clase hermana {@link OutputManager} se encarga.</p>
 */
public final class FileLockDetector {

    /**
     * Comprueba que un archivo de entrada es accesible para lectura.
     * Si esta abierto en Excel (lock file ~$..., o error de bloqueo del OS),
     * lanza {@link InputValidationException} con mensaje claro.
     */
    public void assertNotLocked(File f) {
        // 1. Fichero lock de Excel (~$<nombre>) en la misma carpeta
        File parent = f.getParentFile();
        if (parent != null) {
            File lock = new File(parent, "~$" + f.getName());
            if (lock.exists()) {
                throw new InputValidationException("Cierra '" + f.getName()
                        + "' antes de ejecutar (parece abierto en Excel: existe el lock '"
                        + lock.getName() + "').");
            }
        }
        // 2. Apertura de prueba
        try (FileInputStream probe = new FileInputStream(f)) {
            probe.read();
        } catch (FileNotFoundException e) {
            if (looksLikeLocked(e)) {
                throw new InputValidationException("Cierra '" + f.getName()
                        + "' antes de ejecutar (parece abierto en Excel).", e);
            }
            throw new InputValidationException("No se puede leer '" + f.getName() + "': " + e.getMessage(), e);
        } catch (IOException e) {
            if (looksLikeLocked(e)) {
                throw new InputValidationException("Cierra '" + f.getName()
                        + "' antes de ejecutar (parece abierto en Excel).", e);
            }
            throw new InputValidationException("No se puede leer '" + f.getName() + "': " + e.getMessage(), e);
        }
    }

    /**
     * Abre un {@link FileInputStream} traduciendo errores de bloqueo a mensaje
     * claro. Se usa como fallback por si el fichero se bloquea en la ventana
     * entre {@link #assertNotLocked} y la apertura real.
     */
    public FileInputStream openForRead(File f) {
        try {
            return new FileInputStream(f);
        } catch (FileNotFoundException e) {
            if (looksLikeLocked(e)) {
                throw new InputValidationException("Cierra '" + f.getName()
                        + "' antes de ejecutar (parece abierto en Excel).", e);
            }
            throw new InputValidationException("No se puede abrir '" + f.getName() + "': " + e.getMessage(), e);
        }
    }

    /**
     * Heuristica para identificar errores de fichero bloqueado en distintas
     * plataformas (Windows y Linux producen mensajes diferentes).
     */
    public static boolean looksLikeLocked(Throwable e) {
        Throwable cur = e;
        while (cur != null) {
            String msg = cur.getMessage();
            if (msg != null) {
                String low = msg.toLowerCase();
                if (low.contains("being used by another process")
                        || low.contains("used by another process")
                        || low.contains("the process cannot access")
                        || low.contains("access is denied")
                        || low.contains("permission denied")
                        || low.contains("sharing violation")) {
                    return true;
                }
            }
            cur = cur.getCause();
        }
        return false;
    }
}

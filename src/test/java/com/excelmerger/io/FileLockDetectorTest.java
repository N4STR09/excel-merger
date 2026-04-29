package com.excelmerger.io;

import com.excelmerger.exception.InputValidationException;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledOnOs;
import org.junit.jupiter.api.condition.OS;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.channels.FileChannel;
import java.nio.channels.FileLock;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.PosixFilePermission;
import java.nio.file.attribute.PosixFilePermissions;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Tests de {@link FileLockDetector}: verifica deteccion de lock {@code ~$...},
 * traduccion de errores de bloqueo a {@link InputValidationException}, y la
 * heuristica {@link FileLockDetector#looksLikeLocked(Throwable)} con sus seis
 * patrones de mensaje y la cause-chain.
 *
 * <p>Los tests que requieren un fichero realmente bloqueado a nivel SO se
 * marcan con {@link EnabledOnOs}: en Windows se usa {@link FileChannel#tryLock()}
 * sobre un {@link RandomAccessFile}; en Linux/Mac se quitan los permisos POSIX
 * de lectura.</p>
 */
class FileLockDetectorTest {

    private final FileLockDetector detector = new FileLockDetector();

    // -------------------------------------------------------------------------
    // assertNotLocked / openForRead - happy path
    // -------------------------------------------------------------------------

    @Test
    void assertNotLocked_ficheroValido_noLanza(@TempDir Path tmp) throws IOException {
        Path f = tmp.resolve("input.xlsx");
        Files.writeString(f, "dummy", StandardCharsets.UTF_8);

        // Sin excepcion = OK.
        detector.assertNotLocked(f.toFile());
    }

    @Test
    void openForRead_ficheroValido_devuelveStreamLegible(@TempDir Path tmp) throws IOException {
        Path f = tmp.resolve("input.xlsx");
        Files.writeString(f, "hello", StandardCharsets.UTF_8);

        try (FileInputStream fis = detector.openForRead(f.toFile())) {
            assertThat(fis).isNotNull();
            assertThat(fis.read()).isEqualTo((int) 'h');
        }
    }

    // -------------------------------------------------------------------------
    // Lock file ~$ (camino sin abrir el fichero real)
    // -------------------------------------------------------------------------

    @Test
    void assertNotLocked_existeLockTilde_lanzaInputValidationConMensajeCierra(@TempDir Path tmp) throws IOException {
        Path f = tmp.resolve("datos.xlsx");
        Files.writeString(f, "dummy", StandardCharsets.UTF_8);
        Path lock = tmp.resolve("~$datos.xlsx");
        Files.writeString(lock, "", StandardCharsets.UTF_8);

        assertThatThrownBy(() -> detector.assertNotLocked(f.toFile()))
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("Cierra 'datos.xlsx'")
                .hasMessageContaining("~$datos.xlsx");
    }

    // -------------------------------------------------------------------------
    // Catch genericos: mensaje de error que NO matchea looksLikeLocked
    // -------------------------------------------------------------------------

    @Test
    void assertNotLocked_ficheroNoExiste_lanzaInputValidationGenerica(@TempDir Path tmp) {
        File missing = tmp.resolve("no-existe.xlsx").toFile();

        assertThatThrownBy(() -> detector.assertNotLocked(missing))
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("No se puede leer 'no-existe.xlsx'")
                .hasCauseInstanceOf(IOException.class);
    }

    @Test
    void openForRead_ficheroNoExiste_lanzaInputValidationGenerica(@TempDir Path tmp) {
        File missing = tmp.resolve("no-existe.xlsx").toFile();

        assertThatThrownBy(() -> detector.openForRead(missing))
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("No se puede abrir 'no-existe.xlsx'")
                .hasCauseInstanceOf(IOException.class);
    }

    // -------------------------------------------------------------------------
    // Catch "locked-like": el catch matchea looksLikeLocked y reescribe el msg.
    // Estos tests son SO-especificos porque crear un fichero realmente bloqueado
    // requiere mecanismos de OS distintos en Windows vs POSIX.
    // -------------------------------------------------------------------------

    @Test
    @EnabledOnOs(OS.WINDOWS)
    void assertNotLocked_ficheroBloqueadoEnWindows_lanzaInputValidationConMensajeCierra(@TempDir Path tmp)
            throws IOException {
        Path f = tmp.resolve("locked.xlsx");
        Files.writeString(f, "dummy", StandardCharsets.UTF_8);

        // En Windows un FileChannel con lock exclusivo sobre todo el rango
        // produce que un FileInputStream concurrente lance FileNotFoundException
        // con el mensaje "being used by another process".
        try (RandomAccessFile raf = new RandomAccessFile(f.toFile(), "rw");
             FileChannel ch = raf.getChannel();
             FileLock ignored = ch.lock()) {

            assertThatThrownBy(() -> detector.assertNotLocked(f.toFile()))
                    .isInstanceOf(InputValidationException.class)
                    .hasMessageContaining("Cierra 'locked.xlsx'")
                    .hasMessageContaining("parece abierto en Excel");
        }
    }

    @Test
    @EnabledOnOs({OS.LINUX, OS.MAC})
    void assertNotLocked_ficheroSinPermisosDeLecturaEnLinux_lanzaInputValidationConMensajeCierra(@TempDir Path tmp)
            throws IOException {
        // En POSIX, root ignora los permisos -> el test no es ejecutable como root.
        // Si lo es, salimos sin asertar (no falla).
        if ("root".equals(System.getProperty("user.name"))) return;

        Path f = tmp.resolve("locked.xlsx");
        Files.writeString(f, "dummy", StandardCharsets.UTF_8);

        // Quitar todos los permisos POSIX -> FileInputStream falla con
        // "Permission denied", que matchea looksLikeLocked.
        Set<PosixFilePermission> empty = PosixFilePermissions.fromString("---------");
        Files.setPosixFilePermissions(f, empty);

        try {
            assertThatThrownBy(() -> detector.assertNotLocked(f.toFile()))
                    .isInstanceOf(InputValidationException.class)
                    .hasMessageContaining("Cierra 'locked.xlsx'")
                    .hasMessageContaining("parece abierto en Excel");
        } finally {
            // Restaurar permisos para que @TempDir pueda limpiar.
            Files.setPosixFilePermissions(f, PosixFilePermissions.fromString("rw-------"));
        }
    }

    // -------------------------------------------------------------------------
    // looksLikeLocked - heuristica de mensajes
    // -------------------------------------------------------------------------

    @Test
    void looksLikeLocked_excepcionConMensajeNull_devuelveFalse() {
        // Throwable cuya getMessage() devuelve null -> recorre la cadena sin
        // matchear nada y termina devolviendo false.
        IOException sinMensaje = new IOException();

        assertThat(FileLockDetector.looksLikeLocked(sinMensaje)).isFalse();
    }

    @Test
    void looksLikeLocked_mensajeNoRelacionado_devuelveFalse() {
        IOException otro = new IOException("disk full");

        assertThat(FileLockDetector.looksLikeLocked(otro)).isFalse();
    }

    @Test
    void looksLikeLocked_mensajesQueIndicanLock_devuelvenTrue() {
        // Cubre los ocho sub-branches del OR en looksLikeLocked: seis patrones
        // en ingles (mensajes de JDK Windows/Linux/Mac en locale en) y dos
        // patrones en espanol (Windows con locale es).
        assertThat(FileLockDetector.looksLikeLocked(new IOException("File is being used by another process")))
                .isTrue();
        assertThat(FileLockDetector.looksLikeLocked(new IOException("used by another process")))
                .isTrue();
        assertThat(FileLockDetector.looksLikeLocked(new IOException("The process cannot access the file")))
                .isTrue();
        assertThat(FileLockDetector.looksLikeLocked(new IOException("Access is denied")))
                .isTrue();
        assertThat(FileLockDetector.looksLikeLocked(new IOException("Permission denied")))
                .isTrue();
        assertThat(FileLockDetector.looksLikeLocked(new IOException("Sharing violation on path")))
                .isTrue();
        // Mensaje real visto en Windows ES: "El proceso no tiene acceso al
        // archivo porque otro proceso tiene bloqueada una parte del archivo".
        // Para cubrir la cláusula "otro proceso tiene bloqueada" usamos un
        // mensaje que la contenga pero NO contenga "el proceso no tiene
        // acceso", de modo que el short-circuit del OR llegue hasta esa
        // clausula y la evalue a true.
        assertThat(FileLockDetector.looksLikeLocked(new IOException("otro proceso tiene bloqueada una parte")))
                .isTrue();
        // Y para la octava clausula: "el proceso no tiene acceso" sin
        // que aparezca "otro proceso tiene bloqueada" antes en el OR.
        assertThat(FileLockDetector.looksLikeLocked(new IOException("el proceso no tiene acceso al archivo")))
                .isTrue();
    }

    @Test
    void looksLikeLocked_mensajeEnCausaAnidada_devuelveTrue() {
        // El mensaje "locked-like" esta en la causa, no en la excepcion top-level.
        // Esto fuerza el while a recorrer cur = cur.getCause().
        IOException root = new IOException("access is denied");
        IOException wrapper = new IOException("wrapper sin pistas", root);
        IOException outer = new IOException("otro wrapper inocuo", wrapper);

        assertThat(FileLockDetector.looksLikeLocked(outer)).isTrue();
    }
}

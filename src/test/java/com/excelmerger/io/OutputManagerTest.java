package com.excelmerger.io;

import com.excelmerger.exception.OutputException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledOnOs;
import org.junit.jupiter.api.condition.OS;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.PosixFilePermission;
import java.nio.file.attribute.PosixFilePermissions;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Tests de {@link OutputManager}: cubre {@code assertOutputWritable},
 * {@code prepareOutputFile} (con y sin overwrite/backup), {@code backupOutput}
 * (extension/sin extension/colision/path raiz/error de creacion del directorio
 * history) y {@code writeResult} (happy path, error generico, error
 * "locked-like").
 *
 * <p>Algunos tests son SO-especificos: la rama "locked-like" en {@code writeResult}
 * y la rama {@code getFileName() == null} se cubren con {@link EnabledOnOs}
 * porque requieren mecanismos del sistema de ficheros que difieren entre
 * Windows y POSIX.</p>
 */
class OutputManagerTest {

    private final OutputManager om = new OutputManager();

    /**
     * Lista de paths que pueden quedar fuera de {@code @TempDir} y deben
     * limpiarse manualmente al final de cada test (ej: tests que usan paths
     * relativos sin padre y crean {@code ./history/}).
     */
    private final List<Path> manualCleanupPaths = new ArrayList<>();

    @AfterEach
    void cleanupManualPaths() throws IOException {
        for (Path p : manualCleanupPaths) {
            deleteRecursively(p);
        }
        manualCleanupPaths.clear();
    }

    private static void deleteRecursively(Path p) throws IOException {
        if (!Files.exists(p)) return;
        if (Files.isDirectory(p)) {
            try (var stream = Files.list(p)) {
                for (Path child : stream.toList()) {
                    deleteRecursively(child);
                }
            }
        }
        Files.deleteIfExists(p);
    }

    // -------------------------------------------------------------------------
    // assertOutputWritable
    // -------------------------------------------------------------------------

    @Test
    void assertOutputWritable_sinLock_noLanza(@TempDir Path tmp) {
        Path out = tmp.resolve("salida.xlsx");

        om.assertOutputWritable(out.toString());
    }

    @Test
    void assertOutputWritable_conLockTilde_lanzaOutputExceptionConMensajeCierra(@TempDir Path tmp) throws IOException {
        Path out = tmp.resolve("salida.xlsx");
        Path lock = tmp.resolve("~$salida.xlsx");
        Files.writeString(lock, "", StandardCharsets.UTF_8);

        assertThatThrownBy(() -> om.assertOutputWritable(out.toString()))
                .isInstanceOf(OutputException.class)
                .hasMessageContaining("Cierra 'salida.xlsx'")
                .hasMessageContaining("~$salida.xlsx");
    }

    @Test
    @EnabledOnOs({OS.LINUX, OS.MAC})
    void assertOutputWritable_pathSinNombreDeFichero_lanzaOutputException() {
        // En POSIX, Paths.get("/").getFileName() == null. En Windows
        // Paths.get("C:\\").getFileName() tambien es null pero el test se
        // duplicaria sin valor; cubrimos la rama en Linux/Mac.
        assertThatThrownBy(() -> om.assertOutputWritable("/"))
                .isInstanceOf(OutputException.class)
                .hasMessageContaining("output.file no apunta a un fichero");
    }

    @Test
    void assertOutputWritable_pathSinPadre_usaDirectorioActualParaLock(@TempDir Path tmp) {
        // Path relativo sin separador -> getParent() == null, OutputManager usa
        // Paths.get(".") como directorio. No deberia lanzar (no existe lock).
        // Ejecutamos desde un directorio limpio (tmp) cambiando al test scope:
        // basta con verificar que no existe '~$nombre-improbable.xlsx' en cwd.
        Path possibleLock = Paths.get(".").resolve("~$nombre-improbable-test-12345.xlsx");
        // Pre-condicion: nadie ha dejado un lock con ese nombre en el cwd.
        assertThat(Files.exists(possibleLock)).isFalse();

        // No lanza porque no existe lock en cwd.
        om.assertOutputWritable("nombre-improbable-test-12345.xlsx");
    }

    // -------------------------------------------------------------------------
    // prepareOutputFile
    // -------------------------------------------------------------------------

    @Test
    void prepareOutputFile_outputNoExisteYDirectorioPadreNoExiste_creaPadreYNoLanza(@TempDir Path tmp) {
        Path padreNuevo = tmp.resolve("subdir-nueva");
        Path out = padreNuevo.resolve("salida.xlsx");

        om.prepareOutputFile(out.toString(), false, false);

        assertThat(Files.isDirectory(padreNuevo)).isTrue();
        assertThat(Files.exists(out)).isFalse();
    }

    @Test
    void prepareOutputFile_outputNoExisteYPadreYaExiste_noLanza(@TempDir Path tmp) {
        Path out = tmp.resolve("salida.xlsx");

        om.prepareOutputFile(out.toString(), false, false);

        assertThat(Files.exists(out)).isFalse();
    }

    @Test
    void prepareOutputFile_outputExisteSinOverwriteSinBackup_lanzaOutputException(@TempDir Path tmp) throws IOException {
        Path out = tmp.resolve("salida.xlsx");
        Files.writeString(out, "viejo", StandardCharsets.UTF_8);

        assertThatThrownBy(() -> om.prepareOutputFile(out.toString(), false, false))
                .isInstanceOf(OutputException.class)
                .hasMessageContaining("ya existe")
                .hasMessageContaining("output.overwrite=false");
    }

    @Test
    void prepareOutputFile_outputExisteConOverwriteSinBackup_noLanza(@TempDir Path tmp) throws IOException {
        Path out = tmp.resolve("salida.xlsx");
        Files.writeString(out, "viejo", StandardCharsets.UTF_8);

        // overwrite=true -> no aborta, pero tampoco mueve a history.
        om.prepareOutputFile(out.toString(), true, false);

        assertThat(Files.exists(out)).isTrue();
        assertThat(Files.exists(tmp.resolve("history"))).isFalse();
    }

    @Test
    void prepareOutputFile_outputExisteConBackup_mueveAHistory(@TempDir Path tmp) throws IOException {
        Path out = tmp.resolve("salida.xlsx");
        Files.writeString(out, "viejo", StandardCharsets.UTF_8);

        om.prepareOutputFile(out.toString(), false, true);

        assertThat(Files.exists(out)).isFalse();
        Path history = tmp.resolve("history");
        assertThat(Files.isDirectory(history)).isTrue();
        try (var stream = Files.list(history)) {
            assertThat(stream.toList())
                    .hasSize(1)
                    .allMatch(p -> p.getFileName().toString().startsWith("salida_")
                            && p.getFileName().toString().endsWith(".xlsx"));
        }
    }

    @Test
    void prepareOutputFile_padreIntermedioEsFicheroNoCarpeta_lanzaOutputException(@TempDir Path tmp) throws IOException {
        // Crear un fichero llamado 'bloqueador' en el path intermedio. El parent
        // del output ('tmp/bloqueador/sub') no existe, asi que prepareOutputFile
        // entra al try-catch y Files.createDirectories falla con
        // FileAlreadyExistsException porque 'bloqueador' es fichero, no directorio.
        Path bloqueador = tmp.resolve("bloqueador");
        Files.writeString(bloqueador, "soy un fichero", StandardCharsets.UTF_8);
        Path out = bloqueador.resolve("sub").resolve("dentro.xlsx");

        assertThatThrownBy(() -> om.prepareOutputFile(out.toString(), true, false))
                .isInstanceOf(OutputException.class)
                .hasMessageContaining("No se pudo crear el directorio de salida");
    }

    // -------------------------------------------------------------------------
    // backupOutput
    // -------------------------------------------------------------------------

    @Test
    void backupOutput_ficheroConExtension_mueveAHistoryConTimestamp(@TempDir Path tmp) throws IOException {
        Path out = tmp.resolve("salida.xlsx");
        Files.writeString(out, "viejo", StandardCharsets.UTF_8);

        Path destino = om.backupOutput(out);

        assertThat(Files.exists(out)).isFalse();
        assertThat(Files.exists(destino)).isTrue();
        assertThat(destino.getParent().getFileName()).hasToString("history");
        assertThat(destino.getFileName().toString())
                .startsWith("salida_")
                .endsWith(".xlsx");
    }

    @Test
    void backupOutput_ficheroSinExtension_mueveSinSufijoExt(@TempDir Path tmp) throws IOException {
        Path out = tmp.resolve("salidaSinExt");
        Files.writeString(out, "viejo", StandardCharsets.UTF_8);

        Path destino = om.backupOutput(out);

        assertThat(Files.exists(destino)).isTrue();
        assertThat(destino.getFileName().toString())
                .startsWith("salidaSinExt_")
                .doesNotContain(".");  // no extension
    }

    @Test
    void backupOutput_ficheroConPuntoInicial_consideraSinExtension(@TempDir Path tmp) throws IOException {
        // ".hidden": lastIndexOf('.') == 0, condicion (dot > 0) es falsa, asi que
        // base = ".hidden" entera y ext = "". Cubre ese branch del ternario.
        Path out = tmp.resolve(".hidden");
        Files.writeString(out, "viejo", StandardCharsets.UTF_8);

        Path destino = om.backupOutput(out);

        assertThat(Files.exists(destino)).isTrue();
        assertThat(destino.getFileName().toString()).startsWith(".hidden_");
    }

    @Test
    void backupOutput_pathSinPadre_usaDirectorioActualParaHistory() throws IOException {
        // Path relativo "out.xlsx" -> getParent() == null -> usa Paths.get(".").
        // Creamos el fichero en cwd, lo movemos, y limpiamos en @AfterEach.
        Path out = Paths.get("out-test-pathSinPadre.xlsx");
        Path historyDir = Paths.get("history");
        manualCleanupPaths.add(out);
        manualCleanupPaths.add(historyDir);

        Files.writeString(out, "viejo", StandardCharsets.UTF_8);

        Path destino = om.backupOutput(out);

        assertThat(Files.exists(out)).isFalse();
        assertThat(Files.exists(destino)).isTrue();
        assertThat(destino.getParent().getFileName()).hasToString("history");
    }

    @Test
    @EnabledOnOs({OS.LINUX, OS.MAC})
    void backupOutput_pathSinNombreDeFichero_lanzaOutputException() {
        // Paths.get("/") tiene getFileName() == null en POSIX.
        Path raiz = Paths.get("/");

        assertThatThrownBy(() -> om.backupOutput(raiz))
                .isInstanceOf(OutputException.class)
                .hasMessageContaining("path sin nombre de fichero");
    }

    @Test
    void backupOutput_historyEsFicheroNoCarpeta_lanzaOutputException(@TempDir Path tmp) throws IOException {
        // Pre-crear 'history' como FICHERO (no directorio): createDirectories falla.
        Path out = tmp.resolve("salida.xlsx");
        Files.writeString(out, "viejo", StandardCharsets.UTF_8);
        Path historyAsFile = tmp.resolve("history");
        Files.writeString(historyAsFile, "soy un fichero", StandardCharsets.UTF_8);

        assertThatThrownBy(() -> om.backupOutput(out))
                .isInstanceOf(OutputException.class)
                .hasMessageContaining("No se pudo crear la carpeta de history");
    }

    @Test
    void backupOutput_colisionDeTimestamp_aplicaSufijoNumerico(@TempDir Path tmp) throws IOException {
        // Inyectamos un reloj fijo: ambas llamadas a backupOutput formatean el
        // mismo ts, asi que la segunda colisiona y debe llevar sufijo "_2".
        OutputManager omFijo = new OutputManager();
        omFijo.clockForTesting = () -> LocalDateTime.of(2025, 1, 15, 12, 0, 0);

        // Primera llamada
        Path out1 = tmp.resolve("salida.xlsx");
        Files.writeString(out1, "v1", StandardCharsets.UTF_8);
        Path destino1 = omFijo.backupOutput(out1);

        // Segunda llamada: recreamos el output original (mismo nombre, mismo
        // padre) y volvemos a llamar. El destino chocara con destino1.
        Files.writeString(out1, "v2", StandardCharsets.UTF_8);
        Path destino2 = omFijo.backupOutput(out1);

        assertThat(destino1.getFileName().toString())
                .isEqualTo("salida_2025-01-15_120000.xlsx");
        assertThat(destino2.getFileName().toString())
                .isEqualTo("salida_2025-01-15_120000_2.xlsx");
        assertThat(Files.exists(destino1)).isTrue();
        assertThat(Files.exists(destino2)).isTrue();
    }

    // -------------------------------------------------------------------------
    // writeResult
    // -------------------------------------------------------------------------

    @Test
    void writeResult_outputValido_escribeWorkbook(@TempDir Path tmp) throws IOException {
        Path out = tmp.resolve("resultado.xlsx");

        try (Workbook wb = new XSSFWorkbook()) {
            wb.createSheet("hoja1");
            om.writeResult(wb, out.toString());
        }

        assertThat(Files.exists(out)).isTrue();
        assertThat(Files.size(out)).isPositive();
    }

    @Test
    void writeResult_directorioPadreNoExiste_lanzaOutputExceptionGenerica(@TempDir Path tmp) throws IOException {
        // Apuntamos a un fichero dentro de un directorio que NO existe -> el
        // FileOutputStream lanza FileNotFoundException con un mensaje generico
        // ("No such file or directory" o "El sistema no puede encontrar la
        // ruta") que NO matchea looksLikeLocked en ninguna locale soportada.
        Path out = tmp.resolve("subdir-no-existe").resolve("salida.xlsx");

        try (Workbook wb = new XSSFWorkbook()) {
            wb.createSheet("hoja1");
            assertThatThrownBy(() -> om.writeResult(wb, out.toString()))
                    .isInstanceOf(OutputException.class)
                    .hasMessageContaining("No se pudo escribir el fichero de salida");
        }
    }

    @Test
    @EnabledOnOs({OS.LINUX, OS.MAC})
    void writeResult_outputEnDirectorioSinPermisosLinux_lanzaOutputExceptionLockedLike(@TempDir Path tmp)
            throws IOException {
        // root ignora permisos POSIX -> no ejecutar como root (ej. CI Docker).
        if ("root".equals(System.getProperty("user.name"))) return;

        // Quitar permiso de escritura al directorio padre -> FileNotFoundException
        // con mensaje "Permission denied" -> rama looksLikeLocked = true.
        Set<PosixFilePermission> readOnly = PosixFilePermissions.fromString("r-xr-xr-x");
        Files.setPosixFilePermissions(tmp, readOnly);

        Path out = tmp.resolve("salida.xlsx");

        try (Workbook wb = new XSSFWorkbook()) {
            wb.createSheet("hoja1");
            assertThatThrownBy(() -> om.writeResult(wb, out.toString()))
                    .isInstanceOf(OutputException.class)
                    .hasMessageContaining("Cierra 'salida.xlsx'")
                    .hasMessageContaining("parece abierto en Excel");
        } finally {
            // Restaurar permisos para que @TempDir pueda limpiar.
            Files.setPosixFilePermissions(tmp, PosixFilePermissions.fromString("rwxr-xr-x"));
        }
    }

    // NOTA: NO existe test "writeResult_outputBloqueadoEnWindows" porque la
    // rama "locked-like" en writeResult NO es ejercitable de forma
    // determinista. En Windows, FileChannel.lock() NO impide que un
    // FileOutputStream concurrente abra el descriptor (solo bloquea I/O), y
    // el fallo se manifiesta dentro de POI como OpenXML4JRuntimeException
    // (RuntimeException) durante result.write(fos), NO como IOException en
    // la apertura. El catch de writeResult solo captura
    // FileNotFoundException/IOException, asi que el RuntimeException se
    // propagaria sin envolver en OutputException. Esto es una limitacion
    // conocida documentada en CHANGELOG [2.5.1] como "rama defensiva no
    // testeable sin mockear POI". La rama "locked-like" del catch de
    // writeResult queda cubierta indirectamente por el test
    // writeResult_outputEnDirectorioSinPermisosLinux_lanzaOutputExceptionLockedLike
    // cuando se ejecuta en Linux/Mac.
}

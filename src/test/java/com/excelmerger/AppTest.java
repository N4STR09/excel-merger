package com.excelmerger;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests de {@link App}, la clase con la logica extraida de
 * {@code Main.main(args)} en v3.0.0.
 *
 * <p>Estos tests ejercitan el "happy path" y las ramas principales de
 * error de {@link App#run(String)} sin pasar por el menu interactivo
 * ni por {@link System#exit(int)}.</p>
 *
 * <p>Para el flujo completo de fusion existe ya una bateria amplia en
 * {@link ExcelMergerIntegrationTest}; aqui nos centramos en lo que es
 * propio de {@code App}: el wiring entre {@code ConfigLoader},
 * {@code ConfigValidator}, {@code ExcelMerger} y los exit codes.</p>
 */
class AppTest {

    // =================================================================
    //  Happy path
    // =================================================================

    @Test
    void runConConfigValidoDevuelveExitOk(@TempDir Path tmp) throws IOException {
        Path cfg = preparedConfig(tmp);
        int exitCode = App.run(cfg.toString());
        assertThat(exitCode).isEqualTo(App.EXIT_OK);

        // Verifica que el output se ha escrito (no estamos en dry-run).
        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        assertThat(output).exists();
    }

    @Test
    void runConDryRunNoEscribeOutput(@TempDir Path tmp) throws IOException {
        Path cfg = preparedConfig(tmp);
        // Anadimos output.dryRun=true al config ya renderizado.
        Files.writeString(cfg,
                Files.readString(cfg) + "\noutput.dryRun=true\n");

        int exitCode = App.run(cfg.toString());
        assertThat(exitCode).isEqualTo(App.EXIT_OK);

        // En dry-run no se escribe el output.
        Path output = tmp.resolve("output").resolve("resultado.xlsx");
        assertThat(output).doesNotExist();
    }

    // =================================================================
    //  Ramas de error
    // =================================================================

    @Test
    void runConConfigInexistenteDevuelveExitConfig() {
        // Path absoluto, no existe en disco ni en classpath.
        int exitCode = App.run("/no/existe/este-config.properties");
        assertThat(exitCode).isEqualTo(App.EXIT_CONFIG);
    }

    @Test
    void runConDirectorioDeEntradaInexistenteDevuelveExitInputInvalid(@TempDir Path tmp)
            throws IOException {
        Path cfg = preparedConfig(tmp);
        // Borramos el input para forzar el fallo de validacion de entrada.
        deleteRecursively(tmp.resolve("input"));

        int exitCode = App.run(cfg.toString());
        // El directorio existe a nivel de path (el ConfigValidator lo
        // valida ANTES) — pero al borrarlo despues de renderizar,
        // dependiendo de donde explote puede caer en CONFIG (si lo pilla
        // ConfigValidator) o INPUT_INVALID (si lo pilla ExcelMerger).
        // Aceptamos cualquiera de los dos errores tipados.
        assertThat(exitCode).isIn(App.EXIT_CONFIG, App.EXIT_INPUT_INVALID);
    }

    @Test
    void runConStrictValidationFalseAvanzaAunqueHayaWarnings(@TempDir Path tmp)
            throws IOException {
        Path cfg = preparedConfig(tmp);
        // Anadimos al config una clave invalida que ConfigValidator
        // marcaria como error en strict, mas el opt-out a no-strict.
        // output.mode con valor invalido es un error de validacion
        // tipico (lo controla IoConfigSection).
        Files.writeString(cfg,
                Files.readString(cfg)
                        + "\nconfig.strictValidation=false\n"
                        + "output.mode=valor_inventado_que_no_existe\n");

        int exitCode = App.run(cfg.toString());
        // En no-strict, el error de validacion se degrada a warning y
        // se intenta seguir. Output.mode invalido se tolera y el
        // fallback es el modo por defecto (cierre); el merge debe
        // completarse OK sobre los fixtures.
        assertThat(exitCode).isEqualTo(App.EXIT_OK);
    }

    @Test
    void runConStrictValidationTrueYConfigInvalidoDevuelveExitConfig(@TempDir Path tmp)
            throws IOException {
        Path cfg = preparedConfig(tmp);
        // strict=true (default) y output.mode invalido -> exit 2.
        Files.writeString(cfg,
                Files.readString(cfg)
                        + "\noutput.mode=valor_inventado_que_no_existe\n");

        int exitCode = App.run(cfg.toString());
        assertThat(exitCode).isEqualTo(App.EXIT_CONFIG);
    }

    // =================================================================
    //  Helpers
    // =================================================================

    /** Renderiza el test-config estandar y devuelve la ruta del .properties. */
    private static Path preparedConfig(Path tmp) throws IOException {
        Path inputDir = tmp.resolve("input");
        Path outputFile = tmp.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(outputFile.getParent());
        TestFixtures.copyFixturesTo(inputDir);
        return TestFixtures.renderTestConfig(
                tmp.resolve("test-config.properties"), inputDir, outputFile);
    }

    /** Borrado recursivo de un directorio (suficiente para tests). */
    private static void deleteRecursively(Path path) throws IOException {
        if (!Files.exists(path)) return;
        try (var stream = Files.walk(path)) {
            stream.sorted((a, b) -> b.getNameCount() - a.getNameCount())
                    .forEach(p -> {
                        try {
                            Files.deleteIfExists(p);
                        } catch (IOException ignored) {
                            // best-effort en tests
                        }
                    });
        }
    }
}

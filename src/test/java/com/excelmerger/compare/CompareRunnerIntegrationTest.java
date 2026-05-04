package com.excelmerger.compare;

import com.excelmerger.App;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Test de integracion del flujo completo de la Opcion 2 del menu (v3.1.0).
 *
 * <p>Construye un setup realista en {@code tempDir}:</p>
 * <ul>
 *   <li>{@code input/}: los 3 fixtures CSV (99001, 99002, 99003).</li>
 *   <li>{@code output/resultado_fusion.xlsx}: Resultado sintetico con
 *       las claves descritas en el blueprint del prompt para que el
 *       conteo total de discrepancias sea exactamente 7.</li>
 * </ul>
 *
 * <p>Tras ejecutar {@link CompareRunner#run()}, comprueba:</p>
 * <ul>
 *   <li>Exit code 0.</li>
 *   <li>Existe un fichero {@code output/discrepancias_*.xlsx}.</li>
 *   <li>Tiene 7 filas de datos.</li>
 *   <li>Distribucion correcta por tipo: 1 DIFERENCIA + 4 SOLO_CSV + 2 SOLO_RESULTADO.</li>
 * </ul>
 */
class CompareRunnerIntegrationTest {

    private static final String FIXTURE_99001 = "/fixtures/csv/fixture_99001.csv";
    private static final String FIXTURE_99002 = "/fixtures/csv/fixture_99002.csv";
    private static final String FIXTURE_99003 = "/fixtures/csv/fixture_99003.csv";

    private static void copyFixture(Path destDir, String classpathPath, String name) throws IOException {
        try (InputStream in = CompareRunnerIntegrationTest.class.getResourceAsStream(classpathPath)) {
            assertThat(in).as("Fixture %s no encontrado", classpathPath).isNotNull();
            Files.copy(in, destDir.resolve(name));
        }
    }

    private static Path buildResultado(Path outputDir) throws IOException {
        Files.createDirectories(outputDir);
        Path xlsx = outputDir.resolve("resultado_fusion.xlsx");
        new ResultadoFixtureBuilder()
                .add("100001", "99001", "RE", 5.0)   // match con fixture_99001 -> nada
                .add("100002", "99001", "AN", 8.0)   // csv=10.0 -> DIFERENCIA
                .add("100004", "99001", "PR", 0.0)   // match -> nada
                .add("100005", "99001", "RE", 0.5)   // match -> nada
                .add("100006", "99001", "OT", 7.0)   // SOLO_RESULTADO
                .add("100007", "99001", "RE", 2.0)   // SOLO_RESULTADO
                .add("300001", "99003", "RE", 20.0)  // match con fixture_99003
                .writeTo(outputDir, "resultado_fusion.xlsx");
        return xlsx;
    }

    @Test
    void flujoCompletoConTresCsvProduceSieteDiscrepancias(@TempDir Path tempDir) throws IOException {
        Path inputDir = tempDir.resolve("input");
        Path outputDir = tempDir.resolve("output");
        Files.createDirectories(inputDir);
        copyFixture(inputDir, FIXTURE_99001, "99001.CSV");
        copyFixture(inputDir, FIXTURE_99002, "99002.CSV");
        copyFixture(inputDir, FIXTURE_99003, "99003.CSV");
        Path resultadoPath = buildResultado(outputDir);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        PrintStream out = new PrintStream(baos, true, StandardCharsets.UTF_8);

        CompareRunner runner = new CompareRunner(inputDir, resultadoPath, outputDir, out);
        int exitCode = runner.run();

        assertThat(exitCode).isEqualTo(App.EXIT_OK);

        // Localiza el Excel generado: discrepancias_<timestamp>.xlsx
        List<Path> generated;
        try (Stream<Path> s = Files.list(outputDir)) {
            generated = s.filter(p -> p.getFileName().toString().startsWith("discrepancias_")
                    && p.getFileName().toString().endsWith(".xlsx"))
                    .toList();
        }
        assertThat(generated).hasSize(1);

        // Lee el Excel generado y cuenta tipos.
        Map<String, Integer> tipoCount = new HashMap<>();
        try (InputStream in = Files.newInputStream(generated.get(0));
             Workbook wb = new XSSFWorkbook(in)) {
            Sheet sheet = wb.getSheet(DiscrepancyExporter.SHEET_NAME);
            assertThat(sheet).isNotNull();
            // Fila 0 es cabecera; fila 1..N son datos.
            int last = sheet.getLastRowNum();
            assertThat(last).as("Fila de cabecera + 7 datos").isEqualTo(7);

            for (int r = 1; r <= last; r++) {
                Row row = sheet.getRow(r);
                String tipo = row.getCell(1).getStringCellValue();
                tipoCount.merge(tipo, 1, Integer::sum);
            }
        }

        // Esperado:
        //   fixture_99001: 1 DIFERENCIA + 1 SOLO_CSV + 2 SOLO_RESULTADO = 4
        //   fixture_99002: 3 SOLO_CSV (matricula 99002 no esta en Resultado)
        //   fixture_99003: 0
        // Total: 1 DIFERENCIA, 4 SOLO_CSV, 2 SOLO_RESULTADO.
        assertThat(tipoCount.getOrDefault("DIFERENCIA", 0)).isEqualTo(1);
        assertThat(tipoCount.getOrDefault("SOLO_CSV", 0)).isEqualTo(4);
        assertThat(tipoCount.getOrDefault("SOLO_RESULTADO", 0)).isEqualTo(2);
    }

    @Test
    void siNoHayCsvDevuelveOkConMensajeInformativo(@TempDir Path tempDir) throws IOException {
        Path inputDir = tempDir.resolve("input");
        Path outputDir = tempDir.resolve("output");
        Files.createDirectories(inputDir);
        Path resultadoPath = buildResultado(outputDir);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        PrintStream out = new PrintStream(baos, true, StandardCharsets.UTF_8);

        int exitCode = new CompareRunner(inputDir, resultadoPath, outputDir, out).run();

        assertThat(exitCode).isEqualTo(App.EXIT_OK);
        String output = baos.toString(StandardCharsets.UTF_8);
        assertThat(output).contains("No se encontraron CSV");
        // No deberia haber generado ningun discrepancias_*.xlsx.
        try (Stream<Path> s = Files.list(outputDir)) {
            List<Path> generated = s.filter(p -> p.getFileName().toString().startsWith("discrepancias_"))
                    .toList();
            assertThat(generated).isEmpty();
        }
    }

    @Test
    void siNoExisteResultadoDevuelveOkConMensaje(@TempDir Path tempDir) throws IOException {
        Path inputDir = tempDir.resolve("input");
        Path outputDir = tempDir.resolve("output");
        Files.createDirectories(inputDir);
        Files.createDirectories(outputDir);
        copyFixture(inputDir, FIXTURE_99003, "99003.CSV");
        // NO se crea resultado_fusion.xlsx
        Path resultadoPath = outputDir.resolve("resultado_fusion.xlsx");

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        PrintStream out = new PrintStream(baos, true, StandardCharsets.UTF_8);

        int exitCode = new CompareRunner(inputDir, resultadoPath, outputDir, out).run();

        assertThat(exitCode).isEqualTo(App.EXIT_OK);
        String output = baos.toString(StandardCharsets.UTF_8);
        assertThat(output).contains("Ejecuta primero la Opcion 1");
    }

    @Test
    void siInputDirNoExisteSeTrataComoVacio(@TempDir Path tempDir) throws IOException {
        Path inputDir = tempDir.resolve("noexiste");
        Path outputDir = tempDir.resolve("output");
        Path resultadoPath = buildResultado(outputDir);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        PrintStream out = new PrintStream(baos, true, StandardCharsets.UTF_8);

        int exitCode = new CompareRunner(inputDir, resultadoPath, outputDir, out).run();

        assertThat(exitCode).isEqualTo(App.EXIT_OK);
        assertThat(baos.toString(StandardCharsets.UTF_8)).contains("No se encontraron CSV");
    }

    @Test
    void unCsvCorruptoNoAbortaElProcesadoDeLosDemas(@TempDir Path tempDir) throws IOException {
        Path inputDir = tempDir.resolve("input");
        Path outputDir = tempDir.resolve("output");
        Files.createDirectories(inputDir);
        // Un CSV bien formado y un CSV vacio (sin cabeceras).
        copyFixture(inputDir, FIXTURE_99003, "99003.CSV");
        Files.write(inputDir.resolve("99999.CSV"), new byte[0]);
        Path resultadoPath = buildResultado(outputDir);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        PrintStream out = new PrintStream(baos, true, StandardCharsets.UTF_8);

        // Debe terminar OK pese al CSV corrupto: el log mostrara error
        // en el corrupto y procesara el bueno.
        int exitCode = new CompareRunner(inputDir, resultadoPath, outputDir, out).run();
        assertThat(exitCode).isEqualTo(App.EXIT_OK);

        String output = baos.toString(StandardCharsets.UTF_8);
        assertThat(output).contains("[ERROR]");

        // El Excel se ha generado igualmente.
        try (Stream<Path> s = Files.list(outputDir)) {
            List<Path> generated = s.filter(p -> p.getFileName().toString().startsWith("discrepancias_"))
                    .toList();
            assertThat(generated).hasSize(1);
        }
    }

    @Test
    void elNombreDelExcelDeSalidaIncluyeTimestamp(@TempDir Path tempDir) throws IOException {
        Path inputDir = tempDir.resolve("input");
        Path outputDir = tempDir.resolve("output");
        Files.createDirectories(inputDir);
        copyFixture(inputDir, FIXTURE_99003, "99003.CSV");
        Path resultadoPath = buildResultado(outputDir);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        PrintStream out = new PrintStream(baos, true, StandardCharsets.UTF_8);

        new CompareRunner(inputDir, resultadoPath, outputDir, out).run();

        try (Stream<Path> s = Files.list(outputDir)) {
            List<Path> matches = s.filter(p -> p.getFileName().toString()
                            .matches("discrepancias_\\d{4}-\\d{2}-\\d{2}_\\d{6}\\.xlsx"))
                    .toList();
            assertThat(matches).hasSize(1);
        }
    }
}

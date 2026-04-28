package com.excelmerger;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Reader;
import java.io.StringReader;
import java.io.UncheckedIOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Properties;

/**
 * Utilidad compartida por los tests de la Sesion C.
 *
 * <p>Funcionalidad:</p>
 * <ul>
 *   <li>Copiar los fixtures Excel ({@code extraccion.xlsx}, {@code cierre.xlsx})
 *       desde {@code src/test/resources/fixtures/} a un directorio temporal.</li>
 *   <li>Renderizar {@code test-config.properties} sustituyendo los placeholders
 *       {@code ${TEST_INPUT_DIR}} y {@code ${TEST_OUTPUT_FILE}} por rutas reales
 *       (las del {@code @TempDir} de cada test).</li>
 *   <li>Construir un {@link ConfigLoader} en memoria a partir de un {@link Properties}
 *       sin necesidad de escribir a disco (util para tests unitarios de
 *       {@link com.excelmerger.config.ConfigValidator} y similares).</li>
 * </ul>
 *
 * <p>No usa logging para no introducir ruido en la salida de los tests.</p>
 */
public final class TestFixtures {

    public static final String FIXTURE_EXTRACCION = "fixtures/extraccion.xlsx";
    public static final String FIXTURE_CIERRE     = "fixtures/cierre.xlsx";
    public static final String FIXTURE_DEUDA      = "fixtures/deuda.xlsx";
    public static final String TEST_CONFIG        = "test-config.properties";

    private TestFixtures() {
        // utility
    }

    /**
     * Copia {@code extraccion.xlsx} y {@code cierre.xlsx} desde el classpath
     * al directorio indicado. Devuelve la ruta del directorio para encadenar.
     */
    public static Path copyFixturesTo(Path inputDir) throws IOException {
        Files.createDirectories(inputDir);
        copyClasspathResource(FIXTURE_EXTRACCION, inputDir.resolve("extraccion.xlsx"));
        copyClasspathResource(FIXTURE_CIERRE,     inputDir.resolve("cierre.xlsx"));
        return inputDir;
    }

    /**
     * v2.2.0: variante que ademas copia {@code deuda.xlsx}. Util para los
     * tests que ejercitan la columna "PDCL + Deuda" con suma real de horas.
     */
    public static Path copyFixturesWithDeudaTo(Path inputDir) throws IOException {
        copyFixturesTo(inputDir);
        copyClasspathResource(FIXTURE_DEUDA, inputDir.resolve("deuda.xlsx"));
        return inputDir;
    }

    /**
     * Lee {@code test-config.properties}, sustituye los placeholders por
     * rutas reales y escribe el resultado en {@code targetConfig}. Devuelve
     * la ruta escrita.
     */
    public static Path renderTestConfig(Path targetConfig, Path inputDir, Path outputFile) throws IOException {
        String raw = readClasspathResourceAsString(TEST_CONFIG);
        String rendered = raw
                .replace("${TEST_INPUT_DIR}",  inputDir.toAbsolutePath().toString().replace('\\', '/'))
                .replace("${TEST_OUTPUT_FILE}", outputFile.toAbsolutePath().toString().replace('\\', '/'));
        Files.writeString(targetConfig, rendered, StandardCharsets.UTF_8);
        return targetConfig;
    }

    /**
     * Atajo: copia fixtures al directorio de entrada, renderiza el config
     * apuntando a un output dentro del mismo {@code baseDir} y devuelve un
     * {@link ConfigLoader} ya cargado listo para pasar a {@link ExcelMerger}.
     */
    public static ConfigLoader buildRealisticConfig(Path baseDir) throws IOException {
        Path inputDir = baseDir.resolve("input");
        Path outputFile = baseDir.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(outputFile.getParent());
        copyFixturesTo(inputDir);
        Path cfg = renderTestConfig(baseDir.resolve("test-config.properties"), inputDir, outputFile);
        return new ConfigLoader(cfg.toString());
    }

    /**
     * v2.3.0: variante de {@link #buildRealisticConfig(Path)} que ademas
     * sobreescribe la clave {@code output.mode} con el valor indicado, sin
     * tocar el resto del rendering. Util para tests del modo de salida.
     *
     * <p>Implementacion: renderiza el test-config como hace el atajo normal,
     * lee el resultado, sustituye o anade {@code output.mode=<modo>}, y
     * vuelve a escribir el fichero antes de cargarlo. Asi cubrimos los dos
     * casos: clave existente (la sustituye) y clave ausente (la anade al
     * final).</p>
     */
    public static ConfigLoader buildRealisticConfigWithOutputMode(Path baseDir, String mode) throws IOException {
        Path inputDir = baseDir.resolve("input");
        Path outputFile = baseDir.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(outputFile.getParent());
        copyFixturesTo(inputDir);
        Path cfg = renderTestConfig(baseDir.resolve("test-config.properties"), inputDir, outputFile);
        overrideOutputMode(cfg, mode);
        return new ConfigLoader(cfg.toString());
    }

    /**
     * v2.3.0: variante con fixtures de Deuda incluidos para tests del modo
     * COMPLETO (que usa Deuda) y RESPONSABLES (que la ignora explicitamente).
     */
    public static ConfigLoader buildRealisticConfigWithDeudaAndOutputMode(Path baseDir, String mode) throws IOException {
        Path inputDir = baseDir.resolve("input");
        Path outputFile = baseDir.resolve("output").resolve("resultado.xlsx");
        Files.createDirectories(outputFile.getParent());
        copyFixturesWithDeudaTo(inputDir);
        Path cfg = renderTestConfig(baseDir.resolve("test-config.properties"), inputDir, outputFile);
        overrideOutputMode(cfg, mode);
        return new ConfigLoader(cfg.toString());
    }

    private static void overrideOutputMode(Path cfg, String mode) throws IOException {
        String content = Files.readString(cfg, StandardCharsets.UTF_8);
        java.util.regex.Pattern p =
                java.util.regex.Pattern.compile("^output\\.mode=.*$", java.util.regex.Pattern.MULTILINE);
        String replacement = "output.mode=" + mode;
        if (p.matcher(content).find()) {
            content = p.matcher(content).replaceAll(java.util.regex.Matcher.quoteReplacement(replacement));
        } else {
            content = content + System.lineSeparator() + replacement + System.lineSeparator();
        }
        Files.writeString(cfg, content, StandardCharsets.UTF_8);
    }

    /**
     * Construye un {@link ConfigLoader} a partir de un {@link Properties}.
     * Util para tests unitarios: el Properties se serializa a un fichero
     * temporal y se carga via el constructor publico de ConfigLoader.
     *
     * <p>Usa {@link Properties#store(java.io.Writer, String)} para que los
     * caracteres especiales (incluidos los backslashes de rutas Windows) se
     * escapen correctamente. Si lo escribieramos "a mano", un
     * {@code C:\Users\ERLANT~1} acabaria leyendose como {@code C:UsersERLANT~1}
     * porque Properties interpreta {@code \U}, {@code \n}, {@code \A}, etc.
     * como secuencias de escape.</p>
     */
    public static ConfigLoader configFromProperties(Properties props) {
        try {
            Path tmp = Files.createTempFile("excelmerger-test-cfg-", ".properties");
            tmp.toFile().deleteOnExit();
            try (java.io.Writer w = Files.newBufferedWriter(tmp, StandardCharsets.UTF_8)) {
                props.store(w, null);
            }
            return new ConfigLoader(tmp.toString());
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    /**
     * Version de {@link #configFromProperties(Properties)} que acepta pares
     * variadicos para tests cortos.
     *
     * <pre>configFromPairs("input.directory", "/tmp", "output.file", "/tmp/out.xlsx")</pre>
     */
    public static ConfigLoader configFromPairs(String... keyValuePairs) {
        if (keyValuePairs.length % 2 != 0) {
            throw new IllegalArgumentException("Se esperaban pares clave/valor.");
        }
        Properties p = new Properties();
        for (int i = 0; i < keyValuePairs.length; i += 2) {
            p.setProperty(keyValuePairs[i], keyValuePairs[i + 1]);
        }
        return configFromProperties(p);
    }

    // ==================================================================
    //  Helpers internos
    // ==================================================================
    private static void copyClasspathResource(String resource, Path target) throws IOException {
        try (InputStream in = TestFixtures.class.getClassLoader().getResourceAsStream(resource)) {
            if (in == null) {
                throw new IOException("Fixture no encontrada en classpath: " + resource);
            }
            Files.createDirectories(target.getParent());
            try (OutputStream out = new FileOutputStream(target.toFile())) {
                in.transferTo(out);
            }
        }
    }

    private static String readClasspathResourceAsString(String resource) throws IOException {
        try (InputStream in = TestFixtures.class.getClassLoader().getResourceAsStream(resource)) {
            if (in == null) {
                throw new IOException("Recurso no encontrado en classpath: " + resource);
            }
            return new String(in.readAllBytes(), StandardCharsets.UTF_8);
        }
    }

    /**
     * Carga el contenido de un Properties a partir de una cadena (para tests
     * que quieran verificar que el rendering de test-config ha producido lo esperado).
     */
    public static Properties parsePropertiesString(String content) throws IOException {
        Properties p = new Properties();
        try (Reader r = new StringReader(content)) {
            p.load(r);
        }
        return p;
    }
}

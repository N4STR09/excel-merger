package com.excelmerger.compare;

import com.excelmerger.exception.InputValidationException;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Tests unitarios de {@link CsvParser}. v3.1.0.
 *
 * <p>Cubre:</p>
 * <ul>
 *   <li>Parseo de los 3 fixtures CSV ({@code fixture_99001}, {@code fixture_99002},
 *       {@code fixture_99003}) — sinteticos pero con formato binario identico
 *       al del ERP real (encoding cp1252, CRLF, NUL en {@code DR-Marca},
 *       padding decimal).</li>
 *   <li>Parseo del fixture real {@code real_90054.csv} (copia byte-a-byte
 *       de un export real del ERP). Verifica que los acentos y la "ñ" de
 *       las cabeceras se decodifican correctamente.</li>
 *   <li>Helpers internos: {@link CsvParser#cleanField},
 *       {@link CsvParser#stripJPrefix}, {@link CsvParser#parseSpanishDecimal},
 *       {@link CsvParser#originFromFileName}, {@link CsvParser#isCsvFileName}.</li>
 *   <li>Errores: fichero ausente, cabecera obligatoria faltante.</li>
 * </ul>
 */
class CsvParserTest {

    private static final String FIXTURE_99001 = "/fixtures/csv/fixture_99001.csv";
    private static final String FIXTURE_99002 = "/fixtures/csv/fixture_99002.csv";
    private static final String FIXTURE_99003 = "/fixtures/csv/fixture_99003.csv";
    private static final String FIXTURE_REAL_90054 = "/fixtures/csv/real_90054.csv";

    /** Copia un fixture del classpath a un path temporal y lo devuelve. */
    private static Path copyFixtureToTemp(Path tempDir, String classpathPath, String name)
            throws IOException {
        Path out = tempDir.resolve(name);
        try (InputStream in = CsvParserTest.class.getResourceAsStream(classpathPath)) {
            assertThat(in).as("Fixture %s no encontrado", classpathPath).isNotNull();
            Files.copy(in, out);
        }
        return out;
    }

    // =================================================================
    //  Fixture sintetico 99001: 5 filas, todos los matices
    // =================================================================

    @Test
    void parsea99001ConLas5ImputacionesNormalizadas(@TempDir Path tempDir) throws IOException {
        Path csv = copyFixtureToTemp(tempDir, FIXTURE_99001, "99001.CSV");
        List<CsvImputacion> imps = new CsvParser().parse(csv);

        assertThat(imps).hasSize(5);
        // Fila 1: J100001, RE, 5,0 -> peticion=100001, funcion=RE, hours=5.0
        assertThat(imps.get(0).getPeticion()).isEqualTo("100001");
        assertThat(imps.get(0).getMatricula()).isEqualTo("99001");
        assertThat(imps.get(0).getFuncion()).isEqualTo("RE");
        assertThat(imps.get(0).getRealizadoHoras()).isEqualTo(5.0);

        // Fila 2: AN, 10.0
        assertThat(imps.get(1).getPeticion()).isEqualTo("100002");
        assertThat(imps.get(1).getFuncion()).isEqualTo("AN");
        assertThat(imps.get(1).getRealizadoHoras()).isEqualTo(10.0);

        // Fila 3: OT, 3.5
        assertThat(imps.get(2).getRealizadoHoras()).isEqualTo(3.5);

        // Fila 4: PR, ",0" (sin parte entera) -> 0.0
        assertThat(imps.get(3).getFuncion()).isEqualTo("PR");
        assertThat(imps.get(3).getRealizadoHoras()).isEqualTo(0.0);

        // Fila 5: ",5" -> 0.5
        assertThat(imps.get(4).getRealizadoHoras()).isEqualTo(0.5);
    }

    @Test
    void laJDePeticionSeQuitaSiempre(@TempDir Path tempDir) throws IOException {
        Path csv = copyFixtureToTemp(tempDir, FIXTURE_99001, "99001.CSV");
        List<CsvImputacion> imps = new CsvParser().parse(csv);
        for (CsvImputacion imp : imps) {
            assertThat(imp.getPeticion())
                    .as("Peticion no debe empezar por J")
                    .doesNotStartWith("J");
        }
    }

    @Test
    void elNULDelDRMarcaNoApareceEnLosCamposExtraidos(@TempDir Path tempDir) throws IOException {
        // Aunque DR-Marca no es uno de los campos que extraemos, el test
        // garantiza que el parser no falla / no propaga el caracter NUL
        // en los campos que SI extraemos.
        Path csv = copyFixtureToTemp(tempDir, FIXTURE_99001, "99001.CSV");
        List<CsvImputacion> imps = new CsvParser().parse(csv);
        for (CsvImputacion imp : imps) {
            assertThat(imp.getPeticion()).doesNotContain("\u0000");
            assertThat(imp.getMatricula()).doesNotContain("\u0000");
            assertThat(imp.getFuncion()).doesNotContain("\u0000");
        }
    }

    // =================================================================
    //  Fixture sintetico 99002: 3 filas con tildes y ñ
    // =================================================================

    @Test
    void parsea99002Con3Imputaciones(@TempDir Path tempDir) throws IOException {
        Path csv = copyFixtureToTemp(tempDir, FIXTURE_99002, "99002.CSV");
        List<CsvImputacion> imps = new CsvParser().parse(csv);
        assertThat(imps).hasSize(3);
        assertThat(imps.get(0).getPeticion()).isEqualTo("200001");
        assertThat(imps.get(0).getMatricula()).isEqualTo("99002");
        assertThat(imps.get(0).getFuncion()).isEqualTo("RE");
        assertThat(imps.get(0).getRealizadoHoras()).isEqualTo(7.0);
        assertThat(imps.get(1).getRealizadoHoras()).isEqualTo(14.5);
        assertThat(imps.get(2).getRealizadoHoras()).isEqualTo(0.5);
    }

    // =================================================================
    //  Fixture real 90054.CSV (copia exacta de un export real)
    // =================================================================

    @Test
    void parsea90054RealCorrectamente(@TempDir Path tempDir) throws IOException {
        Path csv = copyFixtureToTemp(tempDir, FIXTURE_REAL_90054, "90054.CSV");
        List<CsvImputacion> imps = new CsvParser().parse(csv);
        assertThat(imps).hasSize(1);
        CsvImputacion imp = imps.get(0);
        assertThat(imp.getPeticion()).isEqualTo("136261");
        assertThat(imp.getMatricula()).isEqualTo("90054");
        assertThat(imp.getFuncion()).isEqualTo("OT");
        // Realizado Horas: ",0" -> 0.0
        assertThat(imp.getRealizadoHoras()).isEqualTo(0.0);
    }

    @Test
    void parsearEncodingLatin1RespetaTildesYEnyes(@TempDir Path tempDir) throws IOException {
        // Verifica indirectamente que el encoding no es UTF-8: si fuese
        // UTF-8, "Petición" no se localizaria como cabecera porque el
        // byte 0xF3 (ó) seria una secuencia invalida y el campo se
        // leeria como "Petici?n" o lanzaria excepcion.
        // Si llega hasta aqui sin excepcion, el encoding cp1252 ha
        // funcionado (la cabecera "Petición" se ha encontrado).
        Path csv = copyFixtureToTemp(tempDir, FIXTURE_99001, "99001.CSV");
        List<CsvImputacion> imps = new CsvParser().parse(csv);
        assertThat(imps).isNotEmpty();
    }

    // =================================================================
    //  Helpers internos
    // =================================================================

    @Test
    void cleanFieldQuitaNUL() {
        assertThat(CsvParser.cleanField("hello\u0000world")).isEqualTo("helloworld");
        assertThat(CsvParser.cleanField("\u0000")).isEmpty();
        assertThat(CsvParser.cleanField("  hola \u0000 mundo  ")).isEqualTo("hola  mundo");
    }

    @Test
    void cleanFieldHaceTrim() {
        assertThat(CsvParser.cleanField("   hello   ")).isEqualTo("hello");
        assertThat(CsvParser.cleanField("")).isEmpty();
        assertThat(CsvParser.cleanField(null)).isEmpty();
    }

    @Test
    void stripJPrefixQuitaJMayusculaYMinuscula() {
        assertThat(CsvParser.stripJPrefix("J137791")).isEqualTo("137791");
        assertThat(CsvParser.stripJPrefix("j137791")).isEqualTo("137791");
        assertThat(CsvParser.stripJPrefix("137791")).isEqualTo("137791");
        assertThat(CsvParser.stripJPrefix("")).isEmpty();
        assertThat(CsvParser.stripJPrefix(null)).isEmpty();
    }

    @Test
    void parseSpanishDecimalManejaCasoSinParteEntera() {
        assertThat(CsvParser.parseSpanishDecimal("5,0")).isEqualTo(5.0);
        assertThat(CsvParser.parseSpanishDecimal(",0")).isEqualTo(0.0);
        assertThat(CsvParser.parseSpanishDecimal(",5")).isEqualTo(0.5);
        assertThat(CsvParser.parseSpanishDecimal("-5,5")).isEqualTo(-5.5);
        assertThat(CsvParser.parseSpanishDecimal("-,5")).isEqualTo(-0.5);
        assertThat(CsvParser.parseSpanishDecimal("100")).isEqualTo(100.0);
    }

    @Test
    void parseSpanishDecimalLanzaSiVacio() {
        assertThatThrownBy(() -> CsvParser.parseSpanishDecimal(""))
                .isInstanceOf(NumberFormatException.class);
    }

    @Test
    void parseSpanishDecimalLanzaSiNoEsNumero() {
        assertThatThrownBy(() -> CsvParser.parseSpanishDecimal("abc"))
                .isInstanceOf(NumberFormatException.class);
    }

    @Test
    void originFromFileNameQuitaExtension() {
        assertThat(CsvParser.originFromFileName("90014.CSV")).isEqualTo("90014");
        assertThat(CsvParser.originFromFileName("90014.csv")).isEqualTo("90014");
        assertThat(CsvParser.originFromFileName("90014")).isEqualTo("90014");
        assertThat(CsvParser.originFromFileName("foo.bar.csv")).isEqualTo("foo.bar");
        assertThat(CsvParser.originFromFileName("")).isEmpty();
        assertThat(CsvParser.originFromFileName(null)).isEmpty();
    }

    @Test
    void isCsvFileNameAceptaUpperLowerYRechazaOtros() {
        assertThat(CsvParser.isCsvFileName("90014.CSV")).isTrue();
        assertThat(CsvParser.isCsvFileName("90014.csv")).isTrue();
        assertThat(CsvParser.isCsvFileName("90014.Csv")).isTrue();
        assertThat(CsvParser.isCsvFileName("90014.xlsx")).isFalse();
        assertThat(CsvParser.isCsvFileName("90014")).isFalse();
        assertThat(CsvParser.isCsvFileName("")).isFalse();
        assertThat(CsvParser.isCsvFileName(null)).isFalse();
    }

    // =================================================================
    //  Errores
    // =================================================================

    @Test
    void parserLanzaSiFicheroNoExiste(@TempDir Path tempDir) {
        Path missing = tempDir.resolve("nope.csv");
        assertThatThrownBy(() -> new CsvParser().parse(missing))
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("no encontrado");
    }

    @Test
    void parserLanzaSiFaltaCabeceraObligatoria() {
        // CSV minimo sin la cabecera "Fu"
        String csv = "\"Año\";\"Mes\";\"Matricula \";\"Petición\";\"Realizado Horas\"\r\n"
                + "\"2026\";\" 4\";\"99001\";\"J1\";\"5,0\"\r\n";
        assertThatThrownBy(() -> new CsvParser().parse(new StringReader(csv), "minimal"))
                .isInstanceOf(InputValidationException.class)
                .hasMessageContaining("\"Fu\"");
    }

    @Test
    void parserAceptaLecturaDesdeReader() {
        // Cabecera minima con todas las columnas obligatorias y 1 fila.
        String csv = "\"Matricula \";\"Petición\";\"Fu\";\"Realizado Horas\"\r\n"
                + "\"  90014  \";\"J137791\";\"RE\";\"     12,5\"\r\n";
        List<CsvImputacion> imps = new CsvParser().parse(new StringReader(csv), "test");
        assertThat(imps).hasSize(1);
        assertThat(imps.get(0).getMatricula()).isEqualTo("90014");
        assertThat(imps.get(0).getPeticion()).isEqualTo("137791");
        assertThat(imps.get(0).getFuncion()).isEqualTo("RE");
        assertThat(imps.get(0).getRealizadoHoras()).isEqualTo(12.5);
    }

    @Test
    void parserIgnoraFilasConRealizadoNoParseable() {
        String csv = "\"Matricula \";\"Petición\";\"Fu\";\"Realizado Horas\"\r\n"
                + "\"99001\";\"J1\";\"RE\";\"basura\"\r\n"
                + "\"99001\";\"J2\";\"AN\";\"5,0\"\r\n";
        List<CsvImputacion> imps = new CsvParser().parse(new StringReader(csv), "test");
        // La fila con "basura" se ignora con warning.
        assertThat(imps).hasSize(1);
        assertThat(imps.get(0).getPeticion()).isEqualTo("2");
    }

    @Test
    void parserIgnoraFilasCompletamenteVacias() {
        String csv = "\"Matricula \";\"Petición\";\"Fu\";\"Realizado Horas\"\r\n"
                + "\"\";\"\";\"\";\"\"\r\n"
                + "\"99001\";\"J1\";\"RE\";\"3,0\"\r\n";
        List<CsvImputacion> imps = new CsvParser().parse(new StringReader(csv), "test");
        assertThat(imps).hasSize(1);
    }
}

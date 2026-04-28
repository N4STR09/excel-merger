package com.excelmerger.config;

import com.excelmerger.ConfigLoader;
import com.excelmerger.TestFixtures;

import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests directos de {@link ValidationHelpers}. Bonus de la Sesión F (v2.5.0):
 * los helpers ya estaban indirectamente cubiertos vía {@code ConfigValidatorTest},
 * pero al haberlos extraído a una clase propia tiene sentido anclar su
 * contrato con tests específicos. En particular, el contrato de mutación de
 * {@link ValidationHelpers#parsePositiveInt} (añade error si el valor no es
 * numérico, devuelve el default si no hay valor) merece tests directos.
 */
class ValidationHelpersTest {

    private static ConfigLoader configWith(String key, String value) {
        Properties p = new Properties();
        if (value != null) {
            p.setProperty(key, value);
        }
        return TestFixtures.configFromProperties(p);
    }

    // ==================================================================
    //  parseCsv
    // ==================================================================

    @Test
    void parseCsvDevuelveListaVaciaParaNull() {
        assertThat(ValidationHelpers.parseCsv(null)).isEmpty();
    }

    @Test
    void parseCsvDevuelveListaVaciaParaCadenaVacia() {
        assertThat(ValidationHelpers.parseCsv("")).isEmpty();
    }

    @Test
    void parseCsvHaceTrimYDescartaEntradasVacias() {
        assertThat(ValidationHelpers.parseCsv(" a , ,b ,  ,c"))
                .containsExactly("a", "b", "c");
    }

    @Test
    void parseCsvUnSoloElemento() {
        assertThat(ValidationHelpers.parseCsv("solo")).containsExactly("solo");
    }

    // ==================================================================
    //  isBlank
    // ==================================================================

    @Test
    void isBlankReconoceNullCadenaVaciaYEspacios() {
        assertThat(ValidationHelpers.isBlank(null)).isTrue();
        assertThat(ValidationHelpers.isBlank("")).isTrue();
        assertThat(ValidationHelpers.isBlank("   ")).isTrue();
        assertThat(ValidationHelpers.isBlank("\t\n")).isTrue();
    }

    @Test
    void isBlankFalseCuandoHayContenido() {
        assertThat(ValidationHelpers.isBlank("x")).isFalse();
        assertThat(ValidationHelpers.isBlank(" x ")).isFalse();
    }

    // ==================================================================
    //  parsePositiveInt
    // ==================================================================

    @Test
    void parsePositiveIntDevuelveDefaultSiClaveAusente() {
        ConfigLoader cfg = configWith("otra.clave", "10");
        List<String> errors = new ArrayList<>();

        Integer result = ValidationHelpers.parsePositiveInt(cfg, "key.que.no.existe", 7, errors);

        assertThat(result).isEqualTo(7);
        assertThat(errors).isEmpty();
    }

    @Test
    void parsePositiveIntDevuelveNullSiClaveAusenteYDefaultNull() {
        ConfigLoader cfg = configWith("otra.clave", "10");
        List<String> errors = new ArrayList<>();

        Integer result = ValidationHelpers.parsePositiveInt(cfg, "key.que.no.existe", null, errors);

        assertThat(result).isNull();
        assertThat(errors).isEmpty();
    }

    @Test
    void parsePositiveIntParsea() {
        ConfigLoader cfg = configWith("k", "42");
        List<String> errors = new ArrayList<>();

        assertThat(ValidationHelpers.parsePositiveInt(cfg, "k", 1, errors)).isEqualTo(42);
        assertThat(errors).isEmpty();
    }

    @Test
    void parsePositiveIntAcumulaErrorSiNoEsNumerico() {
        ConfigLoader cfg = configWith("k", "doce");
        List<String> errors = new ArrayList<>();

        Integer result = ValidationHelpers.parsePositiveInt(cfg, "k", 99, errors);

        // Devuelve el default y ANUNCIA el error en la lista compartida.
        assertThat(result).isEqualTo(99);
        assertThat(errors).hasSize(1);
        assertThat(errors.get(0)).contains("k").contains("doce").contains("no numerico");
    }

    @Test
    void parsePositiveIntAcumulaErroresMultiples() {
        // Dos llamadas, dos valores no numéricos -> dos entradas en errors.
        Properties p = new Properties();
        p.setProperty("a", "uno");
        p.setProperty("b", "dos");
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        List<String> errors = new ArrayList<>();

        ValidationHelpers.parsePositiveInt(cfg, "a", 0, errors);
        ValidationHelpers.parsePositiveInt(cfg, "b", 0, errors);

        assertThat(errors).hasSize(2);
        assertThat(errors).anyMatch(e -> e.contains("a"));
        assertThat(errors).anyMatch(e -> e.contains("b"));
    }
}

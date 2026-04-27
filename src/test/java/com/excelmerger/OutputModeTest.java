package com.excelmerger;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Tests unitarios del enum {@link OutputMode} y su parser estricto.
 */
class OutputModeTest {

    @Test
    void parseStrict_aceptaLosTresValoresValidosEnMinusculas() {
        assertThat(OutputMode.parseStrict("cierre")).isEqualTo(OutputMode.CIERRE);
        assertThat(OutputMode.parseStrict("responsables")).isEqualTo(OutputMode.RESPONSABLES);
        assertThat(OutputMode.parseStrict("completo")).isEqualTo(OutputMode.COMPLETO);
    }

    @Test
    void parseStrict_rechazaMayusculas() {
        assertThatThrownBy(() -> OutputMode.parseStrict("CIERRE"))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("invalido")
                .hasMessageContaining("CIERRE");
        assertThatThrownBy(() -> OutputMode.parseStrict("Cierre"))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> OutputMode.parseStrict("Responsables"))
                .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void parseStrict_rechazaValoresInventados() {
        assertThatThrownBy(() -> OutputMode.parseStrict("foo"))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> OutputMode.parseStrict(""))
                .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void parseStrict_rechazaNull() {
        assertThatThrownBy(() -> OutputMode.parseStrict(null))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("null");
    }

    @Test
    void parseStrict_mensajeListaLosTresValoresValidos() {
        try {
            OutputMode.parseStrict("xxx");
        } catch (IllegalArgumentException e) {
            assertThat(e.getMessage())
                    .contains("cierre")
                    .contains("responsables")
                    .contains("completo");
        }
    }
}

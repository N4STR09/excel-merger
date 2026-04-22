package com.excelmerger;

import org.junit.jupiter.api.Test;

import java.time.Duration;
import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests de {@link RunReport}: acumulacion de hojas/warnings, orden de
 * insercion, contabilidad y formato del resumen final.
 */
class RunReportTest {

    @Test
    void nuevoReportArrancaVacio() {
        RunReport r = new RunReport();

        assertThat(r.sheets()).isEmpty();
        assertThat(r.warnings()).isEmpty();
        assertThat(r.warningCount()).isZero();
    }

    @Test
    void addSheetRegistraNombreYNumeroDeFilas() {
        RunReport r = new RunReport();

        r.addSheet("Extraccion", 15);
        r.addSheet("MES", 14);

        assertThat(r.sheets())
                .containsEntry("Extraccion", 15)
                .containsEntry("MES", 14)
                .hasSize(2);
    }

    @Test
    void sheetsPreservaOrdenDeInsercion() {
        RunReport r = new RunReport();

        r.addSheet("primera", 1);
        r.addSheet("segunda", 2);
        r.addSheet("tercera", 3);

        assertThat(r.sheets().keySet()).containsExactly("primera", "segunda", "tercera");
    }

    @Test
    void addWarningAcumulaEnOrden() {
        RunReport r = new RunReport();

        r.addWarning("CABECERA", "primer aviso");
        r.addWarning("HOJA",     "segundo aviso");
        r.addWarning("CABECERA", "tercer aviso");

        assertThat(r.warningCount()).isEqualTo(3);
        assertThat(r.warnings()).hasSize(3);
        assertThat(r.warnings().get(0).category).isEqualTo("CABECERA");
        assertThat(r.warnings().get(0).message).isEqualTo("primer aviso");
        assertThat(r.warnings().get(1).category).isEqualTo("HOJA");
        assertThat(r.warnings().get(2).message).isEqualTo("tercer aviso");
    }

    @Test
    void warningsDevuelveListaInmodificable() {
        RunReport r = new RunReport();
        r.addWarning("X", "y");

        assertThat(r.warnings()).hasSize(1);
        // La lista devuelta no debe permitir modificaciones externas
        org.assertj.core.api.Assertions
                .assertThatThrownBy(() -> r.warnings().add(new RunReport.Warning("Z", "z")))
                .isInstanceOf(UnsupportedOperationException.class);
    }

    @Test
    void warningToStringLlevaCategoriaYMensaje() {
        RunReport.Warning w = new RunReport.Warning("PERFIL", "algo falla");

        assertThat(w.toString()).isEqualTo("[PERFIL] algo falla");
    }

    @Test
    void registerLookupKeysYhasLookup() {
        RunReport r = new RunReport();
        Set<String> keys = new LinkedHashSet<>(Arrays.asList("DF", "HE", "EW"));

        r.registerLookupKeys("Equipos", keys);

        assertThat(r.hasLookup("Equipos")).isTrue();
        assertThat(r.hasLookup("Otra")).isFalse();
    }

    @Test
    void getLookupKeysDevuelveCopiaInmodificable() {
        RunReport r = new RunReport();
        Set<String> original = new LinkedHashSet<>(Arrays.asList("A", "B", "C"));
        r.registerLookupKeys("Equipos", original);

        Set<String> retrieved = r.getLookupKeys("Equipos");

        assertThat(retrieved).containsExactly("A", "B", "C");
        org.assertj.core.api.Assertions.assertThatThrownBy(() -> retrieved.add("Z"))
                .isInstanceOf(UnsupportedOperationException.class);
    }

    @Test
    void getLookupKeysDevuelveNullSiNoExisteElLookup() {
        RunReport r = new RunReport();

        assertThat(r.getLookupKeys("inexistente")).isNull();
    }

    @Test
    void registerLookupKeysNoGuardaReferenciaDirecta() {
        // Si registramos una Set mutable y luego la mutamos por fuera,
        // el report debe seguir viendo las claves originales (copia defensiva).
        RunReport r = new RunReport();
        Set<String> keys = new LinkedHashSet<>(Arrays.asList("DF", "HE"));
        r.registerLookupKeys("Equipos", keys);

        keys.add("ZZ"); // mutacion externa

        assertThat(r.getLookupKeys("Equipos")).containsExactly("DF", "HE");
    }

    @Test
    void formatSummaryIncluyeTiempoTotal() {
        RunReport r = new RunReport();

        String summary = r.formatSummary(Duration.ofMillis(1234));

        assertThat(summary).contains("RESUMEN DE EJECUCION");
        assertThat(summary).contains("Tiempo total: 1234 ms");
    }

    @Test
    void formatSummaryIncluyeHojasYFilas() {
        RunReport r = new RunReport();
        r.addSheet("Extraccion", 15);
        r.addSheet("MES", 14);

        String summary = r.formatSummary(Duration.ofMillis(10));

        assertThat(summary).contains("Hojas generadas (2)");
        assertThat(summary).contains("Extraccion");
        assertThat(summary).contains("(15 filas)");
        assertThat(summary).contains("MES");
        assertThat(summary).contains("(14 filas)");
    }

    @Test
    void formatSummaryIncluyeWarningsCategorizados() {
        RunReport r = new RunReport();
        r.addWarning("PERFIL", "ningun perfil coincide");
        r.addWarning("LOOKUP", "3 valores sin mapeo");

        String summary = r.formatSummary(Duration.ZERO);

        assertThat(summary).contains("Warnings (2)");
        assertThat(summary).contains("[PERFIL] ningun perfil coincide");
        assertThat(summary).contains("[LOOKUP] 3 valores sin mapeo");
    }

    @Test
    void formatSummaryMuestraNingunoCuandoNoHayWarnings() {
        RunReport r = new RunReport();
        r.addSheet("solo", 1);

        String summary = r.formatSummary(Duration.ZERO);

        assertThat(summary).contains("Warnings (0)");
        assertThat(summary).contains("(ninguno)");
    }

    @Test
    void formatSummaryMuestraNingunaCuandoNoHayHojas() {
        RunReport r = new RunReport();

        String summary = r.formatSummary(Duration.ZERO);

        assertThat(summary).contains("Hojas generadas (0)");
        assertThat(summary).contains("(ninguna)");
    }
}

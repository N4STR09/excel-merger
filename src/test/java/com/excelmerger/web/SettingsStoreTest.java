package com.excelmerger.web;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Tests de {@link SettingsStore}: default que restaura la salida completa
 * de v3.1.0, carga resiliente desde fichero, persistencia en
 * web-settings.properties, validacion de modos, acoplamiento
 * Resumen/por-responsable y serializacion JSON para la API.
 */
class SettingsStoreTest {

    @Test
    void sinFicheroUsaLosValoresPorDefecto(@TempDir Path tmp) {
        SettingsStore store = SettingsStore.load(tmp.resolve("ausente.properties"));

        // Default v4.1.0: la salida completa que tenia el usuario antes
        // de v4.0.0 (modo completo + Resumen + matriz por responsable).
        assertThat(store.mode()).isEqualTo("completo");
        assertThat(store.summaryEnabled()).isTrue();
        assertThat(store.byResponsibleEnabled()).isTrue();
    }

    @Test
    void leeLosAjustesGuardados(@TempDir Path tmp) throws IOException {
        Path file = tmp.resolve("web-settings.properties");
        Files.writeString(file,
                "output.mode=cierre\n"
                        + "summary.enabled=false\n"
                        + "summary.byResponsible.enabled=false\n",
                StandardCharsets.UTF_8);

        SettingsStore store = SettingsStore.load(file);

        assertThat(store.mode()).isEqualTo("cierre");
        assertThat(store.summaryEnabled()).isFalse();
        assertThat(store.byResponsibleEnabled()).isFalse();
        assertThat(store.toOverrides().getProperty("output.mode")).isEqualTo("cierre");
        assertThat(store.toOverrides().getProperty("summary.enabled")).isEqualTo("false");
        assertThat(store.toOverrides()
                .getProperty("summary.byResponsible.enabled")).isEqualTo("false");
    }

    @Test
    void ficheroIlegibleCaeADefaults(@TempDir Path tmp) throws IOException {
        Path file = tmp.resolve("web-settings.properties");
        Files.writeString(file, "esto no es un properties {{{", StandardCharsets.UTF_8);

        SettingsStore store = SettingsStore.load(file);

        assertThat(store.mode()).isEqualTo("completo");
        assertThat(store.summaryEnabled()).isTrue();
        assertThat(store.byResponsibleEnabled()).isTrue();
    }

    @Test
    void modoInvalidoEnFicheroCaeADefault(@TempDir Path tmp) throws IOException {
        Path file = tmp.resolve("web-settings.properties");
        Files.writeString(file, "output.mode=EXPERIMENTAL\n", StandardCharsets.UTF_8);

        SettingsStore store = SettingsStore.load(file);

        assertThat(store.mode()).isEqualTo("completo");
    }

    @Test
    void updatePersisteYSeRecarga(@TempDir Path tmp) throws IOException {
        Path file = tmp.resolve("web-settings.properties");
        SettingsStore store = SettingsStore.load(file);
        store.update("responsables", false, false);

        SettingsStore reloaded = SettingsStore.load(file);
        assertThat(reloaded.mode()).isEqualTo("responsables");
        assertThat(reloaded.summaryEnabled()).isFalse();
        assertThat(reloaded.byResponsibleEnabled()).isFalse();
    }

    @Test
    void toJsonSerializaLosTresAjustes(@TempDir Path tmp) {
        SettingsStore store = SettingsStore.load(tmp.resolve("x.properties"));

        assertThat(store.toJson())
                .isEqualTo("{\"mode\":\"completo\",\"summaryEnabled\":true,"
                        + "\"byResponsibleEnabled\":true}");
    }

    @Test
    void updateConModoInvalidoLanzaIllegalArgumentException(@TempDir Path tmp) {
        SettingsStore store = SettingsStore.load(tmp.resolve("x.properties"));

        assertThatThrownBy(() -> store.update("EXPERIMENTAL", true, true))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("Modo de salida invalido");
    }

    @Test
    void porResponsableObligaAActivarElResumen(@TempDir Path tmp) throws IOException {
        Path file = tmp.resolve("web-settings.properties");
        SettingsStore store = SettingsStore.load(file);
        store.update("completo", false, true);

        // Regla del validador del motor: matriz por responsable requiere
        // hoja Resumen. La interfaz la resuelve en origen.
        assertThat(store.summaryEnabled()).isTrue();
        assertThat(store.byResponsibleEnabled()).isTrue();
        assertThat(SettingsStore.load(file).summaryEnabled()).isTrue();
    }

    @Test
    void updateConNullsMantieneLosValoresActuales(@TempDir Path tmp) {
        SettingsStore store = SettingsStore.load(tmp.resolve("x.properties"));
        store.update(null, null, null);

        assertThat(store.mode()).isEqualTo("completo");
        assertThat(store.summaryEnabled()).isTrue();
        assertThat(store.byResponsibleEnabled()).isTrue();
    }

    @Test
    void updateConNullsConservaCambiosParciales(@TempDir Path tmp) throws IOException {
        Path file = tmp.resolve("web-settings.properties");
        SettingsStore store = SettingsStore.load(file);
        store.update("cierre", null, null);

        assertThat(store.mode()).isEqualTo("cierre");
        assertThat(store.summaryEnabled()).isTrue();
        // La matriz por responsable sigue activa: con el Resumen activo es
        // coherente y el modo cierre la incluye.
        assertThat(store.byResponsibleEnabled()).isTrue();
        assertThat(SettingsStore.load(file).mode()).isEqualTo("cierre");
    }

    @Test
    void modosYValidez() {
        assertThat(SettingsStore.modes())
                .containsExactly("cierre", "responsables", "completo");
        assertThat(SettingsStore.isValidMode("completo")).isTrue();
        assertThat(SettingsStore.isValidMode("responsables")).isTrue();
        assertThat(SettingsStore.isValidMode("cierre")).isTrue();
        assertThat(SettingsStore.isValidMode("Completo")).isFalse();
        assertThat(SettingsStore.isValidMode("")).isFalse();
        assertThat(SettingsStore.isValidMode(null)).isFalse();
    }
}

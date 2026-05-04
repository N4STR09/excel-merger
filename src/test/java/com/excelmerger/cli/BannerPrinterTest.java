package com.excelmerger.cli;

import com.excelmerger.Main;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests del banner ASCII de bienvenida (v3.0.0).
 *
 * <p>{@link BannerPrinter} es package-private. Estos tests viven en el
 * mismo paquete para tener acceso.</p>
 */
class BannerPrinterTest {

    private static String capture() {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        PrintStream ps = new PrintStream(baos, true, StandardCharsets.UTF_8);
        new BannerPrinter(ps).print();
        return baos.toString(StandardCharsets.UTF_8);
    }

    @Test
    void elBannerContieneLaVersionDelMain() {
        String out = capture();
        assertThat(out).contains("v" + Main.APP_VERSION);
    }

    @Test
    void elBannerContieneLaLineaDescriptiva() {
        String out = capture();
        assertThat(out).contains("Fusion de exports ERP + Jira para cierre mensual");
    }

    @Test
    void elBannerContieneSecuenciasAnsiCyan() {
        String out = capture();
        assertThat(out).contains("\u001B[36m"); // ANSI cyan ON
        assertThat(out).contains("\u001B[0m");  // ANSI reset
    }

    @Test
    void elBannerContieneAlMenosCincoLineasDeAsciiArt() {
        String out = capture();
        // El ASCII art tiene 6 lineas. Cuento cuantas tienen el caracter
        // pipe que es signature de las letras Figlet standard.
        long pipeLines = out.lines().filter(l -> l.contains("|")).count();
        assertThat(pipeLines).isGreaterThanOrEqualTo(5);
    }
}

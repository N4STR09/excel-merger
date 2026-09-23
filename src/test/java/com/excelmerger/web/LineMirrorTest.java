package com.excelmerger.web;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests del espejo de consola: reparto buffer/delegado, UTF-8, corte de
 * retorno de carro, techo por linea y pass-through transparente.
 */
class LineMirrorTest {

    @Test
    void reparteLineasEntreBufferYDelegado() throws Exception {
        LogBuffer buffer = new LogBuffer();
        ByteArrayOutputStream sink = new ByteArrayOutputStream();
        PrintStream delegate = new PrintStream(sink, true, StandardCharsets.UTF_8);
        LineMirror mirror = new LineMirror(buffer, delegate, "OUT");

        mirror.write("hola\nsegunda línea ñ\n".getBytes(StandardCharsets.UTF_8));
        mirror.flush();

        List<LogBuffer.Event> events = buffer.since(0);
        assertThat(events).hasSize(2);
        assertThat(events.get(0).message).isEqualTo("hola");
        assertThat(events.get(0).level).isEqualTo("OUT");
        assertThat(events.get(1).message).isEqualTo("segunda línea ñ");
        assertThat(sink.toString(StandardCharsets.UTF_8))
                .isEqualTo("hola\nsegunda línea ñ\n");
    }

    @Test
    void eliminaElRetornoDeCarroDeWindows() throws Exception {
        LogBuffer buffer = new LogBuffer();
        PrintStream delegate = new PrintStream(new ByteArrayOutputStream());
        LineMirror mirror = new LineMirror(buffer, delegate, "OUT");

        mirror.write("INFO  aviso con retorno\r\n".getBytes(StandardCharsets.UTF_8));

        List<LogBuffer.Event> events = buffer.since(0);
        assertThat(events).hasSize(1);
        assertThat(events.get(0).message).isEqualTo("INFO  aviso con retorno");
    }

    @Test
    void emiteEnTrozosSiLaLineaSuperaElTecho() throws Exception {
        LogBuffer buffer = new LogBuffer();
        ByteArrayOutputStream sink = new ByteArrayOutputStream();
        PrintStream delegate = new PrintStream(sink, true, StandardCharsets.UTF_8);
        LineMirror mirror = new LineMirror(buffer, delegate, "OUT");
        byte[] chunk = new byte[LineMirror.MAX_LINE_BYTES + 100];
        java.util.Arrays.fill(chunk, (byte) 'a');

        mirror.write(chunk);

        List<LogBuffer.Event> events = buffer.since(0);
        assertThat(events).hasSize(1);
        assertThat(events.get(0).message).hasSize(LineMirror.MAX_LINE_BYTES);
        assertThat(sink.size()).isEqualTo(LineMirror.MAX_LINE_BYTES + 100);
    }

    @Test
    void ignoraLineasEnBlancoPeroDejaElPassThrough() throws Exception {
        LogBuffer buffer = new LogBuffer();
        ByteArrayOutputStream sink = new ByteArrayOutputStream();
        PrintStream delegate = new PrintStream(sink, true, StandardCharsets.UTF_8);
        LineMirror mirror = new LineMirror(buffer, delegate, "OUT");

        mirror.write(" \n\n".getBytes(StandardCharsets.UTF_8));

        assertThat(buffer.since(0)).isEmpty();
        assertThat(sink.toString(StandardCharsets.UTF_8)).isEqualTo(" \n\n");
    }

    @Test
    void elCierreVueltaPendienteSinSalto() throws Exception {
        LogBuffer buffer = new LogBuffer();
        PrintStream delegate = new PrintStream(new ByteArrayOutputStream());
        LineMirror mirror = new LineMirror(buffer, delegate, "OUT");

        mirror.write("sin newline final".getBytes(StandardCharsets.UTF_8));
        assertThat(buffer.since(0)).isEmpty();
        mirror.close();

        List<LogBuffer.Event> events = buffer.since(0);
        assertThat(events).hasSize(1);
        assertThat(events.get(0).message).isEqualTo("sin newline final");
    }
}

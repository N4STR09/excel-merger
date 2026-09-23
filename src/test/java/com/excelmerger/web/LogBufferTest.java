package com.excelmerger.web;

import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/** Tests del buffer circular de registro que alimenta la interfaz. */
class LogBufferTest {

    @Test
    void asignaSecuenciasYSinceFiltralaPorCursor() {
        LogBuffer logBuffer = new LogBuffer();
        logBuffer.add("INFO", "a");
        logBuffer.add("WARN", "b");

        List<LogBuffer.Event> all = logBuffer.since(0);
        List<LogBuffer.Event> tail = logBuffer.since(1);

        assertThat(all).hasSize(2);
        assertThat(all.get(0).seq).isEqualTo(1);
        assertThat(all.get(0).message).isEqualTo("a");
        assertThat(tail).hasSize(1);
        assertThat(tail.get(0).message).isEqualTo("b");
        assertThat(logBuffer.since(99)).isEmpty();
    }

    @Test
    void descartaElMasAntiguoAlSuperarElMaximo() {
        LogBuffer logBuffer = new LogBuffer();
        int total = LogBuffer.MAX_EVENTS + 5;
        for (int i = 0; i < total; i++) {
            logBuffer.add("INFO", "m" + i);
        }

        List<LogBuffer.Event> events = logBuffer.since(0);

        assertThat(events).hasSize(LogBuffer.MAX_EVENTS);
        assertThat(events.get(0).seq).isEqualTo(6);
        assertThat(events.get(0).message).isEqualTo("m5");
        assertThat(events.get(events.size() - 1).message).isEqualTo("m" + (total - 1));
    }
}

package com.excelmerger.web;

import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.List;

/**
 * Buffer circular en memoria de los eventos que la aplicacion vuelca a
 * consola (logback via su appender de consola, y System.out directo),
 * expuesto por la API web para pintar el panel de log de la interfaz
 * (v4.0.0).
 *
 * <p>Mantiene como maximo {@value #MAX_EVENTS} eventos: los mas antiguos
 * se descartan. Es la memoria que ve el navegador; no guarda nada en
 * disco (el historial completo sigue estando en {@code logs/excel-merger.log}
 * via logback).</p>
 *
 * <p>Todos los metodos de lectura y escritura son {@code synchronized}
 * porque el buffer se escribe desde el hilo de logback y desde el hilo
 * espejo de {@code System.out}, y se lee desde los hilos del servidor
 * HTTP.</p>
 */
public final class LogBuffer {

    /** Numero maximo de eventos retenidos. */
    static final int MAX_EVENTS = 2000;

    private final ArrayDeque<Event> events = new ArrayDeque<>();

    /** Secuencia global (1-based) del ultimo evento anadido. Guardado por {@code this}. */
    private long lastSeq;

    /**
     * Anade un evento al final del buffer, descartando el mas antiguo si
     * se supera {@value #MAX_EVENTS}.
     *
     * @param level   nivel del evento ({@code INFO}, {@code OUT}, ...).
     * @param message mensaje ya formateado (sin patron SLF4J).
     */
    public synchronized void add(String level, String message) {
        lastSeq++;
        if (events.size() >= MAX_EVENTS) {
            events.removeFirst();
        }
        events.addLast(new Event(lastSeq, level, message));
    }

    /**
     * Devuelve (en orden) los eventos con secuencia estrictamente mayor que
     * {@code afterSeq}. La lista devuelta es una copia: el caller puede
     * recorrerla sin bloquear al resto del sistema.
     *
     * @param afterSeq cursor del cliente (0 = desde el principio).
     * @return copia de los eventos posteriores al cursor.
     */
    public synchronized List<Event> since(long afterSeq) {
        List<Event> result = new ArrayList<>();
        for (Event event : events) {
            if (event.seq > afterSeq) {
                result.add(event);
            }
        }
        return result;
    }

    /**
     * Evento inmutable de log: secuencia, nivel y mensaje.
     *
     * <p>Los campos son package-private y finales a proposito: los lee
     * {@link WebServer} para serializarlos y no hay mutadores que
     * exponer.</p>
     */
    public static final class Event {

        /** Secuencia global del evento (1-based). */
        final long seq;

        /** Nivel textual del evento. */
        final String level;

        /** Mensaje ya formateado. */
        final String message;

        Event(long seq, String level, String message) {
            this.seq = seq;
            this.level = level;
            this.message = message;
        }
    }
}

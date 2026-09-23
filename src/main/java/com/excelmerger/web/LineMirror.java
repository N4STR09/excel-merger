package com.excelmerger.web;

import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;

/**
 * {@link OutputStream} espejo que acumula bytes hasta encontrar un salto
 * de linea y reparte cada linea completa entre el buffer (para el panel
 * de la interfaz web) y el stream original (terminal/archivo, que no
 * pierde nada) — v4.0.0.
 *
 * <p>La linea se decodifica como UTF-8 al cerrarse, que es el charset
 * con el que {@link PrintStream} codifica, de modo que los acentos
 * llegan enteros al panel; el retorno de carro final de Windows se
 * descarta. Si una linea no trae salto y supera
 * {@value #MAX_LINE_BYTES} bytes se emite igualmente (en trozos), de
 * forma que el acumulador no puede crecer sin limite.</p>
 *
 * <p>{@link #close()} nunca cierra el stream delegado: es
 * {@code System.out}.</p>
 */
final class LineMirror extends OutputStream {

    private static final int NEWLINE = '\n';
    private static final String CR = "\r";

    /** Techo por linea: al alcanzarse se emite aunque no haya salto. */
    static final int MAX_LINE_BYTES = 8192;

    private final LogBuffer buffer;
    private final PrintStream delegate;
    private final String level;
    private final ByteArrayOutputStream line = new ByteArrayOutputStream();

    /**
     * @param buffer   destino de las lineas completas.
     * @param delegate stream original al que hacer pass-through.
     * @param level    etiqueta de canal del evento ({@code OUT}/{@code ERR}).
     */
    LineMirror(LogBuffer buffer, PrintStream delegate, String level) {
        this.buffer = buffer;
        this.delegate = delegate;
        this.level = level;
    }

    @Override
    public synchronized void write(int b) {
        delegate.write(b);
        if (b == NEWLINE) {
            flushLine();
        } else {
            line.write(b);
            flushIfFull();
        }
    }

    @Override
    public synchronized void write(byte[] bytes, int off, int len) {
        delegate.write(bytes, off, len);
        for (int i = off; i < off + len; i++) {
            if (bytes[i] == NEWLINE) {
                flushLine();
            } else {
                line.write(bytes[i]);
                flushIfFull();
            }
        }
    }

    @Override
    public synchronized void flush() {
        delegate.flush();
    }

    @Override
    public synchronized void close() {
        flushLine();
        delegate.flush();
        // No cerramos delegate: es System.out.
    }

    /** Empuja la linea acumulada al buffer (si no esta vacia) y la limpia. */
    private void flushLine() {
        if (line.size() > 0) {
            String text = line.toString(StandardCharsets.UTF_8);
            line.reset();
            if (text.endsWith(CR)) {
                text = text.substring(0, text.length() - 1);
            }
            if (!text.isBlank()) {
                buffer.add(level, text);
            }
            delegate.flush();
        }
    }

    /** Emite la linea si el acumulador ha alcanzado su techo. */
    private void flushIfFull() {
        if (line.size() >= MAX_LINE_BYTES) {
            flushLine();
        }
    }
}

package com.excelmerger.web;

import org.slf4j.LoggerFactory;

import java.io.PrintStream;
import java.nio.charset.StandardCharsets;

/**
 * Captura en memoria lo que la aplicacion imprime en consola para que la
 * interfaz web pueda mostrarlo en su panel de registro (v4.0.0).
 *
 * <p><b>Un solo canal:</b> se reemplazan {@code System.out/err} por
 * espejos ({@link LineMirror}) que reparten cada linea completa entre el
 * buffer y el stream original. Todo lo que la terminal ve llega una sola
 * vez al buffer:</p>
 * <ul>
 *   <li>logback escribe su consola contra {@code System.out}, de modo que
 *       sus eventos (con su prefijo {@code %-5level}) entran por aqui;</li>
 *   <li>los textos que
 *       {@link com.excelmerger.compare.CompareRunner} imprime en su
 *       {@code PrintStream} entran por aqui;</li>
 *   <li>la terminal no pierde nada: el espejo es transparente
 *       (pass-through) ademas de capturar.</li>
 * </ul>
 *
 * <p>Se descarto colgar ademas un appender propio de logback
 * (BufferAppender): la consola de logback aterriza tambien en el espejo,
 * y con los dos canales cada evento habria aparecido DOS veces en el
 * panel. El nivel real de logback se recupera en la UI leyendo el
 * prefijo del propio mensaje cuando existe (el nivel del evento queda en
 * {@code OUT}/{@code ERR}, que identifica el canal).</p>
 *
 * <p><b>Invariantes:</b></p>
 * <ol>
 *   <li>Nada debe emitir un evento de log ANTES de {@link #install()}.
 *       Si el {@code ConsoleAppender} de logback arrancara antes,
 *       cachearia el {@code System.out} original y el panel no veria el
 *       log. {@code WebLauncher.main} llama a {@code install()} como
 *       primera instruccion y ninguno de sus campos loguea; crear un
 *       logger (como el campo {@code log} de esta clase) solo inicializa
 *       el provider, no emite eventos.</li>
 *   <li>{@code logback.xml} debe conservar su appender de consola: es la
 *       via de alimentacion del panel. El fichero
 *       {@code logs/excel-merger.log} no cambia.</li>
 * </ol>
 *
 * <p><b>Nota de exclusion SpotBugs:</b> {@code MS_EXPOSE_REP} sobre
 * {@link #install()} esta excluido en {@code spotbugs-exclude.xml} con
 * justificacion: el buffer devuelto es el recurso compartido del
 * patron.</p>
 */
public final class LogCapture {

    /**
     * Logger estatico: su creacion inicializa el provider SLF4J pero NO
     * emite eventos, condicion necesaria de la invariante 1 de la
     * documentacion de la clase.
     */
    private static final org.slf4j.Logger log = LoggerFactory.getLogger(LogCapture.class);

    /**
     * Buffer unico de la instancia de JVM. Final y sin inicializacion
     * perezosa (SpotBugs LI_LAZY_INIT_STATIC): crearlo al cargar la
     * clase es trivial y evita el patron lazy estatico.
     */
    private static final LogBuffer BUFFER = new LogBuffer();

    /** Si {@code System.out/err} ya fueron reemplazados. Guardado por {@code LogCapture.class}. */
    private static boolean streamsRedirected;

    private LogCapture() {
        // Clase de utilidad
    }

    /**
     * Prepara la captura (espejo de {@code System.out/err}) y devuelve el
     * buffer comun.
     *
     * <p>Idempotente: las sucesivas llamadas devuelven siempre el mismo
     * buffer y no duplican espejos.</p>
     *
     * @return el buffer de eventos compartido por toda la JVM.
     */
    public static synchronized LogBuffer install() {
        redirectStreams(BUFFER);
        return BUFFER;
    }

    /**
     * Reemplaza {@code System.out/err} por espejos hacia el buffer.
     *
     * <p>{@code CloseResource} se suprime a proposito: los
     * {@code PrintStream} que se crean aqui no son recursos que haya
     * que cerrar, sino el nuevo {@code System.out}/{@code System.err}
     * del proceso; cerrarlos destruiria la salida estandar. Su ciclo de
     * vida es el de la JVM, como el de la consola de logback.</p>
     */
    @SuppressWarnings("PMD.CloseResource")
    private static void redirectStreams(LogBuffer buffer) {
        if (streamsRedirected) {
            return;
        }
        PrintStream originalOut = System.out;
        PrintStream originalErr = System.err;
        System.setOut(new PrintStream(new LineMirror(buffer, originalOut, "OUT"),
                true, StandardCharsets.UTF_8));
        System.setErr(new PrintStream(new LineMirror(buffer, originalErr, "ERR"),
                true, StandardCharsets.UTF_8));
        streamsRedirected = true;
        log.debug("LogCapture: System.out y System.err redirigidos al buffer");
    }
}

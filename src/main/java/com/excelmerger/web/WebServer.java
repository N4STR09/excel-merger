package com.excelmerger.web;

import com.excelmerger.App;
import com.excelmerger.Main;
import com.excelmerger.compare.CompareRunner;

import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpServer;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.InetAddress;
import java.net.InetSocketAddress;
import java.nio.charset.StandardCharsets;
import java.security.SecureRandom;
import java.util.Base64;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Properties;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadFactory;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Function;
import java.util.function.IntSupplier;

/**
 * Servidor HTTP local de Excel Merger (v4.0.0): sirve la interfaz
 * web y expone la <b>misma</b> logica que el menu de terminal, sin
 * cambios en el motor.
 *
 * <h2>Mapeo 1:1 con el menu actual</h2>
 * <ul>
 *   <li>{@code POST /api/merge} → {@link App#run(String)} con config por
 *       defecto (Opcion 1).</li>
 *   <li>{@code POST /api/compare} → {@code new CompareRunner().run()}
 *       (Opcion 2, que en terminal vuelve al menu).</li>
 *   <li>{@code POST /api/shutdown} → exit code 0 (Opcion 3, "Salir sin
 *       hacer nada"), via {@code shutdownAction} inyectado.</li>
 * </ul>
 *
 * <h2>Contrato de exit codes (decision v4.0.0: "split")</h2>
 * Cada ejecucion devuelve su propio codigo en el cuerpo JSON
 * ({@code exitCode}, valores 0-4 identicos a los de {@link App}); el
 * proceso como tal solo devuelve 0 (salida ordenada) o 1 (fallo de
 * arranque). Los modos headless {@code --merge}/{@code --compare}, que
 * preservan los 0-4 a nivel de proceso, los decide
 * {@link com.excelmerger.Main} desde v4.0.0.
 *
 * <h2>Seguridad local</h2>
 * Solo escucha en loopback (nunca en todas las interfaces), exige el
 * token de sesion en toda ruta {@code /api/} (persistido en
 * {@link WebSession} para que la direccion sea siempre la misma),
 * valida el header {@code Host} contra la direccion enlazada y rechaza
 * {@code Origin} ajeno. Asi ninguna pagina web del navegador del usuario
 * puede disparar ejecuciones ni leer el registro.
 *
 * <h2>Concurrencia</h2>
 * Hay una unica plaza de ejecucion: mientras una fusion/comprobacion
 * corre, la segunda peticion recibe 409 en lugar de pelearse por los
 * bloqueos de Excel sobre {@code output/}.
 */
public final class WebServer {

    private static final Logger log = LoggerFactory.getLogger(WebServer.class);

    /** RNG compartido para los tokens: crear un {@code SecureRandom} es caro. */
    private static final SecureRandom RANDOM = new SecureRandom();

    /** Header con el token de acceso. */
    static final String TOKEN_HEADER = "X-Em-Token";

    /** Nombre del parametro de query alternativo al header. */
    static final String TOKEN_PARAM = "token";

    /**
     * Puerto preferido de la interfaz: fijo para que la direccion sea
     * siempre la misma (marcable). Si esta ocupado se prueban los
     * siguientes {@link #PORT_SPAN} puertos y, en ultimo extremo, uno
     * efimero: el arranque nunca falla por culpa del puerto.
     */
    static final int PREFERRED_PORT = 7420;

    /** Puertos consecutivos a probar a partir del preferido. */
    static final int PORT_SPAN = 10;

    private static final String API_PREFIX = "/api/";
    private static final String HEADER_HOST = "Host";
    private static final String HEADER_ORIGIN = "Origin";
    private static final String LOCALHOST = "localhost";
    private static final String JSON_UTF8 = "application/json; charset=utf-8";
    private static final String TEXT_UTF8 = "text/plain; charset=utf-8";
    private static final String HTML_UTF8 = "text/html; charset=utf-8";
    private static final String CSS_UTF8 = "text/css; charset=utf-8";
    private static final String JS_UTF8 = "text/javascript; charset=utf-8";
    private static final String CACHE_NO_STORE = "no-store";
    private static final String RESOURCE_INDEX = "/web/index.html";
    private static final String RESOURCE_CSS = "/web/app.css";
    private static final String RESOURCE_JS = "/web/app.js";
    private static final String METHOD_GET = "GET";
    private static final String METHOD_POST = "POST";
    private static final String ERROR_JSON_PREFIX = "{\"error\":";
    private static final String DENIED_MESSAGE = "no autorizado";
    private static final String BUSY_MESSAGE = "ya hay una ejecucion en curso";
    private static final String SETTINGS_UNAVAILABLE = "ajustes no disponibles";
    private static final int TOKEN_BYTES = 16;
    private static final int POOL_SIZE = 4;
    private static final int SHUTDOWN_DELAY_MS = 400;
    private static final int MAX_BODY_BYTES = 2048;
    private static final int STATUS_OK = 200;
    private static final int STATUS_BAD_REQUEST = 400;
    private static final int STATUS_FORBIDDEN = 403;
    private static final int STATUS_NOT_FOUND = 404;
    private static final int STATUS_METHOD_NOT_ALLOWED = 405;
    private static final int STATUS_CONFLICT = 409;

    /** Nombres de campo del JSON de {@code /api/settings}. */
    private static final String JSON_MODE = "mode";
    private static final String JSON_SUMMARY = "summaryEnabled";
    private static final String JSON_BY_RESPONSIBLE = "byResponsibleEnabled";

    private final HttpServer server;
    private final ExecutorService pool;
    private final String token;
    private final Function<String, Integer> mergeRunner;
    private final IntSupplier compareRunner;
    private final Runnable shutdownAction;
    private final LogBuffer logBuffer;
    private final SettingsStore settings;
    private final AtomicBoolean busy = new AtomicBoolean(false);
    private final String boundAddress;
    private final String boundHostName;
    private final int port;
    private final String inputDir;
    private final String outputDir;

    /**
     * Constructor de pruebas: bindea en loopback con puerto efimero y
     * arranca el servidor con runners y accion de salida inyectados (mismo
     * patron que {@code InteractiveMenu} y {@code CompareRunner}).
     *
     * @param mergeRunner     accion de la Opcion 1; recibe la ruta de
     *                        config ({@code null} = por defecto) y devuelve
     *                        el exit code.
     * @param compareRunner   accion de la Opcion 2; devuelve el exit code.
     * @param shutdownAction  accion de "Salir" (en produccion, exit 0).
     * @param logBuffer       buffer que alimenta el panel de registro.
     * @throws IOException si no se puede enlazar el puerto.
     */
    WebServer(Function<String, Integer> mergeRunner, IntSupplier compareRunner,
              Runnable shutdownAction, LogBuffer logBuffer) throws IOException {
        this(mergeRunner, compareRunner, shutdownAction, logBuffer,
                List.of(0), null, null, null);
    }

    /**
     * Variante sin ajustes de salida (usada por los tests): delega en el
     * constructor general con {@code settings = null}, desactivando los
     * endpoints de ajustes.
     */
    WebServer(Function<String, Integer> mergeRunner, IntSupplier compareRunner,
              Runnable shutdownAction, LogBuffer logBuffer, List<Integer> candidatePorts,
              String sessionToken, String inputDir, String outputDir) throws IOException {
        this(mergeRunner, compareRunner, shutdownAction, logBuffer, candidatePorts,
                sessionToken, inputDir, outputDir, null);
    }

    /**
     * Constructor general: prueba los puertos indicados en orden (0 =
     * efimero) y acepta token de sesion, carpetas de trabajo y ajustes de
     * salida opcionales.
     *
     * @param mergeRunner     accion de la Opcion 1; recibe la ruta de
     *                        config ({@code null} = por defecto) y devuelve
     *                        el exit code.
     * @param compareRunner   accion de la Opcion 2; devuelve el exit code.
     * @param shutdownAction  accion de "Salir" (en produccion, exit 0).
     * @param logBuffer       buffer que alimenta el panel de registro.
     * @param candidatePorts  puertos a probar en orden; 0 = efimero.
     * @param sessionToken    token persistido ({@code null} = generar uno).
     * @param inputDir        carpeta de entrada en absoluto (null = oculta).
     * @param outputDir       carpeta de salida en absoluto (null = oculta).
     * @param settings        ajustes de salida configurables (null =
     *                        endpoints de ajustes desactivados).
     * @throws IOException si ningun puerto candidato es enlazable.
     */
    WebServer(Function<String, Integer> mergeRunner, IntSupplier compareRunner,
              Runnable shutdownAction, LogBuffer logBuffer, List<Integer> candidatePorts,
              String sessionToken, String inputDir, String outputDir, SettingsStore settings)
            throws IOException {
        this.mergeRunner = Objects.requireNonNull(mergeRunner);
        this.compareRunner = Objects.requireNonNull(compareRunner);
        this.shutdownAction = Objects.requireNonNull(shutdownAction);
        this.logBuffer = Objects.requireNonNull(logBuffer);
        this.settings = settings;
        this.token = sessionToken == null ? newToken() : sessionToken;
        this.server = bind(candidatePorts);
        InetAddress bound = server.getAddress().getAddress();
        this.boundAddress = bound.getHostAddress();
        this.boundHostName = bound.getHostName();
        this.port = server.getAddress().getPort();
        this.inputDir = inputDir;
        this.outputDir = outputDir;
        this.pool = Executors.newFixedThreadPool(POOL_SIZE, new WebThreadFactory());
        server.setExecutor(pool);
        server.createContext("/", this::route);
        server.start();
    }

    /**
     * Enlaza el primer puerto de la lista que este libre. Si el puerto
     * preferido lo ocupa otro proceso, se prueba el siguiente y, al
     * final, uno efimero.
     *
     * @param candidatePorts puertos a probar en orden (0 = efimero).
     * @return el servidor HTTP ya enlazado, sin arrancar.
     * @throws IOException si ninguno de los puertos se puede enlazar.
     */
    private static HttpServer bind(List<Integer> candidatePorts) throws IOException {
        IOException lastError = null;
        for (Integer candidate : candidatePorts) {
            try {
                return HttpServer.create(
                        new InetSocketAddress(InetAddress.getLoopbackAddress(), candidate), 0);
            } catch (IOException e) {
                lastError = e;
            }
        }
        throw lastError == null ? new IOException("Sin puertos candidatos") : lastError;
    }

    /**
     * Arranque de produccion: ata los endpoints a la logica real de la
     * aplicacion ({@link App#run(String)}, {@link CompareRunner} y salida
     * con exit code 0 tras confirmar en la interfaz). Prueba primero el
     * puerto de la sesion, despues {@value #PREFERRED_PORT} y los
     * {@value #PORT_SPAN} siguientes, y por ultimo uno efimero.
     *
     * @param logBuffer    buffer de registro que alimenta la interfaz.
     * @param preferred    ultimo puerto utilizado (0 = ninguno).
     * @param sessionToken token persistido (null = generar uno nuevo).
     * @param inputDir     carpeta de entrada en absoluto (null = oculta).
     * @param outputDir    carpeta de salida en absoluto (null = oculta).
     * @param settings     ajustes de salida que se aplican a cada fusion
     *                     ({@code null} = fusion sin overrides).
     * @return el servidor ya arrancado.
     * @throws IOException si no se puede enlazar ningun puerto.
     */
    public static WebServer start(LogBuffer logBuffer, int preferred, String sessionToken,
                                  String inputDir, String outputDir, SettingsStore settings)
            throws IOException {
        LinkedHashSet<Integer> candidates = new LinkedHashSet<>();
        if (preferred > 0) {
            candidates.add(preferred);
        }
        for (int candidate = PREFERRED_PORT; candidate <= PREFERRED_PORT + PORT_SPAN; candidate++) {
            candidates.add(candidate);
        }
        candidates.add(0);
        Function<String, Integer> merge = configPath -> {
            Properties overrides = settings == null ? null : settings.toOverrides();
            if (overrides != null) {
                log.info("Fusion con ajustes de la interfaz: modo {}, resumen {}, "
                        + "por responsable {}.", settings.mode(), settings.summaryEnabled(),
                        settings.byResponsibleEnabled());
            }
            return App.run(configPath, overrides);
        };
        WebServer webServer = new WebServer(merge, () -> new CompareRunner().run(),
                WebLauncher::exitOk, logBuffer, List.copyOf(candidates), sessionToken,
                inputDir, outputDir, settings);
        log.info("Servidor web local listo: {}", webServer.url());
        return webServer;
    }

    /** URL completa de la interfaz, ya con el token de acceso. */
    public String url() {
        return baseUrl() + "/?" + TOKEN_PARAM + "=" + token;
    }

    /** URL base ({@code http://<loopback>:<puerto>}) sin token. */
    public String baseUrl() {
        return "http://" + formatHost(boundAddress) + ":" + port;
    }

    /** Token de acceso: el de la sesion persistida o uno nuevo si no lo habia. */
    public String token() {
        return token;
    }

    /** Puerto sobre el que se ha enlazado (el preferido si estaba libre). */
    public int port() {
        return port;
    }

    /** Detiene el servidor y su pool de hilos. Idempotente. */
    public void stop() {
        server.stop(0);
        pool.shutdownNow();
    }

    /**
     * Decide si el header {@code Host} (con o sin puerto, con corchetes
     * si es IPv6) corresponde a la direccion de loopback enlazada, a
     * {@code localhost} o al nombre inverso de la direccion. Se usa contra
     * ataques DNS-rebinding: un dominio ajeno que resuelva a loopback
     * llegaria con {@code Host} propio y es rechazado.
     *
     * @param hostHeader valor del header {@code Host} (puede ser null).
     * @return {@code true} si el host es aceptable.
     */
    boolean isAllowedHost(String hostHeader) {
        if (hostHeader == null || hostHeader.isEmpty()) {
            return false;
        }
        String name = stripPort(hostHeader);
        return LOCALHOST.equalsIgnoreCase(name)
                || name.equalsIgnoreCase(boundAddress)
                || name.equalsIgnoreCase(boundHostName);
    }

    // ================================================================
    // Enrutado
    // ================================================================

    /**
     * Enruta la peticion, aplica la autenticacion a las rutas
     * {@code /api/} y cierra siempre el intercambio.
     *
     * @param exchange peticion/respuesta HTTP.
     */
    private void route(HttpExchange exchange) {
        // HttpExchange es AutoCloseable: el try-with-resources cierra el
        // intercambio siempre, con la misma semantica que el finally
        // anterior (el cierre ocurre tras el catch).
        try (exchange) {
            String path = exchange.getRequestURI().getPath();
            if (path.startsWith(API_PREFIX)) {
                String denial = authFailure(exchange);
                if (denial != null) {
                    log.warn("Peticion API rechazada: {}", denial);
                    sendError(exchange, STATUS_FORBIDDEN, DENIED_MESSAGE);
                    return;
                }
                routeApi(exchange, path);
            } else {
                routeStatic(exchange, path);
            }
        } catch (IOException e) {
            // Flujo de respuesta cortado a mitad (navegador que cierra la
            // pestana, por ejemplo): no hay nada mas que hacer.
            log.debug("Flujo HTTP interrumpido: {}", e.getMessage());
        }
    }

    /** Enrutado de los ficheros estaticos y del favicon inline. */
    private void routeStatic(HttpExchange exchange, String path) throws IOException {
        if (!METHOD_GET.equals(exchange.getRequestMethod())) {
            sendStatus(exchange, STATUS_METHOD_NOT_ALLOWED, "Metodo no permitido", TEXT_UTF8);
            return;
        }
        switch (path) {
            case "/" -> sendResource(exchange, RESOURCE_INDEX, HTML_UTF8);
            case "/web/app.css" -> sendResource(exchange, RESOURCE_CSS, CSS_UTF8);
            case "/web/app.js" -> sendResource(exchange, RESOURCE_JS, JS_UTF8);
            default -> sendStatus(exchange, STATUS_NOT_FOUND, "No encontrado", TEXT_UTF8);
        }
    }

    /** Enrutado de la API (ya autenticada por {@link #route}). */
    private void routeApi(HttpExchange exchange, String path) throws IOException {
        String method = exchange.getRequestMethod();
        if (METHOD_GET.equals(method)) {
            switch (path) {
                case "/api/info" -> sendInfo(exchange);
                case "/api/logs" -> sendLogs(exchange);
                case "/api/settings" -> sendSettings(exchange);
                default -> sendError(exchange, STATUS_NOT_FOUND, "No encontrado");
            }
        } else if (METHOD_POST.equals(method)) {
            switch (path) {
                case "/api/merge" -> runMerge(exchange);
                case "/api/compare" -> runCompare(exchange);
                case "/api/shutdown" -> runShutdown(exchange);
                case "/api/settings" -> updateSettings(exchange);
                default -> sendError(exchange, STATUS_NOT_FOUND, "No encontrado");
            }
        } else {
            sendError(exchange, STATUS_METHOD_NOT_ALLOWED, "Metodo no permitido");
        }
    }

    // ================================================================
    // Endpoints
    // ================================================================

    /**
     * {@code GET /api/info}: version, build y, si se conocen, las
     * carpetas absolutas de entrada y salida: con ellas la pagina
     * indica al usuario donde dejar sus ficheros y donde saldra el
     * resultado.
     */
    private void sendInfo(HttpExchange exchange) throws IOException {
        String body = "{\"version\":" + Json.quote(Main.APP_VERSION)
                + ",\"build\":" + Json.quote(Main.buildInfoString())
                + optionalField("inputDir", inputDir)
                + optionalField("outputDir", outputDir) + "}";
        sendJson(exchange, STATUS_OK, body);
    }

    /** Campo opcional del JSON de info: omitido si no hay valor. */
    private static String optionalField(String name, String value) {
        if (value == null || value.isEmpty()) {
            return "";
        }
        return ",\"" + name + "\":" + Json.quote(value);
    }

    /** {@code GET /api/logs?after=N}: eventos nuevos desde el cursor. */
    private void sendLogs(HttpExchange exchange) throws IOException {
        long after = parseCursor(exchange.getRequestURI().getQuery());
        List<LogBuffer.Event> events = logBuffer.since(after);
        StringBuilder list = new StringBuilder();
        long next = after;
        for (LogBuffer.Event event : events) {
            next = event.seq;
            if (list.length() > 0) {
                list.append(',');
            }
            list.append("{\"seq\":").append(event.seq)
                    .append(",\"level\":").append(Json.quote(event.level))
                    .append(",\"msg\":").append(Json.quote(event.message))
                    .append('}');
        }
        String body = "{\"next\":" + next + ",\"events\":[" + list + "]}";
        sendJson(exchange, STATUS_OK, body);
    }

    /** {@code GET /api/settings}: los ajustes de salida en vigor. */
    private void sendSettings(HttpExchange exchange) throws IOException {
        if (settings == null) {
            sendError(exchange, STATUS_NOT_FOUND, SETTINGS_UNAVAILABLE);
            return;
        }
        sendJson(exchange, STATUS_OK, settings.toJson());
    }

    /**
     * {@code POST /api/settings}: actualiza y persiste los ajustes de
     * salida. El cuerpo debe ser JSON con las tres claves
     * ({@code mode}, {@code summaryEnabled}, {@code byResponsibleEnabled});
     * cualquier clave ausente o invalida devuelve 400 y no cambia nada.
     */
    private void updateSettings(HttpExchange exchange) throws IOException {
        if (settings == null) {
            sendError(exchange, STATUS_NOT_FOUND, SETTINGS_UNAVAILABLE);
            return;
        }
        String body;
        try {
            body = readBody(exchange);
        } catch (IOException e) {
            sendError(exchange, STATUS_BAD_REQUEST, "cuerpo de ajustes invalido");
            return;
        }
        SettingsRequest request = parseSettingsRequest(body);
        if (request == null) {
            sendError(exchange, STATUS_BAD_REQUEST,
                    "cuerpo de ajustes invalido: se esperan mode, summaryEnabled y "
                            + "byResponsibleEnabled");
            return;
        }
        try {
            settings.update(request.mode, request.summaryEnabled, request.byResponsibleEnabled);
        } catch (IllegalArgumentException e) {
            sendError(exchange, STATUS_BAD_REQUEST, e.getMessage());
            return;
        }
        log.info("Ajustes de salida actualizados desde la interfaz: {}", settings.toJson());
        sendJson(exchange, STATUS_OK, settings.toJson());
    }

    /**
     * Lee el cuerpo de la peticion con un limite de
     * {@value #MAX_BODY_BYTES} bytes: el JSON de ajustes es diminuto y un
     * cuerpo mayor indica una peticion rara (o un cliente corrupto).
     *
     * @param exchange peticion entrante.
     * @return cuerpo como texto UTF-8.
     * @throws IOException si el cuerpo excede el limite o falla la lectura.
     */
    private static String readBody(HttpExchange exchange) throws IOException {
        byte[] raw = exchange.getRequestBody().readNBytes(MAX_BODY_BYTES + 1);
        if (raw.length > MAX_BODY_BYTES) {
            throw new IOException("cuerpo demasiado grande");
        }
        return new String(raw, StandardCharsets.UTF_8);
    }

    /**
     * Parsea el JSON minimo de {@code POST /api/settings}. Estricto y
     * simple: las tres claves deben estar presentes y bien formadas; el
     * orden da igual y las claves desconocidas se ignoran.
     *
     * @param body cuerpo de la peticion.
     * @return la peticion validada, o {@code null} si el cuerpo no es
     *         valido.
     */
    private static SettingsRequest parseSettingsRequest(String body) {
        if (body == null || body.isBlank()) {
            return null;
        }
        String mode = quotedField(body, JSON_MODE);
        String summaryToken = booleanToken(body, JSON_SUMMARY);
        String byResponsibleToken = booleanToken(body, JSON_BY_RESPONSIBLE);
        if (mode == null || summaryToken == null || byResponsibleToken == null
                || !isBoolean(summaryToken) || !isBoolean(byResponsibleToken)
                || !SettingsStore.isValidMode(mode)) {
            return null;
        }
        return new SettingsRequest(mode, Boolean.parseBoolean(summaryToken),
                Boolean.parseBoolean(byResponsibleToken));
    }

    /**
     * Extrae el valor entre comillas de una clave JSON (p. ej.
     * {@code "mode": "completo"}). {@code null} si la clave o el valor
     * no estan.
     */
    private static String quotedField(String body, String key) {
        int keyAt = body.indexOf('"' + key + '"');
        if (keyAt < 0) {
            return null;
        }
        int colon = body.indexOf(':', keyAt + key.length() + 2);
        if (colon < 0) {
            return null;
        }
        int open = body.indexOf('"', colon + 1);
        if (open < 0) {
            return null;
        }
        int close = body.indexOf('"', open + 1);
        if (close < 0) {
            return null;
        }
        String value = body.substring(open + 1, close);
        return value.isEmpty() ? null : value;
    }

    /**
     * Extrae el literal booleano de una clave JSON (p. ej.
     * {@code "summaryEnabled": true}). {@code null} si la clave no esta o
     * el valor no existe.
     */
    private static String booleanToken(String body, String key) {
        int keyAt = body.indexOf('"' + key + '"');
        if (keyAt < 0) {
            return null;
        }
        int colon = body.indexOf(':', keyAt + key.length() + 2);
        if (colon < 0) {
            return null;
        }
        int start = colon + 1;
        while (start < body.length() && Character.isWhitespace(body.charAt(start))) {
            start++;
        }
        int end = start;
        while (end < body.length() && (Character.isLetter(body.charAt(end))
                || Character.isDigit(body.charAt(end)))) {
            end++;
        }
        String token = body.substring(start, end);
        return token.isEmpty() ? null : token;
    }

    /** {@code true} si {@code token} es el literal {@code true} o {@code false}. */
    private static boolean isBoolean(String token) {
        return "true".equals(token) || "false".equals(token);
    }

    /** Peticion {@code POST /api/settings} ya validada. */
    private static final class SettingsRequest {

        private final String mode;
        private final boolean summaryEnabled;
        private final boolean byResponsibleEnabled;

        private SettingsRequest(String mode, boolean summaryEnabled,
                                boolean byResponsibleEnabled) {
            this.mode = mode;
            this.summaryEnabled = summaryEnabled;
            this.byResponsibleEnabled = byResponsibleEnabled;
        }
    }

    /** {@code POST /api/merge}: la Opcion 1 del menu. */
    private void runMerge(HttpExchange exchange) throws IOException {
        if (!busy.compareAndSet(false, true)) {
            sendError(exchange, STATUS_CONFLICT, BUSY_MESSAGE);
            return;
        }
        long startedAt = System.currentTimeMillis();
        int exitCode;
        try {
            exitCode = mergeRunner.apply(null);
        } catch (RuntimeException e) {
            // App.run ya captura lo suyo; esto es la red de seguridad del
            // servidor para no tumbar el hilo HTTP.
            log.error("La fusion ha lanzado una excepcion no capturada: {}", e.getMessage(), e);
            exitCode = App.EXIT_RUNTIME;
        } finally {
            busy.set(false);
        }
        sendRunResult(exchange, exitCode, System.currentTimeMillis() - startedAt);
    }

    /** {@code POST /api/compare}: la Opcion 2 del menu (no cierra nada). */
    private void runCompare(HttpExchange exchange) throws IOException {
        if (!busy.compareAndSet(false, true)) {
            sendError(exchange, STATUS_CONFLICT, BUSY_MESSAGE);
            return;
        }
        long startedAt = System.currentTimeMillis();
        int exitCode;
        try {
            exitCode = compareRunner.getAsInt();
        } catch (RuntimeException e) {
            log.error("El comprobador ha lanzado una excepcion no capturada: {}", e.getMessage(), e);
            exitCode = App.EXIT_RUNTIME;
        } finally {
            busy.set(false);
        }
        sendRunResult(exchange, exitCode, System.currentTimeMillis() - startedAt);
    }

    /** {@code POST /api/shutdown}: la Opcion 3, "Salir sin hacer nada". */
    private void runShutdown(HttpExchange exchange) throws IOException {
        sendJson(exchange, STATUS_OK, "{\"ok\":true}");
        Thread closer = new Thread(() -> {
            try {
                Thread.sleep(SHUTDOWN_DELAY_MS);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                return;
            }
            shutdownAction.run();
        }, "web-shutdown");
        closer.setDaemon(true);
        closer.start();
    }

    // ================================================================
    // Autenticacion
    // ================================================================

    /**
     * Motivo de rechazo de la peticion, o {@code null} si esta permitida.
     *
     * <p>Comprueba en orden: header {@code Host}, header {@code Origin}
     * (solo si el navegador lo envia; debe ser exactamente nuestro
     * esquema+host) y token (header o query).</p>
     *
     * @param exchange peticion entrante.
     * @return {@code null} si pasa, o el motivo del rechazo.
     */
    private String authFailure(HttpExchange exchange) {
        String host = exchange.getRequestHeaders().getFirst(HEADER_HOST);
        if (!isAllowedHost(host)) {
            return "host no permitido: " + host;
        }
        String origin = exchange.getRequestHeaders().getFirst(HEADER_ORIGIN);
        if (origin != null && !origin.equalsIgnoreCase("http://" + host)) {
            return "origen no permitido: " + origin;
        }
        if (!hasValidToken(exchange)) {
            return "token no valido";
        }
        return null;
    }

    /** Compara el token presentado (header o query) con el del servidor. */
    private boolean hasValidToken(HttpExchange exchange) {
        String presented = exchange.getRequestHeaders().getFirst(TOKEN_HEADER);
        if (presented == null) {
            presented = queryParam(exchange.getRequestURI().getQuery(), TOKEN_PARAM);
        }
        return token.equals(presented);
    }

    // ================================================================
    // Serializacion de respuestas y utilidades
    // ================================================================

    /** Cuerpo comun de ejecuciones completadas: exit code + duracion. */
    private static String runResultBody(int exitCode, long elapsedMs) {
        return "{\"exitCode\":" + exitCode + ",\"ms\":" + elapsedMs + "}";
    }

    private static void sendRunResult(HttpExchange exchange, int exitCode, long elapsedMs)
            throws IOException {
        sendJson(exchange, STATUS_OK, runResultBody(exitCode, elapsedMs));
    }

    private static void sendError(HttpExchange exchange, int status, String message)
            throws IOException {
        sendJson(exchange, status, ERROR_JSON_PREFIX + Json.quote(message) + "}");
    }

    private static void sendJson(HttpExchange exchange, int status, String body)
            throws IOException {
        sendBytes(exchange, status, body.getBytes(StandardCharsets.UTF_8), JSON_UTF8);
    }

    private static void sendStatus(HttpExchange exchange, int status, String body, String contentType)
            throws IOException {
        sendBytes(exchange, status, body.getBytes(StandardCharsets.UTF_8), contentType);
    }

    private static void sendResource(HttpExchange exchange, String resource, String contentType)
            throws IOException {
        try (InputStream in = WebServer.class.getResourceAsStream(resource)) {
            if (in == null) {
                sendStatus(exchange, STATUS_NOT_FOUND, "Recurso ausente: " + resource, TEXT_UTF8);
                return;
            }
            sendBytes(exchange, STATUS_OK, in.readAllBytes(), contentType);
        }
    }

    private static void sendBytes(HttpExchange exchange, int status, byte[] bytes, String contentType)
            throws IOException {
        exchange.getResponseHeaders().set("Content-Type", contentType);
        exchange.getResponseHeaders().set("Cache-Control", CACHE_NO_STORE);
        exchange.sendResponseHeaders(status, bytes.length);
        try (OutputStream out = exchange.getResponseBody()) {
            out.write(bytes);
        }
    }

    /** Lee el cursor {@code after} de la query; no numerico o ausente = 0. */
    private static long parseCursor(String query) {
        String raw = queryParam(query, "after");
        if (raw == null) {
            return 0L;
        }
        try {
            return Long.parseLong(raw);
        } catch (NumberFormatException e) {
            return 0L;
        }
    }

    /** Extrae un parametro de una query cruda ({@code a=1&b=2}). */
    private static String queryParam(String query, String name) {
        if (query == null || query.isEmpty()) {
            return null;
        }
        for (String pair : query.split("&")) {
            int eq = pair.indexOf('=');
            if (eq > 0 && pair.substring(0, eq).equals(name)) {
                return pair.substring(eq + 1);
            }
        }
        return null;
    }

    /** Quita el puerto de un header Host (con su manejo de IPv6 entre corchetes). */
    private static String stripPort(String hostHeader) {
        if (hostHeader.startsWith("[")) {
            int close = hostHeader.indexOf(']');
            return close > 0 ? hostHeader.substring(1, close) : hostHeader;
        }
        int colon = hostHeader.lastIndexOf(':');
        if (colon >= 0 && hostHeader.indexOf(':') == colon) {
            return hostHeader.substring(0, colon);
        }
        return hostHeader;
    }

    /** Formatea una direccion para URL: IPv6 entre corchetes, IPv4 tal cual. */
    private static String formatHost(String address) {
        if (address.indexOf(':') >= 0) {
            return "[" + address + "]";
        }
        return address;
    }

    /** Token aleatorio de 128 bits en base64url sin relleno. */
    private static String newToken() {
        byte[] raw = new byte[TOKEN_BYTES];
        RANDOM.nextBytes(raw);
        return Base64.getUrlEncoder().withoutPadding().encodeToString(raw);
    }

    /** Hilos del pool: con nombre y daemon, para no retener la JVM. */
    private static final class WebThreadFactory implements ThreadFactory {

        private static final AtomicInteger COUNT = new AtomicInteger(1);

        @Override
        public Thread newThread(Runnable task) {
            Thread thread = new Thread(task, "excel-merger-web-" + COUNT.getAndIncrement());
            thread.setDaemon(true);
            return thread;
        }
    }
}

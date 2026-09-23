package com.excelmerger.web;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.net.ServerSocket;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests del servidor web local: autenticacion, mapeo 1:1 con las
 * opciones del menu de terminal, plaza unica de ejecucion, registro y
 * salida limpia.
 */
class WebServerTest {

    private WebServer server;
    private HttpClient client;
    private LogBuffer buffer;

    private final AtomicBoolean shutdownCalled = new AtomicBoolean(false);
    private final AtomicInteger mergeResult = new AtomicInteger(0);
    private final AtomicInteger compareResult = new AtomicInteger(0);
    private final AtomicReference<String> mergeConfig = new AtomicReference<>("sin-llamar");
    private volatile CountDownLatch mergeEntered;
    private volatile CountDownLatch mergeRelease;

    @BeforeEach
    void setUp() throws IOException {
        buffer = new LogBuffer();
        server = new WebServer(this::mergeRunner, this::compareRunner,
                () -> shutdownCalled.set(true), buffer);
        client = HttpClient.newHttpClient();
    }

    @AfterEach
    void tearDown() {
        server.stop();
    }

    // ================================================================
    // Ficheros estaticos
    // ================================================================

    @Test
    void indexSeSirveSinToken() throws Exception {
        HttpResponse<String> response = get("/", null);

        assertThat(response.statusCode()).isEqualTo(200);
        assertThat(response.body()).contains("<title>Excel Merger · Torre de mando</title>");
    }

    @Test
    void recursosEstaticosIncluyenContentType() throws Exception {
        HttpResponse<String> css = get("/web/app.css", null);

        assertThat(css.statusCode()).isEqualTo(200);
        assertThat(css.headers().firstValue("Content-Type").orElse(""))
                .startsWith("text/css");
    }

    // ================================================================
    // Autenticacion de la API
    // ================================================================

    @Test
    void apiRechazaSinToken() throws Exception {
        assertThat(get("/api/info", null).statusCode()).isEqualTo(403);
    }

    @Test
    void apiRechazaTokenInvalido() throws Exception {
        assertThat(get("/api/info", "token-falso").statusCode()).isEqualTo(403);
    }

    @Test
    void apiRechazaOriginExterno() throws Exception {
        HttpResponse<String> response = postWithOrigin("/api/merge",
                server.token(), "https://evil.example");

        assertThat(response.statusCode()).isEqualTo(403);
        assertThat(mergeConfig.get()).isEqualTo("sin-llamar");
    }

    @Test
    void hostDenegadoParaDominiosAjenosYOkParaLoopback() {
        String loopbackHost = server.baseUrl().substring("http://".length());

        assertThat(server.isAllowedHost("localhost:" + server.port())).isTrue();
        assertThat(server.isAllowedHost(loopbackHost)).isTrue();
        assertThat(server.isAllowedHost("evil.example:" + server.port())).isFalse();
        assertThat(server.isAllowedHost(null)).isFalse();
        assertThat(server.isAllowedHost("")).isFalse();
    }

    // ================================================================
    // Endpoints de la API autenticada
    // ================================================================

    @Test
    void infoDevuelveLaVersionActual() throws Exception {
        HttpResponse<String> response = get("/api/info", server.token());

        assertThat(response.statusCode()).isEqualTo(200);
        assertThat(response.body()).contains("\"version\":\"" + com.excelmerger.Main.APP_VERSION + "\"");
    }

    @Test
    void infoOmiteLasCarpetasPorDefecto() throws Exception {
        HttpResponse<String> response = get("/api/info", server.token());

        assertThat(response.body()).doesNotContain("inputDir").doesNotContain("outputDir");
    }

    @Test
    void infoIncluyeLasCarpetasCuandoSeIndican() throws Exception {
        WebServer conCarpetas = new WebServer(this::mergeRunner, this::compareRunner,
                () -> shutdownCalled.set(true), buffer, List.of(0), null,
                "/mi/entrada", "/mi/salida");
        try {
            HttpRequest request = HttpRequest.newBuilder(
                    URI.create(conCarpetas.baseUrl() + "/api/info"))
                    .header(WebServer.TOKEN_HEADER, conCarpetas.token())
                    .GET()
                    .build();
            HttpResponse<String> response = client.send(request,
                    HttpResponse.BodyHandlers.ofString());

            assertThat(response.statusCode()).isEqualTo(200);
            assertThat(response.body())
                    .contains("\"inputDir\":\"/mi/entrada\"")
                    .contains("\"outputDir\":\"/mi/salida\"");
        } finally {
            conCarpetas.stop();
        }
    }

    @Test
    void puertoOcupadoSeSaltaAlSiguiente() throws Exception {
        try (ServerSocket ocupado = new ServerSocket(0)) {
            int busyPort = ocupado.getLocalPort();
            WebServer desplazado = new WebServer(this::mergeRunner, this::compareRunner,
                    () -> shutdownCalled.set(true), buffer, List.of(busyPort, 0),
                    null, null, null);
            try {
                assertThat(desplazado.port()).isNotEqualTo(busyPort);
            } finally {
                desplazado.stop();
            }
        }
    }

    @Test
    void mergeUsaRunnerConConfigPorDefecto() throws Exception {
        HttpResponse<String> response = post("/api/merge", server.token());

        assertThat(response.statusCode()).isEqualTo(200);
        assertThat(response.body()).contains("\"exitCode\":0").contains("\"ms\":");
        assertThat(mergeConfig.get()).isNull();
    }

    @Test
    void mergeDevuelveSuExitCodePropio() throws Exception {
        mergeResult.set(2);

        HttpResponse<String> response = post("/api/merge", server.token());

        assertThat(response.statusCode()).isEqualTo(200);
        assertThat(response.body()).contains("\"exitCode\":2");
    }

    @Test
    void mergeSegundaLlamadaMientrasOcupadoDevuelve409() throws Exception {
        mergeEntered = new CountDownLatch(1);
        mergeRelease = new CountDownLatch(1);
        AtomicReference<HttpResponse<String>> first = new AtomicReference<>();
        AtomicReference<Throwable> failure = new AtomicReference<>();

        Thread caller = new Thread(() -> {
            try {
                first.set(post("/api/merge", server.token()));
            } catch (Throwable t) {
                failure.set(t);
            }
        });
        caller.start();
        try {
            assertThat(mergeEntered.await(5, TimeUnit.SECONDS)).isTrue();

            HttpResponse<String> second = post("/api/merge", server.token());
            assertThat(second.statusCode()).isEqualTo(409);
        } finally {
            mergeRelease.countDown();
        }
        caller.join(TimeUnit.SECONDS.toMillis(10));

        assertThat(failure.get()).isNull();
        assertThat(first.get()).isNotNull();
        assertThat(first.get().statusCode()).isEqualTo(200);
    }

    @Test
    void compareUsaSuRunnerYDevuelveSuExitCode() throws Exception {
        compareResult.set(4);

        HttpResponse<String> response = post("/api/compare", server.token());

        assertThat(response.statusCode()).isEqualTo(200);
        assertThat(response.body()).contains("\"exitCode\":4");
    }

    @Test
    void logsDevuelveEventosConJsonEscapado() throws Exception {
        buffer.add("INFO", "comilla \" y salto\n de linea");

        HttpResponse<String> response = get("/api/logs?after=0", server.token());

        assertThat(response.statusCode()).isEqualTo(200);
        assertThat(response.body())
                .contains("\"level\":\"INFO\"")
                .contains("comilla \\\" y salto\\n de linea");
    }

    @Test
    void logsFiltranPorElCursorDelCliente() throws Exception {
        buffer.add("INFO", "primero-aaaa");
        buffer.add("INFO", "segundo-bbbb");

        HttpResponse<String> response = get("/api/logs?after=1", server.token());

        assertThat(response.statusCode()).isEqualTo(200);
        assertThat(response.body())
                .doesNotContain("primero-aaaa")
                .contains("segundo-bbbb")
                .contains("\"next\":2");
    }

    @Test
    void logsCursorNoNumericoVuelveAlPrincipio() throws Exception {
        buffer.add("INFO", "evento-unico");

        HttpResponse<String> response = get("/api/logs?after=zzz", server.token());

        assertThat(response.statusCode()).isEqualTo(200);
        assertThat(response.body()).contains("evento-unico");
    }

    @Test
    void shutdownInvocaLaAccionDeSalir() throws Exception {
        HttpResponse<String> response = post("/api/shutdown", server.token());

        assertThat(response.statusCode()).isEqualTo(200);
        assertThat(response.body()).isEqualTo("{\"ok\":true}");

        long deadline = System.currentTimeMillis() + 3000;
        while (!shutdownCalled.get() && System.currentTimeMillis() < deadline) {
            Thread.sleep(50);
        }
        assertThat(shutdownCalled).isTrue();
    }

    @Test
    void metodoNoSoportadoDevuelve405() throws Exception {
        HttpRequest request = HttpRequest.newBuilder(URI.create(server.baseUrl() + "/api/info"))
                .DELETE()
                .header(WebServer.TOKEN_HEADER, server.token())
                .build();

        HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

        assertThat(response.statusCode()).isEqualTo(405);
    }

    // ================================================================
    // v4.1.0: ajustes de salida en runtime (/api/settings)
    // ================================================================

    @Test
    void settingsSinStoreDevuelve404() throws Exception {
        // El servidor de pruebas por defecto no lleva SettingsStore: los
        // ajustes no estan disponibles y la API responde 404 (nunca
        // inventa valores).
        HttpResponse<String> response = get("/api/settings", server.token());

        assertThat(response.statusCode()).isEqualTo(404);
    }

    @Test
    void settingsDevuelvenLosValoresEnVigor(@TempDir Path tmp) throws Exception {
        WebServer conAjustes = serverConAjustes(SettingsStore.load(tmp.resolve("s.properties")));
        try {
            HttpResponse<String> response = get(conAjustes, "/api/settings", conAjustes.token());

            assertThat(response.statusCode()).isEqualTo(200);
            assertThat(response.body())
                    .contains("\"mode\":\"completo\"")
                    .contains("\"summaryEnabled\":true")
                    .contains("\"byResponsibleEnabled\":true");
        } finally {
            conAjustes.stop();
        }
    }

    @Test
    void postSettingsActualizaYPersiste(@TempDir Path tmp) throws Exception {
        Path file = tmp.resolve("web-settings.properties");
        WebServer conAjustes = serverConAjustes(SettingsStore.load(file));
        try {
            String body = "{\"mode\":\"responsables\",\"summaryEnabled\":false,"
                    + "\"byResponsibleEnabled\":false}";
            HttpResponse<String> response =
                    postJson(conAjustes, "/api/settings", conAjustes.token(), body);

            assertThat(response.statusCode()).isEqualTo(200);
            assertThat(response.body())
                    .contains("\"mode\":\"responsables\"")
                    .contains("\"summaryEnabled\":false")
                    .contains("\"byResponsibleEnabled\":false");

            // Persistido: una recarga del store lo recupera, y el GET lo confirma.
            SettingsStore reloaded = SettingsStore.load(file);
            assertThat(reloaded.mode()).isEqualTo("responsables");
            assertThat(reloaded.summaryEnabled()).isFalse();

            HttpResponse<String> fresh = get(conAjustes, "/api/settings", conAjustes.token());
            assertThat(fresh.body()).contains("\"mode\":\"responsables\"");
        } finally {
            conAjustes.stop();
        }
    }

    @Test
    void postSettingsConModoInvalidoDevuelve400(@TempDir Path tmp) throws Exception {
        Path file = tmp.resolve("web-settings.properties");
        WebServer conAjustes = serverConAjustes(SettingsStore.load(file));
        try {
            String body = "{\"mode\":\"EXPERIMENTAL\",\"summaryEnabled\":true,"
                    + "\"byResponsibleEnabled\":true}";
            HttpResponse<String> response =
                    postJson(conAjustes, "/api/settings", conAjustes.token(), body);

            assertThat(response.statusCode()).isEqualTo(400);
            // Nada cambia: sigue el default.
            assertThat(SettingsStore.load(file).mode()).isEqualTo("completo");
            HttpResponse<String> fresh = get(conAjustes, "/api/settings", conAjustes.token());
            assertThat(fresh.body()).contains("\"mode\":\"completo\"");
        } finally {
            conAjustes.stop();
        }
    }

    @Test
    void postSettingsConJsonMalformadoDevuelve400(@TempDir Path tmp) throws Exception {
        WebServer conAjustes = serverConAjustes(SettingsStore.load(tmp.resolve("s.properties")));
        try {
            HttpResponse<String> response =
                    postJson(conAjustes, "/api/settings", conAjustes.token(), "{\"mode\":");

            assertThat(response.statusCode()).isEqualTo(400);
        } finally {
            conAjustes.stop();
        }
    }

    @Test
    void postSettingsConCampoAusenteDevuelve400(@TempDir Path tmp) throws Exception {
        WebServer conAjustes = serverConAjustes(SettingsStore.load(tmp.resolve("s.properties")));
        try {
            // Faltan summaryEnabled y byResponsibleEnabled: la API es
            // estricta (el cliente siempre envia las tres claves).
            String body = "{\"mode\":\"completo\"}";
            HttpResponse<String> response =
                    postJson(conAjustes, "/api/settings", conAjustes.token(), body);

            assertThat(response.statusCode()).isEqualTo(400);
        } finally {
            conAjustes.stop();
        }
    }

    // ================================================================
    // Helpers
    // ================================================================

    private int mergeRunner(String configPath) {
        mergeConfig.set(configPath);
        CountDownLatch entered = mergeEntered;
        CountDownLatch release = mergeRelease;
        if (entered != null && release != null) {
            entered.countDown();
            try {
                release.await(5, TimeUnit.SECONDS);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
            }
        }
        return mergeResult.get();
    }

    private int compareRunner() {
        return compareResult.get();
    }

    private HttpResponse<String> get(String path, String tokenValue)
            throws IOException, InterruptedException {
        return get(server, path, tokenValue);
    }

    private HttpResponse<String> get(WebServer target, String path, String tokenValue)
            throws IOException, InterruptedException {
        HttpRequest.Builder builder =
                HttpRequest.newBuilder(URI.create(target.baseUrl() + path)).GET();
        if (tokenValue != null) {
            builder.header(WebServer.TOKEN_HEADER, tokenValue);
        }
        return client.send(builder.build(), HttpResponse.BodyHandlers.ofString());
    }

    private HttpResponse<String> post(String path, String tokenValue)
            throws IOException, InterruptedException {
        HttpRequest.Builder builder = HttpRequest.newBuilder(URI.create(server.baseUrl() + path))
                .POST(HttpRequest.BodyPublishers.noBody());
        if (tokenValue != null) {
            builder.header(WebServer.TOKEN_HEADER, tokenValue);
        }
        return client.send(builder.build(), HttpResponse.BodyHandlers.ofString());
    }

    private HttpResponse<String> postWithOrigin(String path, String tokenValue, String origin)
            throws IOException, InterruptedException {
        HttpRequest request = HttpRequest.newBuilder(URI.create(server.baseUrl() + path))
                .POST(HttpRequest.BodyPublishers.noBody())
                .header(WebServer.TOKEN_HEADER, tokenValue)
                .header("Origin", origin)
                .build();
        return client.send(request, HttpResponse.BodyHandlers.ofString());
    }

    /** Servidor de pruebas con ajustes de salida y puerto efimero. */
    private WebServer serverConAjustes(SettingsStore store) throws IOException {
        return new WebServer(this::mergeRunner, this::compareRunner,
                () -> shutdownCalled.set(true), buffer, List.of(0), null, null, null, store);
    }

    private HttpResponse<String> postJson(String path, String tokenValue, String body)
            throws IOException, InterruptedException {
        return postJson(server, path, tokenValue, body);
    }

    private HttpResponse<String> postJson(WebServer target, String path, String tokenValue,
                                          String body) throws IOException, InterruptedException {
        HttpRequest request = HttpRequest.newBuilder(URI.create(target.baseUrl() + path))
                .POST(HttpRequest.BodyPublishers.ofString(body, StandardCharsets.UTF_8))
                .header(WebServer.TOKEN_HEADER, tokenValue)
                .build();
        return client.send(request, HttpResponse.BodyHandlers.ofString());
    }
}

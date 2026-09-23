package com.excelmerger.web;

import com.excelmerger.App;
import com.excelmerger.ConfigLoader;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Desktop;
import java.awt.GraphicsEnvironment;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.util.Base64;
import java.util.Locale;
import java.util.concurrent.TimeUnit;

/**
 * Punto de entrada de la interfaz web local de Excel Merger (v4.0.0).
 * Sustituye a la experiencia del menu de {@code run.bat} por un
 * servidor HTTP en loopback que se sirve a si mismo en el navegador.
 *
 * <p>Flujo de arranque pensado para que "hacer que funcione" sea un
 * solo doble clic:</p>
 * <ol>
 *   <li>{@link LogCapture#install()}: log y {@code System.out} pasan a
 *       alimentar el buffer de la interfaz (ademas de la terminal y el
 *       fichero de log, que no cambian).</li>
 *   <li>Si la {@link WebSession} guardada responde, el servidor ya esta
 *       vivo: se abre el navegador y el proceso termina (el segundo
 *       doble clic no levanta otra copia).</li>
 *   <li>Se preparan las carpetas de trabajo (entrada y salida, segun
 *       {@code config.properties}; se crean si faltan) y se muestran en
 *       la pagina y en la terminal.</li>
 *   <li>{@link WebServer#start}: direccion estable (puerto de la
 *       sesion, luego {@code 7420..7430}, luego uno libre) y token
 *       persistido en {@code web.properties}: misma URL siempre.</li>
 *   <li>Se imprime la URL y se abre el navegador por defecto. Si es la
 *       imagen empaquetada en Windows, ademas se crea (o repone) el
 *       acceso directo en el escritorio.</li>
 *   <li>El proceso queda a la espera: la opcion "Salir" de la interfaz
 *       cierra con exit code 0 (equivalente a la Opcion 3 del menu) y
 *       Ctrl+C en terminal corta por signal.</li>
 * </ol>
 *
 * <p>v4.0.0: la deteccion de modos vive en
 * {@link com.excelmerger.Main} (sin argumentos = esta interfaz,
 * {@code --cli} = menu de terminal) y el empaquetado portable se hace
 * con {@code package.bat} (jpackage app-image + zip, nunca
 * instalador).</p>
 */
public final class WebLauncher {

    /**
     * Logger estatico: su creacion solo inicializa el provider SLF4J y
     * NO emite eventos, condicion necesaria de la invariante de
     * {@link LogCapture#install()} (nada loguea antes de montar el
     * espejo).
     */
    private static final Logger log = LoggerFactory.getLogger(WebLauncher.class);

    private static final String EXE_NAME = "ExcelMerger.exe";
    private static final String SHORTCUT_NAME = "ExcelMerger.lnk";
    private static final int PROBE_TIMEOUT_MS = 900;
    private static final int SHORTCUT_TIMEOUT_S = 8;

    private WebLauncher() {
        // Clase de utilidad
    }

    /**
     * Arranca (o reapunta al) la interfaz web local. Exit codes: 0
     * via "Salir" de la interfaz (accion {@link #exitOk()} invocada por
     * el servidor), 0 tambien al detectar una instancia ya viva, y 1 si
     * el servidor no puede arrancar.
     *
     * @param args ignorados: el modo ya lo ha decidido {@code Main}.
     * @throws InterruptedException si el hilo principal se interrumpe
     *         esperando (se restaura el flag y se sale limpiamente).
     */
    public static void main(String[] args) throws InterruptedException {
        LogBuffer buffer = LogCapture.install();

        // 1) Direccion de la sesion anterior: si el servidor sigue vivo,
        //    no arrancamos otra copia: abrimos el navegador y salimos.
        WebSession session = WebSession.load();
        String running = runningUrl(session);
        if (running != null) {
            System.out.println();
            System.out.println("Excel Merger ya esta en ejecucion.");
            System.out.println("  Abriendo " + running);
            System.out.println();
            openBrowser(running);
            return;
        }

        // 2) Carpetas de trabajo (creandolas si faltan): la pagina y la
        //    terminal muestran donde dejar los ficheros y donde saldra
        //    el resultado.
        String[] workDirs = prepareWorkDirs();

        // 3) Servidor con direccion estable y token persistido. Los
        //    ajustes de salida (modo y resumen) los carga la interfaz
        //    desde web-settings.properties y se aplican a cada fusion.
        SettingsStore settings = SettingsStore.load();
        WebServer server;
        try {
            server = WebServer.start(buffer, session.lastPort(), session.token(),
                    workDirs[0], workDirs[1], settings);
        } catch (IOException e) {
            log.error("No se pudo arrancar el servidor web local: {}", e.getMessage(), e);
            System.exit(App.EXIT_RUNTIME);
            return;
        }
        WebSession.save(server.port(), server.token());

        String url = server.url();
        System.out.println();
        System.out.println("Excel Merger (interfaz web local)");
        System.out.println("  " + url);
        if (workDirs[0] != null) {
            System.out.println("  Tus Excel: " + workDirs[0]);
        }
        if (workDirs[1] != null) {
            System.out.println("  Resultado: " + workDirs[1]);
        }
        System.out.println("  Salida: modo '" + settings.mode() + "', hoja Resumen "
                + (settings.summaryEnabled() ? "si" : "no")
                + ", tabla por responsable " + (settings.byResponsibleEnabled() ? "si" : "no")
                + " (cambiable desde la interfaz)");
        System.out.println("  El 'Salir' de la interfaz cierra el proceso con codigo 0.");
        System.out.println("  Ctrl+C en esta terminal tambien termina.");
        System.out.println();

        createDesktopShortcut();
        openBrowser(url);
        Thread.currentThread().join();
    }

    /**
     * Devuelve la URL si la sesion apunta a un servidor vivo de esta
     * misma maquina (responde a {@code /api/info} con su token);
     * {@code null} en caso contrario. Es lo que hace que un segundo
     * doble clic solo abra el navegador en lugar de levantar otra
     * copia del servidor.
     *
     * @param session sesion persistida.
     * @return la URL viva o {@code null}.
     */
    private static String runningUrl(WebSession session) {
        String infoUrl = session.infoUrl();
        if (infoUrl == null) {
            return null;
        }
        try (HttpClient client = HttpClient.newBuilder()
                .connectTimeout(Duration.ofMillis(PROBE_TIMEOUT_MS))
                .build()) {
            HttpRequest request = HttpRequest.newBuilder(URI.create(infoUrl))
                    .timeout(Duration.ofMillis(PROBE_TIMEOUT_MS))
                    .GET()
                    .build();
            HttpResponse<String> response =
                    client.send(request, HttpResponse.BodyHandlers.ofString());
            if (response.statusCode() == 200 && response.body().contains("\"version\"")) {
                return session.savedUrl();
            }
            return null;
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            return null;
        } catch (IOException | RuntimeException e) {
            return null;
        }
    }

    /**
     * Prepara las carpetas de trabajo y devuelve sus rutas absolutas
     * (posicion 0 = entrada de {@code input.directory}, posicion 1 =
     * carpeta madre de {@code output.file}), creandolas si faltan.
     * Best effort: si el config no carga, devuelve dos {@code null} y
     * la pagina omite el panel de rutas; el motor no depende de esto
     * (el, valdose por si, carga y valida el config en cada ejecucion).
     *
     * @return par {entrada, salida} en rutas absolutas.
     */
    private static String[] prepareWorkDirs() {
        try {
            ConfigLoader config = new ConfigLoader();
            String input = config.get("input.directory", null);
            String output = config.get("output.file", null);
            Path inputDir = isBlank(input)
                    ? null : Paths.get(input).toAbsolutePath().normalize();
            Path outputDir = isBlank(output)
                    ? null : Paths.get(output).toAbsolutePath().normalize().getParent();
            if (inputDir != null) {
                Files.createDirectories(inputDir);
            }
            if (outputDir != null) {
                Files.createDirectories(outputDir);
            }
            return new String[] {toString(inputDir), toString(outputDir)};
        } catch (RuntimeException | IOException e) {
            log.warn("No se pudieron preparar las carpetas de trabajo: {}", e.getMessage());
            return new String[] {null, null};
        }
    }

    /**
     * Crea (o repone) el acceso directo {@code ExcelMerger.lnk} en el
     * escritorio apuntando a este {@code ExcelMerger.exe}, unicamente
     * cuando se ejecuta la imagen empaquetada en Windows (en desarrollo
     * no hay .exe junto al working directory: se omite). Best effort:
     * cualquier fallo solo deja sin icono, nunca rompe el arranque.
     */
    private static void createDesktopShortcut() {
        Path exe = Paths.get(EXE_NAME).toAbsolutePath();
        Path parent = exe.getParent();
        if (!isWindows() || parent == null || !Files.isRegularFile(exe)) {
            return;
        }
        try {
            String folder = parent.toString();
            String script = "$d=[Environment]::GetFolderPath('Desktop');"
                    + "if($d){$s=New-Object -ComObject WScript.Shell;"
                    + "$l=$s.CreateShortcut((Join-Path $d '" + SHORTCUT_NAME + "'));"
                    + "$l.TargetPath='" + psQuote(exe.toString()) + "';"
                    + "$l.WorkingDirectory='" + psQuote(folder) + "';$l.Save()}";
            String encoded = Base64.getEncoder()
                    .encodeToString(script.getBytes(StandardCharsets.UTF_16LE));
            Process process = new ProcessBuilder("powershell.exe", "-NoProfile",
                    "-NonInteractive", "-EncodedCommand", encoded)
                    .redirectOutput(ProcessBuilder.Redirect.DISCARD)
                    .redirectError(ProcessBuilder.Redirect.DISCARD)
                    .start();
            if (process.waitFor(SHORTCUT_TIMEOUT_S, TimeUnit.SECONDS)
                    && process.exitValue() == 0) {
                log.info("Acceso directo en el escritorio: {} (repite este icono "
                        + "para abrir la interfaz).", SHORTCUT_NAME);
            }
        } catch (IOException e) {
            log.debug("Acceso directo no disponible ({}): se omite.", e.getMessage());
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }
    }

    private static boolean isWindows() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT)
                .startsWith("windows");
    }

    private static String psQuote(String value) {
        return value.replace("'", "''");
    }

    private static boolean isBlank(String value) {
        return value == null || value.isBlank();
    }

    private static String toString(Path path) {
        return path == null ? null : path.toString();
    }

    /**
     * Accion inyectada en {@link WebServer#start} para el boton "Salir"
     * de la interfaz: cierra el proceso con exit code 0, el equivalente
     * exacto a la Opcion 3 del menu de terminal ("Salir sin hacer
     * nada"). La espera previa a la peticion de cierre la gestiona el
     * servidor para que el navegador reciba la respuesta.
     *
     * <p>{@code DoNotTerminateVM} se suprime a proposito: cerrar el
     * proceso ES la accion solicitada (contrato 1:1 con la Opcion 3 del
     * menu y con {@code Main.main}, que tambien llama a
     * {@code System.exit}); la regla apunta a apps J2EE y esta CLI local
     * no lo es.</p>
     */
    @SuppressWarnings("PMD.DoNotTerminateVM")
    static void exitOk() {
        System.exit(App.EXIT_OK);
    }

    /**
     * Abre la URL en el navegador por defecto. Si no hay entorno grafico
     * o el navegador no esta disponible, no falla: la URL ya se ha
     * impreso en la terminal y queda en el log.
     *
     * @param url URL de la interfaz, ya con el token.
     */
    private static void openBrowser(String url) {
        try {
            if (!GraphicsEnvironment.isHeadless() && Desktop.isDesktopSupported()
                    && Desktop.getDesktop().isSupported(Desktop.Action.BROWSE)) {
                Desktop.getDesktop().browse(URI.create(cacheBusted(url)));
                return;
            }
            log.info("Entorno sin navegador automatico; abre a mano: {}", url);
        } catch (IOException | RuntimeException e) {
            log.warn("No se pudo abrir el navegador automaticamente ({}); abre a mano: {}",
                    e.getMessage(), url);
        }
    }

    /**
     * Anade un parametro de version a la URL abierta en el navegador.
     * Como el token es estable, sin esto el navegador podria reutilizar
     * una pagina antigua en cache y el rediseno no se veria tras una
     * actualizacion; el parametro fuerza una carga fresca de index.html
     * en cada arranque (los assets llevan su propia version en las URL).
     *
     * @param url URL de la interfaz, ya con el token.
     * @return la misma URL con {@code v=4.4.0} anadido a la query.
     */
    private static String cacheBusted(String url) {
        return url + (url.indexOf('?') >= 0 ? "&v=4.4.0" : "?v=4.4.0");
    }
}

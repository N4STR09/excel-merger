package com.excelmerger.web;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.regex.Pattern;

/**
 * Sesion persistente de la interfaz web local (v4.0.0): conserva el
 * puerto y el token entre arranques para que la direccion sea
 * <b>siempre la misma</b> (marcable y compartible) en lugar de un
 * puerto efimero y un token nuevos en cada ejecucion.
 *
 * <p>Se guarda en {@value #FILE_NAME} del directorio de trabajo, junto
 * al {@code config.properties}: en la imagen portatil viaja con la
 * carpeta y en ejecuciones desde codigo fuente se ignora en git. El
 * fichero es opcional: si no existe, esta corrupto o contiene valores
 * invalidos, se parte de una sesion vacia (token nuevo y puerto por
 * defecto) sin que el arranque falle.</p>
 */
final class WebSession {

    /** Fichero de sesion en el directorio de trabajo. */
    static final String FILE_NAME = "web.properties";

    private static final Logger log = LoggerFactory.getLogger(WebSession.class);

    /** Token aceptado: mismo alfabeto que el generado, minimo 8. */
    private static final Pattern TOKEN_OK = Pattern.compile("[A-Za-z0-9_-]{8,}");

    private static final String PORT_KEY = "port=";
    private static final String TOKEN_KEY = "token=";
    private static final String FILE_HEADER =
            "# Direccion estable de la interfaz web local (se genera solo; no editar).";
    private static final String NL = "\n";
    private static final int MAX_PORT = 65535;

    private final int lastPort;
    private final String token;

    private WebSession(int lastPort, String token) {
        this.lastPort = lastPort;
        this.token = token;
    }

    /**
     * Carga la sesion del fichero por defecto del directorio de trabajo.
     *
     * @return la sesion; vacia si el fichero no existe o es invalido.
     */
    static WebSession load() {
        return load(Path.of(FILE_NAME));
    }

    /**
     * Carga la sesion desde una ruta concreta (usado por los tests).
     *
     * @param file fichero de sesion (puede no existir).
     * @return la sesion; vacia si el fichero no existe o es invalido.
     */
    static WebSession load(Path file) {
        int port = 0;
        String token = null;
        try {
            for (String raw : Files.readAllLines(file, StandardCharsets.UTF_8)) {
                String line = raw.trim();
                if (line.startsWith(PORT_KEY)) {
                    port = validPort(line.substring(PORT_KEY.length()).trim());
                } else if (line.startsWith(TOKEN_KEY)) {
                    String value = line.substring(TOKEN_KEY.length()).trim();
                    if (TOKEN_OK.matcher(value).matches()) {
                        token = value;
                    }
                }
            }
        } catch (IOException e) {
            // Sin fichero (o ilegible): sesion nueva; el arranque sigue.
            log.debug("Sin sesion web previa ({}): se generara una nueva.", e.getMessage());
        }
        return new WebSession(port, token);
    }

    /** Ultimo puerto enlazado; 0 si no hay sesion utilizable. */
    int lastPort() {
        return lastPort;
    }

    /** Token persistido; {@code null} si no hay sesion utilizable. */
    String token() {
        return token;
    }

    /**
     * URL de la interfaz guardada.
     *
     * @return {@code http://127.0.0.1:<puerto>/?token=...} o
     *         {@code null} si la sesion esta incompleta.
     */
    String savedUrl() {
        return hasAddress() ? url("http://127.0.0.1:" + lastPort + "/?token=" + token) : null;
    }

    /**
     * URL de comprobacion de vida ({@code GET /api/info}).
     *
     * @return la URL o {@code null} si la sesion esta incompleta.
     */
    String infoUrl() {
        return hasAddress()
                ? url("http://127.0.0.1:" + lastPort + "/api/info?token=" + token) : null;
    }

    private boolean hasAddress() {
        return lastPort > 0 && token != null;
    }

    private static String url(String value) {
        return value;
    }

    /**
     * Persiste la sesion en el fichero por defecto.
     *
     * @param port  puerto enlazado.
     * @param token token en uso.
     */
    static void save(int port, String token) {
        save(Path.of(FILE_NAME), port, token);
    }

    /**
     * Persiste la sesion en una ruta concreta (usado por los tests).
     * Best effort: si no se puede escribir, la proxima sesion sera nueva
     * pero el arranque actual no se ve afectado.
     *
     * @param file  fichero de sesion.
     * @param port  puerto enlazado.
     * @param token token en uso.
     */
    static void save(Path file, int port, String token) {
        if (port <= 0 || token == null || !TOKEN_OK.matcher(token).matches()) {
            log.debug("Sesion no utilizable (puerto {}): no se guarda.", port);
            return;
        }
        String content = FILE_HEADER + NL + "port=" + port + NL + "token=" + token + NL;
        try {
            if (Files.exists(file)
                    && content.equals(Files.readString(file, StandardCharsets.UTF_8))) {
                return;
            }
            Files.writeString(file, content, StandardCharsets.UTF_8);
        } catch (IOException e) {
            log.warn("No se pudo guardar '{}' ({}): proxima vez habra nueva direccion.",
                    file, e.getMessage());
        }
    }

    private static int validPort(String value) {
        try {
            int port = Integer.parseInt(value);
            return (port > 0 && port <= MAX_PORT) ? port : 0;
        } catch (NumberFormatException e) {
            return 0;
        }
    }
}

package com.excelmerger.web;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests de la sesion persistente de la interfaz web local: la direccion
 * (puerto + token) sobrevive entre arranques y cualquier fichero
 * corrupto degrada a sesion vacia sin tumbar el arranque.
 */
class WebSessionTest {

    @TempDir
    Path dir;

    @Test
    void sinFicheroLaSesionEstaVacia() {
        WebSession session = WebSession.load(dir.resolve("no-existe.properties"));

        assertThat(session.lastPort()).isZero();
        assertThat(session.token()).isNull();
        assertThat(session.savedUrl()).isNull();
        assertThat(session.infoUrl()).isNull();
    }

    @Test
    void saveYLoadConservanPuertoYToken() {
        Path file = dir.resolve(WebSession.FILE_NAME);

        WebSession.save(file, 7420, "abcdef-123456");
        WebSession loaded = WebSession.load(file);

        assertThat(loaded.lastPort()).isEqualTo(7420);
        assertThat(loaded.token()).isEqualTo("abcdef-123456");
        assertThat(loaded.savedUrl()).isEqualTo("http://127.0.0.1:7420/?token=abcdef-123456");
        assertThat(loaded.infoUrl())
                .isEqualTo("http://127.0.0.1:7420/api/info?token=abcdef-123456");
    }

    @Test
    void tokenInvalidoSeDescarta() throws Exception {
        Path file = dir.resolve(WebSession.FILE_NAME);
        Files.writeString(file, "port=7420\ntoken=\"; rm -rf /\n", StandardCharsets.UTF_8);

        WebSession session = WebSession.load(file);

        assertThat(session.token()).isNull();
        assertThat(session.lastPort()).isEqualTo(7420);
        assertThat(session.savedUrl()).isNull();
    }

    @Test
    void puertoFueraDeRangoSeIgnora() throws Exception {
        Path file = dir.resolve(WebSession.FILE_NAME);
        Files.writeString(file, "port=99999\ntoken=abcdefgh\n", StandardCharsets.UTF_8);

        WebSession session = WebSession.load(file);

        assertThat(session.lastPort()).isZero();
        assertThat(session.token()).isEqualTo("abcdefgh");
    }

    @Test
    void comentariosYParesParcialesSeToleran() throws Exception {
        Path file = dir.resolve(WebSession.FILE_NAME);
        Files.writeString(file,
                "# comentario\n\notra-clave=ignorada\ntoken=12345678\n",
                StandardCharsets.UTF_8);

        WebSession session = WebSession.load(file);

        assertThat(session.token()).isEqualTo("12345678");
        assertThat(session.lastPort()).isZero();
        assertThat(session.savedUrl()).isNull();
    }

    @Test
    void saveNoReescribeSiElContenidoEsElMismo() throws Exception {
        Path file = dir.resolve(WebSession.FILE_NAME);
        WebSession.save(file, 7420, "abcdef-123456");
        long firstWrite = Files.getLastModifiedTime(file).toMillis();

        Thread.sleep(25);
        WebSession.save(file, 7420, "abcdef-123456");

        assertThat(Files.getLastModifiedTime(file).toMillis()).isEqualTo(firstWrite);
    }

    @Test
    void saveRechazaValoresNoUtilizables() throws Exception {
        Path file = dir.resolve("nuevo.properties");

        WebSession.save(file, 0, "abcdef-123456");

        assertThat(Files.exists(file)).isFalse();
    }
}

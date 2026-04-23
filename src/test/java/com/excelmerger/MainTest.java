package com.excelmerger;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests del banner de versión y de la carga de {@code git.properties}.
 *
 * <p>Notas:</p>
 * <ul>
 *   <li>No se invoca {@link Main#main(String[])} porque usa {@code System.exit(...)} y
 *       destruiría la JVM del runner de tests.</li>
 *   <li>En entorno de tests el recurso {@code /git.properties} puede estar presente
 *       (si el build ha corrido el plugin) o ausente. El test debe verificar que
 *       {@link Main#buildInfoString()} NO rompe en ninguno de los dos casos y que el
 *       formato cumple el contrato de la Fase E/Sesión E.</li>
 * </ul>
 */
class MainTest {

    @Test
    void buildInfoStringNoLanzaCuandoFaltaGitProperties() {
        // Smoke test: la llamada debe completarse sin excepción, independientemente
        // de si /git.properties está presente en el classpath o no.
        String info = Main.buildInfoString();

        assertThat(info).isNotNull();
    }

    @Test
    void buildInfoStringIncluyeNombreYVersion() {
        String info = Main.buildInfoString();

        assertThat(info).startsWith("Excel Merger v" + Main.APP_VERSION);
    }

    @Test
    void buildInfoStringFormatoValido() {
        String info = Main.buildInfoString();

        // Dos formatos posibles y solo dos:
        //   a) "Excel Merger v1.5.0" (sin git.properties, o con claves vacías)
        //   b) "Excel Merger v1.5.0 (build <hash>, <fecha>)"
        String base = "Excel Merger v" + Main.APP_VERSION;
        boolean plano = info.equals(base);
        boolean enriquecido = info.matches("^" + java.util.regex.Pattern.quote(base)
                + " \\(build [0-9a-f]+, \\d{4}-\\d{2}-\\d{2}\\)$");

        assertThat(plano || enriquecido)
                .as("Formato inesperado: '%s'", info)
                .isTrue();
    }

    @Test
    void appVersionEsLaEsperadaPorLaSesionE() {
        // Este test fija la expectativa de versión al bump acordado.
        // Si se sube la versión, este test debe actualizarse conscientemente.
        // 1.5.2: patch de la Sesión E2 (segunda pasada de calidad + SpotBugs Medium).
        assertThat(Main.APP_VERSION).isEqualTo("1.5.2");
    }
}

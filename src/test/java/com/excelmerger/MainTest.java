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
        // 1.6.0: hoja "Resumen" (sumatorio por matrícula sobre Resultado).
        // 1.6.2: fix del SUMIFS de Jira que dejaba fuera imputaciones por
        //        mismatch numerico/textual entre Extraccion y Cierre.
        // 1.7.0: filas huerfanas en Resultado (imputaciones de Cierre sin
        //        contrapartida en Extraccion). Opt-in via mes.orphans.enabled.
        // 1.7.1: fix del bug C. SummarySheetBuilder escribe matriculas como
        //        STRING siempre; antes las todo-digito eran NUMERIC y el
        //        SUMIFS no casaba contra la columna Matricula de Resultado
        //        (que es STRING tras el fix 1.6.2), dando Jira=0 para todas
        //        las matriculas numericas en Resumen.
        // 1.8.0: segunda tabla opcional en Resumen — matriz cruzada
        //        Matricula x Res. Tecnico con la metrica PDCL (u otra
        //        configurable). Opt-in via summary.byResponsible.enabled.
        //        Normaliza responsables a MAYUSCULAS para colapsar
        //        variantes de capitalizacion del Excel original.
        // 1.8.1: fix del bug del padding de espacios en Usuario_Resp_Tecnico.
        //        El export ERP trae los codigos alineados con espacios
        //        ("MG002   "), que rompian el SUMIFS de la segunda tabla
        //        de Resumen (case-insensitive pero no trim-insensitive).
        //        Nueva clave de config profile.<id>.trim.columns aplica
        //        trim() en la capa de copia; todas las hojas aguas abajo
        //        consumen valores limpios. Fix via copyCellValueAsTextTrimmed.
        // 2.0.0: BREAKING CHANGE. Swap de nombres de perfil: el perfil que
        //        contenia las peticiones del ERP se llamaba "Extraccion"
        //        hasta 1.8.1; en 2.0.0 pasa a llamarse "Cierre". El perfil
        //        con el export de Jira se llamaba "Cierre"; pasa a llamarse
        //        "Extraccion". Los nombres ahora coinciden con los de los
        //        ficheros de entrada habituales. Deteccion sigue siendo
        //        por contenido. Ver CHANGELOG [2.0.0] para guia de migracion.
        // 2.0.1: segunda vuelta de limpieza PMD (sin cambios funcionales).
        // 2.1.0: columna Funcion en Resultado (tras Matricula, copia desde
        //        Cierre.Funcion). Cambio data-driven en config; sin tocar
        //        codigo de produccion.
        // 2.2.0: fichero opcional de Deuda. Nuevo tipo de columna
        //        FORMULA_PLUS_SUMIFS; retrocompat de input.strictTwoFiles
        //        via input.strictMinFiles/MaxFiles. Ver CHANGELOG [2.2.0].
        assertThat(Main.APP_VERSION).isEqualTo("2.2.0");
    }
}

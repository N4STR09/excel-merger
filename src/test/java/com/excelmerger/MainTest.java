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
        // 2.3.0: modos de generacion (output.mode=cierre|responsables|completo).
        //        Default cierre preserva 100% el comportamiento de v2.2.0.
        //        Nuevo enum OutputMode + ResponsablesSheetBuilder (N hojas
        //        vacias por responsable distinto en Resultado.Res. Tecnico,
        //        trim + case-insensitive, orden Collator es_ES). Switch de
        //        modos en ExcelMerger.merge: en RESPONSABLES omite Resumen
        //        y la copia del input Deuda. Ver CHANGELOG [2.3.0].
        // 2.4.0: tablas pivot Peticion x Matricula en cada hoja de
        //        responsable (modos responsables y completo). Dos tablas
        //        SUMIFS (Jira y Facturar) filtradas por A1, opt-out via
        //        responsables.tables.enabled=false. Nuevo helper
        //        ResponsablePivotBuilder. Ver CHANGELOG [2.4.0].
        // 2.5.0: refactor estructural de ConfigValidator (Sesion F.1).
        //        Cambio interno sin efecto en la salida, en la CLI ni en
        //        el config.properties. ConfigValidator monolitico (715
        //        LoC) -> orquestador delgado + 7 *ConfigSection en
        //        com.excelmerger.config. FQN cambia de
        //        com.excelmerger.ConfigValidator a
        //        com.excelmerger.config.ConfigValidator; firma intacta.
        //        Ver CHANGELOG [2.5.0].
        // 2.5.1: cobertura del paquete com.excelmerger.io + fix de
        //        deteccion de locks en Windows ES. Sin cambios funcionales
        //        salvo el caso concreto de Windows con locale espanol y un
        //        fichero abierto en Excel: antes se mostraba "No se puede
        //        leer..." (mensaje generico), ahora el mensaje correcto
        //        "Cierra '<fichero>.xlsx' antes de ejecutar".
        //        Anade FileLockDetectorTest (11 tests) y OutputManagerTest
        //        (20 tests). En produccion: un Supplier<LocalDateTime>
        //        package-private en OutputManager para testar la rama de
        //        colision de timestamp en backupOutput, y dos patrones en
        //        espanol ("otro proceso tiene bloqueada", "el proceso no
        //        tiene acceso") en FileLockDetector.looksLikeLocked.
        //        Ver CHANGELOG [2.5.1].
        // 2.6.0: rename cosmetico REAL -> Facturar en el nombre de la
        //        columna calculada (mes.col.11.name), propagado a todos los
        //        sitios (config.properties, summary.valueColumns,
        //        responsables.tables.realTitle -> facturarTitle, codigo,
        //        tests, README). Cambio de fuente de mes.col.14: ahora
        //        Horas_RealizadoTot copia desde Total_Horas_Realizadas_Recurso
        //        (Modif 1). Diagnostico de "Realizadas_Horas_Mes a 0" como
        //        realidad del ERP, no bug del programa (Modif 2). Sin alias
        //        retrocompatibles: configs antiguos requieren actualizacion
        //        manual (ver seccion Migracion en CHANGELOG [2.6.0]).
        // 2.7.0: tres modificaciones independientes:
        //        Modif 1 — freeze de la primera fila en el output, opt-out
        //        con output.freezeTopRow=false. Excluye Equipos (oculta) y
        //        las hojas de responsable (Fase 0 P1).
        //        Modif 2 — rename de mes.col.15: ahora Horas_Mes desde
        //        Cierre.UltimaPrevision_Horas_Mes (antes Realizadas_Horas_Mes
        //        desde Realizadas_Horas_Mes; el ERP real rellena la columna
        //        antigua con 0 en el 100% de las filas).
        //        Modif 3 — orden invertido de las pivots en hojas de
        //        responsable: ahora primera tabla = PDCL, segunda = Jira
        //        (antes: Jira primero, Facturar segunda). La fuente de la
        //        primera pivot pasa de la columna Facturar (mes.col.11) a
        //        PDCL (mes.col.12). Rename limpio de la clave config:
        //        responsables.tables.facturarTitle -> .pdclTitle. Configs
        //        con la clave antigua (o realTitle de ≤v2.5.1) son
        //        rechazados con error de migracion.
        //        Sin alias retrocompatibles: ver CHANGELOG [2.7.0] seccion
        //        Migracion.
        // 2.7.1: filtrado fisico de filas con las 5 columnas (Jira, Facturar,
        //        PDCL, PDCL + Deuda, Horas_Mes) evaluadas TODAS a 0. Nueva
        //        clave mes.removeEmptyRows=true (default). Si se pone a
        //        false, comportamiento identico al de v2.7.0. Implementacion
        //        post-build con FormulaEvaluator (regla inquebrantable 4).
        //        Efecto secundario positivo: Resumen y pivots de responsable
        //        no muestran filas con totales 0 porque sus claves descubren
        //        a partir de las filas FISICAS de Resultado, ya filtrado.
        // 3.0.0: BREAKING CHANGE. Eliminada la CLI argumentada (--help,
        //        --version, --dry-run, <configPath>). Al arrancar el JAR
        //        aparece SIEMPRE un menu interactivo (JLine) con tres
        //        opciones: 1) Fusion de Excel, 2) Otra opcion (placeholder),
        //        3) Salir. La logica de fusion se ha extraido a la clase
        //        com.excelmerger.App para testabilidad. El antiguo
        //        --dry-run se reemplaza por la clave de config
        //        output.dryRun (default false). Ver CHANGELOG [3.0.0]
        //        seccion Migracion.
        assertThat(Main.APP_VERSION).isEqualTo("3.0.0");
    }
}

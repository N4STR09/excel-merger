package com.excelmerger.config;

import com.excelmerger.ConfigLoader;
import com.excelmerger.TestFixtures;

import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Properties;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests directos de {@link SummaryConfigSection}. Bonus de la Sesión F (v2.5.0).
 *
 * <p>Esta sección tenía hasta v2.4.0 cobertura solo indirecta a través de
 * {@code ConfigValidatorTest} (no había {@code @Test} dirigido a
 * {@code summary.*}). El refactor es un buen momento para anclar el
 * comportamiento con tests específicos.</p>
 *
 * <p>Estilo: instanciar la sección directamente, sin pasar por
 * {@code ConfigValidator}, para verificar el contrato aislado.</p>
 */
class SummaryConfigSectionTest {

    private static SummaryConfigSection sectionFor(Properties p, List<String> errors) {
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        return new SummaryConfigSection(cfg, errors);
    }

    private static Set<String> sheets(String... names) {
        return new LinkedHashSet<>(java.util.Arrays.asList(names));
    }

    private static Properties baseSummaryEnabled() {
        Properties p = new Properties();
        p.setProperty("summary.enabled", "true");
        p.setProperty("summary.sheetName", "Resumen");
        p.setProperty("summary.sumSheet", "Resultado");
        p.setProperty("summary.matriculaColumn", "Matricula");
        p.setProperty("summary.valueColumns", "Horas,REAL");
        return p;
    }

    // ==================================================================
    //  summary.enabled=false
    // ==================================================================

    @Test
    void summaryDeshabilitadoNoAcumulaNada() {
        Properties p = new Properties();
        p.setProperty("summary.enabled", "false");
        // Aunque las claves del resumen estén malformadas, no se valida.
        p.setProperty("summary.sumSheet", "");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Resultado"));

        assertThat(errors).isEmpty();
    }

    // ==================================================================
    //  summary.enabled=true: claves obligatorias
    // ==================================================================

    @Test
    void summaryHabilitadoSinSumSheetEsError() {
        Properties p = baseSummaryEnabled();
        p.remove("summary.sumSheet");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Resultado"));

        assertThat(errors).anyMatch(e -> e.contains("summary.sumSheet"));
    }

    @Test
    void summaryConSumSheetDesconocidaEsError() {
        Properties p = baseSummaryEnabled();
        p.setProperty("summary.sumSheet", "HojaQueNoExiste");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Resultado"));

        assertThat(errors).anyMatch(e ->
                e.contains("summary.sumSheet")
                        && e.contains("HojaQueNoExiste")
                        && e.contains("desconocida"));
    }

    @Test
    void summarySheetNameColisionEsError() {
        Properties p = baseSummaryEnabled();
        p.setProperty("summary.sheetName", "Resultado"); // ya existe
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Resultado"));

        assertThat(errors).anyMatch(e ->
                e.contains("summary.sheetName") && e.contains("colisiona"));
    }

    @Test
    void summarySinValueColumnsEsError() {
        Properties p = baseSummaryEnabled();
        p.remove("summary.valueColumns");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Resultado"));

        assertThat(errors).anyMatch(e -> e.contains("summary.valueColumns"));
    }

    @Test
    void summaryCompletoYValidoNoLevantaErrores() {
        Properties p = baseSummaryEnabled();
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Resultado"));

        assertThat(errors).isEmpty();
    }

    // ==================================================================
    //  summary.byResponsible.*: opt-in
    // ==================================================================

    @Test
    void byResponsibleSinSummaryEnabledNoLevantaErrorPorEarlyReturn() {
        // ANCLAJE DE COMPORTAMIENTO HEREDADO (v2.4.0 → v2.5.0):
        //
        // Aunque el Javadoc de validateSummaryByResponsible dice "Si esta
        // habilitada sin que summary.enabled tambien lo este, se trata como
        // error", el chequeo es de hecho INALCANZABLE en este caso: cuando
        // summary.enabled=false, validateSummary() retorna en su primer
        // guard, sin llamar a validateSummaryByResponsible(). Por tanto el
        // error "byResponsible.enabled=true requiere summary.enabled=true"
        // solo se emite si summary.enabled cambia de true a false entre la
        // entrada de validateSummary y la llamada anidada — lo cual no
        // puede ocurrir.
        //
        // El refactor de la Sesion F.1 (v2.5.0) preserva este
        // comportamiento bit-a-bit. Si una sesion futura decide tratar
        // esto como bug y emitir el error, este test fallara y debera
        // actualizarse conscientemente.
        Properties p = new Properties();
        // summary.enabled NO está
        p.setProperty("summary.byResponsible.enabled", "true");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Resultado"));

        assertThat(errors).noneMatch(e ->
                e.contains("summary.byResponsible.enabled=true")
                        && e.contains("summary.enabled=true"));
    }

    @Test
    void byResponsibleSinClavesObligatoriasEsError() {
        Properties p = baseSummaryEnabled();
        p.setProperty("summary.byResponsible.enabled", "true");
        // Faltan column, valueColumn, title.
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Resultado"));

        assertThat(errors).anyMatch(e -> e.contains("summary.byResponsible.column"));
        assertThat(errors).anyMatch(e -> e.contains("summary.byResponsible.valueColumn"));
        assertThat(errors).anyMatch(e -> e.contains("summary.byResponsible.title"));
    }

    @Test
    void byResponsibleGapRowsAceptaCero() {
        Properties p = baseSummaryEnabled();
        p.setProperty("summary.byResponsible.enabled", "true");
        p.setProperty("summary.byResponsible.column", "Responsable");
        p.setProperty("summary.byResponsible.valueColumn", "Horas");
        p.setProperty("summary.byResponsible.title", "Por responsable");
        p.setProperty("summary.byResponsible.gapRows", "0");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Resultado"));

        assertThat(errors).noneMatch(e -> e.contains("byResponsible.gapRows"));
    }

    @Test
    void byResponsibleGapRowsNegativoEsError() {
        Properties p = baseSummaryEnabled();
        p.setProperty("summary.byResponsible.enabled", "true");
        p.setProperty("summary.byResponsible.column", "Responsable");
        p.setProperty("summary.byResponsible.valueColumn", "Horas");
        p.setProperty("summary.byResponsible.title", "Por responsable");
        p.setProperty("summary.byResponsible.gapRows", "-3");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Resultado"));

        assertThat(errors).anyMatch(e ->
                e.contains("summary.byResponsible.gapRows") && e.contains(">= 0"));
    }

    @Test
    void byResponsibleGapRowsNoNumericoEsError() {
        Properties p = baseSummaryEnabled();
        p.setProperty("summary.byResponsible.enabled", "true");
        p.setProperty("summary.byResponsible.column", "Responsable");
        p.setProperty("summary.byResponsible.valueColumn", "Horas");
        p.setProperty("summary.byResponsible.title", "Por responsable");
        p.setProperty("summary.byResponsible.gapRows", "tres");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Resultado"));

        assertThat(errors).anyMatch(e ->
                e.contains("summary.byResponsible.gapRows")
                        && e.contains("no numerico"));
    }
}

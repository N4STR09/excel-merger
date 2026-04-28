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
 * Tests directos de {@link MesConfigSection}. Bonus de la Sesión F (v2.5.0).
 *
 * <p>Estilo: instanciar la sección directamente, sin pasar por
 * {@code ConfigValidator}, para verificar el contrato aislado. Un test
 * representativo por cada uno de los cinco tipos de columna MES, más los
 * casos transversales (mes deshabilitado, sin columnas, sourceSheet
 * desconocida).</p>
 *
 * <p>El comportamiento "end to end" sigue cubierto por
 * {@code ConfigValidatorTest}; aquí solo nos aseguramos de que la sección
 * se comporta igual cuando se invoca aislada.</p>
 */
class MesConfigSectionTest {

    private static MesConfigSection sectionFor(Properties p, List<String> errors) {
        ConfigLoader cfg = TestFixtures.configFromProperties(p);
        return new MesConfigSection(cfg, errors);
    }

    private static Set<String> sheets(String... names) {
        return new LinkedHashSet<>(java.util.Arrays.asList(names));
    }

    private static Properties baseEnabledMes() {
        Properties p = new Properties();
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sourceSheet", "Origen");
        p.setProperty("mes.anchorColumn", "Peticion");
        return p;
    }

    // ==================================================================
    //  Trivial: deshabilitado / sin columnas
    // ==================================================================

    @Test
    void mesDeshabilitadoNoAcumulaNada() {
        Properties p = new Properties();
        p.setProperty("mes.enabled", "false");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Origen"));

        assertThat(errors).isEmpty();
    }

    @Test
    void mesHabilitadoSinColumnasEsError() {
        Properties p = baseEnabledMes();
        // Ninguna mes.col.N.name definida.
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Origen"));

        assertThat(errors).anyMatch(e ->
                e.contains("mes.enabled=true") && e.contains("ninguna columna"));
    }

    @Test
    void sourceSheetDesconocidaEsError() {
        Properties p = baseEnabledMes();
        p.setProperty("mes.col.1.name", "X");
        p.setProperty("mes.col.1.type", "EMPTY");
        List<String> errors = new ArrayList<>();

        // Origen no está en knownSheets.
        sectionFor(p, errors).validate(sheets("OtraHoja"));

        assertThat(errors).anyMatch(e ->
                e.contains("mes.sourceSheet") && e.contains("Origen"));
    }

    // ==================================================================
    //  Tipo COPY: requiere from
    // ==================================================================

    @Test
    void copySinFromEsError() {
        Properties p = baseEnabledMes();
        p.setProperty("mes.col.1.name", "ColA");
        p.setProperty("mes.col.1.type", "COPY");
        // Falta mes.col.1.from
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Origen"));

        assertThat(errors).anyMatch(e ->
                e.contains("mes.col.1.from") && e.contains("COPY"));
    }

    // ==================================================================
    //  Tipo SUMIFS: requiere from (hoja conocida), sum, match con pares
    // ==================================================================

    @Test
    void sumifsConFromDesconocidoEsError() {
        Properties p = baseEnabledMes();
        p.setProperty("mes.col.1.name", "Horas");
        p.setProperty("mes.col.1.type", "SUMIFS");
        p.setProperty("mes.col.1.from", "HojaQueNoExiste");
        p.setProperty("mes.col.1.sum", "Horas");
        p.setProperty("mes.col.1.match", "A:B");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Origen"));

        assertThat(errors).anyMatch(e ->
                e.contains("mes.col.1.from")
                        && e.contains("HojaQueNoExiste")
                        && e.contains("desconocida"));
    }

    @Test
    void sumifsConMatchSinParesEsError() {
        Properties p = baseEnabledMes();
        p.setProperty("mes.col.1.name", "Horas");
        p.setProperty("mes.col.1.type", "SUMIFS");
        p.setProperty("mes.col.1.from", "Origen");
        p.setProperty("mes.col.1.sum", "Horas");
        p.setProperty("mes.col.1.match", "sin_dos_puntos, ni_aqui");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Origen"));

        assertThat(errors).anyMatch(e ->
                e.contains("mes.col.1.match")
                        && e.contains("malformada"));
    }

    // ==================================================================
    //  Tipo FORMULA: placeholders {col:X} validados
    // ==================================================================

    @Test
    void formulaConPlaceholderQueNoExisteEsError() {
        Properties p = baseEnabledMes();
        p.setProperty("mes.col.1.name", "A");
        p.setProperty("mes.col.1.type", "EMPTY");
        p.setProperty("mes.col.2.name", "B");
        p.setProperty("mes.col.2.type", "FORMULA");
        p.setProperty("mes.col.2.formula", "={col:A}+{col:NoExiste}");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Origen"));

        // {col:A} es válido (existe), {col:NoExiste} no -> 1 error
        assertThat(errors).anyMatch(e ->
                e.contains("placeholder") && e.contains("NoExiste"));
        assertThat(errors).noneMatch(e -> e.contains("placeholder {col:A}"));
    }

    // ==================================================================
    //  Tipo FORMULA_PLUS_SUMIFS: combinación
    // ==================================================================

    @Test
    void formulaPlusSumifsCompletoNoLevantaErrores() {
        Properties p = baseEnabledMes();
        p.setProperty("mes.col.1.name", "Base");
        p.setProperty("mes.col.1.type", "EMPTY");
        p.setProperty("mes.col.2.name", "Calculo");
        p.setProperty("mes.col.2.type", "FORMULA_PLUS_SUMIFS");
        p.setProperty("mes.col.2.baseFormula", "={col:Base}*2");
        p.setProperty("mes.col.2.from", "Origen");
        p.setProperty("mes.col.2.sum", "Horas");
        p.setProperty("mes.col.2.match", "Petic:Peticion");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Origen"));

        assertThat(errors).isEmpty();
    }

    // ==================================================================
    //  Tipo EMPTY + transversales (fill, redIfNotEqualTo)
    // ==================================================================

    @Test
    void emptyConFillInvalidoEsError() {
        Properties p = baseEnabledMes();
        p.setProperty("mes.col.1.name", "X");
        p.setProperty("mes.col.1.type", "EMPTY");
        p.setProperty("mes.col.1.fill", "FUCSIA_BRILLANTE");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Origen"));

        assertThat(errors).anyMatch(e ->
                e.contains("mes.col.1.fill") && e.contains("FUCSIA_BRILLANTE"));
    }

    @Test
    void redIfNotEqualToApuntandoASiMismaEsError() {
        Properties p = baseEnabledMes();
        p.setProperty("mes.col.1.name", "X");
        p.setProperty("mes.col.1.type", "EMPTY");
        p.setProperty("mes.col.1.redIfNotEqualTo", "X");
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Origen"));

        assertThat(errors).anyMatch(e ->
                e.contains("redIfNotEqualTo") && e.contains("si misma"));
    }

    // ==================================================================
    //  Tipo inválido: salta el resto de validaciones de esa columna
    // ==================================================================

    @Test
    void tipoInvalidoEsErrorYNoCascada() {
        Properties p = baseEnabledMes();
        p.setProperty("mes.col.1.name", "X");
        p.setProperty("mes.col.1.type", "RAYO_LASER");
        // Como el tipo es inválido, el switch sale; no se queja de
        // que falte ".from" ni similares.
        List<String> errors = new ArrayList<>();

        sectionFor(p, errors).validate(sheets("Origen"));

        assertThat(errors).anyMatch(e ->
                e.contains("mes.col.1.type") && e.contains("RAYO_LASER"));
        // No debe quejarse de from porque el tipo ya falló
        assertThat(errors).noneMatch(e ->
                e.contains("mes.col.1.from") && e.contains("requerido"));
    }
}

package com.excelmerger;

import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Properties;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests de {@link ConfigValidator}. Puntos cubiertos:
 * <ul>
 *   <li>Config minimo valido no acumula errores.</li>
 *   <li>Input/output ausentes -> errores reportados.</li>
 *   <li>Modo de fusion invalido.</li>
 *   <li>Perfiles sin criterios de deteccion.</li>
 *   <li>Numeros negativos/invalidos en headerRow, sheetIndex, minMatches.</li>
 *   <li>MES: tipo invalido, campos requeridos por tipo, placeholders que
 *       no casan con ninguna columna definida, referencias a hojas
 *       desconocidas.</li>
 *   <li>Lookups sin data.</li>
 *   <li>Derivadas: tipo invalido, AGGREGATION con sourceSheet desconocido,
 *       letras de columna invalidas, aggregation invalido.</li>
 *   <li>Acumulacion: multiples errores se devuelven todos en una sola
 *       invocacion de validate().</li>
 *   <li>validate() es idempotente: llamar varias veces no duplica errores.</li>
 * </ul>
 */
class ConfigValidatorTest {

    // ==================================================================
    //  Helpers de construccion de config
    // ==================================================================
    private static Properties minimalValid() {
        Properties p = new Properties();
        p.setProperty("input.directory", "/tmp/in");
        p.setProperty("output.file", "/tmp/out.xlsx");
        p.setProperty("merge.mode", "SHEETS_SEPARATE");
        return p;
    }

    private static ConfigValidator validatorFor(Properties p) {
        return new ConfigValidator(TestFixtures.configFromProperties(p));
    }

    // ==================================================================
    //  Input / output
    // ==================================================================
    @Test
    void configMinimoValidoNoDevuelveErrores() {
        ConfigValidator v = validatorFor(minimalValid());

        assertThat(v.validate()).isEmpty();
    }

    @Test
    void inputDirectoryAusenteEsError() {
        Properties p = minimalValid();
        p.remove("input.directory");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("input.directory"));
    }

    @Test
    void outputFileAusenteEsError() {
        Properties p = minimalValid();
        p.remove("output.file");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("output.file"));
    }

    @Test
    void mergeModeInvalidoEsError() {
        Properties p = minimalValid();
        p.setProperty("merge.mode", "FUNKY_MODE");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("merge.mode") && e.contains("FUNKY_MODE"));
    }

    // ==================================================================
    //  Perfiles
    // ==================================================================
    @Test
    void perfilSinHeadersNiCellValueEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "MiPerfil");
        p.setProperty("profile.MiPerfil.sheetName", "Datos");
        // Intencionalmente: sin detect.headers y sin detect.cellValue.*

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("'MiPerfil'") && e.contains("incompleto"));
    }

    @Test
    void perfilConCellValueSatisfaceLaDeteccion() {
        Properties p = minimalValid();
        p.setProperty("profiles", "MiPerfil");
        p.setProperty("profile.MiPerfil.sheetName", "Datos");
        p.setProperty("profile.MiPerfil.detect.cellValue.A1", "TITULO");
        // Sin detect.headers, pero con cellValue -> valido

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).noneMatch(e -> e.contains("'MiPerfil'"));
    }

    @Test
    void headerRowNegativoEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "X");
        p.setProperty("profile.X.detect.headers", "a,b,c");
        p.setProperty("profile.X.detect.headerRow", "-1");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("headerRow") && e.contains("-1"));
    }

    @Test
    void minMatchesMayorQueCabecerasEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "X");
        p.setProperty("profile.X.detect.headers", "a,b");
        p.setProperty("profile.X.detect.minMatches", "5");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("minMatches")
                && e.contains("5") && e.contains("2"));
    }

    @Test
    void valorNoNumericoEnHeaderRowEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "X");
        p.setProperty("profile.X.detect.headers", "a,b");
        p.setProperty("profile.X.detect.headerRow", "not-a-number");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("headerRow") && e.contains("no numerico"));
    }

    // ==================================================================
    //  Lookups
    // ==================================================================
    @Test
    void lookupSinDataEsError() {
        Properties p = minimalValid();
        p.setProperty("lookup.sheets", "Equipos");
        // lookup.Equipos.data ausente

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("Lookup 'Equipos'") && e.contains("sin datos"));
    }

    // ==================================================================
    //  MES
    // ==================================================================
    @Test
    void mesEnabledSinSourceSheetEsError() {
        Properties p = minimalValid();
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.anchorColumn", "X");
        p.setProperty("mes.col.1.name", "col1");
        p.setProperty("mes.col.1.type", "EMPTY");
        // sin mes.sourceSheet

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("mes.sourceSheet"));
    }

    @Test
    void mesSourceSheetDesconocidaEsError() {
        Properties p = minimalValid();
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sourceSheet", "NoExiste");
        p.setProperty("mes.anchorColumn", "X");
        p.setProperty("mes.col.1.name", "col1");
        p.setProperty("mes.col.1.type", "EMPTY");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("mes.sourceSheet")
                && e.contains("NoExiste"));
    }

    @Test
    void mesEnabledSinColumnasEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "Extraccion");
        p.setProperty("profile.Extraccion.sheetName", "Extraccion");
        p.setProperty("profile.Extraccion.detect.headers", "a,b,c");
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sourceSheet", "Extraccion");
        p.setProperty("mes.anchorColumn", "Peticion");
        // NINGUNA mes.col.N.name definida

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("no se ha definido ninguna columna"));
    }

    @Test
    void mesColumnaTipoInvalidoEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "Extraccion");
        p.setProperty("profile.Extraccion.sheetName", "Extraccion");
        p.setProperty("profile.Extraccion.detect.headers", "a,b,c");
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sourceSheet", "Extraccion");
        p.setProperty("mes.anchorColumn", "Peticion");
        p.setProperty("mes.col.1.name", "col1");
        p.setProperty("mes.col.1.type", "INVENTADO");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("mes.col.1.type")
                && e.contains("INVENTADO"));
    }

    @Test
    void mesColumnaNombreDuplicadoEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "Extraccion");
        p.setProperty("profile.Extraccion.sheetName", "Extraccion");
        p.setProperty("profile.Extraccion.detect.headers", "a,b,c");
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sourceSheet", "Extraccion");
        p.setProperty("mes.anchorColumn", "Peticion");
        p.setProperty("mes.col.1.name", "dup");
        p.setProperty("mes.col.1.type", "EMPTY");
        p.setProperty("mes.col.2.name", "dup");
        p.setProperty("mes.col.2.type", "EMPTY");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("nombre duplicado")
                && e.contains("'dup'"));
    }

    @Test
    void sumIfsReferenciaHojaDesconocidaEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "Extraccion");
        p.setProperty("profile.Extraccion.sheetName", "Extraccion");
        p.setProperty("profile.Extraccion.detect.headers", "a,b,c");
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sourceSheet", "Extraccion");
        p.setProperty("mes.anchorColumn", "Peticion");
        p.setProperty("mes.col.1.name", "Jira");
        p.setProperty("mes.col.1.type", "SUMIFS");
        p.setProperty("mes.col.1.from", "HojaInexistente");
        p.setProperty("mes.col.1.sum", "Hours");
        p.setProperty("mes.col.1.match", "X:Y");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("mes.col.1.from")
                && e.contains("HojaInexistente"));
    }

    @Test
    void sumIfsSinMatchEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "Extraccion,Cierre");
        p.setProperty("profile.Extraccion.sheetName", "Extraccion");
        p.setProperty("profile.Extraccion.detect.headers", "a,b,c");
        p.setProperty("profile.Cierre.sheetName", "Cierre");
        p.setProperty("profile.Cierre.detect.headers", "x,y,z");
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sourceSheet", "Extraccion");
        p.setProperty("mes.anchorColumn", "Peticion");
        p.setProperty("mes.col.1.name", "Jira");
        p.setProperty("mes.col.1.type", "SUMIFS");
        p.setProperty("mes.col.1.from", "Cierre");
        p.setProperty("mes.col.1.sum", "Hours");
        // sin match

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("mes.col.1.match")
                && e.contains("requerido"));
    }

    @Test
    void sumIfsMatchMalformadoEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "Extraccion,Cierre");
        p.setProperty("profile.Extraccion.sheetName", "Extraccion");
        p.setProperty("profile.Extraccion.detect.headers", "a,b,c");
        p.setProperty("profile.Cierre.sheetName", "Cierre");
        p.setProperty("profile.Cierre.detect.headers", "x,y,z");
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sourceSheet", "Extraccion");
        p.setProperty("mes.anchorColumn", "Peticion");
        p.setProperty("mes.col.1.name", "Jira");
        p.setProperty("mes.col.1.type", "SUMIFS");
        p.setProperty("mes.col.1.from", "Cierre");
        p.setProperty("mes.col.1.sum", "Hours");
        p.setProperty("mes.col.1.match", "loquesea-sin-dos-puntos");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("mes.col.1.match"));
    }

    @Test
    void formulaPlaceholderInvalidoEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "Extraccion");
        p.setProperty("profile.Extraccion.sheetName", "Extraccion");
        p.setProperty("profile.Extraccion.detect.headers", "a,b,c");
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sourceSheet", "Extraccion");
        p.setProperty("mes.anchorColumn", "Peticion");
        p.setProperty("mes.col.1.name", "Jira");
        p.setProperty("mes.col.1.type", "EMPTY");
        p.setProperty("mes.col.2.name", "Desfase");
        p.setProperty("mes.col.2.type", "FORMULA");
        p.setProperty("mes.col.2.formula", "{col:NoExiste}-{col:Jira}");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("{col:NoExiste}"));
    }

    @Test
    void formulaSinFormulaEsError() {
        Properties p = minimalValid();
        p.setProperty("profiles", "Extraccion");
        p.setProperty("profile.Extraccion.sheetName", "Extraccion");
        p.setProperty("profile.Extraccion.detect.headers", "a,b,c");
        p.setProperty("mes.enabled", "true");
        p.setProperty("mes.sourceSheet", "Extraccion");
        p.setProperty("mes.anchorColumn", "Peticion");
        p.setProperty("mes.col.1.name", "X");
        p.setProperty("mes.col.1.type", "FORMULA");
        // sin formula

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("mes.col.1.formula")
                && e.contains("requerido"));
    }

    // ==================================================================
    //  Derivadas
    // ==================================================================
    @Test
    void derivadaTipoInvalidoEsError() {
        Properties p = minimalValid();
        p.setProperty("derived.sheets", "mi_hoja");
        p.setProperty("sheet.mi_hoja.type", "CUALQUIERA");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("sheet.mi_hoja.type")
                && e.contains("CUALQUIERA"));
    }

    @Test
    void aggregationConSourceSheetDesconocidoEsError() {
        Properties p = minimalValid();
        p.setProperty("derived.sheets", "resumen");
        p.setProperty("sheet.resumen.type", "AGGREGATION");
        p.setProperty("sheet.resumen.sourceSheet", "NoExiste");
        p.setProperty("sheet.resumen.groupByColumn", "A");
        p.setProperty("sheet.resumen.valueColumn", "B");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("sheet.resumen.sourceSheet")
                && e.contains("NoExiste"));
    }

    @Test
    void aggregationColumnLetterInvalidaEsError() {
        Properties p = minimalValid();
        p.setProperty("derived.sheets", "resumen");
        p.setProperty("sheet.resumen.type", "AGGREGATION");
        p.setProperty("sheet.resumen.sourceSheet", "resumen");
        p.setProperty("sheet.resumen.groupByColumn", "12"); // numero, no letra
        p.setProperty("sheet.resumen.valueColumn", "B");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("groupByColumn")
                && e.contains("letra"));
    }

    @Test
    void aggregationFuncionInvalidaEsError() {
        Properties p = minimalValid();
        p.setProperty("derived.sheets", "resumen");
        p.setProperty("sheet.resumen.type", "AGGREGATION");
        p.setProperty("sheet.resumen.sourceSheet", "resumen");
        p.setProperty("sheet.resumen.groupByColumn", "A");
        p.setProperty("sheet.resumen.valueColumn", "B");
        p.setProperty("sheet.resumen.aggregation", "MEDIAN"); // no soportada

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).anyMatch(e -> e.contains("aggregation")
                && e.contains("MEDIAN"));
    }

    @Test
    void aggregationCompletaYValidaNoLevantaErrores() {
        Properties p = minimalValid();
        p.setProperty("profiles", "Extraccion");
        p.setProperty("profile.Extraccion.sheetName", "Extraccion");
        p.setProperty("profile.Extraccion.detect.headers", "a,b,c");
        p.setProperty("derived.sheets", "resumen");
        p.setProperty("sheet.resumen.type", "AGGREGATION");
        p.setProperty("sheet.resumen.sourceSheet", "Extraccion"); // conocida
        p.setProperty("sheet.resumen.groupByColumn", "A");
        p.setProperty("sheet.resumen.valueColumn", "B");
        p.setProperty("sheet.resumen.aggregation", "SUM");

        List<String> errors = validatorFor(p).validate();

        assertThat(errors).isEmpty();
    }

    // ==================================================================
    //  Acumulacion de errores
    // ==================================================================
    @Test
    void acumulaVariosErroresEnUnaSolaInvocacion() {
        Properties p = new Properties();
        // Sin input.directory (1)
        // Sin output.file (2)
        p.setProperty("merge.mode", "FOO");                              // (3)
        p.setProperty("profiles", "X");
        // Perfil X sin headers ni cellValue                              // (4)
        p.setProperty("lookup.sheets", "LL");
        // Lookup LL sin data                                             // (5)

        List<String> errors = validatorFor(p).validate();

        assertThat(errors.size()).isGreaterThanOrEqualTo(5);
        assertThat(errors).anyMatch(e -> e.contains("input.directory"));
        assertThat(errors).anyMatch(e -> e.contains("output.file"));
        assertThat(errors).anyMatch(e -> e.contains("merge.mode"));
        assertThat(errors).anyMatch(e -> e.contains("'X'"));
        assertThat(errors).anyMatch(e -> e.contains("Lookup 'LL'"));
    }

    @Test
    void validateEsIdempotente() {
        Properties p = new Properties();
        // Deliberadamente vacio para provocar varios errores.
        ConfigValidator v = validatorFor(p);

        List<String> first  = v.validate();
        List<String> second = v.validate();

        assertThat(second).hasSameSizeAs(first).containsExactlyElementsOf(first);
    }

    @Test
    void validateDevuelveListaInmodificable() {
        List<String> errors = validatorFor(new Properties()).validate();

        org.assertj.core.api.Assertions.assertThatThrownBy(() -> errors.add("boom"))
                .isInstanceOf(UnsupportedOperationException.class);
    }
}

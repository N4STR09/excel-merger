package com.excelmerger.compare;

import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests unitarios de {@link DiscrepancyComparator}. v3.1.0.
 *
 * <p>Cubre los 3 tipos de discrepancia (DIFERENCIA, SOLO_CSV,
 * SOLO_RESULTADO), el orden de salida, la tolerancia cero y la
 * agregacion de imputaciones CSV duplicadas por la misma clave.</p>
 */
class DiscrepancyComparatorTest {

    private final DiscrepancyComparator comparator = new DiscrepancyComparator();

    private static Map<DiscrepancyKey, Double> resultadoMap(Object... pairs) {
        // pairs: clave, valor, clave, valor, ...
        Map<DiscrepancyKey, Double> map = new LinkedHashMap<>();
        for (int i = 0; i < pairs.length; i += 2) {
            map.put((DiscrepancyKey) pairs[i], (Double) pairs[i + 1]);
        }
        return map;
    }

    @Test
    void detectaDiferenciaCuandoLosValoresDifieren() {
        List<CsvImputacion> csv = Collections.singletonList(
                new CsvImputacion("100001", "99001", "RE", 10.0));
        Map<DiscrepancyKey, Double> res = resultadoMap(
                new DiscrepancyKey("99001", "100001", "RE"), 8.0);

        List<Discrepancy> result = comparator.compare("99001", csv, res);

        assertThat(result).hasSize(1);
        Discrepancy d = result.get(0);
        assertThat(d.getTipo()).isEqualTo(Discrepancy.Type.DIFERENCIA);
        assertThat(d.getRealizadoHoras()).isEqualTo(10.0);
        assertThat(d.getPdclDeuda()).isEqualTo(8.0);
        assertThat(d.getDiferencia()).isEqualTo(2.0);
    }

    @Test
    void noHayDiscrepanciaSiLosValoresCoinciden() {
        List<CsvImputacion> csv = Collections.singletonList(
                new CsvImputacion("100001", "99001", "RE", 5.0));
        Map<DiscrepancyKey, Double> res = resultadoMap(
                new DiscrepancyKey("99001", "100001", "RE"), 5.0);

        assertThat(comparator.compare("99001", csv, res)).isEmpty();
    }

    @Test
    void detectaSoloCsvCuandoLaClaveNoExisteEnResultado() {
        List<CsvImputacion> csv = Collections.singletonList(
                new CsvImputacion("100003", "99001", "OT", 3.5));
        Map<DiscrepancyKey, Double> res = resultadoMap();

        List<Discrepancy> result = comparator.compare("99001", csv, res);

        assertThat(result).hasSize(1);
        Discrepancy d = result.get(0);
        assertThat(d.getTipo()).isEqualTo(Discrepancy.Type.SOLO_CSV);
        assertThat(d.getRealizadoHoras()).isEqualTo(3.5);
        assertThat(d.getPdclDeuda()).isNaN();
        assertThat(d.getDiferencia()).isNaN();
    }

    @Test
    void detectaSoloResultadoSiClaveDelOrigenNoEstaEnElCsv() {
        List<CsvImputacion> csv = Collections.emptyList();
        Map<DiscrepancyKey, Double> res = resultadoMap(
                new DiscrepancyKey("99001", "100006", "OT"), 7.0);

        List<Discrepancy> result = comparator.compare("99001", csv, res);

        assertThat(result).hasSize(1);
        Discrepancy d = result.get(0);
        assertThat(d.getTipo()).isEqualTo(Discrepancy.Type.SOLO_RESULTADO);
        assertThat(d.getRealizadoHoras()).isNaN();
        assertThat(d.getPdclDeuda()).isEqualTo(7.0);
        assertThat(d.getDiferencia()).isNaN();
    }

    @Test
    void clavesDeOtraMatriculaNoSeReportanComoSoloResultado() {
        // El comparador procesa el CSV de matricula 99001. Si Resultado
        // contiene una clave para 99002, no debe aparecer como
        // SOLO_RESULTADO en este origen (le tocaria al CSV de 99002).
        List<CsvImputacion> csv = Collections.emptyList();
        Map<DiscrepancyKey, Double> res = resultadoMap(
                new DiscrepancyKey("99002", "200001", "RE"), 7.0);

        List<Discrepancy> result = comparator.compare("99001", csv, res);

        assertThat(result).isEmpty();
    }

    @Test
    void elOrdenDeSalidaEsDiferenciasLuegoSoloCsvLuegoSoloResultado() {
        List<CsvImputacion> csv = Arrays.asList(
                new CsvImputacion("100001", "99001", "RE", 5.0),  // match -> nada
                new CsvImputacion("100002", "99001", "AN", 10.0), // DIFERENCIA
                new CsvImputacion("100003", "99001", "OT", 3.5)); // SOLO_CSV
        Map<DiscrepancyKey, Double> res = resultadoMap(
                new DiscrepancyKey("99001", "100001", "RE"), 5.0,
                new DiscrepancyKey("99001", "100002", "AN"), 8.0,
                new DiscrepancyKey("99001", "100006", "OT"), 7.0,  // SOLO_RESULTADO
                new DiscrepancyKey("99001", "100007", "RE"), 2.0); // SOLO_RESULTADO

        List<Discrepancy> result = comparator.compare("99001", csv, res);

        assertThat(result).hasSize(4);
        assertThat(result.get(0).getTipo()).isEqualTo(Discrepancy.Type.DIFERENCIA);
        assertThat(result.get(0).getPeticion()).isEqualTo("100002");
        assertThat(result.get(1).getTipo()).isEqualTo(Discrepancy.Type.SOLO_CSV);
        assertThat(result.get(1).getPeticion()).isEqualTo("100003");
        assertThat(result.get(2).getTipo()).isEqualTo(Discrepancy.Type.SOLO_RESULTADO);
        assertThat(result.get(2).getPeticion()).isEqualTo("100006");
        assertThat(result.get(3).getTipo()).isEqualTo(Discrepancy.Type.SOLO_RESULTADO);
        assertThat(result.get(3).getPeticion()).isEqualTo("100007");
    }

    @Test
    void agregaSumandoImputacionesCsvDuplicadasPorClave() {
        // Dos lineas en el CSV con la misma clave: deben sumarse antes
        // de comparar contra Resultado.
        List<CsvImputacion> csv = Arrays.asList(
                new CsvImputacion("100001", "99001", "RE", 3.0),
                new CsvImputacion("100001", "99001", "RE", 2.0));
        Map<DiscrepancyKey, Double> res = resultadoMap(
                new DiscrepancyKey("99001", "100001", "RE"), 5.0);

        // Suma 3+2 = 5, coincide con Resultado: NO hay discrepancia.
        assertThat(comparator.compare("99001", csv, res)).isEmpty();
    }

    @Test
    void toleranciaCeroDetectaDiferenciaDe001() {
        List<CsvImputacion> csv = Collections.singletonList(
                new CsvImputacion("100001", "99001", "RE", 5.0));
        Map<DiscrepancyKey, Double> res = resultadoMap(
                new DiscrepancyKey("99001", "100001", "RE"), 5.01);

        List<Discrepancy> result = comparator.compare("99001", csv, res);

        assertThat(result).hasSize(1);
        assertThat(result.get(0).getTipo()).isEqualTo(Discrepancy.Type.DIFERENCIA);
    }

    @Test
    void csvVacioYResultadoVacioNoProducenDiscrepancias() {
        List<Discrepancy> result = comparator.compare(
                "99001", Collections.emptyList(), resultadoMap());
        assertThat(result).isEmpty();
    }

    @Test
    void elOrigenSeAdjuntaACadaDiscrepancia() {
        List<CsvImputacion> csv = Collections.singletonList(
                new CsvImputacion("100003", "99001", "OT", 3.5));
        List<Discrepancy> result = comparator.compare("90014", csv, resultadoMap());
        assertThat(result).allSatisfy(d -> assertThat(d.getOrigen()).isEqualTo("90014"));
    }
}

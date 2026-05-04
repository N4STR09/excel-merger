package com.excelmerger.compare;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Cruza una lista de imputaciones del CSV con el mapa del Resultado y
 * devuelve la lista de discrepancias detectadas.
 *
 * <p>Tolerancia de comparacion: <b>cero</b>. Cualquier diferencia, aunque
 * sea de 0.01 horas, es discrepancia.</p>
 *
 * <h2>Algoritmo</h2>
 * <ol>
 *   <li>Para cada imputacion del CSV:
 *     <ul>
 *       <li>Busca la clave (matricula, peticion, funcion) en el mapa de
 *           Resultado.</li>
 *       <li>Si existe y los valores coinciden -&gt; no hay discrepancia.</li>
 *       <li>Si existe pero los valores difieren -&gt; {@code DIFERENCIA}.</li>
 *       <li>Si no existe -&gt; {@code SOLO_CSV}.</li>
 *     </ul>
 *   </li>
 *   <li>Tras procesar todas las imputaciones del CSV, las claves del
 *       mapa de Resultado correspondientes a la matricula del CSV que NO
 *       fueron consumidas se reportan como {@code SOLO_RESULTADO}. La
 *       matricula del CSV se infiere del nombre del fichero
 *       ({@link CsvParser#originFromFileName(String)}). Una matricula =
 *       un CSV.</li>
 * </ol>
 *
 * <h2>Imputaciones duplicadas en el CSV</h2>
 * <p>Si el CSV tiene dos lineas con la misma clave (poco esperable pero
 * posible), se agregan sumando {@code Realizado Horas}. Mismo enfoque que
 * en {@link ResultadoReader}.</p>
 */
public final class DiscrepancyComparator {

    private static final Logger log = LoggerFactory.getLogger(DiscrepancyComparator.class);

    /**
     * Compara las imputaciones de un CSV (perteneciente a una sola
     * matricula, indicada por {@code origen}) contra el mapa global del
     * Resultado.
     *
     * @param origen          identificador de origen, normalmente la
     *                        matricula del CSV (nombre del fichero sin
     *                        extension).
     * @param csvImputaciones imputaciones leidas del CSV de esa matricula.
     * @param resultadoMap    mapa global (Matricula, Peticion, Funcion) -&gt;
     *                        PDCL+Deuda del Resultado completo.
     * @return lista de discrepancias detectadas, en este orden:
     *         primero todas las DIFERENCIA, luego todas las SOLO_CSV
     *         (en el orden del CSV), luego todas las SOLO_RESULTADO (en
     *         el orden del mapa de Resultado).
     */
    public List<Discrepancy> compare(String origen,
                                     List<CsvImputacion> csvImputaciones,
                                     Map<DiscrepancyKey, Double> resultadoMap) {
        // 1. Agregar imputaciones del CSV por clave (suma si duplicadas).
        Map<DiscrepancyKey, Double> csvMap = new HashMap<>();
        // Mantener tambien el orden de aparicion para ordenar los SOLO_CSV.
        List<DiscrepancyKey> csvOrder = new ArrayList<>();
        for (CsvImputacion imp : csvImputaciones) {
            DiscrepancyKey key = new DiscrepancyKey(
                    imp.getMatricula(), imp.getPeticion(), imp.getFuncion());
            if (!csvMap.containsKey(key)) {
                csvOrder.add(key);
            }
            csvMap.merge(key, imp.getRealizadoHoras(), Double::sum);
        }

        List<Discrepancy> diferencias = new ArrayList<>();
        List<Discrepancy> soloCsv = new ArrayList<>();
        Set<DiscrepancyKey> matched = new HashSet<>();

        // 2. Recorrer el CSV en orden y clasificar.
        for (DiscrepancyKey key : csvOrder) {
            double horasCsv = csvMap.get(key);
            Double horasResultado = resultadoMap.get(key);
            if (horasResultado == null) {
                soloCsv.add(Discrepancy.onlyInCsv(origen, key, horasCsv));
            } else {
                matched.add(key);
                // Tolerancia cero: equals exacto sobre double.
                if (Double.compare(horasCsv, horasResultado) != 0) {
                    diferencias.add(Discrepancy.difference(origen, key,
                            horasCsv, horasResultado));
                }
                // Si coinciden exacto: no hay discrepancia, no se anade.
            }
        }

        // 3. SOLO_RESULTADO: claves del mapa con la matricula correspondiente
        // al origen actual que NO fueron consumidas. La matricula clave es
        // el nombre del fichero (origen), porque cada CSV es de una sola
        // matricula. Comparamos string-a-string.
        List<Discrepancy> soloResultado = new ArrayList<>();
        for (Map.Entry<DiscrepancyKey, Double> e : resultadoMap.entrySet()) {
            DiscrepancyKey key = e.getKey();
            if (!origen.equals(key.getMatricula())) {
                continue; // clave de otra matricula, no nos toca.
            }
            if (matched.contains(key)) {
                continue;
            }
            soloResultado.add(Discrepancy.onlyInResultado(origen, key, e.getValue()));
        }

        log.info("[DiscrepancyComparator] origen={}: {} diferencias, "
                + "{} solo en CSV, {} solo en Resultado.",
                origen, diferencias.size(), soloCsv.size(), soloResultado.size());

        // 4. Concatenar en el orden documentado.
        List<Discrepancy> all = new ArrayList<>(
                diferencias.size() + soloCsv.size() + soloResultado.size());
        all.addAll(diferencias);
        all.addAll(soloCsv);
        all.addAll(soloResultado);
        return all;
    }
}

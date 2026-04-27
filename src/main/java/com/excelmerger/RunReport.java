package com.excelmerger;

import java.time.Duration;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Acumula informacion de diagnostico de una ejecucion: hojas creadas
 * (con numero de filas) y warnings categorizados. Al final, Main vuelca
 * un resumen en consola y en el log.
 *
 * <p>Categorias habituales de warning:</p>
 * <ul>
 *   <li>{@code PERFIL}   — perfiles sin coincidencia.</li>
 *   <li>{@code CABECERA} — cabeceras esperadas no encontradas.</li>
 *   <li>{@code FORMULA}  — placeholders {@code {col:X}} invalidos.</li>
 *   <li>{@code LOOKUP}   — entradas malformadas o apps sin mapeo.</li>
 *   <li>{@code HOJA}     — referencias a hojas inexistentes.</li>
 *   <li>{@code CONFIG}   — otros problemas de configuracion detectados en runtime.</li>
 * </ul>
 */
public class RunReport {

    /** Hojas creadas en el libro resultado, en orden de creacion -> nº de filas (incluida cabecera). */
    private final Map<String, Integer> sheets = new LinkedHashMap<>();

    /** Warnings acumulados durante toda la ejecucion, en el orden en el que aparecen. */
    private final List<Warning> warnings = new ArrayList<>();

    /**
     * Tabla de claves de cada hoja de lookup creada por {@link LookupSheetBuilder}.
     * Se usa para detectar en {@link MesSheetBuilder} valores de Excel (p. ej. codigos de
     * aplicacion) que se referencian con VLOOKUP pero no tienen mapeo.
     */
    private final Map<String, Set<String>> lookupKeys = new LinkedHashMap<>();

    /** v2.3.0: modo de salida usado en esta ejecucion (cierre / responsables / completo). */
    private OutputMode outputMode;

    /** Registra una hoja generada en el libro resultado, con su numero de filas. */
    public void addSheet(String name, int rows) {
        sheets.put(name, rows);
    }

    /** Anade un warning clasificado. Categoria corta en mayusculas (ver Javadoc de clase). */
    public void addWarning(String category, String message) {
        warnings.add(new Warning(category, message));
    }

    /** Asocia el conjunto de claves de una hoja de lookup (columna A, sin cabecera). */
    public void registerLookupKeys(String sheetName, Set<String> keys) {
        lookupKeys.put(sheetName, new LinkedHashSet<>(keys));
    }

    /** Devuelve las claves de un lookup, o null si no existe esa hoja. */
    public Set<String> getLookupKeys(String sheetName) {
        Set<String> s = lookupKeys.get(sheetName);
        return s == null ? null : Collections.unmodifiableSet(s);
    }

    public boolean hasLookup(String sheetName) {
        return lookupKeys.containsKey(sheetName);
    }

    /**
     * v2.3.0: registra el {@link OutputMode} efectivo usado en la ejecucion.
     * Se vuelca en {@link #formatSummary(Duration)} para trazabilidad.
     */
    public void setOutputMode(OutputMode mode) {
        this.outputMode = mode;
    }

    /** v2.3.0: devuelve el {@link OutputMode} efectivo, o {@code null} si no se ha fijado. */
    public OutputMode getOutputMode() {
        return outputMode;
    }

    public Map<String, Integer> sheets() {
        return Collections.unmodifiableMap(sheets);
    }

    public List<Warning> warnings() {
        return Collections.unmodifiableList(warnings);
    }

    public int warningCount() {
        return warnings.size();
    }

    /**
     * Formatea el resumen final listo para volcar por log.info (o System.out).
     * Multilinea con separadores '===' equivalentes a los del resto de la app.
     */
    public String formatSummary(Duration elapsed) {
        String nl = System.lineSeparator();
        StringBuilder sb = new StringBuilder();
        sb.append(nl);
        sb.append("====================================").append(nl);
        sb.append("   RESUMEN DE EJECUCION").append(nl);
        sb.append("====================================").append(nl);
        sb.append("Tiempo total: ").append(elapsed.toMillis()).append(" ms").append(nl);

        if (outputMode != null) {
            sb.append("Modo: ").append(outputMode).append(nl);
        }

        sb.append("Hojas generadas (").append(sheets.size()).append("):").append(nl);
        if (sheets.isEmpty()) {
            sb.append("  (ninguna)").append(nl);
        } else {
            for (Map.Entry<String, Integer> e : sheets.entrySet()) {
                sb.append("  - ").append(e.getKey())
                  .append(" (").append(e.getValue()).append(" filas)")
                  .append(nl);
            }
        }

        sb.append("Warnings (").append(warnings.size()).append("):").append(nl);
        if (warnings.isEmpty()) {
            sb.append("  (ninguno)").append(nl);
        } else {
            for (Warning w : warnings) {
                sb.append("  - [").append(w.category).append("] ")
                  .append(w.message).append(nl);
            }
        }
        sb.append("====================================");
        return sb.toString();
    }

    // ------------------------------------------------------------------
    //  Tipo interno
    // ------------------------------------------------------------------
    public static final class Warning {
        public final String category;
        public final String message;

        Warning(String category, String message) {
            this.category = category;
            this.message = message;
        }

        @Override
        public String toString() {
            return "[" + category + "] " + message;
        }
    }
}

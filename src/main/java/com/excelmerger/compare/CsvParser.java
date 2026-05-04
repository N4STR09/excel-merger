package com.excelmerger.compare;

import com.excelmerger.exception.InputValidationException;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

/**
 * Parser de CSV exportados por el ERP, formato {@code <matricula>.CSV}, para
 * el comprobador de discrepancias (Opcion 2 del menu, v3.1.0).
 *
 * <h2>Formato confirmado por inspeccion de los CSV reales</h2>
 * <ul>
 *   <li>Encoding: <b>Windows-1252 / ISO-8859-1</b> (cp1252). NO es UTF-8;
 *       leerlo como UTF-8 rompe las tildes de "Petición" y "Descripción".</li>
 *   <li>Separador: {@code ;}.</li>
 *   <li>Quoting: caracter {@code "}.</li>
 *   <li>Decimal: coma. Numericos vienen con padding de espacios delante
 *       (p. ej. {@code "       5,0"}, {@code "        ,0"} para cero).</li>
 *   <li>Caracter especial: aparece {@code \u0000} (byte NUL) en el campo
 *       {@code DR-Marca} cuando esta vacio. Limpieza preventiva global.</li>
 *   <li>Fin de linea: {@code \r\n}.</li>
 *   <li>Cada fila tiene 27 campos: las 26 cabeceras declaradas + 1 campo
 *       "fantasma" final lleno de espacios. Se ignora.</li>
 *   <li>Cabeceras con higiene cuestionable: {@code "Matricula "} viene con
 *       espacio al final. Se hace trim para localizar las columnas.</li>
 * </ul>
 *
 * <h2>Normalizacion aplicada</h2>
 * <ul>
 *   <li>Trim global de cabeceras y valores.</li>
 *   <li>Eliminacion de TODOS los caracteres {@code \u0000} de cada valor
 *       (limpieza preventiva, no solo en {@code DR-Marca}).</li>
 *   <li>{@code Petición} pierde la {@code J} prefijada: {@code "J137791"}
 *       se convierte en {@code "137791"}.</li>
 *   <li>{@code Realizado Horas} se parsea de string {@code "       5,0"}
 *       a double {@code 5.0} con la coma como separador decimal.</li>
 * </ul>
 *
 * <h2>Mapeo por indice (no por nombre directo)</h2>
 * <p>Apache Commons CSV permite acceso por nombre de cabecera, pero la
 * cabecera real {@code "Matricula "} (con espacio) y el campo fantasma
 * final hacen mas robusto leer la primera fila como datos y construir un
 * mapa nombre-normalizado -&gt; indice. Asi el parser sigue funcionando si
 * el ERP arregla la higiene de la cabecera o anade columnas nuevas.</p>
 *
 * <h2>Errores</h2>
 * <p>Lanza {@link InputValidationException} si el fichero no se puede leer
 * o si faltan cabeceras obligatorias. Filas individuales con
 * {@code Realizado Horas} no parseable se logan en WARN y se ignoran (no
 * abortan el parsing del fichero entero).</p>
 */
public final class CsvParser {

    private static final Logger log = LoggerFactory.getLogger(CsvParser.class);

    /**
     * Encoding de los CSV del ERP. Windows-1252 (cp1252). Es lo que
     * Excel "exporta como CSV" en Windows en español por defecto.
     */
    public static final Charset DEFAULT_CHARSET = Charset.forName("windows-1252");

    /** Cabecera obligatoria: matricula del recurso. */
    private static final String COL_MATRICULA = "Matricula";
    /** Cabecera obligatoria: codigo de peticion (con J prefijada). */
    private static final String COL_PETICION = "Petición";
    /** Cabecera obligatoria: codigo de funcion (2 letras). */
    private static final String COL_FUNCION = "Fu";
    /** Cabecera obligatoria: horas realizadas (decimal con coma). */
    private static final String COL_REALIZADO = "Realizado Horas";

    /** Caracter NUL ({@code \u0000}) que aparece en {@code DR-Marca}. */
    private static final String NUL_STRING = "\u0000";

    /**
     * Lee el CSV en {@code path} y devuelve la lista de imputaciones
     * normalizadas. Usa el encoding por defecto ({@link #DEFAULT_CHARSET}).
     *
     * @throws InputValidationException si el fichero no existe, no se puede
     *         leer o no contiene las cabeceras obligatorias.
     */
    public List<CsvImputacion> parse(Path path) {
        return parse(path, DEFAULT_CHARSET);
    }

    /**
     * Lee el CSV en {@code path} con el encoding indicado. Permite forzar
     * otro encoding en tests si se proporciona un fixture en UTF-8.
     */
    public List<CsvImputacion> parse(Path path, Charset charset) {
        if (path == null) {
            throw new InputValidationException("Path al CSV es null.");
        }
        if (!Files.exists(path)) {
            throw new InputValidationException("CSV no encontrado: " + path.toAbsolutePath());
        }
        if (!Files.isRegularFile(path)) {
            throw new InputValidationException("Ruta no es un fichero regular: " + path.toAbsolutePath());
        }
        try (InputStream in = Files.newInputStream(path);
             Reader reader = new InputStreamReader(in, charset)) {
            return parseInternal(reader, path.toString());
        } catch (IOException e) {
            throw new InputValidationException(
                    "Error leyendo CSV " + path.toAbsolutePath() + ": " + e.getMessage(), e);
        }
    }

    /**
     * Variante para tests: parsea desde un {@link Reader} arbitrario.
     * El {@code sourceLabel} es solo para mensajes de log/error.
     */
    public List<CsvImputacion> parse(Reader reader, String sourceLabel) {
        try {
            return parseInternal(reader, sourceLabel);
        } catch (IOException e) {
            throw new InputValidationException(
                    "Error leyendo CSV " + sourceLabel + ": " + e.getMessage(), e);
        }
    }

    private List<CsvImputacion> parseInternal(Reader reader, String source) throws IOException {
        // Formato: separador ; quoting con ", sin trim automatico (lo
        // hacemos nosotros tras limpiar el NUL para garantizar el orden).
        // No usamos setHeader() para evitar el problema del campo fantasma
        // final de espacios y de "Matricula " con trailing space.
        // Nota: en commons-csv 1.11.0 el metodo terminal del builder es
        // .build() (en 1.13+ se renombra a .get()).
        CSVFormat format = CSVFormat.DEFAULT.builder()
                .setDelimiter(';')
                .setQuote('"')
                .setIgnoreEmptyLines(true)
                .build();

        try (CSVParser parser = CSVParser.parse(reader, format)) {
            Iterator<CSVRecord> it = parser.iterator();
            if (!it.hasNext()) {
                throw new InputValidationException(
                        "CSV vacio (sin cabecera): " + source);
            }
            CSVRecord headerRow = it.next();
            int idxMatricula = findHeaderIndex(headerRow, COL_MATRICULA, source);
            int idxPeticion  = findHeaderIndex(headerRow, COL_PETICION, source);
            int idxFuncion   = findHeaderIndex(headerRow, COL_FUNCION, source);
            int idxRealizado = findHeaderIndex(headerRow, COL_REALIZADO, source);

            List<CsvImputacion> result = new ArrayList<>();
            int lineNum = 1; // cabecera era linea 1
            while (it.hasNext()) {
                CSVRecord record = it.next();
                lineNum++;
                CsvImputacion imp = parseRow(record, lineNum, source,
                        idxMatricula, idxPeticion, idxFuncion, idxRealizado);
                if (imp != null) {
                    result.add(imp);
                }
            }
            log.info("[CsvParser] Leidas {} imputaciones de {}", result.size(), source);
            return result;
        }
    }

    /**
     * Localiza el indice (0-based) de una cabecera por nombre, comparando
     * tras aplicar {@link #cleanField(String)} (trim + descarte de NUL).
     * Esto absorbe el {@code "Matricula "} con espacio al final y cualquier
     * NUL pegado a una cabecera por contaminacion del export.
     */
    private static int findHeaderIndex(CSVRecord headerRow, String wanted, String source) {
        for (int i = 0; i < headerRow.size(); i++) {
            String h = cleanField(headerRow.get(i));
            if (h.equals(wanted)) {
                return i;
            }
        }
        // Lista de cabeceras encontradas para diagnostico.
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < headerRow.size(); i++) {
            if (i > 0) sb.append(", ");
            sb.append('"').append(cleanField(headerRow.get(i))).append('"');
        }
        throw new InputValidationException(
                "El CSV " + source + " no contiene la cabecera obligatoria \""
                        + wanted + "\". Cabeceras encontradas: [" + sb + "]");
    }

    /**
     * Parsea una fila de datos. Devuelve {@code null} si la fila esta vacia
     * (todos los campos relevantes en blanco) o si {@code Realizado Horas}
     * no se puede parsear (en cuyo caso ademas se loga warning).
     */
    private static CsvImputacion parseRow(CSVRecord record, int lineNum, String source,
                                          int idxMatricula, int idxPeticion,
                                          int idxFuncion, int idxRealizado) {
        // record.size() puede ser menor que el numero de cabeceras si hay
        // filas truncadas (raro, pero defensivo).
        int max = record.size();
        if (idxMatricula >= max || idxPeticion >= max
                || idxFuncion >= max || idxRealizado >= max) {
            log.warn("[CsvParser] {} linea {}: fila truncada (size={}). Ignorada.",
                    source, lineNum, max);
            return null;
        }

        String matricula = cleanField(record.get(idxMatricula));
        String peticionRaw = cleanField(record.get(idxPeticion));
        String funcion = cleanField(record.get(idxFuncion));
        String realizadoStr = cleanField(record.get(idxRealizado));

        // Quitar la J prefijada de la Peticion (case-insensitive por
        // defensiva: el ERP siempre usa mayuscula pero no cuesta nada).
        String peticion = stripJPrefix(peticionRaw);

        // Si los 4 campos clave estan vacios, fila inutil -> ignora.
        if (matricula.isEmpty() && peticion.isEmpty()
                && funcion.isEmpty() && realizadoStr.isEmpty()) {
            return null;
        }

        // Realizado Horas: decimal con coma. Padding ya quitado por trim.
        // Los CSV reales muestran ",0" para cero (sin parte entera).
        double realizadoHoras;
        try {
            realizadoHoras = parseSpanishDecimal(realizadoStr);
        } catch (NumberFormatException e) {
            log.warn("[CsvParser] {} linea {}: Realizado Horas no parseable: \"{}\". Fila ignorada.",
                    source, lineNum, realizadoStr);
            return null;
        }

        return new CsvImputacion(peticion, matricula, funcion, realizadoHoras);
    }

    /**
     * Limpieza global de un campo del CSV: descarta TODOS los {@code \u0000}
     * y aplica {@code trim()}. Devuelve cadena vacia si el resultado es
     * solo whitespace.
     */
    static String cleanField(String raw) {
        if (raw == null) return "";
        String noNul = raw.replace(NUL_STRING, "");
        return noNul.trim();
    }

    /**
     * Quita la {@code J} prefijada de un codigo de peticion. Funciona con
     * mayuscula y minuscula. Si no hay prefijo, devuelve la cadena tal cual.
     */
    static String stripJPrefix(String peticion) {
        if (peticion == null || peticion.isEmpty()) return "";
        char first = peticion.charAt(0);
        if (first == 'J' || first == 'j') {
            return peticion.substring(1);
        }
        return peticion;
    }

    /**
     * Parsea un decimal con coma (estilo español). Acepta el caso especial
     * {@code ",0"} (cero sin parte entera). Reemplaza la coma por punto y
     * delega en {@link Double#parseDouble(String)}.
     *
     * @throws NumberFormatException si el string no es un numero valido.
     */
    static double parseSpanishDecimal(String raw) {
        if (raw == null || raw.isEmpty()) {
            throw new NumberFormatException("cadena vacia");
        }
        String normalized = raw.replace(',', '.');
        // Si empieza por ".", anade el cero ("., 0" -> "0.0"). Lo mismo
        // para signo+punto: "-.5" -> "-0.5". Locale.ROOT no aplica aqui
        // porque ya hemos sustituido la coma a mano.
        if (normalized.startsWith(".")) {
            normalized = "0" + normalized;
        } else if ((normalized.startsWith("-") || normalized.startsWith("+"))
                && normalized.length() > 1 && normalized.charAt(1) == '.') {
            normalized = normalized.charAt(0) + "0" + normalized.substring(1);
        }
        return Double.parseDouble(normalized);
    }

    /**
     * Helper: dado un nombre de fichero (p. ej. {@code "90014.CSV"}),
     * extrae el origen (matricula del CSV). Nombre sin extension. Se usa
     * para la columna {@code Origen} del Excel de discrepancias.
     */
    public static String originFromFileName(String fileName) {
        if (fileName == null) return "";
        String name = fileName;
        int dot = name.lastIndexOf('.');
        if (dot > 0) {
            name = name.substring(0, dot);
        }
        return name.trim();
    }

    /**
     * Detecta si un fichero del directorio de entrada es un CSV (por
     * extension). Case-insensitive, acepta {@code .csv}, {@code .CSV},
     * {@code .Csv}, etc.
     */
    public static boolean isCsvFileName(String fileName) {
        return fileName != null && fileName.toLowerCase(Locale.ROOT).endsWith(".csv");
    }
}

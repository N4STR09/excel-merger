package com.excelmerger;

import com.excelmerger.util.PoiUtils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/**
 * v2.7.1 — Filtra de la hoja {@code Resultado} las filas en las que las 5
 * columnas numericas relevantes (Jira, Facturar, PDCL, PDCL + Deuda y
 * Horas_Mes) evaluan TODAS a {@code 0}. Las filas se eliminan FISICAMENTE
 * del workbook (no son ocultacion via {@code setZeroHeight}).
 *
 * <p>Activacion: clave {@code mes.removeEmptyRows} (default {@code true}).
 * Si la clave esta a {@code false}, el comportamiento es identico al de
 * v2.7.0 (no se filtra nada).</p>
 *
 * <h2>Por que post-build con FormulaEvaluator</h2>
 *
 * <p>Las 5 columnas son fundamentalmente formulas (SUMIFS para Jira,
 * FORMULA {@code {col:Jira}*1.2} para Facturar y PDCL,
 * FORMULA_PLUS_SUMIFS para PDCL + Deuda; Horas_Mes es un COPY directo
 * desde la hoja origen pero puede ser numerico o vacio). Replicar la
 * logica de los SUMIFS y de las multiplicaciones en Java seria deuda
 * tecnica que duplicaria reglas: si manyana cambia una formula en
 * config.properties, el filtrado divergiria. Usando FormulaEvaluator,
 * la "verdad" es exactamente lo que Excel calculara al abrir el fichero.
 * Es tambien coherente con la regla inquebrantable 4 (los tests usan
 * FormulaEvaluator para todo lo que involucre formulas).</p>
 *
 * <h2>Criterio "vale 0"</h2>
 *
 * <p>Una columna se considera "vale 0" si:</p>
 * <ul>
 *   <li>la celda es {@code BLANK} o no existe;</li>
 *   <li>la celda evalua a {@code NUMERIC} con valor exactamente {@code 0.0};</li>
 *   <li>la celda evalua a {@code STRING} con valor vacio, {@code "-"} (sentinela
 *       de filas huerfanas), o que parsea a {@code 0.0} con punto o coma
 *       decimal — {@code "0"}, {@code "0.0"}, {@code "0.00"}, {@code "0,00"}.
 *       Esto ultimo cubre el caso real de la columna {@code Horas_Mes}, que el
 *       ERP serializa como texto formateado en lugar de numerico (v2.7.1
 *       diagnostico post-release).</li>
 * </ul>
 *
 * <p>Cualquier otra evaluacion (NUMERIC distinto de 0, STRING con texto,
 * BOOLEAN, ERROR) cuenta como "no es 0" y por tanto la fila NO se filtra.
 * El criterio es conservador: ante la duda, conservar la fila.</p>
 *
 * <h2>Mecanica de eliminacion</h2>
 *
 * <p>Identificamos las filas a eliminar evaluandolas in-situ. Despues
 * recorremos la hoja desde el principio compactando: para cada fila i,
 * si esta marcada para borrar la saltamos; si no, la copiamos al
 * siguiente slot disponible. Tras la pasada, eliminamos el rabo (filas
 * sobrantes al final). Esto evita las trampas conocidas de
 * {@link Sheet#shiftRows} con merged regions y formato condicional, y es
 * lineal en el numero de filas.</p>
 *
 * <p>POI XSSF actualiza referencias relativas dentro de cada fila al
 * usar {@link Sheet#shiftRows}. Como aqui NO usamos shiftRows sino que
 * reescribimos las celdas en su posicion final, las formulas se
 * reescriben tambien — y como las formulas que generamos en
 * {@link MesSheetBuilder} usan referencias absolutas a otras hojas y
 * referencias relativas a la propia fila por NUMERO de fila Excel
 * (ej. {@code J5*1.2}), debemos REGENERAR las formulas relativas locales
 * para apuntar al numero de fila final. Para mantenerlo simple, usamos
 * una transformacion textual: reemplazamos los numeros de fila locales
 * en la formula por el nuevo numero de fila destino. La transformacion
 * solo aplica a referencias <i>sin nombre de hoja delante</i> (las que
 * apuntan a la hoja actual); las referencias con prefijo {@code Hoja!}
 * quedan intactas.</p>
 *
 * <p>El formato condicional {@code redIfNotEqualTo} y {@code greenIfPositive}
 * se aplica DESPUES del filtrado en {@link MesSheetBuilder#build}, por lo
 * que sus rangos cubren ya solo las filas supervivientes y no requieren
 * ajuste adicional aqui.</p>
 *
 * <p>El formato CORAL (rojo si != PDCL) que aplica POI usa formula relativa
 * sobre rangos de columna; se reemite tras el filtrado y por tanto no
 * sufre desfase.</p>
 */
final class EmptyRowFilter {

    private static final Logger log = LoggerFactory.getLogger(EmptyRowFilter.class);

    /**
     * Nombres de las 5 columnas (en {@code mes.col.N.name}) que componen
     * el criterio AND de "fila vacia". Cableado en v2.7.1 (decision Fase 0,
     * P7); si en el futuro hace falta parametrizar, sera por una clave nueva
     * tipo {@code mes.emptyRowColumns=...} sin romper retrocompatibilidad.
     */
    static final List<String> EMPTY_ROW_COLUMN_NAMES = List.of(
            "Jira", "Facturar", "PDCL", "PDCL + Deuda", "Horas_Mes");

    /**
     * Categoria de warning emitida en RunReport cuando alguna de las 5
     * columnas no existe en el config actual. Se usa la categoria CONFIG
     * existente en lugar de crear una nueva, para no sumar ruido.
     */
    private static final String WARN_CATEGORY = "CONFIG";

    private EmptyRowFilter() {
        // Utility class
    }

    /**
     * Aplica el filtro a la hoja {@code mes}. Si las 5 columnas no se
     * encuentran (todas o algunas), emite un warning y no filtra nada
     * (fail-safe: ante config no estandar, mantener comportamiento v2.7.0).
     *
     * @param workbook libro destino (necesario para crear FormulaEvaluator).
     * @param mes hoja Resultado ya construida (cabecera en fila 0, datos en
     *            filas 1..lastDataRow).
     * @param lastDataRow indice 0-based de la ultima fila de datos antes
     *            de filtrar (es decir, {@code mesRowIdx} en
     *            {@link MesSheetBuilder#build}).
     * @param report acumulador de warnings.
     * @return numero de filas eliminadas (0 si no se filtra nada).
     */
    static int apply(Workbook workbook, Sheet mes, int lastDataRow, RunReport report) {
        if (lastDataRow < 1) {
            // Hoja sin datos: nada que filtrar.
            return 0;
        }
        Row header = mes.getRow(0);
        if (header == null) {
            return 0;
        }

        // Resolver indices de las 5 columnas por nombre. Si alguna falta,
        // warning y abortar el filtrado (no es seguro filtrar parcialmente).
        int[] colIdx = new int[EMPTY_ROW_COLUMN_NAMES.size()];
        List<String> missing = new ArrayList<>();
        for (int i = 0; i < EMPTY_ROW_COLUMN_NAMES.size(); i++) {
            String name = EMPTY_ROW_COLUMN_NAMES.get(i);
            int idx = PoiUtils.findColumnIndex(header, name);
            colIdx[i] = idx;
            if (idx < 0) {
                missing.add(name);
            }
        }
        if (!missing.isEmpty()) {
            String msg = "mes.removeEmptyRows=true pero no se encontraron las columnas "
                    + missing + " en la hoja '" + mes.getSheetName()
                    + "'. Filtrado de filas vacias omitido (comportamiento v2.7.0).";
            log.warn("{}", msg);
            report.addWarning(WARN_CATEGORY, msg);
            return 0;
        }

        FormulaEvaluator evaluator;
        try {
            evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            // v2.7.1.1: pre-evaluar todas las formulas del workbook. POI a
            // veces falla al evaluar individualmente formulas SUMIFS
            // cross-sheet en libros recien construidos (no hay cache previa
            // y las hojas referenciadas pueden no estar registradas todavia
            // en el evaluator). evaluateAll() las recompila en bloque y
            // poblando la cache evita timing issues en evaluate() individuales.
            evaluator.evaluateAll();
        } catch (RuntimeException e) {
            log.warn("No se pudo crear/inicializar FormulaEvaluator para filtrar filas vacias: {}",
                    e.getMessage());
            report.addWarning(WARN_CATEGORY,
                    "No se pudo crear FormulaEvaluator para filtrar filas vacias: "
                            + e.getMessage() + ". Filtrado omitido.");
            return 0;
        }

        // Identificar filas a eliminar (set de indices 0-based).
        Set<Integer> toRemove = new LinkedHashSet<>();
        for (int r = 1; r <= lastDataRow; r++) {
            Row row = mes.getRow(r);
            if (row == null) continue;
            if (rowHasAllZeroes(row, colIdx, evaluator)) {
                toRemove.add(r);
            }
        }
        if (toRemove.isEmpty()) {
            return 0;
        }

        // Compactar: copiar filas supervivientes hacia arriba.
        compactRows(mes, lastDataRow, toRemove);

        log.info("[EmptyRowFilter] {} fila(s) eliminadas de '{}' por tener las 5 columnas {} a 0.",
                toRemove.size(), mes.getSheetName(), EMPTY_ROW_COLUMN_NAMES);
        return toRemove.size();
    }

    /**
     * Devuelve {@code true} sii TODAS las celdas indicadas en la fila
     * evaluan a "0" segun el criterio definido en la JavaDoc de la clase.
     * Si alguna celda no se puede evaluar (formula con error, por ejemplo),
     * devuelve {@code false} (conservador: no filtrar ante incertidumbre).
     */
    private static boolean rowHasAllZeroes(Row row, int[] colIdx, FormulaEvaluator evaluator) {
        for (int idx : colIdx) {
            Cell cell = row.getCell(idx);
            if (!cellIsZero(cell, evaluator)) {
                return false;
            }
        }
        return true;
    }

    /**
     * Decide si una celda concreta cuenta como "0" para el criterio del
     * filtro. Ver criterio en JavaDoc de la clase.
     */
    private static boolean cellIsZero(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) return true;
        CellType type = cell.getCellType();
        if (type == CellType.BLANK) return true;

        if (type == CellType.FORMULA) {
            CellValue cv;
            try {
                cv = evaluator.evaluate(cell);
            } catch (RuntimeException e) {
                // No se pudo evaluar: conservador, no es 0.
                return false;
            }
            if (cv == null) return true;
            CellType resultType = cv.getCellType();
            switch (resultType) {
                case BLANK:
                    return true;
                case NUMERIC:
                    return cv.getNumberValue() == 0.0;
                case STRING:
                    return isStringZero(cv.getStringValue());
                default:
                    // BOOLEAN, ERROR, etc. -> no es 0
                    return false;
            }
        }
        switch (type) {
            case NUMERIC:
                return cell.getNumericCellValue() == 0.0;
            case STRING:
                return isStringZero(cell.getStringCellValue());
            default:
                return false;
        }
    }

    /**
     * Una STRING cuenta como "0" si:
     * <ul>
     *   <li>esta vacia (ignorando espacios);</li>
     *   <li>coincide exactamente con el sentinela {@code "-"} (con trim) — usado
     *       en filas huerfanas para columnas COPY sin valor;</li>
     *   <li>parsea numericamente a {@code 0.0}, sea con punto o coma decimal.
     *       Esto cubre los strings que vienen del ERP como {@code "0.00"},
     *       {@code "0,00"}, {@code "0"}, {@code "0.0"} en la columna
     *       {@code Horas_Mes} (la fuente de datos del ERP serializa la columna
     *       {@code UltimaPrevision_Horas_Mes} como texto formateado, no como
     *       numerico). Una string como {@code "0.5"} NO cuenta como cero.</li>
     * </ul>
     */
    private static boolean isStringZero(String s) {
        if (s == null) return true;
        String trimmed = s.trim();
        if (trimmed.isEmpty() || "-".equals(trimmed)) {
            return true;
        }
        // Aceptar coma decimal del locale es_ES y punto del locale en. No
        // intentamos manejar separadores de millares: en este dato concreto
        // (Horas_Mes en horas, valores tipicos < 1000) no aparecen.
        String normalized = trimmed.replace(',', '.');
        try {
            return Double.parseDouble(normalized) == 0.0;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /**
     * Compacta {@code mes} eliminando las filas indicadas en {@code toRemove}.
     * Implementacion sin {@link Sheet#shiftRows}: copia los valores de las
     * filas supervivientes a su nueva posicion (rellenando huecos), y al
     * final elimina las filas sobrantes del rabo. Asi:
     * <ul>
     *   <li>se preservan estilos de celda (fill, conditional fills, etc.);</li>
     *   <li>las formulas con referencias a la propia fila se REESCRIBEN
     *       traduciendo los numeros de fila origen al destino;</li>
     *   <li>las referencias absolutas y referencias a otras hojas no se
     *       tocan (la regex no las mata).</li>
     * </ul>
     */
    private static void compactRows(Sheet mes, int lastDataRow, Set<Integer> toRemove) {
        // Numero de columnas a copiar: el maximo lastCellNum visto en la
        // cabecera (asumimos cabecera siempre completa) o en cualquier fila
        // de datos (por seguridad).
        int maxCells = 0;
        Row header = mes.getRow(0);
        if (header != null) {
            maxCells = Math.max(maxCells, header.getLastCellNum());
        }
        for (int r = 1; r <= lastDataRow; r++) {
            Row row = mes.getRow(r);
            if (row != null) {
                maxCells = Math.max(maxCells, row.getLastCellNum());
            }
        }

        int writeIdx = 1;
        for (int readIdx = 1; readIdx <= lastDataRow; readIdx++) {
            if (toRemove.contains(readIdx)) {
                continue;
            }
            if (writeIdx == readIdx) {
                writeIdx++;
                continue;
            }
            // Copiar fila readIdx -> writeIdx. Si writeIdx ya existe (no
            // deberia, pero por seguridad si la hoja tiene huecos previos),
            // la eliminamos antes de recrearla.
            Row src = mes.getRow(readIdx);
            Row existingDst = mes.getRow(writeIdx);
            if (existingDst != null) {
                mes.removeRow(existingDst);
            }
            Row dst = mes.createRow(writeIdx);
            if (src != null) {
                copyRowShallow(src, dst, maxCells, readIdx, writeIdx);
            }
            writeIdx++;
        }

        // Eliminar el rabo: filas con indice >= writeIdx hasta lastDataRow.
        for (int r = lastDataRow; r >= writeIdx; r--) {
            Row row = mes.getRow(r);
            if (row != null) {
                mes.removeRow(row);
            }
        }
    }

    /**
     * Copia la fila {@code src} a la fila {@code dst}, reescribiendo las
     * formulas locales con la traduccion de numero de fila
     * {@code srcExcelRow -> dstExcelRow}.
     *
     * <p>Indices Excel (1-based): {@code srcExcelRow = readIdx0 + 1},
     * {@code dstExcelRow = writeIdx0 + 1}.</p>
     */
    private static void copyRowShallow(Row src, Row dst, int maxCells,
                                       int readIdx0, int writeIdx0) {
        int srcExcelRow = readIdx0 + 1;
        int dstExcelRow = writeIdx0 + 1;
        // v2.7.1.1: el nombre de la hoja actual se necesita para traducir
        // tambien referencias absolutas a la PROPIA hoja en formulas
        // FORMULA_PLUS_SUMIFS, que generan "Resultado!$A$<row>" etc.
        String currentSheetName = src.getSheet().getSheetName();
        for (int c = 0; c < maxCells; c++) {
            Cell sCell = src.getCell(c);
            if (sCell == null) continue;
            Cell dCell = dst.createCell(c);
            // Estilo: compartido es seguro porque los CellStyle son por
            // workbook y POI los cachea internamente; al asignarlos en la
            // celda destino quedan referenciados.
            if (sCell.getCellStyle() != null) {
                dCell.setCellStyle(sCell.getCellStyle());
            }
            switch (sCell.getCellType()) {
                case STRING:
                    dCell.setCellValue(sCell.getStringCellValue());
                    break;
                case NUMERIC:
                    dCell.setCellValue(sCell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    dCell.setCellValue(sCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    String original = sCell.getCellFormula();
                    String translated = translateLocalRowRefs(
                            original, srcExcelRow, dstExcelRow, currentSheetName);
                    dCell.setCellFormula(translated);
                    break;
                case BLANK:
                    // nada
                    break;
                case ERROR:
                    dCell.setCellErrorValue(sCell.getErrorCellValue());
                    break;
                default:
                    // tipos nuevos de POI futuros: no copiar valor
                    break;
            }
        }
    }

    /**
     * Traduce las referencias de fila en una formula tras una compactacion
     * que mueve el contenido de la fila Excel {@code srcExcelRow} a
     * {@code dstExcelRow}. Hay dos clases de referencias que se reescriben:
     *
     * <ol>
     *   <li><b>Locales sin prefijo de hoja</b>: patrones tipo {@code COL+ROWNUM}
     *       (no absolutas, no precedidos de identificador), por ejemplo
     *       {@code J5*1.2} -> {@code J4*1.2}.</li>
     *   <li><b>Auto-referencias a la propia hoja con prefijo nombre</b>: patrones
     *       {@code <currentSheetName>!{$?}COL{$?}ROWNUM}, por ejemplo
     *       {@code Resultado!$A$5} -> {@code Resultado!$A$4} cuando
     *       {@code currentSheetName == "Resultado"}. <i>Aqui SI traducimos
     *       referencias absolutas de fila</i> porque las genera
     *       {@link com.excelmerger.sheet.column.FormulaPlusSumIfsColumnStrategy}
     *       como {@code Resultado!$<col>$<row>} para los criterios de SUMIFS
     *       contra Deuda; esos criterios deben seguir apuntando a la propia
     *       fila tras la compactacion.</li>
     * </ol>
     *
     * <p>Las referencias a OTRAS hojas (Cierre, Deuda, Equipos, etc.) NO se
     * tocan en ningun caso: si la fila Resultado original 66 fue construida
     * para apuntar a {@code Cierre!$A$66}, esa correspondencia se preserva
     * tras la compactacion porque la informacion en Cierre no se mueve.</p>
     *
     * <p>Heuristica conservadora: ante una formula con sintaxis no
     * contemplada, se devuelve sin tocar (idempotente respecto a v2.7.0).</p>
     */
    static String translateLocalRowRefs(String formula, int srcExcelRow,
                                        int dstExcelRow, String currentSheetName) {
        if (formula == null || formula.isEmpty() || srcExcelRow == dstExcelRow) {
            return formula;
        }
        String src = String.valueOf(srcExcelRow);
        String dst = String.valueOf(dstExcelRow);
        // Forma "quoted" del nombre de hoja para el matcheo: PoiUtils.quoteSheetName
        // envuelve en comillas si hay espacios o caracteres especiales. La
        // forma sin comillas (nombre simple) tambien debe matchearse para
        // hojas como "Resultado".
        String selfPrefixSimple = currentSheetName == null ? null : currentSheetName + "!";
        String selfPrefixQuoted = currentSheetName == null ? null : "'" + currentSheetName + "'!";

        StringBuilder out = new StringBuilder(formula.length() + 8);
        int i = 0;
        int n = formula.length();
        while (i < n) {
            // 1. Intentar match de auto-referencia con prefijo de hoja propia.
            int next = currentSheetName == null
                    ? i
                    : tryMatchSelfRef(formula, i, n, src, dst, out,
                            selfPrefixSimple, selfPrefixQuoted);
            if (next > i) {
                i = next;
                continue;
            }
            // 2. Intentar match de referencia local (sin prefijo de hoja).
            next = tryMatchLocalRef(formula, i, n, src, dst, out);
            if (next > i) {
                i = next;
            } else {
                out.append(formula.charAt(i));
                i++;
            }
        }
        return out.toString();
    }

    /**
     * Sobrecarga sin nombre de hoja: equivalente a llamar con
     * {@code currentSheetName=null}. Solo traduce referencias locales sin
     * prefijo. Mantenida para compatibilidad de los tests unitarios que no
     * dependen del nombre de hoja.
     */
    static String translateLocalRowRefs(String formula, int srcExcelRow, int dstExcelRow) {
        return translateLocalRowRefs(formula, srcExcelRow, dstExcelRow, null);
    }

    /**
     * Intenta matchear en la posicion {@code i} una auto-referencia tipo
     * {@code <currentSheet>!{$?}COL{$?}srcRow} (con o sin comillas alrededor
     * del nombre de hoja). Si match, escribe la sustitucion en {@code out}
     * (preservando los {@code $} originales del prefijo y de la columna,
     * cambiando solo el numero de fila si no era absoluto, y traduciendolo
     * tambien si lo era). Devuelve la nueva posicion o {@code i} si no hay
     * match.
     *
     * <p>Detalle clave: <b>siempre</b> se traduce el numero de fila si esta
     * presente y coincide con {@code src}, incluso si lleva {@code $} delante
     * (referencia absoluta). El caso real de
     * {@link com.excelmerger.sheet.column.FormulaPlusSumIfsColumnStrategy} es
     * {@code Resultado!$A$66} (fila absoluta) y debe re-numerarse al filtrar.</p>
     */
    private static int tryMatchSelfRef(String formula, int i, int n,
                                       String src, String dst, StringBuilder out,
                                       String selfPrefixSimple, String selfPrefixQuoted) {
        int prefixLen = detectSelfPrefixLen(formula, i, selfPrefixSimple, selfPrefixQuoted);
        if (prefixLen <= 0) {
            return i;
        }
        return tryMatchColRowAfterPrefix(formula, i, i + prefixLen, n, src, dst, out);
    }

    /**
     * Devuelve la longitud del prefijo de auto-referencia detectado en la
     * posicion {@code i} de {@code formula}, o {@code 0} si no hay match.
     * Si el prefijo simple {@code currentSheet!} encaja, comprueba que el
     * caracter previo no sea alfanumerico (para evitar matchear
     * {@code MiResultado!} cuando la hoja actual es {@code Resultado}). El
     * prefijo entrecomillado {@code 'currentSheet'!} no necesita esa
     * comprobacion porque la comilla simple ya delimita.
     */
    private static int detectSelfPrefixLen(String formula, int i,
                                            String selfPrefixSimple, String selfPrefixQuoted) {
        if (selfPrefixSimple != null
                && formula.regionMatches(i, selfPrefixSimple, 0, selfPrefixSimple.length())
                && !precedingCharIsIdentifierPart(formula, i)) {
            return selfPrefixSimple.length();
        }
        if (selfPrefixQuoted != null
                && formula.regionMatches(i, selfPrefixQuoted, 0, selfPrefixQuoted.length())) {
            return selfPrefixQuoted.length();
        }
        return 0;
    }

    /** Devuelve {@code true} si el caracter anterior a {@code i} podria
     *  hacer que la posicion {@code i} sea continuacion de un identificador
     *  mas largo (letra, digito o {@code _}). En la posicion 0 devuelve
     *  {@code false}. */
    private static boolean precedingCharIsIdentifierPart(String formula, int i) {
        if (i <= 0) return false;
        char prevCh = formula.charAt(i - 1);
        return isLetter(prevCh) || isDigit(prevCh) || prevCh == '_';
    }

    /**
     * Tras detectar prefijo {@code <hoja>!} en posicion {@code i..j-1},
     * intenta consumir {@code {$?}COL{$?}ROW} y reescribir si la fila
     * coincide con {@code src}. Devuelve nueva posicion o {@code i} si no.
     */
    private static int tryMatchColRowAfterPrefix(String formula, int i, int j, int n,
                                                  String src, String dst,
                                                  StringBuilder out) {
        if (j >= n) return i;
        // Posible '$' antes de la columna
        boolean dollarBeforeCol = formula.charAt(j) == '$';
        int colStart = dollarBeforeCol ? j + 1 : j;
        if (colStart >= n || !isLetter(formula.charAt(colStart))) {
            return i;
        }
        int colEnd = colStart;
        while (colEnd < n && isLetter(formula.charAt(colEnd))) colEnd++;
        if (colEnd >= n) return i;
        // Posible '$' antes del numero de fila
        boolean dollarBeforeRow = formula.charAt(colEnd) == '$';
        int rowStart = dollarBeforeRow ? colEnd + 1 : colEnd;
        if (rowStart >= n || !isDigit(formula.charAt(rowStart))) {
            return i;
        }
        int rowEnd = rowStart;
        while (rowEnd < n && isDigit(formula.charAt(rowEnd))) rowEnd++;
        String num = formula.substring(rowStart, rowEnd);
        if (!num.equals(src)) {
            return i;
        }
        // Match: emitir prefijo + col + row traducida.
        out.append(formula, i, rowStart);
        out.append(dst);
        return rowEnd;
    }

    /**
     * Si en la posicion {@code i} de {@code formula} comienza una referencia
     * local del tipo {@code COL+srcRow} valida (no precedida de {@code '!'}
     * ni de caracter alfanumerico, sin {@code '$'} entre columna y numero,
     * con el numero coincidiendo exactamente con {@code src}), escribe en
     * {@code out} la sustitucion {@code COL+dstRow} y devuelve la nueva
     * posicion (justo despues del numero original). Si no hay match, devuelve
     * {@code i} (el caller sabe que debe avanzar un caracter).
     *
     * <p>Helper extraido para evitar el if-anidado profundo de PMD
     * {@code AvoidDeeplyNestedIfStmts} en {@link #translateLocalRowRefs}.</p>
     */
    private static int tryMatchLocalRef(String formula, int i, int n,
                                        String src, String dst, StringBuilder out) {
        char ch = formula.charAt(i);
        if (!isLetter(ch)) {
            return i;
        }
        int prev = i - 1;
        if (prev >= 0) {
            char prevCh = formula.charAt(prev);
            if (prevCh == '!' || isLetter(prevCh) || isDigit(prevCh) || prevCh == '_') {
                return i;
            }
        }
        // Avanzar mientras sean letras (columna A..XFD)
        int colStart = i;
        int j = i;
        while (j < n && isLetter(formula.charAt(j))) j++;
        // No puede haber '$' entre la columna y el numero (eso seria fila absoluta).
        if (j >= n || formula.charAt(j) == '$') {
            return i;
        }
        // Leer digitos del numero de fila.
        int numStart = j;
        while (j < n && isDigit(formula.charAt(j))) j++;
        int numEnd = j;
        if (numEnd <= numStart) {
            return i;
        }
        String num = formula.substring(numStart, numEnd);
        if (!num.equals(src)) {
            return i;
        }
        // Match: emitir col + dst y devolver la posicion siguiente al numero.
        out.append(formula, colStart, numStart);
        out.append(dst);
        return numEnd;
    }

    private static boolean isLetter(char c) {
        return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
    }

    private static boolean isDigit(char c) {
        return c >= '0' && c <= '9';
    }

    // ------------------------------------------------------------------
    //  Test hooks: package-private, solo para tests unitarios
    // ------------------------------------------------------------------

    /**
     * Hook para {@code EmptyRowFilterTest}: empaqueta la decision "vale 0"
     * para una unica celda no-formula sin depender de FormulaEvaluator.
     * Las celdas de tipo FORMULA no son evaluables por esta via (devuelve
     * {@code false} para ellas).
     */
    static boolean cellIsZeroForTest(Cell cell) {
        return cell == null
                || (cell.getCellType() != CellType.FORMULA && cellIsZero(cell, null));
    }

    /** Lista inmutable, estable. Expuesta para tests. */
    static List<String> emptyRowColumnNames() {
        return EMPTY_ROW_COLUMN_NAMES;
    }
}

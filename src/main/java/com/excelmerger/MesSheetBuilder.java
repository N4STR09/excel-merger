package com.excelmerger;

import com.excelmerger.sheet.column.FormulaColumnStrategy;
import com.excelmerger.sheet.column.MesColumnStrategy;
import com.excelmerger.sheet.column.MesColumnStrategyFactory;
import com.excelmerger.sheet.column.VlookupLink;
import com.excelmerger.util.PoiUtils;
import com.excelmerger.util.StyleFactory;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Construye la hoja "MES" orquestando una lista de
 * {@link MesColumnStrategy}. Cada estrategia encapsula la logica de un tipo
 * concreto ({@code COPY}, {@code SUMIFS}, {@code FORMULA}, {@code EMPTY}).
 *
 * <p>Las filas de MES se generan por cada fila de la hoja origen que tenga
 * valor en la "columna ancla" (por defecto, Peticion del perfil Cierre —
 * v2.0.0; antes perfil Extraccion).</p>
 *
 * <p>Responsabilidades restantes de este orquestador:</p>
 * <ul>
 *   <li>Lee la config y crea la lista de strategies via
 *       {@link MesColumnStrategyFactory}.</li>
 *   <li>Mantiene el mapa nombre-&gt;indice de columnas
 *       (<b>case-sensitive</b> a proposito; ver CHANGELOG 1.3.0).</li>
 *   <li>Aplica el formato condicional "verde si &gt;= 0" a columnas marcadas.</li>
 *   <li>Detecta apps sin mapeo en formulas VLOOKUP contra hojas de lookup
 *       registradas en el {@link RunReport}.</li>
 * </ul>
 */
public class MesSheetBuilder {

    private static final Logger log = LoggerFactory.getLogger(MesSheetBuilder.class);

    /**
     * Detecta patrones tipo:
     *   VLOOKUP({col:Aplicación},Equipos!$A:$B,...)
     *   VLOOKUP({col:X}, 'Hoja con espacios'!$A:$B,...)
     */
    private static final Pattern VLOOKUP_PATTERN = Pattern.compile(
            "VLOOKUP\\s*\\(\\s*\\{col:([^}]+)\\}\\s*,\\s*(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))\\s*!",
            Pattern.CASE_INSENSITIVE);

    private final ConfigLoader config;
    private final RunReport report;

    public MesSheetBuilder(ConfigLoader config, RunReport report) {
        this.config = config;
        this.report = report;
    }

    public void build(Workbook workbook) {
        if (!config.getBoolean("mes.enabled", false)) {
            log.info("Deshabilitado (mes.enabled=false).");
            return;
        }

        String mesName = config.get("mes.sheetName", "MES");
        // v2.0.0: default ajustado a "Cierre" (antes "Extraccion") tras el
        // swap de nombres de perfil.
        String sourceName = config.get("mes.sourceSheet", "Cierre");
        int sourceHeaderRow0 = config.getInt("mes.sourceHeaderRow", 1) - 1;
        String anchorHeader = config.get("mes.anchorColumn", "Peticion");

        Sheet source = workbook.getSheet(sourceName);
        if (source == null) {
            log.info("La hoja origen '" + sourceName + "' no existe. MES omitida.");
            report.addWarning("HOJA",
                    "La hoja origen '" + sourceName + "' no existe. MES omitida.");
            return;
        }
        if (workbook.getSheet(mesName) != null) {
            log.info("Ya existe una hoja '" + mesName + "'. MES omitida.");
            report.addWarning("HOJA",
                    "Ya existe una hoja '" + mesName + "'. MES omitida.");
            return;
        }

        Row srcHeader = source.getRow(sourceHeaderRow0);
        if (srcHeader == null) {
            log.info("No se encuentran cabeceras en '" + sourceName + "'. MES omitida.");
            report.addWarning("CABECERA",
                    "No se encuentran cabeceras en '" + sourceName + "'. MES omitida.");
            return;
        }

        int anchorCol0 = PoiUtils.findColumnIndex(srcHeader, anchorHeader);
        if (anchorCol0 < 0) {
            log.warn("No se encuentra la columna ancla '{}' en '{}'. MES omitida.",
                    anchorHeader, sourceName);
            report.addWarning("CABECERA",
                    "Columna ancla '" + anchorHeader + "' no encontrada en '"
                            + sourceName + "'. MES omitida.");
            return;
        }

        List<MesColumnStrategy> columns = loadColumns();
        if (columns.isEmpty()) {
            log.info("No hay columnas definidas (mes.col.N.*). MES omitida.");
            report.addWarning("CONFIG",
                    "No hay columnas MES definidas (mes.col.N.*). MES omitida.");
            return;
        }

        Sheet mes = workbook.createSheet(mesName);
        log.info("Creando '" + mesName + "' con " + columns.size() + " columnas.");

        // Mapa nombre -> indice de columna, para resolver placeholders {col:...}
        // CASE-SENSITIVE: permite distinguir columnas que solo difieren en
        // mayusculas/minusculas (p. ej. "Real" vs "REAL"). Feature consciente;
        // ver CHANGELOG 1.3.0 "Decisiones conservadas".
        Map<String, Integer> colByName = new LinkedHashMap<>();
        for (int i = 0; i < columns.size(); i++) {
            colByName.put(columns.get(i).getName(), i);
        }

        // Pre-validar cada columna (avisa UNA vez por problema y marca disabled)
        for (MesColumnStrategy col : columns) {
            col.preValidate(source, sourceHeaderRow0, workbook, colByName, report);
        }

        // Extraer enlaces VLOOKUP para detectar apps sin mapeo
        List<VlookupLink> vlookupLinks = extractVlookupLinks(columns, colByName);

        // --- Cabeceras ---
        Row mesHeader = mes.createRow(0);
        CellStyle headerStyle = StyleFactory.header(workbook);
        for (int i = 0; i < columns.size(); i++) {
            Cell cell = mesHeader.createCell(i);
            cell.setCellValue(columns.get(i).getName());
            cell.setCellStyle(headerStyle);
        }

        // --- Datos: primero recolectamos todas las fuentes de fila (del
        //     perfil Cierre y, opcionalmente, huerfanos del perfil Extraccion),
        //     luego las ordenamos, luego las escribimos. Dos pasadas son
        //     necesarias porque la ordenacion afecta el numero de fila Excel
        //     que cada formula SUMIFS emite como referencia local.
        //     v2.0.0: swap de nombres de perfil respecto a 1.x.
        //     ---
        int sourceLastRow = source.getLastRowNum();

        List<RowSource> rowSources = new ArrayList<>();
        for (int srcR = sourceHeaderRow0 + 1; srcR <= sourceLastRow; srcR++) {
            Row srcRow = source.getRow(srcR);
            if (srcRow == null) continue;
            Cell anchorCell = srcRow.getCell(anchorCol0);
            if (PoiUtils.isBlank(anchorCell)) continue;
            rowSources.add(RowSource.ofCierre(srcRow, srcR + 1));
        }

        boolean orphansEnabled = config.getBoolean("mes.orphans.enabled", false);
        if (orphansEnabled) {
            List<RowSource> orphans = collectOrphans(workbook, source, sourceHeaderRow0);
            rowSources.addAll(orphans);
            sortRowSources(rowSources);
        }

        // Cache de estilos de fill por nombre de color, reutilizable en todas
        // las filas. POI no permite compartir CellStyle entre workbooks; dentro
        // del mismo workbook sí, por eso cacheamos aquí.
        Map<String, CellStyle> fillStyleCache = new LinkedHashMap<>();

        int mesRowIdx = 0;
        for (RowSource rs : rowSources) {
            mesRowIdx++;
            Row mesRow = mes.createRow(mesRowIdx);
            int mesExcelRow = mesRowIdx + 1;

            for (int c = 0; c < columns.size(); c++) {
                MesColumnStrategy col = columns.get(c);
                Cell target = mesRow.createCell(c);
                if (rs.isOrphan()) {
                    writeOrphanCell(target, col, rs, workbook, colByName, mesExcelRow);
                } else {
                    col.writeCell(target, rs.srcRow, source, sourceHeaderRow0,
                            workbook, rs.sourceExcelRow, colByName, mesExcelRow);
                }

                // Fill permanente configurado con mes.col.N.fill=...
                String fillName = col.getFillColor();
                if (fillName != null) {
                    CellStyle fillStyle = fillStyleCache.computeIfAbsent(
                            fillName, k -> buildFillStyle(workbook, k));
                    if (fillStyle != null) {
                        target.setCellStyle(fillStyle);
                    }
                }
            }
        }

        // --- Formato condicional: verde si >= 0 ---
        applyConditionalFormatting(mes, columns, mesRowIdx);

        // --- Formato condicional: rojo si valor != otra columna ---
        applyRedIfNotEqualTo(mes, columns, colByName, mesRowIdx);

        // --- Detectar apps sin mapeo en VLOOKUP ---
        detectUnmappedVlookupKeys(mes, vlookupLinks, mesRowIdx);

        int totalRows = mesRowIdx + 1; // + cabecera
        report.addSheet(mesName, totalRows);
        log.info("{} filas generadas.", mesRowIdx);
    }

    private List<MesColumnStrategy> loadColumns() {
        List<MesColumnStrategy> list = new ArrayList<>();
        int i = 1;
        while (true) {
            String name = config.get("mes.col." + i + ".name", null);
            if (name == null || name.trim().isEmpty()) break;
            String type = config.get("mes.col." + i + ".type", "EMPTY").toUpperCase();
            list.add(MesColumnStrategyFactory.fromConfig(config, i, name.trim(), type, report));
            i++;
        }
        return list;
    }

    /**
     * Aplica una regla de formato condicional a cada columna marcada con
     * greenIfPositive=true: fondo verde claro si la celda es >= 0.
     */
    private void applyConditionalFormatting(Sheet mes, List<MesColumnStrategy> columns, int lastDataRow) {
        if (lastDataRow < 1) return;
        SheetConditionalFormatting scf = mes.getSheetConditionalFormatting();

        for (int i = 0; i < columns.size(); i++) {
            MesColumnStrategy col = columns.get(i);
            if (!col.isGreenIfPositive()) continue;

            ConditionalFormattingRule rule =
                    scf.createConditionalFormattingRule(ComparisonOperator.GE, "0");
            PatternFormatting pattern = rule.createPatternFormatting();
            pattern.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            pattern.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

            String colLetter = CellReference.convertNumToColString(i);
            CellRangeAddress[] range = {
                    CellRangeAddress.valueOf(colLetter + "2:" + colLetter + (lastDataRow + 1))
            };
            scf.addConditionalFormatting(range, rule);
            log.info("Formato condicional 'verde si >=0' aplicado a columna '" + col.getName() + "'.");
        }
    }

    /**
     * Construye un {@link CellStyle} con fill solido de uno de los nombres
     * de color soportados para {@code mes.col.N.fill=...}. Si el nombre no
     * se reconoce, registra un warning y devuelve {@code null}.
     *
     * <p>Colores soportados (alineados con IndexedColors):</p>
     * <ul>
     *   <li>{@code LIGHT_GREEN}  ({@code #E2EFDA})</li>
     *   <li>{@code MEDIUM_GREEN} ({@code #C6E0B4})</li>
     *   <li>{@code LIGHT_BLUE}   ({@code #DDEBF7})</li>
     *   <li>{@code LIGHT_YELLOW} ({@code #FFF2CC})</li>
     *   <li>{@code LIGHT_RED}    ({@code #FCE4D6})</li>
     *   <li>{@code LIGHT_LAVENDER} ({@code #E4DFEC})</li>
     * </ul>
     */
    private CellStyle buildFillStyle(Workbook workbook, String colorName) {
        String hex = resolveFillHex(colorName);
        if (hex == null) {
            report.addWarning("CONFIG",
                    "Color de fill desconocido '" + colorName
                            + "'. Valores permitidos: LIGHT_GREEN, MEDIUM_GREEN, "
                            + "LIGHT_BLUE, LIGHT_YELLOW, LIGHT_RED, LIGHT_LAVENDER.");
            return null;
        }
        return StyleFactory.solidFill(workbook, hex);
    }

    private static String resolveFillHex(String colorName) {
        if (colorName == null) return null;
        switch (colorName.trim().toUpperCase()) {
            case "LIGHT_GREEN":    return "FFE2EFDA";
            case "MEDIUM_GREEN":   return "FFC6E0B4";
            case "LIGHT_BLUE":     return "FFDDEBF7";
            case "LIGHT_YELLOW":   return "FFFFF2CC";
            case "LIGHT_RED":      return "FFFCE4D6";
            case "LIGHT_LAVENDER": return "FFE4DFEC";
            default:               return null;
        }
    }

    /**
     * Aplica un formato condicional "rojo claro si la celda difiere de otra
     * columna" a cada columna marcada con {@code mes.col.N.redIfNotEqualTo=X}.
     * La comparacion es por celda, fila a fila, usando referencias relativas
     * en la formula del CF.
     */
    private void applyRedIfNotEqualTo(Sheet mes, List<MesColumnStrategy> columns,
                                      Map<String, Integer> colByName, int lastDataRow) {
        if (lastDataRow < 1) return;
        SheetConditionalFormatting scf = mes.getSheetConditionalFormatting();

        for (int i = 0; i < columns.size(); i++) {
            MesColumnStrategy col = columns.get(i);
            String refName = col.getRedIfNotEqualTo();
            if (refName == null) continue;

            Integer refIdx = colByName.get(refName);
            if (refIdx == null) {
                report.addWarning("CONFIG",
                        "Columna '" + col.getName() + "' tiene redIfNotEqualTo='" + refName
                                + "' pero esa columna no existe en MES. Regla omitida.");
                continue;
            }

            String thisLetter = CellReference.convertNumToColString(i);
            String refLetter  = CellReference.convertNumToColString(refIdx);

            // Formula relativa: compara la celda actual con la celda homologa de
            // la columna referenciada en la misma fila (primera fila de datos).
            // POI traslada la referencia relativa al resto del rango automaticamente.
            String formula = thisLetter + "2<>" + refLetter + "2";

            ConditionalFormattingRule rule = scf.createConditionalFormattingRule(formula);
            PatternFormatting pattern = rule.createPatternFormatting();
            // IndexedColors.CORAL es un rojo-naranja saturado (aprox. #FF8080) que
            // se ve claramente mas rojo que ROSE (mas rosa) y se lee bien como
            // fondo sin ser agresivo. Se usa indexado para no depender de XSSF
            // en el codigo de CF (mas simple y portable).
            pattern.setFillBackgroundColor(IndexedColors.CORAL.getIndex());
            pattern.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

            CellRangeAddress[] range = {
                    CellRangeAddress.valueOf(thisLetter + "2:" + thisLetter + (lastDataRow + 1))
            };
            scf.addConditionalFormatting(range, rule);
            log.info("Formato condicional 'rojo si != {}' aplicado a columna '{}'.",
                    refName, col.getName());
        }
    }

    /**
     * Extrae todos los enlaces {@code VLOOKUP({col:X}, Hoja!...)} de las
     * formulas de las columnas FORMULA. Solo guarda enlaces cuya hoja destino
     * sea un lookup registrado en el {@link RunReport}.
     */
    private List<VlookupLink> extractVlookupLinks(List<MesColumnStrategy> columns,
                                                  Map<String, Integer> colByName) {
        List<VlookupLink> links = new ArrayList<>();
        for (MesColumnStrategy col : columns) {
            if (!(col instanceof FormulaColumnStrategy)) continue;
            String template = col.formulaTemplate().orElse(null);
            if (template == null) continue;
            Matcher m = VLOOKUP_PATTERN.matcher(template);
            while (m.find()) {
                String sourceColName = m.group(1).trim();
                String lookupSheet = m.group(2) != null ? m.group(2) : m.group(3);
                Integer sourceColIdx = colByName.get(sourceColName);
                if (sourceColIdx == null) continue; // ya se avisa por preValidate
                if (!report.hasLookup(lookupSheet)) continue;
                links.add(new VlookupLink(sourceColName, sourceColIdx, lookupSheet));
            }
        }
        return links;
    }

    /**
     * Recorre todas las filas de MES comparando el valor de cada "columna key"
     * contra las claves de su lookup asociado. Agrega un unico warning por
     * lookupSheet con el listado de valores sin mapeo (hasta 30, despues corta).
     */
    private void detectUnmappedVlookupKeys(Sheet mes, List<VlookupLink> links, int lastDataRow) {
        if (links.isEmpty() || lastDataRow < 1) return;

        Map<String, Set<String>> missing = new LinkedHashMap<>();
        Map<String, String> keyColByLookup = new LinkedHashMap<>();

        for (int r = 1; r <= lastDataRow; r++) {
            Row row = mes.getRow(r);
            if (row == null) continue;
            for (VlookupLink link : links) {
                Set<String> lookupKeys = report.getLookupKeys(link.lookupSheet);
                if (lookupKeys == null) continue;
                Cell keyCell = row.getCell(link.sourceColIdx);
                String value = PoiUtils.cellAsString(keyCell);
                if (value == null || value.trim().isEmpty()) continue;
                value = value.trim();
                // Saltar el sentinela "-" que se usa en filas huerfanas para
                // marcar "sin valor". No es una app real sin mapeo. (v1.7.0)
                if ("-".equals(value)) continue;
                if (!containsIgnoreCase(lookupKeys, value)) {
                    missing.computeIfAbsent(link.lookupSheet, k -> new LinkedHashSet<>()).add(value);
                    keyColByLookup.putIfAbsent(link.lookupSheet, link.sourceColName);
                }
            }
        }

        for (Map.Entry<String, Set<String>> e : missing.entrySet()) {
            Set<String> values = e.getValue();
            String sample = sampleOf(values, 30);
            String msg = values.size() + " valor(es) de '" + keyColByLookup.get(e.getKey())
                    + "' sin mapeo en '" + e.getKey() + "': " + sample;
            log.warn("{}", msg);
            report.addWarning("LOOKUP", msg);
        }
    }

    private static boolean containsIgnoreCase(Set<String> set, String value) {
        for (String s : set) {
            if (s.equalsIgnoreCase(value)) return true;
        }
        return false;
    }

    private static String sampleOf(Set<String> values, int max) {
        StringBuilder sb = new StringBuilder();
        int i = 0;
        for (String v : values) {
            if (i >= max) { sb.append(", ..."); break; }
            if (i > 0) sb.append(", ");
            sb.append(v);
            i++;
        }
        return sb.toString();
    }

    // =============================================================
    //  Huerfanos (v1.7.0) — filas de Resultado para imputaciones del
    //  perfil Extraccion cuya (Component Name, Matricula) no tiene fila
    //  equivalente en Cierre.(Peticion, Recurso). Ver CHANGELOG 1.7.0.
    //  v2.0.0: swap de nombres de perfil.
    // =============================================================

    /**
     * Fuente de una fila de Resultado. Dos tipos: una fila real del
     * perfil {@code Cierre} (con su {@code srcRow} y su numero Excel en
     * origen), o una fila huerfana derivada de agregar imputaciones del
     * perfil {@code Extraccion} para un par {@code (Component Name,
     * Matricula)} sin contrapartida.
     *
     * <p>Se usa una clase pequenya con flag {@code orphan} en lugar de
     * una jerarquia {@code sealed} para mantener el diff minimo y evitar
     * tocar el resto del package.</p>
     */
    private static final class RowSource {
        final boolean orphan;
        // Solo para non-orphan:
        final Row srcRow;
        final int sourceExcelRow;
        // Solo para orphan:
        final String orphanPeticion;
        final String orphanMatricula;
        final double orphanHours;

        private RowSource(boolean orphan, Row srcRow, int sourceExcelRow,
                          String orphanPeticion, String orphanMatricula, double orphanHours) {
            this.orphan = orphan;
            this.srcRow = srcRow;
            this.sourceExcelRow = sourceExcelRow;
            this.orphanPeticion = orphanPeticion;
            this.orphanMatricula = orphanMatricula;
            this.orphanHours = orphanHours;
        }

        static RowSource ofCierre(Row srcRow, int sourceExcelRow) {
            return new RowSource(false, srcRow, sourceExcelRow, null, null, 0.0);
        }

        static RowSource ofOrphan(String peticion, String matricula, double hours) {
            return new RowSource(true, null, -1, peticion, matricula, hours);
        }

        boolean isOrphan() { return orphan; }
    }

    /**
     * Recolecta huerfanos: imputaciones de la hoja configurada en
     * {@code mes.orphans.sourceSheet} cuya {@code (Component Name, Matricula)}
     * no existe como {@code (Peticion, Recurso)} en la hoja {@code source}.
     * Las imputaciones se agrupan por pareja {@code (CN, Mat)} y las horas
     * se suman.
     *
     * <p>Si la hoja de huerfanos no existe o no tiene cabeceras
     * identificables, se emite un warning y se devuelve una lista vacia
     * (opt-in permisivo).</p>
     */
    private List<RowSource> collectOrphans(Workbook workbook, Sheet source,
                                           int sourceHeaderRow0) {
        // v2.0.0: default ajustado a "Extraccion" (antes "Cierre") tras el
        // swap de nombres de perfil.
        String orphanSheetName = config.get("mes.orphans.sourceSheet", "Extraccion");
        String matchCN  = config.get("mes.orphans.matchComponent", "Component Name");
        String matchMat = config.get("mes.orphans.matchMatricula", "Matricula");
        String sumCol   = config.get("mes.orphans.sumColumn", "Hours");

        Sheet orphanSheet = workbook.getSheet(orphanSheetName);
        if (orphanSheet == null) {
            report.addWarning("HOJA",
                    "Hoja de huerfanos '" + orphanSheetName + "' no existe; "
                            + "la seccion de huerfanos se omite.");
            return new ArrayList<>();
        }
        int orphanHeaderRow0 = PoiUtils.detectHeaderRow(orphanSheet);
        Row orphanHeader = orphanSheet.getRow(orphanHeaderRow0);
        if (orphanHeader == null) {
            report.addWarning("CABECERA",
                    "No se encontraron cabeceras en '" + orphanSheetName
                            + "'; la seccion de huerfanos se omite.");
            return new ArrayList<>();
        }
        int cnIdx  = PoiUtils.findColumnIndex(orphanHeader, matchCN);
        int matIdx = PoiUtils.findColumnIndex(orphanHeader, matchMat);
        int hIdx   = PoiUtils.findColumnIndex(orphanHeader, sumCol);
        if (cnIdx < 0 || matIdx < 0 || hIdx < 0) {
            report.addWarning("CABECERA",
                    "Faltan cabeceras en '" + orphanSheetName + "' para calcular"
                            + " huerfanos (buscadas: '" + matchCN + "', '" + matchMat
                            + "', '" + sumCol + "'); la seccion se omite.");
            return new ArrayList<>();
        }

        // Set (Peticion, Recurso) existentes en el perfil Cierre (v2.0.0;
        // antes Extraccion), normalizados como STRING (coherente con el
        // fix 1.6.2: asText.columns=Peticion,Recurso).
        Set<String> extKeys = loadExtractionPairKeys(source, sourceHeaderRow0);

        // Agrupar por (CN, Mat) sumando Hours.
        Map<String, Double> hoursByPair = new LinkedHashMap<>();
        int orphanLastRow = orphanSheet.getLastRowNum();
        for (int r = orphanHeaderRow0 + 1; r <= orphanLastRow; r++) {
            Row row = orphanSheet.getRow(r);
            if (row == null) continue;
            String cn  = cellAsPlainString(row.getCell(cnIdx));
            String mat = cellAsPlainString(row.getCell(matIdx));
            if (cn == null || cn.isEmpty()) continue;
            String matKey = mat == null ? "" : mat;
            String pairKey = pairKey(cn, matKey);
            if (extKeys.contains(pairKey)) continue; // no es huerfano
            double hours = cellAsDouble(row.getCell(hIdx));
            hoursByPair.merge(pairKey, hours, Double::sum);
        }

        // Convertir el mapa a lista de RowSource. Ordenamos primero por CN
        // y despues por Matricula para que al hacer el sort final (numerico
        // ASC + no-numerico al final) sea estable dentro de cada CN.
        TreeMap<String, Double> sorted = new TreeMap<>(hoursByPair);
        List<RowSource> out = new ArrayList<>();
        for (Map.Entry<String, Double> e : sorted.entrySet()) {
            String[] parts = unpairKey(e.getKey());
            out.add(RowSource.ofOrphan(parts[0], parts[1], e.getValue()));
        }
        log.info("Huerfanos detectados: {} filas (CN, Matricula) sin contrapartida en '{}'.",
                out.size(), source.getSheetName());
        return out;
    }

    /**
     * Construye el set de claves {@code "pet|rec"} a partir de las filas
     * del perfil {@code Cierre} (v2.0.0; antes perfil {@code Extraccion}).
     * Coherente con la normalizacion del fix 1.6.2: las celdas numericas
     * enteras se serializan sin decimales.
     */
    private Set<String> loadExtractionPairKeys(Sheet source, int sourceHeaderRow0) {
        Row header = source.getRow(sourceHeaderRow0);
        if (header == null) return new LinkedHashSet<>();
        String peticionHeader = config.get("mes.anchorColumn", "Peticion");
        // El nombre "Recurso" viene implicitamente del SUMIFS de Jira
        // (match=Component Name:Peticion,Matricula:Recurso,...). Para no
        // introducir otra clave de config, buscamos la columna "Recurso"
        // por nombre literal en la hoja origen. Si no existe, devolvemos
        // set vacio: todos los Component Name seran huerfanos, cosa que
        // el usuario notara en el resultado.
        int petIdx = PoiUtils.findColumnIndex(header, peticionHeader);
        int recIdx = PoiUtils.findColumnIndex(header, "Recurso");
        Set<String> keys = new LinkedHashSet<>();
        if (petIdx < 0 || recIdx < 0) {
            report.addWarning("CABECERA",
                    "No se pudieron localizar columnas 'Peticion'/'Recurso' en '"
                            + source.getSheetName() + "' para calcular huerfanos.");
            return keys;
        }
        int last = source.getLastRowNum();
        for (int r = sourceHeaderRow0 + 1; r <= last; r++) {
            Row row = source.getRow(r);
            if (row == null) continue;
            String pet = cellAsPlainString(row.getCell(petIdx));
            if (pet == null || pet.isEmpty()) continue;
            String rec = cellAsPlainString(row.getCell(recIdx));
            if (rec == null) rec = "";
            keys.add(pairKey(pet, rec));
        }
        return keys;
    }

    /**
     * Escribe la celda {@code target} para una fila de tipo huerfano. La logica
     * se basa en el <b>nombre de columna</b>, no en el tipo de estrategia:
     * las columnas "clave" (Peticion, Matricula) reciben los valores del
     * huerfano, la columna "Jira" recibe las horas sumadas, el resto de
     * columnas FORMULA se evaluan con sus plantillas originales (funcionan
     * igual al resolver {@code col:X} contra la misma fila MES), y todas las
     * demas columnas (COPY de otras cosas, EMPTY) reciben un literal
     * {@code "-"} para cumplir el contrato del usuario.
     */
    private void writeOrphanCell(Cell target, MesColumnStrategy col, RowSource rs,
                                 Workbook workbook, Map<String, Integer> colByName,
                                 int mesExcelRow) {
        String colPeticion  = config.get("mes.orphans.colPeticion",  "Petición");
        String colMatricula = config.get("mes.orphans.colMatricula", "Matrícula");
        String colJira      = config.get("mes.orphans.colJira",      "Jira");

        String name = col.getName();
        if (name.equals(colPeticion)) {
            target.setCellValue(rs.orphanPeticion);
            return;
        }
        if (name.equals(colMatricula)) {
            target.setCellValue(rs.orphanMatricula);
            return;
        }
        if (name.equals(colJira)) {
            target.setCellValue(rs.orphanHours);
            return;
        }
        if (col instanceof FormulaColumnStrategy) {
            // Para columnas FORMULA (REAL, PDCL, PDCL+Deuda, Equipo)
            // dejamos que la estrategia resuelva la formula normalmente:
            // {col:Jira}*1.2 se convertira en p.ej. "D16*1.2" que
            // referencia la celda Jira de esta misma fila huerfana. La
            // implementacion de FormulaColumnStrategy.doWriteCell no lee
            // srcRow/source, solo necesita workbook, colByName y mesExcelRow.
            col.writeCell(target, null, null, -1, workbook,
                    -1, colByName, mesExcelRow);
            return;
        }
        // Resto (COPY de otras columnas, EMPTY, SUMIFS que no es Jira): literal "-".
        target.setCellValue("-");
    }

    /**
     * Ordena {@code rowSources} in-place: primero las peticiones numericas
     * (parseables como Long) en orden ascendente, despues las no numericas
     * en orden alfabetico. Dentro de un mismo Peticion mantiene orden por
     * Matricula (tambien numericas ASC primero, el resto alfabeticas).
     *
     * <p>La clave "numerica" se obtiene de la celda Peticion: para filas
     * del perfil Cierre, leyendo la celda ancla; para huerfanos,
     * directamente {@code orphanPeticion}.</p>
     */
    private void sortRowSources(List<RowSource> rowSources) {
        String anchorHeader = config.get("mes.anchorColumn", "Peticion");
        // Para RowSource no-huerfano, necesitamos saber en que columna
        // de srcRow esta la peticion. La hoja origen tiene un header; lo
        // resolvemos una sola vez.
        Sheet anySource = null;
        for (RowSource rs : rowSources) {
            if (!rs.orphan && rs.srcRow != null) {
                anySource = rs.srcRow.getSheet();
                break;
            }
        }
        int sourceHeaderRow0 = config.getInt("mes.sourceHeaderRow", 1) - 1;
        final int anchorCol0 = anySource == null
                ? -1
                : PoiUtils.findColumnIndex(anySource.getRow(sourceHeaderRow0), anchorHeader);

        Comparator<RowSource> byPeticion = Comparator.comparing(
                (RowSource rs) -> extractPeticionKey(rs, anchorCol0),
                Comparator.comparing((SortKey k) -> !k.numeric)  // numericos primero
                        .thenComparingLong(k -> k.numericValue)
                        .thenComparing(k -> k.textValue));
        rowSources.sort(byPeticion);
    }

    private SortKey extractPeticionKey(RowSource rs, int anchorCol0) {
        String raw;
        if (rs.orphan) {
            raw = rs.orphanPeticion;
        } else if (anchorCol0 < 0) {
            raw = "";
        } else {
            raw = cellAsPlainString(rs.srcRow.getCell(anchorCol0));
        }
        return SortKey.of(raw == null ? "" : raw);
    }

    /** Clave de ordenacion: primero numericos ASC, luego textuales A-Z. */
    private static final class SortKey {
        final boolean numeric;
        final long numericValue;
        final String textValue;

        private SortKey(boolean numeric, long numericValue, String textValue) {
            this.numeric = numeric;
            this.numericValue = numericValue;
            this.textValue = textValue;
        }

        static SortKey of(String s) {
            try {
                return new SortKey(true, Long.parseLong(s), s);
            } catch (NumberFormatException e) {
                return new SortKey(false, 0L, s);
            }
        }
    }

    // --- Helpers de celda ---

    /** Devuelve la celda como string "plano" (sin decimales si es entero). */
    private static String cellAsPlainString(Cell cell) {
        if (cell == null) return null;
        if (cell.getCellType() == CellType.NUMERIC) {
            double d = cell.getNumericCellValue();
            return (d == (long) d) ? String.valueOf((long) d) : String.valueOf(d);
        }
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }
        return PoiUtils.cellAsString(cell);
    }

    private static double cellAsDouble(Cell cell) {
        if (cell == null) return 0.0;
        if (cell.getCellType() == CellType.NUMERIC) return cell.getNumericCellValue();
        if (cell.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(cell.getStringCellValue().trim());
            } catch (NumberFormatException e) {
                return 0.0;
            }
        }
        return 0.0;
    }

    private static String pairKey(String pet, String rec) {
        // Separador '\u0001' para evitar colisiones con valores reales.
        return pet + "\u0001" + rec;
    }

    private static String[] unpairKey(String key) {
        int i = key.indexOf('\u0001');
        if (i < 0) return new String[] {key, ""};
        return new String[] {key.substring(0, i), key.substring(i + 1)};
    }

    // =============================================================
    //  VlookupLink vive en com.excelmerger.sheet.column.VlookupLink
    //  (extraida en Fase 5 del refactor 1.3.0).
    // =============================================================
}

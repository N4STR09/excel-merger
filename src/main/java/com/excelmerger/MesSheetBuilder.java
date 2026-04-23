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
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Construye la hoja "MES" orquestando una lista de
 * {@link MesColumnStrategy}. Cada estrategia encapsula la logica de un tipo
 * concreto ({@code COPY}, {@code SUMIFS}, {@code FORMULA}, {@code EMPTY}).
 *
 * <p>Las filas de MES se generan por cada fila de la hoja origen que tenga
 * valor en la "columna ancla" (por defecto, Peticion de Extraccion).</p>
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
        String sourceName = config.get("mes.sourceSheet", "Extraccion");
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

        // --- Datos: una fila de MES por cada fila de origen con valor en la ancla ---
        int sourceLastRow = source.getLastRowNum();
        int mesRowIdx = 0;

        // Cache de estilos de fill por nombre de color, reutilizable en todas
        // las filas. POI no permite compartir CellStyle entre workbooks; dentro
        // del mismo workbook sí, por eso cacheamos aquí.
        Map<String, CellStyle> fillStyleCache = new LinkedHashMap<>();

        for (int srcR = sourceHeaderRow0 + 1; srcR <= sourceLastRow; srcR++) {
            Row srcRow = source.getRow(srcR);
            if (srcRow == null) continue;
            Cell anchorCell = srcRow.getCell(anchorCol0);
            if (PoiUtils.isBlank(anchorCell)) continue;

            mesRowIdx++;
            Row mesRow = mes.createRow(mesRowIdx);

            int sourceExcelRow = srcR + 1;   // 1-indexado para las formulas
            int mesExcelRow = mesRowIdx + 1; // idem para referencias internas de MES

            for (int c = 0; c < columns.size(); c++) {
                MesColumnStrategy col = columns.get(c);
                Cell target = mesRow.createCell(c);
                col.writeCell(target, srcRow, source, sourceHeaderRow0,
                        workbook, sourceExcelRow, colByName, mesExcelRow);

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
    //  VlookupLink vive en com.excelmerger.sheet.column.VlookupLink
    //  (extraida en Fase 5 del refactor 1.3.0).
    // =============================================================
}

package com.excelmerger.sheet.column;

import com.excelmerger.RunReport;
import com.excelmerger.util.PoiUtils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;

/**
 * Columna FORMULA_PLUS_SUMIFS: combina una formula base con placeholders
 * {@code {col:X}} y, opcionalmente, la suma de una hoja remota via
 * {@code IFERROR(SUMIFS(...),0)}.
 *
 * <p>Comportamiento:</p>
 * <ul>
 *   <li>Si la hoja {@code fromSheet} existe en el workbook, se escribe
 *       {@code base+IFERROR(SUMIFS(...),0)}. Las letras de columna del
 *       remoto se resuelven dinamicamente leyendo la cabecera (mismo
 *       patron que {@link SumIfsColumnStrategy}).</li>
 *   <li>Si la hoja no existe (escenario: el tercer fichero de Deuda no se
 *       aporto), se escribe <b>solo</b> la formula base. Sin warnings: es un
 *       degradado silencioso por diseno.</li>
 * </ul>
 *
 * <p>Este tipo existe para soportar la columna {@code PDCL + Deuda} de la
 * hoja Resultado, que en v2.2.0 pasa a sumar horas reales de deuda cuando
 * el usuario aporta el fichero opcional correspondiente. Antes de v2.2.0
 * la columna usaba el tipo FORMULA con plantilla {@code {col:PDCL}} y ese
 * comportamiento se mantiene intacto cuando no hay Deuda.</p>
 */
public final class FormulaPlusSumIfsColumnStrategy extends AbstractMesColumnStrategy {

    private final String baseFormulaTemplate;
    private final String fromSheet;
    private final String sumHeader;
    private final List<String[]> matches;

    public FormulaPlusSumIfsColumnStrategy(String name, boolean greenIfPositive,
                                           String fillColor, String redIfNotEqualTo,
                                           String baseFormulaTemplate,
                                           String fromSheet, String sumHeader,
                                           List<String[]> matches) {
        super(name, greenIfPositive, fillColor, redIfNotEqualTo);
        this.baseFormulaTemplate = baseFormulaTemplate;
        this.fromSheet = fromSheet;
        this.sumHeader = sumHeader;
        this.matches = Collections.unmodifiableList(matches);
    }

    @Override
    public Optional<String> formulaTemplate() {
        return Optional.ofNullable(baseFormulaTemplate);
    }

    @Override
    public void preValidate(Sheet source, int sourceHeaderRow0, Workbook workbook,
                            Map<String, Integer> mesColIndexByName, RunReport report) {
        // 1. Validar la formula base (placeholders {col:X}).
        if (baseFormulaTemplate == null) {
            disable();
            return;
        }
        int guardStart = 0;
        String s = baseFormulaTemplate;
        while (true) {
            int start = s.indexOf("{col:", guardStart);
            if (start < 0) break;
            int end = s.indexOf("}", start);
            if (end < 0) break;
            String colName = s.substring(start + 5, end).trim();
            if (!mesColIndexByName.containsKey(colName)) {
                report.addWarning("FORMULA",
                        "Placeholder {col:" + colName + "} en '" + name
                                + "' no coincide con ninguna columna MES.");
                disable();
            }
            guardStart = end + 1;
        }

        // 2. Validar referencias a la hoja remota SOLO si la hoja existe.
        //    Si no existe, la columna sigue activa (degrada a base pura
        //    en runtime). Sin warnings: es la via "sin fichero Deuda".
        Sheet remote = workbook.getSheet(fromSheet);
        if (remote == null) {
            return;
        }
        int remoteHeaderRow0 = PoiUtils.detectHeaderRow(remote);
        if (PoiUtils.columnLetter(remote, remoteHeaderRow0, sumHeader) == null) {
            report.addWarning("CABECERA",
                    "Columna '" + sumHeader + "' no encontrada en '"
                            + fromSheet + "' (columna '" + name + "').");
            disable();
            return;
        }
        for (String[] m : matches) {
            String remoteHeader = m[0];
            String localHeader = m[1];
            if (PoiUtils.columnLetter(remote, remoteHeaderRow0, remoteHeader) == null) {
                report.addWarning("CABECERA",
                        "Columna '" + remoteHeader + "' no encontrada en '"
                                + fromSheet + "' (columna '" + name + "').");
                disable();
            }
            // IMPORTANTE: para FORMULA_PLUS_SUMIFS, localHeader es el NOMBRE
            // de una columna MES (de Resultado), no una cabecera de la hoja
            // origen (Cierre). Validamos contra mesColIndexByName porque la
            // fórmula generada referenciará Resultado!$<letraMES>$<filaMES>,
            // NO la hoja origen.
            if (!mesColIndexByName.containsKey(localHeader)) {
                report.addWarning("CABECERA",
                        "Columna MES '" + localHeader + "' no encontrada en "
                                + "Resultado (columna '" + name + "'). "
                                + "Columnas MES disponibles: "
                                + mesColIndexByName.keySet() + ".");
                disable();
            }
        }
    }

    @Override
    protected void doWriteCell(Cell target, Row srcRow, Sheet source, int sourceHeaderRow0,
                               Workbook workbook, int sourceExcelRow,
                               Map<String, Integer> mesColIndexByName, int mesExcelRow) {
        // 1. Resolver placeholders de la formula base en {col:X} y {colLetter:X}.
        String base = resolvePlaceholders(baseFormulaTemplate, mesColIndexByName, mesExcelRow);
        if (base == null) {
            target.setBlank();
            return;
        }
        if (base.startsWith("=")) base = base.substring(1);

        // 2. Si la hoja remota no existe, escribir solo la base.
        Sheet remote = workbook.getSheet(fromSheet);
        if (remote == null) {
            target.setCellFormula(base);
            return;
        }

        // 3. Construir el SUMIFS resolviendo letras en runtime.
        int remoteHeaderRow0 = PoiUtils.detectHeaderRow(remote);
        String sumCol = PoiUtils.columnLetter(remote, remoteHeaderRow0, sumHeader);
        if (sumCol == null) {
            target.setCellFormula(base);
            return;
        }

        // La hoja MES (Resultado) donde se escribe la celda target. Los
        // criterios locales del SUMIFS apuntan a esta misma fila de MES.
        String mesSheetName = target.getSheet().getSheetName();

        StringBuilder sumifs = new StringBuilder();
        sumifs.append("SUMIFS(").append(PoiUtils.quoteSheetName(fromSheet))
                .append("!$").append(sumCol).append(":$").append(sumCol);
        for (String[] m : matches) {
            String remoteHeader = m[0];
            String localHeader = m[1];
            String remoteCol = PoiUtils.columnLetter(remote, remoteHeaderRow0, remoteHeader);
            Integer localIdx = mesColIndexByName.get(localHeader);
            if (remoteCol == null || localIdx == null) {
                // Defensa final: si algo no se puede resolver, degradar a base.
                target.setCellFormula(base);
                return;
            }
            String localCol = CellReference.convertNumToColString(localIdx);
            sumifs.append(",").append(PoiUtils.quoteSheetName(fromSheet))
                    .append("!$").append(remoteCol).append(":$").append(remoteCol);
            sumifs.append(",").append(PoiUtils.quoteSheetName(mesSheetName))
                    .append("!$").append(localCol).append("$").append(mesExcelRow);
        }
        sumifs.append(")");

        String full = base + "+IFERROR(" + sumifs + ",0)";
        target.setCellFormula(full);
    }

    /**
     * Resuelve placeholders {@code {col:Nombre}} y {@code {colLetter:X}}
     * contra el mapa de columnas MES. Devuelve {@code null} si algun
     * {@code {col:...}} no puede resolverse.
     * Misma semantica que {@link FormulaColumnStrategy}.
     */
    private static String resolvePlaceholders(String template,
                                              Map<String, Integer> mesColIndexByName,
                                              int mesExcelRow) {
        String resolved = template;
        int guard = 0;
        while (resolved.contains("{col:") && guard < 50) {
            guard++;
            int start = resolved.indexOf("{col:");
            int end = resolved.indexOf("}", start);
            if (end < 0) break;
            String colName = resolved.substring(start + 5, end).trim();
            Integer idx = mesColIndexByName.get(colName);
            if (idx == null) {
                return null;
            }
            String ref = CellReference.convertNumToColString(idx) + mesExcelRow;
            resolved = resolved.substring(0, start) + ref + resolved.substring(end + 1);
        }
        guard = 0;
        while (resolved.contains("{colLetter:") && guard < 50) {
            guard++;
            int start = resolved.indexOf("{colLetter:");
            int end = resolved.indexOf("}", start);
            if (end < 0) break;
            String letter = resolved.substring(start + 11, end).trim().toUpperCase(Locale.ROOT);
            String ref = letter + mesExcelRow;
            resolved = resolved.substring(0, start) + ref + resolved.substring(end + 1);
        }
        return resolved;
    }
}

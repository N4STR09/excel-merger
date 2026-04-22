package com.excelmerger.sheet.column;

import com.excelmerger.RunReport;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.util.Map;
import java.util.Optional;

/**
 * Columna FORMULA: resuelve una plantilla con placeholders
 * <ul>
 *   <li>{@code {col:NombreColumna}} -&gt; letra de esa columna MES + fila actual.
 *       <b>Case-sensitive</b> a proposito (permite distinguir {@code Real} vs
 *       {@code REAL}); v. documentacion del proyecto.</li>
 *   <li>{@code {colLetter:X}} -&gt; letra literal {@code X} + fila actual.</li>
 * </ul>
 */
public final class FormulaColumnStrategy extends AbstractMesColumnStrategy {

    private final String formulaTemplate;

    public FormulaColumnStrategy(String name, boolean greenIfPositive, String formulaTemplate) {
        super(name, greenIfPositive);
        this.formulaTemplate = formulaTemplate;
    }

    @Override
    public Optional<String> formulaTemplate() {
        return Optional.of(formulaTemplate);
    }

    @Override
    public void preValidate(Sheet source, int sourceHeaderRow0, Workbook workbook,
                            Map<String, Integer> mesColIndexByName, RunReport report) {
        if (formulaTemplate == null) { disable(); return; }
        int guardStart = 0;
        String s = formulaTemplate;
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
    }

    @Override
    protected void doWriteCell(Cell target, Row srcRow, Sheet source, int sourceHeaderRow0,
                               Workbook workbook, int sourceExcelRow,
                               Map<String, Integer> mesColIndexByName, int mesExcelRow) {
        String resolved = formulaTemplate;

        // Reemplazar {col:Nombre}. La validez ya se comprobo en preValidate.
        // NOTA: lookup case-sensitive (feature consciente; ver CHANGELOG).
        int guard = 0;
        while (resolved.contains("{col:") && guard++ < 50) {
            int start = resolved.indexOf("{col:");
            int end = resolved.indexOf("}", start);
            if (end < 0) break;
            String colName = resolved.substring(start + 5, end).trim();
            Integer idx = mesColIndexByName.get(colName);
            if (idx == null) {
                target.setBlank();
                return;
            }
            String ref = CellReference.convertNumToColString(idx) + mesExcelRow;
            resolved = resolved.substring(0, start) + ref + resolved.substring(end + 1);
        }

        // Reemplazar {colLetter:X}
        guard = 0;
        while (resolved.contains("{colLetter:") && guard++ < 50) {
            int start = resolved.indexOf("{colLetter:");
            int end = resolved.indexOf("}", start);
            if (end < 0) break;
            String letter = resolved.substring(start + 11, end).trim().toUpperCase();
            String ref = letter + mesExcelRow;
            resolved = resolved.substring(0, start) + ref + resolved.substring(end + 1);
        }

        if (resolved.startsWith("=")) resolved = resolved.substring(1);
        target.setCellFormula(resolved);
    }
}

package com.excelmerger;

import com.excelmerger.util.PoiUtils;
import com.excelmerger.util.StyleFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.List;

/**
 * v2.4.0 — Helper que escribe UNA tabla pivot SUMIFS Petición × Matrícula
 * en una hoja de responsable. Se invoca dos veces por hoja: una para Jira,
 * otra para REAL.
 *
 * <p>La tabla pivot tiene esta forma (los rangos son ilustrativos):</p>
 * <pre>
 *   Fila titleRow0:    [Titulo merged sobre todas las columnas]
 *   Fila titleRow0+1:  [cabecera] "Petición" | M-1001 | M-1002 | ... | Total
 *   Fila titleRow0+2:  P-001     | SUMIFS  | SUMIFS  | ... | SUM(fila)
 *   ...                ...        | ...     | ...     | ... | ...
 *   Fila ultima:       Total     | SUM     | SUM     | ... | SUM(grand)
 * </pre>
 *
 * <p><b>Forma de las fórmulas SUMIFS</b> (3 criterios):</p>
 * <pre>
 *   SUMIFS(
 *     Resultado!&lt;valLetter&gt;2:&lt;valLetter&gt;maxRow,    // sum_range (Jira o REAL)
 *     Resultado!&lt;peticionLetter&gt;2:&lt;peticionLetter&gt;maxRow,  $A&lt;dataRow&gt;,
 *     Resultado!&lt;matrLetter&gt;2:&lt;matrLetter&gt;maxRow,           &lt;matrColLetter&gt;$&lt;headerRow&gt;,
 *     Resultado!&lt;respLetter&gt;2:&lt;respLetter&gt;maxRow,           $A$1
 *   )
 * </pre>
 *
 * <p>El criterio del responsable es <b>siempre {@code $A$1}</b> de la hoja
 * destino. Eso evita problemas de escapado de caracteres especiales
 * ({@code '}, {@code "}, {@code @}) en el nombre del responsable y permite
 * que un usuario edite A1 sin romper las fórmulas.</p>
 *
 * <p><b>Lección 1.7.1 (mismatch numerico/textual)</b>: las cabeceras de
 * matrícula y las celdas clave de petición se escriben SIEMPRE como STRING
 * (con {@link Cell#setCellValue(String)}). Las columnas Petición y Matrícula
 * de Resultado son {@code asText.columns} desde 1.6.2, asi que el SUMIFS
 * compara string-string en ambos lados. Si escribiéramos las matrículas
 * todo-dígito como NUMERIC, el SUMIFS daría 0.</p>
 *
 * <p>Esta clase es <b>package-private</b>; solo {@link ResponsablesSheetBuilder}
 * la invoca. Es stateless: cada llamada a {@link #writeTable} es independiente.</p>
 */
final class ResponsablePivotBuilder {

    /** Letra de columna de la celda clave (Petición). Siempre A. */
    private static final int KEY_COL_IDX = 0;
    /** Etiqueta de la celda inferior izquierda en la fila de totales. */
    private static final String LABEL_TOTAL = "Total";
    /** Etiqueta de la columna clave (cabecera de la primera columna). */
    private static final String LABEL_PETICION = "Petición";
    /** Etiqueta usada cuando un responsable no tiene datos de pivot que mostrar. */
    private static final String LABEL_NO_DATA = "(Sin datos)";

    private final ConfigLoader config;

    ResponsablePivotBuilder(ConfigLoader config) {
        this.config = config;
    }

    /**
     * Datos consumidos por el builder. Inmutable, construido por
     * {@link ResponsablesSheetBuilder} tras la pasada única sobre Resultado.
     */
    static final class PivotInputs {
        /** Nombre de la hoja origen (típicamente "Resultado"). */
        final String sourceSheet;
        /** Letra Excel de la columna Petición en la hoja origen. */
        final String peticionLetter;
        /** Letra Excel de la columna Matrícula en la hoja origen. */
        final String matriculaLetter;
        /** Letra Excel de la columna Res. Tecnico en la hoja origen. */
        final String responsableLetter;
        /** Letra Excel de la columna sumada (Jira o REAL). */
        final String valueLetter;
        /** Título visible de la tabla (fila merged). */
        final String title;
        /** Peticiones únicas del responsable, ya ordenadas. */
        final List<String> peticiones;
        /** Matriculas únicas del responsable, ya ordenadas. */
        final List<String> matriculas;
        /** Tope superior del rango SUMIFS (ej. 10000). */
        final int sumifsMaxRow;

        // Constructor con muchos parametros: es un POJO inmutable que agrupa
        // los datos de entrada de la pivot. La regla PMD ExcessiveParameterList
        // esta excluida en el ruleset del proyecto (pmd-ruleset.xml) y un
        // builder con setters romperia la inmutabilidad, por lo que no se
        // necesita supresion explicita.
        PivotInputs(String sourceSheet,
                    String peticionLetter,
                    String matriculaLetter,
                    String responsableLetter,
                    String valueLetter,
                    String title,
                    List<String> peticiones,
                    List<String> matriculas,
                    int sumifsMaxRow) {
            this.sourceSheet = sourceSheet;
            this.peticionLetter = peticionLetter;
            this.matriculaLetter = matriculaLetter;
            this.responsableLetter = responsableLetter;
            this.valueLetter = valueLetter;
            this.title = title;
            this.peticiones = peticiones;
            this.matriculas = matriculas;
            this.sumifsMaxRow = sumifsMaxRow;
        }
    }

    /**
     * Resultado de escribir una tabla pivot. Se devuelve para que el
     * orquestador encadene la siguiente con el {@code gapRows} adecuado.
     */
    static final class PivotResult {
        /** Indice 0-based de la primera fila ocupada (la del titulo). */
        final int firstRow0;
        /** Indice 0-based de la última fila ocupada (la de totales o "Sin datos"). */
        final int lastRow0;
        /** Numero de fórmulas SUMIFS escritas (peticiones × matrículas). 0 si "Sin datos". */
        final int sumifsCount;

        PivotResult(int firstRow0, int lastRow0, int sumifsCount) {
            this.firstRow0 = firstRow0;
            this.lastRow0 = lastRow0;
            this.sumifsCount = sumifsCount;
        }
    }

    /**
     * Escribe una tabla pivot completa en {@code targetSheet}, comenzando en
     * la fila {@code titleRow0} (0-based). No toca filas anteriores.
     *
     * @param workbook    workbook destino (debe ser el mismo que contiene targetSheet).
     * @param targetSheet hoja del responsable, ya creada y con cabecera A1.
     * @param titleRow0   fila 0-based donde colocar el titulo de esta tabla.
     * @param inputs      datos pre-calculados (peticiones/matriculas únicas + letras + titulo).
     * @return resumen de la tabla escrita ({@link PivotResult}).
     */
    PivotResult writeTable(Workbook workbook, Sheet targetSheet, int titleRow0, PivotInputs inputs) {
        // Caso degenerado: sin peticiones o sin matriculas, no hay pivot
        // que escribir. Igualmente colocamos el título y "(Sin datos)" para
        // que la hoja sea legible y el caller sepa cuántas filas ha
        // ocupado el bloque.
        if (inputs.peticiones.isEmpty() || inputs.matriculas.isEmpty()) {
            return writeEmpty(workbook, targetSheet, titleRow0, inputs.title);
        }
        return writeFull(workbook, targetSheet, titleRow0, inputs);
    }

    // ==================================================================
    //  Escritura de tabla con datos
    // ==================================================================

    private PivotResult writeFull(Workbook wb, Sheet sheet, int titleRow0, PivotInputs inputs) {
        CellStyle titleStyle = StyleFactory.summaryBlockHeader(wb);
        CellStyle headerStyle = StyleFactory.summarySubHeaderGray(wb, true);
        CellStyle valueCellStyle = StyleFactory.summaryValueCell(wb);
        CellStyle numericCellStyle = StyleFactory.summaryNumericCell(wb);
        CellStyle totalsStyle = StyleFactory.summaryTotalCell(wb, true);

        int nP = inputs.peticiones.size();
        int nM = inputs.matriculas.size();

        // Total de columnas Excel: 1 (clave Petición) + nM matriculas + 1 (Total)
        int totalColsExcel = 1 + nM + 1;
        int totalColIdx0 = 1 + nM; // 0-based: indice de la columna "Total"

        // 1) Fila de titulo (merged)
        writeTitleRow(sheet, titleRow0, totalColsExcel, inputs.title, titleStyle);

        // 2) Fila de cabecera (titleRow0 + 1)
        int headerRow0 = titleRow0 + 1;
        writeHeaderRow(sheet, headerRow0, inputs.matriculas, totalColIdx0, headerStyle);

        // 3) Filas de datos (peticiones)
        int firstDataRow0 = headerRow0 + 1;
        String quotedSum = PoiUtils.quoteSheetName(inputs.sourceSheet);
        int maxRow = inputs.sumifsMaxRow;

        int sumifsCount = 0;
        for (int i = 0; i < nP; i++) {
            String peticion = inputs.peticiones.get(i);
            int rowIdx0 = firstDataRow0 + i;
            Row r = sheet.createRow(rowIdx0);

            // Celda clave Petición — STRING obligatorio (lección 1.7.1)
            Cell keyCell = r.createCell(KEY_COL_IDX);
            keyCell.setCellValue(peticion);
            keyCell.setCellStyle(valueCellStyle);

            // SUMIFS por cada matrícula
            for (int j = 0; j < nM; j++) {
                Cell c = r.createCell(1 + j);
                String matrColLetter = CellReference.convertNumToColString(1 + j);
                String matrHeaderRef = matrColLetter + "$" + (headerRow0 + 1);

                String formula = "SUMIFS("
                        + quotedSum + "!" + inputs.valueLetter + "2:"
                        + inputs.valueLetter + maxRow + ","
                        + quotedSum + "!" + inputs.peticionLetter + "2:"
                        + inputs.peticionLetter + maxRow + ","
                        + "$A" + (rowIdx0 + 1) + ","
                        + quotedSum + "!" + inputs.matriculaLetter + "2:"
                        + inputs.matriculaLetter + maxRow + ","
                        + matrHeaderRef + ","
                        + quotedSum + "!" + inputs.responsableLetter + "2:"
                        + inputs.responsableLetter + maxRow + ","
                        + "$A$1)";
                c.setCellFormula(formula);
                c.setCellStyle(numericCellStyle);
                sumifsCount++;
            }

            // Columna Total por fila: SUM(B:lastMatr)
            Cell rowTotal = r.createCell(totalColIdx0);
            String firstMatrLetter = CellReference.convertNumToColString(1);
            String lastMatrLetter = CellReference.convertNumToColString(nM);
            rowTotal.setCellFormula("SUM(" + firstMatrLetter + (rowIdx0 + 1)
                    + ":" + lastMatrLetter + (rowIdx0 + 1) + ")");
            rowTotal.setCellStyle(totalsStyle);
        }

        // 4) Fila de totales
        int totalsRow0 = firstDataRow0 + nP;
        writeTotalsRow(sheet, totalsRow0, firstDataRow0, nP, nM, totalColIdx0, totalsStyle);

        return new PivotResult(titleRow0, totalsRow0, sumifsCount);
    }

    private void writeTitleRow(Sheet sheet, int titleRow0, int totalColsExcel,
                                      String title, CellStyle titleStyle) {
        Row titleRow = sheet.createRow(titleRow0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(title);
        titleCell.setCellStyle(titleStyle);
        for (int c = 1; c < totalColsExcel; c++) {
            titleRow.createCell(c).setCellStyle(titleStyle);
        }
        if (totalColsExcel > 1) {
            sheet.addMergedRegion(new CellRangeAddress(
                    titleRow0, titleRow0, 0, totalColsExcel - 1));
        }
    }

    private void writeHeaderRow(Sheet sheet, int headerRow0, List<String> matriculas,
                                       int totalColIdx0, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(headerRow0);
        Cell keyHeader = headerRow.createCell(0);
        keyHeader.setCellValue(LABEL_PETICION);
        keyHeader.setCellStyle(headerStyle);
        for (int i = 0; i < matriculas.size(); i++) {
            Cell c = headerRow.createCell(1 + i);
            // STRING obligatorio (lección 1.7.1) — incluso si la matrícula
            // es todo dígitos, el SUMIFS la compara contra la columna
            // Matrícula de Resultado, que está marcada asText.
            c.setCellValue(matriculas.get(i));
            c.setCellStyle(headerStyle);
        }
        Cell totalHeader = headerRow.createCell(totalColIdx0);
        totalHeader.setCellValue(LABEL_TOTAL);
        totalHeader.setCellStyle(headerStyle);
    }

    private void writeTotalsRow(Sheet sheet, int totalsRow0, int firstDataRow0,
                                       int nP, int nM, int totalColIdx0,
                                       CellStyle totalsStyle) {
        Row totalsRow = sheet.createRow(totalsRow0);
        Cell totLabel = totalsRow.createCell(0);
        totLabel.setCellValue(LABEL_TOTAL);
        totLabel.setCellStyle(totalsStyle);

        int firstDataExcel = firstDataRow0 + 1;
        int lastDataExcel = firstDataRow0 + nP;

        for (int j = 0; j < nM; j++) {
            String letter = CellReference.convertNumToColString(1 + j);
            Cell c = totalsRow.createCell(1 + j);
            c.setCellFormula("SUM(" + letter + firstDataExcel
                    + ":" + letter + lastDataExcel + ")");
            c.setCellStyle(totalsStyle);
        }

        Cell grandTotal = totalsRow.createCell(totalColIdx0);
        String totalLetter = CellReference.convertNumToColString(totalColIdx0);
        grandTotal.setCellFormula("SUM(" + totalLetter + firstDataExcel
                + ":" + totalLetter + lastDataExcel + ")");
        grandTotal.setCellStyle(totalsStyle);
    }

    // ==================================================================
    //  Caso degenerado (sin peticiones o sin matriculas)
    // ==================================================================

    private PivotResult writeEmpty(Workbook wb, Sheet sheet, int titleRow0, String title) {
        // Aun en el caso degenerado, mantener la consistencia visual:
        // titulo con su estilo, y debajo un literal "(Sin datos)".
        CellStyle titleStyle = StyleFactory.summaryBlockHeader(wb);
        CellStyle valueCellStyle = StyleFactory.summaryValueCell(wb);

        Row titleRow = sheet.createRow(titleRow0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(title);
        titleCell.setCellStyle(titleStyle);

        int noDataRow0 = titleRow0 + 1;
        Row noDataRow = sheet.createRow(noDataRow0);
        Cell noDataCell = noDataRow.createCell(0);
        noDataCell.setCellValue(LABEL_NO_DATA);
        noDataCell.setCellStyle(valueCellStyle);

        return new PivotResult(titleRow0, noDataRow0, 0);
    }

    /**
     * Helper estatico para que el orquestador {@link ResponsablesSheetBuilder}
     * resuelva la letra de cada columna de Resultado una sola vez (todas las
     * tablas comparten origen). Devuelve {@code null} si la columna no existe.
     */
    static String letterOrNull(Row headerRow, String columnName) {
        if (headerRow == null) return null;
        int idx = PoiUtils.findColumnIndex(headerRow, columnName);
        return idx < 0 ? null : CellReference.convertNumToColString(idx);
    }

    /** Acceso al ConfigLoader para los tests; no usado actualmente fuera. */
    ConfigLoader config() { return config; }
}

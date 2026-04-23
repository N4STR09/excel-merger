package com.excelmerger.sheet.column;

import com.excelmerger.RunReport;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;
import java.util.Optional;

/**
 * Estrategia de escritura de una columna de la hoja MES.
 *
 * <p>Cada implementacion encapsula la logica de un tipo concreto
 * ({@code COPY}, {@code SUMIFS}, {@code FORMULA}, {@code EMPTY}). El
 * orquestador ({@code MesSheetBuilder}) solo itera la lista de estrategias
 * y no conoce los detalles de cada tipo.</p>
 *
 * <p>Ciclo de vida esperado:</p>
 * <ol>
 *   <li>Se instancia via {@link MesColumnStrategyFactory#fromConfig}.</li>
 *   <li>{@code preValidate} comprueba referencias; si falla, marca disabled
 *       y anade warnings en el {@link RunReport}.</li>
 *   <li>Por cada fila de datos se llama a {@code writeCell}; si la columna
 *       esta {@code disabled}, la celda se escribe como blank.</li>
 * </ol>
 */
public interface MesColumnStrategy {

    /** Nombre de la columna, tal y como se escribira en la cabecera. */
    String getName();

    /** {@code true} si esta columna debe tener formato condicional "verde si &gt;= 0". */
    boolean isGreenIfPositive();

    /**
     * Color de fondo permanente aplicado a las celdas de esta columna
     * (no formato condicional). {@code null} si la columna no lleva fill.
     * Soportados: {@code LIGHT_GREEN}, {@code LIGHT_BLUE}, {@code LIGHT_YELLOW},
     * {@code LIGHT_RED}, {@code LIGHT_LAVENDER}.
     */
    String getFillColor();

    /**
     * Nombre de otra columna MES con la que se compara en runtime: si
     * el valor de ESTA celda es distinto del de la columna referenciada
     * en la misma fila, se pinta el fondo de rojo claro. {@code null}
     * si no aplica.
     */
    String getRedIfNotEqualTo();

    /**
     * {@code true} si {@code preValidate} detecto un problema y las celdas
     * de esta columna deben escribirse como blank.
     */
    boolean isDisabled();

    /**
     * Comprueba en tiempo de construccion que todas las referencias de la
     * columna se pueden resolver. Puede marcar la columna como disabled y
     * anadir warnings al {@link RunReport}.
     */
    void preValidate(Sheet source, int sourceHeaderRow0, Workbook workbook,
                     Map<String, Integer> mesColIndexByName, RunReport report);

    /**
     * Escribe el valor de la celda correspondiente a esta columna en la
     * fila de MES indicada. Si la columna esta disabled, escribe blank.
     *
     * @param target            celda destino en MES
     * @param srcRow            fila origen en la hoja de extraccion
     * @param source            hoja de extraccion
     * @param sourceHeaderRow0  indice (0-based) de la fila de cabeceras en origen
     * @param workbook          workbook destino
     * @param sourceExcelRow    indice (1-based) de la fila origen en Excel
     * @param mesColIndexByName mapa nombre-&gt;indice de columnas MES
     * @param mesExcelRow       indice (1-based) de la fila MES en Excel
     */
    void writeCell(Cell target, Row srcRow, Sheet source, int sourceHeaderRow0,
                   Workbook workbook, int sourceExcelRow,
                   Map<String, Integer> mesColIndexByName, int mesExcelRow);

    /**
     * Plantilla de formula original, si este tipo de columna la tiene.
     * Usado por {@code MesSheetBuilder} para extraer enlaces VLOOKUP y
     * detectar "apps sin mapeo". La implementacion por defecto devuelve
     * {@code Optional.empty()}; solo {@code FormulaColumnStrategy} lo expone.
     */
    default Optional<String> formulaTemplate() {
        return Optional.empty();
    }
}

package com.excelmerger;

import com.excelmerger.util.PoiUtils;
import com.excelmerger.util.StyleFactory;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.text.Collator;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * v2.3.0: construye una hoja vacia por cada responsable distinto que
 * aparezca en la columna {@code Res. Tecnico} de la hoja {@code Resultado}.
 *
 * <p>Solo se invoca desde {@link ExcelMerger} cuando el modo de salida es
 * {@link OutputMode#RESPONSABLES} o {@link OutputMode#COMPLETO}. En modo
 * {@link OutputMode#CIERRE} no se llama y por tanto el comportamiento
 * v2.2.0 queda preservado al 100%.</p>
 *
 * <p><b>Reglas de extraccion</b>:</p>
 * <ul>
 *   <li><b>Trim + case-insensitive</b>: valores como {@code "tresp1@x"},
 *       {@code "TRESP1@x"} y {@code " tresp1@x "} colapsan en una unica
 *       hoja. Se usa la primera ocurrencia (en orden de filas de
 *       {@code Resultado}) como nombre canonico, coherente con la
 *       politica del builder de Resumen.</li>
 *   <li><b>Sin filtros ni exclusiones</b>: si el valor (tras trim) es
 *       no-vacio, genera hoja. Decisiones del usuario en Fase 0.</li>
 *   <li><b>Orden de creacion</b>: alfabetico por nombre canonico, con
 *       {@link Collator} {@code es_ES} y {@code PRIMARY strength} para
 *       que las tildes ordenen como un humano espera y la salida sea
 *       determinista entre runs.</li>
 * </ul>
 *
 * <p><b>Reglas de naming de la hoja</b>:</p>
 * <ul>
 *   <li>El nombre se sanea con {@link FileProfileResolver#safeSheetName(String)}
 *       (reemplaza {@code \\ / ? * [ ] :} por {@code _}, trunca a 31 chars).
 *       Si el saneado difiere del original, se emite warning categoria
 *       {@code RESPONSABLE}.</li>
 *   <li>Si el nombre saneado colisiona con una hoja ya existente
 *       (incluyendo otra hoja de responsable o las builtin), se sufija
 *       {@code _2}, {@code _3}, ... via
 *       {@link FileProfileResolver#ensureUniqueSheetName(Workbook, String)}
 *       y se emite warning {@code RESPONSABLE}.</li>
 * </ul>
 *
 * <p><b>Contenido de cada hoja</b>: por ahora solo {@code A1} con el
 * nombre canonico del responsable, en estilo titulo (bold, 14 pt). En
 * sesiones posteriores se anadiran dos tablas de resumen (decision
 * Fase 0 #5).</p>
 *
 * <p><b>Casos borde</b>:</p>
 * <ul>
 *   <li>Hoja {@code Resultado} ausente -> warning {@code RESPONSABLE},
 *       0 hojas creadas, sin error.</li>
 *   <li>Hoja {@code Resultado} sin filas de datos -> warning
 *       {@code RESPONSABLE}, 0 hojas, sin error.</li>
 *   <li>Cabecera {@code Res. Tecnico} no encontrada -> warning
 *       {@code RESPONSABLE}, 0 hojas, sin error.</li>
 *   <li>Todas las celdas de la columna vacias -> 0 hojas, sin warning
 *       (caso degenerado pero legal: simplemente no hay nada que generar).</li>
 * </ul>
 */
public class ResponsablesSheetBuilder {

    private static final Logger log = LoggerFactory.getLogger(ResponsablesSheetBuilder.class);

    private static final String WARN_CATEGORY = "RESPONSABLE";

    private final ConfigLoader config;
    private final RunReport report;

    public ResponsablesSheetBuilder(ConfigLoader config, RunReport report) {
        this.config = config;
        this.report = report;
    }

    /**
     * Recorre la hoja {@code Resultado} y crea N hojas vacias, una por
     * responsable distinto. Idempotente respecto a la presencia previa
     * de hojas: si una colision ocurre, se aplica sufijo (no se sobreescribe).
     */
    public void buildAll(Workbook workbook) {
        // El nombre de la hoja Resultado y de la columna se leen de la
        // misma config que usan los builders existentes, asi seguimos
        // funcionando aunque alguien renombre la hoja por config.
        String resultadoName = config.get("mes.sheetName", "MES");
        String responsibleColumn = config.get("summary.byResponsible.column", "Res. Tecnico");

        Sheet resultado = workbook.getSheet(resultadoName);
        if (resultado == null) {
            log.warn("[Responsables] La hoja '{}' no existe; no se generan hojas por responsable.",
                    resultadoName);
            report.addWarning(WARN_CATEGORY,
                    "No se pudo construir hojas por responsable: hoja '"
                            + resultadoName + "' ausente.");
            return;
        }
        Row header = resultado.getRow(0);
        if (header == null) {
            log.warn("[Responsables] La hoja '{}' no tiene cabecera; no se generan hojas por responsable.",
                    resultadoName);
            report.addWarning(WARN_CATEGORY,
                    "No se pudo construir hojas por responsable: hoja '"
                            + resultadoName + "' sin cabecera.");
            return;
        }
        int colIdx = PoiUtils.findColumnIndex(header, responsibleColumn);
        if (colIdx < 0) {
            log.warn("[Responsables] Columna '{}' no encontrada en '{}'; no se generan hojas por responsable.",
                    responsibleColumn, resultadoName);
            report.addWarning(WARN_CATEGORY,
                    "No se pudo construir hojas por responsable: columna '"
                            + responsibleColumn + "' no encontrada en '" + resultadoName + "'.");
            return;
        }
        if (resultado.getLastRowNum() < 1) {
            log.info("[Responsables] La hoja '{}' no tiene filas de datos; nada que generar.",
                    resultadoName);
            report.addWarning(WARN_CATEGORY,
                    "Hoja '" + resultadoName + "' sin filas de datos; 0 hojas por responsable generadas.");
            return;
        }

        // Recolectar responsables unicos (clave lower-case, valor primer
        // literal visto). LinkedHashMap preserva el orden de insercion
        // para que "primera aparicion" sea consultable en pruebas, aunque
        // luego ordenemos alfabeticamente para crear las hojas.
        Map<String, String> uniqueByLower = new LinkedHashMap<>();
        for (int r = 1; r <= resultado.getLastRowNum(); r++) {
            Row row = resultado.getRow(r);
            if (row == null) continue;
            Cell cell = row.getCell(colIdx);
            if (cell == null) continue;
            String raw = PoiUtils.cellAsString(cell);
            if (raw == null) continue;
            String trimmed = raw.trim();
            if (trimmed.isEmpty()) continue;
            String key = trimmed.toLowerCase(Locale.ROOT);
            uniqueByLower.putIfAbsent(key, trimmed);
        }

        if (uniqueByLower.isEmpty()) {
            log.info("[Responsables] No hay responsables no-vacios en '{}'; 0 hojas generadas.",
                    resultadoName);
            return;
        }

        // Ordenar alfabeticamente por nombre canonico, con Collator
        // es_ES strength PRIMARY para tratar tildes/case como humano.
        List<String> canonicalNames = new ArrayList<>(uniqueByLower.values());
        Collator collator = Collator.getInstance(Locale.of("es", "ES"));
        collator.setStrength(Collator.PRIMARY);
        Collections.sort(canonicalNames, collator);

        log.info("[Responsables] {} responsable(s) distinto(s) detectado(s) en '{}': {}",
                canonicalNames.size(), resultadoName, canonicalNames);

        CellStyle titleStyle = StyleFactory.title(workbook);

        for (String canonical : canonicalNames) {
            String safe = FileProfileResolver.safeSheetName(canonical);
            if (!safe.equals(canonical)) {
                report.addWarning(WARN_CATEGORY,
                        "Nombre de responsable saneado para uso como nombre de hoja: '"
                                + canonical + "' -> '" + safe + "'.");
            }
            String unique = FileProfileResolver.ensureUniqueSheetName(workbook, safe);
            if (!unique.equals(safe)) {
                report.addWarning(WARN_CATEGORY,
                        "Colision de nombre de hoja para responsable; resuelto con sufijo: '"
                                + safe + "' -> '" + unique + "'.");
            }

            Sheet sheet = workbook.createSheet(unique);
            Row headerRow = sheet.createRow(0);
            Cell a1 = headerRow.createCell(0);
            a1.setCellValue(canonical);
            a1.setCellStyle(titleStyle);
            try { sheet.autoSizeColumn(0); } catch (Exception ignored) { /* ignorado */ }

            report.addSheet(unique, 1);
            log.info("[Responsables] Hoja creada: '{}' (responsable canonico: '{}')", unique, canonical);
        }
    }
}

package com.excelmerger.sheet.column;

import com.excelmerger.ConfigLoader;
import com.excelmerger.RunReport;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * Factoria de {@link MesColumnStrategy} a partir de las claves del config
 * {@code mes.col.<N>.*}. Si la configuracion de una columna es invalida o
 * esta incompleta, devuelve {@link EmptyColumnStrategy} como fallback
 * (equivalente al comportamiento anterior del builder).
 */
public final class MesColumnStrategyFactory {

    private static final Logger log = LoggerFactory.getLogger(MesColumnStrategyFactory.class);

    private MesColumnStrategyFactory() {
        // Utility class
    }

    public static MesColumnStrategy fromConfig(ConfigLoader config, int idx,
                                               String name, String type, RunReport report) {
        String prefix = "mes.col." + idx + ".";
        boolean greenIfPositive = config.getBoolean(prefix + "greenIfPositive", false);
        String fillColor = nullIfBlank(config.get(prefix + "fill", null));
        String redIfNotEqualTo = nullIfBlank(config.get(prefix + "redIfNotEqualTo", null));
        String normalizedType = type == null ? "" : type.trim().toUpperCase();

        switch (normalizedType) {
            case "COPY": {
                String from = config.get(prefix + "from", null);
                if (from == null) {
                    log.info("Columna '{}' (COPY) sin 'from'. Tratada como EMPTY.", name);
                    return new EmptyColumnStrategy(name, greenIfPositive, fillColor, redIfNotEqualTo);
                }
                return new CopyColumnStrategy(name, greenIfPositive, fillColor, redIfNotEqualTo, from);
            }
            case "SUMIFS": {
                String fromSheet = config.get(prefix + "from", null);
                String sumHeader = config.get(prefix + "sum", null);
                List<String[]> matches = parseMatchExpression(config.get(prefix + "match", ""));
                if (fromSheet == null || sumHeader == null || matches.isEmpty()) {
                    log.info("Columna '{}' (SUMIFS) mal configurada. Tratada como EMPTY.", name);
                    return new EmptyColumnStrategy(name, greenIfPositive, fillColor, redIfNotEqualTo);
                }
                return new SumIfsColumnStrategy(name, greenIfPositive, fillColor, redIfNotEqualTo,
                        fromSheet, sumHeader, matches);
            }
            case "FORMULA": {
                String tpl = config.get(prefix + "formula", null);
                if (tpl == null || tpl.trim().isEmpty()) {
                    log.info("Columna '{}' (FORMULA) sin 'formula'. Tratada como EMPTY.", name);
                    return new EmptyColumnStrategy(name, greenIfPositive, fillColor, redIfNotEqualTo);
                }
                return new FormulaColumnStrategy(name, greenIfPositive, fillColor, redIfNotEqualTo, tpl);
            }
            case "EMPTY":
                return new EmptyColumnStrategy(name, greenIfPositive, fillColor, redIfNotEqualTo);
            default:
                log.info("Tipo '{}' desconocido en '{}'. Tratada como EMPTY.", type, name);
                return new EmptyColumnStrategy(name, greenIfPositive, fillColor, redIfNotEqualTo);
        }
    }

    private static String nullIfBlank(String s) {
        if (s == null) return null;
        String t = s.trim();
        return t.isEmpty() ? null : t;
    }

    /**
     * Parsea una expresion {@code match=remoteHeader1:localHeader1,remoteHeader2:localHeader2}
     * en una lista de pares. Los vacios se descartan silenciosamente.
     */
    private static List<String[]> parseMatchExpression(String expr) {
        List<String[]> out = new ArrayList<>();
        if (expr == null || expr.trim().isEmpty()) return out;
        for (String pair : expr.split(",")) {
            String[] parts = pair.split(":", 2);
            if (parts.length == 2) {
                out.add(new String[]{parts[0].trim(), parts[1].trim()});
            }
        }
        return out;
    }
}

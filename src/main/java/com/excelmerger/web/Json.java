package com.excelmerger.web;

/**
 * Serializador JSON minimo para las respuestas de la API local (v4.0.0).
 *
 * <p>El proyecto no depende de Jackson/Gson y las respuestas tienen
 * siempre la misma forma pequena, asi que el JSON se construya a mano.
 * Si la API crece, toca incorporar una libreria real en lugar de
 * ampliar esta clase.</p>
 *
 * <p>Nota: se destinan a {@code fetch}, no a embeberse en un
 * {@code <script>}, por eso no se escapan {@code <} ni {@code >} y si se
 * escapan los separadores de linea {@code U+2028/U+2029} por robustez
 * JSON estandar.</p>
 */
final class Json {

    /** Digitos hexadecimales para componer escapes Unicode de 4 digitos (p. ej. u005f). */
    private static final char[] HEX = "0123456789abcdef".toCharArray();

    private Json() {
        // Clase de utilidad
    }

    /**
     * Devuelve {@code value} como literal JSON entre comillas, escapando
     * comillas, barras invertidas, controles y los separadores de linea
     * unicode. {@code null} serializa como {@code null}.
     *
     * @param value cadena de entrada (puede ser {@code null}).
     * @return literal JSON.
     */
    static String quote(String value) {
        if (value == null) {
            return "null";
        }
        StringBuilder sb = new StringBuilder(value.length() + 16);
        sb.append('"');
        for (int i = 0; i < value.length(); i++) {
            char c = value.charAt(i);
            switch (c) {
                case '"' -> sb.append("\\\"");
                case '\\' -> sb.append("\\\\");
                case '\n' -> sb.append("\\n");
                case '\r' -> sb.append("\\r");
                case '\t' -> sb.append("\\t");
                case '\f' -> sb.append("\\f");
                case '\b' -> sb.append("\\b");
                case '\u2028' -> sb.append("\\u2028");
                case '\u2029' -> sb.append("\\u2029");
                default -> appendOther(sb, c);
            }
        }
        sb.append('"');
        return sb.toString();
    }

    /**
     * Anade {@code c} tal cual si es imprimible normal, o como un
     * escape Unicode de 4 digitos hexadecimales (u00XX) si es un
     * control {@code < 0x20} no contemplado arriba.
     */
    private static void appendOther(StringBuilder sb, char c) {
        if (c < 0x20) {
            sb.append("\\u00");
            sb.append(HEX[(c >> 4) & 0xF]);
            sb.append(HEX[c & 0xF]);
        } else {
            sb.append(c);
        }
    }
}

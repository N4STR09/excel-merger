package com.excelmerger;

/**
 * v2.3.0: modo de generacion de hojas en el libro de salida.
 *
 * <p>Tres modos seleccionables via la clave {@code output.mode} en
 * {@code config.properties}:</p>
 *
 * <ul>
 *   <li>{@link #CIERRE} — comportamiento historico v2.2.0. Genera
 *       {@code Cierre}, {@code Extraccion}, {@code Deuda} (si el usuario
 *       aporta el 3er fichero), {@code Equipos} (oculta, lookup),
 *       {@code Resultado} y {@code Resumen}.</li>
 *   <li>{@link #RESPONSABLES} — genera {@code Cierre}, {@code Extraccion},
 *       {@code Equipos} (oculta), {@code Resultado} y N hojas vacias, una
 *       por cada responsable distinto que aparezca en
 *       {@code Resultado.Res. Tecnico} (trim + case-insensitive). NO genera
 *       {@code Deuda} (la copia del input se omite) ni {@code Resumen}.</li>
 *   <li>{@link #COMPLETO} — la suma de los dos anteriores: todas las hojas
 *       de {@link #CIERRE} mas las hojas por responsable.</li>
 * </ul>
 *
 * <p>El default cuando la clave esta ausente o vacia es {@link #CIERRE},
 * preservando 100% el comportamiento de v2.2.0 sin tocar configs
 * existentes.</p>
 */
public enum OutputMode {
    CIERRE,
    RESPONSABLES,
    COMPLETO;

    /**
     * Parsea el valor crudo leido de {@code config.properties}. Estricto,
     * case-sensitive, solo admite las cadenas exactas {@code "cierre"},
     * {@code "responsables"}, {@code "completo"}. Cualquier otro valor
     * (incluido {@code null}, mayusculas, traducciones, espacios internos)
     * lanza {@link IllegalArgumentException}.
     *
     * <p>El espacio al inicio/final lo gestiona {@code ConfigLoader.get},
     * que ya hace {@code trim()} en el valor devuelto. La logica
     * "ausente o vacio -> CIERRE" NO vive aqui: vive en el llamador
     * (ver {@link ExcelMerger#merge()}).</p>
     */
    public static OutputMode parseStrict(String raw) {
        if (raw == null) {
            throw new IllegalArgumentException(
                    "Valor null para output.mode. Validos: cierre, responsables, completo");
        }
        switch (raw) {
            case "cierre":       return CIERRE;
            case "responsables": return RESPONSABLES;
            case "completo":     return COMPLETO;
            default:
                throw new IllegalArgumentException(
                        "Valor invalido '" + raw + "' para output.mode. "
                                + "Validos: cierre, responsables, completo");
        }
    }
}

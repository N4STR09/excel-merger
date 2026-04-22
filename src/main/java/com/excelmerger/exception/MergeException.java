package com.excelmerger.exception;

/**
 * Fallo durante la fusion de contenidos: error de POI al leer/escribir
 * workbooks, inconsistencias inesperadas entre hojas, etc. Es el
 * "catch-all" del proceso de merge cuando el fallo no encaja como
 * configuracion, entrada u output.
 * Se mapea a <b>exit code 1</b> en {@code Main}, igual que cualquier
 * otra {@link Exception} no tipada.
 */
public class MergeException extends ExcelMergerException {

    private static final long serialVersionUID = 1L;

    public MergeException(String message) {
        super(message);
    }

    public MergeException(String message, Throwable cause) {
        super(message, cause);
    }
}

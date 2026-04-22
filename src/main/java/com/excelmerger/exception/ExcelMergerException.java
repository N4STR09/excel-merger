package com.excelmerger.exception;

/**
 * Excepcion base (unchecked) de Excel Merger. Todas las excepciones de
 * negocio de la aplicacion extienden de esta.
 *
 * <p>Se modela como {@link RuntimeException} para no contaminar las firmas
 * del pipeline de merge con {@code throws} checked en cada capa intermedia.
 * {@code Main} hace un catch por tipo y mapea a codigos de salida documentados.</p>
 */
public class ExcelMergerException extends RuntimeException {

    private static final long serialVersionUID = 1L;

    public ExcelMergerException(String message) {
        super(message);
    }

    public ExcelMergerException(String message, Throwable cause) {
        super(message, cause);
    }
}

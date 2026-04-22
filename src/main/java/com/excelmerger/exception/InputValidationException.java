package com.excelmerger.exception;

/**
 * Fallo en la validacion de las entradas: directorio de entrada ausente,
 * sin los ficheros Excel esperados, ficheros bloqueados (~$...), o numero
 * de ficheros fuera del rango permitido.
 * Se mapea a <b>exit code 3</b> en {@code Main}.
 */
public class InputValidationException extends ExcelMergerException {

    private static final long serialVersionUID = 1L;

    public InputValidationException(String message) {
        super(message);
    }

    public InputValidationException(String message, Throwable cause) {
        super(message, cause);
    }
}

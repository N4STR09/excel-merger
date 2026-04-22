package com.excelmerger.exception;

/**
 * Fallo al preparar o escribir el fichero de salida: lock ~$... sobre el
 * output, fichero existente con {@code output.overwrite=false}, error de
 * escritura en disco, fallo del backup.
 * Se mapea a <b>exit code 4</b> en {@code Main}.
 */
public class OutputException extends ExcelMergerException {

    private static final long serialVersionUID = 1L;

    public OutputException(String message) {
        super(message);
    }

    public OutputException(String message, Throwable cause) {
        super(message, cause);
    }
}

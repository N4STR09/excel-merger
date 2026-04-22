package com.excelmerger.exception;

/**
 * Fallo relativo a la configuracion: fichero {@code config.properties}
 * ausente, ilegible, o con valores que impiden arrancar el proceso.
 * Se mapea a <b>exit code 2</b> en {@code Main}.
 */
public class ConfigurationException extends ExcelMergerException {

    private static final long serialVersionUID = 1L;

    public ConfigurationException(String message) {
        super(message);
    }

    public ConfigurationException(String message, Throwable cause) {
        super(message, cause);
    }
}

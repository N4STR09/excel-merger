package com.excelmerger;

import com.excelmerger.exception.ConfigurationException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;

/**
 * Gestiona la carga y acceso a la configuracion de la aplicacion.
 * La configuracion se carga desde un fichero externo 'config.properties'.
 *
 * <p>Si el fichero no existe ni como ruta externa ni en el classpath, o
 * falla la lectura, lanza {@link ConfigurationException} (unchecked).</p>
 */
public class ConfigLoader {

    private static final Logger log = LoggerFactory.getLogger(ConfigLoader.class);
    private static final String DEFAULT_CONFIG_FILE = "config.properties";
    private final Properties properties;

    public ConfigLoader() {
        this(DEFAULT_CONFIG_FILE);
    }

    public ConfigLoader(String configPath) {
        this.properties = new Properties();
        loadProperties(configPath);
    }

    private void loadProperties(String configPath) {
        Path externalPath = Paths.get(configPath);

        // 1. Intentar cargar desde fichero externo (preferido) en UTF-8
        if (Files.exists(externalPath)) {
            try (Reader reader = new InputStreamReader(
                    new FileInputStream(externalPath.toFile()), StandardCharsets.UTF_8)) {
                properties.load(reader);
                log.info("Configuracion cargada desde: {}", externalPath.toAbsolutePath());
                return;
            } catch (IOException e) {
                throw new ConfigurationException(
                        "No se pudo leer el fichero de configuracion '" + configPath + "': "
                                + e.getMessage(), e);
            }
        }

        // 2. Fallback: cargar desde recursos del classpath en UTF-8
        try (InputStream input = getClass().getClassLoader().getResourceAsStream(configPath)) {
            if (input == null) {
                throw new ConfigurationException(
                        "No se encontro el fichero de configuracion: " + configPath);
            }
            try (Reader reader = new InputStreamReader(input, StandardCharsets.UTF_8)) {
                properties.load(reader);
            }
            log.info("Configuracion cargada desde classpath (recursos internos)");
        } catch (IOException e) {
            throw new ConfigurationException(
                    "No se pudo leer el fichero de configuracion '" + configPath
                            + "' desde classpath: " + e.getMessage(), e);
        }
    }

    public String get(String key) {
        String value = properties.getProperty(key);
        if (value == null) {
            throw new IllegalArgumentException("Propiedad no encontrada en config.properties: " + key);
        }
        return value.trim();
    }

    public String get(String key, String defaultValue) {
        String value = properties.getProperty(key, defaultValue);
        return value == null ? null : value.trim();
    }

    public boolean getBoolean(String key, boolean defaultValue) {
        String value = properties.getProperty(key);
        if (value == null) return defaultValue;
        return Boolean.parseBoolean(value.trim());
    }

    public int getInt(String key, int defaultValue) {
        String value = properties.getProperty(key);
        if (value == null) return defaultValue;
        try {
            return Integer.parseInt(value.trim());
        } catch (NumberFormatException e) {
            return defaultValue;
        }
    }

    /**
     * Devuelve el objeto Properties subyacente para iterar todas las claves.
     * Usado por componentes que necesitan descubrir propiedades dinamicas
     * (p. ej. sheet.&lt;id&gt;.cell.&lt;CELDA&gt;).
     */
    public Properties getRawProperties() {
        return properties;
    }
}

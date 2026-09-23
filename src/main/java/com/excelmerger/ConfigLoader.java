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

    /**
     * v4.1.0: carga la configuracion base (externa o classpath, igual que
     * {@link #ConfigLoader()}) y aplica encima los overrides indicados.
     *
     * <p>Es el mecanismo de los ajustes configurables en runtime de la
     * interfaz web: la pagina persiste tres claves de salida
     * ({@code output.mode}, {@code summary.enabled} y
     * {@code summary.byResponsible.enabled}) y esta carga las aplica sobre
     * el config que sea, sin tocar el fichero. Linea de comandos y menu
     * de terminal siguen cargando sin overrides (comportamiento identico
     * al de v3.x).</p>
     *
     * @param overrides claves a sobreescribir tras cargar la base; si es
     *        {@code null} o vacio, el resultado es identico a la carga sin
     *        overrides.
     */
    public ConfigLoader(Properties overrides) {
        this(DEFAULT_CONFIG_FILE, overrides);
    }

    /**
     * v4.1.0: como {@link #ConfigLoader(Properties)} pero con ruta de
     * configuracion explicita ({@code null} = default).
     *
     * @param configPath ruta al fichero de configuracion, o {@code null}
     *        para usar el default ({@code config.properties}).
     * @param overrides  claves a sobreescribir tras cargar la base.
     */
    public ConfigLoader(String configPath, Properties overrides) {
        this.properties = new Properties();
        loadProperties(configPath == null ? DEFAULT_CONFIG_FILE : configPath);
        if (overrides != null && !overrides.isEmpty()) {
            properties.putAll(overrides);
            log.debug("Aplicadas {} claves de override sobre la configuracion base.",
                    overrides.size());
        }
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
     * {@code true} si la clave esta presente (aunque su valor sea vacio)
     * en el fichero de configuracion. Usado para distinguir "clave no
     * definida" de "clave definida explicitamente con el valor default".
     */
    public boolean has(String key) {
        return properties.getProperty(key) != null;
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

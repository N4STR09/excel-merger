package com.excelmerger;

import com.excelmerger.exception.ConfigurationException;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Tests de {@link ConfigLoader}: carga desde fichero externo, carga desde
 * classpath, acceso con/ sin default, coerciones a boolean/int, trim, y
 * lectura de claves no existentes.
 */
class ConfigLoaderTest {

    @Test
    void cargaDesdeFicheroExterno(@TempDir Path tmp) throws IOException {
        Path cfg = tmp.resolve("my.properties");
        Files.writeString(cfg, "input.directory=/data/in\noutput.file=/data/out.xlsx\n",
                StandardCharsets.UTF_8);

        ConfigLoader loader = new ConfigLoader(cfg.toString());

        assertThat(loader.get("input.directory")).isEqualTo("/data/in");
        assertThat(loader.get("output.file")).isEqualTo("/data/out.xlsx");
    }

    @Test
    void getConDefaultDevuelveDefaultCuandoFalta(@TempDir Path tmp) throws IOException {
        Path cfg = tmp.resolve("my.properties");
        Files.writeString(cfg, "foo=bar\n", StandardCharsets.UTF_8);

        ConfigLoader loader = new ConfigLoader(cfg.toString());

        assertThat(loader.get("foo", "DEFAULT")).isEqualTo("bar");
        assertThat(loader.get("missing", "DEFAULT")).isEqualTo("DEFAULT");
    }

    @Test
    void getSinDefaultLanzaSiFaltaLaClave(@TempDir Path tmp) throws IOException {
        Path cfg = tmp.resolve("my.properties");
        Files.writeString(cfg, "only.one=yes\n", StandardCharsets.UTF_8);

        ConfigLoader loader = new ConfigLoader(cfg.toString());

        assertThatThrownBy(() -> loader.get("absent.key"))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("absent.key");
    }

    @Test
    void getRecortaEspaciosAlrededorDelValor(@TempDir Path tmp) throws IOException {
        Path cfg = tmp.resolve("my.properties");
        // Properties ya elimina el espacio inicial, pero el final depende del parser.
        // Forzamos un caso de trim usando una clave con espacios dentro del valor.
        Files.writeString(cfg, "key=  hola mundo   \n", StandardCharsets.UTF_8);

        ConfigLoader loader = new ConfigLoader(cfg.toString());

        assertThat(loader.get("key")).isEqualTo("hola mundo");
    }

    @Test
    void getBooleanInterpretaValoresValidos(@TempDir Path tmp) throws IOException {
        Path cfg = tmp.resolve("my.properties");
        Files.writeString(cfg, "a=true\nb=false\nc=TRUE\nd=no-es-bool\n",
                StandardCharsets.UTF_8);

        ConfigLoader loader = new ConfigLoader(cfg.toString());

        assertThat(loader.getBoolean("a", false)).isTrue();
        assertThat(loader.getBoolean("b", true)).isFalse();
        assertThat(loader.getBoolean("c", false)).isTrue();
        // Boolean.parseBoolean devuelve false si no es "true"
        assertThat(loader.getBoolean("d", true)).isFalse();
        // Clave ausente: devuelve default
        assertThat(loader.getBoolean("missing", true)).isTrue();
        assertThat(loader.getBoolean("missing", false)).isFalse();
    }

    @Test
    void getIntDevuelveDefaultCuandoValorNoNumerico(@TempDir Path tmp) throws IOException {
        Path cfg = tmp.resolve("my.properties");
        Files.writeString(cfg, "good=42\nbad=abc\n", StandardCharsets.UTF_8);

        ConfigLoader loader = new ConfigLoader(cfg.toString());

        assertThat(loader.getInt("good", 7)).isEqualTo(42);
        assertThat(loader.getInt("bad", 7)).isEqualTo(7);
        assertThat(loader.getInt("missing", 7)).isEqualTo(7);
    }

    @Test
    void fallbackAClasspathSiFicheroExternoNoExiste() throws IOException {
        // El resource test-config.properties existe en el classpath de tests.
        ConfigLoader loader = new ConfigLoader("test-config.properties");

        // Cualquier clave conocida del test-config
        assertThat(loader.get("merge.mode")).isEqualTo("SHEETS_SEPARATE");
        assertThat(loader.getBoolean("mes.enabled", false)).isTrue();
    }

    @Test
    void classpathFallbackLanzaSiNoEstaNiEnDiscoNiEnClasspath() {
        assertThatThrownBy(() -> new ConfigLoader("no-existe-en-ningun-sitio.properties"))
                .isInstanceOf(ConfigurationException.class)
                .hasMessageContaining("no-existe-en-ningun-sitio.properties");
    }

    @Test
    void getRawPropertiesExponeTodasLasClaves(@TempDir Path tmp) throws IOException {
        Path cfg = tmp.resolve("my.properties");
        Files.writeString(cfg, "a.b.c=1\na.b.d=2\nother=3\n", StandardCharsets.UTF_8);

        ConfigLoader loader = new ConfigLoader(cfg.toString());

        assertThat(loader.getRawProperties().stringPropertyNames())
                .containsExactlyInAnyOrder("a.b.c", "a.b.d", "other");
    }

    @Test
    void leeUtf8Correctamente(@TempDir Path tmp) throws IOException {
        Path cfg = tmp.resolve("utf8.properties");
        Files.writeString(cfg, "nombre=José Ñoño áéíóú\n", StandardCharsets.UTF_8);

        ConfigLoader loader = new ConfigLoader(cfg.toString());

        assertThat(loader.get("nombre")).isEqualTo("José Ñoño áéíóú");
    }

    // ==================================================================
    //  Contrato usado por run.bat para configs por entorno
    //    run.bat contabilidad -> config-contabilidad.properties
    //  El .bat resuelve el nombre y pasa la ruta al JAR; aqui blindamos
    //  que ConfigLoader acepta esa ruta y lanza ConfigurationException
    //  con mensaje claro si el fichero no existe.
    // ==================================================================

    @Test
    void cargaConfigDeEntornoPorNombreConvencional(@TempDir Path tmp) throws IOException {
        // Simulamos lo que run.bat genera al ejecutarse como 'run.bat contabilidad':
        // un fichero 'config-contabilidad.properties' en la carpeta del .bat.
        Path cfg = tmp.resolve("config-contabilidad.properties");
        Files.writeString(cfg,
                "input.directory=input-contabilidad\n"
              + "output.file=output/contabilidad.xlsx\n",
                StandardCharsets.UTF_8);

        ConfigLoader loader = new ConfigLoader(cfg.toString());

        // Las claves vienen del config de entorno, no del default.
        assertThat(loader.get("input.directory")).isEqualTo("input-contabilidad");
        assertThat(loader.get("output.file")).isEqualTo("output/contabilidad.xlsx");
    }

    @Test
    void configDeEntornoInexistenteLanzaConfigurationException() {
        // Si el usuario invoca 'run.bat pre' y no existe 'config-pre.properties',
        // el .bat aborta antes de llegar aqui. Este test blinda el caso degenerado
        // de que alguien pase la ruta directamente al JAR sin pasar por el .bat.
        assertThatThrownBy(() -> new ConfigLoader("config-pre.properties"))
                .isInstanceOf(ConfigurationException.class)
                .hasMessageContaining("config-pre.properties");
    }
}

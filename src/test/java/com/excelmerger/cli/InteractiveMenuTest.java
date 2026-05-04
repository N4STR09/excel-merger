package com.excelmerger.cli;

import com.excelmerger.App;

import org.jline.terminal.Terminal;
import org.jline.terminal.impl.DumbTerminal;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Function;
import java.util.function.IntSupplier;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests del menu interactivo introducido en v3.0.0.
 *
 * <p>Estrategia: se inyecta un {@link DumbTerminal} de JLine cuyo stdin
 * lee de un {@link ByteArrayInputStream} con el input simulado de cada
 * caso de test, y cuyo stdout escribe en un {@link ByteArrayOutputStream}
 * que luego se inspecciona. Asi no se depende de stdin real ni de un TTY
 * real, lo que permitiria fallos en CI.</p>
 *
 * <p>El runner de la Opcion 1 se inyecta como una lambda
 * {@code Function<String, Integer>} que registra que se ha llamado y
 * devuelve el exit code que el test quiere simular (0 = OK; 1-4 = error).
 * De este modo no se ejerce {@link App#run(String)} real, lo que
 * convertiria estos tests en tests de integracion costosos.</p>
 */
class InteractiveMenuTest {

    /**
     * Construye un menu con stdin simulado y captura de stdout. El
     * runner de la Opcion 2 por defecto devuelve {@link App#EXIT_OK} y
     * no hace nada, util para tests que no estan mirando la Opcion 2.
     * Para tests especificos de Opcion 2, usar
     * {@link #harness(String, Function, IntSupplier)}.
     */
    private static Harness harness(String stdinInput,
                                   Function<String, Integer> runner) throws IOException {
        return harness(stdinInput, runner, () -> App.EXIT_OK);
    }

    /** Variante con runner explicito para Opcion 2. */
    private static Harness harness(String stdinInput,
                                   Function<String, Integer> optionOneRunner,
                                   IntSupplier optionTwoRunner) throws IOException {
        ByteArrayInputStream in = new ByteArrayInputStream(
                stdinInput.getBytes(StandardCharsets.UTF_8));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        // DumbTerminal: terminal sin capacidades de cursor ni colores,
        // perfecto para tests deterministas. JLine sigue funcionando.
        Terminal terminal = new DumbTerminal(
                "test", "dumb", in, out, StandardCharsets.UTF_8);
        PrintStream printStream = new PrintStream(out, true, StandardCharsets.UTF_8);
        InteractiveMenu menu = new InteractiveMenu(
                terminal, optionOneRunner, optionTwoRunner, printStream);
        return new Harness(menu, out, terminal);
    }

    private static final class Harness {
        final InteractiveMenu menu;
        final ByteArrayOutputStream out;
        final Terminal terminal;

        Harness(InteractiveMenu menu, ByteArrayOutputStream out, Terminal terminal) {
            this.menu = menu;
            this.out = out;
            this.terminal = terminal;
        }

        String capturedOutput() {
            return out.toString(StandardCharsets.UTF_8);
        }
    }

    // =================================================================
    //  Banner
    // =================================================================

    @Test
    void elBannerSeMuestraAlArrancar() throws IOException {
        Harness h = harness("3\n", configPath -> App.EXIT_OK);
        h.menu.run();

        String output = h.capturedOutput();
        // El banner es ASCII art: las palabras "Excel" y "Merger" se
        // dibujan con caracteres '_', '|', '\\', '/' en varias lineas,
        // no aparecen como string literal. Lo que SI aparece literal en
        // el output del banner es la version (linea final del ASCII) y
        // la linea descriptiva debajo.
        assertThat(output).contains("v3.1.0");
        assertThat(output).contains("Fusion de exports ERP + Jira");
        // Verifica tambien que hay multiples lineas con caracter '|',
        // que es signature del ASCII art Figlet standard.
        long pipeLines = output.lines().filter(l -> l.contains("|")).count();
        assertThat(pipeLines).as("el ASCII art deberia tener varias lineas con '|'")
                .isGreaterThanOrEqualTo(5);
    }

    @Test
    void elBannerSeMuestraSoloUnaVezAunqueElMenuRecicleVarias() throws IOException {
        // Input: opcion invalida, opcion invalida, salir.
        Harness h = harness("xx\n99\n3\n", configPath -> App.EXIT_OK);
        h.menu.run();

        String output = h.capturedOutput();
        // El header descriptivo del banner es unico y especifico.
        int firstIdx = output.indexOf("Fusion de exports ERP + Jira");
        int secondIdx = output.indexOf("Fusion de exports ERP + Jira", firstIdx + 1);
        assertThat(firstIdx).as("banner debe aparecer al menos una vez").isGreaterThanOrEqualTo(0);
        assertThat(secondIdx).as("banner no debe repetirse al volver al menu").isEqualTo(-1);
    }

    // =================================================================
    //  Opcion 1: dispara la fusion
    // =================================================================

    @Test
    void opcion1DisparaElFlujoDeFusion() throws IOException {
        AtomicInteger calls = new AtomicInteger();
        Harness h = harness("1\n", configPath -> {
            calls.incrementAndGet();
            return App.EXIT_OK;
        });
        int exitCode = h.menu.run();

        assertThat(calls.get()).isEqualTo(1);
        assertThat(exitCode).isEqualTo(App.EXIT_OK);
    }

    @Test
    void opcion1OkDevuelveExitCero() throws IOException {
        Harness h = harness("1\n", configPath -> App.EXIT_OK);
        int exitCode = h.menu.run();
        assertThat(exitCode).isEqualTo(App.EXIT_OK);
    }

    @Test
    void opcion1ConErrorEsperaEnterAntesDeSalir() throws IOException {
        // Tras el "1" simulamos un enter para confirmar la pausa.
        Harness h = harness("1\n\n", configPath -> App.EXIT_RUNTIME);
        int exitCode = h.menu.run();

        assertThat(exitCode).isEqualTo(App.EXIT_RUNTIME);
        String output = h.capturedOutput();
        assertThat(output).contains("[ERROR] El proceso ha terminado con codigo 1");
        assertThat(output).contains("Pulsa Enter para continuar");
    }

    @Test
    void opcion1PreservaExitCodesEspecificos() throws IOException {
        int[] codes = {App.EXIT_CONFIG, App.EXIT_INPUT_INVALID, App.EXIT_OUTPUT_INVALID};
        for (int code : codes) {
            Harness h = harness("1\n\n", configPath -> code);
            int exitCode = h.menu.run();
            assertThat(exitCode).as("exit code esperado para %s", code).isEqualTo(code);
        }
    }

    // =================================================================
    //  Opcion 2: comprobador de discrepancias (v3.1.0)
    // =================================================================

    @Test
    void opcion2InvocaElRunnerInyectadoYVuelveAlMenu() throws IOException {
        // Tras pulsar 2 -> ejecuta runner -> espera Enter ("\n") -> menu
        // de nuevo -> 3 (salir).
        AtomicInteger calls = new AtomicInteger();
        Harness h = harness("2\n\n3\n",
                configPath -> {
                    throw new AssertionError("Opcion 1 no debe ejecutarse");
                },
                () -> {
                    calls.incrementAndGet();
                    return App.EXIT_OK;
                });

        int exitCode = h.menu.run();

        assertThat(exitCode).isEqualTo(App.EXIT_OK);
        assertThat(calls.get()).as("el runner de Opcion 2 debe haberse llamado una vez")
                .isEqualTo(1);
        // El menu se redibuja: la lista de opciones aparece >=2 veces.
        String output = h.capturedOutput();
        assertThat(output).contains("Comprobador de discrepancias contra CSV");
        assertThat(output).contains("Pulsa Enter para volver al menu");
    }

    @Test
    void opcion2NoTerminaElProgramaAunqueElRunnerFalle() throws IOException {
        // El runner de Opcion 2 devuelve EXIT_INPUT_INVALID. El menu
        // muestra el aviso, pide Enter, y vuelve al menu. Salimos con 3.
        AtomicInteger calls = new AtomicInteger();
        // Stdin: 2 -> runner falla -> Enter (aviso) -> Enter (volver) -> 3.
        Harness h = harness("2\n\n\n3\n",
                configPath -> App.EXIT_OK,
                () -> {
                    calls.incrementAndGet();
                    return App.EXIT_INPUT_INVALID;
                });

        int exitCode = h.menu.run();

        // Importante: aunque el runner devolvio !=OK, el menu termina con
        // EXIT_OK porque el usuario eligio Salir despues. Es decir, la
        // Opcion 2 NO termina el programa.
        assertThat(exitCode).isEqualTo(App.EXIT_OK);
        assertThat(calls.get()).isEqualTo(1);

        String output = h.capturedOutput();
        assertThat(output).contains("[AVISO]");
        assertThat(output).contains("codigo " + App.EXIT_INPUT_INVALID);
    }

    @Test
    void opcion2MultipleVecesSeguidasInvocaElRunnerVariasVeces() throws IOException {
        // Tres opcion 2 seguidas (con sus respectivos Enter para volver
        // al menu) y luego salir con 3.
        AtomicInteger calls = new AtomicInteger();
        Harness h = harness("2\n\n2\n\n2\n\n3\n",
                configPath -> App.EXIT_OK,
                () -> {
                    calls.incrementAndGet();
                    return App.EXIT_OK;
                });

        int exitCode = h.menu.run();

        assertThat(exitCode).isEqualTo(App.EXIT_OK);
        assertThat(calls.get()).isEqualTo(3);
    }

    // =================================================================
    //  Opcion 3: salir limpiamente
    // =================================================================

    @Test
    void opcion3TerminaConExitCero() throws IOException {
        Harness h = harness("3\n", configPath -> App.EXIT_OK);
        int exitCode = h.menu.run();
        assertThat(exitCode).isEqualTo(App.EXIT_OK);
    }

    @Test
    void opcion3NoEjecutaLaFusion() throws IOException {
        AtomicInteger calls = new AtomicInteger();
        Harness h = harness("3\n", configPath -> {
            calls.incrementAndGet();
            return App.EXIT_OK;
        });
        h.menu.run();
        assertThat(calls.get()).isZero();
    }

    // =================================================================
    //  Validacion de input
    // =================================================================

    @Test
    void inputInvalidoNumericoFueraDeRangoRepiteElMenu() throws IOException {
        Harness h = harness("4\n5\n99\n3\n", configPath -> App.EXIT_OK);
        int exitCode = h.menu.run();

        assertThat(exitCode).isEqualTo(App.EXIT_OK);
        String output = h.capturedOutput();
        // El mensaje "Opcion invalida" debe aparecer 3 veces (4, 5, 99).
        int count = 0;
        int idx = 0;
        while ((idx = output.indexOf(InteractiveMenu.INVALID_OPTION_MESSAGE, idx)) != -1) {
            count++;
            idx++;
        }
        assertThat(count).isEqualTo(3);
    }

    @Test
    void inputInvalidoLetrasRepiteElMenu() throws IOException {
        Harness h = harness("a\nx\nfoo\n3\n", configPath -> App.EXIT_OK);
        int exitCode = h.menu.run();

        assertThat(exitCode).isEqualTo(App.EXIT_OK);
        String output = h.capturedOutput();
        int count = 0;
        int idx = 0;
        while ((idx = output.indexOf(InteractiveMenu.INVALID_OPTION_MESSAGE, idx)) != -1) {
            count++;
            idx++;
        }
        assertThat(count).isEqualTo(3);
    }

    @Test
    void inputInvalidoVacioRepiteElMenu() throws IOException {
        // 2 lineas vacias y luego 3 (salir).
        Harness h = harness("\n\n3\n", configPath -> App.EXIT_OK);
        int exitCode = h.menu.run();

        assertThat(exitCode).isEqualTo(App.EXIT_OK);
        String output = h.capturedOutput();
        int count = 0;
        int idx = 0;
        while ((idx = output.indexOf(InteractiveMenu.INVALID_OPTION_MESSAGE, idx)) != -1) {
            count++;
            idx++;
        }
        assertThat(count).isEqualTo(2);
    }

    @Test
    void inputConEspaciosSeNormalizaConTrim() throws IOException {
        // "  1  " debe interpretarse como "1".
        AtomicInteger calls = new AtomicInteger();
        Harness h = harness("  1  \n", configPath -> {
            calls.incrementAndGet();
            return App.EXIT_OK;
        });
        int exitCode = h.menu.run();

        assertThat(exitCode).isEqualTo(App.EXIT_OK);
        assertThat(calls.get()).isEqualTo(1);
    }

    @Test
    void inputCon3PrecedidoDeEspaciosTerminaLimpiamente() throws IOException {
        Harness h = harness("   3\n", configPath -> App.EXIT_OK);
        int exitCode = h.menu.run();
        assertThat(exitCode).isEqualTo(App.EXIT_OK);
    }

    // =================================================================
    //  EOF / Ctrl+D
    // =================================================================

    @Test
    void stdinAgotadoSinSeleccionTerminaLimpiamente() throws IOException {
        // No hay ningun "\n" => readLine de JLine lanza EndOfFileException.
        // El menu lo trata como salida limpia (codigo 0).
        Harness h = harness("", configPath -> App.EXIT_OK);
        int exitCode = h.menu.run();
        assertThat(exitCode).isEqualTo(App.EXIT_OK);
    }
}

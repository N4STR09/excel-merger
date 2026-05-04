package com.excelmerger.cli;

import com.excelmerger.App;
import com.excelmerger.compare.CompareRunner;

import org.jline.reader.LineReader;
import org.jline.reader.LineReaderBuilder;
import org.jline.terminal.Terminal;
import org.jline.terminal.TerminalBuilder;

import java.io.IOException;
import java.io.PrintStream;
import java.util.function.IntSupplier;

/**
 * Menu interactivo de Excel Merger v3.0.0+. Sustituye a la antigua CLI
 * argumentada de v2.7.1 (eliminada).
 *
 * <p>Aparece SIEMPRE al arrancar {@link com.excelmerger.Main} y ofrece
 * tres opciones:</p>
 * <ol>
 *   <li>Fusion de Excel — ejecuta el flujo completo via {@link App#run(String)}.</li>
 *   <li>Comprobador de discrepancias contra CSV — invoca
 *       {@link CompareRunner} (v3.1.0). Tras ejecutar, vuelve al menu.</li>
 *   <li>Salir sin hacer nada — devuelve {@link App#EXIT_OK}.</li>
 * </ol>
 *
 * <p>Validacion de input: cualquier valor distinto de 1, 2 o 3 (incluido
 * vacio, espacios, letras, numeros fuera de rango) hace que se muestre
 * "Opcion invalida, introduce 1, 2 o 3" y se vuelva a pedir.</p>
 *
 * <p>Comportamiento al fallar la fusion (excepcion en Opcion 1, exit
 * code != 0): se muestra el error en consola, se espera enter del
 * usuario para que pueda leerlo (util en doble-click sobre el JAR en
 * Windows, donde la ventana cerraria) y se devuelve el exit code para
 * que {@code Main} llame a {@code System.exit(code)}.</p>
 *
 * <p>La Opcion 2 nunca termina el programa: si falla, muestra el error
 * y vuelve al menu. Asi el usuario puede corregir los CSV y reintentar
 * sin relanzar el JAR.</p>
 *
 * <p>La clase admite inyeccion de {@link Terminal}, {@link LineReader},
 * {@link PrintStream}, un runner para Opcion 1 y un runner para Opcion 2
 * (ambos como {@link IntSupplier}/lambda), de modo que los tests no
 * necesitan stdin real ni entrar al merger / comparador reales.</p>
 */
public final class InteractiveMenu {

    /** Texto exacto del aviso de input invalido. */
    static final String INVALID_OPTION_MESSAGE = "Opcion invalida, introduce 1, 2 o 3";

    /** Prompt que pide la opcion. */
    static final String PROMPT = "Selecciona una opcion [1-3]: ";

    private final LineReader reader;
    private final PrintStream out;
    private final IntSupplier optionOneRunner;
    private final IntSupplier optionTwoRunner;
    private final BannerPrinter bannerPrinter;

    /**
     * Constructor publico de uso normal en {@link com.excelmerger.Main}.
     * Crea un {@link Terminal} de sistema (stdin/stdout reales), enlaza
     * la Opcion 1 con {@link App#run(String)} y la Opcion 2 con
     * {@link CompareRunner#run()}.
     *
     * @throws IOException si no se puede inicializar el terminal JLine.
     */
    public InteractiveMenu() throws IOException {
        this(buildSystemTerminal(), App::run, () -> new CompareRunner().run(),
                buildSystemOut());
    }

    /**
     * Constructor para tests. Permite inyectar un terminal con stdin/stdout
     * controlados, runners falsos para Opcion 1 y Opcion 2, y un
     * PrintStream de captura.
     *
     * @param terminal       terminal JLine ya construido (real o de test).
     * @param optionOneRunner accion que ejecuta la Opcion 1. Devuelve el
     *                        exit code del merge.
     * @param optionTwoRunner accion que ejecuta la Opcion 2. Devuelve el
     *                        exit code del comparador.
     * @param out            stream donde escribir banner, menu y mensajes
     *                       (los tests pasan un ByteArrayOutputStream).
     */
    InteractiveMenu(Terminal terminal,
                    java.util.function.Function<String, Integer> optionOneRunner,
                    IntSupplier optionTwoRunner,
                    PrintStream out) {
        this.reader = LineReaderBuilder.builder().terminal(terminal).build();
        this.out = out;
        // null = config por defecto. La firma queda preparada por si se
        // amplia en el futuro a un sub-flujo que pregunte la ruta.
        this.optionOneRunner = () -> optionOneRunner.apply(null);
        this.optionTwoRunner = optionTwoRunner;
        this.bannerPrinter = new BannerPrinter(out);
    }

    private static Terminal buildSystemTerminal() throws IOException {
        return TerminalBuilder.builder().system(true).build();
    }

    /**
     * Construye un {@link PrintStream} sobre {@link System#out} con encoding
     * UTF-8 explicito. Asi evitamos que SpotBugs avise por
     * {@code DM_DEFAULT_ENCODING} y nos aseguramos de que el banner cyan
     * y los acentos del menu se escriben con el mismo encoding que usa
     * el resto del proyecto.
     */
    private static PrintStream buildSystemOut() {
        return new PrintStream(System.out, true, java.nio.charset.StandardCharsets.UTF_8);
    }

    /**
     * Ejecuta el bucle del menu. Devuelve el exit code que {@link com.excelmerger.Main}
     * pasara a {@link System#exit(int)}:
     * <ul>
     *   <li>Tras Opcion 1 OK: {@link App#EXIT_OK}.</li>
     *   <li>Tras Opcion 1 con error: el exit code de {@link App#run(String)},
     *       previa pausa para que el usuario lea el mensaje.</li>
     *   <li>Tras Opcion 3: {@link App#EXIT_OK}.</li>
     *   <li>La Opcion 2 NO devuelve: vuelve al menu independientemente del
     *       exit code interno (los errores se loguean y se muestran al
     *       usuario, pero el programa sigue para que pueda reintentar).</li>
     * </ul>
     */
    public int run() {
        bannerPrinter.print();
        while (true) {
            printMenu();
            String input = readLineSafe();
            String choice = input == null ? "" : input.trim();

            switch (choice) {
                case "1":
                    return runOptionOne();
                case "2":
                    runOptionTwo();
                    // Vuelve al menu en la siguiente iteracion del while.
                    break;
                case "3":
                    out.println();
                    out.println("Saliendo sin hacer nada. Hasta pronto.");
                    return App.EXIT_OK;
                default:
                    out.println();
                    out.println(INVALID_OPTION_MESSAGE);
                    out.println();
                    // Vuelve al menu en la siguiente iteracion del while.
                    break;
            }
        }
    }

    /**
     * Lee una linea del usuario con el prompt principal del menu. Si
     * JLine lanza {@code EndOfFileException} (Ctrl+D, stdin agotado en
     * tests) o {@code UserInterruptException} (Ctrl+C), devuelve
     * {@code "3"} para que el bucle salga limpiamente con codigo 0
     * (equivalente a haber elegido Salir).
     */
    private String readLineSafe() {
        return readLineWithPrompt(PROMPT);
    }

    /** Ejecuta la Opcion 1 y aplica el comportamiento OK / KO descrito en la clase. */
    private int runOptionOne() {
        out.println();
        int exitCode = optionOneRunner.getAsInt();
        if (exitCode != App.EXIT_OK) {
            // Pausa para que el usuario pueda leer el error antes de que
            // se cierre la ventana en doble-click sobre el JAR (Windows).
            out.println();
            out.println("[ERROR] El proceso ha terminado con codigo " + exitCode + ".");
            out.println("Pulsa Enter para continuar...");
            waitForEnter();
        }
        return exitCode;
    }

    /**
     * Ejecuta la Opcion 2 (comprobador de discrepancias). v3.1.0.
     * <p>A diferencia de la Opcion 1, NO termina el programa: tras la
     * ejecucion vuelve al menu. Si el comparador devuelve un exit code
     * distinto de OK se muestra un aviso y se espera Enter, pero el
     * bucle del menu continua.</p>
     */
    private void runOptionTwo() {
        int exitCode = optionTwoRunner.getAsInt();
        if (exitCode != App.EXIT_OK) {
            out.println();
            out.println("[AVISO] El comprobador ha terminado con codigo " + exitCode + ".");
            out.println("Pulsa Enter para continuar...");
            waitForEnter();
        }
        // Pausa breve sin error: dejar que el usuario lea el resumen
        // antes de redibujar el menu.
        out.println();
        out.println("Pulsa Enter para volver al menu...");
        waitForEnter();
    }

    /**
     * Espera a que el usuario pulse Enter. Si JLine lanza
     * {@code EndOfFileException} (Ctrl+D, stdin agotado en tests) o
     * {@code UserInterruptException} (Ctrl+C), devuelve igualmente sin
     * fallar: el caller seguira con su exit code. Delega en
     * {@link #readLineWithPrompt(String)} con prompt vacio, lo que
     * tambien evita un catch-vacio que dispararia la regla
     * {@code EmptyCatchBlock} de PMD.
     */
    private void waitForEnter() {
        // El valor devuelto se descarta a proposito: solo nos interesa
        // que la lectura bloquee hasta que el usuario pulse Enter (o
        // hasta que stdin se cierre / el usuario haga Ctrl+C, casos
        // ambos cubiertos por readLineWithPrompt sin lanzar excepcion).
        readLineWithPrompt("");
    }

    /**
     * Variante de {@link #readLineSafe} que acepta un prompt arbitrario
     * y devuelve la cadena leida, o {@code "3"} (interpretado como
     * "salir") si JLine senaliza EOF/Ctrl+C. Centraliza el manejo de
     * estas excepciones en un solo punto.
     */
    private String readLineWithPrompt(String prompt) {
        try {
            return reader.readLine(prompt);
        } catch (org.jline.reader.EndOfFileException | org.jline.reader.UserInterruptException eofOrInterrupt) {
            return "3";
        }
    }

    /** Muestra el menu con las tres opciones activas. */
    private void printMenu() {
        out.println();
        out.println("¿Que quieres hacer?");
        out.println();
        out.println("  1) Fusion de Excel");
        // v3.1.0: la Opcion 2 ya no es placeholder. Estilo normal.
        out.println("  2) Comprobador de discrepancias contra CSV");
        out.println("  3) Salir sin hacer nada");
        out.println();
    }
}

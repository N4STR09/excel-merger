package com.excelmerger.cli;

import com.excelmerger.Main;

import java.io.PrintStream;

/**
 * Imprime el banner ASCII art de bienvenida de Excel Merger v3.0.0.
 *
 * <p>El banner es estilo "standard Figlet" escrito a mano (no se usa la
 * libreria Figlet para no anadir mas dependencias), 67 columnas de ancho,
 * con la version integrada en el espacio libre del subrayado.</p>
 *
 * <p>Color cyan via secuencias ANSI ({@code ESC[36m} ... {@code ESC[0m}).
 * En terminales que no soportan ANSI las secuencias se imprimen como
 * caracteres literales; JLine resuelve esto antes a nivel de Terminal,
 * pero el banner se imprime sobre el {@link PrintStream} crudo para que
 * sea probable a su vez sin instanciar JLine. Si esto da problemas en
 * Windows muy antiguo, el fix es enviar el banner via
 * {@code AttributedString.print(terminal)} en lugar de {@code out.println}.</p>
 */
final class BannerPrinter {

    private static final String ANSI_CYAN = "\u001B[36m";
    private static final String ANSI_RESET = "\u001B[0m";

    /** Banner ASCII (sin la version, que se inyecta al imprimir). */
    private static final String[] BANNER_LINES = {
        "  _____                _   __  __                              ",
        " | ____|_  _____ ___  | | |  \\/  | ___ _ __ __ _  ___ _ __    ",
        " |  _| \\ \\/ / __/ _ \\ | | | |\\/| |/ _ \\ '__/ _` |/ _ \\ '__|   ",
        " | |___ >  < (_|  __/ | | | |  | |  __/ | | (_| |  __/ |      ",
        " |_____/_/\\_\\___\\___| |_| |_|  |_|\\___|_|  \\__, |\\___|_|      ",
        "                                           |___/     v%s   "
    };

    private static final String DESCRIPTION =
            " Fusion de exports ERP + Jira para cierre mensual";

    private final PrintStream out;

    BannerPrinter(PrintStream out) {
        this.out = out;
    }

    void print() {
        out.println();
        out.print(ANSI_CYAN);
        for (int i = 0; i < BANNER_LINES.length - 1; i++) {
            out.println(BANNER_LINES[i]);
        }
        // Ultima linea: insertamos la version. La version se trunca o
        // rellena segun longitud para no descuadrar el ancho. Para
        // versiones tipo "3.0.0" (5 chars) encaja perfecto. Si se sube
        // a algo mas largo (p.ej. "3.10.0" = 6), perdera 1 espacio
        // de padding pero seguira siendo legible.
        out.println(String.format(BANNER_LINES[BANNER_LINES.length - 1], Main.APP_VERSION));
        out.print(ANSI_RESET);
        out.println(DESCRIPTION);
        out.println();
    }
}

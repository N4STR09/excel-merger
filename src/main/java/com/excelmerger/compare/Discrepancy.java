package com.excelmerger.compare;

import java.util.Objects;

/**
 * Una discrepancia detectada entre un CSV del ERP y la hoja {@code Resultado}
 * del Excel del programa. Inmutable. v3.1.0.
 *
 * <h2>Tipos</h2>
 * <ul>
 *   <li>{@link Type#DIFERENCIA}: la combinacion (matricula, peticion,
 *       funcion) existe en ambos lados pero el valor numerico difiere.
 *       Tolerancia 0: cualquier diferencia, aunque sea de 0.01, es
 *       discrepancia.</li>
 *   <li>{@link Type#SOLO_CSV}: la combinacion existe en el CSV pero NO en
 *       Resultado. {@code pdclDeuda} se marca como
 *       {@code Double.NaN} (se exporta como celda vacia).</li>
 *   <li>{@link Type#SOLO_RESULTADO}: la combinacion existe en Resultado
 *       (con la matricula del CSV de origen) pero NO en el CSV.
 *       {@code realizadoHoras} se marca como {@code Double.NaN}.</li>
 * </ul>
 */
public final class Discrepancy {

    /** Tipo de discrepancia. */
    public enum Type {
        /** Existe en CSV y Resultado pero los valores difieren. */
        DIFERENCIA,
        /** Solo en el CSV. */
        SOLO_CSV,
        /** Solo en Resultado. */
        SOLO_RESULTADO
    }

    private final String origen;
    private final Type tipo;
    private final String peticion;
    private final String matricula;
    private final String funcion;
    private final double realizadoHoras;
    private final double pdclDeuda;

    /**
     * Constructor canonico. Se recomienda usar las factorias estaticas
     * {@link #difference}, {@link #onlyInCsv}, {@link #onlyInResultado}
     * para mayor claridad.
     */
    public Discrepancy(String origen, Type tipo, String peticion, String matricula,
                       String funcion, double realizadoHoras, double pdclDeuda) {
        this.origen = origen;
        this.tipo = tipo;
        this.peticion = peticion;
        this.matricula = matricula;
        this.funcion = funcion;
        this.realizadoHoras = realizadoHoras;
        this.pdclDeuda = pdclDeuda;
    }

    /**
     * Factoria para una discrepancia de tipo {@link Type#DIFERENCIA}.
     *
     * <p>Acepta {@link DiscrepancyKey} en lugar de los 3 strings sueltos
     * (matricula, peticion, funcion) por dos razones: (1) coherencia con
     * el dominio — esa terna es exactamente la clave de cruce y ya esta
     * encapsulada en {@link DiscrepancyKey}; (2) PMD-Design
     * {@code UseObjectForClearerAPI}: pasar 3+ strings posicionales hace
     * facil equivocarse de orden en la llamada.</p>
     */
    public static Discrepancy difference(String origen, DiscrepancyKey key,
                                         double realizadoHoras, double pdclDeuda) {
        return new Discrepancy(origen, Type.DIFERENCIA,
                key.getPeticion(), key.getMatricula(), key.getFuncion(),
                realizadoHoras, pdclDeuda);
    }

    /** Factoria para una discrepancia de tipo {@link Type#SOLO_CSV}. */
    public static Discrepancy onlyInCsv(String origen, DiscrepancyKey key,
                                        double realizadoHoras) {
        return new Discrepancy(origen, Type.SOLO_CSV,
                key.getPeticion(), key.getMatricula(), key.getFuncion(),
                realizadoHoras, Double.NaN);
    }

    /** Factoria para una discrepancia de tipo {@link Type#SOLO_RESULTADO}. */
    public static Discrepancy onlyInResultado(String origen, DiscrepancyKey key,
                                              double pdclDeuda) {
        return new Discrepancy(origen, Type.SOLO_RESULTADO,
                key.getPeticion(), key.getMatricula(), key.getFuncion(),
                Double.NaN, pdclDeuda);
    }

    public String getOrigen() {
        return origen;
    }

    public Type getTipo() {
        return tipo;
    }

    public String getPeticion() {
        return peticion;
    }

    public String getMatricula() {
        return matricula;
    }

    public String getFuncion() {
        return funcion;
    }

    /** Realizado Horas del CSV. {@link Double#NaN} para {@link Type#SOLO_RESULTADO}. */
    public double getRealizadoHoras() {
        return realizadoHoras;
    }

    /** PDCL + Deuda del Resultado. {@link Double#NaN} para {@link Type#SOLO_CSV}. */
    public double getPdclDeuda() {
        return pdclDeuda;
    }

    /**
     * Diferencia numerica {@code Realizado Horas - (PDCL + Deuda)}. Solo
     * tiene sentido para {@link Type#DIFERENCIA}; en los otros tipos
     * devuelve {@link Double#NaN}.
     */
    public double getDiferencia() {
        if (tipo != Type.DIFERENCIA) return Double.NaN;
        return realizadoHoras - pdclDeuda;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof Discrepancy)) return false;
        Discrepancy that = (Discrepancy) o;
        return Double.compare(that.realizadoHoras, realizadoHoras) == 0
                && Double.compare(that.pdclDeuda, pdclDeuda) == 0
                && Objects.equals(origen, that.origen)
                && tipo == that.tipo
                && Objects.equals(peticion, that.peticion)
                && Objects.equals(matricula, that.matricula)
                && Objects.equals(funcion, that.funcion);
    }

    @Override
    public int hashCode() {
        return Objects.hash(origen, tipo, peticion, matricula, funcion,
                realizadoHoras, pdclDeuda);
    }

    @Override
    public String toString() {
        return "Discrepancy{origen=" + origen + ", tipo=" + tipo
                + ", peticion=" + peticion + ", matricula=" + matricula
                + ", funcion=" + funcion + ", realizado=" + realizadoHoras
                + ", pdclDeuda=" + pdclDeuda + "}";
    }
}

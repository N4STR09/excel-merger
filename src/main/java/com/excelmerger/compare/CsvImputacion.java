package com.excelmerger.compare;

import java.util.Objects;

/**
 * Una imputacion individual leida de un CSV del ERP (formato
 * {@code <matricula>.CSV}). Inmutable.
 *
 * <p>Cada CSV trae una fila por imputacion: una persona (matricula) imputa
 * un numero de horas a una peticion concreta con una funcion concreta. Esta
 * clase representa esa fila tras la normalizacion descrita en
 * {@link CsvParser}: trim de strings, quitar la {@code J} prefijada de
 * {@code Petición}, decimal con coma -> double, descarte de {@code \u0000}.</p>
 *
 * <p>v3.1.0: introducida para la Opcion 2 del menu (comprobador de
 * discrepancias contra Excel del programa).</p>
 */
public final class CsvImputacion {

    private final String peticion;
    private final String matricula;
    private final String funcion;
    private final double realizadoHoras;

    /**
     * @param peticion       codigo de peticion sin la {@code J} prefijada
     *                       (p. ej. {@code "137791"}).
     * @param matricula      matricula del recurso (p. ej. {@code "90014"}).
     * @param funcion        codigo de funcion de 2 letras (p. ej. {@code "RE"}).
     * @param realizadoHoras horas realizadas, ya parseadas como double con
     *                       el separador decimal en punto.
     */
    public CsvImputacion(String peticion, String matricula, String funcion,
                         double realizadoHoras) {
        this.peticion = peticion;
        this.matricula = matricula;
        this.funcion = funcion;
        this.realizadoHoras = realizadoHoras;
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

    public double getRealizadoHoras() {
        return realizadoHoras;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof CsvImputacion)) return false;
        CsvImputacion that = (CsvImputacion) o;
        return Double.compare(that.realizadoHoras, realizadoHoras) == 0
                && Objects.equals(peticion, that.peticion)
                && Objects.equals(matricula, that.matricula)
                && Objects.equals(funcion, that.funcion);
    }

    @Override
    public int hashCode() {
        return Objects.hash(peticion, matricula, funcion, realizadoHoras);
    }

    @Override
    public String toString() {
        return "CsvImputacion{peticion=" + peticion
                + ", matricula=" + matricula
                + ", funcion=" + funcion
                + ", realizadoHoras=" + realizadoHoras
                + "}";
    }
}

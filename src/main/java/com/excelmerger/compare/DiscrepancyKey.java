package com.excelmerger.compare;

import java.util.Objects;

/**
 * Clave compuesta {@code (matricula, peticion, funcion)} usada para cruzar
 * lineas de CSV con filas de la hoja {@code Resultado} en el comprobador de
 * discrepancias (Opcion 2 del menu).
 *
 * <p>Inmutable. Todos los componentes son strings ya normalizados (trim,
 * sin la {@code J} prefijada en peticion, sin {@code \u0000}). El cruce
 * es exact-match string-a-string sobre los tres componentes.</p>
 *
 * <p>v3.1.0: introducida para la Opcion 2 del menu.</p>
 */
public final class DiscrepancyKey {

    private final String matricula;
    private final String peticion;
    private final String funcion;

    public DiscrepancyKey(String matricula, String peticion, String funcion) {
        this.matricula = matricula;
        this.peticion = peticion;
        this.funcion = funcion;
    }

    public String getMatricula() {
        return matricula;
    }

    public String getPeticion() {
        return peticion;
    }

    public String getFuncion() {
        return funcion;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof DiscrepancyKey)) return false;
        DiscrepancyKey that = (DiscrepancyKey) o;
        return Objects.equals(matricula, that.matricula)
                && Objects.equals(peticion, that.peticion)
                && Objects.equals(funcion, that.funcion);
    }

    @Override
    public int hashCode() {
        return Objects.hash(matricula, peticion, funcion);
    }

    @Override
    public String toString() {
        return "DiscrepancyKey{matricula=" + matricula
                + ", peticion=" + peticion
                + ", funcion=" + funcion
                + "}";
    }
}

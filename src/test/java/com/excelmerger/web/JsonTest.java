package com.excelmerger.web;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/** Tests del serializador JSON minimo de respuestas. */
class JsonTest {

    @Test
    void escapaComillasBarrasYControles() {
        assertThat(Json.quote(null)).isEqualTo("null");
        assertThat(Json.quote("texto")).isEqualTo("\"texto\"");
        assertThat(Json.quote("a\"b")).isEqualTo("\"a\\\"b\"");
        assertThat(Json.quote("a\\b")).isEqualTo("\"a\\\\b\"");
        assertThat(Json.quote("l1\nl2\tfin")).isEqualTo("\"l1\\nl2\\tfin\"");
        assertThat(Json.quote("\001")).isEqualTo("\"\\u0001\"");
        assertThat(Json.quote("caracter eñe")).isEqualTo("\"caracter eñe\"");
    }
}

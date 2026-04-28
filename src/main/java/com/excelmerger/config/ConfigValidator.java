package com.excelmerger.config;

import com.excelmerger.ConfigLoader;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/**
 * Validacion estatica del config.properties ANTES de abrir ningun archivo.
 *
 * <p>Acumula TODOS los errores detectados y los devuelve. {@code Main} decide
 * si aborta con exit code 2 (strictValidation=true) o solo los registra como
 * warnings en el {@code RunReport} (strictValidation=false).</p>
 *
 * <p>Esta clase es un ORQUESTADOR delgado: cada bloque lógico se delega a una
 * {@code *ConfigSection} package-private del mismo paquete. El refactor
 * estructural (Sesión F, v2.5.0) partió un único validador de 715 LoC en un
 * coordinador y siete secciones colaboradoras, sin cambio funcional ni de
 * API.</p>
 *
 * <p>Chequeos soportados:</p>
 * <ul>
 *   <li>Entrada/salida minimas.</li>
 *   <li>{@code merge.mode} valido.</li>
 *   <li>Perfiles: cada perfil listado en {@code profiles=} tiene al menos un
 *       criterio de deteccion ({@code detect.headers} o {@code detect.cellValue.*}).</li>
 *   <li>Columnas MES: tipo valido y campos obligatorios por tipo.</li>
 *   <li>Placeholders {@code {col:X}} de fórmulas MES referencian nombres existentes.</li>
 *   <li>Referencias a hojas: SUMIFS {@code from} y AGGREGATION {@code sourceSheet}
 *       deben apuntar a una hoja conocida (perfil, lookup, derivada o resultado de fusion).</li>
 *   <li>Lookups: cada hoja listada tiene datos.</li>
 *   <li>Derivadas: tipo valido y campos requeridos por tipo.</li>
 * </ul>
 */
public class ConfigValidator {

    private final ConfigLoader config;
    private final List<String> errors = new ArrayList<>();

    public ConfigValidator(ConfigLoader config) {
        this.config = config;
    }

    /**
     * Ejecuta todas las validaciones y devuelve la lista completa de errores.
     * Orden estable: se acumulan en el orden en el que se detectan, idéntico
     * al del validador monolítico previo a la Sesión F.
     */
    public List<String> validate() {
        errors.clear();

        IoConfigSection io = new IoConfigSection(config, errors);
        io.validateInputOutput();
        String mergeMode = io.validateMergeMode();
        io.validateOutputMode();

        ProfilesConfigSection profiles = new ProfilesConfigSection(config, errors);
        Set<String> profileSheetNames = profiles.validateProfiles();
        Set<String> lookupIds        = profiles.validateLookups();
        Set<String> derivedIds       = profiles.collectDerivedIds();

        Set<String> knownSheets = buildKnownSheets(
                mergeMode, profileSheetNames, lookupIds, derivedIds);

        new MesConfigSection(config, errors).validate(knownSheets);
        new DerivedConfigSection(config, errors).validate(knownSheets);
        new SummaryConfigSection(config, errors).validate(knownSheets);
        new OrphansConfigSection(config, errors).validate(knownSheets);
        new ResponsablesTablesConfigSection(config, errors).validate();

        return Collections.unmodifiableList(new ArrayList<>(errors));
    }

    private Set<String> buildKnownSheets(String mergeMode,
                                         Set<String> profileSheetNames,
                                         Set<String> lookupIds,
                                         Set<String> derivedIds) {
        Set<String> known = new LinkedHashSet<>();
        if ("APPEND_ROWS".equals(mergeMode)) {
            known.add(config.get("merge.resultSheetName", "Datos_Fusionados"));
        } else {
            // En SHEETS_SEPARATE cada perfil aporta una hoja con su sheetName
            known.addAll(profileSheetNames);
        }
        known.addAll(lookupIds);
        known.addAll(derivedIds);
        // La propia hoja MES tambien es referenciable
        if (config.getBoolean("mes.enabled", false)) {
            known.add(config.get("mes.sheetName", "MES"));
        }
        return known;
    }
}

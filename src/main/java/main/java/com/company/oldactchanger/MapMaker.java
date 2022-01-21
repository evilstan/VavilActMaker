package main.java.com.company.oldactchanger;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

public class MapMaker {

    Map<String, String> values;
    Map<String, String> addresses;

    public MapMaker() {
        values = new LinkedHashMap<>();
        addresses = new LinkedHashMap<>();
        setValues();
        setAddresses();
    }

    public Map<String, String> getValues() {
        return values;
    }

    public Map<String, String> getAddresses() {
        return addresses;
    }

    private void setValues() {
        values.put("ACT_TYPE", "");
        values.put("CURRENT_AXES_CELL", "");
        values.put("HEIGHT_CELL", "");
        values.put("DRAWING_CELL", "");
        values.put("REBAR_MATERIALS_CELL", "");
        values.put("REBAR_CERTIFICATE_CELL", "");
        values.put("CONCRETE_MATERIALS_CELL", "");
        values.put("CONCRETE_CERTIFICATES_CELL", "");
        values.put("REINFORCEMENT_DATE_CELL", "");
        values.put("CONCRETING_DATE_CELL", "");
        values.put("NEXT_HEIGHT_CELL", "");
        values.put("SECTION_CELL", "");
    }

    private void setAddresses() {
        addresses.put("ACT_TYPE", "E15");
        addresses.put("CURRENT_AXES_CELL", "E16");
        addresses.put("HEIGHT_CELL", "E17");
        addresses.put("DRAWING_CELL", "E18");
        addresses.put("REBAR_MATERIALS_CELL", "E19");
        addresses.put("REBAR_CERTIFICATE_CELL", "E20");
        addresses.put("CONCRETE_MATERIALS_CELL", "E21");
        addresses.put("CONCRETE_CERTIFICATES_CELL", "E22");
        addresses.put("REINFORCEMENT_DATE_CELL", "E25");
        addresses.put("CONCRETING_DATE_CELL", "E26");
        addresses.put("NEXT_HEIGHT_CELL", "E29");
        addresses.put("SECTION_CELL", "E30");
    }
}

package main.java.com.company.oldactchanger;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

public class RebarParser {
    private XSSFWorkbook rebarBook;
    private PassportSelector passportSelector;

    public RebarParser() {
        passportSelector = new PassportSelector();
    }

    public List<String> parse(String rebar1) {
        List<String> diameters = Arrays.asList(rebar1.split("Ø"));
        List<String> parsedDiameters = deleteCommas(diameters);
        return setPassport(parsedDiameters);
    }

    private List<String> setPassport(List<String> rebarList) {
        List<String> result = new ArrayList<>();
        Collections.sort(rebarList);
        for (String s : rebarList) {
            result.add("Ø" + s + "(" + passportSelector.passport(s) + ")");
        }
        return result;
    }

    private List<String> deleteCommas(List<String> diameters) {
        List<String> result = new ArrayList<>();
        for (String diameter : diameters) {
            if (diameter.contains("А240")) {
                diameter = diameter.trim();
                String s = diameter.split(" ")[0];
                s = s.replaceAll(",", "").trim();
                s = s.replaceAll(";", "").trim();
                result.add(s);
                continue;
            } else if (diameter.contains("А500")) continue;

            if (diameter.contains(",")) {
                String s = diameter.replaceAll(",", "").trim();
                s = s.replaceAll(";", "");
                result.add(s);
            } else result.add(diameter.trim());
        }
        return result;
    }

}

package main.java.com.company.newactcreator;

import main.java.com.company.helpers.FileChanger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellsFiller {
    private final FileChanger fileChanger;
    final int MAX_UPPER_CHARACTERS = 60;
    final int MAX_LOWER_CHARACTERS = 95;

    final String SECTION = "E30";
    final String CONSTRUCTION = "E15";
    final String AXES = "E16";
    final String CURRENT_HEIGHT = "E17";
    final String NEXT_HEIGHT = "E29";
    final String DRAWINGS = "E18";
    final String REBAR = "E19";
    final String REBAR_CERTIFICATES = "E20";
    final String CONCRETE = "E21";
    final String CONCRETE_CERTIFICATES = "E22";
    final String REBAR_DATE = "E25";
    final String CONCRETE_DATE = "E26";

    public CellsFiller() {
        fileChanger = new FileChanger();
    }


    public XSSFWorkbook fillCells(XSSFWorkbook workbook, ActDataObject dataObject) {
        XSSFSheet sheet = workbook.getSheetAt(0);
        fileChanger.setCellValue(sheet, CONSTRUCTION, dataObject.getConstruction());
        fileChanger.setCellValue(sheet, AXES, dataObject.getAxes());
        fileChanger.setCellValue(sheet, CURRENT_HEIGHT, dataObject.getCurrentHeight());
        fileChanger.setCellValue(sheet, DRAWINGS, dataObject.getDrawings());
        fileChanger.setCellValue(sheet, CONCRETE, dataObject.getConcrete());
        fileChanger.setCellValue(sheet, CONCRETE_CERTIFICATES, dataObject.getConcreteCertificates());
        fileChanger.setCellValue(sheet, REBAR_DATE, dataObject.getRebarDate());
        fileChanger.setCellValue(sheet, CONCRETE_DATE, dataObject.getConcreteDate());
        fileChanger.setCellValue(sheet, NEXT_HEIGHT, dataObject.getNextHeight());
        fileChanger.setCellValue(sheet, SECTION, dataObject.getSection());

        fileChanger.setCellValue(sheet, REBAR, formatRebar(dataObject.getRebar())[0]);
        fileChanger.setCellValue(sheet, REBAR_CERTIFICATES, formatRebar(dataObject.getRebar())[1]);
        return workbook;
    }

    private String[] formatRebar(String s) {

        if (s.length() > MAX_LOWER_CHARACTERS) {
            for (int j = MAX_UPPER_CHARACTERS; j >= 0; j--) {
                String splitter = Character.toString(s.charAt(j));
                if (splitter.equals(";") || splitter.equals(",")) {
                    return new String[]{s.substring(0, j + 1).trim(), "\n", s.substring(j + 2, s.length() - 1).trim()};
                }
            }
        }
        return new String[]{"", s};
    }
}

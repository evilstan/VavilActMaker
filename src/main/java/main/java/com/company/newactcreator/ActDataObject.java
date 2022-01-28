package main.java.com.company.newactcreator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.List;

public class ActDataObject {
    DataFormatter formatter;

    private int section;
    private String construction;
    private String axes;
    private String currentHeight;
    private String nextHeight;
    private String drawings;
    private String rebar;
    private String concrete;
    private String[] concreteCertificates;
    private String rebarDate;
    private String concreteDate;

    private final boolean isVertical;

    private final List<String> rowData;
    static int counter = 0;

    public ActDataObject(Row row) {
        counter++;
        System.out.println("Object number " + counter);
        formatter = new DataFormatter();
        rowData = new ArrayList<>();

        fillArray(row);
        arrayToVariables();
        isVertical = construction.contains("верт");
    }

    private void fillArray(Row row) {
        for (Cell cell : row) {
            rowData.add(formatter.formatCellValue(cell));
        }
    }

    private void arrayToVariables() {
        section = Integer.parseInt(rowData.get(0));
        construction = rowData.get(1);
        axes = rowData.get(2);
        currentHeight = rowData.get(3);
        nextHeight = rowData.get(4);
        drawings = rowData.get(5);
        rebar = rowData.get(6);
        concrete = rowData.get(7);
        concreteCertificates = stringToArray(rowData.get(8));
        rebarDate = rowData.get(9);
        rebarDate = formatDate(rowData.get(9));
        concreteDate = formatDate(rowData.get(10));
    }

    private String formatDate(String inputDate) {
        String[] list = inputDate.split("/");
        return list[1] + "." + list[0] + "." + list[2];
    }

    private String[] stringToArray(String s) {
        String[] result = s.split(",");

        for (int i = 0; i < result.length; i++) {
            result[i] = result[i].trim();
        }
        return result;
    }

    public boolean isVertical() {
        return isVertical;
    }

    public String getSection() {
        return "секції №" + section;
    }

    public String getConstruction() {
        if (construction.contains("верт")) {
            return "вертикальних конструкцій";
        } else if (construction.contains("плит")) {
            return "плити перекриття";
        }
        return construction;
    }

    public String getAxes() {
        return axes;
    }

    public String getCurrentHeight() {
        return currentHeight;
    }

    public String getNextHeight() {
        if (isVertical) {
            return nextHeight;
        }
        return currentHeight;
    }

    public String getDrawings() {
        return drawings;
    }

    public String getRebar() {
        return rebar;
    }

    public String getConcrete() {
        return concrete;
    }

    public String getConcreteCertificates() {
        StringBuilder certificates = new StringBuilder();
        String buf = ", ";
        for (String cert : concreteCertificates) {
            certificates.append(cert += buf);
        }
        String result = certificates.toString().trim();
        result = result.substring(0, result.length() - 1);
        return result;
    }

    public String getRebarDate() {
        return rebarDate;
    }

    public String getConcreteDate() {
        return concreteDate;
    }
}

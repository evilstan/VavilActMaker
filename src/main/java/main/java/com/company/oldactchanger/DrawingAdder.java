package main.java.com.company.oldactchanger;

import main.java.com.company.helpers.FileChanger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

public class DrawingAdder {
    private XSSFWorkbook drawings;
    private FileChanger fileChanger;


    public DrawingAdder() {
        try {
            drawings = new XSSFWorkbook(new File("C:\\vavil\\draws.xlsx"));
            fileChanger = new FileChanger();
        } catch (IOException | InvalidFormatException e) {
            System.out.println("Error opening drawing file " + e.getMessage());
        }
    }

    public XSSFWorkbook addDrawings(XSSFWorkbook workbook) {
        XSSFSheet actSheet = workbook.getSheetAt(0);
        String constructionName = fileChanger.getCellValue(actSheet, "E15");
        String height = fileChanger.getCellValue(actSheet, "E17");
        String section = fileChanger.getCellValue(actSheet, "E30");
        String drawings = getDrawings(constructionName, height, section);
        fileChanger.setCellValue(actSheet, "E18", drawings);
        return workbook;
    }

    private String getDrawings(String constructionName, String height, String section) {
        XSSFSheet drawingSheet = drawings.getSheetAt(0);
        XSSFRow row;
        String constructionCell;
        double heightCell;
        String sectionCell;
        String cryptoCell;
        String sheetCell;

        for (int i = 2; i < 85; i++) {
            row = drawingSheet.getRow(i);
            constructionCell = fileChanger.getCellValue(drawingSheet, ("A" + i));
            heightCell = fileChanger.getCellDoubleValue(drawingSheet, ("B" + i));
            sectionCell = fileChanger.getCellValue(drawingSheet, ("C" + i));

            if (constructionName.toLowerCase(Locale.ROOT).contains(constructionCell)) {
                if (height.contains(Double.toString(heightCell))) {
                    System.out.println("constructionName " + constructionName + " constructionCell= " + constructionCell);
                    System.out.println("height = " + height + " heightCell= " + heightCell);
                    System.out.println("section = " + section + " sectionCell = " + sectionCell);

                    if (section.equals(sectionCell)) {
                        System.out.println();
                        cryptoCell = fileChanger.getCellValue(drawingSheet, ("D" + i));
                        sheetCell = fileChanger.getCellValue(drawingSheet, ("E" + i));
                        return cryptoCell + "; арк. " + sheetCell;
                    }
                }
            }
        }
        System.out.println(constructionName + " " + height+" section " + section + " drawing not found");
        return "";
    }
}

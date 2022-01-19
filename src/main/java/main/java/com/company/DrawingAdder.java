package main.java.com.company;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

public class DrawingAdder {
    XSSFWorkbook drawings;
    FileChanger fileChanger;

    public DrawingAdder() {
        try {
            drawings = new XSSFWorkbook(new File("C:\\vavil\\draws.xlsx"));
            fileChanger = new FileChanger();
        } catch (IOException | InvalidFormatException e) {
            System.out.println("Error opening drawing file " + e.getMessage());
        }
    }

    public XSSFSheet addDrawings(XSSFSheet actSheet) {

    }

    private String getDrawings(String constructionName, String height, int section) {
        XSSFSheet drawingSheet = drawings.getSheetAt(0);
        XSSFRow row;
        String constructionCell;
        String heightCell;
        String sectionCell;
        String codeCell;
        String drawingCell;

        for (int i = 1; i < 45; i++) {
            row = drawingSheet.getRow(i);
            constructionCell = fileChanger.getCellValue(drawingSheet, ("A" + i));
            heightCell = fileChanger.getCellValue(drawingSheet, ("B" + i));
            sectionCell = fileChanger.getCellValue(drawingSheet, ("C" + i));
            codeCell = fileChanger.getCellValue(drawingSheet, ("D" + i));
            drawingCell = fileChanger.getCellValue(drawingSheet, ("E" + i));

            if (constructionName.contains(constructionCell) &&
                    height.contains(heightCell) &&
                    section == Integer.parseInt(sectionCell)) {

            }


        }
    }
}

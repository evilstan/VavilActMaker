package main.java.com.company;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Locale;

public class FileChanger {
    DataFormatter formatter;

    public FileChanger() {
        formatter = new DataFormatter();
    }

    private void sheets(XSSFWorkbook workbook) {
        int sheetsNumber = workbook.getNumberOfSheets();
        System.out.println("Sheet " + sheetsNumber);
        for (int i = 1; i < sheetsNumber; i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);
            setCellValue(sheet, "A2", "HelloWorld");
        }
    }

    public String getCellValue(XSSFSheet sheet, String address) {
        int[] indexes = cellAddressConverter(address);

        XSSFRow row = sheet.getRow(indexes[0]);
        XSSFCell cell = row.getCell(indexes[1]);

        return formatter.formatCellValue(cell);
    }

    public void setCellValue(XSSFSheet sheet, String address, String value) {
        int[] indexes = cellAddressConverter(address);


        XSSFRow row = sheet.getRow(indexes[0]);
        XSSFCell cell = row.getCell(indexes[1]);
        cell.setCellValue(value);
    }

    public void addRow(XSSFSheet sheet, int startRow) {
        sheet.shiftRows(startRow, sheet.getLastRowNum(), 1);
        sheet.createRow(startRow);
    }

    public void copyRows(XSSFSheet sheet, int startRow) {
        sheet.shiftRows(startRow, sheet.getLastRowNum(), 1);
        sheet.copyRows(0, 0, 1, new CellCopyPolicy());
    }


    private int[] cellAddressConverter(String address) {
        address = address.replaceAll("(?<=[A-Z])(?=[A-Z])|(?<=[a-z])(?=[A-Z])|(?<=\\D)$", "1");
        String[] atoms = address.split("(?<=\\D)(?=\\d)|(?<=\\d)(?=\\D)");

        int row = Integer.parseInt(atoms[1]) - 1;
        int column = letterToColumnIndex(atoms[0].toLowerCase(Locale.ROOT).toCharArray()[0]) - 1;

        if (column < 0) {
            throw new NumberFormatException("Wrong input");
        }

        return new int[]{row, column};
    }

    private int letterToColumnIndex(char letter) {
        return (int) letter - 96;
    }


}

package main.java.com.company.helpers;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.ParseException;
import java.text.SimpleDateFormat;
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

    public double getCellDoubleValue(XSSFSheet sheet, String address){
        int[] indexes = cellAddressConverter(address);

        XSSFRow row = sheet.getRow(indexes[0]);
        XSSFCell cell = row.getCell(indexes[1]);
        return cell.getNumericCellValue();
    }

    public void setCellValue(XSSFSheet sheet, String address, String value) {
        int[] indexes = cellAddressConverter(address);


        XSSFRow row = sheet.getRow(indexes[0]);
        XSSFCell cell = row.getCell(indexes[1]);

        if (isNumeric(value)) {
            cell.setCellValue(Double.parseDouble(value));
        } else if (isDate(value)) {
            setDate(cell, value);
        } else cell.setCellValue(value);
    }

    public static boolean isDate(String str) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yy");
        dateFormat.setLenient(false);
        try {
            dateFormat.parse(str.trim());
        } catch (ParseException pe) {
            return false;
        }
        return true;
    }

    public static boolean isNumeric(String str) {
        try {
            Double.parseDouble(str);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private void setDate(Cell cell, String str) {
        try {
            SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yy");
            dateFormat.setLenient(false);
            cell.setCellValue(dateFormat.parse(str.trim()));
        } catch (Exception e) {
            cell.setCellValue(str);
        }
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

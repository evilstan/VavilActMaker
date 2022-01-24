package main.java.com.company.oldactchanger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class PassportSelector {
    private List<Long> passport6;
    private List<Long> passport8;
    private List<Long> passport10;
    private List<Long> passport12;
    private List<Long> passport16;
    private List<Long> passport18;
    private List<Long> passport20;
    private List<Long> passport25;
    private List<Long> passport28;
    private List<Long> passport32;

    XSSFWorkbook passportBook;
    Row diameterRow;

    public PassportSelector() {
        try {
            passportBook = new XSSFWorkbook(new File("C:\\vavil\\rebar.xlsx"));
        } catch (IOException e) {
            System.out.println("Error opening file " + e.getMessage());
        } catch (InvalidFormatException e) {
            System.out.println("File not excel format " + e.getMessage());
        }
        fillArrays(passportBook.getSheetAt(1));
    }

    public long passport(String diameter) {
        return getPassport(diameter);
    }

    private void fillArrays(XSSFSheet sheet) {
        passport6 = new ArrayList<>();
        passport8 = new ArrayList<>();
        passport10 = new ArrayList<>();
        passport12 = new ArrayList<>();
        passport16 = new ArrayList<>();
        passport20 = new ArrayList<>();
        passport25 = new ArrayList<>();
        passport28 = new ArrayList<>();
        passport32 = new ArrayList<>();

        for (int i = 0; i < 10; i++) {
            diameterRow = sheet.getRow(i);
            int diameter = (int) diameterRow.getCell(0).getNumericCellValue();
            switch (diameter) {
                case 6 -> passport6 = putPassport();
                case 8 -> passport8 = putPassport();
                case 10 -> passport10 = putPassport();
                case 12 -> passport12 = putPassport();
                case 16 -> passport16 = putPassport();
                case 18 -> passport18 = putPassport();
                case 20 -> passport20 = putPassport();
                case 25 -> passport25 = putPassport();
                case 28 -> passport28 = putPassport();
                case 32 -> passport32 = putPassport();
            }
        }
    }

    private List<Long> putPassport() {
        List<Long> result = new ArrayList<>();
        for (Cell cell : diameterRow) {
            if (cell.getColumnIndex() == 0) continue;
            try {
                if (cell.getNumericCellValue() == 0) throw new Exception();
                if (cell.getNumericCellValue() == 2147483647) continue;
                result.add((long)cell.getNumericCellValue());
            } catch (Exception e) {
                System.out.println("Diameter " + diameterRow.getCell(0).getNumericCellValue() + "passports" + result);

                return result;
            }
        }
        System.out.println("Diameter " + diameterRow.getCell(0).getNumericCellValue() + "passports" + result);
        return result;
    }

    private Long getPassport(String diameter) {
        switch (diameter) {
            case "6":
                return passport6.get((int) (Math.random() * passport6.size() - 1));
            case "8":
                return passport8.get((int) (Math.random() * passport8.size() - 1));
            case "10":
                return passport10.get((int) (Math.random() * passport10.size() - 1));
            case "12":
                return passport12.get((int) (Math.random() * passport12.size() - 1));
            case "18":
                return passport18.get((int) (Math.random() * passport18.size() - 1));
            case "16":
                return passport16.get((int) (Math.random() * passport16.size() - 1));
            case "20":
                return passport20.get((int) (Math.random() * passport20.size() - 1));
            case "25":
                return passport25.get((int) (Math.random() * passport25.size() - 1));
            case "28":
                return passport28.get((int) (Math.random() * passport28.size() - 1));
            case "32":
                return passport32.get((int) (Math.random() * passport32.size() - 1));
            default:
                System.out.println("Default");
        }
        return (long)(Math.random() * 1000000);
    }

}

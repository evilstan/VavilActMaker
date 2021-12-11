package main.java.com.company;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.*;

public class FileOpener {
    private String path;

    public FileOpener(String path) {
        this.path = path;
        try {
            workbooks(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public List<XSSFWorkbook> workbooks(String path) throws Exception {
        final File folder = new File(path);
        List<File> files = parseFile(folder);
        List<XSSFWorkbook> workbooks = new ArrayList<>();
        for (File file : files) {
            workbooks.add(new XSSFWorkbook(file));
        }

        return workbooks;
    }

    private List<File> parseFile(final File folder) {
        List<File> files = new ArrayList<>();
        for (final File excelFile : folder.listFiles()) {
            if (excelFile.isDirectory()) {
                parseFile(excelFile);
            } else {
                System.out.println(excelFile.getName());
                if (checkExtension(excelFile)) {
                    files.add(excelFile);
                }
            }
        }
        return files; //TODO fix
    }

    private void openFile(String path) {

        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook("c://1.xlsx");
        } catch (IOException e) {
            System.out.println("File " + path + " not found");
        }

        XSSFSheet wSheet = null;
        int index = 0;
        try {
            wSheet = workbook.getSheetAt(index);
        } catch (Exception e) {
            System.out.println("Sheet " + path + " not found");
        }

        XSSFRow row = wSheet.getRow(0);
        XSSFCell cell = row.getCell(0);
        System.out.println("value = " + cell.getRichStringCellValue().getString());

        final File folder = new File("C:\\1.xlsx");
        parseFile(folder);
    }

    private boolean checkExtension(File file) {
        String extension = "";

        int i = file.getName().lastIndexOf('.');
        if (i > 0) {
            extension = file.getName().substring(i + 1);
        }
        return extension.toLowerCase(Locale.ROOT).contains("xls");
    }

    private int[] cellAddressConverter(String address) {

        address = address.replaceAll("(?<=[A-Z])(?=[A-Z])|(?<=[a-z])(?=[A-Z])|(?<=\\D)$", "1");
        String[] atoms = address.split("(?<=\\D)(?=\\d)|(?<=\\d)(?=\\D)");

        int row = Integer.parseInt(atoms[1]);
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

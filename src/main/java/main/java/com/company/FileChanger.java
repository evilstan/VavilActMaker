package main.java.com.company;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public class FileChanger {
    private String path;
    private FileOpener fileOpener;


    public FileChanger(String path) {
        this.path = path;
        fileOpener = new FileOpener(path);

    }

    private void books() throws Exception {
        List<XSSFWorkbook> workbooks = fileOpener.workbooks(path);
        for (XSSFWorkbook workbook : workbooks) {
            sheets(workbook);
        }
    }

    private void sheets(XSSFWorkbook workbook) {
        int sheetsNumber = workbook.getNumberOfSheets();
        for (int i = 1; i < sheetsNumber; i++) {
            Sheet currentSheet = workbook.getSheetAt(i);
        }
    }

    private void changeSheet(Sheet sheet) {

    }
}

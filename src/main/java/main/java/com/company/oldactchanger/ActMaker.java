package main.java.com.company.oldactchanger;

import main.java.com.company.helpers.FileChanger;
import main.java.com.company.helpers.FileOpener;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ActMaker {

    final String HOR_TEMPLATE_FILENAME = "hor_template.xlsx";
    final String VERT_TEMPLATE_FILENAME = "vert_template.xlsx";
    Map<String, String> address;
    Map<String, String> value;

    private final String filePath;
    private String templatePath;
    private String filename;

    private List<String> filenames;

    private List<XSSFWorkbook> workbooks;
    private XSSFWorkbook horTemplate;
    private XSSFWorkbook vertTemplate;

    private FileChanger fileChanger;

    public ActMaker(String filePath) {
        this.filePath = filePath;
        init();
    }

    private void init() {
        fileChanger = new FileChanger();
        templatePath = "C:\\vavil\\";
        openFiles();
        writeDataToAddressMap();
        changeBooks();
    }

    private void openFiles() {
        FileOpener fileOpener = new FileOpener(filePath);
        try {
            fileOpener.openBooks(filePath);
            workbooks = fileOpener.getWorkbooks();
            System.out.println("workbooks size " + workbooks.size());

            filenames = fileOpener.getFilenames();
            System.out.println("filenames size " + filenames.size());

            horTemplate = fileOpener.openBook(templatePath, HOR_TEMPLATE_FILENAME);
            vertTemplate = fileOpener.openBook(templatePath, VERT_TEMPLATE_FILENAME);
        } catch (Exception e) {
            System.out.println("Exception while opening " + e.getMessage());
        }
    }

    private void writeDataToAddressMap() {
        MapMaker mapMaker = new MapMaker();
        value = mapMaker.getValues();
        address = mapMaker.getAddresses();
    }

    private void changeBooks() {
        DrawingAdder drawingAdder = new DrawingAdder();
        for (XSSFWorkbook workbook : workbooks) {
            filename = filenames.get(workbooks.indexOf(workbook));
/*            writeDataToValueMap(workbook.getSheetAt(0));
            writeMapToBook();*/
            workbook = drawingAdder.addDrawings(workbook);

            try {
                saveFile(workbook, filename);
                workbook.close();
            } catch (IOException e) {
                System.out.println("Error saving file " + e.getMessage());
            }
        }
    }

    private void writeDataToValueMap(XSSFSheet sheet) {
        for (String key : value.keySet()) {
            String address = this.address.get(key);
            String value = fileChanger.getCellValue(sheet, address);
            this.value.put(key, value);
        }
        value.put("REINFORCEMENT_DATE_CELL", formatDate(value.get("REINFORCEMENT_DATE_CELL")));
        value.put("CONCRETING_DATE_CELL", formatDate(value.get("CONCRETING_DATE_CELL")));
        value.put("NEXT_HEIGHT_CELL", formatHeight(value.get("NEXT_HEIGHT_CELL")));
        value.put("HEIGHT_CELL", formatHeight(value.get("HEIGHT_CELL")));
    }

    private String formatDate(String inputDate) {
        String[] list = inputDate.split("/");
        return list[1] + "." + list[0] + "." + list[2];
    }

    private String formatHeight(String input) {
        String height = input.trim();
        height = height.replace("\\.", "\\,");
        if (height.endsWith("м") || height.endsWith(".")) {
            return input;
        } else {
            return height + " м.";
        }
    }

    private void writeMapToBook() {
        XSSFWorkbook workbook;
        XSSFSheet sheet;

        if (value.get("ACT_TYPE").contains("вертикал")) {
            workbook = vertTemplate;
        } else {
            workbook = horTemplate;
        }

        sheet = workbook.getSheetAt(0);

        for (String key : address.keySet()) {
            String adr = address.get(key);
            String val = value.get(key);
            System.out.println("Saving " + val + " to " + adr);
            fileChanger.setCellValue(sheet, adr, val);
        }


    }

    private void saveFile(XSSFWorkbook workbook, String filename) throws IOException {
        String savePath = filename.replace("old", "new");
        FileOutputStream saveStream = new FileOutputStream(savePath);
        XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
        workbook.write(saveStream);
    }

}

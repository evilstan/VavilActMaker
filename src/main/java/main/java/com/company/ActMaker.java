package main.java.com.company;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ActMaker {

    final String HOR_TEMPLATE_FILENAME = "hor_template.xlsx";
    final String VERT_TEMPLATE_FILENAME = "vert_template.xlsx";
    final String NEW = "\\new\\";
    final String TEMPLATES = "\\templates\\";
    Map<String, String> address;
    Map<String, String> value;

    private String filePath;
    private String templatePath;
    private List<String> filenames;

    private List<XSSFWorkbook> workbooks;
    private XSSFWorkbook horTemplate;
    private XSSFWorkbook vertTemplate;
    private XSSFWorkbook temp;

    private FileChanger fileChanger;

    public ActMaker(String filePath) {
        this.filePath = filePath;
        init();
    }

    private void init() {
        fileChanger = new FileChanger();
        templatePath = "C:\\vavil\\";
        openFiles();
        initMaps();
        // changeBooks();
        setValuesMap(workbooks.get(0).getSheetAt(0));
        writeNewBook();
    }

    private void openFiles() {
        FileOpener fileOpener = new FileOpener(filePath);
        try {
            workbooks = fileOpener.openBooks(filePath);
            filenames = fileOpener.getFilenames();

            horTemplate = fileOpener.openBook(templatePath, HOR_TEMPLATE_FILENAME);
            vertTemplate = fileOpener.openBook(templatePath, VERT_TEMPLATE_FILENAME);
        } catch (Exception e) {
            System.out.println("Exception while opening " + e.getMessage());
        }
    }

    private void initMaps() {
        MapMaker mapMaker = new MapMaker();
        value = mapMaker.getValues();
        address = mapMaker.getAddresses();
    }

    private void changeBooks() {
        for (XSSFWorkbook workbook : workbooks) {
            setValuesMap(workbook.getSheetAt(0));
        }
    }

    private void setValuesMap(XSSFSheet sheet) {
        for (String key : value.keySet()) {
            String address = this.address.get(key);
            String value = fileChanger.getCellValue(sheet, address);
            this.value.put(key, value);
        }
        value.put("REINFORCEMENT_DATE_CELL", formatDate(value.get("REINFORCEMENT_DATE_CELL")));
        value.put("CONCRETING_DATE_CELL", formatDate(value.get("CONCRETING_DATE_CELL")));

        for (String key : value.keySet()) {
            System.out.println(key + " = " + value.get(key));
        }
    }

    private String formatDate(String inputDate) {
        String[] list = inputDate.split("/");
        return list[1] + "." + list[0] + "." + list[2];
        //return inputDate.replace("\\/", "\\a");
    }

    private void writeNewBook() {
        //XSSFWorkbook newBook = new XSSFWorkbook(new File(path + NEW + "file.xlsx"));
        XSSFWorkbook workbook = vertTemplate;

        for (String key : address.keySet()) {
            XSSFSheet sheet = vertTemplate.getSheetAt(0);
            String adr = address.get(key);
            String val = value.get(key);
            fileChanger.setCellValue(sheet, adr, val);
        }

        System.out.println(filenames.get(0));
        try {
            saveFile(vertTemplate, filenames.get(0));
            System.out.println();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    private void saveFile(XSSFWorkbook workbook, String filename) throws IOException {
        String savePath = templatePath + NEW + filename;
        FileOutputStream saveStream = new FileOutputStream(savePath);
        workbook.write(saveStream);
    }

}

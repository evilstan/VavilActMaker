package main.java.com.company.newactcreator;

import main.java.com.company.helpers.FileOpener;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ActFileCreator {
    private final FileOpener fileOpener;
    private final CellsFiller cellsFiller;

    final String HOR_TEMPLATE_FILENAME = "hor_template.xlsx";
    final String VERT_TEMPLATE_FILENAME = "vert_template.xlsx";
    final String REG_FILENAME = "registry.xlsx";
    final String PARENT_PATH = "C:\\vavil\\";

    private XSSFWorkbook horTemplate;
    private XSSFWorkbook vertTemplate;
    private XSSFWorkbook registry;

    private final List<ActDataObject> dataObjects;

    public ActFileCreator(List<ActDataObject> dataObjects) {
        this.dataObjects = dataObjects;
        fileOpener = new FileOpener();
        cellsFiller = new CellsFiller();
        try {
            openFiles();
        } catch (Exception e) {
            System.out.println("Failed to open files: " + e.getMessage());
        }
    }

    //TODO add registry
    private void openFiles() throws Exception {
        horTemplate = fileOpener.openBook(PARENT_PATH, HOR_TEMPLATE_FILENAME);
        vertTemplate = fileOpener.openBook(PARENT_PATH, VERT_TEMPLATE_FILENAME);
        // registry = fileOpener.openBook(PARENT_PATH, REG_FILENAME);
    }

    public void makeActs() {
        for (ActDataObject dataObject : dataObjects) {
            XSSFWorkbook newAct = cellsFiller.fillCells(
                    dataObject.isVertical() ? vertTemplate : horTemplate
                    , dataObject);
            String filename = generateFileName(dataObject);
            String path = generateFilePath(dataObject);
            saveFile("C:\\vavil\\Fresh\\", filename, newAct);
        }
    }


    private String generateFileName(ActDataObject dataObject) {
        String result = "ПП ";
        if (dataObject.isVertical()) {
            result = "Вертикал ";
        }
        result += dataObject.getSection() + " з відм. " + dataObject.getCurrentHeight() + ".xlsx";
        return result;
    }

    //TODO correctly generate path!
    private String generateFilePath(ActDataObject dataObject) {
        return "C:\\vavil\\Fresh\\" + dataObject.getSection() + "\\" + dataObject.getCurrentHeight() + "\\";
    }

    private void saveFile(String path, String filename, XSSFWorkbook workbook) {

        try {
            FileOutputStream saveStream = new FileOutputStream(path + filename);
            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            workbook.write(saveStream);
        } catch (IOException e) {
            System.out.println("Error saving file " + e.getMessage());
        }

/*        try {
            workbook.close();
        } catch (IOException e) {
            System.out.println("Error closing the book: " + e.getMessage());
        }*/
    }

}

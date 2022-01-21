package main.java.com.company.newactcreator;

import main.java.com.company.helpers.FileOpener;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;

public class ActCreator {
    final String DB_FILENAME = "database.xlsx";
    final String PARENT_PATH = "C:\\vavil\\";

    private XSSFWorkbook database;
    private XSSFSheet sheet;

    private List<ActDataObject> dataObjects;

    private FileOpener fileOpener;
    private ActFileCreator actFileCreator;

    public ActCreator() throws Exception {
        fileOpener = new FileOpener();
        openFiles();
        createDataObjects();
        createFiles();
    }

    private void openFiles() throws Exception {
        database = fileOpener.openBook(PARENT_PATH, DB_FILENAME);
    }

    private void createDataObjects() {
        dataObjects = new ArrayList<>();

        for (int i = 0; i < 44; i++) {
            dataObjects.add(new ActDataObject(database.getSheetAt(0).getRow(i)));
        }

/*        for (Row row : database.getSheetAt(0)) {
            if (row == null) return;
            dataObjects.add(new ActDataObject(row));
        }*/
    }

    private void createFiles() {
        actFileCreator = new ActFileCreator(dataObjects);

    }


}

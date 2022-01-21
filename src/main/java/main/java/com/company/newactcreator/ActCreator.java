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
    final int LAST_ROW = 45;

    private XSSFWorkbook database;

    private List<ActDataObject> dataObjects;

    private final FileOpener fileOpener;

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

        for (int i = 0; i < LAST_ROW; i++) {
            dataObjects.add(new ActDataObject(database.getSheetAt(0).getRow(i)));
        }
    }

    private void createFiles() {
        ActFileCreator actFileCreator = new ActFileCreator(dataObjects);
        actFileCreator.makeActs();
    }


}

package main.java.com.company.helpers;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

public class FileOpener {
    private List<File> files;
    private List<XSSFWorkbook> workbooks;
    private List<String> filenames;
    private List<File> usedFiles;

    public FileOpener(String path) {
        init();
        try {
            openBooks(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public FileOpener() {
        init();
    }

    private void init() {
        files = new ArrayList<>();
        workbooks = new ArrayList<>();
        filenames = new ArrayList<>();
        usedFiles = new ArrayList<>();
    }

    public XSSFWorkbook openBook(String path, String filename) throws Exception {
        return new XSSFWorkbook(new File(path + filename));
    }

    public void openBooks(String path) throws Exception {
        final File folder = new File(path);
        List<File> files = parseFile(folder);
        System.out.println("Files = " + files.size());
        workbooks.clear();
        convertFileToExcel();
    }

    private void convertFileToExcel() throws IOException, InvalidFormatException {

        for (File file : files) {

            System.out.println("Books  = " + workbooks.size());
            FileInputStream fileInputStream = new FileInputStream(file);
            workbooks.add(new XSSFWorkbook(fileInputStream));
            fileInputStream.close();
        }
    }

    int count = 0;

    private List<File> parseFile(final File folder) {
        for (final File excelFile : Objects.requireNonNull(folder.listFiles())) {
            if (usedFiles.contains(excelFile)) {
                System.out.println("Used file found");
                continue;
            }
            usedFiles.add(excelFile);
            if (excelFile.isDirectory()) {
                parseFile(excelFile);
            } else {
                System.out.println(excelFile.getName());
                if (checkExtension(excelFile)) {
                    count++;
                    files.add(excelFile);
                    System.out.println("FilePath = " + excelFile.getPath() + "Count = " + count);
                    filenames.add(excelFile.getPath());
                }
            }

        }
        return files;
    }

    private boolean checkExtension(File file) {
        String extension = "";

        int i = file.getName().lastIndexOf('.');
        if (i > 0) {
            extension = file.getName().substring(i + 1);
        }
        return extension.toLowerCase(Locale.ROOT).contains("xls")
                && !file.getName().toLowerCase(Locale.ROOT).contains("~");
    }

    public List<String> getFilenames() {
        return filenames;
    }

    public List<XSSFWorkbook> getWorkbooks() {
        return workbooks;
    }
}

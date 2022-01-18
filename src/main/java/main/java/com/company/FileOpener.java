package main.java.com.company;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

public class FileOpener {
    private final String path;
    private List<File> files;
    private List<XSSFWorkbook> workbooks;
    private List<String> filenames;

    public FileOpener(String path) {
        this.path = path;
        init();
        try {
            openBooks(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void init() {
        files = new ArrayList<>();
        workbooks = new ArrayList<>();
        filenames = new ArrayList<>();
    }

    public XSSFWorkbook openBook(String path, String filename) throws Exception {
        return new XSSFWorkbook(new File(path + filename));
    }

    public void openBooks(String path) throws Exception {
        final File folder = new File(path);
        List<File> files = parseFile(folder);
        workbooks.clear();
        for (File file : files) {
            workbooks.add(new XSSFWorkbook(file));
        }
    }

    private List<File> parseFile(final File folder) {
        for (final File excelFile : Objects.requireNonNull(folder.listFiles())) {
            if (excelFile.isDirectory()) {
                parseFile(excelFile);
            } else {
                System.out.println(excelFile.getName());
                if (checkExtension(excelFile)) {
                    files.add(excelFile);

/*                    String parent3 = excelFile.getParentFile().getName() + "\\";
                    String parent2 = excelFile.getParentFile().getParentFile().getName() + "\\";
                    String parent1 = excelFile.getParentFile().getParentFile().getParentFile().getName() + "\\";
                    filenames.add(parent1 + parent2 + parent3 + excelFile.getName());*/
                    System.out.println("FilePath = " + excelFile.getPath());
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
        return extension.toLowerCase(Locale.ROOT).contains("xls");
    }

    public List<String> getFilenames() {
        return filenames;
    }

    public List<XSSFWorkbook> getWorkbooks(){
        return workbooks;
    }
}

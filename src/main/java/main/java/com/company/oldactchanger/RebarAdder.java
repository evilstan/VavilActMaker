package main.java.com.company.oldactchanger;

import main.java.com.company.helpers.FileChanger;
import main.java.com.company.helpers.FileOpener;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class RebarAdder {


    private List<String> filenames;
    private List<XSSFWorkbook> workbooks;
    private FileChanger fileChanger;
    private RebarParser rebarParser;

    public RebarAdder() {
        init();
    }

    private void init() {
      //  toPdf();
        fileChanger = new FileChanger();
        rebarParser = new RebarParser();
        openFiles();
        changeFiles();
    }

    private void openFiles() {
        try {
            FileOpener fileOpener = new FileOpener("C:\\vavil\\new");

            //fileOpener.openBooks("C:\\vavil\\new");
            workbooks = fileOpener.getWorkbooks();
            System.out.println("workbooks size " + workbooks.size());

            filenames = fileOpener.getFilenames();
            System.out.println("filenames size " + filenames.size());

        } catch (Exception e) {
            System.out.println("Exception while opening " + e.getMessage());
        }
    }

    private void changeFiles() {
        for (XSSFWorkbook workbook : workbooks) {
            addRebar(workbook);
            String filename = filenames.get(workbooks.indexOf(workbook)).replace("new", "rebar");
            saveFile(workbook, filename);
        }
    }

    private void saveFile(XSSFWorkbook workbook, String filename) {
        FileOutputStream fileOutputStream;
        XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
        try {
            fileOutputStream = new FileOutputStream(filename);
            workbook.write(fileOutputStream);
            fileOutputStream.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void addRebar(XSSFWorkbook workbook) {
        XSSFSheet sheet = workbook.getSheetAt(0);
        String oldRebar = fileChanger.getCellValue(sheet, "E19");
        List<String> newRebar = rebarParser.parse(oldRebar);
        String totalRebar = "Арматура класу А500С, А240С: " + String.join(", ", newRebar);
        String[] balancedRebar = balanceRebar(totalRebar);
        fileChanger.setCellValue(sheet, "E19", balancedRebar[0]);
        fileChanger.setCellValue(sheet, "E20", balancedRebar[1]);
    }

    private String[] balanceRebar(String str) {
        if (str.length() <= 85) return new String[]{"", str};
        for (int i = 60; i > 0; i--) {
            if (Character.toString(str.charAt(i)).equals(","))
                return new String[]{str.substring(0, i), str.substring(i + 1)};
        }
        return new String[2];
    }

    private void toPdf(){
        try {
            //create a temporary file and grab the path for it
            Path tempScript = Files.createTempFile("script", ".vbs");

            //read all the lines of the .vbs script into memory as a list
            //here we pull from the resources of a Gradle build, where the vbs script is stored
           // System.out.println("Path for vbs script is: '" +  main.java.com.company.Main.class.getResource("xl2pdf.vbs").toString().substring(6) + "'");
            List<String> script = Files.readAllLines(Paths.get("C:\\vavil\\xl2pdf.vbs"));

            // append test.xlsm for file name. savePath was passed to this function
            String templateFile = "C:\\1.xls";
            templateFile = templateFile.replace("\\", "\\\\");
            String pdfFile = "C:\\vavil\\new" + "\\test.pdf";
            pdfFile = pdfFile.replace("\\", "\\\\");
            System.out.println("templateFile is: " + templateFile);
            System.out.println("pdfFile is: " + pdfFile);

            //replace the placeholders in the vbs script with the chosen file paths
            for (int i = 0; i < script.size(); i++) {
                script.set(i, script.get(i).replaceAll("XL_FILE", templateFile));
                script.set(i, script.get(i).replaceAll("PDF_FILE", pdfFile));
                System.out.println("Line " + i + " is: " + script.get(i));
            }

            //write the modified code to the temporary script
            Files.write(tempScript, script);

            //create a processBuilder for starting an operating system process
            ProcessBuilder pb = new ProcessBuilder("wscript", tempScript.toString());

            //start the process on the operating system
            Process process = pb.start();

            //tell the process how long to wait for timeout
            Boolean success = process.waitFor(5, TimeUnit.SECONDS);
            if(!success) {
                System.out.println("Error: Could not print PDF within " + 5 + TimeUnit.SECONDS);
            } else {
                System.out.println("Process to run visual basic script for pdf conversion succeeded.");
            }

        } catch (Exception e) {
            e.printStackTrace();
/*            Alert saveAsPdfAlert = new Alert(AlertType.ERROR);
            saveAsPdfAlert.setTitle("ERROR: Error converting to pdf.");
            saveAsPdfAlert.setHeaderText("Exception message is:");
            saveAsPdfAlert.setContentText(e.getMessage());
            saveAsPdfAlert.showAndWait();*/
        }
    }

}

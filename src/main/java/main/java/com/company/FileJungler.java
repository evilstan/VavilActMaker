package main.java.com.company;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class FileJungler {

    private final String REBAR_CLASS = "Арматура класу ";
    private final String path;
    private List<XSSFWorkbook> workbooks;
    private List<String> filenames;
    private FileOpener fileOpener;
    private FileChanger fileChanger;

    private List<Integer> passport8;
    private List<Integer> passport10;
    private List<Integer> passport12;
    private List<Integer> passport16;
    private List<Integer> passport18;
    private List<Integer> passport20;
    private List<Integer> passport22;
    private List<Integer> passport25;
    private List<Integer> passport28;
    private List<Integer> passport32;


    public FileJungler(String path) {
        this.path = path;
        init();
    }

    private void init() {
        fileOpener = new FileOpener(path);
        fileChanger = new FileChanger();
        putRandom();
        try {
            workbooks = fileOpener.workbooks(path);
            filenames = fileOpener.getFilenames();
        } catch (Exception e) {
            System.out.println("Exception while opening " + e.getMessage());
        }
        XSSFSheet sheet = workbooks.get(0).getSheetAt(1);
        setNewRebarValues(sheet);
        doBoring(sheet);

        try {
            saveFile(workbooks.get(0), filenames.get(0));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void putRandom() {
        passport8 = new ArrayList<>();
        passport10 = new ArrayList<>();
        passport12 = new ArrayList<>();
        passport16 = new ArrayList<>();
        passport18 = new ArrayList<>();
        passport20 = new ArrayList<>();
        passport22 = new ArrayList<>();
        passport25 = new ArrayList<>();
        passport28 = new ArrayList<>();
        passport32 = new ArrayList<>();
        List<List<Integer>> list = Arrays.asList(passport8, passport10, passport12, passport16, passport18, passport20, passport22, passport25, passport28, passport32);

        for (List<Integer> integers : list) {
            for (int i = 0; i < 5; i++) {
                integers.add((int) (Math.random() * 1000000));
            }
        }
    }


    private void setNewRebarValues(XSSFSheet sheet) {
        List<String> firstLine = splitRebar(fileChanger.getCellValue(sheet, "D30"));
        List<String> secondLine = splitRebar(fileChanger.getCellValue(sheet, "A32"));

        String line1 = createStringRebar(firstLine);
        String line2 = createStringRebar(secondLine);

        fileChanger.setCellValue(sheet, "D30", line2); //swap!
        fileChanger.setCellValue(sheet, "A32", line1);
    }

    private List<String> splitRebar(String value) {
        int posDiameterStart = value.indexOf("Ø");
        int posDiameterEnd = value.lastIndexOf("Ø") + 3;
        int posClass = value.indexOf("класу ");

        String diameters = value.substring(posDiameterStart, posDiameterEnd);
        if (!Character.isDigit(diameters.charAt(diameters.length() - 1))) {
            diameters = diameters.substring(0, diameters.length() - 1);
        }

        String clazz = value.substring(posClass + 6, posClass + 11);
        List<String> rebars = new ArrayList<>(Arrays.asList(diameters.split(", ")));
        rebars.add(0, clazz);
        System.out.println(rebars);
        System.out.println(filenames.get(0));
        return rebars;
    }

    private String createStringRebar(List<String> line) {
        StringBuilder result = new StringBuilder(REBAR_CLASS + line.get(0));
        for (int i = 1; i < line.size(); i++) {
            String diam = line.get(i);
            result.append(", ").append(diam).append(" (").append(getPassport(diam).toString()).append(")");
        }
        return result.toString();
    }

    private Integer getPassport(String diameter) {
        switch (diameter) {
            case "Ø8":
                return passport8.get((int) (Math.random() * passport8.size() - 1));
            case "Ø10":
                return passport10.get((int) (Math.random() * passport10.size() - 1));
            case "Ø12":
                return passport12.get((int) (Math.random() * passport12.size() - 1));
            case "Ø16":
                return passport16.get((int) (Math.random() * passport16.size() - 1));
            case "Ø18":
                return passport18.get((int) (Math.random() * passport18.size() - 1));
            case "Ø20":
                return passport20.get((int) (Math.random() * passport20.size() - 1));
            case "Ø22":
                return passport22.get((int) (Math.random() * passport22.size() - 1));
            case "Ø25":
                return passport25.get((int) (Math.random() * passport25.size() - 1));
            case "Ø28":
                return passport28.get((int) (Math.random() * passport28.size() - 1));
            case "Ø32":
                return passport32.get((int) (Math.random() * passport32.size() - 1));
            default:
                System.out.println("Default");
        }
        return (int) (Math.random() * 1000000);
    }

    private void saveFile(XSSFWorkbook workbook, String filename) throws Exception {
        String savePath = path + "\\new\\" + filename;
        FileOutputStream saveStream = new FileOutputStream(savePath);
        workbook.write(saveStream);
    }

    private void doBoring(XSSFSheet sheet) {
        fileChanger.setCellValue(sheet, "B4", "\"Будівництво житлово-офісного комплексу з об'єктами торгово-розважальної, ринкової, ");
        fileChanger.setCellValue(sheet, "A5", "соціальної інфраструктури та паркінгами на вул. Сім'ї Хохлових, 8 у Шевченківському районі м. Києва\"");


        fileChanger.setCellValue(sheet, "h45", "Лук'янчук Олег Миколайович");
        fileChanger.setCellValue(sheet, "E9", "Лук'янчук Олег Миколайович");

        fileChanger.setCellValue(sheet, "h48", "Перестюк Вадим Володимирович");
        fileChanger.setCellValue(sheet, "F11", "Перестюк Вадим Володимирович");

        fileChanger.setCellValue(sheet, "A48", "Представник генпідрядної будівельної");
        fileChanger.setCellValue(sheet, "A11", "Представник генпідрядної будівельної організації");

        fileChanger.setCellValue(sheet, "A57", "Примітка. Керівник генпідрядної організації не пізніше, як за 5 днів інформує учасників про дату і");
        fileChanger.setCellValue(sheet, "A58", "місце проведення робіт");

        fileChanger.addRow(sheet, 55);
        fileChanger.copyRows(sheet, 1);
        fileChanger.setCellValue(sheet, "h2", "Додаток В");

    }

}

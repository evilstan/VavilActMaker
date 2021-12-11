

package main.java.com.company;

import org.apache.log4j.BasicConfigurator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Objects;

public class Main {

    public static void main(String[] args) throws IOException {
        BasicConfigurator.configure();
        System.setProperty("log4j.configurationFile", "./path_to_the_log4j2_config_file/log4j2.xml");
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        String path = "";
        System.out.println("Type path with files, Enter to default (C:\\2)");

        try {
            path = reader.readLine();
            if (path.equals("")) path = "C:\\2";
        } catch (Exception e) {
            System.out.println("Wrong path entered");
        }

        System.out.println(path);
        FileOpener fileOpener = new FileOpener(path);
    }

}

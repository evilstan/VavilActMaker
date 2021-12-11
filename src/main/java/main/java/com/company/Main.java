package main.java.com.company;

import org.apache.log4j.BasicConfigurator;

import java.io.BufferedReader;
import java.io.InputStreamReader;

public class Main {

    public static void main(String[] args) {
        BasicConfigurator.configure();
        System.setProperty("log4j.configurationFile", "./path_to_the_log4j2_config_file/log4j2.xml");
        String path;

/*        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        System.out.println("Type path with files, Enter to default (C:\\2)");

        try {
            path = reader.readLine();
            if (path.equals("")) path = "C:\\2";
        } catch (Exception e) {
            System.out.println("Wrong path entered");
        }*/
        path = "C:\\2";
        System.out.println(path);
        FileJungler fileChanger = new FileJungler(path);
    }

}

package main.java.com.company;

import main.java.com.company.newactcreator.ActCreator;
import main.java.com.company.oldactchanger.ActMaker;
import org.apache.log4j.BasicConfigurator;

public class Main {
    final static String PATH = "C:\\vavil\\old";

    public static void main(String[] args) {
        BasicConfigurator.configure();
        System.setProperty("log4j.configurationFile", "./path_to_the_log4j2_config_file/log4j2.xml");

/*        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        System.out.println("Type path with files, Enter to default (C:\\2)");

        try {
            path = reader.readLine();
            if (path.equals("")) path = "C:\\2";
        } catch (Exception e) {
            System.out.println("Wrong path entered");
        }*/

       // FileJuggler fileChanger = new FileJuggler("c:\\2\\");
        //ActMaker actMaker = new ActMaker(PATH);

        try {
            ActCreator actCreator = new ActCreator();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}

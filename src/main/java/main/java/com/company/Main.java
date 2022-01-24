package main.java.com.company;

import main.java.com.company.oldactchanger.RebarAdder;
import org.apache.log4j.BasicConfigurator;

public class Main {
    final static String PATH = "C:\\vavil\\old";

    public static void main(String[] args) {
        BasicConfigurator.configure();
        System.setProperty("log4j.configurationFile", "./path_to_the_log4j2_config_file/log4j2.xml");

       // FileJuggler fileChanger = new FileJuggler("c:\\2\\");
        //ActMaker actMaker = new ActMaker(PATH);
/*        RebarParser rebarParser = new RebarParser();
        List<String> list = rebarParser.parse("Арматура класу А500С Ø10,Ø12,Ø16,Ø20 А240С Ø6,Ø8");
        System.out.println("List size = " + list.size());

        for (String s : list) {
            System.out.println(s);
        }*/
        RebarAdder rebarAdder = new RebarAdder();

/*        try {
            ActCreator actCreator = new ActCreator();
        } catch (Exception e) {
            e.printStackTrace();
        }*/
    }

}

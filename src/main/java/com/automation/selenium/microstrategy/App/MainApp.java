package com.automation.selenium.microstrategy.App;

import com.automation.selenium.microstrategy.io.XlsXlsxConverter3;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;


public class MainApp {


    public static void main(String[] args) {
        ZipSecureFile.setMinInflateRatio(0);


        //Params verify
        if (args.length < 4) {
            System.out.println("\nIngrese el comando sugerido: App.jar -Type -x1 -x2 -y1, donde \n" +
                    "\nType = \"SR\" or \"IN\" " +
                    "\nx1 = path \"last_report.xsls\" " +
                    "\nx2 = path \"older_report.xsls\" " +
                    "\ny1 = output path \"output_name.xsls\" " );
            System.exit(0);
        }



        //Catching params
        final String report_type = (args[0].charAt(0) == '-') ? args[0].substring(1) : args[0];

        final String destination_dir =  (args[1].charAt(0) == '-') ? args[1].substring(1) : args[1];

        final String source_dir = (args[2].charAt(0) == '-') ? args[2].substring(1) : args[2];

        final String output_dir = (args[3].charAt(0) == '-') ? args[3].substring(1) : args[3];

        //Get files
        try {
            OPCPackage opc1 = OPCPackage.open(destination_dir);
            OPCPackage opc2 = OPCPackage.open(source_dir);
            Workbook wb_1 = new XSSFWorkbook(opc1);
            Workbook wb_2 = new XSSFWorkbook(opc2);

            Workbook wb_final = new XSSFWorkbook();

            wb_final.createSheet("Actual Week");
            wb_final.createSheet("Last Week");
            wb_final.createSheet("Final Report");

            XlsXlsxConverter3.run(report_type, wb_1, wb_2, wb_final);


            OutputStream os = new FileOutputStream(output_dir);

            //Parsear nombre con fecha
            System.out.print ("If you arrived here, it means you're good boy");
            wb_final.write(os);
            os.flush();
            os.close();
            wb_final.close();

        } catch (Exception e) {
            System.out.println("No se puede encontrar los archivos especificados. Asegurese de escribir la ruta");
            e.printStackTrace();
        }

    }
}

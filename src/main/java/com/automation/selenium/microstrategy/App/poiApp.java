package com.automation.selenium.microstrategy.App;


import com.automation.selenium.microstrategy.io.XlsXlsxConverter3;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/***
 * This app will storage all the data in memory... Please do not fuck it up with the amount of data you have.
 */

public class poiApp {

    public static final String dir_1 =
            "C:" + File.separator
            + "Users" + File.separator
            + "ivan.monzon" + File.separator
            + "Documents" + File.separator
            + "Directv_Auto" + File.separator
            + "SR_Phermosi_01jun2018_0701AM.xlsx";

    public static final String dir_2 =
            "C:" + File.separator
                    + "Users" + File.separator
                    + "ivan.monzon" + File.separator
                    + "Documents" + File.separator
                    + "Directv_Auto" + File.separator
                    + "SR_Phermosi_11may2018_0700AM.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        File excel_1 = new File(dir_1);
        File excel_2 = new File(dir_2);
        Workbook wb_1 = new XSSFWorkbook(OPCPackage.open(dir_1));
        Workbook wb_2 = new XSSFWorkbook(OPCPackage.open(dir_2));

        wb_1.setSheetName(0, "Actual");

    /*Get sheets from the temp file*/
        Sheet sheet_1 = wb_1.getSheetAt(0);
        Sheet sheet_2 = wb_1.createSheet("Later");




        OutputStream os = new FileOutputStream("hello.xlsx");

        wb_1.write(os);
        os.flush();
        os.close();
        wb_1.close();
    }


}

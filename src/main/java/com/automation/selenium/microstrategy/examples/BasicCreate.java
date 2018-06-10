package com.automation.selenium.microstrategy.examples;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class BasicCreate {

    public static final String dir = "C:\\Users\\ivan.monzon\\Documents\\Directv_Auto\\SR_Phermosi_01jun2018_0701AM.xls";

    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();




        Sheet sheet = wb.createSheet("Sample");

        System.out.print(dir);

        //the following three statements are required only for HSSF
        sheet.setAutobreaks(true);

        Row header = sheet.createRow(1);
        Cell cell = header.createCell(1);
        cell.setCellValue("Hello world bitches");



        OutputStream os = new FileOutputStream("hello.xls");

        wb.write(os);
        os.close();
        wb.close();

    }

    private Workbook getLastWorkbook() throws Exception{
        return new HSSFWorkbook();
    }


}

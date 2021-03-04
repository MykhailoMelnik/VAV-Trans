package com.company.apache;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Map;

public class ExelFiles {

    String SHEET = "01.02.2021";
    String INVOICE = "/Users/mihajlomelnik/Documents/VAV TRANS/Інвойси 2021 LKW !.xlsx";
    String data;


    File miFile = new File(INVOICE);
    FileInputStream fileInputStream = new FileInputStream(miFile);
    Workbook workbook = new XSSFWorkbook(fileInputStream);
    XSSFSheet sheet = (XSSFSheet) workbook.getSheet(SHEET);

    public ExelFiles() throws IOException {

    }

    public void foundEmptyRow() throws IOException {
        data = String.valueOf(sheet.getRow(4).getCell(13));
        for (int i = 4; data.length() != 0; i++) {
            data = String.valueOf(sheet.getRow(i).getCell(13));
            if ((String.valueOf(sheet.getRow(i).getCell(10))).equals("")) {
                System.out.println(i);
//                TollCollect tollCollect = new TollCollect();
            }

        }
        fileInputStream.close();
        System.out.println("clouse");



        }
    }


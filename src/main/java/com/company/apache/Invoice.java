package com.company.apache;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;

public class Invoice {

    String SHEET = "01.02.2021";
    String INVOICE = "C:\\Users\\Professional\\Desktop\\VAV TRANS\\Інвойси 2021 LKW test.xlsx";
    //  String INVOICE = "/Users/mihajlomelnik/Documents/VAV TRANS/Інвойси 2021 LKW !.xlsx";

    File miFile = new File(INVOICE);
    FileInputStream fileInputStream = new FileInputStream(miFile);
    Workbook workbook = new XSSFWorkbook(fileInputStream);
    XSSFSheet sheet = (XSSFSheet) workbook.getSheet(SHEET);


    public Invoice() throws IOException {
        foundEmptyRow();
    }

    public void foundEmptyRow() throws IOException {
        TollCollect tollCollect = new TollCollect();
        for (int i = 4; String.valueOf(sheet.getRow(i).getCell(13)).length() != 0; i++) {
            if ((String.valueOf(sheet.getRow(i).getCell(10))).equals("")) {
                double euroInInvoiceWithEmptyCell = Math.abs(
                        Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(15))));
                Date dataOfCases = sheet.getRow(i).getCell(13).getDateCellValue();
                tollCollect.searchInTollCollect(euroInInvoiceWithEmptyCell, dataOfCases);
            }
        }
        tollCollect.closeInputStreamTollCollect();
        closeInputStreamInvoice();
    }

    public void closeInputStreamInvoice() throws IOException {
        fileInputStream.close();
        System.out.println("close stream Invoice");
    }
}


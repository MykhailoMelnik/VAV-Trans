package com.company.apache;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Date;
import java.util.Iterator;


public class TollCollect {
    String TOLL_COLLECT = "/Users/mihajlomelnik/Documents/VAV TRANS/Z_EFN_1256149_6023019663.xlsx";

    File miFile = new File(TOLL_COLLECT);
    FileInputStream fileInputStream = new FileInputStream(miFile);
    Workbook workbook = new XSSFWorkbook(fileInputStream);
    XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);

    public TollCollect() throws IOException {

    }

    public void searchInTollCollect(double euroInInvoice, Date dateInInvoice) {
        System.out.println(euroInInvoice);
        for (int i = 1; i < sheet.getLastRowNum() - 3; i++) {
            Date dateTollCollect = workbook.getSheetAt(0).getRow(i).getCell(2).getDateCellValue();
            double euroInTollCollect = Double.parseDouble(String.valueOf(workbook.getSheetAt(0).getRow(1).getCell(17)));

//            if (dateInInvoice.getTime() - dateTollCollect.getTime() / 86400000 < 6 &&
//                    euroInTollCollect == euroInInvoice) {

//                System.out.println(dateInInvoice.getTime() - dateTollCollect.getTime() / 86400000);
//                System.out.println(euroInTollCollect + " " + euroInInvoice);
//            }
        }
    }

    public void closeInputStreamTollCollect() throws IOException {
        fileInputStream.close();
        System.out.println("close stream TollCollect");
    }
}

package com.company.apache;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Date;
import java.util.Iterator;


public class TollCollect {
    String TOLL_COLLECT = "C:\\Users\\Professional\\Desktop\\VAV TRANS\\6023019663.xlsx";
    //String TOLL_COLLECT = "/Users/mihajlomelnik/Documents/VAV TRANS/Z_EFN_1256149_6023019663.xlsx";

    File miFile = new File(TOLL_COLLECT);
    FileInputStream fileInputStream = new FileInputStream(miFile);
    Workbook workbook = new XSSFWorkbook(fileInputStream);
    XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);

    public TollCollect() throws IOException {

    }

    public boolean searchInTollCollect(double euroInInvoice, Date dateInInvoice) throws IOException {
        boolean euroInTollCol = false;
        for (int i = 1; i < sheet.getLastRowNum() - 3; i++) {
            Date dateTollCollect = sheet.getRow(i).getCell(2).getDateCellValue();

            double euroInTollCollect = Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17)));
            if ((dateInInvoice.getTime() - dateTollCollect.getTime()) / 86400000 < 6 &&
                    (dateInInvoice.getTime() - dateTollCollect.getTime()) / 86400000 > 0 &&
                    euroInTollCollect == euroInInvoice) {
                System.out.println("сумма найдена " + euroInInvoice);
                // delete row in toll collect if the number was found
                euroInTollCol = true;
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
                FileOutputStream outputStream = new FileOutputStream(TOLL_COLLECT);
                workbook.write(outputStream);
                outputStream.close();
                break;
            }
        }
        return euroInTollCol;
    }


    public void closeInputStreamTollCollect() throws IOException {
        fileInputStream.close();
        System.out.println("close stream TollCollect");
    }
}

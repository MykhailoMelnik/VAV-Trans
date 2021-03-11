package com.company.apache;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;


public class TollCollect {
    String TOLL_COLLECT = "C:\\Users\\Professional\\Desktop\\Git\\6023019663.xlsx";
    //String TOLL_COLLECT = "/Users/mihajlomelnik/Documents/VAV TRANS/Z_EFN_1256149_6023019663.xlsx";
    public String DATA_TOLL_COLLECT;
    public Double numberTollCollect;

    File miFile = new File(TOLL_COLLECT);
    FileInputStream fileInputStream = new FileInputStream(miFile);
    Workbook workbook = new XSSFWorkbook(fileInputStream);
    XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);

    public TollCollect(String DATA_TOLL_COLLECT) throws IOException {
        this.DATA_TOLL_COLLECT = DATA_TOLL_COLLECT;
        numberTollCollect = sheet.getRow(0).getCell(0).getNumericCellValue();
    }

    public boolean searchInTollCollect(double euroInInvoice, Date dateInInvoice) throws IOException {

        for (int i = 1; i < sheet.getLastRowNum() - 3; i++) {
            Date dateTollCollect = sheet.getRow(i).getCell(2).getDateCellValue();

            double euroInTollCollect = Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17)));
            if ((dateInInvoice.getTime() - dateTollCollect.getTime()) / 86400000 < 6 &&
                    (dateInInvoice.getTime() - dateTollCollect.getTime()) / 86400000 > 0 &&
                    euroInTollCollect == euroInInvoice) {
                System.out.println("сумма найдена " + euroInInvoice);
                // delete row in toll collect if the number was found
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
                FileOutputStream outputStream = new FileOutputStream(TOLL_COLLECT);
                workbook.write(outputStream);
                outputStream.close();
                return true;
            }
        }
        return false;
    }



    public void closeInputStreamTollCollect() throws IOException {
        fileInputStream.close();
    }
}

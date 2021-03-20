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
    //    String TOLL_COLLECT = "C:\\Users\\Professional\\Desktop\\Git\\6023170227.xlsx";
    String TOLL_COLLECT = "/Users/mihajlomelnik/Documents/VAV TRANS/1.xlsx";
    public String DATA_TOLL_COLLECT = "16.03.2021";
    public Double numberTollCollect;
    private int calculationAmounts;

    File miFile = new File(TOLL_COLLECT);
    FileInputStream fileInputStream = new FileInputStream(miFile);
    Workbook workbook = new XSSFWorkbook(fileInputStream);
    XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);

    public TollCollect() throws IOException {
        numberTollCollect = sheet.getRow(0).getCell(0).getNumericCellValue();
    }


    public boolean searchInTollCollect(double euroInInvoice, Date dateInInvoice) throws IOException {

        for (int i = 1; i < sheet.getLastRowNum() - 3; i++) {
            Date dateTollCollect = sheet.getRow(i).getCell(2).getDateCellValue();

            double euroInTollCollect = Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17)));
            if ((dateInInvoice.getTime() - dateTollCollect.getTime()) / 86400000 < 10 &&
                    (dateInInvoice.getTime() - dateTollCollect.getTime()) / 86400000 > 0 &&
                    euroInTollCollect == euroInInvoice) {
                calculationAmounts += 1;
                System.out.println(calculationAmounts + " сумма найдена " + euroInInvoice);
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

    public void summationTollCollectByDates() throws IOException {

        for (int i = 1; i < sheet.getLastRowNum() - 3; i++) {
            if (sheet.getRow(i).getCell(2).getDateCellValue()
                    .equals(sheet.getRow(i + 1).getCell(2).getDateCellValue())) {
                double calculation = Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17)))
                        + Double.parseDouble(String.valueOf(sheet.getRow(i + 1).getCell(17)));

                sheet.getRow(i + 1).getCell(17).setCellValue(calculation);
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);

                FileOutputStream outputStream = new FileOutputStream(TOLL_COLLECT);
                workbook.write(outputStream);
                outputStream.close();
                summationTollCollectByDates();
            }
        }
    }

    public void closeInputStreamTollCollect() throws IOException {
        fileInputStream.close();
    }
}

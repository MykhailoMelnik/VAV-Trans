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
    String TOLL_COLLECT = "C:\\Users\\Professional\\Desktop\\Git\\6023170227.xlsx";
    //String TOLL_COLLECT = "/Users/mihajlomelnik/Documents/VAV TRANS/Z_EFN_1256149_6023019663.xlsx";
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
        Date dateTollCollect;
        double calculation;
        for (int i = 1; i < sheet.getLastRowNum() - 4; i++) {
            calculation = Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17)));
            dateTollCollect = sheet.getRow(i).getCell(2).getDateCellValue();

            for (int j = i + 1; j < sheet.getLastRowNum()-4; j++) {
                if (dateTollCollect == sheet.getRow(j).getCell(2).getDateCellValue()) {
                    calculation += Double.parseDouble(String.valueOf(sheet.getRow(j).getCell(17)));
                    sheet.shiftRows(j, sheet.getLastRowNum(), -1);
                    System.out.println("ok");
                } else {
                    System.out.println("nex data");
                    sheet.getRow(i).getCell(17).setCellValue(calculation);
                    calculation = 0;
                    FileOutputStream outputStream = new FileOutputStream(TOLL_COLLECT);
                    workbook.write(outputStream);
                    outputStream.close();
                }
            }
        }
    }

        public void closeInputStreamTollCollect () throws IOException {
            fileInputStream.close();
        }
    }

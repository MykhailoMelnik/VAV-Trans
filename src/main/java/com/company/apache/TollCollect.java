package com.company.apache;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;


public class TollCollect {
//    static String TOLL_COLLECT = "C:\\Users\\Professional\\Desktop\\Git\\6023306181.xlsx";
      String TOLL_COLLECT = "/Users/mihajlomelnik/IdeaProjects/VAV-Trans/6023306181.xlsx";
    public String DATA_TOLL_COLLECT = "16.04.2021";
    public Double numberTollCollect;
    private int calculationAmounts;

    File miFile = new File(TOLL_COLLECT);
    FileInputStream fileInputStream = new FileInputStream(miFile);
    Workbook workbook = new XSSFWorkbook(fileInputStream);
    XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);

    public TollCollect() throws IOException {
        numberTollCollect = sheet.getRow(0).getCell(0).getNumericCellValue();
    }

    // TODO: 25.03.2021 try do 
    public void transformTollCollect() throws IOException {
        FileOutputStream outputStream = new FileOutputStream(TOLL_COLLECT);

        for (int i = 1; i < sheet.getLastRowNum() - 3; i++) {
            for (int j = 3; j < 17; j++) {
                sheet.getRow(i).getCell(j).setCellValue("");
            }
            sheet.getRow(i).getCell(3)
                    .setCellValue(String.valueOf(sheet.getRow(i).getCell(17)));
        }
        workbook.write(outputStream);
        outputStream.close();
    }


    public boolean searchInTollCollect(double euroInInvoice, Date dateInInvoice) throws IOException {

        for (int i = 1; i < sheet.getLastRowNum() - 3; i++) {
            Date dateTollCollect = sheet.getRow(i).getCell(2).getDateCellValue();

            double euroInTollCollect = Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17)));
            if ((dateInInvoice.getTime() - dateTollCollect.getTime()) / 86400000 < 7 &&
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

    // TODO: 25.03.2021 fix 
    public void summationTollCollectByDates(Walter walter) throws IOException {
        double calculation = 0;
        for (int i = 1; i < sheet.getLastRowNum() - 3; i++) {

            int rowStart = 0;
            if (sheet.getRow(i).getCell(2).getDateCellValue()
                    .equals(sheet.getRow(i + 1).getCell(2).getDateCellValue())) {
                rowStart -= 1;
                calculation += Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17)));
            } else {
                calculation += Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17)));
                Date dataDayInToll = sheet.getRow(i).getCell(2).getDateCellValue();
                rowStart -= 1;
                if (walter.foundNumberInWalterAfterSummation(calculation, dataDayInToll)) {
                    System.out.println("нашло суму вернуло тру " + calculation);

                    sheet.shiftRows(i+1, sheet.getLastRowNum(), rowStart-2);
                    FileOutputStream outputStream = new FileOutputStream(TOLL_COLLECT);
                    workbook.write(outputStream);
                    outputStream.close();

                }
                calculation = 0;
            }




            /*double calculation = 0;
            int rowStart;
            if (sheet.getRow(i).getCell(2).getDateCellValue()
                    .equals(sheet.getRow(i + 1).getCell(2).getDateCellValue())) {
                rowStart = i;
                calculation = Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17)))
                        + Double.parseDouble(String.valueOf(sheet.getRow(i + 1).getCell(17)));

            } else if (walter.foundNumberInWalterAfterSummation(calculation)) {
                sheet.getRow(i + 1).getCell(17).setCellValue(calculation);
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
            }
            FileOutputStream outputStream = new FileOutputStream(TOLL_COLLECT);
            workbook.write(outputStream);
            outputStream.close();
            summationTollCollectByDates();*/
        }
    }

    public void closeInputStreamTollCollect() throws IOException {
        fileInputStream.close();
    }
}

package com.company.apache;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.Date;

public class Walter {

    String SHEET;
    String INVOICE = "C:\\Users\\Professional\\Desktop\\Git\\Інвойси 2021 LKW.xlsx";
//    String INVOICE = "/Users/mihajlomelnik/Documents/VAV TRANS/Інвойси 2021 LKW !.xlsx";

    File miFile = new File(INVOICE);
    FileInputStream fileInputStream = new FileInputStream(miFile);
    Workbook workbook = new XSSFWorkbook(fileInputStream);
    XSSFSheet sheet;
    TollCollect tollCollect = new TollCollect();
    double euroInInvoiceWithEmptyCell;

    public Walter() throws IOException {
    }

    public void start(String sheetStart) throws IOException {
        SHEET = sheetStart;
        //tollCollect.transformTollCollect();
        startFound();
        System.out.println("суммирование toll collect начало");
        //tollCollect.summationTollCollectByDates();
        System.out.println("суммирование toll collect закончено");
        SHEET = sheetStart;
        startFound();
        totalAmountTollCollectInWalter();

    }

    public void startFound() throws IOException {
        sheet = (XSSFSheet) workbook.getSheet(SHEET);
        System.out.println("Страница " + SHEET);
        for (int i = 4; String.valueOf(sheet.getRow(i).getCell(13)).length() != 0; i++) {
            if ((String.valueOf(sheet.getRow(i).getCell(10))).equals("")) {
                double euroInInvoiceWithEmptyCell = Math.abs(
                        Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(15))));
                Date dataOfCases = sheet.getRow(i).getCell(13).getDateCellValue();
                if (tollCollect.searchInTollCollect(euroInInvoiceWithEmptyCell, dataOfCases)) {
                    writeValueToInvoice(i);
                }
            }
        }
//        Recursion
        SHEET = workbook.getSheetName(workbook.getSheetIndex(sheet) + 1);
        if (sheet.iterator().hasNext() && SHEET.length() > 6) {
            startFound();
        }
        tollCollect.closeInputStreamTollCollect();
        closeInputStreamInvoice();
    }

    public void totalAmountTollCollectInWalter() throws IOException {

        sheet = (XSSFSheet) workbook.getSheet(SHEET);
        System.out.println("Страница " + SHEET);
        DecimalFormat decimalFormat = new DecimalFormat( "#.###" );
        String result = decimalFormat.format(euroInInvoiceWithEmptyCell);
        String resultMinusSumTollCollect = decimalFormat
                .format(euroInInvoiceWithEmptyCell - tollCollect.sheet.getRow(1).getCell(0).getNumericCellValue());
        System.out.println(result);
        System.out.println(resultMinusSumTollCollect);

        for (int i = 4; String.valueOf(sheet.getRow(i).getCell(13)).length() != 0; i++) {
            try {
                if (tollCollect.numberTollCollect == sheet.getRow(i).getCell(10).getNumericCellValue()) {
                    euroInInvoiceWithEmptyCell += Math.abs(Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(15))));
                }
            } catch (Exception e) {
            }
        }
//        Recursion
        SHEET = workbook.getSheetName(workbook.getSheetIndex(sheet) + 1);
        if (sheet.iterator().hasNext() && SHEET.length() > 6) {
            totalAmountTollCollectInWalter();
        }
        closeInputStreamInvoice();
    }


    public void writeValueToInvoice(int row) throws IOException {

        sheet.getRow(row).getCell(10).setCellValue(tollCollect.numberTollCollect);
        sheet.getRow(row).getCell(11).setCellValue(tollCollect.DATA_TOLL_COLLECT);
        sheet.getRow(row).getCell(12).setCellValue("Toll Collect");

        FileOutputStream out = new FileOutputStream(INVOICE);
        workbook.write(out);
        out.close();
    }

    public void closeInputStreamInvoice() throws IOException {
        fileInputStream.close();
    }
}


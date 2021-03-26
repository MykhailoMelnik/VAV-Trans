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
    String numberDocForCount;
    double sumDoc;
    int countAmountDoc;

    public Walter() throws IOException {
    }

    public void start(String sheetStart) throws IOException {
        SHEET = sheetStart;
        //tollCollect.transformTollCollect();
        // startFound();
        // System.out.println("суммирование toll collect начало");
        //tollCollect.summationTollCollectByDates();
        // System.out.println("суммирование toll collect закончено");
        // SHEET = sheetStart;
        // startFound();
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

    public void totalAmountNumberDocInWalter(String numberDoc) throws IOException {
        numberDocForCount = numberDoc;
        sheet = (XSSFSheet) workbook.getSheet(SHEET);
        for (int i = 4; String.valueOf(sheet.getRow(i).getCell(13)).length() != 0; i++) {
            try {
                if (Long.parseLong(numberDoc) == sheet.getRow(i).getCell(10).getNumericCellValue()) {
                    countAmountDoc += 1;;
                    sumDoc += Math.abs(Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(15))));
                }
            } catch (Exception e) {
                if (String.valueOf(sheet.getRow(i).getCell(10)).equals(numberDoc)) {
                    countAmountDoc += 1;
                    sumDoc += Math.abs(Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(15))));
                }
            }
        }
        SHEET = workbook.getSheetName(workbook.getSheetIndex(sheet) + 1);
        if (sheet.iterator().hasNext() && SHEET.length() > 6) {
            totalAmountNumberDocInWalter(numberDoc);
        }
        closeInputStreamInvoice();
    }

    public void printAfterTotalAmount() {

        DecimalFormat decimalFormat = new DecimalFormat("#.###");
        String result = decimalFormat.format(sumDoc);
        System.out.println("\n--------------------------------------");
        System.out.println("документ с номером \" " + numberDocForCount + "\"");
        System.out.println("--------------------------------------");
        System.out.println("рознесен на сумму \t" + result + " EUR");
        System.out.println("\t\t\t\t\t\t  " + countAmountDoc + " шт");
        System.out.println("--------------------------------------");
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


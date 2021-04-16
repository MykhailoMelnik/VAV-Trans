package com.company.apache;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.Date;
import java.util.Scanner;

public class Walter {

    String SHEET;
    String SHEETStart;
    static String INVOICE = "C:\\Users\\Professional\\Desktop\\Git\\Інвойси 2021 LKW.xlsx";
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
        SHEETStart = sheetStart;
        //SHEET = sheetStart;
        //startFound();
        SHEET = sheetStart;
        tollCollect.summationTollCollectByDates(this);
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
                        Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17))));
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
                    countAmountDoc += 1;

                    sumDoc += Math.abs(Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17))));
                }
            } catch (Exception e) {
                if (String.valueOf(sheet.getRow(i).getCell(10)).equals(numberDoc)) {
                    countAmountDoc += 1;
                    sumDoc += Math.abs(Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17))));
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

    public boolean foundNumberInWalterAfterSummation(double summaInTollAfterCalculateByDate, Date dateTollCollect) throws IOException {

        SHEET = SHEETStart;
        while (SHEET.length() > 6) {
            sheet = (XSSFSheet) workbook.getSheet(SHEET);
            for (int i = 4; String.valueOf(sheet.getRow(i).getCell(9)).length() != 0; i++) {
                if ((String.valueOf(sheet.getRow(i).getCell(10))).equals("")) {

                    double euroInInvoiceWithEmptyCell = Math.abs(
                            Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17))));
                    Date dateInInvoice = sheet.getRow(i).getCell(13).getDateCellValue();

                    if (euroInInvoiceWithEmptyCell == summaInTollAfterCalculateByDate
                            && (dateInInvoice.getTime() - dateTollCollect.getTime()) / 86400000 < 7 &&
                            (dateInInvoice.getTime() - dateTollCollect.getTime()) / 86400000 > 0) {

                        writeValueToInvoice(i);
                        closeInputStreamInvoice();
                        System.out.println(summaInTollAfterCalculateByDate + " найдено на странице " + SHEET);
                        return true;
                    }
                }
            }
            SHEET = workbook.getSheetName(workbook.getSheetIndex(sheet) + 1);
        }
        return false;
    }

    public void foundAndWritIDS() throws IOException {
        while (true) {
            SHEET = "29.03.2021";
            sheet = (XSSFSheet) workbook.getSheet(SHEET);

            System.out.println("введіть число");
            Scanner scanner = new Scanner(System.in);
            double sum = scanner.nextDouble();
            System.out.println("1-DKK, 2-GER, 3-LUX, 4-SPAIN, 5-FRAN");
            double scale = Math.pow(10, 2);
            switch (scanner.nextInt()) {
                case (1):
                    System.out.println(sum);
                    break;
                case (2):
                    sum += sum / 100 * 19;
                    sum = Math.ceil(sum * scale) / scale;
                    sum = new BigDecimal(sum).setScale(2, RoundingMode.HALF_UP).doubleValue();
                    System.out.println(sum);
                    break;
                case (3):
                    sum += sum / 100 * 17;
                    sum = Math.ceil(sum * scale) / scale;
                    sum = new BigDecimal(sum).setScale(2, RoundingMode.HALF_UP).doubleValue();
                    System.out.println(sum);
                    break;
                case (4):
                    sum += sum / 100 * 21;
                    sum = Math.ceil(sum * scale) / scale;
                    sum = new BigDecimal(sum).setScale(2, RoundingMode.HALF_UP).doubleValue();
                    System.out.println(sum);
                    break;
                case (5):
                    sum += sum / 100 * 20;
                    sum = Math.ceil(sum * scale) / scale;
                    sum = new BigDecimal(sum).setScale(2, RoundingMode.HALF_UP).doubleValue();
                    System.out.println(sum);
                    break;
            }

            while (sheet.iterator().hasNext() && SHEET.length() > 6) {

                for (int i = 4; String.valueOf(sheet.getRow(i).getCell(13)).length() != 0; i++) {

                    if ((String.valueOf(sheet.getRow(i).getCell(10))).equals("") && Math.abs(
                            Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17)))) == sum) {
                        System.err.println("Страница " + SHEET);
                        System.err.println("нашов " + Math.abs(
                                Double.parseDouble(String.valueOf(sheet.getRow(i).getCell(17)))));

                        sheet.getRow(i).getCell(10).setCellValue("DE00574498");
                        sheet.getRow(i).getCell(11).setCellValue("2021-03-31");
                        sheet.getRow(i).getCell(12).setCellValue("Q8");

                        FileOutputStream out = new FileOutputStream(INVOICE);
                        workbook.write(out);
                        out.close();
                        break;
                    }
                }

                SHEET = workbook.getSheetName(workbook.getSheetIndex(sheet) + 1);
                sheet = (XSSFSheet) workbook.getSheet(SHEET);
                closeInputStreamInvoice();
            }
        }
    }

    public void closeInputStreamInvoice() throws IOException {
        fileInputStream.close();
    }
}


package com.company.apache;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;


public class WriteTollCollect {

    public static String SHEET = "01.02.2021";
    public static String INVOICE = "/Users/mihajlomelnik/Documents/VAV TRANS/Інвойси 2021 LKW !.xlsx";
    public static String TOLL_COLLECT = "/Users/mihajlomelnik/Documents/VAV TRANS/6023019663.xlsx";
    public static String data;
    public static String dataInvoice;


    {
        try {
            File miFile = new File(INVOICE);
            FileInputStream fileInputStream = new FileInputStream(miFile);
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = (XSSFSheet) workbook.getSheet(SHEET);
            String result = "Start";

            for (int i = 4; result.length() != 0; i++) {
                dataInvoice = String.valueOf(sheet.getRow(i).getCell(13));
                result = String.valueOf(workbook.getSheet(SHEET).getRow(i).getCell(15));
//  search for an empty field
                if (String.valueOf(workbook.getSheet(SHEET).getRow(i).getCell(10)).equals("")) {

                    if (!foundTol(result).equals("")) {
                        System.out.println(result);
                        setNumberInInvoice(foundTol(result), i);
                    }
                }
            }
            fileInputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String foundTol(String invoiceNumber) throws IOException, ParseException {

        File miFile = new File(TOLL_COLLECT);
        FileInputStream fileInputStream = new FileInputStream(miFile);
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        String result = NumberToTextConverter.toText(
                workbook.getSheetAt(0).getRow(0).getCell(0).getNumericCellValue());
        int marcerRowToll = 0;
        String found = "start";
        for (int i = 1; !found.equals("stop"); i++) {

            data = String.valueOf(workbook.getSheetAt(0).getRow(i).getCell(1));
            SimpleDateFormat formatter2 = new SimpleDateFormat("dd-MMM-yyyy");
            Date date = formatter2.parse(data);
            Date dateInv = formatter2.parse(dataInvoice);
            int day = (int) ((date.getTime() - dateInv.getTime()) / 86400000);

            found = String.valueOf(
                    workbook.getSheetAt(0).getRow(i).getCell(2));

            if (invoiceNumber.length() > 0 && invoiceNumber.substring(
                    invoiceNumber.indexOf('-') + 1).equals(found) && day < 0 && day > -12) {
                result = NumberToTextConverter.toText(
                        workbook.getSheetAt(0).getRow(0).getCell(0).getNumericCellValue());
                workbook.getSheetAt(0).getRow(i).getCell(2).setCellValue(7777);
                marcerRowToll = i;
//  set "ok" in toll collect!

                CellStyle style = workbook.createCellStyle();
                style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
                style.setFillPattern(FillPatternType.BRICKS);

                workbook.getSheetAt(0).getRow(i).getCell(2).setCellStyle(style);

                break;
            } else {
                result = "";
            }
        }
        fileInputStream.close();
        if (!result.equals("")) {
            setInTollColMarker(marcerRowToll);
        }
        return result;
    }


    public static void setNumberInInvoice(String numberTollCollect, int i) throws IOException, ParseException {
        InputStream ExcelFileToRead = new FileInputStream(INVOICE);
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
        XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
        XSSFSheet sheet = wb.getSheet(SHEET);
//        Integer parse = parseInt(numberTollCollect);
//        CellStyle numberStyle = wb.createCellStyle();
//        numberStyle.getFontIndexAsInt();
        Long parse = Long.parseLong(numberTollCollect);

        sheet.getRow(i).getCell(10).setCellValue(parse);
        sheet.getRow(i).getCell(12).setCellValue("Toll Collect");
        sheet.getRow(i).getCell(11).setCellValue("16.02.2021");

        FileOutputStream out = new FileOutputStream(new File(INVOICE));
        wb.write(out);
        out.close();
    }

    public static void setInTollColMarker(int i) throws IOException {
        InputStream ExcelFileToRead = new FileInputStream(TOLL_COLLECT);
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
        XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
        XSSFSheet sheet = wb.getSheet("6023019663");

        CellStyle style = wb.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.YELLOW1.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);


        sheet.getRow(i).getCell(2).setCellStyle(style);

        FileOutputStream out = new FileOutputStream(new File(TOLL_COLLECT));
        wb.write(out);
        out.close();
    }
}

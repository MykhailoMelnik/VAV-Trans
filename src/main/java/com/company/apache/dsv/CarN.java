package com.company.apache.dsv;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class CarN extends Car {

    File miFileN = new File("C:\\Users\\Professional\\Desktop\\Git\\dsv\\N.xlsx");
    FileInputStream fileInputStream = new FileInputStream(miFileN);
    Workbook workbookN = new XSSFWorkbook(fileInputStream);
    XSSFSheet sheetN;
    String numberCar;


    File AllCarsFile = new File("C:\\Users\\Professional\\Desktop\\Git\\dsv\\3.2021.xlsx");
    FileInputStream allCarsFileInputStream = new FileInputStream(AllCarsFile);
    Workbook allCarsFileWorkbook = new XSSFWorkbook(allCarsFileInputStream);

    public CarN() throws IOException {
        copyAndSetCarN();
    }

    public void copyAndSetCarN() throws IOException {
        for (int i = 1; workbookN.getSheetName(workbookN.getSheetIndex(sheetN) + i).length() > 7; i++) {
            numberCar = workbookN.getSheetName(workbookN.getSheetIndex(sheetN) + i);
            System.out.println(numberCar);

            for (int j = 1; j < 43; j++) {

                int cell = (int) workbookN.getSheet(numberCar).getRow(j).getCell(0).getNumericCellValue();
                if (cell != 0) allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(0).setCellValue(
                        workbookN.getSheet(numberCar).getRow(j).getCell(0).getNumericCellValue());

                cell = (int) workbookN.getSheet(numberCar).getRow(j).getCell(1).getNumericCellValue();
                if (cell != 0) allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(1).setCellValue(
                        workbookN.getSheet(numberCar).getRow(j).getCell(1).getNumericCellValue());

                allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(2).setCellValue(
                        String.valueOf(workbookN.getSheet(numberCar).getRow(j).getCell(2)));

                allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(3).setCellValue(
                        String.valueOf(workbookN.getSheet(numberCar).getRow(j).getCell(3)));

                cell = (int) workbookN.getSheet(numberCar).getRow(j).getCell(4).getNumericCellValue();
                if (cell != 0) allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(4).setCellValue(
                        workbookN.getSheet(numberCar).getRow(j).getCell(4).getNumericCellValue());
                System.out.println(cell);

                cell = (int) workbookN.getSheet(numberCar).getRow(j).getCell(5).getNumericCellValue();
                if (cell != 0) allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(5).setCellValue(
                        workbookN.getSheet(numberCar).getRow(j).getCell(5).getNumericCellValue());

                allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(11).setCellValue(
                        String.valueOf(workbookN.getSheet(numberCar).getRow(j).getCell(12)));

                System.out.println("Срока " + j + " скопирована");
                FileOutputStream fileOutputStream = new FileOutputStream(AllCarsFile);
                allCarsFileWorkbook.write(fileOutputStream);
                fileOutputStream.close();
            }
        }
        fileInputStream.close();
    }
}

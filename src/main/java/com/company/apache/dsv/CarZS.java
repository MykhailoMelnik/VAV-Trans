package com.company.apache.dsv;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Map;

public class CarZS extends Car {

    File miFileZS = new File("C:\\Users\\Professional\\Desktop\\Git\\dsv\\ZS.xlsx");
    FileInputStream fileInputStream = new FileInputStream(miFileZS);
    Workbook workbookZS = new XSSFWorkbook(fileInputStream);
    XSSFSheet sheetZS;
    String numberCar;


    File AllCarsFile = new File("C:\\Users\\Professional\\Desktop\\Git\\dsv\\3.2021.xlsx");
    FileInputStream allCarsFileInputStream = new FileInputStream(AllCarsFile);
    Workbook allCarsFileWorkbook = new XSSFWorkbook(allCarsFileInputStream);


    public CarZS() throws IOException {
        copyAndSetCarZS();


    }

    public void copyAndSetCarZS() throws IOException {
        for (int i = 1; workbookZS.getSheetName(workbookZS.getSheetIndex(sheetZS) + i).length() > 7; i++) {
            numberCar = workbookZS.getSheetName(workbookZS.getSheetIndex(sheetZS) + i);
            System.out.println(numberCar);

            for (int j = 1; j < 43; j++) {
                System.out.println("Срока " + j + " скопирована");
                int cell0 = (int) workbookZS.getSheet(numberCar).getRow(j).getCell(0).getNumericCellValue();
                if (cell0 != 0) allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(0).setCellValue(
                        workbookZS.getSheet(numberCar).getRow(j).getCell(0).getNumericCellValue());

                int cell1 = (int) workbookZS.getSheet(numberCar).getRow(j).getCell(1).getNumericCellValue();
                if (cell1 != 0) allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(1).setCellValue(
                        workbookZS.getSheet(numberCar).getRow(j).getCell(1).getNumericCellValue());

                allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(2).setCellValue(
                        String.valueOf(workbookZS.getSheet(numberCar).getRow(j).getCell(2)));

                allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(3).setCellValue(
                        String.valueOf(workbookZS.getSheet(numberCar).getRow(j).getCell(3)));

                int cell4 = (int) workbookZS.getSheet(numberCar).getRow(j).getCell(4).getNumericCellValue();
                if (cell4 != 0) allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(4).setCellValue(
                        workbookZS.getSheet(numberCar).getRow(j).getCell(4).getNumericCellValue());

                int cell5 = (int) workbookZS.getSheet(numberCar).getRow(j).getCell(5).getNumericCellValue();
                if (cell5 != 0) allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(5).setCellValue(
                        workbookZS.getSheet(numberCar).getRow(j).getCell(5).getNumericCellValue());

                allCarsFileWorkbook.getSheet(numberCar).getRow(j).getCell(14).setCellValue(
                        String.valueOf(workbookZS.getSheet(numberCar).getRow(j).getCell(12)));

                FileOutputStream fileOutputStream = new FileOutputStream(AllCarsFile);
                allCarsFileWorkbook.write(fileOutputStream);
                fileOutputStream.close();
            }
        }
        fileInputStream.close();
    }
}

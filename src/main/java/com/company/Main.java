package com.company;

import com.company.apache.WriteTollCollect;
import com.company.swingMenu.SwingFirstMenu;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Scanner;


public class Main {


    public static void main(String[] args) throws IOException {

//        WriteTollCollect writeTollCollect = new WriteTollCollect();
//
//        SwingFirstMenu swingFirstMenu = new SwingFirstMenu();
//        swingFirstMenu.setVisible(true);

        File miFile = new File("/Users/mihajlomelnik/Documents/VAV TRANS/Інвойси 2021 LKW !.xlsx");
        FileInputStream fileInputStream = new FileInputStream(miFile);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        String result = String.valueOf(workbook.getSheet("01.02.2021").getRow(5).getCell(15));
        System.out.println(result);


    }
}

package com.company;

import com.company.apache.ExelFiles;
import com.company.apache.Invoice;
import com.company.apache.TollCollect;

import java.io.*;


public class Main {


    public static void main(String[] args) throws IOException {
//        SwingFirstMenu swingFirstMenu = new SwingFirstMenu();
//        swingFirstMenu.setVisible(true);

        Invoice invoice = new Invoice(new TollCollect("16.02.2021"));
        invoice.startFound();

    }

}

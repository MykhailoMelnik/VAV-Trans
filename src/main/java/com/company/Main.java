package com.company;

import com.company.apache.Walter;

import java.io.*;


public class Main {

    public static void main(String[] args) throws IOException {
//        SwingFirstMenu swingFirstMenu = new SwingFirstMenu();
//        swingFirstMenu.setVisible(true);

        Walter walter = new Walter();
        walter.start("19.01.2021");
        walter.totalAmountNumberDocInWalter("6023170227");
        walter.printAfterTotalAmount();
//        tollCollect.summationTollCollectByDates();

    }
}

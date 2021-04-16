package com.company;

import com.company.apache.TollCollect;
import com.company.apache.Walter;

import java.io.*;
import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;


public class Main {

    public static void main(String[] args) throws IOException {
//        SwingFirstMenu swingFirstMenu = new SwingFirstMenu();
//        swingFirstMenu.setVisible(true);


        Walter walter = new Walter();
        //  walter.foundAndWritIDS();
        walter.start("15.03.2021");

        // walter.totalAmountNumberDocInWalter("6023170227");
        //   walter.printAfterTotalAmount();
        //     tollCollect.summationTollCollectByDates();

        //   double sum = 222.4454;
        //   sum = new BigDecimal(sum).setScale(2, RoundingMode.HALF_UP).doubleValue();

    }
}

package com.company;

import com.company.apache.Walter;
import com.company.apache.TollCollect;

import java.io.*;


public class Main {

    public static void main(String[] args) throws IOException {
//        SwingFirstMenu swingFirstMenu = new SwingFirstMenu();
//        swingFirstMenu.setVisible(true);
        TollCollect tollCollect = new TollCollect();
//        Walter invoice = new Walter(tollCollect, "22.03.2021");
 //       invoice.startFound();
      tollCollect.summationTollCollectByDates();

    }
}

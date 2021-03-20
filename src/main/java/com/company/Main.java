package com.company;

import com.company.apache.Walter;
import com.company.apache.TollCollect;

import java.io.*;


public class Main {

    public static void main(String[] args) throws IOException {
//        SwingFirstMenu swingFirstMenu = new SwingFirstMenu();
//        swingFirstMenu.setVisible(true);

        Walter walter = new Walter();
        walter.start("15.02.2021");
//        tollCollect.summationTollCollectByDates();

    }
}

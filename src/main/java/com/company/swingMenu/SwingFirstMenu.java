package com.company.swingMenu;

import javax.swing.*;
import java.awt.*;
import java.io.File;

public class SwingFirstMenu extends JFrame {
    private JButton tollCollect = new JButton("Toll Collect");
    private JButton walterExel = new JButton("Walter exel");

    public SwingFirstMenu() {
        super("VAV TRANS");
        this.setBounds(100, 100, 250, 100);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        Container container = this.getContentPane();
        container.setLayout(new GridLayout(3, 2, 2, 2));

        container.add(tollCollect);
        tollCollect.addActionListener(e -> selectFile());
        container.add(walterExel);
        walterExel.addActionListener(e -> selectFile());
    }

    public void selectFile() {
        JFileChooser chooser = new JFileChooser();
        // optionally set chooser options ...
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File f = chooser.getSelectedFile();
            // read  and/or display the file somehow. ....
        } else {
            // user changed their mind
        }
    }
}

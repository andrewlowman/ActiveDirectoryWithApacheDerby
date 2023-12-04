package com.example.activedirectorywithapachederby.Utility;

import javax.swing.*;
import java.awt.*;

public class DialogPopUp {
    public void showDialog(Frame frame, String text){
        int width = 400;
        int height = 200;
        JDialog dialog = new JDialog(frame,"Help!");
        JLabel label = new JLabel(text);
        label.setHorizontalAlignment(JLabel.CENTER);
        label.setVerticalAlignment(JLabel.CENTER);
        JButton button = new JButton();

        button.addActionListener(actionEvent -> dialog.dispose());

        button.setVisible(true);
        button.setBounds(170,105,60,40);
        button.setText("OK");
        dialog.add(button);
        dialog.add(label);
        dialog.setSize(width,height);
        dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        dialog.setLocationRelativeTo(frame);
        dialog.setVisible(true);
    }
}
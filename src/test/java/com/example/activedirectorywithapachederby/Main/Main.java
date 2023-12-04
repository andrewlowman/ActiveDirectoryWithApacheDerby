package com.example.activedirectorywithapachederby.Main;

import com.example.activedirectorywithapachederby.Selenium.FillForm;
import com.example.activedirectorywithapachederby.Utility.DialogPopUp;
import com.example.activedirectorywithapachederby.Utility.LoadDepartmentList;
import com.example.activedirectorywithapachederby.Utility.LoadExcel;
import com.example.activedirectorywithapachederby.Utility.SearchExcelDepartmentList;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

public class Main extends JFrame {
    private JPanel mainPanel;
    private JTextField nameTextField;
    private JTextField deptTextField;
    private JTextField locationTextField;
    private JTextField mailCodeTextField;
    private JTextField userIDTextField;
    private JPasswordField pwTextField;
    private JButton openWindowButton;
    private JButton deptButton;
    private JButton excelButton;
    private JButton nextButton;
    private JTextField sheetTextField;
    private File excel = null;

    private File deptExcel = null;

    private LoadExcel loadExcel;

    private int counter = 0;
    private LoadDepartmentList loadDepartmentList = new LoadDepartmentList();
    private SearchExcelDepartmentList searchExcelDepartmentList = new SearchExcelDepartmentList();


    public Main(){
        FillForm fillForm = new FillForm();
        DialogPopUp dialogPopUp = new DialogPopUp();

        JFrame frame = new JFrame();
        frame.setContentPane(mainPanel);
        frame.setTitle("Update Active Directory");
        frame.setSize(750,600);

        //I think the program wasn't closing when I closed the window
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

        //placing the window to the right
        GraphicsConfiguration configuration = frame.getGraphicsConfiguration();
        Rectangle bounds =configuration.getBounds();
        Insets insets = Toolkit.getDefaultToolkit().getScreenInsets(configuration);
        int x = bounds.x + bounds.width - insets.right - frame.getWidth();
        int y = bounds.y + insets.top;

        frame.setLocation(x, y);

        frame.setVisible(true);

        deptButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                JFileChooser jFileChooser = new JFileChooser("C:\\Users\\low85\\Desktop");

                int jFile = jFileChooser.showOpenDialog(Main.this);

                if(jFile == JFileChooser.APPROVE_OPTION){
                   deptExcel = jFileChooser.getSelectedFile();

                }
            }
        });

        excelButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                if(sheetTextField.getText().isEmpty()){
                    dialogPopUp.showDialog(Main.this,"There is no sheet number");

                    return;
                }
                JFileChooser fileChooser = new JFileChooser("C:\\Users\\low85\\Desktop");

                int file = fileChooser.showOpenDialog(Main.this);

                if(file == JFileChooser.APPROVE_OPTION){
                    excel = fileChooser.getSelectedFile();
                    int sheetNumber = Integer.parseInt(sheetTextField.getText());
                    loadExcel = new LoadExcel(excel,sheetNumber - 1);
                }
            }
        });

        nextButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                int noResultCtr = 0;

                String firstName = loadExcel.getEmployeeFirstName(counter);
                String lastName = loadExcel.getEmployeeLastName(counter);

                if(firstName.equals("empty")||lastName.equals("empty")){
                    dialogPopUp.showDialog(Main.this,"You've reached the end");
                    return;
                }else{
                    nameTextField.setText(firstName + " " + lastName);
                }

                int emplID = loadExcel.getEmployeeID(counter);

                if(emplID==0){
                    dialogPopUp.showDialog(Main.this,"You already did this one");
                    counter++;
                    return;
                }

                String deptCode = String.valueOf(loadExcel.getDeptCode(counter));
                String deptName = loadExcel.getNameOfDept(counter);
                //trim dept name so it matches
                if(deptName.contains("'")){
                    deptName = deptName.replace("'","");
                }

                String locCode = searchExcelDepartmentList.getDeptLocationCodeByDeptName(deptName,deptExcel);
                String mailCode = searchExcelDepartmentList.getMailCodeByDeptName(deptName, deptExcel);

                //if no results from dept search add to no results counter
                if(locCode.equals("0")||mailCode.equals("0")||mailCode.equals("VARIOUS MAIL CODE")){
                    noResultCtr++;

                    locCode = searchExcelDepartmentList.getLocationCodeByDeptCode(deptCode,deptExcel);
                    mailCode = searchExcelDepartmentList.getMailCodeByDeptCode(deptCode, deptExcel);

                    if(locCode.equals("0")||mailCode.equals("0")||mailCode.equals("VARIOUS MAIL CODE")){
                        noResultCtr++;
                    }

                }

                if(noResultCtr >= 2){
                    dialogPopUp.showDialog(Main.this,"You're going to need to do it yourself");
                    counter++;
                    deptTextField.setText(deptName);
                    locationTextField.setText("");
                    mailCodeTextField.setText("");
                    fillForm.noDepartment(emplID);
                    return;
                }

                deptTextField.setText(deptName);
                locationTextField.setText(splitString(locCode));
                mailCodeTextField.setText(splitString(mailCode));
                fillForm.nextEntry(emplID,splitString(locCode),splitString(mailCode));

                counter++;
            }
        });

        openWindowButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                if(userIDTextField.getText().isEmpty()||pwTextField.getText().isEmpty()){
                    dialogPopUp.showDialog(Main.this,"Please enter a userid or password");
                }else{
                    fillForm.openWindow(userIDTextField.getText(),pwTextField.getText());
                }
            }
        });

    }

    public String splitString(String text){
        String newText = text;
        if(text.contains("/")){
            newText = text.split("/")[0];
        }

        return newText;
    }

    public static void main(String[] args) {
        new Main();
    }
}

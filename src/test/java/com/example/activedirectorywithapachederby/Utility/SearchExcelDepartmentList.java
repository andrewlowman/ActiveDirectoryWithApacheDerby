package com.example.activedirectorywithapachederby.Utility;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class SearchExcelDepartmentList {

//-------------------------------------------------------By Department Name----------------------------------------------
    public String getDeptLocationCodeByDeptName(String deptName, File excelFile) {
        try {
            FileInputStream fileInputStream = new FileInputStream(excelFile);

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet sheet = workbook.getSheetAt(0);

            int lastRow = sheet.getLastRowNum();

            Iterator<Row> rowIterator = sheet.iterator();

            rowIterator.next();

            //boolean done = false;

            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();

                String cellContents = nextRow.getCell(1).getStringCellValue();

                if (cellContents.equals(deptName)) {
                    return nextRow.getCell(2).getStringCellValue();
                }

            }

            workbook.close();
            return "0";

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /*public String getDeptLocationCodeByDeptNameAlternative(String deptName, File excelFile) {
        try {
            FileInputStream fileInputStream = new FileInputStream(excelFile);

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet sheet = workbook.getSheetAt(0);

            int lastRow = sheet.getLastRowNum();

            Iterator<Row> rowIterator = sheet.iterator();

            rowIterator.next();

            //boolean done = false;

            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();

                Cell cell = nextRow.getCell(1);

                String cellContents;

                if(cell != null){
                    switch(cell.getCellType()){
                        case STRING:
                            cellContents = nextRow.getCell(1).getStringCellValue();
                            break;
                        case NUMERIC:

                    }
                }



                if (cellContents.equals(deptName)) {
                    return nextRow.getCell(2).getStringCellValue();
                }

            }

            workbook.close();
            return "0";

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }*/

    public String getMailCodeByDeptName(String deptName, File excelFile) {
        try {
            FileInputStream fileInputStream = new FileInputStream(excelFile);

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet sheet = workbook.getSheetAt(0);

            int lastRow = sheet.getLastRowNum();

            Iterator<Row> rowIterator = sheet.iterator();

            rowIterator.next();

            //boolean done = false;

            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();

                String cellContents = nextRow.getCell(1).getStringCellValue();

                if (cellContents.equals(deptName)) {
                    return nextRow.getCell(4).getStringCellValue();
                }

            }

            workbook.close();
            return "0";

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public String getLocationDescriptionByDeptName(String deptName, File excelFile) {
        try {
            FileInputStream fileInputStream = new FileInputStream(excelFile);

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet sheet = workbook.getSheetAt(0);

            int lastRow = sheet.getLastRowNum();

            Iterator<Row> rowIterator = sheet.iterator();

            rowIterator.next();

            //boolean done = false;

            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();

                String cellContents = nextRow.getCell(1).getStringCellValue();

                if (cellContents.equals(deptName)) {
                    return nextRow.getCell(3).getStringCellValue();
                }

            }

            workbook.close();
            return "0";

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


    //--------------------------------------------------By Department Code---------------------------------------


    public String getLocationCodeByDeptCode(String code, File excelFile){

        try{
            FileInputStream fileInputStream = new FileInputStream(excelFile);

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();

            rowIterator.next();

            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();

                //have to trim leading zeroes to match other sheet
                String cellContents = nextRow.getCell(0).getStringCellValue();
                String finalString = cellContents.replaceFirst("^0","");
                //System.out.println("Final String " + finalString);

                if (finalString.equals(code)) {
                    return nextRow.getCell(2).getStringCellValue();
                }

            }

            workbook.close();
            return "0";
        }catch (IOException e){
            throw new RuntimeException(e);
        }
    }

    public String getMailCodeByDeptCode(String code, File excelFile){

        try{
            FileInputStream fileInputStream = new FileInputStream(excelFile);

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();

            rowIterator.next();

            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();

                //have to trim leading zeroes to match other sheet
                String cellContents = nextRow.getCell(0).getStringCellValue();
                String finalString = cellContents.replaceFirst("^0","");
                //System.out.println("Final String " + finalString);

                if (finalString.equals(code)) {
                    return nextRow.getCell(4).getStringCellValue();
                }

            }

            workbook.close();
            return "0";
        }catch (IOException e){
            throw new RuntimeException(e);
        }
    }

    public String getLocationDescriptionByDeptCode(String code, File excelFile){

        try{
            FileInputStream fileInputStream = new FileInputStream(excelFile);

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();

            rowIterator.next();

            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();

                //have to trim leading zeroes to match other sheet
                String cellContents = nextRow.getCell(0).getStringCellValue();
                String finalString = cellContents.replaceFirst("^0","");
                //System.out.println("Final String " + finalString);

                if (finalString.equals(code)) {
                    return nextRow.getCell(3).getStringCellValue();
                }

            }

            workbook.close();
            return "0";
        }catch (IOException e){
            throw new RuntimeException(e);
        }
    }
}

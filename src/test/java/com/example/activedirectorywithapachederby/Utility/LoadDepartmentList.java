package com.example.activedirectorywithapachederby.Utility;

import io.netty.handler.codec.DateFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Types;
import java.util.Iterator;

public class LoadDepartmentList {

    Connection conn = null;
    File excelDeptFile;

    int sheetNumber = 1;

    int batchSize = 20;

    public LoadDepartmentList(){

    }

    public File getExcelDeptFile() {
        return excelDeptFile;
    }

    public void setExcelDeptFile(File excelDeptFile) {
        this.excelDeptFile = excelDeptFile;
    }

    public int getSheetNumber() {
        return sheetNumber;
    }

    public void setSheetNumber(int sheetNumber) {
        this.sheetNumber = sheetNumber;
    }

    //SEARCH BY DEPT NAME--------------------------------------------------------------------------

    public int getRowNumberByName(String deptName){
        try{
            FileInputStream fileInputStream = new FileInputStream(excelDeptFile);

            XSSFWorkbook workbook = new XSSFWorkbook((fileInputStream));

            XSSFSheet sheet = workbook.getSheetAt(sheetNumber);

            DataFormatter formatter = new DataFormatter();

            for(Row row:sheet){
                for(Cell c:row){
                    String cellContents = formatter.formatCellValue(c);
                    if(cellContents.equals(deptName)){
                        return row.getRowNum();
                    }
                }
            }
        }catch(IOException e){
            throw new RuntimeException();
        }

        return -1;
    }

    //SEARCH BY NUMBER--------------------------------------------------------------------

    public int getRowNumberByDeptNumber(int deptCode){
        try{
            FileInputStream fileInputStream = new FileInputStream(excelDeptFile);

            XSSFWorkbook workbook = new XSSFWorkbook((fileInputStream));

            XSSFSheet sheet = workbook.getSheetAt(sheetNumber);

            //DataFormatter formatter = new DataFormatter();

            for(Row row:sheet){
                for(Cell c:row){
                    if(c.getColumnIndex() == 1){
                        if(c.getNumericCellValue() == deptCode){
                            return row.getRowNum();
                        }
                    }
                }
            }
        }catch(IOException e){
            throw new RuntimeException();
        }

        return -1;
    }

    /*public void loadDepartment(File file) {
        conn = DBConnection.startConnection();

        try {

            FileInputStream fileInputStream = new FileInputStream(file);

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();

            //conn = DriverManager.getConnection(jdbcURL,username,password);
            conn.setAutoCommit(false);

            String sql = "INSERT INTO DEPT_LIST_MC (Dept_Code, Dept_Name, Location_Code, Location_Desc, Mail_Code) VALUES (?,?,?,?,?)";

            PreparedStatement preparedStatement = conn.prepareStatement(sql);

            int count = 0;

            rowIterator.next();

            while(rowIterator.hasNext()){

                Row nextRow = rowIterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();

                while(cellIterator.hasNext()){
                    Cell nextCell = cellIterator.next();

                    int columnIndex = nextCell.getColumnIndex();

                    switch(columnIndex){
                        case 0:
                            String deptCode = nextCell.getStringCellValue();
                            preparedStatement.setString(1,deptCode);
                            break;
                        case 1:
                            if(nextCell.getCellType() == CellType.BLANK){
                                preparedStatement.setNull(2, Types.VARCHAR);
                            }else{String deptName = nextCell.getStringCellValue();
                                preparedStatement.setString(2,deptName);
                            }
                            break;
                        case 2:
                            if(nextCell.getCellType() == CellType.BLANK){
                                preparedStatement.setNull(3,Types.VARCHAR);
                            }else{
                                String locationCode = nextCell.getStringCellValue();
                                preparedStatement.setString(3,locationCode);
                            }
                            break;
                        case 3:
                            if(nextCell.getCellType() == CellType.BLANK){
                                preparedStatement.setNull(4,Types.VARCHAR);
                            }else {
                                String locationDesc = nextCell.getStringCellValue();
                                preparedStatement.setString(4, locationDesc);
                            }
                            break;
                        case 4:
                            if(nextCell.getCellType() == CellType.BLANK){
                                preparedStatement.setNull(5,Types.VARCHAR);
                            }else {
                                String mailCode = nextCell.getStringCellValue();
                                preparedStatement.setString(5, mailCode);
                            }
                            break;
                        *//*case 5:
                            if(nextCell.getCellType() == CellType.BLANK){
                                preparedStatement.setNull(6,Types.VARCHAR);
                            }else {
                                String mailCodeTwo = nextCell.getStringCellValue();
                                preparedStatement.setString(6, mailCodeTwo);
                            }
                            break;
                        case 6:
                            if(nextCell.getCellType() == CellType.BLANK){
                                preparedStatement.setNull(7,Types.VARCHAR);
                            }else {
                                String mailCodeThree = nextCell.getStringCellValue();
                                preparedStatement.setString(7, mailCodeThree);
                            }
                            break;
                        case 7:
                            if(nextCell.getCellType() == CellType.BLANK){
                                preparedStatement.setNull(8,Types.VARCHAR);
                            }else {
                                String otherLocation = nextCell.getStringCellValue();
                                preparedStatement.setString(8, otherLocation);
                            }
                            break;
                        case 8:
                            if(nextCell.getCellType() == CellType.BLANK){
                                preparedStatement.setNull(9,Types.VARCHAR);
                            }else {
                                String hospitalMC = nextCell.getStringCellValue();
                                preparedStatement.setString(9, hospitalMC);
                            }
                            break;*//*
                        default:
                            break;
                    }

                }
                preparedStatement.addBatch();

                if(count % batchSize == 0){
                    preparedStatement.executeBatch();
                }


                //JDialog dialog = new JDialog()
            }
            workbook.close();

            preparedStatement.executeBatch();

            conn.commit();
            DBConnection.closeConnection();

            long end = System.currentTimeMillis();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }
*/
}

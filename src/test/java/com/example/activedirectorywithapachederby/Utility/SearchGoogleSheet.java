package com.example.activedirectorywithapachederby.Utility;

public class SearchGoogleSheet {

    String spreadSheetID = "1DR3OM0CxCDx_qzQ0v9II4LM5YXhyrf3hxO0ZaJfDHI8";

    String sheetName = "Dept Location/MC";

    String range = "A1:E567";

    public SearchGoogleSheet(){

    }

    public String[][] getDataFromSheet() throws Exception {
        String[][] data = GoogleAuthorizeUtil.getData(spreadSheetID, sheetName, range);

        return data;
    }
//have to check array length to make sure no blank entries
    //-------------------------------------BY NAME--------------------------------------------
    public String getLocationCodeFromDeptNameFromGoogleSheet(String[][] data, String name){
        for(int i = 0; i < data.length; i++){
            if(data[i][1].equals(name)){
                //have to check array length to make sure no blank entries
                if(data[i].length < 5){
                    return "0";
                }else{
                    return data[i][2];
                }
            }
        }
        return "0";
    }

    public String getLocationDescriptionFromDeptNameFromGoogleSheet(String[][] data, String name){
        for(int i = 0; i < data.length; i++){
            if(data[i][1].equals(name)){
                return data[i][3];
            }
        }
        return "0";
    }

    public String getMailCodeFromDeptNameFromGoogleSheet(String[][] data, String name){
        for(int i = 0; i < data.length; i++){
            if(data[i][1].equals(name)){
                //have to check array length to make sure no blank entries
                if(data[i].length < 5){
                    return "0";
                }else{
                    return data[i][4];
                }
            }
        }
        return "0";
    }

    //---------------------------------- BY CODE--------------------------------------

    public String getLocCodeFromDepartmentCodeFromGoogleSheet(String[][] data, String name){
        for(int i = 0; i < data.length; i++){
            String dataString = data[i][0].replaceFirst("^0","");
            if(dataString.equals(name)){
                //have to check array length to make sure no blank entries
                if(data[i].length < 5){
                    return "0";
                }else{
                    return data[i][2];
                }
            }
        }
        return "0";
    }

    public String getLocDescFromDepartmentCodeFromGoogleSheet(String[][] data, String name){
        for(int i = 0; i < data.length; i++){
            if(data[i][0].equals(name)){
                return data[i][3];
            }
        }
        return "0";
    }

    public String getMailCodeFromDepartmentCodeFromGoogleSheet(String[][] data, String name){
        for(int i = 0; i < data.length; i++){
            String dataString = data[i][0].replaceFirst("^0","");
            if(dataString.equals(name)){
                //have to check array length to make sure no blank entries
                if(data[i].length < 5){
                    return "0";
                }else{
                    return data[i][4];
                }
            }
        }
        return "0";
    }
}

package com.example.activedirectorywithapachederby.Utility;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class DBConnection {

    private static final String apachedbURL = "jdbc:derby:memory:;create=false";

    private static Connection conn = null;


    public static Connection startConnection(){
        try{
            conn = DriverManager.getConnection(apachedbURL);

        }catch(SQLException e){
            throw new RuntimeException();
        }

        return conn;
    }

    public static void closeConnection(){
        try{
            conn.close();
        }catch(SQLException e){
            throw new RuntimeException();
        }
    }
}

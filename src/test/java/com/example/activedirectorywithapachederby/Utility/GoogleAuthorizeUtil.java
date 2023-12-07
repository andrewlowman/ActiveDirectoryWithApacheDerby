package com.example.activedirectorywithapachederby.Utility;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

public class GoogleAuthorizeUtil {

    public static Credential authorize() throws IOException, GeneralSecurityException {

        String credentialLocation = System.getProperty("user.dir") + "\\src\\test\\resources\\google-sheets-client-secret.json";

        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JacksonFactory.getDefaultInstance(), new FileReader(credentialLocation));

        List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS);

        GoogleAuthorizationCodeFlow googleAuthorizationCodeFlow = new GoogleAuthorizationCodeFlow
                .Builder(GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(), clientSecrets, scopes)
                .setDataStoreFactory(new FileDataStoreFactory(new File(System.getProperty("user.dir") + "/resources")))
                .setAccessType("offline")
                .build();

        return new AuthorizationCodeInstalledApp(googleAuthorizationCodeFlow, new LocalServerReceiver()).authorize("user");
    }

    public static String[][] getData(String spreadSheetID, String sheetName, String rangeDataToRead) throws Exception{

        Sheets sheet = new Sheets(GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(), authorize());

        List<List<Object>> data = sheet.spreadsheets().values().get(spreadSheetID, sheetName + "!" + rangeDataToRead).execute().getValues();

        return convertToArray(data);
    }

    private static String[][] convertToArray(List<List<Object>> data){
        String[][] array = new String[data.size()][];

        int i = 0;
        for(List<Object> row : data){
            array[i++] = row.toArray(new String[row.size()]);
        }

        return array;
    }
}

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class App {
    //Google sample code for quick start and efficient code 👇
    private static final String APPLICATION_NAME = "Google Sheets Tunts.Rocks challenge";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";

    /**
     * Global instance of the scopes required by this quickstart.
     * If modifying these scopes, delete your previously saved tokens/ folder.
     */
    private static final List<String> SCOPES =
            Collections.singletonList(SheetsScopes.SPREADSHEETS);
    private static final String CREDENTIALS_FILE_PATH = "/credentials.json";

    /**
     * Creates an authorized Credential object.
     *
     * @param HTTP_TRANSPORT The network HTTP Transport.
     * @return An authorized Credential object.
     * @throws IOException If the credentials.json file cannot be found.
     */
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
        InputStream in = App.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }
    //End of Google Code


    /**
     * Prints the names and average grade of students in a sample spreadsheet:
     * https://docs.google.com/spreadsheets/d/121NOYZS2cyFx0Lx0p8NgxXoj7ihjYW3MRvrUQEUxIWI/edit
     */
    public static void main(String... args) throws IOException, GeneralSecurityException {
        // Build a new authorized API client service.
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        final String spreadsheetId = "121NOYZS2cyFx0Lx0p8NgxXoj7ihjYW3MRvrUQEUxIWI"; //Id can be found on spreadsheet link
        final String range = "engenharia_de_software!A4:F"; //Range of cells in specific worksheet.

        final int numberOfClasses = 60;//Number of classes during the year
        final int maxMissings = (int)Math.ceil(numberOfClasses*0.25); //Max of missing classes during the year

        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT)).setApplicationName(APPLICATION_NAME).build(); //Connecting to the spreadsheet service
        ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).execute(); //Accessing spreadsheet values
        List<List<Object>> values = response.getValues(); //Getting values

        List<List<Object>> finalOutput = new ArrayList<>(); //Output variable, for rows and columns

        if (values == null || values.isEmpty()) {
            System.out.println("No data found."); //Empty spreasheet
        } else {
            System.out.println("Students situations");
            for (List row : values) {
                List<Object> tmpList = new ArrayList<>(); //Temporary list for rows
                int avgGrade = (Integer.parseInt(row.get(3).toString())+Integer.parseInt(row.get(4).toString())+Integer.parseInt(row.get(5).toString()))/3;
                int finalGrade = (int)Math.ceil(avgGrade*0.1);
                System.out.print(row.get(1)+" has "+row.get(2)+" absences, final grade of "+finalGrade+". Situation: ");
                if (Integer.parseInt(row.get(2).toString())>=maxMissings) { //Check number of absences
                    tmpList.add("Reprovado por Faltas");
                    tmpList.add("0");
                    System.out.print("Failed due to absence\n");
                } else {
                    if (finalGrade>=7) {
                        tmpList.add("Aprovado");
                        tmpList.add("0");
                        System.out.print("Approved\n");
                    } else if (finalGrade<5) {
                        tmpList.add("Reprovado por Nota");
                        tmpList.add("0");
                        System.out.print("Failed due to grade\n");
                    } else {
                        int gradeForApproval = (5*2-(finalGrade));
                        tmpList.add("Exame Final");
                        tmpList.add(Integer.toString(gradeForApproval));
                        System.out.print("Final exam, need "+gradeForApproval+" points for approval\n");
                    }
                }
                finalOutput.add(tmpList); //Add row of situation and grade for approval
            }
            ValueRange sitBody = new ValueRange().setValues(finalOutput); //Create body for value update
            service.spreadsheets().values().update(spreadsheetId, "G4",sitBody).setValueInputOption("RAW").execute(); //Runs the update
        }
    }
}
package sample;

import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonType;
import javafx.scene.control.TextInputDialog;

import java.util.Optional;


public class DisplayInformation {

    // Show a Information Alert with header Text
    private void showAlertWithHeaderText() {

        Alert alert = new Alert(AlertType.INFORMATION);

        alert.setTitle("Test Connection");

        alert.setHeaderText("Results:");

        alert.setContentText("Connect to the database successfully!");

        alert.showAndWait();

    }


    // Show a Information Alert with default header Text
    public static void showErrorAlertWithHeaderText(String title, String header, String message) {

        Alert alert = new Alert(AlertType.INFORMATION);

        alert.setTitle(title);

        // alert.setHeaderText("Results:");
        alert.setContentText(message);

        // Header Text: null
        alert.setHeaderText(header);

        alert.showAndWait();

    }

    // Show a Information Alert with default header Text
    public static void showAlertWithHeaderText(String title, String header, String message) {

        Alert alert = new Alert(AlertType.INFORMATION);

        alert.setTitle(title);

        // alert.setHeaderText("Results:");
        alert.setContentText(message);

        // Header Text: null
        alert.setHeaderText(header);

        alert.showAndWait();

    }


    // Show a Information Alert without Header Text
    public static void showInformationMessage(String title, String message) {

        Alert alert = new Alert(AlertType.INFORMATION);

        alert.setTitle(title);

        alert.setHeaderText(null);

        alert.setContentText(message);

        alert.showAndWait();

    }



    // Show a Information Alert without Header Text
    public static void showErrorMessage(String title, String message) {

        Alert alert = new Alert(AlertType.ERROR);

        alert.setTitle(title);

        // Header Text: null
        alert.setHeaderText(null);

        alert.setContentText(message);

        alert.showAndWait();

    }


    //Show confirmation dialog
    public  static boolean confirmAction(String title, String message) {

        Alert alert = new Alert(AlertType.CONFIRMATION, message, ButtonType.NO, ButtonType.YES);

        alert.setTitle(title);

        alert.setHeaderText(null);

        alert.showAndWait();

        return alert.getResult() == ButtonType.YES;
    }



    //Show confirmation dialog
    public  static boolean getUserResponse(String title, String message) {

        Alert alert = new Alert(AlertType.CONFIRMATION, message, ButtonType.NO, ButtonType.YES);

        alert.setTitle(title);

        alert.setHeaderText(null);

        alert.showAndWait();

        return alert.getResult() == ButtonType.YES;
    }


    static String r = "";

    public static String getUserInput(String title, String message) {

        TextInputDialog pDialog  = new TextInputDialog();

        pDialog.setTitle( title);

        pDialog.setHeaderText(null);
        pDialog.setContentText(message);

        Optional<String> result = pDialog.showAndWait();

        result.ifPresent( name -> {

            r = name;

        });

        return r;

    }


    public static void showDirectErrorMessage() {

        Alert alert = new Alert(AlertType.INFORMATION);

        alert.setTitle("Error");

        alert.setHeaderText(null);

        alert.setContentText("Oops!! Something went Wrong");

        alert.showAndWait();

    }

}
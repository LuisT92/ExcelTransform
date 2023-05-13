package com.example.exceltransform;

import javafx.fxml.FXML;
import javafx.scene.control.Label;

public class ExcelTransformController {
    @FXML
    private Label welcomeText;

    @FXML
    protected void onHelloButtonClick() {
        welcomeText.setText("Welcome to JavaFX Application!");
    }
}
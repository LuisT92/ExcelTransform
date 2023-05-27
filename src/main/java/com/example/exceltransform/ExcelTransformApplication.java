package com.example.exceltransform;

/**
 * @version 1.0
 * @autor Luis Humberto Torres Escogido
 * @since 27/05/2023
 * @see ExcelTransformController
 */
//importaciones de librerias de javafx para la interfaz grafica de usuario (GUI) y para la ejecucion de la aplicacion
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;
import java.io.IOException;


public class ExcelTransformApplication extends Application { //clase principal de la aplicacion
    @Override
    public void start(Stage stage) throws IOException { //metodo que se ejecuta al iniciar la aplicacion
        FXMLLoader fxmlLoader = new FXMLLoader(ExcelTransformApplication.class.getResource("ExcelTransform-view.fxml")); //carga el archivo fxml que contiene la interfaz grafica de usuario
        Scene scene = new Scene(fxmlLoader.load(), 600, 400); //crea la escena con el archivo fxml cargado
        stage.setTitle("Separa Nombres"); //titulo de la ventana
        stage.setScene(scene); //agrega la escena a la ventana
        stage.show(); //muestra la ventana
    }

    public static void main(String[] args) { //metodo principal de la aplicacion
        launch(); //inicia la aplicacion
    }
}
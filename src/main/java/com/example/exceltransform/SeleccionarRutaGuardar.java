package com.example.exceltransform;

import javax.swing.*;

public class SeleccionarRutaGuardar {
    public String seleccionarRutaGuardar() {
        JFileChooser rutaCarpeta = new JFileChooser();
        rutaCarpeta.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int resultado = rutaCarpeta.showOpenDialog(null);

        return resultado == JFileChooser.APPROVE_OPTION ? rutaCarpeta.getSelectedFile().getAbsolutePath() : "";

    }
}

package com.example.exceltransform;


import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

public class ObtenerRutaArchivo {
    public String obtenerRutaArchivo() {
        JFileChooser rutaArchivo = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivos XLSX", "xlsx");
        rutaArchivo.setFileFilter(filter);
        int resultado = rutaArchivo.showOpenDialog(null);
        return resultado == JFileChooser.APPROVE_OPTION ? rutaArchivo.getSelectedFile().getAbsolutePath() : "";
    }
}

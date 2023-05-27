package com.example.exceltransform;

import javafx.fxml.FXML;
import javafx.scene.control.TextField;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;



public class ExcelTransformController {

    String ruta = "";
    String carpetaGuardar = "";
    List<String> celdasRFC12 = new ArrayList<>();
    List<String> celdasNombresPFoPM12 = new ArrayList<>();

    List<String> celdasRFC13 = new ArrayList<>();
    List<String> celdasNombresPFoPM13 = new ArrayList<>();
    List <String> celdasTipoPM = new ArrayList<>();

    @FXML
    private TextField PathArchivo;
    @FXML
    private TextField PathCarpeta;
    @FXML
    protected void buttonObtenerRutaArchivo() {
        FileChooser rutaArchivo = new FileChooser();
        rutaArchivo.setTitle("Seleccionar archivo");
        rutaArchivo.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Excel", "*.xlsx")
        );
        File archivoSeleccionado = rutaArchivo.showOpenDialog(null);
        if (archivoSeleccionado != null) {
            ruta = archivoSeleccionado.getAbsolutePath(); // Asignar la ruta seleccionada al atributo ruta
            PathArchivo.setText(ruta);
        }
    }


    @FXML
    protected void buttonSeleccionarCarpetaGuardar() {
        DirectoryChooser rutaCarpeta = new DirectoryChooser();
        rutaCarpeta.setTitle("Seleccionar carpeta");
        carpetaGuardar = rutaCarpeta.showDialog(null).getAbsolutePath();
        PathCarpeta.setText(carpetaGuardar);
        }

    @FXML
    protected void buttonGenerarArchivos() {
        SepararPMdePF separarPMdePF = new SepararPMdePF();
        PMTipo pmTipo = new PMTipo();
        PFSepararNombre pfSepararNombre = new PFSepararNombre();
        separarPMdePF.SepararPMdePF();
        pfSepararNombre.PFSepararNombreyGuardar();
        pmTipo.PMTipo();
        pmTipo.generarArchivoYGuardar();
        //cerrar programa
        System.exit(0);
    }

    public class SepararPMdePF {

        public void SepararPMdePF() {
            FileInputStream archivo;

            try {
                archivo = new FileInputStream(new File(ruta));
                Workbook libro = new XSSFWorkbook(archivo);
                Sheet hoja = libro.getSheetAt(0);

                for (Row row : hoja) {
                    Cell celdaRFC = row.getCell(0);
                    Cell celdaNombre = row.getCell(1);

                    String rfc = celdaRFC.getStringCellValue();

                    String name = "";
                    if (celdaNombre != null) {
                        name = celdaNombre.getStringCellValue();
                    }

                    int longitudRFC = rfc.length();

                    if (longitudRFC == 12) {
                        celdasRFC12.add(rfc);
                        celdasNombresPFoPM12.add(name);
                    }else{
                        celdasRFC13.add(rfc);
                        celdasNombresPFoPM13.add(name);
                    }
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }


    public class PMTipo{

        List<String> tiposSociedades = Arrays.asList("S. DE R.L. DE C.V.", "S.A. DE C.V. SOFOM. ENR.", "SA DE CV SOFOM ENR",
                "S.A.P.I. de C.V.", "S DE RL DE CV", "S A P I de C.V.",
                "S.A. de C.V.", "S.A.P.I. DE C.V. SOFOM. ENR.", "S.C. de C.V.", "SA DE CV SOFOM ENR",
                "S.C. de CV", "S.A.S. de C.V.", "S.A.", "S.C.", "SC de CV", "SC", "SAPI de CV",
                "SA DE CV", "SAS de CV", "S.A.S. de C.V.", "SAPI", "SAS", "AC", "A.C.");
        public void PMTipo() {
            for (int i = 0; i < celdasNombresPFoPM12.size(); i++) {
                for (int j = 0; j < tiposSociedades.size(); j++) {
                    if (celdasNombresPFoPM12.get(i).contains(tiposSociedades.get(j))) {
                        celdasTipoPM.add(tiposSociedades.get(j));
                        celdasNombresPFoPM12.set(i, celdasNombresPFoPM12.get(i).replace(tiposSociedades.get(j), ""));
                    }
                }
            }
        }

        public void generarArchivoYGuardar(){
            Workbook libro = new XSSFWorkbook();
            Sheet hoja = libro.createSheet("PM");
            int maxSize = Math.min(celdasRFC12.size(), Math.min(celdasNombresPFoPM12.size(), celdasTipoPM.size()));
            for(int i = 0; i < maxSize; i++){
                Row fila = hoja.createRow(i);
                Cell celdaRFC = fila.createCell(0);
                Cell celdaNombre = fila.createCell(1);
                Cell celdaTipo = fila.createCell(2);
                celdaRFC.setCellValue(celdasRFC12.get(i));
                celdaNombre.setCellValue(celdasNombresPFoPM12.get(i));
                celdaTipo.setCellValue(celdasTipoPM.get(i));
            }
            try {
                FileOutputStream archivo = new FileOutputStream(carpetaGuardar + "/PM.xlsx");
                libro.write(archivo);
                archivo.close();
            } catch (IOException e) {
                System.out.println("Error al guardar archivo");
            }
        }
    }
    public class PFSepararNombre{

        public void PFSepararNombreyGuardar(){
            Workbook libro = new XSSFWorkbook();
            Sheet hoja = libro.createSheet("PF");

            for(int i=0; i < celdasRFC13.size(); i++){
                Row fila = hoja.createRow(i);
                Cell celdaRFC = fila.createCell(0);
                celdaRFC.setCellValue(celdasRFC13.get(i));
            }

            for(int i =0; i < celdasNombresPFoPM13.size(); i++){
                String [] nombreSeparado = celdasNombresPFoPM13.get(i).split(" ");
                Row fila = hoja.getRow(i);

                if(fila == null){
                    fila = hoja.createRow(i);
                }

                for(int j = 0; j < nombreSeparado.length; j++){
                    Cell celdaNombre = fila.createCell(j+1);
                    celdaNombre.setCellValue(nombreSeparado[j]);
                }
            }
            try {
                FileOutputStream archivo = new FileOutputStream(carpetaGuardar + "/PF.xlsx");
                libro.write(archivo);
                archivo.close();
            } catch (IOException e) {
                System.out.println("Error al guardar archivo");
            }
        }
    }
}
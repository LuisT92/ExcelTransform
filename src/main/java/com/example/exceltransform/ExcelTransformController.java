package com.example.exceltransform;

/**
 * @version 1.0
 * @autor Luis Humberto Torres Escogido
 * @since 27/05/2023
 * Programa para separar las personas morales de las personas fisicas de un archivo de Excel
 * y generar archivos de Excel con los RFC y nombres de las personas morales y fisicas por separado
 */

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
//Librerias Usadas en el proyecto


public class ExcelTransformController {

    String ruta = ""; // Atributo para guardar la ruta del archivo seleccionado
    String carpetaGuardar = ""; // Atributo para guardar la ruta de la carpeta seleccionada
    List<String> celdasRFC12 = new ArrayList<>(); // Lista para guardar los RFC de 12 caracteres de las personas morales
    List<String> celdasNombresPFoPM12 = new ArrayList<>(); // Lista para guardar los nombres de las personas morales

    List<String> celdasRFC13 = new ArrayList<>(); // Lista para guardar los RFC de 13 caracteres de las personas fisicas
    List<String> celdasNombresPFoPM13 = new ArrayList<>(); // Lista para guardar los nombres de las personas fisicas
    List <String> celdasTipoPM = new ArrayList<>(); // Lista para guardar los tipos de personas morales

    @FXML
    private TextField PathArchivo; // Campo de texto para mostrar la ruta del archivo seleccionado
    @FXML
    private TextField PathCarpeta; // Campo de texto para mostrar la ruta de la carpeta seleccionada
    @FXML
    protected void buttonObtenerRutaArchivo() { // Método para obtener la ruta del archivo seleccionado
        FileChooser rutaArchivo = new FileChooser();
        rutaArchivo.setTitle("Seleccionar archivo");
        rutaArchivo.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Excel", "*.xlsx") // Filtro para mostrar solo archivos con extensión .xlsx
        );
        File archivoSeleccionado = rutaArchivo.showOpenDialog(null);
        if (archivoSeleccionado != null) {
            ruta = archivoSeleccionado.getAbsolutePath(); // Asignar la ruta seleccionada al atributo ruta del archivo asociado al botón "Seleccionar archivo"
            PathArchivo.setText(ruta); // Mostrar la ruta seleccionada en el campo de texto
        }
    }


    @FXML
    protected void buttonSeleccionarCarpetaGuardar() { // Método para obtener la ruta de la carpeta seleccionada asociado al botón "Seleccionar carpeta"
        DirectoryChooser rutaCarpeta = new DirectoryChooser();
        rutaCarpeta.setTitle("Seleccionar carpeta");
        carpetaGuardar = rutaCarpeta.showDialog(null).getAbsolutePath(); // Asignar la ruta seleccionada al atributo carpetaGuardar
        PathCarpeta.setText(carpetaGuardar); // Mostrar la ruta seleccionada en el campo de texto
        }

    @FXML
    protected void buttonGenerarArchivos() { // Método para generar los archivos asociado al botón "Generar archivos"
        SepararPMdePF separarPMdePF = new SepararPMdePF(); // Crear objeto de la clase SepararPMdePF
        PMTipo pmTipo = new PMTipo(); // Crear objeto de la clase PMTipo
        PFSepararNombre pfSepararNombre = new PFSepararNombre(); // Crear objeto de la clase PFSepararNombre
        separarPMdePF.SepararPMdePF(); // Llamar al método SepararPMdePF
        pfSepararNombre.PFSepararNombreyGuardar(); // Llamar al método PFSepararNombreyGuardar
        pmTipo.PMTipo(); // Llamar al método PMTipo
        pmTipo.generarArchivoYGuardar(); // Llamar al método generarArchivoYGuardar
        System.exit(0); // Cerrar la aplicación
    }

    public class SepararPMdePF { // Clase para separar las personas morales de las personas fisicas segun el tamaño de su RFC

        public void SepararPMdePF() { // Método para separar las personas morales de las personas fisicas segun el tamaño de su RFC
            FileInputStream archivo;

            try {
                archivo = new FileInputStream(new File(ruta)); // Abrir el archivo seleccionado
                Workbook libro = new XSSFWorkbook(archivo);
                Sheet hoja = libro.getSheetAt(0);

                for (Row row : hoja) {
                    Cell celdaRFC = row.getCell(0); // Obtener la celda de la columna RFC
                    Cell celdaNombre = row.getCell(1); // Obtener la celda de la columna Nombre

                    String rfc = celdaRFC.getStringCellValue(); // Obtener el valor de la celda RFC

                    String name = "";
                    if (celdaNombre != null) {
                        name = celdaNombre.getStringCellValue(); // Obtener el valor de la celda Nombre
                    }

                    int longitudRFC = rfc.length(); // Obtener la longitud del RFC

                    if (longitudRFC == 12) { // Si la longitud del RFC es 12, se trata de una persona moral
                        celdasRFC12.add(rfc); // Agregar el RFC a la lista de RFC de personas morales
                        celdasNombresPFoPM12.add(name); // Agregar el nombre a la lista de nombres de personas morales
                    }else{
                        celdasRFC13.add(rfc); // Agregar el RFC a la lista de RFC de personas fisicas
                        celdasNombresPFoPM13.add(name); // Agregar el nombre a la lista de nombres de personas fisicas
                    }
                }
            } catch (IOException e) {
                throw new RuntimeException(e); // Lanzar excepción en caso de error
            }
        }
    }


    public class PMTipo{ // Clase para separar los tipos de personas morales

        List<String> tiposSociedades = Arrays.asList("S. DE R.L. DE C.V.", "S.A. DE C.V. SOFOM. ENR.", "SA DE CV SOFOM ENR",
                "S.A.P.I. de C.V.", "S DE RL DE CV", "S A P I de C.V.",
                "S.A. de C.V.", "S.A.P.I. DE C.V. SOFOM. ENR.", "S.C. de C.V.", "SA DE CV SOFOM ENR",
                "S.C. de CV", "S.A.S. de C.V.", "S.A.", "S.C.", "SC de CV", "SC", "SAPI de CV",
                "SA DE CV", "SAS de CV", "S.A.S. de C.V.", "SAPI", "SAS", "AC", "A.C."); // Lista de tipos de personas morales (Agregar más tipos de personas morales en caso de ser necesario)
        public void PMTipo() { // Método para separar los tipos de personas morales
            for (int i = 0; i < celdasNombresPFoPM12.size(); i++) { // Recorrer la lista de nombres de personas morales
                for (int j = 0; j < tiposSociedades.size(); j++) { // Recorrer la lista de tipos de personas morales
                    if (celdasNombresPFoPM12.get(i).contains(tiposSociedades.get(j))) { // Si el nombre de la persona moral contiene el tipo de persona moral en la posición j de la lista de tipos de personas morales (Ejemplo: "S. DE R.L. DE C.V.") entonces...
                        celdasTipoPM.add(tiposSociedades.get(j)); // Agregar el tipo de persona moral a la lista de tipos de personas morales
                        celdasNombresPFoPM12.set(i, celdasNombresPFoPM12.get(i).replace(tiposSociedades.get(j), "")); // Remover el tipo de persona moral del nombre de la persona moral
                    }
                }
            }
        }

        public void generarArchivoYGuardar(){ // Método para generar el archivo de personas morales y guardar
            Workbook libro = new XSSFWorkbook(); // Crear el libro de Excel
            Sheet hoja = libro.createSheet("PM"); // Crear la hoja de Excel
            int maxSize = Math.min(celdasRFC12.size(), Math.min(celdasNombresPFoPM12.size(), celdasTipoPM.size())); // Obtener el tamaño máximo de las listas de RFC, nombres y tipos de personas morales (Se obtiene el mínimo de los tamaños de las listas) para evitar errores
            for(int i = 0; i < maxSize; i++){ // Recorrer las listas de RFC, nombres y tipos de personas morales
                Row fila = hoja.createRow(i); // Crear una fila en la hoja de Excel
                Cell celdaRFC = fila.createCell(0); // Crear una celda en la fila de la columna RFC
                Cell celdaNombre = fila.createCell(1); // Crear una celda en la fila de la columna Nombre
                Cell celdaTipo = fila.createCell(2); // Crear una celda en la fila de la columna Tipo
                celdaRFC.setCellValue(celdasRFC12.get(i)); // Agregar el RFC a la celda de la columna RFC
                celdaNombre.setCellValue(celdasNombresPFoPM12.get(i)); // Agregar el nombre a la celda de la columna Nombre
                celdaTipo.setCellValue(celdasTipoPM.get(i)); // Agregar el tipo de persona moral a la celda de la columna Tipo
            }
            try {
                FileOutputStream archivo = new FileOutputStream(carpetaGuardar + "/PM.xlsx"); // Crear el archivo de Excel en la carpeta de guardar con el nombre "PM.xlsx" (Ejemplo: C:/Users/Usuario/Desktop/PM.xlsx), se puede cambiar el nombre del archivo si se desea
                libro.write(archivo); // Escribir el libro de Excel en el archivo
                archivo.close(); // Cerrar el archivo
            } catch (IOException e) {
                System.out.println("Error al guardar archivo"); // Mostrar mensaje en caso de error
            }
        }
    }
    public class PFSepararNombre{ // Clase para separar el nombre de las personas fisicas

        public void PFSepararNombreyGuardar(){ // Método para separar el nombre de las personas fisicas y guardar
            Workbook libro = new XSSFWorkbook(); // Crear el libro de Excel
            Sheet hoja = libro.createSheet("PF"); // Crear la hoja de Excel

            for(int i=0; i < celdasRFC13.size(); i++){ // Recorrer la lista de RFC de personas fisicas
                Row fila = hoja.createRow(i); // Crear una fila en la hoja de Excel
                Cell celdaRFC = fila.createCell(0); // Crear una celda en la fila de la columna RFC
                celdaRFC.setCellValue(celdasRFC13.get(i)); // Agregar el RFC a la celda de la columna RFC
            }

            for(int i =0; i < celdasNombresPFoPM13.size(); i++){ // Recorrer la lista de nombres de personas fisicas
                String [] nombreSeparado = celdasNombresPFoPM13.get(i).split(" "); // Separar el nombre de la persona fisica por espacios (Ejemplo: "Juan Perez" -> "Juan", "Perez") y agregarlo a un arreglo de Strings (Ejemplo: ["Juan", "Perez"])
                Row fila = hoja.getRow(i); // Obtener la fila de la hoja de Excel en la posición i

                if(fila == null){ // Si la fila es nula entonces...
                    fila = hoja.createRow(i); // Crear una fila en la hoja de Excel en la posición i
                }

                for(int j = 0; j < nombreSeparado.length; j++){ // Recorrer el arreglo de Strings con el nombre separado
                    Cell celdaNombre = fila.createCell(j+1); // Crear una celda en la fila de la columna Nombre
                    celdaNombre.setCellValue(nombreSeparado[j]); // Agregar el nombre separado a la celda de la columna Nombre
                }
            }
            try {
                FileOutputStream archivo = new FileOutputStream(carpetaGuardar + "/PF.xlsx"); // Crear el archivo de Excel en la carpeta de guardar con el nombre "PF.xlsx" (Ejemplo: C:/Users/Usuario/Desktop/PF.xlsx), se puede cambiar el nombre del archivo si se desea
                libro.write(archivo); // Escribir el libro de Excel en el archivo
                archivo.close(); // Cerrar el archivo
            } catch (IOException e) {
                System.out.println("Error al guardar archivo"); // Mostrar mensaje en caso de error
            }
        }
    }
}
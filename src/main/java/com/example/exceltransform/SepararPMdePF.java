package com.example.exceltransform;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SepararPMdePF {
    public void separarPMdePF(String rutaArchivo) {

        List<String> celdasRFC12 = new ArrayList<>();
        List<String> celdasNombresPFoPM12 = new ArrayList<>();

        List<String> celdasRFC13 = new ArrayList<>();
        List<String> celdasNombresPFoPM13 = new ArrayList<>();

        try{
            File archivo = new File(rutaArchivo);
            FileInputStream archivoLeido = new FileInputStream(archivo);

            XSSFWorkbook libroExcel = new XSSFWorkbook(archivoLeido);
            XSSFSheet hoja = libroExcel.getSheetAt(0);

            for (Row fila : hoja) {
                Cell RFC = fila.getCell(0);
                Cell NombrePFoPM = fila.getCell(1);

                if (RFC != null) {
                    String contenidoRFC = RFC.getStringCellValue();
                    int longitudRFC = contenidoRFC.length();

                    if (longitudRFC == 12) {
                        celdasRFC12.add(contenidoRFC);
                        if (NombrePFoPM != null) {
                            String contenidoNombres = NombrePFoPM.getStringCellValue();
                            celdasNombresPFoPM12.add(contenidoNombres);
                        } else {
                            celdasNombresPFoPM12.add("");
                        }
                    }
                    if (longitudRFC == 13) {
                        celdasRFC13.add(contenidoRFC);
                        if (NombrePFoPM != null) {
                            String contenidoNombres = NombrePFoPM.getStringCellValue();
                            celdasNombresPFoPM13.add(contenidoNombres);
                        } else {
                            celdasNombresPFoPM13.add("");
                        }
                    }
                }
            }
            archivoLeido.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        PMTipo pmTipo12 = new PMTipo(celdasRFC12, celdasNombresPFoPM12);
        PFSepararNombre pfSepararNombre13 = new PFSepararNombre(celdasRFC13, celdasNombresPFoPM13);
    }
}

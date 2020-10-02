/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package prueba_tecnica;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author janer
 */
public class Prueba_Tecnica {

    public Prueba_Tecnica(File fileName, File file) {
        List cellData = new ArrayList();
        try {
            //Obtiene bytes de entrada desde un archivo en un sistema de archivos
            FileInputStream fileInpuStream = new FileInputStream(fileName);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInpuStream);

            //Obtiene hoja en la posicion dada
            XSSFSheet hssfSheet = workbook.getSheetAt(0);
            Iterator rowIterator = hssfSheet.rowIterator();

            while (rowIterator.hasNext()) {
                //Obtenemos los datos de la fila 0
                XSSFRow hssfRow = (XSSFRow) rowIterator.next();
                Iterator iterator = hssfRow.cellIterator();
                List cellTemp = new ArrayList();

                //Con el cliclo while nos movemos por los datos de cada fila
                while (iterator.hasNext()) {
                    //Se almacenan los datos de cada celda en el hssfCell
                    XSSFCell hssfCell = (XSSFCell) iterator.next();
                    //Y los datos almacenados en el hssfCell los almacenamos en el cellTemp
                    cellTemp.add(hssfCell);
                }
                //Se almacenan los datos en la lista cellData obtenidos de la lista cellTemp
                cellData.add(cellTemp);
            }
            //Llamamos el metodo obtener y se le pasa como parametro el array cellData
            obtener(cellData);

            //Se añade nueva hoja en el libro de excel y se le asigna un nombre
            XSSFSheet sheet1 = workbook.createSheet("Output");

            //contenido de la nueva hoja de excel
            int[][] document = new int[][]{
                {0}, {3}, {5}, {0}, {0}, {3}, {3}, {5}, {0}, {0}, {0}, {0}, {6},
                {5}, {2}, {0}, {0}, {6}, {6}, {4}, {0}, {6}, {6}, {6}, {2}, {0},
                {-1}, {-1}, {2}, {0}, {-1}, {-1}, {-1}, {3}, {0}, {-1}, {0},
                {-1}, {3}, {0}, {-1}, {-1}, {2}, {0}, {34}, {38}, {40}, {40},
                {0}, {38}, {0}, {99000000}, {0}, {1}, {0}, {0}
            };

            //generar los datos para la nueva hoja
            for (int i = 0; i < document.length; i++) {
                //se crean las filas
                XSSFRow row = sheet1.createRow(i);
                for (int j = 0; j < document[i].length; j++) {
                    //se crean las celdas, junto con la posición
                    XSSFCell cell = row.createCell(j);
                    //se añade el contenido
                    cell.setCellValue(document[i][j]);
                }
            }

            //Se crea la nueva hoja en el libro de excel
            try {
                FileOutputStream fileOuS = new FileOutputStream(file);
                workbook.write(fileOuS);
                fileOuS.flush();
                fileOuS.close();
                System.out.println();
                System.out.println("Archivo output fue creado, revisar archivo "
                        + "medium.xlsx dentro de la carpeta del proyecto");
            } catch (Exception e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //Metodo para leer el archivo excel de entrada
    private void obtener(List cellDataList) {
        for (int i = 0; i < cellDataList.size(); i++) {
            //Se obtienen los datos de cellDataList y almacenan en cellTempList
            List cellTempList = (List) cellDataList.get(i);
            for (int j = 0; j < cellTempList.size(); j++) {
                XSSFCell hssfCell = (XSSFCell) cellTempList.get(j);
                //Se convierten los datos a String
                String stringCellValue = hssfCell.toString();
                //Se imprimen los datos de la hoja por consola
                System.out.print(stringCellValue + "   |   ");
            }
            System.out.println();
        }
    }

    public static void main(String[] args) {
        //Ruta donde esta creado el libro de excel medium.xlsx
        File f = new File("C:/Users/janer/OneDrive/Escritorio/Prueba_Tecnica/medium.xlsx");

        //Ruta donde se creara la nueva hoja de excel en el libro medium.xlsx
        File fe = new File("C:\\Users\\janer\\OneDrive\\Escritorio\\Prueba_Tecnica\\medium.xlsx");
        if (f.exists()) {
            Prueba_Tecnica obj = new Prueba_Tecnica(f, fe);
        } else {
            System.out.println("Archivo no existe");
        }
    }
}
